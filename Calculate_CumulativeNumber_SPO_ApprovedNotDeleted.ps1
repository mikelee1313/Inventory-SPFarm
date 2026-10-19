<#
.SYNOPSIS
    Counts Microsoft Purview disposition items whose ReviewAction is
    RecordReviewed and exports one summary row per retention label to CSV.

.DESCRIPTION
    This script:

    1. Connects interactively to Exchange Online PowerShell.
    2. Displays matching effective Exchange Online Mailbox Import Export
       role assignments as a diagnostic check.
    3. Connects interactively to Microsoft Security and Compliance PowerShell.
    4. Retrieves Microsoft Purview retention labels.
    5. Processes either one configured label or all labels identified as having
       disposition review.
    6. Uses Get-ReviewItems -Disposed $true to retrieve disposed items.
    7. Processes every page returned by Get-ReviewItems.
    8. Counts items whose ReviewAction equals RecordReviewed.
    9. Does not filter or evaluate items based on DeletedDate.
    10. Does not retain item-level results in memory after processing each row.
    11. Writes one summary row per processed label to a CSV.
    12. Renames the CSV at completion using the completion date and time.

.NOTES
    Important:

    Get-ReviewItems must not be called with -ErrorAction Stop in this
    environment. Testing confirmed that the command succeeds with its default
    error behavior but returns a generic server-side error when
    -ErrorAction Stop is supplied.

    The global $ErrorActionPreference can remain set to Stop.

    Retry behavior:

    - The reported server-side "Please try again after some time" error is
      retried on the same page, without advancing the cookie or item counts.
    - ReviewItemRetryCount defaults to 10 retries after the initial request
      (11 attempts maximum per page). Set it to 0 to disable retries.
    - ReviewItemRetryDelaySeconds defaults to 30 seconds between retries.
    - Other errors are not retried. After exhaustion, the original error is
      passed to the existing per-label failure and CSV error handling.

    Role diagnostics:

    - Exchange Online role enumeration is diagnostic only. No matching rows
      does not prove that the compliance session lacks cmdlet access.
    - Required commands are checked after Connect-IPPSSession succeeds.

    Label behavior:

    - Set $Label to a label name, GUID, immutable ID, or identity to process
      only that label.
    - Leave $Label blank to inspect all labels and process those whose returned
      configuration indicates disposition review.

    Counting behavior:

    - ReviewAction comparisons are case-insensitive.
    - Leading and trailing whitespace is removed before comparison.
    - No DeletedDate cutoff is applied.
    - Only items returned from the disposed view are examined.

    CSV behavior:

    - One row is written for each processed label.
    - Successful and failed labels are both included.
    - No cumulative-total row is generated.
    - The completed filename contains the script completion timestamp.
#>

[CmdletBinding()]
param()

#region Script configuration

# Optional retention label name, GUID, immutable ID, or identity.
#
# Use a label name for single-label testing:
$Label = "Internal"

# Leave blank to process all labels detected as having disposition review:
# $Label = ""

# ReviewAction value to count.
$TargetReviewAction = "RecordReviewed"

# Number of seconds to pause between Get-ReviewItems page requests.
$PageDelaySeconds = 1

# Additional attempts per page for the transient server-side retry error.
$ReviewItemRetryCount = 10

# Seconds to wait before each retry, separate from normal paging delays.
$ReviewItemRetryDelaySeconds = 30

# Number of seconds to pause between detailed Get-ComplianceTag requests when
# inspecting all labels.
$LabelDelaySeconds = 0

# Full directory where the completed CSV report will be saved.
#
# The script creates this directory automatically if it does not exist.
#
# This can also be a network path, for example:
# $OutputDirectory = "C:\Temp\PurviewReports"
$OutputDirectory = "c:\temp"

# Base filename for the CSV report.
#
# The completion timestamp and .csv extension are added automatically.
# Example:
# Disposition_ApprovedNotDeletedSummary-2026-09-29_08-15-42.csv
$OutputFileBaseName = "Disposition_ApprovedNotDeletedSummary"

#endregion Script configuration

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Helper functions

function Test-ModuleAvailable
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    $module = Get-Module `
        -ListAvailable `
        -Name $Name

    if ($null -eq $module)
    {
        return $false
    }

    return $true
}

function Test-StringHasValue
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )

    if ($null -eq $Value)
    {
        return $false
    }

    $stringValue = [string]$Value

    if ($stringValue.Trim().Length -eq 0)
    {
        return $false
    }

    return $true
}

function Get-PropertyValue
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$PropertyName
    )

    foreach ($name in $PropertyName)
    {
        $property = $InputObject.PSObject.Properties[$name]

        if ($null -ne $property)
        {
            return $property.Value
        }
    }

    return $null
}

function Test-HasDispositionReview
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$ComplianceTag
    )

    # Check properties commonly used for single-stage disposition review.
    $reviewers = Get-PropertyValue `
        -InputObject $ComplianceTag `
        -PropertyName @(
            "ReviewerEmail",
            "ReviewerEmails",
            "DispositionReviewers"
        )

    if ($null -ne $reviewers)
    {
        if ($reviewers -is [string])
        {
            if (Test-StringHasValue -Value $reviewers)
            {
                return $true
            }
        }
        elseif ($reviewers -is [System.Collections.IEnumerable])
        {
            foreach ($reviewer in $reviewers)
            {
                if (Test-StringHasValue -Value $reviewer)
                {
                    return $true
                }
            }
        }
        elseif (Test-StringHasValue -Value $reviewers)
        {
            return $true
        }
    }

    # Check properties commonly used for multi-stage disposition review.
    $multiStageReview = Get-PropertyValue `
        -InputObject $ComplianceTag `
        -PropertyName @(
            "MultiStageReviewProperty",
            "MultiStageReviewProperties"
        )

    if ($null -ne $multiStageReview)
    {
        if ($multiStageReview -is [string])
        {
            if (Test-StringHasValue -Value $multiStageReview)
            {
                return $true
            }
        }
        elseif ($multiStageReview -is [System.Collections.IEnumerable])
        {
            foreach ($reviewStage in $multiStageReview)
            {
                if (Test-StringHasValue -Value $reviewStage)
                {
                    return $true
                }
            }
        }
        elseif (Test-StringHasValue -Value $multiStageReview)
        {
            return $true
        }
    }

    return $false
}

function Get-LabelLookupIdentity
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$ComplianceTag
    )

    $lookupIdentity = Get-PropertyValue `
        -InputObject $ComplianceTag `
        -PropertyName @(
            "Guid",
            "ImmutableId",
            "Identity",
            "Name"
        )

    if (-not (Test-StringHasValue -Value $lookupIdentity))
    {
        throw "Unable to determine a lookup identity for a retention label."
    }

    return ([string]$lookupIdentity).Trim()
}

function Get-DetailedComplianceTag
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$ComplianceTag
    )

    $lookupIdentity = Get-LabelLookupIdentity `
        -ComplianceTag $ComplianceTag

    $detailedLabelResult = Get-ComplianceTag `
        -Identity $lookupIdentity `
        -ErrorAction Stop

    if ($null -eq $detailedLabelResult)
    {
        throw (
            "Get-ComplianceTag returned no detailed information for label " +
            "'$($ComplianceTag.Name)'."
        )
    }

    $detailedLabelCollection = @($detailedLabelResult)

    if ($detailedLabelCollection.Count -eq 1)
    {
        return $detailedLabelCollection[0]
    }

    $originalName = Get-PropertyValue `
        -InputObject $ComplianceTag `
        -PropertyName @("Name")

    $originalGuid = Get-PropertyValue `
        -InputObject $ComplianceTag `
        -PropertyName @("Guid")

    $originalImmutableId = Get-PropertyValue `
        -InputObject $ComplianceTag `
        -PropertyName @("ImmutableId")

    $exactMatches = @(
        $detailedLabelCollection | Where-Object {
            $nameMatches = (
                (Test-StringHasValue -Value $originalName) -and
                [string]$_.Name -eq [string]$originalName
            )

            $guidMatches = (
                (Test-StringHasValue -Value $originalGuid) -and
                [string]$_.Guid -eq [string]$originalGuid
            )

            $immutableIdMatches = (
                (Test-StringHasValue -Value $originalImmutableId) -and
                [string]$_.ImmutableId -eq [string]$originalImmutableId
            )

            $nameMatches -or $guidMatches -or $immutableIdMatches
        }
    )

    if ($exactMatches.Count -eq 1)
    {
        return $exactMatches[0]
    }

    throw (
        "Unable to uniquely retrieve detailed configuration for retention " +
        "label '$originalName'."
    )
}

function Get-LabelImmutableId
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$ComplianceTag
    )

    # Get-ReviewItems specifically requires the label's ImmutableId.
    $immutableId = Get-PropertyValue `
        -InputObject $ComplianceTag `
        -PropertyName @("ImmutableId")

    if (-not (Test-StringHasValue -Value $immutableId))
    {
        throw (
            "The retention label '$($ComplianceTag.Name)' did not return an " +
            "ImmutableId. Get-ReviewItems requires the exact ImmutableId."
        )
    }

    return ([string]$immutableId).Trim()
}

function Get-ReviewActionCount
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetLabelId,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ReviewAction,

        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 300)]
        [int]$DelaySeconds = 1,

        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 100)]
        [int]$RetryCount = 10,

        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 3600)]
        [int]$RetryDelaySeconds = 30
    )

    [long]$matchingCount = 0
    [long]$totalDisposedCount = 0
    [long]$missingReviewActionCount = 0
    [int]$pageNumber = 0

    $pagingCookie = $null
    $headers = $null

    $normalizedTargetReviewAction = $ReviewAction.Trim()

    do
    {
        $pageNumber++

        # Do not add ErrorAction to this hashtable.
        #
        # Testing confirmed that Get-ReviewItems fails in this environment
        # when -ErrorAction Stop is supplied directly to the command.
        $reviewItemParameters = @{
            TargetLabelId  = $TargetLabelId
            Disposed       = $true
            IncludeHeaders = $true
        }

        if (Test-StringHasValue -Value $pagingCookie)
        {
            $reviewItemParameters["PagingCookie"] = $pagingCookie
        }

        Write-Verbose (
            "Requesting page $pageNumber for label ID '$TargetLabelId'."
        )

        $retriesUsed = 0

        while ($true)
        {
            $page = $null

            try
            {
                # Deliberately do not add -ErrorAction Stop.
                $page = Get-ReviewItems @reviewItemParameters
                break
            }
            catch
            {
                $isTransientReviewError = (
                    $_.Exception.Message -like "*A server side error has occurred*" -and
                    $_.Exception.Message -like "*Please try again after some time*"
                )

                if (-not $isTransientReviewError)
                {
                    throw
                }

                if ($retriesUsed -ge $RetryCount)
                {
                    Write-Warning (
                        "Get-ReviewItems exhausted $RetryCount retries for " +
                        "label ID '$TargetLabelId', page $pageNumber. " +
                        "The label will be recorded as Failed."
                    )
                    throw
                }

                $retriesUsed++

                Write-Warning (
                    "Transient Get-ReviewItems error for label ID " +
                    "'$TargetLabelId', page $pageNumber. Retry " +
                    "$retriesUsed of $RetryCount in $RetryDelaySeconds " +
                    "seconds. $($_.Exception.Message)"
                )

                if ($RetryDelaySeconds -gt 0)
                {
                    Start-Sleep -Seconds $RetryDelaySeconds
                }
            }
        }

        if ($null -eq $page)
        {
            throw (
                "Get-ReviewItems returned no response for label ID " +
                "'$TargetLabelId' on page $pageNumber."
            )
        }

        if ($null -eq $headers)
        {
            $headers = @($page.Headers)
        }

        $exportItems = @($page.ExportItems)

        # A valid label can have no disposed items.
        if ($exportItems.Count -eq 0)
        {
            Write-Verbose (
                "No disposed items were returned for label ID " +
                "'$TargetLabelId' on page $pageNumber."
            )

            $pagingCookie = [string]$page.PaginationCookie

            $exportItems = $null
            $page = $null

            if (-not (Test-StringHasValue -Value $pagingCookie))
            {
                break
            }

            if ($DelaySeconds -gt 0)
            {
                Start-Sleep -Seconds $DelaySeconds
            }

            continue
        }

        if ($null -eq $headers -or $headers.Count -eq 0)
        {
            throw (
                "Get-ReviewItems returned disposed items but did not return " +
                "CSV headers for label ID '$TargetLabelId'."
            )
        }

        # Confirm that ReviewAction exists in the returned headers.
        $reviewActionHeaderExists = $false

        foreach ($header in $headers)
        {
            if ([string]$header -eq "ReviewAction")
            {
                $reviewActionHeaderExists = $true
                break
            }
        }

        if (-not $reviewActionHeaderExists)
        {
            throw (
                "Get-ReviewItems did not return a ReviewAction column for " +
                "label ID '$TargetLabelId'."
            )
        }

        foreach ($csvRow in $exportItems)
        {
            if (-not (Test-StringHasValue -Value $csvRow))
            {
                continue
            }

            # Process one CSV row at a time so item-level data is not retained.
            $item = $csvRow |
                ConvertFrom-Csv -Header $headers |
                Select-Object -First 1

            if ($null -eq $item)
            {
                continue
            }

            $totalDisposedCount++

            $itemReviewAction = Get-PropertyValue `
                -InputObject $item `
                -PropertyName @("ReviewAction")

            if (-not (Test-StringHasValue -Value $itemReviewAction))
            {
                $missingReviewActionCount++

                $item = $null
                continue
            }

            $normalizedItemReviewAction = (
                [string]$itemReviewAction
            ).Trim()

            # PowerShell string equality is case-insensitive by default.
            if ($normalizedItemReviewAction -eq $normalizedTargetReviewAction)
            {
                $matchingCount++
            }

            $item = $null
        }

        $pagingCookie = [string]$page.PaginationCookie

        Write-Verbose (
            "Label ID ${TargetLabelId}: processed page $pageNumber; " +
            "disposed items examined: $totalDisposedCount; " +
            "$normalizedTargetReviewAction items: $matchingCount"
        )

        $exportItems = $null
        $page = $null

        if (
            (Test-StringHasValue -Value $pagingCookie) -and
            $DelaySeconds -gt 0
        )
        {
            Start-Sleep -Seconds $DelaySeconds
        }
    }
    while (Test-StringHasValue -Value $pagingCookie)

    return [pscustomobject]@{
        RecordReviewedCount      = $matchingCount
        TotalDisposedExamined    = $totalDisposedCount
        MissingReviewActionCount = $missingReviewActionCount
        PagesProcessed           = $pageNumber
    }
}

function Write-SummaryToCsv
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$Summary,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    # Select the CSV columns in the required order.
    $csvRow = $Summary |
        Select-Object `
            LabelName,
            LabelId,
            TargetReviewAction,
            RecordReviewedCount,
            TotalDisposedExamined,
            MissingReviewAction,
            PagesProcessed,
            Status

    if (Test-Path -LiteralPath $Path)
    {
        $csvRow |
            Export-Csv `
                -LiteralPath $Path `
                -NoTypeInformation `
                -Encoding UTF8 `
                -Append `
                -ErrorAction Stop
    }
    else
    {
        $csvRow |
            Export-Csv `
                -LiteralPath $Path `
                -NoTypeInformation `
                -Encoding UTF8 `
                -ErrorAction Stop
    }
}

#endregion Helper functions

#region CSV report initialization

if (-not (Test-StringHasValue -Value $OutputDirectory))
{
    throw "OutputDirectory cannot be empty."
}

$OutputDirectory = ([string]$OutputDirectory).Trim()

if (-not (Test-StringHasValue -Value $OutputFileBaseName))
{
    throw "OutputFileBaseName cannot be empty."
}

$OutputFileBaseName = $OutputFileBaseName.Trim()

if (-not (Test-StringHasValue -Value $TargetReviewAction))
{
    throw "TargetReviewAction cannot be empty."
}

$TargetReviewAction = $TargetReviewAction.Trim()

# Create the configured export directory if it does not already exist.
if (-not (Test-Path -LiteralPath $OutputDirectory))
{
    Write-Host ""
    Write-Host (
        "Creating CSV output directory: $OutputDirectory"
    ) -ForegroundColor Cyan

    New-Item `
        -Path $OutputDirectory `
        -ItemType Directory `
        -Force `
        -ErrorAction Stop |
        Out-Null
}

# Verify the configured output location is a directory.
$outputDirectoryItem = Get-Item `
    -LiteralPath $OutputDirectory `
    -ErrorAction Stop

if (-not $outputDirectoryItem.PSIsContainer)
{
    throw (
        "The configured OutputDirectory '$OutputDirectory' is not a directory."
    )
}

# Use an in-progress filename while the script is running.
$temporaryCsvFileName = "$OutputFileBaseName-InProgress.csv"

$temporaryCsvPath = Join-Path `
    -Path $OutputDirectory `
    -ChildPath $temporaryCsvFileName

# Remove a stale temporary report from a previous interrupted execution.
if (Test-Path -LiteralPath $temporaryCsvPath)
{
    Remove-Item `
        -LiteralPath $temporaryCsvPath `
        -Force `
        -ErrorAction Stop
}

$completedCsvPath = $null

Write-Host ""
Write-Host "CSV output directory:" `
    -ForegroundColor Cyan

Write-Host $OutputDirectory `
    -ForegroundColor Green

Write-Host (
    "Target ReviewAction: $TargetReviewAction"
) -ForegroundColor Green

#endregion CSV report initialization

#region Prerequisite validation

if (-not (Test-ModuleAvailable -Name "ExchangeOnlineManagement"))
{
    throw @"
The ExchangeOnlineManagement module is not installed.

Install it with:

    Install-Module ExchangeOnlineManagement -Scope CurrentUser
"@
}

Import-Module ExchangeOnlineManagement `
    -ErrorAction Stop

$loadedExchangeModule = Get-Module `
    -Name ExchangeOnlineManagement

if ($null -eq $loadedExchangeModule)
{
    throw "ExchangeOnlineManagement was not loaded successfully."
}

Write-Host ""
Write-Host "Loaded module:" -ForegroundColor Cyan

$loadedExchangeModule |
    Select-Object Name, Version, Path |
    Format-Table -AutoSize

#endregion Prerequisite validation

#region Validate script configuration

if ($PageDelaySeconds -lt 0 -or $PageDelaySeconds -gt 300)
{
    throw "PageDelaySeconds must be between 0 and 300."
}

if ($LabelDelaySeconds -lt 0 -or $LabelDelaySeconds -gt 300)
{
    throw "LabelDelaySeconds must be between 0 and 300."
}

if ($ReviewItemRetryCount -lt 0 -or $ReviewItemRetryCount -gt 100)
{
    throw "ReviewItemRetryCount must be between 0 and 100."
}

if (
    $ReviewItemRetryDelaySeconds -lt 0 -or
    $ReviewItemRetryDelaySeconds -gt 3600
)
{
    throw "ReviewItemRetryDelaySeconds must be between 0 and 3600."
}

if (Test-StringHasValue -Value $Label)
{
    $Label = $Label.Trim()

    Write-Host (
        "Label mode: Single label, '$Label'"
    ) -ForegroundColor Green
}
else
{
    $Label = ""

    Write-Host (
        "Label mode: All labels configured for disposition review"
    ) -ForegroundColor Green
}

#endregion Validate script configuration

$exchangeOnlineConnected = $false
$complianceConnected = $false

try
{
    #region Connect to Exchange Online

    Write-Host ""
    Write-Host "Connecting interactively to Exchange Online PowerShell..." `
        -ForegroundColor Cyan

    Connect-ExchangeOnline `
        -ShowBanner:$false `
        -ShowProgress $false `
        -ErrorAction Stop

    $exchangeOnlineConnected = $true

    Write-Host "Exchange Online PowerShell connection established." `
        -ForegroundColor Green

    #endregion Connect to Exchange Online

    #region Validate Exchange Online role assignment

    $requiredExchangeRole = "Mailbox Import Export"

    Write-Host ""
    Write-Host (
        "Checking the signed-in account's effective Exchange Online role " +
        "assignments..."
    ) -ForegroundColor Cyan

    $exchangeConnection = @(
        Get-ConnectionInformation -ErrorAction Stop |
            Where-Object {
                $_.State -eq "Connected" -and
                (Test-StringHasValue -Value $_.UserPrincipalName)
            }
    ) |
        Select-Object -Last 1

    if ($null -eq $exchangeConnection)
    {
        throw (
            "The signed-in Exchange Online account could not be determined " +
            "from Get-ConnectionInformation."
        )
    }

    $signedInUserPrincipalName = [string]$exchangeConnection.UserPrincipalName
    $exchangeIdentityValues = @($signedInUserPrincipalName)

    try
    {
        $signedInExchangeUser = Get-User `
            -Identity $signedInUserPrincipalName `
            -ErrorAction Stop

        foreach ($identityProperty in @("Name", "DisplayName", "Alias"))
        {
            $exchangeIdentityValues += [string](Get-PropertyValue `
                -InputObject $signedInExchangeUser `
                -PropertyName $identityProperty)
        }
    }
    catch
    {
        Write-Warning (
            "Could not retrieve identity details for " +
            "'$signedInUserPrincipalName'. The role check will use the " +
            "signed-in user principal name only. $($_.Exception.Message)"
        )
    }

    $exchangeIdentityValues = @(
        $exchangeIdentityValues |
            Where-Object { Test-StringHasValue -Value $_ } |
            Select-Object -Unique
    )

    try
    {
        $effectiveRoleAssignments = @(
            Get-ManagementRoleAssignment `
                -Role $requiredExchangeRole `
                -GetEffectiveUsers `
                -ErrorAction Stop |
                Where-Object {
                    $effectiveUserName = [string](Get-PropertyValue `
                        -InputObject $_ `
                        -PropertyName "EffectiveUserName")

                    $exchangeIdentityValues -icontains $effectiveUserName
                }
        )

        if ($effectiveRoleAssignments.Count -eq 0)
        {
            Write-Warning (
                "No matching effective '$requiredExchangeRole' Exchange " +
                "Online role assignments were returned for " +
                "'$signedInUserPrincipalName'. This diagnostic does not " +
                "establish access in Security and Compliance PowerShell. " +
                "The script will connect with Connect-IPPSSession and then " +
                "check whether Get-ReviewItems is available."
            )
        }
        else
        {
            Write-Host (
                "Matching effective '$requiredExchangeRole' assignments for " +
                "'$signedInUserPrincipalName':"
            ) -ForegroundColor Green

            $effectiveRoleAssignments |
                Select-Object `
                    Role,
                    RoleAssigneeName,
                    EffectiveUserName,
                    AssignmentMethod |
                Format-Table -AutoSize |
                Out-Host
        }
    }
    catch
    {
        Write-Warning (
            "The effective '$requiredExchangeRole' role assignment could not " +
            "be enumerated for '$signedInUserPrincipalName'. This query can " +
            "require role-assignment visibility that the account does not " +
            "have. The script will continue and verify access by checking " +
            "whether Get-ReviewItems is available. $($_.Exception.Message)"
        )
    }

    #endregion Validate Exchange Online role assignment

    #region Connect to Security and Compliance PowerShell

    Write-Host ""
    Write-Host (
        "Connecting interactively to Microsoft Security and Compliance " +
        "PowerShell..."
    ) -ForegroundColor Cyan

    Connect-IPPSSession `
        -ShowBanner:$false `
        -ErrorAction Stop

    $complianceConnected = $true

    Write-Host (
        "Security and Compliance PowerShell connection established."
    ) -ForegroundColor Green

    #endregion Connect to Security and Compliance PowerShell

    #region Validate required commands

    $requiredCommands = @(
        "Get-ComplianceTag",
        "Get-ReviewItems"
    )

    $missingCommands = @()

    foreach ($requiredCommand in $requiredCommands)
    {
        $command = Get-Command `
            -Name $requiredCommand `
            -ErrorAction SilentlyContinue

        if ($null -eq $command)
        {
            $missingCommands += $requiredCommand
            Write-Warning "Command unavailable: $requiredCommand"
        }
        else
        {
            Write-Host "Command available: $requiredCommand" `
                -ForegroundColor Green
        }
    }

    if ($missingCommands.Count -gt 0)
    {
        $missingCommandList = $missingCommands -join ", "

        throw (
            "The following command(s) are not available in the current " +
            "PowerShell sessions: $missingCommandList. Verify that the " +
            "signed-in account has the required Exchange Online and " +
            "Microsoft Purview RBAC role assignments."
        )
    }

    Write-Host "All required commands are available." `
        -ForegroundColor Green

    #endregion Validate required commands

    #region Retrieve retention labels

    Write-Host ""
    Write-Host "Retrieving Microsoft Purview retention labels..." `
        -ForegroundColor Cyan

    $allLabels = @(Get-ComplianceTag -ErrorAction Stop)

    if ($allLabels.Count -eq 0)
    {
        throw "No retention labels were returned by Get-ComplianceTag."
    }

    Write-Host (
        "Retrieved $($allLabels.Count) total retention label(s)."
    ) -ForegroundColor Green

    #endregion Retrieve retention labels

    #region Select labels to process

    $labelsToProcess = @()

    if (Test-StringHasValue -Value $Label)
    {
        $matchingLabels = @(
            $allLabels | Where-Object {
                [string]$_.Name -eq $Label -or
                [string]$_.Guid -eq $Label -or
                [string]$_.ImmutableId -eq $Label -or
                [string]$_.Identity -eq $Label
            }
        )

        if ($matchingLabels.Count -eq 0)
        {
            throw "The retention label '$Label' was not found."
        }

        if ($matchingLabels.Count -gt 1)
        {
            throw (
                "More than one retention label matched '$Label'. Use the " +
                "label GUID or immutable ID instead."
            )
        }

        Write-Host (
            "Retrieving detailed configuration for label " +
            "'$($matchingLabels[0].Name)'..."
        ) -ForegroundColor Cyan

        $detailedLabel = Get-DetailedComplianceTag `
            -ComplianceTag $matchingLabels[0]

        # Do not reject a specifically selected label based on reviewer
        # metadata. Get-ReviewItems is the authoritative operation.
        $labelsToProcess = @($detailedLabel)

        Write-Host (
            "Selected retention label '$($detailedLabel.Name)' for processing."
        ) -ForegroundColor Green
    }
    else
    {
        [int]$labelNumber = 0
        [int]$inspectionFailureCount = 0

        foreach ($summaryLabel in $allLabels)
        {
            $labelNumber++

            $summaryLabelName = Get-PropertyValue `
                -InputObject $summaryLabel `
                -PropertyName @("Name")

            Write-Host (
                "Inspecting label $labelNumber of $($allLabels.Count): " +
                "$summaryLabelName"
            ) -ForegroundColor DarkCyan

            try
            {
                $detailedLabel = Get-DetailedComplianceTag `
                    -ComplianceTag $summaryLabel

                if (
                    $null -ne $detailedLabel -and
                    (
                        Test-HasDispositionReview `
                            -ComplianceTag $detailedLabel
                    )
                )
                {
                    $labelsToProcess += $detailedLabel

                    Write-Host (
                        "Disposition-review label found: " +
                        "$($detailedLabel.Name)"
                    ) -ForegroundColor Green
                }
                else
                {
                    Write-Verbose (
                        "Label '$summaryLabelName' was not identified as a " +
                        "disposition-review label."
                    )
                }

                $detailedLabel = $null
            }
            catch
            {
                $inspectionFailureCount++

                Write-Warning (
                    "Unable to inspect retention label '$summaryLabelName'. " +
                    "$($_.Exception.Message)"
                )
            }

            if (
                $LabelDelaySeconds -gt 0 -and
                $labelNumber -lt $allLabels.Count
            )
            {
                Start-Sleep -Seconds $LabelDelaySeconds
            }
        }

        if ($inspectionFailureCount -gt 0)
        {
            Write-Warning (
                "$inspectionFailureCount retention label(s) could not be " +
                "inspected. Those labels will not be processed."
            )
        }
    }

    $allLabels = $null

    if ($labelsToProcess.Count -eq 0)
    {
        Write-Warning (
            "No retention labels configured for disposition review were found."
        )

        return
    }

    Write-Host ""
    Write-Host (
        "Found $($labelsToProcess.Count) retention label(s) to process."
    ) -ForegroundColor Green

    #endregion Select labels to process

    #region Process retention labels

    Write-Host ""

    foreach ($retentionLabel in $labelsToProcess)
    {
        $labelName = [string]$retentionLabel.Name
        $labelId = $null

        try
        {
            $labelId = Get-LabelImmutableId `
                -ComplianceTag $retentionLabel

            Write-Host (
                "Processing label: $labelName"
            ) -ForegroundColor Yellow

            Write-Host (
                "TargetLabelId: [$labelId]"
            ) -ForegroundColor DarkCyan

            $countResult = Get-ReviewActionCount `
                -TargetLabelId $labelId `
                -ReviewAction $TargetReviewAction `
                -DelaySeconds $PageDelaySeconds `
                -RetryCount $ReviewItemRetryCount `
                -RetryDelaySeconds $ReviewItemRetryDelaySeconds

            $summary = [pscustomobject]@{
                LabelName             = $labelName
                LabelId               = $labelId
                TargetReviewAction    = $TargetReviewAction
                RecordReviewedCount   = $countResult.RecordReviewedCount
                TotalDisposedExamined = $countResult.TotalDisposedExamined
                MissingReviewAction   = $countResult.MissingReviewActionCount
                PagesProcessed        = $countResult.PagesProcessed
                Status                = "Completed"
                Error                 = $null
            }

            Write-SummaryToCsv `
                -Summary $summary `
                -Path $temporaryCsvPath

            Write-Host (
                "Completed label '$labelName'. " +
                "$TargetReviewAction count: " +
                "$($countResult.RecordReviewedCount)"
            ) -ForegroundColor Green

            # Continue displaying the same summary in the console.
            $summary

            $countResult = $null
            $summary = $null
        }
        catch
        {
            $errorMessage = $_.Exception.Message

            Write-Warning (
                "Unable to process label '$labelName'. $errorMessage"
            )

            $failedSummary = [pscustomobject]@{
                LabelName             = $labelName
                LabelId               = $labelId
                TargetReviewAction    = $TargetReviewAction
                RecordReviewedCount   = $null
                TotalDisposedExamined = $null
                MissingReviewAction   = $null
                PagesProcessed        = $null
                Status                = "Failed"
                Error                 = $errorMessage
            }

            Write-SummaryToCsv `
                -Summary $failedSummary `
                -Path $temporaryCsvPath

            # Continue displaying the failed summary in the console.
            $failedSummary

            $failedSummary = $null
        }
    }

    #endregion Process retention labels
}
finally
{
    #region Disconnect

    if ($complianceConnected -or $exchangeOnlineConnected)
    {
        try
        {
            Write-Host ""
            Write-Host "Disconnecting Exchange Online sessions..." `
                -ForegroundColor Cyan

            Disconnect-ExchangeOnline `
                -Confirm:$false `
                -ErrorAction SilentlyContinue

            Write-Host "Exchange Online sessions disconnected." `
                -ForegroundColor Green
        }
        catch
        {
            Write-Verbose (
                "Disconnect returned an error: $($_.Exception.Message)"
            )
        }
    }

    #endregion Disconnect

    #region Finalize CSV report

    if (Test-Path -LiteralPath $temporaryCsvPath)
    {
        try
        {
            $completionTimestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

            $completedCsvFileName = (
                "$OutputFileBaseName-$completionTimestamp.csv"
            )

            $completedCsvPath = Join-Path `
                -Path $OutputDirectory `
                -ChildPath $completedCsvFileName

            Move-Item `
                -LiteralPath $temporaryCsvPath `
                -Destination $completedCsvPath `
                -Force `
                -ErrorAction Stop

            Write-Host ""
            Write-Host "CSV report created successfully:" `
                -ForegroundColor Green

            Write-Host $completedCsvPath `
                -ForegroundColor Cyan
        }
        catch
        {
            Write-Warning (
                "The CSV was created, but it could not be renamed with the " +
                "completion timestamp. Temporary report path: " +
                "$temporaryCsvPath. Error: $($_.Exception.Message)"
            )
        }
    }
    else
    {
        Write-Warning (
            "No CSV report was created because no label summary rows were " +
            "generated."
        )
    }

    #endregion Finalize CSV report
}
