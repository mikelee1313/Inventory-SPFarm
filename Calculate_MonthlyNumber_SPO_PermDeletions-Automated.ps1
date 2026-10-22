<#
.SYNOPSIS
    Counts Microsoft Purview disposition items permanently deleted on or after
    a configured date and exports one summary row per retention label to CSV.

.DESCRIPTION
    This script:

    1. Connects to Exchange Online PowerShell using a saved credential.
    2. Displays matching effective Exchange Online Mailbox Import Export
       role assignments as a diagnostic check.
    3. Connects to Microsoft Security and Compliance PowerShell using a
       separately loaded copy of the saved credential.
    4. Retrieves all Microsoft Purview retention labels.
    5. Processes either one configured label or every label that returns an
       ImmutableId.
    6. Uses Get-ReviewItems -Disposed $true to retrieve disposed items.
    7. Processes every page returned by Get-ReviewItems.
    8. Counts items whose DeletedDate is on or after the configured SinceDate.
    9. Does not retain all item-level results in memory.
    10. Writes one CSV row per processed, skipped, or failed label.
    11. Includes readable and JSON-formatted error information.
    12. Renames the CSV at completion using the completion date and time.

.NOTES
    Authors: David O'Saben | Mike Lee
    Updated: 10/22/26
    
    Important:

    Get-ReviewItems must not be called with -ErrorAction Stop in this
    environment. Testing confirmed that the command succeeds with its default
    error behavior but returns a generic server-side error when
    -ErrorAction Stop is supplied.

    The global $ErrorActionPreference can remain set to Stop.

    Unattended authentication:

    - Requires an approved account whose existing sign-in policies allow
      password-only authentication. This does not bypass MFA or Conditional
      Access and does not use an app registration.
    - On the scheduled-task computer, run the following once as the Windows
      account that will execute the task:

          New-Item -Path "C:\ProgramData\PurviewReporting" `
              -ItemType Directory -Force | Out-Null
          Get-Credential |
              Export-Clixml -LiteralPath `
                  "C:\ProgramData\PurviewReporting\Credential.xml"

    - Windows encrypts the exported password using DPAPI. Only the same
      Windows account on the same computer can decrypt it. Restrict NTFS
      access to the credential directory to the task account and authorized
      administrators. Re-export the credential after password changes.
    - Configure the task to run whether the user is logged on or not, using
      -NoProfile -NonInteractive -File "<full script path>". Do not allow
      overlapping runs because the in-progress CSV filename is shared.
    - Credentials are loaded separately for each connection because
      Connect-IPPSSession clears the supplied credential's password.
    - Missing or invalid credentials stop execution; no interactive
      authentication fallback is attempted.

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

    - Set $Label to a label name, GUID, ImmutableId, or identity to process
      only that label.
    - Leave $Label blank to process every retention label that returns an
      ImmutableId.
    - ReviewerEmail and MultiStageReviewProperty are not used to prefilter
      labels because those properties are not consistently populated.
    - Get-ReviewItems is treated as the authoritative operation.

    Status behavior:

    - Completed: Get-ReviewItems successfully processed the label.
    - Skipped: Get-ReviewItems reported that the disposition-review folder was
      unavailable for the label.
    - Failed: An unexpected service or processing failure occurred.

    CSV behavior:

    - One row is written for every label attempted.
    - Rows are written as each label finishes.
    - The Error column contains a readable error description.
    - The ErrorJson column contains structured PowerShell error details.
    - The final filename contains the script completion timestamp.
#>

[CmdletBinding()]
param()

#region Script configuration

# Windows-encrypted credential exported by the scheduled-task account.
$CredentialPath = "C:\ProgramData\PurviewReporting\Credential.xml"

# Count only items whose DeletedDate is on or after this date.
$SinceDate = Get-Date "2026-09-01"

# Optional retention label name, GUID, ImmutableId, or identity.
#
# Example for a single-label run:
# $Label = "ADM100"
#
# Leave blank to process all labels that return an ImmutableId:
$Label = "Internal"

# Number of seconds to pause between Get-ReviewItems page requests.
$PageDelaySeconds = 1

# Additional attempts per page for the transient server-side retry error.
$ReviewItemRetryCount = 10

# Seconds to wait before each retry, separate from normal paging delays.
$ReviewItemRetryDelaySeconds = 30

# Full directory where the completed CSV report will be saved.
#
# The script creates this directory automatically if it does not exist.
$OutputDirectory = "c:\temp"

# Base filename for the CSV report.
#
# A completion timestamp and the .csv extension are added automatically.
# Example:
# Disposition_PermDeletionsSummary-2026-10-01_14-25-30.csv
$OutputFileBaseName = "Disposition_PermDeletionsSummary"

#endregion Script configuration

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Helper functions

function Get-AutomationCredential
{
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCredential])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf -ErrorAction Stop))
    {
        throw (
            "Credential file '$Path' was not found. Create it with " +
            "Get-Credential | Export-Clixml on this computer as the Windows " +
            "account that will run the scheduled task."
        )
    }

    try
    {
        $credential = Import-Clixml -LiteralPath $Path -ErrorAction Stop
    }
    catch
    {
        throw (
            "Could not load credential file '$Path'. Verify file access and " +
            "re-export it on this computer as the Windows account that will " +
            "run the scheduled task. $($_.Exception.Message)"
        )
    }

    if ($credential -isnot [System.Management.Automation.PSCredential])
    {
        throw "Credential file '$Path' does not contain a PSCredential."
    }

    if (
        [string]::IsNullOrWhiteSpace($credential.UserName) -or
        $credential.Password.Length -eq 0
    )
    {
        throw "Credential file '$Path' must contain a username and password."
    }

    return $credential
}

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

function ConvertTo-DeletedDate
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
        return $null
    }

    if ($Value -is [datetimeoffset])
    {
        return $Value
    }

    if ($Value -is [datetime])
    {
        return [datetimeoffset]$Value
    }

    $textValue = ([string]$Value).Trim()

    if ($textValue.Length -eq 0)
    {
        return $null
    }

    try
    {
        return [datetimeoffset]$textValue
    }
    catch
    {
        return $null
    }
}

function Get-LabelImmutableId
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [object]$ComplianceTag
    )

    $immutableId = Get-PropertyValue `
        -InputObject $ComplianceTag `
        -PropertyName @("ImmutableId")

    if (-not (Test-StringHasValue -Value $immutableId))
    {
        $labelName = Get-PropertyValue `
            -InputObject $ComplianceTag `
            -PropertyName @("Name")

        throw (
            "The retention label '$labelName' did not return an ImmutableId. " +
            "Get-ReviewItems requires the exact ImmutableId."
        )
    }

    return ([string]$immutableId).Trim()
}

function Get-DisposedItemCount
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetLabelId,

        [Parameter(Mandatory = $true)]
        [datetimeoffset]$CutoffDate,

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
    [long]$missingDeletedDateCount = 0
    [long]$unparseableDeletedDateCount = 0
    [int]$pageNumber = 0

    $pagingCookie = $null
    $headers = $null

    do
    {
        $pageNumber++

        # Do not add ErrorAction to this hashtable.
        #
        # Get-ReviewItems returns a generic service error in this environment
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

        # A queryable label can have no items in the disposed view.
        # This is treated as a successful zero-item result.
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

        foreach ($csvRow in $exportItems)
        {
            if (-not (Test-StringHasValue -Value $csvRow))
            {
                continue
            }

            # Convert and process one returned row at a time.
            $item = $csvRow |
                ConvertFrom-Csv -Header $headers |
                Select-Object -First 1

            if ($null -eq $item)
            {
                continue
            }

            $totalDisposedCount++

            $deletedDateText = Get-PropertyValue `
                -InputObject $item `
                -PropertyName @("DeletedDate")

            if (-not (Test-StringHasValue -Value $deletedDateText))
            {
                $missingDeletedDateCount++
                $item = $null
                continue
            }

            $deletedDate = ConvertTo-DeletedDate `
                -Value $deletedDateText

            if ($null -eq $deletedDate)
            {
                $unparseableDeletedDateCount++
                $item = $null
                continue
            }

            if ($deletedDate -ge $CutoffDate)
            {
                $matchingCount++
            }

            $item = $null
            $deletedDate = $null
        }

        $pagingCookie = [string]$page.PaginationCookie

        Write-Verbose (
            "Label ID ${TargetLabelId}: processed page $pageNumber; " +
            "disposed items examined: $totalDisposedCount; " +
            "matching items: $matchingCount"
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
        MatchingCount               = $matchingCount
        TotalDisposedExamined       = $totalDisposedCount
        MissingDeletedDateCount     = $missingDeletedDateCount
        UnparseableDeletedDateCount = $unparseableDeletedDateCount
        PagesProcessed              = $pageNumber
    }
}

function Convert-ErrorRecordToJson
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    # Capture inner-exception details without attempting to serialize the
    # exception object directly.
    $innerExceptionData = $null

    if ($null -ne $ErrorRecord.Exception.InnerException)
    {
        $innerExceptionData = [ordered]@{
            Type = (
                $ErrorRecord.Exception.InnerException.GetType().FullName
            )
            Message = (
                $ErrorRecord.Exception.InnerException.Message
            )
            StackTrace = (
                $ErrorRecord.Exception.InnerException.StackTrace
            )
        }
    }

    # Capture ErrorDetails safely.
    $errorDetailsMessage = $null
    $recommendedAction = $null

    if ($null -ne $ErrorRecord.ErrorDetails)
    {
        $errorDetailsMessage = $ErrorRecord.ErrorDetails.Message
        $recommendedAction = $ErrorRecord.ErrorDetails.RecommendedAction
    }

    # Capture InvocationInfo safely.
    $invocationName = $null
    $scriptName = $null
    $scriptLineNumber = $null
    $offsetInLine = $null
    $line = $null
    $positionMessage = $null

    if ($null -ne $ErrorRecord.InvocationInfo)
    {
        $invocationName = $ErrorRecord.InvocationInfo.InvocationName
        $scriptName = $ErrorRecord.InvocationInfo.ScriptName
        $scriptLineNumber = $ErrorRecord.InvocationInfo.ScriptLineNumber
        $offsetInLine = $ErrorRecord.InvocationInfo.OffsetInLine
        $line = $ErrorRecord.InvocationInfo.Line
        $positionMessage = $ErrorRecord.InvocationInfo.PositionMessage
    }

    # Convert TargetObject to text rather than serializing the raw object.
    $targetObjectText = $null

    if ($null -ne $ErrorRecord.TargetObject)
    {
        try
        {
            $targetObjectText = [string]$ErrorRecord.TargetObject
        }
        catch
        {
            $targetObjectText = (
                "<Unable to convert TargetObject to text>"
            )
        }
    }

    $errorData = [ordered]@{
        Timestamp = (
            Get-Date
        ).ToString("o")

        ExceptionType = (
            $ErrorRecord.Exception.GetType().FullName
        )

        ExceptionMessage = (
            $ErrorRecord.Exception.Message
        )

        ExceptionStackTrace = (
            $ErrorRecord.Exception.StackTrace
        )

        InnerException = $innerExceptionData

        FullyQualifiedErrorId = (
            $ErrorRecord.FullyQualifiedErrorId
        )

        Category = (
            [string]$ErrorRecord.CategoryInfo.Category
        )

        CategoryActivity = (
            $ErrorRecord.CategoryInfo.Activity
        )

        CategoryReason = (
            $ErrorRecord.CategoryInfo.Reason
        )

        CategoryTargetName = (
            $ErrorRecord.CategoryInfo.TargetName
        )

        CategoryTargetType = (
            $ErrorRecord.CategoryInfo.TargetType
        )

        ErrorDetailsMessage = $errorDetailsMessage
        RecommendedAction = $recommendedAction
        TargetObject = $targetObjectText

        ScriptStackTrace = (
            $ErrorRecord.ScriptStackTrace
        )

        InvocationName = $invocationName
        ScriptName = $scriptName
        ScriptLineNumber = $scriptLineNumber
        OffsetInLine = $offsetInLine
        Line = $line
        PositionMessage = $positionMessage
    }

    try
    {
        return $errorData |
            ConvertTo-Json `
                -Depth 8 `
                -Compress
    }
    catch
    {
        # Return simpler valid JSON if the detailed JSON conversion fails.
        $jsonConversionError = $_.Exception.Message

        $fallbackErrorData = [ordered]@{
            Timestamp = (
                Get-Date
            ).ToString("o")

            ExceptionType = (
                $ErrorRecord.Exception.GetType().FullName
            )

            ExceptionMessage = (
                $ErrorRecord.Exception.Message
            )

            FullyQualifiedErrorId = (
                $ErrorRecord.FullyQualifiedErrorId
            )

            JsonConversionError = $jsonConversionError
        }

        return $fallbackErrorData |
            ConvertTo-Json `
                -Depth 4 `
                -Compress
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
            SinceDate,
            PermanentlyDeletedSinceDate,
            TotalDisposedExamined,
            MissingDeletedDate,
            UnparseableDeletedDate,
            PagesProcessed,
            Status,
            Error,
            ErrorJson

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

# Validate credentials before creating or removing any report files.
$exchangeCredential = Get-AutomationCredential -Path $CredentialPath

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

$outputDirectoryItem = Get-Item `
    -LiteralPath $OutputDirectory `
    -ErrorAction Stop

if (-not $outputDirectoryItem.PSIsContainer)
{
    throw (
        "The configured OutputDirectory '$OutputDirectory' is not a directory."
    )
}

# Use an in-progress filename while processing.
$temporaryCsvFileName = "$OutputFileBaseName-InProgress.csv"

$temporaryCsvPath = Join-Path `
    -Path $OutputDirectory `
    -ChildPath $temporaryCsvFileName

# Remove a stale temporary report from an interrupted earlier run.
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
Write-Host "Loaded module:" `
    -ForegroundColor Cyan

$loadedExchangeModule |
    Select-Object Name, Version, Path |
    Format-Table -AutoSize

#endregion Prerequisite validation

#region Validate script configuration

if ($null -eq $SinceDate)
{
    throw "SinceDate cannot be empty."
}

if ($PageDelaySeconds -lt 0 -or $PageDelaySeconds -gt 300)
{
    throw "PageDelaySeconds must be between 0 and 300."
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

try
{
    $cutoffDate = [datetimeoffset]$SinceDate
}
catch
{
    throw "SinceDate '$SinceDate' could not be converted to a valid date."
}

Write-Host (
    "DeletedDate cutoff: $($cutoffDate.ToString('o'))"
) -ForegroundColor Green

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
        "Label mode: All retention labels with an ImmutableId"
    ) -ForegroundColor Green
}

#endregion Validate script configuration

$exchangeOnlineConnected = $false
$complianceConnected = $false

try
{
    #region Connect to Exchange Online

    Write-Host ""
    Write-Host "Connecting to Exchange Online using the saved credential..." `
        -ForegroundColor Cyan

    Connect-ExchangeOnline `
        -Credential $exchangeCredential `
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
        "Connecting to Microsoft Security and Compliance PowerShell " +
        "using the saved credential..."
    ) -ForegroundColor Cyan

    Connect-IPPSSession `
        -Credential (Get-AutomationCredential -Path $CredentialPath) `
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

    # Retrieve the tenant-wide label collection once.
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
    [int]$labelsWithoutImmutableId = 0

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
                "label GUID or ImmutableId instead."
            )
        }

        $selectedImmutableId = Get-PropertyValue `
            -InputObject $matchingLabels[0] `
            -PropertyName @("ImmutableId")

        if (-not (Test-StringHasValue -Value $selectedImmutableId))
        {
            throw (
                "The selected retention label '$Label' did not return an " +
                "ImmutableId and cannot be passed to Get-ReviewItems."
            )
        }

        $labelsToProcess = @($matchingLabels[0])

        Write-Host (
            "Selected retention label '$($matchingLabels[0].Name)' for " +
            "processing."
        ) -ForegroundColor Green
    }
    else
    {
        # Do not filter using ReviewerEmail or MultiStageReviewProperty.
        #
        # Every label with an ImmutableId is passed to Get-ReviewItems.
        foreach ($summaryLabel in $allLabels)
        {
            $summaryLabelName = Get-PropertyValue `
                -InputObject $summaryLabel `
                -PropertyName @("Name")

            $summaryLabelImmutableId = Get-PropertyValue `
                -InputObject $summaryLabel `
                -PropertyName @("ImmutableId")

            if (Test-StringHasValue -Value $summaryLabelImmutableId)
            {
                $labelsToProcess += $summaryLabel
            }
            else
            {
                $labelsWithoutImmutableId++

                Write-Warning (
                    "Skipping retention label '$summaryLabelName' because it " +
                    "did not return an ImmutableId."
                )
            }
        }
    }

    $allLabels = $null

    if ($labelsWithoutImmutableId -gt 0)
    {
        Write-Warning (
            "$labelsWithoutImmutableId retention label(s) were skipped " +
            "because they did not return an ImmutableId."
        )
    }

    if ($labelsToProcess.Count -eq 0)
    {
        Write-Warning "No retention labels are available to process."
        return
    }

    Write-Host ""
    Write-Host (
        "Found $($labelsToProcess.Count) retention label(s) to process."
    ) -ForegroundColor Green

    #endregion Select labels to process

    #region Process retention labels

    Write-Host ""

    [int]$currentLabelNumber = 0
    [int]$completedLabelCount = 0
    [int]$skippedLabelCount = 0
    [int]$failedLabelCount = 0

    foreach ($retentionLabel in $labelsToProcess)
    {
        $currentLabelNumber++

        $labelName = Get-PropertyValue `
            -InputObject $retentionLabel `
            -PropertyName @("Name")

        $labelId = $null

        try
        {
            $labelId = Get-LabelImmutableId `
                -ComplianceTag $retentionLabel

            Write-Host ""
            Write-Host (
                "Processing label $currentLabelNumber of " +
                "$($labelsToProcess.Count): $labelName"
            ) -ForegroundColor Yellow

            Write-Host (
                "TargetLabelId: [$labelId]"
            ) -ForegroundColor DarkCyan

            $countResult = Get-DisposedItemCount `
                -TargetLabelId $labelId `
                -CutoffDate $cutoffDate `
                -DelaySeconds $PageDelaySeconds `
                -RetryCount $ReviewItemRetryCount `
                -RetryDelaySeconds $ReviewItemRetryDelaySeconds

            $summary = [pscustomobject]@{
                LabelName                   = $labelName
                LabelId                     = $labelId
                SinceDate                   = $cutoffDate
                PermanentlyDeletedSinceDate = $countResult.MatchingCount
                TotalDisposedExamined       = $countResult.TotalDisposedExamined
                MissingDeletedDate          = $countResult.MissingDeletedDateCount
                UnparseableDeletedDate      = $countResult.UnparseableDeletedDateCount
                PagesProcessed              = $countResult.PagesProcessed
                Status                      = "Completed"
                Error                       = $null
                ErrorJson                   = $null
            }

            Write-SummaryToCsv `
                -Summary $summary `
                -Path $temporaryCsvPath

            $completedLabelCount++

            Write-Host (
                "Completed label '$labelName'. Permanently deleted since " +
                "cutoff: $($countResult.MatchingCount)"
            ) -ForegroundColor Green

            # Continue displaying the same summary in the console.
            $summary

            $countResult = $null
            $summary = $null
        }
        catch
        {
            $errorRecord = $_
            $errorMessage = $errorRecord.Exception.Message

            $errorJson = Convert-ErrorRecordToJson `
                -ErrorRecord $errorRecord

            # Labels without a disposition-review backing folder can return:
            #
            # Value cannot be null. (Parameter 'folderId')
            #
            # Treat that condition as Skipped rather than Failed.
            $isMissingDispositionFolder = (
                $errorMessage -like "*Value cannot be null*" -and
                $errorMessage -like "*folderId*"
            )

            if ($isMissingDispositionFolder)
            {
                $skippedLabelCount++

                $status = "Skipped"

                $csvErrorMessage = (
                    "Get-ReviewItems could not query this label because no " +
                    "disposition-review folder was available."
                )

                Write-Warning (
                    "Skipping label '$labelName'. Get-ReviewItems did not " +
                    "find a disposition-review folder for this label."
                )
            }
            else
            {
                $failedLabelCount++

                $status = "Failed"
                $csvErrorMessage = $errorMessage

                Write-Warning (
                    "Unable to process label '$labelName'. $errorMessage"
                )
            }

            $unsuccessfulSummary = [pscustomobject]@{
                LabelName                   = $labelName
                LabelId                     = $labelId
                SinceDate                   = $cutoffDate
                PermanentlyDeletedSinceDate = $null
                TotalDisposedExamined       = $null
                MissingDeletedDate          = $null
                UnparseableDeletedDate      = $null
                PagesProcessed              = $null
                Status                      = $status
                Error                       = $csvErrorMessage
                ErrorJson                   = $errorJson
            }

            Write-SummaryToCsv `
                -Summary $unsuccessfulSummary `
                -Path $temporaryCsvPath

            # Continue displaying the same summary in the console.
            $unsuccessfulSummary

            $unsuccessfulSummary = $null
            $errorRecord = $null
            $errorJson = $null
        }
    }

    Write-Host ""
    Write-Host "Label processing complete." `
        -ForegroundColor Cyan

    Write-Host (
        "Completed labels: $completedLabelCount"
    ) -ForegroundColor Green

    if ($skippedLabelCount -gt 0)
    {
        Write-Host (
            "Skipped labels: $skippedLabelCount"
        ) -ForegroundColor Yellow
    }
    else
    {
        Write-Host "Skipped labels: 0" `
            -ForegroundColor Green
    }

    if ($failedLabelCount -gt 0)
    {
        Write-Warning "Failed labels: $failedLabelCount"
    }
    else
    {
        Write-Host "Failed labels: 0" `
            -ForegroundColor Green
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
