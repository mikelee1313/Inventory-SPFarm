<#
.SYNOPSIS
Enumerates all licensed users in the tenant and reports their OneDrive for Business URLs.

.DESCRIPTION
This script authenticates with Microsoft Graph API, retrieves all users that have at least
one assigned license, queries the OneDrive drive resource for each user, and exports the
results (URL, storage quota, drive ID, etc.) to a CSV file.

.PARAMETER None
This script does not accept parameters through the command line. Configuration is done through
variables at the beginning of the script.

.NOTES
File Name       : get-onedriveinfo - graph.ps1
Author          : Mike Lee
Date Created    : 4/28/25
Prerequisites   :
- PowerShell 5.1 or higher
- Appropriate permissions in Azure AD

API Permissions Required:
- User.Read.All          (enumerate users and license assignments)
- Files.Read.All         (read drive metadata for every user)

.EXAMPLE
PS> .\get-onedriveinfo - graph.ps1
Queries all licensed users and exports their OneDrive URLs to a CSV in %TEMP%.

.OUTPUTS
CSV file with one row per licensed user containing:
  UserPrincipalName, DisplayName, UserId, DriveId, DriveWebUrl,
  DriveType, StorageUsedGB, StorageTotalGB, DriveLastModified, Notes

.LINK
https://learn.microsoft.com/en-us/graph/api/user-list
https://learn.microsoft.com/en-us/graph/api/drive-get

.COMPONENT
Microsoft Graph API

.FUNCTIONALITY
- Authenticates with Microsoft Graph API using client credentials (secret or certificate)
- Retrieves all users with at least one assigned license, handling OData pagination
- Queries GET /users/{id}/drive for each licensed user
- Handles throttling with Retry-After / exponential backoff
- Exports results to a timestamped CSV file
#>

#region Configuration
##############################################################
#                  CONFIGURATION SECTION                     #
##############################################################
# Modify these values according to your environment

# Enable or disable verbose debug output
# Set to $true for detailed logging, $false for basic info only
$debug = $false

# Set the tenant ID and client ID for authentication
$tenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3'
$clientId = 'abc64618-283f-47ba-a185-50d935d51d57'

# Authentication type: Choose 'ClientSecret' or 'Certificate'
$AuthType = 'Certificate'  # Valid values: 'ClientSecret' or 'Certificate'

# Client Secret authentication (used when $AuthType = 'ClientSecret')
$clientSecret = ''

# Certificate authentication (used when $AuthType = 'Certificate')
$Thumbprint = 'B696FDCFE1453F3FBC6031F54DE988DA0ED905A9'

# Certificate store location: Choose 'LocalMachine' or 'CurrentUser'
$CertStore = 'LocalMachine'  # Valid values: 'LocalMachine' or 'CurrentUser'

# Delay (seconds) between each user's drive request to be gentle on the API.
# Increase this value if throttling is observed.
$delayBetweenRequests = 0

# --- Optional: scope report to a specific Entra group ---
# Leave empty ('') to report on ALL licensed users in the tenant.
# Set to an Entra group Object ID to report only on members of that group.
# Example: $groupId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
$groupId = 'ed046cb9-86bc-47e7-95f5-912cfe343fc2'

# When using group mode: set to $true to skip unlicensed group members (recommended),
# or $false to include all group members regardless of license assignment.
$groupLicensedOnly = $false

##############################################################
#                  CONFIGURATION SECTION                     #
##############################################################
#endregion Configuration

#region Initialization
# This ensures each output file has a unique name
$date = Get-Date -Format 'yyyyMMddHHmmss'
$LogName = Join-Path -Path $env:TEMP -ChildPath ("OneDrive_URL_Report_$date.csv")

# Initialize global variables for the token
$global:token = $null
$global:tokenExpiry = $null
#endregion Initialization

#region Helper Functions
# Handles throttling for Microsoft Graph requests.
# Implements Retry-After header support and exponential backoff per
# https://learn.microsoft.com/en-us/graph/throttling
function Invoke-GraphRequestWithThrottleHandling {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $true)]
        [string]$Method,

        [Parameter(Mandatory = $false)]
        [hashtable]$Headers = @{},

        [Parameter(Mandatory = $false)]
        [string]$Body = $null,

        [Parameter(Mandatory = $false)]
        [string]$ContentType = 'application/json',

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 15,

        [Parameter(Mandatory = $false)]
        [int]$InitialBackoffSeconds = 3,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 300
    )

    $retryCount = 0
    $backoffSeconds = $InitialBackoffSeconds
    $success = $false
    $result = $null

    if ($debug) { Write-Host "Graph request -> $Uri" -ForegroundColor Gray }

    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            $invokeParams = @{
                Uri         = $Uri
                Method      = $Method
                Headers     = $Headers
                ContentType = $ContentType
                TimeoutSec  = $TimeoutSeconds
                ErrorAction = 'Stop'
            }
            if ($Body) { $invokeParams['Body'] = $Body }

            $result = Invoke-RestMethod @invokeParams
            $success = $true
        }
        catch [System.Net.WebException] {
            $webEx = $_.Exception
            $statusCode = $null
            if ($webEx.Response) { $statusCode = [int]$webEx.Response.StatusCode }

            $isTransient = (
                $webEx.Status -eq [System.Net.WebExceptionStatus]::Timeout -or
                $webEx.Status -eq [System.Net.WebExceptionStatus]::ConnectionClosed -or
                $webEx.Status -eq [System.Net.WebExceptionStatus]::ConnectFailure -or
                $statusCode -in @(429, 502, 503, 504)
            )

            if ($isTransient) {
                $waitTime = if ($statusCode -eq 429 -and $webEx.Response.Headers['Retry-After']) {
                    [int]$webEx.Response.Headers['Retry-After']
                }
                else {
                    [Math]::Min($backoffSeconds, 300)
                }

                $retryCount++
                Write-Host "Transient error (status $statusCode). Waiting $waitTime s. Attempt $retryCount/$MaxRetries..." -ForegroundColor Yellow

                if ($retryCount -lt $MaxRetries) {
                    Start-Sleep -Seconds $waitTime
                    $backoffSeconds = [Math]::Min($backoffSeconds * 2, 300)
                }
                else {
                    Write-Host 'Max retries reached.' -ForegroundColor Red
                    throw $_
                }
            }
            else {
                throw $_
            }
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) { $statusCode = $_.Exception.Response.StatusCode.value__ }

            if ($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -le 599)) {
                $retryAfter = $backoffSeconds
                if ($statusCode -eq 429 -and $_.Exception.Response.Headers.Contains('Retry-After')) {
                    $retryAfter = [int]($_.Exception.Response.Headers.GetValues('Retry-After') | Select-Object -First 1)
                }

                $retryCount++
                Write-Host "Retryable error ($statusCode). Waiting $retryAfter s. Attempt $retryCount/$MaxRetries..." -ForegroundColor Yellow

                if ($retryCount -lt $MaxRetries) {
                    Start-Sleep -Seconds $retryAfter
                    $backoffSeconds = [Math]::Min($backoffSeconds * 2, 300)
                }
                else {
                    Write-Host 'Max retries reached.' -ForegroundColor Red
                    throw $_
                }
            }
            else {
                throw $_
            }
        }
    }

    return $result
}
#endregion Helper Functions

#region Authentication Functions
function AcquireToken {
    Write-Host "Connecting to Microsoft Graph using $AuthType authentication..." -ForegroundColor Cyan

    if ($AuthType -eq 'ClientSecret') {
        $uri = "https://login.microsoftonline.com/$tenantId/oauth2/token"
        $body = @{
            grant_type    = 'client_credentials'
            client_id     = $clientId
            client_secret = $clientSecret
            resource      = 'https://graph.microsoft.com'
            scope         = 'https://graph.microsoft.com/.default'
        }

        try {
            $loginResponse = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
            $global:token = $loginResponse.access_token
            $expiresIn = if ($loginResponse.expires_in) { $loginResponse.expires_in } else { 3600 }
            $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)
            Write-Host "Connected via Client Secret. Token expires: $($global:tokenExpiry)" -ForegroundColor Green
        }
        catch {
            Write-Host 'Failed to connect using Client Secret authentication.' -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            Exit
        }
    }
    elseif ($AuthType -eq 'Certificate') {
        $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

        try {
            $cert = Get-Item -Path "Cert:\$CertStore\My\$Thumbprint" -ErrorAction Stop
        }
        catch {
            Write-Host "Certificate $Thumbprint not found in $CertStore\My store." -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            Exit
        }

        $now = [System.DateTimeOffset]::UtcNow
        $exp = $now.AddMinutes(10).ToUnixTimeSeconds()
        $nbf = $now.ToUnixTimeSeconds()
        $aud = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

        $header = @{ alg = 'RS256'; typ = 'JWT'; x5t = [Convert]::ToBase64String($cert.GetCertHash()).TrimEnd('=').Replace('+', '-').Replace('/', '_') } | ConvertTo-Json -Compress
        $payload = @{ aud = $aud; exp = $exp; iss = $clientId; jti = [System.Guid]::NewGuid().ToString(); nbf = $nbf; sub = $clientId } | ConvertTo-Json -Compress

        $headerB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $payloadB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $toSign = "$headerB64.$payloadB64"
        $sigBytes = $cert.PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($toSign), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $sigB64 = [Convert]::ToBase64String($sigBytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $jwt = "$toSign.$sigB64"

        $body = @{
            client_id             = $clientId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwt
            scope                 = 'https://graph.microsoft.com/.default'
            grant_type            = 'client_credentials'
        }

        try {
            $loginResponse = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
            $global:token = $loginResponse.access_token
            $expiresIn = if ($loginResponse.expires_in) { $loginResponse.expires_in } else { 3600 }
            $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)
            Write-Host "Connected via Certificate. Token expires: $($global:tokenExpiry)" -ForegroundColor Green
        }
        catch {
            Write-Host 'Failed to connect using Certificate authentication.' -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            Exit
        }
    }
    else {
        Write-Host "Invalid AuthType '$AuthType'. Use 'ClientSecret' or 'Certificate'." -ForegroundColor Red
        Exit
    }
}

function Test-ValidToken {
    if ($null -eq $global:tokenExpiry -or (Get-Date) -gt $global:tokenExpiry) {
        Write-Host 'Token expired or expiring soon. Refreshing...' -ForegroundColor Yellow
        AcquireToken
    }
}
#endregion Authentication Functions

#region Core Functions
# Returns all users that have at least one assigned license.
# Handles OData paging (@odata.nextLink) automatically.
function Get-LicensedUsers {
    Write-Host 'Retrieving all licensed users (paginated)...' -ForegroundColor Cyan

    $users = [System.Collections.Generic.List[object]]::new()
    $nextUri = 'https://graph.microsoft.com/v1.0/users?$select=id,userPrincipalName,displayName,assignedLicenses,city,department,employeeId,officeLocation,preferredDataLocation,onPremisesExtensionAttributes&$top=999'

    do {
        Test-ValidToken
        $headers = @{ Authorization = "Bearer $global:token" }
        $response = Invoke-GraphRequestWithThrottleHandling -Uri $nextUri -Method GET -Headers $headers

        foreach ($user in $response.value) {
            if ($user.assignedLicenses -and $user.assignedLicenses.Count -gt 0) {
                $users.Add($user)
            }
        }

        $nextUri = if ($response.'@odata.nextLink') { $response.'@odata.nextLink' } else { $null }

        Write-Host "  Page retrieved. Licensed users so far: $($users.Count)" -ForegroundColor Gray
    } while ($nextUri)

    Write-Host "Total licensed users found: $($users.Count)" -ForegroundColor Green
    return $users
}

# Returns user-type members of an Entra group.
# Uses the /microsoft.graph.user type-cast endpoint so Graph server-side filters
# to user objects only -- no need to check @odata.type client-side.
# Handles OData paging (@odata.nextLink) automatically.
# Optionally filters to only members that have at least one assigned license.
function Get-GroupMembers {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $false)]
        [bool]$LicensedOnly = $true
    )

    Write-Host "Retrieving members of group '$GroupId'..." -ForegroundColor Cyan

    $users = [System.Collections.Generic.List[object]]::new()
    # /members/microsoft.graph.user casts server-side to users only, so pagination
    # pages contain only user objects and $select applies cleanly.
    $nextUri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/microsoft.graph.user?`$select=id,userPrincipalName,displayName,assignedLicenses,city,department,employeeId,officeLocation,preferredDataLocation,onPremisesExtensionAttributes&`$top=999"

    do {
        Test-ValidToken
        $headers = @{ Authorization = "Bearer $global:token" }
        $response = Invoke-GraphRequestWithThrottleHandling -Uri $nextUri -Method GET -Headers $headers

        foreach ($member in $response.value) {
            if ($LicensedOnly -and (-not $member.assignedLicenses -or $member.assignedLicenses.Count -eq 0)) {
                if ($debug) { Write-Host "  Skipping unlicensed user: $($member.userPrincipalName)" -ForegroundColor Gray }
                continue
            }

            $users.Add($member)
        }

        $nextUri = if ($response.'@odata.nextLink') { $response.'@odata.nextLink' } else { $null }

        Write-Host "  Page retrieved. Qualifying members so far: $($users.Count)" -ForegroundColor Gray
    } while ($nextUri)

    Write-Host "Total qualifying group members found: $($users.Count)" -ForegroundColor Green
    return $users
}

# Queries GET /users/{id}/drive for a single user.
# Returns a result object with drive details or an error note.
function Get-UserOneDrive {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [Parameter(Mandatory = $false)]
        [string]$UserGEO = '',

        [Parameter(Mandatory = $false)]
        [string]$City = '',

        [Parameter(Mandatory = $false)]
        [string]$Department = '',

        [Parameter(Mandatory = $false)]
        [string]$EmployeeID = '',

        [Parameter(Mandatory = $false)]
        [string]$OfficeLocation = '',

        [Parameter(Mandatory = $false)]
        [string]$HomeDrive = ''
    )

    Test-ValidToken
    $headers = @{ Authorization = "Bearer $global:token" }
    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/drive"

    $row = [PSCustomObject]@{
        UserPrincipalName = $UserPrincipalName
        DisplayName       = $DisplayName
        UserId            = $UserId
        UserGEO           = $UserGEO
        City              = $City
        Department        = $Department
        EmployeeID        = $EmployeeID
        OfficeLocation    = $OfficeLocation
        HomeDrive         = $HomeDrive
        DriveWebUrl       = ''
        StorageUsedGB     = ''
        StorageTotalGB    = ''
        DriveLastModified = ''
    }

    try {
        $drive = Invoke-GraphRequestWithThrottleHandling -Uri $uri -Method GET -Headers $headers

        $row.DriveWebUrl = $drive.webUrl -replace '/Documents$', ''
        $row.DriveLastModified = $drive.lastModifiedDateTime

        if ($drive.quota) {
            $row.StorageUsedGB = [Math]::Round($drive.quota.used / 1GB, 2)
            $row.StorageTotalGB = [Math]::Round($drive.quota.total / 1GB, 2)
        }

        Write-Host "  [OK] $UserPrincipalName -> $($drive.webUrl)" -ForegroundColor Green
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response) { $statusCode = $_.Exception.Response.StatusCode.value__ }

        $note = switch ($statusCode) {
            404 { 'OneDrive not provisioned or not found (404)' }
            403 { 'Access denied (403)' }
            $null { "Error (no HTTP response): $($_.Exception.Message)" }
            default { "Error $statusCode : $($_.Exception.Message)" }
        }

        Write-Host "  [--] $UserPrincipalName : $note" -ForegroundColor Yellow
    }

    return $row
}
#endregion Core Functions

#region Main Execution
Write-Host '=======================================' -ForegroundColor Cyan
Write-Host '  OneDrive URL Report - Microsoft Graph' -ForegroundColor Cyan
Write-Host '=======================================' -ForegroundColor Cyan

# Authenticate
AcquireToken

# Step 1: Retrieve users - either from a specific group or all licensed users in the tenant
if ($groupId -ne '') {
    Write-Host "Mode: Group members only (Group ID: $groupId)" -ForegroundColor Cyan
    $licensedUsers = Get-GroupMembers -GroupId $groupId -LicensedOnly $groupLicensedOnly
}
else {
    Write-Host 'Mode: All licensed users in the tenant' -ForegroundColor Cyan
    $licensedUsers = Get-LicensedUsers
}

if ($licensedUsers.Count -eq 0) {
    Write-Host 'No qualifying users found. Exiting.' -ForegroundColor Yellow
    Exit
}

# Step 2: Query each user's OneDrive and collect results
Write-Host "`nQuerying OneDrive for $($licensedUsers.Count) licensed users..." -ForegroundColor Cyan

$report = [System.Collections.Generic.List[object]]::new()
$total = $licensedUsers.Count
$current = 0

foreach ($user in $licensedUsers) {
    $current++
    Write-Host "[$current/$total] $($user.userPrincipalName)" -ForegroundColor Gray

    $row = Get-UserOneDrive -UserId $user.id -UserPrincipalName $user.userPrincipalName -DisplayName $user.displayName `
        -UserGEO ([string]$user.preferredDataLocation) `
        -City ([string]$user.city) `
        -Department ([string]$user.department) `
        -EmployeeID ([string]$user.employeeId) `
        -OfficeLocation ([string]$user.officeLocation) `
        -HomeDrive ([string]$user.onPremisesExtensionAttributes.extensionAttribute13)
    $report.Add($row)

    if ($delayBetweenRequests -gt 0) { Start-Sleep -Seconds $delayBetweenRequests }
}

# Step 3: Export to CSV
$report | Export-Csv -Path $LogName -NoTypeInformation -Encoding UTF8

Write-Host "`n=======================================" -ForegroundColor Cyan
Write-Host "Report complete. $($report.Count) rows exported." -ForegroundColor Green
Write-Host "Output file: $LogName" -ForegroundColor Green
Write-Host '=======================================' -ForegroundColor Cyan
#endregion Main Execution
