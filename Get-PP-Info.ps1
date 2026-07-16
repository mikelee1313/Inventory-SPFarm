<#
.SYNOPSIS
    Comprehensive Power Automate & Power Apps Environment Scanner
    Scans all Power Automate environments to extract flows, apps, connectors, and URLs.

.DESCRIPTION
    Provides complete inventory of Power Automate cloud/desktop flows and Power Apps across
    all accessible environments using delegated (user) authentication with OAuth 2.0 PKCE flow.
    Exports comprehensive CSV reports with all available metadata from API responses.

    FEATURES
    --------
    - Scans all Power Automate environments accessible to signed-in user
    - Retrieves cloud flows (automated, instant, scheduled) and desktop flows (RPA)
    - Retrieves Power Apps (canvas apps) from each environment
    - Extracts flow definitions to identify HTTP/API connectors and URLs
    - Parses connectionReferences to identify connector dependencies
    - Deduplicates flows (API may return same flow multiple times)
    - Exports two separate CSV files:
      * PowerApps_Report_[timestamp].csv — 19 columns including description, version, connectors
      * PowerAutomate_Flows_Report_[timestamp].csv — 16 columns including URLs, connectors, status

    OAUTH AUTHENTICATION
    --------------------
    Uses dual-scope OAuth 2.0 Authorization Code + PKCE (Proof Key for Code Exchange):
    - Flows Scope: https://service.flow.microsoft.com/.default (Power Automate API)
    - Apps Scope: https://service.powerapps.com/.default (Power Apps API)
    
    Scopes are requested separately because Azure AD limits one .default scope per token.
    Both scopes are silently refreshed using refresh tokens; interactive sign-in only required
    when refresh token expires.

    PREREQUISITES
    -------------
    - Entra ID app registration configured as a Public Client
    - Delegated permissions granted (and consented):
      * ProcessSimple.Environment.Read (Power Automate)
      * Flow.Read (Power Automate flows)
      * PowerApps.ReadAll (Power Apps)
    - OAuth redirect URI (http://localhost:8080) registered under "Mobile and desktop applications"
    - Port 8080 must be available when script runs
    - User must have access to target Power Automate environments

    CONFIGURATION
    --------------
    Edit the Configuration section (lines 50-80) to set:
    - $tenantId: Azure AD tenant ID
    - $clientId: Entra ID app registration Client ID
    - $redirectUri: OAuth callback URI (default: http://localhost:8080)
    - $OutputFolder: Where to save CSV files (default: $env:TEMP)
    - $MaxRetries, $InitialBackoffSec, $RequestTimeoutSec: Throttle/retry settings

    OUTPUT
    ------
    Two timestamped CSV files in $OutputFolder:
    1. PowerApps_Report_yyyyMMdd_HHmmss.csv
       Columns: AppName, AppId, Owner, Environment, AppType, State, CreatedTime, LastModifiedTime,
                Description, AppVersion, SharedUsers, SharedGroups, IsFeatured, UsesPremium,
                UsesCustomConnector, UsesOnPremise, IsCustomizable, BypassConsent, Connectors
    
    2. PowerAutomate_Flows_Report_yyyyMMdd_HHmmss.csv
       Columns: FlowName, FlowId, FlowType, FlowState, Owner, Environment, CreatedTime,
                LastModifiedTime, TemplateName, ProvisioningMethod, Plan, UserType,
                FlowFailureAlert, IsManaged, Connectors, URLs

    EXAMPLE
    -------
    PS> .\Get-PP-Info.ps1
    [Browser opens for interactive sign-in]
    [Script scans environments, retrieves flows/apps, extracts URLs]
    [Two CSV files generated in C:\Users\<user>\AppData\Local\Temp\]
#>

#region Configuration
##############################################################
#                  CONFIGURATION SECTION                     #
##############################################################

# ---- Debug output (set to $true for verbose Graph call tracing) ----
$debug = $true

# ---- Tenant & App Registration ----
$tenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3'   # Tenant ID or verified domain, e.g. 'contoso.onmicrosoft.com'
$clientId = 'abc64618-283f-47ba-a185-50d935d51d57'   # Application (client) ID of the Entra ID app registration

# ---- Redirect URI — must match a registered redirect URI on the app registration ----
$redirectUri = 'http://localhost:8080'

# ---- OAuth scopes (space-separated) ----
# Use specific scopes instead of /.default to request only what is needed (least privilege).
# Examples: 'User.Read', 'Mail.Read', 'Files.ReadWrite'
# Power Automate scope - used for retrieving flows and environments
$scopeFlows = 'https://service.flow.microsoft.com/.default'

# Power Apps scope - used separately for retrieving Power Apps (cannot be combined with Power Automate scope)
$scopeApps = 'https://service.powerapps.com/.default'

# ---- Output folder for any exported files ----
$OutputFolder = $env:TEMP

# ---- Throttle / retry settings ----
$MaxRetries = 15    # Maximum retry attempts per request
$InitialBackoffSec = 3     # Starting back-off in seconds (doubles each retry, caps at 300)
$RequestTimeoutSec = 300   # Per-request timeout in seconds

# ---- Browser sign-in listener timeout (seconds) ----
$AuthListenerTimeoutSec = 120

##############################################################
#                END CONFIGURATION SECTION                   #
##############################################################
#endregion Configuration

#region Initialization
$date = Get-Date -Format 'yyyyMMddHHmmss'
$today = (Get-Date).Date

$global:token = $null
$global:tokenExpiry = $null
$global:refreshToken = $null
#endregion Initialization

# Global token variables for dual-scope authentication
$global:tokenFlows = $null
$global:tokenFlowsExpiry = $null
$global:refreshTokenFlows = $null

$global:tokenApps = $null
$global:tokenAppsExpiry = $null
$global:refreshTokenApps = $null

#region Helper Functions

function New-PKCEParameters {
    <#
    .SYNOPSIS
        Generates a cryptographically random PKCE code verifier and its SHA-256 challenge.
        The verifier is a 32-byte random value encoded as base64url (43 characters, no padding).
        The challenge is SHA-256(verifier) encoded as base64url, per RFC 7636.
    #>
    $randomBytes = [byte[]]::new(32)
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($randomBytes)
    $codeVerifier = [Convert]::ToBase64String($randomBytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    $hasher = [System.Security.Cryptography.SHA256]::Create()
    $codeChallenge = [Convert]::ToBase64String(
        $hasher.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($codeVerifier))
    ).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    return @{
        CodeVerifier  = $codeVerifier
        CodeChallenge = $codeChallenge
    }
}

function Invoke-GraphRequestWithThrottleHandling {
    <#
    .SYNOPSIS
        Wraps Invoke-RestMethod with Retry-After / exponential-backoff throttle handling
        for Microsoft Graph API calls (429, 502, 503, 504, and network timeouts).
    .EXAMPLE
        $headers = Get-GraphAuthHeaders
        $result  = Invoke-GraphRequestWithThrottleHandling `
                       -Uri    'https://graph.microsoft.com/v1.0/me' `
                       -Method 'GET' `
                       -Headers $headers
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string]    $Uri,
        [Parameter(Mandatory)] [string]    $Method,
        [Parameter()]          [hashtable] $Headers = @{},
        [Parameter()]          [string]    $Body = $null,
        [Parameter()]          [string]    $ContentType = 'application/json',
        [Parameter()]          [int]       $MaxRetries = $script:MaxRetries,
        [Parameter()]          [int]       $InitialBackoffSeconds = $script:InitialBackoffSec,
        [Parameter()]          [int]       $TimeoutSeconds = $script:RequestTimeoutSec
    )

    $retryCount = 0
    $backoffSec = $InitialBackoffSeconds

    if ($debug) { Write-Host "  Graph -> $Method $Uri" -ForegroundColor DarkGray }

    while ($retryCount -le $MaxRetries) {
        try {
            $invokeParams = @{
                Uri         = $Uri
                Method      = $Method
                Headers     = $Headers
                ContentType = $ContentType
                TimeoutSec  = $TimeoutSeconds
                ErrorAction = 'Stop'
                Verbose     = $false
            }
            if ($Body) { $invokeParams['Body'] = $Body }

            return Invoke-RestMethod @invokeParams
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            $isRetryable = ($statusCode -in @(429, 502, 503, 504)) -or
            ($_.Exception -is [System.Net.WebException] -and
            $_.Exception.Status -in @(
                [System.Net.WebExceptionStatus]::Timeout,
                [System.Net.WebExceptionStatus]::ConnectionClosed
            ))

            if (-not $isRetryable) { throw $_ }

            if ($retryCount -ge $MaxRetries) {
                Write-Warning "Max retries ($MaxRetries) reached for: $Uri"
                throw $_
            }

            # Honour the Retry-After header when present (common on 429 and 503)
            $waitSec = $backoffSec
            if ($statusCode -in @(429, 503)) {
                try {
                    $ra = $_.Exception.Response.Headers['Retry-After']
                    if ($ra) { $waitSec = [int]$ra }
                }
                catch {}
            }

            $retryCount++
            Write-Host "    Throttled ($statusCode). Waiting ${waitSec}s (attempt $retryCount/$MaxRetries)..." `
                -ForegroundColor Yellow
            Start-Sleep -Seconds $waitSec
            $backoffSec = [Math]::Min($backoffSec * 2, 300)
        }
    }
}

function Invoke-GraphPagedRequest {
    <#
    .SYNOPSIS
        Executes a Graph GET request and automatically follows @odata.nextLink pages,
        returning all results as a single flat list.
    .EXAMPLE
        $messages = Invoke-GraphPagedRequest `
                        -Uri 'https://graph.microsoft.com/v1.0/me/messages?$select=id,subject&$top=100'
        Write-Host "Total messages: $($messages.Count)"
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string] $Uri
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $nextLink = $Uri

    do {
        Test-ValidToken
        $headers = Get-GraphAuthHeaders
        $response = Invoke-GraphRequestWithThrottleHandling -Uri $nextLink -Method 'GET' -Headers $headers

        if ($null -ne $response.value) {
            $results.AddRange([object[]]$response.value)
        }
        else {
            # Single-object response (no value array) — return directly
            return $response
        }

        $nextLink = $response.'@odata.nextLink'
        if ($debug -and $nextLink) { Write-Host '  Fetching next page...' -ForegroundColor DarkGray }
    } while ($nextLink)

    return $results
}

#endregion Helper Functions

#region Authentication Functions

function Get-TokenWithPKCE {
    <#
    .SYNOPSIS
        Performs an interactive OAuth 2.0 Authorization Code + PKCE flow to acquire a
        Microsoft Graph access token and refresh token for the signed-in user.
        Opens the default browser, listens on $redirectUri for the callback, then
        exchanges the authorization code for tokens.
        
    .PARAMETER Scope
        The OAuth scope to request (e.g., 'https://service.flow.microsoft.com/.default')
        
    .PARAMETER TokenType
        'Flows' for Power Automate token, 'Apps' for Power Apps token
    #>
    param (
        [Parameter()] [string] $Scope = $script:scopeFlows,
        [Parameter()] [ValidateSet('Flows', 'Apps')] [string] $TokenType = 'Flows'
    )
    
    Write-Host "Starting interactive PKCE authentication for $TokenType..." -ForegroundColor Cyan

    $pkce = New-PKCEParameters
    $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Build authorization URL
    $authUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/authorize" +
    "?client_id=$clientId" +
    "&response_type=code" +
    "&redirect_uri=$([uri]::EscapeDataString($redirectUri))" +
    "&scope=$([uri]::EscapeDataString($Scope))" +
    "&code_challenge=$($pkce.CodeChallenge)" +
    "&code_challenge_method=S256" +
    "&response_mode=query"

    Write-Host "Opening browser for sign-in ($TokenType scope)..." -ForegroundColor Yellow
    Write-Host "If the browser does not open automatically, navigate to:`n$authUri" -ForegroundColor Cyan
    Start-Process $authUri

    # Start local HTTP listener to capture the redirect
    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add("$redirectUri/")
    $listener.Start()

    Write-Host "`nWaiting for browser sign-in (timeout: ${AuthListenerTimeoutSec}s)..." -ForegroundColor Yellow

    # Use async GetContext so we can enforce a timeout
    $asyncResult = $listener.BeginGetContext($null, $null)
    $completed = $asyncResult.AsyncWaitHandle.WaitOne(($AuthListenerTimeoutSec * 1000))

    if (-not $completed) {
        $listener.Stop()
        throw "Authentication timed out after ${AuthListenerTimeoutSec} seconds. No response received from browser."
    }

    $context = $listener.EndGetContext($asyncResult)
    $request = $context.Request
    $response = $context.Response

    # Parse the authorization code (or error) from the redirect query string
    $authCode = $null
    if ($request.QueryString['code']) {
        $authCode = $request.QueryString['code']
        $responseHtml = '<html><body><h1>Authentication Successful</h1><p>You may close this window and return to PowerShell.</p></body></html>'
        Write-Host 'Authorization code received.' -ForegroundColor Green
    }
    else {
        $authError = $request.QueryString['error']
        $authErrorDesc = $request.QueryString['error_description']
        $responseHtml = "<html><body><h1>Authentication Failed</h1><p>$authError</p><p>$authErrorDesc</p></body></html>"
        Write-Host "Authentication failed: $authError — $authErrorDesc" -ForegroundColor Red
    }

    # Send confirmation page to the browser then shut down the listener
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseHtml)
    $response.ContentLength64 = $buffer.Length
    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.OutputStream.Close()
    $listener.Stop()

    if (-not $authCode) {
        throw 'Failed to obtain an authorization code. Check the browser for error details.'
    }

    # Exchange authorization code for tokens
    $tokenBody = @{
        grant_type    = 'authorization_code'
        client_id     = $clientId
        code          = $authCode
        redirect_uri  = $redirectUri
        code_verifier = $pkce.CodeVerifier
        scope         = $Scope
    }

    try {
        $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $tokenBody `
            -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false
        
        # Store tokens in the appropriate global variable
        if ($TokenType -eq 'Flows') {
            $global:tokenFlows = $resp.access_token
            $global:refreshTokenFlows = $resp.refresh_token
            $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
            $global:tokenFlowsExpiry = (Get-Date).AddSeconds($expiresIn - 300)
            Write-Host "  Signed in (Flows). Token valid until: $($global:tokenFlowsExpiry)" -ForegroundColor Green
        }
        else {
            $global:tokenApps = $resp.access_token
            $global:refreshTokenApps = $resp.refresh_token
            $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
            $global:tokenAppsExpiry = (Get-Date).AddSeconds($expiresIn - 300)
            Write-Host "  Signed in (Apps). Token valid until: $($global:tokenAppsExpiry)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "  Token exchange failed: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

function Update-TokenFromRefreshToken {
    <#
    .SYNOPSIS
        Silently refreshes the access token using the stored refresh token for either Flows or Apps.
        Falls back to a full interactive PKCE sign-in if no refresh token is available.
        
    .PARAMETER TokenType
        'Flows' for Power Automate token, 'Apps' for Power Apps token
    #>
    param (
        [Parameter()] [ValidateSet('Flows', 'Apps')] [string] $TokenType = 'Flows'
    )
    
    $refreshToken = if ($TokenType -eq 'Flows') { $global:refreshTokenFlows } else { $global:refreshTokenApps }
    $scope = if ($TokenType -eq 'Flows') { $script:scopeFlows } else { $script:scopeApps }
    
    if ([string]::IsNullOrWhiteSpace($refreshToken)) {
        Write-Host "No refresh token available for $TokenType — starting interactive sign-in..." -ForegroundColor Yellow
        Get-TokenWithPKCE -Scope $scope -TokenType $TokenType
        return
    }

    Write-Host "Refreshing $TokenType access token silently..." -ForegroundColor Yellow

    $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $tokenBody = @{
        grant_type    = 'refresh_token'
        client_id     = $clientId
        refresh_token = $refreshToken
        scope         = $scope
    }

    try {
        $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $tokenBody `
            -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false
        
        if ($TokenType -eq 'Flows') {
            $global:tokenFlows = $resp.access_token
            if ($resp.refresh_token) { $global:refreshTokenFlows = $resp.refresh_token }
            $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
            $global:tokenFlowsExpiry = (Get-Date).AddSeconds($expiresIn - 300)
            Write-Host "  $TokenType token refreshed. Valid until: $($global:tokenFlowsExpiry)" -ForegroundColor Green
        }
        else {
            $global:tokenApps = $resp.access_token
            if ($resp.refresh_token) { $global:refreshTokenApps = $resp.refresh_token }
            $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
            $global:tokenAppsExpiry = (Get-Date).AddSeconds($expiresIn - 300)
            Write-Host "  $TokenType token refreshed. Valid until: $($global:tokenAppsExpiry)" -ForegroundColor Green
        }
    }
    catch {
        # Refresh token may be expired or revoked — fall back to interactive sign-in
        Write-Host "  Silent refresh failed: $($_.Exception.Message). Falling back to interactive sign-in..." -ForegroundColor Yellow
        if ($TokenType -eq 'Flows') {
            $global:refreshTokenFlows = $null
        } else {
            $global:refreshTokenApps = $null
        }
        Get-TokenWithPKCE -Scope $scope -TokenType $TokenType
    }
}

function Test-ValidToken {
    <#
    .SYNOPSIS
        Checks whether the cached access token is still valid for both Flows and Apps;
        refreshes silently (or interactively as a last resort) if expired or missing.
        The token is considered stale 5 minutes before its actual expiry.
        
    .PARAMETER TokenType
        'Flows' for Power Automate token, 'Apps' for Power Apps token, 'Both' for both
    #>
    param (
        [Parameter()] [ValidateSet('Flows', 'Apps', 'Both')] [string] $TokenType = 'Flows'
    )
    
    if ($TokenType -in @('Flows', 'Both')) {
        if ($null -eq $global:tokenFlowsExpiry -or (Get-Date) -gt $global:tokenFlowsExpiry) {
            Update-TokenFromRefreshToken -TokenType 'Flows'
        }
    }
    
    if ($TokenType -in @('Apps', 'Both')) {
        if ($null -eq $global:tokenAppsExpiry -or (Get-Date) -gt $global:tokenAppsExpiry) {
            Update-TokenFromRefreshToken -TokenType 'Apps'
        }
    }
}

function Get-GraphAuthHeaders {
    <#
    .SYNOPSIS
        Returns a hashtable containing the Authorization bearer header for API calls.
        Automatically refreshes the token when it is expired or about to expire.
        
    .PARAMETER TokenType
        'Flows' for Power Automate token, 'Apps' for Power Apps token
        
    .EXAMPLE
        $headers = Get-GraphAuthHeaders -TokenType 'Flows'
        Invoke-GraphRequestWithThrottleHandling -Uri '...' -Method 'GET' -Headers $headers
    #>
    param (
        [Parameter()] [ValidateSet('Flows', 'Apps')] [string] $TokenType = 'Flows'
    )
    
    Test-ValidToken -TokenType $TokenType
    
    $token = if ($TokenType -eq 'Flows') { $global:tokenFlows } else { $global:tokenApps }
    return @{ Authorization = "Bearer $token" }
}

#endregion Authentication Functions

#region Power Automate Helper Functions

function Get-PowerAutomateEnvironments {
    <#
    .SYNOPSIS
        Retrieves all Power Automate environments accessible to the signed-in user.
    #>
    $uri = 'https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01'
    $headers = Get-GraphAuthHeaders
    $headers['Content-Type'] = 'application/json'
    
    try {
        $response = Invoke-GraphRequestWithThrottleHandling `
            -Uri $uri `
            -Method 'GET' `
            -Headers $headers
        return $response.value
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Warning "Failed to retrieve environments: $errorMsg"
        
        if ($errorMsg -like '*401*' -or $errorMsg -like '*Unauthorized*') {
            Write-Host "`n[DIAGNOSTICS]" -ForegroundColor Yellow
            Write-Host "Authentication failed. Possible causes:" -ForegroundColor Yellow
            Write-Host "1. Scope is incorrect. Current scope: $scope" -ForegroundColor Yellow
            Write-Host "2. App registration does not have Power Automate permissions" -ForegroundColor Yellow
            Write-Host "3. User has not consented to Power Automate access" -ForegroundColor Yellow
            Write-Host "`nTry re-running to trigger fresh authentication." -ForegroundColor Yellow
            Write-Host "If error persists, verify your app registration has these permissions:" -ForegroundColor Yellow
            Write-Host "  - ProcessSimple.Environment.Read" -ForegroundColor Cyan
            Write-Host "  - Flow.Read" -ForegroundColor Cyan
        }
        
        return @()
    }
}

function Get-PowerAutomateFlows {
    <#
    .SYNOPSIS
        Retrieves all cloud flows (automated, instant, scheduled) in a specific environment.
    #>
    param (
        [Parameter(Mandatory)] [string] $EnvironmentId,
        [Parameter()] [string] $FlowType = 'all'  # 'all', 'automated', 'instant', 'scheduled'
    )
    
    $uri = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/flows?api-version=2016-11-01"
    if ($FlowType -ne 'all') {
        $uri += "&`$filter=properties/flowType eq '$FlowType'"
    }
    
    $headers = Get-GraphAuthHeaders
    $headers['Content-Type'] = 'application/json'
    
    try {
        if ($debug) { Write-Host "    [DEBUG] Fetching flows from: $uri" -ForegroundColor DarkGray }
        $response = Invoke-GraphRequestWithThrottleHandling `
            -Uri $uri `
            -Method 'GET' `
            -Headers $headers
        if ($debug) { Write-Host "    [DEBUG] Response value count: $($response.value.Count)" -ForegroundColor DarkGray }
        return $response.value
    }
    catch {
        Write-Warning "Failed to retrieve flows for environment $EnvironmentId : $($_.Exception.Message)"
        return @()
    }
}

function Get-PowerAutomateDesktopFlows {
    <#
    .SYNOPSIS
        Retrieves all desktop flows (RPA) in a specific environment.
    #>
    param (
        [Parameter(Mandatory)] [string] $EnvironmentId
    )
    
    $uri = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/flows?api-version=2016-11-01&`$filter=properties/flowType eq 'DesktopFlow'"
    $headers = Get-GraphAuthHeaders
    $headers['Content-Type'] = 'application/json'
    
    try {
        if ($debug) { Write-Host "    [DEBUG] Fetching desktop flows from: $uri" -ForegroundColor DarkGray }
        $response = Invoke-GraphRequestWithThrottleHandling `
            -Uri $uri `
            -Method 'GET' `
            -Headers $headers
        return $response.value
    }
    catch {
        # Desktop flows may not be available in all environments - fail silently
        if ($debug) { Write-Host "    [DEBUG] Desktop flows unavailable: $($_.Exception.Message)" -ForegroundColor DarkGray }
        return @()
    }
}

function Get-PowerApps {
    <#
    .SYNOPSIS
        Retrieves all canvas apps in a specific environment using Power Apps API.
        Returns empty array if Power Apps token is not available.
    #>
    param (
        [Parameter(Mandatory)] [string] $EnvironmentId
    )
    
    # Check if Power Apps token is available
    if ([string]::IsNullOrWhiteSpace($global:tokenApps)) {
        if ($debug) { Write-Host "    [DEBUG] Power Apps token not available - skipping Power Apps retrieval" -ForegroundColor Yellow }
        return @()
    }
    
    # Try multiple endpoints for Power Apps
    $uris = @(
        # Primary: Power Apps API with /environments/ path
        "https://api.powerapps.com/providers/Microsoft.PowerApps/environments/$EnvironmentId/apps?api-version=2016-11-01",
        
        # Secondary: Direct Power Apps endpoint without environment scope
        "https://api.powerapps.com/providers/Microsoft.PowerApps/apps?api-version=2016-11-01",
        
        # Tertiary: Power Automate admin portal endpoint (sometimes has app access)
        "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/apps?api-version=2016-11-01"
    )
    
    $headers = Get-GraphAuthHeaders -TokenType 'Apps'
    $headers['Content-Type'] = 'application/json'
    
    foreach ($uri in $uris) {
        try {
            if ($debug) { Write-Host "    [DEBUG] Trying Power Apps endpoint: $uri" -ForegroundColor DarkGray }
            $response = Invoke-GraphRequestWithThrottleHandling `
                -Uri $uri `
                -Method 'GET' `
                -Headers $headers
            
            if ($response.value -and $response.value.Count -gt 0) {
                if ($debug) { Write-Host "    [DEBUG] Success! Found $($response.value.Count) Power Apps" -ForegroundColor Green }
                # Filter to current environment if needed
                return $response.value | Where-Object { 
                    $_.properties.environment -eq $EnvironmentId -or 
                    $_.tags.environment -eq $EnvironmentId -or 
                    $true  # Include all if no env filtering available
                }
            }
        }
        catch {
            if ($debug) { Write-Host "    [DEBUG] Endpoint failed: $($_.Exception.Message)" -ForegroundColor DarkGray }
            continue
        }
    }
    
    # If all endpoints fail, return empty
    if ($debug) { Write-Host "    [DEBUG] No Power Apps endpoints returned data" -ForegroundColor Yellow }
    return @()
}

function Get-PowerAppDetails {
    <#
    .SYNOPSIS
        Retrieves detailed definition of a Power App to extract connectors.
    #>
    param (
        [Parameter(Mandatory)] [string] $AppId,
        [Parameter(Mandatory)] [string] $EnvironmentId
    )
    
    # Try to get app definition from the app resource
    $uri = "https://api.powerapps.com/providers/Microsoft.PowerApps/apps/$AppId/definition?api-version=2020-10-01"
    $headers = Get-GraphAuthHeaders
    $headers['Content-Type'] = 'application/json'
    
    try {
        if ($debug) { Write-Host "    [DEBUG] Fetching app details from: $uri" -ForegroundColor DarkGray }
        $response = Invoke-GraphRequestWithThrottleHandling `
            -Uri $uri `
            -Method 'GET' `
            -Headers $headers
        return $response
    }
    catch {
        if ($debug) { Write-Host "    [DEBUG] Could not retrieve app definition: $($_.Exception.Message)" -ForegroundColor DarkGray }
        return $null
    }
}

function Extract-AppConnectors {
    <#
    .SYNOPSIS
        Parses Power App definition to extract connector references.
    #>
    param (
        [Parameter(Mandatory)] [PSCustomObject] $AppDefinition
    )
    
    $connectors = @()
    $urls = @()
    
    if ($null -eq $AppDefinition) { 
        return @{ Connectors = $connectors; Urls = $urls }
    }
    
    # Power Apps definition structure - look for connectorReferences
    if ($AppDefinition.connectorReferences) {
        foreach ($connRef in $AppDefinition.connectorReferences.PSObject.Properties) {
            $connectorName = $connRef.Name
            $connectorValue = $connRef.Value
            
            $connectors += @{
                Name = $connectorName
                Type = if ($connectorValue.id) { $connectorValue.id } else { 'Unknown' }
            }
        }
    }
    
    # Also check for datasources that might have URLs
    if ($AppDefinition.dataSources) {
        foreach ($ds in $AppDefinition.dataSources.PSObject.Properties) {
            $dsName = $ds.Name
            $dsValue = $ds.Value
            
            # Look for connection info
            if ($dsValue.connectionReferenceLogicalName) {
                $connectors += @{
                    Name = $dsName
                    Type = 'DataSource'
                    ConnectionRef = $dsValue.connectionReferenceLogicalName
                }
            }
            
            # Extract URLs if present
            $extractedUrls = Extract-UrlFromObject -Object $dsValue
            $urls += $extractedUrls
        }
    }
    
    return @{
        Connectors = $connectors | Select-Object -Unique
        Urls = $urls | Select-Object -Unique
    }
}

function Get-FlowDetails {
    <#
    .SYNOPSIS
        Retrieves detailed definition of a flow to extract connectors and URLs.
    #>
    param (
        [Parameter(Mandatory)] [string] $EnvironmentId,
        [Parameter(Mandatory)] [string] $FlowId
    )
    
    # Validate parameters
    if ([string]::IsNullOrWhiteSpace($FlowId)) {
        Write-Warning "FlowId parameter is empty or null"
        return $null
    }
    
    $uri = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/flows/$FlowId`?api-version=2016-11-01"
    $headers = Get-GraphAuthHeaders
    $headers['Content-Type'] = 'application/json'
    
    try {
        if ($debug) { 
            Write-Host "    [DEBUG] FlowId: '$FlowId'" -ForegroundColor DarkGray
            Write-Host "    [DEBUG] Fetching flow details from: $uri" -ForegroundColor DarkGray 
        }
        $response = Invoke-GraphRequestWithThrottleHandling `
            -Uri $uri `
            -Method 'GET' `
            -Headers $headers
        
        # Return the definition property if it exists, otherwise return the full response
        if ($response.properties.definition) {
            if ($debug) { 
                Write-Host "    [DEBUG] Flow definition structure (first 2000 chars):" -ForegroundColor DarkGray
                $defJson = $response.properties.definition | ConvertTo-Json -Depth 5 | Select-Object -First 100
                Write-Host "    $defJson" -ForegroundColor DarkGray
            }
            return $response.properties.definition
        }
        elseif ($response.definition) {
            if ($debug) { 
                Write-Host "    [DEBUG] Definition found in response.definition" -ForegroundColor DarkGray
                Write-Host "    [DEBUG] Keys: $($response.definition.PSObject.Properties.Name -join ', ')" -ForegroundColor DarkGray
            }
            return $response.definition
        }
        elseif ($response.triggers -or $response.actions) {
            # Response is already the definition
            if ($debug) { 
                Write-Host "    [DEBUG] Response is already the definition" -ForegroundColor DarkGray
                Write-Host "    [DEBUG] Triggers: $($response.triggers.PSObject.Properties.Name -join ', ')" -ForegroundColor DarkGray
                Write-Host "    [DEBUG] Actions: $($response.actions.PSObject.Properties.Name -join ', ')" -ForegroundColor DarkGray
            }
            return $response
        }
        
        if ($debug) { Write-Host "    [DEBUG] Unexpected response structure, dumping full response" -ForegroundColor DarkGray }
        return $response
    }
    catch {
        Write-Warning "Failed to retrieve flow details for flow $FlowId : $($_.Exception.Message)"
        return $null
    }
}

function Extract-UrlFromObject {
    <#
    .SYNOPSIS
        Recursively searches for URL-like values in an object.
    #>
    param (
        [Parameter(Mandatory)] [PSCustomObject] $Object,
        [int] $Depth = 0,
        [int] $MaxDepth = 10
    )
    
    $urls = @()
    
    if ($null -eq $Object -or $Depth -gt $MaxDepth) { return $urls }
    
    # Handle strings
    if ($Object -is [string]) {
        if ($Object -match 'https?://' -or $Object -match '@\{' -or $Object -match '\[') {
            $urls += $Object
        }
        return $urls
    }
    
    # Handle collections
    if ($Object -is [array]) {
        foreach ($item in $Object) {
            $urls += Extract-UrlFromObject -Object $item -Depth ($Depth + 1) -MaxDepth $MaxDepth
        }
        return $urls
    }
    
    # Handle objects - search properties
    if ($Object -is [PSCustomObject] -or $Object -is [hashtable]) {
        $properties = if ($Object -is [PSCustomObject]) { 
            $Object.PSObject.Properties 
        } else { 
            $Object.GetEnumerator() 
        }
        
        foreach ($prop in $properties) {
            $value = $prop.Value
            $propName = $prop.Name
            
            # High-priority URL properties (typically contain the actual URLs)
            if ($propName -in @('uri', 'url', 'endpoint', 'path', 'host', 'dataset')) {
                if ($value -is [string] -and $value.Length -gt 0) {
                    if ($debug -and $Depth -eq 0) { Write-Host "          [DEBUG-URL] Found URL in property '$propName': $($value.Substring(0, [Math]::Min(100, $value.Length)))" -ForegroundColor DarkGray }
                    $urls += $value
                }
            }
            
            # For SharePoint and Office365 connectors, look for list/site IDs and operation details
            if ($propName -in @('operationId', 'operationName', 'method', 'runAfter')) {
                if ($value -is [string] -and $value -notmatch '^\s*$') {
                    if ($debug -and $Depth -eq 0) { Write-Host "          [DEBUG-URL] Found operation detail '$propName': $value" -ForegroundColor DarkGray }
                    $urls += "[$propName=$value]"
                }
            }
            
            # Recursively search nested objects (max depth 10 to avoid infinite recursion)
            if ($value -is [PSCustomObject] -or $value -is [hashtable] -or $value -is [array]) {
                $urls += Extract-UrlFromObject -Object $value -Depth ($Depth + 1) -MaxDepth $MaxDepth
            }
        }
    }
    
    return $urls | Select-Object -Unique
}

function Extract-HttpConnectorsAndUrls {
    <#
    .SYNOPSIS
        Parses a flow definition to extract ALL actions and HTTP/API connector references with URLs being called.
    #>
    param (
        [Parameter(Mandatory)] [PSCustomObject] $FlowDefinition
    )
    
    $allActions = @()
    $httpConnectors = @()
    $urls = @()
    
    if ($debug) { Write-Host "      [DEBUG-EXTRACT] Starting flow definition extraction" -ForegroundColor DarkGray }
    
    # Parse triggers
    if ($FlowDefinition.triggers) {
        foreach ($trigger in $FlowDefinition.triggers.PSObject.Properties) {
            $triggerObj = $trigger.Value
            $triggerName = $trigger.Name
            
            # Collect all trigger details
            $triggerDetails = @{
                Name = $triggerName
                Type = 'Trigger'
                Kind = $triggerObj.type
                ConnectorName = if ($triggerObj.inputs.host.connection.name) { $triggerObj.inputs.host.connection.name } else { 'N/A' }
                Uri = ''
            }
            
            # Extract URL from trigger
            $extractedUrls = Extract-UrlFromObject -Object $triggerObj.inputs
            if ($extractedUrls -and $extractedUrls.Count -gt 0) {
                if ($debug) { Write-Host "        [DEBUG-EXTRACT] Trigger $triggerName found URLs: $($extractedUrls -join ', ')" -ForegroundColor DarkGray }
                # Get first valid URL
                foreach ($url in $extractedUrls) {
                    if ($url -match 'https?://') {
                        $triggerDetails['Uri'] = $url
                        break
                    }
                }
            }
            $urls += $extractedUrls
            
            # Check if it's HTTP-related or API connection
            $isApiCall = ($triggerObj.type -eq 'Http' -or 
                         $triggerObj.type -eq 'HttpWebhook' -or 
                         $triggerObj.type -eq 'OpenApiConnection' -or
                         $triggerObj.inputs.host.connection.name -like '*http*' -or
                         $triggerObj.inputs.host.connection.name -like '*sharepoint*' -or
                         $triggerObj.inputs.host.connection.name -like '*office365*')
            
            if ($isApiCall) {
                $httpConnectors += $triggerDetails
                $triggerDetails['IsHttp'] = $true
            }
            
            $allActions += $triggerDetails
        }
    }
    
    # Parse actions
    if ($FlowDefinition.actions) {
        foreach ($action in $FlowDefinition.actions.PSObject.Properties) {
            $actionObj = $action.Value
            $actionName = $action.Name
            
            # Collect all action details
            $actionDetails = @{
                Name = $actionName
                Type = 'Action'
                Kind = $actionObj.type
                ConnectorName = if ($actionObj.inputs.host.connection.name) { 
                    $actionObj.inputs.host.connection.name 
                } else { 
                    'Built-in' 
                }
                Uri = ''
                Method = ''
            }
            
            # Add method if present (for HTTP actions)
            if ($actionObj.inputs.method) {
                $actionDetails['Method'] = $actionObj.inputs.method
            }
            
            # Extract URLs from action
            $extractedUrls = Extract-UrlFromObject -Object $actionObj.inputs
            if ($extractedUrls -and $extractedUrls.Count -gt 0) {
                if ($debug) { 
                    $urlsToShow = ($extractedUrls | Select-Object -First 3) -join ', '
                    Write-Host "        [DEBUG-EXTRACT] Action $actionName extracted URLs: $urlsToShow" -ForegroundColor DarkGray 
                }
                # Get first valid HTTP URL
                foreach ($url in $extractedUrls) {
                    if ($url -match 'https?://') {
                        $actionDetails['Uri'] = $url
                        if ($debug) { Write-Host "          [DEBUG-EXTRACT] Assigned URI to $actionName : $url" -ForegroundColor DarkGray }
                        break
                    }
                }
            }
            $urls += $extractedUrls
            
            # Also check operationId for SharePoint/Office365 operations
            if ($actionObj.inputs.operationId) {
                if ($debug) { Write-Host "        [DEBUG-EXTRACT] Action $actionName has operationId: $($actionObj.inputs.operationId)" -ForegroundColor DarkGray }
                $actionDetails['Uri'] = $actionObj.inputs.operationId
            }
            
            # Check if it's HTTP-related or API connection
            $isApiCall = ($actionObj.type -eq 'Http' -or 
                         $actionObj.type -eq 'HttpWebhook' -or
                         $actionObj.type -eq 'OpenApiConnection' -or
                         $actionObj.inputs.host.connection.name -like '*http*' -or
                         $actionObj.inputs.host.connection.name -like '*sharepoint*' -or
                         $actionObj.inputs.host.connection.name -like '*office365*' -or
                         $actionObj.inputs.host.connection.connectionProperties.connectionKind -like '*http*')
            
            if ($isApiCall) {
                $httpConnectors += $actionDetails
                $actionDetails['IsHttp'] = $true
            }
            
            # Add runAfter for dependency tracking
            if ($actionObj.runAfter) {
                $dependsOn = @($actionObj.runAfter.PSObject.Properties | Select-Object -ExpandProperty Name)
                if ($dependsOn.Count -gt 0) {
                    $actionDetails['DependsOn'] = ($dependsOn -join ', ')
                }
            }
            
            $allActions += $actionDetails
        }
    }
    
    if ($debug) { Write-Host "      [DEBUG-EXTRACT] Extracted $($allActions.Count) actions, $($urls.Count) URLs" -ForegroundColor DarkGray }
    
    return @{
        AllActions = $allActions
        HttpConnectors = $httpConnectors
        Urls = $urls | Select-Object -Unique
    }
}

function Get-ConnectorStatus {
    <#
    .SYNOPSIS
        Retrieves connector status/health information for an environment.
    #>
    param (
        [Parameter(Mandatory)] [string] $EnvironmentId
    )
    
    $uri = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/connectors?api-version=2016-11-01"
    $headers = Get-GraphAuthHeaders
    $headers['Content-Type'] = 'application/json'
    
    try {
        if ($debug) { Write-Host "    [DEBUG] Fetching connectors from: $uri" -ForegroundColor DarkGray }
        $response = Invoke-GraphRequestWithThrottleHandling `
            -Uri $uri `
            -Method 'GET' `
            -Headers $headers
        return $response.value
    }
    catch {
        # Silently fail - connector endpoint may not be available in all environments
        if ($debug) { Write-Host "    [DEBUG] Connector endpoint unavailable: $($_.Exception.Message)" -ForegroundColor DarkGray }
        return @()
    }
}

#endregion Power Automate Helper Functions

#region CSV Export Functions

function Export-PowerAppsToCSV {
    <#
    .SYNOPSIS
        Exports Power Apps and their connectors to CSV format with ALL available properties from API response.
    #>
    param (
        [Parameter(Mandatory)] [PSCustomObject[]] $Apps,
        [Parameter(Mandatory)] [string] $EnvironmentName,
        [Parameter(Mandatory)] [string] $EnvironmentId
    )
    
    $csvRows = @()
    
    foreach ($app in $Apps) {
        # Extract all available properties from the app object
        $appName = if ($app.displayName) { $app.displayName } elseif ($app.properties.displayName) { $app.properties.displayName } elseif ($app.name) { $app.name } else { '' }
        $appId = if ($app.appId) { $app.appId } elseif ($app.properties.appId) { $app.properties.appId } elseif ($app.name) { $app.name } elseif ($app.id) { $app.id } else { '' }
        $owner = if ($app.owner.displayName) { $app.owner.displayName } elseif ($app.properties.owner.displayName) { $app.properties.owner.displayName } else { 'Unknown' }
        $appType = if ($app.appType) { $app.appType } elseif ($app.properties.appType) { $app.properties.appType } else { 'Unknown' }
        $appState = if ($app.state) { $app.state } elseif ($app.properties.state) { $app.properties.state } else { 'Unknown' }
        $createdTime = if ($app.properties.createdTime) { $app.properties.createdTime } elseif ($app.createdTime) { $app.createdTime } else { '' }
        $lastModifiedTime = if ($app.properties.lastModifiedTime) { $app.properties.lastModifiedTime } elseif ($app.lastModifiedTime) { $app.lastModifiedTime } else { '' }
        
        # Skip apps with empty IDs
        if ([string]::IsNullOrWhiteSpace($appId)) { continue }
        
        # Extract additional properties from app.properties
        $description = if ($app.properties.description) { $app.properties.description } else { '' }
        $appVersion = if ($app.properties.appVersion) { $app.properties.appVersion } else { '' }
        $sharedUsersCount = if ($app.properties.sharedUsersCount) { $app.properties.sharedUsersCount } else { '0' }
        $sharedGroupsCount = if ($app.properties.sharedGroupsCount) { $app.properties.sharedGroupsCount } else { '0' }
        $isFeaturedApp = if ($app.properties.isFeaturedApp) { $app.properties.isFeaturedApp } else { 'False' }
        $usesPremiumApi = if ($app.properties.usesPremiumApi) { $app.properties.usesPremiumApi } else { 'False' }
        $usesCustomApi = if ($app.properties.usesCustomApi) { $app.properties.usesCustomApi } else { 'False' }
        $usesOnPremiseGateway = if ($app.properties.usesOnPremiseGateway) { $app.properties.usesOnPremiseGateway } else { 'False' }
        $isCustomizable = if ($app.properties.isCustomizable) { $app.properties.isCustomizable } else { 'False' }
        $bypassConsent = if ($app.properties.bypassConsent) { $app.properties.bypassConsent } else { 'False' }
        
        # Extract connection references (connectors used)
        $connectors = @()
        if ($app.properties.connectionReferences) {
            $connectors = @($app.properties.connectionReferences.PSObject.Properties.Name)
        }
        $connectorList = if ($connectors.Count -gt 0) { $connectors -join '; ' } else { '' }
        
        # Build CSV row with ALL meaningful fields from API response
        $csvRows += [PSCustomObject]@{
            'AppName' = $appName
            'AppId' = $appId
            'Owner' = $owner
            'Environment' = $EnvironmentName
            'AppType' = $appType
            'State' = $appState
            'CreatedTime' = $createdTime
            'LastModifiedTime' = $lastModifiedTime
            'Description' = $description
            'AppVersion' = $appVersion
            'SharedUsers' = $sharedUsersCount
            'SharedGroups' = $sharedGroupsCount
            'IsFeatured' = $isFeaturedApp
            'UsesPremium' = $usesPremiumApi
            'UsesCustomConnector' = $usesCustomApi
            'UsesOnPremise' = $usesOnPremiseGateway
            'IsCustomizable' = $isCustomizable
            'BypassConsent' = $bypassConsent
            'Connectors' = $connectorList
        }
    }
    
    return $csvRows
}

function Export-PowerAutomateFlowsToCSV {
    <#
    .SYNOPSIS
        Exports Power Automate flows and their properties to CSV format with ALL available metadata.
        Deduplicates by FlowId to avoid duplicate entries for the same flow.
    #>
    param (
        [Parameter(Mandatory)] [PSCustomObject[]] $Flows,
        [Parameter(Mandatory)] [string] $EnvironmentName,
        [Parameter(Mandatory)] [string] $EnvironmentId,
        [Parameter()] [hashtable] $FlowUrls = @{}
    )
    
    $csvRows = @()
    $seenFlowIds = @()
    
    foreach ($flow in $Flows) {
        $flowId = $flow.name
        
        # Skip duplicate FlowIds (same flow may appear multiple times in API response)
        if ($flowId -in $seenFlowIds) {
            if ($debug) { Write-Host "    [DEBUG-EXPORT] Skipping duplicate FlowId: $flowId" -ForegroundColor DarkGray }
            continue
        }
        $seenFlowIds += $flowId
        
        # Extract all available properties from the flow object
        $flowName = $flow.properties.displayName
        $flowType = if ($flow.properties.flowType) { $flow.properties.flowType } else { 'Unknown' }
        $flowState = if ($flow.properties.state) { $flow.properties.state } else { 'Unknown' }
        $flowOwner = if ($flow.properties.owner.displayName) { $flow.properties.owner.displayName } else { 'Unknown' }
        $flowCreatedTime = if ($flow.properties.createdTime) { $flow.properties.createdTime } else { '' }
        $flowModifiedTime = if ($flow.properties.lastModifiedTime) { $flow.properties.lastModifiedTime } else { '' }
        $flowDefinitionUri = if ($flow.properties.definitionUri) { $flow.properties.definitionUri } else { '' }
        
        # Extract additional flow properties
        $templateName = if ($flow.properties.templateName) { $flow.properties.templateName } else { '' }
        $provisioningMethod = if ($flow.properties.provisioningMethod) { $flow.properties.provisioningMethod } else { '' }
        $flowPlan = if ($flow.properties.plan) { $flow.properties.plan } else { '' }
        $flowFailureAlert = if ($flow.properties.flowFailureAlertSubscribed) { $flow.properties.flowFailureAlertSubscribed } else { 'False' }
        $isManaged = if ($flow.properties.isManaged) { $flow.properties.isManaged } else { 'False' }
        $userType = if ($flow.properties.userType) { $flow.properties.userType } else { '' }
        
        # Extract connection references (connectors used in flow)
        $flowConnectors = @()
        if ($flow.properties.connectionReferences) {
            $flowConnectors = @($flow.properties.connectionReferences.PSObject.Properties | Select-Object -ExpandProperty Name)
        }
        $flowConnectorList = if ($flowConnectors.Count -gt 0) { $flowConnectors -join '; ' } else { '' }
        
        # Get URLs for this flow (extracted during flow processing)
        $flowUrlList = if ($FlowUrls.ContainsKey($flowId) -and $FlowUrls[$flowId].Count -gt 0) { 
            # Filter to only actual URLs (not operation details)
            $actualUrls = @($FlowUrls[$flowId] | Where-Object { $_ -match 'https?://' })
            if ($actualUrls.Count -gt 0) { $actualUrls -join '; ' } else { '' }
        } else { 
            '' 
        }
        
        # Build CSV row with ALL available flow properties
        $csvRows += [PSCustomObject]@{
            'FlowName' = $flowName
            'FlowId' = $flowId
            'FlowType' = $flowType
            'FlowState' = $flowState
            'Owner' = $flowOwner
            'Environment' = $EnvironmentName
            'CreatedTime' = $flowCreatedTime
            'LastModifiedTime' = $flowModifiedTime
            'TemplateName' = $templateName
            'ProvisioningMethod' = $provisioningMethod
            'Plan' = $flowPlan
            'UserType' = $userType
            'FlowFailureAlert' = $flowFailureAlert
            'IsManaged' = $isManaged
            'Connectors' = $flowConnectorList
            'URLs' = $flowUrlList
        }
    }
    
    return $csvRows
}

#endregion CSV Export Functions

#region Main
##############################################################
#                    MAIN EXECUTION                          #
##############################################################

try {
    # Clear any cached tokens to force fresh authentication
    Write-Host "Initializing authentication for Power Automate and Power Apps..." -ForegroundColor Cyan
    $global:tokenFlows = $null
    $global:tokenFlowsExpiry = $null
    $global:refreshTokenFlows = $null
    $global:tokenApps = $null
    $global:tokenAppsExpiry = $null
    $global:refreshTokenApps = $null
    
    # Request token for Power Automate Flows (required)
    Get-TokenWithPKCE -Scope $scopeFlows -TokenType 'Flows'
    
    # Try Power Apps auth (optional) - skip on any error
    Write-Host "`nAttempting Power Apps authentication (optional)..." -ForegroundColor Yellow
    try {
        Get-TokenWithPKCE -Scope $scopeApps -TokenType 'Apps'
        Write-Host "  Power Apps token acquired successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "  ⚠ Power Apps authentication unavailable: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "  Will attempt to retrieve Power Apps using Power Automate scope instead..." -ForegroundColor Cyan
        $global:tokenApps = $global:tokenFlows  # Use Flows token as fallback for Power Apps
    }

    Write-Host "`n=========================================" -ForegroundColor Cyan
    Write-Host "Power Automate & Power Apps Report" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
    Write-Host "=========================================`n" -ForegroundColor Cyan

    # Initialize collections for CSV export
    $allAppsData = @()
    $allFlowsData = @()

    # Get all environments
    Write-Host "Retrieving Power Automate environments..." -ForegroundColor Yellow
    $environments = Get-PowerAutomateEnvironments

    if ($environments.Count -eq 0) {
        Write-Host "No environments found." -ForegroundColor Yellow
        exit
    }

    Write-Host "Found $($environments.Count) environment(s):`n" -ForegroundColor Green

    # Process each environment
    foreach ($env in $environments) {
        $envName = $env.properties.displayName
        $envId = $env.name
        
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Magenta
        Write-Host "ENVIRONMENT: $envName" -ForegroundColor Magenta
        Write-Host "ID: $envId" -ForegroundColor DarkGray
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor Magenta

        # Get flows in this environment (cloud flows + desktop flows)
        Write-Host "  Retrieving cloud flows..." -ForegroundColor Yellow
        $cloudFlows = Get-PowerAutomateFlows -EnvironmentId $envId
        
        Write-Host "  Retrieving desktop flows..." -ForegroundColor Yellow
        $desktopFlows = Get-PowerAutomateDesktopFlows -EnvironmentId $envId
        
        Write-Host "  Retrieving Power Apps..." -ForegroundColor Yellow
        $powerApps = Get-PowerApps -EnvironmentId $envId
        
        $allFlows = @()
        if ($cloudFlows) { $allFlows += $cloudFlows }
        if ($desktopFlows) { $allFlows += $desktopFlows }

        if ($allFlows.Count -eq 0 -and $powerApps.Count -eq 0) {
            Write-Host "  No flows or apps found in this environment.`n" -ForegroundColor DarkGray
            continue
        }

        if ($allFlows.Count -gt 0) {
            Write-Host "  Found $($allFlows.Count) flow(s): [$($cloudFlows.Count) cloud, $($desktopFlows.Count) desktop]" -ForegroundColor Green
        }
        
        if ($powerApps.Count -gt 0) {
            Write-Host "  Found $($powerApps.Count) Power App(s)" -ForegroundColor Green
        }
        
        Write-Host ""

        # Get connector status (optional - may not be available in all environments)
        Write-Host "  Retrieving connector information..." -ForegroundColor Yellow
        $connectors = Get-ConnectorStatus -EnvironmentId $envId
        if ($connectors.Count -gt 0) {
            Write-Host "  Found $($connectors.Count) connector(s)`n" -ForegroundColor Green
        }
        else {
            Write-Host "  (Connector details unavailable for this environment)`n" -ForegroundColor DarkGray
        }

        # Process each flow
        foreach ($flow in $allFlows) {
            $flowName = $flow.properties.displayName
            $flowId = $flow.name
            $flowState = $flow.properties.state
            $flowType = $flow.properties.flowType
            if ([string]::IsNullOrWhiteSpace($flowType)) { $flowType = 'Unknown' }
            
            Write-Host "    ┌─ Flow: $flowName" -ForegroundColor Cyan
            Write-Host "    │  ID: $flowId" -ForegroundColor DarkGray
            Write-Host "    │  Type: $flowType | State: $flowState" -ForegroundColor DarkGray

            # Get flow definition - check if flowId is valid
            if ([string]::IsNullOrWhiteSpace($flowId)) {
                Write-Warning "Flow ID is empty, skipping flow details"
                Write-Host "    │" -ForegroundColor Cyan
                Write-Host "    └─────────────────────────────────────`n" -ForegroundColor Cyan
                continue
            }
            
            $flowDef = Get-FlowDetails -EnvironmentId $envId -FlowId $flowId

            if ($flowDef) {
                # Extract HTTP connectors and URLs
                $extracted = Extract-HttpConnectorsAndUrls -FlowDefinition $flowDef
                $allActions = $extracted.AllActions
                $httpConnectors = $extracted.HttpConnectors
                $urls = $extracted.Urls

                # Display ALL Actions
                if ($allActions.Count -gt 0) {
                    Write-Host "    │" -ForegroundColor Cyan
                    Write-Host "    ├─ All Actions & Triggers ($($allActions.Count)):" -ForegroundColor Yellow
                    foreach ($action in $allActions) {
                        $actionType = $action.Type
                        $actionKind = $action.Kind
                        $connName = $action.ConnectorName
                        $actionName = $action.Name
                        $isHttp = if ($action.IsHttp) { ' [HTTP]' } else { '' }
                        
                        Write-Host "    │  ├─ $actionName$isHttp" -ForegroundColor White
                        Write-Host "    │  │  └─ Type: $actionType | Kind: $actionKind | Connector: $connName" -ForegroundColor DarkGray
                        
                        if ($action.Method) {
                            Write-Host "    │  │     Method: $($action.Method)" -ForegroundColor DarkGray
                        }
                        if ($action.DependsOn) {
                            Write-Host "    │  │     DependsOn: $($action.DependsOn)" -ForegroundColor DarkGray
                        }
                    }
                }

                # Display HTTP Connectors (highlighted)
                if ($httpConnectors.Count -gt 0) {
                    Write-Host "    │" -ForegroundColor Cyan
                    Write-Host "    ├─ HTTP/API Connectors ($($httpConnectors.Count)) [HIGHLIGHTED]:" -ForegroundColor Green
                    foreach ($connector in $httpConnectors) {
                        Write-Host "    │  ├─ $($connector.Name)" -ForegroundColor Green
                        Write-Host "    │  │  ├─ Type: $($connector.Type)" -ForegroundColor White
                        Write-Host "    │  │  ├─ Kind: $($connector.Kind)" -ForegroundColor White
                        Write-Host "    │  │  └─ Connector: $($connector.ConnectorName)" -ForegroundColor White
                        if ($connector.Method) {
                            Write-Host "    │  │     Method: $($connector.Method)" -ForegroundColor White
                        }
                        if ($connector.DependsOn) {
                            Write-Host "    │  │     DependsOn: $($connector.DependsOn)" -ForegroundColor DarkGray
                        }
                    }
                }
                else {
                    # Show if there are OpenApiConnection actions even if no HTTP detected
                    $hasApiConnections = $allActions | Where-Object { $_.Kind -eq 'OpenApiConnection' }
                    if ($hasApiConnections) {
                        Write-Host "    │" -ForegroundColor Cyan
                        Write-Host "    ├─ API Connectors (OpenApiConnection): $($hasApiConnections | Measure-Object | Select-Object -ExpandProperty Count)" -ForegroundColor Cyan
                        foreach ($apiAction in $hasApiConnections) {
                            Write-Host "    │  ├─ $($apiAction.Name) [$($apiAction.ConnectorName)]" -ForegroundColor Cyan
                        }
                    }
                }

                # Display URLs
                if ($urls.Count -gt 0) {
                    Write-Host "    │" -ForegroundColor Cyan
                    Write-Host "    ├─ URLs & Endpoints ($($urls.Count)):" -ForegroundColor Cyan
                    foreach ($url in $urls) {
                        $urlDisplay = if ($url.Length -gt 120) { "$($url.Substring(0, 117))..." } else { $url }
                        Write-Host "    │  ├─ $urlDisplay" -ForegroundColor White
                    }
                }
            }
            else {
                Write-Host "    │" -ForegroundColor Cyan
                Write-Host "    ├─ (Could not retrieve flow definition)" -ForegroundColor Yellow
            }

            Write-Host "    │" -ForegroundColor Cyan
            Write-Host "    └─────────────────────────────────────`n" -ForegroundColor Cyan
        }

        # Process Power Apps
        if ($powerApps.Count -gt 0) {
            Write-Host "  Power Apps:" -ForegroundColor Yellow
            foreach ($app in $powerApps) {
                # Handle different property names from different API endpoints
                $appName = if ($app.displayName) { $app.displayName } elseif ($app.properties.displayName) { $app.properties.displayName } elseif ($app.name) { $app.name } else { '(Unknown)' }
                $appId = if ($app.appId) { $app.appId } elseif ($app.properties.appId) { $app.properties.appId } elseif ($app.name) { $app.name } elseif ($app.id) { $app.id } else { '' }
                
                # Skip apps with empty IDs
                if ([string]::IsNullOrWhiteSpace($appId)) {
                    if ($debug) { Write-Host "    [DEBUG] Skipping app with empty ID: $appName" -ForegroundColor Yellow }
                    continue
                }
                
                $appType = if ($app.appType) { $app.appType } elseif ($app.properties.appType) { $app.properties.appType } else { 'Unknown' }
                $appState = if ($app.state) { $app.state } elseif ($app.properties.state) { $app.properties.state } else { 'Unknown' }
                
                Write-Host "    ┌─ App: $appName" -ForegroundColor Magenta
                Write-Host "    │  ID: $appId" -ForegroundColor DarkGray
                Write-Host "    │  Type: $appType | State: $appState" -ForegroundColor DarkGray
                
                # Try to get app details
                $appDef = Get-PowerAppDetails -AppId $appId -EnvironmentId $envId
                
                if ($appDef) {
                    $appConnectorInfo = Extract-AppConnectors -AppDefinition $appDef
                    $appConnectors = $appConnectorInfo.Connectors
                    $appUrls = $appConnectorInfo.Urls
                    
                    # Display connectors
                    if ($appConnectors.Count -gt 0) {
                        Write-Host "    │" -ForegroundColor Magenta
                        Write-Host "    ├─ Connectors ($($appConnectors.Count)):" -ForegroundColor Green
                        foreach ($connector in $appConnectors) {
                            Write-Host "    │  ├─ $($connector.Name)" -ForegroundColor White
                            if ($connector.Type -and $connector.Type -ne 'Unknown') {
                                Write-Host "    │  │  └─ Type: $($connector.Type)" -ForegroundColor DarkGray
                            }
                            if ($connector.ConnectionRef) {
                                Write-Host "    │  │  └─ ConnectionRef: $($connector.ConnectionRef)" -ForegroundColor DarkGray
                            }
                        }
                    }
                    
                    # Display URLs
                    if ($appUrls.Count -gt 0) {
                        Write-Host "    │" -ForegroundColor Magenta
                        Write-Host "    ├─ URLs & Endpoints ($($appUrls.Count)):" -ForegroundColor Cyan
                        foreach ($url in $appUrls) {
                            $urlDisplay = if ($url.Length -gt 120) { "$($url.Substring(0, 117))..." } else { $url }
                            Write-Host "    │  ├─ $urlDisplay" -ForegroundColor White
                        }
                    }
                }
                
                Write-Host "    │" -ForegroundColor Magenta
                Write-Host "    └─────────────────────────────────────`n" -ForegroundColor Magenta
            }
        }

        # Collect data for CSV export
        Write-Host "  Collecting data for export..." -ForegroundColor Yellow
        
        # Collect Power Apps data
        if ($powerApps.Count -gt 0) {
            if ($debug) { Write-Host "    [DEBUG-EXPORT] Processing $($powerApps.Count) Power Apps" -ForegroundColor DarkGray }
            $appsExport = Export-PowerAppsToCSV -Apps $powerApps -EnvironmentName $envName -EnvironmentId $envId
            $allAppsData += $appsExport
            if ($debug) { Write-Host "    [DEBUG-EXPORT] Exported $($appsExport.Count) Power App rows" -ForegroundColor DarkGray }
        }
        
        # Collect Power Automate flows data with URLs
        if ($allFlows.Count -gt 0) {
            if ($debug) { Write-Host "    [DEBUG-EXPORT] Processing $($allFlows.Count) flows" -ForegroundColor DarkGray }
            
            # Build hashtable of flow URLs for export function
            $flowUrlMap = @{}
            foreach ($flow in $allFlows) {
                $flowId = $flow.name
                if (-not [string]::IsNullOrWhiteSpace($flowId)) {
                    $flowDef = Get-FlowDetails -EnvironmentId $envId -FlowId $flowId
                    if ($flowDef) {
                        $extracted = Extract-HttpConnectorsAndUrls -FlowDefinition $flowDef
                        $flowUrlMap[$flowId] = $extracted.Urls
                    }
                }
            }
            
            $flowsExport = Export-PowerAutomateFlowsToCSV -Flows $allFlows -EnvironmentName $envName -EnvironmentId $envId -FlowUrls $flowUrlMap
            $allFlowsData += $flowsExport
            if ($debug) { Write-Host "    [DEBUG-EXPORT] Exported $($flowsExport.Count) flow action rows" -ForegroundColor DarkGray }
            if ($debug) { 
                foreach ($row in $flowsExport | Select-Object -First 3) {
                    Write-Host "      Row: Name=$($row.FlowName), URLs=$($row.URLs)" -ForegroundColor DarkGray
                }
            }
        }

        Write-Host "`n" -ForegroundColor Magenta
    }

    # Export all data to CSV
    Write-Host "`n=========================================" -ForegroundColor Green
    Write-Host "Exporting to separate CSV files..." -ForegroundColor Yellow
    Write-Host "=========================================`n" -ForegroundColor Green
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    
    # Export Power Apps to separate file
    if ($allAppsData.Count -gt 0) {
        $appsOutputPath = Join-Path $OutputFolder "PowerApps_Report_$timestamp.csv"
        $allAppsData | Export-Csv -Path $appsOutputPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "Power Apps report exported to: $appsOutputPath" -ForegroundColor Green
        Write-Host "  Total Power Apps exported: $($allAppsData.Count) rows" -ForegroundColor Cyan
    }
    else {
        Write-Host "  No Power Apps to export." -ForegroundColor DarkGray
    }
    
    # Export Power Automate Flows to separate file
    if ($allFlowsData.Count -gt 0) {
        $flowsOutputPath = Join-Path $OutputFolder "PowerAutomate_Flows_Report_$timestamp.csv"
        $allFlowsData | Export-Csv -Path $flowsOutputPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "Power Automate flows report exported to: $flowsOutputPath" -ForegroundColor Green
        Write-Host "  Total flow actions exported: $($allFlowsData.Count) rows" -ForegroundColor Cyan
    }
    else {
        Write-Host "  No Power Automate flows to export." -ForegroundColor DarkGray
    }
    
    Write-Host "`nScript completed successfully." -ForegroundColor Green
}
catch {
    Write-Host "`nScript failed: $($_.Exception.Message)" -ForegroundColor Red
    if ($debug) {
        Write-Host $_.Exception.StackTrace -ForegroundColor Red
    }
    throw
}

#endregion Main
