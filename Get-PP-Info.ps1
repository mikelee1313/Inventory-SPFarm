<#
.SYNOPSIS
        Power Platform Inventory Scanner (Power Automate + Power Apps)
        Discovers environments, flows, apps, connectors, and extracted endpoint URLs.

.DESCRIPTION
        Enumerates Power Automate and Power Apps assets across all environments the signed-in
        user can access. Uses delegated OAuth 2.0 Authorization Code + PKCE and exports
        detailed CSV reports.

        The script supports modern Power Platform endpoints and automatically falls back to
        legacy Microsoft.ProcessSimple endpoints when required by tenant/environment behavior.
        This enables discovery of flows that may not appear on newer endpoints.

    FEATURES
    --------
        - Retrieves all accessible environments with include/exclude filtering
        - Collects cloud flows, desktop flows, and legacy ProcessSimple flows
        - Collects Power Apps metadata and connector references
        - Resolves flow details from modern and legacy detail endpoints
        - Extracts endpoint URLs from flow definitions (including SharePoint dataset URLs when present)
        - Normalizes flow types to business-friendly labels (for example: Cloud - Automated)
        - Deduplicates flow entries returned by multiple endpoints
        - Supports explicit flow URL force-include via `$IncludeFlowUrls`
        - Produces diagnostic JSON samples when debug dump toggles are enabled

    OAUTH AUTHENTICATION
    --------------------
        Uses one OAuth sign-in for both Flows and Apps:
        - Scope: `https://api.powerplatform.com/.default offline_access`

        The `offline_access` scope is required so the script can obtain a refresh token,
        silently refresh access tokens, and exchange for legacy Flow scopes when needed.

    PREREQUISITES
    -------------
    - Entra ID app registration configured as a Public Client
    - Delegated permissions granted (and consented):
      * ProcessSimple.Environment.Read (Power Automate)
      * Flow.Read (Power Automate flows)
      * PowerApps.ReadAll (Power Apps)
        - OAuth redirect URI registered under "Mobile and desktop applications":
            * `http://localhost:8081`
        - Local machine must be able to open browser auth and receive callback on redirect port
    - User and APP must have access to target Power Automate environments

    CONFIGURATION
        -------------
        Edit the Configuration section to set:
    - $tenantId: Azure AD tenant ID
    - $clientId: Entra ID app registration Client ID
        - $redirectUri: OAuth callback URI
        - $collectionScope: `Flows`, `Apps`, or `Both`
        - $Environment / $ExcludeEnvironment: optional environment filters
        - $IncludeFlowUrls: optional direct flow include list from maker URLs
    - $OutputFolder: Where to save CSV files (default: $env:TEMP)
    - $MaxRetries, $InitialBackoffSec, $RequestTimeoutSec: Throttle/retry settings
        - $EnableLegacyFlowDiscoveryFallback / $LegacyFlowScopes: legacy discovery behavior
        - $DumpFlowPayloadSamples / $DumpOneFlowObjectPerEnvironment: debug sample dumps

    OUTPUT
    ------
    Two timestamped CSV files in $OutputFolder:
     1. PowerApps_Report_yyyyMMdd_HHmmss.csv
         Columns: AppName, AppId, Owner, Environment, EnvironmentId, AppType, State, CreatedTime, LastModifiedTime,
                Description, AppVersion, SharedUsers, SharedGroups, IsFeatured, UsesPremium,
                UsesCustomConnector, UsesOnPremise, IsCustomizable, BypassConsent, AppDetailFetchStatus,
                Connectors, DatasetUrls
    
    2. PowerAutomate_Flows_Report_yyyyMMdd_HHmmss.csv
         Columns: FlowName, FlowId, FlowType, FlowState, Owner, Environment, EnvironmentId, CreatedTime,
                LastModifiedTime, TemplateName, ProvisioningMethod, Plan, UserType, OwnerIds,
                FlowFailureAlert, IsManaged, Connectors, URLs

                 FlowType values are normalized labels, including:
                 - Cloud - Automated
                 - Cloud - Instant
                 - Cloud - Scheduled
                 - Cloud - Other
                 - Desktop - RPA

    EXAMPLE
    -------
    PS> .\Get-PP-Info.ps1
        [Interactive sign-in opens in browser]
        [Script discovers environments, scans flows/apps, and exports CSV files]

    Set `$Environment = @('Default-Contoso', 'Sandbox')`
    PS> .\Get-PP-Info.ps1
    [Only the specified environments are scanned]

    Set `$ExcludeEnvironment = @('Trial', 'Personal Productivity')`
    PS> .\Get-PP-Info.ps1
    [All accessible environments except the excluded ones are scanned]

    Set `$IncludeFlowUrls = @('https://make.powerautomate.com/environments/<envId>/flows/<flowId>/details')`
    PS> .\Get-PP-Info.ps1
    [Flow is force-included by direct lookup even if not returned by list endpoints]
#>

#region Configuration
##############################################################
#                  CONFIGURATION SECTION                     #
##############################################################

# ---- Debug output (set to $true for verbose Graph call tracing) ----
$debug = $false

# ---- Tenant & App Registration ----
$tenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3'   # Tenant ID or verified domain, e.g. 'contoso.onmicrosoft.com'
$clientId = 'abc64618-283f-47ba-a185-50d935d51d57'   # Application (client) ID of the Entra ID app registration

# ---- Redirect URI — must match a registered redirect URI on the app registration ----
$redirectUri = 'http://localhost:8081'

# ---- OAuth scope (single token for both Flows and Apps) ----
# offline_access is required so the script receives a refresh token and can
# exchange it for legacy ProcessSimple scopes when newer flow endpoints 404.
$scope = 'https://api.powerplatform.com/.default offline_access'

# ---- Output folder for any exported files ----
$OutputFolder = $env:TEMP

# ---- Throttle / retry settings ----
$MaxRetries = 15    # Maximum retry attempts per request
$InitialBackoffSec = 3     # Starting back-off in seconds (doubles each retry, caps at 300)
$RequestTimeoutSec = 300   # Per-request timeout in seconds

# ---- Browser sign-in listener timeout (seconds) ----
$AuthListenerTimeoutSec = 120

# ---- Heartbeat progress output interval ----
# Prints a progress line every N items while processing flows/apps in an environment.
$ProgressUpdateEvery = 25

# ---- Optional debug payload dump ----
# When enabled, writes one sample cloud flow payload per environment to TEMP for schema analysis.
$DumpFlowPayloadSamples = $false

# When enabled, writes one detailed flow object JSON per environment (raw object + parsed definition + extracted URLs).
$DumpOneFlowObjectPerEnvironment = $false

# Enable legacy flow discovery fallbacks for broader tenant coverage.
# Some customer tenants expose additional flows through Microsoft.ProcessSimple endpoints.
$EnableLegacyFlowDiscoveryFallback = $true

# Alternate scopes to try when querying legacy ProcessSimple endpoints.
# These are used only for legacy endpoint calls and are acquired silently via refresh token when possible.
$LegacyFlowScopes = @(
    'https://service.flow.microsoft.com/.default',
    'https://api.flow.microsoft.com/.default'
)

# ---- Collection scope: what resources to collect ----
# Valid values: 'Flows', 'Apps', or 'Both'
# 'Flows' - collect only Power Automate flows
# 'Apps' - collect only Power Apps
# 'Both' - collect both flows and apps (default)
$collectionScope = 'Both'

# ---- Environment filters ----
# Match by environment display name or environment ID.
# Leave empty arrays to scan all accessible environments.
$Environment =  @()
$ExcludeEnvironment = @()

# ---- Explicit flow URLs to force-include (optional) ----
# Use this when a flow is visible in maker portal but missing from list endpoints.
$IncludeFlowUrls = @()



##############################################################
#                END CONFIGURATION SECTION                   #
##############################################################
#endregion Configuration

#region Initialization
$global:token = $null
$global:tokenExpiry = $null
$global:refreshToken = $null
#endregion Initialization

# Unified token variables for Power Platform API
$global:tokenPP = $null
$global:tokenPPExpiry = $null
$global:refreshTokenPP = $null
$global:tokenLegacyFlow = $null
$global:tokenLegacyFlowExpiry = $null
$global:tokenLegacyFlowScope = $null

function Set-DebugTokenContext {
    <#
    .SYNOPSIS
        Mirrors the active token into the legacy token variables for easier debugging.
    #>
    $global:token = $global:tokenPP
    $global:tokenExpiry = $global:tokenPPExpiry
    $global:refreshToken = $global:refreshTokenPP
}

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
        Power Platform access token and refresh token for the signed-in user.
        Opens the default browser, listens on the configured redirect URI for the callback, then
        exchanges the authorization code for tokens.
        
    .PARAMETER Scope
        The OAuth scope to request (e.g., 'https://api.powerplatform.com/.default offline_access')
        
    .PARAMETER TokenType
        'Flows' for Power Automate token, 'Apps' for Power Apps token
    #>
    param (
        [Parameter()] [string] $Scope = $script:scope,
        [Parameter()] [ValidateSet('Flows', 'Apps')] [string] $TokenType = 'Flows'
    )

    $redirectUri = $script:redirectUri
    Write-Host "Starting interactive PKCE authentication for Power Platform..." -ForegroundColor Cyan

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

    Write-Host "Opening browser for sign-in..." -ForegroundColor Yellow
    Write-Host "Using redirect URI: $redirectUri" -ForegroundColor DarkGray
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
        
        $global:tokenPP = $resp.access_token
        $global:refreshTokenPP = $resp.refresh_token
        $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
        $global:tokenPPExpiry = (Get-Date).AddSeconds($expiresIn - 300)
        Set-DebugTokenContext
        Write-Host "  Signed in (Power Platform). Token valid until: $($global:tokenPPExpiry)" -ForegroundColor Green
        if ([string]::IsNullOrWhiteSpace($global:refreshTokenPP)) {
            Write-Host '  No refresh token was returned. Legacy flow fallback endpoints will be unavailable for this sign-in.' -ForegroundColor Yellow
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
        Silently refreshes the unified Power Platform access token.
        Falls back to a full interactive PKCE sign-in if no refresh token is available.
        
    .PARAMETER TokenType
        'Flows' for Power Automate token, 'Apps' for Power Apps token
    #>
    param (
        [Parameter()] [ValidateSet('Flows', 'Apps')] [string] $TokenType = 'Flows'
    )
    
    $refreshToken = $global:refreshTokenPP
    $scope = $script:scope
    
    if ([string]::IsNullOrWhiteSpace($refreshToken)) {
        Write-Host "No refresh token available for $TokenType — starting interactive sign-in..." -ForegroundColor Yellow
        Get-TokenWithPKCE -Scope $scope -TokenType $TokenType
        return
    }

    Write-Host "Refreshing Power Platform access token silently..." -ForegroundColor Yellow

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
        
        $global:tokenPP = $resp.access_token
        if ($resp.refresh_token) { $global:refreshTokenPP = $resp.refresh_token }
        $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
        $global:tokenPPExpiry = (Get-Date).AddSeconds($expiresIn - 300)
        Set-DebugTokenContext
        Write-Host "  Token refreshed. Valid until: $($global:tokenPPExpiry)" -ForegroundColor Green
        if ([string]::IsNullOrWhiteSpace($global:refreshTokenPP)) {
            Write-Host '  No refresh token is cached after refresh. Legacy flow fallback endpoints remain unavailable.' -ForegroundColor Yellow
        }
    }
    catch {
        # Refresh token may be expired or revoked — fall back to interactive sign-in
        Write-Host "  Silent refresh failed: $($_.Exception.Message). Falling back to interactive sign-in..." -ForegroundColor Yellow
        $global:refreshTokenPP = $null
        Get-TokenWithPKCE -Scope $scope -TokenType $TokenType
    }
}

function Test-ValidToken {
    <#
    .SYNOPSIS
        Checks whether the cached access token is still valid;
        refreshes silently (or interactively as a last resort) if expired or missing.
        The token is considered stale 5 minutes before its actual expiry.
        
    .PARAMETER TokenType
        'Flows' for Power Automate token, 'Apps' for Power Apps token, 'Both' for both
    #>
    param (
        [Parameter()] [ValidateSet('Flows', 'Apps', 'Both')] [string] $TokenType = 'Flows'
    )
    
    if ($null -eq $global:tokenPPExpiry -or (Get-Date) -gt $global:tokenPPExpiry) {
        Update-TokenFromRefreshToken -TokenType 'Flows'
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

    Set-DebugTokenContext
    $token = $global:token
    return @{ Authorization = "Bearer $token" }
}

function Test-IsLegacyFlowEndpoint {
    param (
        [Parameter(Mandatory)] [string] $Uri
    )

    return ($Uri -like 'https://api.flow.microsoft.com/*' -or $Uri -match '/providers/Microsoft\.ProcessSimple/')
}

function Get-LegacyFlowAuthHeaders {
    <#
    .SYNOPSIS
        Returns auth headers for legacy ProcessSimple endpoints by exchanging refresh token
        against legacy Flow scopes when needed.
    #>

    if ($global:tokenLegacyFlow -and $global:tokenLegacyFlowExpiry -and (Get-Date) -lt $global:tokenLegacyFlowExpiry) {
        return @{ Authorization = "Bearer $($global:tokenLegacyFlow)" }
    }

    if ([string]::IsNullOrWhiteSpace($global:refreshTokenPP)) {
        if ($debug) { Write-Host '    [DEBUG] No refresh token available for legacy flow scope exchange.' -ForegroundColor DarkGray }
        return $null
    }

    $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    foreach ($legacyScope in @($LegacyFlowScopes | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        try {
            if ($debug) { Write-Host "    [DEBUG] Attempting legacy scope exchange: $legacyScope" -ForegroundColor DarkGray }
            $tokenBody = @{
                grant_type    = 'refresh_token'
                client_id     = $clientId
                refresh_token = $global:refreshTokenPP
                scope         = $legacyScope
            }

            $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $tokenBody `
                -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false

            if ($resp.access_token) {
                $global:tokenLegacyFlow = $resp.access_token
                $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
                $global:tokenLegacyFlowExpiry = (Get-Date).AddSeconds($expiresIn - 300)
                $global:tokenLegacyFlowScope = $legacyScope
                if ($resp.refresh_token) { $global:refreshTokenPP = $resp.refresh_token }
                if ($debug) { Write-Host "    [DEBUG] Legacy token acquired using scope: $legacyScope" -ForegroundColor DarkGray }
                return @{ Authorization = "Bearer $($global:tokenLegacyFlow)" }
            }
        }
        catch {
            if ($debug) { Write-Host "    [DEBUG] Legacy scope exchange failed for $legacyScope : $($_.Exception.Message)" -ForegroundColor DarkGray }
        }
    }

    return $null
}

#endregion Authentication Functions

#region Power Automate Helper Functions

function Get-PowerAutomateEnvironments {
    <#
    .SYNOPSIS
        Retrieves all Power Automate environments accessible to the signed-in user.
    #>
    $headers = Get-GraphAuthHeaders -TokenType 'Flows'
    $headers['Content-Type'] = 'application/json'

    # Environment discovery is done via Power Apps endpoints; flow endpoints require
    # an environment ID in the path and do not reliably expose a tenant-level list route.
    $uris = @(
        'https://api.powerplatform.com/powerapps/environments?api-version=2024-10-01',
        'https://api.powerapps.com/providers/Microsoft.PowerApps/environments?api-version=2016-11-01'
    )
    
    foreach ($uri in $uris) {
        try {
            $environments = [System.Collections.Generic.List[object]]::new()
            $seenEnvironmentIds = @{}
            $nextUri = $uri

            while ($nextUri) {
                $response = Invoke-GraphRequestWithThrottleHandling `
                    -Uri $nextUri `
                    -Method 'GET' `
                    -Headers $headers

                foreach ($environmentEntry in @($response.value)) {
                    $environmentId = try { [string]$environmentEntry.name } catch { $null }
                    if ([string]::IsNullOrWhiteSpace($environmentId)) {
                        $environmentId = [guid]::NewGuid().ToString()
                    }

                    if (-not $seenEnvironmentIds.ContainsKey($environmentId)) {
                        $seenEnvironmentIds[$environmentId] = $true
                        $environments.Add($environmentEntry)
                    }
                }

                $nextUri = $response.'@odata.nextLink'
                if (-not $nextUri) {
                    $nextUri = try { $response.nextLink } catch { $null }
                }
            }

            return @($environments)
        }
        catch {
            $errorMsg = $_.Exception.Message
            if ($debug) { Write-Host "    [DEBUG] Environment discovery endpoint failed: $uri : $errorMsg" -ForegroundColor DarkGray }

            if ($errorMsg -like '*401*' -or $errorMsg -like '*Unauthorized*') {
                Write-Host "`n[DIAGNOSTICS]" -ForegroundColor Yellow
                Write-Host "Authentication failed. Possible causes:" -ForegroundColor Yellow
                Write-Host "1. Scope is incorrect. Current scope: $scope" -ForegroundColor Yellow
                Write-Host "2. App registration does not have Power Platform API delegated permissions" -ForegroundColor Yellow
                Write-Host "3. User has not consented to Power Platform access" -ForegroundColor Yellow
                Write-Host "`nTry re-running to trigger fresh authentication." -ForegroundColor Yellow
            }
        }
    }

    Write-Warning "Failed to retrieve environments from available endpoints."
    return @()
}

function Get-PowerAppsEnvironments {
    <#
    .SYNOPSIS
        Retrieves all Power Apps environments accessible to the signed-in user.
    #>
    $headers = Get-GraphAuthHeaders -TokenType 'Apps'
    $headers['Content-Type'] = 'application/json'

    $uris = @(
        'https://api.powerplatform.com/powerapps/environments?api-version=2024-10-01'
    )

    foreach ($uri in $uris) {
        try {
            $environments = [System.Collections.Generic.List[object]]::new()
            $seenEnvironmentIds = @{}
            $nextUri = $uri

            while ($nextUri) {
                $response = Invoke-GraphRequestWithThrottleHandling `
                    -Uri $nextUri `
                    -Method 'GET' `
                    -Headers $headers

                foreach ($environmentEntry in @($response.value)) {
                    $environmentId = try { [string]$environmentEntry.name } catch { $null }
                    if ([string]::IsNullOrWhiteSpace($environmentId)) {
                        $environmentId = [guid]::NewGuid().ToString()
                    }

                    if (-not $seenEnvironmentIds.ContainsKey($environmentId)) {
                        $seenEnvironmentIds[$environmentId] = $true
                        $environments.Add($environmentEntry)
                    }
                }

                $nextUri = $response.'@odata.nextLink'
                if (-not $nextUri) {
                    $nextUri = try { $response.nextLink } catch { $null }
                }
            }

            return @($environments)
        }
        catch {
            if ($debug) { Write-Host "    [DEBUG] Power Apps environments endpoint failed: $uri : $($_.Exception.Message)" -ForegroundColor DarkGray }
        }
    }

    Write-Warning "Failed to retrieve Power Apps environments from available endpoints."
    return @()
}

function Select-Environments {
    <#
    .SYNOPSIS
        Filters environments by include and exclude lists using display name or environment ID.
    #>
    param (
        [Parameter()] [AllowNull()] [object[]] $Environments = @(),
        [Parameter()] [string[]] $IncludeEnvironment,
        [Parameter()] [string[]] $ExcludeEnvironment
    )

    if ($null -eq $Environments -or @($Environments).Count -eq 0) {
        return @()
    }

    $includeLookup = @{}
    foreach ($value in @($IncludeEnvironment)) {
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $includeLookup[$value.Trim().ToLowerInvariant()] = $true
        }
    }

    $excludeLookup = @{}
    foreach ($value in @($ExcludeEnvironment)) {
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $excludeLookup[$value.Trim().ToLowerInvariant()] = $true
        }
    }

    $filteredEnvironments = foreach ($environmentEntry in $Environments) {
        $environmentId = [string]$environmentEntry.name
        $environmentName = [string]$environmentEntry.properties.displayName
        $candidateKeys = @($environmentId, $environmentName) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { $_.Trim().ToLowerInvariant() }

        $isIncluded = ($includeLookup.Count -eq 0)
        if (-not $isIncluded) {
            $isIncluded = @($candidateKeys | Where-Object { $includeLookup.ContainsKey($_) }).Count -gt 0
        }

        if (-not $isIncluded) {
            continue
        }

        $isExcluded = @($candidateKeys | Where-Object { $excludeLookup.ContainsKey($_) }).Count -gt 0
        if ($isExcluded) {
            continue
        }

        $environmentEntry
    }

    return @($filteredEnvironments)
}

function Get-FlowTargetsFromUrls {
    <#
    .SYNOPSIS
        Parses maker portal flow URLs into environment/flow ID targets.
    #>
    param (
        [Parameter()] [string[]] $FlowUrls = @()
    )

    $targets = [System.Collections.Generic.List[object]]::new()
    foreach ($url in @($FlowUrls)) {
        if ([string]::IsNullOrWhiteSpace($url)) { continue }

        if ($url -match '/environments/([^/]+)/flows/([^/?#]+)') {
            $envId = [uri]::UnescapeDataString($matches[1])
            $flowId = [uri]::UnescapeDataString($matches[2])
            if (-not [string]::IsNullOrWhiteSpace($envId) -and -not [string]::IsNullOrWhiteSpace($flowId)) {
                $targets.Add([PSCustomObject]@{
                        EnvironmentId = $envId
                        FlowId        = $flowId
                        SourceUrl     = $url
                    })
            }
        }
    }

    return @($targets)
}

function Get-PowerAutomateFlowById {
    <#
    .SYNOPSIS
        Retrieves a single flow by ID from environment-scoped Power Platform endpoints.
    #>
    param (
        [Parameter(Mandatory)] [string] $EnvironmentId,
        [Parameter(Mandatory)] [string] $FlowId
    )

    $headers = Get-GraphAuthHeaders -TokenType 'Flows'
    $headers['Content-Type'] = 'application/json'

    $encodedFlowId = [uri]::EscapeDataString($FlowId)
    $uris = @(
        "https://api.powerplatform.com/powerautomate/environments/$EnvironmentId/cloudFlows/$encodedFlowId?api-version=2024-10-01",
        "https://api.powerplatform.com/powerautomate/environments/$EnvironmentId/flows/$encodedFlowId?api-version=2024-10-01"
    )

    if ($EnableLegacyFlowDiscoveryFallback) {
        $uris += @(
            "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/flows/$encodedFlowId?api-version=2016-11-01",
            "https://api.powerapps.com/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/flows/$encodedFlowId?api-version=2016-11-01"
        )
    }

    foreach ($uri in $uris) {
        try {
            if ($debug) { Write-Host "    [DEBUG] Direct flow lookup: $uri" -ForegroundColor DarkGray }

            $activeHeaders = $headers
            if (Test-IsLegacyFlowEndpoint -Uri $uri) {
                $legacyHeaders = Get-LegacyFlowAuthHeaders
                if ($legacyHeaders) {
                    $activeHeaders = $legacyHeaders
                    $activeHeaders['Content-Type'] = 'application/json'
                }
            }

            $response = Invoke-GraphRequestWithThrottleHandling -Uri $uri -Method 'GET' -Headers $activeHeaders
            if ($null -ne $response) {
                return $response
            }
        }
        catch {
            if ($debug) { Write-Host "    [DEBUG] Direct flow lookup failed: $uri : $($_.Exception.Message)" -ForegroundColor DarkGray }
        }
    }

    return $null
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
    
    $headers = Get-GraphAuthHeaders -TokenType 'Flows'
    $headers['Content-Type'] = 'application/json'

    $uris = @(
        "https://api.powerplatform.com/powerautomate/environments/$EnvironmentId/cloudFlows?api-version=2024-10-01",
        "https://api.powerplatform.com/powerautomate/environments/$EnvironmentId/flows?api-version=2024-10-01"
    )

    if ($EnableLegacyFlowDiscoveryFallback) {
        $uris += @(
            "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/flows?api-version=2016-11-01",
            "https://api.powerapps.com/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/flows?api-version=2016-11-01"
        )
    }

    if ($FlowType -ne 'all') {
        $uris = @($uris | ForEach-Object { "$_&flowType=$FlowType" })
    }

    $collectedFlows = [System.Collections.Generic.List[object]]::new()
    $seenFlowKeys = @{}
    
    foreach ($uri in $uris) {
        $nextUri = $uri
        $pageNumber = 0

        while ($nextUri) {
            try {
                if ($debug) { Write-Host "    [DEBUG] Fetching flows from: $nextUri" -ForegroundColor DarkGray }

                $activeHeaders = $headers
                if (Test-IsLegacyFlowEndpoint -Uri $nextUri) {
                    $legacyHeaders = Get-LegacyFlowAuthHeaders
                    if ($legacyHeaders) {
                        $activeHeaders = $legacyHeaders
                        $activeHeaders['Content-Type'] = 'application/json'
                        if ($debug) {
                            Write-Host "    [DEBUG] Using legacy auth scope: $($global:tokenLegacyFlowScope)" -ForegroundColor DarkGray
                        }
                    }
                }

                $response = Invoke-GraphRequestWithThrottleHandling `
                    -Uri $nextUri `
                    -Method 'GET' `
                    -Headers $activeHeaders
                $flowsFromEndpoint = @($response.value)
                if ($debug) { Write-Host "    [DEBUG] Response value count: $($flowsFromEndpoint.Count)" -ForegroundColor DarkGray }

                foreach ($flow in $flowsFromEndpoint) {
                    $flowKey = try { [string]$flow.workflowId } catch { $null }
                    if ([string]::IsNullOrWhiteSpace($flowKey)) { $flowKey = try { [string]$flow.name } catch { $null } }
                    if ([string]::IsNullOrWhiteSpace($flowKey)) { $flowKey = try { [string]$flow.id } catch { $null } }
                    if ([string]::IsNullOrWhiteSpace($flowKey)) { $flowKey = [guid]::NewGuid().ToString() }

                    if (-not $seenFlowKeys.ContainsKey($flowKey)) {
                        $seenFlowKeys[$flowKey] = $true
                        $collectedFlows.Add($flow)
                    }
                }

                if ($DumpFlowPayloadSamples -and $pageNumber -eq 0 -and $response.value -and @($response.value).Count -gt 0) {
                    try {
                        $samplePath = Join-Path $OutputFolder ("FlowSample_{0}_{1}.json" -f ($EnvironmentId -replace '[^a-zA-Z0-9_-]', '_'), (Get-Date -Format 'yyyyMMdd_HHmmss'))
                        @($response.value | Select-Object -First 1) | ConvertTo-Json -Depth 100 | Out-File -FilePath $samplePath -Encoding utf8
                        Write-Host "    [DEBUG] Wrote flow payload sample: $samplePath" -ForegroundColor DarkGray
                    }
                    catch {
                        if ($debug) { Write-Host "    [DEBUG] Failed to write flow payload sample: $($_.Exception.Message)" -ForegroundColor DarkGray }
                    }
                }

                $nextUri = $response.'@odata.nextLink'
                if (-not $nextUri) {
                    $nextUri = try { $response.nextLink } catch { $null }
                }
                if ($nextUri -and $debug) {
                    Write-Host "    [DEBUG] Fetching next flow page..." -ForegroundColor DarkGray
                }
                $pageNumber++
            }
            catch {
                if ($debug) { Write-Host "    [DEBUG] Flow endpoint failed: $nextUri : $($_.Exception.Message)" -ForegroundColor DarkGray }
                $nextUri = $null
            }
        }
    }

    if ($collectedFlows.Count -gt 0) {
        if ($debug) { Write-Host "    [DEBUG] Total deduplicated flows across endpoints: $($collectedFlows.Count)" -ForegroundColor DarkGray }
        return @($collectedFlows)
    }

    Write-Warning "Failed to retrieve flows for environment $EnvironmentId from available endpoints."
    return @()
}

function Get-PowerAutomateDesktopFlows {
    <#
    .SYNOPSIS
        Retrieves all desktop flows (RPA) in a specific environment.
    #>
    param (
        [Parameter(Mandatory)] [string] $EnvironmentId
    )
    
    $headers = Get-GraphAuthHeaders -TokenType 'Flows'
    $headers['Content-Type'] = 'application/json'

    $uris = @(
        "https://api.powerplatform.com/powerautomate/environments/$EnvironmentId/desktopFlows?api-version=2024-10-01",
        "https://api.powerplatform.com/powerautomate/environments/$EnvironmentId/flows?api-version=2024-10-01&flowType=DesktopFlow"
    )
    
    $desktopFlows = [System.Collections.Generic.List[object]]::new()
    $seenDesktopFlowKeys = @{}

    foreach ($uri in $uris) {
        $nextUri = $uri

        while ($nextUri) {
            try {
                if ($debug) { Write-Host "    [DEBUG] Fetching desktop flows from: $nextUri" -ForegroundColor DarkGray }
                $response = Invoke-GraphRequestWithThrottleHandling `
                    -Uri $nextUri `
                    -Method 'GET' `
                    -Headers $headers

                foreach ($flow in @($response.value)) {
                    $flowKey = try { [string]$flow.workflowId } catch { $null }
                    if ([string]::IsNullOrWhiteSpace($flowKey)) { $flowKey = try { [string]$flow.name } catch { $null } }
                    if ([string]::IsNullOrWhiteSpace($flowKey)) { $flowKey = try { [string]$flow.id } catch { $null } }
                    if ([string]::IsNullOrWhiteSpace($flowKey)) { $flowKey = [guid]::NewGuid().ToString() }

                    if (-not $seenDesktopFlowKeys.ContainsKey($flowKey)) {
                        $seenDesktopFlowKeys[$flowKey] = $true
                        $desktopFlows.Add($flow)
                    }
                }

                $nextUri = $response.'@odata.nextLink'
                if (-not $nextUri) {
                    $nextUri = try { $response.nextLink } catch { $null }
                }

                if ($nextUri -and $debug) {
                    Write-Host "    [DEBUG] Fetching next desktop flow page..." -ForegroundColor DarkGray
                }
            }
            catch {
                if ($debug) { Write-Host "    [DEBUG] Desktop flow endpoint failed: $nextUri : $($_.Exception.Message)" -ForegroundColor DarkGray }
                $nextUri = $null
            }
        }
    }

    return @($desktopFlows)
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
    
    # Check if unified token is available
    if ([string]::IsNullOrWhiteSpace($global:tokenPP)) {
        if ($debug) { Write-Host "    [DEBUG] Power Platform token not available - skipping Power Apps retrieval" -ForegroundColor Yellow }
        return @()
    }
    
    # Try environment-scoped endpoints only so per-environment scans stay bounded.
    $uris = @(
        # Power Platform apps endpoint (environment-scoped)
        "https://api.powerplatform.com/powerapps/environments/$EnvironmentId/apps?api-version=2024-10-01"
    )
    
    $headers = Get-GraphAuthHeaders -TokenType 'Apps'
    $headers['Content-Type'] = 'application/json'
    
    $apps = [System.Collections.Generic.List[object]]::new()
    $seenAppKeys = @{}

    foreach ($uri in $uris) {
        $nextUri = $uri

        while ($nextUri) {
            try {
                if ($debug) { Write-Host "    [DEBUG] Trying Power Apps endpoint: $nextUri" -ForegroundColor DarkGray }
                $response = Invoke-GraphRequestWithThrottleHandling `
                    -Uri $nextUri `
                    -Method 'GET' `
                    -Headers $headers
            
            $responseValues = @($response.value)
                foreach ($app in $responseValues) {
                    $appKey = try { [string]$app.appId } catch { $null }
                    if ([string]::IsNullOrWhiteSpace($appKey)) { $appKey = try { [string]$app.properties.appId } catch { $null } }
                    if ([string]::IsNullOrWhiteSpace($appKey)) { $appKey = try { [string]$app.name } catch { $null } }
                    if ([string]::IsNullOrWhiteSpace($appKey)) { $appKey = try { [string]$app.id } catch { $null } }
                    if ([string]::IsNullOrWhiteSpace($appKey)) { $appKey = [guid]::NewGuid().ToString() }

                    if (-not $seenAppKeys.ContainsKey($appKey)) {
                        $seenAppKeys[$appKey] = $true
                        $apps.Add($app)
                    }
                }

                $nextUri = $response.'@odata.nextLink'
                if (-not $nextUri) {
                    $nextUri = try { $response.nextLink } catch { $null }
                }

                if ($nextUri -and $debug) {
                    Write-Host "    [DEBUG] Fetching next Power Apps page..." -ForegroundColor DarkGray
                }
            }
            catch {
                if ($debug) { Write-Host "    [DEBUG] Endpoint failed: $($_.Exception.Message)" -ForegroundColor DarkGray }
                $nextUri = $null
                continue
            }
        }
    }
    
    if ($debug) {
        if ($apps.Count -gt 0) {
            Write-Host "    [DEBUG] Success! Found $($apps.Count) Power Apps" -ForegroundColor Green
        }
        else {
            Write-Host "    [DEBUG] No Power Apps endpoints returned data" -ForegroundColor Yellow
        }
    }

    return @($apps)
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
    
    # Try multiple app definition endpoints for cross-tenant compatibility.
    $uris = @(
        "https://api.powerplatform.com/powerapps/environments/$EnvironmentId/apps/$AppId/definition?api-version=2024-10-01",
        "https://api.powerapps.com/providers/Microsoft.PowerApps/environments/$EnvironmentId/apps/$AppId/definition?api-version=2016-11-01",
        "https://api.powerapps.com/providers/Microsoft.PowerApps/apps/$AppId/definition?api-version=2022-11-01",
        "https://api.powerapps.com/providers/Microsoft.PowerApps/apps/$AppId/definition?api-version=2020-10-01"
    )

    $headers = Get-GraphAuthHeaders -TokenType 'Apps'
    $headers['Content-Type'] = 'application/json'

    foreach ($uri in $uris) {
        try {
            if ($debug) { Write-Host "    [DEBUG] Fetching app details from: $uri" -ForegroundColor DarkGray }
            $response = Invoke-GraphRequestWithThrottleHandling `
                -Uri $uri `
                -Method 'GET' `
                -Headers $headers

            return [PSCustomObject]@{
                Success  = $true
                Status   = 'Success'
                Endpoint = $uri
                Error    = ''
                Definition = $response
            }
        }
        catch {
            if ($debug) { Write-Host "    [DEBUG] Could not retrieve app definition from endpoint: $uri : $($_.Exception.Message)" -ForegroundColor DarkGray }
        }
    }

    return [PSCustomObject]@{
        Success  = $false
        Status   = 'DefinitionUnavailable'
        Endpoint = ''
        Error    = 'All app definition endpoints failed'
        Definition = $null
    }
}

function Get-AppConnectorsFromDefinition {
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
                    Name          = $dsName
                    Type          = 'DataSource'
                    ConnectionRef = $dsValue.connectionReferenceLogicalName
                }
            }
            
            # Extract URLs if present
            $extractedUrls = Get-UrlsFromObject -Object $dsValue
            $urls += $extractedUrls
        }
    }
    
    return @{
        Connectors = $connectors | Select-Object -Unique
        Urls       = $urls | Select-Object -Unique
    }
}

function Get-FlowDetails {
    <#
    .SYNOPSIS
        Retrieves detailed definition of a flow to extract connectors and URLs.
    #>
    param (
        [Parameter(Mandatory)] [string] $EnvironmentId,
        [Parameter(Mandatory)] [string[]] $FlowIds
    )
    
    # Validate parameters
    if ($null -eq $FlowIds -or @($FlowIds).Count -eq 0) {
        Write-Warning "FlowIds parameter is empty or null"
        return $null
    }
    
    $headers = Get-GraphAuthHeaders -TokenType 'Flows'
    $headers['Content-Type'] = 'application/json'

    foreach ($flowId in @($FlowIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        # Flow IDs can contain spaces and special characters; encode for path safety.
        $encodedFlowId = [uri]::EscapeDataString($flowId)

        $uris = @(
            "https://api.powerplatform.com/powerautomate/environments/$EnvironmentId/cloudFlows/${encodedFlowId}?api-version=2024-10-01",
            "https://api.powerplatform.com/powerautomate/environments/$EnvironmentId/flows/${encodedFlowId}?api-version=2024-10-01"
        )

        if ($EnableLegacyFlowDiscoveryFallback) {
            $uris += @(
                "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/flows/${encodedFlowId}?api-version=2016-11-01",
                "https://api.powerapps.com/providers/Microsoft.ProcessSimple/environments/$EnvironmentId/flows/${encodedFlowId}?api-version=2016-11-01"
            )
        }

        foreach ($uri in $uris) {
            try {
                if ($debug) {
                    Write-Host "    [DEBUG] FlowId candidate: '$flowId'" -ForegroundColor DarkGray
                    Write-Host "    [DEBUG] Fetching flow details from: $uri" -ForegroundColor DarkGray
                }

                $activeHeaders = $headers
                if (Test-IsLegacyFlowEndpoint -Uri $uri) {
                    $legacyHeaders = Get-LegacyFlowAuthHeaders
                    if ($legacyHeaders) {
                        $activeHeaders = $legacyHeaders
                        $activeHeaders['Content-Type'] = 'application/json'
                    }
                }

                $response = Invoke-GraphRequestWithThrottleHandling `
                    -Uri $uri `
                    -Method 'GET' `
                    -Headers $activeHeaders
        
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
                if ($debug) { Write-Host "    [DEBUG] Flow details endpoint failed: $uri : $($_.Exception.Message)" -ForegroundColor DarkGray }
            }
        }
    }

    if ($debug) { Write-Host "    [DEBUG] Could not retrieve flow details for candidate IDs: $(@($FlowIds) -join ', ')" -ForegroundColor DarkGray }
    return $null
}

function Get-FlowIdCandidates {
    <#
    .SYNOPSIS
        Builds a prioritized set of candidate flow identifiers from a flow list item.
    #>
    param (
        [Parameter(Mandatory)] [PSCustomObject] $Flow
    )

    $candidates = [System.Collections.Generic.List[string]]::new()

    function Get-StringPropertyValue {
        param (
            [Parameter()] $InputObject,
            [Parameter(Mandatory)] [string] $PropertyName
        )

        if ($null -eq $InputObject) { return $null }

        $prop = $InputObject.PSObject.Properties[$PropertyName]
        if ($null -eq $prop) { return $null }

        $value = [string]$prop.Value
        if ([string]::IsNullOrWhiteSpace($value)) { return $null }
        return $value
    }

    $properties = $null
    if ($Flow.PSObject.Properties['properties']) {
        $properties = $Flow.properties
    }

    $rawCandidates = @(
        (Get-StringPropertyValue -InputObject $Flow -PropertyName 'workflowId'),
        (Get-StringPropertyValue -InputObject $properties -PropertyName 'flowId'),
        (Get-StringPropertyValue -InputObject $properties -PropertyName 'workflowId'),
        (Get-StringPropertyValue -InputObject $properties -PropertyName 'workflowEntityId'),
        (Get-StringPropertyValue -InputObject $properties -PropertyName 'name'),
        (Get-StringPropertyValue -InputObject $Flow -PropertyName 'name'),
        (Get-StringPropertyValue -InputObject $Flow -PropertyName 'id'),
        (Get-StringPropertyValue -InputObject $properties -PropertyName 'id'),
        (Get-StringPropertyValue -InputObject $properties -PropertyName 'resourceId')
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($candidate in $rawCandidates) {
        if (-not $candidates.Contains($candidate)) {
            $candidates.Add($candidate)
        }

        # Extract terminal ID from resource paths such as /.../cloudFlows/{id} or /.../flows/{id}
        if ($candidate -match '/(?:cloudFlows|flows)/([^/?]+)') {
            $pathId = [uri]::UnescapeDataString($matches[1])
            if (-not [string]::IsNullOrWhiteSpace($pathId) -and -not $candidates.Contains($pathId)) {
                $candidates.Add($pathId)
            }
        }
    }

    return @($candidates)
}

function Get-EmbeddedFlowDefinition {
    <#
    .SYNOPSIS
        Extracts workflow definition from a flow list item when the API returns definition inline.
    #>
    param (
        [Parameter(Mandatory)] [PSCustomObject] $Flow
    )

    $rawDefinition = $null
    if ($Flow.PSObject.Properties['definition']) {
        $rawDefinition = $Flow.definition
    }
    elseif ($Flow.PSObject.Properties['properties'] -and $Flow.properties -and $Flow.properties.PSObject.Properties['definition']) {
        $rawDefinition = $Flow.properties.definition
    }

    if ($null -eq $rawDefinition) {
        return $null
    }

    $parsed = $null
    if ($rawDefinition -is [string]) {
        $trimmed = $rawDefinition.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            return $null
        }

        try {
            $parsed = $trimmed | ConvertFrom-Json -Depth 100
        }
        catch {
            if ($debug) { Write-Host "    [DEBUG] Could not parse embedded flow definition JSON: $($_.Exception.Message)" -ForegroundColor DarkGray }
            return $null
        }
    }
    elseif ($rawDefinition -is [PSCustomObject] -or $rawDefinition -is [hashtable]) {
        $parsed = $rawDefinition
    }
    else {
        return $null
    }

    if ($parsed.PSObject.Properties['properties'] -and $parsed.properties -and $parsed.properties.PSObject.Properties['definition']) {
        return $parsed.properties.definition
    }
    if ($parsed.PSObject.Properties['definition']) {
        return $parsed.definition
    }
    if ($parsed.PSObject.Properties['triggers'] -or $parsed.PSObject.Properties['actions']) {
        return $parsed
    }

    return $null
}

function Get-EmbeddedFlowConnectionReferences {
    <#
    .SYNOPSIS
        Extracts connector reference names and API names from an embedded flow definition.
    #>
    param (
        [Parameter(Mandatory)] [PSCustomObject] $Flow
    )

    $references = @()

    $rawDefinition = $null
    if ($Flow.PSObject.Properties['definition']) {
        $rawDefinition = $Flow.definition
    }
    elseif ($Flow.PSObject.Properties['properties'] -and $Flow.properties -and $Flow.properties.PSObject.Properties['definition']) {
        $rawDefinition = $Flow.properties.definition
    }

    if ($rawDefinition -isnot [string] -or [string]::IsNullOrWhiteSpace($rawDefinition)) {
        return @()
    }

    try {
        $embeddedObj = $rawDefinition | ConvertFrom-Json -Depth 100
        if ($embeddedObj.properties -and $embeddedObj.properties.connectionReferences) {
            foreach ($refProp in $embeddedObj.properties.connectionReferences.PSObject.Properties) {
                $refValue = $refProp.Value
                $apiName = try { [string]$refValue.api.name } catch { '' }
                $apiId = try { [string]$refValue.api.id } catch { '' }
                $references += [PSCustomObject]@{
                    ReferenceName = [string]$refProp.Name
                    ApiName       = $apiName
                    ApiId         = $apiId
                }
            }
        }
    }
    catch {
        if ($debug) { Write-Host "    [DEBUG] Could not parse embedded connectionReferences: $($_.Exception.Message)" -ForegroundColor DarkGray }
    }

    return @($references)
}

function Resolve-FlowType {
    <#
    .SYNOPSIS
        Normalizes flow type values into stable, user-friendly categories.
    #>
    param (
        [Parameter(Mandatory)] [PSCustomObject] $Flow,
        [Parameter()] $FlowDefinition = $null
    )

    $properties = $null
    if ($Flow.PSObject.Properties['properties']) {
        $properties = $Flow.properties
    }

    function Get-OptionalStringValue {
        param (
            [Parameter()] $InputObject,
            [Parameter(Mandatory)] [string] $PropertyPath
        )

        if ($null -eq $InputObject) { return $null }

        try {
            $value = $InputObject
            foreach ($segment in $PropertyPath -split '\.') {
                if ($null -eq $value) { return $null }
                $prop = $value.PSObject.Properties[$segment]
                if ($null -eq $prop) { return $null }
                $value = $prop.Value
            }

            $stringValue = [string]$value
            if ([string]::IsNullOrWhiteSpace($stringValue)) { return $null }
            return $stringValue
        }
        catch {
            return $null
        }
    }

    $rawTypeCandidates = @(
        (Get-OptionalStringValue -InputObject $properties -PropertyPath 'flowType'),
        (Get-OptionalStringValue -InputObject $Flow -PropertyPath 'flowType'),
        (Get-OptionalStringValue -InputObject $Flow -PropertyPath 'modernFlowType'),
        (Get-OptionalStringValue -InputObject $Flow -PropertyPath 'type'),
        (Get-OptionalStringValue -InputObject $properties -PropertyPath 'type')
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($candidate in $rawTypeCandidates) {
        switch -Regex ($candidate) {
            'DesktopFlow' { return 'Desktop - RPA' }
            '^Automated$' { return 'Cloud - Automated' }
            '^Instant$' { return 'Cloud - Instant' }
            '^Scheduled$' { return 'Cloud - Scheduled' }
        }
    }

    $definitionToInspect = $FlowDefinition
    if (-not $definitionToInspect) {
        $definitionToInspect = Get-EmbeddedFlowDefinition -Flow $Flow
    }

    $triggerEntries = @()
    if ($definitionToInspect -and $definitionToInspect.PSObject.Properties['triggers'] -and $definitionToInspect.triggers) {
        $triggerEntries = @($definitionToInspect.triggers.PSObject.Properties)
    }

    if ($triggerEntries.Count -gt 0) {
        foreach ($triggerEntry in $triggerEntries) {
            $triggerName = [string]$triggerEntry.Name
            $triggerObj = $triggerEntry.Value
            $triggerType = Get-OptionalStringValue -InputObject $triggerObj -PropertyPath 'type'
            if (-not $triggerType) { $triggerType = '' }
            $operationId = Get-OptionalStringValue -InputObject $triggerObj -PropertyPath 'inputs.host.operationId'
            if (-not $operationId) { $operationId = '' }

            if ($triggerType -eq 'Recurrence') {
                return 'Cloud - Scheduled'
            }

            if ($triggerType -in @('Request', 'Manual')) {
                return 'Cloud - Instant'
            }

            if ($triggerName -match '^(manual|button)$' -or $operationId -match 'manual|button|powerapp') {
                return 'Cloud - Instant'
            }
        }

        return 'Cloud - Automated'
    }

    foreach ($candidate in $rawTypeCandidates) {
        if ($candidate -match 'PowerAutomateFlow|Microsoft\.ProcessSimple/environments/flows') {
            return 'Cloud - Other'
        }
    }

    return 'Unknown'
}

function Get-UrlsFromObject {
    <#
    .SYNOPSIS
        Recursively searches for URL-like values in an object.
    #>
    param (
        [Parameter()] $Object,
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
            $urls += Get-UrlsFromObject -Object $item -Depth ($Depth + 1) -MaxDepth $MaxDepth
        }
        return $urls
    }
    
    # Handle objects - search properties
    if ($Object -is [PSCustomObject] -or $Object -is [hashtable]) {
        $properties = if ($Object -is [PSCustomObject]) { 
            $Object.PSObject.Properties 
        }
        else { 
            $Object.GetEnumerator() 
        }
        
        foreach ($prop in $properties) {
            $value = $prop.Value
            $propName = $prop.Name

            # Some payloads (for example connectionReferences.dataSets) store URLs as hashtable keys.
            if ($propName -is [string] -and $propName -match '^https?://') {
                $urls += $propName
            }
            
            # High-priority URL properties (typically contain the actual URLs)
            if ($propName -in @('uri', 'url', 'endpoint', 'path', 'host', 'dataset')) {
                if ($value -is [string] -and $value.Length -gt 0) {
                    if ($debug -and $Depth -eq 0) { Write-Host "          [DEBUG-URL] Found URL in property '$propName': $($value.Substring(0, [Math]::Min(100, $value.Length)))" -ForegroundColor DarkGray }
                    $urls += $value
                }
            }

            # Capture URL-like values regardless of property name (important for connector parameter bags).
            if ($value -is [string] -and $value -match 'https?://') {
                $urls += $value
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
                $urls += Get-UrlsFromObject -Object $value -Depth ($Depth + 1) -MaxDepth $MaxDepth
            }
        }
    }
    
    return $urls | Select-Object -Unique
}

function Get-HttpConnectorsAndUrlsFromFlow {
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

    # Also scan the serialized definition text for explicit SharePoint/HTTP URLs.
    try {
        $definitionText = $FlowDefinition | ConvertTo-Json -Depth 100 -Compress
        if (-not [string]::IsNullOrWhiteSpace($definitionText)) {
            $urlMatches = [regex]::Matches($definitionText, 'https?://[^\"''\s,]+')
            foreach ($match in $urlMatches) {
                $urls += $match.Value
            }
        }
    }
    catch {}
    
    # Parse triggers
    if ($FlowDefinition.triggers) {
        foreach ($trigger in $FlowDefinition.triggers.PSObject.Properties) {
            $triggerObj = $trigger.Value
            $triggerName = $trigger.Name
            $triggerConnName = try { $triggerObj.inputs.host.connection.name } catch { $null }
            
            # Collect all trigger details
            $triggerDetails = @{
                Name          = $triggerName
                Type          = 'Trigger'
                Kind          = $triggerObj.type
                ConnectorName = if ($triggerConnName) { $triggerConnName } else { 'N/A' }
                Uri           = ''
            }
            
            # Extract URL from trigger
            $triggerInputs = try { $triggerObj.inputs } catch { $null }
            $extractedUrls = if ($null -ne $triggerInputs) {
                Get-UrlsFromObject -Object $triggerInputs
            }
            else {
                @()
            }
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
                $triggerConnName -like '*http*' -or
                $triggerConnName -like '*sharepoint*' -or
                $triggerConnName -like '*office365*')
            
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
            $actionConnName = try { $actionObj.inputs.host.connection.name } catch { $null }
            $actionConnKind = try { $actionObj.inputs.host.connection.connectionProperties.connectionKind } catch { $null }
            
            # Collect all action details
            $actionDetails = @{
                Name          = $actionName
                Type          = 'Action'
                Kind          = $actionObj.type
                ConnectorName = if ($actionConnName) { 
                    $actionConnName 
                }
                else { 
                    'Built-in' 
                }
                Uri           = ''
                Method        = ''
            }
            
            # Add method if present (for HTTP actions)
            $inputsMethod = try { $actionObj.inputs.method } catch { $null }
            if ($inputsMethod) {
                $actionDetails['Method'] = $inputsMethod
            }
            
            # Extract URLs from action
            $actionInputs = try { $actionObj.inputs } catch { $null }
            $extractedUrls = if ($null -ne $actionInputs) {
                Get-UrlsFromObject -Object $actionInputs
            }
            else {
                @()
            }
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
            $inputsOpId = try { $actionObj.inputs.operationId } catch { $null }
            if ($inputsOpId) {
                if ($debug) { Write-Host "        [DEBUG-EXTRACT] Action $actionName has operationId: $inputsOpId" -ForegroundColor DarkGray }
                $actionDetails['Uri'] = $inputsOpId
            }
            
            # Check if it's HTTP-related or API connection
            $isApiCall = ($actionObj.type -eq 'Http' -or 
                $actionObj.type -eq 'HttpWebhook' -or
                $actionObj.type -eq 'OpenApiConnection' -or
                $actionConnName -like '*http*' -or
                $actionConnName -like '*sharepoint*' -or
                $actionConnName -like '*office365*' -or
                $actionConnKind -like '*http*')
            
            if ($isApiCall) {
                $httpConnectors += $actionDetails
                $actionDetails['IsHttp'] = $true
            }
            
            # Add runAfter for dependency tracking
            if ($actionObj.runAfter) {
                $dependsOn = @()
                try {
                    $dependsOn = @($actionObj.runAfter.PSObject.Properties | Select-Object -ExpandProperty Name)
                }
                catch {
                    $dependsOn = @()
                }
                if ($dependsOn.Count -gt 0) {
                    $actionDetails['DependsOn'] = ($dependsOn -join ', ')
                }
            }
            
            $allActions += $actionDetails
        }
    }
    
    if ($debug) { Write-Host "      [DEBUG-EXTRACT] Extracted $($allActions.Count) actions, $($urls.Count) URLs" -ForegroundColor DarkGray }
    
    return @{
        AllActions     = $allActions
        HttpConnectors = $httpConnectors
        Urls           = $urls | Select-Object -Unique
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
    
    $headers = Get-GraphAuthHeaders -TokenType 'Flows'
    $headers['Content-Type'] = 'application/json'

    $uris = @(
        "https://api.powerplatform.com/powerautomate/environments/$EnvironmentId/connectors?api-version=2024-10-01"
    )
    
    foreach ($uri in $uris) {
        try {
            if ($debug) { Write-Host "    [DEBUG] Fetching connectors from: $uri" -ForegroundColor DarkGray }
            $response = Invoke-GraphRequestWithThrottleHandling `
                -Uri $uri `
                -Method 'GET' `
                -Headers $headers
            return @($response.value)
        }
        catch {
            # Silently fail - connector endpoint may not be available in all environments
            if ($debug) { Write-Host "    [DEBUG] Connector endpoint unavailable: $($_.Exception.Message)" -ForegroundColor DarkGray }
        }
    }

    return @()
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
        $appName = try { $app.displayName } catch { $null }
        if (-not $appName) { $appName = try { $app.properties.displayName } catch { $null } }
        if (-not $appName) { $appName = try { $app.name } catch { $null } }
        if (-not $appName) { $appName = '' }

        $appId = try { $app.appId } catch { $null }
        if (-not $appId) { $appId = try { $app.properties.appId } catch { $null } }
        if (-not $appId) { $appId = try { $app.name } catch { $null } }
        if (-not $appId) { $appId = try { $app.id } catch { $null } }
        if (-not $appId) { $appId = '' }

        $owner = try { $app.owner.displayName } catch { $null }
        if (-not $owner) { $owner = try { $app.properties.owner.displayName } catch { $null } }
        if (-not $owner) { $owner = 'Unknown' }

        $appType = try { $app.appType } catch { $null }
        if (-not $appType) { $appType = try { $app.properties.appType } catch { $null } }
        if (-not $appType) { $appType = 'Unknown' }

        $appState = try { $app.state } catch { $null }
        if (-not $appState) { $appState = try { $app.properties.state } catch { $null } }
        if (-not $appState) { $appState = try { $app.properties.status } catch { $null } }
        if (-not $appState) { $appState = 'Unknown' }

        $createdTime = try { $app.properties.createdTime } catch { $null }
        if (-not $createdTime) { $createdTime = try { $app.createdTime } catch { $null } }
        if (-not $createdTime) { $createdTime = '' }

        $lastModifiedTime = try { $app.properties.lastModifiedTime } catch { $null }
        if (-not $lastModifiedTime) { $lastModifiedTime = try { $app.lastModifiedTime } catch { $null } }
        if (-not $lastModifiedTime) { $lastModifiedTime = '' }
        
        # Skip apps with empty IDs
        if ([string]::IsNullOrWhiteSpace($appId)) { continue }
        
        # Extract additional properties from app.properties
        $description = try { $app.properties.description } catch { $null }; if (-not $description) { $description = '' }
        $appVersion = try { $app.properties.appVersion } catch { $null }; if (-not $appVersion) { $appVersion = '' }
        $sharedUsersCount = try { $app.properties.sharedUsersCount } catch { $null }; if ($null -eq $sharedUsersCount) { $sharedUsersCount = '0' }
        $sharedGroupsCount = try { $app.properties.sharedGroupsCount } catch { $null }; if ($null -eq $sharedGroupsCount) { $sharedGroupsCount = '0' }
        $isFeaturedApp = try { $app.properties.isFeaturedApp } catch { $null }; if ($null -eq $isFeaturedApp) { $isFeaturedApp = 'False' }
        $usesPremiumApi = try { $app.properties.usesPremiumApi } catch { $null }; if ($null -eq $usesPremiumApi) { $usesPremiumApi = 'False' }
        $usesCustomApi = try { $app.properties.usesCustomApi } catch { $null }; if ($null -eq $usesCustomApi) { $usesCustomApi = 'False' }
        $usesOnPremiseGateway = try { $app.properties.usesOnPremiseGateway } catch { $null }; if ($null -eq $usesOnPremiseGateway) { $usesOnPremiseGateway = 'False' }
        $isCustomizable = try { $app.properties.isCustomizable } catch { $null }; if ($null -eq $isCustomizable) { $isCustomizable = 'False' }
        $bypassConsent = try { $app.properties.bypassConsent } catch { $null }; if ($null -eq $bypassConsent) { $bypassConsent = 'False' }
        
        # Extract connection references (connectors used)
        $connectors = @()
        if ($app.properties.connectionReferences) {
            $connectors = @($app.properties.connectionReferences.PSObject.Properties.Name)
        }
        $connectorList = if ($connectors.Count -gt 0) { $connectors -join '; ' } else { '' }

        # Extract dataset URLs from connectionReferences.*.dataSets where URLs are often object keys.
        $datasetUrls = @()
        if ($app.properties.connectionReferences) {
            $datasetUrls = @(Get-UrlsFromObject -Object $app.properties.connectionReferences)
        }
        $datasetUrlList = if ($datasetUrls.Count -gt 0) {
            @($datasetUrls | Where-Object { $_ -is [string] -and $_ -match '^https?://' } | Select-Object -Unique) -join '; '
        }
        else {
            ''
        }
        
        $appDetailFetchStatus = try { [string]$app.PSObject.Properties['__AppDetailFetchStatus'].Value } catch { $null }
        if ([string]::IsNullOrWhiteSpace($appDetailFetchStatus)) { $appDetailFetchStatus = 'NotAttempted' }

        # Build CSV row with ALL meaningful fields from API response
        $csvRows += [PSCustomObject]@{
            'AppName'             = $appName
            'AppId'               = $appId
            'Owner'               = $owner
            'Environment'         = $EnvironmentName
            'EnvironmentId'       = $EnvironmentId
            'AppType'             = $appType
            'State'               = $appState
            'CreatedTime'         = $createdTime
            'LastModifiedTime'    = $lastModifiedTime
            'Description'         = $description
            'AppVersion'          = $appVersion
            'SharedUsers'         = $sharedUsersCount
            'SharedGroups'        = $sharedGroupsCount
            'IsFeatured'          = $isFeaturedApp
            'UsesPremium'         = $usesPremiumApi
            'UsesCustomConnector' = $usesCustomApi
            'UsesOnPremise'       = $usesOnPremiseGateway
            'IsCustomizable'      = $isCustomizable
            'BypassConsent'       = $bypassConsent
            'AppDetailFetchStatus' = $appDetailFetchStatus
            'Connectors'          = $connectorList
            'DatasetUrls'         = $datasetUrlList
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

    function Add-UniqueTextValue {
        param (
            [Parameter()] $Collection,
            [Parameter()] $Value
        )

        if ($null -eq $Collection) { return }
        if ($null -eq $Value) { return }
        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text)) { return }

        # Guard method calls so an empty/odd payload shape cannot break export.
        if ($Collection.PSObject.Methods['Contains'] -and $Collection.PSObject.Methods['Add']) {
            if (-not $Collection.Contains($text)) {
                $Collection.Add($text) | Out-Null
            }
        }
    }

    function Get-OptionalFlowProperty {
        param (
            [Parameter()] $InputObject,
            [Parameter(Mandatory)] [string] $PropertyPath
        )

        if ($null -eq $InputObject) { return $null }

        try {
            $value = $InputObject
            foreach ($segment in $PropertyPath -split '\.') {
                if ($null -eq $value) { return $null }
                $prop = $value.PSObject.Properties[$segment]
                if ($null -eq $prop) { return $null }
                $value = $prop.Value
            }
            return $value
        }
        catch {
            return $null
        }
    }
    
    foreach ($flow in $Flows) {
        $flowId = try { $flow.workflowId } catch { $null }
        if (-not $flowId) { $flowId = try { $flow.name } catch { $null } }
        if (-not $flowId) { $flowId = try { $flow.id } catch { $null } }
        if ([string]::IsNullOrWhiteSpace($flowId)) { continue }
        
        # Skip duplicate FlowIds (same flow may appear multiple times in API response)
        if ($flowId -in $seenFlowIds) {
            if ($debug) { Write-Host "    [DEBUG-EXPORT] Skipping duplicate FlowId: $flowId" -ForegroundColor DarkGray }
            continue
        }
        $seenFlowIds += $flowId
        
        # Extract all available properties from the flow object
        $flowName = try { $flow.properties.displayName } catch { $null }
        if (-not $flowName) { $flowName = try { $flow.displayName } catch { $null } }
        if (-not $flowName) { $flowName = try { $flow.name } catch { $null } }
        if (-not $flowName) { $flowName = '' }

        $flowType = try { [string]$flow.PSObject.Properties['__NormalizedFlowType'].Value } catch { $null }
        if ([string]::IsNullOrWhiteSpace($flowType)) {
            $flowType = Resolve-FlowType -Flow $flow
        }

        $flowState = try { $flow.properties.state } catch { $null }
        if (-not $flowState) { $flowState = try { $flow.state } catch { $null } }
        if (-not $flowState) { $flowState = try { $flow.stateCode } catch { $null } }
        if (-not $flowState) { $flowState = try { $flow.properties.status } catch { $null } }
        if (-not $flowState) { $flowState = try { $flow.statusCode } catch { $null } }
        if (-not $flowState) { $flowState = 'Unknown' }

        $ownerNames = [System.Collections.Generic.List[string]]::new()
        $ownerIds = [System.Collections.Generic.List[string]]::new()

        # Preferred human-readable owner hints
        Add-UniqueTextValue -Collection $ownerNames -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'properties.owner.displayName')
        Add-UniqueTextValue -Collection $ownerNames -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'owner.displayName')
        Add-UniqueTextValue -Collection $ownerNames -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'properties.createdBy.displayName')
        Add-UniqueTextValue -Collection $ownerNames -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'properties.creator.displayName')
        Add-UniqueTextValue -Collection $ownerNames -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'properties.creator.userPrincipalName')
        Add-UniqueTextValue -Collection $ownerNames -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'properties.creator.email')

        # Stable identity hints for tenants where display names are not returned
        Add-UniqueTextValue -Collection $ownerIds -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'ownerId')
        Add-UniqueTextValue -Collection $ownerIds -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'createdBy')
        Add-UniqueTextValue -Collection $ownerIds -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'properties.owner.id')
        Add-UniqueTextValue -Collection $ownerIds -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'properties.owner.objectId')
        Add-UniqueTextValue -Collection $ownerIds -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'properties.createdBy.id')
        Add-UniqueTextValue -Collection $ownerIds -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'properties.createdBy.objectId')
        Add-UniqueTextValue -Collection $ownerIds -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'properties.createdBy.userId')
        Add-UniqueTextValue -Collection $ownerIds -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'properties.creator.objectId')
        Add-UniqueTextValue -Collection $ownerIds -Value (Get-OptionalFlowProperty -InputObject $flow -PropertyPath 'properties.creator.userId')

        $ownerIdList = if ($ownerIds.Count -gt 0) { $ownerIds -join '; ' } else { '' }
        $flowOwner = if ($ownerNames.Count -gt 0) {
            $ownerNames -join '; '
        }
        elseif (-not [string]::IsNullOrWhiteSpace($ownerIdList)) {
            $ownerIdList
        }
        else {
            'Unknown'
        }

        $flowCreatedTime = try { $flow.properties.createdTime } catch { $null }
        if (-not $flowCreatedTime) { $flowCreatedTime = try { $flow.createdOn } catch { $null } }
        if (-not $flowCreatedTime) { $flowCreatedTime = try { $flow.createdTime } catch { $null } }
        if (-not $flowCreatedTime) { $flowCreatedTime = '' }

        $flowModifiedTime = try { $flow.properties.lastModifiedTime } catch { $null }
        if (-not $flowModifiedTime) { $flowModifiedTime = try { $flow.modifiedOn } catch { $null } }
        if (-not $flowModifiedTime) { $flowModifiedTime = try { $flow.lastModifiedTime } catch { $null } }
        if (-not $flowModifiedTime) { $flowModifiedTime = '' }
        $flowDefinitionUri = try { $flow.properties.definitionUri } catch { $null }; if (-not $flowDefinitionUri) { $flowDefinitionUri = '' }
        
        # Extract additional flow properties
        $templateName = try { $flow.properties.templateName } catch { $null }; if (-not $templateName) { $templateName = '' }
        $provisioningMethod = try { $flow.properties.provisioningMethod } catch { $null }; if (-not $provisioningMethod) { $provisioningMethod = '' }
        $flowPlan = try { $flow.properties.plan } catch { $null }; if (-not $flowPlan) { $flowPlan = '' }
        $flowFailureAlert = try { $flow.properties.flowFailureAlertSubscribed } catch { $null }; if ($null -eq $flowFailureAlert) { $flowFailureAlert = 'False' }
        $isManaged = try { $flow.properties.isManaged } catch { $null }; if ($null -eq $isManaged) { $isManaged = 'False' }
        $userType = try { $flow.properties.userType } catch { $null }; if (-not $userType) { $userType = '' }
        
        # Extract connection references (connectors used in flow)
        $flowConnectors = @()
        if ($flow.properties.connectionReferences) {
            $flowConnectors = @($flow.properties.connectionReferences.PSObject.Properties | Select-Object -ExpandProperty Name)
        }
        $embeddedRefs = @(Get-EmbeddedFlowConnectionReferences -Flow $flow)
        if ($embeddedRefs.Count -gt 0) {
            $flowConnectors += @($embeddedRefs | ForEach-Object {
                    if (-not [string]::IsNullOrWhiteSpace($_.ApiName)) {
                        "{0} ({1})" -f $_.ReferenceName, $_.ApiName
                    }
                    else {
                        $_.ReferenceName
                    }
                })
        }
        if ($flowConnectors.Count -eq 0) {
            $embeddedDef = Get-EmbeddedFlowDefinition -Flow $flow
            if ($embeddedDef -and $embeddedDef.parameters -and $embeddedDef.parameters.'$connections' -and $embeddedDef.parameters.'$connections'.defaultValue) {
                $flowConnectors = @($embeddedDef.parameters.'$connections'.defaultValue.PSObject.Properties.Name)
            }
        }
        $flowConnectors = @($flowConnectors | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
        $flowConnectorList = if ($flowConnectors.Count -gt 0) { $flowConnectors -join '; ' } else { '' }
        
        # Get URLs for this flow (extracted during flow processing)
        $flowUrlList = if ($FlowUrls.ContainsKey($flowId) -and $FlowUrls[$flowId].Count -gt 0) { 
            # Filter to only actual URLs (not operation details)
            $actualUrls = @($FlowUrls[$flowId] | Where-Object { $_ -match 'https?://' })
            $sharePointUrls = @($actualUrls | Where-Object { $_ -match 'https?://[^\s]*sharepoint\.com' } | Select-Object -Unique)
            if ($sharePointUrls.Count -gt 0) { $sharePointUrls -join '; ' }
            elseif ($actualUrls.Count -gt 0) { @($actualUrls | Select-Object -Unique) -join '; ' }
            else { '' }
        }
        else { 
            # Fallback: scan embedded definition text for explicit SharePoint URLs.
            $embeddedDefText = try { [string]$flow.definition } catch { '' }
            if (-not [string]::IsNullOrWhiteSpace($embeddedDefText)) {
                $urlMatches = [regex]::Matches($embeddedDefText, 'https?://[^\"''\s,]+')
                $rawUrls = @($urlMatches | ForEach-Object { $_.Value } | Select-Object -Unique)
                $spUrls = @($rawUrls | Where-Object { $_ -match 'https?://[^\s]*sharepoint\.com' } | Select-Object -Unique)
                if ($spUrls.Count -gt 0) { $spUrls -join '; ' }
                elseif ($rawUrls.Count -gt 0) { $rawUrls -join '; ' }
                else { '' }
            }
            else {
                ''
            }
        }
        
        # Build CSV row with ALL available flow properties
        $csvRows += [PSCustomObject]@{
            'FlowName'           = $flowName
            'FlowId'             = $flowId
            'FlowType'           = $flowType
            'FlowState'          = $flowState
            'Owner'              = $flowOwner
            'OwnerIds'           = $ownerIdList
            'Environment'        = $EnvironmentName
            'EnvironmentId'      = $EnvironmentId
            'CreatedTime'        = $flowCreatedTime
            'LastModifiedTime'   = $flowModifiedTime
            'TemplateName'       = $templateName
            'ProvisioningMethod' = $provisioningMethod
            'Plan'               = $flowPlan
            'UserType'           = $userType
            'FlowFailureAlert'   = $flowFailureAlert
            'IsManaged'          = $isManaged
            'Connectors'         = $flowConnectorList
            'URLs'               = $flowUrlList
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
    # Normalize collection scope for case-insensitive input.
    $collectionScope = if ($null -ne $collectionScope) { $collectionScope.ToString().Trim() } else { '' }
    switch ($collectionScope.ToLowerInvariant()) {
        'flows' { $collectionScope = 'Flows'; break }
        'apps' { $collectionScope = 'Apps'; break }
        'both' { $collectionScope = 'Both'; break }
    }

    # Validate collection scope parameter
    if ($collectionScope -notin @('Flows', 'Apps', 'Both')) {
        Write-Host "Error: Invalid collection scope '$collectionScope'" -ForegroundColor Red
        Write-Host "Valid values are: 'Flows', 'Apps', or 'Both'" -ForegroundColor Red
        exit 1
    }

    if ($ProgressUpdateEvery -lt 1) {
        Write-Host "ProgressUpdateEvery must be 1 or greater. Using 25." -ForegroundColor Yellow
        $ProgressUpdateEvery = 25
    }
    
    # Clear any cached tokens to force fresh authentication
    Write-Host "Initializing authentication for Power Platform..." -ForegroundColor Cyan
    Write-Host "Collection scope: $collectionScope" -ForegroundColor Yellow
    $global:tokenPP = $null
    $global:tokenPPExpiry = $null
    $global:refreshTokenPP = $null

    # Single sign-in for both Flows and Apps.
    Get-TokenWithPKCE -Scope $scope -TokenType 'Flows'

    Write-Host "`n=========================================" -ForegroundColor Cyan
    Write-Host "Power Automate & Power Apps Report" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
    Write-Host "=========================================`n" -ForegroundColor Cyan

    # Initialize streaming export targets and running totals.
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $appsOutputPath = if ($collectionScope -in @('Apps', 'Both')) { Join-Path $OutputFolder "PowerApps_Report_$timestamp.csv" } else { $null }
    $flowsOutputPath = if ($collectionScope -in @('Flows', 'Both')) { Join-Path $OutputFolder "PowerAutomate_Flows_Report_$timestamp.csv" } else { $null }
    $totalAppsExported = 0
    $totalFlowsExported = 0

    if ($appsOutputPath) {
        Write-Host "Power Apps output file: $appsOutputPath" -ForegroundColor DarkGray
    }
    if ($flowsOutputPath) {
        Write-Host "Power Automate flows output file: $flowsOutputPath" -ForegroundColor DarkGray
    }
    Write-Host ""

    # Get all environments
    if ($collectionScope -eq 'Apps') {
        Write-Host "Retrieving Power Apps environments..." -ForegroundColor Yellow
        $environments = Get-PowerAppsEnvironments
    }
    else {
        Write-Host "Retrieving Power Automate environments..." -ForegroundColor Yellow
        $environments = Get-PowerAutomateEnvironments
    }

    $environments = @($environments)
    $environments = @(Select-Environments -Environments $environments -IncludeEnvironment $Environment -ExcludeEnvironment $ExcludeEnvironment)

    if ($environments.Count -eq 0) {
        if (@($Environment).Count -gt 0 -or @($ExcludeEnvironment).Count -gt 0) {
            Write-Host "No environments matched the specified include/exclude filters." -ForegroundColor Yellow
        }
        else {
            Write-Host "No environments found." -ForegroundColor Yellow
        }
        exit
    }

    Write-Host "Found $($environments.Count) environment(s):`n" -ForegroundColor Green

    if (@($Environment).Count -gt 0) {
        Write-Host "Environment filter: $($Environment -join ', ')" -ForegroundColor DarkGray
    }
    if (@($ExcludeEnvironment).Count -gt 0) {
        Write-Host "Excluded environments: $($ExcludeEnvironment -join ', ')" -ForegroundColor DarkGray
    }
    if (@($Environment).Count -gt 0 -or @($ExcludeEnvironment).Count -gt 0) {
        Write-Host "" 
    }

    $explicitFlowTargets = @(Get-FlowTargetsFromUrls -FlowUrls $IncludeFlowUrls)
    if ($explicitFlowTargets.Count -gt 0 -and $debug) {
        Write-Host "Explicit flow include targets: $($explicitFlowTargets.Count)" -ForegroundColor DarkGray
    }

    # Process each environment
    foreach ($env in $environments) {
        $envName = $env.properties.displayName
        $envId = $env.name
        $flowUrlMap = @{}
        $flowObjectDumped = $false
        
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Magenta
        Write-Host "ENVIRONMENT: $envName" -ForegroundColor Magenta
        Write-Host "ID: $envId" -ForegroundColor DarkGray
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor Magenta

        # Get flows in this environment (cloud flows + desktop flows) based on collection scope
        $cloudFlows = @()
        $desktopFlows = @()
        if ($collectionScope -in @('Flows', 'Both')) {
            Write-Host "  Retrieving cloud flows..." -ForegroundColor Yellow
            $cloudFlows = @(Get-PowerAutomateFlows -EnvironmentId $envId)
            
            Write-Host "  Retrieving desktop flows..." -ForegroundColor Yellow
            $desktopFlows = @(Get-PowerAutomateDesktopFlows -EnvironmentId $envId)
        }
        
        $powerApps = @()
        if ($collectionScope -in @('Apps', 'Both')) {
            Write-Host "  Retrieving Power Apps..." -ForegroundColor Yellow
            $powerApps = @(Get-PowerApps -EnvironmentId $envId)
        }
        
        $allFlows = @()
        if ($cloudFlows) { $allFlows += $cloudFlows }
        if ($desktopFlows) { $allFlows += $desktopFlows }

        if ($collectionScope -in @('Flows', 'Both') -and $explicitFlowTargets.Count -gt 0) {
            $targetsForEnvironment = @($explicitFlowTargets | Where-Object { $_.EnvironmentId -eq $envId })
            foreach ($target in $targetsForEnvironment) {
                $alreadyPresent = $false
                foreach ($existingFlow in $allFlows) {
                    $existingIds = @()

                    $candidateId = $null
                    try { $candidateId = [string]$existingFlow.workflowId } catch {}
                    if (-not [string]::IsNullOrWhiteSpace($candidateId)) { $existingIds += $candidateId }

                    $candidateId = $null
                    try { $candidateId = [string]$existingFlow.name } catch {}
                    if (-not [string]::IsNullOrWhiteSpace($candidateId)) { $existingIds += $candidateId }

                    $candidateId = $null
                    try { $candidateId = [string]$existingFlow.id } catch {}
                    if (-not [string]::IsNullOrWhiteSpace($candidateId)) { $existingIds += $candidateId }

                    if ($existingIds -contains $target.FlowId) {
                        $alreadyPresent = $true
                        break
                    }
                }

                if (-not $alreadyPresent) {
                    $directFlow = Get-PowerAutomateFlowById -EnvironmentId $envId -FlowId $target.FlowId
                    if ($directFlow) {
                        $allFlows += $directFlow
                        $cloudFlows += $directFlow
                        Write-Host "  Included explicit flow by URL target: $($target.FlowId)" -ForegroundColor Green
                    }
                    elseif ($debug) {
                        Write-Host "  [DEBUG] Explicit flow target not found via API: $($target.FlowId)" -ForegroundColor DarkGray
                    }
                }
            }
        }

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
        if ($collectionScope -in @('Flows', 'Both')) {
            Write-Host "  Retrieving connector information..." -ForegroundColor Yellow
            $connectors = @(Get-ConnectorStatus -EnvironmentId $envId)
            if ($connectors.Count -gt 0) {
                Write-Host "  Found $($connectors.Count) connector(s)`n" -ForegroundColor Green
            }
            else {
                Write-Host "  (Connector details unavailable for this environment)`n" -ForegroundColor DarkGray
            }
        }
        else {
            Write-Host "  Skipping connector status for Apps-only scope.`n" -ForegroundColor DarkGray
        }

        # Process each flow (only if collecting flows)
        if ($collectionScope -in @('Flows', 'Both')) {
            $flowTotal = $allFlows.Count
            $flowProcessed = 0
            foreach ($flow in $allFlows) {
                $flowProcessed++
                $flowName = try { $flow.properties.displayName } catch { $null }
                if (-not $flowName) { $flowName = try { $flow.displayName } catch { $null } }
                if (-not $flowName) { $flowName = try { $flow.name } catch { $null } }
                if (-not $flowName) { $flowName = '(Unknown)' }

                $flowId = try { $flow.workflowId } catch { $null }
                if (-not $flowId) { $flowId = try { $flow.name } catch { $null } }
                if (-not $flowId) { $flowId = try { $flow.id } catch { $null } }

                $flowIdCandidates = @(Get-FlowIdCandidates -Flow $flow)
                if ($flowIdCandidates.Count -gt 0) {
                    $flowId = $flowIdCandidates[0]
                }

                $flowState = try { $flow.properties.state } catch { $null }
                if (-not $flowState) { $flowState = try { $flow.state } catch { $null } }
                if (-not $flowState) { $flowState = try { $flow.properties.status } catch { $null } }
                if (-not $flowState) { $flowState = 'Unknown' }

                $flowType = Resolve-FlowType -Flow $flow
            
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
            
                $flowDef = Get-EmbeddedFlowDefinition -Flow $flow
                if (-not $flowDef) {
                    $flowDef = Get-FlowDetails -EnvironmentId $envId -FlowIds $flowIdCandidates
                }

                $flowType = Resolve-FlowType -Flow $flow -FlowDefinition $flowDef
                if ($flow.PSObject.Properties['__NormalizedFlowType']) {
                    $flow.PSObject.Properties['__NormalizedFlowType'].Value = $flowType
                }
                else {
                    $flow | Add-Member -NotePropertyName '__NormalizedFlowType' -NotePropertyValue $flowType
                }

                Write-Host "    │  Normalized Type: $flowType" -ForegroundColor DarkGray

                if ($flowDef) {
                    # Extract HTTP connectors and URLs
                    $extracted = Get-HttpConnectorsAndUrlsFromFlow -FlowDefinition $flowDef
                    $allActions = $extracted.AllActions
                    $httpConnectors = $extracted.HttpConnectors
                    $urls = $extracted.Urls
                    $flowUrlMap[$flowId] = $urls

                    if ($DumpOneFlowObjectPerEnvironment -and -not $flowObjectDumped) {
                        try {
                            $debugFlowPath = Join-Path $OutputFolder ("FlowObjectDebug_{0}_{1}.json" -f ($envId -replace '[^a-zA-Z0-9_-]', '_'), (Get-Date -Format 'yyyyMMdd_HHmmss'))
                            $debugFlowObject = [PSCustomObject]@{
                                EnvironmentName      = $envName
                                EnvironmentId        = $envId
                                FlowName             = $flowName
                                FlowId               = $flowId
                                FlowIdCandidates     = $flowIdCandidates
                                ExtractedUrls        = @($urls)
                                SharePointUrls       = @($urls | Where-Object { $_ -is [string] -and $_ -match 'https?://[^\s]*sharepoint\.com' } | Select-Object -Unique)
                                RawFlow              = $flow
                                ParsedFlowDefinition = $flowDef
                            }
                            $debugFlowObject | ConvertTo-Json -Depth 100 | Out-File -FilePath $debugFlowPath -Encoding utf8
                            Write-Host "    [DEBUG] Wrote flow debug object: $debugFlowPath" -ForegroundColor DarkGray
                            $flowObjectDumped = $true
                        }
                        catch {
                            if ($debug) { Write-Host "    [DEBUG] Failed to write flow debug object: $($_.Exception.Message)" -ForegroundColor DarkGray }
                        }
                    }

                    # Display ALL Actions
                    if ($allActions.Count -gt 0) {
                        Write-Host "    │" -ForegroundColor Cyan
                        Write-Host "    ├─ All Actions & Triggers ($($allActions.Count)):" -ForegroundColor Yellow
                        foreach ($action in $allActions) {
                            $actionType = $action.Type
                            $actionKind = $action.Kind
                            $connName = $action.ConnectorName
                            $actionName = $action.Name
                            $isHttp = if ($action.ContainsKey('IsHttp') -and $action.IsHttp) { ' [HTTP]' } else { '' }
                        
                            Write-Host "    │  ├─ $actionName$isHttp" -ForegroundColor White
                            Write-Host "    │  │  └─ Type: $actionType | Kind: $actionKind | Connector: $connName" -ForegroundColor DarkGray
                        
                            if ($action.ContainsKey('Method') -and $action.Method) {
                                Write-Host "    │  │     Method: $($action.Method)" -ForegroundColor DarkGray
                            }
                            if ($action.ContainsKey('DependsOn') -and $action.DependsOn) {
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
                            if ($connector.ContainsKey('Method') -and $connector.Method) {
                                Write-Host "    │  │     Method: $($connector.Method)" -ForegroundColor White
                            }
                            if ($connector.ContainsKey('DependsOn') -and $connector.DependsOn) {
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

                if (($flowProcessed % $ProgressUpdateEvery -eq 0) -or ($flowProcessed -eq $flowTotal)) {
                    Write-Host "  Progress (Flows): $flowProcessed/$flowTotal processed in '$envName'" -ForegroundColor DarkCyan
                }
            }
        }

        # Process Power Apps (only if collecting apps)
        if ($collectionScope -in @('Apps', 'Both')) {
            if ($powerApps.Count -gt 0) {
                Write-Host "  Power Apps:" -ForegroundColor Yellow
                $appTotal = $powerApps.Count
                $appProcessed = 0
                foreach ($app in $powerApps) {
                    $appProcessed++
                    # Handle different property names from different API endpoints
                    $appName = try { $app.displayName } catch { $null }
                    if (-not $appName) { $appName = try { $app.properties.displayName } catch { $null } }
                    if (-not $appName) { $appName = try { $app.name } catch { $null } }
                    if (-not $appName) { $appName = '(Unknown)' }

                    $appId = try { $app.appId } catch { $null }
                    if (-not $appId) { $appId = try { $app.properties.appId } catch { $null } }
                    if (-not $appId) { $appId = try { $app.name } catch { $null } }
                    if (-not $appId) { $appId = try { $app.id } catch { $null } }
                    if (-not $appId) { $appId = '' }
                
                    # Skip apps with empty IDs
                    if ([string]::IsNullOrWhiteSpace($appId)) {
                        if ($debug) { Write-Host "    [DEBUG] Skipping app with empty ID: $appName" -ForegroundColor Yellow }
                        continue
                    }
                
                    $appType = try { $app.appType } catch { $null }
                    if (-not $appType) { $appType = try { $app.properties.appType } catch { $null } }
                    if (-not $appType) { $appType = 'Unknown' }

                    $appState = try { $app.state } catch { $null }
                    if (-not $appState) { $appState = try { $app.properties.state } catch { $null } }
                    if (-not $appState) { $appState = try { $app.properties.status } catch { $null } }
                    if (-not $appState) { $appState = 'Unknown' }
                
                    Write-Host "    ┌─ App: $appName" -ForegroundColor Magenta
                    Write-Host "    │  ID: $appId" -ForegroundColor DarkGray
                    Write-Host "    │  Type: $appType | State: $appState" -ForegroundColor DarkGray
                
                    # Try to get app details
                    $appDetailResult = Get-PowerAppDetails -AppId $appId -EnvironmentId $envId
                    $appDetailFetchStatus = if ($appDetailResult -and $appDetailResult.Status) { [string]$appDetailResult.Status } else { 'Unknown' }
                    if ($app.PSObject.Properties['__AppDetailFetchStatus']) {
                        $app.PSObject.Properties['__AppDetailFetchStatus'].Value = $appDetailFetchStatus
                    }
                    else {
                        $app | Add-Member -NotePropertyName '__AppDetailFetchStatus' -NotePropertyValue $appDetailFetchStatus
                    }

                    $appDef = if ($appDetailResult) { $appDetailResult.Definition } else { $null }

                    if ($appDef) {
                        $appConnectorInfo = Get-AppConnectorsFromDefinition -AppDefinition $appDef
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
                    else {
                        Write-Host "    │  App detail status: $appDetailFetchStatus" -ForegroundColor DarkYellow
                    }
                
                    Write-Host "    │" -ForegroundColor Magenta
                    Write-Host "    └─────────────────────────────────────`n" -ForegroundColor Magenta

                    if (($appProcessed % $ProgressUpdateEvery -eq 0) -or ($appProcessed -eq $appTotal)) {
                        Write-Host "  Progress (Apps): $appProcessed/$appTotal processed in '$envName'" -ForegroundColor DarkMagenta
                    }
                }
            }
        }

        # Collect data for CSV export
        Write-Host "  Collecting data for export..." -ForegroundColor Yellow
        
        # Collect Power Apps data (only if collecting apps)
        if ($collectionScope -in @('Apps', 'Both')) {
            if ($powerApps.Count -gt 0) {
                if ($debug) { Write-Host "    [DEBUG-EXPORT] Processing $($powerApps.Count) Power Apps" -ForegroundColor DarkGray }
                $appsExport = Export-PowerAppsToCSV -Apps $powerApps -EnvironmentName $envName -EnvironmentId $envId
                if ($appsExport.Count -gt 0) {
                    if (Test-Path -Path $appsOutputPath) {
                        $appsExport | Export-Csv -Path $appsOutputPath -NoTypeInformation -Encoding UTF8 -Append
                    }
                    else {
                        $appsExport | Export-Csv -Path $appsOutputPath -NoTypeInformation -Encoding UTF8
                    }
                    $totalAppsExported += $appsExport.Count
                }
                if ($debug) { Write-Host "    [DEBUG-EXPORT] Exported $($appsExport.Count) Power App rows" -ForegroundColor DarkGray }
            }
        }
        
        # Collect Power Automate flows data with URLs (only if collecting flows)
        if ($collectionScope -in @('Flows', 'Both')) {
            if ($allFlows.Count -gt 0) {
                if ($debug) { Write-Host "    [DEBUG-EXPORT] Processing $($allFlows.Count) flows" -ForegroundColor DarkGray }

                $flowsExport = Export-PowerAutomateFlowsToCSV -Flows $allFlows -EnvironmentName $envName -EnvironmentId $envId -FlowUrls $flowUrlMap
                if ($flowsExport.Count -gt 0) {
                    if (Test-Path -Path $flowsOutputPath) {
                        $flowsExport | Export-Csv -Path $flowsOutputPath -NoTypeInformation -Encoding UTF8 -Append
                    }
                    else {
                        $flowsExport | Export-Csv -Path $flowsOutputPath -NoTypeInformation -Encoding UTF8
                    }
                    $totalFlowsExported += $flowsExport.Count
                }
                if ($debug) { Write-Host "    [DEBUG-EXPORT] Exported $($flowsExport.Count) flow action rows" -ForegroundColor DarkGray }
                if ($debug) { 
                    foreach ($row in $flowsExport | Select-Object -First 3) {
                        Write-Host "      Row: Name=$($row.FlowName), URLs=$($row.URLs)" -ForegroundColor DarkGray
                    }
                }
            }
        }

        Write-Host "  Environment complete. Current totals written to disk: Apps=$totalAppsExported, Flows=$totalFlowsExported" -ForegroundColor Green

        Write-Host "`n" -ForegroundColor Magenta
    }

    # Final export summary
    Write-Host "`n=========================================" -ForegroundColor Green
    Write-Host "Export complete" -ForegroundColor Yellow
    Write-Host "=========================================`n" -ForegroundColor Green

    # Power Apps summary (only if collecting apps)
    if ($collectionScope -in @('Apps', 'Both')) {
        if ($totalAppsExported -gt 0) {
            Write-Host "Power Apps report exported to: $appsOutputPath" -ForegroundColor Green
            Write-Host "  Total Power Apps exported: $totalAppsExported rows" -ForegroundColor Cyan
        }
        else {
            Write-Host "  No Power Apps to export." -ForegroundColor DarkGray
        }
    }
    
    # Power Automate summary (only if collecting flows)
    if ($collectionScope -in @('Flows', 'Both')) {
        if ($totalFlowsExported -gt 0) {
            Write-Host "Power Automate flows report exported to: $flowsOutputPath" -ForegroundColor Green
            Write-Host "  Total flows exported: $totalFlowsExported rows" -ForegroundColor Cyan
        }
        else {
            Write-Host "  No Power Automate flows to export." -ForegroundColor DarkGray
        }
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
