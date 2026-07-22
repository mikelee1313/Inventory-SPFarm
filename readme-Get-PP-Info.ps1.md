## Prerequisites

- PowerShell 5.1 or later
- SharePoint Management Shell (for on-prem environments)
- Appropriate SharePoint farm administrative permissions
- Update variables (such as SharePoint URLs or output paths) inside each script as needed

## Get-PP-Info.ps1 Power Platform Inventory

`Get-PP-Info.ps1` scans Power Platform resources that the signed-in user can access. It uses OAuth 2.0 authorization code with PKCE and delegated permissions, which means the Entra app registration acts as the signed-in user. The app permissions and the user's Power Platform/environment permissions are both required.

### What the script does

- Opens an interactive browser sign-in and requests `https://api.powerplatform.com/.default offline_access`.
- Retrieves Power Platform environments visible to the signed-in user.
- Supports include and exclude filters by environment display name or environment ID.
- Collects cloud flows, desktop flows, and legacy `Microsoft.ProcessSimple` flows.
- Collects Power Apps metadata and app connector references.
- Retrieves flow and app definitions where the signed-in user has enough access.
- Extracts connector references, HTTP/API actions, SharePoint dataset URLs, and other endpoint URLs from flow and app definitions.
- Normalizes flow types into readable categories such as `Cloud - Automated`, `Cloud - Instant`, `Cloud - Scheduled`, `Cloud - Other`, and `Desktop - RPA`.
- Deduplicates flows returned by multiple modern and legacy endpoints.
- Supports force-including specific flows from maker portal URLs with `$IncludeFlowUrls`.
- Optionally writes diagnostic JSON payload samples when debug dump settings are enabled.
- Exports timestamped CSV reports to `$OutputFolder`.

### Output files

When `$collectionScope` includes apps, the script writes:

- `PowerApps_Report_yyyyMMdd_HHmmss.csv`

When `$collectionScope` includes flows, the script writes:

- `PowerAutomate_Flows_Report_yyyyMMdd_HHmmss.csv`

### Script configuration

Before running the script, update the configuration section in `Get-PP-Info.ps1`:

- `$tenantId` - Entra tenant ID or verified tenant domain.
- `$clientId` - Application/client ID of the Entra app registration.
- `$redirectUri` - OAuth callback URI. The default is `http://localhost:8081`.
- `$collectionScope` - `Flows`, `Apps`, or `Both`.
- `$Environment` - Optional list of environment names or IDs to include.
- `$ExcludeEnvironment` - Optional list of environment names or IDs to exclude.
- `$IncludeFlowUrls` - Optional list of direct Power Automate maker portal flow URLs to force include.
- `$OutputFolder` - Folder for CSV and optional debug output. The default is `$env:TEMP`.
- `$EnableLegacyFlowDiscoveryFallback` - Enables fallback to legacy Flow/ProcessSimple endpoints.
- `$LegacyFlowScopes` - Legacy Flow scopes used for token exchange when fallback endpoints are needed.
- `$DumpFlowPayloadSamples` and `$DumpOneFlowObjectPerEnvironment` - Optional debug output toggles.

### Entra app registration setup

Create or update an Entra app registration for delegated interactive authentication.

- Supported account type: usually single tenant.
- Platform: `Mobile and desktop applications`.
- Redirect URI: `http://localhost:8081`.
- Client secret or certificate: not required for this script.
- Authentication flow: OAuth 2.0 authorization code with PKCE.
- Admin consent: recommended, and usually required for tenant-wide/admin usage.

The redirect URI in the app registration must exactly match `$redirectUri` in the script. The local machine must be able to open the browser sign-in page and listen for the callback on the redirect port.

### Delegated API permissions

Add the following delegated permissions to the app registration and grant admin consent.

#### Power Platform API

These are the current Power Platform API delegated permissions used by the modern endpoints:

| Permission | Purpose |
| --- | --- |
| `EnvironmentManagement.Environments.Read` | Read Power Platform environments. |
| `PowerAutomate.Flows.Read` | Read Power Automate cloud and desktop flows. |
| `PowerApps.Apps.Read` | Read Power Apps metadata and definitions. |
| `Connectivity.Connectors.Read` | Read connector metadata. |

#### Microsoft Flow Service

These permissions support legacy Microsoft Flow and `Microsoft.ProcessSimple` endpoints used by the script's fallback logic:

| Permission shown in portal | Purpose |
| --- | --- |
| `User` - Access Microsoft Flow as signed-in user | Allows delegated access to Microsoft Flow as the signed-in user. This is the portal-visible legacy permission that covers the older `ProcessSimple.Environment.Read` style access. |
| `Flows.Read.All` | Read flows visible to the signed-in user/admin. |
| `Flows.Read.Plans` | Read flow plan metadata. |

The older script comments and some legacy documentation may reference permissions such as `ProcessSimple.Environment.Read`, `Flow.Read`, or `PowerApps.ReadAll`. In many tenants, these are now surfaced in the portal under Microsoft Flow Service or Power Apps Service with friendlier names rather than the exact legacy scope names.

#### Power Apps Service

If `$collectionScope` is `Apps` or `Both`, also add delegated Power Apps Service permissions when they are available in the tenant UI:

| Permission shown in portal | Purpose |
| --- | --- |
| `User` - Access Power Apps as signed-in user | Allows delegated access to Power Apps as the signed-in user. |
| Power Apps delegated read permission, if available | Reads Power Apps and app metadata through legacy Power Apps endpoints. |

#### Microsoft Graph

Add this delegated Microsoft Graph permission:

| Permission | Purpose |
| --- | --- |
| `offline_access` | Allows the token response to include a refresh token so the script can silently refresh access tokens and exchange for legacy Flow scopes. |

To confirm `offline_access` is configured:

1. Open Microsoft Entra ID.
2. Go to App registrations.
3. Open the app registration used by the script.
4. Select API permissions.
5. Select Add a permission.
6. Select Microsoft Graph.
7. Select Delegated permissions.
8. Search for `offline_access`.
9. Check `offline_access` - Maintain access to data you have given it access to.
10. Select Add permissions.
11. Select Grant admin consent.

The script already requests `offline_access` in this scope value:

```powershell
$scope = 'https://api.powerplatform.com/.default offline_access'
```

After successful sign-in, the OAuth token response should include a `refresh_token`. If no refresh token is returned, modern calls may still work until the access token expires, but silent refresh and legacy Flow fallback token exchange will not work.

### Signed-in user requirements

Because the script uses delegated permissions, the signed-in user controls what data can be inventoried.

For limited inventory, the signed-in user needs:

- Power Platform access/license.
- Access to each target environment.
- Owner, co-owner, shared user, or admin access to the target apps and flows.

For full inventory of an environment, the signed-in user should have:

- Environment Admin, System Administrator, or equivalent administrative access in the target environment.
- Permission to read flow and app definitions.

For tenant-wide inventory, the signed-in user should have one of the following tenant-level roles:

- Power Platform Administrator.
- Dynamics 365 Administrator.
- Global Administrator.

The user must also be allowed into any environments restricted by security groups.

### Manual Power Platform environment setup

This script does not require a service principal application user because it does not use app-only/client credentials authentication. It signs in interactively and acts as the user.

Manual setup that may still be required:

- Add the signed-in user to any environment security group.
- Assign the signed-in user the appropriate environment role.
- Make the signed-in user owner, co-owner, or admin of flows and apps where full definitions are required.
- Grant admin consent for the app registration API permissions.
- If the Enterprise Application has `User assignment required` enabled, assign the user or a group containing the user to the enterprise application.

If the script is rewritten to use service principal/app-only authentication, additional Power Platform setup is required. Register the app with Power Platform, for example with `pac admin create-service-principal --environment <environment id>`, then assign the appropriate Power Platform RBAC/environment access. That setup is not required for the delegated script as written.

### How to run

Run the script from PowerShell after updating the configuration values:

```powershell
.\Get-PP-Info.ps1
```

The script opens a browser window for sign-in, scans the environments and resources visible to the signed-in user, and writes CSV reports to `$OutputFolder`.
