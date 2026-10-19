<#
.Disclaimer: The sample scripts are provided AS IS without warranty of any kind. 
    Microsoft further disclaims all implied warranties including, without limitation, 
    any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
    In no event shall Microsoft, its authors, or anyone else involved in the creation, 
    production, or delivery of the scripts be liable for any damages whatsoever 
    (including, without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use of or inability 
    to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.
.SYNOPSIS
    Inventories Entra ID app registrations that use the "Sites.Selected" permission
    (Microsoft Graph and/or SharePoint Online) and reports every SharePoint site each
    app has been granted access to.

.DESCRIPTION
    1. Connects to Microsoft Graph with APP-ONLY certificate authentication using the
       TenantId, ClientId and CertificateThumbprint variables (CONFIGURATION section).
    2. Resolves the Sites.Selected permission IDs dynamically from the Microsoft Graph and
       Office 365 SharePoint Online service principals in the tenant.
    3. Finds every app registration whose requiredResourceAccess (API permissions) includes
       Sites.Selected, and checks whether admin consent has actually been granted.
       Optionally (-IncludeEnterpriseApps) also finds third-party / multi-tenant enterprise
       apps (service principals) that hold a consented Sites.Selected app role.
    4. Enumerates all SharePoint sites (/sites/getAllSites) and reads each site's
       application permission grants (/sites/{id}/permissions).
    5. Matches site grants to the apps found in step 3 and exports a CSV with:
       AppDisplayName, AppId, AppObjectId, PermissionSource, PermissionType, AdminConsented,
       SiteUrl, SiteId, SiteName, Roles, SitePermissionId.
       Apps with Sites.Selected but no site grants get one row with SiteUrl = "<no site grants found>".
    6. Exports a second CSV listing any sites whose permissions could not be read.

.REQUIRED APP-ONLY PERMISSIONS (on the app registration that RUNS this script)
    API: Microsoft Graph -> Application permissions
      - Application.Read.All     Read app registrations, service principals and their
                                 app role assignments (consent status).
      - Sites.FullControl.All    Required to read /sites/{id}/permissions (Sites.Selected
                                 grants). This is the least-privileged permission that
                                 supports listing site permissions app-only. It also
                                 covers enumerating all sites via /sites/getAllSites.
    Admin consent must be granted for both permissions.
    Credential: a certificate uploaded to the app registration (private key on the
    machine running the script).
    See the README section at the bottom of this file for setup steps.

.PARAMETER TenantId
    Optional. Overrides $ConfigTenantId in the CONFIGURATION section.

.PARAMETER ClientId
    Optional. Overrides $ConfigClientId in the CONFIGURATION section.

.PARAMETER CertificateThumbprint
    Optional. Overrides $ConfigCertificateThumbprint in the CONFIGURATION section.
    The certificate (with private key) must be in CurrentUser\My or LocalMachine\My.

.PARAMETER Environment
    Graph cloud: Global (default), USGov (GCC High), USGovDoD, China.

.PARAMETER OutputPath
    Path of the result CSV. Default: .\SitesSelected-AppGrants-<timestamp>.csv

.PARAMETER IncludeOneDrive
    Also scan OneDrive (personal) sites. Off by default to reduce run time.

.PARAMETER IncludeEnterpriseApps
    Also include service principals (e.g., third-party / multi-tenant apps not registered
    in this tenant) that have a consented Sites.Selected application role.

.PARAMETER SiteUrlFilter
    Optional. One or more wildcard patterns (string array) to limit which sites are
    scanned; a site is included if it matches ANY pattern. Overrides $ConfigSiteUrlFilter.
    e.g. -SiteUrlFilter 'https://contoso.sharepoint.com/sites/Finance','*/sites/HR*'

.EXAMPLE
    # 1) Edit the CONFIGURATION section ($ConfigTenantId, $ConfigClientId,
    #    $ConfigCertificateThumbprint), then simply run:
    .\Get-SitesSelectedAppGrants.ps1

.EXAMPLE
    # 2) Or supply/override the values at run time:
    .\Get-SitesSelectedAppGrants.ps1 -TenantId contoso.onmicrosoft.com `
        -ClientId 11111111-2222-3333-4444-555555555555 `
        -CertificateThumbprint A1B2C3D4E5F6... -IncludeEnterpriseApps

.NOTES
    Requires PowerShell 7.2+ (Windows PowerShell 5.1 also works) and:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
    Only the Authentication module is needed; all calls use Invoke-MgGraphRequest.
    Read-only: the script makes no changes to the tenant.
#>

[CmdletBinding()]
param(
    # Optional overrides. If omitted, the values in the CONFIGURATION section below are used.
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,

    [ValidateSet('Global', 'USGov', 'USGovDoD', 'China')]
    [string] $Environment = 'Global',

    [string] $OutputPath = (Join-Path (Get-Location) ("SitesSelected-AppGrants-{0:yyyyMMdd-HHmmss}.csv" -f (Get-Date))),

    [switch] $IncludeOneDrive,
    [switch] $IncludeEnterpriseApps,
    [string[]] $SiteUrlFilter
)

# =============================================================================
#  CONFIGURATION - set your tenant connection values here
# =============================================================================
$ConfigTenantId              = '<your-tenant-id-or-contoso.onmicrosoft.com>'   # Directory (tenant) ID
$ConfigClientId              = '<your-app-registration-client-id>'             # Application (client) ID
$ConfigCertificateThumbprint = '<your-certificate-thumbprint>'                 # Cert in CurrentUser\My or LocalMachine\My

# Optional: limit the scan to specific sites. Leave as @() to scan ALL sites.
# Each entry is a wildcard pattern matched against the site URL (-like, case-insensitive).
# Use an exact URL for one site, or * wildcards for a group of sites. Examples:
# $ConfigSiteUrlFilter = @(
#     'https://contoso.sharepoint.com/sites/Finance'        # exact site
#     'https://contoso.sharepoint.com/sites/HR*'           # every site starting with HR
#     '*/sites/Project-*'                                  # any Project- site
#     '*/teams/Legal'                                      # a /teams/ site
# )
$ConfigSiteUrlFilter = @()
# =============================================================================

# Command-line values win; otherwise use the configuration variables
if (-not $TenantId)              { $TenantId              = $ConfigTenantId }
if (-not $ClientId)              { $ClientId              = $ConfigClientId }
if (-not $CertificateThumbprint) { $CertificateThumbprint = $ConfigCertificateThumbprint }
if (-not $SiteUrlFilter)         { $SiteUrlFilter         = $ConfigSiteUrlFilter }

$ErrorActionPreference = 'Stop'

# Well-known resource app IDs (identical in every tenant)
$GraphResourceAppId      = '00000003-0000-0000-c000-000000000000'   # Microsoft Graph
$SharePointResourceAppId = '00000003-0000-0ff1-ce00-000000000000'   # Office 365 SharePoint Online
$TargetPermission        = 'Sites.Selected'

#region ---------- Helpers ----------

function Invoke-GraphWithRetry {
    <#  GET with handling for HTTP 429 / 503 / 504 throttling: honors Retry-After,
        otherwise exponential back-off.  #>
    param(
        [Parameter(Mandatory)] [string] $Uri,
        [int] $MaxRetries = 6
    )
    $attempt = 0
    while ($true) {
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
        }
        catch {
            $status = $null
            $retryAfter = $null
            $resp = $_.Exception.Response
            if ($resp) {
                try { $status = [int]$resp.StatusCode } catch { }
                try {
                    if ($resp.Headers.RetryAfter.Delta) { $retryAfter = [int]$resp.Headers.RetryAfter.Delta.TotalSeconds }
                } catch { }
            }
            if (-not $status -and $_.Exception.Message -match '\b(429|503|504)\b') { $status = [int]$Matches[1] }

            if ($status -in 429, 503, 504 -and $attempt -lt $MaxRetries) {
                $attempt++
                $wait = if ($retryAfter) { $retryAfter } else { [math]::Min(120, [math]::Pow(2, $attempt) * 2) }
                Write-Warning "HTTP $status (throttled/busy). Waiting $wait s before retry $attempt of $MaxRetries ..."
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
}

function Get-GraphPaged {
    # Follows @odata.nextLink and returns every item from .value
    param([Parameter(Mandatory)] [string] $Uri)
    $next = $Uri
    while ($next) {
        $page = Invoke-GraphWithRetry -Uri $next
        if ($page.value) { $page.value }
        $next = $page.'@odata.nextLink'
    }
}

#endregion

#region ---------- 1. Connect (app-only) ----------

# Validate connection variables before connecting
$connVars = [ordered]@{ TenantId = $TenantId; ClientId = $ClientId; CertificateThumbprint = $CertificateThumbprint }
foreach ($k in $connVars.Keys) {
    if ([string]::IsNullOrWhiteSpace($connVars[$k]) -or $connVars[$k] -match '^<.*>$') {
        throw "Connection variable '$k' is not set. Update the CONFIGURATION section or pass -$k."
    }
}
$CertificateThumbprint = ($CertificateThumbprint -replace '\s', '').ToUpper()

# Confirm the certificate (with private key) is installed on this machine
$cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$CertificateThumbprint", "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction SilentlyContinue |
    Select-Object -First 1
if (-not $cert)               { throw "Certificate $CertificateThumbprint was not found in CurrentUser\My or LocalMachine\My." }
if (-not $cert.HasPrivateKey) { throw "Certificate $CertificateThumbprint has no private key on this machine." }
if ($cert.NotAfter -lt (Get-Date)) { throw "Certificate $CertificateThumbprint expired on $($cert.NotAfter)." }

# App-only connection to the Microsoft 365 tenant using the variables
$connectParams = @{
    TenantId              = $TenantId
    ClientId              = $ClientId
    CertificateThumbprint = $CertificateThumbprint
    Environment           = $Environment
    NoWelcome             = $true
}
Write-Host "Connecting to Microsoft Graph ($Environment) app-only: tenant $TenantId, app $ClientId ..." -ForegroundColor Cyan
Connect-MgGraph @connectParams

$ctx = Get-MgContext
if ($ctx.AuthType -ne 'AppOnly') {
    throw "Expected an app-only (AppOnly) session but got '$($ctx.AuthType)'. Check the connection variables."
}
$missing = @('Application.Read.All', 'Sites.FullControl.All') | Where-Object { $ctx.Scopes -notcontains $_ }
if ($missing) {
    Write-Warning "The access token does not contain: $($missing -join ', '). Grant these application permissions with admin consent, or the script will fail."
}

# Graph endpoint for the selected cloud (e.g. https://graph.microsoft.us for USGov)
$graphBase = (Get-MgEnvironment -Name $Environment).GraphEndpoint.TrimEnd('/')
$v1 = "$graphBase/v1.0"

#endregion

#region ---------- 2. Resolve Sites.Selected permission IDs ----------

Write-Host "Resolving '$TargetPermission' permission IDs..." -ForegroundColor Cyan
$resourceInfo = @{}   # resourceAppId -> @{ Name; SpId; RoleIds; ScopeIds }
foreach ($res in @(
        @{ AppId = $GraphResourceAppId;      Name = 'Microsoft Graph' },
        @{ AppId = $SharePointResourceAppId; Name = 'SharePoint' })) {

    $sp = (Invoke-GraphWithRetry -Uri "$v1/servicePrincipals(appId='$($res.AppId)')?`$select=id,appId,appRoles,oauth2PermissionScopes")
    $resourceInfo[$res.AppId] = [pscustomobject]@{
        Name     = $res.Name
        SpId     = $sp.id
        RoleIds  = @($sp.appRoles               | Where-Object value -eq $TargetPermission | ForEach-Object id)
        ScopeIds = @($sp.oauth2PermissionScopes | Where-Object value -eq $TargetPermission | ForEach-Object id)
    }
    Write-Host ("  {0}: application role id(s) {1}; delegated scope id(s) {2}" -f $res.Name,
        ($resourceInfo[$res.AppId].RoleIds -join ','), ($resourceInfo[$res.AppId].ScopeIds -join ','))
}

#endregion

#region ---------- 3. Find app registrations / enterprise apps using Sites.Selected ----------

Write-Host "Scanning app registrations..." -ForegroundColor Cyan
$apps = @{}   # keyed by appId (client ID) - this is what site grants reference

$allApps = Get-GraphPaged -Uri "$v1/applications?`$select=id,appId,displayName,requiredResourceAccess&`$top=999"
foreach ($app in $allApps) {
    $sources = [System.Collections.Generic.List[string]]::new()
    $types   = [System.Collections.Generic.List[string]]::new()
    foreach ($rra in $app.requiredResourceAccess) {
        $info = $resourceInfo[$rra.resourceAppId]
        if (-not $info) { continue }
        foreach ($ra in $rra.resourceAccess) {
            $isRole  = $ra.type -eq 'Role'  -and $info.RoleIds  -contains $ra.id
            $isScope = $ra.type -eq 'Scope' -and $info.ScopeIds -contains $ra.id
            if ($isRole -or $isScope) {
                if (-not $sources.Contains($info.Name)) { $sources.Add($info.Name) }
                $t = if ($isRole) { 'Application' } else { 'Delegated' }
                if (-not $types.Contains($t)) { $types.Add($t) }
            }
        }
    }
    if ($sources.Count -gt 0) {
        $apps[$app.appId] = [pscustomobject]@{
            AppDisplayName   = $app.displayName
            AppId            = $app.appId
            AppObjectId      = $app.id
            AppType          = 'App registration'
            PermissionSource = ($sources -join '; ')
            PermissionType   = ($types -join '; ')
            AdminConsented   = 'No'
        }
    }
}
Write-Host "  App registrations requesting $TargetPermission : $($apps.Count)"

# Consent check: app role assignments made on the Graph / SharePoint service principals
$allRoleIds = $resourceInfo.Values | ForEach-Object { $_.RoleIds }
foreach ($info in $resourceInfo.Values) {
    $assignments = Get-GraphPaged -Uri "$v1/servicePrincipals/$($info.SpId)/appRoleAssignedTo?`$top=999"
    foreach ($a in ($assignments | Where-Object { $allRoleIds -contains $_.appRoleId })) {
        # principalId is the client app's service principal object ID -> resolve its appId
        $clientSp = Invoke-GraphWithRetry -Uri "$v1/servicePrincipals/$($a.principalId)?`$select=id,appId,displayName,appOwnerOrganizationId"
        if ($apps.ContainsKey($clientSp.appId)) {
            $apps[$clientSp.appId].AdminConsented = 'Yes'
        }
        elseif ($IncludeEnterpriseApps) {
            $apps[$clientSp.appId] = [pscustomobject]@{
                AppDisplayName   = $clientSp.displayName
                AppId            = $clientSp.appId
                AppObjectId      = $clientSp.id
                AppType          = 'Enterprise app (service principal)'
                PermissionSource = $info.Name
                PermissionType   = 'Application'
                AdminConsented   = 'Yes'
            }
        }
        else {
            Write-Verbose "Enterprise app '$($clientSp.displayName)' has consented Sites.Selected (use -IncludeEnterpriseApps to report it)."
        }
    }
}
Write-Host "  Total apps in scope: $($apps.Count)"

if ($apps.Count -eq 0) {
    Write-Host "No apps use $TargetPermission. Nothing to export." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    return
}

#endregion

#region ---------- 4. Enumerate sites and read their app permission grants ----------

Write-Host "Enumerating SharePoint sites (this can take a while in large tenants)..." -ForegroundColor Cyan
$sites = Get-GraphPaged -Uri "$v1/sites/getAllSites?`$select=id,webUrl,displayName&`$top=999"
if (-not $IncludeOneDrive) { $sites = $sites | Where-Object { $_.webUrl -notmatch '-my\.sharepoint\.(com|us|cn)' } }
if ($SiteUrlFilter) {
    # Keep a site if its URL matches ANY pattern; trailing '/' is ignored on both sides
    $patterns = @($SiteUrlFilter | Where-Object { $_ } | ForEach-Object { $_.Trim().TrimEnd('/') })
    $sites = $sites | Where-Object {
        $url = $_.webUrl.TrimEnd('/')
        foreach ($p in $patterns) { if ($url -like $p) { return $true } }
        $false
    }
    Write-Host "  Site URL filter applied ($($patterns.Count) pattern(s)): $($patterns -join ', ')"
}
$sites = @($sites)
Write-Host "  Sites to scan: $($sites.Count)"

$results = [System.Collections.Generic.List[object]]::new()
$errors  = [System.Collections.Generic.List[object]]::new()
$appsWithGrants = [System.Collections.Generic.HashSet[string]]::new()
$i = 0

foreach ($site in $sites) {
    $i++
    if ($i % 25 -eq 0 -or $i -eq $sites.Count) {
        Write-Progress -Activity "Reading site permissions" -Status "$i of $($sites.Count): $($site.webUrl)" `
            -PercentComplete ([int](100 * $i / $sites.Count))
    }
    try {
        $perms = Get-GraphPaged -Uri "$v1/sites/$($site.id)/permissions"
    }
    catch {
        $errors.Add([pscustomobject]@{ SiteUrl = $site.webUrl; SiteId = $site.id; Error = $_.Exception.Message })
        continue
    }

    foreach ($perm in $perms) {
        # grantedToIdentitiesV2 is current; grantedToIdentities is the legacy property
        $identities = @($perm.grantedToIdentitiesV2) + @($perm.grantedToIdentities) | Where-Object { $_ }
        $grantAppIds = $identities | ForEach-Object { $_.application.id } | Where-Object { $_ } | Select-Object -Unique

        foreach ($grantAppId in $grantAppIds) {
            if (-not $apps.ContainsKey($grantAppId)) { continue }

            # The list call may omit roles; fetch the individual permission if needed
            $roles = $perm.roles
            if (-not $roles) {
                try { $roles = (Invoke-GraphWithRetry -Uri "$v1/sites/$($site.id)/permissions/$($perm.id)").roles } catch { $roles = @('<unknown>') }
            }

            $a = $apps[$grantAppId]
            $results.Add([pscustomobject]@{
                AppDisplayName   = $a.AppDisplayName
                AppId            = $a.AppId
                AppObjectId      = $a.AppObjectId
                AppType          = $a.AppType
                PermissionSource = $a.PermissionSource
                PermissionType   = $a.PermissionType
                AdminConsented   = $a.AdminConsented
                SiteUrl          = $site.webUrl
                SiteId           = $site.id
                SiteName         = $site.displayName
                Roles            = ($roles -join '; ')
                SitePermissionId = $perm.id
            })
            [void]$appsWithGrants.Add($grantAppId)
        }
    }
}
Write-Progress -Activity "Reading site permissions" -Completed

# Apps that have Sites.Selected but no site grants (useful for clean-up)
foreach ($a in $apps.Values | Where-Object { -not $appsWithGrants.Contains($_.AppId) }) {
    $results.Add([pscustomobject]@{
        AppDisplayName   = $a.AppDisplayName
        AppId            = $a.AppId
        AppObjectId      = $a.AppObjectId
        AppType          = $a.AppType
        PermissionSource = $a.PermissionSource
        PermissionType   = $a.PermissionType
        AdminConsented   = $a.AdminConsented
        SiteUrl          = '<no site grants found>'
        SiteId           = ''
        SiteName         = ''
        Roles            = ''
        SitePermissionId = ''
    })
}

#endregion

#region ---------- 5. Export ----------

$results | Sort-Object AppDisplayName, SiteUrl |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Results exported to: $OutputPath ($($results.Count) rows)" -ForegroundColor Green

if ($errors.Count -gt 0) {
    $errPath = [System.IO.Path]::ChangeExtension($OutputPath, $null).TrimEnd('.') + '-errors.csv'
    $errors | Export-Csv -Path $errPath -NoTypeInformation -Encoding UTF8
    Write-Warning "$($errors.Count) site(s) could not be read. Details: $errPath"
}

Write-Host ("Summary: {0} app(s) with {1}; {2} have site grants; {3} site(s) scanned." -f `
        $apps.Count, $TargetPermission, $appsWithGrants.Count, $sites.Count)

Disconnect-MgGraph | Out-Null

#endregion

<# ============================================================================
README - Setting up the scanner app registration (one-time)
==============================================================================

1. Create a certificate (run on the machine that will run the script):

     $cert = New-SelfSignedCertificate -Subject "CN=SitesSelectedScanner" `
               -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable `
               -KeySpec Signature -KeyLength 2048 -HashAlgorithm SHA256 `
               -NotAfter (Get-Date).AddYears(1)
     Export-Certificate -Cert $cert -FilePath .\SitesSelectedScanner.cer
     $cert.Thumbprint     # use as -CertificateThumbprint

   (Use a CA-issued certificate or Azure Key Vault in production.)

2. Register the app:
     Entra admin center > Identity > Applications > App registrations > New registration
       Name: SitesSelected-Inventory-Scanner
       Supported account types: Accounts in this organizational directory only
       Redirect URI: none

3. Upload the certificate:
     App > Certificates & secrets > Certificates > Upload certificate > SitesSelectedScanner.cer

4. Add API permissions:
     App > API permissions > Add a permission > Microsoft Graph > Application permissions
       - Application.Read.All
       - Sites.FullControl.All
     Remove the default delegated User.Read permission (not needed).

5. Grant admin consent:
     App > API permissions > Grant admin consent for <tenant>
     (Requires Global Administrator or Privileged Role Administrator.)

6. Note the Application (client) ID and Directory (tenant) ID from the Overview page.

7. Put the Tenant ID, Client ID and certificate thumbprint into the CONFIGURATION
   section at the top of the script ($ConfigTenantId, $ConfigClientId,
   $ConfigCertificateThumbprint).

8. Install the module and run:
     Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
     .\Get-SitesSelectedAppGrants.ps1

Permission summary
------------------
| API             | Permission             | Type        | Why                                            |
|-----------------|------------------------|-------------|------------------------------------------------|
| Microsoft Graph | Application.Read.All   | Application | Read app registrations, service principals,    |
|                 |                        |             | and app role assignments (consent status)      |
| Microsoft Graph | Sites.FullControl.All  | Application | Enumerate all sites and read site permission   |
|                 |                        |             | grants (/sites/{id}/permissions)               |

Security notes
--------------
- Sites.FullControl.All is highly privileged. Protect the certificate, restrict who
  can run the script, and consider removing the permission or disabling the app
  between audits.
- The script is read-only and makes no changes.
- For GCC High use -Environment USGov; for DoD use -Environment USGovDoD.
============================================================================ #>
