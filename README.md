# Get-SitesSelectedAppGrants

What the script does
Signs in app-only with a certificate (recommended) or a client secret. It stops if the session isn't app-only.
Finds the Sites.Selected permission IDs by reading them from the Microsoft Graph and SharePoint Online service principals in the tenant, rather than hard-coding them.
Finds the app registrations whose API permissions include Sites.Selected, from Graph, SharePoint or both, application or delegated. For each one it checks whether admin consent was actually granted.
Add -IncludeEnterpriseApps to also report third-party or multi-tenant apps that aren't registered in your tenant but have consent for Sites.Selected.
Lists every site and reads each site's app permission grants, then matches those grants to the apps from step 3.
Writes the results to CSV: one row per app and site, plus one row for each app that has Sites.Selected but no site grants. Any sites it couldn't read go into a separate errors.csv.

CSV columns: AppDisplayName, AppId, AppObjectId, AppType, PermissionSource, PermissionType, AdminConsented, SiteUrl, SiteId, SiteName, Roles, SitePermissionId

# App-only permissions the scanner app needs

API	Permission	
Microsoft Graph	Application.Read.All -	Application	Read app registrations, service principals and consent status
Microsoft Graph	Sites.FullControl.All -	Application	List all sites and read each site's permissions. Reading site permissions app-only requires this permission; nothing lower works.
Admin consent is required for both, from a Global Administrator or Privileged Role Administrator. Remove the default delegated User.Read, since the script doesn't use it.

# One-time setup
Create a certificate: New-SelfSignedCertificate, then export the .cer file. The exact commands are in the README block at the end of the script.
Register the scanner app: Entra admin center → App registrations → New registration (single tenant, no redirect URI).
Upload the certificate: Certificates & secrets → upload the .cer file.
Add the permissions: API permissions → add the two Graph application permissions above → grant admin consent.

# Things to know before running it
It makes one Graph call per site. In large tenants that can take hours and can be throttled. The script handles 429, 503 and 504 responses by waiting for the time in the Retry-After header, or waiting longer after each retry if there's no header. Run it off-hours, or use -SiteUrlFilter to scan in batches.
It doesn't check list-, library- or file-level grants. Those use the newer *.SelectedOperations.Selected permissions and aren't included.
Sites.FullControl.All is a high-privilege permission. Protect the certificate, and consider disabling the scanner app between audits. The script itself only reads data and makes no changes.

Set it in the configuration section
In the configuration section near the top, add one site URL or pattern per line inside @( ). No commas are needed when each entry is on its own line:

# Using SiteURL Filter
The $ConfigureSiteUrlFilder can include multiple sites using various parameters as shown below

$ConfigSiteUrlFilter = @(

        'https://contoso.sharepoint.com/sites/Finance'      # exact site
        'https://contoso.sharepoint.com/sites/HR*'         # every site starting with HR
        '*/sites/Project-*'                                # any Project- site
        '*/teams/Legal'                                    # a /teams/ site
    
)

Leave it as $ConfigSiteUrlFilter = @() to scan all sites.
