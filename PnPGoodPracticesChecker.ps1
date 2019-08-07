if (!("LoginTypes" -as [type])) { Add-Type -TypeDefinition "public enum LoginTypes {Simple, ADFS, Multifactor }" }

#-------------------------------------------------------------------------------------------------------------------------------------------------#
# CONFIGURATION START                                                                                                                             #
#v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v-v#

# admin URL for your tenant; alternatively a "normal" site URL might work, but certain checks won't succeed because of missing permissions
$adminSiteOrNormalSiteYouHaveAccessTo = "https://yourtenant-admin.sharepoint.com"

# choose the login type to use - depending on how you need to authenticate to Office 365 this is more or less complicated...
$loginType = [LoginTypes]::Simple

#^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^#
# CONFIGURATION END - NOW RUN THE SCRIPT                                                                                                          #
#-------------------------------------------------------------------------------------------------------------------------------------------------#


Write-Host "Checking that SharePointPnPPowerShellOnline is available..."
# check if cmdlets are available
$cmd = Get-Command | Where-Object { $_.Name -eq "Connect-PnPOnline" -and $_.Source -eq "SharePointPnPPowerShellOnline" }
if (!$cmd) 
{
    Write-Host "PnP PowerShell is not installed or not loaded. Please fix."
    exit
}
Write-Host "Done: Checking that SharePointPnPPowerShellOnline is available"


function Connect($url = $adminSiteOrNormalSiteYouHaveAccessTo)
{
    if ($loginType -eq [LoginTypes]::Simple)
    {
        Connect-PnPOnline $url -Credentials $credentials -ErrorAction Stop
    } elseif ($loginType -eq [LoginTypes]::Multifactor)
    {
        Connect-PnPOnline $url -PnPO365ManagementShell -LaunchBrowser -ErrorAction Stop
    } elseif ($loginType -eq [LoginTypes]::ADFS)
    {
        Connect-PnPOnline $url -UseWebLogin -ErrorAction Stop
    } else
    {
        Write-Error "Don't know how to connect"
        exit
    }

    if (!$?) 
    {
        Write-Host "Could not connect." -ForegroundColor Red
        exit
    }
}

function CheckModifiedTitleProperty
{
    Write-Host "[Search] Checking mapping of managed property 'Title' (might take some time)..."
    try
    {
        $config =Get-PnPSearchConfiguration -Scope Subscription -OutputFormat ManagedPropertyMappings
        if (!$config)
        {
            Write-Host "[Search] Cannot check mapping of managed property 'Title' for unknown reason. Maybe you did not connect to the -admin site for administrative access." -ForegroundColor Red
            return
        }

        $standardMapping = "Title,MetadataExtractorTitle,TermTitle,2,ows_BaseName,Title,MailSubject,5,urn:schemas-microsoft-com:sharepoint:portal:profile:PreferredName,urn:schemas.microsoft.com:fulltextqueryinfo:displaytitle,ows_Title,10,9,MetadataExtractorTitle"
        $modifiedTitleProperty = $config | Where-Object { $_.Name -eq "2" -and [string]::Join(",", $_.Mappings) -ne $standardMapping }
        if ($modifiedTitleProperty) 
        {
            Write-Host "[Search] Mappings of managed property 'Title' have been changed. This might have unintended consequences." -ForegroundColor Yellow
            Write-Host "[Search] See: https://joannecklein.com/2017/08/01/office-365-sharepoint-app-site-titles/" -ForegroundColor DarkYellow

        } else 
        {
            Write-Host "[Search] Managed property 'Title' has default mapping" -ForegroundColor Green
        }

    } finally 
    {
        Write-Host "[Search] Done: Checking mapping of managed property 'Title'"
    }
}

function CheckTenantLanguage
{
    Write-Host "[Tenant] Checking tenant language..."
    try
    {
        # assumption: tenant language is the default language of the root site; this might change when the root site can be replaced
        $uri = New-Object System.Uri $adminSiteOrNormalSiteYouHaveAccessTo
        $rootSiteUrl = "$($uri.Scheme)://$($uri.Host)"
        Write-Host "[Tenant] Trying to get '$rootSiteUrl'"
        Disconnect-PnPOnline
        Connect $rootSiteUrl
        try
        {
            $web = Get-PnPWeb
            $lcid = Get-PnPProperty $web Language

            if ($lcid -ne 1033) 
            {
                Write-Host "[Tenant] The tenant language was not set to US English (instead it's LCID $($lcid)). This might have unintended consequences. You cannot change this." -ForegroundColor Yellow
                Write-Host "[Tenant] See: https://thomy.tech/10-things-you-should-do-with-your-office365-demo-or-dev-tenant." -ForegroundColor DarkYellow

                # maybe check Get-MsolCompanyInformation as well?
                # not sure where the default language matters... have to investigate.
            } else
            {
                Write-Host "[Tenant] The tenant language was (problably) set to US English. This is usually a good choice when using services like the SharePoint Online Provisioning Service (also for newly created sites)." -ForegroundColor Green
                Write-Host "[Tenant] SharePoint Online Provisioning Service: https://provisioning.sharepointpnp.com/" -ForegroundColor DarkGreen
            }
        } finally 
        {
            Disconnect-PnPOnline
            Connect
        }
    } finally
    {
        Write-Host "[Tenant] Done: Checking tenant language"
    }
}

function CheckAppCatalogExistence
{
    Write-Host "[Tenant] Checking tenant app catalog site existence..."
    try
    {
        $web = Get-PnPWeb
        $url = [Microsoft.SharePoint.Client.WebExtensions]::GetAppCatalog($web).AbsoluteUri.ToString()
        if (!$? -or !$url) 
        {
            Write-Host "[Tenant] There is no tenant app catalog site. You really should create one." -ForegroundColor Yellow
            Write-Host "[Tenant] Background: https://thomy.tech/10-things-you-should-do-with-your-office365-demo-or-dev-tenant" -ForegroundColor DarkYellow
            Write-Host "[Tenant] Instructions how to create one: https://docs.microsoft.com/en-us/sharepoint/use-app-catalog#step-1-create-the-app-catalog-site-collection" -ForegroundColor DarkYellow
        
        } else 
        {
            Write-Host "[Tenant] Tenant app catalog site exists." -ForegroundColor Green
        }
    } finally
    {
        Write-Host "[Tenant] Done: Checking tenant app catalog site existence"
    }
}

function CheckAppCatalogAccess
{
    Write-Host "[Tenant] Checking tenant app catalog access..."
    try
    {
        $web = Get-PnPWeb
        # get app catalog URL; note: don't use Get-PnPTenantAppCatalogUrl since this does not work as guest user
        $url = [Microsoft.SharePoint.Client.WebExtensions]::GetAppCatalog($web).AbsoluteUri.ToString()
        if (!$url)
        {
            Write-Host "[Tenant] Skipping check since there is no app catalog site (or you don't have access)." -ForegroundColor Yellow
            return
        }

        Disconnect-PnPOnline
        Connect $url
        try
        {
            $loginNameEveryone = "c:0(.s|true"
            $loginNamePartEveryoneExceptExternal = "rolemanager|spo-grid-all-users"
            $everyoneWasFoundOnAppCatalog = $false
            $everyoneCanAccessAppCatalog = $false
            $everyoneExceptExternalWasFoundOnAppCatalog = $false
            $everyoneExceptExternalCanAccessAppCatalog = $false
            $canCheckPermissions = $true

            $web = Get-PnPWeb
            Get-PnPProperty $web SiteUsers > $null
            $web.SiteUsers | Where-Object { $_.LoginName -eq $loginNameEveryone} | ForEach-Object {
                $everyoneWasFoundOnAppCatalog = $true
                try
                {
                    $permResults = $web.GetUserEffectivePermissions($_.LoginName)
                    Invoke-PnPQuery
                    $everyoneCanAccessAppCatalog = $permResults.Value.Has([Microsoft.SharePoint.Client.PermissionKind]::ViewListItems)
                } catch 
                {
                    $canCheckPermissions = $false
                }
            }

            $web.SiteUsers | Where-Object { $_.LoginName -like "*$loginNamePartEveryoneExceptExternal*"} | ForEach-Object {
                $everyoneExceptExternalWasFoundOnAppCatalog = $true
                if ($canCheckPermissions)
                {
                    $permResults = $web.GetUserEffectivePermissions($_.LoginName)
                    Invoke-PnPQuery
                    $everyoneExceptExternalCanAccessAppCatalog = $permResults.Value.Has([Microsoft.SharePoint.Client.PermissionKind]::ViewListItems)
                }
            }

            if (!$everyoneWasFoundOnAppCatalog -or ($everyoneWasFoundOnAppCatalog -and !$everyoneCanAccessAppCatalog))
            {
                Write-Host "[Tenant] External users cannot access the app catalog. This is the standard setting. But they might not be able to see SharePoint Framework solutions (web parts etc.)." -ForegroundColor Green
                Write-Host "[Tenant] Note: You can allow access to external users by adding 'Everyone' on the app catalog site '$url'." -ForegroundColor DarkGreen
                Write-Host "[Tenant] See: https://laurakokkarinen.com/sharepoint-online-guest-user-troubles-and-how-to-get-past-them/#custom-share-point-apps-dont-load-for-guest-users-by-default" -ForegroundColor DarkGreen
            }
            if ($everyoneWasFoundOnAppCatalog -and !$canCheckPermissions)
            {
                Write-Host "[Tenant] The 'Everyone' group, which is related to external user access, was found, but its permissions could not be checked. Make sure you have enough permissions yourself." -ForegroundColor Yellow
            }

            if ($everyoneCanAccessAppCatalog)
            {
                Write-Host "[Tenant] External users can see the app catalog. This is a non-standard setting but might be desirable to make SharePoint Framework solutions work for external users." -ForegroundColor Green
                Write-Host "[Tenant] Note: You can prevent access for external users by replacing 'Everyone' with 'Everyone except external users' on the app catalog site '$url'." -ForegroundColor DarkGreen
            }

            if (!$everyoneWasFoundOnAppCatalog -and !$everyoneExceptExternalWasFoundOnAppCatalog)
            {
                Write-Host "[Tenant] The 'Everyone except external users' is not allowed to see the app catalog. This is non-standard, odd and should probably be corrected." -ForegroundColor Yellow
                Write-Host "[Tenant] Note: You can allow access to 'Everyone except external users' on the app catalog site '$url'." -ForegroundColor DarkYellow
            }

            if ($everyoneCanAccessAppCatalog -or $everyoneExceptExternalCanAccessAppCatalog)
            {
                Write-Host "[Tenant] All internal users can access the app catalog. This is standard." -ForegroundColor Green
            }

            if (($everyoneWasFoundOnAppCatalog -or $everyoneExceptExternalWasFoundOnAppCatalog) -and !$canCheckPermissions)
            {
                Write-Host "[Tenant] Found some of the 'Everyone' groups, but could not check its permissions. So it cannot be determined if internal users have indeed access to the app catalog. Make sure you have enough permissions yourself." -ForegroundColor Yellow
            }

        } finally 
        {
            # restore initial connection since we are currently connected to the app catalog
            Disconnect-PnPOnline
            Connect
        }
    } finally
    {
        Write-Host "[Tenant] Done: Checking tenant app catalog access"
    }
}

# try to find out if we are term store administrator; don't know how to check this directly so we try to access existing term groups
function CheckManagedMetadataServiceAdmin
{
    $groups = Get-PnPTermGroup
    $groupCount = $groups.Count
    if ($groupCount -eq 0)
    {
        Write-Host "Got 0 term groups, that's odd. There should be system groups available. Something went wrong." -ForegroundColor Yellow
        return
    }
    $errorCount = 0
    foreach ($group in $groups)
    {
        Get-PnPProperty $group GroupManagerPrincipalNames 2>$null
        if (!$?) 
        {
            $errorCount++
            break; # don't try all term groups for now... need to gather experience on this one
        }
    }

    if ($errorCount -gt 0) 
    {
        Write-Host "[Term Store] Looks like you have restricted access to the term store. This might be ok, but not if you are supposed to be the term store administrator." -ForegroundColor Yellow
        Write-Host "[Term Store] See: https://thomy.tech/10-things-you-should-do-with-your-office365-demo-or-dev-tenant" -ForegroundColor DarkYellow
        
    } else
    {
        Write-Host "[Term Store] Either you have access to all $groupCount term groups or you are term store administrator. This check shall pass." -ForegroundColor Green
    }
}

Connect
CheckModifiedTitleProperty
CheckTenantLanguage
CheckAppCatalogExistence
CheckAppCatalogAccess
CheckManagedMetadataServiceAdmin