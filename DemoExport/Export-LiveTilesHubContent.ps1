<#
.SYNOPSIS
    Exports demo content and term set for LiveTiles Intranet modules to SharePoint Online"

.DESCRIPTION
    Requirements:
    - Powershell v5
    - PnP.PowerShell, AzureAD

    All modules will be validated and potentially updated when the script runs.

    The script will perform the following actions:
    - Exports the landing pages, pnp templates, and SharePoint based events
    - Exports the LiveTiles term group (needed for workspaces)
    - Exports the pnp provisioning template for MS Teams
    - Updates the Json files with a placeholder so they can more easily be reused on other tenants
    - Exports the news pages
    - Exports the policy pages
    - Exports the topics pages

    With the -WhatIf switch, only the testing and verification of the above actions will be performed. No CREATE actions are actually performed.

.PARAMETER sourceTenant
    The source tenant name to be exported.

.PARAMETER sourceUrl
    The source relative url to the main site collection e.g. /sites/intranet

.PARAMETER usersToReplace
    A string array of usernames to be replaced with a placeholder for easy import to other environments.

.PARAMETER sourceReachSubscription
    The reach subscription id for the source environment.

.PARAMETER newsUrl
    (Optional) The url to the site collection containing news pages for export e.g. /sites/news

.PARAMETER topicsUrl
    (Optional) The url to the site collection containing topics pages for export e.g. /sites/topics

.PARAMETER policiesUrl
    (Optional) The url to the site collection containing policy pages for export e.g. /sites/policies

.EXAMPLE
    .\Export-LiveTilesHubContent -sourceTenant "TryLiveTilesXX" -sourceUrl "/sites/intranet" -usersToReplace "garry.sinclair@trylivetilesxx.onmicrosoft.com","christoffer.soltau@trylivetilesxx.onmicrosoft.com" -sourceReachSubscription "c9d53069-5baa-4ae8-b87c-d0fa325dcf3e" -newsUrl "/sites/news" -topicsUrl "/sites/topics" -policiesUrl "/sites/policies"
    
    Exports demo content on the specified tenant and site collections. After the script has been completed, the Json configuration files must up exported manually to the JsonFiles folder 

.NOTES
    AUTHOR: Garry Sinclair
    LASTEDIT: 30-06-2021 
    v1.0
        First Release. Identified 2Do´s:
            - Automatically export the LiveTiles Json configuration. Requires generation of access token.

.LINK
    Updated versions of this script will be available on the LiveTiles Partner Portal
#>
param (
    [Parameter(Mandatory=$true)]
    [string]$sourceTenant,
    [Parameter(Mandatory=$true)]
    [string]$sourceUrl,
    [Parameter(Mandatory=$true)]
    [string[]]$usersToReplace,
    [Parameter(Mandatory=$true)]
    [string]$sourceReachSubscription,
    [Parameter(Mandatory=$false)]
    [string]$newsUrl,
    [Parameter(Mandatory=$false)]
    [string]$topicsUrl,
    [Parameter(Mandatory=$false)]
    [string]$policiesUrl
)

function Export-LiveTilesSitePageData {
    param(
        [Parameter(Mandatory=$true)]
        [String]$siteUrl,
        [Parameter(Mandatory=$true)]
        [String]$siteName
    )

    $pageData= New-Object -TypeName 'System.Collections.ArrayList'

    Connect-PnPOnline -Url $siteUrl -Interactive
    
    $list = Get-PnPList -Identity "SitePages"
    $items = Get-PnPListItem -List $list

    $items |?  {
        $_.FieldValues.FileLeafRef -ne "Home.aspx"
    } |% {
        $filename = $_.FieldValues.FileLeafRef
        Write-Host "Exporting $filename"
        #Export-PnPPage -Identity $filename -Out "$siteName-$filename.xml" -PersistBrandingFiles

        $pageItem = @{
            Filename=$filename
            BannerImageUrl=$_.FieldValues.BannerImageUrl
        }
        $pageData.Add($pageItem)
    }

    #Export-PnPListToSiteTemplate -List "SiteAssets"
    $pageData | ConvertTo-Json | Set-Content -Path "$siteName/$siteName-pageData.json"
}

function Export-LiveTilesSite {
    param(
        [Parameter(Mandatory=$true)]
        [String]$siteUrl,
        [Parameter(Mandatory=$true)]
        [String]$siteName,
        [Parameter(Mandatory=$false)]
        [String[]]$usersToReplace = $null,
        [Parameter(Mandatory=$false)]
        [String[]]$listsToExtract = $null
    )

    Connect-PnPOnline -Url $siteUrl -Interactive
    
    Write-Host "Exporting $siteName ..."
    if($listsToExtract -ne $null) {
        Get-PnPSiteTemplate -Out "$siteName/$siteName-Template.xml" -Handlers PageContents, Pages, Lists -IncludeAllPages -IncludeNativePublishingFiles -PersistPublishingFiles -PersistBrandingFiles -ListsToExtract $listsToExtract
    } else {
        Get-PnPSiteTemplate -Out "$siteName/$siteName-Template.xml" -Handlers PageContents, Pages -IncludeAllPages -IncludeNativePublishingFiles -PersistPublishingFiles -PersistBrandingFiles
    }
    Write-Host "Done ..."

    $usersToReplace |% { 
        Write-Host "Replacing lvtmaster reference $_ ..."
        ((Get-Content -Path "$siteName/$siteName-Template.xml" -Raw) -replace $_, "USER_PLACEHOLDER") | Set-Content -Path "$siteName/$siteName-Template.xml"
        Write-Host "Done ..."    
    }

    Export-LiveTilesSitePageData -siteUrl $siteUrl -siteName $siteName
}

function Export-LiveTilesTermGroup {
    param(
        [Parameter(Mandatory=$true)]
        [String]$siteUrl,
        [Parameter(Mandatory=$true)]
        [String]$termGroupId,
        [Parameter(Mandatory=$false)]
        [String[]]$usersToReplace = $null
    )

    Connect-PnPOnline -Url $siteUrl -Interactive
    
    Write-Host "Exporting term group $termGroupId ..."
    Export-PnPTermGroupToXml -Identity $termGroupId -Out "Terms/Terms.xml"
    Write-Host "Done ..."

    $usersToReplace |% { 
        Write-Host "Replacing lvtmaster reference $_ ..."
        ((Get-Content -Path "Terms/Terms.xml" -Raw) -replace $_, "USER_PLACEHOLDER") | Set-Content -Path "Terms/Terms.xml"
        Write-Host "Done ..."    
    }
}


function Update-LiveTilesJsonFiles {
    param(
        [Parameter(Mandatory=$true)]
        [String]$sourceTenant,
        [Parameter(Mandatory=$true)]
        [String]$sourceSubscription
    )

    Write-Host "Updating hub.json - Replacing source subscription placeholder ..."
    ((Get-Content -Path "JsonFiles/original-hub.json" -Raw) -replace $sourceSubscription, "SUBSCRIPTION_PLACEHOLDER") | Set-Content -Path "JsonFiles/original-hub.json"
    Write-Host "Done ..."

    Write-Host "Updating hub.json - Replacing $sourceTenant with placeholder ..."
    ((Get-Content -Path "JsonFiles/original-hub.json" -Raw) -replace "$sourceTenant", "TENANT_PLACEHOLDER") | Set-Content -Path "JsonFiles/original-hub.json"
    Write-Host "Done ..."

    Write-Host "Updating siteType-Community.json - Replacing $sourceTenant with placeholder ..."
    ((Get-Content -Path "JsonFiles/original-siteType-Community.json" -Raw) -replace "$sourceTenant", "TENANT_PLACEHOLDER") | Set-Content -Path "JsonFiles/original-siteType-Community.json"
    Write-Host "Done ..."

    Write-Host "Updating siteType-Project.json - Replacing $sourceTenant with placeholder ..."
    ((Get-Content -Path "JsonFiles/original-siteType-Project.json" -Raw) -replace "$sourceTenant", "TENANT_PLACEHOLDER") | Set-Content -Path "JsonFiles/original-siteType-Project.json"
    Write-Host "Done ..."

    Write-Host "Updating siteType-Team.json - Replacing $sourceTenant with placeholder ..."
    ((Get-Content -Path "JsonFiles/original-siteType-Team.json" -Raw) -replace "$sourceTenant", "TENANT_PLACEHOLDER") | Set-Content -Path "JsonFiles/original-siteType-Team.json"
    Write-Host "Done ..."
}

$tenantUrl = "https://$sourceTenant.sharepoint.com"

Export-LiveTilesSite -siteUrl "$tenantUrl$sourceUrl" -siteName "Intranet" -usersToReplace $usersToReplace -listsToExtract "PnPTemplates","Events"
Add-PnPDataRowsToSiteTemplate -Path ".\Intranet\Intranet-Template.xml" -List "Events"

Export-LiveTilesTermGroup -siteUrl "$tenantUrl$sourceUrl" -termGroupId 467cf726-dd98-411b-a88e-2d4965fbe3b5 -usersToReplace $usersToReplace
Get-PnPFile -Url "$sourceUrl/PnPTemplates/Team.tenant.xml"  -Path ".\Intranet" -FileName "Team.tenant.xml" -AsFile -Force
Update-LiveTilesJsonFiles -sourceTenant $sourceTenant -sourceSubscription $sourceReachSubscription

if($newsUrl -ne $null){
    Export-LiveTilesSite -siteUrl "$tenantUrl$newsUrl" -siteName "News" -usersToReplace $usersToReplace
}
if($policiesUrl -ne $null){
    Export-LiveTilesSite -siteUrl "$tenantUrl$policiesUrl" -siteName "Policies" -usersToReplace $usersToReplace
}
if($topicsUrl -ne $null){
    Export-LiveTilesSite -siteUrl "$tenantUrl$topicsUrl" -siteName "Topics" -usersToReplace $usersToReplace
}

# Consider this as another way of migrating SP page images
#Add-PnPDataRowsToSiteTemplate -Path ".\News\News-Template.xml" -List "Site Pages" -Fields "FileLeafRef","BannerImageUrl"
#Add-PnPDataRowsToSiteTemplate -Path ".\Policies\Policies-Template.xml" -List "Site Pages" -Fields "Title","BannerImageUrl"
#Add-PnPDataRowsToSiteTemplate -Path ".\Topics\Topics-Template.xml" -List "Site Pages"