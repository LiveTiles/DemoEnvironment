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
    ((Get-Content -Path "JsonFiles/hub.json" -Raw) -replace $sourceSubscription, "SUBSCRIPTION_PLACEHOLDER") | Set-Content -Path "JsonFiles/hub.json"
    Write-Host "Done ..."

    Write-Host "Updating hub.json - Replacing $sourceTenant with placeholder ..."
    ((Get-Content -Path "JsonFiles/hub.json" -Raw) -replace "$sourceTenant", "TENANT_PLACEHOLDER") | Set-Content -Path "JsonFiles/hub.json"
    Write-Host "Done ..."

    Write-Host "Updating siteType-Community.json - Replacing $sourceTenant with placeholder ..."
    ((Get-Content -Path "JsonFiles/siteType-Community.json" -Raw) -replace "$sourceTenant", "TENANT_PLACEHOLDER") | Set-Content -Path "JsonFiles/siteType-Community.json"
    Write-Host "Done ..."

    Write-Host "Updating siteType-Project.json - Replacing $sourceTenant with placeholder ..."
    ((Get-Content -Path "JsonFiles/siteType-Project.json" -Raw) -replace "$sourceTenant", "TENANT_PLACEHOLDER") | Set-Content -Path "JsonFiles/siteType-Project.json"
    Write-Host "Done ..."

    Write-Host "Updating siteType-Team.json - Replacing $sourceTenant with placeholder ..."
    ((Get-Content -Path "JsonFiles/siteType-Team.json" -Raw) -replace "$sourceTenant", "TENANT_PLACEHOLDER") | Set-Content -Path "JsonFiles/siteType-Team.json"
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