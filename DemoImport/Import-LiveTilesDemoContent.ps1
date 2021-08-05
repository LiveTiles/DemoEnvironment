<#
.SYNOPSIS
    Imports demo content and term set for LiveTiles Intranet modules to SharePoint Online"

.DESCRIPTION
    Requirements:
    - Powershell v5
    - PnP.PowerShell, AzureAD

    All modules will be validated and potentially updated when the script runs.

    The script will perform the following actions:
    - Connect to SharePoint tenant and check site collection exists
    - Import LiveTiles theme 
    - Import LiveTiles TermSet Group
    - Check target site collection exists
    - Set the LiveTiles theme on the site
    - Import landing page content and SharePoint events
    - Import news page content
    - Import policies page content
    - Import topics page content
    - Import demo documents
    - Site configurations e.g. set default home page
    - Update LiveTiles Intranet Json configuration files so they are ready for import

    With the -WhatIf switch, only the testing and verification of the above actions will be performed. No CREATE actions are actually performed.

.PARAMETER tenantName
    The tenant name to be updated.

.PARAMETER importUrl
    The relative url to the site collection e.g. /sites/intranet

.PARAMETER targetUser
    The username of a valid user on the target tenant.
    
.EXAMPLE
    .\Import-LiveTilesDemoContent -tenantName "TryLiveTilesXX" -importUrl "/sites/intranet" -targetUser "admin@TryLiveTilesXX.onmicrosoft.com"
    
    Imports demo content on the specified tenant and site collection. After the script has been completed, the Json configuration files must up added to the LiveTiles intranet configuration manually 

.NOTES
    AUTHOR: Garry Sinclair
    LASTEDIT: 30-06-2021 
    v1.0
        First Release. Identified 2Do´s:
            - Automatically upload the LiveTiles Json configuration. Requires generation of access token.

.LINK
    Updated versions of this script will be available on the LiveTiles Partner Portal
#>
param (
    [Parameter(Mandatory=$true)]
    [string]$tenantName,
    [Parameter(Mandatory=$true)]
    [string]$importUrl,
    [Parameter(Mandatory=$true)]
    [string]$targetUser
)

if (($PSVersionTable.PSVersion.Major -lt 5) -or (($PSVersionTable.PSVersion.Major -eq 5) -and ($PSVersionTable.PSVersion.Minor -lt 1))) {
    Write-Host "PowerShell version 5.1 or greater is required to run this script." -ForegroundColor Red
    Write-Host "Please find the download packages here: https://docs.microsoft.com/en-us/powershell/scripting/install/installing-windows-powershell?view=powershell-6#upgrading-existing-windows-powershell"
    Exit
}

function Import-LiveTilesSite {
    param(
        [Parameter(Mandatory=$true)]
        [String]$siteName,
        [Parameter(Mandatory=$true)]
        [String]$targetUser
    )

    Write-Host "Importing $siteName ..."

    ((Get-Content -Path "$siteName/$siteName-Template.xml" -Raw) -replace "USER_PLACEHOLDER", "$targetUser") | Set-Content -Path "$siteName/$siteName-TemplateReplaced.xml"

    Invoke-PnPSiteTemplate -Path "$siteName/$siteName-TemplateReplaced.xml"
    Remove-Item -Path "$siteName/$siteName-TemplateReplaced.xml"

    Publish-LiveTilesPages -siteName $siteName

    Write-Host "Done"
}

function Import-LiveTilesTermGroup {
    param(
        [Parameter(Mandatory=$true)]
        [String]$targetUser
    )

    Write-Host "Importing term group $termGroupId ..."

    ((Get-Content -Path "Terms/Terms.xml" -Raw) -replace "USER_PLACEHOLDER", $targetUser) | Set-Content -Path "Terms/TermsReplaced.xml"
    
    Import-PnPTermGroupFromXml -Path "Terms/TermsReplaced.xml"
    Remove-Item -Path "Terms/TermsReplaced.xml"
    Write-Host "Done"
}

function Configure-LiveTilesSite {
    param(    )
    
    Write-Host "Configuring site ..."

    #Get the Site
    $Site = Get-PnPSite –Includes CustomScriptSafeDomains  
 
    #Add domain
    $Domain = [Microsoft.SharePoint.Client.ScriptSafeDomainEntityData]::new()
    $Domain.DomainName = "reach.livetiles.io"
    $null = $Site.CustomScriptSafeDomains.Create($Domain)
    Invoke-PnPQuery

    Set-PnPHomePage -RootFolderRelativeUrl "SitePages/HomeReachContent.aspx"

    Write-Host "Done"
}

function Publish-LiveTilesPages {
    param(
        [Parameter(Mandatory=$true)]
        [String]$siteName
    )

    $pageDataArray = Get-Content -Path "$siteName/$siteName-pageData.json" | ConvertFrom-Json

    $list = Get-PnPList -Identity "SitePages"

    $pageDataArray |% {
        $pageDataItem = $_
        $filename = $pageDataItem.Filename
        $page = Get-PnPPage -Identity $filename
        Write-Host "Publishing $filename ..."
        $null = Set-PnPListItem -List $list -Identity $page.PageId -Values @{"BannerImageUrl" = $pageDataItem.BannerImageUrl.Url; }
        if($filename.StartsWith("Home")){
            $null = Set-PnPPage -Identity $filename -Publish -LayoutType Home
        } else {
            $null = Set-PnPPage -Identity $filename -Publish
        }
    }
}

function Import-LiveTilesTheme {

    $themepalette = @{
        "themePrimary" = "#7c4dff";
        "themeLighterAlt" = "#faf8ff";
        "themeLighter" = "#eae2ff";
        "themeLight" = "#d8c9ff";
        "themeTertiary" = "#b094ff";
        "themeSecondary" = "#8c62ff";
        "themeDarkAlt" = "#7045e6";
        "themeDark" = "#5e3ac2";
        "themeDarker" = "#452b8f";
        "neutralLighterAlt" = "#f8f8f8";
        "neutralLighter" = "#f4f4f4";
        "neutralLight" = "#eaeaea";
        "neutralQuaternaryAlt" = "#dadada";
        "neutralQuaternary" = "#d0d0d0";
        "neutralTertiaryAlt" = "#c8c8c8";
        "neutralTertiary" = "#c2c2c2";
        "neutralSecondary" = "#858585";
        "neutralPrimaryAlt" = "#4b4b4b";
        "neutralPrimary" = "#333333";
        "neutralDark" = "#272727";
        "black" = "#1d1d1d";
        "white" = "#ffffff";
        "primaryBackground" = "#ffffff";
        "primaryText" = "#333333";
        "bodyBackground" = "#ffffff";
        "bodyText" = "#333333";
        "disabledBackground" = "#f4f4f4";
        "disabledText" = "#c8c8c8";
    }

    Add-PnPTenantTheme -Identity "LiveTiles" -Palette $themepalette -IsInverted $false -Overwrite
}

function Import-LiveTilesDemoDocuments {
    
    Write-Host "Importing demo documents ..."

    $null = Add-PnPFile -Path ".\Intranet\Team.tenant.xml" -Folder "PnPTemplates"
    $null = Add-PnPFile -Path ".\SourceFiles\Event Proposal Template.docx" -Folder "Shared Documents"
    $null = Add-PnPFile -Path ".\SourceFiles\LiveTiles - PPT Template.pptx" -Folder "Shared Documents"
    $null = Add-PnPFile -Path ".\SourceFiles\LiveTiles - The power of employee communication - eBook.pdf" -Folder "Shared Documents"
    $null = Add-PnPFile -Path ".\SourceFiles\LiveTiles Reach for Schools.pdf" -Folder "Shared Documents"
    $null = Add-PnPFile -Path ".\SourceFiles\LiveTiles_Reach_Evaluation_Guide.pdf" -Folder "Shared Documents"

    Write-Host "Done"
}


# Begin execution here

$tenantUrl = "https://$tenantName.sharepoint.com"
Connect-PnPOnline -Url $tenantUrl -Interactive

Import-LiveTilesTheme
Import-LiveTilesTermGroup -targetUser $targetUser


$importUrl = "$tenantUrl$importUrl"

Connect-PnPOnline -Url $importUrl -Interactive
$site = Get-PnPSite -ErrorAction SilentlyContinue

if($site -eq $null) {
    Write-Host "Site $importUrl does not exist. Create before continueing."
    Exit
}

Set-PnPWebTheme -Theme LiveTiles
Import-LiveTilesSite -siteName "Intranet" -targetUser $targetUser
Import-LiveTilesSite -siteName "News" -targetUser $targetUser
Import-LiveTilesSite -siteName "Policies" -targetUser $targetUser
Import-LiveTilesSite -siteName "Topics" -targetUser $targetUser
Import-LiveTilesDemoDocuments
Configure-LiveTilesSite