param (
    [Parameter(Mandatory=$true)]
    [string]$tenantName,
    [Parameter(Mandatory=$true)]
    [string]$importUrl,
    [Parameter(Mandatory=$true)]
    [string]$targetUser,
    [Parameter(Mandatory=$false)]
    [string]$targetReachSubscription
)

if (($PSVersionTable.PSVersion.Major -lt 5) -or (($PSVersionTable.PSVersion.Major -eq 5) -and ($PSVersionTable.PSVersion.Minor -lt 1))) {
    Write-Host "PowerShell version 5.1 or greater is required to run this script." -ForegroundColor Red
    Write-Host "Please find the download packages here: https://docs.microsoft.com/en-us/powershell/scripting/install/installing-windows-powershell?view=powershell-6#upgrading-existing-windows-powershell"
    Exit
}

# 

#$tenantName = "trywizdom07"
#$importUrl = "/sites/intranet"
#$targetUser = "admin@trywizdom07.onmicrosoft.com"
#$targetReachSubscription = "fe3f9bf2-bf77-4c9b-b02f-c39b5ddc66fe"

#.\importContent.ps1 -tenantName "trywizdom07" -importUrl "/sites/intranet" -targetUser "admin@trywizdom07.onmicrosoft.com" -targetReachSubscription "fe3f9bf2-bf77-4c9b-b02f-c39b5ddc66fe"
#.\importContent.ps1 -tenantName "trywizdom25" -importUrl "/sites/intranet" -targetUser "admin@trywizdom25.onmicrosoft.com" -targetReachSubscription "23cc5a8c-7457-41ee-9020-9e56d54be18b"

function ImportSite {
    param(
        [Parameter(Mandatory=$true)]
        [String]$siteUrl,
        [Parameter(Mandatory=$true)]
        [String]$siteName,
        [Parameter(Mandatory=$true)]
        [String]$targetUser
    )

    Write-Host "Importing $siteName ..."

    Connect-PnPOnline -Url $importUrl -Interactive

    ((Get-Content -Path "$siteName/$siteName-Template.xml" -Raw) -replace "USER_PLACEHOLDER", "$targetUser") | Set-Content -Path "$siteName/$siteName-TemplateReplaced.xml"

    Invoke-PnPSiteTemplate -Path "$siteName/$siteName-TemplateReplaced.xml"

    PublishPages -siteUrl $siteUrl -siteName $siteName

    Write-Host "Done"
}

function ImportTermGroup {
    param(
        [Parameter(Mandatory=$true)]
        [String]$siteUrl,
        [Parameter(Mandatory=$true)]
        [String]$targetUser
    )

    Write-Host "Importing term group $termGroupId ..."

    ((Get-Content -Path "Terms/Terms.xml" -Raw) -replace "USER_PLACEHOLDER", $targetUser) | Set-Content -Path "Terms/TermsReplaced.xml"

    Connect-PnPOnline -Url $siteUrl -Interactive
    
    Import-PnPTermGroupFromXml -Path "Terms/TermsReplaced.xml"
    Write-Host "Done"
}

function ConfigureSite {
    param(
        [Parameter(Mandatory=$true)]
        [String]$siteUrl
    )
    
    Write-Host "Configuring site ..."

    #Connect to PnP Online
    Connect-PnPOnline -Url $siteUrl -Interactive
 
    #Get the Site
    $Site = Get-PnPSite –Includes CustomScriptSafeDomains  
 
    #Add domain
    $Domain = [Microsoft.SharePoint.Client.ScriptSafeDomainEntityData]::new()
    $Domain.DomainName = "app.condense.ch"
    $null = $Site.CustomScriptSafeDomains.Create($Domain)
    Invoke-PnPQuery

    Set-PnPHomePage -RootFolderRelativeUrl "SitePages/HomeReachContent.aspx"

    Write-Host "Done"
}

function PublishPages {
    param(
        [Parameter(Mandatory=$true)]
        [String]$siteUrl,
        [Parameter(Mandatory=$true)]
        [String]$siteName
    )

    $pageDataArray = Get-Content -Path "$siteName/$siteName-pageData.json" | ConvertFrom-Json

    #Connect to PnP Online
    Connect-PnPOnline -Url $siteUrl -Interactive

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

function ImportDocuments {
    
    Write-Host "Importing demo documents ..."

    $null = Add-PnPFile -Path ".\Intranet\Team.tenant.xml" -Folder "PnPTemplates"
    $null = Add-PnPFile -Path ".\SourceFiles\Event Proposal Template.docx" -Folder "Shared Documents"
    $null = Add-PnPFile -Path ".\SourceFiles\LiveTiles - PPT Template.pptx" -Folder "Shared Documents"
    $null = Add-PnPFile -Path ".\SourceFiles\LiveTiles - The power of employee communication - eBook.pdf" -Folder "Shared Documents"
    $null = Add-PnPFile -Path ".\SourceFiles\LiveTiles Reach for Schools.pdf" -Folder "Shared Documents"
    $null = Add-PnPFile -Path ".\SourceFiles\LiveTiles_Reach_Evaluation_Guide.pdf" -Folder "Shared Documents"

    Write-Host "Done"
}

function UpdateJsonFiles {
    param(
        [Parameter(Mandatory=$true)]
        [String]$tenantName,
        [Parameter(Mandatory=$true)]
        [String]$targetSubscription
    )

    Write-Host "Updating json files ..."

    ((Get-Content -Path "JsonFiles/hub.json" -Raw) -replace "SUBSCRIPTION_PLACEHOLDER", $targetSubscription) | Set-Content -Path "JsonFiles/hubReplaced.json"
    ((Get-Content -Path "JsonFiles/hubReplaced.json" -Raw) -replace "TENANT_PLACEHOLDER", "$tenantName") | Set-Content -Path "JsonFiles/hubReplaced.json"
    ((Get-Content -Path "JsonFiles/siteType-Community.json" -Raw) -replace "TENANT_PLACEHOLDER", "$tenantName") | Set-Content -Path "JsonFiles/siteType-CommunityReplaced.json"
    ((Get-Content -Path "JsonFiles/siteType-Project.json" -Raw) -replace "TENANT_PLACEHOLDER", "$tenantName") | Set-Content -Path "JsonFiles/siteType-ProjectReplaced.json"
    ((Get-Content -Path "JsonFiles/siteType-Team.json" -Raw) -replace "TENANT_PLACEHOLDER", "$tenantName") | Set-Content -Path "JsonFiles/siteType-TeamReplaced.json"

    Write-Host "Done"
}

$tenantUrl = "https://$tenantName.sharepoint.com"
Connect-PnPOnline -Url $tenantUrl -Interactive

$importUrl = "$tenantUrl$importUrl"

$site = Get-PnPTenantSite -Identity $importUrl -ErrorAction SilentlyContinue

if($site -eq $null) {
    Write-Host "Site $importUrl does not exist. Create before continueing."
    Exit
}

ImportSite -siteUrl $importUrl -siteName "Intranet" -targetUser $targetUser
ImportSite -siteUrl $importUrl -siteName "News" -targetUser $targetUser
ImportSite -siteUrl $importUrl -siteName "Policies" -targetUser $targetUser
ImportSite -siteUrl $importUrl -siteName "Topics" -targetUser $targetUser
ImportDocuments

ConfigureSite -siteUrl $importUrl
ImportTermGroup -siteUrl $importUrl -targetUser $targetUser 
if($targetReachSubscription -ne $null){
    UpdateJsonFiles -tenantName $tenantName -targetSubscription $targetReachSubscription
}