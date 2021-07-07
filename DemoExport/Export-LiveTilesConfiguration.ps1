<#
.SYNOPSIS
    Exports demo configuration and for LiveTiles Intranet"

.DESCRIPTION
    Requirements:
    - Powershell v5
    - PnP.PowerShell, AzureAD

    All modules will be validated and potentially updated when the script runs.

    The script will perform the following actions:
    - Exports the LiveTiles Intranet Configuration as Json

    With the -WhatIf switch, only the testing and verification of the above actions will be performed. No CREATE actions are actually performed.

.PARAMETER tenantName
    The source tenant name to be exported.

.PARAMETER siteUrl
    The site relative url to the main site collection e.g. /sites/intranet

.PARAMETER sourceReachSubscription
    The reach subscription id for the source environment.

.PARAMETER SharePointTenantAdmin
    An intranet admin user. Needed for generating the access token.

.PARAMETER SharePointTenantAdminPassword
    Password for the SharePointTenantAdmin.

.EXAMPLE
    .\Export-LiveTilesConfiguration -tenantName "TryLiveTilesXX" -siteUrl "/sites/intranet" -SharePointTenantAdmin "admin@trylivetilesxx.onmicrosoft.com" -SharePointTenantAdminPassword "P@ssw0rd" 
    
    Exports demo configuration for the specified tenant.  

.NOTES
    AUTHOR: Garry Sinclair
    LASTEDIT: 02-07-2021 
    v1.0
        First Release. Identified 2Do´s:
            - Automatically export the LiveTiles Metadata and Workspaces Json configuration.

.LINK
    Updated versions of this script will be available on the LiveTiles Partner Portal
#>
param (
    [Parameter(Mandatory=$true)]
    [string]$tenantName,
    [Parameter(Mandatory=$true)]
    [string]$siteUrl,
    [Parameter(Mandatory=$true)]
    [string]$sourceReachSubscription,
    [Parameter(Mandatory=$true)]
    [string]$SharePointTenantAdmin,
    [Parameter(Mandatory=$true)]
    [string]$SharePointTenantAdminPassword
)

function Get-LiveTilesIntranetConfig {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$accessToken,
        [Parameter(Mandatory=$true)]
        [string]$tenantName,
        [Parameter(Mandatory=$true)]
        [string]$siteUrl
    )

    Write-Host "Exporting LiveTiles Intranet Config..."

    $header = @{
        'accept-encoding' = 'gzip, deflate, br'
        'accept' = '*/*'
        'accept-language' = 'en-US,en;q=0.9'
        'Authorization' = 'Bearer ' + $accessToken
        'cache-control' = 'no-cache'
        'client-url' = 'https://$($tenantName).sharepoint.com$($siteUrl)'
        'Host' = 'hub.matchpoint365.com'
        'origin' = 'https://$($tenantName).sharepoint.com'
        'pragma' = 'no-cache'
        'referer' = 'https://$($tenantName).sharepoint.com/'
        'sec-ch-ua' = '" Not;A Brand";v="99", "Google Chrome";v="91", "Chromium";v="91"'
        'sec-ch-ua-mobile' = '?0'
        'sec-fetch-dest' = 'empty'
        'sec-fetch-mode' = 'cors'
        'sec-fetch-site' = 'cross-site'
        'user-agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36'
    }

    $url = "https://hub.matchpoint365.com/api/configs/default_hub"
    $result = Invoke-RestMethod –Uri $url -Method Get -Headers $header

    Write-Host "Done ..."

    return $result
}

function Get-LiveTilesMetadataConfig {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$accessToken,
        [Parameter(Mandatory=$true)]
        [string]$tenantName,
        [Parameter(Mandatory=$true)]
        [string]$siteUrl
    )
    $header = @{
        'accept-encoding' = 'gzip, deflate, br'
        'accept' = '*/*'
        'accept-language' = 'en-US,en;q=0.9'
        'Authorization' = 'Bearer ' + $accessToken
        'cache-control' = 'no-cache'
        'client-url' = 'https://$($tenantName).sharepoint.com$($siteUrl)'
        'Host' = 'metadata.matchpoint365.com'
        'origin' = 'https://$($tenantName).sharepoint.com'
        'pragma' = 'no-cache'
        'referer' = 'https://$($tenantName).sharepoint.com/'
        'sec-ch-ua' = '" Not;A Brand";v="99", "Google Chrome";v="91", "Chromium";v="91"'
        'sec-ch-ua-mobile' = '?0'
        'sec-fetch-dest' = 'empty'
        'sec-fetch-mode' = 'cors'
        'sec-fetch-site' = 'cross-site'
        'user-agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36'
    }

    $url = "https://metadata.matchpoint365.com/api/configs"
    $result = Invoke-RestMethod –Uri $url -Method Get -Headers $header
    return $result
}

function Get-LiveTilesHubToken {
    Param(
        [Parameter(Mandatory=$true)][string]$SharePointTenantAdmin,
        [Parameter(Mandatory=$true)][string]$SharePointTenantAdminPassword
    )
      
    $clientId = "https://iurcycl.onmicrosoft.com/14a6b046-c3d9-4988-aa99-f0804587f299"
    $resource = "https://iurcycl.onmicrosoft.com/14a6b046-c3d9-4988-aa99-f0804587f299"

    $AADapiToken = (Invoke-RestMethod "https://login.windows.net/common/oauth2/token" -Method POST -Body "resource=$($resource)&grant_type=password&client_id=$($clientId)&username=$($SharePointTenantAdmin)&password=$($SharePointTenantAdminPassword)").access_token
    return $AADapiToken
}

function Update-LiveTilesJsonFiles {
    param(
        [Parameter(Mandatory=$true)]
        [String]$sourceTenant,
        [Parameter(Mandatory=$true)]
        [String]$sourceSubscription
    )

    Write-Host "Updating hub.json - Replacing source subscription placeholder ..."
    ((Get-Content -Path "../DemoImport/JsonFiles/original-hub.json" -Raw) -replace $sourceSubscription, "SUBSCRIPTION_PLACEHOLDER") | Set-Content -Path "../DemoImport/JsonFiles/original-hub.json"
    Write-Host "Done ..."

    Write-Host "Updating hub.json - Replacing $sourceTenant with placeholder ..."
    ((Get-Content -Path "../DemoImport/JsonFiles/original-hub.json" -Raw) -replace "$sourceTenant", "TENANT_PLACEHOLDER") | Set-Content -Path "../DemoImport/JsonFiles/original-hub.json"
    Write-Host "Done ..."

    Write-Host "Updating siteType-Community.json - Replacing $sourceTenant with placeholder ..."
    ((Get-Content -Path "../DemoImport/JsonFiles/original-siteType-Community.json" -Raw) -replace "$sourceTenant", "TENANT_PLACEHOLDER") | Set-Content -Path "../DemoImport/JsonFiles/original-siteType-Community.json"
    Write-Host "Done ..."

    Write-Host "Updating siteType-Project.json - Replacing $sourceTenant with placeholder ..."
    ((Get-Content -Path "../DemoImport/JsonFiles/original-siteType-Project.json" -Raw) -replace "$sourceTenant", "TENANT_PLACEHOLDER") | Set-Content -Path "../DemoImport/JsonFiles/original-siteType-Project.json"
    Write-Host "Done ..."

    Write-Host "Updating siteType-Team.json - Replacing $sourceTenant with placeholder ..."
    ((Get-Content -Path "../DemoImport/JsonFiles/original-siteType-Team.json" -Raw) -replace "$sourceTenant", "TENANT_PLACEHOLDER") | Set-Content -Path "../DemoImport/JsonFiles/original-siteType-Team.json"
    Write-Host "Done ..."
}


# Execution begins here


$accessToken = Get-LiveTilesHubToken -SharePointTenantAdmin $SharePointTenantAdmin -SharePointTenantAdminPassword $SharePointTenantAdminPassword
$jsonConfig = Get-LiveTilesIntranetConfig -accessToken $accessToken -tenantName $tenantName -siteUrl $siteUrl

# Use this to store the config as a file
$jsonConfig.json | ConvertTo-Json -Depth 100 |  Set-Content -Path "..\DemoImport\JsonFiles\original-hub.json"

# Replace source tenant name with place holder
Update-LiveTilesJsonFiles -sourceTenant $tenantName -sourceSubscription $sourceReachSubscription
