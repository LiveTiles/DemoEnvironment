<#
.SYNOPSIS
    Imports demo configuration and for LiveTiles Intranet"

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

.PARAMETER SharePointTenantAdmin
    An intranet admin user. Needed for generating the access token.

.PARAMETER SharePointTenantAdminPassword
    Password for the SharePointTenantAdmin.

.PARAMETER targetReachSubscription
    (Optional) The reach subscription id for the target environment. If provided, this Id will be inserted into the relevant configuration files.

.EXAMPLE
    .\Import-LiveTilesConfiguration -tenantName "TryLiveTilesXX" -siteUrl "/sites/intranet" -SharePointTenantAdmin "admin@trylivetilesxx.onmicrosoft.com" -SharePointTenantAdminPassword "P@ssw0rd" -targetReachSubscription "fe3f9bf2-bf77-4c9b-b02f-c39b5ddc66fe"
    
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
    [string]$SharePointTenantAdmin,
    [Parameter(Mandatory=$true)]
    [string]$SharePointTenantAdminPassword,
    [Parameter(Mandatory=$false)]
    [string]$targetReachSubscription = "SUBSCRIPTION_PLACEHOLDER"
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
    return $result
}

function Set-LiveTilesIntranetConfig {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$accessToken,
        [Parameter(Mandatory=$true)]
        [string]$tenantName,
        [Parameter(Mandatory=$true)]
        [string]$siteUrl
    )
    
    Write-Host "Updating LiveTiles intranet config ..."

    # Static variables
    $url = "https://hub.matchpoint365.com/api/configs/default_hub"

    # Get the existing config data
    $jsonConfig = Get-LiveTilesIntranetConfig -accessToken $accessToken -tenantName $tenantName -siteUrl $siteUrl

    # Use this to get the new config
    $newConfig = Get-Content -Path .\JsonFiles\$tenantName-hub.json | ConvertFrom-Json
        
    # Merge the new config into a new object
    $bodyObj = @{
        'key' = $jsonConfig.key
        'changeToken' = $jsonConfig.changeToken
        'appVersion' = $jsonConfig.appVersion
        'json' = $newConfig
    }
    
    # Convery the body to json
    $body = $bodyObj | ConvertTo-Json -Depth 100 -Compress

    # Set the header, including the body length    
    $header = @{
        'accept' = '*/*'
        'accept-encoding' = 'gzip, deflate, br'
        'accept-language' = 'en-US,en;q=0.9'
        'Authorization' = 'Bearer ' + $accessToken
        'cache-control' = 'no-cache'
        'client-url' = 'https://$($tenantName).sharepoint.com$($siteUrl)'
        'Content-Length' = $body.Length
        'content-type' = 'application/json'
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

    # Post the config to hub
    $result = Invoke-RestMethod –Uri $url -Method Post -Headers $header -Body $body

    Write-Host "Done"
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
        [String]$tenantName,
        [Parameter(Mandatory=$true)]
        [String]$targetSubscription
    )

    Write-Host "Updating json files ..."

    ((Get-Content -Path "JsonFiles/original-hub.json" -Raw) -replace "SUBSCRIPTION_PLACEHOLDER", $targetSubscription) | Set-Content -Path "JsonFiles/$tenantName-hub.json"
    ((Get-Content -Path "JsonFiles/$tenantName-hub.json" -Raw) -replace "TENANT_PLACEHOLDER", "$tenantName") | Set-Content -Path "JsonFiles/$tenantName-hub.json"
    ((Get-Content -Path "JsonFiles/original-siteType-Community.json" -Raw) -replace "TENANT_PLACEHOLDER", "$tenantName") | Set-Content -Path "JsonFiles/$tenantName-siteType-Community.json"
    ((Get-Content -Path "JsonFiles/original-siteType-Project.json" -Raw) -replace "TENANT_PLACEHOLDER", "$tenantName") | Set-Content -Path "JsonFiles/$tenantName-siteType-Project.json"
    ((Get-Content -Path "JsonFiles/original-siteType-Team.json" -Raw) -replace "TENANT_PLACEHOLDER", "$tenantName") | Set-Content -Path "JsonFiles/$tenantName-siteType-Team.json"
    (Get-Content -Path "JsonFiles/original-metadata-Department.json" -Raw) | Set-Content -Path "JsonFiles/$tenantName-metadata-Department.json"
    (Get-Content -Path "JsonFiles/original-metadata-Project.json" -Raw) | Set-Content -Path "JsonFiles/$tenantName-metadata-Project.json"
    (Get-Content -Path "JsonFiles/original-metadata-Team.json" -Raw) | Set-Content -Path "JsonFiles/$tenantName-metadata-Team.json"

    Write-Host "Done"
}



#$pnpSharePointAdminContext = Connect-PnPOnline -ReturnConnection -Url $SharePointAdminUrl -Credentials $credentialsOffice365Admin

# Execution begins here

Update-LiveTilesJsonFiles -tenantName $tenantName -targetSubscription $targetReachSubscription

# Get access token for updating LiveTiles config
$accessToken = Get-LiveTilesHubToken -SharePointTenantAdmin $SharePointTenantAdmin -SharePointTenantAdminPassword $SharePointTenantAdminPassword

Set-LiveTilesIntranetConfig -accessToken $accessToken -tenantName $tenantName -siteUrl $siteUrl