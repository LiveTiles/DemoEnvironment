param (
    [Parameter(Mandatory=$true)]
    [string]$tenantName,
    [Parameter(Mandatory=$true)]
    [string]$siteUrl,
    [Parameter(Mandatory=$true)]
    [string]$SharePointTenantAdmin,
    [Parameter(Mandatory=$true)]
    [string]$SharePointTenantAdminPassword
)

function Get-LiveTilesJsonFiles {
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

    $url = "https://hub.matchpoint365.com/api/configs/default_hub/versions"
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

    $url = "https://hub.matchpoint365.com/api/configs/default_hub/versions"
    $url = "https://hub.matchpoint365.com/api/configs/default_hub"
    $body = ""
    $result = Invoke-RestMethod –Uri $url -Method Post -Headers $header -Body $body
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

$accessToken = Get-LiveTilesHubToken -SharePointTenantAdmin $SharePointTenantAdmin -SharePointTenantAdminPassword $SharePointTenantAdminPassword
$jsonConfig = Get-LiveTilesJsonFiles -accessToken $accessToken -tenantName $tenantName -siteUrl $siteUrl



# Use this to store the config as a file
$jsonConfig.json | ConvertTo-Json -Depth 100 -Compress | Set-Content -Path ".\original-hub.json"