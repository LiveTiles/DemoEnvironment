param (
    [Parameter(Mandatory=$true)]
    [string]$sourceTenant,
    [Parameter(Mandatory=$true)]
    [string]$sourceUrl,
    [Parameter(Mandatory=$true)]
    [string]$SharePointTenantAdmin,
    [Parameter(Mandatory=$false)]
    [string]$SharePointTenantAdminPassword
)
function Get-LiveTilesJsonFiles {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$accessToken,
        [Parameter(Mandatory=$true)]
        [string]$tenantName
    )
    $header = @{
        'accept-encoding' = 'gzip, deflate, br'
        'accept' = '*/*'
        'accept-language' = 'en-US,en;q=0.9'
        'Authorization' = 'Bearer ' + $accessToken
        'cache-control' = 'no-cache'
        'client-url' = 'https://$($tenantName).sharepoint.com/sites/intranet#/'
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

#        'Connection' = 'keep-alive'
    #        'subscription-id' = 'fe3f9bf2-bf77-4c9b-b02f-c39b5ddc66fe'
    #    'x-condense-client' = 'web'
    #    'x-condense-client-buildid' = '67064'
    #    'x-condense-client-version' = '0.2.415'
    #    'x-condense-preferred-language' = 'en'
    #    'x-condense-preferred-timezone' = 'Europe/Copenhagen'

        

    $url = "https://hub.matchpoint365.com/api/configs/default_hub/versions"
    Invoke-RestMethod –Uri $url -Method Get -Headers $header
}

function Get-LiveTilesHubToken {
    Param(
        [Parameter(Mandatory=$true)][string]$tenantId, 
        [Parameter(Mandatory=$true)][string]$SharePointTenantAdmin,
        [Parameter(Mandatory=$true)][string]$SharePointTenantAdminPassword
    )
    
    $AADapiToken = (Invoke-RestMethod "https://login.windows.net/$($tenantId)/oauth2/token" -Method POST -Body "resource=2d28d563-0baf-49a8-acf3-aacac75d9a00&grant_type=password&client_id=14a6b046-c3d9-4988-aa99-f0804587f299&scope=openid&username=$($SharePointTenantAdmin)&password=$($SharePointTenantAdminPassword)").access_token
    return $AADapiToken
}

Function Grant-OAuth2PermissionsToApp () #Mimics "Grant Permissions" button when setting permissions on the AAD App
{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param(
        [Parameter(Mandatory=$true)][string]$azureAppId, #application ID of the azure application you wish to admin-consent to
        [Parameter(Mandatory=$true)][string]$apiToken
    )
    
    $header = @{
    'Authorization' = 'Bearer ' + $apiToken
    'X-Requested-With'= 'XMLHttpRequest'
    'x-ms-client-request-id'= [guid]::NewGuid()
    'x-ms-correlation-id' = [guid]::NewGuid()}
    $url = "https://main.iam.ad.ext.azure.com/api/RegisteredApplications/$azureAppId/Consent?onBehalfOfAll=true"
    do {
        try {
        Invoke-RestMethod –Uri $url –Headers $header –Method POST # -ErrorAction Stop
            $completed = $true} catch {
            $completed = $false
            Start-Sleep -Seconds 10}
    } until ($completed)
}

$tenantUrl = "https://$sourceTenant.sharepoint.com"
Connect-PnPOnline -Url "$tenantUrl$sourceUrl" -Interactive
$tenantId = Get-PnPTenantId
$accessToken = Get-LiveTilesHubToken -tenantId $tenantId -SharePointTenantAdmin $SharePointTenantAdmin -SharePointTenantAdminPassword $SharePointTenantAdminPassword
Get-LiveTilesJsonFiles -accessToken $accessToken -tenantName $sourceTenant