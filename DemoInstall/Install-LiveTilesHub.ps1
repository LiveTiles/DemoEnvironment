<#
.SYNOPSIS
    Deploys LiveTiles Reach, Hub and Hub modules to SharePoint Online"

.DESCRIPTION
    Requirements:
    - Powershell v5
    - PnP.PowerShell, AzureAD and AzureRM.profile or Az.Accounts

    All modules will be validated and potentially updated when the script runs.

    The script will perform the following actions:
    - Install or Update required PowerShell Modules and components
    - Create the Application Catalog if it doesn't exist
    - Download the LiveTiles Packages from LiveTiles and upload them to the SharePoint App Catalog
    - Import Tenant level search configuration (managed properties)
    - Create LiveTiles Service Principals
    - Create Livetiles Subscriptions
    - Add cross LiveTiles Service Principal permissions    
    - Approve API requests in SharePoint
    - Add the LiveTiles Apps to the root site collection (or HubSiteCollection if that parameter is used)
    - Add the required webparts to the new Hub frontpage to make the Hub modules visible in admin.

    With the -WhatIf switch, only the testing and verification of the above actions will be performed. No CREATE actions are actually performed.

.PARAMETER UseWebLogin
    (Optional) (Switch) Information that the script will be running in Multifactor Authentication mode for the SharePoint Tenant Admin User.

.PARAMETER SharePointTenantAdmin
    The username of a user with Tenant Admin rights for Office 365 (including admin rights to SharePoint)

.PARAMETER SharePointTenantAdminPassword
    (Optional) The password of the SharePoint Tenant Admin user. For unattended execution. If not supplied the script will query for it as a secure string

.PARAMETER o365Tenant
    (Optional) If the SharePointTenantAdmin is supplied and is a .onmicrosoft.com adress, the tenant will be derived from that - if not, this parameter is required.

.PARAMETER hubSiteCollection 
    (Optional) Hub will be installed on this sitecollection, that is presumed to exist. If not, the script will fail. Default install is the root site.

.PARAMETER unInstall
    (Optional) (Switch) Removes the LiveTiles installation from the m365 tenant

.PARAMETER noGovernance
    (Optional) (Switch) Excludes the Governance module from the installation

.PARAMETER noMetadata
    (Optional) (Switch) Excludes the Metadata module from the installation

.PARAMETER noWorkspaces
    (Optional) (Switch) Excludes the Workspaces module from the installation

.PARAMETER noEverywhere
    (Optional) (Switch) Excludes the Everywhere module from the installation

.PARAMETER Reach
    (Optional) (Choice) Decide whether to create a Reach instance, and what demo content to include
        - NoInstall
        - DefaultDemo (default)
        - BroadcastLtd
        - CareNet
        - FoodInc
      

.PARAMETER EnableCDN
    (Optional) (Switch) Automatically enables O365 CDN if it's not already enabled

.EXAMPLE
    .\Install-LiveTilesHub -UseWebLogin
    
    Creates Reach Subscription and Installs Hub and all modules on the root site collection using MFA enabled login 

.EXAMPLE
    .\Install-LiveTilesHub -SharePointTenantAdmin john.doe@fabricam.onmicrosoft.com

    Creates Reach Subscription and Installs Hub and all modules on the root site collection using the John Doe credentials. Will prompt for password. 
    
.EXAMPLE
    .\Install-LiveTilesHub -SharePointTenantAdmin john.doe@fabricam.com -o365Tenant "othertenant" -hubSiteCollection "livetileshub"
    
    Creates Reach Subscription and Installs Hub and all modules on <othertenant>/sites/<hubsitecollection> using the John Doe credentials. Will prompt for password. 

.EXAMPLE
    .\Install-LiveTilesHub -SharePointTenantAdmin john.doe@fabricam.com -unInstall
    
    UnInstalls Hub and all modules using the John Doe credentials. Will prompt for password. 

.EXAMPLE
    .\Install-LiveTilesHub -SharePointTenantAdmin john.doe@fabricam.com -Reach NoInstall
    
    Installs Hub and all modules on the root site collection using the John Doe credentials. Will prompt for password. Reach will not be installed

.EXAMPLE
    .\Install-LiveTilesHub -SharePointTenantAdmin john.doe@fabricam.com -Reach BroadcastLtd
    
    Installs Hub and all modules on the root site collection using the John Doe credentials. Will prompt for password. Reach will be provisioned with the "Broadcast Ltd" demo set

.NOTES
    AUTHOR: Christoffer Soltau
    LASTEDIT: 30-06-2021 
    v1.0
        First Release. Identified 2Do´s:
            - Actually test the MFA process :-)
	        - After search config import, Better error message - perhaps a description of what needs to be done manually?
                - Can't see that there's a reason it would fail with current settings - no changes made to custom managed properties where it usually fails.
	        - Creation of Governance license creation. Untested runtime - may have to be placed before subscription link??
            - Check if user is app cat admin - otherwise we can't upload packages
            - Check if user is site coll admin - otherwise we can't add apps.

.LINK
    Updated versions of this script will be available on the LiveTiles Partner Portal
#>
[cmdletbinding(SupportsShouldProcess=$True)]
param (
    #2Do - Actually test the MFA process :-)
    [Parameter(Mandatory=$true, ParameterSetName="MFA.uninstall")]
    [Parameter(Mandatory=$true, ParameterSetName="MFA.partial.install")]
    [Parameter(Mandatory=$true, ParameterSetName="MFA.full.install")]
    [switch]$UseWebLogin,
    [Parameter(Mandatory=$false, ParameterSetName = "nonMFA.uninstall")]
    [Parameter(Mandatory=$false, ParameterSetName="nonMFA.partial.install")]
    [Parameter(Mandatory=$false, ParameterSetName="nonMFA.full.install")]
    [string]$SharePointTenantAdmin = "admin@M365x510994.onmicrosoft.com",
    [Parameter(Mandatory=$false, ParameterSetName = "nonMFA.uninstall")]
    [Parameter(Mandatory=$false, ParameterSetName="nonMFA.partial.install")]
    [Parameter(Mandatory=$false, ParameterSetName="nonMFA.full.install")]
    [string]$SharePointTenantAdminPassword = "Rd6E0x7e1N",
    [Parameter(Mandatory=$false)]
    [string]$O365Tenant = "",
    [Parameter(Mandatory=$false)]
    [string]$HubSiteCollection, #if not used, assume root site coll?
    [Parameter(Mandatory=$true, ParameterSetName="MFA.uninstall")]
    [Parameter(Mandatory=$true, ParameterSetName="nonMFA.uninstall")]
    [switch]$Uninstall,
    [Parameter(Mandatory=$false, ParameterSetName="MFA.partial.install")]
    [Parameter(Mandatory=$false, ParameterSetName="nonMFA.partial.install")]
    [switch]$NoGovernance,
    [Parameter(Mandatory=$false, ParameterSetName="MFA.partial.install")]
    [Parameter(Mandatory=$false, ParameterSetName="nonMFA.partial.install")]
    [switch]$NoMetadata,
    [Parameter(Mandatory=$false, ParameterSetName="MFA.partial.install")]
    [Parameter(Mandatory=$false, ParameterSetName="nonMFA.partial.install")]
    [switch]$NoWorkspaces,
    [Parameter(Mandatory=$false, ParameterSetName="MFA.partial.install")]
    [Parameter(Mandatory=$false, ParameterSetName="nonMFA.partial.install")]
    [switch]$NoEverywhere,
    [Parameter(Mandatory=$false, ParameterSetName="MFA.partial.install")]
    [Parameter(Mandatory=$false, ParameterSetName="nonMFA.partial.install")]
    [switch]$NoProvisioning,
    [Parameter(DontShow, Mandatory=$true, ParameterSetName="MFA.full.install")]
    [Parameter(DontShow, Mandatory=$true, ParameterSetName="nonMFA.full.install")]
    [switch]$FullInstall = $True,
    [switch]$EnableCDN,
    [ValidateSet('NoInstall', 'DefaultDemo', 'BroadcastLtd', 'CareNet', 'FoodInc')]$Reach = "DefaultDemo"
)

if (($PSVersionTable.PSVersion.Major -lt 5) -or (($PSVersionTable.PSVersion.Major -eq 5) -and ($PSVersionTable.PSVersion.Minor -lt 1))) {
    Write-Host "PowerShell version 5.1 or greater is required to run this script." -ForegroundColor Red
    Write-Host "Please find the download packages here: https://docs.microsoft.com/en-us/powershell/scripting/install/installing-windows-powershell?view=powershell-6#upgrading-existing-windows-powershell"
    Exit
}

#region Support functions
Function Write-LiveTilesHost () #Helper Function: write formatted output
{
    [CmdletBinding(SupportsShouldProcess=$True)]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet('PROCESS', 'PROCESSATTENTION', 'OK', 'WARNING', 'ERROR')]$messageType,
        [Parameter(Mandatory=$false)]
        [String]$message,
        [Parameter(Mandatory=$false)]
        [int]$outputMaxLength,
        [Parameter(Mandatory=$false)]
        [String]$afterOutputMessage,
        [Parameter(Mandatory=$false)]
        [int]$initialStringLength

    )
        switch ($messageType) {
        PROCESS {
            if ($WhatIfPreference) {
                Write-Host $message -ForegroundColor DarkGray
            } else {
                Write-Host $message -ForegroundColor White -NoNewline
            }
            return $message.Length
        }
        PROCESSATTENTION {
            if ($WhatIfPreference) {
                Write-Host $message -ForegroundColor DarkYellow
            } else {
                Write-Host $message -ForegroundColor Yellow -NoNewline
            }
            return $message.Length
        }
        OK {
            if (-not $WhatIfPreference) {
                if ($message -ne "") {
                    Write-Host $message -ForegroundColor Yellow -NoNewline
                }
                Write-Host ("." * [math]::max(($outputMaxLength - $initialStringLength - $message.Length),0)) -ForegroundColor White -NoNewline
                Write-Host "OK" -ForegroundColor Green
            }
        }
        WARNING {
            if (-not $WhatIfPreference) {
                if ($message -ne "") {
                    Write-Host $message -ForegroundColor Yellow -NoNewline
                }
                Write-Host ("." * [math]::max(($outputMaxLength - $initialStringLength - $message.Length - 5),0)) -ForegroundColor White -NoNewline
                Write-Host "WARNING" -ForegroundColor Yellow
                if ($afterOutputMessage -ne "") {
                    Write-Host $afterOutputMessage -ForegroundColor White            
                }
            }
        }
        ERROR {
            if (-not $WhatIfPreference) {
                if ($message -ne "") {
                    Write-Host $message -ForegroundColor Red -NoNewline
                }
                Write-Host ("." * [math]::max(($outputMaxLength - $initialStringLength - $message.Length - 3),0)) -ForegroundColor White -NoNewline
                Write-Host "ERROR" -ForegroundColor Red
                if ($afterOutputMessage -ne "") {
                    Write-Host $afterOutputMessage -ForegroundColor White
                }
            }
        }
    }
}

Function pause ($message = "Press any key to continue...") #Helper Function: "Press anykey to continue"
{
    # Check if running in PowerShell ISE
    If ($psISE) {
        # "ReadKey" not supported in PowerShell ISE.
        # Show MessageBox UI
        $Shell = New-Object -ComObject "WScript.Shell"
        $Button = $Shell.Popup("Click OK to continue.", 0, "LiveTiles Installer", 0)
        Return
    } else {
 
    $Ignore =
        16,  # Shift (left or right)
        17,  # Ctrl (left or right)
        18,  # Alt (left or right)
        20,  # Caps lock
        91,  # Windows key (left)
        92,  # Windows key (right)
        93,  # Menu key
        144, # Num lock
        145, # Scroll lock
        166, # Back
        167, # Forward
        168, # Refresh
        169, # Stop
        170, # Search
        171, # Favorites
        172, # Start/Home
        173, # Mute
        174, # Volume Down
        175, # Volume Up
        176, # Next Track
        177, # Previous Track
        178, # Stop Media
        179, # Play
        180, # Mail
        181, # Select Media
        182, # Application 1
        183  # Application 2
 
        Write-Host -NoNewline $Message
        While ($KeyInfo.VirtualKeyCode -Eq $Null -Or $Ignore -Contains $KeyInfo.VirtualKeyCode) {
            $KeyInfo = $Host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")
        }
    }
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

Function Get-PortalAPIAccessToken {
    #piggybacks on connect-azuread to get user login and dll
    param(
        [string]$ClientId,
        [string]$RedirectUri = "urn:ietf:wg:oauth:2.0:oob",
        [string]$TenantId,
        [string]$Resource
    )
    $authority = "https://login.microsoftonline.com/$($TenantId)"
    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority, $false
    $PromptBehavior = [Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::RefreshSession
    $PlatformParameters = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList $PromptBehavior
    $authResult = $authContext.AcquireTokenAsync($Resource, $ClientId, $redirectUri, $PlatformParameters).Result
    $token = $authResult.AccessToken
    return $token;
}
#endregion

$dateStart = get-date
Write-Host ("Installation starttime: " + $dateStart) -ForegroundColor White

#region Set Constants
    $waitVeryLong = 60
    $waitLong = 10
    $waitShort = 5

    $tenantDomain = $SharePointTenantAdmin.Split("@")[1].split(".")
    if ($tenantDomain[1] -eq "onmicrosoft") {
        $SharePointAdminUrl = "https://" + $tenantDomain[0] + "-admin.sharepoint.com"
        $SharePointUrl = "https://" + $tenantDomain[0] + ".sharepoint.com"
    } elseif ($o365Tenant -ne "") {
        if ($o365Tenant -match "^https?://(.*?\.)+?") {
            $o365Tenant = $o365Tenant.split("/")[2].split(".")[0]
        }
        $SharePointAdminUrl = "https://" + $o365Tenant + "-admin.sharepoint.com"
        $SharePointUrl = "https://" + $o365Tenant + ".sharepoint.com"
    } else {
        Write-Host "Tenant url missing, please use the o365Tenant parameter. Exiting installation script."
        Exit
    }

    if (($hubSiteCollection -eq "") -or ($hubSiteCollection -eq $null) -or $hubSiteCollection -eq $SharePointUrl) {
        $hubSiteCollection = $SharePointUrl
    } else {
        $hubSiteCollection = $SharePointUrl + "/sites/" + $hubSiteCollection.Split("/")[-1]
    }

    $outputMaxLength = 45 + $SharePointTenantAdmin.Length #"Logging in to SharePoint Admin as admin@M365x510994.onmicrosoft.com"
    $outputMaxUninstall = 46 + $hubSiteCollection.Length #"Removing local LiveTiles Apps from https://m365x510994.sharepoint.com/"
    if ($outputMaxLength -lt $outputMaxUninstall) {$outputMaxLength = $outputMaxUninstall}
    if ($outputMaxLength -lt 84) {$outputMaxLength = 84}  #"Getting LiveTiles installation files and deploying packages to SharePoint"
 
    $appIdsLiveTiles = @()
        $appIdsLiveTiles += "14a6b046-c3d9-4988-aa99-f0804587f299"
        if ($Reach -ne "NoInstall") {$appIdsLiveTiles += "d492530a-8cff-481c-90da-9c3c3f1be7da"}
        if ($Reach -ne "NoInstall") {$appIdsLiveTiles += "e40ee8da-4b99-4273-b99f-a6b2f10ac29e"}
        if ($Reach -ne "NoInstall") {$appIdsLiveTiles += "02de0a9e-a2a9-42af-a6c3-2ed77e913a4b"}
        if (-not $noGovernance.IsPresent) {$appIdsLiveTiles += "b6a4f91d-466c-4afe-bb29-3f41ea615da7"}
        if (-not $noMetadata.IsPresent) {$appIdsLiveTiles += "337ea7aa-6b44-4d3c-b5c3-52aaeb1d5dd8"}
        if (-not $noProvisioning.IsPresent) {$appIdsLiveTiles += "3b8f4e7e-93b5-43ed-830d-eada7c8ff81f"}
        if (-not $noWorkspaces.IsPresent) {$appIdsLiveTiles += "a9a0e8f6-2cee-42d2-b00f-c9299e509958"}
        #MatchPoint Hub AppId "14a6b046-c3d9-4988-aa99-f0804587f299"
        #Reach AppId "d492530a-8cff-481c-90da-9c3c3f1be7da"
        #Condense API AppId "e40ee8da-4b99-4273-b99f-a6b2f10ac29e"
        #Condense AppId "02de0a9e-a2a9-42af-a6c3-2ed77e913a4b"
        #MatchPoint Governance AppId "b6a4f91d-466c-4afe-bb29-3f41ea615da7"
        #MatchPoint Metadata AppId "337ea7aa-6b44-4d3c-b5c3-52aaeb1d5dd8"
        #MatchPoint Provisioning AppId "3b8f4e7e-93b5-43ed-830d-eada7c8ff81f"
        #MatchPoint Workspaces AppId "a9a0e8f6-2cee-42d2-b00f-c9299e509958"
#endregion

#region Install or Update required PowerShell Modules and components
    #if (-not $skipCheckPrerequisites.IsPresent) {
        $output = Write-LiveTilesHost -messageType PROCESS -message ("Checking Powershell Module prerequisites")
        $dotCount = 0
        $moduleList = @(("PnP.PowerShell", "1.6.0"), ("AzureAD", "2.0.2.135"), ("Microsoft.Online.SharePoint.Powershell", "16.0.21213.12000"))#, ("PowerShellGet","2.2.5"), ("MSAL.PS", "4.21.0.1")) #, ("Az.Accounts","2.3.0"), ("AzureRM.profile","5.8.4"))
        $availableModules = Get-Module
        foreach ($module in $moduleList) {
            if (($output + $dotCount) -le ($outputMaxLength - 2)) {
                Write-Host -NoNewline "." -ForegroundColor White
                $dotCount++
            }
            $moduleOnline = Find-Module -Name $module[0] -MinimumVersion $module[1]
            $moduleInstalled = @(Get-Module -ListAvailable -Name $module[0])
            if (-not $moduleInstalled -and $module[0] -eq "AzureAD") {
                $moduleInstalled = @(Get-Module -ListAvailable -Name "AzureADPreview")
                if ($moduleInstalled) {
                    $module[0] = "AzureADPreview"
                    $module[1] = "2.0.2.136"
                }
            }

            if ($moduleInstalled) {
	            if ($moduleInstalled[0].Version -eq $moduleOnline.Version) {
		            Write-Verbose "$($module[0]) module is installed and latest version" #-foregroundcolor green
	            } elseif ($moduleInstalled[0].Version -ge [Version]$module[1]) {
                    Write-Verbose "$($module[0]) module is installed and of a later or equal version than required for the script" #-foregroundcolor yellow
                } else {
                    Write-Host
                    $inputString = "Confirm`n`r" + $module[0] + " module is currently version " + $moduleInstalled[0].Version.ToString() + ". The latest version is " + $moduleOnline.Version.ToString() + ", and $($Module[1]) is required" + " - do you want to Update?`n`rThe process will open in a new window to instal as 'admin'`r`n[Y] Yes  [N] No (default is 'Y')"
                    $input = Read-Host -prompt $inputString
                    if ($input -eq $null) {$input = "y"}
                    if ($input.ToString().ToLower() -ne "n") { 
                        Remove-Module $module[0]
                        start-process powershell -Verb RunAs "try {Install-Module $($module[0]) -AllowClobber -Force} catch {Install-Module $($module[0]) -AllowClobber -Force -SkipPublisherCheck}" -wait
                    }
	            }
            } else {
                Write-Host
                $inputString = "Confirm`n`r" + $module[0] + " module is not installed. Do you want to Install?`n`rThe process will open in a new window to instal as 'admin'`r`n[Y] Yes  [N] No (default is 'Y')"
                $input = Read-Host -prompt $inputString
                if ($input -eq $null) {$input = "y"}
                if ($input.ToString().ToLower() -ne "n") { 
    		        start-process powershell -Verb RunAs "try {Install-Module $($module[0]) -AllowClobber -Force} catch {Install-Module $($module[0]) -AllowClobber -Force -SkipPublisherCheck}"
                } else {
                    Write-Host "Module prerequisites not met. Exiting installation script."
                    Exit
                }
            }
        }
        Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength ($output+$dotCount) -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
#   }
#endregion

#region Logging in to SharePoint Admin via PnP

    try {
        if (-not $UseWebLogin.IsPresent) {
            $output = Write-LiveTilesHost -messageType PROCESS -message ("Logging in to SharePoint Admin using PnP as " + $SharePointTenantAdmin)
                
                if ($SharePointTenantAdminPassword -eq "") {
                    Write-Host
                    $SPOsecurePassword = Read-Host -Prompt ("Please enter the password for " + $SharePointTenantAdmin) -AsSecureString
                } else {$SPOsecurePassword = $SharePointTenantAdminPassword | ConvertTo-SecureString -AsPlainText -Force}
                $credentialsOffice365Admin = New-Object System.Management.Automation.PSCredential($SharePointTenantAdmin,$SPOsecurePassword)
                
                try {
                    $pnpSharePointAdminContext = Connect-PnPOnline -ReturnConnection -Url $SharePointAdminUrl -Credentials $credentialsOffice365Admin
                } catch {
                    if ($Error.CategoryInfo.Reason -eq "MsalUiRequiredException") {
                        $pnpSharePointAdminContext = Connect-PnPOnline -Url $SharePointAdminUrl -PnPManagementShell 
                    } else {
                        Write-LiveTilesHost -messageType ERROR -afterOutputMessage $Error -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
                    }
                }
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        } else {

            $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "PnPOnline-Module: Sign in to Office 365 with SharePoint Admin rights"
            $pnpSharePointAdminContext = Connect-PnPOnline -ReturnConnection -Url $SharePointAdminUrl -Interactive
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        }
    } catch {
        Write-LiveTilesHost -messageType ERROR -outputMaxLength $outputMaxLength -initialStringLength $output -afterOutputMessage "Cannot login. Insufficient rights? Exiting script"
        Write-Host "Hints:"
        Write-Host "- Insufficient rights?"
        Write-Host "  - User needs to be assigned SharePoint Administrator Role (or global admin...) to add packages and setup sharepoint"
        Write-Host "  - User needs to be assigned Application Administrator Role (or Global Admin...) to add Service Principals (connect to LiveTiles Services)"
        Write-Host "  - User needs to be assigned Priveleged Role Administrator Role (or Global Admin...) to be able to grant consent to Service Principal Permission requests on behalf of all users"
        Write-Host "- MFA Enabled?"
        Write-Host "  - Use the -UseWeblogin parameter)."Exit
    }
    $tenantId = Get-PnPTenantId

#endregion

#region SPO Online module
     try {
        if (-not $UseWebLogin.IsPresent) {
            $output = Write-LiveTilesHost -messageType PROCESS -message ("Logging in to SharePoint Online Module as " + $SharePointTenantAdmin)
                Connect-SPOService -Url $SharePointAdminUrl -Credential $credentialsOffice365Admin
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        } else {
            $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "SharePoint Online Module: Sign in to Office 365 with SharePoint Admin rights"
            Connect-SPOService -Url $SharePointAdminUrl
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        }
    } catch {
        Write-LiveTilesHost -messageType ERROR -outputMaxLength $outputMaxLength -initialStringLength $output -afterOutputMessage "Cannot login. Exiting script"
        Write-Host "Hints:"
        Write-Host "- Insufficient rights?"
        Write-Host "  - User needs to be assigned SharePoint Administrator Role (or global admin...) to add packages and setup sharepoint"
        Write-Host "  - User needs to be assigned Application Administrator Role (or Global Admin...) to add Service Principals (connect to LiveTiles Services)"
        Write-Host "  - User needs to be assigned Priveleged Role Administrator Role (or Global Admin...) to be able to grant consent to Service Principal Permission requests on behalf of all users"
        Write-Host "- MFA Enabled?"
        Write-Host "  - Use the -UseWeblogin parameter)."
        Exit
    }

#endregion

#region azuread module
    try {
        if (-not $UseWebLogin.IsPresent) {
            $output = Write-LiveTilesHost -messageType PROCESS -message ("Logging in to AzureAD Module as " + $SharePointTenantAdmin)
            $AADcontext = Connect-AzureAD -Credential $credentialsOffice365Admin
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        } else {
            $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "AzureAD-Module: Sign in to AzureAD with Application Administrator rights"
            $AADcontext = Connect-AzureAD
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
            
        }  
    } catch {
        Write-LiveTilesHost -messageType ERROR -outputMaxLength $outputMaxLength -initialStringLength $output -afterOutputMessage "Cannot login. Exiting script"
        Write-Host "Hints:"
        Write-Host "- Insufficient rights?"
        Write-Host "  - User needs to be assigned SharePoint Administrator Role (or global admin...) to add packages and setup sharepoint"
        Write-Host "  - User needs to be assigned Application Administrator Role (or Global Admin...) to add Service Principals (connect to LiveTiles Services)"
        Write-Host "  - User needs to be assigned Priveleged Role Administrator Role (or Global Admin...) to be able to grant consent to Service Principal Permission requests on behalf of all users"
        Write-Host "- MFA Enabled?"
        Write-Host "  - Use the -UseWeblogin parameter)."
        Exit
    }

#endregion

#region Get Azure AD internal API token
    $output = Write-LiveTilesHost -messageType PROCESS -message ("Getting Azure AD API Token")
    if (-not $UseWebLogin.IsPresent) {
        $AADapiToken = (Invoke-RestMethod "https://login.windows.net/$($tenantId)/oauth2/token" -Method POST -Body "resource=74658136-14ec-4630-ad9b-26e160ff0fc6&grant_type=password&client_id=1950a258-227b-4e31-a9cf-717495945fc2&scope=openid&username=$($SharePointTenantAdmin)&password=$($SharePointTenantAdminPassword)").access_token
    } else {
        $AADapiToken = Get-PortalAPIAccessToken -ClientId "1950a258-227b-4e31-a9cf-717495945fc2" -TenantId $tenantId -Resource "74658136-14ec-4630-ad9b-26e160ff0fc6"
    }
    Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
#endregion

if (-not $uninstall.IsPresent) {
    #region Checking if Application Catalog exists
        $output = Write-LiveTilesHost -messageType PROCESS -message "Checking if Application Catalog exists" -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
            $appCatalogUrl = Get-PnPTenantAppCatalogUrl -Connection $pnpSharePointAdminContext
            $dotCount = 0
            if ($appCatalogUrl -eq "" -or $appCatalogUrl -eq $null) {
                Write-LiveTilesHost -messageType WARNING -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference -afterOutputMessage "The App Catalog doesn't exist. Creating it..."
                $output = 0

                Get-PnPTimeZoneId | Out-Host
                $appCatalogUrl = $SharePointUrl + "/sites/appcatalog"
                if ($PSCmdlet.ShouldProcess("AppCatalogSite ", "Create")) {
                    Register-PnPAppCatalogSite -Url $appCatalogUrl -Owner $SharePointTenantAdmin -Connection $pnpSharePointAdminContext |Out-Null # -TimeZoneId 4

                    $appCatExists = $False
                    do {
                        if (-not $useWebLogin.IsPresent) {
                            $pnpSharePointAdminContext = Connect-PnPOnline -ReturnConnection -Url $SharePointAdminUrl -Credentials $credentialsOffice365Admin
                        } else {
                            $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "PnPOnline-Module: Sign in to Office 365 with SharePoint Admin rights"
                            $pnpSharePointAdminContext = Connect-PnPOnline -ReturnConnection -Url $SharePointAdminUrl -Interactive
                        }
                        try {
                            $appCatApps = Get-PnPApp -Connection $pnpSharePointAdminContext -Scope Tenant -ErrorAction Stop
                            $appCatExists = $True
                        } catch {
                            if (($output + $dotCount) -le ($outputMaxLength - 2)) {
                                Write-Host -NoNewline "." -ForegroundColor White
                                $dotCount++
                            }
                            Start-Sleep -Seconds $waitLong
                        }
                    } until ($appCatExists)
                }
            }                

        Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength ($output+$dotCount) -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference

    #endregion
    
    #region check app catalog rights
        $output = Write-WizdomHost -messageType PROCESS -message "Checking Application Catalog rights" -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
            $appCatalogAdmin = $null
            $appCatalogAdminGroup = $null
            do {    
                Start-Sleep -Seconds $waitShort
                try {
                    $appCatalogAdmin = Get-SPOUser -site $appCatalogUrl -LoginName $SharePointTenantAdmin
                    $appCatalogAdminGroup = get-spositegroup -Site $appCatalogUrl
                } catch {
                    $appCatalogAdmin = $null
                    $appCatalogAdminGroup = $null
                }
            } while ($null -eq $appCatalogAdmin -and $null -eq $appCatalogAdminGroup)
            if ((($appCatalogAdmin).where({$_.IsSiteAdmin -eq $true}).count -eq 0) -and (($appCatalogAdminGroup.where{($_.roles -match "Full Control" -or $_.roles -match "Contribute") -and $_.users -match $SharePointTenantAdmin}).count -eq 0)) {
                Write-WizdomHost -messageType ERROR -outputMaxLength $outputMaxLength -initialStringLength $output -afterOutputMessage "$($SharePointTenantAdmin) doesn't have access to the App Catalog at $($appCatalogUrl) as either Owner or Site Collection Owner. Stopping installation script" -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
                Exit
            }
        Write-WizdomHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
    #endregion

    #region enable CDN
    if ($EnableCDN.IsPresent) {
        if (-not (Get-PnPTenantCdnEnabled -Connection $pnpSharePointAdminContext -CdnType Public).value) {
            $output = Write-LiveTilesHost -messageType PROCESS -message "Enabling CDN for tenant" -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
            if ($PSCmdlet.ShouldProcess("TenantCDN ", "Enable")) {
                Set-PnPTenantCdnEnabled -CdnType Public -Connection $pnpSharePointAdminContext -Enable:$true
            }
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength ($output) -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        }

    }
    #endregion

    #region add hubsite collection if not existing
        $output = Write-LiveTilesHost -messageType PROCESS -message "Checking if Hub SiteCollection exists" -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        $hubSite = Get-PnPTenantSite -Connection $pnpSharePointAdminContext -Identity $HubSiteCollection -ErrorAction Ignore
        if ($hubSite -eq "" -or $hubSite -eq $null) {
            Write-LiveTilesHost -messageType WARNING -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference -afterOutputMessage "The Hub site doesn't exist. Creating $($HubSiteCollection)"
            $output = 0
            if ($PSCmdlet.ShouldProcess("HubSite ", "Create")) {
                New-PnPSite -Type CommunicationSite -Title Intranet -Url $hubSiteCollection -Connection $pnpSharePointAdminContext -Wait | Write-Verbose
            }
        }                

        Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength ($output) -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
    #endregion

    #region Download / Upload LiveTiles Packages
        $output = Write-LiveTilesHost -messageType PROCESS -message "Getting LiveTiles installation files and deploying packages to SharePoint" -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        try {
            $fileListPublishPackages = @()
                $fileListPublishPackages += "https://install.matchpoint365.com/resources/LiveTiles.Intranet.Hub.sppkg"
                if (-not $noGovernance.IsPresent) {$fileListPublishPackages += "https://install.matchpoint365.com/resources/LiveTiles.Intranet.Governance.sppkg"}
                if (-not $noEverywhere.IsPresent) {$fileListPublishPackages += "https://install.matchpoint365.com/resources/LiveTiles.Everywhere.Panel.sppkg"}
        
            $fileListUploadPackages = @()
                $fileListUploadPackages += "https://install.matchpoint365.com/resources/LiveTiles.Intranet.Hub.LandingPage.sppkg"
                if (-not $noWorkspaces.IsPresent) {$fileListUploadPackages += "https://install.matchpoint365.com/resources/LiveTiles.Intranet.Workspaces.sppkg"}
                if (-not $noMetadata.IsPresent) {$fileListUploadPackages += "https://install.matchpoint365.com/resources/LiveTiles.Intranet.Metadata.sppkg"}
        
            $fileListSearchConfig = @(
                "https://install.matchpoint365.com/resources/SearchSchemaConfiguration.xml"
            )
            $dotCount = 0
            foreach ($file in ($fileListUploadPackages)) {
                if (($output + $dotCount) -le ($outputMaxLength - 2)) {
                    Write-Host -NoNewline "." -ForegroundColor White
                    $dotCount++
                }

                $fileName = ($env:TMP + "\" + $file.split("/")[-1])
                Invoke-WebRequest $file -OutFile $fileName
                if ($PSCmdlet.ShouldProcess("Site Scoped LiveTiles Apps", "Upload and Publish")) {
                    do { #If the App Catalog is newly created, the 
                        try {
                            Add-PnPApp -Path $fileName -Publish -Connection $pnpSharePointAdminContext -Overwrite -ErrorAction Stop| Write-Verbose
                            $appDeployed = $true
                        } catch {$appDeployed = $false}
                    } until ($appDeployed)

                }
                Remove-Item -Path $fileName -Force -ErrorAction Ignore
            }

            foreach ($file in $fileListPublishPackages) {
                if (($output + $dotCount) -le ($outputMaxLength - 2)) {
                    Write-Host -NoNewline "." -ForegroundColor White
                    $dotCount++
                }
                $fileName = ($env:TMP + "\" + $file.split("/")[-1])
                Invoke-WebRequest $file -OutFile $fileName
                if ($PSCmdlet.ShouldProcess("Tenant Scoped LiveTiles Apps", "Upload and Publish")) {
                    $appId = (Add-PnPApp -Path $fileName -Publish -Connection $pnpSharePointAdminContext -Overwrite -SkipFeatureDeployment).Id
                }
                Remove-Item -Path $fileName -Force -ErrorAction Ignore
            }
            if (($output + $dotCount) -le ($outputMaxLength - 2)) {
                Write-Host -NoNewline "." -ForegroundColor White
                $dotCount++
            }
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength ($output+$dotCount) -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        } catch {
            Write-LiveTilesHost -messageType ERROR -afterOutputMessage $Error -outputMaxLength $outputMaxLength -initialStringLength ($output+$dotCount)
        }
    #endregion

    #region Setting Tenant level search configuration (managed properties)
        Write-Host "Settings are not overridden, so manual action may be required on tenants with preexisting configuration:" -ForegroundColor DarkGray
    
        $fileSearchConfig = "https://install.matchpoint365.com/resources/SearchSchemaConfiguration.xml"
        $fileName = ($env:TMP + "\" + $fileSearchConfig.split("/")[-1])
        Invoke-WebRequest $fileSearchConfig -OutFile $fileName

        $output = Write-LiveTilesHost -messageType PROCESS -message "Setting Tenant level search configuration (managed properties)" -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        try {
            if ($PSCmdlet.ShouldProcess("SearchConfiguration", "Upload")) {
                set-PnPSearchConfiguration -path $fileName -Scope Subscription -Connection $pnpSharePointAdminContext -ErrorAction Stop
            }
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        } catch {
            #2Do - Better error message - perhaps a description of what needs to be done manually?
            Write-LiveTilesHost -messageType WARNING -outputMaxLength $outputMaxLength -initialStringLength $output -afterOutputMessage $_.exception.message -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        } finally {
            Remove-Item -Path $fileName -Force -ErrorAction Ignore
        }
    #endregion

    #region Create Subscriptions
        $output = Write-LiveTilesHost -messageType PROCESS -message "Creating Service Principals in AAD." -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        $dotCount = 0


  
        #Creating Service Principals
        foreach ($appId in $appIdsLiveTiles) {
            if (($output + $dotCount) -le ($outputMaxLength - 2)) {
                Write-Host -NoNewline "." -ForegroundColor White
                $dotCount++
            }
            if ($PSCmdlet.ShouldProcess("Service Principal $($appId)", "Create")) {
                Grant-OAuth2PermissionsToApp -azureAppId $appId -apiToken $AADapiToken|write-verbose
            }
        }

        #Waiting until Service Principals are created
        if ($PSCmdlet.ShouldProcess("ServicePricinpal Creation finalized ", "Wait")) {
            $completed = $false
            do {
                if (-not $UseWebLogin.IsPresent) {
                    $AADcontext = Connect-AzureAD -Credential $credentialsOffice365Admin
                } else {
                    Start-Sleep -Seconds $waitLong
                    Write-Host
                    $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "AzureAD-Module: Sign in to AzureAD with Application Administrator rights"
                    $AADcontext = Connect-AzureAD
                    Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
            
                }                
                $splist = Get-AzureADServicePrincipal -All $true
                foreach ($appId in $appIdsLiveTiles) {
                    if ($splist.where({$_.AppId -eq $appId}).count -eq 0) {
                        $completed = $false
                        if (($output + $dotCount) -le ($outputMaxLength - 2)) {
                            Write-Host -NoNewline "." -ForegroundColor White
                            $dotCount++
                        }
                        Start-Sleep -Seconds $waitLong
                        break
                    } else {
                        $completed = $true
                    }
                }

            } until ($completed)
        }

        Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength ($output+$dotCount) -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference

    #Region Creating Subscriptions

        if ($Reach -ne "NoInstall") {
            $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "Opening browser to create Reach Subscription - please return here afterwards." -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
            if ($PSCmdlet.ShouldProcess("Reach Subscription", "Create")) {

                sleep -Seconds $waitShort
                switch ($Reach) {
                    'DefaultDemo' {
                        Start-Process "https://reach.livetiles.io/subscribe"
                    }
                    'BroadcastLtd' {
                        Start-Process "https://reach.livetiles.io/subscribe?specificDemoSubscriptionId=94e7bddc-6842-4b42-aebb-0a558e76cf98"
                    }
                    'CareNet' {
                        Start-Process "https://reach.livetiles.io/subscribe?specificDemoSubscriptionId=e60f81bd-dd26-478d-aee7-3edd794e4273"
                    }
                    'FoodInc' {
                        Start-Process "https://reach.livetiles.io/subscribe?specificDemoSubscriptionId=bc4233b1-de1f-4777-8b51-7a65fdb46850"
                    }
                }
                Write-Host
                pause
                Write-Host
            }
        }

        $output = Write-LiveTilesHost -messageType PROCESS -message "Opening browser to create LiveTiles Subscriptions - No action needed." -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        #sleep -Seconds $waitShort
            if ($PSCmdlet.ShouldProcess("Hub Subscription", "Create")) {
                Start-Process "https://hub.matchpoint365.com/api/manage/register"
            }
            if (-not $noGovernance.IsPresent) {
                if ($PSCmdlet.ShouldProcess("Governance Subscription", "Create")) {
                    Start-Process "https://governance.matchpoint365.com/api/governance/register"
                }
            }
            if (-not $noMetadata.IsPresent) {
                if ($PSCmdlet.ShouldProcess("Metadata Subscription", "Create")) {
                    Start-Process "https://metadata.matchpoint365.com/api/metadata/register"
                }
            }
            if (-not $noProvisioning.IsPresent) {
                if ($PSCmdlet.ShouldProcess("Provisioning Subscription", "Create")) {
                    Start-Process "https://provisioning.matchpoint365.com/api/provision/register"
                }
            }
            if (-not $noWorkspaces.IsPresent) {
                if ($PSCmdlet.ShouldProcess("Workspaces Subscription", "Create")) {
                    Start-Process "https://workspaces.matchpoint365.com/api/workspaces/register"
                }
            }
        Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference

    #endregion

    #region Grant Consent between apps
        if ($Reach -ne "NoInstall") {
            $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "Opening browser to grant Reach consent to access Hub - please return here afterwards." -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
            sleep -Seconds $waitShort
            if ($PSCmdlet.ShouldProcess("Reach -> Hub Access", "Consent")) {
                Start-Process "https://login.microsoftonline.com/common/oauth2/authorize?response_type=id_token&client_id=02de0a9e-a2a9-42af-a6c3-2ed77e913a4b&redirect_uri=https%3A%2F%2Fapp.condense.ch/logout"
            }
            Write-Host
            pause
            Write-Host
        }

        if ((-not $noWorkspaces.IsPresent) -and (-not $noMetadata.IsPresent)) {
            $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "Opening browser to grant Workspaces consent to access Metadata - please return here afterwards." -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
            sleep -Seconds $waitShort
            if ($PSCmdlet.ShouldProcess("Workspaces -> Metadata Access", "Consent")) {
                Start-Process "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=a9a0e8f6-2cee-42d2-b00f-c9299e509958&response_type=id_token%20token&scope=https://iurcycl.onmicrosoft.com/matchpoint-metadata/user_impersonation&redirect_uri=https://workspaces.matchpoint365.com/api/workspaces/register"
            }
            Write-Host
            pause
            Write-Host
        }

        if ((-not $noWorkspaces.IsPresent) -and (-not $noProvisioning.IsPresent)) {
            $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "Opening browser to grant Workspaces consent to access Provisioning - please return here afterwards." -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
            sleep -Seconds $waitShort
            if ($PSCmdlet.ShouldProcess("Workspaces -> Provisioning Access", "Consent")) {
                Start-Process "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=a9a0e8f6-2cee-42d2-b00f-c9299e509958&response_type=id_token%20token&scope=https://iurcycl.onmicrosoft.com/matchpoint-provisioning/user_impersonation&redirect_uri=https://workspaces.matchpoint365.com/api/workspaces/register"
            }
            Write-Host
            pause
            Write-Host
        }

        if (-not $noGovernance.IsPresent) {
            #2Do - Untested runtime - may have to be placed before subscription link??
            $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "Opening browser to grant Governance consent to access SharePoint - please return here afterwards." -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
            sleep -Seconds $waitShort
            if ($PSCmdlet.ShouldProcess("Governance -> SharePoint Access", "Consent")) {
                Start-Process "https://login.microsoftonline.com/common/oauth2/authorize?client_id=b6a4f91d-466c-4afe-bb29-3f41ea615da7&redirect_uri=https%3A%2F%2Fgovernance.matchpoint365.com%2F&response_type=id_token&scope=openid%20profile&response_mode=form_post&nonce=637589315897098371.NDdmM2JjZjMtZjllMy00NTgyLWI3MjctOGE4MmIyNDFiNzBlZjJiYjUwNGYtMjZiYS00OTgwLTk1ZWUtMjE5ZGE2ZDFlYjQ5&state=CfDJ8LRbmffsmjNCqinaRfnCoRvfwFPJ109R2f6GGWj_wM8uZQiKPHIT--3jIsiMFsFChjtSZTJSvWGF9suCktafrskUdJ96RbSCyl4eLHIcQ2Yw6emt-wpDnOmUA9jobArDn0-zd_2HKy6B1GMAIcCGOn9BkntyXx4W56guWvmJatuxvBhTqfaMcBJIr8W3WxUE6cGno5eCM-rGTovP42jMBzrsRWrsBQyXRKGvgKKYHUz8UnHq2Oikp2-gKrygUGL7UK1ZDPlD5zDDn44Yhv2B0RyQfrHVE8nYw3vMJYW--iJt0SFa-pQJhxrC-XBbeh6cRcizmwF70XIG0jBZa5fbtu_UeeJKi021WapSolby_54WkIiIZFxW-oTX3YwS-Ie69A&x-client-SKU=ID_NET461&x-client-ver=5.3.0.0"
            }
            Write-Host
            pause
            Write-Host
        }

    #endregion

    #region Approve API requests
        $output = Write-LiveTilesHost -messageType PROCESS -message "Aproving API requests from Apps" -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        $dotCount = 0

        $requests = Get-PnPTenantServicePrincipalPermissionRequests -Connection $pnpSharePointAdminContext
        $requestsToApprove = $requests | ? { $_.PackageName -eq 'LiveTiles Intranet Workspaces' -or $_.PackageName -eq 'LiveTiles Intranet Metadata' -or $_.PackageName -eq 'LiveTiles Intranet Hub' -or $_.PackageName -eq 'LiveTiles Intranet Governance' }

        if ($requestsToApprove -ne $null -or ($requestsToApprove.Count -eq 0))
        {
            foreach($request in $requestsToApprove)
            {
                if ($PSCmdlet.ShouldProcess("PrincipalPermissionRequest $($request.PackageName)", "Approve")) {
                    Approve-PnPTenantServicePrincipalPermissionRequest -RequestId $request.Id -Force -Connection $pnpSharePointAdminContext -ErrorAction SilentlyContinue | Write-Verbose
                    if (($output + $dotCount) -le ($outputMaxLength - 2)) {
                        Write-Host -NoNewline "." -ForegroundColor White
                        $dotCount++
                    }
                }
            }
        }

        Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength ($output+$dotCount) -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
    #endregion

    #region Logging in to Hub Site Collection
        if ($PSCmdlet.ShouldProcess("Hubsite", "Login")) {    
            try {
                if (-not $useWebLogin.IsPresent) {
                    $output = Write-LiveTilesHost -messageType PROCESS -message ("Logging in to SharePoint Hub Site as " + $SharePointTenantAdmin)
                        $pnpSharePointHubSiteContext = Connect-PnPOnline -ReturnConnection -Url $hubSiteCollection -Credentials $credentialsOffice365Admin
                        $pnpUserName = $pnpSharePointHubSiteContext.PSCredential.UserName
                    Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
                } else {
                    $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "PnPOnline-Module: Sign in to Office 365 with SharePoint Admin rights"
                        $pnpSharePointHubSiteContext = Connect-PnPOnline -ReturnConnection -Url $hubSiteCollection -Interactive
                        $pnpUserName = (Get-PnPAccessToken -Decoded).payload.upn
                    Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
                }
            } catch {
                Write-LiveTilesHost -messageType ERROR -outputMaxLength $outputMaxLength -initialStringLength $output -afterOutputMessage "Cannot login. Insufficient rights? Exiting script"
                Exit
            }
        }
    #endregion

    #region Add Apps to Site Collection
        $output = Write-LiveTilesHost -messageType PROCESS -message ("Adding local apps to Hub SiteCollection " + $pnpSharePointHubSiteContext.Url)
            if ($PSCmdlet.ShouldProcess("SiteCollectionAdmin", "Add")) {
                $hubSiteAdminList = @((Get-PnPSiteCollectionAdmin -Connection $pnpSharePointHubSiteContext).LoginName)
                if (("i:0#.f|membership|"+$pnpUserName) -notin $hubSiteAdminList) {
                    Add-PnPSiteCollectionAdmin -Owners $pnpUserName -Connection $pnpSharePointHubSiteContext
                }
            }

        $availableApps = Get-PnPApp -Scope Tenant -Connection $pnpSharePointHubSiteContext

        if ($PSCmdlet.ShouldProcess("Site Collection Apps", "Install")) {
            Install-PnPApp -Scope Tenant -Connection $pnpSharePointHubSiteContext -Identity $availableApps.where({$_.Title -eq "LiveTiles Intranet Hub Landing Page"}).Id
            if (-not $noMetadata.IsPresent) {Install-PnPApp -Scope Tenant -Connection $pnpSharePointHubSiteContext -Identity $availableApps.where({$_.Title -eq "LiveTiles Intranet Metadata"}).Id}
            if (-not $noWorkspaces.IsPresent) {Install-PnPApp -Scope Tenant -Connection $pnpSharePointHubSiteContext -Identity $availableApps.where({$_.Title -eq "LiveTiles Intranet Workspaces"}).Id}
        }
        $fileExists = $false
        if ($PSCmdlet.ShouldProcess("Hub.aspx creation", "Wait")) {
 
            do {
                #$pnpSharePointHubSiteContext  = Connect-PnPOnline -AccessToken Get-PnPAccessToken -url $hubSiteCollection -ReturnConnection

                $file = Get-PnPFile -Url '/SitePages/Hub.aspx' -Connection $pnpSharePointHubSiteContext -ErrorAction SilentlyContinue
                if ($file -ne $null) {
                    $fileExists = $true
                } else {
                    Start-Sleep -Seconds $waitShort
                }
            } until ($fileExists)
        }
        Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
    #endregion

    #region Add webparts to frontpage
        $output = Write-LiveTilesHost -messageType PROCESS -message ("Adding webparts to Hub Frontpage")
        $xmlHubWebPart = '<webParts>
      <webPart xmlns="http://schemas.microsoft.com/WebPart/v3">
        <metaData>
          <type name="Microsoft.SharePoint.WebPartPages.ClientSideWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
          <importErrorMessage>Cannot Import WebPart</importErrorMessage>
        </metaData>
        <data>
          <properties>
            <property name="AllowZoneChange" type="bool">True</property>
            <property name="HelpUrl" type="string" />
            <property name="Hidden" type="bool">False</property>
            <property name="MissingAssembly" type="string">Cannot Import WebPart</property>
            <property name="Description" type="string">A Client side web part that renders components from the LiveTiles Intranet Hub.</property>
            <property name="AllowHide" type="bool">True</property>
            <property name="AllowMinimize" type="bool">True</property>
            <property name="ExportMode" type="exportmode">All</property>
            <property name="Title" type="string">LiveTiles Intranet Hub</property>
            <property name="TitleUrl" type="string" />
            <property name="ClientSideWebPartData" type="string">&lt;div data-sp-webpart="" data-sp-webpartdataversion=1.0 data-sp-webpartdata="&amp;#123;&amp;quot;id&amp;quot;&amp;#58;&amp;quot;92f5dabd-1a54-43e4-bbb5-34a751d96900&amp;quot;,&amp;quot;instanceId&amp;quot;&amp;#58;null,&amp;quot;title&amp;quot;&amp;#58;&amp;quot;LiveTiles Intranet Hub&amp;quot;,&amp;quot;description&amp;quot;&amp;#58;&amp;quot;A Client side web part that renders components from the LiveTiles Intranet Hub.&amp;quot;,&amp;quot;version&amp;quot;&amp;#58;&amp;quot;0.0.1&amp;quot;,&amp;quot;properties&amp;quot;&amp;#58;&amp;#123;&amp;quot;description&amp;quot;&amp;#58;&amp;quot;Renders LiveTiles Intranet Hub components.&amp;quot;&amp;#125;,&amp;quot;htmlProperties&amp;quot;&amp;#58;null&amp;#125;"&gt;&lt;div data-sp-componentid=""&gt;92f5dabd-1a54-43e4-bbb5-34a751d96900&lt;/div&gt;&lt;div data-sp-htmlproperties=""&gt;&lt;/div&gt;&lt;/div&gt;</property>
            <property name="ChromeType" type="chrometype">None</property>
            <property name="AllowConnect" type="bool">True</property>
            <property name="Width" type="string" />
            <property name="Height" type="string" />
            <property name="CatalogIconImageUrl" type="string" />
            <property name="HelpMode" type="helpmode">Modeless</property>
            <property name="ClientSideWebPartId" type="System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089">92f5dabd-1a54-43e4-bbb5-34a751d96900</property>
            <property name="AllowEdit" type="bool">True</property>
            <property name="TitleIconImageUrl" type="string" />
            <property name="Direction" type="direction">NotSet</property>
            <property name="AllowClose" type="bool">True</property>
            <property name="ChromeState" type="chromestate">Normal</property>
          </properties>
        </data>
      </webPart>
    </webParts>'

    $xmlWorkspacesWebpart = '<webParts>
      <webPart xmlns="http://schemas.microsoft.com/WebPart/v3">
        <metaData>
          <type name="Microsoft.SharePoint.WebPartPages.ClientSideWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
          <importErrorMessage>Cannot Import WebPart</importErrorMessage>
        </metaData>
        <data>
          <properties>
            <property name="AllowZoneChange" type="bool">True</property>
            <property name="HelpUrl" type="string" />
            <property name="Hidden" type="bool">False</property>
            <property name="MissingAssembly" type="string">Cannot Import WebPart</property>
            <property name="Description" type="string">Provides the LiveTiles Intranet Workspaces functionality as a Web Part.</property>
            <property name="AllowHide" type="bool">True</property>
            <property name="AllowMinimize" type="bool">True</property>
            <property name="ExportMode" type="exportmode">All</property>
            <property name="Title" type="string">LiveTiles Intranet Workspaces</property>
            <property name="TitleUrl" type="string" />
            <property name="ClientSideWebPartData" type="string">&lt;div data-sp-webpart="" data-sp-webpartdataversion=1.0 data-sp-webpartdata="&amp;#123;&amp;quot;id&amp;quot;&amp;#58;&amp;quot;5bdc6955-5379-471a-9893-b12adb6aa126&amp;quot;,&amp;quot;instanceId&amp;quot;&amp;#58;null,&amp;quot;title&amp;quot;&amp;#58;&amp;quot;LiveTiles Intranet Workspaces&amp;quot;,&amp;quot;description&amp;quot;&amp;#58;&amp;quot;Provides the LiveTiles Intranet Workspaces functionality as a Web Part.&amp;quot;,&amp;quot;version&amp;quot;&amp;#58;&amp;quot;0.0.1&amp;quot;,&amp;quot;properties&amp;quot;&amp;#58;&amp;#123;&amp;quot;description&amp;quot;&amp;#58;&amp;quot;LiveTiles Intranet Workspaces&amp;quot;&amp;#125;,&amp;quot;htmlProperties&amp;quot;&amp;#58;null&amp;#125;" data-sp-splinksapplied=true&gt;&lt;div data-sp-componentid=""&gt;5bdc6955-5379-471a-9893-b12adb6aa126&lt;/div&gt;&lt;div data-sp-htmlproperties=""&gt;&lt;/div&gt;&lt;/div&gt;</property>
            <property name="ChromeType" type="chrometype">None</property>
            <property name="AllowConnect" type="bool">True</property>
            <property name="Width" type="string" />
            <property name="Height" type="string" />
            <property name="CatalogIconImageUrl" type="string" />
            <property name="HelpMode" type="helpmode">Modeless</property>
            <property name="ClientSideWebPartId" type="System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089">5bdc6955-5379-471a-9893-b12adb6aa126</property>
            <property name="AllowEdit" type="bool">True</property>
            <property name="TitleIconImageUrl" type="string" />
            <property name="Direction" type="direction">NotSet</property>
            <property name="AllowClose" type="bool">True</property>
            <property name="ChromeState" type="chromestate">Normal</property>
          </properties>
        </data>
      </webPart>
    </webParts>'

    $xmlGovernanceWebpart = '<webParts>
      <webPart xmlns="http://schemas.microsoft.com/WebPart/v3">
        <metaData>
          <type name="Microsoft.SharePoint.WebPartPages.ClientSideWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
          <importErrorMessage>Cannot Import WebPart</importErrorMessage>
        </metaData>
        <data>
          <properties>
            <property name="AllowZoneChange" type="bool">True</property>
            <property name="HelpUrl" type="string" />
            <property name="Hidden" type="bool">False</property>
            <property name="MissingAssembly" type="string">Cannot Import WebPart</property>
            <property name="Description" type="string">Provides the LiveTiles Intranet Governance Dashboard functionality as a Web Part.</property>
            <property name="AllowHide" type="bool">True</property>
            <property name="AllowMinimize" type="bool">True</property>
            <property name="ExportMode" type="exportmode">All</property>
            <property name="Title" type="string">LiveTiles Intranet Governance</property>
            <property name="TitleUrl" type="string" />
            <property name="ClientSideWebPartData" type="string">&lt;div data-sp-webpart="" data-sp-webpartdataversion=1.0 data-sp-webpartdata="&amp;#123;&amp;quot;id&amp;quot;&amp;#58;&amp;quot;d1b0ad55-7325-4fee-af7a-80afa1f7f127&amp;quot;,&amp;quot;instanceId&amp;quot;&amp;#58;null,&amp;quot;title&amp;quot;&amp;#58;&amp;quot;LiveTiles Intranet Governance&amp;quot;,&amp;quot;description&amp;quot;&amp;#58;&amp;quot;Provides the LiveTiles Intranet Governance Dashboard functionality as a Web Part.&amp;quot;,&amp;quot;version&amp;quot;&amp;#58;&amp;quot;0.0.1&amp;quot;,&amp;quot;properties&amp;quot;&amp;#58;&amp;#123;&amp;quot;description&amp;quot;&amp;#58;&amp;quot;LiveTiles Intranet Governance Dashboard&amp;quot;&amp;#125;,&amp;quot;htmlProperties&amp;quot;&amp;#58;null&amp;#125;" data-sp-splinksapplied=true&gt;&lt;div data-sp-componentid=""&gt;d1b0ad55-7325-4fee-af7a-80afa1f7f127&lt;/div&gt;&lt;div data-sp-htmlproperties=""&gt;&lt;/div&gt;&lt;/div&gt;</property>
            <property name="ChromeType" type="chrometype">None</property>
            <property name="AllowConnect" type="bool">True</property>
            <property name="Width" type="string" />
            <property name="Height" type="string" />
            <property name="CatalogIconImageUrl" type="string" />
            <property name="HelpMode" type="helpmode">Modeless</property>
            <property name="ClientSideWebPartId" type="System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089">d1b0ad55-7325-4fee-af7a-80afa1f7f127</property>
            <property name="AllowEdit" type="bool">True</property>
            <property name="TitleIconImageUrl" type="string" />
            <property name="Direction" type="direction">NotSet</property>
            <property name="AllowClose" type="bool">True</property>
            <property name="ChromeState" type="chromestate">Normal</property>
          </properties>
        </data>
      </webPart>
    </webParts>'

        if ($PSCmdlet.ShouldProcess("App WebParts for classic Hub page", "Add")) {
            #set pnp-hompage?
            Add-PnPWebPartToWebPartPage -ServerRelativePageUrl "/SitePages/Hub.aspx" -Xml $xmlHubWebPart -ZoneId "Zone 1" -ZoneIndex "0" -Connection $pnpSharePointHubSiteContext
            if (-not $noGovernance.IsPresent) {Add-PnPWebPartToWebPartPage -ServerRelativePageUrl "/SitePages/Hub.aspx" -Xml $xmlGovernanceWebpart -ZoneId "Zone 1" -ZoneIndex "0" -Connection $pnpSharePointHubSiteContext}
            if (-not $noWorkspaces.IsPresent) {Add-PnPWebPartToWebPartPage -ServerRelativePageUrl "/SitePages/Hub.aspx" -Xml $xmlWorkspacesWebpart -ZoneId "Zone 1" -ZoneIndex "0" -Connection $pnpSharePointHubSiteContext}
        }

        Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
    #endregion
} else { #Uninstall
    #region check app catalog rights
        $output = Write-WizdomHost -messageType PROCESS -message "Checking Application Catalog rights" -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
            $appCatalogAdmin = $null
            $appCatalogAdminGroup = $null
            do {    
                Start-Sleep -Seconds $waitShort
                try {
                    $appCatalogAdmin = Get-SPOUser -site $appCatalogUrl -LoginName $SharePointTenantAdmin
                    $appCatalogAdminGroup = get-spositegroup -Site $appCatalogUrl
                } catch {
                    $appCatalogAdmin = $null
                    $appCatalogAdminGroup = $null
                }
            } while ($null -eq $appCatalogAdmin -and $null -eq $appCatalogAdminGroup)
            if ((($appCatalogAdmin).where({$_.IsSiteAdmin -eq $true}).count -eq 0) -and (($appCatalogAdminGroup.where{($_.roles -match "Full Control" -or $_.roles -match "Contribute") -and $_.users -match $SharePointTenantAdmin}).count -eq 0)) {
                Write-WizdomHost -messageType ERROR -outputMaxLength $outputMaxLength -initialStringLength $output -afterOutputMessage "$($SharePointTenantAdmin) doesn't have access to the App Catalog at $($appCatalogUrl) as either Owner or Site Collection Owner. Stopping installation script" -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
                Exit
            }
        Write-WizdomHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
    #endregion
    
    #remove frontpage
    try {
        if (-not $useWebLogin.IsPresent) {
            $output = Write-LiveTilesHost -messageType PROCESS -message ("Logging in to SharePoint Hub Site as " + $SharePointTenantAdmin)
                $pnpSharePointHubSiteContext = Connect-PnPOnline -ReturnConnection -Url $hubSiteCollection -Credentials $credentialsOffice365Admin
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        } else {
            $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "PnPOnline-Module: Sign in to Office 365 with SharePoint Admin rights"
                $pnpSharePointHubSiteContext = Connect-PnPOnline -ReturnConnection -Url $hubSiteCollection -Interactive
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        }
    } catch {
        Write-LiveTilesHost -messageType ERROR -outputMaxLength $outputMaxLength -initialStringLength $output -afterOutputMessage "Cannot login. Insufficient rights? Exiting script"
        Exit
    }
    
    $newHomePage = (Get-PnPList -Connection $pnpSharePointHubSiteContext).where({$_.EntityTypeName -eq "SitePages"}).Id | Get-PnPListItem -Connection $pnpSharePointHubSiteContext | ? {$_.FieldValues.FileLeafRef -ne "Hub.aspx"} | Sort-Object -Property Id | Select-Object -First 1
    $output = Write-LiveTilesHost -messageType PROCESS -message ("Setting new Homepage to " + $newHomePage.FieldValues.FileLeafRef)
        if ($newHomePage -ne $null) {
           if ($PSCmdlet.ShouldProcess("Homepage $($newHomePage)", "Set")) {
               Set-PnPHomePage -Connection $pnpSharePointHubSiteContext -RootFolderRelativeUrl $newHomePage.FieldValues.FileLeafRef
           }
        }
    $f = (get-pnpfile /SitePages/Hub.aspx -AsListItem)
    if ($f -ne $null) {
        if ($PSCmdlet.ShouldProcess("Page /SitePages/Hub.aspx", "Delete")) {
            $f.DeleteObject()
        }
    }

    Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference

    #remove apps from site collection
    $output = Write-LiveTilesHost -messageType PROCESS -message ("Removing local LiveTiles Apps from $($pnpSharePointHubSiteContext.Url)")
        $localapps = (Get-PnPAppInfo -Connection $pnpSharePointAdminContext -Name LiveTiles).ProductId
            if ($localapps.count -gt 0) {
                if ($PSCmdlet.ShouldProcess("Local apps on HubSite", "Uninstall")) {
                    $localapps | Uninstall-PnPApp -Connection $pnpSharePointHubSiteContext
                }
            }
        #(Get-PnPApp -Connection $pnpSharePointHubSiteContext).where({$_.Title -like "LiveTiles*"}) | Remove-PnPApp -Connection $pnpSharePointHubSiteContext
    Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference

    #deny access

    $output = Write-LiveTilesHost -messageType PROCESS -message ("Removing API Access to SharePoint Tenant")
        $requests = (Get-PnPTenantServicePrincipalPermissionRequests -Connection $pnpSharePointAdminContext) | ? { ($_.PackageName -eq 'LiveTiles Intranet Workspaces' -or $_.PackageName -eq 'LiveTiles Intranet Metadata' -or $_.PackageName -eq 'LiveTiles Intranet Hub' -or $_.PackageName -eq 'LiveTiles Intranet Governance' ) }
        if ($requests.Count -gt 0) {
            foreach ($request in $requests) {
                if ($PSCmdlet.ShouldProcess("Pending Service Principal API request $($request.PackageName)", "Deny")) {
                    Deny-PnPTenantServicePrincipalPermissionRequest -RequestId $request.Id -Connection $pnpSharePointAdminContext -Force
                }
            }
        }
        
        $requests = (Get-PnPTenantServicePrincipalPermissionGrants -Connection $pnpSharePointAdminContext) | ? { ($_.Resource -like 'MatchPoint*' -or $_.Resource -like 'Condense*') }
        if ($requests.Count -gt 0) {
            foreach ($request in $requests) {
                if ($PSCmdlet.ShouldProcess("Granted Service Principal API request $($request.Resource)", "Revoke")) {
                    Revoke-SPOTenantServicePrincipalPermission -ObjectId $request.ObjectId
                }
            }
        }
    Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference

    #remove service principals
        if (-not $UseWebLogin.IsPresent) {
            $AADcontext = Connect-AzureAD -Credential $credentialsOffice365Admin
        } else {
            $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "AzureAD-Module: Sign in to AzureAD with Application Administrator rights"
            $AADcontext = Connect-AzureAD
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
            
        }



    $output = Write-LiveTilesHost -messageType PROCESS -message ("Removing Service Principals")
        if ($PSCmdlet.ShouldProcess("Azure AD Service Principals", "Remove")) {
            (Get-AzureADServicePrincipal -All $true).where({$_.AppId -in $appIdsLiveTiles}) | Remove-AzureADServicePrincipal
        }
    Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference

    #remove packages from app catalog
    $output = Write-LiveTilesHost -messageType PROCESS -message ("Removing Apps from Application Site Catalog")
        if ($PSCmdlet.ShouldProcess("Application Site Catalog Packages", "Remove")) {
            (get-pnpapp -Connection $pnpSharePointAdminContext).where({$_.Title -like "LiveTiles*"}) | Remove-PnPApp -Connection $pnpSharePointAdminContext
        }
    Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
}

Disconnect-PnPOnline
Disconnect-SPOService
Write-Host ("Installation endtime: " + (get-date)) -ForegroundColor White


<#
new System.IdentityModel.Tokens.Jwt.JwtSecurityToken(AccessToken)

Id token gennem PnP Management shell app
Invoke-RestMethod "https://login.windows.net/bb9ca92d-d8ac-4810-b19a-753773c9c3bd/oauth2/token" -Method POST -Body "resource=$([System.Web.HttpUtility]::UrlEncode("https://graph.microsoft.com"))&grant_type=password&client_id=31359c7f-bd7e-475c-86db-fdb8c937548e&scope=openid&username=admin@M365x510994.onmicrosoft.com&password=Rd6E0x7e1N"

Invoke-RestMethod "https://login.windows.net/bb9ca92d-d8ac-4810-b19a-753773c9c3bd/oauth2/token" -Method POST -Body "resource=74658136-14ec-4630-ad9b-26e160ff0fc6&grant_type=password&client_id=1950a258-227b-4e31-a9cf-717495945fc2&scope=openid&username=admin@M365x510994.onmicrosoft.com&password=Rd6E0x7e1N"

#>