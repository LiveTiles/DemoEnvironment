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

.EXAMPLE
    .\Install-LiveTilesHub -UseWebLogin
    
    Creates Reach Subscription and Installs Hub on the root site collection using MFA enabled login 

.EXAMPLE
    .\Install-LiveTilesHub -SharePointTenantAdmin john.doe@fabricam.onmicrosoft.com

    Creates Reach Subscription and Installs Hub on the root site collection using the John Doe credentials. Will prompt for password. 
    
.EXAMPLE
    .\Install-LiveTilesHub -SharePointTenantAdmin john.doe@fabricam.com -o365Tenant "othertenant" -hubSiteCollection "livetileshub"
    
    Creates Reach Subscription and Installs Hub on <othertenant>/sites/<hubsitecollection> using the John Doe credentials. Will prompt for password. 


.NOTES
    AUTHOR: Christoffer Soltau
    LASTEDIT: 15-06-2021 
    v0.9
        First Release. Identified 2Do´s:
            - Actually test the MFA process :-)
	        - Add parameters to control which LiveTiles Hub modules are installed, if this is to be used for other than Demo
	        - OutPut formatting. Refine maxlength to match longest possible output
	        - After creating App Catalog, Could a relogin to PnP suffice?
	        - After search config import, Better error message - perhaps a description of what needs to be done manually?
	        - Creation of Governance license creation. Untested runtime - may have to be placed before subscription link??
	        - Check if hubsite collection actually exists - if not, create it?
            - Install Modules dosen't work from x64 ISE. It launches x86 PS and installs 32 bit version, not x64 

.LINK
    Updated versions of this script will be available on the LiveTiles Partner Portal
#>
[cmdletbinding(SupportsShouldProcess=$True)]
param (
    #2Do - Actually test the MFA process :-)
    #2Do - add parameters to control which LiveTiles Hub modules are installed, if this is to be used for other than Demo
    [Parameter(Mandatory=$true, ParameterSetName="MFA")]
    [switch]$UseWebLogin,
    #[Parameter(Mandatory=$true, ParameterSetName="MFA")]
    [Parameter(Mandatory=$false, ParameterSetName = "nonMFA")]
    [string]$SharePointTenantAdmin, # = "admin@M365x510994.onmicrosoft.com",
    #[Parameter(Mandatory=$false, ParameterSetName="MFA")]
    [Parameter(Mandatory=$false, ParameterSetName = "nonMFA")]
    [string]$SharePointTenantAdminPassword, # = "Rd6E0x7e1N",
    #[Parameter(Mandatory=$false)]
    #[switch]$skipCheckPrerequisites = $True,
    [Parameter(Mandatory=$false)]
    [string]$o365Tenant = "",
    [Parameter(Mandatory=$false)]
    [string]$hubSiteCollection #if not used, assume root site coll?
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
#endregion

Write-Host ("Installation starttime: " + (get-date)) -ForegroundColor White

#region Set Constants
    $outputMaxLength = 10 + (77)
    #2Do: Refine maxlength to match longest possible output

    if ($outputMaxLength -le 100) {$outputMaxLength = 100}
    $waitVeryLong = 60
    $waitLong = 10
    $waitShort = 5
#endregion

#region Install or Update required PowerShell Modules and components
    #if (-not $skipCheckPrerequisites.IsPresent) {
        $output = Write-LiveTilesHost -messageType PROCESS -message ("Checking Powershell Module prerequisites")
        $dotCount = 0
        $moduleList = @(("PnP.PowerShell", "1.6.0"), ("AzureAD", "2.0.2.135"), ("Az.Accounts","2.3.0"), ("AzureRM.profile","5.8.4"))
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

            if (-not $moduleInstalled -and $module[0] -eq "AzureRM.profile") {
                $moduleInstalled = @(Get-Module -ListAvailable -Name "Az.Accounts")
                if ($moduleInstalled) {
                    $module[0] = "Az.Accounts"
                    $module[1] = "2.3.0"
                }
            } elseif (-not $moduleInstalled -and $module[0] -eq "Az.Account") {
                $moduleInstalled = @(Get-Module -ListAvailable -Name "AzureRM.profile")
                if ($moduleInstalled) {
                    $module[0] = "AzureRM.profile"
                    $module[1] = "5.8.4"
                }
            }

            switch ($module[0]) {
                "Az.Accounts" {$AzureLoginMethod = "Az"}
                "AzureRM.profile" {
                    $azInstalled = @(Get-Module -ListAvailable -Name "Az.Accounts")
                    if ($azInstalled) {$AzureLoginMethod = "Az"                    
                    } else {$AzureLoginMethod = "AzureRM"}
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
                        start-process powershell -Verb RunAs "try {Install-Module $module[0] -AllowClobber -Force} catch {Install-Module $module[0] -AllowClobber -Force -SkipPublisherCheck}" -wait
                    }
	            }
            } else {
                Write-Host
                $inputString = "Confirm`n`r" + $module[0] + " module is not installed. Do you want to Install?`n`rThe process will open in a new window to instal as 'admin'`r`n[Y] Yes  [N] No (default is 'Y')"
                $input = Read-Host -prompt $inputString
                if ($input -eq $null) {$input = "y"}
                if ($input.ToString().ToLower() -ne "n") { 
    		        start-process powershell -Verb RunAs "try {Install-Module $module[0] -AllowClobber -Force} catch {Install-Module $module[0] -AllowClobber -Force -SkipPublisherCheck}" -Wait
                } else {
                    Write-Host "Module prerequisites not met. Exiting installation script."
                    Exit
                }
            }
        }
        Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength ($output+$dotCount) -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
#   }
#endregion

#region Base variables
    foreach ($module in $moduleList) {
        if ($module[0] -eq "Az.Accounts" -or $module[0] -eq "AzureRM.profile") {
            switch ($AzureLoginMethod) {
                "Az" {remove-module AzureRM.Profile -Force -ErrorAction SilentlyContinue
                    Import-Module Az.Accounts        
                    }
                "AzureRM" {remove-module Az.Accounts -Force -ErrorAction SilentlyContinue
                    Import-Module AzureRM.profile
                    }
            }
        } else {
            Import-Module $module[0]
        }
    }
    
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
#endregion

#region Logging in to SharePoint Admin

    try {
        if (-not $useWebLoginSharePoint.IsPresent) {
            $output = Write-LiveTilesHost -messageType PROCESS -message ("Logging in to SharePoint Admin as " + $SharePointTenantAdmin)
                
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
            $pnpSharePointAdminContext = Connect-PnPOnline -ReturnConnection -Url $SharePointAdminUrl -UseWebLogin #-Interactive instead?
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        }
    } catch {
        Write-LiveTilesHost -messageType ERROR -outputMaxLength $outputMaxLength -initialStringLength $output -afterOutputMessage "Cannot login. Insufficient rights? Exiting script"
        Exit
    }

#endregion

#region Get Azure AD internal API token
    $output = Write-LiveTilesHost -messageType PROCESS -message ("Getting Azure AD API Token")
    if ($AzureLoginMethod -eq "Az") {
        #Avoid dll lock - weird error that occurs when running the script "too fast"?
        do {
            try {
                $o365Context = Login-AzAccount -TenantId $pnpSharePointAdminContext.Tenant -Credential $credentialsOffice365Admin -WhatIf:$false
                $AzcontextCreated = $true
            } catch {
                $AzcontextCreated = $false
                Start-Sleep -Seconds $waitShort
            }
        } until ($AzcontextCreated)
        $Azcontext = get-azcontext
        #$GraphapiToken = ([Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate($Azcontext.Account, $Azcontext.Environment, $Azcontext.Tenant.Id, $null, "Never", $null, "ee62de39-b9b0-4886-aa58-08b89c4e3db3")).AccessToken
        $AADapiToken = ([Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate($Azcontext.Account, $Azcontext.Environment, $Azcontext.Tenant.Id, $null, "Never", $null, "74658136-14ec-4630-ad9b-26e160ff0fc6")).AccessToken
        
        
    } else {
        $o365Context = Login-AzureRmAccount -TenantId $pnpSharePointAdminContext.Tenant -Credential $credentialsOffice365Admin -WhatIf:$false

        $o365Token = $o365Context.Context.TokenCache.ReadItems().where({($_.tenantid -eq $pnpSharePointAdminContext.Tenant) -and ($_.displayableid -eq $o365Context.Context.Account) -and $_.resource -eq "https://management.core.windows.net/"})[-1].AccessToken
        $refreshToken = $o365Context.Context.TokenCache.ReadItems().where({($_.tenantid -eq $pnpSharePointAdminContext.Tenant) -and ($_.displayableid -eq $o365Context.Context.Account)})[-1].refreshtoken

        #Doesn't get token when not using azurermlogin
        $body = "grant_type=refresh_token&refresh_token=$($refreshToken)&resource=74658136-14ec-4630-ad9b-26e160ff0fc6"
        $AADapiToken = (Invoke-RestMethod "https://login.windows.net/$($pnpSharePointAdminContext.Tenant)/oauth2/token" -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded').access_token
    }
    Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
#endregion

#region Checking if Application Catalog exists
    $output = Write-LiveTilesHost -messageType PROCESS -message "Checking if Application Catalog exists" -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
            $appCatalogUrl = Get-PnPTenantAppCatalogUrl -Connection $pnpSharePointAdminContext
        if ($appCatalogUrl -eq "" -or $appCatalogUrl -eq $null) {
            Write-LiveTilesHost -messageType WARNING -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference -afterOutputMessage "The App Catalog doesn't exist. Creating it..."
            
            Get-PnPTimeZoneId | Out-Host
            $appCatalogUrl = $SharePointUrl + "/sites/appcatalog"
            if ($PSCmdlet.ShouldProcess("AppCatalogSite ", "Create")) {
                Register-PnPAppCatalogSite -Url $appCatalogUrl -Owner $SharePointTenantAdmin -Connection $pnpSharePointAdminContext # -TimeZoneId 4
            }

            write-host "App Catalog created successfully. Please rerun the script..."
            #2Do - Could a relogin to PnP suffice?
            Exit
        } else {
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        }

#endregion

#region Download / Upload LiveTiles Packages
    $output = Write-LiveTilesHost -messageType PROCESS -message "Getting LiveTiles installation files and deploying packages to SharePoint" -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
    try {
        $fileListPublishPackages = @(
            "https://install.matchpoint365.com/resources/LiveTiles.Intranet.Hub.sppkg",
            "https://install.matchpoint365.com/resources/LiveTiles.Intranet.Governance.sppkg",
            "https://install.matchpoint365.com/resources/LiveTiles.Everywhere.Panel.sppkg"
        )
        $fileListUploadPackages = @(
            "https://install.matchpoint365.com/resources/LiveTiles.Intranet.Hub.LandingPage.sppkg",
            "https://install.matchpoint365.com/resources/LiveTiles.Intranet.Workspaces.sppkg",
            "https://install.matchpoint365.com/resources/LiveTiles.Intranet.Metadata.sppkg"
        )
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
                Add-PnPApp -Path $fileName -Publish -Connection $pnpSharePointAdminContext -Overwrite | Write-Verbose
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
    #Reach AppId "d492530a-8cff-481c-90da-9c3c3f1be7da"
    #Condense API AppId "e40ee8da-4b99-4273-b99f-a6b2f10ac29e"
    #Condense AppId "02de0a9e-a2a9-42af-a6c3-2ed77e913a4b"
    #MatchPoint Hub AppId "14a6b046-c3d9-4988-aa99-f0804587f299"
    #MatchPoint Governance AppId "b6a4f91d-466c-4afe-bb29-3f41ea615da7"
    #MatchPoint Metadata AppId "337ea7aa-6b44-4d3c-b5c3-52aaeb1d5dd8"
    #MatchPoint Provisioning AppId "3b8f4e7e-93b5-43ed-830d-eada7c8ff81f"
    #MatchPoint Workspaces AppId "a9a0e8f6-2cee-42d2-b00f-c9299e509958"

    $appIdsLiveTiles = @(
        "d492530a-8cff-481c-90da-9c3c3f1be7da",
        "e40ee8da-4b99-4273-b99f-a6b2f10ac29e",
        "02de0a9e-a2a9-42af-a6c3-2ed77e913a4b",
        "14a6b046-c3d9-4988-aa99-f0804587f299",
        "b6a4f91d-466c-4afe-bb29-3f41ea615da7",
        "337ea7aa-6b44-4d3c-b5c3-52aaeb1d5dd8",
        "3b8f4e7e-93b5-43ed-830d-eada7c8ff81f",
        "a9a0e8f6-2cee-42d2-b00f-c9299e509958"
    )
    
    #Creating Service Principals
    foreach ($appId in $appIdsLiveTiles) {
        if (($output + $dotCount) -le ($outputMaxLength - 2)) {
            Write-Host -NoNewline "." -ForegroundColor White
            $dotCount++
        }
        Grant-OAuth2PermissionsToApp -azureAppId $appId -apiToken $AADapiToken|write-verbose
    }

    #Waiting until Service Principals are created
    $completed = $false
    do {
        $AADcontext = Connect-AzureAD -Credential $credentialsOffice365Admin
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

    Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength ($output+$dotCount) -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference

#Region Creating Subscriptions

    $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "Opening browser to create Reach Subscription - please return here afterwards." -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference

    sleep -Seconds $waitShort
    Start-Process "https://app.condense.ch/subscribe"
    Write-Host
    pause
    Write-Host

    $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "Opening browser to create LiveTiles Subscriptions - No action needed." -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
    #sleep -Seconds $waitShort
    Start-Process "https://hub.matchpoint365.com/api/manage/register"
    Start-Process "https://governance.matchpoint365.com/api/governance/register"
    Start-Process "https://metadata.matchpoint365.com/api/metadata/register"
    Start-Process "https://provisioning.matchpoint365.com/api/provision/register"
    Start-Process "https://workspaces.matchpoint365.com/api/workspaces/register"
    
#endregion

#region Grant Consent between apps
    $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "Opening browser to grant Reach consent to access Hub - please return here afterwards." -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
    sleep -Seconds $waitShort
    Start-Process "https://login.microsoftonline.com/common/oauth2/authorize?response_type=id_token&client_id=02de0a9e-a2a9-42af-a6c3-2ed77e913a4b&redirect_uri=https%3A%2F%2Fapp.condense.ch/logout"
    Write-Host
    pause
    Write-Host

    $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "Opening browser to grant Workspaces consent to access Metadata - please return here afterwards." -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
    sleep -Seconds $waitShort
    Start-Process "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=a9a0e8f6-2cee-42d2-b00f-c9299e509958&response_type=id_token%20token&scope=https://iurcycl.onmicrosoft.com/matchpoint-metadata/user_impersonation&redirect_uri=https://workspaces.matchpoint365.com/api/workspaces/register"
    Write-Host
    pause
    Write-Host

    $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "Opening browser to grant Workspaces consent to access Provisioning - please return here afterwards." -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
    sleep -Seconds $waitShort
    Start-Process "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=a9a0e8f6-2cee-42d2-b00f-c9299e509958&response_type=id_token%20token&scope=https://iurcycl.onmicrosoft.com/matchpoint-provisioning/user_impersonation&redirect_uri=https://workspaces.matchpoint365.com/api/workspaces/register"
    Write-Host
    pause
    Write-Host

    #2Do - Untested runtime - may have to be placed before subscription link??
    $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "Opening browser to grant Governance consent to access SharePoint - please return here afterwards." -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
    sleep -Seconds $waitShort
    Start-Process "https://login.microsoftonline.com/common/oauth2/authorize?client_id=b6a4f91d-466c-4afe-bb29-3f41ea615da7&redirect_uri=https%3A%2F%2Fgovernance.matchpoint365.com%2F&response_type=id_token&scope=openid%20profile&response_mode=form_post&nonce=637589315897098371.NDdmM2JjZjMtZjllMy00NTgyLWI3MjctOGE4MmIyNDFiNzBlZjJiYjUwNGYtMjZiYS00OTgwLTk1ZWUtMjE5ZGE2ZDFlYjQ5&state=CfDJ8LRbmffsmjNCqinaRfnCoRvfwFPJ109R2f6GGWj_wM8uZQiKPHIT--3jIsiMFsFChjtSZTJSvWGF9suCktafrskUdJ96RbSCyl4eLHIcQ2Yw6emt-wpDnOmUA9jobArDn0-zd_2HKy6B1GMAIcCGOn9BkntyXx4W56guWvmJatuxvBhTqfaMcBJIr8W3WxUE6cGno5eCM-rGTovP42jMBzrsRWrsBQyXRKGvgKKYHUz8UnHq2Oikp2-gKrygUGL7UK1ZDPlD5zDDn44Yhv2B0RyQfrHVE8nYw3vMJYW--iJt0SFa-pQJhxrC-XBbeh6cRcizmwF70XIG0jBZa5fbtu_UeeJKi021WapSolby_54WkIiIZFxW-oTX3YwS-Ie69A&x-client-SKU=ID_NET461&x-client-ver=5.3.0.0"
    Write-Host
    pause
    Write-Host

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
            if ($PSCmdlet.ShouldProcess("PrincipalPermissionRequest", "Approve")) {
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
    
    #2Do: Check if hubsite collection actually exists - if not, create it?
    try {
        if (-not $useWebLoginSharePoint.IsPresent) {
            $output = Write-LiveTilesHost -messageType PROCESS -message ("Logging in to SharePoint Hub Site as " + $SharePointTenantAdmin)
                $pnpSharePointHubSiteContext = Connect-PnPOnline -ReturnConnection -Url $hubSiteCollection -Credentials $credentialsOffice365Admin
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        } else {
            $output = Write-LiveTilesHost -messageType PROCESSATTENTION -message "PnPOnline-Module: Sign in to Office 365 with SharePoint Admin rights"
                $pnpSharePointHubSiteContext = Connect-PnPOnline -ReturnConnection -Url $hubSiteCollection -UseWebLogin
            Write-LiveTilesHost -messageType OK -outputMaxLength $outputMaxLength -initialStringLength $output -WhatIf:$WhatIfPreference -Verbose:$VerbosePreference
        }
    } catch {
        Write-LiveTilesHost -messageType ERROR -outputMaxLength $outputMaxLength -initialStringLength $output -afterOutputMessage "Cannot login. Insufficient rights? Exiting script"
        Exit
    }
#endregion

#region Add Apps to Site Collection
    $hubSiteAdminList = @((Get-PnPSiteCollectionAdmin -Connection $pnpSharePointHubSiteContext).LoginName)
    if (("i:0#.f|membership|"+$pnpSharePointHubSiteContext.PSCredential.UserName) -notin $hubSiteAdminList) {
        if ($PSCmdlet.ShouldProcess("SiteCollectionAdmin", "Add")) {
            Add-PnPSiteCollectionAdmin -Owners $pnpSharePointHubSiteContext.PSCredential.UserName -Connection $pnpSharePointHubSiteContext
        }
    }

    $availableApps = Get-PnPApp -Scope Tenant -Connection $pnpSharePointHubSiteContext

    if ($PSCmdlet.ShouldProcess("Site Collection Apps", "Install")) {
        Install-PnPApp -Scope Tenant -Connection $pnpSharePointHubSiteContext -Identity $availableApps.where({$_.Title -eq "LiveTiles Intranet Hub Landing Page"}).Id
        Install-PnPApp -Scope Tenant -Connection $pnpSharePointHubSiteContext -Identity $availableApps.where({$_.Title -eq "LiveTiles Intranet Metadata"}).Id
        Install-PnPApp -Scope Tenant -Connection $pnpSharePointHubSiteContext -Identity $availableApps.where({$_.Title -eq "LiveTiles Intranet Workspaces"}).Id
    }
#endregion

#region Add webparts to frontpage

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
        Add-PnPWebPartToWebPartPage -ServerRelativePageUrl "/SitePages/Hub.aspx" -Xml $xmlGovernanceWebpart -ZoneId "Zone 1" -ZoneIndex "0" -Connection $pnpSharePointHubSiteContext
        Add-PnPWebPartToWebPartPage -ServerRelativePageUrl "/SitePages/Hub.aspx" -Xml $xmlWorkspacesWebpart -ZoneId "Zone 1" -ZoneIndex "0" -Connection $pnpSharePointHubSiteContext
        Add-PnPWebPartToWebPartPage -ServerRelativePageUrl "/SitePages/Hub.aspx" -Xml $xmlHubWebPart -ZoneId "Zone 1" -ZoneIndex "0" -Connection $pnpSharePointHubSiteContext
    }
#endregion

Write-Host ("Installation endtime: " + (get-date)) -ForegroundColor White