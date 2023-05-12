<#
    ###########################################################################
    # Script Name: Outlook_Master_Tool.ps1
    # Author: Callum Stones & Luke Jackson
    # Company: HBP Systems
    # Date Created: 12/05/2023 
    # Version: 1.4
    #
    # Copyright (c) Callum Stones, HBP Systems. All rights reserved.
    # This script is provided "AS IS" without any warranties and is intended
    # for use solely by the author and HBP Systems. Unauthorized copying,
    # reproduction, modification, or distribution is strictly prohibited
    # without the express written consent of Callum Stones and HBP Systems.
    ###########################################################################
#>

$omtLastUpdate = "May 2023"

$LocalPath = (Get-CimInstance Win32_ComputerSystem).UserName

Function UserEndMSProcess {
    do {
        $response = Read-Host "Script will terminate all MS Apps, continue? (Y/N)"
    } until ($response -eq "Y" -or $response -eq "N")

    if ($response -eq "N") {
        Write-Host "Script terminated." -ForegroundColor Red
        Start-Sleep -Seconds 2
    exit
    }
}

UserEndMSProcess

function EndMSProcess {

    $msprocesses = @("WINWORD", "EXCEL", "POWERPNT", "OUTLOOK", "TEAMS", "MSACCESS", "MSPUB")

    foreach ($msprocess in $msprocesses) {
        $runningmsProcesses = Get-Process -Name $msprocess -ErrorAction SilentlyContinue
    
        if ($runningmsProcesses) {
            Write-Host "Killing process: $process"
            $runningmsProcesses | Foreach-Object { $_.Kill() }
        } else {
            Write-Host "Process not found: $process"
        }
    }
    
}

EndMSProcess

function Show-LastUpdate {
    Write-Host "Outlook Master Tool (v1.4) by Callum Stones & Luke Jackson"
    Write-Host "Tool last updated: $omtLastUpdate" -ForegroundColour Green
}

function Show-Uptime {
    $OS = Get-WmiObject Win32_OperatingSystem
    $uptime = (Get-Date) - $OS.ConvertToDateTime($OS.LastBootUpTime)
    $UpTimeInt = [int]$uptime.Days
    if ($UpTimeInt -gt 3) {
        Write-Host "Device has been up $UpTimeInt days. It is recommended to reboot before proceeding." -ForegroundColor Red
    } else {
        Write-Host "Device has been up $UpTimeInt days" -ForegroundColor Yellow
    }
}

function Show-LastInstallDate {
    $LastInstallDate = Get-HotFix | Sort-Object -Property InstalledOn | Select-Object -Last 1 -ExpandProperty InstalledOn
    Write-Host "The last installed update or hotfix was installed on $($LastInstallDate.ToString('dd MMMM yyyy'))" -ForegroundColor Yellow
}

function Show-Menu {
    
    Write-Host ""
    Write-Host "Select the desired tool:"
    Write-Host "0. Exit"
    Write-Host "1. Outlook ADAL Registry Keys"
    Write-Host "2. Outlook WAM Registry Keys"
    Write-Host "3. MSOAuth For AutoDiscover Registry Keys"
    Write-Host "4. Rewrite AAD Broker Plugin"
    Write-Host "5. Clear Outlook Profiles & Cache"
    Write-Host "6. Credential Manager Tool"
    Write-Host "7. Clear Full Office Cache"
    Write-Host "8. Force Log Out 365 Connected Services On Workstation"
    Write-Host "9. Force Log Out 365 Connected Services On Azure AD"
    Write-Host "10. Help & Further Options"
}

function Invoke-ADAL {
    
    New-Item -Path HKCU:\SOFTWARE\Microsoft\Office\15.0\Common\Identity\ -Force | Out-Null
    New-ItemProperty –Path HKCU:\SOFTWARE\Microsoft\Office\15.0\Common\Identity\ –Name EnableADAL -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty –Path HKCU:\SOFTWARE\Microsoft\Office\15.0\Common\Identity\ –Name Version -Value 1 -PropertyType DWord -Force | Out-Null
    Write-Host "ADAL Keys written. Restart Outlook." -ForegroundColor Green
    Start-Sleep -Seconds 2
}

function Invoke-WAM {
    
    New-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity\ -Force | Out-Null
    New-ItemProperty –Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity\ –Name DisableADALatopWAMOverride -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty –Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity\ –Name DisableAADWAM -Value 1 -PropertyType DWord -Force | Out-Null
    Write-Host "WAM Keys written. Restart Outlook." -ForegroundColor Green
    Start-Sleep -Seconds 2
}

function Invoke-MSOAuth {
    
    New-Item -Path HKCU:\SOFTWARE\Microsoft\Exchange\ -Force | Out-Null
    New-ItemProperty –Path HKCU:\SOFTWARE\Microsoft\Exchange\ –Name AlwaysUseMSOAuthForAutoDiscover -Value 1 -PropertyType DWord -Force | Out-Null
    Write-Host "MSOAuth Keys written. Restart Outlook." -ForegroundColor Green
    Start-Sleep -Seconds 2
}   
   
function Invoke-AADBroker {
    
    Write-Host "Rewriting AAD Broker Plugin. This can take up to a minute. Please wait..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    
    Write-Host "Stopping CryptoGraphic Services..." -ForegroundColor Yellow
    Stop-Service Crypto | Out-Null
    Start-Sleep -Seconds 2
    
    Write-Host "Removing AAD Broker Cache..." -ForegroundColor Yellow
    Remove-Item -Path "C:\Users\$LocalPath\AppData\Local\Packages*AAD*" -Force -Recurse | Out-Null
    Start-Sleep -Seconds 2
    
    Write-Host "Starting CryptoGraphic Services..." -ForegroundColor Yellow
    Start-Service Crypto | Out-Null
    Start-Sleep -Seconds 2
    
    Write-Host "AAD Broker Plugin Rewritten. Restart Outlook." -ForegroundColor Green
    Start-Sleep -Seconds 2
}
    
function Invoke-ClearProfilesCache {
    Write-Host "If either the cache and/or profile does not exist, an error will occur, this is normal..."
    Start-Sleep -Seconds 2
    
    Write-Host "Removing Profile..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2

    try {
        Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\ -Force -Recurse | Out-Null
        New-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\ -Force | Out-Null
        Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\15.0\Outlook\Profiles\ -Force -Recurse | Out-Null
        New-Item -Path HKCU:\SOFTWARE\Microsoft\Office\15.0\Outlook\Profiles\ -Force | Out-Null
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Host "Check Log - Clear Profile Exception Thrown" -ForegroundColor Red
    }
    
    Start-Sleep -Seconds 2

    Write-Host "Clearing Cache..."

    try {
        Remove-Item -Path "C:\Users\$LocalPath\AppData\Local\Microsoft\Outlook" -Force -Recurse | Out-Null
    }
    catch {
        Write-Host "Check Log - Clear Cache Exception Thrown" -ForegroundColor Red
    }
    Start-Sleep -Seconds 2
    
    Write-Host "Profile/Cache Function Compelte. Restart Outlook." -ForegroundColor Green

    Start-Sleep -Seconds 2
}
    
function Invoke-CredentialManager {
    
    Write-Host "A new window will appear with all stored credentials, delete appropriate entries as required." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    Start-Process rundll32.exe keymgr.dll, KRShowKeyMgr
    Start-Sleep -Seconds 2
}

function Invoke-ClearOfficeCache {
        
        Write-Host "This will clear the Office Cache. Continue? (Y/N)" -ForegroundColor Yellow
        $confirm7 = Read-Host
        
        if ($confirm7 -eq 'Y') {
        Write-Host "Deleting Cache..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Remove-Item -Path "C:\Users\$LocalPath\AppData\Local\Microsoft\Office" -Force -Recurse | Out-Null
        
        Write-Host "Cache Deleted." -ForegroundColor Green
        Start-Sleep -Seconds 2
        }
} 

function Invoke-ForceLogout365ConnectedServicesOnWorkstation {

    Write-Host "Forcing Log Out. This can take a while. Please wait..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    Stop-Service Crypto | Out-Null
    Start-Sleep -Seconds 2
    Remove-Item -Path "C:\Users\$LocalPath\AppData\Local\Packages*AAD*" -Force -Recurse | Out-Null
    Start-Sleep -Seconds 2
    Start-Service Crypto | Out-Null
    Start-Sleep -Seconds 2

    Write-Host "Log out forced. You must now check that 'Connected Work + School Accounts' has disconnected. Ensure you open an Office Application and sign out of any connected accounts using the GUI as this is not possible via script." -ForegroundColor Green
    Read-Host -Prompt "Press any key to continue..." -ForegroundColor Green
}

function Invoke-ForceLogout365ConnectedServicesOnAzureAD {

    Write-Host "Please only run this sub tool on an engineer's device! AzureAD Module should not be connected on an end-user's machine!" -ForegroundColor Yellow
    $confirm9 = Read-Host "Confirm you have read the above (Y to continue)" -ForegroundColor Yellow
    if ($confirm9 -eq 'Y') {

    Install-Module AzureAD

    $azureuser = Read-Host "Type full email of Azure user to sign out" -ForegroundColor Yellow

    Write-Host "Session will prompt for admin credentials for the relevant Azure Environment..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    Connect-AzureAD
    Start-Sleep -Seconds 2

    Write-Host "Revoking sessions for $azureuser ..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    Get-AzureADUser -SearchString $azureuser | Revoke-AzureADUserAllRefreshToken

    Write-Host "All Sessions Revoked!" -ForegroundColor Green
    Start-Sleep -Seconds 2
    }
}

function Invoke-HelpAndFurtherOptions {

    Write-Host "If you have tried all of these tools and it has not fixed your issue, you may want to try some other methods manually:"
    Write-Host "Check Machine Uptime"
    Write-Host "Run Windows Updates"
    Write-Host "Reinstall Office Products"
    Write-Host "Recreate The User Profile"
    Write-Host "Use Microsoft's SARA Tool: https://aka.ms/SaRA-FirstScreen"
    Write-Host "Check Duo 2FA is not interfering with authentication, set a bypass"
    Write-Host "Try the SFC & DISM commandset"
    Write-Host "Try disabling connected experiences in Outlook: https://www.thewindowsclub.com/disable-connected-experiences-in-microsoft-office-365"
    Write-Host ""
    Write-Host "If you have any feedback for the tool please email me at cstones@thehbpgroup.co.uk"
    Write-Host ""
    Read-Host -Prompt "Press any key to go back to the main menu" -ForegroundColor Yellow
}

while ($true) {
    Show-LastUpdate
    Show-Uptime
    Show-LastInstallDate
    Show-Menu
    $selection = Read-Host "Select Tool (0-10)" -ForegroundColor Green
    switch ($selection) {
    '0' { break }
    '1' { Invoke-ADAL }
    '2' { Invoke-WAM }
    '3' { Invoke-MSOAuth }
    '4' { Invoke-AADBroker }
    '5' { Invoke-ClearProfilesCache }
    '6' { Invoke-CredentialManager }
    '7' { Invoke-ClearOfficeCache }
    '8' { Invoke-ForceLogout365ConnectedServicesOnWorkstation }
    '9' { Invoke-ForceLogout365ConnectedServicesOnAzureAD }
    '10' { Invoke-HelpAndFurtherOptions }
    default { Write-Host "Invalid selection. Please try again." -ForegroundColor Red}
    }
    }
