$omtLastUpdate = "September 2022"
Write-Host "Outlook Master Tool (v1.2 Team Beta) by Callum Stones"
Write-Host "QA by Anthony Wood & Luke Jackson"
Write-Host "Tool last updated: $omtLastUpdate"

Start-Sleep -Seconds 2

Clear-Host

$bootuptime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime | Out-Null

$CurrentDate = Get-Date | Out-Null

$uptime = $CurrentDate - $bootuptime | Out-Null

$OS = Get-WmiObject Win32_OperatingSystem

$Uptime = (Get-Date) - $OS.ConvertToDateTime($OS.LastBootUpTime)

$Uptime = ([String]$Uptime.Days)

$UpTimeInt = [int]$Uptime

if ($UpTimeInt -gt 3) {
    Write-Host "Device has been up" $UpTimeInt "days. It is recommended to reboot before proceeding."
}
else {
    Write-Host "Device has been up" $UpTimeInt "days"
}

$LastInstallDate = Get-HotFix | Sort-Object -Property InstalledOn | Select-Object -Last 1 -ExpandProperty InstalledOn
Write-Host "The last installed update or hotfix was installed on $($LastInstallDate.ToString('dd MMMM yyyy'))"

Write-Host ""

Do {
Write-Host "Ensure all Office Apps are closed or errors will occur!"
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

$selection = Read-Host "Select Tool (1-10)"

if ($selection -eq '1')
{
Clear-Host

New-Item -Path HKCU:\SOFTWARE\Microsoft\Office\15.0\Common\Identity\ -Force | Out-Null

New-ItemProperty –Path HKCU:\SOFTWARE\Microsoft\Office\15.0\Common\Identity\ –Name EnableADAL -Value 1 -PropertyType DWord -Force | Out-Null

New-ItemProperty –Path HKCU:\SOFTWARE\Microsoft\Office\15.0\Common\Identity\ –Name Version -Value 1 -PropertyType DWord -Force | Out-Null

Write-Host "ADAL Keys written. Restart Outlook."

Start-Sleep -Seconds 3
}

if ($selection -eq '2') 
{
Clear-Host

New-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity\ -Force | Out-Null

New-ItemProperty –Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity\ –Name DisableADALatopWAMOverride -Value 1 -PropertyType DWord -Force | Out-Null

New-ItemProperty –Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity\ –Name DisableAADWAM -Value 1 -PropertyType DWord -Force | Out-Null

Write-Host "WAM Keys written. Restart Outlook."

Start-Sleep -Seconds 3
}

if ($selection -eq '3') 
{
Clear-Host

New-Item -Path HKCU:\SOFTWARE\Microsoft\Exchange\ -Force | Out-Null

New-ItemProperty –Path HKCU:\SOFTWARE\Microsoft\Exchange\ –Name AlwaysUseMSOAuthForAutoDiscover -Value 1 -PropertyType DWord -Force | Out-Null

Write-Host "MSOAuth Keys written. Restart Outlook."

Start-Sleep -Seconds 3
}

if ($selection -eq '4') 
{
Clear-Host

Write-Host "Rewriting AAD Broker Plugin. This can take upto a minute. Please wait..."

Start-Sleep -Seconds 3

Clear-Host

Write-Host "Stopping CryptoGraphic Services..."

Stop-Service *Crypto* | Out-Null

Start-Sleep -Seconds 15

Clear-Host

Write-Host "Removing AAD Broker Cache..."

Remove-Item -Path "C:\Users\$env:USERNAME\AppData\Local\Packages\*AAD*" -Force -Recurse | Out-Null

Start-Sleep -Seconds 10

Clear-Host

Write-Host "Starting CryptoGraphic Services..."

Start-Service *Crypto* | Out-Null

Start-Sleep -Seconds 10

Clear-Host

Write-Host "AAD Broker Plugin Rewritten. Restart Outlook."

Start-Sleep -Seconds 3
}

if ($selection -eq '5')
{

Write-Host "If either the cache and/or profile does not exist, an error will occur, this is normal..."

Start-Sleep -Seconds 5

Clear-Host

Write-Host "Removing Profile..."

Start-Sleep -Seconds 2

Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\ -Force -Recurse | Out-Null

New-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\ -Force | Out-Null

Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\15.0\Outlook\Profiles\ -Force -Recurse | Out-Null

New-Item -Path HKCU:\SOFTWARE\Microsoft\Office\15.0\Outlook\Profiles\ -Force | Out-Null

Start-Sleep -Seconds 3

Clear-Host

Write-Host "Clearing Cache..."

Remove-Item -Path "C:\Users\$env:USERNAME\AppData\Local\Microsoft\Outlook" -Force -Recurse | Out-Null

Start-Sleep -Seconds 3

Clear-Host

Write-Host "Profile Removed & Cache Cleared. Restart Outlook."

Start-Sleep -Seconds 3
}

if ($selection -eq '6')
{
Clear-Host

Write-Host "A new window will appear with all stored credentials, delete appropriate entries as required."

Start-Sleep -Seconds 4

Start-Process rundll32.exe keymgr.dll, KRShowKeyMgr

Start-Sleep -Seconds 3
}
 
if ($selection -eq '7')
{
Clear-Host

Write-Host "This will clear the Office Cache. Continue? (Y/N)"

$confirm7 = Read-Host

Clear-Host

if ($confirm7 -ne 'Y' ) {

}Else{

Write-Host "Deleting Cache..."

Start-Sleep -Seconds 2

Remove-Item -Path "C:\Users\$env:USERNAME\AppData\Local\Microsoft\Office" -Force -Recurse | Out-Null

Clear-Host

Write-Host "Cache Deleted."

Start-Sleep -Seconds 3
}

}

if ($selection -eq '8') 
{
Clear-Host

Write-Host "Forcing Log Out. This can take a while. Please wait..."

Start-Sleep -Seconds 3

Stop-Service *Crypto* | Out-Null

Start-Sleep -Seconds 15

Remove-Item -Path "C:\Users\$env:USERNAME\AppData\Local\Packages\*AAD*" -Force -Recurse | Out-Null

Start-Sleep -Seconds 10

Start-Service *Crypto* | Out-Null

Start-Sleep -Seconds 10

Clear-Host

Write-Host "Log out forced. You must now check that 'Connected Work + School Accounts' has disconnected. Ensure you open an Office Application and sign out of any connected accounts using the GUI as this is not possible via script."

Read-Host -Prompt "Press any key to continue..."
}

if ($selection -eq '9')
{

Write-Host "Please only run this sub tool on an engineers device! AzureAD Module should not be connected on an end users machine!"
$confirm9 = Read-Host "Confirm you have read the above (Y to continue)"

if ($confirm9 -ne 'Y') {

}else{

Clear-Host

Install-Module AzureAD

Clear-Host

$azureuser = Read-Host "Type full email of Azure user to sign out"

Clear-Host

Write-Host "Session will prompt for admin credentials for the relevant Azure Environment..."

Start-Sleep -Seconds 3

Connect-AzureAD

Start-Sleep -Seconds 8

Clear-Host

Write-Host "Revoking sessions for $azureuser ..."

Start-Sleep -Seconds 3

Get-AzureADUser -SearchString $azureuser | Revoke-AzureADUserAllRefreshToken

Clear-Host

Write-Host "All Sessions Revoked!"

Start-Sleep -Seconds 3
}
}

if ($selection -eq '10') 
{

Clear-Host

Write-Host "If you have tried all of these tools and it has not fixed your issue, you may want to try some other methods manually:"
Write-Host "Check Machine Uptime"
Write-Host "Run Windows Updates"
Write-Host "Reinstall Office Products"
Write-Host "Recreate The User Profile"
Write-Host "Use Microsoft's SARA Tool: https://aka.ms/SaRA-FirstScreen"
Write-Host "Check Duo 2FA is not intefering with authentication, set a bypass"
Write-Host "Try the SFC & DISM commandset"
Write-Host "Try disabling connected experiences in Outlook: https://www.thewindowsclub.com/disable-connected-experiences-in-microsoft-office-365"
Write-Host ""
Write-Host "If you have any feedback for the tool please email me at cstones@thehbpgroup.co.uk"
Write-Host ""
Read-Host -Prompt "Press any key to go back to the main menu" 

}

Clear-Host

Start-Sleep -Seconds 1

}While ($selection -ne 0)




