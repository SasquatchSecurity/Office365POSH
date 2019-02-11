try
{
    Get-SPOTenant -ErrorAction Stop > $null
    
}
catch 
{
$UserCredential = Get-Credential
$tenant = Read-Host -Prompt "Enter your tenant name"
$tenanturl = "https://" + $tenant + "-admin.sharepoint.com"
Connect-AzureAD -Credential $UserCredential -ErrorAction Stop
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Connect-SPOService -Url $tenanturl -credential $UserCredential
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $UserCredential -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession
}




#Clear O365 Sessions for affected user
$officeuser = Read-Host -Prompt 'Enter the full BGSU username with @bgsu.edu'
Revoke-SPOUserSession -User $officeuser -Confirm:$false

#$session = Connect-AzureAD -Credential $UserCredential -ErrorAction Stop
Get-AzureADUser -SearchString $officeuser | Revoke-AzureADUserAllRefreshToken

#run the move mailbox to hopefully clear out the session quicker
#clear error so if it catches a moverequest pending we can clear it
$error.clear()
New-MoveRequest -Identity $officeuser | Select-Object -Property MailboxIdentity,DisplayName,Status

#run the final loop to catch if there is a pending move request
if ($error)
{
Write-Output "Found a pending move request, removing the request and running movemailbox again"
Remove-MoveRequest -Identity $officeuser -Confirm:$false
New-MoveRequest -Identity $officeuser
}