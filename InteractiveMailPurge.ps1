$localpath = Get-Location
$diditrun = Get-PSSession
if ($diditrun)
{
#dont do anything if already connected to 365 powershell
Write-Host "Already connected to 365, proceeding" -ForegroundColor green

}
else
{
# Import EXOP Modules
# Import AD Module
Try {
Import-Module $localpath\EXOP\Microsoft.IdentityModel.Clients.ActiveDirectory.dll -Verbose
} Catch {
	Write-Error -Message "Error. $_" -ErrorAction Stop
	}	
# Import EXOP Module
Try {
Import-Module $localpath\EXOP\Microsoft.Exchange.Management.ExoPowershellModule.dll -Verbose
} Catch {
	Write-Error -Message "Error. $_" -ErrorAction Stop
	}	

# Get login credentials 
Write-Host "Enter your full shell username" -ForegroundColor green
$UPN = Read-Host -Prompt 'Username'
Try {
$Session = New-ExoPSSession -UserPrincipalName $UPN -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid
} Catch {
        Write-Error -Message "Error. $_" -ErrorAction Stop
}
Import-PSSession $Session -AllowClobber -DisableNameChecking 
$Host.UI.RawUI.WindowTitle = "Office 365 Security & Compliance Center" 
}

#Grab the name of the search you want to purge

#other way of getting it inline 
#$Purgename = Read-Host -Prompt 'Input the search name here'

#New method to grab search name
Write-Host "Enter the Compliance Search Name" -ForegroundColor green
$Purgename = Read-Host -Prompt 'Name'

#Run the purge below searchname must be exact

New-ComplianceSearchAction -SearchName "$Purgename" -Purge -PurgeType SoftDelete
