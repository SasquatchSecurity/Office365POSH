if(Get-PSSession){
    Write-Host "Already connected to 365, proceeding" -ForegroundColor green
}
else{
    $MFAExchangeModule = ((Get-ChildItem $Env:LOCALAPPDATA\Apps\2.0\*\CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Target -First 1).Replace("CreateExoPSSession.ps1", ""))
    Write-Host "Importing Exchange MFA Module"
    . "$MFAExchangeModule\CreateExoPSSession.ps1"

    Write-Host "Connecting to Exchange Online with MFA"
    Connect-EXOPSSession 
}

$openFileDialog = New-Object windows.forms.openfiledialog   
$openFileDialog.initialDirectory = (Get-Location).path
$openFileDialog.title = "Select CSV File with list of users"   
$openFileDialog.filter = "CSV Files (*.csv)|*.csv|All files (*.*)|*.*"   
$openFileDialog.ShowHelp = $True   
Write-Host "Select CSV File... (see File Open Dialog)" -ForegroundColor Green  
$open = $openFileDialog.ShowDialog()   
# Display the Dialog / Wait for user response 
$open 
if($open -eq "OK")    {    
        Write-Host "Selected CSV File:"  -ForegroundColor Green  
        $OpenFileDialog.filename   
        # $OpenFileDialog.CheckFileExists 
        Write-Host "CSV File Imported!" -ForegroundColor Green 
    } 
    else { Write-Host "CSV File Import Cancelled! Cancelling." -ForegroundColor Yellow
exit
}

# set where to save output
$saveFileDialog = New-Object windows.forms.savefiledialog
$saveFileDialog.title = "Save output to file"   
$saveFileDialog.filter = "CSV Files (*.csv)|*.csv|All files (*.*)|*.*"   
$saveFileDialog.ShowHelp = $True
Write-Host "Please Name Your File... (see File Save Dialog)" -ForegroundColor Green  
$save = $saveFileDialog.ShowDialog()   
# Display the Dialog / Wait for user response 
$save 
if($save -eq "OK")    {    
        Write-Host "Selected CSV file:"  -ForegroundColor Green  
        $saveFileDialog.filename   
        Write-Host "Filename selected" -ForegroundColor Green 
    } 
    else { Write-Host "Save location selection cancelled!" -ForegroundColor Yellow}

# set CSV as variable to parse through
$userlist = Get-Content $openFileDialog.filename

# Run Inbox Rule against users in CSV
if($saveFileDialog.filename)
{
foreach($username in $userlist) 
{
Get-InboxRule -Mailbox $username | Select-Object MailboxOwnerId,Name,Enabled,DeleteMessage,Description,SubjectorBodyContainsWords | Where DeleteMessage -Match "True" | Export-CSV $saveFileDialog.filename -NoTypeInformation -Append 
    }
# Write status
Write-Host "Rule export complete, opening" -ForegroundColor Green
# Open file
Invoke-Item $saveFileDialog.filename
}
    else {Write-Host "Save location not selected. Ending." -ForegroundColor Yellow}

