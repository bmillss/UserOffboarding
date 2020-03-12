#Asks for the off boarded user until it is confirmed correct by the user
#Include domain for the user (@sentryone.com, @sentryone.ie, @sentryone.de)
Do {
    $email = Read-Host -prompt 'What is the email of the offboarded user?'
    #Defining variables used in the prompt
    $question1    = 'Confirm this is correct?'
    $choices1  = '&Yes', '&No'
    
    #Creates a UI prompt with 2 options for the user to select
    $answer1 = $Host.UI.PromptForChoice($question1, $email, $choices1, 1)
    }
    Until ($answer1 -eq '0')
    
    #Asks for Admin Username and password
    Write-Host "`nPrompting for O365 login, use your @sentryone.com address" -ForegroundColor Blue
    Start-Sleep 1
    Do {
    If($Null -eq $O365Cred){
        $O365Cred = $Host.ui.PromptForCredential("","Enter your OFFICE 365 admin creds","","")
    }
    $Auser = $o365Cred | Select-Object Username | ForEach-Object {$_.Username}
    #Defining variables used in the prompt
    $question2    = 'Confirm this is correct?'
    $choices2  = '&Yes', '&No'
    
    #Creates a UI prompt with 2 options for the user to select
    $answer1 = $Host.UI.PromptForChoice($question2, $Auser, $choices2, 1)
    }
    Until ($answer1 -eq '0')
        #Seperates username from O365creds
        $Auser = $o365Cred | Select-Object Username | ForEach-Object {$_.Username}
    #Username and credentials for a global admin
    $departinguser = $email
    $destinationuser = "IT.Offboarding@sentryone.com"
    $globaladmin = $Auser
    $credentials = Get-Credential -Credential $globaladmin
    
    #Connects to Microsoft Online service through powershell
    Connect-MsolService -Credential $credentials
    
    #Used to determine default domain for the organization
    $InitialDomain = Get-MsolDomain | Where-Object {$_.IsInitial -eq $true}
    
    #Setting sharepointURL with initial domain.
    $SharePointAdminURL = "https://$($InitialDomain.Name.Split(".")[0])-admin.sharepoint.com"
    
    #Formats the names for both the destination OnedDrive and the offboarding user
    $departingUserUnderscore = $departinguser -replace "[^a-zA-Z]", "_"
    $destinationUserUnderscore = $destinationuser -replace "[^a-zA-Z]", "_"
    
    #Defines both the destination OneDrive and the offboarding user
    $departingOneDriveSite = "https://$($InitialDomain.Name.Split(".")[0])-my.sharepoint.com/personal/$departingUserUnderscore"
    $destinationOneDriveSite = "https://$($InitialDomain.Name.Split(".")[0])-my.sharepoint.com/personal/$destinationUserUnderscore"
    
    #Connects to Sharepoint Online and authenticates using O365 admin credentials
    Write-Host "`nConnecting to SharePoint Online" -ForegroundColor Blue
    Connect-SPOService -Url $SharePointAdminURL -Credential $credentials
    
    # Set current admin as a Site Collection Admin on both OneDrive Site Collections
    Write-Host "`nAdding $globaladmin as site collection admin on both OneDrive site collections" -ForegroundColor Blue
    Set-SPOUser -Site $departingOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $true
    Set-SPOUser -Site $destinationOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $true
      
    Write-Host "`nConnecting to $departinguser's OneDrive via SharePoint Online PNP module" -ForegroundColor Blue
    
    #Connects to the offboarding user's OneDrive via the PnP module built into the sharepoint module for powershell
    Connect-PnPOnline -Url $departingOneDriveSite -Credentials $credentials
      
    Write-Host "`nGetting display name of $departinguser" -ForegroundColor Blue
    
    # Get name of departing user to create folder name.
    $departingOwner = Get-PnPSiteCollectionAdmin | Where-Object {$_.loginname -match $departinguser}
      
    # If there's an issue retrieving the departing user's display name, set this one.
    if ($departingOwner -contains $null) {
        $departingOwner = @{
            Title = "Departing User"
        }
    }
      
    # Define relative folder locations for OneDrive source and destination
    $departingOneDrivePath = "/personal/$departingUserUnderscore/Documents"
    $destinationOneDrivePath = "/personal/$destinationUserUnderscore/Documents/$($departingOwner.Title)'s Files"
    $destinationOneDriveSiteRelativePath = "Documents/$($departingOwner.Title)'s Files"
      
    Write-Host "`nGetting all items from $($departingOwner.Title)" -ForegroundColor Blue
    
    # Get all items from source OneDrive
    $items = Get-PnPListItem -List Documents -PageSize 1000
    
    #This will write all files too large to transfer via script to a .txt file
    $largeItems = $items | Where-Object {[long]$_.fieldvalues.SMTotalFileStreamSize -ge 261095424 -and $_.FileSystemObjectType -contains "File"}
    if ($largeItems) {
        $largeexport = @()
        foreach ($item in $largeitems) {
            $largeexport += "$(Get-Date) - Size: $([math]::Round(($item.FieldValues.SMTotalFileStreamSize / 1MB),2)) MB Path: $($item.FieldValues.FileRef)"
            Write-Host "File too large to copy: $($item.FieldValues.FileRef)" -ForegroundColor DarkYellow
        } 
        #This is defining the path in which the file files that the files that failed to transfer will be saved
        $largeexport | Out-file C:\temp\largefiles.txt -Append
        Write-Host "A list of files too large to be copied from $($departingOwner.Title) have been exported to C:\temp\LargeFiles.txt" -ForegroundColor Yellow
    }
    
    #Gets the files that are able to be transfered with a script  
    $rightSizeItems = $items | Where-Object {[long]$_.fieldvalues.SMTotalFileStreamSize -lt 261095424 -or $_.FileSystemObjectType -contains "Folder"}
    
    #Connects to offboarding user's Onedrive  
    Write-Host "`nConnecting to $destinationuser via SharePoint PNP PowerShell module" -ForegroundColor Blue
    Connect-PnPOnline -Url $destinationOneDriveSite -Credentials $credentials
    
    # Filter by Folders to create directory structure
    Write-Host "`nFilter by folders" -ForegroundColor Blue
    $folders = $rightSizeItems | Where-Object {$_.FileSystemObjectType -contains "Folder"}
    #Recreates the directory structure from the offboarding user to the IT.Offboarding OneDrive
    Write-Host "`nCreating Directory Structure" -ForegroundColor Blue
    foreach ($folder in $folders) {
        $path = ('{0}{1}' -f $destinationOneDriveSiteRelativePath, $folder.fieldvalues.FileRef).Replace($departingOneDrivePath, '')
        Write-Host "Creating folder in $path" -ForegroundColor Green
        $newfolder = Resolve-PnPFolder -SiteRelativePath $path
    }
      
     #Begins copying files from the offboarding user to the IT.Offboarding's OneDrive
    Write-Host "`nCopying Files" -ForegroundColor Blue
    $files = $rightSizeItems | Where-Object {$_.FileSystemObjectType -contains "File"}
    $fileerrors = ""
    foreach ($file in $files) {
        #Sets the destination path using previous variable's, also logs any files that fail to copy 
        $destpath = ("$destinationOneDrivePath$($file.fieldvalues.FileDirRef)").Replace($departingOneDrivePath, "")
        Write-Host "Copying $($file.fieldvalues.FileLeafRef) to $destpath" -ForegroundColor Green
        $newfile = Copy-PnPFile -SourceUrl $file.fieldvalues.FileRef -TargetUrl $destpath -OverwriteIfAlreadyExists -Force -ErrorVariable errors -ErrorAction SilentlyContinue
        $fileerrors += $errors
    }
    
    #Defines details for the logfile containing files that failed to copy
    $user = $email.Split("@")[0]
    $path = ("c:\temp\" + $user + "_fileerrors.txt")
    $fileerrors | Out-File -append -literalpath $path
      
    # Remove Global Admin from Site Collection Admin role for both users
    Write-Host "`nRemoving $globaladmin from OneDrive site collections" -ForegroundColor Blue
    Set-SPOUser -Site $departingOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $false
    Set-SPOUser -Site $destinationOneDriveSite -LoginName $globaladmin -IsSiteCollectionAdmin $false
    Write-Host "`Copy Complete!" -ForegroundColor Green
    
    #Defining variables used in the prompt
    $title    = 'Offboarding'
    $question3 = "Would you like to export the user's email as well?"
    $choices3  = '&Yes','&No'
    
    #Creates a UI prompt with 2 options for the user to select
    $decision = $Host.UI.PromptForChoice($title, $question3, $choices3, 1)
    
    Switch ($decision)
    {
    {$_ -match '0'} { Write-host "Starting script"}
    {$_ -match '1'} { Write-host "Hit CTRL+C to stop the script" | sleep 3000}
    }
    
        #Seperates username from email address
        $user = $email.Split("@")[0]
        
            #Starts a seperate PS sessions with and imports O365
            Write-Host "`nStarting O365 Session" -ForegroundColor Blue
        Try{Get-O365Mailbox aguerot -ErrorAction Stop > $Null}
        Catch{
        $O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365Cred -Authentication Basic -AllowRedirection
        Import-PSSession $O365Session –Prefix o365
        }
    
        #Logs into MS Online
        Write-Host "`nLogging into MS Online" -ForegroundColor Blue
        Try{Get-MsolUser -UserPrincipalName $user -ErrorAction Stop > $Null}
        Catch{Connect-MsolService -Credential $O365Cred}
        
        Write-Host "`nStarting Compliance session" -ForegroundColor Blue
        
        Try{Get-ComplianceSearch > $Null}
        Catch{
        #Get login credentials
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $O365Cred -Authentication Basic -AllowRedirection
        Import-PSSession $Session -AllowClobber -DisableNameChecking
        }
        Write-Host "`nConfiguring logfile" -ForegroundColor Blue
        
        $Logfile = "C:\log\$user.txt"
        #Function to setup a logfile to write
        Function LogWrite
        {
        Param ([string]$logstring)
        Write-Host $LogString
        Add-content $Logfile -value $logstring
        }
        $Date = Get-Date -Format "MM/dd/yyyy"
        LogWrite $Date
        LogWrite " "
        LogWrite "Username: $user"
        
        #Begin Compliance Search
        $UPN = Get-ADUser $user | ForEach-Object{$_.UserPrincipalName}
        $SearchName = $User + "_Search"
        New-ComplianceSearch -Name $SearchName -ExchangeLocation $UPN
        Start-ComplianceSearch $SearchName
        Logwrite " "
        LogWrite "Compliance search $SearchName started"
        #Updates the user on status of the compliance search
        Do{
        $complianceSearch = Get-ComplianceSearch $SearchName | ForEach-Object{$_.Status}
        Write-Host "`nCompliance search in progress" -ForegroundColor Blue
        Start-Sleep -s 30
        }
        While ($complianceSearch -ne 'Completed')
        $Size = Get-ComplianceSearch $SearchName | ForEach-Object{$_.Size}
        $Size = $Size / 1048576 | Out-String
        $Size = $Size.SubString(0,6)
        
        Write-Host "`nFormatting Search for Export" -ForegroundColor Blue
        
        # Begins an export of the search
        New-ComplianceSearchAction -SearchName $SearchName -EnableDedupe $true -Export -Format FxStream -ArchiveFormat PerUserPST > $Null
        
        #Will give you updates for exports and what the status is (When testing this appears to be broken)
        Write-Host "`nExporting .pst" -ForegroundColor Blue
        
        #Wait for Export to complete
        $ExportName = $SearchName + "_Export"
        Start-Sleep -s 20
        do{
        $SearchAction = Get-ComplianceSearchAction -Identity $ExportName | Select-Object Status,JobProgress
        $Status = $SearchAction.Status
        $ExportProgress = $SearchAction.JobProgress
        Write-Host "Export in progress, $ExportProgress complete"
        If($Status -ne "Completed"){
        Start-Sleep -s 60
        }
        }
        while ($Status -ne 'Completed')
        #Writes to log giving you proof the export was completed
        LogWrite "Compliance search completed"
        #Script is complete
        Write-Host "`nCompliance Search Completed and Export Started" -ForegroundColor Blue