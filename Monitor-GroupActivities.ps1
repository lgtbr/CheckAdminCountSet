###############################################################################################
#                                                                                             #
# DESCRIPTION                                                                                 #
#                                                                                             #
# Ce script permet la sauvegarde de la configration d'un FortiSwitch (ou d'un autre           #
# équipement permettant la connexion en SSH puis l'export via TFTP).                          #
#                                                                                             #
# Il fonctionne uniquement pour un équipement. Pour l'utiliser sur plusieurs équipements, il  #
# faut faire une copie de ce fichier puis le reconfigurer pour l'équipement en question.      #
#                                                                                             #
###############################################################################################

###############################################################################################
#                                                                                             #
# 0 - FUNCTIONS                                                                               #
#                                                                                             #
###############################################################################################

function Get-AdminAccounts {
    Get-ADGroupMember -Identity $domainAdminGroupName -Recursive | Select-Object Name, SamAccountName, DistinguishedName
}

function Log-NewChange {
    $logChange = [PSCustomObject]@{
        Date = (Get-Date -Format "MM/dd/yyyy HH:mm:ss")
        Name = $_.Name
        SamAccountName = $_.SamAccountName
        State = $sideIndicatorTranslaste
        DistinguishedName = $_.DistinguishedName
    }

    $logChange | Sort-Object -Property Date -Descending | Export-Csv "$scriptDirectory\$logsFoldername\$logsFilename" -Append -Encoding UTF8 -NoTypeInformation
    $updatedAdminList | Export-Csv -Path "$scriptDirectory\$adminListFoldername\$adminListFilename" -Encoding UTF8 -NoTypeInformation
}

function Check-InitialSetup {
    if (Get-Module -Name ActiveDirectory) {
        Import-Module -Name ActiveDirectory
        Write-Host "OK - Module ActiveDirectory has been imported" -ForegroundColor green
    } else {
        Write-Host "NOK - Module ActiveDirectory not detected" -ForegroundColor red
        Install-WindowsFeature -Name "RSAT-AD-PowerShell" -IncludeAllSubFeature
        Import-Module -Name ActiveDirectory
        Write-Host "OK - Module ActiveDirectory has been installed and imported" -ForegroundColor green
    }

    foreach ($folder in $foldersList) {
        if ((Test-Path $folder) -eq $false){
            Write-Host "NOK - Folder $folder does not exist in $scriptDirectory !" -ForegroundColor red
            New-Item "$folder" -ItemType Directory | Out-Null
        } 
        Write-Host "OK - Folder $folder has been created in $scriptDirectory" -ForegroundColor green
    }

    if ((Test-Path -Path $scriptDirectory\$configFoldername\$configFilename) -eq $false ) {
        Write-Host "NOK - File $configFilename is missing in $scriptDirectory\$configFoldername\$configFile" -ForegroundColor red
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/lgtbr/MonitorAdminCountSet/main/parameters.json" -OutFile "$scriptDirectory\$configFoldername\$configFilename"
        Write-Host "OK - File $configFilename has been downloaded" -ForegroundColor green
    }

    $parametersJSON = Get-Content -Path $scriptDirectory\$configFoldername\$configFilename | Out-String | ConvertFrom-Json

    if ($parametersJSON.initialized -eq 0) {
        $domainAdminGroupName = Read-Host -Prompt "Name of your Donain Administrators group?"
        $logsRetentionDays = Read-Host -Prompt "Logs retention in days? (30 is adviced)"

        do { 
            $sendingMailMethod = Read-Host -Prompt "Please refer to the GitHub page for more infornations about sending methods`nChoose your mail reporting method : (MailSMTPCredentials | DirectMailOffice365 | MailMicrosoftGraph)" 
        } until ('MailSMTPCredentials','DirectMailOffice365','MailMicrosoftGraph' -ccontains $sendingMailMethod)

        switch ($sendingMailMethod) {
            'MailSMTPCredentials' { Setup-MailSMTPCredentials } 
            'DirectMailOffice365' { Setup-DirectMailOffice365 } 
            'MailMicrosoftGraph' { Setup-MailMicrosoftGraph } 
        }

        $parametersJSON.domainAdminGroupName = $domainAdminGroupName;
        $parametersJSON.logsRetentionDays = $logsRetentionDays;
        $parametersJSON.initialized = 1;

        foreach($method in @('MailSMTPCredentials', 'DirectMailOffice365', 'MailMicrosoftGraph')){   # loop through all methods and set the enable flag to true for the selected one and false for all others  
            if($method -eq $sendingMailMethod){   # check if the current method is the selected one  
                $parametersJSON.$method[0].enable = [boolean]$True;   # set enable flag to true  
            } else{   # current method is not selected one  
                $parametersJSON.$method[0].enable = [boolean]$False;   # set enable flag to false  
            }   # end of if statement  

            $parametersJSON | ConvertTo-Json | Out-File $scriptDirectory\$configFoldername\$configFilename;     # write changes to config file after each iteration of loop    													     
            # end of foreach loop    	     
        }
    } else {
        Write-Host "OK - File $configFilename is already setup" -ForegroundColor green
    } 

    foreach ($file in $filesList) {
        if ((Test-Path $file) -eq $false) {
            if ($file.Contains($adminListFilename)) {
                Write-Host "NOK - File $adminListFilename is missing or empty !" -ForegroundColor red
                Get-AdminAccounts | Export-Csv -Path "$scriptDirectory\$adminListFoldername\$adminListFilename" -Encoding UTF8 -NoTypeInformation
                Write-Host "OK - File $adminListFilename has been created / refilled" -ForegroundColor green
            } elseif ($file.Contains($configFilename)) {
                Write-Host "NOK - File $configFilename is missing in $scriptDirectory\$configFoldername\$configFile" -ForegroundColor red
                Invoke-WebRequest -Uri "https://raw.githubusercontent.com/lgtbr/MonitorAdminCountSet/main/parameters.json" -OutFile "$scriptDirectory\$configFoldername\$configFilename"
                Write-Host "OK - File $configFilename has been downloaded" -ForegroundColor green
            } elseif ($file.Contains($htmlTemplateFilename)) {
                Write-Host "NOK - File $configFile is missing in $scriptDirectory\$htmlTemplateFoldername\$htmlTemplateFilename" -ForegroundColor red
                Invoke-WebRequest -Uri "https://raw.githubusercontent.com/lgtbr/MonitorAdminCountSet/main/htmlTemplate/model.html" -OutFile "$scriptDirectory\$htmlTemplateFoldername\$htmlTemplateFilename"
                Write-Host "OK - File $configFile has been downloaded" -ForegroundColor green 
            } elseif ($file.Contains($logsFilename)) { 
                Write-Host "NOK - File $logsFilename is missing" -ForegroundColor red 
                New-Item "$scriptDirectory\$logsFoldername\$logsFilename" –ItemType File | Out-Null 
                Write-Host "OK – File $logsFilename has been created” –ForegroundColor green 
            }  
        } else {   
            Write-Host “OK – File $file exist” –ForegroundColor green
        }  
    }	        
}

function Setup-MailSMTPCredentials {
    $parametersJSON.MailSMTPCredentials[0].authUsername = Read-Host -Prompt "Mail address used to authenticate to the SMTP server"
    $parametersJSON.MailSMTPCredentials[0].authPassword = Read-Host -Prompt "Password used to authenticate to the SMTP server" -AsSecureString | ConvertFrom-SecureString 
    $parametersJSON.MailSMTPCredentials[0].smtpServer  = Read-Host -Prompt "Name of the SMTP server" 
    $parametersJSON.MailSMTPCredentials[0].port        = Read-Host -Prompt "Port of the SMTP server"
    $msgBoxInput = [System.Windows.MessageBox]::Show('Does your SMTP server use SSL?','SMTP SSL','YesNo','Question')
    switch ($msgBoxInput) {
        'Yes' {$parametersJSON.MailSMTPCredentials[0].useSSL = [boolean]$true} 
        'No' {$parametersJSON.MailSMTPCredentials[0].useSSL = [boolean]$false} 
    }
    $parametersJSON.MailSMTPCredentials[0].from        = Read-Host -Prompt "Sender’s email address"
    $parametersJSON.MailSMTPCredentials[0].to          = Read-Host -Prompt "Email address of a recipient or recipients separated by , value" 

    $parametersJSON | ConvertTo-Json | Out-File $scriptDirectory\$configFoldername\$configFilename
}

function Send-MailSMTPCredentials {
    $mailCredentials = New-Object System.Management.Automation.PSCredential -argumentlist $parametersJSON.MailSMTPCredentials[0].authUsername, (ConvertTo-SecureString $parametersJSON.MailSMTPCredentials[0].authPassword -Force)
    $mailParams = @{
        SmtpServer                 = $mailSMTPCredentialsSmtpServer 
        Port                       = $mailSMTPCredentialsPort
        UseSSL                     = $mailSMTPCredentialsUseSSL
        Credential                 = $mailCredentials     
        From                       = $mailSMTPCredentialsFrom
        To                         = $mailSMTPCredentialsTo 
        Subject                    = $mailSubject
        Body                       = [string]$modifiedHTMLContent
        BodyAsHtml                 = $true
    }
    Send-MailMessage @mailParams
}

function Setup-DirectMailOffice365 {
    $parametersJSON.DirectMailOffice365[0].smtpServer = Read-Host -Prompt "Name of the SMTP server" 
    $parametersJSON.DirectMailOffice365[0].port       = Read-Host -Prompt "Port of the SMTP server"
    $msgBoxInput = [System.Windows.MessageBox]::Show('Does your SMTP server use SSL?','SMTP SSL','YesNo','Question')
    switch ($msgBoxInput) {
        'Yes' {$parametersJSON.DirectMailOffice365[0].useSSL = [boolean]$true}
        'No' {$parametersJSON.DirectMailOffice365[0].useSSL = [boolean]$false}
    }
    $parametersJSON.DirectMailOffice365[0].from       = Read-Host -Prompt "Sender’s email address"
    $parametersJSON.DirectMailOffice365[0].to         = Read-Host -Prompt "Email address of a recipient or recipients separated by , value"

    $parametersJSON | ConvertTo-Json | Out-File $scriptDirectory\$configFoldername\$configFilename
}

function Send-DirectMailOffice365 {
    $mailParams = @{
        SmtpServer                 = $directMailOffice365SmtpServer
        Port                       = $directMailOffice365Port
        UseSSL                     = $directMailOffice365UseSSL
        From                       = $directMailOffice365From
        To                         = $directMailOffice365TO
        Subject                    = $mailSubject
        Body                       = [string]$modifiedHTMLContent
        BodyAsHtml                 = $true
    }
    Send-MailMessage @mailParams
}

function Setup-MailMicrosoftGraph {
    $parametersJSON.MailMicrosoftGraph[0].clientID = Read-Host -Prompt "Application (client) ID of the Azure App registrations" 
    $parametersJSON.MailMicrosoftGraph[0].clientsecret = Read-Host -Prompt "App registration client secret values " -AsSecureString | ConvertFrom-SecureString
    $parametersJSON.MailMicrosoftGraph[0].tenantID = Read-Host -Prompt "Azure Tenant ID"
    $parametersJSON.MailMicrosoftGraph[0].from = Read-Host -Prompt "Sender’s email address"
    $parametersJSON.MailMicrosoftGraph[0].to = Read-Host -Prompt "Email address of a recipient or recipients separated by , value"
    $parametersJSON | ConvertTo-Json | Out-File $scriptDirectory\$configFoldername\$configFilename  
}  

function Send-MailMicrosoftGraph {
    if (-not (Get-Module -ListAvailable -Name "Mailozaurr"))  {
        Write-Host "NOK - Module Mailozaurr not installed" -ForegroundColor red
        Install-Module Mailozaurr -Force
        Write-Host "OK - Module Mailozaurr has been installed" -ForegroundColor green
    } else {
        Write-Host "OK - Module Mailozaurr has been detected" -ForegroundColor green
    }
    $sendEmailMessageSplat = @{
        From                   = $mailMicrosoftGraphFrom
        To                     = $mailMicrosoftGraphTo
        Credential             = ConvertTo-GraphCredential -ClientID $mailMicrosoftGraphClientID -ClientSecretEncrypted $mailMicrosoftGraphClientsecret -DirectoryID $mailMicrosoftGraphTenantID
        HTML                   = $modifiedHTMLContent
        Subject                = $mailSubject
        Graph                  = $true
        Verbose                = $false
        RequestReadReceipt     = $false
        RequestDeliveryReceipt = $false
        DoNotSaveToSentItems   = $true
    }

    Send-EmailMessage @sendEmailMessageSplat 
} 

###############################################################################################
#                                                                                             #
# 0 - VARIABLES                                                                               #
#                                                                                             #
###############################################################################################

# Set the execution path where the .ps1 is located
$scriptDirectory = Get-Location
Set-Location $scriptDirectory

# Variables needed for the script execution
$adminListFoldername = "adminList"
$adminListFilename = "actualAdmin.csv"
$tmpUpdatedAdminListFilename = "tmpUpdatedAdminList.csv"

$configFoldername = "config"
$configFilename = "parameters.json"

$htmlTemplateFoldername = "htmlTemplate"
$htmlTemplateFilename = "model.html"

$logsFoldername = "logs"
$logsFilename = "logsHistory.csv"

# Create an array of folders and files
$foldersList = @("$scriptDirectory\$adminListFoldername","$scriptDirectory\$configFoldername","$scriptDirectory\$htmlTemplateFoldername","$scriptDirectory\$logsFoldername")
$filesList = @("$scriptDirectory\$adminListFoldername\$adminListFilename","$scriptDirectory\$configFoldername\$configFilename","$scriptDirectory\$htmlTemplateFoldername\$htmlTemplateFilename","$scriptDirectory\$logsFoldername\$logsFilename")

# Subject of mail to be sent in case of membership change(s) on group 
$mailSubject = "WARNING - Membership change(s) on group $domainAdminGroupName"

# Call Check-InitialSetup function
Check-InitialSetup

$parametersJSON = Get-Content -Path $scriptDirectory\$configFoldername\$configFilename | Out-String | ConvertFrom-Json

$domainAdminGroupName            = $parametersJSON.domainAdminGroupName
$logsRetentionDays               = $parametersJSON.logsRetentionDays
            
$mailSMTPCredentialsEnable       = $parametersJSON.MailSMTPCredentials.enable
$mailSMTPCredentialsAuthUsername = $parametersJSON.MailSMTPCredentials.authUsername
$mailSMTPCredentialsAuthPassword = $parametersJSON.MailSMTPCredentials.authPassword
$mailSMTPCredentialsSmtpServer   = $parametersJSON.MailSMTPCredentials.smtpServer
$mailSMTPCredentialsPort         = $parametersJSON.MailSMTPCredentials.port
$mailSMTPCredentialsUseSSL       = $parametersJSON.MailSMTPCredentials.useSSL
$mailSMTPCredentialsFrom         = $parametersJSON.MailSMTPCredentials.from
$mailSMTPCredentialsTo           = $parametersJSON.MailSMTPCredentials.to

$directMailOffice365Enable       = $parametersJSON.DirectMailOffice365.enable
$directMailOffice365SmtpServer   = $parametersJSON.DirectMailOffice365.smtpServer
$directMailOffice365Port         = $parametersJSON.DirectMailOffice365.port
$directMailOffice365UseSSL       = $parametersJSON.DirectMailOffice365.useSSL
$directMailOffice365From         = $parametersJSON.DirectMailOffice365.from
$directMailOffice365TO           = $parametersJSON.DirectMailOffice365.to

$mailMicrosoftGraphEnable        = $parametersJSON.MailMicrosoftGraph.enable
$mailMicrosoftGraphClientID      = $parametersJSON.MailMicrosoftGraph.clientID
$mailMicrosoftGraphClientsecret  = $parametersJSON.MailMicrosoftGraph.clientsecret
$mailMicrosoftGraphTenantID      = $parametersJSON.MailMicrosoftGraph.tenantID
$mailMicrosoftGraphFrom          = $parametersJSON.MailMicrosoftGraph.from
$mailMicrosoftGraphTo            = $parametersJSON.MailMicrosoftGraph.to

###############################################################################################
#                                                                                             #
# 3 - CREATE TEMPORARY FILE WITH ALL USERS ADMINCOUNTSET SET TO 1 AND COMPARE WITH ACTUAL LIST#
#                                                                                             #
###############################################################################################

$actualAdminList = Import-Csv "$scriptDirectory\$adminListFoldername\$adminListFilename"
$updatedAdminList = Get-AdminAccounts

$actualDate = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
$changeCounter = 0

if (-not (diff $actualAdminList $updatedAdminList)) {
    Write-Host "OK - No changes detected" -ForegroundColor green
} elseif (Compare-Object -ReferenceObject $updatedAdminList -DifferenceObject $actualAdminList -Property DistinguishedName, Name, SamAccountName, SideIndicator, DistinguishedName | Select-Object DistinguishedName, Name, SamAccountName, SideIndicator | `
    ForEach-Object {
        if ($_.SideIndicator -eq "=>") {
            $changeCounter = ++$changeCounter 
            Write-Host "[-] User" $_.Name "("$_.SamAccountName") has been removed from $domainAdminGroupName group" -ForegroundColor red
            Write-Host "[-] The attribute adminCount of the account" $_.Name "("$_.SamAccountName")" "has been reset to 0" -ForegroundColor yellow
            $sideIndicatorTranslaste = "Removed"
            Log-NewChange
            Set-ADUser -Identity $_.SamAccountName -Replace @{adminCount=0}  #added this line to set the adminCount attribute to 0 for removed users 
        } elseif ($_.SideIndicator -eq "<=") {  #changed this line to include setting the adminCount attribute to 1 for added users 
            $changeCounter = ++$changeCounter  #added this line to increment change counter before logging new change 
            Write-Host "[+] User" $_.Name $_.SamAccountName "has been added in $domainAdminGroupName group" -ForegroundColor green
            Set-ADUser -Identity $_.SamAccountName -Replace @{adminCount=1} #added this line to set the adminCount attribute to 1 for added users  
            $sideIndicatorTranslaste = "Added"
            Log-NewChange #moved this line after setting admin count attribute  
        }   #added closing bracket here
    }) {  #changed this bracket from closing ForEach loop to closing If statement  
    
 }   #added closing bracket here 

$historyFile = Import-CSV "$scriptDirectory\$logsFoldername\$logsFilename"
$lastChanges = $historyFile | Sort-Object -Property Date -Descending | Select-Object -First $changeCounter
$changeHistory = $historyFile | Where { [datetime]::ParseExact($_.Date, "MM/dd/yyyy HH:mm:ss", $null) -gt (Get-Date).date.adddays(-$logsRetentionDays)} | Select-Object -SkipLast $changeCounter | Sort-Object -Property Date -Descending

###############################################################################################
#                                                                                             #
# 4 - HTML body                                                                               #
#                                                                                             #
###############################################################################################

$membershipChangesContent = ForEach ($changes in $lastChanges) {
    switch -RegEx ($changes.State) {
        "Added" {$changes.State = "<font color='28a745'>" + $changes.State + "</font>"} 
        "Removed" {$changes.State = "<font color='dc3545'>" + $changes.State + "</font>"} 
    }
    "<tr class='text-center'>"
    "<td class='text-center'>" + $changes.Date + "</td>"
    "<td>" + $changes.Name + "</td>"
    "<td>" + $changes.SamAccountName + "</td>"
    "<td>" + $changes.State + "</td>"
    "<td>" + $changes.DistinguishedName + "</td>"
    "</tr>"
} 

$historyContent = ForEach ($changes in $changeHistory) {
    switch ($changes.State) {
        "Added" {$changes.State = "<font color='28a745'>" + $changes.State + "</font>"} 
        "Removed" {$changes.State = "<font color='dc3545'>" + $changes.State + "</font>"} 
    }

    "<tr class='text-center'>"
    "<td class='text-center'>" + $changes.Date + "</td>"
    "<td>" + $changes.Name + "</td>"
    "<td>" + $changes.SamAccountName + "</td>"
    "<td class='font-weight-bold'>" + $changes.State + "</td>"
    "<td>" + $changes.DistinguishedName + "</td>"
    "</tr>"
} 

$modifiedHTMLContent = $(Switch -RegEx (Get-Content "$scriptDirectory\$htmlTemplateFoldername\$htmlTemplateFilename"){
    "membershipChangesContent" {$_ -replace "membershipChangesContent", "$membershipChangesContent"; Continue}
    "daysHistoryContent" {$_ -replace "daysHistoryContent", "$logsRetentionDays"; Continue}
    "groupNameContent" {$_ -replace "groupNameContent", "$domainAdminGroupName"; Continue}
    "historyContent" {$_ -replace "historyContent", "$historyContent"; Continue}
    "actualDateContent" {$_ -replace "actualDateContent", "$actualDate"; Continue}
    default {$_}
}) 

###############################################################################################
#                                                                                             #
# 4 - Sending mail                                                                            #
#                                                                                             #
###############################################################################################

if ($changeCounter -ne 0) {
    if ($mailSMTPCredentialsEnable) {
        Send-MailSMTPCredentials
    } elseif ($directMailOffice365Enable) {
        Send-DirectMailOffice365
    } elseif ($mailMicrosoftGraphEnable) {
        Send-MailMicrosoftGraph
    }
}
