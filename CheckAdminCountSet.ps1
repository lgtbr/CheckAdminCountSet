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
# 0 - PREREQUISITES                                                                           #
#                                                                                             #
###############################################################################################

# Location where the .ps1 file is located
$scriptDirectory = "C:\Scripts"
$adminListFilename = "actualAdmin"
$tmpUpdatedAdminListFilename = "tmpUpdatedAdminList"

$logsFolderName = "adminsActivitiesLogs"
$logsFilename = "history"
$logsRetentionDays = 30

$domainAdminGroupName = "Admins du domaine"



$enableLogs = $true



###############################################################################################
#                                                                                             #
# 0 - PARAMETERS                                                                              #
#                                                                                             #
###############################################################################################

param ([int] $logsRetentionDays,[string] $domainAdminGroupName, [boolean] $enableLogs, [string] $mailSMTPMethod)

###############################################################################################
#                                                                                             #
# 1 - FUNCTIONS                                                                               #
#                                                                                             #
###############################################################################################

function Get-AdminAccounts {
    Get-ADGroupMember -Identity $domainAdminGroupName -Recursive | ForEach-Object {  
        Get-ADUser -Identity $_.SamAccountName -Properties adminCount | Where-Object {$_.adminCount -gt 0} | Select-Object -Property Name, SamAccountName, adminCount, DistinguishedName
    }
}

$mailParameters = Get-Content "C:\Scripting\network_equipment.json" | ConvertFrom-Json

function Send-MailSMTPCredentials {
    # Get the credential
    $password = ConvertTo-SecureString "y'2hfb6h*a" -AsPlainText -Force
    $Cred = New-Object System.Management.Automation.PSCredential ("qimapp@qiminfo.ch", $password)

    ## Define the Send-MailMessage parameters
    $mailParams = @{
        SmtpServer                 = 'smtp.office365.com'
        Port                       = '587' # or '25' if not using TLS
        UseSSL                     = $true ## or not if using non-TLS
        Credential                 = $Cred     
        From                       = 'qimapp@qiminfo.ch'
        To                         = 'louis.gattabrusi@qiminfo.ch'
        Subject                    = "SMTP Client Submission - $(Get-Date -Format g)"
        Body                       = 'This is a test email using SMTP Client Submission'
        DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
    }
}

function Send-DirectMailOffice365 {
    ## Send the message
    Send-MailMessage @mailParams

    ## Build parameters
    $mailParams = @{
        SmtpServer                 = 'qiminfo.mail.protection.outlook.com'
        Port                       = '25'
        UseSSL                     = $true   
        From                       = 'qimapp@qiminfo.ch'
        To                         = 'louis.gattabrusi@qiminfo.ch'
        Subject                    = "Direct Send $(Get-Date -Format g)"
        Body                       = 'This is a test email using SMTP Client Submission'
        DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
    }

    ## Send the email
    Send-MailMessage @mailParams
}

function Send-MailMicrosoftGraph {
    $clientID = "b0829ad3-a5aa-40d8-9a0e-673761283fc6"
    $Clientsecret = "0Mm8Q~TCKWvySzdCDuSCSUnc0cx49PkYVbrmfdeG"
    $tenantID = "38d9b3f4-6481-4de8-9988-9e284f15e845"

    $Clientsecret = ConvertTo-SecureString -String '0Mm8Q~TCKWvySzdCDuSCSUnc0cx49PkYVbrmfdeG' -AsPlainText -Force | ConvertFrom-SecureString | Set-Clipboard
    $Credential = ConvertTo-GraphCredential -ClientID $clientID -ClientSecretEncrypted '01000000d08c9ddf0115d1118c7a00c04fc297eb01000000ca9cbe7796086c42905c6357165266160000000002000000000003660000c0000000100000003e23b0ecb94103cef4473b93473425290000000004800000a00000001000000092b16e5ff30e61b54de506e1bcdc3bff580000008b915aac2f9c86a146a4d8d7c6ba0b53daeff525366b9aa19798a566faa3a8e41eb6223a1ee02f8d1b34b65a5515ed3c6b6c5ec5f42ab98625bcb4863e452d9d02dd2ff831993dcded90a9f0258846ff80a0df7b9f89daab14000000606bdfe90e3e8d265fe2dd0f516aa717d5bc0273' -DirectoryID $tenantID


    try {
        $sendEmailMessageSplat = @{
            From                   = 'qimapp@qiminfo.ch'
            To                     = 'louis.gattabrusi@qiminfo.ch'
            Credential             = $Credential
            HTML                   = $x
            Subject                = "WARNING - Membership change(s) on group $domainAdminGroupName"
            Graph                  = $true
            Verbose                = $true
            RequestReadReceipt     = $false
            RequestDeliveryReceipt = $false
            DoNotSaveToSentItems   = $true
        }
        Send-EmailMessage @sendEmailMessageSplat
    } catch {
        Write-HOst "erroe"
    }
}

###############################################################################################
#                                                                                             #
# 2 - CHECK IF EXIST OR CREATE A LIST OF THE ACTUAL USER WITH ADMINCOUNTSET SET TO 1          #
#                                                                                             #
###############################################################################################

Set-Location $scriptDirectory

if ((Test-Path "$scriptDirectory\$adminListFilename.csv") -and -not (((Get-Content "$scriptDirectory\$adminListFilename.csv") -eq $null) -eq $true)) {
    Write-Host "OK - File $adminListFilename.csv exist and is not empty" -ForegroundColor green
} else {
    Write-Host "NOK - File $adminListFilename.csv is missing or empty !" -ForegroundColor red
    Get-AdminAccounts | Export-Csv -Path "$scriptDirectory\$adminListFilename.csv" -Encoding UTF8 -NoTypeInformation
    Write-Host "OK - File $adminListFilename.csv has been created / refilled" -ForegroundColor green
}

if ($enableLogs -eq $true) {
    if ((Test-Path "$scriptDirectory\$logsFolderName") -eq $false) {
        Write-Host "NOK - Folder $logsFolderName does not exist !" -ForegroundColor red
        New-Item "$scriptDirectory\$logsFolderName" -ItemType Directory | Out-Null
        Write-Host "OK - Folder $logsFolderName has been created" -ForegroundColor green
    } else {
        Write-Host "OK - Folder $logsFolderName exist" -ForegroundColor green
    }

    if ((Test-Path "$scriptDirectory\$logsFolderName\$logsFilename.csv") -eq $false) {
        Write-Host "NOK - File $logsFilename.csv is missing" -ForegroundColor red
        New-Item "$scriptDirectory\$logsFolderName\$logsFilename.csv" -ItemType File | Out-Null
        Write-Host "OK - File $logsFilename.csv has been created" -ForegroundColor green
    } else {
        Write-Host "OK - File $logsFolderName.csv exist" -ForegroundColor green
    }
}

###############################################################################################
#                                                                                             #
# 3 - CREATE TEMPORARY FILE WITH ALL USERS ADMINCOUNTSET SET TO 1 AND COMPARE WITH ACTUAL LIST#
#                                                                                             #
###############################################################################################

Get-AdminAccounts | Export-Csv -Path "$scriptDirectory\$tmpUpdatedAdminListFilename.csv" -Encoding UTF8 -NoTypeInformation

$actualAdminList = Import-Csv "$scriptDirectory\$adminListFilename.csv"
$updatedAdminList = Import-Csv "$scriptDirectory\$tmpUpdatedAdminListFilename.csv"

$actualDate = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
$changeCounter = 0

if (Compare-Object -ReferenceObject $updatedAdminList -DifferenceObject $actualAdminList -Property DistinguishedName, Name, SamAccountName, SideIndicator, DistinguishedName | Select-Object DistinguishedName, Name, SamAccountName, SideIndicator | `
    ForEach-Object {
        if ($_.SideIndicator -eq "=>") {
            $changeCounter = ++$changeCounter 
            Write-Host "[-] User" $_.Name "("$_.SamAccountName") has been removed from $domainAdminGroupName group" -ForegroundColor red
            Write-Host "[-] The attribute adminCount of the account" $_.Name "("$_.SamAccountName")" "has been reset to 0" -ForegroundColor yellow
            $sideIndicatorTranslaste = "Removed"
            #Set-ADUser -Identity $_.SamAccountName -Replace @{adminCount=0}
        } elseif ($_.SideIndicator -eq "<=") {
            $changeCounter = ++$changeCounter 
            Write-Host "[+] User" $_.Name $_.SamAccountName "has been added in $domainAdminGroupName group" -ForegroundColor green
            $sideIndicatorTranslaste = "Added"
        } if ($enableLogs -eq $true) { 
            $logChange = New-Object PSObject
            $logChange | Add-Member Noteproperty -Name Date -value $actualDate
            $logChange | Add-Member Noteproperty -Name Name -value $_.Name
            $logChange | Add-Member Noteproperty -Name SamAccountName -value $_.SamAccountName
            $logChange | Add-Member Noteproperty -Name State -value $sideIndicatorTranslaste
            $logChange | Add-Member Noteproperty -Name DistinguishedName -value $_.DistinguishedName
            $logChange | Sort-Object -Property Date -Descending | Export-Csv "$scriptDirectory\$logsFolderName\$logsFilename.csv" -Append -Encoding UTF8 -NoTypeInformation
        } 
    } ) {
    #Set-Content -Path "$scriptDirectory\$adminListFilename.csv" -Value (Get-Content "$scriptDirectory\$tmpUpdatedAdminListFilename.csv")
    #Remove-Item -Path "$scriptDirectory\$tmpUpdatedAdminListFilename.csv"
} elseif (-not (diff $updatedAdminList $actualAdminList)) {
    Write-Host "OK - No changes detected, the file $tmpUpdatedAdminListFilename.csv will be deleted..." -ForegroundColor green
    #Remove-Item -Path "$scriptDirectory\$tmpUpdatedAdminListFilename.csv"
}

$historyFile = Import-CSV "$scriptDirectory\$logsFolderName\$logsFilename.csv"
$lastChanges = $historyFile | Sort-Object -Property Date -Descending | Select-Object -First $changeCounter
$changeHistory = $historyFile | Where { [datetime]::ParseExact($_.Date, "MM/dd/yyyy HH:mm:ss", $null) -gt (Get-Date).date.adddays(-$logsRetentionDays)} | Select-Object -Skip $changeCounter | Sort-Object -Property Date -Descending

$membershipChangesContent = ForEach ($users in $lastChanges) {
    if ($users.State -eq "Added") {
        $users.State = "<font color='28a745'>" + $users.State + "</font>"
    } elseif ($users.State -eq "Removed") {
        $users.State = "<font color='dc3545'>" + $users.State + "</font>"
    }
    "<tr class='text-center'>"
    "<td class='text-center'>" + $users.Date + "</td>"
    "<td>" + $users.Name + "</td>"
    "<td>" + $users.SamAccountName + "</td>"
    "<td>" + $users.State + "</td>"
    "<td>" + $users.DistinguishedName + "</td>"
    "</tr>"
}

$historyContent = ForEach ($changes in $changeHistory) {
    if ($changes.State -eq "Added") {
        $changes.State = "<font color='28a745'>" + $changes.State + "</font>"
    } elseif ($changes.State -eq "Removed") {
        $changes.State = "<font color='dc3545'>" + $changes.State + "</font>"
    }
    "<tr class='text-center'>"
    "<td class='text-center'>" + $changes.Date + "</td>"
    "<td>" + $changes.Name + "</td>"
    "<td>" + $changes.SamAccountName + "</td>"
    "<td class='font-weight-bold'>" + $changes.State + "</td>"
    "<td>" + $changes.DistinguishedName + "</td>"
    "</tr>"
}



###############################################################################################
#                                                                                             #
# 4 - HTML body                                                                               #
#                                                                                             #
###############################################################################################

$x = $(Switch -RegEx (Get-Content "C:\Scripts\11\index.html"){
    "membershipChangesContent" {$_ -replace "membershipChangesContent", "$membershipChangesContent"; Continue}
    "daysHistoryContent" {$_ -replace "daysHistoryContent", "$logsRetentionDays"; Continue}
    "groupNameContent" {$_ -replace "groupNameContent", "$domainAdminGroupName"; Continue}
    "historyContent" {$_ -replace "historyContent", "$historyContent"; Continue}
    "actualDateContent" {$_ -replace "actualDateContent", "$actualDate"; Continue}
    default {$_}
})
