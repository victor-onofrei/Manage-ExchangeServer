$ts = Get-Date -Format yyyyMMdd_hhmmss
$Path="\\path\outputs"
$ProjName = "mailboxes_list"
$FileName = "Mailboxes_list_complex_onprem.full.$ts.xls"
New-Item -Name $ProjName -Path $Path -Type Directory -ErrorAction SilentlyContinue
New-Item -Name $FileName -Path $Path\$ProjName -Type File -ErrorAction SilentlyContinue
$PathtoAddressesOutfile = "$Path\$ProjName\$FileName"

Start-Transcript "$Path\$ProjName\$FileName.txt"

# Adresses of the users in scope
# On premise: Which mailbox type should be searched for? Possible Values: "DiscoveryMailbox, EquipmentMailbox, GroupMailbox, LegacyMailbox, LinkedMailbox, LinkedRoomMailbox, RoomMailbox, SchedulingMailbox, SharedMailbox, TeamMailbox, UserMailbox"
$RecipientType = "UserMailbox" 

# Import AD module
Import-Module Activedirectory

# Get the Dataset of Mailboxes and RemoteMailboxes to work with
Write-Host "Getting OnPremise $RecipientType"
$Timer = [System.diagnostics.stopwatch]::startNew()
$cp = $Timer.ElapsedMilliseconds
$OnPrem_Mailboxes = Get-Recipient -RecipientType $RecipientType -ResultSize Unlimited -ErrorAction SilentlyContinue
Write-Host "Found" @($OnPrem_Mailboxes).Count "$RecipientType Mailboxes"
Write-Host "`t`tOnPrem_Mailboxes: $($Timer.ElapsedMilliseconds - $cp) milliseconds"

Write-Host "Getting OnPremise TotalMailboxSizes"
$cp = $Timer.ElapsedMilliseconds
$OnPrem_Mailboxes_TotalSizeinMB = @()
$OnPrem_Mailboxes_TotalSizeinMB = Get-Recipient -RecipientType $RecipientType -ResultSize Unlimited -ErrorAction SilentlyContinue | Select -expandproperty Alias | Get-MailboxStatistics -ErrorAction SilentlyContinue | select MailboxGuid,@{name="TotalItemSizeinMB"; expression={[math]::Round( `
    ($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB))}}
Write-Host "`t`tOnPrem_Mailboxes_TotalSizeinMB: $($Timer.ElapsedMilliseconds - $cp) milliseconds"

Write-Host "Getting OnPremise Archive TotalMailboxSizes"
$cp = $Timer.ElapsedMilliseconds
$OnPrem_Mailboxes_Archive_TotalSizeinMB = @()
$OnPrem_Mailboxes_Archive_TotalSizeinMB = Get-Recipient -RecipientType $RecipientType -ResultSize Unlimited -ErrorAction SilentlyContinue -Filter "ArchiveState -ne 'None'" | Select -expandproperty Alias | Get-MailboxStatistics -Archive -ErrorAction SilentlyContinue | select MailboxGuid,@{name="TotalItemSizeinMB"; expression={[math]::Round( `
    ($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB))}}
Write-Host "`t`tOnPrem_Mailboxes_Archive_TotalSizeinMB: $($Timer.ElapsedMilliseconds - $cp) milliseconds"

$OnPrem_MailboxesCount = @($OnPrem_Mailboxes).Count
Write-Host "To process:" $OnPrem_MailboxesCount "mailboxes"

Add-Content $PathtoAddressesOutfile SamAccountName'>'UPN'>'PrimarySMTPAddress'>'MailDomain'>'Realm'>'World'>'MbxType'>'DisplayName'>'Company'>'Department'>'intExt'>'Owner'>'MbxSizeMB'>'HasArchive'>'ArchiveSizeMB

for ($index = 0; $index -lt $OnPrem_MailboxesCount; $index++) {
    Write-Host "`tProcessing mailbox: " ($index + 1) "/ $OnPrem_MailboxesCount"

    $OnPremMailbox = $OnPrem_Mailboxes[$index]

    $SamAccountName = $OnPremMailbox.SamAccountName
    $PrimarySMTPAddress = $OnPremMailbox.PrimarySMTPAddress
    $MailDomain = $PrimarySMTPAddress.ToString().split('@')[1]
    $Realm = "OP"
    $UPN = Get-Mailbox $OnPremMailbox.Alias | select -expandproperty UserPrincipalName
    $OnPremMailbox_CA8 = $OnPremMailbox.CustomAttribute8
    if ($OnPremMailbox_CA8) {
        $World = $OnPremMailbox_CA8.ToString().split('=')[1][0]
    }
    $MbxType = $OnPremMailbox.RecipientTypeDetails
    if ($MbxType -like "UserMailbox" -and $SamAccountName -like "SRV*") {
        $MbxType = "SrvMailbox"
    }
    $DisplayName = $OnPremMailbox.DisplayName
    $Company = $OnPremMailbox.CustomAttribute9.split('#')[0]
    $Department = $OnPremMailbox.Department
    $intExt = $OnPremMailbox.CustomAttribute10
    $Owner = $OnPremMailbox.CustomAttribute3
    $MbxSizeMB = $OnPrem_Mailboxes_TotalSizeinMB | ? {$_.MailboxGuid -eq $OnPremMailbox.ExchangeGuid} | select -expandproperty TotalItemSizeinMB
    if ($OnPremMailbox.ArchiveState -ne "None") {
        $HasArchive = "1"
        $ArchiveSizeMB = $OnPrem_Mailboxes_Archive_TotalSizeinMB | ? {$_.MailboxGuid -eq $OnPremMailbox.ArchiveGuid} | select -expandproperty TotalItemSizeinMB
    } else {
        $HasArchive = "0"
        $ArchiveSizeMB = "0"
    }
        
    Add-Content $PathtoAddressesOutfile $SamAccountName">"$UPN">"$PrimarySMTPAddress">"$MailDomain">"$Realm">"$World">"$MbxType">"$DisplayName">"$Company">"$Department">"$intExt">"$Owner">"$MbxSizeMB">"$HasArchive">"$ArchiveSizeMB
}

Write-Host "`t`tOnPrem_Mailboxes_for: $($Timer.ElapsedMilliseconds - $cp) milliseconds"
Write-Host "`t`tOnPrem_Mailboxes_TOTAL: $($Timer.ElapsedMilliseconds) milliseconds"

$SmtpServer = "smtp.domain.com"
$att = new-object Net.Mail.Attachment($PathtoAddressesOutfile)
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($SmtpServer)
$msg.From = "noreply_mailbox_details@domain1.com"
$msg.Cc.Add("user@domain1.com")
$msg.To.Add("user1@domain1.com")
$msg.To.Add("user2@domain1.com")
$msg.Subject = "Mailbox list complex OnPrem report is ready"
$msg.Body = "Attached is the mailbox list complex OnPrem report"
$msg.Attachments.Add($att)
$smtp.Send($msg)
Stop-Transcript
