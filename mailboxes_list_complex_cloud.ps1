$ts = Get-Date -Format yyyyMMdd_hhmmss
$Path = "\\path\outputs"
$ProjName = "mailboxes_list"
$FileName = "Mailboxes_list_complex_cloud.full.$ts.xls"
New-Item -Name $ProjName -Path $Path -Type Directory -ErrorAction SilentlyContinue
New-Item -Name $FileName -Path $Path\$ProjName -Type File -ErrorAction SilentlyContinue
$PathtoAddressesOutfile = "$Path\$ProjName\$FileName"

Start-Transcript "$Path\$ProjName\Faster\$FileName.txt"

# Adresses of the users in scope
# On premise: Which mailbox type should be searched for? Possible Values: "DiscoveryMailbox, EquipmentMailbox, GroupMailbox, LegacyMailbox, LinkedMailbox, LinkedRoomMailbox, RoomMailbox, SchedulingMailbox, SharedMailbox, TeamMailbox, UserMailbox"
$RecipientType = "UserMailbox"
# Get the Dataset of Mailboxes and RemoteMailboxes to work with
Write-Host "Getting Cloud $RecipientType"
$Timer = [System.diagnostics.stopwatch]::startNew()
$cp = $Timer.ElapsedMilliseconds
$Cloud_Mailboxes = Get-EXORecipient -RecipientType $RecipientType -Properties SamAccountName, PrimarySMTPAddress, WindowsLiveId, CustomAttribute8, RecipientTypeDetails, DisplayName, CustomAttribute9, Department, CustomAttribute10, CustomAttribute3, ExchangeGuid, ArchiveGuid, ArchiveState -ResultSize Unlimited -ErrorAction SilentlyContinue
Write-Host "Found" @($Cloud_Mailboxes).Count "$RecipientType Mailboxes"
Write-Host "`t`tCloud_Mailboxes: $($Timer.ElapsedMilliseconds - $cp)"

Write-Host "Getting Cloud TotalMailboxSizes"
$cp = $Timer.ElapsedMilliseconds
$Cloud_Mailboxes_TotalSizeinMB = @()
$Cloud_Mailboxes_TotalSizeinMB = Get-EXORecipient -RecipientType $RecipientType -Properties ExchangeGuid -ResultSize Unlimited -ErrorAction SilentlyContinue | Get-EXOMailboxStatistics $_.ExchangeGuid -Properties MailboxGuid, TotalItemSize -ErrorAction SilentlyContinue | select MailboxGuid, @{name = "TotalItemSizeinMB"; expression = { [math]::Round( `
            ($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB)) }
}
Write-Host "`t`tCloud_Mailboxes_TotalSizeinMB: $($Timer.ElapsedMilliseconds - $cp)"

Write-Host "Getting Cloud Archive TotalMailboxSizes"
$cp = $Timer.ElapsedMilliseconds
$Cloud_Mailboxes_Archive_TotalSizeinMB = @()
$Cloud_Mailboxes_Archive_TotalSizeinMB = Get-EXORecipient -RecipientType $RecipientType -Properties ExchangeGuid -ResultSize Unlimited -ErrorAction SilentlyContinue -Filter "ArchiveState -ne 'None'" | Get-EXOMailboxStatistics $_.ExchangeGuid -Archive -Properties MailboxGuid, TotalItemSize -ErrorAction SilentlyContinue | select MailboxGuid, @{name = "TotalItemSizeinMB"; expression = { [math]::Round( `
            ($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB)) }
}
Write-Host "`t`tCloud_Mailboxes_Archive_TotalSizeinMB: $($Timer.ElapsedMilliseconds - $cp)"

$Cloud_MailboxesCount = @($Cloud_Mailboxes).Count
Write-Host "To process:" $Cloud_MailboxesCount "mailboxes"

Add-Content $PathtoAddressesOutfile SamAccountName'>'UPN'>'PrimarySMTPAddress'>'MailDomain'>'Realm'>'World'>'MbxType'>'DisplayName'>'Company'>'Department'>'intExt'>'Owner'>'MbxSizeMB'>'HasArchive'>'ArchiveSizeMB

$cp = $Timer.ElapsedMilliseconds
for ($index = 0; $index -lt $Cloud_MailboxesCount; $index++) {
    Write-Host "`tProcessing mailbox: " ($index + 1) "/ $Cloud_MailboxesCount"

    $CloudMailbox = $Cloud_Mailboxes[$index]

    $SamAccountName = $CloudMailbox.SamAccountName
    $PrimarySMTPAddress = $CloudMailbox.PrimarySMTPAddress
    $MailDomain = $PrimarySMTPAddress.ToString().split('@')[1]
    $Realm = "O365"
    $UPN = $CloudMailbox.WindowsLiveId
    $CloudMailbox_CA8 = $CloudMailbox.CustomAttribute8
    if ($CloudMailbox_CA8) {
        $World = $CloudMailbox_CA8.ToString().split('=')[1][0]
    }
    $MbxType = $CloudMailbox.RecipientTypeDetails
    if ($MbxType -like "UserMailbox" -and $SamAccountName -like "service*") {
        $MbxType = "SrvMailbox"
    }
    $DisplayName = $CloudMailbox.DisplayName
    $Company = $CloudMailbox.CustomAttribute9.split('#')[0]
    $Department = $CloudMailbox.Department
    $intExt = $CloudMailbox.CustomAttribute10
    $Owner = $CloudMailbox.CustomAttribute3
    $MbxSizeMB = $Cloud_Mailboxes_TotalSizeinMB | ? { $_.MailboxGuid -eq $CloudMailbox.ExchangeGuid } | select -expandproperty TotalItemSizeinMB
    if ($CloudMailbox.ArchiveState -ne "None") {
        $HasArchive = "1"
        $ArchiveSizeMB = $Cloud_Mailboxes_Archive_TotalSizeinMB | ? { $_.MailboxGuid -eq $CloudMailbox.ArchiveGuid } | select -expandproperty TotalItemSizeinMB
    } else {
        $HasArchive = "0"
        $ArchiveSizeMB = "0"
    }

    Add-Content $PathtoAddressesOutfile $SamAccountName'>'$UPN'>'$PrimarySMTPAddress'>'$MailDomain'>'$Realm'>'$World'>'$MbxType'>'$DisplayName'>'$Company'>'$Department'>'$intExt'>'$Owner'>'$MbxSizeMB'>'$HasArchive'>'$ArchiveSizeMB
}

Write-Host "`t`tCloud_Mailboxes_for: $($Timer.ElapsedMilliseconds - $cp)"
Write-Host "`t`tCloud_Mailboxes_TOTAL: $($Timer.ElapsedMilliseconds)"

$SmtpServer = "smtp.domain.com"
$att = new-object Net.Mail.Attachment($PathtoAddressesOutfile)
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($SmtpServer)
$msg.From = "noreply_mailbox_details@domain1.com"
$msg.Cc.Add("user@domain1.com")
$msg.To.Add("user1@domain1.com")
$msg.To.Add("user2@domain1.com")
$msg.Subject = "Mailbox list complex Cloud report is ready"
$msg.Body = "Attached is the mailbox list complex Cloud report"
$msg.Attachments.Add($att)
$smtp.Send($msg)

Stop-Transcript
