$ts = Get-Date -Format yyyyMMdd_hhmmss
$Path = "\\path\outputs"
$ProjName = "mailboxes_list"
$FileName = "Mailboxes_list_complex_cloud.custom.$ts.xls"
New-Item -Name $ProjName -Path $Path -Type Directory -ErrorAction SilentlyContinue
New-Item -Name $FileName -Path $Path\$ProjName -Type File -ErrorAction SilentlyContinue
$PathtoAddressesOutfile = "$Path\$ProjName\$FileName"

Start-Transcript "$Path\$ProjName\$FileName.txt"


$MailboxPool = Get-Content "\\path\inputs\mapping_export.csv"

# Adresses of the users in scope
# On premise: Which mailbox type should be searched for? Possible Values: "DiscoveryMailbox, EquipmentMailbox, GroupMailbox, LegacyMailbox, LinkedMailbox, LinkedRoomMailbox, RoomMailbox, SchedulingMailbox, SharedMailbox, TeamMailbox, UserMailbox"
$RecipientType = "UserMailbox"
# Get the Dataset of Mailboxes and RemoteMailboxes to work with
Write-Host "Getting Cloud $RecipientType"
$Timer = [System.diagnostics.stopwatch]::startNew()
$cp = $Timer.ElapsedMilliseconds
$Cloud_Mailboxes = $MailboxPool | Get-EXORecipient -RecipientType $RecipientType -Properties SamAccountName,PrimarySMTPAddress,WindowsLiveId,CustomAttribute8,RecipientTypeDetails,DisplayName,CustomAttribute9,Department,CustomAttribute10,CustomAttribute3,ExchangeGuid,ArchiveGuid,ArchiveState -ErrorAction SilentlyContinue
Write-Host "Found" @($Cloud_Mailboxes).Count "$RecipientType Mailboxes"
Write-Host "`t`tCloud_Mailboxes: $($Timer.ElapsedMilliseconds - $cp)"

Write-Host "Getting Cloud TotalMailboxSizes"
$cp = $Timer.ElapsedMilliseconds
$Cloud_Mailboxes_TotalSizeinMB = @()
$Cloud_Mailboxes_TotalSizeinMB = $MailboxPool | Get-EXORecipient -Properties ExchangeGuid -ErrorAction SilentlyContinue | Get-EXOMailboxStatistics $_.ExchangeGuid -Properties MailboxGuid,TotalItemSize -ErrorAction SilentlyContinue | select MailboxGuid,@{name="TotalItemSizeinMB"; expression={[math]::Round( `
    ($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB))}}
$Cloud_Mailboxes_TotalDeletedItemSizeinMB = @()
$Cloud_Mailboxes_TotalDeletedItemSizeinMB = $MailboxPool | Get-EXORecipient -Properties ExchangeGuid -ErrorAction SilentlyContinue | Get-EXOMailboxStatistics $_.ExchangeGuid -Properties MailboxGuid,TotalDeletedItemSize -ErrorAction SilentlyContinue | select MailboxGuid,@{name="TotalDeletedItemSizeinMB"; expression={[math]::Round( `
    ($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB))}}

Write-Host "`t`tCloud_Mailboxes_TotalSizeinMB: $($Timer.ElapsedMilliseconds - $cp)"

Write-Host "Getting Cloud Archive TotalMailboxSizes"
$cp = $Timer.ElapsedMilliseconds
$Cloud_Mailboxes_Archive = $MailboxPool | Get-EXOMailbox -Properties ExchangeGuid,ArchiveName,ArchiveGuid -ErrorAction SilentlyContinue | select ExchangeGuid,ArchiveName,ArchiveGuid
$Cloud_Mailboxes_Archive_TotalSizeinMB = @()
$Cloud_Mailboxes_Archive_TotalSizeinMB = $MailboxPool | Get-Recipient -Properties ExchangeGuid -ErrorAction SilentlyContinue -Filter "ArchiveState -ne 'None'" | Get-ExoMailboxStatistics $_.ExchangeGuid -Archive -Properties MailboxGuid,TotalItemSize -ErrorAction SilentlyContinue | select MailboxGuid,@{name="TotalItemSizeinMB"; expression={[math]::Round( `
    ($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB))}}
$Cloud_Mailboxes_Archive_TotalDeletedItemSizeinMB = @()
$Cloud_Mailboxes_Archive_TotalDeletedItemSizeinMB = $MailboxPool | Get-Recipient -Properties ExchangeGuid -ErrorAction SilentlyContinue -Filter "ArchiveState -ne 'None'" | Get-ExoMailboxStatistics $_.ExchangeGuid -Archive -Properties MailboxGuid,TotalDeletedItemSize -ErrorAction SilentlyContinue | select MailboxGuid,@{name="TotalDeletedItemSizeinMB"; expression={[math]::Round( `
    ($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB))}}
Write-Host "`t`tCloud_Mailboxes_Archive_TotalSizeinMB: $($Timer.ElapsedMilliseconds - $cp)"

$Cloud_MailboxesCount = @($Cloud_Mailboxes).Count
Write-Host "To process:" $Cloud_MailboxesCount "mailboxes"

    Add-Content $PathtoAddressesOutfile AliasUID'>'DisplayName'>'SamAccountName'>'UPN'>'PrimarySMTPAddress'>'RecipientType'>'MbxType'>'Mailboxplan'>'ArchiveState'>'MaxSendSize'>'MaxReceiveSize'>'IssueWarningQuota'>'ProhibitSendQuota'>'ProhibitSendReceiveQuota'>'TotalDeletedItemSizeinMB'>'MbxSizeMB'>'ArchiveDisplayName'>'ArchiveGuid'>'ArchiveDeletedItemSizeMB'>'ArchiveSizeMB

$cp = $Timer.ElapsedMilliseconds
for ($index = 0; $index -lt $Cloud_MailboxesCount; $index++) {
    Write-Host "`tProcessing mailbox: " ($index + 1) "/ $Cloud_MailboxesCount"

    $CloudMailbox = $Cloud_Mailboxes[$index]

    $AliasUID = $CloudMailbox.Alias
    $SamAccountName = $CloudMailbox.SamAccountName
    $PrimarySMTPAddress = $CloudMailbox.PrimarySMTPAddress
    $UPN = $CloudMailbox.WindowsLiveId
    $RecipientType = $CloudMailbox.RecipientType
    $MbxType = $CloudMailbox.RecipientTypeDetails
    if ($MbxType -like "UserMailbox" -and $SamAccountName -like "SRV*") {
        $MbxType = "SrvMailbox"
    }
    $DisplayName = $CloudMailbox.DisplayName
    $Mailboxplan = Get-EXOMailbox $AliasUID -Properties MailboxPlan | select -expandproperty Mailboxplan
    $MaxSendSize = Get-EXOMailbox $AliasUID -Properties MaxSendSize | select -expandproperty MaxSendSize
    $MaxReceiveSize = Get-EXOMailbox $AliasUID -Properties MaxReceiveSize | select -expandproperty MaxReceiveSize
    $IssueWarningQuota = Get-EXOMailbox $AliasUID -Properties IssueWarningQuota | select -expandproperty IssueWarningQuota
    $ProhibitSendQuota = Get-EXOMailbox $AliasUID -Properties ProhibitSendQuota | select -expandproperty ProhibitSendQuota
    $ProhibitSendReceiveQuota = Get-EXOMailbox $AliasUID -Properties ProhibitSendReceiveQuota | select -expandproperty ProhibitSendReceiveQuota

    $ArchiveState = $CloudMailbox.ArchiveState
    $MbxSizeMB = $Cloud_Mailboxes_TotalSizeinMB | ? {$_.MailboxGuid -eq $CloudMailbox.ExchangeGuid} | select -expandproperty TotalItemSizeinMB
    $TotalDeletedItemSizeinMB = $Cloud_Mailboxes_TotalDeletedItemSizeinMB | ? {$_.MailboxGuid -eq $CloudMailbox.ExchangeGuid} | select -expandproperty TotalDeletedItemSizeinMB

    $ArchiveDisplayName = $null
    $ArchiveGuid = $null
    $ArchiveDeletedItemSizeMB = $null
    $ArchiveSizeMB = $null
    if ($CloudMailbox.ArchiveState -ne "None") {
        $ArchiveDisplayName = $Cloud_Mailboxes_Archive | ? {$_.ExchangeGuid -eq $CloudMailbox.ExchangeGuid} | select -expandproperty ArchiveName
        $ArchiveGuid = $Cloud_Mailboxes_Archive | ? {$_.ExchangeGuid -eq $CloudMailbox.ExchangeGuid} | select -expandproperty ArchiveGuid
        $ArchiveDeletedItemSizeMB = $Cloud_Mailboxes_Archive_TotalDeletedItemSizeinMB | ? {$_.MailboxGuid -eq $CloudMailbox.ArchiveGuid} | select -expandproperty TotalDeletedItemSizeinMB
        $ArchiveSizeMB = $Cloud_Mailboxes_Archive_TotalSizeinMB | ? {$_.MailboxGuid -eq $CloudMailbox.ArchiveGuid} | select -expandproperty TotalItemSizeinMB
    }

    Add-Content $PathtoAddressesOutfile $AliasUID'>'$DisplayName'>'$SamAccountName'>'$UPN'>'$PrimarySMTPAddress'>'$RecipientType'>'$MbxType'>'$Mailboxplan'>'$ArchiveState'>'$MaxSendSize'>'$MaxReceiveSize'>'$IssueWarningQuota'>'$ProhibitSendQuota'>'$ProhibitSendReceiveQuota'>'$TotalDeletedItemSizeinMB'>'$MbxSizeMB'>'$ArchiveDisplayName'>'$ArchiveGuid'>'$ArchiveDeletedItemSizeMB'>'$ArchiveSizeMB
}

Write-Host "`t`tCloud_Mailboxes_for: $($Timer.ElapsedMilliseconds - $cp)"
Write-Host "`t`tCloud_Mailboxes_TOTAL: $($Timer.ElapsedMilliseconds)"

$SmtpServer = "smtp.domain.com"
$att = new-object Net.Mail.Attachment($PathtoAddressesOutfile)
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($SmtpServer)
$msg.From = "noreply_mailbox_details@domain1.com"
$msg.To.Add("user@domain1.com")
$msg.Subject = "Mailbox list complex Cloud report is ready"
$msg.Body = "Attached is the mailbox list complex Cloud report"
$msg.Attachments.Add($att)
$smtp.Send($msg)

Stop-Transcript
