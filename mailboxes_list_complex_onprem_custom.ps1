$ts = Get-Date -Format yyyyMMdd_hhmmss
$Path = "\\path\outputs"
$ProjName = "mailboxes_list"
$FileName = "Mailboxes_list_complex_onprem.custom.$ts.xls"
New-Item -Name $ProjName -Path $Path -Type Directory -ErrorAction SilentlyContinue
New-Item -Name $FileName -Path $Path\$ProjName -Type File -ErrorAction SilentlyContinue
$PathtoAddressesOutfile = "$Path\$ProjName\$FileName"

Start-Transcript "$Path\$ProjName\$FileName.txt"


$MailboxPool = Get-Content "\\path\inputs\mapping_export.csv"

# Adresses of the users in scope
# On premise: Which mailbox type should be searched for? Possible Values: "DiscoveryMailbox, EquipmentMailbox, GroupMailbox, LegacyMailbox, LinkedMailbox, LinkedRoomMailbox, RoomMailbox, SchedulingMailbox, SharedMailbox, TeamMailbox, UserMailbox"
$RecipientType = "UserMailbox"
# Get the Dataset of Mailboxes and RemoteMailboxes to work with
Write-Host "Getting Onprem $RecipientType"
$Timer = [System.diagnostics.stopwatch]::startNew()
$cp = $Timer.ElapsedMilliseconds
$Onprem_Mailboxes = $MailboxPool | Get-Recipient -RecipientType $RecipientType -ResultSize Unlimited -ErrorAction SilentlyContinue
Write-Host "Found" @($Onprem_Mailboxes).Count "$RecipientType Mailboxes"
Write-Host "`t`tOnprem_Mailboxes: $($Timer.ElapsedMilliseconds - $cp)"

Write-Host "Getting Onprem TotalMailboxSizes"
$cp = $Timer.ElapsedMilliseconds
$Onprem_Mailboxes_TotalSizeinMB = @()
$Onprem_Mailboxes_TotalSizeinMB = $MailboxPool | Get-Recipient -RecipientType $RecipientType -ResultSize Unlimited -ErrorAction SilentlyContinue | Select -expandproperty Alias | Get-MailboxStatistics -ErrorAction SilentlyContinue | select MailboxGuid,@{name="TotalItemSizeinMB"; expression={[math]::Round( `
    ($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB))}}
$Onprem_Mailboxes_TotalDeletedItemSizeinMB = @()
$Onprem_Mailboxes_TotalDeletedItemSizeinMB = $MailboxPool | Get-Recipient -RecipientType $RecipientType -ResultSize Unlimited -ErrorAction SilentlyContinue | Get-MailboxStatistics -ErrorAction SilentlyContinue | select MailboxGuid,@{name="TotalDeletedItemSizeinMB"; expression={[math]::Round( `
    ($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB))}}

Write-Host "`t`tOnprem_Mailboxes_TotalSizeinMB: $($Timer.ElapsedMilliseconds - $cp)"

Write-Host "Getting Onprem Archive TotalMailboxSizes"
$cp = $Timer.ElapsedMilliseconds
$Onprem_Mailboxes_Archive = $MailboxPool | Get-Mailbox -ResultSize Unlimited -ErrorAction SilentlyContinue | select ExchangeGuid,ArchiveName,ArchiveGuid
$Onprem_Mailboxes_Archive_TotalSizeinMB = @()
$Onprem_Mailboxes_Archive_TotalSizeinMB = $MailboxPool | Get-Recipient -RecipientType $RecipientType -ErrorAction SilentlyContinue -Filter "ArchiveState -ne 'None'" | Select -expandproperty Alias | Get-MailboxStatistics -Archive -ErrorAction SilentlyContinue | select MailboxGuid,@{name="TotalItemSizeinMB"; expression={[math]::Round( `
    ($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB))}}
$Onprem_Mailboxes_Archive_TotalDeletedItemSizeinMB = @()
$Onprem_Mailboxes_Archive_TotalDeletedItemSizeinMB = $MailboxPool |  Get-Recipient -RecipientType $RecipientType -ErrorAction SilentlyContinue -Filter "ArchiveState -ne 'None'" | Select -expandproperty Alias | Get-MailboxStatistics -Archive -ErrorAction SilentlyContinue | select MailboxGuid,@{name="TotalDeletedItemSizeinMB"; expression={[math]::Round( `
    ($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB))}}
Write-Host "`t`tOnprem_Mailboxes_Archive_TotalSizeinMB: $($Timer.ElapsedMilliseconds - $cp)"

$Onprem_MailboxesCount = @($Onprem_Mailboxes).Count
Write-Host "To process:" $Onprem_MailboxesCount "mailboxes"

    Add-Content $PathtoAddressesOutfile AliasUID'>'DisplayName'>'SamAccountName'>'UPN'>'PrimarySMTPAddress'>'RecipientType'>'MbxType'>'Mailboxplan'>'ArchiveState'>'MaxSendSize'>'MaxReceiveSize'>'IssueWarningQuota'>'ProhibitSendQuota'>'ProhibitSendReceiveQuota'>'TotalDeletedItemSizeinMB'>'MbxSizeMB'>'ArchiveDisplayName'>'ArchiveGuid'>'ArchiveDeletedItemSizeMB'>'ArchiveSizeMB

$cp = $Timer.ElapsedMilliseconds
for ($index = 0; $index -lt $Onprem_MailboxesCount; $index++) {
    Write-Host "`tProcessing mailbox: " ($index + 1) "/ $Onprem_MailboxesCount"

    $OnpremMailbox = $Onprem_Mailboxes[$index]

    $AliasUID = $OnpremMailbox.Alias
    $SamAccountName = $OnpremMailbox.SamAccountName
    $PrimarySMTPAddress = $OnpremMailbox.PrimarySMTPAddress
    $UPN = Get-Mailbox $AliasUID | select -expandproperty UserPrincipalName
    $RecipientType = $OnpremMailbox.RecipientType
    $MbxType = $OnpremMailbox.RecipientTypeDetails
    if ($MbxType -like "UserMailbox" -and $SamAccountName -like "SRV*") {
        $MbxType = "SrvMailbox"
    }
    $DisplayName = $OnpremMailbox.DisplayName

    $Mailboxplan = Get-Mailbox $AliasUID | select -expandproperty Mailboxplan
    $MaxSendSize = Get-Mailbox $AliasUID | select -expandproperty MaxSendSize
    $MaxReceiveSize = Get-Mailbox $AliasUID | select -expandproperty MaxReceiveSize
    $IssueWarningQuota = Get-Mailbox $AliasUID | select -expandproperty IssueWarningQuota
    $ProhibitSendQuota = Get-Mailbox $AliasUID | select -expandproperty ProhibitSendQuota
    $ProhibitSendReceiveQuota = Get-Mailbox $AliasUID | select -expandproperty ProhibitSendReceiveQuota

    $ArchiveState = $OnpremMailbox.ArchiveState
    $MbxSizeMB = $Onprem_Mailboxes_TotalSizeinMB | ? {$_.MailboxGuid -eq $OnpremMailbox.ExchangeGuid} | select -expandproperty TotalItemSizeinMB
    $TotalDeletedItemSizeinMB = $Onprem_Mailboxes_TotalDeletedItemSizeinMB | ? {$_.MailboxGuid -eq $OnpremMailbox.ExchangeGuid} | select -expandproperty TotalDeletedItemSizeinMB

    $ArchiveDisplayName = $null
    $ArchiveGuid = $null
    $ArchiveDeletedItemSizeMB = $null
    $ArchiveSizeMB = $null
    if ($OnpremMailbox.ArchiveState -ne "None") {
        $ArchiveDisplayName = $Onprem_Mailboxes_Archive | ? {$_.ExchangeGuid -eq $OnpremMailbox.ExchangeGuid} | select -expandproperty ArchiveName
        $ArchiveGuid = $Onprem_Mailboxes_Archive | ? {$_.ExchangeGuid -eq $OnpremMailbox.ExchangeGuid} | select -expandproperty ArchiveGuid
        $ArchiveDeletedItemSizeMB = $Onprem_Mailboxes_Archive_TotalDeletedItemSizeinMB | ? {$_.MailboxGuid -eq $OnpremMailbox.ArchiveGuid} | select -expandproperty TotalDeletedItemSizeinMB
        $ArchiveSizeMB = $Onprem_Mailboxes_Archive_TotalSizeinMB | ? {$_.MailboxGuid -eq $OnpremMailbox.ArchiveGuid} | select -expandproperty TotalItemSizeinMB
    }

    Add-Content $PathtoAddressesOutfile $AliasUID'>'$DisplayName'>'$SamAccountName'>'$UPN'>'$PrimarySMTPAddress'>'$RecipientType'>'$MbxType'>'$Mailboxplan'>'$ArchiveState'>'$MaxSendSize'>'$MaxReceiveSize'>'$IssueWarningQuota'>'$ProhibitSendQuota'>'$ProhibitSendReceiveQuota'>'$TotalDeletedItemSizeinMB'>'$MbxSizeMB'>'$ArchiveDisplayName'>'$ArchiveGuid'>'$ArchiveDeletedItemSizeMB'>'$ArchiveSizeMB
}

Write-Host "`t`tOnprem_Mailboxes_for: $($Timer.ElapsedMilliseconds - $cp)"
Write-Host "`t`tOnprem_Mailboxes_TOTAL: $($Timer.ElapsedMilliseconds)"

$SmtpServer = "smtp.domain.com"
$att = new-object Net.Mail.Attachment($PathtoAddressesOutfile)
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($SmtpServer)
$msg.From = "noreply_mailbox_details@domain1.com"
$msg.To.Add("user@domain1.com")
$msg.Subject = "Mailbox list complex Onprem report is ready"
$msg.Body = "Attached is the mailbox list complex Onprem report"
$msg.Attachments.Add($att)
$smtp.Send($msg)

Stop-Transcript
