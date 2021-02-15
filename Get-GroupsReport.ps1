begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"
}

Start-Transcript

$Groups = Get-DistributionGroup -ResultSize Unlimited
$GroupsCount = @($Groups).Count
Write-Host "To process:" $GroupsCount "groups"

Add-Content $PathtoAddressesOutfile Displayname"`";`""Name"`";`""Alias"`";`""LegacyExchangeDN"`";`""SMTP"`";`""SMTPPrefix"`";`""Company"`";`""RequireSenderAuthenticationEnabled"`";`""HiddenFromAddressListsEnabled"`";`""Category"`";`""ManagedBy_SMTP"`";`""ManagedBy_ExtAtt"`";`""EmailAddresses"`";`""AcceptMessagesOnlyFromExtAtt"`";`""AcceptMessagesOnlyFromDLMembersExtAtt"`";`""AcceptMessagesOnlyFromSendersOrMembersExtAtt"`";`""RejectMessagesFromExtAtt"`";`""RejectMessagesFromDLMembersExtAtt"`";`""RejectMessagesFromSendersOrMembersExtAtt

for ($index = 0; $index -lt $GroupsCount; $index++) {
    Write-Host "`tProcessing group: " ($index + 1) "/ $GroupsCount"

    $Group = $Groups[$index]

    $DLManagers = $Group.ManagedBy | Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue

    $Group_AD = Get-ADGroup -Identity $Group.DistinguishedName
    $Group_Members = Get-ADGroupMember -identity $Group_AD -ErrorAction SilentlyContinue -recursive | ? {(Get-ADUser -Identity $_.objectGUID).Enabled -eq "true" -and ((Get-ADUser -Identity $_.objectGUID -Properties mail).mail)}

    if ($DLManagers) {
        $DLManagerscount = ($DLManagers | measure-object).count
        $RecipientUserProperties = $DLManagers.customattribute8 # | Select customattribute8 -ExpandProperty customattribute8
        $compBManagerscount = ($RecipientUserProperties | ? {$_ -like "CAB*"} | measure-object).count
        $compAManagerscount = ($RecipientUserProperties | ? {$_ -like "CAA*"} | measure-object).count
        $RecipientUserProperties = $RecipientUserProperties -join ";"
        $UserProperties = $RecipientUserProperties
        $DLcountORDLManagerscount = $DLManagerscount
        $compBcountORcompBManagerscount = $compBManagerscount
        $compAcountORcompAManagerscount = $compAManagerscount
    } else {
        $DLcount = ($Group_Members | measure-object).count
        $ADUserProperties = $Group_Members | Get-ADUser -Properties extensionattribute8 -ErrorAction SilentlyContinue | ? {($_.extensionattribute8)} | Select extensionattribute8 -ExpandProperty extensionattribute8
        $compBcount = ($ADUserProperties | ? {$_ -like "CAB*"} | measure-object).count
        $compAcount = ($ADUserProperties | ? {$_ -like "CAA*"} | measure-object).count
        $ADUserProperties = $ADUserProperties -join ";"
        $UserProperties = $ADUserProperties
        $DLcountORDLManagerscount = $DLcount
        $compBcountORcompBManagerscount = $compBcount
        $compAcountORcompAManagerscount = $compAcount
    }

    if ($compBcountORcompBManagerscount -eq 0 -and $DLcountORDLManagerscount -eq 0 -and $compAcountORcompAManagerscount -eq 0) {
        $Group_company = "None"
    } elseif ($compBcount -and $compAcount) {
        $Group_company = "Mixed Users"
    } elseif ($compBManagerscount -and $compAManagerscount) {
        $Group_company = "Mixed Owners"
    } elseif (($compBcountORcompBManagerscount -eq $DLcountORDLManagerscount) -or ($compBcountORcompBManagerscount -eq 0 -and $compAcountORcompAManagerscount -eq 0 -and $Group_ManagedBy_Company -match "compB" -and $Group_ManagedBy_Company -notmatch "compA") -or (($compBcountORcompBManagerscount) -and $compAcountORcompAManagerscount -eq 0)) {
        $Group_company = "compB"
    } elseif ($compAcountORcompAManagerscount -eq $DLcountORDLManagerscount -or ($compBcountORcompBManagerscount -eq 0 -and $compAcountORcompAManagerscount -eq 0 -and $Group_ManagedBy_Company -match "compA" -and $Group_ManagedBy_Company -notmatch "compB") -or (($compAcountORcompAManagerscount) -and $compBcountORcompBManagerscount -eq 0)) {
        $Group_company = "compA"
    }

    if ($Group_company -eq "compA" -or $Group_company -like "Mixed*") {
        $Group_ManagedBy_SMTP = @()
        # $Group_ManagedBy_Company = @()
        # $Group_ManagedBy_CA8 = @()
        $Group_ManagedBy_ExtAtt = @()

        Foreach ($Manager in $DLManagers) {
            $Group_ManagedBy_SMTP += $Manager | select PrimarySMTPAddress -ErrorAction SilentlyContinue | % {
                $_.PrimarySMTPAddress.ToString()
            }
            # $Group_ManagedBy_Company += $Manager | select Company | % {
            #     $_.Company.ToString()
            # }
            # $Group_ManagedBy_CA8 += $Manager | select customattribute8 | % {
            #     $_.customattribute8.ToString()
            # }
            $Group_ManagedBy_ExtAtt += $Manager.Alias | Get-ADUser -Properties msDS-SourceAnchor | select -ExpandProperty msDS-SourceAnchor
        }

        $Group_ManagedBy_SMTP = $Group_ManagedBy_SMTP -join "#"
        # $Group_ManagedBy_Company = $Group_ManagedBy_Company -join "#"
        # $Group_ManagedBy_CA8 = $Group_ManagedBy_CA8 -join "#"
        $Group_ManagedBy_ExtAtt = $Group_ManagedBy_ExtAtt -join "#"

        $Group_CA3 = $Group.customattribute3
        $Group_SMTP = $Group.PrimarySMTPAddress
        $Group_SMTPPrefix = $Group_SMTP.Split("@")[0]
        $Group_Displayname = $Group.DisplayName
        $Group_Name = $Group.Name
        $Group_Alias = $Group.Alias
        $Group_Category = $Group.RecipientType
        $Group_GUID = $Group.GUID
        $Group_LegacyExchangeDN = $Group.LegacyExchangeDN
        $Group_RequireSenderAuthenticationEnabled = $Group.RequireSenderAuthenticationEnabled
        $Group_MembersEmail = $Group_Members | Get-ADUser -Property mail | Select-Object -ExpandProperty mail
        $Group_MembersEmail = $GroupMembersEmail -join "#"
        $Group_MembersExtAtt = $Group_Members | Get-ADUser -Property msDS-SourceAnchor | Select-Object -ExpandProperty msDS-SourceAnchor
        $Group_MembersExtAtt = $GroupMembersExtAtt -join "#"
        $Group_HiddenFromAddressListsEnabled = $Group.HiddenFromAddressListsEnabled
        $Group_EmailAddresses = ($Group.EmailAddresses | ? {$_ -notlike "*compB*" -and $_ -match "smtp"}).Replace("SMTP","smtp")
        $Group_EmailAddresses = $Group_EmailAddresses -join "#"

        $Group_AcceptMessagesOnlyFrom = $Group.AcceptMessagesOnlyFrom | Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue | Select -ExpandProperty Alias
        $Group_AcceptMessagesOnlyFromExtAtt = $Group_AcceptMessagesOnlyFrom | Get-ADUser -Properties msDS-SourceAnchor | select -ExpandProperty msDS-SourceAnchor
        $Group_AcceptMessagesOnlyFromExtAtt = $Group_AcceptMessagesOnlyFromExtAtt -join "#"

        $Group_AcceptMessagesOnlyFromDLMembers = $Group.AcceptMessagesOnlyFromDLMembers | Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue | Select -ExpandProperty Alias
        $Group_AcceptMessagesOnlyFromDLMembersExtAtt = $Group_AcceptMessagesOnlyFromDLMembers | Get-ADUser -Properties msDS-SourceAnchor | select -ExpandProperty msDS-SourceAnchor
        $Group_AcceptMessagesOnlyFromDLMembersExtAtt = $Group_AcceptMessagesOnlyFromDLMembersExtAtt -join "#"

        $Group_AcceptMessagesOnlyFromSendersOrMembers = $Group.AcceptMessagesOnlyFromSendersOrMembers | Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue | Select -ExpandProperty Alias
        $Group_AcceptMessagesOnlyFromSendersOrMembersExtAtt = $Group_AcceptMessagesOnlyFromSendersOrMembers | Get-ADUser -Properties msDS-SourceAnchor | select -ExpandProperty msDS-SourceAnchor
        $Group_AcceptMessagesOnlyFromSendersOrMembersExtAtt = $Group_AcceptMessagesOnlyFromSendersOrMembersExtAtt -join "#"

        $Group_RejectMessagesFrom = $Group.RejectMessagesFrom | Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue | Select -ExpandProperty Alias
        $Group_RejectMessagesFromExtAtt = $Group_RejectMessagesFrom | Get-ADUser -Properties msDS-SourceAnchor | select -ExpandProperty msDS-SourceAnchor
        $Group_RejectMessagesFromExtAtt = $Group_RejectMessagesFromExtAtt -join "#"

        $Group_RejectMessagesFromDLMembers = $Group.RejectMessagesFromDLMembers | Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue | Select -ExpandProperty Alias
        $Group_RejectMessagesFromDLMembersExtAtt = $Group_RejectMessagesFromDLMembers | Get-ADUser -Properties msDS-SourceAnchor | select -ExpandProperty msDS-SourceAnchor
        $Group_RejectMessagesFromDLMembersExtAtt = $Group_RejectMessagesFromDLMembersExtAtt -join "#"

        $Group_RejectMessagesFromSendersOrMembers = $Group.RejectMessagesFromSendersOrMembers | Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue | Select -ExpandProperty Alias
        $Group_RejectMessagesFromSendersOrMembersExtAtt = $Group_RejectMessagesFromSendersOrMembers | Get-ADUser -Properties msDS-SourceAnchor | select -ExpandProperty msDS-SourceAnchor
        $Group_RejectMessagesFromSendersOrMembersExtAtt = $Group_RejectMessagesFromSendersOrMembersExtAtt -join "#"

        Add-Content $PathtoAddressesOutfile $Group_Displayname"`";`""$Group_Name"`";`""$Group_Alias"`";`""$Group_LegacyExchangeDN"`";`""$Group_SMTP"`";`""$Group_SMTPPrefix"`";`""$Group_company"`";`""$Group_RequireSenderAuthenticationEnabled"`";`""$Group_HiddenFromAddressListsEnabled"`";`""$Group_Category"`";`""$Group_ManagedBy_SMTP"`";`""$Group_ManagedBy_ExtAtt"`";`""$Group_EmailAddresses"`";`""$Group_AcceptMessagesOnlyFromExtAtt"`";`""$Group_AcceptMessagesOnlyFromDLMembersExtAtt"`";`""$Group_AcceptMessagesOnlyFromSendersOrMembersExtAtt"`";`""$Group_RejectMessagesFromExtAtt"`";`""$Group_RejectMessagesFromDLMembersExtAtt"`";`""$Group_RejectMessagesFromSendersOrMembersExtAtt

    }
    $Group_Members = $null

}

$SmtpServer = "smtp.mail.com"
$att = new-object Net.Mail.Attachment($PathtoAddressesOutfile)
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($SmtpServer)
$msg.From = "noreply_get_groupsreport@compA.com"
$msg.To.Add("user1@compA.com")
$msg.Cc.Add("user2@compA.com")
$msg.Subject = "Groups report is ready"
$msg.Body = "Attached is the groups report"
$msg.Attachments.Add($att)
$smtp.Send($msg)
Stop-Transcript

