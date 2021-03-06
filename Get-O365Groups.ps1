using namespace System.Management.Automation

begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"
}

process {
    Start-Transcript "$($params.outputFilePath).txt"

    $groups = Get-Group -ResultSize Unlimited -Filter "RecipientTypeDetails -eq 'GroupMailbox'" |
        Select-Object WindowsEmailAddress, ManagedBy, Name, RecipientType, GUID
    $groupsCount = @($groups).Count

    Write-Output "To process: $groupsCount groups"

    for ($index = 0; $index -lt $groupsCount; $index++) {
        Write-Output "`tProcessing group: $($index + 1) / $groupsCount"

        Start-Sleep -Milliseconds 500

        $group = $groups[$index]
        $groupSMTP = $group.WindowsEmailAddress

        $groupManagers = $group.ManagedBy |
            Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue |
            Select-Object CustomAttribute8, PrimarySmtpAddress, Company

        $groupMembers = Get-Group -Identity $groupSMTP -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty Members |
            Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue |
            Select-Object CustomAttribute8, PrimarySmtpAddress

        $areManagersInBothCompanies = false
        $areMembersInBothCompanies = false

        if ($groupManagers) {
            $groupManagersCount = ($groupManagers | Measure-Object).Count
            $groupManagerProperties = $groupManagers.CustomAttribute8

            $firstCompanyManagersCount = (
                $groupManagerProperties | Where-Object { $_ -like 'CAA*' } | Measure-Object
            ).Count
            $secondCompanyManagersCount = (
                $groupManagerProperties | Where-Object { $_ -like 'CAB*' } | Measure-Object
            ).Count

            $areManagersInBothCompanies = (
                $firstCompanyManagersCount -and $secondCompanyManagersCount
            )

            $groupUserProperties = $groupManagerProperties -join ';'
            $groupUsersCount = $groupManagersCount

            $firstCompanyUsersCount = $firstCompanyManagersCount
            $secondCompanyUsersCount = $secondCompanyManagersCount

            $groupManagersSMTPAddresses = (
                $groupManagers | Select-Object -ExpandProperty PrimarySmtpAddress
            ) -join ';'
            $groupManagersCompanies = (
                $groupManagers | Select-Object -ExpandProperty Company
            ) -join ';'

            $groupMembersEmails = ''
        } elseif ($groupMembers) {
            $groupMembersCount = ($groupMembers | Measure-Object).Count
            $groupMemberProperties = $groupMembers.CustomAttribute8

            $firstCompanyMembersCount = (
                $groupMemberProperties | Where-Object { $_ -like 'CAA*' } | Measure-Object
            ).Count
            $secondCompanyMembersCount = (
                $groupMemberProperties | Where-Object { $_ -like 'CAB*' } | Measure-Object
            ).Count

            $areMembersInBothCompanies = (
                $firstCompanyMembersCount -and $secondCompanyMembersCount
            )

            $groupUserProperties = $groupMemberProperties -join ';'
            $groupUsersCount = $groupMembersCount

            $firstCompanyUsersCount = $firstCompanyMembersCount
            $secondCompanyUsersCount = $secondCompanyMembersCount

            $groupManagersSMTPAddresses = ''
            $groupManagersCompanies = ''

            $groupMembersEmails = $groupMembers.PrimarySmtpAddress -join ';'
        } else {
            $groupUserProperties = ''
            $groupUsersCount = 0

            $firstCompanyUsersCount = 0
            $secondCompanyUsersCount = 0

            $groupManagersSMTPAddresses = ''
            $groupManagersCompanies = ''

            $groupMembersEmails = ''
        }

        $hasFirstCompanyUsers = $firstCompanyUsersCount -gt 0
        $hasSecondCompanyUsers = $secondCompanyUsersCount -gt 0
        $hasFirstOrSecondCompanyUsers = $hasFirstCompanyUsers -or $hasSecondCompanyUsers

        $hasOnlyFirstCompanyUsers = $hasFirstCompanyUsers -and (-not $hasSecondCompanyUsers)
        $hasOnlySecondCompanyUsers = (-not $hasFirstCompanyUsers) -and $hasSecondCompanyUsers

        if (-not $hasFirstOrSecondCompanyUsers) {
            $groupCompany = 'None'
        } elseif ($areManagersInBothCompanies) {
            $groupCompany = 'Mixed Owners'
        } elseif ($areMembersInBothCompanies) {
            $groupCompany = 'Mixed Users'
        } elseif ($firstCompanyUsersCount -eq $groupUsersCount -or $hasOnlyFirstCompanyUsers) {
            $groupCompany = 'compA'
        } elseif ($secondCompanyUsersCount -eq $groupUsersCount -or $hasOnlySecondCompanyUsers) {
            $groupCompany = 'compB'
        } else {
            $groupCompany = 'N/A'
        }

        if ($groupCompany -eq 'compA' -or $groupCompany -like 'Mixed*') {
            $groupName = $group.Name
            $groupCategory = $group.RecipientType
            $groupGUID = $group.GUID

            [PSCustomObject]@{
                'Group' = $group
                'Group Name' = $groupName
                'Group GUID' = $groupGUID
                'Group SMTP' = $groupSMTP
                'Group Category' = $groupCategory
                'Group Company' = $groupCompany
                'Group Manager or Member Properties' = $groupUserProperties
                'Group Managers or Members Count' = $groupUsersCount
                'First Company Managers Or Members Count' = $firstCompanyUsersCount
                'Second Company Managers Or Members Count' = $secondCompanyUsersCount
                'Group Managers SMTP Addresses' = $groupManagersSMTPAddresses
                'Group Managers Companies' = $groupManagersCompanies
                'Group Members Emails' = $groupMembersEmails
            } | Export-Csv $params.outputFilePath -Append -NoTypeInformation
        }
    }

    $attachment = New-Object Net.Mail.Attachment($params.outputFilePath)

    $message = New-Object Net.Mail.MailMessage
    $message.From = 'noreply_group_details@compA.com'
    $message.Cc.Add('user1@compA.com')
    $message.To.Add('user2@compA.com')
    $message.Subject = "$outputFileName report is ready"
    $message.Body = "Attached is the $outputFileName report"
    $message.Attachments.Add($attachment)

    $smtpServer = 'smtp.compB.com'
    $smtp = New-Object Net.Mail.SmtpClient($smtpServer)
    $smtp.Send($message)

    Stop-Transcript
}
