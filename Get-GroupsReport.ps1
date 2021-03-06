param (
    [ValidateSet('Distribution Groups', 'Office 365 Groups')]
    [String]
    $Type
)

begin {
    enum GroupsType {
        none
        distribution
        office365
    }

    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"

    $typeParams = @{
        Name = 'Type'
        Value = $Type
        Config = $params.config
        ScriptName = $params.scriptName
    }

    $Type = Read-Param @typeParams

    switch ($Type) {
        'Distribution Groups' { $groupsType = [GroupsType]::distribution }
        'Office 365 Groups' { $groupsType = [GroupsType]::office365 }
        Default { $groupsType = [GroupsType]::none }
    }
}

process {
    if ($groupsType -eq [GroupsType]::none) {
        Write-Error "Filtering by the groups type '$Type' is not implemented!"

        return
    }

    Start-Transcript "$($params.outputFilePath).txt"

    $groups = Get-Group -ResultSize Unlimited -Filter { RecipientTypeDetails -eq 'GroupMailbox' } |
        Select-Object WindowsEmailAddress, ManagedBy, Name, RecipientType, Guid
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

        $areManagersInBothCompanies = $false
        $areMembersInBothCompanies = $false

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
        } else {
            $groupUserProperties = ''
            $groupUsersCount = 0

            $firstCompanyUsersCount = 0
            $secondCompanyUsersCount = 0

            $groupManagersSMTPAddresses = ''
            $groupManagersCompanies = ''
        }

        if ($groupMembers) {
            $groupMembersEmails = $groupMembers.PrimarySmtpAddress -join ';'
        } else {
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
        } elseif ($groupUsersCount -eq $firstCompanyUsersCount -or $hasOnlyFirstCompanyUsers) {
            $groupCompany = 'compA'
        } elseif ($groupUsersCount -eq $secondCompanyUsersCount -or $hasOnlySecondCompanyUsers) {
            $groupCompany = 'compB'
        } else {
            $groupCompany = 'N/A'
        }

        if ($groupCompany -eq 'compA' -or $groupCompany -like 'Mixed*') {
            [PSCustomObject]@{
                'Group' = $group

                'Group Name' = $group.Name
                'Group GUID' = $group.Guid
                'Group SMTP' = $groupSMTP
                'Group Category' = $group.RecipientType
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
