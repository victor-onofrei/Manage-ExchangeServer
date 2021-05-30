using namespace System.Management.Automation

param (
    [ValidateSet('Distribution Groups', 'Microsoft 365 Groups')]
    [String]
    $Type
)

begin {
    enum GroupsType {
        none
        distribution
        unified
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
        'Microsoft 365 Groups' { $groupsType = [GroupsType]::unified }
        Default { $groupsType = [GroupsType]::none }
    }

    function Get-ManagersFromList {
        param (
            [Object[]]$List
        )

        $managers = $List |
            Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue |
            Select-Object CustomAttribute8, PrimarySmtpAddress, Company

        return $managers
    }

    function Get-MembersListFromGroup {
        param (
            [PSObject]$Group
        )

        switch ($groupsType) {
            ([GroupsType]::distribution) {
                $getADGroupMemberParams = @{
                    Identity = $Group.Guid
                    Recursive = $true
                    ErrorAction = [ActionPreference]::SilentlyContinue
                }

                $list = Get-ADGroupMember @getADGroupMemberParams |
                    Get-ADUser -Identity $_.ObjectGUID -Properties mail |
                    Where-Object { $_.Enabled -eq 'True' -and $_.mail }
            }
            ([GroupsType]::unified) {
                $getGroupParams = @{
                    Identity = $Group.Guid
                    ErrorAction = [ActionPreference]::SilentlyContinue
                }

                $list = Get-Group @getGroupParams |
                    Select-Object -ExpandProperty Members
            }
        }

        return $list
    }

    function Get-MembersFromList {
        param (
            [Object[]]$List
        )

        $members = $List |
            Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue |
            Select-Object CustomAttribute8, PrimarySmtpAddress

        return $members
    }

    function Get-MembersFromGroup {
        param (
            [PSObject]$Group
        )

        $list = Get-MembersListFromGroup $Group

        if (-not $list) {
            return $null
        }

        $members = Get-MembersFromList $list

        if (-not $members) {
            return $null
        }

        return $members
    }
}

process {
    if ($groupsType -eq [GroupsType]::none) {
        Write-Error "Filtering by the groups type '$Type' is not implemented!"

        return
    }

    Start-Transcript "$($params.outputFilePath).txt"

    switch ($groupsType) {
        ([GroupsType]::distribution) {
            $groups = Get-DistributionGroup -ResultSize Unlimited |
                Select-Object WindowsEmailAddress, ManagedBy, Name, RecipientType, Guid
        }
        ([GroupsType]::unified) {
            $groups = Get-Group -ResultSize Unlimited -Filter {
                RecipientTypeDetails -eq 'GroupMailbox'
            } |
                Select-Object WindowsEmailAddress, ManagedBy, Name, RecipientType, Guid
        }
    }

    $groupsCount = @($groups).Count

    Write-Output "To process: $groupsCount groups"

    for ($index = 0; $index -lt $groupsCount; $index++) {
        Write-Output "`tProcessing group: $($index + 1) / $groupsCount"

        Start-Sleep -Milliseconds 500

        $group = $groups[$index]

        $groupManagersList = $group.ManagedBy

        $areManagersInBothCompanies = $false
        $areMembersInBothCompanies = $false

        $groupManagers = $null
        $groupMembers = $null

        $groupManagerProperties = $null
        $groupMemberProperties = $null

        $groupManagersCount = 0
        $groupMembersCount = 0

        $firstCompanyManagersCount = 0
        $firstCompanyMembersCount = 0

        $secondCompanyManagersCount = 0
        $secondCompanyMembersCount = 0

        if ($groupManagersList) {
            $groupManagers = Get-ManagersFromList $groupManagersList

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

            $groupUsersCount = $groupManagersCount

            $firstCompanyUsersCount = $firstCompanyManagersCount
            $secondCompanyUsersCount = $secondCompanyManagersCount
        } else {
            $groupMembers = Get-MembersFromGroup $group

            if ($groupMembers) {
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

                $groupUsersCount = $groupMembersCount

                $firstCompanyUsersCount = $firstCompanyMembersCount
                $secondCompanyUsersCount = $secondCompanyMembersCount
            } else {
                $groupUsersCount = 0

                $firstCompanyUsersCount = 0
                $secondCompanyUsersCount = 0
            }
        }

        if ($groupManagersList) {
            if (-not $groupManagers) {
                $groupManagers = Get-ManagersFromList $groupManagersList
            }

            $groupManagersSMTPAddresses = (
                $groupManagers | Select-Object -ExpandProperty PrimarySmtpAddress
            ) -join ';'
            $groupManagersCompanies = (
                $groupManagers | Select-Object -ExpandProperty Company
            ) -join ';'
        } else {
            $groupManagersSMTPAddresses = ''
            $groupManagersCompanies = ''
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
            if (-not $groupMembers) {
                $groupMembers = Get-MembersFromGroup $group
            }

            if ($groupMembers) {
                $groupMembersSMTPAddresses = $groupMembers.PrimarySmtpAddress -join ';'
            } else {
                $groupMembersSMTPAddresses = ''
            }

            [PSCustomObject]@{
                'Group' = $group

                'Group Name' = $group.Name
                'Group GUID' = $group.Guid
                'Group SMTP' = $group.WindowsEmailAddress
                'Group Category' = $group.RecipientType
                'Group Company' = $groupCompany

                'Group Manager Properties' = $groupManagerProperties -join ';'
                'Group Member Properties' = $groupMemberProperties -join ';'
                'Group Managers Count' = $groupManagersCount
                'Group Members Count' = $groupMembersCount

                'First Company Managers Count' = $firstCompanyManagersCount
                'First Company Members Count' = $firstCompanyMembersCount
                'Second Company Managers Count' = $secondCompanyManagersCount
                'Second Company Members Count' = $secondCompanyMembersCount

                'Group Managers SMTP Addresses' = $groupManagersSMTPAddresses
                'Group Managers Companies' = $groupManagersCompanies

                'Group Members SMTP Addresses' = $groupMembersSMTPAddresses
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