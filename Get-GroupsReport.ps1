using namespace Microsoft.Exchange.Data.Directory.Management
using namespace System.Management.Automation

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
            [ADPresentationObject]$Group
        )

        $getGroupParams = @{
            Identity = $Group.Guid
            ErrorAction = [ActionPreference]::SilentlyContinue
        }

        $list = Get-Group @getGroupParams |
            Select-Object -ExpandProperty Members

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
            [ADPresentationObject]$Group
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
        ([GroupsType]::office365) {
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

            $groupUserProperties = $groupManagerProperties -join ';'
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

                $groupUserProperties = $groupMemberProperties -join ';'
                $groupUsersCount = $groupMembersCount

                $firstCompanyUsersCount = $firstCompanyMembersCount
                $secondCompanyUsersCount = $secondCompanyMembersCount
            } else {
                $groupUserProperties = ''
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
                $groupMembersEmails = $groupMembers.PrimarySmtpAddress -join ';'
            } else {
                $groupMembersEmails = ''
            }

            [PSCustomObject]@{
                'Group' = $group

                'Group Name' = $group.Name
                'Group GUID' = $group.Guid
                'Group SMTP' = $group.WindowsEmailAddress
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
