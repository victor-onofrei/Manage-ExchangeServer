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
            Select-Object CustomAttribute8, PrimarySMTPAddress, Company

        $groupMembers = Get-Group -Identity $groupSMTP -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty Members |
            Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue |
            Select-Object CustomAttribute8, PrimarySMTPAddress

        if ($groupManagers) {
            $groupManagersCount = ($groupManagers | Measure-Object).Count
            $groupManagerProperties = $groupManagers.CustomAttribute8

            $firstCompanyManagersCount = (
                $groupManagerProperties | Where-Object { $_ -like 'CAA*' } | Measure-Object
            ).Count
            $secondCompanyManagersCount = (
                $groupManagerProperties | Where-Object { $_ -like 'CAB*' } | Measure-Object
            ).Count

            $groupManagerProperties = $groupManagerProperties -join ';'
            $groupManagerOrMemberCustomAttribute8 = $groupManagerProperties
            $groupManagersOrMembersCount = $groupManagersCount

            $firstCompanyManagersOrMembersCount = $firstCompanyManagersCount
            $secondCompanyManagersOrMembersCount = $secondCompanyManagersCount

            $groupsManagedBySMTP = @()
            $groupsManagedByCompany = @()

            foreach ($manager in $groupManagers) {
                $groupsManagedBySMTP += $manager |
                    Select-Object PrimarySMTPAddress -ExpandProperty PrimarySMTPAddress

                $groupsManagedByCompany += $manager | Select-Object Company -ExpandProperty Company
            }

            $groupsManagedBySMTP = $groupsManagedBySMTP -join ';'
            $groupsManagedByCompany = $groupsManagedByCompany -join ';'
        } else {
            $groupMembersCount = ($groupMembers | Measure-Object).Count
            $groupMemberProperties = $groupMembers.CustomAttribute8

            $firstCompanyMembersCount = (
                $groupMemberProperties | Where-Object { $_ -like 'CAA*' } | Measure-Object
            ).Count
            $secondCompanyMembersCount = (
                $groupMemberProperties | Where-Object { $_ -like 'CAB*' } | Measure-Object
            ).Count

            $groupMemberProperties = $groupMemberProperties -join ';'
            $groupManagerOrMemberCustomAttribute8 = $groupMemberProperties
            $groupManagersOrMembersCount = $groupMembersCount

            $firstCompanyManagersOrMembersCount = $firstCompanyMembersCount
            $secondCompanyManagersOrMembersCount = $secondCompanyMembersCount

            $groupsManagedBySMTP = ''
            $groupsManagedByCompany = ''
        }

        if (
            $secondCompanyManagersOrMembersCount -eq 0 -and
            $groupManagersOrMembersCount -eq 0 -and
            $firstCompanyManagersOrMembersCount -eq 0
        ) {
            $groupCompany = 'None'
        } elseif ($secondCompanyManagersCount -and $firstCompanyManagersCount) {
            $groupCompany = 'Mixed Owners'
        } elseif ($secondCompanyMembersCount -and $firstCompanyMembersCount) {
            $groupCompany = 'Mixed Users'
        } elseif (
            $secondCompanyManagersOrMembersCount -eq $groupManagersOrMembersCount -or (
                $secondCompanyManagersOrMembersCount -eq 0 -and
                $firstCompanyManagersOrMembersCount -eq 0 -and
                $groupsManagedByCompany -match 'compB' -and
                $groupsManagedByCompany -notmatch 'compA'
            ) -or (
                $secondCompanyManagersOrMembersCount -and
                $firstCompanyManagersOrMembersCount -eq 0
            )
        ) {
            $groupCompany = 'compB'
        } elseif (
            $firstCompanyManagersOrMembersCount -eq $groupManagersOrMembersCount -or (
                $secondCompanyManagersOrMembersCount -eq 0 -and
                $firstCompanyManagersOrMembersCount -eq 0 -and
                $groupsManagedByCompany -match 'compA' -and
                $groupsManagedByCompany -notmatch 'compB'
            ) -or (
                $firstCompanyManagersOrMembersCount -and
                $secondCompanyManagersOrMembersCount -eq 0
            )
        ) {
            $groupCompany = 'compA'
        } else {
            $groupCompany = 'N/A'
        }

        if ($groupCompany -eq 'compA' -or $groupCompany -like 'Mixed*') {
            $groupName = $group.Name
            $groupCategory = $group.RecipientType
            $groupGUID = $group.GUID

            $groupMembersEmails = $groupMembers.PrimarySMTPAddress
            $groupMembersEmails = $groupMembersEmails -join ';'

            [PSCustomObject]@{
                'Group' = $group
                'Group Name' = $groupName
                'Group GUID' = $groupGUID
                'Group SMTP' = $groupSMTP
                'Group Category' = $groupCategory
                'Group Company' = $groupCompany
                'Group Manager or Member Custom Attribute 8' = $groupManagerOrMemberCustomAttribute8
                'Group Managers or Members Count' = $groupManagersOrMembersCount
                'First Company Managers Or Members Count' = $firstCompanyManagersOrMembersCount
                'Second Company Managers Or Members Count' = $secondCompanyManagersOrMembersCount
                'Groups Managed By SMTP' = $groupsManagedBySMTP
                'Groups Managed By Company' = $groupsManagedByCompany
                'Group Members Emails' = $groupMembersEmails
            } | Export-Csv $params.outputFilePath -Append
        }

        $groupMembers = $null

        $firstCompanyManagersCount = $null
        $secondCompanyManagersCount = $null

        $firstCompanyMembersCount = $null
        $secondCompanyMembersCount = $null
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
