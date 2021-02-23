using namespace System.Management.Automation

process {
    $timestamp = Get-Date -Format "yyyyMMdd_hhmmss"
    $outputDir = "\\\Download\generic\outputs"
    $projectName = "migration"
    $outputFileName = "O365_groups.$timestamp.xls"

    $newOutputDirectoryParams = @{
        Path = $outputDir
        Name = $projectName
        ItemType = "directory"
        ErrorAction = [ActionPreference]::SilentlyContinue
    }

    New-Item @newOutputDirectoryParams

    $outputFilePathParams = @{
        Path = $outputDir
        ChildPath = $projectName
        AdditionalChildPath = $outputFileName
    }

    $outputFilePath = Join-Path @outputFilePathParams

    Start-Transcript "$outputFilePath.txt"

    $groups = Get-Group -ResultSize Unlimited -Filter "RecipientTypeDetails -eq 'GroupMailbox'" |
        Select-Object WindowsEmailAddress, ManagedBy, Name, RecipientType, GUID
    $groupsCount = @($groups).Count

    $header = -join (
        "Group>Group Name>Group GUID>Group SMTP>Group Category>Group Company>",
        "Group Member Properties>Group Members Or Managers Count>",
        "First Company Members Or Managers Count>Second Company Members Or Managers Count>",
        "Groups Managed By SMTP>Groups Managed By Company>Manager Custom Attribute 8>",
        "Group Members Emails"
    )

    $header >> $outputFilePath

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
                $groupManagerProperties | Where-Object { $_ -like "CAA*" } | Measure-Object
            ).Count
            $secondCompanyManagersCount = (
                $groupManagerProperties | Where-Object { $_ -like "CAB*" } | Measure-Object
            ).Count

            $groupManagerProperties = $groupManagerProperties -join ";"
            $managerCustomAttribute8 = $groupManagerProperties
            $groupMembersOrManagersCount = $groupManagersCount

            $firstCompanyMembersOrManagersCount = $firstCompanyManagersCount
            $secondCompanyMembersOrManagersCount = $secondCompanyManagersCount

            $groupsManagedBySMTP = @()
            $groupsManagedByCompany = @()

            foreach ($manager in $groupManagers) {
                $groupsManagedBySMTP += $manager |
                    Select-Object PrimarySMTPAddress -ExpandProperty PrimarySMTPAddress

                $groupsManagedByCompany += $manager | Select-Object Company -ExpandProperty Company
            }

            $groupsManagedBySMTP = $groupsManagedBySMTP -join ";"
            $groupsManagedByCompany = $groupsManagedByCompany -join ";"
        } else {
            $groupMembersCount = ($groupMembers | Measure-Object).Count
            $groupMemberProperties = $groupMembers.CustomAttribute8

            $firstCompanyMembersCount = (
                $groupMemberProperties | Where-Object { $_ -like "CAA*" } | Measure-Object
            ).Count
            $secondCompanyMembersCount = (
                $groupMemberProperties | Where-Object { $_ -like "CAB*" } | Measure-Object
            ).Count

            $groupMemberProperties = $groupMemberProperties -join ";"
            $groupMembersOrManagersCount = $groupMembersCount

            $firstCompanyMembersOrManagersCount = $firstCompanyMembersCount
            $secondCompanyMembersOrManagersCount = $secondCompanyMembersCount
        }

        if (
            $secondCompanyMembersOrManagersCount -eq 0 -and
            $groupMembersOrManagersCount -eq 0 -and
            $firstCompanyMembersOrManagersCount -eq 0
        ) {
            $groupCompany = "None"
        } elseif ($secondCompanyManagersCount -and $firstCompanyManagersCount) {
            $groupCompany = "Mixed Owners"
        } elseif ($secondCompanyMembersCount -and $firstCompanyMembersCount) {
            $groupCompany = "Mixed Users"
        } elseif (
            $secondCompanyMembersOrManagersCount -eq $groupMembersOrManagersCount -or (
                $secondCompanyMembersOrManagersCount -eq 0 -and
                $firstCompanyMembersOrManagersCount -eq 0 -and
                $groupsManagedByCompany -match "compB" -and
                $groupsManagedByCompany -notmatch "compA"
            ) -or (
                $secondCompanyMembersOrManagersCount -and
                $firstCompanyMembersOrManagersCount -eq 0
            )
        ) {
            $groupCompany = "compB"
        } elseif (
            $firstCompanyMembersOrManagersCount -eq $groupMembersOrManagersCount -or (
                $secondCompanyMembersOrManagersCount -eq 0 -and
                $firstCompanyMembersOrManagersCount -eq 0 -and
                $groupsManagedByCompany -match "compA" -and
                $groupsManagedByCompany -notmatch "compB"
            ) -or (
                $firstCompanyMembersOrManagersCount -and
                $secondCompanyMembersOrManagersCount -eq 0
            )
        ) {
            $groupCompany = "compA"
        }

        if ($groupCompany -eq "compA" -or $groupCompany -like "Mixed*") {
            $groupName = $group.Name
            $groupCategory = $group.RecipientType
            $groupGUID = $group.GUID

            $groupMembersEmails = $groupMembers.PrimarySMTPAddress
            $groupMembersEmails = $groupMembersEmails -join ";"

            $row = -join (
                "$group>$groupName>$groupGUID>$groupSMTP>$groupCategory>$groupCompany>",
                "$groupMemberProperties>$groupMembersOrManagersCount>",
                "$firstCompanyMembersOrManagersCount>$secondCompanyMembersOrManagersCount>",
                "$groupsManagedBySMTP>$groupsManagedByCompany>$managerCustomAttribute8>",
                "$groupMembersEmails"
            )

            $row >> $outputFilePath
        }

        $groupMembers = $null

        $firstCompanyManagersCount = $null
        $secondCompanyManagersCount = $null

        $firstCompanyMembersCount = $null
        $secondCompanyMembersCount = $null
    }

    $attachment = New-Object Net.Mail.Attachment($outputFilePath)

    $message = New-Object Net.Mail.MailMessage
    $message.From = "noreply_group_details@compA.com"
    $message.Cc.Add("user1@compA.com")
    $message.To.Add("user2@compA.com")
    $message.Subject = "$($outputFileName) report is ready"
    $message.Body = "Attached is the $($outputFileName) report"
    $message.Attachments.Add($attachment)

    $smtpServer = "smtp.compB.com"
    $smtp = New-Object Net.Mail.SmtpClient($smtpServer)
    $smtp.Send($message)

    Stop-Transcript
}
