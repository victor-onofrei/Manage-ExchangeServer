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
        "Group>Group Name>GroupGUID>Group SMTP>Group Category>Group Company>",
        "Group Member Properties>Group Members Or Managers Count>",
        "First Company Members Or Managers Count>Second Company Members Or Managers Count>",
        "Groups Managed By SMTP>Groups Managed By Company>Manager Custom Attribute 8>",
        "GroupMembersEmail"
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
            $GroupGUID = $group.GUID

            $GroupMembersEmail = $groupMembers.PrimarySMTPAddress
            $GroupMembersEmail = $GroupMembersEmail -join ";"

            Add-Content $outputFilePath $group">"$groupName">"$GroupGUID">"$groupSMTP">"$groupCategory">"$groupCompany">"$groupMemberProperties">"$groupMembersOrManagersCount">"$firstCompanyMembersOrManagersCount">"$secondCompanyMembersOrManagersCount">"$groupsManagedBySMTP">"$groupsManagedByCompany">"$managerCustomAttribute8">"$GroupMembersEmail
        }
        $firstCompanyMembersCount = $null
        $secondCompanyMembersCount = $null
        $firstCompanyManagersCount = $null
        $secondCompanyManagersCount = $null
        $groupMembers = $null
    }

    $SmtpServer = "smtp.compB.com"
    $att = new-object Net.Mail.Attachment($outputFilePath)
    $msg = new-object Net.Mail.MailMessage
    $smtp = new-object Net.Mail.SmtpClient($SmtpServer)
    $msg.From = "noreply_group_details@compA.com"
    $msg.Cc.Add("user1@compA.com")
    $msg.To.Add("user2@compA.com")
    $msg.Subject = "$($outputFileName) report is ready"
    $msg.Body = "Attached is the $($outputFileName) report"
    $msg.Attachments.Add($att)
    $smtp.Send($msg)
    Stop-Transcript
}
