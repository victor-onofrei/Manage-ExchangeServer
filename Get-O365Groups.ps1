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
        "Group>Groupname>GroupGUID>Group SMTP>Groupcategory>Group_company>Group_Members_CA8>",
        "Group Members Or Managers Count>First Company Members Or Managers Count>",
        "Second Company Members Or Managers Count>Groups Managed By SMTP>",
        "Groups Managed By Company>Manager Custom Attribute 8>GroupMembersEmail"
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
            $groupManagersCount = ($groupManagers | Measure-Object).count
            $groupManagerProperties = $groupManagers.CustomAttribute8

            $firstCompanyManagersCount = (
                $groupManagerProperties | Where-Object { $_ -like "CAA*" } | Measure-Object
            ).count
            $secondCompanyManagersCount = (
                $groupManagerProperties | Where-Object { $_ -like "CAB*" } | Measure-Object
            ).count

            $groupManagerProperties = $groupManagerProperties -join ";"
            $managerCustomAttribute8 = $groupManagerProperties
            $groupMembersOrManagersCount = $groupManagersCount

            $firstCompanyMembersOrManagersCount = $firstCompanyManagersCount
            $secondCompanyMembersOrManagersCount = $secondCompanyManagersCount

            $groupsManagedBySMTP = @()
            $groupsManagedByCompany = @()

            Foreach ($Manager in $groupManagers) {
                $groupsManagedBySMTP += $Manager | select PrimarySMTPAddress -ExpandProperty PrimarySMTPAddress
                # | % { $_.PrimarySMTPAddress.ToString() }
                $groupsManagedByCompany += $Manager | select Company -ExpandProperty Company
                # | % { $_.Company.ToString() }
                # $managerCustomAttribute8 += $Manager | select customattribute8 -ExpandProperty customattribute8
                # | % { $_.customattribute8.ToString() }
            }
            $groupsManagedBySMTP = $groupsManagedBySMTP -join ";"
            $groupsManagedByCompany = $groupsManagedByCompany -join ";"
        } else {
            $DLcount = ($groupMembers | Measure-Object).count
            $ADUserProperties = $groupMembers.CustomAttribute8
            $compBcount = ($ADUserProperties | ? {$_ -like "CAB*"} | Measure-Object).count
            $compAcount = ($ADUserProperties | ? {$_ -like "CAA*"} | Measure-Object).count
            $ADUserProperties = $ADUserProperties -join ";"
            # $UserProperties = $ADUserProperties
            $groupMembersOrManagersCount = $DLcount
            $secondCompanyMembersOrManagersCount = $compBcount
            $firstCompanyMembersOrManagersCount = $compAcount
        }

        if ($secondCompanyMembersOrManagersCount -eq 0 -and $groupMembersOrManagersCount -eq 0 -and $firstCompanyMembersOrManagersCount -eq 0) {
            $Group_company = "None"
        } elseif ($secondCompanyManagersCount -and $firstCompanyManagersCount) {
            $Group_company = "Mixed Owners"
        } elseif ($compBcount -and $compAcount) {
            $Group_company = "Mixed Users"
        } elseif (($secondCompanyMembersOrManagersCount -eq $groupMembersOrManagersCount) -or ($secondCompanyMembersOrManagersCount -eq 0 -and $firstCompanyMembersOrManagersCount -eq 0 -and $groupsManagedByCompany -match "compB" -and $groupsManagedByCompany -notmatch "compA") -or (($secondCompanyMembersOrManagersCount) -and $firstCompanyMembersOrManagersCount -eq 0)) {
            $Group_company = "compB"
        } elseif ($firstCompanyMembersOrManagersCount -eq $groupMembersOrManagersCount -or ($secondCompanyMembersOrManagersCount -eq 0 -and $firstCompanyMembersOrManagersCount -eq 0 -and $groupsManagedByCompany -match "compA" -and $groupsManagedByCompany -notmatch "compB") -or (($firstCompanyMembersOrManagersCount) -and $secondCompanyMembersOrManagersCount -eq 0)) {
            $Group_company = "compA"
        }

        if ($Group_company -eq "compA" -or $Group_company -like "Mixed*") {
            $Groupname = $group.Name
            $Groupcategory = $group.RecipientType
            $GroupGUID = $group.GUID

            $GroupMembersEmail = $groupMembers.PrimarySMTPAddress
            $GroupMembersEmail = $GroupMembersEmail -join ";"

            Add-Content $outputFilePath $group">"$Groupname">"$GroupGUID">"$groupSMTP">"$Groupcategory">"$Group_company">"$ADUserProperties">"$groupMembersOrManagersCount">"$firstCompanyMembersOrManagersCount">"$secondCompanyMembersOrManagersCount">"$groupsManagedBySMTP">"$groupsManagedByCompany">"$managerCustomAttribute8">"$GroupMembersEmail
        }
        $compAcount = $null
        $compBcount = $null
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
