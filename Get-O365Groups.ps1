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
        "DLcountORDLManagerscount>compBcountORcompBManagerscount>compAcountORcompAManagerscount>",
        "Group_ManagedBy_SMTP>Group_ManagedBy_Company>Group_ManagedBy_CA8>GroupMembersEmail"
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
            $RecipientUserProperties = $groupManagers.CustomAttribute8
            $compBManagerscount = ($RecipientUserProperties | ? {$_ -like "CAB*"} | Measure-Object).count
            $compAManagerscount = ($RecipientUserProperties | ? {$_ -like "CAA*"} | Measure-Object).count
            $RecipientUserProperties = $RecipientUserProperties -join ";"
            $Group_ManagedBy_CA8 = $RecipientUserProperties
            $DLcountORDLManagerscount = $groupManagersCount
            $compBcountORcompBManagerscount = $compBManagerscount
            $compAcountORcompAManagerscount = $compAManagerscount

            $Group_ManagedBy_SMTP = @()
            $Group_ManagedBy_Company = @()
            Foreach ($Manager in $groupManagers) {
                $Group_ManagedBy_SMTP += $Manager | select PrimarySMTPAddress -ExpandProperty PrimarySMTPAddress
                # | % { $_.PrimarySMTPAddress.ToString() }
                $Group_ManagedBy_Company += $Manager | select Company -ExpandProperty Company
                # | % { $_.Company.ToString() }
                # $Group_ManagedBy_CA8 += $Manager | select customattribute8 -ExpandProperty customattribute8
                # | % { $_.customattribute8.ToString() }
            }
            $Group_ManagedBy_SMTP = $Group_ManagedBy_SMTP -join ";"
            $Group_ManagedBy_Company = $Group_ManagedBy_Company -join ";"
        } else {
            $DLcount = ($groupMembers | Measure-Object).count
            $ADUserProperties = $groupMembers.CustomAttribute8
            $compBcount = ($ADUserProperties | ? {$_ -like "CAB*"} | Measure-Object).count
            $compAcount = ($ADUserProperties | ? {$_ -like "CAA*"} | Measure-Object).count
            $ADUserProperties = $ADUserProperties -join ";"
            # $UserProperties = $ADUserProperties
            $DLcountORDLManagerscount = $DLcount
            $compBcountORcompBManagerscount = $compBcount
            $compAcountORcompAManagerscount = $compAcount
        }

        if ($compBcountORcompBManagerscount -eq 0 -and $DLcountORDLManagerscount -eq 0 -and $compAcountORcompAManagerscount -eq 0) {
            $Group_company = "None"
        } elseif ($compBManagerscount -and $compAManagerscount) {
            $Group_company = "Mixed Owners"
        } elseif ($compBcount -and $compAcount) {
            $Group_company = "Mixed Users"
        } elseif (($compBcountORcompBManagerscount -eq $DLcountORDLManagerscount) -or ($compBcountORcompBManagerscount -eq 0 -and $compAcountORcompAManagerscount -eq 0 -and $Group_ManagedBy_Company -match "compB" -and $Group_ManagedBy_Company -notmatch "compA") -or (($compBcountORcompBManagerscount) -and $compAcountORcompAManagerscount -eq 0)) {
            $Group_company = "compB"
        } elseif ($compAcountORcompAManagerscount -eq $DLcountORDLManagerscount -or ($compBcountORcompBManagerscount -eq 0 -and $compAcountORcompAManagerscount -eq 0 -and $Group_ManagedBy_Company -match "compA" -and $Group_ManagedBy_Company -notmatch "compB") -or (($compAcountORcompAManagerscount) -and $compBcountORcompBManagerscount -eq 0)) {
            $Group_company = "compA"
        }

        if ($Group_company -eq "compA" -or $Group_company -like "Mixed*") {
            $Groupname = $group.Name
            $Groupcategory = $group.RecipientType
            $GroupGUID = $group.GUID

            $GroupMembersEmail = $groupMembers.PrimarySMTPAddress
            $GroupMembersEmail = $GroupMembersEmail -join ";"

            Add-Content $outputFilePath $group">"$Groupname">"$GroupGUID">"$groupSMTP">"$Groupcategory">"$Group_company">"$ADUserProperties">"$DLcountORDLManagerscount">"$compBcountORcompBManagerscount">"$compAcountORcompAManagerscount">"$Group_ManagedBy_SMTP">"$Group_ManagedBy_Company">"$Group_ManagedBy_CA8">"$GroupMembersEmail
        }
        $compAcount = $null
        $compBcount = $null
        $compAManagerscount = $null
        $compBManagerscount = $null
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
