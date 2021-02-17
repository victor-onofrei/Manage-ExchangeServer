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

    Start-Transcript "$outputDir\$projectName\$outputFileName.txt"

    $Groups = Get-Group -ResultSize Unlimited -Filter "RecipientTypeDetails -eq 'GroupMailbox'" | Select-Object WindowsEmailAddress, ManagedBy, Name, RecipientType, GUID # | ? {$_.RecipientTypeDetails -eq "GroupMailbox"}
    $GroupsCount = @($Groups).Count
    Write-Host "To process:" $GroupsCount "groups"

    Add-Content $outputFilePath Group">"Groupname">"GroupGUID">"Group_SMTP">"Groupcategory">"Group_company">"Group_Members_CA8">"DLcountORDLManagerscount">"compBcountORcompBManagerscount">"compAcountORcompAManagerscount">"Group_ManagedBy_SMTP">"Group_ManagedBy_Company">"Group_ManagedBy_CA8">"GroupMembersEmai


    for ($index = 0; $index -lt $GroupsCount; $index++) {
        Write-Host "`tProcessing group: " ($index + 1) "/ $GroupsCount"
        Start-Sleep -milliseconds 500
        $connectionstatus = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange"}
        if ($connectionstatus.State -ne "Opened") {
            exo
        }

        $Group = $Groups[$index]
        $Group_SMTP = $Group.WindowsEmailAddress

        $DLManagers = $Group.ManagedBy | Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue | Select-Object CustomAttribute8, PrimarySMTPAddress, Company

        $GroupMembers = Get-Group -Identity $Group_SMTP -ErrorAction SilentlyContinue | Select -ExpandProperty Members | Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue | Select-Object CustomAttribute8, PrimarySMTPAddress

        if ($DLManagers) {
            $DLManagerscount = ($DLManagers | Measure-Object).count
            $RecipientUserProperties = $DLManagers.CustomAttribute8
            $compBManagerscount = ($RecipientUserProperties | ? {$_ -like "CAB*"} | Measure-Object).count
            $compAManagerscount = ($RecipientUserProperties | ? {$_ -like "CAA*"} | Measure-Object).count
            $RecipientUserProperties = $RecipientUserProperties -join ";"
            $Group_ManagedBy_CA8 = $RecipientUserProperties
            $DLcountORDLManagerscount = $DLManagerscount
            $compBcountORcompBManagerscount = $compBManagerscount
            $compAcountORcompAManagerscount = $compAManagerscount

            $Group_ManagedBy_SMTP = @()
            $Group_ManagedBy_Company = @()
            Foreach ($Manager in $DLManagers) {
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
            $DLcount = ($GroupMembers | Measure-Object).count
            $ADUserProperties = $GroupMembers.CustomAttribute8
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
            $Groupname = $Group.Name
            $Groupcategory = $Group.RecipientType
            $GroupGUID = $Group.GUID

            $GroupMembersEmail = $GroupMembers.PrimarySMTPAddress
            $GroupMembersEmail = $GroupMembersEmail -join ";"

            Add-Content $outputFilePath $Group">"$Groupname">"$GroupGUID">"$Group_SMTP">"$Groupcategory">"$Group_company">"$ADUserProperties">"$DLcountORDLManagerscount">"$compBcountORcompBManagerscount">"$compAcountORcompAManagerscount">"$Group_ManagedBy_SMTP">"$Group_ManagedBy_Company">"$Group_ManagedBy_CA8">"$GroupMembersEmail
        }
        $compAcount = $null
        $compBcount = $null
        $compAManagerscount = $null
        $compBManagerscount = $null
        $GroupMembers = $null
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
