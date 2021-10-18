begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"
    $mailboxPool = $params.exchangeObjects

    Write-Output "Getting Exchange Online mailboxes..."
    $cloudMailboxes = $mailboxPool |
        Get-EXORecipient -RecipientType UserMailbox -ErrorAction SilentlyContinue
    $cloudMailboxesCount = @($cloudMailboxes).Count
    Write-Output "Found $cloudMailboxesCount Exchange Online mailboxes"

    Write-Output "Getting Exchange Online mailboxes total sizes..."
    $cloudMailboxesTotalSizeMB = @()
    $cloudMailboxesTotalSizeMB = $mailboxPool |
        Get-EXORecipient -Properties ExchangeGuid -ErrorAction SilentlyContinue |
        Get-EXOMailboxStatistics $_.ExchangeGuid -ErrorAction SilentlyContinue |
        Select-Object MailboxGuid, @{
            Name = "TotalItemSizeinMB";
            Expression = {
                [math]::Round(
                    ($_.TotalItemSize.ToString().
                        Split("(")[1].
                        Split(" ")[0].
                        Replace(",", "") / 1MB)
                )
            }
        }
    $cloudMailboxesTotalDeletedItemSizeMB = @()
    $cloudMailboxesTotalDeletedItemSizeMB = $mailboxPool |
        Get-EXORecipient -Properties ExchangeGuid -ErrorAction SilentlyContinue |
        Get-EXOMailboxStatistics $_.ExchangeGuid -ErrorAction SilentlyContinue |
        Select-Object MailboxGuid, @{
            Name = "TotalDeletedItemSizeinMB";
            Expression = {
                [math]::Round(
                    ($_.TotalDeletedItemSize.ToString().
                        Split("(")[1].
                        Split(" ")[0].
                        Replace(",", "") / 1MB)
                )
            }
        }

    Write-Output "Getting Exchange Online mailboxes' archives total size..."
    $cloudMailboxesArchive = $mailboxPool |
        Get-EXOMailbox -ErrorAction SilentlyContinue |
        Select-Object ExchangeGuid, ArchiveName, ArchiveGuid
    $cloudMailboxesArchiveTotalSizeMB = @()
    $cloudMailboxesArchiveTotalSizeMB = $mailboxPool |
        Get-Recipient -ErrorAction SilentlyContinue -Filter "ArchiveState -ne 'None'" |
        Get-EXOMailboxStatistics $_.ExchangeGuid -Archive -ErrorAction SilentlyContinue |
        Select-Object MailboxGuid, @{
            Name = "TotalItemSizeinMB";
            Expression = {
                [math]::Round(
                    ($_.TotalItemSize.ToString().
                        Split("(")[1].
                        Split(" ")[0].
                        Replace(",", "") / 1MB)
                )
            }
        }
    $cloudMailboxesArchiveTotalDeletedItemSizeMB = @()
    $cloudMailboxesArchiveTotalDeletedItemSizeMB = $mailboxPool |
        Get-Recipient -ErrorAction SilentlyContinue -Filter "ArchiveState -ne 'None'" |
        Get-EXOMailboxStatistics $_.ExchangeGuid -Archive -ErrorAction SilentlyContinue |
        Select-Object MailboxGuid, @{
            Name = "TotalDeletedItemSizeinMB";
            Expression = {
                [math]::Round(
                    ($_.TotalDeletedItemSize.ToString().
                        Split("(")[1].
                        Split(" ")[0].
                        Replace(",", "") / 1MB)
                )
            }
        }
}

process {
    Write-Output "To process $cloudMailboxesCount Exchange Online mailboxes"
    for ($index = 0; $index -lt $cloudMailboxesCount; $index++) {
        $cloudMailbox = $cloudMailboxes[$index]
        Write-Output (
            "`tProcessing mailbox $($index + 1) / $cloudMailboxesCount | $cloudMailbox"
        )
        $cloudMailboxAliasUID = $cloudMailbox.Alias
        $cloudMailboxSamAccountName = $cloudMailbox.SamAccountName
        $cloudMailboxPrimarySMTPAddress = $cloudMailbox.PrimarySMTPAddress
        $cloudMailboxMailDomain = $cloudMailboxPrimarySMTPAddress.ToString().split('@')[1]
        $cloudMailboxUPN = $cloudMailbox.WindowsLiveId
        $cloudMailboxType = $cloudMailbox.RecipientTypeDetails
        if ($cloudMailboxType -like "UserMailbox" -and $cloudMailboxSamAccountName -like "SRV*") {
            $cloudMailboxType = "ServiceMailbox"
        }
        $cloudMailboxDisplayName = $cloudMailbox.DisplayName
        $cloudMailboxEXOMailbox = Get-EXOMailbox $cloudMailboxAliasUID
        $cloudMailboxPlan = $cloudMailboxEXOMailbox |
            Select-Object -ExpandProperty Mailboxplan
        $cloudMailboxMaxSendSize = $cloudMailboxEXOMailbox |
            Select-Object -ExpandProperty MaxSendSize
        $cloudMailboxMaxReceiveSize = $cloudMailboxEXOMailbox |
            Select-Object -ExpandProperty MaxReceiveSize
        $cloudMailboxIssueWarningQuota = $cloudMailboxEXOMailbox |
            Select-Object -ExpandProperty IssueWarningQuota
        $cloudMailboxProhibitSendQuota = $cloudMailboxEXOMailbox |
            Select-Object -ExpandProperty ProhibitSendQuota
        $cloudMailboxProhibitSendReceiveQuota = $cloudMailboxEXOMailbox |
            Select-Object -ExpandProperty ProhibitSendReceiveQuota

        $cloudMailboxArchiveState = $cloudMailbox.ArchiveState
        $cloudMailboxSizeMB = $cloudMailboxesTotalSizeMB |
            Where-Object {
                $_.MailboxGuid -eq $cloudMailbox.ExchangeGuid
            } |
            Select-Object -ExpandProperty TotalItemSizeinMB
        $cloudMailboxTotalDeletedItemSizeMB = $cloudMailboxesTotalDeletedItemSizeMB |
            Where-Object {
                $_.MailboxGuid -eq $cloudMailbox.ExchangeGuid
            } |
            Select-Object -ExpandProperty TotalDeletedItemSizeinMB

        $cloudMailboxArchiveDisplayName = $null
        $cloudMailboxArchiveGuid = $null
        $cloudMailboxArchiveDeletedItemSizeMB = $null
        $cloudMailboxArchiveSizeMB = $null
        if ($cloudMailbox.ArchiveState -ne "None") {
            $cloudMailboxArchiveDisplayName =
            $cloudMailboxesArchive |
                Where-Object {
                    $_.ExchangeGuid -eq $cloudMailbox.ExchangeGuid
                } |
                Select-Object -ExpandProperty ArchiveName
            $cloudMailboxArchiveGuid =
            $cloudMailboxesArchive |
                Where-Object {
                    $_.ExchangeGuid -eq $cloudMailbox.ExchangeGuid
                } |
                Select-Object -ExpandProperty ArchiveGuid
            $cloudMailboxArchiveDeletedItemSizeMB =
            $cloudMailboxesArchiveTotalDeletedItemSizeMB |
                Where-Object {
                    $_.MailboxGuid -eq $cloudMailbox.ArchiveGuid
                } |
                Select-Object -ExpandProperty TotalDeletedItemSizeinMB
            $cloudMailboxArchiveSizeMB =
            $cloudMailboxesArchiveTotalSizeMB |
                Where-Object {
                    $_.MailboxGuid -eq $cloudMailbox.ArchiveGuid
                } |
                Select-Object -ExpandProperty TotalItemSizeinMB
        }

        [PSCustomObject]@{
            cloudMailboxAliasUID = $cloudMailboxAliasUID
            cloudMailboxDisplayName = $cloudMailboxDisplayName
            cloudMailboxSamAccountName = $cloudMailboxSamAccountName
            cloudMailboxUPN = $cloudMailboxUPN
            cloudMailboxPrimarySMTPAddress = $cloudMailboxPrimarySMTPAddress
            cloudMailboxMailDomain = $cloudMailboxMailDomain
            cloudMailboxType = $cloudMailboxType
            cloudMailboxPlan = $cloudMailboxPlan
            cloudMailboxArchiveState = $cloudMailboxArchiveState
            cloudMailboxMaxSendSize = $cloudMailboxMaxSendSize
            cloudMailboxMaxReceiveSize = $cloudMailboxMaxReceiveSize
            cloudMailboxIssueWarningQuota = $cloudMailboxIssueWarningQuota
            cloudMailboxProhibitSendQuota = $cloudMailboxProhibitSendQuota
            cloudMailboxProhibitSendReceiveQuota = $cloudMailboxProhibitSendReceiveQuota
            cloudMailboxTotalDeletedItemSizeMB = $cloudMailboxTotalDeletedItemSizeMB
            cloudMailboxSizeMB = $cloudMailboxSizeMB
            cloudMailboxArchiveDisplayName = $cloudMailboxArchiveDisplayName
            cloudMailboxArchiveGuid = $cloudMailboxArchiveGuid
            cloudMailboxArchiveDeletedItemSizeMB = $cloudMailboxArchiveDeletedItemSizeMB
            cloudMailboxArchiveSizeMB = $cloudMailboxArchiveSizeMB
        } | Export-Csv $params.outputFilePath -Append -NoTypeInformation
    }
}

end {
    Send-DefaultReportMail -ScriptParams $params
}