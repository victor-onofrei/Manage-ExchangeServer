param (
    [String]$SizeFieldName = "TotalItemSizeInGB",

    [String]$DefaultProhibitSendReceiveQuota = "100GB",
    [String]$DefaultRecoverableItemsQuota = "30GB",
    [String]$DefaultRecoverableItemsWarningQuota = "20GB",

    [Int]$MailboxMinimumQuotaDifference = 1,
    [Int]$MailboxMinimumQuota = 2,
    [Float]$MailboxQuotaStep = 0.5,

    [Int]$ArchiveMinimumQuotaDifference = 1,
    [Int]$ArchiveMinimumQuota = 5,
    [Float]$ArchiveQuotaStep = 5.0
)

begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"
}

process {
    function Get-QuotaForSize {
        <#
            .SYNOPSIS
                Compute and return a quota based on the input size based on specific
                step, minimum quota and minimum difference between size and quota.

            .EXAMPLE
                Get-QuotaForSize
                    -Size 1.2
                    -MinimumDifference 1
                    -MinimumQuota 2
                    -Step 0.5

                This returns 2.5.
        #>
        param (
            [Double]$Size,
            [Double]$MinimumDifference,
            [Double]$MinimumQuota,
            [Double]$Step
        )

        $upscaledQuotaStep = [Math]::Ceiling($Step) / $Step

        $upscaledSize = ($Size + $MinimumDifference) * $upscaledQuotaStep
        $downscaledQuota = [Math]::Ceiling($upscaledSize) / $upscaledQuotaStep

        [Math]::Max($downscaledQuota, $MinimumQuota)
    }

    function Get-BytesFromGigaBytes {
        <#
            .SYNOPSIS
                Compute and return a quota based on the input size based on specific
                step, minimum quota and minimum difference between size and quota.

            .EXAMPLE
                Get-QuotaFromSize -GigaBytes 1.2

                This returns 2.5.
        #>
        param (
            [Double]$GigaBytes
        )

        [Math]::Round($GigaBytes * [Math]::Pow(2, 30))
    }

    foreach ($exchangeObject in $params.exchangeObjects) {
        $mailboxSizeGigaBytes =
            Get-MailboxStatistics -Identity $exchangeObject |
            Select-Object @{
                Name = $SizeFieldName
                Expression = {
                    [Math]::Round(
                        (
                            $_.
                                TotalItemSize.
                                ToString().
                                Split("(")[1].
                                Split(" ")[0].
                                Replace(",", "") / 1GB
                        ),
                        2
                    )
                }
            } |
            Select-Object $SizeFieldName -ExpandProperty $SizeFieldName

        $mailboxDesiredQuotaGigaBytes = Get-QuotaForSize `
            -Size $mailboxSizeGigaBytes `
            -MinimumDifference $MailboxMinimumQuotaDifference `
            -MinimumQuota $MailboxMinimumQuota `
            -Step $MailboxQuotaStep
        $mailboxDesiredQuotaBytes = Get-BytesFromGigaBytes -GigaBytes $mailboxDesiredQuotaGigaBytes

        $movingProhibitSendQuota = $mailboxDesiredQuotaBytes
        $movingIssueWarningQuota = [Math]::Round($mailboxDesiredQuotaBytes * 0.9)

        Set-Mailbox $exchangeObject `
            -UseDatabaseQuotaDefaults $false `
            -ProhibitSendQuota $movingProhibitSendQuota `
            -ProhibitSendReceiveQuota $DefaultProhibitSendReceiveQuota `
            -RecoverableItemsQuota $DefaultRecoverableItemsQuota `
            -RecoverableItemsWarningQuota $DefaultRecoverableItemsWarningQuota `
            -IssueWarningQuota $movingIssueWarningQuota

        $mailboxInfo = Get-Mailbox -Identity $exchangeObject
        $hasArchiveGuid = $mailboxInfo.archiveGuid -ne "00000000-0000-0000-0000-000000000000"
        $hasArchive = $hasArchiveGuid -and $mailboxInfo.archiveDatabase

        if ($hasArchive) {
            $archiveSizeGigaBytes =
                Get-MailboxStatistics -Identity $exchangeObject -Archive |
                Select-Object @{
                    Name = $SizeFieldName
                    Expression = {
                        [Math]::Round(
                            (
                                $_.
                                    TotalItemSize.
                                    ToString().
                                    Split("(")[1].
                                    Split(" ")[0].
                                    Replace(",", "") / 1GB
                            ),
                            2
                        )
                    }
                } |
                Select-Object $SizeFieldName -ExpandProperty $SizeFieldName
        } else {
            $archiveSizeGigaBytes = 0
        }

        $archiveDesiredQuotaGigaBytes = Get-QuotaForSize `
            -Size $archiveSizeGigaBytes `
            -MinimumDifference $ArchiveMinimumQuotaDifference `
            -MinimumQuota $ArchiveMinimumQuota `
            -Step $ArchiveQuotaStep
        $archiveDesiredQuotaBytes = Get-BytesFromGigaBytes -GigaBytes $archiveDesiredQuotaGigaBytes

        $movingArchiveQuota = $archiveDesiredQuotaBytes
        $movingArchiveWarningQuota = [Math]::Round($archiveDesiredQuotaBytes * 0.9)

        Set-Mailbox $exchangeObject `
            -UseDatabaseQuotaDefaults $false `
            -ArchiveQuota $movingArchiveQuota `
            -ArchiveWarningQuota $movingArchiveWarningQuota
    }
}
