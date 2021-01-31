. "$PSScriptRoot\Initializer.ps1"
$params = Invoke-Expression "Initialize-DefaultParams $args"

$sizeFieldName = "TotalItemSizeInGB"

$defaultProhibitSendReceiveQuota = "100GB"
$defaultRecoverableItemsQuota = "30GB"
$defaultRecoverableItemsWarningQuota = "20GB"

$mailboxMinimumQuotaDifference = 1
$mailboxMinimumQuota = 2
$mailboxQuotaStep = 0.5

$archiveMinimumQuotaDifference = 1
$archiveMinimumQuota = 5
$archiveQuotaStep = 5.0

function Get-QuotaForSize {
    <#
        .SYNOPSIS
            Compute and return a quota based on the input size based on specific
            step, minimum quota and minimum difference between size and quota.

        .EXAMPLE
            Get-QuotaForSize
                -size 1.2
                -minimumDifference 1
                -minimumQuota 2
                -step 0.5

            This returns 2.5.
    #>
    param(
        [Double]$size,
        [Double]$minimumDifference,
        [Double]$minimumQuota,
        [Double]$step
    )

    $upscaledQuotaStep = [math]::Ceiling($step) / $step

    $upscaledSize = ($size + $minimumDifference) * $upscaledQuotaStep
    $downscaledQuota = [math]::Ceiling($upscaledSize) / $upscaledQuotaStep

    [math]::Max($downscaledQuota, $minimumQuota)
}

function Get-BytesFromGigaBytes {
    <#
        .SYNOPSIS
            Compute and return a quota based on the input size based on specific
            step, minimum quota and minimum difference between size and quota.

        .EXAMPLE
            Get-QuotaFromSize -gigaBytes 1.2

            This returns 2.5.
    #>
    param(
        [Double]$gigaBytes
    )

    [math]::Round($gigaBytes * [math]::Pow(2, 30))
}

foreach ($exchangeObject in $params.exchangeObjects) {
    $mailboxSizeGigaBytes =
        Get-MailboxStatistics -Identity $exchangeObject | `
        Select-Object @{
            Name = $sizeFieldName
            Expression = {
                [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1GB), 2)
            }
        } | `
        Select-Object $sizeFieldName -ExpandProperty $sizeFieldName

    $mailboxDesiredQuotaGigaBytes = Get-QuotaForSize `
        -size $mailboxSizeGigaBytes `
        -minimumDifference $mailboxMinimumQuotaDifference `
        -minimumQuota $mailboxMinimumQuota `
        -step $mailboxQuotaStep
    $mailboxDesiredQuotaBytes = Get-BytesFromGigaBytes -gigaBytes $mailboxDesiredQuotaGigaBytes

    $movingProhibitSendQuota = $mailboxDesiredQuotaBytes
    $movingIssueWarningQuota = [math]::Round($mailboxDesiredQuotaBytes * 0.9)

    Set-Mailbox $exchangeObject `
        -UseDatabaseQuotaDefaults $false `
        -ProhibitSendQuota $movingProhibitSendQuota `
        -ProhibitSendReceiveQuota $defaultProhibitSendReceiveQuota `
        -RecoverableItemsQuota $defaultRecoverableItemsQuota `
        -RecoverableItemsWarningQuota $defaultRecoverableItemsWarningQuota `
        -IssueWarningQuota $movingIssueWarningQuota

    $mailboxInfo = Get-Mailbox -Identity $exchangeObject
    $hasArchive = ($mailboxInfo.archiveGuid -ne "00000000-0000-0000-0000-000000000000") -and $mailboxInfo.archiveDatabase

    if ($hasArchive) {
        $archiveSizeGigaBytes =
            Get-MailboxStatistics -Identity $exchangeObject -Archive | `
            Select-Object @{
                Name = $sizeFieldName
                Expression = {
                    [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1GB), 2)
                }
            } | `
            Select-Object $sizeFieldName -ExpandProperty $sizeFieldName
    } else {
        $archiveSizeGigaBytes = 0
    }

    $archiveDesiredQuotaGigaBytes = Get-QuotaForSize `
        -size $archiveSizeGigaBytes `
        -minimumDifference $archiveMinimumQuotaDifference `
        -minimumQuota $archiveMinimumQuota `
        -step $archiveQuotaStep
    $archiveDesiredQuotaBytes = Get-BytesFromGigaBytes -gigaBytes $archiveDesiredQuotaGigaBytes

    $movingArchiveQuota = $archiveDesiredQuotaBytes
    $movingArchiveWarningQuota = [math]::Round($archiveDesiredQuotaBytes * 0.9)

    Set-Mailbox $exchangeObject `
        -UseDatabaseQuotaDefaults $false `
        -ArchiveQuota $movingArchiveQuota `
        -ArchiveWarningQuota $movingArchiveWarningQuota
}
