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
        Get-MailboxStatistics -Identity $exchangeObject | `
        Select-Object @{
            Name = $sizeFieldName
            Expression = {
                [Math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1GB), 2)
            }
        } | `
        Select-Object $sizeFieldName -ExpandProperty $sizeFieldName

    $mailboxDesiredQuotaGigaBytes = Get-QuotaForSize `
        -Size $mailboxSizeGigaBytes `
        -MinimumDifference $mailboxMinimumQuotaDifference `
        -MinimumQuota $mailboxMinimumQuota `
        -Step $mailboxQuotaStep
    $mailboxDesiredQuotaBytes = Get-BytesFromGigaBytes -GigaBytes $mailboxDesiredQuotaGigaBytes

    $movingProhibitSendQuota = $mailboxDesiredQuotaBytes
    $movingIssueWarningQuota = [Math]::Round($mailboxDesiredQuotaBytes * 0.9)

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
                    [Math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1GB), 2)
                }
            } | `
            Select-Object $sizeFieldName -ExpandProperty $sizeFieldName
    } else {
        $archiveSizeGigaBytes = 0
    }

    $archiveDesiredQuotaGigaBytes = Get-QuotaForSize `
        -Size $archiveSizeGigaBytes `
        -MinimumDifference $archiveMinimumQuotaDifference `
        -MinimumQuota $archiveMinimumQuota `
        -Step $archiveQuotaStep
    $archiveDesiredQuotaBytes = Get-BytesFromGigaBytes -GigaBytes $archiveDesiredQuotaGigaBytes

    $movingArchiveQuota = $archiveDesiredQuotaBytes
    $movingArchiveWarningQuota = [Math]::Round($archiveDesiredQuotaBytes * 0.9)

    Set-Mailbox $exchangeObject `
        -UseDatabaseQuotaDefaults $false `
        -ArchiveQuota $movingArchiveQuota `
        -ArchiveWarningQuota $movingArchiveWarningQuota
}
