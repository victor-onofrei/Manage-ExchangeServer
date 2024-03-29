begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParam $args"
}

process {
    Start-Transcript "$($params.outputFilePath).txt"

    Write-Output 'To gather all recipient information'
    $allRecipients = Get-Recipient -ResultSize Unlimited |
        Where-Object { $_.RecipientType -match 'User' }

    $domains = Get-AcceptedDomain
    $domainsCount = $domains | Measure-Object |
        Select-Object -ExpandProperty Count
    Write-Output "To process $domainsCount domains"

    for ($index = 0; $index -lt $domainsCount; $index++) {
        $domain = $domains[$index]
        Write-Output (
            "Processing domain $($index + 1) / $domainsCount | $domain"
        )
        $matchMailboxes = $null
        $matchMailboxes = $allRecipients |
            Where-Object { $_.EmailAddresses -match $domain }
        $matchMailboxesAlias = $matchMailboxes |
            Where-Object { $_.PrimarySmtpAddress -notmatch $domain -and
                $_.EmailAddresses -match $domain }
        $matchMailboxesAliasCount = $matchMailboxesAlias |
            Measure-Object | Select-Object -ExpandProperty Count
        $matchMailboxesCount = $matchMailboxes | Measure-Object |
            Select-Object -ExpandProperty Count
        $matchMailboxesEXP = 0
        $matchMailboxesEXO = 0
        Write-Output "To process $matchMailboxesCount exchange objects"
        for ($jindex = 0; $jindex -lt $matchMailboxesCount; $jindex++) {
            $matchMailbox = $matchMailboxes[$jindex]
            Write-Output (
                "Processing domain $($index + 1) / $domainsCount | $domain |",
                "Processing exchange object $($jindex + 1) / $matchMailboxesCount | $matchMailbox"
            )
            $location = Get-ExchangeObjectLocation -ExchangeObject $matchMailbox
            if ($location -eq 'exchangeOnPremises') {
                $matchMailboxesEXP++
            } elseif ($location -eq 'exchangeOnline') {
                $matchMailboxesEXO++
            }
        }

        [PSCustomObject]@{
            Domain = $domain
            MailboxesCount = $matchMailboxesCount
            EXPMailboxes = $matchMailboxesEXP
            EXOMailboxes = $matchMailboxesEXO
            Aliases = $matchMailboxesAliasCount
        } | Export-Csv $params.outputFilePath -Append -NoTypeInformation
    }

    Stop-Transcript
}

end {
    Send-DefaultReportMail -ScriptParams $params
}