begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParam $args"
}

process {
    Get-Mailbox -ResultSize Unlimited |
        Get-MailboxStatistics |
        Select-Object -ExpandProperty TotalItemSize |
        Select-Object -ExpandProperty Value |
        Select-Object @{
            Name = 'TotalItemSizeinMB';
            Expression = {
                [math]::Round(
                    ($_.ToString().
                    Split('(')[1].
                    Split(' ')[0].
                    Replace(',', '') / 1MB)
                )
            }
        } |
        Select-Object -ExpandProperty TotalItemSizeinMB |
        Measure-Object -Average -Sum -Minimum -Maximum |
        Export-Csv $params.outputFilePath -NoTypeInformation
}

end {
    Send-DefaultReportMail -ScriptParams $params
}