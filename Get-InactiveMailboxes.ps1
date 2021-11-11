param (
    [Alias('IT')][Int]$InactiveThreshold
)

begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParam $args"

    $inactiveThresholdParams = @{
        Name = 'InactiveThreshold'
        Value = $InactiveThreshold
        DefaultValue = 90
        Config = $params.config
        ScriptName = $params.scriptName
    }
    $inactiveThreshold = Read-Param @inactiveThresholdParams

    $inactiveSpan = New-TimeSpan -Days $inactiveThreshold
    $today = Get-Date
}

process {
    $allMailboxes = Get-Mailbox -ResultSize Unlimited |
        Where-Object { $_.WhenMailboxCreated -lt ($today - $inactiveSpan) }
    $allMailboxesStats = $allMailboxes |
        Get-MailboxStatistics -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    $inactiveMailboxes = $allMailboxesStats |
        Where-Object { ($_.LastLogonTime -lt ($today - $inactiveSpan)) -or
            ($_.LastLogonTime -eq $Null) }
    $inactiveMailboxesCount = $inactiveMailboxes | Measure-Object |
        Select-Object -ExpandProperty Count

    $output = @()
    Write-Output "To process $inactiveMailboxesCount inactive mailboxes"

    for ($index = 0; $index -lt $inactiveMailboxesCount; $index++) {
        $inactiveMailbox = $inactiveMailboxes[$index]
        Write-Output (
            "Processing mailbox $($index + 1) / $inactiveMailboxesCount | " +
            "$($inactiveMailbox.DisplayName)"
        )

        $mailbox = ($allMailboxes |
                Where-Object { $_.DisplayName -eq $inactiveMailbox.DisplayName })

        if ($Null -eq $inactiveMailbox.LastLogonTime) {
            $inactiveSpanInactive = ($today - $mailbox.WhenMailboxCreated).Days
        } else {
            $inactiveSpanInactive = ($today - $inactiveMailbox.LastLogonTime).Days
        }

        $instance = ($inactiveMailbox |
                Select-Object @{
                    Label = 'Displayname';
                    Expression = { $mailbox.DisplayName }
                },
                @{
                    Label = 'UserPrincipalName';
                    Expression = { $mailbox.UserPrincipalName }
                },
                @{
                    Label = 'PrimarySmtpAddress';
                    Expression = { $mailbox.PrimarySmtpAddress }
                },
                @{
                    Label = 'MailboxType';
                    Expression = { $mailbox.RecipientTypeDetails }
                },
                @{
                    Label = 'MailboxCreatedon';
                    Expression = { $mailbox.WhenMailboxCreated }
                },
                @{
                    Label = 'Lastloggedon';
                    Expression = { $Inactivemailbox.LastLogonTime }
                },
                @{
                    Label = 'Inactive';
                    Expression = { $inactiveSpanInactive }
                }
        )
        $output += $instance
    }
}

end {
    $output | Export-Csv $params.outputFilePath -Encoding UTF8 -NoTypeInformation

    Send-DefaultReportMail -ScriptParams $params
}