param (
    [String]$inputPath = "$env:homeshare\VDI-UserData\Download\generic\inputs\",
    [String]$fileName = "pending_migrations.csv",
    [String]$outputPath = "$env:homeshare\VDI-UserData\Download\generic\outputs\pending_migrations",
    [String]$user = $null
)

if ($user) {
    $allMailboxes = @($user)
} else {
    $allMailboxes = Get-Content "$inputPath\$fileName"
}

$routingAddress = (
    Get-OrganizationConfig | `
    Select-Object -ExpandProperty MicrosoftExchangeRecipientEmailAddresses | `
    Where-Object {$_ -like "*mail.onmicrosoft.com"}
).Split("@")[1]

foreach ($mailbox in $allMailboxes) {
    $itemCount = Get-MailboxStatistics $mailbox | Select-Object -ExpandProperty ItemCount
    if ($itemCount -le "300") {
        $mailboxInfo = Get-Mailbox -Identity $mailbox
        $mailboxInfo.EmailAddresses > $outputPath\$mailbox.txt
        $hasArchive = ($mailboxInfo.archiveGuid -ne "00000000-0000-0000-0000-000000000000") -and $mailboxInfo.archiveDatabase
        if (!$hasArchive) {
            Disable-Mailbox -Identity $mailboxInfo.Alias -Confirm:$false
            Enable-RemoteMailbox $mailboxInfo.Alias -RemoteRoutingAddress "$mailbox@$routingAddress"
            Set-RemoteMailbox $mailboxInfo.UserPrincipalName -EmailAddresses $mailboxInfo.EmailAddresses `
             -EmailAddressPolicyEnabled $false
        } else {
            [System.String]$message = "Mailbox $mailbox has on-premise archive enabled. Processing the script requires permanently " +
                "disabling the archive which will result in data loss. Consider merging the 2 mailbox objects manually after backup. Skipping this mailbox..."
            [System.Management.Automation.PSInvalidCastException]$exception = New-Object `
                -TypeName System.Management.Automation.PSInvalidCastException `
                -ArgumentList $message
            [System.Management.Automation.ErrorRecord]$errorRecord = New-Object `
                -TypeName System.Management.Automation.ErrorRecord `
                -ArgumentList `
                    $exception,
                    'ArchiveEnabled',
                    ([System.Management.Automation.ErrorCategory]::PermissionDenied),
                    $mailbox
            Write-Error -ErrorRecord $errorRecord
        }
    } else {
        [System.String]$message = "Mailbox $mailbox has on-premise content above threshold. Processing the script requires permanently " +
            "disabling the mailbox which will result in data loss. Consider merging the 2 mailbox objects manually after backup. Skipping this mailbox..."
        [System.Management.Automation.PSInvalidCastException]$exception = New-Object `
            -TypeName System.Management.Automation.PSInvalidCastException `
            -ArgumentList $message
        [System.Management.Automation.ErrorRecord]$errorRecord = New-Object `
            -TypeName System.Management.Automation.ErrorRecord `
            -ArgumentList `
                $exception,
                'ContentAboveThreshold',
                ([System.Management.Automation.ErrorCategory]::PermissionDenied),
                $mailbox
        Write-Error -ErrorRecord $errorRecord
    }
}
