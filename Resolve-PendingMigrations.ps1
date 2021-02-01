param(
    [Switch]$BypassItemCount
)

. "$PSScriptRoot\Initializer.ps1"
$params = Invoke-Expression "Initialize-DefaultParams $args"

$routingAddress = (
    Get-OrganizationConfig | `
    Select-Object -ExpandProperty MicrosoftExchangeRecipientEmailAddresses | `
    Where-Object {$_ -like "*mail.onmicrosoft.com"}
).Split("@")[1]

foreach ($exchangeObject in $params.exchangeObjects) {
    $itemCount = Get-MailboxFolderStatistics $exchangeObject | Where-Object {$_.FolderType -eq "Root"} | Select-Object -ExpandProperty ItemsInFolderAndSubfolders
    if ($itemCount -le "100" -or $BypassItemCount) {
        $mailboxInfo = Get-Mailbox -Identity $exchangeObject

        foreach ($emailAddress in $mailboxInfo.EmailAddresses) {
            "$exchangeObject,$emailAddress" >> $params.outputFilePath
        }

        $hasArchive = ($mailboxInfo.archiveGuid -ne "00000000-0000-0000-0000-000000000000") -and $mailboxInfo.archiveDatabase
        if (!$hasArchive) {
            Disable-Mailbox -Identity $mailboxInfo.Alias -Confirm:$false
            Enable-RemoteMailbox $mailboxInfo.Alias -RemoteRoutingAddress "$exchangeObject@$routingAddress"
            Set-RemoteMailbox $mailboxInfo.UserPrincipalName -EmailAddresses $mailboxInfo.EmailAddresses `
             -EmailAddressPolicyEnabled $false
        } else {
            [System.String]$message = "Mailbox $exchangeObject has on-premise archive enabled. Processing the script requires permanently " +
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
                    $exchangeObject
            Write-Error -ErrorRecord $errorRecord
        }
    } else {
        [System.String]$message = "Mailbox $exchangeObject has on-premise content above threshold. Processing the script requires permanently " +
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
                $exchangeObject
        Write-Error -ErrorRecord $errorRecord
    }
}
