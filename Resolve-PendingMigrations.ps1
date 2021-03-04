param (
    [Alias('BIC')][Switch]$BypassItemCount,
    [Alias('ICT')][Int]$ItemCountThreshold
)

begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"

    $itemCountThreshold = Read-Param 'ItemCountThreshold' `
        -Value $ItemCountThreshold `
        -DefaultValue 50 `
        -Config $params.config `
        -ScriptName $params.scriptName

    $routingAddress = (
        Get-OrganizationConfig |
            Select-Object -ExpandProperty MicrosoftExchangeRecipientEmailAddresses |
            Where-Object { $_ -like '*mail.onmicrosoft.com' }
    ).Split('@')[1]
}

process {
    foreach ($exchangeObject in $params.exchangeObjects) {
        $itemCount = Get-MailboxFolderStatistics $exchangeObject |
            Where-Object { $_.FolderType -eq 'Root' } |
            Select-Object -ExpandProperty ItemsInFolderAndSubfolders

        if ($itemCount -le $itemCountThreshold -or $BypassItemCount) {
            $mailboxInfo = Get-Mailbox -Identity $exchangeObject

            foreach ($emailAddress in $mailboxInfo.EmailAddresses) {
                [PSCustomObject]@{
                    exchangeObject = $exchangeObject
                    emailAddress = $emailAddress
                } | Export-Csv $params.outputFilePath -Append
            }

            $hasArchiveGuid = $mailboxInfo.archiveGuid -ne '00000000-0000-0000-0000-000000000000'
            $hasArchive = $hasArchiveGuid -and $mailboxInfo.archiveDatabase

            if (!$hasArchive) {
                Disable-Mailbox -Identity $mailboxInfo.Alias -Confirm:$false
                Enable-RemoteMailbox $mailboxInfo.Alias `
                    -RemoteRoutingAddress "$exchangeObject@$routingAddress"
                Set-RemoteMailbox $mailboxInfo.UserPrincipalName `
                    -EmailAddresses $mailboxInfo.EmailAddresses `
                    -EmailAddressPolicyEnabled $false
            } else {
                $errorMessage = -join (
                    "Mailbox $exchangeObject has on-premise archive enabled. Processing the ",
                    'script requires permanently disabling the archive which will result in data ',
                    'loss. Consider merging the 2 mailbox objects manually after backup. Skipping ',
                    'this mailbox...'
                )

                Write-Error $errorMessage
            }
        } else {
            $errorMessage = -join (
                "Mailbox $exchangeObject has on-premise content above threshold. Processing the ",
                'script requires permanently disabling the mailbox which will result in data ',
                'loss. Consider merging the 2 mailbox objects manually after backup. Skipping ',
                'this mailbox...'
            )

            Write-Error $errorMessage
        }
    }
}
