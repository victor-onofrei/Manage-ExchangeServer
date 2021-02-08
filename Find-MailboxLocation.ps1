begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"
}

process {
    "exchangeObject,mailboxLocation" >> $params.outputFilePath
    foreach ($exchangeObject in $params.exchangeObjects) {
        $exchangeObjectTypeDetails = (
            Get-Recipient `
                -Identity $exchangeObject `
                -ErrorAction SilentlyContinue
        ).RecipientTypeDetails

        $mailboxLocation = $null
        $isLocal = $exchangeObjectTypeDetails -like "*Mailbox"
        $isRemote = $exchangeObjectTypeDetails -like "Remote*"

        if ($isLocal -and (-not $isRemote)) {
            $mailboxLocation = "EXP"
        } elseif ($isRemote) {
            $mailboxLocation = "EXO"
        } else {
            $mailboxLocation = "N/A"
        }

        "$exchangeObject,$mailboxLocation" >> $params.outputFilePath
    }
}
