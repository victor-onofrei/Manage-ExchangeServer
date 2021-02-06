. "$PSScriptRoot\Initializer.ps1"
$params = Invoke-Expression "Initialize-DefaultParams $args"

"exchangeObject,mailboxLocation" >> $params.outputFilePath
foreach ($exchangeObject in $params.exchangeObjects) {
    $exchangeObjectTypeDetails = (Get-Recipient -Identity $exchangeObject -ErrorAction SilentlyContinue).RecipientTypeDetails
    $mailboxLocation = $null
    $isRemote = $exchangeObjectTypeDetails -like "Remote*"

    if ($exchangeObjectTypeDetails -like "*Mailbox" -and (-not $isRemote)) {
        $mailboxLocation = "EXP"
    } elseif ($isRemote) {
        $mailboxLocation = "EXO"
    } else {
        $mailboxLocation = "N/A"
    }

    "$exchangeObject,$mailboxLocation" >> $params.outputFilePath
}
