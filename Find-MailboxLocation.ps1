. "$PSScriptRoot\Initializer.ps1"
$params = Invoke-Expression "Initialize-DefaultParams $args"

"exchangeObject,mailboxLocation" >> $params.outputFilePath
foreach ($exchangeObject in $params.exchangeObjects) {
    $exchangeObjectTypeDetails = (Get-Recipient -Identity $exchangeObject -ErrorAction SilentlyContinue).RecipientTypeDetails
    $mailboxLocation = $null
    if ($exchangeObjectTypeDetails -like "*Mailbox" -and $exchangeObjectTypeDetails -notlike "Remote*") {
        $mailboxLocation = "EXP"
    } elseif ($exchangeObjectTypeDetails -like "Remote*") {
        $mailboxLocation = "EXO"
    } else {
        $mailboxLocation = "N/A"
    }
    "$exchangeObject,$mailboxLocation" >> $params.outputFilePath
}
