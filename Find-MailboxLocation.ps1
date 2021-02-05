. "$PSScriptRoot\Initializer.ps1"
$params = Invoke-Expression "Initialize-DefaultParams $args"

foreach ($exchangeObject in $params.exchangeObjects) {
    $exchangeObjectTypeDetails = Get-Recipient -Identity $exchangeObject -ErrorAction SilentlyContinue | Select -ExpandProperty RecipientTypeDetails -ErrorAction SilentlyContinue
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
