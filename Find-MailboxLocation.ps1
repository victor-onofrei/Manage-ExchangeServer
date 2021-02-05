. "$PSScriptRoot\Initializer.ps1"
$params = Invoke-Expression "Initialize-DefaultParams $args"

# Write-Output $params.exchangeObjects

Add-Content -Path $params.outputFilePath exchangeObject','mailboxLocation
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
    Add-Content -Path $params.outputFilePath $exchangeObject','$mailboxLocation
    # Write-Output $exchangeObject','$mailboxLocation
}
