begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"
}

process {
    "exchangeObject,mailboxLocation" >> $params.outputFilePath

    foreach ($exchangeObject in $params.exchangeObjects) {
        $location = Get-ExchangeObjectLocation -ExchangeObject $exchangeObject
        $locationDescription = Get-ExchangeObjectLocationDescription -Location $location

        "$exchangeObject,$locationDescription" >> $params.outputFilePath
    }
}
