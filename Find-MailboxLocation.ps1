begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"
}

process {
    "exchangeObject,mailboxLocation" >> $params.outputFilePath

    for ($index = 0; $index -lt $params.exchangeObjects.Count; $index++) {
        $exchangeObject = $params.exchangeObjects[$index]

        Write-Output "Processing object: $exchangeObject"

        $location = Get-ExchangeObjectLocation -ExchangeObject $exchangeObject

        "$exchangeObject,$location" >> $params.outputFilePath
    }
}
