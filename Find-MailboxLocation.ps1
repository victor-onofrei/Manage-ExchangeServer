begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"
}

process {
    "exchangeObject,mailboxLocation" >> $params.outputFilePath

    $exchangeObjectsCount = $params.exchangeObjects.Count

    Write-Output "To process $exchangeObjectsCount exchange objects"

    for ($index = 0; $index -lt $exchangeObjectsCount; $index++) {
        $exchangeObject = $params.exchangeObjects[$index]

        Write-Output "Processing object: $exchangeObject"

        $location = Get-ExchangeObjectLocation -ExchangeObject $exchangeObject

        "$exchangeObject,$location" >> $params.outputFilePath
    }
}
