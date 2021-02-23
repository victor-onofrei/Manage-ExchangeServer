begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"
}

process {
    "exchangeObject,mailboxLocation" >> $params.outputFilePath

    foreach ($exchangeObject in $params.exchangeObjects) {
        Write-Output "Processing object: $exchangeObject"

        $location = Get-ExchangeObjectLocation -ExchangeObject $exchangeObject

        "$exchangeObject,$location" >> $params.outputFilePath
    }
}
