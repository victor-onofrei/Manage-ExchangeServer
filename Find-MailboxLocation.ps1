begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParam $args"
}

process {
    $exchangeObjectsCount = $params.exchangeObjects.Count

    Write-Output "To process $exchangeObjectsCount exchange objects"

    for ($index = 0; $index -lt $exchangeObjectsCount; $index++) {
        $exchangeObject = $params.exchangeObjects[$index]

        $location = Get-ExchangeObjectLocation -ExchangeObject $exchangeObject

        Write-Output (
            "Processed exchange object $($index + 1) / $exchangeObjectsCount | " +
            "$exchangeObject | $location"
        )

        [PSCustomObject]@{
            exchangeObject = $exchangeObject
            mailboxLocation = $location
        } | Export-Csv $params.outputFilePath -Append -NoTypeInformation
    }
}
