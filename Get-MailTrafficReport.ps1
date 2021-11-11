begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParam $args"
}

process {
    $serverName = $env:computername
    $beginDay = Get-Date -Hour 0 -Minute 00 -Second 00
    $endDay = $beginDay.AddDays(1).AddSeconds(-1)
    $outputFileName = $params.$intermediateOutputFilePath +
    $serverName + '_' + $beginDay.Day + '.' + $beginDay.Month + '.' + $beginDay.Year + '.csv'

    Get-MessageTrackingLog -Start $beginDay -End $endDay -EventID 'Receive' -ResultSize Unlimited |
        Where-Object { (Get-ExchangeServer).Fqdn -notcontains $_.ClientHostname } |
        Group-Object -Property ClientHostname, Sender |
        ForEach-Object -Process {
            New-Object -TypeName PSObject -Property @{
                Timestamp = ($_.Group.Timestamp | Select-Object -First 1).ToString().Split(' ')[0]
                Count = $_.Count
                ClientHostname = $_.Name.Split(',')[0]
                Sender = $_.Name.Split(',')[1].Split(' ')[1]
                ClientComputedIP =
                (Resolve-DnsName `
                    -Name $_.Name.Split(',')[0] `
                    -Type A `
                    -ErrorAction SilentlyContinue |
                    Select-Object `
                        -ExpandProperty IPAddress `
                        -ErrorAction SilentlyContinue) -join ';'
                }
            } |
            Sort-Object -Property Count -Descending |
            Select-Object Timestamp, ClientHostname, ClientComputedIP, Sender, Count |
            Export-Csv -Path $outputFileName -NoTypeInformation
}
