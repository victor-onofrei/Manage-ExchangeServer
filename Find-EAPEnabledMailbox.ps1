param (
    [Alias("DEAP")][Switch]$DisableEAP
)

begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"
}

process {
    if ($DisableEAP) {
        $recipients = (
            Get-Recipient -ResultSize 100) |
            Where-Object {
                $_.EmailAddressPolicyEnabled -eq $true -and
                $_.CustomAttribute8 -like "world=H*" -or
                $_.CustomAttribute8 -like "world=N*"
            }
        $recipientsCount = @($recipients).count
        Write-Output "To process $recipientsCount recipients"
        for ($index = 0; $index -lt $recipientsCount; $index++) {
            Write-Output -InputObject `
                "`tProcessing recipient " ($index + 1) " / $recipientsCount | $recipient"
            $recipient = $recipients[$index]
            $location = Get-ExchangeObjectLocation -ExchangeObject $recipient
            switch ($location) {
                ([ExchangeObjectLocation]::notAvailable) {
                    Write-Output "Mailbox does not exist anymore. Skipping mailbox..."
                }
                ([ExchangeObjectLocation]::exchangeOnPremises) {
                    # Set-Mailbox $recipient -EmailAddressPolicyEnabled $false -WhatIf
                    Write-Output "$recipient is on prem"
                }
                ([ExchangeObjectLocation]::exchangeOnline) {
                    # Set-RemoteMailbox $recipient -EmailAddressPolicyEnabled $false -WhatIf
                    Write-Output "$recipient is in cloud"
                }
            }
        }
    }
    $FormatEnumerationLimit = -1
    Get-Recipient -ResultSize Unlimited |
        Where-Object {
            $_.EmailAddressPolicyEnabled -eq $true -and
            $_.CustomAttribute8 -like "world=H*"
        } |
        Select-Object SamAccountName, PrimarySmtpAddress, Company > $params.outputFilePath
}
