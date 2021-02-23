param (
    [Alias("DEAP")][Switch]$DisableEAP,
    [Alias("OUT")][Switch]$Output
)

begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"
}

process {
    $FormatEnumerationLimit = -1
    $recipients = (
        Get-Recipient -ResultSize Unlimited |
            Where-Object {
                $_.RecipientTypeDetails -like "*Mailbox" -and
                $_.EmailAddressPolicyEnabled -eq $true -and
                (
                    $_.CustomAttribute8 -like "world=H*" -or
                    $_.CustomAttribute8 -like "world=N*"
                )
            } |
            Select-Object -ExpandProperty SamAccountName
    )

    if ($Output) {
        $recipients |
            Select-Object SamAccountName, PrimarySmtpAddress, Company > $params.outputFilePath
    }

    if ($DisableEAP) {
        $recipientsCount = @($recipients).count
        Write-Output "To process $recipientsCount recipients"
        for ($index = 0; $index -lt $recipientsCount; $index++) {
            $recipient = $recipients[$index]
            Write-Output "Processing recipient $($index + 1) / $recipientsCount | $recipient"
            $location = Get-ExchangeObjectLocation -ExchangeObject $recipient
            switch ($location) {
                ([ExchangeObjectLocation]::notAvailable) {
                    Write-Output "`tMailbox does not exist anymore. Skipping mailbox..."; Break
                }
                ([ExchangeObjectLocation]::exchangeOnPremises) {
                    Set-Mailbox $recipient -EmailAddressPolicyEnabled $false; Break
                }
                ([ExchangeObjectLocation]::exchangeOnline) {
                    Set-RemoteMailbox $recipient -EmailAddressPolicyEnabled $false; Break
                }
            }
        }
    }
}
