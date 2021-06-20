param (
    [Alias('DEAP')][Switch]$DisableEAP,
    [Alias('OUT')][Switch]$Output
)

begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"
}

process {
    $recipients = (
        Get-Recipient -ResultSize Unlimited |
            Where-Object {
                $_.RecipientTypeDetails -like '*Mailbox' -and
                $_.EmailAddressPolicyEnabled -eq $true
            } |
            Select-Object -ExpandProperty SamAccountName
    )

    if ($null -eq $recipients) {
        $recipients = @()
    }

    if ($Output) {
        $recipients |
            Select-Object SamAccountName, PrimarySmtpAddress, Company |
            Export-Csv $params.outputFilePath -NoTypeInformation
    }

    if ($DisableEAP) {
        $recipientsCount = $recipients.Count

        Write-Output "To process $recipientsCount recipients"

        for ($index = 0; $index -lt $recipientsCount; $index++) {
            $recipient = $recipients[$index]

            Write-Output "Processing recipient $($index + 1) / $recipientsCount | $recipient"

            $location = Get-ExchangeObjectLocation -ExchangeObject $recipient

            switch ($location) {
                ([ExchangeObjectLocation]::notAvailable) {
                    Write-Output "`tMailbox does not exist anymore. Skipping mailbox..."

                    break
                }
                ([ExchangeObjectLocation]::exchangeOnPremises) {
                    Set-Mailbox $recipient -EmailAddressPolicyEnabled $false

                    break
                }
                ([ExchangeObjectLocation]::exchangeOnline) {
                    Set-RemoteMailbox $recipient -EmailAddressPolicyEnabled $false

                    break
                }
            }
        }
    }
}
