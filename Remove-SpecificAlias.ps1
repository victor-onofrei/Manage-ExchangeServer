param (
    [Alias('AI')][SupportsWildcards()][string]$AliasIdentifier
)

begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"

    $aliasIdentifierParams = @{
        Name = 'AliasIdentifier'
        Value = $AliasIdentifier
        Config = $params.config
        ScriptName = $params.scriptName
    }
    $aliasIdentifier = Read-Param @aliasIdentifierParams
}

process {
    foreach ($exchangeObject in $params.exchangeObjects) {
        $aliasToRemove = Get-Recipient $exchangeObject |
            Select-Object -Property @{
                Name = 'SpecificEmailAddresses';
                Expression = {
                    $_.EmailAddresses |
                        Where-Object { $_ -like "$aliasIdentifier" }
                    }
                } | Select-Object -ExpandProperty SpecificEmailAddresses
        foreach ($alias in $aliasToRemove) {
            $aliasToRemoveParams = @{
                EmailAddresses = @{ remove = "$alias" }
                EmailAddressPolicyEnabled = $false
            }
            Set-RemoteMailbox $exchangeObject @aliasToRemoveParams
        }
    }
}