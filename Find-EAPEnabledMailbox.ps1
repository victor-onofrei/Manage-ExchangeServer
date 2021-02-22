begin {
    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"
}

process {
    Get-Recipient -ResultSize Unlimited |
        Where-Object {
            $_.EmailAddressPolicyEnabled -eq $true
            -and
            $_.CustomAttribute8 -like "world=H*" } |
        Select-Object SamAccountName, PrimarySmtpAddress, Company > $params.outputFilePath
}