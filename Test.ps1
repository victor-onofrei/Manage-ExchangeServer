. "$PSScriptRoot\Initializers.ps1"
$params = Invoke-Expression "Initialize-DefaultParams -scriptName Test $args"

Write-Output "name: $($MyInvocation.MyCommand.Name)"
Write-Output "name2: $($MyInvocation.ScriptName)"
Write-Output "inputPath: $($params.inputPath)"
