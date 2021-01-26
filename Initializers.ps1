function Initialize-DefaultParams {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$scriptName,

        [String]$inputPath = "$env:homeshare\VDI-UserData\Download\generic\inputs\",
        [String]$inputFilename = "input.csv",

        [String]$outputPath = "$env:homeshare\VDI-UserData\Download\generic\outputs\",
        [String]$outputDir = $null,
        [String]$outputFilename = "output.csv"

        # [ValidateNotNullOrEmpty()]
        # [String[]]$mailboxes = (Get-Content "$inputPath\$inputFilename")
    )

    # TODO: Find a way to make this work
    $cStack = @(Get-PSCallStack)
    $script = $cStack[$cStack.Length-1].InvocationInfo.MyCommand

    Write-Host "scriptName: $($script)"

    return @{
        inputPath = $inputPath;
        inputFilename = $inputFilename;

        outputPath = $outputPath;
        outputDir = $outputDir;
        outputFilename = $outputFilename;

        mailboxes = $mailboxes;
    }
}
