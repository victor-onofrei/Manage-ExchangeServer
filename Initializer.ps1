. "$PSScriptRoot\Utilities.ps1"

Set-StrictMode -Version Latest

function Get-ScriptName {
    $callStack = Get-PSCallStack
    $scriptFileName = $callStack[1].Command

    return [IO.Path]::GetFileNameWithoutExtension($scriptFileName)
}

function Initialize-DefaultParam {
    [CmdletBinding()]
    [OutputType([Hashtable])]
    param (
        [String]$_ScriptName = (Get-ScriptName),

        [Alias('IP')][String]$InputPath,
        [Alias('ID')][String]$InputDir,
        [Alias('IFN')][String]$InputFileName,

        [Alias('OP')][String]$OutputPath,
        [Alias('OD')][String]$OutputDir,
        [Alias('OFN')][String]$OutputFileName,

        [Alias('EO')][String[]]$ExchangeObjects
    )

    begin {
        function Initialize-IniModule {
            Set-Variable 'iniModule' -Option Constant -Value 'PsIni'

            $isModuleInstalled = Get-Module $iniModule -ListAvailable

            if (-not $isModuleInstalled) {
                Install-Module $iniModule -AcceptLicense
            }

            Import-Module $iniModule -Verbose:$false
        }

        function Get-Config {
            Set-Variable 'configDirectoryPath' -Option Constant -Value "$HOME\.config"
            Set-Variable 'configFilePath' -Option Constant -Value (
                Join-Path $configDirectoryPath -ChildPath 'manage-exchange_server.ini'
            )

            Initialize-IniModule

            if (Test-Path $configFilePath -PathType Leaf) {
                # Read the existing config file.
                $config = Get-IniContent $configFilePath
            } else {
                # Create the `.config` folder if it doesn't exist.
                New-Item $configDirectoryPath `
                    -ItemType Directory `
                    -ErrorAction SilentlyContinue > $null

                # Create a default, empty config data.
                $config = [Ordered]@{
                    $configGlobalCategory = @{}
                }

                # Save the created config data in the file.
                Out-IniFile $configFilePath -InputObject $config
            }

            return $config
        }
    }

    process {
        # Load the config.
        $config = Get-Config

        # Set timestamp variable.
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

        # Read the params.
        $inputPathParams = @{
            Name = 'InputPath'
            Value = $InputPath
            DefaultValue = "$HOME"
            Config = $config
            ScriptName = $_ScriptName
            AllowGlobal = $true
        }
        $inputPath = Read-Param @inputPathParams

        $inputDirParams = @{
            Name = 'InputDir'
            Value = $InputDir
            Config = $config
            ScriptName = $_ScriptName
            AllowGlobal = $true
        }
        $inputDir = Read-Param @$inputDirParams

        $inputFileNameParams = @{
            Name = 'InputFileName'
            Value = $InputFileName
            DefaultValue = "input_$_ScriptName.csv"
            Config = $config
            ScriptName = $_ScriptName
            AllowGlobal = $true
        }
        $inputFileName = Read-Param @inputFileNameParams

        $outputPathParams = @{
            Name = 'OutputPath'
            Value = $OutputPath
            DefaultValue = "$HOME"
            Config = $config
            ScriptName = $_ScriptName
            AllowGlobal = $true
        }
        $outputPath = Read-Param @outputPathParams

        $outputDirParams = @{
            Name = 'OutputDir'
            Value = $OutputDir
            Config = $config
            ScriptName = $_ScriptName
            AllowGlobal = $true
        }
        $outputDir = Read-Param @outputDirParams

        $outputFileNameParams = @{
            Name = 'OutputFileName'
            Value = $OutputFileName
            DefaultValue = "output_$_ScriptName.$timestamp.csv"
            Config = $config
            ScriptName = $_ScriptName
            AllowGlobal = $true
        }
        $outputFileName = Read-Param @outputFileNameParams

        $intermediateInputFilePath = Join-Path $inputPath -ChildPath $inputDir
        $intermediateOutputFilePath = Join-Path $outputPath -ChildPath $outputDir
        $inputFilePath = Join-Path $intermediateInputFilePath -ChildPath $inputFileName
        $outputFilePath = Join-Path $intermediateOutputFilePath -ChildPath $outputFileName

        if (-not (Test-Path $intermediateOutputFilePath -PathType Container)) {
            New-Item $intermediateOutputFilePath `
                -ItemType Directory `
                -ErrorAction SilentlyContinue > $null
        }

        $exchangeObjectsParams = @{
            Name = 'ExchangeObjects'
            Value = $ExchangeObjects
            DefaultValue = (Get-Content $inputFilePath -ErrorAction SilentlyContinue)
            AllowGlobal = $true
        }
        $exchangeObjects = Read-Param @exchangeObjectsParams

        if ($null -eq $exchangeObjects) {
            $exchangeObjects = @()
        }
    }

    end {
        Write-Verbose "scriptName: $_ScriptName"
        Write-Verbose "config: $($config | ConvertTo-Json)"
        Write-Verbose "inputFilePath: $inputFilePath"
        Write-Verbose "outputFilePath: $outputFilePath"
        Write-Verbose "exchangeObjects.Count: $($exchangeObjects.Count)"

        return @{
            scriptName = $_ScriptName
            config = $config

            inputPath = $inputPath
            inputDir = $inputDir
            inputFileName = $inputFileName
            inputFilePath = $inputFilePath

            outputPath = $outputPath
            outputDir = $outputDir
            outputFileName = $outputFileName
            outputFilePath = $outputFilePath

            exchangeObjects = $exchangeObjects
        }
    }
}
