Set-StrictMode -Version Latest

function Get-ScriptName {
    $callStack = Get-PSCallStack
    $scriptFileName = $callStack[1].Command

    return [IO.Path]::GetFileNameWithoutExtension($scriptFileName)
}

function Initialize-DefaultParams {
    [CmdletBinding()]
    param (
        [String]$_ScriptName = (Get-ScriptName),

        [Alias("IP")][String]$InputPath,
        [Alias("ID")][String]$InputDir,
        [Alias("IFN")][String]$InputFileName,

        [Alias("OP")][String]$OutputPath,
        [Alias("OD")][String]$OutputDir,
        [Alias("OFN")][String]$OutputFileName,

        [Alias("EO")][String[]]$ExchangeObjects,

        [Alias("ICT")][Int]$ItemCountThreshold
    )

    begin {
        Set-Variable "configGlobalCategory" -Option Constant -Value "Global"

        function Initialize-IniModule {
            Set-Variable "iniModule" -Option Constant -Value "PsIni"

            $isModuleInstalled = Get-Module $iniModule -ListAvailable

            if (-not $isModuleInstalled) {
                Install-Module $iniModule -AcceptLicense
            }

            Import-Module $iniModule -Verbose:$false
        }

        function Get-Config {
            Set-Variable "configDirectoryPath" -Option Constant -Value "$HOME\.config"
            Set-Variable "configFilePath" -Option Constant -Value (Join-Path $configDirectoryPath -ChildPath "manage-exchange_server.ini")

            Initialize-IniModule

            if (Test-Path $configFilePath -PathType Leaf) {
                # Read the existing config file.
                $config = Get-IniContent $configFilePath
            } else {
                # Create the `.config` folder if it doesn't exist.
                New-Item $configDirectoryPath -ItemType Directory -ErrorAction SilentlyContinue > $null

                # Create a default, empty config data.
                $config = [Ordered]@{
                    $configGlobalCategory = @{}
                }

                # Save the created config data in the file.
                Out-IniFile $configFilePath -InputObject $config
            }

            return $config
        }

        function Read-Param {
            param (
                [String]$Name,
                [String[]]$Value,
                [String[]]$DefaultValue,

                [Hashtable]$Config,
                [String]$ScriptName
            )

            if ($Value) {
                # Return the actual value if it exists.
                return $Value
            }

            if ($Config) {
                # Return a value from the config file if it exists.

                $key = $Name

                $specificCategory = $ScriptName

                if ($Config[$specificCategory]) {
                    $specificValue = $Config[$specificCategory][$key]

                    if ($specificValue) {
                        # Return the specific value from the config file if it
                        # exists.
                        return $specificValue
                    }
                }

                $globalValue = $Config[$configGlobalCategory][$key]

                if ($globalValue) {
                    # Return the global value from the config file if it exists.
                    return $globalValue
                }
            }

            if ($DefaultValue) {
                # Return the default value if it exists.
                return $DefaultValue
            }

            # Return null otherwise.
            return $null
        }
    }

    process {
        # Load the config.
        $config = Get-Config

        # Set timestamp variable.
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

        # Read the params.

        $InputPath = Read-Param "InputPath" -Value $InputPath -DefaultValue "$HOME" -Config $config -ScriptName $_ScriptName
        $InputDir = Read-Param "InputDir" -Value $InputDir -Config $config -ScriptName $_ScriptName
        $InputFileName = Read-Param "InputFileName" -Value $InputFileName -DefaultValue "input_$_ScriptName.csv" -Config $config -ScriptName $_ScriptName

        $OutputPath = Read-Param "OutputPath" -Value $OutputPath -DefaultValue "$HOME" -Config $config -ScriptName $_ScriptName
        $OutputDir = Read-Param "OutputDir" -Value $OutputDir -Config $config -ScriptName $_ScriptName
        $OutputFileName = Read-Param "OutputFileName" -Value $OutputFileName -DefaultValue "output_$_ScriptName.$timestamp.csv" -Config $config -ScriptName $_ScriptName

        $intermediateInputFilePath = Join-Path $InputPath -ChildPath $InputDir
        $intermediateOutputFilePath = Join-Path $OutputPath -ChildPath $OutputDir
        $inputFilePath = Join-Path $intermediateInputFilePath -ChildPath $InputFileName
        $outputFilePath = Join-Path $intermediateOutputFilePath -ChildPath $OutputFileName

        if (-not (Test-Path $intermediateOutputFilePath -PathType Container)) {
            New-Item $intermediateOutputFilePath -ItemType Directory -ErrorAction SilentlyContinue > $null
        }

        $ExchangeObjects = Read-Param "ExchangeObjects" -Value $ExchangeObjects -DefaultValue (Get-Content $inputFilePath -ErrorAction SilentlyContinue)
        $ItemCountThreshold = Read-Param "ItemCountThreshold" -Value $ItemCountThreshold -DefaultValue "50" -Config $config -ScriptName $_ScriptName
    }

    end {
        Write-Verbose "inputFilePath: $inputFilePath"
        Write-Verbose "outputFilePath: $outputFilePath"
        Write-Verbose "exchangeObjects: $exchangeObjects"

        return @{
            inputPath = $InputPath
            inputDir = $InputDir
            inputFileName = $InputFileName
            inputFilePath = $inputFilePath

            outputPath = $OutputPath
            outputDir = $OutputDir
            outputFileName = $OutputFileName
            outputFilePath = $outputFilePath

            exchangeObjects = $ExchangeObjects
            itemCountThreshold = $ItemCountThreshold
        }
    }
}
