function Get-ScriptName {
    $callStack = Get-PSCallStack
    $scriptFileName = $callStack[1].Command

    return [io.path]::GetFileNameWithoutExtension($scriptFileName)
}

function Initialize-DefaultParams {
    param(
        [String]$_ScriptName = (Get-ScriptName),

        [Alias("IP")][String]$InputPath,
        [Alias("ID")][String]$InputDir,
        [Alias("IFN")][String]$InputFileName,

        [Alias("OP")][String]$OutputPath,
        [Alias("OD")][String]$OutputDir,
        [Alias("OFN")][String]$OutputFileName,

        [Alias("EO")][String[]]$ExchangeObjects
    )

    begin {
        Set-Variable "configGlobalCategory" -Option Constant -Value "Global"

        function Initialize-IniModule {
            Set-Variable "iniModule" -Option Constant -Value "PsIni"

            $isModuleInstalled = Get-Module $iniModule -ListAvailable

            if (-not $isModuleInstalled) {
                Install-Module $iniModule -AcceptLicense
            }

            Import-Module $iniModule
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
                New-Item $configDirectoryPath -Type Directory -ErrorAction SilentlyContinue > $null

                # Create a default, empty config data.
                $config = [ordered]@{
                    $configGlobalCategory = @{}
                }

                # Save the created config data in the file.
                Out-IniFile $configFilePath -InputObject $config
            }

            return $config
        }

        function Read-Param {
            param(
                [String]$Name,
                [String]$Value,
                [String[]]$DefaultValue,

                [hashtable]$Config,
                [String]$ScriptName
            )

            if ($Value) {
                # Return the actual value if it exists.
                return $Value
            }

            $key = $Name

            $specificCategory = $ScriptName

            if ($config[$specificCategory]) {
                $specificValue = $config[$specificCategory][$key]

                if ($specificValue) {
                    # Return the specific value from the config file if it
                    # exists.
                    return $specificValue
                }
            }

            $globalValue = $config[$configGlobalCategory][$key]

            if ($globalValue) {
                # Return the global value from the config file if it exists.
                return $globalValue
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

        # Read the params.

        $InputPath = Read-Param "InputPath" -Value $InputPath -DefaultValue "$HOME" -Config $config -ScriptName $_ScriptName
        $InputDir = Read-Param "InputDir" -Value $InputDir -Config $config -ScriptName $_ScriptName
        $InputFileName = Read-Param "InputFileName" -Value $InputFileName -DefaultValue "input_$_ScriptName.csv" -Config $config -ScriptName $_ScriptName

        $OutputPath = Read-Param "OutputPath" -Value $OutputPath -DefaultValue "$HOME" -Config $config -ScriptName $_ScriptName
        $OutputDir = Read-Param "OutputDir" -Value $OutputDir -Config $config -ScriptName $_ScriptName
        $OutputFileName = Read-Param "OutputFileName" -Value $OutputFileName -DefaultValue "output_$_ScriptName.csv" -Config $config -ScriptName $_ScriptName

        $intermediateInputFilePath = Join-Path $InputPath -ChildPath $InputDir
        $intermediateOutputFilePath = Join-Path $OutputPath -ChildPath $OutputDir
        $inputFilePath = Join-Path $intermediateInputFilePath -ChildPath $InputFileName
        $outputFilePath = Join-Path $intermediateOutputFilePath -ChildPath $OutputFileName

        $ExchangeObjects = Read-Param "ExchangeObjects" -Value $ExchangeObjects -DefaultValue (Get-Content $inputFilePath -ErrorAction SilentlyContinue) -Config $config -ScriptName $_ScriptName
    }

    end {
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
        }
    }
}
