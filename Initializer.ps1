function Get-ScriptName {
    $callStack = Get-PSCallStack
    $scriptFileName = $callStack[1].Command

    return [io.path]::GetFileNameWithoutExtension($scriptFileName)
}

function Initialize-DefaultParams {
    param(
        [String]$_ScriptName = (Get-ScriptName),

        [String]$InputPath,
        [String]$InputDir,
        [String]$InputFileName,

        [String]$OutputPath,
        [String]$OutputDir,
        [String]$OutputFileName,

        [String[]]$ExchangeObjects
    )

    begin {
        Set-Variable "configGenericCategory" -Option Constant -Value "Generic"
        Set-Variable "configSpecificCategory" -Option Constant -Value "Specific"

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
                New-Item $configDirectoryPath -Type Directory > $null

                # Create a default, empty config data.
                $config = [ordered]@{
                    $configGenericCategory = @{}
                    $configSpecificCategory = @{}
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
                [string]$DefaultValue,

                [hashtable]$Config,
                [String]$ScriptName
            )

            if ($Value) {
                # Return the actual value if it exists.
                return $Value
            }

            $specificKey = "$($Name)_$ScriptName"
            $specificValue = $config[$configSpecificCategory][$specificKey]

            if ($specificValue) {
                # Return the specific value from the config file if it exists.
                return $specificValue
            }

            $genericKey = $Name
            $genericValue = $config[$configGenericCategory][$genericKey]

            if ($genericValue) {
                # Return the generic value from the config file if it exists.
                return $genericValue
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
