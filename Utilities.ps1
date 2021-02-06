Set-Variable "configGlobalCategory" -Option Constant -Value "Global"

function Read-Param {
    param (
        [String]$Name,
        [Object[]]$Value,
        [Object[]]$DefaultValue,

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
                # Return the specific value from the config file if it exists.
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
