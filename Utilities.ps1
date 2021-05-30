using namespace System.Management.Automation

if (-not (Test-Path variable:global:configGlobalCategory)) {
    $params = @{
        Name = 'configGlobalCategory'
        Option = [ScopedItemOptions]::Constant
        Scope = 'Global'
        Value = 'Global'
        Visibility = [SessionStateEntryVisibility]::Private
    }

    Set-Variable @params
}

function Read-Param {
    param (
        [String]$Name,
        [Object[]]$Value,
        [Object[]]$DefaultValue,

        [Hashtable]$Config,
        [String]$ScriptName,
        [Switch]$AllowGlobal
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

        if ($AllowGlobal) {
            $globalValue = $Config[$configGlobalCategory][$key]

            if ($globalValue) {
                # Return the global value from the config file if it exists.
                return $globalValue
            }
        }
    }

    if ($DefaultValue) {
        # Return the default value if it exists.
        return $DefaultValue
    }

    # Return null otherwise.
    return $null
}

enum ExchangeObjectLocation {
    notAvailable
    exchangeOnPremises
    exchangeOnline
}

function Get-ExchangeObjectLocation {
    param (
        [String]$ExchangeObject
    )

    $exchangeObjectTypeDetails = (
        Get-Recipient -Identity $ExchangeObject -ErrorAction SilentlyContinue
    ) | Select-Object -ExpandProperty RecipientTypeDetails -ErrorAction SilentlyContinue

    $isLocal = $exchangeObjectTypeDetails -like '*Mailbox'
    $isRemote = $exchangeObjectTypeDetails -like 'Remote*'

    if ($isLocal -and (-not $isRemote)) {
        return [ExchangeObjectLocation]::exchangeOnPremises
    } elseif ($isRemote) {
        return [ExchangeObjectLocation]::exchangeOnline
    } else {
        return [ExchangeObjectLocation]::notAvailable
    }
}

function Send-ReportMail {
    param (
        [String]$From,
        [String]$To,
        [String]$CC,

        [String]$AttachmentFilePath,
        [String]$AttachmentFileName,

        [String]$SMTPHost
    )

    $attachment = New-Object Net.Mail.Attachment($AttachmentFilePath)

    $message = New-Object Net.Mail.MailMessage

    $message.From = $From
    $message.To.Add($To)
    $message.Cc.Add($CC)

    $message.Subject = "$AttachmentFileName report is ready"
    $message.Body = "Attached is the $AttachmentFileName report"

    $message.Attachments.Add($attachment)

    $smtp = New-Object Net.Mail.SmtpClient($SMTPHost)
    $smtp.Send($message)
}
