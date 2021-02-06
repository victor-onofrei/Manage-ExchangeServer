# Manage-ExchangeServer

PowerShell scripts for Management / Reporting within Exchange On Premises /
Online.

## Config file

Upon running any script for the first time, a default config file will be
created for you. You can also manually create it beforehand at the path
`$HOME/.config/manage-exchange_server.ini`, making sure that the required
sections are also included.

The config file is used to define default values for the common parameters.
These can be global, which apply to all the scripts, or specific, which apply to
a specific script.

These are used as a list of fall-backs, meaning, a valid value will be searched
in a specific order until one is found.

So the value used for a parameter will be:
1. The value specified on the command line if available,
2. Else, the value specified in the specific section of the config file for the
   respective script if available,
3. Else, the value specified in the `Global` section of the config file if
   available,
4. Else, a sensible default value defined by this project.

The syntax for a parameter value is:

```ini
ParamName=value
```

As a reference, here is a sample config file specifying global values for all
the parameters and specific values for the `Set-MailboxQuota` script:

```ini
[Global]
InputPath=C:\exchange
InputDir=inputs
InputFileName=input.csv

OutputPath=C:\exchange
OutputDir=outputs
OutputFileName=output.csv

[Set-MailboxQuota]
InputPath=C:\quotas
InputDir=inputs
InputFileName=users_list.csv

OutputPath=C:\quotas
OutputDir=outputs
OutputFileName=result.csv
```

## Parameters

There are some parameters that are available to all the scripts, for which you
can also configure default values in the config file as described in the
[`Config` section](#config).

Here is the list of these generally available parameters, along with their
respective aliases:

Name | Alias
--- | ---
`-InputPath` | IP
`-InputDir` | ID
`-InputFileName` | IFN
`-OutputPath` | OP
`-OutputDir` | OD
`-OutputFileName` | OFN
`-ExchangeObjects` | EO

_Note: It's not guaranteed that all of the scripts are using all of these
parameters. Some scripts might even decide not to use any of them._
