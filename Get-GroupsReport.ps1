<#
.SYNOPSIS
This report computes the companies groups belong to based on its managers' and members' companies.

.DESCRIPTION
This report has a mandatory parameter named -Type.
Value must be 'Distribution Groups' or 'Microsoft 365 Groups'.

.EXAMPLE
Get-GroupsReport -Type 'Distribution Groups' -CompanyIdentifierAttribute 'CustomAttribute1' -FirstCompanyName 'contoso' -FirstCompanyIdentifier 'contoso#employee' -SecondCompanyName 'fabrikam' -SecondCompanyIdentifier 'fabrikam#employee'

.PARAMETER Type
Specifies the type of groups the report will run for.
You may use 'Distribution Groups' or 'Microsoft 365 Groups'.
#>

using namespace System.Management.Automation

param (
    [Parameter(Mandatory)]
    [ValidateSet('Distribution Groups', 'Microsoft 365 Groups')]
    [String]
    $Type,

    [Alias('FCI')]
    [String]
    $FirstCompanyIdentifier,

    [Alias('FCN')]
    [String]
    $FirstCompanyName,

    [Alias('SCI')]
    [String]
    $SecondCompanyIdentifier,

    [Alias('SCN')]
    [String]
    $SecondCompanyName,

    [Alias('CIA')]
    [String]
    $CompanyIdentifierAttribute
)

begin {
    enum GroupsType {
        none
        distribution
        unified
    }

    . "$PSScriptRoot\Initializer.ps1"
    $params = Invoke-Expression "Initialize-DefaultParams $args"

    $typeParams = @{
        Name = 'Type'
        Value = $Type
        Config = $params.config
        ScriptName = $params.scriptName
    }
    $Type = Read-Param @typeParams

    $firstCompanyIdentifierParams = @{
        Name = 'FirstCompanyIdentifier'
        Value = $FirstCompanyIdentifier
        Config = $params.config
        ScriptName = $params.scriptName
    }
    $FirstCompanyIdentifier = Read-Param @firstCompanyIdentifierParams

    $firstCompanyNameParams = @{
        Name = 'FirstCompanyName'
        Value = $FirstCompanyName
        Config = $params.config
        ScriptName = $params.scriptName
    }
    $FirstCompanyName = Read-Param @firstCompanyNameParams

    $secondCompanyIdentifierParams = @{
        Name = 'SecondCompanyIdentifier'
        Value = $SecondCompanyIdentifier
        Config = $params.config
        ScriptName = $params.scriptName
    }
    $SecondCompanyIdentifier = Read-Param @secondCompanyIdentifierParams

    $secondCompanyNameParams = @{
        Name = 'SecondCompanyName'
        Value = $SecondCompanyName
        Config = $params.config
        ScriptName = $params.scriptName
    }
    $SecondCompanyName = Read-Param @secondCompanyNameParams

    $companyIdentifierAttributeParams = @{
        Name = 'CompanyIdentifierAttribute'
        Value = $CompanyIdentifierAttribute
        Config = $params.config
        ScriptName = $params.scriptName
    }
    $CompanyIdentifierAttribute = Read-Param @companyIdentifierAttributeParams

    switch ($Type) {
        'Distribution Groups' { $groupsType = [GroupsType]::distribution }
        'Microsoft 365 Groups' { $groupsType = [GroupsType]::unified }
        Default { $groupsType = [GroupsType]::none }
    }

    function Get-ManagersFromList {
        param (
            [Object[]]$List
        )

        $managers = $List |
            Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue |
            Select-Object $CompanyIdentifierAttribute, PrimarySmtpAddress, Company

        return $managers
    }

    function Get-MembersListFromGroup {
        param (
            [PSObject]$Group
        )

        switch ($groupsType) {
            ([GroupsType]::distribution) {
                $getADGroupMemberParams = @{
                    Identity = $Group.SamAccountName
                    Recursive = $true
                    ErrorAction = [ActionPreference]::SilentlyContinue
                }
                $list = Get-ADGroupMember @getADGroupMemberParams |
                    Get-ADUser -Properties mail |
                    Where-Object { $_.Enabled -eq 'True' -and $_.mail } |
                    Select-Object -ExpandProperty mail
            }
            ([GroupsType]::unified) {
                $getGroupParams = @{
                    Identity = $Group.SamAccountName
                    ErrorAction = [ActionPreference]::SilentlyContinue
                }
                $list = Get-Group @getGroupParams |
                    Select-Object -ExpandProperty Members
            }
        }

        return $list
    }

    function Get-MembersFromList {
        param (
            [Object[]]$List
        )

        $members = $List |
            Get-Recipient -ResultSize Unlimited -ErrorAction SilentlyContinue |
            Select-Object $CompanyIdentifierAttribute, PrimarySmtpAddress

        return $members
    }

    function Get-MembersFromGroup {
        param (
            [PSObject]$Group
        )

        $list = Get-MembersListFromGroup $Group

        if (-not $list) {
            return $null
        }

        $members = Get-MembersFromList $list

        if (-not $members) {
            return $null
        }

        return $members
    }
}

process {
    # Workaround for https://github.com/PowerShell/PSScriptAnalyzer/issues/1643.
    Write-Verbose $FirstCompanyIdentifier
    Write-Verbose $SecondCompanyIdentifier

    if ($groupsType -eq [GroupsType]::none) {
        Write-Error "Filtering by the groups type '$Type' is not implemented!"

        return
    }

    Start-Transcript "$($params.outputFilePath).txt"

    switch ($groupsType) {
        ([GroupsType]::distribution) {
            $groups = Get-DistributionGroup -ResultSize Unlimited |
                Select-Object WindowsEmailAddress, ManagedBy, Name, RecipientType, Guid,
                SamAccountName
        }
        ([GroupsType]::unified) {
            $groups = Get-Group -ResultSize Unlimited -Filter {
                RecipientTypeDetails -eq 'GroupMailbox'
            } |
                Select-Object WindowsEmailAddress, ManagedBy, Name, RecipientType, Guid,
                SamAccountName
        }
    }

    $groupsCount = @($groups).Count

    Write-Output "To process: $groupsCount groups"

    for ($index = 0; $index -lt $groupsCount; $index++) {
        Write-Output "`tProcessing group: $($index + 1) / $groupsCount"

        Start-Sleep -Milliseconds 500

        $group = $groups[$index]

        $groupManagersList = $group.ManagedBy

        $areManagersInBothCompanies = $false
        $areMembersInBothCompanies = $false

        $groupManagers = $null
        $groupMembers = $null

        $groupManagerProperties = $null
        $groupMemberProperties = $null

        $groupManagersCount = 0
        $groupMembersCount = 0

        $firstCompanyManagersCount = 0
        $firstCompanyMembersCount = 0

        $secondCompanyManagersCount = 0
        $secondCompanyMembersCount = 0

        if ($groupManagersList) {
            $groupManagers = Get-ManagersFromList $groupManagersList

            $groupManagersCount = ($groupManagers | Measure-Object).Count
            $groupManagerProperties = $groupManagers.$CompanyIdentifierAttribute

            $firstCompanyManagersCount = (
                $groupManagerProperties |
                    Where-Object { $_ -like $FirstCompanyIdentifier } |
                    Measure-Object
            ).Count
            $secondCompanyManagersCount = (
                $groupManagerProperties |
                    Where-Object { $_ -like $SecondCompanyIdentifier } |
                    Measure-Object
            ).Count

            $areManagersInBothCompanies = (
                $firstCompanyManagersCount -and $secondCompanyManagersCount
            )

            $groupUsersCount = $groupManagersCount

            $firstCompanyUsersCount = $firstCompanyManagersCount
            $secondCompanyUsersCount = $secondCompanyManagersCount
        } else {
            $groupMembers = Get-MembersFromGroup $group

            if ($groupMembers) {
                $groupMembersCount = ($groupMembers | Measure-Object).Count
                $groupMemberProperties = $groupMembers.$CompanyIdentifierAttribute

                $firstCompanyMembersCount = (
                    $groupMemberProperties |
                        Where-Object { $_ -like $FirstCompanyIdentifier } |
                        Measure-Object
                ).Count
                $secondCompanyMembersCount = (
                    $groupMemberProperties |
                        Where-Object { $_ -like $SecondCompanyIdentifier } |
                        Measure-Object
                ).Count

                $areMembersInBothCompanies = (
                    $firstCompanyMembersCount -and $secondCompanyMembersCount
                )

                $groupUsersCount = $groupMembersCount

                $firstCompanyUsersCount = $firstCompanyMembersCount
                $secondCompanyUsersCount = $secondCompanyMembersCount
            } else {
                $groupUsersCount = 0

                $firstCompanyUsersCount = 0
                $secondCompanyUsersCount = 0
            }
        }

        if ($groupManagersList) {
            if (-not $groupManagers) {
                $groupManagers = Get-ManagersFromList $groupManagersList
            }

            $groupManagersSMTPAddresses = (
                $groupManagers | Select-Object -ExpandProperty PrimarySmtpAddress
            ) -join ';'
            $groupManagersCompanies = (
                $groupManagers | Select-Object -ExpandProperty Company
            ) -join ';'
        } else {
            $groupManagersSMTPAddresses = ''
            $groupManagersCompanies = ''
        }

        $hasFirstCompanyUsers = $firstCompanyUsersCount -gt 0
        $hasSecondCompanyUsers = $secondCompanyUsersCount -gt 0
        $hasFirstOrSecondCompanyUsers = $hasFirstCompanyUsers -or $hasSecondCompanyUsers

        $hasOnlyFirstCompanyUsers = $hasFirstCompanyUsers -and (-not $hasSecondCompanyUsers)
        $hasOnlySecondCompanyUsers = (-not $hasFirstCompanyUsers) -and $hasSecondCompanyUsers

        if (-not $hasFirstOrSecondCompanyUsers) {
            $groupCompany = 'None'
        } elseif ($areManagersInBothCompanies) {
            $groupCompany = 'Mixed Owners'
        } elseif ($areMembersInBothCompanies) {
            $groupCompany = 'Mixed Users'
        } elseif ($groupUsersCount -eq $firstCompanyUsersCount -or $hasOnlyFirstCompanyUsers) {
            $groupCompany = $FirstCompanyName
        } elseif ($groupUsersCount -eq $secondCompanyUsersCount -or $hasOnlySecondCompanyUsers) {
            $groupCompany = $SecondCompanyName
        } else {
            $groupCompany = 'N/A'
        }

        if ($groupCompany -eq $FirstCompanyName -or $groupCompany -like 'Mixed*') {
            if (-not $groupMembers) {
                $groupMembers = Get-MembersFromGroup $group
            }

            if ($groupMembers) {
                $groupMembersSMTPAddresses = $groupMembers.PrimarySmtpAddress -join ';'
            } else {
                $groupMembersSMTPAddresses = ''
            }

            [PSCustomObject]@{
                'Group' = $group

                'Group Name' = $group.Name
                'Group GUID' = $group.Guid
                'Group SMTP' = $group.WindowsEmailAddress
                'Group Category' = $group.RecipientType
                'Group Company' = $groupCompany

                'Group Manager Properties' = $groupManagerProperties -join ';'
                'Group Member Properties' = $groupMemberProperties -join ';'
                'Group Managers Count' = $groupManagersCount
                'Group Members Count' = $groupMembersCount

                'First Company Managers Count' = $firstCompanyManagersCount
                'First Company Members Count' = $firstCompanyMembersCount
                'Second Company Managers Count' = $secondCompanyManagersCount
                'Second Company Members Count' = $secondCompanyMembersCount

                'Group Managers SMTP Addresses' = $groupManagersSMTPAddresses
                'Group Managers Companies' = $groupManagersCompanies

                'Group Members SMTP Addresses' = $groupMembersSMTPAddresses
            } | Export-Csv $params.outputFilePath -Append -NoTypeInformation
        }
    }

    Stop-Transcript
}

end {
    Send-DefaultReportMail -ScriptParams $params
}