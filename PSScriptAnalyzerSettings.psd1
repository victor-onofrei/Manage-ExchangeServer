@{
    ExcludeRules = @(
        "PSAvoidUsingInvokeExpression"
    )
    Rules = @{
        PSAvoidLongLines = @{
            Enable = $true
            MaximumLineLength = 100
        }
        PSAvoidUsingDoubleQuotesForConstantString = @{
            Enable = $true
        }
        PSPlaceCloseBrace = @{
            Enable = $true
            NewLineAfter = $false
            NoEmptyLineBefore = $true
        }
        PSPlaceOpenBrace = @{
            Enable = $true
        }
        PSUseCompatibleSyntax = @{
            Enable = $true
            TargetVersions = @(
                "5",
                "6",
                "7"
            )
        }
        PSUseConsistentIndentation = @{
            Enable = $true
        }
        PSUseConsistentWhitespace = @{
            Enable = $true
            # CheckParameter = $true
            CheckPipeForRedundantWhitespace = $true
            # IgnoreAssignmentOperatorInsideHashTable = $true
        }
    }
}
