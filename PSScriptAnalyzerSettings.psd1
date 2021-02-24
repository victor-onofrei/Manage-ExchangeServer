@{
    ExcludeRules = @(
        'PSAvoidUsingInvokeExpression'
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
                '5.0',
                '6.0',
                '7.0'
            )
        }
        PSUseConsistentIndentation = @{
            Enable = $true
        }
        PSUseConsistentWhitespace = @{
            Enable = $true
            CheckPipeForRedundantWhitespace = $true
        }
    }
}
