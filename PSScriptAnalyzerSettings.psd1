@{
    ExcludeRules = @(
        "PSAvoidUsingInvokeExpression"
    )
    Rules = @{
        PSAvoidLongLines  = @{
            Enable = $true
            MaximumLineLength = 100
        }
    }
}
