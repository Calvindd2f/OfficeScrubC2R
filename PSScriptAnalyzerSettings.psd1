@{
    Severity     = @('Error', 'Warning')
    ExcludeRules = @(
        'PSAvoidUsingWriteHost'
    )
    Rules        = @{
        PSAvoidUsingCmdletAliases = @{
            AllowList = @()
        }
    }
}
