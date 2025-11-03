@{
    'ExcludeRules' = @('*')
    'Rules'        = @{
        'PSAvoidUsingCmdletAliases'        = @{
            'allowlist' = @('set', 'get', 'nv', 'Add-Type')
        }
        'PSAvoidUsingPlainTextForPassword' = @{
            'allowlist' = @('set', 'get', 'nv', 'Add-Type')
        }
    }
}