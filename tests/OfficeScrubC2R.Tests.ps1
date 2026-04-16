$ErrorActionPreference = 'Stop'

Describe 'OfficeScrubC2R binary module contract' {
    BeforeAll {
        $repoRoot = Split-Path -Parent $PSScriptRoot
        & (Join-Path $repoRoot 'build.ps1') -Clean
        Import-Module (Join-Path $repoRoot 'OfficeScrubC2R.psd1') -Force
    }

    It 'exports exactly the supported public cmdlets' {
        $commands = Get-Command -Module OfficeScrubC2R | Select-Object -ExpandProperty Name | Sort-Object
        $commands | Should -Be @(
            'Get-InstalledOfficeProducts',
            'Invoke-OfficeScrubC2R',
            'Test-OfficeC2RState'
        )
    }

    It 'validates the manifest' {
        $repoRoot = Split-Path -Parent $PSScriptRoot
        $manifest = Test-ModuleManifest -Path (Join-Path $repoRoot 'OfficeScrubC2R.psd1')

        $manifest.Version.ToString() | Should -Be '3.0.0'
        $manifest.ExportedCmdlets.Keys | Sort-Object | Should -Be @(
            'Get-InstalledOfficeProducts',
            'Invoke-OfficeScrubC2R',
            'Test-OfficeC2RState'
        )
        $manifest.ExportedFunctions.Keys.Count | Should -Be 0
    }

    It 'runs product detection without requiring elevation' {
        { Get-InstalledOfficeProducts } | Should -Not -Throw
    }

    It 'returns a preflight state object' {
        $state = Test-OfficeC2RState

        $state.PSObject.TypeNames[0] | Should -Be 'OfficeScrubC2R.OfficeC2RState'
        ($state.InstalledProducts -is [System.Collections.IEnumerable]) | Should -BeTrue
        ($state.Issues -is [System.Collections.IEnumerable]) | Should -BeTrue
    }

    It 'returns a planning-only scrub plan' {
        $plan = Invoke-OfficeScrubC2R -PlanOnly

        $plan.PSObject.TypeNames[0] | Should -Be 'OfficeScrubC2R.ScrubPlan'
        $plan.PlanOnly | Should -BeTrue
        $plan.PlannedOperations | Should -Not -BeNullOrEmpty
        $plan.PlannedOperations |
            Where-Object { $_.Step -eq 'CompanionApps' -and $_.Action -eq 'RemoveTeamsAndCopilot' } |
            Should -Not -BeNullOrEmpty
    }

    It 'supports keeping Teams and Copilot during planning' {
        $plan = Invoke-OfficeScrubC2R -PlanOnly -KeepTeams -KeepCopilot

        $plan.KeepTeams | Should -BeTrue
        $plan.KeepCopilot | Should -BeTrue
        $plan.PlannedOperations |
            Where-Object { $_.Step -eq 'CompanionApps' -and $_.Status -eq 'Skipped' } |
            Should -Not -BeNullOrEmpty
    }

    It 'supports WhatIf without destructive execution' {
        $plan = Invoke-OfficeScrubC2R -WhatIf

        $plan.PSObject.TypeNames[0] | Should -Be 'OfficeScrubC2R.ScrubPlan'
        $plan.ExecutionStatus | Should -Be 'WhatIf'
        $plan.PlannedOperations | Should -Not -BeNullOrEmpty
    }

    It 'accepts legacy v2 invocation switches for compatibility' {
        { Invoke-OfficeScrubC2R -Quiet -Force -RemoveAll -WhatIf } | Should -Not -Throw
    }

    It 'blocks real destructive execution when not elevated with a stable error id' {
        $state = Test-OfficeC2RState
        if ($state.IsElevated) {
            Set-ItResult -Skipped -Because 'This test must not run destructive cleanup on an elevated host.'
            return
        }

        { Invoke-OfficeScrubC2R -Confirm:$false -ErrorAction Stop } |
            Should -Throw -ErrorId 'OfficeScrubC2R.AdminRequired,OfficeScrubC2R.PowerShell.InvokeOfficeScrubC2RCommand'
    }
}
