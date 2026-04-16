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

    It 'returns a non-destructive scrub plan' {
        $plan = Invoke-OfficeScrubC2R -PlanOnly

        $plan.PSObject.TypeNames[0] | Should -Be 'OfficeScrubC2R.ScrubPlan'
        $plan.PlanOnly | Should -BeTrue
        $plan.PlannedOperations | Should -Not -BeNullOrEmpty
    }

    It 'supports WhatIf without destructive execution' {
        $plan = Invoke-OfficeScrubC2R -WhatIf

        $plan.PSObject.TypeNames[0] | Should -Be 'OfficeScrubC2R.ScrubPlan'
        $plan.PlannedOperations | Should -Not -BeNullOrEmpty
    }

    It 'blocks real destructive execution with a stable error id' {
        { Invoke-OfficeScrubC2R -Confirm:$false -ErrorAction Stop } |
            Should -Throw -ErrorId 'OfficeScrubC2R.DestructiveExecutionNotSupported,OfficeScrubC2R.PowerShell.InvokeOfficeScrubC2RCommand'
    }
}
