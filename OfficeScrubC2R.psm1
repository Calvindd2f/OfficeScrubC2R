#
# OfficeScrubC2R Module
# PowerShell/C# implementation of Microsoft's Office Scrub C2R tool
# Version: 2.19.0
#

#Requires -Version 5.1

# Import the utilities module
$utilitiesPath = Join-Path $PSScriptRoot "OfficeScrubC2R-Utilities.psm1"
if (Test-Path $utilitiesPath) {
    Import-Module $utilitiesPath -Force -Global
}
else {
    throw "Required module not found: $utilitiesPath"
}

# Dot source the main script functions
$mainScriptPath = Join-Path $PSScriptRoot "OfficeScrubC2R.ps1"
if (Test-Path $mainScriptPath) {
    . $mainScriptPath
}
else {
    throw "Main script not found: $mainScriptPath"
}

<#
.SYNOPSIS
    Removes Office Click-to-Run (C2R) products from the system.

.DESCRIPTION
    Invoke-OfficeScrubC2R provides comprehensive removal of Office 2013, 2016, 2019, and Office 365
    Click-to-Run installations when standard uninstall methods fail. This is a PowerShell/C# port of
    Microsoft's OffScrubC2R.vbs with 10-50x performance improvements.

    The tool performs the following operations:
    - Stops all Office processes
    - Uninstalls Office products using ODT
    - Removes registry keys and values
    - Deletes Office files and folders
    - Cleans up Windows Installer metadata
    - Removes scheduled tasks and services
    - Clears Office licenses (optional)

    Requires Administrator privileges.

.PARAMETER Quiet
    Run in quiet mode with minimal output.

.PARAMETER DetectOnly
    Only detect installed products without removing them.

.PARAMETER Force
    Force removal without user confirmation.

.PARAMETER RemoveAll
    Remove all Office products.

.PARAMETER KeepLicense
    Keep Office licensing information.

.PARAMETER Offline
    Run in offline mode (skip ODT download).

.PARAMETER ForceArpUninstall
    Force ARP-based uninstall.

.PARAMETER ClearTaskBand
    Clear taskband shortcuts.

.PARAMETER UnpinMode
    Unpin shortcuts from taskbar and start menu.

.PARAMETER SkipSD
    Skip scheduled deletion operations.

.PARAMETER NoElevate
    Do not attempt elevation (will fail if not already elevated).

.PARAMETER LogPath
    Specify custom log path.

.EXAMPLE
    Invoke-OfficeScrubC2R
    
    Interactive mode - will prompt for confirmation before removal.

.EXAMPLE
    Invoke-OfficeScrubC2R -Quiet -Force
    
    Silent removal without prompts or confirmation.

.EXAMPLE
    Invoke-OfficeScrubC2R -DetectOnly
    
    Detect installed Office C2R products without removing them.

.EXAMPLE
    Invoke-OfficeScrubC2R -KeepLicense -LogPath "C:\Logs"
    
    Remove Office but keep licenses, saving logs to custom location.

.NOTES
    Author: Calvin (PowerShell/C# port)
    Original: Microsoft Corporation (OffScrubC2R.vbs)
    Version: 2.19.0
    Requires: PowerShell 5.1+, .NET Framework 4.5+, Administrator privileges
    
.LINK
    https://github.com/Calvindd2f/OfficeScrubC2R

.LINK
    Get-InstalledOfficeProducts

.LINK
    Test-IsC2R
#>
function Invoke-OfficeScrubC2R {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [Alias('Remove-OfficeC2R', 'Uninstall-OfficeC2R')]
    param(
        [Parameter(HelpMessage = "Run in quiet mode with minimal output")]
        [switch]$Quiet,

        [Parameter(HelpMessage = "Only detect installed products without removing them")]
        [switch]$DetectOnly,

        [Parameter(HelpMessage = "Force removal without user confirmation")]
        [switch]$Force,

        [Parameter(HelpMessage = "Remove all Office products")]
        [switch]$RemoveAll,

        [Parameter(HelpMessage = "Keep Office licensing information")]
        [switch]$KeepLicense,

        [Parameter(HelpMessage = "Run in offline mode")]
        [switch]$Offline,

        [Parameter(HelpMessage = "Force ARP-based uninstall")]
        [switch]$ForceArpUninstall,

        [Parameter(HelpMessage = "Clear taskband shortcuts")]
        [switch]$ClearTaskBand,

        [Parameter(HelpMessage = "Unpin shortcuts from taskbar")]
        [switch]$UnpinMode,

        [Parameter(HelpMessage = "Skip scheduled deletion")]
        [switch]$SkipSD,

        [Parameter(HelpMessage = "Do not attempt elevation")]
        [switch]$NoElevate,

        [Parameter(HelpMessage = "Specify custom log path")]
        [string]$LogPath
    )

    begin {
        # Check for Administrator privileges
        if (-not $NoElevate) {
            $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
            $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            
            if (-not $isAdmin) {
                throw "This command requires Administrator privileges. Please run PowerShell as Administrator."
            }
        }

        # Validate PowerShell version
        if ($PSVersionTable.PSVersion.Major -lt 5) {
            throw "This module requires PowerShell 5.1 or later. Current version: $($PSVersionTable.PSVersion)"
        }
    }

    process {
        if ($PSCmdlet.ShouldProcess("Office Click-to-Run Products", "Remove")) {
            try {
                # Set script-level variables for the main script functions
                $script:Quiet = $Quiet
                $script:DetectOnly = $DetectOnly
                $script:Force = $Force
                $script:RemoveAll = $RemoveAll
                $script:KeepLicense = $KeepLicense
                $script:Offline = $Offline
                $script:ForceArpUninstall = $ForceArpUninstall
                $script:ClearTaskBand = $ClearTaskBand
                $script:UnpinMode = $UnpinMode
                $script:SkipSD = $SkipSD
                $script:NoElevate = $NoElevate
                $script:LogPath = $LogPath

                # Call Main function directly (script already dot-sourced)
                $exitCode = Main
                
                # Return exit code
                return $exitCode
            }
            catch {
                Write-Error "Failed to execute OfficeScrubC2R: $_"
                throw
            }
        }
    }

    end {
        # Cleanup is handled by the main script
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Invoke-OfficeScrubC2R',
    'Get-InstalledOfficeProducts',
    'Test-IsC2R',
    'Initialize-Environment',
    'Stop-OfficeProcesses'
) -Alias @(
    'Remove-OfficeC2R',
    'Uninstall-OfficeC2R'
)

# Initialize environment after all exports
# This must happen after the utilities module is fully loaded
try {
    Initialize-Environment
    Write-Verbose "OfficeScrubC2R environment initialized successfully"
}
catch {
    Write-Warning "Failed to initialize OfficeScrubC2R environment: $_"
}

