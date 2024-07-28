<#
.SYNOPSIS
    Zoom.

.DESCRIPTION
    Script to install or uninstall Zoom.

.PARAMETER Mode
Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\zoominstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\zoominstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 2.0
    - Date: 30.05.2024
#>

#Region Parameters
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Install", "Uninstall", IgnoreCase = $true)]
    [string]$Mode
)

# Reference functions.ps1 (Assuming it contains necessary functions like Initialize-Directories and Write-LogEntry)
. "$PSScriptRoot\functions.ps1"

# Download function
function Start-DownloadFile {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$URL,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )
    begin {
        $WebClient = New-Object -TypeName System.Net.WebClient
    }
    process {
        if (-not (Test-Path -Path $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }
        $WebClient.DownloadFile($URL, (Join-Path -Path $Path -ChildPath $Name))
    }
    end {
        $WebClient.Dispose()
    }
}

# Initialize Directories
$folderPaths = Initialize-Directories -HomeFolder C:\HUD\
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$validationFolderVar = $folderPaths.ValidationFolder
$Date = Get-Date -Format "MM-dd-yyyy"
$AppName = "Zoom"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "2.0"
$LogFileName = "$($AppName)_${Mode}_$Date.log"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1
$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName
Write-LogEntry -Value "Setup folder has been created at: $SetupFolder." -Severity 1

# Install/Uninstall
if ($Mode -eq "Install") {
    try {
        $installerFileName = "ZoomInstallerFull.msi"
        Start-DownloadFile -URL "https://cdn.zoom.us/prod/6.0.11.39959/x64/ZoomInstallerFull.msi" -Path $SetupFolder -Name $installerFileName -ErrorAction Stop
        Write-LogEntry -Value "Downloaded Zoom installer to $installerFileName." -Severity 1

        $SetupFilePath = "$SetupFolder\$installerFileName"
        if (-not (Test-Path $SetupFilePath)) {
            throw "Installer file not found."
        }
        Write-LogEntry -Value "Found installer file at $SetupFilePath." -Severity 1

        $Arguments = "/quiet /qn /norestart ZoomAutoUpdate=1"
        $Process = Start-Process -FilePath $SetupFilePath -ArgumentList $Arguments -Wait -PassThru -ErrorAction Stop

        if ($Process.ExitCode -eq 0) {
            New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
            Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
            Write-LogEntry -Value "Install of $AppName is complete" -Severity 1
        } else {
            Write-LogEntry -Value "Install of $AppName failed with ExitCode: $($Process.ExitCode)" -Severity 3
        }

        # Cleanup 
        if (Test-Path "$SetupFolder") {
            Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
            Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
        }

    } catch [System.Exception] {
        Write-LogEntry -Value "Error running installer. Errormessage: $($_.Exception.Message)" -Severity 3
        return # Stop execution of the script after logging a critical error
    }

} elseif ($Mode -eq "Uninstall") {
    try {
        # Find Zoom Uninstaller
        $MyApp = Get-InstalledApps -App "workplace*"

        # Uninstall App
        $uninstall_command = 'MsiExec.exe'
        $Result = (($MyApp.UninstallString -split ' ')[1] -replace '/I','/X ') + ' /quiet'
        $uninstall_args = [string]$Result
        $uninstallProcess = Start-Process $uninstall_command -ArgumentList $uninstall_args -PassThru -Wait -ErrorAction Stop

        # Post Uninstall Actions
        if ($uninstallProcess.ExitCode -eq 0) {
            # Delete validation file
            try {
                Remove-Item -Path $AppValidationFile -Force -Confirm:$false
                Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1

                # Cleanup 
                if (Test-Path "$SetupFolder") {
                    Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
                    Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
                }
            } catch [System.Exception] {
                Write-LogEntry -Value "Error deleting validation file. Errormessage: $($_.Exception.Message)" -Severity 3
            }
        } else {
            throw "Uninstallation failed with exit code $($uninstallProcess.ExitCode)"
        }
    } catch [System.Exception] {
        Write-LogEntry -Value "Error completing uninstall. Errormessage: $($_.Exception.Message)" -Severity 3
        throw "Uninstallation halted due to an error"
    }

    Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
}
