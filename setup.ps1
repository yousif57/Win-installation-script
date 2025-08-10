# Enhanced Setup Script for Chocolatey, Winget, GPU drivers, and MAS activation
# Initialize log file
$logFile = "C:\SetupLog.txt"
$ErrorActionPreference = "Continue"

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] - $Message"
    $logEntry | Out-File -FilePath $logFile -Append -Encoding UTF8
    
    # Also write to console with color coding
    switch ($Level) {
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        default { Write-Host $logEntry -ForegroundColor White }
    }
}

function Test-InternetConnection {
    try {
        $response = Test-NetConnection -ComputerName "8.8.8.8" -Port 53 -InformationLevel Quiet -WarningAction SilentlyContinue
        return $response
    } catch {
        return $false
    }
}

function Install-Chocolatey {
    Write-Log "Starting Chocolatey installation" "INFO"
    
    try {
        # Check if Chocolatey is already installed
        $chocoInstalled = Get-Command choco -ErrorAction SilentlyContinue
        if ($chocoInstalled) {
            Write-Log "Chocolatey is already installed. Version: $(choco --version)" "SUCCESS"
            return $true
        }
        
        Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction Stop
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        
        Write-Log "Downloading and executing Chocolatey installation script" "INFO"
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        
        # Verify installation
        $chocoInstalled = Get-Command choco -ErrorAction SilentlyContinue
        if ($chocoInstalled) {
            Write-Log "Chocolatey installed successfully. Version: $(choco --version)" "SUCCESS"
            return $true
        } else {
            Write-Log "Chocolatey installation failed - command not found after installation" "ERROR"
            return $false
        }
    } catch {
        Write-Log "Error installing Chocolatey: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Install-Winget {
    Write-Log "Starting Winget installation" "INFO"
    
    try {
        # Check if Winget is already installed
        $wingetInstalled = Get-Command winget -ErrorAction SilentlyContinue
        if ($wingetInstalled) {
            Write-Log "Winget is already installed. Version: $(winget --version)" "SUCCESS"
            return $true
        }
        
        Write-Log "Winget not found, attempting multiple installation methods" "INFO"
        
        # Method 1: Check if already installed but not in PATH
        $wingetPaths = @(
            "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe",
            "${env:ProgramFiles}\WindowsApps\Microsoft.DesktopAppInstaller_*\winget.exe"
        )
        
        foreach ($path in $wingetPaths) {
            $resolvedPaths = Get-ChildItem $path -ErrorAction SilentlyContinue
            if ($resolvedPaths) {
                Write-Log "Found winget at: $($resolvedPaths[0].FullName)" "SUCCESS"
                return $true
            }
        }
        
        # Method 2: Try to install via direct download (more reliable)
        try {
            Write-Log "Downloading winget directly from GitHub releases" "INFO"
            $apiUrl = "https://api.github.com/repos/microsoft/winget-cli/releases/latest"
            $releaseInfo = Invoke-RestMethod -Uri $apiUrl -UseBasicParsing
            $downloadUrl = $releaseInfo.assets | Where-Object { $_.name -like "*.msixbundle" } | Select-Object -First 1 -ExpandProperty browser_download_url
            
            if ($downloadUrl) {
                $tempPath = "$env:TEMP\winget.msixbundle"
                Write-Log "Downloading from: $downloadUrl" "INFO"
                Invoke-WebRequest -Uri $downloadUrl -OutFile $tempPath -UseBasicParsing
                
                Write-Log "Installing winget package" "INFO"
                Add-AppxPackage -Path $tempPath -ErrorAction Stop
                Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
                
                Start-Sleep -Seconds 10
                
                # Refresh PATH
                $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
                
                $wingetInstalled = Get-Command winget -ErrorAction SilentlyContinue
                if ($wingetInstalled) {
                    Write-Log "Winget installed successfully via direct download. Version: $(winget --version)" "SUCCESS"
                    return $true
                }
            }
        } catch {
            Write-Log "Method 2 failed: $($_.Exception.Message)" "WARNING"
        }
        
        # Method 3: Try App Installer registration
        try {
            Write-Log "Attempting App Installer registration method" "INFO"
            $appxPackage = Get-AppxPackage -Name "Microsoft.DesktopAppInstaller" -ErrorAction SilentlyContinue
            if (-not $appxPackage) {
                Write-Log "Installing Microsoft.DesktopAppInstaller" "INFO"
                Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop
                Start-Sleep -Seconds 15
                
                # Refresh PATH again
                $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
                
                $wingetInstalled = Get-Command winget -ErrorAction SilentlyContinue
                if ($wingetInstalled) {
                    Write-Log "Winget installed successfully via App Installer registration. Version: $(winget --version)" "SUCCESS"
                    return $true
                }
            } else {
                Write-Log "Microsoft.DesktopAppInstaller already installed but winget not accessible" "WARNING"
            }
        } catch {
            Write-Log "Method 3 failed: $($_.Exception.Message)" "WARNING"
        }
        
        Write-Log "All winget installation methods failed" "ERROR"
        return $false
        
    } catch {
        Write-Log "Error in Winget installation process: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Install-ChocoApps {
    param([array]$Apps)
    
    Write-Log "Starting Chocolatey apps installation: $($Apps -join ', ')" "INFO"
    
    # First verify Chocolatey is accessible
    $chocoCmd = Get-Command choco -ErrorAction SilentlyContinue
    if (-not $chocoCmd) {
        Write-Log "Chocolatey command not found in PATH, attempting to refresh environment" "WARNING"
        
        # Try to refresh environment variables
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
        $chocoCmd = Get-Command choco -ErrorAction SilentlyContinue
        
        if (-not $chocoCmd) {
            # Try common Chocolatey installation paths
            $chocoPaths = @(
                "${env:ProgramData}\chocolatey\bin\choco.exe",
                "${env:ALLUSERSPROFILE}\chocolatey\bin\choco.exe"
            )
            
            foreach ($path in $chocoPaths) {
                if (Test-Path $path) {
                    Write-Log "Found Chocolatey at: $path" "INFO"
                    $chocoCmd = $path
                    break
                }
            }
        }
        
        if (-not $chocoCmd) {
            Write-Log "Chocolatey not accessible, cannot install apps" "ERROR"
            return @{ Successful = @(); Failed = $Apps }
        }
    }
    
    $successfulInstalls = @()
    $failedInstalls = @()
    
    foreach ($app in $Apps) {
        try {
            Write-Log "Installing $app via Chocolatey" "INFO"
            
            # Use the found choco command (either from PATH or direct path)
            $chocoPath = if ($chocoCmd.Source) { $chocoCmd.Source } else { $chocoCmd }
            
            Write-Log "Using Chocolatey at: $chocoPath" "INFO"
            Write-Log "Executing: $chocoPath install $app -y --force" "INFO"
            
            # Use cmd.exe to ensure proper environment loading
            $arguments = "/c `"$chocoPath`" install $app -y --force --no-progress --acceptlicense"
            $process = Start-Process -FilePath "cmd.exe" -ArgumentList $arguments -Wait -NoNewWindow -PassThru -RedirectStandardOutput "$env:TEMP\choco_$app_output.txt" -RedirectStandardError "$env:TEMP\choco_$app_error.txt"
            
            # Wait a moment for files to be written
            Start-Sleep -Seconds 2
            
            $output = Get-Content "$env:TEMP\choco_$app_output.txt" -Raw -ErrorAction SilentlyContinue
            $errorOutput = Get-Content "$env:TEMP\choco_$app_error.txt" -Raw -ErrorAction SilentlyContinue
            
            Write-Log "Process exit code for $app : $($process.ExitCode)" "INFO"
            
            if ($output) {
                Write-Log "Chocolatey output for $app : $($output.Substring(0, [Math]::Min(200, $output.Length)))" "INFO"
            }
            
            if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 1) {
                # Exit code 1 can also mean success with warnings in Chocolatey
                Write-Log "$app installed successfully via Chocolatey (Exit Code: $($process.ExitCode))" "SUCCESS"
                $successfulInstalls += $app
            } else {
                Write-Log "$app installation failed via Chocolatey (Exit Code: $($process.ExitCode))" "ERROR"
                if ($errorOutput) { 
                    Write-Log "Error details: $($errorOutput.Substring(0, [Math]::Min(500, $errorOutput.Length)))" "ERROR" 
                }
                $failedInstalls += $app
            }
            
            # Clean up temp files
            Remove-Item "$env:TEMP\choco_$app_output.txt" -Force -ErrorAction SilentlyContinue
            Remove-Item "$env:TEMP\choco_$app_error.txt" -Force -ErrorAction SilentlyContinue
            
        } catch {
            Write-Log "Exception installing $app via Chocolatey: $($_.Exception.Message)" "ERROR"
            $failedInstalls += $app
        }
        
        # Add small delay between installations
        Start-Sleep -Seconds 3
    }
    
    Write-Log "Chocolatey installation summary - Successful: $($successfulInstalls.Count), Failed: $($failedInstalls.Count)" "INFO"
    return @{ Successful = $successfulInstalls; Failed = $failedInstalls }
}

function Install-WingetApp {
    param([string]$AppId, [string]$AppName)
    
    try {
        Write-Log "Installing $AppName ($AppId) via Winget" "INFO"
        
        $process = Start-Process -FilePath "winget" -ArgumentList "install --id $AppId --silent --accept-package-agreements --accept-source-agreements" -Wait -NoNewWindow -PassThru -RedirectStandardOutput "$env:TEMP\winget_output.txt" -RedirectStandardError "$env:TEMP\winget_error.txt"
        
        $output = Get-Content "$env:TEMP\winget_output.txt" -Raw -ErrorAction SilentlyContinue
        $errorOutput = Get-Content "$env:TEMP\winget_error.txt" -Raw -ErrorAction SilentlyContinue
        
        if ($process.ExitCode -eq 0) {
            Write-Log "$AppName installed successfully via Winget (Exit Code: $($process.ExitCode))" "SUCCESS"
            $result = $true
        } else {
            Write-Log "$AppName installation failed via Winget (Exit Code: $($process.ExitCode))" "ERROR"
            if ($errorOutput) { Write-Log "Error details: $errorOutput" "ERROR" }
            $result = $false
        }
        
        # Clean up temp files
        Remove-Item "$env:TEMP\winget_output.txt" -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:TEMP\winget_error.txt" -Force -ErrorAction SilentlyContinue
        
        return $result
        
    } catch {
        Write-Log "Exception installing $AppName via Winget: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Get-GPUVendor {
    Write-Log "Detecting GPU vendor" "INFO"
    
    try {
        $gpus = Get-WmiObject Win32_VideoController | Where-Object { $_.PNPDeviceID -match "PCI\\VEN_" }
        $vendorMap = @{
            "10DE" = "NVIDIA"
            "1002" = "AMD"
            "8086" = "Intel"
        }
        
        foreach ($gpu in $gpus) {
            Write-Log "Found GPU: $($gpu.Name)" "INFO"
            
            if ($gpu.PNPDeviceID -match "VEN_([0-9A-F]{4})") {
                $vendorId = $matches[1].ToUpper()
                $vendor = $vendorMap[$vendorId]
                
                Write-Log "GPU Vendor ID: $vendorId, Vendor: $vendor" "INFO"
                
                if ($vendor) {
                    return $vendor
                }
            }
        }
        
        Write-Log "No supported GPU vendor detected" "WARNING"
        return $null
        
    } catch {
        Write-Log "Error detecting GPU vendor: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Install-GPUDriver {
    param([string]$Vendor)
    
    if (-not $Vendor) {
        Write-Log "No GPU vendor specified, skipping driver installation" "WARNING"
        return
    }
    
    Write-Log "Installing driver for $Vendor GPU" "INFO"
    
    try {
        switch ($Vendor) {
            "NVIDIA" {
                Write-Log "Installing NVIDIA drivers via Chocolatey" "INFO"
                $process = Start-Process -FilePath "choco" -ArgumentList "install nvidia-display-driver -y --no-progress" -Wait -NoNewWindow -PassThru
                
                if ($process.ExitCode -eq 0) {
                    Write-Log "NVIDIA drivers installed successfully (Exit Code: $($process.ExitCode))" "SUCCESS"
                } else {
                    Write-Log "NVIDIA drivers installation failed (Exit Code: $($process.ExitCode))" "ERROR"
                }
            }
            
            "Intel" {
                Write-Log "Downloading Intel Graphics installer" "INFO"
                $installerUrl = "https://dsadata.intel.com/installer"
                $installerPath = "$env:TEMP\Intel-Graphics-Installer.exe"
                
                try {
                    (New-Object System.Net.WebClient).DownloadFile($installerUrl, $installerPath)
                    Write-Log "Intel Graphics installer downloaded successfully" "SUCCESS"
                    
                    Write-Log "Running Intel Graphics installer" "INFO"
                    $process = Start-Process -FilePath $installerPath -ArgumentList "-s -noreboot" -Wait -NoNewWindow -PassThru
                    
                    if ($process.ExitCode -eq 0) {
                        Write-Log "Intel Graphics drivers installed successfully (Exit Code: $($process.ExitCode))" "SUCCESS"
                    } else {
                        Write-Log "Intel Graphics drivers installation failed (Exit Code: $($process.ExitCode))" "ERROR"
                    }
                    
                    Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
                    
                } catch {
                    Write-Log "Error downloading/installing Intel Graphics drivers: $($_.Exception.Message)" "ERROR"
                }
            }
            
            "AMD" {
                Write-Log "AMD driver installation not implemented in this script" "WARNING"
                Write-Log "Please download AMD drivers manually from AMD website" "INFO"
            }
        }
    } catch {
        Write-Log "Error installing $Vendor GPU driver: $($_.Exception.Message)" "ERROR"
    }
}

function Invoke-MASActivation {
    Write-Log "Starting MAS for Windows and Office activation" "INFO"
    
    try {
        Write-Log "Executing MAS activation script" "INFO"
        $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-Command & ([ScriptBlock]::Create((irm https://get.activated.win))) /HWID /Ohook" -Wait -NoNewWindow -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-Log "MAS activation completed successfully (Exit Code: $($process.ExitCode))" "SUCCESS"
        } else {
            Write-Log "MAS activation failed or completed with warnings (Exit Code: $($process.ExitCode))" "WARNING"
        }
        
    } catch {
        Write-Log "Error running MAS activation: $($_.Exception.Message)" "ERROR"
    }
}

# ==================== MAIN SCRIPT EXECUTION ====================

Write-Log "Starting enhanced setup script for Chocolatey, Winget, GPU drivers, and MAS activation" "INFO"
Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)" "INFO"
Write-Log "Operating System: $((Get-WmiObject Win32_OperatingSystem).Caption)" "INFO"
Write-Log "Running as Administrator: $(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))" "INFO"
Write-Log "Execution Policy: $(Get-ExecutionPolicy)" "INFO"

# Check internet connectivity
Write-Log "Checking internet connectivity" "INFO"
if (-not (Test-InternetConnection)) {
    Write-Log "No internet connection detected. Script cannot continue." "ERROR"
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Log "Internet connection verified" "SUCCESS"

# Install Chocolatey
$chocoSuccess = Install-Chocolatey

# Wait for Chocolatey to fully initialize
if ($chocoSuccess) {
    Write-Log "Waiting for Chocolatey environment to initialize..." "INFO"
    Start-Sleep -Seconds 10
    
    # Refresh environment variables after Chocolatey installation
    $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
    
    # Verify Chocolatey is accessible
    $chocoTest = Get-Command choco -ErrorAction SilentlyContinue
    if ($chocoTest) {
        Write-Log "Chocolatey is accessible and ready for use" "SUCCESS"
    } else {
        Write-Log "Chocolatey installed but not accessible in PATH" "WARNING"
        $chocoSuccess = $false
    }
}

if (-not $chocoSuccess) {
    Write-Log "Chocolatey installation failed, but continuing with script" "WARNING"
}

# Install Winget
$wingetSuccess = Install-Winget
if (-not $wingetSuccess) {
    Write-Log "Winget installation failed, but continuing with script" "WARNING"
}

# Install apps via Chocolatey (excluding Chrome since we'll use Winget for that)
if ($chocoSuccess) {
    $chocoApps = @("vlc", "discord", "k-litecodecpackfull", "filezilla", "winrar")
    $chocoResults = Install-ChocoApps -Apps $chocoApps
    
    if ($chocoResults.Failed.Count -gt 0) {
        Write-Log "Some Chocolatey apps failed to install: $($chocoResults.Failed -join ', ')" "WARNING"
    }
} else {
    Write-Log "Skipping Chocolatey app installations due to Chocolatey installation failure" "WARNING"
}

# Install Google Chrome via Winget
if ($wingetSuccess) {
    $chromeSuccess = Install-WingetApp -AppId "Google.Chrome" -AppName "Google Chrome"
    if (-not $chromeSuccess) {
        Write-Log "Failed to install Google Chrome via Winget, trying Chocolatey as backup" "WARNING"
        if ($chocoSuccess) {
            $chromeBackup = Install-ChocoApps -Apps @("googlechrome")
            if ($chromeBackup.Successful.Count -gt 0) {
                Write-Log "Google Chrome installed successfully via Chocolatey backup method" "SUCCESS"
            }
        }
    }
} else {
    Write-Log "Skipping Google Chrome installation via Winget due to Winget installation failure" "WARNING"
    # Try Chrome via Chocolatey as backup
    if ($chocoSuccess) {
        Write-Log "Attempting to install Google Chrome via Chocolatey as backup" "INFO"
        $chromeBackup = Install-ChocoApps -Apps @("googlechrome")
        if ($chromeBackup.Successful.Count -gt 0) {
            Write-Log "Google Chrome installed successfully via Chocolatey backup method" "SUCCESS"
        }
    }
}

# Detect GPU and install drivers
$gpuVendor = Get-GPUVendor
if ($gpuVendor) {
    Install-GPUDriver -Vendor $gpuVendor
} else {
    Write-Log "No supported GPU detected, skipping GPU driver installation" "WARNING"
}

# Run MAS activation
Invoke-MASActivation

Write-Log "Enhanced setup script completed" "SUCCESS"
Write-Log "Please check the log file at $logFile for detailed information" "INFO"

# Display summary
Write-Log "=== SETUP SUMMARY ===" "INFO"
Write-Log "Chocolatey: $(if ($chocoSuccess) { 'Installed' } else { 'Failed' })" "INFO"
Write-Log "Winget: $(if ($wingetSuccess) { 'Installed' } else { 'Failed' })" "INFO"
Write-Log "GPU Vendor: $(if ($gpuVendor) { $gpuVendor } else { 'Not detected' })" "INFO"
Write-Log "Log file location: $logFile" "INFO"

# Pause to let user review results
Write-Log "Script execution complete. Review the log above for any issues." "INFO"
Read-Host "Press Enter to exit"
