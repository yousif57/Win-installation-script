# Initialize log file
$logFile = "C:\SetupLog.txt"
function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append
}

Write-Log "Starting setup script for Chocolatey, GPU driver, and MAS activation"

# Install Chocolatey
Write-Log "Installing Chocolatey"
try {
    Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction Stop
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    Write-Log "Chocolatey installed successfully"
} catch {
    Write-Log "Error installing Chocolatey: $_"
}

# Install apps via Chocolatey
Write-Log "Installing apps via Chocolatey (vlc discord k-litecodecpackfull filezilla googlechrome winrar)"
try {
    choco install -y vlc discord k-litecodecpackfull filezilla googlechrome winrar
    choco install googlechrome -y --ignore-checksums
    Write-Log "Chocolatey apps installed successfully"
} catch {
    Write-Log "Error installing Chocolatey apps: $_"
}

# Detect GPU vendor using Vendor ID
Write-Log "Detecting GPU vendor"
try {
    $gpu = Get-WmiObject Win32_VideoController | Select-Object Name, PNPDeviceID
    $vendorId = $null
    if ($gpu.PNPDeviceID -match "VEN_([0-9A-F]{4})") {
        $vendorId = $matches[1].ToUpper()
    }
    Write-Log "GPU Name: $($gpu.Name), Vendor ID: $vendorId"

    # Map Vendor IDs to manufacturers
    $vendorMap = @{
        "10DE" = "NVIDIA"
        "1002" = "AMD"
        "8086" = "Intel"
    }
    $vendor = $vendorMap[$vendorId]
    Write-Log "Detected vendor: $vendor"
} catch {
    Write-Log "Error detecting GPU vendor: $_"
    $vendor = $null
}

# Install GPU driver app based on vendor
if ($vendor) {
    Write-Log "Installing driver app for $vendor"
    try {
        if ($vendor -eq "NVIDIA") {
            Write-Log "Installing NVIDIA App via Chocolatey (nvidia-display-driver)"
            $process = Start-Process -FilePath "choco" -ArgumentList "install nvidia-display-driver -y" -Wait -NoNewWindow -PassThru
            Write-Log "NVIDIA App installer exit code: $($process.ExitCode)"

       } elseif ($vendor -eq "AMD") {
            # Create temp directory for AMD driver
            if (-not (Test-Path $tempPath)) {
                New-Item -ItemType Directory -Path $tempPath -Force | Out-Null
                Write-Log "Created temp directory: $tempPath"
            }
            $amdInstaller = "$tempPath\amd-software-adrenalin-edition-25.8.1-minimalsetup-250801_web.exe"
            $amdUrl = "https://raw.githubusercontent.com/yousif57/Win-installation-script/main/amd-software-adrenalin-edition-25.8.1-minimalsetup-250801_web.exe"
            Write-Log "Downloading AMD driver from $amdUrl"
            try {
                (New-Object System.Net.WebClient).DownloadFile($amdUrl, $amdInstaller)
                Write-Log "AMD driver downloaded successfully to $amdInstaller"
            } catch {
                Write-Log "Error downloading AMD driver: $_"
                throw
            }
            Write-Log "Running AMD Adrenalin installer: $amdInstaller /S"
            $process = Start-Process -FilePath $amdInstaller -ArgumentList "/S" -Wait -NoNewWindow -PassThru
            Write-Log "AMD Adrenalin installer exit code: $($process.ExitCode)"

        } elseif ($vendor -eq "Intel") {
            $installerUrl = "https://dsadata.intel.com/installer"
            $installerPath = "$env:TEMP\Intel-Graphics-Installer.exe"
            Write-Log "Downloading Intel Graphics installer from $installerUrl"
            (New-Object System.Net.WebClient).DownloadFile($installerUrl, $installerPath)
            Write-Log "Running Intel Graphics installer: $installerPath -s -noreboot"
            $process = Start-Process -FilePath $installerPath -ArgumentList "-s -noreboot" -Wait -NoNewWindow -PassThru
            Write-Log "Intel Graphics installer exit code: $($process.ExitCode)"
        }
    } catch {
        Write-Log "Error installing $vendor driver app: $_"
    }
} else {
    Write-Log "No supported GPU vendor detected or detection failed; skipping driver installation"
}

# Run MAS for Windows and Office activation
Write-Log "Running MAS for Windows and Office activation"
try {
    Write-Log "Executing MAS: & ([ScriptBlock]::Create((irm https://get.activated.win))) /HWID /z-office"
    $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-Command & ([ScriptBlock]::Create((irm https://get.activated.win))) /HWID /Ohook" -Wait -NoNewWindow -PassThru
    Write-Log "MAS exit code: $($process.ExitCode)"
} catch {
    Write-Log "Error running MAS: $_"
}

Write-Log "Setup script completed"
