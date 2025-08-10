# Initialize log file
$logFile = "C:\SetupLog.txt"
function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

Write-Log "Starting setup script for winget, Chocolatey, GPU driver, and MAS activation"

# Function to check internet connectivity
function Test-InternetConnection {
    try {
        Test-Connection -ComputerName "www.google.com" -Count 1 -Quiet -ErrorAction Stop
        Write-Log "Internet connection confirmed"
        return $true
    } catch {
        Write-Log "No internet connection: $_"
        return $false
    }
}

# Function to install winget
function Install-Winget {
    Write-Log "Checking for winget installation"
    try {
        if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
            Write-Log "winget not found, attempting to install"
            $wingetUrl = "https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
            $wingetPath = "$env:TEMP\Microsoft.DesktopAppInstaller.msixbundle"
            Write-Log "Downloading winget from $wingetUrl"
            (New-Object System.Net.WebClient).DownloadFile($wingetUrl, $wingetPath)
            Write-Log "Installing winget"
            Add-AppxPackage -Path $wingetPath -ErrorAction Stop
            Write-Log "winget installed successfully"
        } else {
            Write-Log "winget already installed"
        }
    } catch {
        Write-Log "Error installing winget: $_"
        throw
    }
}

# Function to install Chocolatey
function Install-Choco {
    Write-Log "Installing Chocolatey"
    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction Stop
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        $installOutput = iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')) 2>&1
        Write-Log "Chocolatey installation output: $installOutput"
        Write-Log "Chocolatey installed successfully"
    } catch {
        Write-Log "Error installing Chocolatey: $_"
        throw
    }
}

# Function to install apps via Chocolatey
function Install-ChocoApps {
    param($Packages)
    foreach ($package in $Packages) {
        Write-Log "Installing Chocolatey package: $package"
        try {
            $installOutput = choco install -y $package 2>&1
            Write-Log "Chocolatey package $package installation output: $installOutput"
            Write-Log "Chocolatey package $package installed successfully"
        } catch {
            Write-Log "Error installing Chocolatey package $package: $_"
        }
    }
}

# Function to install apps via winget
function Install-WingetApps {
    param($Packages)
    foreach ($package in $Packages) {
        Write-Log "Installing winget package: $package"
        try {
            $installOutput = winget install --id $package -e --silent --accept-package-agreements --accept-source-agreements 2>&1
            Write-Log "winget package $package installation output: $installOutput"
            Write-Log "winget package $package installed successfully"
        } catch {
            Write-Log "Error installing winget package $package: $_"
        }
    }
}

# Function to detect GPU vendor
function Get-GPUVendor {
    Write-Log "Detecting GPU vendor"
    try {
        $gpu = Get-WmiObject Win32_VideoController -ErrorAction Stop | Select-Object Name, PNPDeviceID
        $vendorId = $null
        if ($gpu.PNPDeviceID -match "VEN_([0-9A-F]{4})") {
            $vendorId = $matches[1].ToUpper()
        }
        Write-Log "GPU Name: $($gpu.Name), Vendor ID: $vendorId"

        $vendorMap = @{
            "10DE" = "NVIDIA"
            "1002" = "AMD"
            "8086" = "Intel"
        }
        $vendor = $vendorMap[$vendorId]
        if ($vendor) {
            Write-Log "Detected vendor: $vendor"
            return $vendor
        } else {
            Write-Log "Unknown Vendor ID: $vendorId"
            return $null
        }
    } catch {
        Write-Log "Error detecting GPU vendor: $_"
        return $null
    }
}

# Function to install GPU driver
function Install-GPUDriver {
    param($Vendor)
    $driversPath = "D:\Drivers"
    if ($Vendor) {
        Write-Log "Installing driver app for $Vendor"
        try {
            if ($Vendor -eq "NVIDIA") {
                Write-Log "Installing NVIDIA App via Chocolatey (nvidia-display-driver)"
                $installOutput = choco install -y nvidia-display-driver 2>&1
                Write-Log "NVIDIA App installation output: $installOutput"
                Write-Log "NVIDIA App installer exit code: $LASTEXITCODE"
            } elseif ($Vendor -eq "AMD") {
                $installer = "$driversPath\AMD-Adrenalin-Installer.exe"
                Write-Log "Running AMD Adrenalin installer: $installer /S"
                $process = Start-Process -FilePath $installer -ArgumentList "/S" -Wait -NoNewWindow -PassThru
                Write-Log "AMD Adrenalin installer exit code: $($process.ExitCode)"
            } elseif ($Vendor -eq "Intel") {
                $installerUrl = "https://dsadata.intel.com/installer"
                $installerPath = "$env:TEMP\Intel-Graphics-Installer.exe"
                Write-Log "Downloading Intel Graphics installer from $installerUrl"
                (New-Object System.Net.WebClient).DownloadFile($installerUrl, $installerPath)
                Write-Log "Running Intel Graphics installer: $installerPath -s -noreboot"
                $process = Start-Process -FilePath $installerPath -ArgumentList "-s -noreboot" -Wait -NoNewWindow -PassThru
                Write-Log "Intel Graphics installer exit code: $($process.ExitCode)"
            }
        } catch {
            Write-Log "Error installing $Vendor driver app: $_"
        }
    } else {
        Write-Log "No supported GPU vendor detected or detection failed; skipping driver installation"
    }
}

# Function to run MAS activation
function Invoke-MASActivation {
    Write-Log "Running MAS for Windows and Office activation"
    try {
        Write-Log "Executing MAS: & ([ScriptBlock]::Create((irm https://get.activated.win))) /HWID /Ohook"
        $installOutput = & ([ScriptBlock]::Create((Invoke-RestMethod -Uri "https://get.activated.win" -ErrorAction Stop))) /HWID /Ohook 2>&1
        Write-Log "MAS execution output: $installOutput"
        Write-Log "MAS exit code: $LASTEXITCODE"
    } catch {
        Write-Log "Error running MAS: $_"
    }
}

# Main execution
try {
    # Check internet connection
    $hasInternet = Test-InternetConnection

    if ($hasInternet) {
        # Install winget
        Install-Winget

        # Install winget apps
        $wingetPackages = @("Google.Chrome")
        Install-WingetApps -Packages $wingetPackages

        # Install Chocolatey
        Install-Choco

        # Install Chocolatey apps
        $chocoPackages = @("notepadplusplus", "7zip")
        Install-ChocoApps -Packages $chocoPackages
    } else {
        Write-Log "Skipping winget and Chocolatey installations due to no internet connection"
    }

    # Detect and install GPU driver
    $vendor = Get-GPUVendor
    Install-GPUDriver -Vendor $vendor

    # Run MAS activation (requires internet for irm)
    if ($hasInternet) {
        Invoke-MASActivation
    } else {
        Write-Log "Skipping MAS activation due to no internet connection"
    }

    Write-Log "Setup script completed successfully"
} catch {
    Write-Log "Critical error in main execution: $_"
}
