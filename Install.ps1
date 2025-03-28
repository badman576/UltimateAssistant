<# 
ULTIMATE ASSISTANT INSTALLER v4.0
Zero-Error Edition
#>

# Configuration
$RepoBase = "https://raw.githubusercontent.com/badman576/UltimateAssistant/main"
$InstallDir = "$env:ProgramFiles\UltimateAssistant"
$RequiredFiles = @("UltimateAssistant.exe", "config.json")

# Create installation directory
if (-not (Test-Path $InstallDir)) {
    New-Item -Path $InstallDir -ItemType Directory -Force | Out-Null
    Write-Host "Created installation directory" -ForegroundColor Green
}

# Download function with retries
function Get-AssistantFile {
    param($fileName)
    $attempts = 3
    $urls = @(
        "$RepoBase/bin/$fileName",
        "https://gitlab.com/badman576/UltimateAssistant-Mirror/-/raw/main/bin/$fileName"
    )
    
    foreach ($url in $urls) {
        for ($i = 1; $i -le $attempts; $i++) {
            try {
                Invoke-WebRequest $url -OutFile "$InstallDir\$fileName" -ErrorAction Stop
                Write-Host "Downloaded $fileName successfully" -ForegroundColor Green
                return $true
            } catch {
                Write-Host "Attempt $i failed for $url" -ForegroundColor Yellow
                Start-Sleep -Seconds 2
            }
        }
    }
    return $false
}

# Main installation
try {
    # Download all required files
    $allSuccess = $true
    foreach ($file in $RequiredFiles) {
        if (-not (Get-AssistantFile $file)) {
            Write-Host "Failed to download $file" -ForegroundColor Red
            $allSuccess = $false
        }
    }

    if ($allSuccess) {
        # Create shortcut
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut("$env:USERPROFILE\Desktop\Ultimate Assistant.lnk")
        $shortcut.TargetPath = "$InstallDir\UltimateAssistant.exe"
        $shortcut.WorkingDirectory = $InstallDir
        $shortcut.Save()

        Write-Host @"
        
        🎉 Installation Complete!
        ========================
        1. Double-click 'Ultimate Assistant' on your desktop
        2. First-run setup will begin automatically
        3. Say 'Hey Assistant' to activate

        Location: $InstallDir
        "@ -ForegroundColor Cyan
    } else {
        Write-Host "Installation failed - check internet connection" -ForegroundColor Red
    }
} catch {
    Write-Host "Critical error: $_" -ForegroundColor Red
} finally {
    # Cleanup COM object
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
}
