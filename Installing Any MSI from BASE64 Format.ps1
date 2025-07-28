# 1. Enter the name of the application for logging purposes.
$appName = "My Custom Application"

# 2. Enter a temporary filename for the installer.
$tempMsiFileName = "installer.msi"

# 3. Paste the Base64 string generated from your .msi file between the quotes.
$base64MsiString = "PASTE_YOUR_LONG_BASE64_MSI_STRING_HERE"

# ===================================================================================
# --- SCRIPT LOGIC: NO EDITS NEEDED BELOW THIS LINE ---
# ===================================================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message"
}

if ($base64MsiString -like "PASTE*") {
    Write-Log "Script has not been configured. Please edit the script to add the Base64 string for your MSI." -Level "ERROR"
    exit 1
}

$tempMsiPath = Join-Path -Path $env:TEMP -ChildPath $tempMsiFileName

try {
    # Decode the string and write the .msi file to the temp directory
    Write-Log "Decoding embedded installer for '$appName' to: $tempMsiPath"
    $msiBytes = [System.Convert]::FromBase64String($base64MsiString)
    [System.IO.File]::WriteAllBytes($tempMsiPath, $msiBytes)
    Write-Log "Installer file created successfully."

    # Silently install the .msi
    Write-Log "Starting silent installation of '$appName'..."
    $installArgs = "/i `"$tempMsiPath`" /qn /norestart"
    
    $proc = Start-Process -FilePath "msiexec.exe" -ArgumentList $installArgs -Wait -PassThru -ErrorAction Stop
    if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
        Write-Log "'$appName' installation completed successfully. Exit Code: $($proc.ExitCode)"
    } else {
        Write-Log "'$appName' installation failed with Exit Code: $($proc.ExitCode)." -Level "ERROR"
        exit 1
    }

} catch {
    Write-Log "A critical error occurred: $_" -Level "ERROR"
    exit 1
} finally {
    # Clean up the temporary installer file
    if (Test-Path -Path $tempMsiPath) {
        Write-Log "Cleaning up temporary installer file."
        Remove-Item -Path $tempMsiPath -Force
    }
}

exit 0