# Replace with your actual Render service URL
$renderUrl = "https://outlook-action.onrender.com"

# Optional: path to log file
$logFile = "$PSScriptRoot\render-wake-log.txt"

# Timestamp for logging
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

try {
    $response = Invoke-WebRequest -Uri $renderUrl -UseBasicParsing -TimeoutSec 30
    $logEntry = "$timestamp - Server responded with status: $($response.StatusCode)"
} catch {
    $logEntry = "$timestamp - Failed to contact server: $($_.Exception.Message)"
}

# Write to log and console
Add-Content -Path $logFile -Value $logEntry
Write-Host $logEntry
