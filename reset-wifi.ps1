# Reset WiFi - Disable, wait 5 seconds, then re-enable
Disable-NetAdapter -Name "Wi-Fi 6" -Confirm:$false
Start-Sleep -Seconds 5
Enable-NetAdapter -Name "Wi-Fi 6" -Confirm:$false
