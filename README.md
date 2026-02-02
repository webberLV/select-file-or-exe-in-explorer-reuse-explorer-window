cli cmd use 



-NoProfile -STA -Command "Start-Process ctrlF3-sel-explorer-cli  -WindowStyle Hidden; Start-Sleep 1; Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('^{F3}')"
