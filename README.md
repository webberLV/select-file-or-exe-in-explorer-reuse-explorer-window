cli cmd use 

powershell -NoProfile -STA -Command "Start-Process f3sel.exe -WindowStyle Hidden; Start-Sleep 1; Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('^{F3}')"
