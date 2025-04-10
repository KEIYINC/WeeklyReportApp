# Uygulama klasörünü oluştur
$appFolder = "C:\Program Files\WeeklyReportApp"
New-Item -ItemType Directory -Force -Path $appFolder

# Yayınlanan dosyaları kopyala
Copy-Item -Path "bin\Release\net6.0-windows\win-x64\publish\*" -Destination $appFolder -Recurse -Force

# Kısayol oluştur
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\WeeklyReportApp.lnk")
$Shortcut.TargetPath = "$appFolder\WeeklyReportApp.exe"
$Shortcut.Save()

# Zamanlanmış görevi oluştur
$action = New-ScheduledTaskAction -Execute "$appFolder\WeeklyReportApp.exe"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Friday -At 6PM
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive

Register-ScheduledTask -TaskName "WeeklyReportApp" -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "Weekly Report Application - Runs every Friday at 6 PM" -Force

Write-Host "WeeklyReportApp başarıyla kuruldu!"
Write-Host "Masaüstünde bir kısayol oluşturuldu."
Write-Host "Her Cuma saat 18:00'de otomatik olarak çalışacak şekilde ayarlandı." 