REM minimized current batch window
powershell -window minimized -command ""

Taskkill /IM "TokenUpdater.exe" /F
TokenUpdater.exe > log.txt