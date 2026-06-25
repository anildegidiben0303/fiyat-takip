@echo off
netsh advfirewall firewall add rule name="Fiyat Takip 3000" dir=in action=allow protocol=TCP localport=3000
netsh advfirewall firewall add rule name="Node.js Fiyat" dir=in action=allow program="%ProgramFiles%\nodejs\node.exe" enable=yes
echo.
echo Kurallar eklendi! Arkadasin tekrar denesin.
pause
