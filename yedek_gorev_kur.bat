@echo off
:: Fiyat Takip Sistemi - Günlük Yedekleme Görevi
:: Bu dosyaya sağ tıkla → Yönetici olarak çalıştır
::
:: Her firma için SUNUCU ve PAROLA değiştirilmeli!
set SUNUCU=https://esamoda.fly.dev
set PAROLA=admin123
set SCRIPT=%~dp0yedekle.ps1

echo === %SUNUCU% için yedekleme görevi kuruluyor ===
schtasks /create /tn "FiyatTakipYedekleme_esamoda" /tr "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File \"%SCRIPT%\" -sunucu %SUNUCU% -parola %PAROLA%" /sc daily /st 21:00 /rl HIGHEST /f
echo.
echo === Görev bilgisi ===
schtasks /query /tn "FiyatTakipYedekleme_esamoda" /fo LIST
echo.
echo Yedekleme görevi her gün 21:00'de çalışacak.
echo.
echo Baska bir firma icin: SUNUCU ve PAROLA degiskenlerini guncelleyip tekrar calistirin.
pause
