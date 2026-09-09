# ============================================
# Fiyat Takip - Canlı Sunucudan Otomatik Yedek
# Fly.io'daki gerçek veriyi API üzerinden çeker
# Her akşam çalıştırılır
#
# Kullanim: .\yedekle.ps1 -sunucu "https://esamoda.fly.dev" -parola "admin123"
# ============================================

param(
  [string]$sunucu = "",
  [string]$kullanici = "admin",
  [string]$parola = ""
)

if (-not $sunucu) {
  Write-Host "HATA: -sunucu parametresi zorunludur. Ornek: .\yedekle.ps1 -sunucu https://esamoda.fly.dev -parola admin123"
  exit 1
}
if (-not $parola) {
  Write-Host "HATA: -parola parametresi zorunludur."
  exit 1
}

$yedekKlasor = "$PSScriptRoot\..\fiyat-takip-yedekler"
$tarih = Get-Date -Format "yyyy-MM-dd"
$saat  = Get-Date -Format "HH-mm"
$hedefKlasor = "$yedekKlasor\$tarih"

Write-Host "=== Fiyat Takip Canlı Yedekleme - $tarih $saat ==="

# Yedek klasörünü oluştur
if (-not (Test-Path $hedefKlasor)) {
    New-Item -ItemType Directory -Path $hedefKlasor -Force | Out-Null
}

# --- 1. Giriş yap, token al ---
Write-Host "Sunucuya bağlanılıyor..."
try {
    $loginBody = @{ username = $kullanici; password = $parola } | ConvertTo-Json
    $loginRes = Invoke-RestMethod -Uri "$sunucu/api/login" -Method Post -Body $loginBody -ContentType "application/json" -ErrorAction Stop
    $token = $loginRes.token
    $headers = @{ "x-auth-token" = $token }
    Write-Host "  OK: Giriş başarılı"
} catch {
    Write-Host "  HATA: Sunucuya bağlanılamadı! $_"
    exit 1
}

# --- 2. Veritabanı yedekle (.db dosyası) ---
Write-Host "Veritabanı indiriliyor..."
try {
    $dbDosya = "$hedefKlasor\fiyatlar.db"
    Invoke-WebRequest -Uri "$sunucu/api/yedekle/db" -Headers $headers -OutFile $dbDosya -ErrorAction Stop
    $dbBoyut = (Get-Item $dbDosya).Length
    Write-Host "  OK: fiyatlar.db ($([math]::Round($dbBoyut/1024,1)) KB)"
} catch {
    Write-Host "  HATA: Veritabanı indirilemedi: $_"
}

# --- 3. Excel yedek (tüm tablolar) ---
Write-Host "Excel yedek indiriliyor..."
try {
    $excelDosya = "$hedefKlasor\fiyat-takip-$tarih.xlsx"
    Invoke-WebRequest -Uri "$sunucu/api/yedekle/excel" -Headers $headers -OutFile $excelDosya -ErrorAction Stop
    Write-Host "  OK: fiyat-takip-$tarih.xlsx"
} catch {
    Write-Host "  HATA: Excel indirilemedi: $_"
}

# --- 4. Uploads klasörünü yedekle (lokalden - sadece varsa) ---
$uploadsKaynak = "$PSScriptRoot\uploads"
if (Test-Path $uploadsKaynak) {
    $uploadsBoyut = (Get-ChildItem $uploadsKaynak -Recurse | Measure-Object -Property Length -Sum).Sum
    Write-Host "Uploads klasörü yedekleniyor ($([math]::Round($uploadsBoyut/1024/1024,1)) MB)..."
    Copy-Item $uploadsKaynak "$hedefKlasor\uploads" -Recurse -Force
    Write-Host "  OK: uploads klasörü"
}

# --- 5. Ayar dosyalarını yedekle ---
$ayarDosyalar = @("fly.toml", "package.json", "Dockerfile", ".dockerignore", ".gitignore", "server.js")
foreach ($dosya in $ayarDosyalar) {
    $kaynak = "$PSScriptRoot\$dosya"
    if (Test-Path $kaynak) {
        Copy-Item $kaynak $hedefKlasor -Force
        Write-Host "  OK: $dosya"
    }
}

# --- 6. Son 7 günden eski yedekleri sil ---
Write-Host "Eski yedekler temizleniyor..."
$sonTarih = (Get-Date).AddDays(-7)
Get-ChildItem $yedekKlasor -Directory | ForEach-Object {
    try {
        $klasorTarih = [DateTime]::ParseExact($_.Name, "yyyy-MM-dd", $null)
        if ($klasorTarih -lt $sonTarih) {
            Remove-Item $_.FullName -Recurse -Force
            Write-Host "  Silindi: $($_.Name)"
        }
    } catch {
        # Tarih formatında olmayan klasörler atlanır
    }
}

Write-Host "=== Yedekleme tamamlandı: $hedefKlasor ==="

# Günlük dosyasına kaydet
"$tarih $saat - Canlı sunucudan yedekleme başarılı -> $hedefKlasor" | Out-File "$yedekKlasor\yedekleme.log" -Append -Encoding utf8
