# ============================================
# Fiyat Takip - Otomatik Yedekleme Scripti
# Her akşam çalıştırılır
# ============================================

$projeYolu  = "C:\Users\win\OneDrive\Desktop\claude\fiyat-takip"
$yedekKlasor = "C:\Users\win\OneDrive\Desktop\claude\fiyat-takip-yedekler"
$tarih = Get-Date -Format "yyyy-MM-dd"
$hedefKlasor = "$yedekKlasor\$tarih"

# Yedek klasörünü oluştur
if (-not (Test-Path $hedefKlasor)) {
    New-Item -ItemType Directory -Path $hedefKlasor -Force | Out-Null
}

Write-Host "=== Fiyat Takip Yedekleme - $tarih ==="
Write-Host "Hedef: $hedefKlasor"

# --- 1. Veritabanı yedekle ---
Write-Host "Veritabanı yedekleniyor..."
$dbDosyalar = @("fiyatlar.db", "fiyatlar.db-shm", "fiyatlar.db-wal")
foreach ($dosya in $dbDosyalar) {
    $kaynak = "$projeYolu\$dosya"
    if (Test-Path $kaynak) {
        Copy-Item $kaynak $hedefKlasor -Force
        Write-Host "  OK: $dosya"
    } else {
        Write-Host "  YOK: $dosya (atlandı)"
    }
}

# --- 2. Uploads klasörünü yedekle ---
$uploadsKaynak = "$projeYolu\uploads"
if (Test-Path $uploadsKaynak) {
    Copy-Item $uploadsKaynak "$hedefKlasor\uploads" -Recurse -Force
    Write-Host "  OK: uploads klasörü"
} else {
    Write-Host "  YOK: uploads klasörü (atlandı)"
}

# --- 3. Ayar dosyalarını yedekle ---
$ayarDosyalar = @("fly.toml", "package.json", "Dockerfile", ".dockerignore", ".gitignore")
foreach ($dosya in $ayarDosyalar) {
    $kaynak = "$projeYolu\$dosya"
    if (Test-Path $kaynak) {
        Copy-Item $kaynak $hedefKlasor -Force
        Write-Host "  OK: $dosya"
    }
}

# --- 4. Excel export ---
Write-Host "Excel dosyası oluşturuluyor..."
try {
    $nodeOutput = & node "$projeYolu\excel_yedekle.js" 2>&1
    Write-Host $nodeOutput
} catch {
    Write-Host "  HATA: Excel oluşturulamadı - $_"
}

# --- 5. Son 7 günden eski yedekleri sil ---
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
        # Tarih formatında olmayan klasörler atlanır (örn: log dosyası)
    }
}

Write-Host "=== Yedekleme tamamlandı: $hedefKlasor ==="

# Günlük dosyasına kaydet
"$tarih - Yedekleme başarıyla tamamlandı -> $hedefKlasor" | Out-File "$yedekKlasor\yedekleme.log" -Append -Encoding utf8
