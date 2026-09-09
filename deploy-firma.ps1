# ============================================
# Fiyat Takip - Yeni Firma Deployment Scripti
# ============================================
param(
  [Parameter(Mandatory=$true)]
  [string]$FirmaKisaAd,              # ornek: "alptekin"

  [Parameter(Mandatory=$true)]
  [string]$FirmaTamAd,               # ornek: "Alp Tekin Tekstil"

  [string]$AdminSifre = "",          # bos birakilirsa rastgele uretilir
  [string]$Bolge = "ams"
)

$slug = $FirmaKisaAd.ToLower() -replace '[^a-z0-9-]','-'
$appUrl = "https://${slug}.fly.dev"
$volumeName = "${slug}_data"

# Rastgele sifre (varsayilan)
if (-not $AdminSifre) {
  $chars = "abcdefghkmnprstuvwxyzABCDEFGHKLMNPRSTUVWXYZ23456789"
  $AdminSifre = -join ((1..10) | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
}

Write-Host ""
Write-Host "============================================"
Write-Host "  Yeni Firma Deployment"
Write-Host "============================================"
Write-Host "  Slug      : $slug"
Write-Host "  Firma     : $FirmaTamAd"
Write-Host "  URL       : $appUrl"
Write-Host "  Volume    : $volumeName"
Write-Host "  Admin     : admin / $AdminSifre"
Write-Host "  Bolge     : $Bolge"
Write-Host "============================================"
Write-Host ""

$onay = Read-Host "Devam etmek icin EVET yazin"
if ($onay -ne "EVET") {
  Write-Host "Iptal edildi."
  exit 0
}

# --- 1. Fly.io uygulamasi olustur ---
Write-Host "`n[1/5] Fly.io uygulamasi olusturuluyor..."
fly apps create $slug --org personal
if ($LASTEXITCODE -ne 0) {
  Write-Host "HATA: Uygulama olusturulamadi."
  exit 1
}

# --- 2. Volume olustur ---
Write-Host "`n[2/5] Volume olusturuluyor..."
fly volumes create $volumeName --region $Bolge --size 1 --app $slug
if ($LASTEXITCODE -ne 0) {
  Write-Host "HATA: Volume olusturulamadi."
  exit 1
}

# --- 3. Env degiskenlerini ayarla ---
Write-Host "`n[3/5] Ortam degiskenleri ayarlaniyor..."
fly secrets set COMPANY_NAME="$FirmaTamAd" COMPANY_SLUG="$slug" APP_URL="$appUrl" ADMIN_DEFAULT_PASSWORD="$AdminSifre" --app $slug
if ($LASTEXITCODE -ne 0) {
  Write-Host "HATA: Env degiskenleri ayarlanamadi."
  exit 1
}

# --- 4. Deploy ---
Write-Host "`n[4/5] Kod deploy ediliyor..."
fly deploy --app $slug
if ($LASTEXITCODE -ne 0) {
  Write-Host "HATA: Deploy basarisiz."
  exit 1
}

# --- 5. Firma kaydini tut ---
Write-Host "`n[5/5] Firma kaydi olusturuluyor..."
$firmaKayit = @{
  firma = $FirmaTamAd
  slug = $slug
  url = $appUrl
  admin_sifre = $AdminSifre
  bolge = $Bolge
  tarih = (Get-Date -Format "yyyy-MM-dd HH:mm")
  durum = "Aktif"
}

$takipDosya = "$PSScriptRoot\..\firmalar.json"
$firmalar = @()
if (Test-Path $takipDosya) {
  try { $firmalar = Get-Content $takipDosya -Encoding UTF8 | ConvertFrom-Json } catch { $firmalar = @() }
  if (-not ($firmalar -is [Array])) { $firmalar = @($firmalar) }
}
$firmalar += $firmaKayit
$firmalar | ConvertTo-Json -Depth 3 | Set-Content $takipDosya -Encoding UTF8
Write-Host "  OK: $takipDosya"

# --- 6. Sonuc ---
Write-Host ""
Write-Host "============================================"
Write-Host "  DEPLOY BASARILI!"
Write-Host "============================================"
Write-Host "  URL       : $appUrl"
Write-Host "  Kullanici : admin"
Write-Host "  Sifre     : $AdminSifre"
Write-Host "============================================"
Write-Host ""
Write-Host "ONEMLI: Firma logosunu public/logo.png olarak"
Write-Host "        degistirip tekrar deploy etmeyi unutma:"
Write-Host "        fly deploy --app $slug"
Write-Host ""
