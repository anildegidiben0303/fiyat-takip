// ============================================
// Fiyat Takip - Veritabanını Excel'e Aktarma
// ============================================

const Database = require('better-sqlite3');
const ExcelJS = require('exceljs');
const path = require('path');

const DB_PATH = path.join(__dirname, 'fiyatlar.db');
const CIKTI_KLASOR = path.join(__dirname, '..', 'fiyat-takip-yedekler');
const TARIH = new Date().toISOString().split('T')[0];

async function excelAktar() {
  const db = new Database(DB_PATH, { readonly: true });
  const workbook = new ExcelJS.Workbook();

  // --- Tablo tanımları ---
  const tablolar = [
    {
      name: 'Fiyatlar',
      query: 'SELECT * FROM fiyatlar ORDER BY created_at DESC',
      columns: [
        { header: 'ID', key: 'id', width: 6 },
        { header: 'Müşteri Adı', key: 'musteri_adi', width: 20 },
        { header: 'Firma Adı', key: 'firma_adi', width: 20 },
        { header: 'Ürün Açıklaması', key: 'urun_aciklamasi', width: 30 },
        { header: 'Fiyat', key: 'fiyat', width: 12 },
        { header: 'Para Birimi', key: 'para_birimi', width: 10 },
        { header: 'Miktar', key: 'miktar', width: 8 },
        { header: 'Tarih', key: 'tarih', width: 12 },
        { header: 'Notlar', key: 'notlar', width: 25 },
        { header: 'Görsel', key: 'gorsel', width: 20 },
        { header: 'Excel Dosya', key: 'excel_dosya', width: 20 },
        { header: 'Kayıt Tarihi', key: 'created_at', width: 16 },
      ],
    },
    {
      name: 'Siparişler',
      query: 'SELECT * FROM siparisler ORDER BY created_at DESC',
      columns: [
        { header: 'ID', key: 'id', width: 6 },
        { header: 'Müşteri Adı', key: 'musteri_adi', width: 20 },
        { header: 'Firma Adı', key: 'firma_adi', width: 20 },
        { header: 'Ürün Açıklaması', key: 'urun_aciklamasi', width: 30 },
        { header: 'Fiyat', key: 'fiyat', width: 12 },
        { header: 'Para Birimi', key: 'para_birimi', width: 10 },
        { header: 'Miktar', key: 'miktar', width: 8 },
        { header: 'Tarih', key: 'tarih', width: 12 },
        { header: 'Notlar', key: 'notlar', width: 25 },
        { header: 'Görsel', key: 'gorsel', width: 20 },
        { header: 'Durum', key: 'durum', width: 14 },
        { header: 'Sipariş Tarihi', key: 'siparis_tarihi', width: 16 },
        { header: 'Termin', key: 'termin', width: 12 },
        { header: 'Renk', key: 'renk', width: 15 },
        { header: 'Renk Detay', key: 'renk_detay', width: 20 },
        { header: 'Kayıt Tarihi', key: 'created_at', width: 16 },
      ],
    },
    {
      name: 'Gerçekleşen',
      query: 'SELECT * FROM gerceklesen ORDER BY created_at DESC',
      columns: [
        { header: 'ID', key: 'id', width: 6 },
        { header: 'Müşteri Adı', key: 'musteri_adi', width: 20 },
        { header: 'Firma Adı', key: 'firma_adi', width: 20 },
        { header: 'Ürün Açıklaması', key: 'urun_aciklamasi', width: 30 },
        { header: 'Fiyat', key: 'fiyat', width: 12 },
        { header: 'Para Birimi', key: 'para_birimi', width: 10 },
        { header: 'Miktar', key: 'miktar', width: 8 },
        { header: 'Görsel', key: 'gorsel', width: 20 },
        { header: 'Kesim Adedi', key: 'kesim_adedi', width: 12 },
        { header: 'Yüklenen Adet', key: 'yuklenen_adet', width: 12 },
        { header: 'İkinci Kalite', key: 'ikinci_kalite', width: 12 },
        { header: 'Kumaş Bedeli', key: 'kumas_bedeli', width: 14 },
        { header: 'Aksesuar Bedeli', key: 'aksesuar_bedeli', width: 14 },
        { header: 'İşçilik Bedeli', key: 'iscilik_bedeli', width: 14 },
        { header: 'Satış Toplam', key: 'satis_toplam', width: 14 },
        { header: 'Termin', key: 'termin', width: 12 },
        { header: 'Excel Dosya', key: 'excel_dosya', width: 20 },
        { header: 'Sipariş ID', key: 'siparis_id', width: 10 },
        { header: 'Kayıt Tarihi', key: 'created_at', width: 16 },
      ],
    },
    {
      name: 'Ödemeler',
      query: 'SELECT * FROM odemeler ORDER BY created_at DESC',
      columns: [
        { header: 'ID', key: 'id', width: 6 },
        { header: 'Müşteri Adı', key: 'musteri_adi', width: 20 },
        { header: 'Firma Adı', key: 'firma_adi', width: 20 },
        { header: 'Ödeme Türü', key: 'odeme_turu', width: 14 },
        { header: 'Tutar', key: 'tutar', width: 12 },
        { header: 'Para Birimi', key: 'para_birimi', width: 10 },
        { header: 'Açıklama', key: 'aciklama', width: 30 },
        { header: 'Durum', key: 'durum', width: 12 },
        { header: 'Vade Tarihi', key: 'vade_tarihi', width: 12 },
        { header: 'Çek No', key: 'cek_no', width: 14 },
        { header: 'Banka', key: 'banka', width: 16 },
        { header: 'Çek Sahibi', key: 'cek_sahibi', width: 20 },
        { header: 'Kredi Grup ID', key: 'kredi_grup_id', width: 16 },
        { header: 'Kayıt Tarihi', key: 'created_at', width: 16 },
      ],
    },
    {
      name: 'Aksesuar Takip',
      query: 'SELECT * FROM aksesuar_takip ORDER BY siparis_id, created_at ASC',
      columns: [
        { header: 'ID', key: 'id', width: 6 },
        { header: 'Sipariş ID', key: 'siparis_id', width: 10 },
        { header: 'Aksesuar Adı', key: 'aksesuar_adi', width: 24 },
        { header: 'Sipariş Adet', key: 'siparis_adet', width: 12 },
        { header: 'Birim Fiyat', key: 'birim_fiyat', width: 12 },
        { header: 'Gelen Adet', key: 'gelen_adet', width: 12 },
        { header: 'Gelen Detay', key: 'gelen_detay', width: 30 },
        { header: 'Notlar', key: 'notlar', width: 25 },
        { header: 'Fatura', key: 'fatura', width: 16 },
        { header: 'Kayıt Tarihi', key: 'created_at', width: 16 },
      ],
    },
    {
      name: 'Kumaş Takip',
      query: 'SELECT * FROM kumas_takip ORDER BY siparis_id, created_at ASC',
      columns: [
        { header: 'ID', key: 'id', width: 6 },
        { header: 'Sipariş ID', key: 'siparis_id', width: 10 },
        { header: 'Kumaş Adı', key: 'kumas_adi', width: 20 },
        { header: 'Renk', key: 'renk', width: 14 },
        { header: 'Adet', key: 'adet', width: 8 },
        { header: 'Pastal Boyu', key: 'pastal_boyu', width: 12 },
        { header: 'Kumaş En', key: 'kumas_en', width: 10 },
        { header: 'Gramaj', key: 'kumas_gramaj', width: 10 },
        { header: 'Fire %', key: 'fire', width: 8 },
        { header: 'Asorti', key: 'asorti', width: 8 },
        { header: 'Biye', key: 'biye', width: 8 },
        { header: 'Birim Gider', key: 'birim_gider', width: 12 },
        { header: 'Kumaş Fiyatı', key: 'kumas_fiyati', width: 12 },
        { header: 'Kesim Adedi', key: 'kesim_adedi', width: 10 },
        { header: 'Kesim %', key: 'kesim_yuzdesi', width: 8 },
        { header: 'Gerç. Birim', key: 'gerceklesen_birim', width: 12 },
        { header: 'Gelen Detay', key: 'gelen_detay', width: 30 },
        { header: 'Notlar', key: 'notlar', width: 20 },
        { header: 'Kayıt Tarihi', key: 'created_at', width: 16 },
      ],
    },
    {
      name: 'Kullanıcılar',
      query: 'SELECT id, username, rol, permissions, created_at FROM kullanicilar ORDER BY id ASC',
      columns: [
        { header: 'ID', key: 'id', width: 6 },
        { header: 'Kullanıcı Adı', key: 'username', width: 18 },
        { header: 'Rol', key: 'rol', width: 14 },
        { header: 'Yetkiler', key: 'permissions', width: 30 },
        { header: 'Kayıt Tarihi', key: 'created_at', width: 16 },
      ],
    },
  ];

  // --- Her tablo için ayrı sheet oluştur ---
  for (const tablo of tablolar) {
    const sheet = workbook.addWorksheet(tablo.name);
    const rows = db.prepare(tablo.query).all();
    const toplam = rows.length;

    // Başlık satırı (mavi arka plan, beyaz yazı)
    const baslikRow = sheet.addRow(tablo.columns.map(c => c.header));
    baslikRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    baslikRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF2563EB' },
    };
    baslikRow.alignment = { horizontal: 'center', vertical: 'middle' };

    // Veri satırları
    for (const row of rows) {
      const degerler = tablo.columns.map(c => {
        let val = row[c.key];
        if (val === null || val === undefined) return '';
        return val;
      });
      sheet.addRow(degerler);
    }

    // Sütun genişlikleri
    tablo.columns.forEach((c, i) => {
      sheet.getColumn(i + 1).width = c.width;
    });

    // İlk satırı dondur
    sheet.views = [{ state: 'frozen', ySplit: 1 }];

    // Otomatik filtre
    const sonKolonHarf = String.fromCharCode(64 + tablo.columns.length);
    sheet.autoFilter = `A1:${sonKolonHarf}${toplam + 1}`;

    console.log(`  ✅ ${tablo.name}: ${toplam} kayıt`);
  }

  // --- Kaydet ---
  const fs = require('fs');
  if (!fs.existsSync(CIKTI_KLASOR)) {
    fs.mkdirSync(CIKTI_KLASOR, { recursive: true });
  }

  const dosyaYolu = path.join(CIKTI_KLASOR, `fiyat-takip-${TARIH}.xlsx`);
  await workbook.xlsx.writeFile(dosyaYolu);

  // Aynı Excel'i proje klasörüne de kopyala (güncel Excel her zaman erişilebilir olsun)
  const guncelYol = path.join(__dirname, 'fiyat-takip-guncel.xlsx');
  await workbook.xlsx.writeFile(guncelYol);

  db.close();

  console.log(`\n📊 Excel dosyası oluşturuldu:`);
  console.log(`   ${dosyaYolu}`);
  console.log(`   ${guncelYol}`);
  return dosyaYolu;
}

excelAktar().catch(err => {
  console.error('❌ Hata:', err.message);
  process.exit(1);
});
