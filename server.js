const express = require('express');
const Database = require('better-sqlite3');
const path = require('path');
const multer = require('multer');
const fs = require('fs');
const ExcelJS = require('exceljs');

const app = express();
const PORT = 3000;

// --- uploads klasörü ---
const uploadsDir = path.join(__dirname, 'public', 'uploads');
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });

// --- multer (görsel yükleme) ---
const storage = multer.diskStorage({
  destination: uploadsDir,
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname);
    const name = Date.now() + '-' + Math.round(Math.random() * 1e9) + ext;
    cb(null, name);
  }
});
const upload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB
  fileFilter: (req, file, cb) => {
    const izinli = /\.(jpg|jpeg|png|gif|webp|bmp)$/i;
    if (izinli.test(path.extname(file.originalname))) {
      cb(null, true);
    } else {
      cb(new Error('Sadece resim dosyaları (jpg, png, gif, webp, bmp) yüklenebilir.'));
    }
  }
});

// Excel yükleme için ayrı multer
const excelUpload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const izinli = /\.(xlsx|xls)$/i;
    if (izinli.test(path.extname(file.originalname))) {
      cb(null, true);
    } else {
      cb(new Error('Sadece Excel dosyaları (.xlsx, .xls) yüklenebilir.'));
    }
  }
});

// hem resim hem excel kabul eden multer (fiyat formu için)
const karmaUpload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname);
    const resim = /\.(jpg|jpeg|png|gif|webp|bmp)$/i;
    const excel = /\.(xlsx|xls)$/i;
    if (resim.test(ext) || excel.test(ext)) {
      cb(null, true);
    } else {
      cb(new Error('Sadece resim (jpg,png,gif,webp,bmp) ve Excel (.xlsx,.xls) dosyaları yüklenebilir.'));
    }
  }
});

// --- veritabanı ---
const dbPath = process.env.DB_PATH || 'fiyatlar.db';
const dbDir = path.dirname(dbPath);
if (dbDir !== '.' && !fs.existsSync(dbDir)) fs.mkdirSync(dbDir, { recursive: true });
const db = new Database(dbPath);
db.pragma('journal_mode = WAL');

db.exec(`
  CREATE TABLE IF NOT EXISTS fiyatlar (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    musteri_adi TEXT NOT NULL,
    firma_adi TEXT DEFAULT '',
    urun_aciklamasi TEXT NOT NULL,
    fiyat REAL NOT NULL,
    para_birimi TEXT DEFAULT 'TL',
    miktar INTEGER DEFAULT 1,
    tarih TEXT NOT NULL DEFAULT (date('now')),
    notlar TEXT DEFAULT '',
    gorsel TEXT DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
  )
`);

db.exec(`
  CREATE TABLE IF NOT EXISTS siparisler (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    musteri_adi TEXT NOT NULL,
    firma_adi TEXT DEFAULT '',
    urun_aciklamasi TEXT NOT NULL,
    fiyat REAL NOT NULL,
    para_birimi TEXT DEFAULT 'TL',
    miktar INTEGER DEFAULT 1,
    tarih TEXT NOT NULL,
    notlar TEXT DEFAULT '',
    gorsel TEXT DEFAULT '',
    durum TEXT DEFAULT 'Hazırlanıyor',
    siparis_tarihi TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
  )
`);

// migration: eski tabloda gorsel sütunu yoksa ekle
try { db.exec('ALTER TABLE fiyatlar ADD COLUMN gorsel TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE fiyatlar ADD COLUMN excel_dosya TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
// migration: siparisler tablosuna durum sütunu
try { db.exec('ALTER TABLE siparisler ADD COLUMN durum TEXT DEFAULT \'Hazırlanıyor\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE siparisler ADD COLUMN siparis_tarihi TEXT NOT NULL DEFAULT (datetime(\'now\',\'localtime\'))'); } catch (e) { /* zaten var */ }
// migration: termin ve renk sütunları
try { db.exec('ALTER TABLE siparisler ADD COLUMN termin TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE siparisler ADD COLUMN renk TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE siparisler ADD COLUMN renk_detay TEXT DEFAULT \'[]\''); } catch (e) { /* zaten var */ }

// --- gerçekleşen (maliyet) tablosu ---
db.exec(`
  CREATE TABLE IF NOT EXISTS gerceklesen (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    musteri_adi TEXT NOT NULL,
    firma_adi TEXT DEFAULT '',
    urun_aciklamasi TEXT NOT NULL,
    fiyat REAL NOT NULL,
    para_birimi TEXT DEFAULT 'TL',
    miktar INTEGER DEFAULT 1,
    gorsel TEXT DEFAULT '',
    kesim_adedi INTEGER DEFAULT 0,
    yuklenen_adet INTEGER DEFAULT 0,
    ikinci_kalite INTEGER DEFAULT 0,
    kumas_bedeli REAL DEFAULT 0,
    aksesuar_bedeli REAL DEFAULT 0,
    iscilik_bedeli REAL DEFAULT 0,
    siparis_id INTEGER,
    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
  )
`);
// migration: gerceklesen tablosu yeni sütunlar
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN ikinci_kalite INTEGER DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN satis_toplam REAL DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN termin TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN excel_dosya TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }

// middleware
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// --- API ---

// tüm fiyatları listele
app.get('/api/fiyatlar', (req, res) => {
  const rows = db.prepare('SELECT * FROM fiyatlar ORDER BY created_at DESC').all();
  res.json(rows);
});

// tek fiyat getir
app.get('/api/fiyatlar/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM fiyatlar WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ hata: 'Kayıt bulunamadı' });
  res.json(row);
});

// yeni fiyat ekle (görselli + excel)
const cpUpload = karmaUpload.fields([{ name: 'gorsel', maxCount: 1 }, { name: 'excel', maxCount: 1 }]);
app.post('/api/fiyatlar', cpUpload, (req, res) => {
  const { musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, tarih, notlar } = req.body;

  if (!musteri_adi || !urun_aciklamasi || fiyat == null) {
    if (req.files) Object.values(req.files).flat().forEach(f => fs.unlinkSync(f.path));
    return res.status(400).json({ hata: 'Müşteri adı, ürün açıklaması ve fiyat zorunludur' });
  }

  const gorselYolu = req.files && req.files.gorsel ? 'uploads/' + req.files.gorsel[0].filename : '';
  let excelYolu = '';
  if (req.files && req.files.excel) {
    excelYolu = 'uploads/' + req.files.excel[0].filename;
  }

  const stmt = db.prepare(`
    INSERT INTO fiyatlar (musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, tarih, notlar, gorsel, excel_dosya)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const info = stmt.run(
    musteri_adi,
    firma_adi || '',
    urun_aciklamasi,
    fiyat,
    para_birimi || 'TL',
    miktar || 1,
    tarih || new Date().toISOString().split('T')[0],
    notlar || '',
    gorselYolu,
    excelYolu
  );

  res.status(201).json({ id: info.lastInsertRowid, mesaj: 'Fiyat kaydedildi' });
});

// fiyat güncelle (görsel + excel)
app.put('/api/fiyatlar/:id', cpUpload, (req, res) => {
  const { musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, tarih, notlar, gorsel_sil } = req.body;

  if (!musteri_adi || !urun_aciklamasi || fiyat == null) {
    if (req.files) Object.values(req.files).flat().forEach(f => fs.unlinkSync(f.path));
    return res.status(400).json({ hata: 'Müşteri adı, ürün açıklaması ve fiyat zorunludur' });
  }

  const eski = db.prepare('SELECT gorsel, excel_dosya FROM fiyatlar WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: 'Kayıt bulunamadı' });

  // görsel
  let gorselYolu;
  if (gorsel_sil === '1') {
    if (eski.gorsel) { const dosya = path.join(__dirname, 'public', eski.gorsel); if (fs.existsSync(dosya)) fs.unlinkSync(dosya); }
    if (req.files && req.files.gorsel) fs.unlinkSync(req.files.gorsel[0].path);
    gorselYolu = '';
  } else if (req.files && req.files.gorsel) {
    if (eski.gorsel) { const dosya = path.join(__dirname, 'public', eski.gorsel); if (fs.existsSync(dosya)) fs.unlinkSync(dosya); }
    gorselYolu = 'uploads/' + req.files.gorsel[0].filename;
  } else {
    gorselYolu = eski.gorsel || '';
  }

  // excel
  let excelYolu = eski.excel_dosya || '';
  if (req.files && req.files.excel) {
    if (eski.excel_dosya) { const dosya = path.join(__dirname, 'public', eski.excel_dosya); if (fs.existsSync(dosya)) fs.unlinkSync(dosya); }
    excelYolu = 'uploads/' + req.files.excel[0].filename;
  }

  db.prepare(`
    UPDATE fiyatlar
    SET musteri_adi=?, firma_adi=?, urun_aciklamasi=?, fiyat=?, para_birimi=?, miktar=?, tarih=?, notlar=?, gorsel=?, excel_dosya=?
    WHERE id=?
  `).run(
    musteri_adi,
    firma_adi || '',
    urun_aciklamasi,
    fiyat,
    para_birimi || 'TL',
    miktar || 1,
    tarih || new Date().toISOString().split('T')[0],
    notlar || '',
    gorselYolu,
    excelYolu,
    req.params.id
  );

  res.json({ mesaj: 'Güncellendi' });
});

// fiyat sil (görseli de sil)
app.delete('/api/fiyatlar/:id', (req, res) => {
  const row = db.prepare('SELECT gorsel FROM fiyatlar WHERE id = ?').get(req.params.id);
  if (row && row.gorsel) {
    const dosya = path.join(__dirname, 'public', row.gorsel);
    if (fs.existsSync(dosya)) fs.unlinkSync(dosya);
  }
  db.prepare('DELETE FROM fiyatlar WHERE id = ?').run(req.params.id);
  res.json({ mesaj: 'Silindi' });
});

// arama
app.get('/api/ara', (req, res) => {
  const q = `%${req.query.q || ''}%`;
  const rows = db.prepare(
    'SELECT * FROM fiyatlar WHERE musteri_adi LIKE ? OR firma_adi LIKE ? OR urun_aciklamasi LIKE ? ORDER BY created_at DESC'
  ).all(q, q, q);
  res.json(rows);
});

// maliyet excel'i indir / görüntüle
app.get('/api/fiyatlar/:id/excel', async (req, res) => {
  const row = db.prepare('SELECT * FROM fiyatlar WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ hata: 'Kayıt bulunamadı' });

  // kayıtlı excel varsa direkt onu gönder
  if (row.excel_dosya) {
    const dosya = path.join(__dirname, 'public', row.excel_dosya);
    if (fs.existsSync(dosya)) {
      return res.sendFile(dosya);
    }
  }

  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Maliyet');

  // --- stiller ---
  const headerStyle = { font: { bold: true, size: 14, color: { argb: 'FF1E293B' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { bottom: { style: 'medium', color: { argb: 'FF3B82F6' } } } };
  const sectionStyle = { font: { bold: true, size: 12, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } }, alignment: { horizontal: 'left' } };
  const labelStyle = { font: { bold: true, size: 11 }, alignment: { horizontal: 'right' } };
  const inputStyle = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFF0' } }, border: { bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } } }, alignment: { horizontal: 'center' } };
  const totalStyle = { font: { bold: true, size: 13, color: { argb: 'FF059669' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } }, alignment: { horizontal: 'center' } };
  const kurStyle = { font: { bold: true, size: 13, color: { argb: 'FFDC2626' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } }, alignment: { horizontal: 'center' } };

  ws.columns = [
    { key: 'A', width: 22 }, { key: 'B', width: 18 }, { key: 'C', width: 18 }, { key: 'D', width: 18 }, { key: 'E', width: 18 }
  ];

  // --- fiyat_id (4 farklı yerde yedekli - Excel silse bile kurtulur) ---
  // 1. gizli satır
  ws.getCell('A2').value = 'FIYAT_ID';
  ws.getCell('B2').value = row.id;
  ws.getCell('A2').font = { color: { argb: 'FFFFFFFF' }, size: 6 };
  ws.getCell('B2').font = { color: { argb: 'FFFFFFFF' }, size: 6 };
  ws.getRow(2).hidden = true;
  // 2. Keywords alanı
  if (!workbook.properties) workbook.properties = {};
  workbook.properties.keywords = 'fiyat_id:' + row.id;
  // 3. başlığa göm (ID:123)
  const titleText = 'MALİYET HESAPLAMA  #' + row.id;

  // --- başlık ---
  ws.mergeCells('A3:E3');
  ws.getCell('A3').value = titleText;
  ws.getCell('A3').font = { bold: true, size: 18, color: { argb: 'FF1E293B' } };
  ws.getCell('A3').alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(3).height = 35;

  // 4. C2 hücresine görünür ama küçük yazıyla (gizli satır değil, silinmez)
  ws.getCell('C2').value = 'ID:' + row.id;
  ws.getCell('C2').font = { color: { argb: 'FF94A3B8' }, size: 8 };
  ws.getCell('C2').alignment = { horizontal: 'right' };

  // --- sipariş bilgileri ---
  ws.mergeCells('A5:E5');
  ws.getCell('A5').value = '📋 SİPARİŞ BİLGİLERİ';
  ['A5'].forEach(c => { ws.getCell(c).font = sectionStyle.font; ws.getCell(c).fill = sectionStyle.fill; ws.getCell(c).alignment = sectionStyle.alignment; });
  ws.getRow(5).height = 24;

  const info = [
    ['Müşteri', row.musteri_adi, '', 'Firma', row.firma_adi],
    ['Ürün', row.urun_aciklamasi, '', 'Tarih', row.tarih],
    ['Birim Fiyat', row.fiyat + ' ' + row.para_birimi, '', 'Miktar', row.miktar],
    ['Toplam Satış', (row.fiyat * row.miktar) + ' ' + row.para_birimi, '', 'Para Birimi', row.para_birimi],
  ];

  info.forEach((r, i) => {
    const rn = 6 + i;
    ws.getCell(`A${rn}`).value = r[0]; ws.getCell(`A${rn}`).font = labelStyle.font;
    ws.getCell(`B${rn}`).value = r[1]; ws.getCell(`B${rn}`).fill = inputStyle.fill; ws.getCell(`B${rn}`).border = inputStyle.border;
    ws.getCell(`D${rn}`).value = r[3]; ws.getCell(`D${rn}`).font = labelStyle.font;
    ws.getCell(`E${rn}`).value = r[4]; ws.getCell(`E${rn}`).fill = inputStyle.fill; ws.getCell(`E${rn}`).border = inputStyle.border;
  });

  // --- üretim bilgileri ---
  const uretimRow = 11;
  ws.mergeCells(`A${uretimRow}:E${uretimRow}`);
  ws.getCell(`A${uretimRow}`).value = '✂️ ÜRETİM BİLGİLERİ';
  ['A' + uretimRow].forEach(c => { ws.getCell(c).font = sectionStyle.font; ws.getCell(c).fill = sectionStyle.fill; ws.getCell(c).alignment = sectionStyle.alignment; });
  ws.getRow(uretimRow).height = 24;

  const uretim = [
    ['Kesim Adedi', '', '', 'Yüklenen Adet', ''],
    ['2. Kalite', '', '', 'Fire / Hata', ''],
  ];

  uretim.forEach((r, i) => {
    const rn = uretimRow + 1 + i;
    ws.getCell(`A${rn}`).value = r[0]; ws.getCell(`A${rn}`).font = labelStyle.font;
    ws.getCell(`B${rn}`).fill = inputStyle.fill; ws.getCell(`B${rn}`).border = inputStyle.border;
    ws.getCell(`D${rn}`).value = r[3]; ws.getCell(`D${rn}`).font = labelStyle.font;
    ws.getCell(`E${rn}`).fill = inputStyle.fill; ws.getCell(`E${rn}`).border = inputStyle.border;
  });

  // --- maliyet kalemleri ---
  const maliyetRow = 14;
  ws.mergeCells(`A${maliyetRow}:E${maliyetRow}`);
  ws.getCell(`A${maliyetRow}`).value = '💰 MALİYET KALEMLERİ';
  ['A' + maliyetRow].forEach(c => { ws.getCell(c).font = sectionStyle.font; ws.getCell(c).fill = sectionStyle.fill; ws.getCell(c).alignment = sectionStyle.alignment; });
  ws.getRow(maliyetRow).height = 24;

  const maliyetK = [
    ['Kumaş Bedeli (TL)', '', '', 'Aksesuar Bedeli (TL)', ''],
    ['İşçilik Bedeli (TL)', '', '', 'Diğer Giderler (TL)', ''],
  ];

  maliyetK.forEach((r, i) => {
    const rn = maliyetRow + 1 + i;
    ws.getCell(`A${rn}`).value = r[0]; ws.getCell(`A${rn}`).font = labelStyle.font;
    ws.getCell(`B${rn}`).fill = inputStyle.fill; ws.getCell(`B${rn}`).border = inputStyle.border;
    ws.getCell(`D${rn}`).value = r[3]; ws.getCell(`D${rn}`).font = labelStyle.font;
    ws.getCell(`E${rn}`).fill = inputStyle.fill; ws.getCell(`E${rn}`).border = inputStyle.border;
  });

  // --- toplam / kâr ---
  const sonucRow = 17;
  ws.mergeCells(`A${sonucRow}:E${sonucRow}`);
  ws.getCell(`A${sonucRow}`).value = '📊 SONUÇ';
  ['A' + sonucRow].forEach(c => { ws.getCell(c).font = sectionStyle.font; ws.getCell(c).fill = sectionStyle.fill; ws.getCell(c).alignment = sectionStyle.alignment; });
  ws.getRow(sonucRow).height = 24;

  // formüller
  ws.getCell('A18').value = 'Toplam Maliyet'; ws.getCell('A18').font = labelStyle.font;
  ws.getCell('B18').value = { formula: 'B15+B16+D15+D16', result: 0 }; ws.getCell('B18').fill = totalStyle.fill; ws.getCell('B18').font = totalStyle.font; ws.getCell('B18').alignment = totalStyle.alignment;

  ws.mergeCells('A19:E19');
  ws.getCell('A19').value = { formula: `"SATIŞ: "&B9&" | MALİYET: "&B18&" | KÂR: "&(VALUE(LEFT(B9,LEN(B9)-3))-B18)&" TL | KÂR %: "&ROUND((VALUE(LEFT(B9,LEN(B9)-3))-B18)/VALUE(LEFT(B9,LEN(B9)-3))*100,1)&"%"`, result: '' };
  ws.getCell('A19').font = { bold: true, size: 12, color: { argb: 'FF1E293B' } };
  ws.getCell('A19').alignment = { horizontal: 'center' };

  // --- notlar ---
  ws.mergeCells('A21:E21');
  ws.getCell('A21').value = '📝 NOTLAR';
  ['A21'].forEach(c => { ws.getCell(c).font = sectionStyle.font; ws.getCell(c).fill = sectionStyle.fill; ws.getCell(c).alignment = sectionStyle.alignment; });
  ws.mergeCells('A22:E24');
  ws.getCell('A22').fill = inputStyle.fill;

  if (row.notlar) {
    ws.getCell('A22').value = row.notlar;
  }

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="maliyet_${row.musteri_adi.replace(/\s+/g,'_')}_${row.urun_aciklamasi.replace(/\s+/g,'_')}_${row.id}.xlsx"`);
  await workbook.xlsx.write(res);
  res.end();
});

// maliyet excel'i yükle ve gerçekleşene aktar
app.post('/api/gerceklesen/excel', excelUpload.single('excel'), async (req, res) => {
  if (!req.file) return res.status(400).json({ hata: 'Excel dosyası zorunludur' });

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(req.file.path);

    // Maliyet sayfasını bul (yoksa ilk sayfayı kullan)
    let ws = workbook.getWorksheet('Maliyet');
    if (!ws) {
      // Excel bazen sayfa adını değiştirebiliyor, ilk sayfayı dene
      ws = workbook.worksheets[0];
    }
    if (!ws) {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ hata: 'Excel dosyasında hiç sayfa bulunamadı' });
    }

    // fiyat_id'yi bul (4 yöntem dene)
    let fiyatId = null;

    // Yöntem 1: B2 hücresi (gizli satır)
    fiyatId = parseInt(ws.getCell('B2').value);

    // Yöntem 2: C2 hücresi "ID:123" formatında
    if (!fiyatId) {
      const c2 = String(ws.getCell('C2').value || '');
      const m = c2.match(/ID:(\d+)/);
      if (m) fiyatId = parseInt(m[1]);
    }

    // Yöntem 3: Başlıktan "#123" formatında
    if (!fiyatId) {
      const title = String(ws.getCell('A3').value || '');
      const m = title.match(/#(\d+)/);
      if (m) fiyatId = parseInt(m[1]);
    }

    // Yöntem 4: Keywords alanı
    if (!fiyatId) {
      try {
        const kw = workbook.properties.keywords || '';
        const m = kw.match(/fiyat_id:(\d+)/);
        if (m) fiyatId = parseInt(m[1]);
      } catch(e) {}
    }

    if (!fiyatId) {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ hata: 'Bu Excel sistem tarafından oluşturulmamış. Lütfen Fiyatlar sayfasındaki 📥 butonuyla indirilen Excel\'i kullanın.' });
    }

    // fiyat kaydını bul
    const fiyat = db.prepare('SELECT * FROM fiyatlar WHERE id = ?').get(fiyatId);
    if (!fiyat) {
      fs.unlinkSync(req.file.path);
      return res.status(404).json({ hata: 'Fiyat kaydı bulunamadı. Önce fiyatı tekrar ekleyin.' });
    }

    // excel'den değerleri oku (sayısal hücreleri parse et)
    const getNum = (cell) => {
      const v = ws.getCell(cell).value;
      if (v === null || v === undefined || v === '') return 0;
      if (typeof v === 'number') return v;
      // "5000 TL" gibi string'lerden sayıyı çıkar
      const num = parseFloat(String(v).replace(/[^0-9.,\-]/g, '').replace(',', '.'));
      return isNaN(num) ? 0 : num;
    };

    const data = {
      musteri_adi: fiyat.musteri_adi,
      firma_adi: fiyat.firma_adi,
      urun_aciklamasi: fiyat.urun_aciklamasi,
      fiyat: fiyat.fiyat,
      para_birimi: fiyat.para_birimi,
      miktar: fiyat.miktar,
      gorsel: fiyat.gorsel,
      kesim_adedi: getNum('B12'),
      yuklenen_adet: getNum('E12'),
      ikinci_kalite: getNum('B13'),
      kumas_bedeli: getNum('B15'),
      aksesuar_bedeli: getNum('D15'),
      iscilik_bedeli: getNum('B16'),
      satis_toplam: getNum('B9'),
      siparis_id: fiyatId,
    };

    // fiyat kaydını sil
    db.prepare('DELETE FROM fiyatlar WHERE id = ?').run(fiyatId);

    // gerçekleşene ekle
    const stmt = db.prepare(`
      INSERT INTO gerceklesen (musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, gorsel,
        kesim_adedi, yuklenen_adet, ikinci_kalite, kumas_bedeli, aksesuar_bedeli, iscilik_bedeli, satis_toplam, siparis_id, excel_dosya)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

    // excel dosyasını uploads klasörüne taşı (herkes erişebilsin)
    const ext = path.extname(req.file.originalname);
    const excelAdi = 'excel_' + Date.now() + '_' + fiyatId + ext;
    const hedefYol = path.join(uploadsDir, excelAdi);
    fs.renameSync(req.file.path, hedefYol);
    const excelYolu = 'uploads/' + excelAdi;

    const info = stmt.run(
      data.musteri_adi, data.firma_adi, data.urun_aciklamasi, data.fiyat, data.para_birimi, data.miktar, data.gorsel,
      data.kesim_adedi, data.yuklenen_adet, data.ikinci_kalite, data.kumas_bedeli, data.aksesuar_bedeli, data.iscilik_bedeli, data.satis_toplam, data.siparis_id, excelYolu
    );

    res.status(201).json({ id: info.lastInsertRowid, excel_url: excelYolu, mesaj: 'Excel yüklendi, gerçekleşene aktarıldı' });
  } catch (e) {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.status(500).json({ hata: 'Excel işlenemedi: ' + e.message });
  }
});

// --- SİPARİŞ API ---

// tüm siparişleri listele
app.get('/api/siparisler', (req, res) => {
  const rows = db.prepare('SELECT * FROM siparisler ORDER BY created_at DESC').all();
  res.json(rows);
});

// tek sipariş getir
app.get('/api/siparisler/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM siparisler WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ hata: 'Sipariş bulunamadı' });
  res.json(row);
});

// fiyat kaydını siparişe aktar (fiyattan sil, siparişe ekle)
app.post('/api/siparisler', (req, res) => {
  const { fiyat_id } = req.body;
  if (!fiyat_id) return res.status(400).json({ hata: 'fiyat_id zorunludur' });

  const fiyat = db.prepare('SELECT * FROM fiyatlar WHERE id = ?').get(fiyat_id);
  if (!fiyat) return res.status(404).json({ hata: 'Fiyat kaydı bulunamadı' });

  const stmt = db.prepare(`
    INSERT INTO siparisler (musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, tarih, notlar, gorsel)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const info = stmt.run(
    fiyat.musteri_adi,
    fiyat.firma_adi,
    fiyat.urun_aciklamasi,
    fiyat.fiyat,
    fiyat.para_birimi,
    fiyat.miktar,
    fiyat.tarih,
    fiyat.notlar,
    fiyat.gorsel
  );

  // fiyat kaydını sil (görsel dosyasına dokunma - siparişte de aynı görsel kullanılacak)
  db.prepare('DELETE FROM fiyatlar WHERE id = ?').run(fiyat_id);

  res.status(201).json({ id: info.lastInsertRowid, mesaj: 'Siparişe aktarıldı' });
});

// sipariş güncelle (durum, not vb.)
app.put('/api/siparisler/:id', upload.single('gorsel'), (req, res) => {
  const { musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, tarih, notlar, durum, termin, renk, renk_detay } = req.body;

  if (!musteri_adi || !urun_aciklamasi || fiyat == null) {
    if (req.file) fs.unlinkSync(req.file.path);
    return res.status(400).json({ hata: 'Müşteri adı, ürün açıklaması ve fiyat zorunludur' });
  }

  const eski = db.prepare('SELECT gorsel FROM siparisler WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: 'Sipariş bulunamadı' });

  let gorselYolu = eski.gorsel || '';

  if (req.file) {
    gorselYolu = 'uploads/' + req.file.filename;
  }

  db.prepare(`
    UPDATE siparisler
    SET musteri_adi=?, firma_adi=?, urun_aciklamasi=?, fiyat=?, para_birimi=?, miktar=?, tarih=?, notlar=?, durum=?, gorsel=?, termin=?, renk=?, renk_detay=?
    WHERE id=?
  `).run(
    musteri_adi,
    firma_adi || '',
    urun_aciklamasi,
    fiyat,
    para_birimi || 'TL',
    miktar || 1,
    tarih || new Date().toISOString().split('T')[0],
    notlar || '',
    durum || 'Hazırlanıyor',
    gorselYolu,
    termin || '',
    renk || '',
    renk_detay || '[]',
    req.params.id
  );

  res.json({ mesaj: 'Sipariş güncellendi' });
});

// sipariş sil (görseli de sil)
app.delete('/api/siparisler/:id', (req, res) => {
  const row = db.prepare('SELECT gorsel FROM siparisler WHERE id = ?').get(req.params.id);
  if (row && row.gorsel) {
    const dosya = path.join(__dirname, 'public', row.gorsel);
    if (fs.existsSync(dosya)) fs.unlinkSync(dosya);
  }
  db.prepare('DELETE FROM siparisler WHERE id = ?').run(req.params.id);
  res.json({ mesaj: 'Sipariş silindi' });
});

// sipariş durumu güncelle (hızlı)
app.patch('/api/siparisler/:id/durum', (req, res) => {
  const { durum } = req.body;
  if (!durum) return res.status(400).json({ hata: 'durum zorunludur' });
  db.prepare('UPDATE siparisler SET durum=? WHERE id=?').run(durum, req.params.id);
  res.json({ mesaj: 'Durum güncellendi' });
});

// --- GERÇEKLEŞEN API ---

// tüm gerçekleşenleri listele
app.get('/api/gerceklesen', (req, res) => {
  const rows = db.prepare('SELECT * FROM gerceklesen ORDER BY created_at DESC').all();
  res.json(rows);
});

// tek gerçekleşen getir
app.get('/api/gerceklesen/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM gerceklesen WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ hata: 'Kayıt bulunamadı' });
  res.json(row);
});

// siparişi gerçekleşene aktar
app.post('/api/gerceklesen', (req, res) => {
  const { siparis_id } = req.body;
  if (!siparis_id) return res.status(400).json({ hata: 'siparis_id zorunludur' });

  const siparis = db.prepare('SELECT * FROM siparisler WHERE id = ?').get(siparis_id);
  if (!siparis) return res.status(404).json({ hata: 'Sipariş bulunamadı' });

  const stmt = db.prepare(`
    INSERT INTO gerceklesen (musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, gorsel, siparis_id)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const info = stmt.run(
    siparis.musteri_adi,
    siparis.firma_adi,
    siparis.urun_aciklamasi,
    siparis.fiyat,
    siparis.para_birimi,
    siparis.miktar,
    siparis.gorsel,
    siparis.id
  );

  // siparişi sil (görsel dosyasına dokunma)
  db.prepare('DELETE FROM siparisler WHERE id = ?').run(siparis_id);

  res.status(201).json({ id: info.lastInsertRowid, mesaj: 'Gerçekleşene aktarıldı' });
});

// gerçekleşen güncelle (maliyet kalemleri + excel yükleme)
app.put('/api/gerceklesen/:id', excelUpload.single('excel_dosya'), (req, res) => {
  const { musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar,
          kesim_adedi, yuklenen_adet, ikinci_kalite, kumas_bedeli, aksesuar_bedeli, iscilik_bedeli, satis_toplam, termin } = req.body;

  if (!musteri_adi || !urun_aciklamasi || fiyat == null) {
    return res.status(400).json({ hata: 'Müşteri adı, ürün açıklaması ve fiyat zorunludur' });
  }

  const eski = db.prepare('SELECT * FROM gerceklesen WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: 'Kayıt bulunamadı' });

  // yeni excel yüklendiyse kaydet
  let excelYolu = eski.excel_dosya || '';
  if (req.file) {
    const ext = path.extname(req.file.originalname);
    const excelAdi = 'excel_' + Date.now() + '_' + req.params.id + ext;
    const hedefYol = path.join(uploadsDir, excelAdi);
    fs.renameSync(req.file.path, hedefYol);
    excelYolu = 'uploads/' + excelAdi;
  }

  db.prepare(`
    UPDATE gerceklesen
    SET musteri_adi=?, firma_adi=?, urun_aciklamasi=?, fiyat=?, para_birimi=?, miktar=?,
        kesim_adedi=?, yuklenen_adet=?, ikinci_kalite=?, kumas_bedeli=?, aksesuar_bedeli=?, iscilik_bedeli=?, satis_toplam=?, termin=?, excel_dosya=?
    WHERE id=?
  `).run(
    musteri_adi,
    firma_adi || '',
    urun_aciklamasi,
    fiyat,
    para_birimi || 'TL',
    miktar || 1,
    kesim_adedi || 0,
    yuklenen_adet || 0,
    ikinci_kalite || 0,
    kumas_bedeli || 0,
    aksesuar_bedeli || 0,
    iscilik_bedeli || 0,
    satis_toplam || 0,
    termin || '',
    excelYolu,
    req.params.id
  );

  res.json({ mesaj: 'Gerçekleşen güncellendi' });
});

// gerçekleşen sil
app.delete('/api/gerceklesen/:id', (req, res) => {
  db.prepare('DELETE FROM gerceklesen WHERE id = ?').run(req.params.id);
  res.json({ mesaj: 'Silindi' });
});

// --- başlat ---
app.listen(PORT, () => {
  console.log(`✅ Fiyat Takip Sistemi çalışıyor → http://localhost:${PORT}`);
  console.log(`📡 Ağdaki diğer bilgisayarlar için → http://${getLocalIP()}:${PORT}`);
});

function getLocalIP() {
  const { networkInterfaces } = require('os');
  const nets = networkInterfaces();
  for (const name of Object.keys(nets)) {
    for (const net of nets[name]) {
      if (net.family === 'IPv4' && !net.internal) {
        return net.address;
      }
    }
  }
  return 'localhost';
}
