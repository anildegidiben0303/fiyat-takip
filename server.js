const express = require('express');
const Database = require('better-sqlite3');
const path = require('path');
const multer = require('multer');
const fs = require('fs');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const { spawn } = require('child_process');
const os = require('os');
const crypto = require('crypto');
const { t } = require('./server-i18n');
const swaggerUi = require('swagger-ui-express');

const app = express();
const PORT = process.env.PORT || 3000;
const COMPANY_NAME = process.env.COMPANY_NAME || 'Fiyat Takip';
const COMPANY_SLUG = process.env.COMPANY_SLUG || 'fiyat-takip';
const APP_URL = process.env.APP_URL || ('http://localhost:' + PORT);
const ADMIN_DEFAULT_PASSWORD = process.env.ADMIN_DEFAULT_PASSWORD || 'admin123';

// dil tespiti
function getLang(req) {
  return req.query.lang || ((req.headers['accept-language'] || '').startsWith('en') ? 'en' : 'tr');
}

// --- çok kullanıcılı auth sistemi ---

const ROLE_PERMISSIONS = {
  admin: ['fiyatlar','siparisler','gerceklesen','odemeler','bant','kullanicilar','gelir_takip','cari','faturalab'],
  muhasebe: ['fiyatlar','gerceklesen','odemeler','gelir_takip','cari','faturalab'],
  operasyon: ['siparisler','gerceklesen','bant'],
  satis: ['fiyatlar','siparisler']
};

function hashPassword(password) {
  return crypto.createHash('sha256').update(password).digest('hex');
}

function getSession(token) {
  if (!token) return null;
  const row = db.prepare('SELECT user_id, username, rol, permissions FROM sessions WHERE token = ?').get(token);
  if (!row) return null;
  return { userId: row.user_id, username: row.username, rol: row.rol, permissions: JSON.parse(row.permissions || '[]') };
}

function authMiddleware(req, res, next) {
  const token = req.headers['x-auth-token'] || req.query.token || '';
  const user = getSession(token);
  if (user) { req.user = user; return next(); }
  res.status(401).json({ hata: t(getLang(req), 'sessionInvalid') });
}

function requirePermission(permission) {
  return (req, res, next) => {
    const token = req.headers['x-auth-token'] || req.query.token || '';
    const user = getSession(token);
    if (!user) return res.status(401).json({ hata: t(getLang(req), 'sessionInvalid') });
    if (!user.permissions || !user.permissions.includes(permission)) {
      return res.status(403).json({ hata: t(getLang(req), 'accessDenied') });
    }
    req.user = user;
    next();
  };
}

// --- uploads klasörü ---
const dataDir = process.env.DB_PATH ? path.dirname(process.env.DB_PATH) : __dirname;
const uploadsDir = path.join(dataDir, 'uploads');
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });
const pubDir = path.join(__dirname, 'public', 'uploads');
if (!fs.existsSync(pubDir)) fs.mkdirSync(pubDir, { recursive: true });

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
      cb(new Error(t(getLang(req), 'imageOnly')));
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
      cb(new Error(t(getLang(req), 'excelOnly')));
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
      cb(new Error(t(getLang(req), 'fileTypeError')));
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
try { db.exec('ALTER TABLE siparisler ADD COLUMN excel_dosya TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE siparisler ADD COLUMN kumas_tamam TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
// migration: fit / ld_test / bulk_test / gold_test
try { db.exec('ALTER TABLE siparisler ADD COLUMN fit TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE siparisler ADD COLUMN ld_test TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE siparisler ADD COLUMN bulk_test TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE siparisler ADD COLUMN gold_test TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE siparisler ADD COLUMN termin_detay TEXT DEFAULT \'[]\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE siparisler ADD COLUMN mt TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE siparisler ADD COLUMN aksesuar TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }

// --- MT (takım eden) kişiler tablosu ---
db.exec(`
  CREATE TABLE IF NOT EXISTS mt_kisiler (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    isim TEXT NOT NULL UNIQUE
  )
`);
// varsayılan MT kişileri
['ANIL', 'NURCAN'].forEach(isim => {
  db.prepare('INSERT OR IGNORE INTO mt_kisiler (isim) VALUES (?)').run(isim);
});

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
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN durum TEXT DEFAULT \'Devam Ediyor\''); } catch (e) { /* zaten var */ }
// migration: sub-item detay JSON sütunları
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN yuklenen_adet_detay TEXT DEFAULT \'[]\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN satis_toplam_detay TEXT DEFAULT \'[]\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN kumas_bedeli_detay TEXT DEFAULT \'[]\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN aksesuar_bedeli_detay TEXT DEFAULT \'[]\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN iscilik_bedeli_detay TEXT DEFAULT \'[]\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN baski_nakis_yikama REAL DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN baski_nakis_yikama_detay TEXT DEFAULT \'[]\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN paketleme_bedeli REAL DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN nakliye_bedeli REAL DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN paketleme_cari_id INTEGER'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN nakliye_cari_id INTEGER'); } catch (e) { /* zaten var */ }

// --- ödeme planı tablosu ---
db.exec(`
  CREATE TABLE IF NOT EXISTS odemeler (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    musteri_adi TEXT NOT NULL,
    firma_adi TEXT DEFAULT '',
    odeme_turu TEXT NOT NULL DEFAULT 'Rotatif Kredi',
    tutar REAL NOT NULL,
    para_birimi TEXT DEFAULT 'TL',
    vade_tarihi TEXT NOT NULL,
    aciklama TEXT DEFAULT '',
    durum TEXT DEFAULT 'Bekliyor',
    cek_no TEXT DEFAULT '',
    banka TEXT DEFAULT '',
    cek_sahibi TEXT DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
  )
`);
// migration: odemeler tablosu yeni sütunlar
try { db.exec('ALTER TABLE odemeler ADD COLUMN cek_no TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE odemeler ADD COLUMN banka TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE odemeler ADD COLUMN cek_sahibi TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE odemeler ADD COLUMN kredi_grup_id TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE odemeler ADD COLUMN ana_para REAL DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE odemeler ADD COLUMN faiz_tutari REAL DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE odemeler ADD COLUMN bsmv_tutari REAL DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE odemeler ADD COLUMN tarih TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE odemeler ADD COLUMN cari_id INTEGER'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE odemeler ADD COLUMN rotatif_ay INTEGER'); } catch (e) { /* zaten var */ }
// migration: "Banka Kredisi" tanımı → "Taksitli Kredi"
try { db.exec("UPDATE odemeler SET odeme_turu = 'Taksitli Kredi' WHERE odeme_turu = 'Banka Kredisi'"); } catch (e) { /* zaten uygulanmış */ }

// --- aksesuar takip tablosu ---
db.exec(`
  CREATE TABLE IF NOT EXISTS aksesuar_takip (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    siparis_id INTEGER NOT NULL,
    aksesuar_adi TEXT NOT NULL,
    siparis_adet INTEGER DEFAULT 0,
    gelen_adet INTEGER DEFAULT 0,
    notlar TEXT DEFAULT '',
    fatura TEXT DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
  )
`);

// migration: aksesuar çoklu sevkiyat detay
try { db.exec('ALTER TABLE aksesuar_takip ADD COLUMN gelen_detay TEXT DEFAULT \'[]\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE aksesuar_takip ADD COLUMN birim_fiyat REAL DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE aksesuar_takip ADD COLUMN renk TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }

// --- kumaş takip tablosu ---
db.exec(`
  CREATE TABLE IF NOT EXISTS kumas_takip (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    siparis_id INTEGER NOT NULL,
    kumas_adi TEXT NOT NULL DEFAULT '',
    renk TEXT DEFAULT '',
    adet INTEGER DEFAULT 0,
    pastal_boyu REAL DEFAULT 0,
    kumas_en REAL DEFAULT 0,
    kumas_gramaj REAL DEFAULT 0,
    fire REAL DEFAULT 0,
    asorti REAL DEFAULT 1,
    biye REAL DEFAULT 0,
    birim_gider REAL DEFAULT 0,
    notlar TEXT DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
  )
`);
// migration: renk ve adet sütunları
try { db.exec('ALTER TABLE kumas_takip ADD COLUMN renk TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE kumas_takip ADD COLUMN adet INTEGER DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE kumas_takip ADD COLUMN gelen_detay TEXT DEFAULT \'[]\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE kumas_takip ADD COLUMN kumas_fiyati REAL DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE kumas_takip ADD COLUMN kesim_adedi INTEGER DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE kumas_takip ADD COLUMN kesim_yuzdesi REAL DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE kumas_takip ADD COLUMN gerceklesen_birim REAL DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE kumas_takip ADD COLUMN fit_birim REAL DEFAULT 0'); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE kumas_takip ADD COLUMN stok REAL DEFAULT 0'); } catch (e) { /* zaten var */ }

// --- gelir takip tablosu ---
db.exec(`
  CREATE TABLE IF NOT EXISTS gelir_takip (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    siparis_id INTEGER NOT NULL,
    musteri_adi TEXT NOT NULL,
    firma_adi TEXT DEFAULT '',
    urun_aciklamasi TEXT NOT NULL,
    toplam_tutar REAL NOT NULL DEFAULT 0,
    para_birimi TEXT DEFAULT 'TL',
    nakit_tutar REAL DEFAULT 0,
    cek_toplam REAL DEFAULT 0,
    cek_detay TEXT DEFAULT '[]',
    durum TEXT DEFAULT 'Bekliyor',
    notlar TEXT DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
  )
`);
try { db.exec('ALTER TABLE gelir_takip ADD COLUMN durum TEXT DEFAULT \'Bekliyor\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gelir_takip ADD COLUMN notlar TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gelir_takip ADD COLUMN odeme_turu TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gelir_takip ADD COLUMN tarih TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gelir_takip ADD COLUMN tfs_detay TEXT DEFAULT \'[]\''); } catch (e) { /* zaten var */ }

// mevcut admin/muhasebe kullanıcılarına gelir_takip yetkisi ekle
try {
  const usersToUpdate = db.prepare("SELECT id, permissions FROM kullanicilar WHERE (rol = 'admin' OR rol = 'muhasebe') AND (permissions IS NULL OR permissions NOT LIKE '%gelir_takip%')").all();
  usersToUpdate.forEach(user => {
    let perms = [];
    try { perms = JSON.parse(user.permissions || '[]'); } catch(e) { perms = []; }
    if (!perms.includes('gelir_takip')) {
      perms.push('gelir_takip');
      db.prepare('UPDATE kullanicilar SET permissions = ? WHERE id = ?').run(JSON.stringify(perms), user.id);
    }
  });
} catch (e) { /* ignore */ }

// --- cari kartlar tablosu (tedarikçi hesapları) ---
db.exec(`
  CREATE TABLE IF NOT EXISTS cari_kartlar (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    hesap_turu TEXT DEFAULT 'Tedarikçi',
    cari_adi TEXT NOT NULL,
    vergi_no TEXT DEFAULT '',
    telefon TEXT DEFAULT '',
    email TEXT DEFAULT '',
    vade TEXT DEFAULT '',
    adres TEXT DEFAULT '',
    notlar TEXT DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
  )
`);

try { db.exec('ALTER TABLE cari_kartlar ADD COLUMN hesap_turu TEXT DEFAULT \'Tedarikçi\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE cari_kartlar ADD COLUMN vade TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }

// --- cari hareketler tablosu (borç/alacak) ---
db.exec(`
  CREATE TABLE IF NOT EXISTS cari_hareketler (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cari_id INTEGER NOT NULL,
    tarih TEXT NOT NULL,
    tur TEXT NOT NULL DEFAULT 'borc',
    tutar REAL NOT NULL DEFAULT 0,
    para_birimi TEXT DEFAULT 'TL',
    aciklama TEXT DEFAULT '',
    belge_no TEXT DEFAULT '',
    odeme_sekli TEXT DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
  )
`);
try { db.exec('ALTER TABLE cari_hareketler ADD COLUMN para_birimi TEXT DEFAULT \'TL\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE cari_hareketler ADD COLUMN belge_no TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE cari_hareketler ADD COLUMN odeme_sekli TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE cari_hareketler ADD COLUMN kaynak TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE cari_hareketler ADD COLUMN vade TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE cari_hareketler ADD COLUMN iskonto_detay TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE cari_hareketler ADD COLUMN durum TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }

// mevcut admin/muhasebe kullanıcılarına cari yetkisi ekle
try {
  const cariUsers = db.prepare("SELECT id, permissions FROM kullanicilar WHERE (rol = 'admin' OR rol = 'muhasebe') AND (permissions IS NULL OR permissions NOT LIKE '%cari%')").all();
  cariUsers.forEach(user => {
    let perms = [];
    try { perms = JSON.parse(user.permissions || '[]'); } catch(e) { perms = []; }
    if (!perms.includes('cari')) {
      perms.push('cari');
      db.prepare('UPDATE kullanicilar SET permissions = ? WHERE id = ?').run(JSON.stringify(perms), user.id);
    }
  });
} catch (e) { /* ignore */ }

// mevcut admin/muhasebe kullanıcılarına faturalab yetkisi ekle
try {
  const flUsers = db.prepare("SELECT id, permissions FROM kullanicilar WHERE (rol = 'admin' OR rol = 'muhasebe') AND (permissions IS NULL OR permissions NOT LIKE '%faturalab%')").all();
  flUsers.forEach(user => {
    let perms = [];
    try { perms = JSON.parse(user.permissions || '[]'); } catch(e) { perms = []; }
    if (!perms.includes('faturalab')) {
      perms.push('faturalab');
      db.prepare('UPDATE kullanicilar SET permissions = ? WHERE id = ?').run(JSON.stringify(perms), user.id);
    }
  });
} catch (e) { /* ignore */ }

// --- ayarlar tablosu (global birim fiyatlar vs) ---
db.exec(`
  CREATE TABLE IF NOT EXISTS ayarlar (
    anahtar TEXT PRIMARY KEY,
    deger TEXT NOT NULL DEFAULT ''
  )
`);
// varsayılanları ekle
db.prepare("INSERT OR IGNORE INTO ayarlar (anahtar, deger) VALUES ('paketleme_birim_fiyat', '2.65')").run();
db.prepare("INSERT OR IGNORE INTO ayarlar (anahtar, deger) VALUES ('nakliye_birim_fiyat', '2.00')").run();

// --- paketleme malzemeleri (stok + tüketim) ---
db.exec(`
  CREATE TABLE IF NOT EXISTS paketleme_malzemeleri (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    adi TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
  )
`);
db.exec(`
  CREATE TABLE IF NOT EXISTS paketleme_alimlar (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    malzeme_id INTEGER NOT NULL,
    adet REAL NOT NULL DEFAULT 0,
    toplam_tutar REAL NOT NULL DEFAULT 0,
    birim_fiyat REAL DEFAULT 0,
    cari_id INTEGER,
    tarih TEXT NOT NULL DEFAULT (date('now')),
    belge_no TEXT DEFAULT '',
    notlar TEXT DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
  )
`);
try { db.exec('ALTER TABLE paketleme_alimlar ADD COLUMN belge_no TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }
try { db.exec('ALTER TABLE gerceklesen ADD COLUMN paketleme_malzeme_detay TEXT DEFAULT \'[]\''); } catch (e) { /* zaten var */ }

// kullanıcılar tablosu
db.exec(`
  CREATE TABLE IF NOT EXISTS kullanicilar (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    rol TEXT NOT NULL DEFAULT 'satis',
    permissions TEXT DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
  )
`);
try { db.exec('ALTER TABLE kullanicilar ADD COLUMN permissions TEXT DEFAULT \'\''); } catch (e) { /* zaten var */ }

// oturum tablosu (kalıcı session)
db.exec(`
  CREATE TABLE IF NOT EXISTS sessions (
    token TEXT PRIMARY KEY,
    user_id INTEGER NOT NULL,
    username TEXT NOT NULL,
    rol TEXT NOT NULL,
    permissions TEXT DEFAULT '[]',
    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
  )
`);

// admin seed
const adminVar = db.prepare('SELECT id FROM kullanicilar WHERE username = ?').get('admin');
if (!adminVar) {
  const allPerms = JSON.stringify(ROLE_PERMISSIONS.admin);
  db.prepare('INSERT INTO kullanicilar (username, password_hash, rol, permissions) VALUES (?, ?, ?, ?)')
    .run('admin', hashPassword(ADMIN_DEFAULT_PASSWORD), 'admin', allPerms);
  console.log('✅ Varsayılan admin kullanıcısı oluşturuldu: admin');
}

// --- Cari senkronizasyon yardımcı fonksiyonu ---
// Kaynak bazlı cari hareketlerini senkronize eder (fatura/gider → borç hareketi).
// prefix: örn. 'kumas:12:' — entries: [{ kaynak, cari_id, tutar, aciklama, belge_no }]
function cariSenkron(prefix, entries) {
  const existing = db.prepare('SELECT * FROM cari_hareketler WHERE kaynak LIKE ?').all(prefix + '%');
  const byKaynak = {};
  existing.forEach(e => { byKaynak[e.kaynak] = e; });
  const seen = new Set();

  const bugun = new Date();
  const today = bugun.getFullYear() + '-' + ('0' + (bugun.getMonth() + 1)).slice(-2) + '-' + ('0' + bugun.getDate()).slice(-2);

  (entries || []).forEach(en => {
    const key = en.kaynak;
    seen.add(key);
    const tutar = Math.round((parseFloat(en.tutar) || 0) * 100) / 100;
    const cariId = en.cari_id ? parseInt(en.cari_id) : null;

    if (!cariId || tutar <= 0) {
      // cari seçilmemiş veya tutar yok → varsa sil
      if (byKaynak[key]) db.prepare('DELETE FROM cari_hareketler WHERE id = ?').run(byKaynak[key].id);
      return;
    }

    // cari kartı hâlâ var mı kontrol et (silinmiş cariye yazma)
    const kart = db.prepare('SELECT id FROM cari_kartlar WHERE id = ?').get(cariId);
    if (!kart) {
      if (byKaynak[key]) db.prepare('DELETE FROM cari_hareketler WHERE id = ?').run(byKaynak[key].id);
      return;
    }

    const tarih = en.tarih || today;
    const tur = en.tur || 'borc';
    const vade = en.vade || '';
    const odemeSekli = en.odeme_sekli || '';
    if (byKaynak[key]) {
      db.prepare(`UPDATE cari_hareketler SET cari_id=?, tarih=?, tur=?, tutar=?, aciklama=?, belge_no=?, vade=?, odeme_sekli=? WHERE id=?`)
        .run(cariId, tarih, tur, tutar, en.aciklama || '', en.belge_no || '', vade, odemeSekli, byKaynak[key].id);
    } else {
      db.prepare(`INSERT INTO cari_hareketler (cari_id, tarih, tur, tutar, para_birimi, aciklama, belge_no, odeme_sekli, vade, kaynak)
        VALUES (?, ?, ?, ?, 'TL', ?, ?, ?, ?, ?)`)
        .run(cariId, tarih, tur, tutar, en.aciklama || '', en.belge_no || '', odemeSekli, vade, key);
    }
  });

  // artık kaynağı olmayan hareketleri sil
  existing.forEach(e => {
    if (!seen.has(e.kaynak)) db.prepare('DELETE FROM cari_hareketler WHERE id = ?').run(e.id);
  });
}

// middleware
app.use(express.json());

// login/logout
app.post('/api/login', (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password) return res.status(401).json({ hata: t(getLang(req), 'userRequired') });

  const user = db.prepare('SELECT * FROM kullanicilar WHERE username = ?').get(username);
  if (!user || user.password_hash !== hashPassword(password)) {
    return res.status(401).json({ hata: t(getLang(req), 'sessionInvalid') });
  }

  let permissions;
  if (user.permissions) {
    try { permissions = JSON.parse(user.permissions); } catch(e) { permissions = ROLE_PERMISSIONS[user.rol] || []; }
  } else {
    permissions = ROLE_PERMISSIONS[user.rol] || [];
  }

  const token = 'tok_' + Date.now() + '_' + user.id + '_' + Math.random().toString(36).substr(2, 8);
  db.prepare('INSERT OR REPLACE INTO sessions (token, user_id, username, rol, permissions) VALUES (?, ?, ?, ?, ?)')
    .run(token, user.id, user.username, user.rol, JSON.stringify(permissions));

  res.json({ token, username: user.username, rol: user.rol, permissions });
});

app.post('/api/logout', (req, res) => {
  const token = req.headers['x-auth-token'] || '';
  db.prepare('DELETE FROM sessions WHERE token = ?').run(token);
  res.json({ mesaj: t(getLang(req), 'logoutOk') });
});

// public config (auth gerektirmez — login ekranı için)
app.get('/api/config', (req, res) => {
  res.json({
    companyName: COMPANY_NAME,
    companySlug: COMPANY_SLUG,
    appUrl: APP_URL
  });
});

// cari adlarını dropdown için döndür (cari yetkisi olmasa da sipariş/gerçekleşen kullanıcıları görebilir)
app.get('/api/cari-secenek', (req, res) => {
  const token = req.headers['x-auth-token'] || req.query.token || '';
  const user = getSession(token);
  if (!user) return res.status(401).json({ hata: t(getLang(req), 'sessionInvalid') });
  const rows = db.prepare('SELECT id, cari_adi, hesap_turu FROM cari_kartlar ORDER BY cari_adi ASC').all();
  res.json(rows);
});

app.use(express.static(path.join(__dirname, 'public')));
app.use('/uploads', express.static(uploadsDir));

// --- API Dökümantasyonu (Swagger) ---
const swaggerDoc = {
  openapi: '3.0.0',
  info: {
    title: COMPANY_NAME + ' API',
    version: '1.0.0',
    description: 'Fiyat Takip Sistemi (' + COMPANY_NAME + ') - Tüm endpoint\'ler **x-auth-token** header\'ı gerektirir (login hariç).'
  },
  servers: [{ url: APP_URL, description: 'Production' }],
  components: {
    securitySchemes: {
      token: { type: 'apiKey', in: 'header', name: 'x-auth-token', description: 'Login ile alınan token' }
    }
  },
  security: [{ token: [] }],
  paths: {
    '/api/login': {
      post: {
        tags: ['Auth'], summary: 'Giriş yap',
        requestBody: { required: true, content: { 'application/json': { schema: { type: 'object', properties: { username: { type: 'string', example: 'admin' }, password: { type: 'string', example: 'admin123' } } } } } },
        responses: { '200': { description: 'Token döner' } }
      }
    },
    '/api/logout': {
      post: { tags: ['Auth'], summary: 'Çıkış yap', responses: { '200': { description: 'Başarılı' } } }
    },
    '/api/fiyatlar': {
      get: { tags: ['Fiyatlar'], summary: 'Tüm fiyatları listele', responses: { '200': { description: 'Fiyat listesi' } } },
      post: { tags: ['Fiyatlar'], summary: 'Yeni fiyat ekle (multipart)', responses: { '201': { description: 'Oluşturuldu' } } }
    },
    '/api/fiyatlar/{id}': {
      get: { tags: ['Fiyatlar'], summary: 'Tek fiyat getir', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Fiyat detayı' } } },
      put: { tags: ['Fiyatlar'], summary: 'Fiyat güncelle (multipart)', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Güncellendi' } } },
      delete: { tags: ['Fiyatlar'], summary: 'Fiyat sil', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Silindi' } } }
    },
    '/api/fiyatlar/{id}/excel': {
      get: { tags: ['Fiyatlar'], summary: 'Maliyet Excel\'i indir', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Excel dosyası' } } }
    },
    '/api/siparisler': {
      get: { tags: ['Siparişler'], summary: 'Siparişleri listele', parameters: [{ name: 'bant', in: 'query', schema: { type: 'string' }, description: 'Bant sıralaması için 1' }], responses: { '200': { description: 'Sipariş listesi' } } },
      post: { tags: ['Siparişler'], summary: 'Fiyatı siparişe aktar', requestBody: { required: true, content: { 'application/json': { schema: { type: 'object', properties: { fiyat_id: { type: 'integer' } } } } } }, responses: { '201': { description: 'Aktarıldı' } } }
    },
    '/api/siparisler/{id}': {
      get: { tags: ['Siparişler'], summary: 'Tek sipariş getir', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Sipariş detayı' } } },
      put: { tags: ['Siparişler'], summary: 'Sipariş güncelle (multipart)', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Güncellendi' } } },
      delete: { tags: ['Siparişler'], summary: 'Sipariş sil', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Silindi' } } }
    },
    '/api/siparisler/{id}/durum': {
      patch: { tags: ['Siparişler'], summary: 'Durum güncelle', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], requestBody: { required: true, content: { 'application/json': { schema: { type: 'object', properties: { durum: { type: 'string' } } } } } }, responses: { '200': { description: 'Güncellendi' } } }
    },
    '/api/siparisler/{id}/aksesuar': {
      get: { tags: ['Aksesuar'], summary: 'Aksesuarları listele', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Liste' } } },
      post: { tags: ['Aksesuar'], summary: 'Aksesuar ekle', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '201': { description: 'Eklendi' } } }
    },
    '/api/siparisler/{id}/kumas': {
      get: { tags: ['Kumaş'], summary: 'Kumaşları listele', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Liste' } } },
      post: { tags: ['Kumaş'], summary: 'Kumaş ekle', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '201': { description: 'Eklendi' } } }
    },
    '/api/kumas/{id}': {
      get: { tags: ['Kumaş'], summary: 'Tek kumaş getir', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Detay' } } },
      put: { tags: ['Kumaş'], summary: 'Kumaş güncelle', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Güncellendi' } } },
      delete: { tags: ['Kumaş'], summary: 'Kumaş sil', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Silindi' } } }
    },
    '/api/kumas/{id}/siparis-formu': {
      get: { tags: ['Kumaş'], summary: 'Kumaş sipariş formu indir (Excel)', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Excel dosyası' } } }
    },
    '/api/gerceklesen': {
      get: { tags: ['Gerçekleşen'], summary: 'Gerçekleşenleri listele', responses: { '200': { description: 'Liste' } } },
      post: { tags: ['Gerçekleşen'], summary: 'Siparişi gerçekleşene aktar', requestBody: { required: true, content: { 'application/json': { schema: { type: 'object', properties: { siparis_id: { type: 'integer' } } } } } }, responses: { '201': { description: 'Aktarıldı' } } }
    },
    '/api/gerceklesen/{id}': {
      get: { tags: ['Gerçekleşen'], summary: 'Tek gerçekleşen getir', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Detay' } } },
      put: { tags: ['Gerçekleşen'], summary: 'Gerçekleşen güncelle (multipart)', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Güncellendi' } } },
      delete: { tags: ['Gerçekleşen'], summary: 'Gerçekleşen sil', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Silindi' } } }
    },
    '/api/gerceklesen/{id}/excel': {
      get: { tags: ['Gerçekleşen'], summary: 'Gerçekleşen maliyet Excel\'i indir', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Excel dosyası' } } }
    },
    '/api/gerceklesen/excel': {
      post: { tags: ['Gerçekleşen'], summary: 'Maliyet Excel\'i yükle (multipart)', responses: { '201': { description: 'Yüklendi' } } }
    },
    '/api/odemeler': {
      get: { tags: ['Ödemeler'], summary: 'Ödemeleri listele', parameters: [{ name: 'durum', in: 'query', schema: { type: 'string' } }, { name: 'odeme_turu', in: 'query', schema: { type: 'string' } }], responses: { '200': { description: 'Liste' } } },
      post: { tags: ['Ödemeler'], summary: 'Yeni ödeme ekle', responses: { '201': { description: 'Eklendi' } } }
    },
    '/api/odemeler/{id}': {
      get: { tags: ['Ödemeler'], summary: 'Tek ödeme getir', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Detay' } } },
      put: { tags: ['Ödemeler'], summary: 'Ödeme güncelle', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Güncellendi' } } },
      delete: { tags: ['Ödemeler'], summary: 'Ödeme sil', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Silindi' } } }
    },
    '/api/odemeler/{id}/durum': {
      patch: { tags: ['Ödemeler'], summary: 'Ödeme durumu güncelle', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], requestBody: { required: true, content: { 'application/json': { schema: { type: 'object', properties: { durum: { type: 'string' } } } } } }, responses: { '200': { description: 'Güncellendi' } } }
    },
    '/api/odemeler/haftalik': {
      get: { tags: ['Ödemeler'], summary: 'Haftalık ödeme planı', responses: { '200': { description: 'Haftalık plan' } } }
    },
    '/api/kullanicilar': {
      get: { tags: ['Kullanıcılar'], summary: 'Kullanıcıları listele (sadece admin)', responses: { '200': { description: 'Liste' } } },
      post: { tags: ['Kullanıcılar'], summary: 'Kullanıcı ekle (sadece admin)', responses: { '201': { description: 'Eklendi' } } }
    },
    '/api/kullanicilar/{id}': {
      put: { tags: ['Kullanıcılar'], summary: 'Kullanıcı güncelle (sadece admin)', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Güncellendi' } } },
      delete: { tags: ['Kullanıcılar'], summary: 'Kullanıcı sil (sadece admin)', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Silindi' } } }
    },
    '/api/ara': {
      get: { tags: ['Fiyatlar'], summary: 'Fiyatlarda arama', parameters: [{ name: 'q', in: 'query', required: true, schema: { type: 'string' } }], responses: { '200': { description: 'Sonuçlar' } } }
    },
    '/api/yedekle': {
      get: { tags: ['Sistem'], summary: 'Tam yedek (DB + uploads + bant) tar.gz', responses: { '200': { description: 'Yedek dosyası (tar.gz)' } } }
    },
    '/api/yedekle/db': {
      get: { tags: ['Sistem'], summary: 'Sadece veritabanı dosyası (.db)', responses: { '200': { description: 'SQLite veritabanı dosyası' } } }
    },
    '/api/yedekle/excel': {
      get: { tags: ['Sistem'], summary: 'Tüm tablolar Excel formatında yedek', responses: { '200': { description: 'Excel yedek dosyası (.xlsx)' } } }
    },
    '/api/gelir-takip': {
      get: { tags: ['Gelir Takip'], summary: 'Gelir takip listesi', parameters: [{ name: 'siparis_id', in: 'query', schema: { type: 'integer' } }], responses: { '200': { description: 'Liste' } } },
      post: { tags: ['Gelir Takip'], summary: 'Yeni gelir takip ekle', responses: { '201': { description: 'Eklendi' } } }
    },
    '/api/gelir-takip/ozet': {
      get: { tags: ['Gelir Takip'], summary: 'Gelir takip özeti', responses: { '200': { description: 'Özet istatistikler' } } }
    },
    '/api/gelir-takip/{id}': {
      get: { tags: ['Gelir Takip'], summary: 'Tek kayıt getir', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Detay' } } },
      put: { tags: ['Gelir Takip'], summary: 'Güncelle', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Güncellendi' } } },
      delete: { tags: ['Gelir Takip'], summary: 'Sil', parameters: [{ name: 'id', in: 'path', required: true, schema: { type: 'integer' } }], responses: { '200': { description: 'Silindi' } } }
    }
  }
};
app.use('/api-docs', swaggerUi.serve, swaggerUi.setup(swaggerDoc, { customCss: '.swagger-ui .topbar { display: none }' }));
app.get('/api-docs.json', (req, res) => res.json(swaggerDoc));

// auth middleware (login/logout hariç tüm /api/*)
app.use('/api', (req, res, next) => {
  if (req.path === '/login' || req.path === '/logout') return next();
  authMiddleware(req, res, next);
});

// permission middleware per route group
app.use('/api/fiyatlar', requirePermission('fiyatlar'));
app.use('/api/siparisler', requirePermission('siparisler'));
app.use('/api/gerceklesen', requirePermission('gerceklesen'));
app.use('/api/odemeler', requirePermission('odemeler'));
app.use('/api/kullanicilar', requirePermission('kullanicilar'));
app.use('/api/gelir-takip', requirePermission('gelir_takip'));
app.use('/api/cari', requirePermission('cari'));

// --- API ---

// tüm fiyatları listele
app.get('/api/fiyatlar', (req, res) => {
  const rows = db.prepare('SELECT * FROM fiyatlar ORDER BY created_at DESC').all();
  res.json(rows);
});

// tek fiyat getir
app.get('/api/fiyatlar/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM fiyatlar WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ hata: t(getLang(req), 'fiyatNotFound') });
  res.json(row);
});

// yeni fiyat ekle (görselli + excel)
const cpUpload = karmaUpload.fields([{ name: 'gorsel', maxCount: 1 }, { name: 'excel', maxCount: 1 }]);
app.post('/api/fiyatlar', cpUpload, (req, res) => {
  const { musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, tarih, notlar } = req.body;

  if (!musteri_adi || !urun_aciklamasi || fiyat == null) {
    if (req.files) Object.values(req.files).flat().forEach(f => fs.unlinkSync(f.path));
    return res.status(400).json({ hata: t(getLang(req), 'fiyatRequired') });
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

  res.status(201).json({ id: info.lastInsertRowid, mesaj: t(getLang(req), 'fiyatSaved') });
});

// fiyat güncelle (görsel + excel)
app.put('/api/fiyatlar/:id', cpUpload, (req, res) => {
  const { musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, tarih, notlar, gorsel_sil } = req.body;

  if (!musteri_adi || !urun_aciklamasi || fiyat == null) {
    if (req.files) Object.values(req.files).flat().forEach(f => fs.unlinkSync(f.path));
    return res.status(400).json({ hata: t(getLang(req), 'fiyatRequired') });
  }

  const eski = db.prepare('SELECT gorsel, excel_dosya FROM fiyatlar WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'fiyatNotFound') });

  // görsel
  let gorselYolu;
  if (gorsel_sil === '1') {
    if (eski.gorsel) { const dosya = path.join(uploadsDir, path.basename(eski.gorsel||'')); if (fs.existsSync(dosya)) fs.unlinkSync(dosya); }
    if (req.files && req.files.gorsel) fs.unlinkSync(req.files.gorsel[0].path);
    gorselYolu = '';
  } else if (req.files && req.files.gorsel) {
    if (eski.gorsel) { const dosya = path.join(uploadsDir, path.basename(eski.gorsel||'')); if (fs.existsSync(dosya)) fs.unlinkSync(dosya); }
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

  res.json({ mesaj: t(getLang(req), 'fiyatUpdated') });
});

// fiyat sil (görseli de sil)
app.delete('/api/fiyatlar/:id', (req, res) => {
  const row = db.prepare('SELECT gorsel FROM fiyatlar WHERE id = ?').get(req.params.id);
  if (row && row.gorsel) {
    const dosya = path.join(uploadsDir, path.basename(row.gorsel||''));
    if (fs.existsSync(dosya)) fs.unlinkSync(dosya);
  }
  db.prepare('DELETE FROM fiyatlar WHERE id = ?').run(req.params.id);
  res.json({ mesaj: t(getLang(req), 'fiyatDeleted') });
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
  if (!row) return res.status(404).json({ hata: t(getLang(req), 'fiyatNotFound') });

  // kayıtlı excel varsa direkt onu gönder
  if (row.excel_dosya) {
    const dosya = path.join(uploadsDir, path.basename(row.excel_dosya || ''));
    if (fs.existsSync(dosya)) {
      return res.download(dosya);
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

  // Türkçe karakterleri temizle
  const safeName = (s) => s.replace(/[^a-zA-Z0-9]/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="maliyet_${safeName(row.musteri_adi)}_${safeName(row.urun_aciklamasi)}_${row.id}.xlsx"`);
  await workbook.xlsx.write(res);
  res.end();
});

// maliyet excel'i yükle ve gerçekleşene aktar
app.post('/api/gerceklesen/excel', excelUpload.single('excel'), async (req, res) => {
  if (!req.file) return res.status(400).json({ hata: t(getLang(req), 'excelNotFound') });

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
      return res.status(400).json({ hata: t(getLang(req), 'excelInvalidSheet') });
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
      return res.status(400).json({ hata: t(getLang(req), 'excelNotSystem') });
    }

    // fiyat kaydını bul
    const fiyat = db.prepare('SELECT * FROM fiyatlar WHERE id = ?').get(fiyatId);
    if (!fiyat) {
      fs.unlinkSync(req.file.path);
      return res.status(404).json({ hata: t(getLang(req), 'fiyatNotFoundById') });
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

    res.status(201).json({ id: info.lastInsertRowid, excel_url: excelYolu, mesaj: t(getLang(req), 'excelUploaded') });
  } catch (e) {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.status(500).json({ hata: t(getLang(req), 'excelProcessError') + ': ' + e.message });
  }
});

// --- SİPARİŞ API ---

app.get('/api/siparisler', async (req, res) => {
  const rows = db.prepare('SELECT * FROM siparisler ORDER BY created_at DESC').all();
  if (req.query.bant === '1') {
    const bd = path.join(uploadsDir, 'bant.xlsx'); if (fs.existsSync(bd)) { try {
      const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(bd); const ws = wb.worksheets[0]; if (ws) {
      const mt = new Map(); ws.eachRow((row, rn) => { if (rn === 1) return; const dc = row.getCell(1).value; if (!dc) return;
        let rd; if (dc instanceof Date) rd = dc; else rd = new Date(dc); if (isNaN(rd.getTime())) return; const rts = rd.getTime();
        row.eachCell((cell, cn) => { if (cn === 1) return; const m = String(cell.value||'').match(/[A-Z]\d{3,4}[A-Z]\d{1,3}/g);
          if (m) m.forEach(c => { if (!mt.has(c) || rts < mt.get(c)) mt.set(c, rts); }); });
      });
      rows.forEach(r => { r._bantSira = null; const u = (r.urun_aciklamasi||'').toUpperCase(); let ek = Infinity;
        mt.forEach((ts, c) => { if (u.includes(c) && ts < ek) ek = ts; }); if (ek < Infinity) r._bantSira = ek; });
      rows.sort((a, b) => (a._bantSira===null?1:b._bantSira===null?-1:a._bantSira-b._bantSira));
    }} catch(e) {} }
  }
  res.json(rows);
});

// fiyat kaydını siparişe aktar (fiyattan sil, siparişe ekle)
app.post('/api/siparisler', (req, res) => {
  const { fiyat_id, termin, renk, miktar, renk_detay, fiyat, mt } = req.body;
  if (!fiyat_id) return res.status(400).json({ hata: t(getLang(req), 'fiyatIdRequired') });
  if (!fiyat) return res.status(400).json({ hata: 'Fiyat zorunludur' });

  const fiyatKaydi = db.prepare('SELECT * FROM fiyatlar WHERE id = ?').get(fiyat_id);
  if (!fiyatKaydi) return res.status(404).json({ hata: t(getLang(req), 'fiyatNotFoundById') });

  let yeniTermin = termin || '';

  // renk_detay işleme
  let yeniRenk = '';
  let finalRenkDetay = '[]';
  if (renk_detay) {
    finalRenkDetay = typeof renk_detay === 'string' ? renk_detay : JSON.stringify(renk_detay);
    try {
      const d = JSON.parse(finalRenkDetay);
      if (d.length > 0) {
        yeniRenk = d[0].renk || '';
        if (!yeniTermin) yeniTermin = d[0].termin || '';
      }
    } catch(e) {}
  } else if (renk) {
    yeniRenk = renk;
    finalRenkDetay = JSON.stringify([{ renk, adet: miktar || fiyatKaydi.miktar }]);
  }

  // toplam adedi renk detaylarından hesapla
  let toplamAdet = miktar || fiyatKaydi.miktar;
  if (renk_detay) {
    try {
      const d = JSON.parse(finalRenkDetay);
      toplamAdet = d.reduce((sum, x) => sum + (parseInt(x.adet) || 0), 0);
    } catch(e) {}
  }

  const terminDetay = req.body.termin_detay || '[]';

  const stmt = db.prepare(`
    INSERT INTO siparisler (musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, tarih, notlar, gorsel, excel_dosya, termin, renk, renk_detay, termin_detay, mt)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const info = stmt.run(
    fiyatKaydi.musteri_adi,
    fiyatKaydi.firma_adi,
    fiyatKaydi.urun_aciklamasi,
    fiyat,
    fiyatKaydi.para_birimi,
    toplamAdet,
    fiyatKaydi.tarih,
    fiyatKaydi.notlar,
    fiyatKaydi.gorsel,
    fiyatKaydi.excel_dosya || '',
    yeniTermin,
    yeniRenk,
    finalRenkDetay,
    terminDetay,
    mt || ''
  );

  // fiyat kaydını sil (görsel dosyasına dokunma - siparişte de aynı görsel kullanılacak)
  db.prepare('DELETE FROM fiyatlar WHERE id = ?').run(fiyat_id);

  res.status(201).json({ id: info.lastInsertRowid, mesaj: 'Siparişe aktarıldı' });
});

// --- BANT PROGRAMI ---
app.get('/api/siparisler/bant-excel', async (req, res) => {
  const bd = path.join(uploadsDir, 'bant.xlsx'); if (fs.existsSync(bd)) return res.download(bd);
  const rows = db.prepare('SELECT * FROM siparisler ORDER BY siparis_tarihi DESC').all();
  const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('Bant');
  ws.columns=[{header:'Sip.No',key:'id',width:8},{header:'Müşteri',key:'musteri',width:20},{header:'Firma',key:'firma',width:14},{header:'Model',key:'model',width:16},{header:'Renk',key:'renk',width:16},{header:'Adet',key:'adet',width:8},{header:'Birim Fiyat',key:'fiyat',width:12},{header:'Toplam',key:'toplam',width:14},{header:'Termin',key:'termin',width:12},{header:'Durum',key:'durum',width:14},{header:'Not',key:'not',width:20}];
  ws.getRow(1).eachCell(c=>{c.font={bold:true,size:11,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1E40AF'}};c.alignment={horizontal:'center',vertical:'middle'};c.border={top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}};});ws.getRow(1).height=22;
  const bg=new Date().toISOString().split('T')[0];let rn=2,brd={top:{style:'thin',color:{argb:'FFE2E8F0'}},bottom:{style:'thin',color:{argb:'FFE2E8F0'}},left:{style:'thin',color:{argb:'FFE2E8F0'}},right:{style:'thin',color:{argb:'FFE2E8F0'}}};
  rows.forEach(sip=>{const rd=[];try{const d=JSON.parse(sip.renk_detay||'[]');if(Array.isArray(d))rd.push(...d);}catch(e){}if(rd.length>0){rd.forEach((x,i)=>{const ad=parseInt(x.adet)||0;ws.addRow({id:i===0?sip.id:'',musteri:i===0?sip.musteri_adi:'',firma:i===0?sip.firma_adi:'',model:i===0?sip.urun_aciklamasi:'',renk:x.renk||'',adet:ad,fiyat:i===0?sip.fiyat+' '+sip.para_birimi:'',toplam:(sip.fiyat*ad).toFixed(2)+' '+sip.para_birimi,termin:i===0?(sip.termin||''):'',durum:i===0?sip.durum:'',not:i===0?(sip.notlar||''):''});if(sip.termin&&sip.termin<bg)ws.getRow(rn).eachCell(c=>{c.font={color:{argb:'FFDC2626'}};});rn++;});}else{ws.addRow({id:sip.id,musteri:sip.musteri_adi,firma:sip.firma_adi,model:sip.urun_aciklamasi,renk:sip.renk||'-',adet:sip.miktar,fiyat:sip.fiyat+' '+sip.para_birimi,toplam:(sip.fiyat*sip.miktar).toFixed(2)+' '+sip.para_birimi,termin:sip.termin||'',durum:sip.durum,not:sip.notlar||''});if(sip.termin&&sip.termin<bg)ws.getRow(rn).eachCell(c=>{c.font={color:{argb:'FFDC2626'}};});rn++;}});
  for(let r=2;r<rn;r++){ws.getRow(r).eachCell(c=>{c.border=brd;c.alignment={vertical:'middle'};});}
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');res.setHeader('Content-Disposition','attachment; filename="Bant_Plani.xlsx"');await wb.xlsx.write(res);res.end();
});
app.post('/api/siparisler/bant-excel',(req,res)=>{excelUpload.single('excel')(req,res,(err)=>{if(err)return res.status(400).json({hata:err.message||t(getLang(req), 'excelNoFile')});if(!req.file)return res.status(400).json({hata:t(getLang(req), 'excelNotSelected')});const h=path.join(uploadsDir,'bant.xlsx');if(fs.existsSync(h))fs.unlinkSync(h);fs.renameSync(req.file.path,h);res.json({mesaj:t(getLang(req), 'excelUploadedOk')});});});
app.delete('/api/siparisler/bant-excel',(req,res)=>{const h=path.join(uploadsDir,'bant.xlsx');if(fs.existsSync(h))fs.unlinkSync(h);res.json({mesaj:t(getLang(req), 'fiyatDeleted')});});
const srd=()=>path.join(uploadsDir,'bant_secenek_renk.json');
app.get('/api/siparisler/bant-secenek-renk',(req,res)=>{try{res.json(JSON.parse(fs.readFileSync(srd(),'utf8')));}catch(e){res.json({});}});
app.post('/api/siparisler/bant-secenek-renk',(req,res)=>{const{key,renk}=req.body;if(!key)return res.status(400).json({hata:t(getLang(req), 'bantRequired')});let d={};try{d=JSON.parse(fs.readFileSync(srd(),'utf8'));}catch(e){}d[key]=renk||'#dbeafe';fs.writeFileSync(srd(),JSON.stringify(d));res.json({mesaj:'ok'});});
const ad=()=>path.join(uploadsDir,'bant_atama.json');
const ao=()=>{try{return JSON.parse(fs.readFileSync(ad(),'utf8'));}catch(e){return{};}};
const ay=(d)=>fs.writeFileSync(ad(),JSON.stringify(d));
app.get('/api/siparisler/bant-atama',(req,res)=>{res.json(ao());});
app.post('/api/siparisler/bant-atama',(req,res)=>{
  const{key,siparis_id,urun,renk,adet,termin,hucreRenk,sure}=req.body;if(!key)return res.status(400).json({hata:t(getLang(req), 'bantRequired')});
  const atamalar=ao();const[sr,col]=key.split('-').map(Number);const sa=isNaN(parseInt(sure))?1:parseInt(sure);
  if(siparis_id||hucreRenk){const es=(atamalar[key]&&atamalar[key].sure)?atamalar[key].sure:0;const esf=es>0?Math.max(0,sa-es):sa;
    if(esf>0&&!isNaN(sr)&&!isNaN(col)){Object.entries(atamalar).filter(([k])=>{const p=k.split('-');return parseInt(p[1])===col&&parseInt(p[0])>sr;}).sort((a,b)=>parseInt(b[0].split('-')[0])-parseInt(a[0].split('-')[0])).forEach(([ok,ov])=>{const nr=parseInt(ok.split('-')[0])+esf;atamalar[nr+'-'+col]=ov;if(ok!==nr+'-'+col)delete atamalar[ok];});}
    atamalar[key]={siparis_id:siparis_id||0,urun:urun||'',renk:renk||'',adet:adet||0,termin:termin||'',hucreRenk:hucreRenk||'',sure:sa};}else{delete atamalar[key];}
  ay(atamalar);res.json({mesaj:t(getLang(req), 'bantSaved')});
});
app.post('/api/siparisler/bant-satir-ekle',(req,res)=>{
  const{satir,adet,kolon}=req.body;const s=parseInt(satir),a=parseInt(adet)||1,k=parseInt(kolon);
  if(!s||s<1||a<1||!k)return res.status(400).json({hata:t(getLang(req), 'bantInvalid')});
  const atamalar=ao(),yeni={};Object.entries(atamalar).forEach(([key,val])=>{const[r,c]=key.split('-').map(Number);if(r>=s&&c===k)yeni[(r+a)+'-'+c]=val;else yeni[key]=val;});ay(yeni);res.json({mesaj:a + ' ' + t(getLang(req), 'bantRowAdded')});
});
app.delete('/api/siparisler/bant-satir-sil',(req,res)=>{
  const{satir,adet,kolon}=req.body;const s=parseInt(satir),a=parseInt(adet)||1,k=parseInt(kolon);
  if(!s||s<1||a<1||!k)return res.status(400).json({hata:t(getLang(req), 'bantInvalid')});
  const atamalar=ao(),yeni={};Object.entries(atamalar).forEach(([key,val])=>{const[r,c]=key.split('-').map(Number);if(c===k){if(r>=s&&r<s+a)return;if(r>=s+a)yeni[(r-a)+'-'+c]=val;else yeni[key]=val;}else yeni[key]=val;});ay(yeni);res.json({mesaj:a + ' ' + t(getLang(req), 'bantRowDeleted')});
});
app.get('/api/siparisler/bant-json',async(req,res)=>{
  const bd=path.join(uploadsDir,'bant.xlsx');if(!fs.existsSync(bd))return res.json({hata:t(getLang(req), 'bantExcelNotFound')});
  try{const wb=new ExcelJS.Workbook();await wb.xlsx.readFile(bd);const ws=wb.worksheets[0];if(!ws)return res.json({hata:t(getLang(req), 'bantSheetNotFound')});
    const sl=db.prepare('SELECT id,urun_aciklamasi FROM siparisler').all();const mm=new Map();
    sl.forEach(s=>{const c=(s.urun_aciklamasi||'').toUpperCase().match(/[A-Z]\d{3,4}[A-Z]\d{1,3}/g);if(c)c.forEach(x=>mm.set(x,s.id));});
    const bl=[],sr=[],mc=ws.columnCount||23;ws.eachRow((row,rn)=>{const v=[];for(let c=1;c<=mc;c++){const cell=row.getCell(c);let val=cell.value;if(val instanceof Date)val=val.toISOString().split('T')[0];else if(val&&typeof val==='object'&&val.richText)val=val.richText.map(t=>t.text).join('');v.push(String(val||''));}if(rn===1)bl.push(...v);else sr.push(v);});
    const bs=new Date('2026-06-01').getTime();const fl=[];sr.forEach(s=>{const ts=s[0];if(!ts)return;const d=new Date(ts);if(isNaN(d.getTime())||d.getTime()<bs)return;fl.push(s);});
    const cc=bl.length;const bosS=[],tarS=[];const dr=/^\d{4}-\d{2}-\d{2}/;
    for(let c=0;c<cc;c++){let bos=true,tc=0;for(let r=0;r<fl.length;r++){const v=fl[r][c];if(v&&v.trim())bos=false;if(dr.test(v))tc++;}if(bos)bosS.push(c+1);if(tc>fl.length*0.7)tarS.push(c+1);}
    const ks=new Set([1,...tarS]);const tm=fl.map(s=>s.map((v,i)=>ks.has(i+1)?v:''));
    res.json({basliklar:bl,satirlar:tm,stiller:[],bosSutunlar:bosS,tarihSutunlar:tarS,modelSiparisMap:Object.fromEntries(mm),atamalar:ao()});}catch(e){res.status(500).json({hata:t(getLang(req), 'bantReadError') + ': ' + e.message});}
});
app.get('/api/siparisler/:id/excel',async(req,res)=>{
  const row=db.prepare('SELECT * FROM siparisler WHERE id = ?').get(req.params.id);if(!row)return res.status(404).json({hata:'Bulunamadı'});
  if(row.excel_dosya){const d=path.join(uploadsDir,path.basename(row.excel_dosya||''));if(fs.existsSync(d))return res.download(d);}
  const wb=new ExcelJS.Workbook();const ws=wb.addWorksheet('Maliyet');ws.columns=[{key:'A',width:22},{key:'B',width:18}];
  ws.getCell('A1').value='Müşteri';ws.getCell('B1').value=row.musteri_adi;ws.getCell('A2').value='Model';ws.getCell('B2').value=row.urun_aciklamasi;
  ws.getCell('A3').value='Fiyat';ws.getCell('B3').value=row.fiyat+' '+(row.para_birimi||'TL');ws.getCell('A4').value='Adet';ws.getCell('B4').value=row.miktar;
  ws.getCell('A5').value='Toplam';ws.getCell('B5').value=(row.fiyat*row.miktar)+' '+(row.para_birimi||'TL');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');res.setHeader('Content-Disposition','attachment; filename="maliyet_'+row.id+'.xlsx"');await wb.xlsx.write(res);res.end();
});
app.get('/api/siparisler/:id',(req,res)=>{const row=db.prepare('SELECT * FROM siparisler WHERE id = ?').get(req.params.id);if(!row)return res.status(404).json({hata:'Bulunamadı'});res.json(row);});

// sipariş güncelle
const siparisUpdateUpload = karmaUpload.fields([{ name: 'gorsel', maxCount: 1 }, { name: 'excel_dosya', maxCount: 1 }]);

app.put('/api/siparisler/:id', siparisUpdateUpload, (req, res) => {
  const { musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, tarih, notlar, durum, termin, renk, renk_detay, kumas_tamam, fit, ld_test, bulk_test, gold_test, termin_detay, mt, aksesuar } = req.body;

  if (!musteri_adi || !urun_aciklamasi || fiyat == null) {
    return res.status(400).json({ hata: t(getLang(req), 'fiyatRequired') });
  }

  const eski = db.prepare('SELECT gorsel, excel_dosya, durum FROM siparisler WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'siparisNotFound') });

  let gorselYolu = eski.gorsel || '';
  if (req.files && req.files.gorsel) {
    gorselYolu = 'uploads/' + req.files.gorsel[0].filename;
  }

  let excelYolu = eski.excel_dosya || '';
  if (req.files && req.files.excel_dosya) {
    const ext = path.extname(req.files.excel_dosya[0].originalname);
    const excelAdi = 'excel_s_' + Date.now() + '_' + req.params.id + ext;
    const hedefYol = path.join(uploadsDir, excelAdi);
    fs.renameSync(req.files.excel_dosya[0].path, hedefYol);
    excelYolu = 'uploads/' + excelAdi;
  }

  db.prepare(`
    UPDATE siparisler
    SET musteri_adi=?, firma_adi=?, urun_aciklamasi=?, fiyat=?, para_birimi=?, miktar=?, tarih=?, notlar=?, durum=?, gorsel=?, termin=?, renk=?, renk_detay=?, excel_dosya=?, kumas_tamam=?, fit=?, ld_test=?, bulk_test=?, gold_test=?, termin_detay=?, mt=?, aksesuar=?
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
    excelYolu,
    kumas_tamam || '',
    fit || '',
    ld_test || '',
    bulk_test || '',
    gold_test || '',
    termin_detay || '[]',
    mt || '',
    aksesuar || '',
    req.params.id
  );

  // gelir_takip: durum Tamamlandı'dan çıkarsa sil, Tamamlandı olursa ekle
  const yeniDurum2 = durum || 'Hazırlanıyor';
  const eskiDurum2 = eski.durum || '';
  if ((eskiDurum2 || '').startsWith('Tamamland') && !(yeniDurum2 || '').startsWith('Tamamland')) {
    db.prepare('DELETE FROM gelir_takip WHERE siparis_id = ?').run(req.params.id);
  }
  if (!(eskiDurum2 || '').startsWith('Tamamland') && (yeniDurum2 || '').startsWith('Tamamland')) {
    const mevcut = db.prepare('SELECT id FROM gelir_takip WHERE siparis_id = ?').get(req.params.id);
    if (!mevcut) {
      const gerc = db.prepare('SELECT satis_toplam FROM gerceklesen WHERE siparis_id = ?').get(req.params.id);
      const satisToplam = gerc ? (gerc.satis_toplam || ((fiyat || 0) * (miktar || 1))) : ((fiyat || 0) * (miktar || 1));
      db.prepare(`
        INSERT INTO gelir_takip (siparis_id, musteri_adi, firma_adi, urun_aciklamasi, toplam_tutar, para_birimi)
        VALUES (?, ?, ?, ?, ?, ?)
      `).run(
        req.params.id,
        firma_adi || '',
        firma_adi || '',
        urun_aciklamasi,
        satisToplam,
        para_birimi || 'TL'
      );
    }
  }

  res.json({ mesaj: t(getLang(req), 'siparisUpdated') });
});

// --- MT (takım eden) kişiler ---
app.get('/api/mt-kisiler', requirePermission('siparisler'), (req, res) => {
  const rows = db.prepare('SELECT * FROM mt_kisiler ORDER BY isim ASC').all();
  res.json(rows);
});

app.post('/api/mt-kisiler', requirePermission('siparisler'), (req, res) => {
  const ad = String((req.body && req.body.isim) || '').trim().toUpperCase();
  if (!ad) return res.status(400).json({ hata: 'İsim boş olamaz' });
  const mevcut = db.prepare('SELECT id FROM mt_kisiler WHERE isim = ?').get(ad);
  if (mevcut) return res.status(400).json({ hata: 'Bu isim zaten var' });
  const info = db.prepare('INSERT INTO mt_kisiler (isim) VALUES (?)').run(ad);
  res.status(201).json({ id: info.lastInsertRowid, isim: ad });
});

app.delete('/api/mt-kisiler/:id', requirePermission('siparisler'), (req, res) => {
  db.prepare('DELETE FROM mt_kisiler WHERE id = ?').run(req.params.id);
  res.json({ mesaj: 'Silindi' });
});

// sipariş sil (görseli de sil)
app.delete('/api/siparisler/:id', (req, res) => {
  const row = db.prepare('SELECT gorsel FROM siparisler WHERE id = ?').get(req.params.id);
  if (row && row.gorsel) {
    const dosya = path.join(uploadsDir, path.basename(row.gorsel||''));
    if (fs.existsSync(dosya)) fs.unlinkSync(dosya);
  }
  db.prepare('DELETE FROM siparisler WHERE id = ?').run(req.params.id);
  res.json({ mesaj: t(getLang(req), 'siparisDeleted') });
});

// sipariş durumu güncelle (hızlı)
app.patch('/api/siparisler/:id/durum', (req, res) => {
  const { durum } = req.body;
  if (!durum) return res.status(400).json({ hata: 'durum zorunludur' });

  const lang = getLang(req);
  const siparis = db.prepare('SELECT * FROM siparisler WHERE id = ?').get(req.params.id);
  if (!siparis) return res.status(404).json({ hata: t(lang, 'siparisNotFound') });

  const eskiDurum = siparis.durum;
  db.prepare('UPDATE siparisler SET durum=? WHERE id=?').run(durum, req.params.id);

  // Auto-create gelir_takip when order is marked as Tamamlandı
  // use startsWith to avoid Turkish character encoding issues
  if ((durum || '').startsWith('Tamamland') && !(eskiDurum || '').startsWith('Tamamland')) {
    const mevcut = db.prepare('SELECT id FROM gelir_takip WHERE siparis_id = ?').get(req.params.id);
    if (!mevcut) {
      // gerçekleşen kaydından satış toplamını al
      const gerc = db.prepare('SELECT satis_toplam FROM gerceklesen WHERE siparis_id = ?').get(req.params.id);
      const satisToplam = gerc ? (gerc.satis_toplam || (siparis.fiyat * (siparis.miktar || 1))) : (siparis.fiyat * (siparis.miktar || 1));
      db.prepare(`
        INSERT INTO gelir_takip (siparis_id, musteri_adi, firma_adi, urun_aciklamasi, toplam_tutar, para_birimi)
        VALUES (?, ?, ?, ?, ?, ?)
      `).run(
        siparis.id,
        siparis.firma_adi,
        siparis.firma_adi,
        siparis.urun_aciklamasi,
        satisToplam,
        siparis.para_birimi
      );
    }
  }

  // Auto-delete gelir_takip when order leaves Tamamlandı
  if ((eskiDurum || '').startsWith('Tamamland') && !(durum || '').startsWith('Tamamland')) {
    db.prepare('DELETE FROM gelir_takip WHERE siparis_id = ?').run(req.params.id);
  }

  res.json({ mesaj: t(lang, 'siparisStatusUpdated') });
});

// siparişi geri al (Tamamlandı → Üretimde)
app.post('/api/siparisler/:id/geri-al', (req, res) => {
  const siparis = db.prepare('SELECT * FROM siparisler WHERE id = ?').get(req.params.id);
  if (!siparis) return res.status(404).json({ hata: t(getLang(req), 'siparisNotFound') });

  db.prepare('UPDATE siparisler SET durum = ? WHERE id = ?').run('Üretimde', req.params.id);

  // gerçekleşen kaydı korunur (maliyet verileri silinmez) - sadece durum güncellenir
  db.prepare('UPDATE gerceklesen SET durum = ? WHERE siparis_id = ?').run('Üretimde', req.params.id);

  // varsa bağlı gelir_takip kaydını sil
  db.prepare('DELETE FROM gelir_takip WHERE siparis_id = ?').run(req.params.id);

  res.json({ mesaj: t(getLang(req), 'siparisReverted') });
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
  if (!row) return res.status(404).json({ hata: t(getLang(req), 'fiyatNotFound') });
  res.json(row);
});

// --- AKSESUAR TAKİP API ---

// siparişe ait aksesuarları listele
app.get('/api/siparisler/:id/aksesuar', (req, res) => {
  const renk = req.query.renk;
  let rows;
  if (renk !== undefined && renk !== '') {
    rows = db.prepare('SELECT * FROM aksesuar_takip WHERE siparis_id = ? AND renk = ? ORDER BY created_at ASC').all(req.params.id, renk);
  } else {
    rows = db.prepare('SELECT * FROM aksesuar_takip WHERE siparis_id = ? ORDER BY created_at ASC').all(req.params.id);
  }
  res.json(rows);
});

// yeni aksesuar ekle
app.post('/api/siparisler/:id/aksesuar', (req, res) => {
  const { aksesuar_adi, siparis_adet, gelen_adet, birim_fiyat, notlar, fatura, renk } = req.body;
  if (!aksesuar_adi) return res.status(400).json({ hata: t(getLang(req), 'aksesuarRequired') });

  const info = db.prepare(`
    INSERT INTO aksesuar_takip (siparis_id, aksesuar_adi, siparis_adet, gelen_adet, birim_fiyat, notlar, fatura, gelen_detay, renk)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(req.params.id, aksesuar_adi, siparis_adet || 0, gelen_adet || 0, birim_fiyat || 0, notlar || '', fatura || '', '[]', renk || '');

  res.status(201).json({ id: info.lastInsertRowid, mesaj: t(getLang(req), 'aksesuarCreated') });
});

// tek aksesuar getir
app.get('/api/aksesuar/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM aksesuar_takip WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ hata: t(getLang(req), 'aksesuarNotFound') });
  res.json(row);
});

// aksesuar güncelle
app.put('/api/aksesuar/:id', (req, res) => {
  const { aksesuar_adi, siparis_adet, birim_fiyat, notlar, gelen_detay } = req.body;
  if (!aksesuar_adi) return res.status(400).json({ hata: t(getLang(req), 'aksesuarRequired') });

  const eski = db.prepare('SELECT * FROM aksesuar_takip WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'aksesuarNotFound') });

  // gelen_detay JSON'dan toplam gelen adedi hesapla
  let gelenAdet = 0;
  let detayJson = '[]';
  try {
    const detay = typeof gelen_detay === 'string' ? JSON.parse(gelen_detay) : (gelen_detay || []);
    if (Array.isArray(detay)) {
      detay.forEach(function(d) { gelenAdet += (parseInt(d.adet) || 0); });
      detayJson = JSON.stringify(detay);
    }
  } catch(e) { detayJson = gelen_detay || '[]'; }

  db.prepare(`
    UPDATE aksesuar_takip SET aksesuar_adi=?, siparis_adet=?, gelen_adet=?, birim_fiyat=?, notlar=?, gelen_detay=? WHERE id=?
  `).run(aksesuar_adi, siparis_adet || 0, gelenAdet, birim_fiyat !== undefined ? birim_fiyat : (eski.birim_fiyat || 0), notlar || '', detayJson, req.params.id);

  // cari senkron: aksesuar sevkiyat faturaları → borç
  let aksDetayArr = [];
  try { aksDetayArr = typeof detayJson === 'string' ? JSON.parse(detayJson) : (detayJson || []); } catch(e) { aksDetayArr = []; }
  if (!Array.isArray(aksDetayArr)) aksDetayArr = [];
  const aksKaynak = 'aksesuar:' + req.params.id + ':';
  const aksEntries = aksDetayArr.map((d, i) => ({
    kaynak: aksKaynak + i,
    cari_id: d.cari_id,
    tutar: ((parseFloat(d.tutar) || 0) + (parseFloat(d.kdv) || 0)) || (parseFloat(d.vergi_dahil) || 0),
    tarih: d.tarih || '',
    aciklama: (aksesuar_adi || '') + (d.aciklama ? ' - ' + d.aciklama : '') + (d.fatura ? ' - ' + d.fatura : ''),
    belge_no: d.belge_no || d.fatura || ''
  }));
  cariSenkron(aksKaynak, aksEntries);

  res.json({ mesaj: t(getLang(req), 'aksesuarUpdated'), gelen_adet: gelenAdet });
});

// aksesuar sil
app.delete('/api/aksesuar/:id', (req, res) => {
  const eski = db.prepare('SELECT * FROM aksesuar_takip WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'aksesuarNotFound') });
  db.prepare('DELETE FROM aksesuar_takip WHERE id = ?').run(req.params.id);
  cariSenkron('aksesuar:' + req.params.id + ':', []);
  res.json({ mesaj: t(getLang(req), 'aksesuarDeleted') });
});

// --- KUMAŞ TAKİP API ---

function kumasBirimGiderHesapla(pb, ke, kg, fr, asorti, biye) {
  const as = parseFloat(asorti) || 1;
  return as > 0 ? (pb * ke * kg * (1 + fr / 100)) / as + biye : 0;
}

// bir siparişin (modelin) tüm renk varyantlarını döndür
function siparisRenkleri(siparisId) {
  const s = db.prepare('SELECT renk, renk_detay FROM siparisler WHERE id = ?').get(siparisId);
  if (!s) return [];
  const set = new Set();
  if (s.renk) set.add(String(s.renk));
  try {
    const d = JSON.parse(s.renk_detay || '[]');
    if (Array.isArray(d)) d.forEach(x => { if (x && x.renk) set.add(String(x.renk)); });
  } catch (e) {}
  return [...set];
}

// kumaşın fabrika alanlarını (ad, pastal boyu, en, gramaj, fire, biye) aynı modelin diğer renk varyantlarına yay
// eskiAd: kaynak satırın eski kumaş adı (isim değişikliğinde eşleştirme için)
function kumasVaryantlaraYay(siparisId, kaynakRenk, eskiAd, d) {
  const yeniAd = d.kumas_adi || '';
  if (!yeniAd) return;
  const hedefler = siparisRenkleri(siparisId).filter(r => r && r !== kaynakRenk);
  if (hedefler.length === 0) return;

  const pb = parseFloat(d.pastal_boyu) || 0;
  const ke = parseFloat(d.kumas_en) || 0;
  const kg = parseFloat(d.kumas_gramaj) || 0;
  const fr = parseFloat(d.fire) || 0;
  const by = parseFloat(d.biye) || 0;
  const asKaynak = parseFloat(d.asorti) || 1;
  const adDegisti = eskiAd && eskiAd !== yeniAd;

  for (const r of hedefler) {
    // isim değiştiyse önce eski isimle eşleşen satırı bul (rename senaryosu)
    const eskiAdaGore = adDegisti
      ? db.prepare('SELECT * FROM kumas_takip WHERE siparis_id = ? AND renk = ? AND kumas_adi = ?').get(siparisId, r, eskiAd)
      : null;
    const yeniAdaGore = db.prepare('SELECT * FROM kumas_takip WHERE siparis_id = ? AND renk = ? AND kumas_adi = ?').get(siparisId, r, yeniAd);
    const hedef = eskiAdaGore || yeniAdaGore;

    if (hedef) {
      const as = parseFloat(hedef.asorti) || 1;
      const birim = kumasBirimGiderHesapla(pb, ke, kg, fr, as, by);
      db.prepare('UPDATE kumas_takip SET kumas_adi=?, pastal_boyu=?, kumas_en=?, kumas_gramaj=?, fire=?, biye=?, birim_gider=? WHERE id=?')
        .run(yeniAd, pb, ke, kg, fr, by, birim, hedef.id);
    } else {
      const birim = kumasBirimGiderHesapla(pb, ke, kg, fr, asKaynak, by);
      db.prepare(`INSERT INTO kumas_takip (siparis_id, kumas_adi, renk, adet, pastal_boyu, kumas_en, kumas_gramaj, fire, asorti, biye, birim_gider, kumas_fiyati, kesim_adedi, kesim_yuzdesi, gerceklesen_birim, fit_birim, stok, notlar, gelen_detay)
        VALUES (?, ?, ?, 0, ?, ?, ?, ?, ?, ?, ?, 0, 0, 0, 0, 0, 0, '', '[]')`)
        .run(siparisId, yeniAd, r, pb, ke, kg, fr, asKaynak, by, birim);
    }
  }
}

// siparişe ait kumaşları listele
app.get('/api/siparisler/:id/kumas', (req, res) => {
  const renk = req.query.renk;
  let rows;
  if (renk !== undefined && renk !== '') {
    rows = db.prepare('SELECT * FROM kumas_takip WHERE siparis_id = ? AND renk = ? ORDER BY created_at ASC').all(req.params.id, renk);
  } else {
    rows = db.prepare('SELECT * FROM kumas_takip WHERE siparis_id = ? ORDER BY created_at ASC').all(req.params.id);
  }
  res.json(rows);
});

// yeni kumaş ekle
app.post('/api/siparisler/:id/kumas', (req, res) => {
  const { kumas_adi, renk, adet, pastal_boyu, kumas_en, kumas_gramaj, fire, asorti, biye, kumas_fiyati, kesim_adedi, kesim_yuzdesi, fit_birim, stok, notlar } = req.body;
  const info = db.prepare(`
    INSERT INTO kumas_takip (siparis_id, kumas_adi, renk, adet, pastal_boyu, kumas_en, kumas_gramaj, fire, asorti, biye, birim_gider, kumas_fiyati, kesim_adedi, kesim_yuzdesi, gerceklesen_birim, fit_birim, stok, notlar)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(req.params.id, kumas_adi || '', renk || '', adet || 0, pastal_boyu || 0, kumas_en || 0, kumas_gramaj || 0, fire || 0, asorti || 1, biye || 0, 0, kumas_fiyati || 0, kesim_adedi || 0, kesim_yuzdesi || 0, 0, fit_birim || 0, stok || 0, notlar || '');
  res.status(201).json({ id: info.lastInsertRowid, mesaj: t(getLang(req), 'kumasCreated') });
});

// tek kumaş getir
app.get('/api/kumas/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM kumas_takip WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ hata: t(getLang(req), 'kumasNotFound') });
  res.json(row);
});

// kumaş güncelle (hesaplamaları yap)
app.put('/api/kumas/:id', (req, res) => {
  const { kumas_adi, renk, adet, pastal_boyu, kumas_en, kumas_gramaj, fire, asorti, biye, kumas_fiyati, kesim_adedi, kesim_yuzdesi, fit_birim, stok, notlar, gelen_detay } = req.body;
  const eski = db.prepare('SELECT * FROM kumas_takip WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'kumasNotFound') });

  const pb = parseFloat(pastal_boyu) || 0;
  const ke = parseFloat(kumas_en) || 0;
  const kg = parseFloat(kumas_gramaj) || 0;
  const fr = parseFloat(fire) || 0;
  const as = parseFloat(asorti) || 1;
  const by = parseFloat(biye) || 0;
  const ad = parseInt(adet) || 0;

  // birim_gider = (pastal_boyu * kumas_en * kumas_gramaj * (1 + fire/100)) / asorti + biye
  const birimGider = as > 0 ? (pb * ke * kg * (1 + fr / 100)) / as + by : 0;
  // kumaş siparişi = birim_gider × adet
  const kumasSiparis = birimGider * ad;
  // kesim yüzdesi = (kesim adedi / adet) × 100
  const ka = parseInt(kesim_adedi) || 0;
  const hesaplananKesimYuzdesi = ad > 0 ? Math.round((ka / ad) * 10000) / 100 : 0;

  // gerçekleşen birim = toplam gelen kumaş miktarı / kesim adedi
  let toplamGelenKumas = 0;
  let detayJson = '[]';
  try {
    const detay = typeof gelen_detay === 'string' ? JSON.parse(gelen_detay) : (gelen_detay || []);
    if (Array.isArray(detay)) {
      detay.forEach(function(d) { toplamGelenKumas += (parseFloat(d.adet) || 0); });
      detayJson = JSON.stringify(detay);
    }
  } catch(e) { detayJson = gelen_detay || '[]'; }
  const stokMiktar = parseFloat(stok) || 0;
  const gerceklesenBirim = ka > 0 ? Math.round(((toplamGelenKumas - stokMiktar) / ka) * 1000) / 1000 : 0;

  db.prepare(`
    UPDATE kumas_takip SET kumas_adi=?, renk=?, adet=?, pastal_boyu=?, kumas_en=?, kumas_gramaj=?, fire=?, asorti=?, biye=?, birim_gider=?, kumas_fiyati=?, kesim_adedi=?, kesim_yuzdesi=?, gerceklesen_birim=?, fit_birim=?, stok=?, notlar=?, gelen_detay=?
    WHERE id=?
  `).run(kumas_adi || '', renk || '', ad, pastal_boyu || 0, kumas_en || 0, kumas_gramaj || 0, fire || 0, asorti || 1, biye || 0, birimGider, kumas_fiyati || 0, ka, hesaplananKesimYuzdesi, gerceklesenBirim, fit_birim || 0, stokMiktar, notlar || '', detayJson, req.params.id);

  // kumaşın fabrika alanları (ad, pastal boyu, en, gramaj, fire, biye) değiştiyse aynı modelin diğer renk varyantlarına yay
  const fabrikaDegisti =
    (eski.kumas_adi || '') !== (kumas_adi || '') ||
    (parseFloat(eski.pastal_boyu) || 0) !== pb ||
    (parseFloat(eski.kumas_en) || 0) !== ke ||
    (parseFloat(eski.kumas_gramaj) || 0) !== kg ||
    (parseFloat(eski.fire) || 0) !== fr ||
    (parseFloat(eski.biye) || 0) !== by;
  if (fabrikaDegisti) {
    kumasVaryantlaraYay(eski.siparis_id, eski.renk || '', eski.kumas_adi || '', { kumas_adi: kumas_adi || '', pastal_boyu: pb, kumas_en: ke, kumas_gramaj: kg, fire: fr, biye: by, asorti: as });
  }

  // gerçekleşen varsa kesim adedini güncelle
  const siparisId = eski.siparis_id;
  const gerceklesen = db.prepare('SELECT id FROM gerceklesen WHERE siparis_id = ?').get(siparisId);
  if (gerceklesen) {
    const toplamKesim = db.prepare('SELECT COALESCE(SUM(kesim_adedi), 0) as toplam FROM kumas_takip WHERE siparis_id = ?').get(siparisId);
    db.prepare('UPDATE gerceklesen SET kesim_adedi = ? WHERE id = ?').run(toplamKesim.toplam, gerceklesen.id);
  }

  // cari senkron: kumaş sevkiyat faturaları → borç
  let kumasDetayArr = [];
  try { kumasDetayArr = typeof detayJson === 'string' ? JSON.parse(detayJson) : (detayJson || []); } catch(e) { kumasDetayArr = []; }
  if (!Array.isArray(kumasDetayArr)) kumasDetayArr = [];
  const kumasKaynak = 'kumas:' + req.params.id + ':';
  const kumasEntries = kumasDetayArr.map((d, i) => ({
    kaynak: kumasKaynak + i,
    cari_id: d.cari_id,
    tutar: ((parseFloat(d.tutar) || 0) + (parseFloat(d.kdv) || 0)) || (parseFloat(d.vergi_dahil) || 0),
    tarih: d.tarih || '',
    aciklama: (kumas_adi || '') + (renk ? ' (' + renk + ')' : '') + (d.fatura ? ' - ' + d.fatura : ''),
    belge_no: d.belge_no || d.fatura || ''
  }));
  cariSenkron(kumasKaynak, kumasEntries);

  res.json({ mesaj: t(getLang(req), 'kumasUpdated'), birim_gider: Math.round(birimGider * 100) / 100, kumas_siparis: Math.round(kumasSiparis * 100) / 100, kesim_yuzdesi: hesaplananKesimYuzdesi, gerceklesen_birim: gerceklesenBirim });
});

// tüm siparişlerin kumaş faturalarını getir (gerçekleşen için)
app.get('/api/kumas-faturalar', (req, res) => {
  const rows = db.prepare('SELECT * FROM kumas_takip ORDER BY siparis_id, created_at ASC').all();
  const grouped = {};
  rows.forEach(k => {
    const sid = k.siparis_id;
    if (!grouped[sid]) grouped[sid] = { faturalar: [], toplam_tutar: 0 };
    let detaylar = [];
    try { detaylar = typeof k.gelen_detay === 'string' ? JSON.parse(k.gelen_detay) : (k.gelen_detay || []); } catch(e) { detaylar = []; }
    if (!Array.isArray(detaylar)) detaylar = [];
    detaylar.forEach(function(d) {
      if (d.fatura && d.fatura.trim()) {
        const tutar = parseFloat(d.tutar) || 0;
        grouped[sid].faturalar.push({
          kumas_adi: k.kumas_adi,
          renk: k.renk || '',
          fatura: d.fatura.trim(),
          adet: d.adet || 0,
          aciklama: d.aciklama || '',
          tutar: tutar
        });
        grouped[sid].toplam_tutar += tutar;
      }
    });
  });
  res.json(grouped);
});

// kumaş sil
app.delete('/api/kumas/:id', (req, res) => {
  const eski = db.prepare('SELECT * FROM kumas_takip WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'kumasNotFound') });
  db.prepare('DELETE FROM kumas_takip WHERE id = ?').run(req.params.id);
  cariSenkron('kumas:' + req.params.id + ':', []);
  res.json({ mesaj: t(getLang(req), 'kumasDeleted') });
});

// --- KUMAŞ SİPARİŞ FORMU (Excel) ---

async function generateKumasSiparisFormu(res, kumas, overrides = {}) {
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Kumaş Sipariş Formu', {
    pageSetup: { orientation: 'portrait', paperSize: 9, margins: { top: 0.4, bottom: 0.4, left: 0.4, right: 0.4 } }
  });

  // Sütun genişlikleri
  ws.columns = [
    { key: 'A', width: 14 }, { key: 'B', width: 11 }, { key: 'C', width: 11 },
    { key: 'D', width: 10 }, { key: 'E', width: 12 }, { key: 'F', width: 9 },
    { key: 'G', width: 9 }, { key: 'H', width: 16 }
  ];

  // Stiller
  const hFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } };
  const hFont = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
  const thinBorder = { top: { style: 'thin', color: { argb: 'FF94A3B8' } }, bottom: { style: 'thin', color: { argb: 'FF94A3B8' } }, left: { style: 'thin', color: { argb: 'FF94A3B8' } }, right: { style: 'thin', color: { argb: 'FF94A3B8' } } };
  const labelFont = { bold: true, size: 9 };
  const valFont = { size: 9 };

  const bugun = new Date();
  const today = ('0' + bugun.getDate()).slice(-2) + '.' + ('0' + (bugun.getMonth() + 1)).slice(-2) + '.' + bugun.getFullYear();

  function formatDateTR(d) {
    if (!d) return today;
    // "YYYY-MM-DD" veya "GG.AA.YYYY" formatındaysa dönüştür
    const parts = String(d).split(/[-.]/);
    if (parts.length === 3) {
      if (parts[0].length === 4) return parts[2] + '.' + parts[1] + '.' + parts[0]; // YYYY-MM-DD → GG.AA.YYYY
      return parts[0] + '.' + parts[1] + '.' + parts[2]; // zaten GG.AA.YYYY
    }
    return String(d);
  }

  // --- Logo ---
  const logoPath = path.join(__dirname, 'public', 'logo.png');
  if (fs.existsSync(logoPath)) {
    const logoId = workbook.addImage({ filename: logoPath, extension: 'png' });
    ws.addImage(logoId, { tl: { col: 0.05, row: 0.05 }, ext: { width: 110, height: 65 } });
  }

  // --- Başlık ---
  ws.mergeCells('A1:H1');
  ws.getCell('A1').value = 'KUMAŞ SİPARİŞ FORMU';
  ws.getCell('A1').font = { bold: true, size: 18, color: { argb: 'FF1E40AF' } };
  ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 60;

  // Form No + Tarih
  ws.mergeCells('G2:H2');
  ws.getCell('G2').value = 'Form No: KSF-' + kumas.id + ' | Tarih: ' + today;
  ws.getCell('G2').font = { size: 8, color: { argb: 'FF64748B' } };
  ws.getCell('G2').alignment = { horizontal: 'right' };

  let r = 3;

  // --- Üst Bilgi Alanı (2 sütun) ---
  const bilgiler = [
    ['MÜŞTERİ', overrides.musteri || COMPANY_NAME, 'SİPARİŞ ADI', overrides.siparis_adi || kumas.urun_aciklamasi || ''],
    ['SİPARİŞİ GEÇEN', overrides.siparisi_gecen || '', 'SATICI FİRMA', overrides.satici_firma || ''],
    ['FİRMA YETKİLİSİ', overrides.firma_yetkilisi || '', 'SİPARİŞ TARİHİ', overrides.siparis_tarihi || today],
  ];

  bilgiler.forEach(([l1, v1, l2, v2]) => {
    ws.mergeCells('A' + r + ':B' + r);
    ws.getCell('A' + r).value = l1 + ': ' + (v1 || '');
    ws.getCell('A' + r).font = { bold: true, size: 9 };
    ws.mergeCells('D' + r + ':E' + r);
    ws.getCell('D' + r).value = l2 + ': ' + (v2 ? formatDateTR(v2) : '');
    ws.getCell('D' + r).font = { bold: true, size: 9 };
    ws.getCell('C' + r).value = '';
    ws.getCell('F' + r).value = '';
    r++;
  });

  // Boş satır
  r++;

  // --- Tablo Başlığı ---
  const headers = ['KUMAŞ ADI', 'RENK', 'MİKTAR', 'FİYAT', 'KULLANIM YERİ', 'EN', 'GRAMAJ', 'SİPARİŞ TERMİNİ'];
  headers.forEach((h, i) => {
    const cell = ws.getCell(String.fromCharCode(65 + i) + r);
    cell.value = h;
    cell.font = hFont;
    cell.fill = hFill;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = thinBorder;
  });
  ws.getRow(r).height = 22;
  r++;

  // --- Tablo Veri Satırları ---
  const rows = overrides.tablo_rows || [{
    kumas_adi: kumas.kumas_adi || '',
    renk: kumas.renk || '',
    miktar: (kumas.adet || 0) + ' ADET',
    fiyat: (kumas.kumas_fiyati || 0) + ' TL',
    kullanim_yeri: '',
    en: kumas.kumas_en || 0,
    gramaj: kumas.kumas_gramaj || 0,
    termin: kumas.termin || ''
  }];

  for (let ri = 0; ri < Math.max(rows.length, 5); ri++) {
    const row = rows[ri] || {};
    const vals = [row.kumas_adi || '', row.renk || '', row.miktar || '', row.fiyat || '', row.kullanim_yeri || '', row.en || '', row.gramaj || '', row.termin || ''];
    vals.forEach((v, i) => {
      const cell = ws.getCell(String.fromCharCode(65 + i) + r);
      cell.value = v;
      cell.font = valFont;
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = thinBorder;
    });
    ws.getRow(r).height = 22;
    r++;
  }

  // Boş satır
  r++;

  // --- Alt Bölüm: Alıcı Bilgileri + Satıcı İmza (yan yana) ---
  const leftStart = r;

  // SOL: ALICININ İLETİŞİM BİLGİLERİ
  ws.mergeCells('A' + r + ':D' + r);
  ws.getCell('A' + r).value = 'ALICININ İLETİŞİM BİLGİLERİ';
  ws.getCell('A' + r).font = { bold: true, size: 10, color: { argb: 'FF1E40AF' } };
  ws.getCell('A' + r).border = { bottom: { style: 'medium', color: { argb: 'FF1E40AF' } } };
  r++;

  const aliciBilgileri = [
    ['Adres', overrides.adres || ''],
    ['Telefon No', overrides.telefon || ''],
    ['Vergi Dairesi Adı', overrides.vergi_dairesi || ''],
    ['Vergi Numarası', overrides.vergi_no || ''],
    ['Teslimat Adresi', overrides.teslimat_adres || ''],
  ];
  aliciBilgileri.forEach(([l, v]) => {
    ws.mergeCells('A' + r + ':B' + r);
    ws.getCell('A' + r).value = l + ':';
    ws.getCell('A' + r).font = { bold: true, size: 8 };
    ws.mergeCells('C' + r + ':D' + r);
    ws.getCell('C' + r).value = v || '';
    ws.getCell('C' + r).font = { size: 8 };
    r++;
  });

  // Ödeme Şekli
  ws.mergeCells('A' + r + ':B' + r);
  ws.getCell('A' + r).value = 'Ödeme Şekli:';
  ws.getCell('A' + r).font = { bold: true, size: 8 };
  ws.mergeCells('C' + r + ':D' + r);
  ws.getCell('C' + r).value = overrides.odeme_sekli || '';
  ws.getCell('C' + r).font = { size: 8 };
  r++;

  const leftEnd = r;

  // SAĞ: SATICI İMZA KAŞE
  r = leftStart + 1;
  ws.mergeCells('F' + r + ':H' + (r + 5));
  ws.getCell('F' + r).value = 'SATICI\nİMZA\nKAŞE';
  ws.getCell('F' + r).font = { bold: true, size: 10, color: { argb: 'FF64748B' } };
  ws.getCell('F' + r).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  ws.getCell('F' + r).border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };

  r = Math.max(leftEnd, r + 6);

  // --- Yasal Metin ---
  r++;
  ws.mergeCells('A' + r + ':H' + (r + 8));
  ws.getCell('A' + r).value = 'Mallar alıcı tarafından, yukarıda belirtilen yerlerden ve belirtilen süre içerisinde teslim alınacaktır. Satıcı, ürünün tamamının ve bir kısmının zamanında teslim edememesi sonucunda alıcının müşterisinin söz konusu kumaşların kullanılacağı nihai ürün siparişlerini iptal etme ve/veya indirimli olarak alma hakkının olduğunun bilincinde olmalıdır. Satıcı, sözleşmede belirtilen teslim sürelerine uyulmaması halinde Alıcı\'nın anlaşılan KDV dahil ürün fiyatı üzerinden aşağıda belirtilen kademelendirilmiş reklamasyon bedellerinin kendisine fatura edileceğini kabul eder. Alıcı\'nın teslimde geciken ürünleri aşağıdaki tabloya göre reklamasyon uygulayarak teslim almak veya siparişi iptal etmek hakkına sahiptir. Alıcının, uğramış olduğu zararları satıcıya yansıtma hakkı saklıdır.';
  ws.getCell('A' + r).font = { size: 7, italic: true };
  ws.getCell('A' + r).alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
  for (let i = 0; i < 9; i++) ws.getRow(r + i).height = 14;
  r += 9;

  // --- Reklamasyon Tablosu ---
  r++;
  ws.mergeCells('A' + r + ':D' + r);
  ws.getCell('A' + r).value = 'Reklamasyon Uygulanacak Siparişler';
  ws.getCell('A' + r).font = { bold: true, size: 9, color: { argb: 'FF1E40AF' } };

  ws.mergeCells('F' + r + ':H' + r);
  ws.getCell('F' + r).value = 'Teslim tarihine göre';
  ws.getCell('F' + r).font = { bold: true, size: 9, color: { argb: 'FF1E40AF' } };
  r++;

  // Reklamasyon tablo başlıkları
  ws.mergeCells('A' + r + ':D' + r);
  ws.getCell('A' + r).value = 'Gecikme Gün Sayısı';
  ws.getCell('A' + r).font = hFont; ws.getCell('A' + r).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF64748B' } };
  ws.getCell('A' + r).alignment = { horizontal: 'center' }; ws.getCell('A' + r).border = thinBorder;

  ws.mergeCells('F' + r + ':H' + r);
  ws.getCell('F' + r).value = 'Uygulanacak Reklamasyon Oranı';
  ws.getCell('F' + r).font = hFont; ws.getCell('F' + r).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF64748B' } };
  ws.getCell('F' + r).alignment = { horizontal: 'center' }; ws.getCell('F' + r).border = thinBorder;
  r++;

  const reklamasyon = [
    ['0-1 gün', '%0'],
    ['2-7 gün', '%3'],
    ['8-14 gün', '%5'],
    ['15+ gün', '%10'],
  ];
  reklamasyon.forEach(([gun, oran]) => {
    ws.mergeCells('A' + r + ':D' + r);
    ws.getCell('A' + r).value = gun;
    ws.getCell('A' + r).font = valFont; ws.getCell('A' + r).alignment = { horizontal: 'center' };
    ws.getCell('A' + r).border = thinBorder;
    ws.mergeCells('F' + r + ':H' + r);
    ws.getCell('F' + r).value = oran;
    ws.getCell('F' + r).font = valFont; ws.getCell('F' + r).alignment = { horizontal: 'center' };
    ws.getCell('F' + r).border = thinBorder;
    r++;
  });

  const safeName = (s) => String(s || '').replace(/[^a-zA-Z0-9]/g, '_').replace(/_+/g, '_').substring(0, 30);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="Kumas_Siparis_Formu_${safeName(kumas.kumas_adi)}_${kumas.id}.xlsx"`);
  await workbook.xlsx.write(res);
  res.end();
}

// GET: hızlı indirme (sadece DB verileriyle)
app.get('/api/kumas/:id/siparis-formu', async (req, res) => {
  try {
    const kumas = db.prepare('SELECT k.*, s.musteri_adi, s.firma_adi, s.urun_aciklamasi, s.termin FROM kumas_takip k LEFT JOIN siparisler s ON k.siparis_id = s.id WHERE k.id = ?').get(req.params.id);
    if (!kumas) return res.status(404).json({ hata: t(getLang(req), 'kumasNotFound') });
    await generateKumasSiparisFormu(res, kumas, {});
  } catch (e) {
    res.status(500).json({ hata: t(getLang(req), 'excelProcessError') + ': ' + e.message });
  }
});

// POST: manuel alanlarla birlikte indirme
app.post('/api/kumas/:id/siparis-formu', async (req, res) => {
  try {
    const kumas = db.prepare('SELECT k.*, s.musteri_adi, s.firma_adi, s.urun_aciklamasi, s.termin FROM kumas_takip k LEFT JOIN siparisler s ON k.siparis_id = s.id WHERE k.id = ?').get(req.params.id);
    if (!kumas) return res.status(404).json({ hata: t(getLang(req), 'kumasNotFound') });
    const overrides = req.body || {};
    await generateKumasSiparisFormu(res, kumas, overrides);
  } catch (e) {
    res.status(500).json({ hata: t(getLang(req), 'excelProcessError') + ': ' + e.message });
  }
});

// tüm siparişlerin aksesuar faturalarını getir (gerçekleşen için)
app.get('/api/aksesuar-faturalar', (req, res) => {
  const rows = db.prepare('SELECT * FROM aksesuar_takip ORDER BY siparis_id, created_at ASC').all();
  const grouped = {};
  rows.forEach(a => {
    const sid = a.siparis_id;
    if (!grouped[sid]) grouped[sid] = { faturalar: [], toplam_tutar: 0 };
    // gelen_detay JSON'ından faturaları çıkar
    let detaylar = [];
    try { detaylar = typeof a.gelen_detay === 'string' ? JSON.parse(a.gelen_detay) : (a.gelen_detay || []); } catch(e) { detaylar = []; }
    if (!Array.isArray(detaylar)) detaylar = [];
    detaylar.forEach(function(d) {
      if (d.fatura && d.fatura.trim()) {
        const tutar = parseFloat(d.tutar) || 0;
        grouped[sid].faturalar.push({
          aksesuar_adi: a.aksesuar_adi,
          fatura: d.fatura.trim(),
          adet: d.adet || 0,
          aciklama: d.aciklama || '',
          tutar: tutar
        });
        grouped[sid].toplam_tutar += tutar;
      }
    });
  });
  res.json(grouped);
});

// siparişi gerçekleşene aktar (sipariş silinmez, her ikisi de kalır)
app.post('/api/gerceklesen', (req, res) => {
  const { siparis_id } = req.body;
  if (!siparis_id) return res.status(400).json({ hata: t(getLang(req), 'siparisIdRequired') });

  const siparis = db.prepare('SELECT * FROM siparisler WHERE id = ?').get(siparis_id);
  if (!siparis) return res.status(404).json({ hata: t(getLang(req), 'siparisNotFound') });

  // Bu sipariş için zaten gerçekleşen kaydı var mı kontrol et
  const mevcut = db.prepare('SELECT id FROM gerceklesen WHERE siparis_id = ?').get(siparis_id);
  if (mevcut) {
    return res.json({ id: mevcut.id, mesaj: t(getLang(req), 'gerceklesenExists') });
  }

  const stmt = db.prepare(`
    INSERT INTO gerceklesen (musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, gorsel, siparis_id, excel_dosya, termin)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const info = stmt.run(
    siparis.musteri_adi,
    siparis.firma_adi,
    siparis.urun_aciklamasi,
    siparis.fiyat,
    siparis.para_birimi,
    siparis.miktar,
    siparis.gorsel,
    siparis.id,
    siparis.excel_dosya || '',
    siparis.termin || ''
  );

  res.status(201).json({ id: info.lastInsertRowid, mesaj: t(getLang(req), 'gerceklesenCreated') });
});

// gerçekleşen güncelle (maliyet kalemleri + excel yükleme)
app.put('/api/gerceklesen/:id', excelUpload.single('excel_dosya'), (req, res) => {
  const { musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar,
          kesim_adedi, yuklenen_adet, ikinci_kalite, kumas_bedeli, aksesuar_bedeli, iscilik_bedeli, satis_toplam, termin,
          yuklenen_adet_detay, satis_toplam_detay, kumas_bedeli_detay, aksesuar_bedeli_detay, iscilik_bedeli_detay, baski_nakis_yikama, baski_nakis_yikama_detay, paketleme_malzeme_detay, paketleme_bedeli, nakliye_bedeli } = req.body;

  if (!musteri_adi || !urun_aciklamasi || fiyat == null) {
    return res.status(400).json({ hata: t(getLang(req), 'fiyatRequired') });
  }

  const eski = db.prepare('SELECT * FROM gerceklesen WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'fiyatNotFound') });

  // yeni excel yüklendiyse kaydet
  let excelYolu = eski.excel_dosya || '';
  if (req.file) {
    const ext = path.extname(req.file.originalname);
    const excelAdi = 'excel_' + Date.now() + '_' + req.params.id + ext;
    const hedefYol = path.join(uploadsDir, excelAdi);
    fs.renameSync(req.file.path, hedefYol);
    excelYolu = 'uploads/' + excelAdi;
  }

  const nakliyeCari = req.body.nakliye_cari_id ? parseInt(req.body.nakliye_cari_id) : null;

  // paketleme malzeme tüketimi (stoktan) → paketleme bedelini hesapla
  let pakDetayArr = [];
  try { pakDetayArr = typeof paketleme_malzeme_detay === 'string' ? JSON.parse(paketleme_malzeme_detay) : (paketleme_malzeme_detay || []); } catch(e) { pakDetayArr = []; }
  if (!Array.isArray(pakDetayArr)) pakDetayArr = [];
  let paketlemeBedeli = parseFloat(paketleme_bedeli) || 0; // legacy/manual yedek
  if (pakDetayArr.length > 0) {
    let pakToplam = 0;
    pakDetayArr = pakDetayArr.map(d => {
      const mid = d.malzeme_id ? parseInt(d.malzeme_id) : null;
      const adet = parseFloat(d.adet) || 0;
      if (mid == null || adet <= 0) return null;
      const mat = db.prepare('SELECT adi FROM paketleme_malzemeleri WHERE id = ?').get(mid);
      const birim = paketlemeBirimFiyat(mid);
      const tutar = Math.round((birim * adet) * 100) / 100;
      pakToplam += tutar;
      return { malzeme_id: mid, adi: mat ? mat.adi : '', adet, birim_fiyat: birim, tutar };
    }).filter(Boolean);
    paketlemeBedeli = Math.round(pakToplam * 100) / 100;
  }
  const pakDetayJson = JSON.stringify(pakDetayArr);

  db.prepare(`
    UPDATE gerceklesen
    SET musteri_adi=?, firma_adi=?, urun_aciklamasi=?, fiyat=?, para_birimi=?, miktar=?,
        kesim_adedi=?, yuklenen_adet=?, ikinci_kalite=?, kumas_bedeli=?, aksesuar_bedeli=?, iscilik_bedeli=?, satis_toplam=?, termin=?, excel_dosya=?,
        yuklenen_adet_detay=?, satis_toplam_detay=?, kumas_bedeli_detay=?, aksesuar_bedeli_detay=?, iscilik_bedeli_detay=?, baski_nakis_yikama=?, baski_nakis_yikama_detay=?, paketleme_bedeli=?, paketleme_malzeme_detay=?, nakliye_bedeli=?, nakliye_cari_id=?
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
    yuklenen_adet_detay || '[]',
    satis_toplam_detay || '[]',
    kumas_bedeli_detay || '[]',
    aksesuar_bedeli_detay || '[]',
    iscilik_bedeli_detay || '[]',
    baski_nakis_yikama || 0,
    baski_nakis_yikama_detay || '[]',
    paketlemeBedeli,
    pakDetayJson,
    nakliye_bedeli || 0,
    nakliyeCari,
    req.params.id
  );

  // cari senkron: gerçekleşen gider kalemleri → borç (paketleme artık stoktan, ayrı cari yok)
  const gid = req.params.id;
  const gkaynak = 'gerceklesen:' + gid + ':';
  const gEntries = [];

  let iscArr = []; try { iscArr = typeof iscilik_bedeli_detay === 'string' ? JSON.parse(iscilik_bedeli_detay) : (iscilik_bedeli_detay || []); } catch(e) { iscArr = []; }
  if (!Array.isArray(iscArr)) iscArr = [];
  iscArr.forEach((d, i) => gEntries.push({ kaynak: gkaynak + 'iscilik:' + i, cari_id: d.cari_id, tutar: ((parseFloat(d.deger) || 0) + (parseFloat(d.kdv) || 0)) || (parseFloat(d.vergi_dahil) || 0), tarih: d.tarih || '', aciklama: 'İşçilik' + (d.aciklama ? ' - ' + d.aciklama : ''), belge_no: d.belge_no || '' }));

  let baskiArr = []; try { baskiArr = typeof baski_nakis_yikama_detay === 'string' ? JSON.parse(baski_nakis_yikama_detay) : (baski_nakis_yikama_detay || []); } catch(e) { baskiArr = []; }
  if (!Array.isArray(baskiArr)) baskiArr = [];
  baskiArr.forEach((d, i) => gEntries.push({ kaynak: gkaynak + 'baski:' + i, cari_id: d.cari_id, tutar: ((parseFloat(d.deger) || 0) + (parseFloat(d.kdv) || 0)) || (parseFloat(d.vergi_dahil) || 0), tarih: d.tarih || '', aciklama: 'Baskı/Nakış/Yıkama' + (d.aciklama ? ' - ' + d.aciklama : ''), belge_no: d.belge_no || '' }));

  // satış toplamı satırları → alacak (müşteri bize borçlu)
  let satisArr = []; try { satisArr = typeof satis_toplam_detay === 'string' ? JSON.parse(satis_toplam_detay) : (satis_toplam_detay || []); } catch(e) { satisArr = []; }
  if (!Array.isArray(satisArr)) satisArr = [];
  satisArr.forEach((d, i) => gEntries.push({ kaynak: gkaynak + 'satis:' + i, tur: 'alacak', cari_id: d.cari_id, tutar: ((parseFloat(d.deger) || 0) + (parseFloat(d.kdv) || 0)) || (parseFloat(d.vergi_dahil) || 0), tarih: d.tarih || '', aciklama: 'Satış' + (d.aciklama ? ' - ' + d.aciklama : ''), belge_no: d.belge_no || '' }));

  gEntries.push({ kaynak: gkaynak + 'nakliye', cari_id: nakliyeCari, tutar: nakliye_bedeli, aciklama: 'Nakliye', belge_no: '' });

  cariSenkron(gkaynak, gEntries);

  res.json({ mesaj: t(getLang(req), 'gerceklesenUpdated'), paketleme_bedeli: paketlemeBedeli });
});

// gerçekleşen excel indir
app.get('/api/gerceklesen/:id/excel', async (req, res) => {
  const row = db.prepare('SELECT * FROM gerceklesen WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ hata: t(getLang(req), 'fiyatNotFound') });

  // kayıtlı excel varsa direkt onu gönder
  if (row.excel_dosya) {
    const dosya = path.join(uploadsDir, path.basename(row.excel_dosya || ''));
    if (fs.existsSync(dosya)) {
      return res.download(dosya);
    }
  }

  // excel yoksa boş maliyet şablonu oluştur
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Maliyet');
  ws.columns = [
    { key: 'A', width: 22 }, { key: 'B', width: 18 }, { key: 'C', width: 18 }, { key: 'D', width: 18 }, { key: 'E', width: 18 }
  ];

  const headerStyle = { font: { bold: true, size: 14, color: { argb: 'FF1E293B' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { bottom: { style: 'medium', color: { argb: 'FF3B82F6' } } } };
  const sectionStyle = { font: { bold: true, size: 12, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } }, alignment: { horizontal: 'left' } };
  const labelStyle = { font: { bold: true, size: 11 }, alignment: { horizontal: 'right' } };
  const inputStyle = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFF0' } }, border: { bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } } }, alignment: { horizontal: 'center' } };
  const totalStyle = { font: { bold: true, size: 13, color: { argb: 'FF059669' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } }, alignment: { horizontal: 'center' } };

  // Başlık
  ws.mergeCells('A1:E1');
  ws.getCell('A1').value = 'MALİYET FORMU';
  ws.getCell('A1').font = { bold: true, size: 16, color: { argb: 'FF1E40AF' } };
  ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 30;

  // Bilgi bölümü
  let r = 3;
  const addRow = (label, value, style) => {
    ws.getCell('A' + r).value = label;
    ws.getCell('A' + r).font = { bold: true };
    ws.mergeCells('B' + r + ':E' + r);
    ws.getCell('B' + r).value = value;
    if (style) ws.getCell('B' + r).style = style;
    r++;
  };

  addRow('Müşteri', row.musteri_adi);
  addRow('Firma', row.firma_adi);
  addRow('Ürün', row.urun_aciklamasi);
  addRow('Birim Fiyat', row.fiyat + ' ' + (row.para_birimi || 'TL'));
  addRow('Adet', row.miktar);

  // Toplam satış
  r++;
  ws.mergeCells('A' + r + ':D' + r);
  ws.getCell('A' + r).value = 'TOPLAM SATIŞ';
  ws.getCell('A' + r).font = { bold: true, size: 12 };
  ws.getCell('E' + r).value = row.satis_toplam || (row.fiyat * row.miktar);
  ws.getCell('E' + r).style = totalStyle;
  r += 2;

  // Maliyet Kalemleri
  ws.mergeCells('A' + r + ':E' + r);
  ws.getCell('A' + r).value = 'MALİYET KALEMLERİ';
  ws.getCell('A' + r).style = sectionStyle;
  r++;

  addRow('Kesim Adedi', row.kesim_adedi || 0);
  addRow('Yüklenen Adet', row.yuklenen_adet || 0);
  addRow('2. Kalite', row.ikinci_kalite || 0);
  addRow('Kumaş Bedeli', (row.kumas_bedeli || 0) + ' TL');
  addRow('Aksesuar Bedeli', (row.aksesuar_bedeli || 0) + ' TL');
  addRow('İşçilik Bedeli', (row.iscilik_bedeli || 0) + ' TL');
  addRow('Baskı/Nakış/Yıkama', (row.baski_nakis_yikama || 0) + ' TL');

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="maliyet_gerceklesen_' + row.id + '.xlsx"');
  await workbook.xlsx.write(res);
  res.end();
});

// gerçekleşen durum güncelle
app.patch('/api/gerceklesen/:id/durum', (req, res) => {
  const { durum } = req.body;
  if (!durum) return res.status(400).json({ hata: t(getLang(req), 'durumRequired') });
  const row = db.prepare('SELECT id FROM gerceklesen WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ hata: t(getLang(req), 'fiyatNotFound') });
  db.prepare('UPDATE gerceklesen SET durum = ? WHERE id = ?').run(durum, req.params.id);
  res.json({ mesaj: t(getLang(req), 'siparisUpdated') });
});

// gerçekleşen sil
app.delete('/api/gerceklesen/:id', (req, res) => {
  db.prepare('DELETE FROM gerceklesen WHERE id = ?').run(req.params.id);
  cariSenkron('gerceklesen:' + req.params.id + ':', []);
  res.json({ mesaj: t(getLang(req), 'fiyatDeleted') });
});

// --- ÖDEME PLANI API ---

// ödemelerdeki "Çek" ve "Nakit" kayıtlarını cariye yansıt (ödeme hareketi olarak)
function odemeCariSenkron(odeme) {
  const prefix = 'odeme:' + odeme.id + ':';
  const entries = [];
  const tur = odeme.odeme_turu || '';
  const yansiyanlar = ['Çek', 'Nakit'];
  if (yansiyanlar.includes(tur)) {
    let cariId = odeme.cari_id ? parseInt(odeme.cari_id) : null;
    // cari seçilmemişse firma adından birebir eşleştir
    if (!cariId && odeme.firma_adi) {
      const kart = db.prepare('SELECT id FROM cari_kartlar WHERE LOWER(cari_adi) = LOWER(?)').get(String(odeme.firma_adi).trim());
      if (kart) cariId = kart.id;
    }
    const isCek = tur === 'Çek';
    entries.push({
      kaynak: prefix + 'odeme',
      tur: 'odeme',
      cari_id: cariId,
      tutar: odeme.tutar,
      tarih: odeme.tarih || odeme.vade_tarihi || '',
      vade: isCek ? (odeme.vade_tarihi || '') : '',
      odeme_sekli: isCek ? 'Çek' : 'Nakit',
      aciklama: tur + (odeme.aciklama ? ' - ' + odeme.aciklama : ''),
      belge_no: isCek ? (odeme.cek_no || '') : ''
    });
  }
  cariSenkron(prefix, entries);
}

// tüm ödemeleri listele (filtrelerle)
app.get('/api/odemeler', (req, res) => {
  const { durum, vade_baslangic, vade_bitis, odeme_turu } = req.query;
  let sql = 'SELECT * FROM odemeler WHERE 1=1';
  const params = [];

  if (durum) { sql += ' AND durum = ?'; params.push(durum); }
  if (odeme_turu) { sql += ' AND odeme_turu = ?'; params.push(odeme_turu); }
  if (vade_baslangic) { sql += ' AND vade_tarihi >= ?'; params.push(vade_baslangic); }
  if (vade_bitis) { sql += ' AND vade_tarihi <= ?'; params.push(vade_bitis); }

  sql += ' ORDER BY vade_tarihi ASC';
  const rows = db.prepare(sql).all(...params);
  res.json(rows);
});

// haftalık gruplandırılmış ödeme planı (geçmiş haftalar dahil)
app.get('/api/odemeler/haftalik', (req, res) => {
  const bugun = new Date();
  const bugunStr = bugun.toISOString().split('T')[0];

  // tüm ödemeler (iptal hariç) — geçmiş haftalarda ödenen/gecikenler de görünsün
  const rows = db.prepare(
    "SELECT * FROM odemeler WHERE durum != 'İptal' ORDER BY vade_tarihi ASC"
  ).all();

  // en erken ve en geç vadeyi bul
  let minDate = bugunStr;
  let maxDate = bugunStr;
  rows.forEach(r => {
    if (r.vade_tarihi && r.vade_tarihi < minDate) minDate = r.vade_tarihi;
    if (r.vade_tarihi && r.vade_tarihi > maxDate) maxDate = r.vade_tarihi;
  });

  // en az 8 hafta ileri; geriye tüm geçmişi göster (veri ne kadar eskiyse o kadar)
  const sonTarih = new Date(maxDate);
  const ileriHafta = Math.ceil((sonTarih - bugun) / (7 * 24 * 60 * 60 * 1000));
  const haftaIleri = Math.max(ileriHafta + 2, 8);
  const minDateObj = new Date(minDate);
  const geriHafta = Math.ceil((bugun - minDateObj) / (7 * 24 * 60 * 60 * 1000));
  const haftaGeri = geriHafta + 1;

  // pazartesi tabanlı hafta dizisi oluştur (en eski vadeden en ileri vadeye)
  const haftalar = [];
  const baslangicBugun = new Date(bugun);
  baslangicBugun.setDate(bugun.getDate() - bugun.getDay() + 1 - (haftaGeri * 7));
  const toplamHafta = haftaGeri + haftaIleri;
  for (let i = 0; i < toplamHafta && i < 1200; i++) {
    const baslangic = new Date(baslangicBugun);
    baslangic.setDate(baslangicBugun.getDate() + (i * 7));
    const bitis = new Date(baslangic);
    bitis.setDate(baslangic.getDate() + 6);

    haftalar.push({
      baslangic: baslangic.toISOString().split('T')[0],
      bitis: bitis.toISOString().split('T')[0],
      odemeler: [],
      toplam_tutar: 0,
      odenen_tutar: 0
    });
  }

  const gecikmis = [];

  rows.forEach(row => {
    const vade = row.vade_tarihi || '';
    // gecikmiş: vadesi geçmiş ama hâlâ ödenmemiş
    if (row.durum === 'Gecikti' || (row.durum === 'Bekliyor' && vade && vade < bugunStr)) {
      gecikmis.push(row);
    }

    for (const h of haftalar) {
      if (vade && vade >= h.baslangic && vade <= h.bitis) {
        h.odemeler.push(row);
        if (row.durum === 'Bekliyor' || row.durum === 'Gecikti') {
          h.toplam_tutar += row.tutar;
        } else if (row.durum === 'Tahsil Edildi') {
          h.odenen_tutar += row.tutar;
        }
        break;
      }
    }
  });

  // gelecek toplam (bugünden sonraki bekleyen/gecikenler)
  const gelecek_toplam = rows
    .filter(r => r.vade_tarihi >= bugunStr && (r.durum === 'Bekliyor' || r.durum === 'Gecikti'))
    .reduce((sum, r) => sum + r.tutar, 0);

  res.json({
    haftalar,
    gecikmis,
    gecikmis_toplam: gecikmis.reduce((sum, r) => sum + r.tutar, 0),
    gelecek_toplam,
    bugun: bugunStr
  });
});

// haftalık birleşik gelir-gider görünümü
app.get('/api/odemeler/haftalik-birlestir', (req, res) => {
  const bugun = new Date();
  const bugunStr = bugun.toISOString().split('T')[0];

  // tüm ödemeleri getir (tüm durumlar, geçmiş görünsün diye)
  const odemeler = db.prepare(
    'SELECT id, firma_adi, odeme_turu, tutar, para_birimi, vade_tarihi, durum, cek_no, banka FROM odemeler ORDER BY vade_tarihi ASC'
  ).all();

  // tüm gelirleri getir
  const gelirler = db.prepare(
    'SELECT id, musteri_adi, firma_adi, toplam_tutar, nakit_tutar, cek_toplam, tarih FROM gelir_takip ORDER BY tarih ASC'
  ).all();

  // tarih aralığını belirle
  let minDate = bugunStr;
  let maxDate = bugunStr;

  odemeler.forEach(o => {
    if (o.vade_tarihi && o.vade_tarihi < minDate) minDate = o.vade_tarihi;
    if (o.vade_tarihi && o.vade_tarihi > maxDate) maxDate = o.vade_tarihi;
  });

  gelirler.forEach(g => {
    if (g.tarih && g.tarih < minDate) minDate = g.tarih;
    if (g.tarih && g.tarih > maxDate) maxDate = g.tarih;
  });

  // en az 8 hafta ileriyi göster, en fazla 52 hafta geriyi
  const sonTarih = new Date(maxDate);
  const ileriHafta = Math.ceil((sonTarih - bugun) / (7 * 24 * 60 * 60 * 1000));
  const haftaIleri = Math.max(ileriHafta + 2, 8);

  const minDateObj = new Date(minDate);
  const geriHafta = Math.ceil((bugun - minDateObj) / (7 * 24 * 60 * 60 * 1000));
  const haftaGeri = Math.min(geriHafta + 1, 52);

  // pazartesi tabanlı hafta dizisi oluştur
  const haftalar = [];
  const baslangicBugun = new Date(bugun);
  baslangicBugun.setDate(bugun.getDate() - bugun.getDay() + 1 - (haftaGeri * 7)); // haftaGeri hafta öncesinin pazartesisi

  const toplamHafta = haftaGeri + haftaIleri;
  for (let i = 0; i < toplamHafta; i++) {
    const baslangic = new Date(baslangicBugun);
    baslangic.setDate(baslangicBugun.getDate() + (i * 7));
    const bitis = new Date(baslangic);
    bitis.setDate(baslangic.getDate() + 6);

    haftalar.push({
      baslangic: baslangic.toISOString().split('T')[0],
      bitis: bitis.toISOString().split('T')[0],
      gelir_kayitlari: [],
      gider_kayitlari: [],
      gelir_toplam: 0,
      gider_toplam: 0,
      bakiye: 0
    });
  }

  // ödemeleri haftalara dağıt
  odemeler.forEach(o => {
    if (!o.vade_tarihi) return;
    for (const h of haftalar) {
      if (o.vade_tarihi >= h.baslangic && o.vade_tarihi <= h.bitis) {
        h.gider_kayitlari.push(o);
        h.gider_toplam += o.tutar;
        break;
      }
    }
  });

  // gelirleri haftalara dağıt
  gelirler.forEach(g => {
    if (!g.tarih) return;
    for (const h of haftalar) {
      if (g.tarih >= h.baslangic && g.tarih <= h.bitis) {
        h.gelir_kayitlari.push(g);
        h.gelir_toplam += g.toplam_tutar || 0;
        break;
      }
    }
  });

  // bakiyeleri hesapla
  haftalar.forEach(h => {
    h.bakiye = h.gelir_toplam - h.gider_toplam;
  });

  // geçmiş ve gelecek haftaları ayır
  const gecmis_haftalar = haftalar.filter(h => h.bitis < bugunStr);
  const gelecek_haftalar = haftalar.filter(h => h.bitis >= bugunStr);

  // özet
  const toplam_gelir = gelirler.reduce((s, g) => s + (g.toplam_tutar || 0), 0);
  const toplam_gider = odemeler.reduce((s, o) => s + o.tutar, 0);

  res.json({
    bugun: bugunStr,
    gecmis_haftalar,
    gelecek_haftalar,
    haftalar,
    ozet: {
      toplam_gelir,
      toplam_gider,
      net_bakiye: toplam_gelir - toplam_gider
    }
  });
});

// aylık birleşik gelir-gider görünümü
app.get('/api/odemeler/aylik-birlestir', (req, res) => {
  const bugun = new Date();
  const bugunStr = bugun.toISOString().split('T')[0];
  const buAy = bugun.getMonth();
  const buYil = bugun.getFullYear();

  // tüm ödemeleri getir
  const odemeler = db.prepare(
    'SELECT id, firma_adi, odeme_turu, tutar, para_birimi, vade_tarihi, durum, cek_no, banka, ana_para, faiz_tutari, bsmv_tutari, tarih, kredi_grup_id FROM odemeler ORDER BY vade_tarihi ASC'
  ).all();

  // kredileri giderden ayır; kredi anaparası gelir, faiz/bsmv'si gider
  const krediTurleri = ['Taksitli Kredi', 'Spot Kredi', 'Rotatif Kredi'];
  const krediler = [];
  const odemelerGider = [];
  odemeler.forEach(o => {
    if (krediTurleri.includes(o.odeme_turu)) krediler.push(o);
    else odemelerGider.push(o);
  });

  const krediGruplari = {};
  krediler.forEach(r => {
    if (!r.tarih) return; // eski (tarihsiz) kredileri atla
    const key = r.kredi_grup_id || ('id_' + r.id);
    if (!krediGruplari[key]) krediGruplari[key] = { firma_adi: r.firma_adi, odeme_turu: r.odeme_turu, tutar: 0, ana_para: 0, tarih: r.tarih };
    const g = krediGruplari[key];
    g.tutar += (r.tutar || 0);
    g.ana_para += (r.ana_para || 0);
    if (r.tarih < g.tarih) g.tarih = r.tarih;
  });
  const krediGelirleri = Object.values(krediGruplari).map(g => ({
    firma_adi: g.firma_adi,
    odeme_turu: g.odeme_turu,
    tutar: g.ana_para > 0 ? g.ana_para : g.tutar,
    tarih: g.tarih
  }));

  // tüm gelirleri getir (cari alınan ödemeler; çekler vade tarihine göre, iskonto edilenler net + iskonto tarihiyle yansır)
  const gelirlerRaw = db.prepare(`
    SELECT h.id, c.cari_adi AS musteri_adi, c.cari_adi AS firma_adi, h.tutar AS toplam_tutar, h.tutar AS nakit_tutar, 0 AS cek_toplam, '' AS cek_detay, COALESCE(NULLIF(h.vade,''), h.tarih) AS tarih, '' AS odeme_turu, '[]' AS tfs_detay, h.durum, h.iskonto_detay
    FROM cari_hareketler h
    LEFT JOIN cari_kartlar c ON c.id = h.cari_id
    WHERE h.tur = 'tahsilat'
    ORDER BY COALESCE(NULLIF(h.vade,''), h.tarih) ASC
  `).all();

  const gelirler = gelirlerRaw.map(g => {
    if (g.durum === 'iskonto_edildi' && g.iskonto_detay) {
      try {
        const d = JSON.parse(g.iskonto_detay);
        if (parseFloat(d.net_tutar) > 0) {
          g.toplam_tutar = d.net_tutar;
          g.nakit_tutar = d.net_tutar;
          if (d.tarih) g.tarih = d.tarih;
        }
      } catch(e) {}
    }
    return g;
  });

  // çekleri kendi vade tarihlerine göre ayrıştır
  const gelirCekleri = [];
  gelirler.forEach(g => {
    if (g.cek_toplam > 0) {
      let cekler = [];
      try { cekler = typeof g.cek_detay === 'string' ? JSON.parse(g.cek_detay) : (g.cek_detay || []); } catch(e) { cekler = []; }
      if (!Array.isArray(cekler)) cekler = [];
      cekler.forEach(c => {
        if (c.tutar > 0) {
          // teminata verilmiş ve kredi kullanılmışsa
          if (c.durum === 'teminata_verildi' && c.kredi_tutar > 0 && c.el_gecis_tarihi) {
            const netTutar = c.net_el_gecen || (c.kredi_tutar - (c.kredi_faiz || 0) - (c.teminat_masraf || 0));
            // 1) Kredi net tutarı eline geçiş tarihinde
            if (netTutar > 0) {
              gelirCekleri.push({
                gelir_id: g.id,
                musteri_adi: g.musteri_adi,
                firma_adi: g.firma_adi,
                tutar: netTutar,
                vade_tarihi: c.el_gecis_tarihi,
                durum: c.durum,
                banka: c.banka || '',
                cek_no: c.cek_no || '',
                aciklama: 'Kredi (net)'
              });
            }
            // 2) Kalan tutar (çek - kredi_tutar) çekin kendi vadesinde
            const kalan = c.tutar - c.kredi_tutar;
            if (kalan > 0 && c.vade_tarihi) {
              gelirCekleri.push({
                gelir_id: g.id,
                musteri_adi: g.musteri_adi,
                firma_adi: g.firma_adi,
                tutar: kalan,
                vade_tarihi: c.vade_tarihi,
                durum: c.durum,
                banka: c.banka || '',
                cek_no: c.cek_no || '',
                aciklama: 'Kredi sonrası kalan'
              });
            }
          } else if (c.vade_tarihi) {
            // normal çek: kendi vade tarihi ve tutarı
            gelirCekleri.push({
              gelir_id: g.id,
              musteri_adi: g.musteri_adi,
              firma_adi: g.firma_adi,
              tutar: c.tutar,
              vade_tarihi: c.vade_tarihi,
              durum: c.durum || 'bekliyor',
              banka: c.banka || '',
              cek_no: c.cek_no || ''
            });
          }
        }
      });
    }
  });

  // tarih aralığını belirle (çek vade tarihlerini de hesaba kat)
  let minDate = bugunStr;
  let maxDate = bugunStr;

  odemeler.forEach(o => {
    if (o.vade_tarihi && o.vade_tarihi < minDate) minDate = o.vade_tarihi;
    if (o.vade_tarihi && o.vade_tarihi > maxDate) maxDate = o.vade_tarihi;
  });

  gelirler.forEach(g => {
    if (g.tarih && g.tarih < minDate) minDate = g.tarih;
    if (g.tarih && g.tarih > maxDate) maxDate = g.tarih;
  });

  gelirCekleri.forEach(c => {
    if (c.vade_tarihi && c.vade_tarihi < minDate) minDate = c.vade_tarihi;
    if (c.vade_tarihi && c.vade_tarihi > maxDate) maxDate = c.vade_tarihi;
  });

  krediGelirleri.forEach(k => {
    if (k.tarih && k.tarih < minDate) minDate = k.tarih;
    if (k.tarih && k.tarih > maxDate) maxDate = k.tarih;
  });

  // en az 12 ay ileriyi, en fazla 24 ay geriyi göster
  const minDateObj = new Date(minDate);
  const maxDateObj = new Date(maxDate);

  const baslangicYil = minDateObj.getFullYear();
  const baslangicAy = minDateObj.getMonth();
  const bitisYil = Math.max(maxDateObj.getFullYear(), buYil + 1);
  const bitisAy = Math.max(maxDateObj.getMonth(), buAy + 1);

  // ay dizisi oluştur
  const aylar = [];
  let y = baslangicYil;
  let a = baslangicAy;
  const toplamAy = (bitisYil - baslangicYil) * 12 + (bitisAy - baslangicAy) + 2;
  const maxGeriAy = 24;

  for (let i = 0; i < Math.min(toplamAy, maxGeriAy + 12); i++) {
    const bStr = y + '-' + String(a + 1).padStart(2, '0') + '-01';
    const sonGun = new Date(y, a + 1, 0).getDate();
    const btStr = y + '-' + String(a + 1).padStart(2, '0') + '-' + String(sonGun).padStart(2, '0');
    aylar.push({
      baslangic: bStr,
      bitis: btStr,
      ay: a + 1,
      yil: y,
      gelir_kayitlari: [],
      gider_kayitlari: [],
      gelir_toplam: 0,
      gider_toplam: 0,
      genel_gider_toplam: 0,
      bakiye: 0
    });
    a++;
    if (a > 11) { a = 0; y++; }
  }

  // ödemeleri aylara dağıt (kredi kapama ve ara ödemeleri dahil gider olarak)
  odemeler.forEach(o => {
    if (!o.vade_tarihi) return;
    for (const h of aylar) {
      if (o.vade_tarihi >= h.baslangic && o.vade_tarihi <= h.bitis) {
        h.gider_kayitlari.push(o);
        h.gider_toplam += o.tutar;
        break;
      }
    }
  });

  // gelirleri aylara dağıt (çekler kendi vade tarihine göre, TFS bölünmüş)
  gelirler.forEach(g => {
    // TFS varsa: TFS kısmı TFS tarihinde, kalan vade tarihinde
    if (g.odeme_turu === 'Vadeli Ödeme') {
      let tfsler = [];
      try { tfsler = typeof g.tfs_detay === 'string' ? JSON.parse(g.tfs_detay) : (g.tfs_detay || []); } catch(e) { tfsler = []; }
      if (!Array.isArray(tfsler)) tfsler = [];
      let toplamTfs = 0;
      tfsler.forEach(t => {
        if (t.tfs_tarihi && t.net_tutar > 0) {
          for (const h of aylar) {
            if (t.tfs_tarihi >= h.baslangic && t.tfs_tarihi <= h.bitis) {
              h.gelir_kayitlari.push({
                id: g.id,
                musteri_adi: g.musteri_adi,
                firma_adi: g.firma_adi,
                toplam_tutar: t.net_tutar,
                nakit_tutar: t.net_tutar,
                tarih: t.tfs_tarihi
              });
              h.gelir_toplam += t.net_tutar;
              break;
            }
          }
          toplamTfs += (t.net_tutar || 0) + (t.komisyon || 0);
        }
      });
      // kalan kısım vade tarihinde
      const kalan = (g.nakit_tutar || g.toplam_tutar || 0) - toplamTfs;
      if (kalan > 0 && g.tarih) {
        for (const h of aylar) {
          if (g.tarih >= h.baslangic && g.tarih <= h.bitis) {
            h.gelir_kayitlari.push({
              id: g.id,
              musteri_adi: g.musteri_adi,
              firma_adi: g.firma_adi,
              toplam_tutar: kalan,
              nakit_tutar: kalan,
              tarih: g.tarih
            });
            h.gelir_toplam += kalan;
            break;
          }
        }
      }
      return; // bu gelir kaydı işlendi
    }

    // nakit kısmı kaydın kendi tarihine
    const nakitTutar = g.nakit_tutar || 0;
    if (nakitTutar > 0 && g.tarih) {
      for (const h of aylar) {
        if (g.tarih >= h.baslangic && g.tarih <= h.bitis) {
          // nakit kısmı için sanal kayıt
          h.gelir_kayitlari.push({
            id: g.id,
            musteri_adi: g.musteri_adi,
            firma_adi: g.firma_adi,
            toplam_tutar: nakitTutar,
            nakit_tutar: nakitTutar,
            cek_toplam: 0,
            tarih: g.tarih
          });
          h.gelir_toplam += nakitTutar;
          break;
        }
      }
    }
  });

  // çekleri kendi vade tarihlerine göre dağıt
  gelirCekleri.forEach(c => {
    for (const h of aylar) {
      if (c.vade_tarihi >= h.baslangic && c.vade_tarihi <= h.bitis) {
        // çek için sanal kayıt
        h.gelir_kayitlari.push({
          id: c.gelir_id,
          musteri_adi: c.musteri_adi,
          firma_adi: c.firma_adi,
          toplam_tutar: c.tutar,
          nakit_tutar: 0,
          cek_toplam: c.tutar,
          tarih: c.vade_tarihi
        });
        h.gelir_toplam += c.tutar;
        break;
      }
    }
  });

  // kredileri aylık gelire ekle (gelir listesinde de görünsün)
  krediGelirleri.forEach(k => {
    for (const h of aylar) {
      if (k.tarih >= h.baslangic && k.tarih <= h.bitis) {
        h.gelir_kayitlari.push({
          id: 0,
          musteri_adi: k.odeme_turu,
          firma_adi: k.firma_adi,
          toplam_tutar: k.tutar,
          nakit_tutar: k.tutar,
          cek_toplam: 0,
          tarih: k.tarih
        });
        h.gelir_toplam += k.tutar;
        break;
      }
    }
  });

  // bakiyeleri hesapla
  aylar.forEach(h => {
    h.bakiye = h.gelir_toplam - h.gider_toplam;
    // genel gider: Kredi Kartı, Kira, Maaş, Vergi
    h.genel_gider_toplam = h.gider_kayitlari
      .filter(o => o.odeme_turu === 'Kredi Kartı' || o.odeme_turu === 'Kira' || o.odeme_turu === 'Maaş' || o.odeme_turu === 'Vergi')
      .reduce((s, o) => s + o.tutar, 0);
  });

  // bu ayı bul (yerel saatle)
  const buAyStr = buYil + '-' + String(buAy + 1).padStart(2, '0') + '-01';

  // geçmiş ve gelecek ayları ayır
  const gecmis_aylar = aylar.filter(h => h.baslangic < buAyStr);
  const gelecek_aylar = aylar.filter(h => h.baslangic >= buAyStr);

  // özet
  const toplam_gelir = gelirler.reduce((s, g) => s + (g.toplam_tutar || 0), 0) + krediGelirleri.reduce((s, k) => s + k.tutar, 0);
  const toplam_gider = odemeler.reduce((s, o) => s + o.tutar, 0);

  res.json({
    bugun: bugunStr,
    buAy: buAyStr,
    gecmis_aylar,
    gelecek_aylar,
    aylar,
    ozet: {
      toplam_gelir,
      toplam_gider,
      net_bakiye: toplam_gelir - toplam_gider
    }
  });
});

// genel gider yıllık toplamları (Kredi Kartı, Kira, Maaş, Vergi)
app.get('/api/odemeler/genel-gider-yillik', (req, res) => {
  const rows = db.prepare(`
    SELECT vade_tarihi, tutar FROM odemeler
    WHERE odeme_turu IN ('Kredi Kartı', 'Kira', 'Maaş', 'Vergi')
  `).all();
  const yillik = {};
  rows.forEach(r => {
    const yil = (r.vade_tarihi || '').substring(0, 4);
    if (!yil || yil === '') return;
    if (!yillik[yil]) yillik[yil] = 0;
    yillik[yil] += r.tutar;
  });
  res.json(yillik);
});

// aylık kredi raporu (Taksitli Kredi + Rotatif Kredi)
app.get('/api/odemeler/kredi-aylik', (req, res) => {
  const rows = db.prepare(`
    SELECT * FROM odemeler
    WHERE (odeme_turu = 'Taksitli Kredi' OR odeme_turu = 'Rotatif Kredi')
      AND vade_tarihi >= '2024-12-01'
    ORDER BY vade_tarihi ASC
  `).all();

  // aya göre grupla
  const aylik = {};
  rows.forEach(r => {
    if (!r.vade_tarihi) return;
    const parts = r.vade_tarihi.split('-');
    const key = parts[0] + '-' + parts[1]; // YYYY-MM
    if (!aylik[key]) {
      aylik[key] = {
        yil: parseInt(parts[0]),
        ay: parseInt(parts[1]),
        ana_para: 0,
        faiz_tutari: 0,
        bsmv_tutari: 0,
        toplam_tutar: 0,
        adet: 0,
        odemeler: []
      };
    }
    aylik[key].ana_para += (r.ana_para || 0);
    aylik[key].faiz_tutari += (r.faiz_tutari || 0);
    aylik[key].bsmv_tutari += (r.bsmv_tutari || 0);
    aylik[key].toplam_tutar += r.tutar;
    aylik[key].adet++;
    aylik[key].odemeler.push(r);
  });

  // sıralı dizi
  const aylar = Object.values(aylik).sort((a, b) => a.yil !== b.yil ? a.yil - b.yil : a.ay - b.ay);

  // toplamlar
  const toplam = {
    ana_para: aylar.reduce((s, a) => s + a.ana_para, 0),
    faiz_tutari: aylar.reduce((s, a) => s + a.faiz_tutari, 0),
    bsmv_tutari: aylar.reduce((s, a) => s + a.bsmv_tutari, 0),
    toplam_tutar: aylar.reduce((s, a) => s + a.toplam_tutar, 0),
    adet: aylar.reduce((s, a) => s + a.adet, 0)
  };

  res.json({ aylar, toplam });
});

// tek ödeme getir
app.get('/api/odemeler/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM odemeler WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ hata: t(getLang(req), 'fiyatNotFound') });
  res.json(row);
});

// yeni ödeme ekle
app.post('/api/odemeler', (req, res) => {
  const { musteri_adi, firma_adi, odeme_turu, tutar, para_birimi, vade_tarihi, aciklama, cek_no, banka, cek_sahibi, kredi_grup_id, ana_para, faiz_tutari, bsmv_tutari, tarih, cari_id, rotatif_ay } = req.body;

  if (!tutar || !vade_tarihi) {
    return res.status(400).json({ hata: t(getLang(req), 'odemeRequired') });
  }

  const cariId = cari_id ? parseInt(cari_id) : null;
  let firma = firma_adi || '';
  if (cariId) {
    const kart = db.prepare('SELECT cari_adi FROM cari_kartlar WHERE id = ?').get(cariId);
    if (kart) firma = kart.cari_adi || '';
  }

  const stmt = db.prepare(`
    INSERT INTO odemeler (musteri_adi, firma_adi, odeme_turu, tutar, para_birimi, vade_tarihi, aciklama, cek_no, banka, cek_sahibi, kredi_grup_id, ana_para, faiz_tutari, bsmv_tutari, tarih, cari_id, rotatif_ay)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const info = stmt.run(
    '',
    firma,
    odeme_turu || 'Rotatif Kredi',
    tutar,
    para_birimi || 'TL',
    vade_tarihi,
    aciklama || '',
    cek_no || '',
    banka || '',
    cek_sahibi || '',
    kredi_grup_id || '',
    ana_para || 0,
    faiz_tutari || 0,
    bsmv_tutari || 0,
    tarih || '',
    cariId,
    rotatif_ay || null
  );

  odemeCariSenkron({ id: info.lastInsertRowid, odeme_turu, firma_adi: firma, tutar, tarih, vade_tarihi, cek_no, aciklama, cari_id: cariId });

  res.status(201).json({ id: info.lastInsertRowid, mesaj: t(getLang(req), 'odemeSaved') });
});

// ödeme güncelle
app.put('/api/odemeler/:id', (req, res) => {
  const { musteri_adi, firma_adi, odeme_turu, tutar, para_birimi, vade_tarihi, aciklama, durum, cek_no, banka, cek_sahibi, kredi_grup_id, ana_para, faiz_tutari, bsmv_tutari, tarih, cari_id, rotatif_ay } = req.body;

  if (!tutar || !vade_tarihi) {
    return res.status(400).json({ hata: t(getLang(req), 'odemeRequired') });
  }

  const eski = db.prepare('SELECT * FROM odemeler WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'fiyatNotFound') });

  const cariId = cari_id ? parseInt(cari_id) : null;
  let firma = firma_adi || '';
  if (cariId) {
    const kart = db.prepare('SELECT cari_adi FROM cari_kartlar WHERE id = ?').get(cariId);
    if (kart) firma = kart.cari_adi || '';
  }

  db.prepare(`
    UPDATE odemeler
    SET musteri_adi='', firma_adi=?, odeme_turu=?, tutar=?, para_birimi=?, vade_tarihi=?, aciklama=?, durum=?, cek_no=?, banka=?, cek_sahibi=?, kredi_grup_id=?, ana_para=?, faiz_tutari=?, bsmv_tutari=?, tarih=?, cari_id=?, rotatif_ay=?
    WHERE id=?
  `).run(
    firma,
    odeme_turu || 'Rotatif Kredi',
    tutar,
    para_birimi || 'TL',
    vade_tarihi,
    aciklama || '',
    durum || 'Bekliyor',
    cek_no || '',
    banka || '',
    cek_sahibi || '',
    kredi_grup_id || '',
    ana_para || 0,
    faiz_tutari || 0,
    bsmv_tutari || 0,
    tarih || '',
    cariId,
    rotatif_ay || null,
    req.params.id
  );

  // Aynı kredi grubundaki diğer taksitlerin carisini de güncelle (tek banka → tek cari)
  const grupId = String(kredi_grup_id || eski.kredi_grup_id || '');
  if (grupId) {
    db.prepare('UPDATE odemeler SET cari_id = ?, firma_adi = ? WHERE kredi_grup_id = ? AND id != ?')
      .run(cariId, firma, grupId, req.params.id);
  }

  odemeCariSenkron({ id: req.params.id, odeme_turu, firma_adi: firma, tutar, tarih, vade_tarihi, cek_no, aciklama, cari_id: cariId });

  res.json({ mesaj: t(getLang(req), 'odemeUpdated') });
});

// ödeme sil
app.delete('/api/odemeler/:id', (req, res) => {
  db.prepare('DELETE FROM odemeler WHERE id = ?').run(req.params.id);
  cariSenkron('odeme:' + req.params.id + ':', []);
  res.json({ mesaj: t(getLang(req), 'fiyatDeleted') });
});

// ödeme durumu güncelle (hızlı)
app.patch('/api/odemeler/:id/durum', (req, res) => {
  const { durum } = req.body;
  if (!durum) return res.status(400).json({ hata: 'durum zorunludur' });
  db.prepare('UPDATE odemeler SET durum=? WHERE id=?').run(durum, req.params.id);
  res.json({ mesaj: t(getLang(req), 'siparisStatusUpdated') });
});

// --- KULLANICI YÖNETİMİ ---

app.get('/api/kullanicilar', (req, res) => {
  const rows = db.prepare('SELECT id, username, rol, permissions, created_at FROM kullanicilar ORDER BY created_at ASC').all();
  res.json(rows);
});

app.post('/api/kullanicilar', (req, res) => {
  const { username, password, rol, permissions } = req.body || {};
  if (!username || !password) return res.status(400).json({ hata: t(getLang(req), 'userRequired') });
  const mevcut = db.prepare('SELECT id FROM kullanicilar WHERE username = ?').get(username);
  if (mevcut) return res.status(400).json({ hata: t(getLang(req), 'userExists') });
  const perms = permissions ? JSON.stringify(permissions) : JSON.stringify(ROLE_PERMISSIONS[rol || 'satis'] || []);
  const info = db.prepare('INSERT INTO kullanicilar (username, password_hash, rol, permissions) VALUES (?, ?, ?, ?)')
    .run(username, hashPassword(password), rol || 'satis', perms);
  res.status(201).json({ id: info.lastInsertRowid, mesaj: t(getLang(req), 'userCreated') });
});

app.put('/api/kullanicilar/:id', (req, res) => {
  const { username, password, rol, permissions } = req.body || {};
  const user = db.prepare('SELECT * FROM kullanicilar WHERE id = ?').get(req.params.id);
  if (!user) return res.status(404).json({ hata: t(getLang(req), 'userNotFound') });
  if (username && username !== user.username) {
    const mevcut = db.prepare('SELECT id FROM kullanicilar WHERE username = ? AND id != ?').get(username, req.params.id);
    if (mevcut) return res.status(400).json({ hata: t(getLang(req), 'userExists') });
  }
  const perms = permissions !== undefined ? JSON.stringify(permissions) : user.permissions;
  const yeniRol = rol || user.rol;
  const yeniUsername = username || user.username;
  if (password) {
    db.prepare('UPDATE kullanicilar SET username=?, password_hash=?, rol=?, permissions=? WHERE id=?')
      .run(yeniUsername, hashPassword(password), yeniRol, perms, req.params.id);
  } else {
    db.prepare('UPDATE kullanicilar SET username=?, rol=?, permissions=? WHERE id=?')
      .run(yeniUsername, yeniRol, perms, req.params.id);
  }
  res.json({ mesaj: t(getLang(req), 'userUpdated') });
});

app.delete('/api/kullanicilar/:id', (req, res) => {
  const user = db.prepare('SELECT * FROM kullanicilar WHERE id = ?').get(req.params.id);
  if (!user) return res.status(404).json({ hata: t(getLang(req), 'userNotFound') });
  if (user.username === 'admin') return res.status(400).json({ hata: t(getLang(req), 'adminCantDelete') });
  db.prepare('DELETE FROM kullanicilar WHERE id = ?').run(req.params.id);
  res.json({ mesaj: t(getLang(req), 'userDeleted') });
});

// mevcut admin yetkilerini güncelle
db.prepare("UPDATE kullanicilar SET permissions = ? WHERE username = 'admin'")
  .run(JSON.stringify(ROLE_PERMISSIONS.admin));

// --- GELİR TAKİP API ---

// tüm gelir takip kayıtlarını listele
app.get('/api/gelir-takip', (req, res) => {
  const { siparis_id } = req.query;
  let sql = 'SELECT * FROM gelir_takip WHERE 1=1';
  const params = [];

  if (siparis_id) { sql += ' AND siparis_id = ?'; params.push(siparis_id); }

  sql += ' ORDER BY created_at DESC';
  const rows = db.prepare(sql).all(...params);
  res.json(rows);
});

// gelir takip özeti
app.get('/api/gelir-takip/ozet', (req, res) => {
  const rows = db.prepare('SELECT * FROM gelir_takip').all();
  let toplamGelir = 0, toplamNakit = 0, toplamCek = 0, teminatCek = 0, kirdirilanCek = 0, netNakit = 0;

  rows.forEach(r => {
    toplamGelir += r.toplam_tutar || 0;
    toplamNakit += r.nakit_tutar || 0;
    toplamCek += r.cek_toplam || 0;

    let cekler = [];
    try { cekler = typeof r.cek_detay === 'string' ? JSON.parse(r.cek_detay) : (r.cek_detay || []); } catch(e) { cekler = []; }
    if (!Array.isArray(cekler)) cekler = [];

    cekler.forEach(c => {
      const ct = c.tutar || 0;
      if (c.durum === 'teminata_verildi') teminatCek += ct;
      if (c.durum === 'kirdirildi') {
        kirdirilanCek += ct;
        netNakit += (c.net_tutar || 0);
      }
    });
  });

  res.json({
    toplam_gelir: toplamGelir,
    toplam_nakit: toplamNakit,
    toplam_cek: toplamCek,
    teminata_verilen: teminatCek,
    kirdirilan_toplam: kirdirilanCek,
    kirdirilan_net: netNakit,
    net_eldeki: toplamNakit + netNakit
  });
});

// tek gelir takip kaydı getir
app.get('/api/gelir-takip/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM gelir_takip WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ hata: t(getLang(req), 'gelirTakipNotFound') });
  res.json(row);
});

// yeni gelir takip kaydı ekle (manuel)
app.post('/api/gelir-takip', (req, res) => {
  const { siparis_id, musteri_adi, firma_adi, urun_aciklamasi, toplam_tutar, para_birimi, nakit_tutar, cek_toplam, cek_detay, notlar, odeme_turu, tarih } = req.body;

  if (!musteri_adi) {
    return res.status(400).json({ hata: 'musteri_adi zorunludur' });
  }

  // siparis_id varsa ve > 0 ise duplicate check yap
  const sid = siparis_id || 0;
  if (sid > 0) {
    const mevcut = db.prepare('SELECT id FROM gelir_takip WHERE siparis_id = ?').get(sid);
    if (mevcut) {
      return res.status(409).json({ hata: t(getLang(req), 'gelirTakipExists'), id: mevcut.id });
    }
  }

  const info = db.prepare(`
    INSERT INTO gelir_takip (siparis_id, musteri_adi, firma_adi, urun_aciklamasi, toplam_tutar, para_birimi, nakit_tutar, cek_toplam, cek_detay, notlar, odeme_turu, tarih)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(
    sid, musteri_adi, firma_adi || '', urun_aciklamasi,
    toplam_tutar || 0, para_birimi || 'TL', nakit_tutar || 0,
    cek_toplam || 0, cek_detay || '[]', notlar || '', odeme_turu || '', tarih || ''
  );

  res.status(201).json({ id: info.lastInsertRowid, mesaj: t(getLang(req), 'gelirTakipCreated') });
});

// gelir takip güncelle
app.put('/api/gelir-takip/:id', (req, res) => {
  const { musteri_adi, firma_adi, urun_aciklamasi, toplam_tutar, para_birimi, nakit_tutar, cek_toplam, cek_detay, notlar, odeme_turu, tarih, tfs_detay } = req.body;

  if (!musteri_adi) {
    return res.status(400).json({ hata: t(getLang(req), 'siparisRequired') });
  }

  const eski = db.prepare('SELECT * FROM gelir_takip WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'gelirTakipNotFound') });

  db.prepare(`
    UPDATE gelir_takip
    SET musteri_adi=?, firma_adi=?, urun_aciklamasi=?, toplam_tutar=?, para_birimi=?, nakit_tutar=?, cek_toplam=?, cek_detay=?, notlar=?, odeme_turu=?, tarih=?, tfs_detay=?
    WHERE id=?
  `).run(
    musteri_adi,
    firma_adi || '',
    urun_aciklamasi,
    toplam_tutar || 0,
    para_birimi || 'TL',
    nakit_tutar || 0,
    cek_toplam || 0,
    cek_detay || '[]',
    notlar || '',
    odeme_turu || '',
    tarih || '',
    tfs_detay || '[]',
    req.params.id
  );

  res.json({ mesaj: t(getLang(req), 'gelirTakipUpdated') });
});

// haftalık çek vade görünümü
app.get('/api/gelir-takip/haftalik', (req, res) => {
  const rows = db.prepare('SELECT * FROM gelir_takip').all();
  const bugun = new Date();
  const bugunStr = bugun.toISOString().split('T')[0];

  // tüm çekleri topla
  const tumCekler = [];
  rows.forEach(r => {
    let cekler = [];
    try { cekler = typeof r.cek_detay === 'string' ? JSON.parse(r.cek_detay) : (r.cek_detay || []); } catch(e) { cekler = []; }
    if (!Array.isArray(cekler)) cekler = [];
    cekler.forEach(c => {
      if (c.vade_tarihi && c.tutar > 0) {
        tumCekler.push({
          gelir_id: r.id,
          musteri_adi: r.musteri_adi,
          tutar: c.tutar,
          vade_tarihi: c.vade_tarihi,
          banka: c.banka || '',
          cek_no: c.cek_no || '',
          durum: c.durum || 'bekliyor'
        });
      }
    });
  });

  // en uzak vadeye göre hafta sayısı
  const enUzak = tumCekler.reduce((max, c) => c.vade_tarihi > max ? c.vade_tarihi : max, bugunStr);
  const sonTarih = new Date(enUzak);
  const haftaFarki = Math.ceil((sonTarih - bugun) / (7 * 24 * 60 * 60 * 1000));
  const haftaSayisi = Math.max(haftaFarki + 2, 8);

  const haftalar = [];
  for (let i = 0; i < haftaSayisi; i++) {
    const baslangic = new Date(bugun);
    baslangic.setDate(bugun.getDate() - bugun.getDay() + 1 + (i * 7));
    const bitis = new Date(baslangic);
    bitis.setDate(baslangic.getDate() + 6);
    haftalar.push({ baslangic: baslangic.toISOString().split('T')[0], bitis: bitis.toISOString().split('T')[0], cekler: [], toplam_tutar: 0 });
  }

  const gecikmis = [];

  tumCekler.forEach(c => {
    if (c.vade_tarihi < bugunStr) { gecikmis.push(c); return; }
    for (const h of haftalar) {
      if (c.vade_tarihi >= h.baslangic && c.vade_tarihi <= h.bitis) {
        h.cekler.push(c);
        h.toplam_tutar += c.tutar;
        break;
      }
    }
  });

  res.json({
    haftalar,
    gecikmis,
    gecikmis_toplam: gecikmis.reduce((s, c) => s + c.tutar, 0),
    bugun: bugunStr
  });
});

// gelir takip sil
app.delete('/api/gelir-takip/:id', (req, res) => {
  const eski = db.prepare('SELECT * FROM gelir_takip WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'gelirTakipNotFound') });
  db.prepare('DELETE FROM gelir_takip WHERE id = ?').run(req.params.id);
  res.json({ mesaj: t(getLang(req), 'gelirTakipDeleted') });
});

// --- CARİ (TEDARİKÇİ HESAPLARI) API ---

// tüm cari kartları listele (borç/alacak/bakiye toplamlarıyla)
app.get('/api/cari', (req, res) => {
  const rows = db.prepare(`
    SELECT c.*,
      COALESCE((SELECT SUM(tutar) FROM cari_hareketler h WHERE h.cari_id = c.id AND h.tur = 'borc'), 0)
      - COALESCE((SELECT SUM(tutar) FROM cari_hareketler h WHERE h.cari_id = c.id AND h.tur = 'odeme'), 0) AS borc,
      COALESCE((SELECT SUM(tutar) FROM cari_hareketler h WHERE h.cari_id = c.id AND h.tur = 'alacak'), 0)
      - COALESCE((SELECT SUM(tutar) FROM cari_hareketler h WHERE h.cari_id = c.id AND h.tur = 'tahsilat'), 0) AS alacak
    FROM cari_kartlar c
    ORDER BY c.cari_adi ASC
  `).all();
  rows.forEach(r => { r.bakiye = (r.borc || 0) - (r.alacak || 0); });
  res.json(rows);
});

// cari dökümü (Excel): cari bazlı + aylık + yıllık + hareketler
app.get('/api/cari/dokum', async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    const headerStyle = { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } }, alignment: { horizontal: 'center', vertical: 'middle' } };
    const num = v => (parseFloat(v) || 0);

    // --- 1) Cari Bazlı ---
    const cariler = db.prepare(`
      SELECT c.hesap_turu, c.cari_adi, c.vergi_no,
        COALESCE((SELECT SUM(tutar) FROM cari_hareketler h WHERE h.cari_id = c.id AND h.tur = 'borc'), 0)
        - COALESCE((SELECT SUM(tutar) FROM cari_hareketler h WHERE h.cari_id = c.id AND h.tur = 'odeme'), 0) AS borc,
        COALESCE((SELECT SUM(tutar) FROM cari_hareketler h WHERE h.cari_id = c.id AND h.tur = 'alacak'), 0)
        - COALESCE((SELECT SUM(tutar) FROM cari_hareketler h WHERE h.cari_id = c.id AND h.tur = 'tahsilat'), 0) AS alacak
      FROM cari_kartlar c
    `).all();

    const turSirasi = ['Tedarikçi', 'Üretici', 'Diğer', 'Müşteri'];
    cariler.sort((a, b) => {
      const oa = turSirasi.indexOf(a.hesap_turu); const ob = turSirasi.indexOf(b.hesap_turu);
      const ra = oa === -1 ? 999 : oa; const rb = ob === -1 ? 999 : ob;
      if (ra !== rb) return ra - rb;
      return (a.cari_adi || '').localeCompare(b.cari_adi || '', 'tr');
    });

    const wsCari = workbook.addWorksheet('Cari Bazlı');
    wsCari.addRow(['Hesap Türü', 'Cari Adı', 'Vergi No', 'Borç (₺)', 'Alacak (₺)', 'Bakiye (₺)']);
    cariler.forEach(c => {
      const borc = num(c.borc), alacak = num(c.alacak);
      wsCari.addRow([c.hesap_turu || 'Diğer', c.cari_adi, c.vergi_no || '', borc, alacak, borc - alacak]);
    });
    wsCari.columns = [{ width: 14 }, { width: 30 }, { width: 16 }, { width: 14 }, { width: 14 }, { width: 14 }];

    // --- 2) Aylık + 3) Yıllık ---
    const hareketler = db.prepare("SELECT tarih, tur, tutar FROM cari_hareketler WHERE tarih IS NOT NULL AND tarih != ''").all();
    const aylik = {}, yillik = {};
    hareketler.forEach(h => {
      const ay = (h.tarih || '').substring(0, 7);
      const yil = (h.tarih || '').substring(0, 4);
      const alan = (h.tur === 'borc') ? 'borc' : (h.tur === 'odeme') ? 'odeme' : (h.tur === 'alacak') ? 'alacak' : (h.tur === 'tahsilat') ? 'tahsilat' : null;
      if (!alan) return;
      if (ay.length === 7) { if (!aylik[ay]) aylik[ay] = { borc: 0, odeme: 0, alacak: 0, tahsilat: 0 }; aylik[ay][alan] += num(h.tutar); }
      if (yil.length === 4) { if (!yillik[yil]) yillik[yil] = { borc: 0, odeme: 0, alacak: 0, tahsilat: 0 }; yillik[yil][alan] += num(h.tutar); }
    });

    const net = a => a.borc - a.odeme - a.alacak + a.tahsilat;

    const wsAylik = workbook.addWorksheet('Aylık');
    wsAylik.addRow(['Ay', 'Borç (₺)', 'Ödeme (₺)', 'Alacak (₺)', 'Tahsilat (₺)', 'Net (₺)']);
    Object.keys(aylik).sort().forEach(ay => { const a = aylik[ay]; wsAylik.addRow([ay, a.borc, a.odeme, a.alacak, a.tahsilat, net(a)]); });
    wsAylik.columns = [{ width: 10 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }];

    const wsYillik = workbook.addWorksheet('Yıllık');
    wsYillik.addRow(['Yıl', 'Borç (₺)', 'Ödeme (₺)', 'Alacak (₺)', 'Tahsilat (₺)', 'Net (₺)']);
    Object.keys(yillik).sort().forEach(yil => { const a = yillik[yil]; wsYillik.addRow([yil, a.borc, a.odeme, a.alacak, a.tahsilat, net(a)]); });
    wsYillik.columns = [{ width: 10 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }];

    // --- 4) Hareketler (detay) ---
    const detay = db.prepare(`
      SELECT h.tarih, c.cari_adi, h.tur, h.aciklama, h.belge_no, h.odeme_sekli, h.tutar
      FROM cari_hareketler h
      LEFT JOIN cari_kartlar c ON c.id = h.cari_id
      ORDER BY h.tarih ASC, h.id ASC
    `).all();
    const wsDetay = workbook.addWorksheet('Hareketler');
    wsDetay.addRow(['Tarih', 'Cari', 'Tür', 'Açıklama', 'Belge No', 'Ödeme Şekli', 'Tutar (₺)']);
    detay.forEach(d => { wsDetay.addRow([d.tarih || '', d.cari_adi || '', d.tur || '', d.aciklama || '', d.belge_no || '', d.odeme_sekli || '', num(d.tutar)]); });
    wsDetay.columns = [{ width: 12 }, { width: 24 }, { width: 12 }, { width: 32 }, { width: 16 }, { width: 14 }, { width: 14 }];

    [wsCari, wsAylik, wsYillik, wsDetay].forEach(ws => {
      const header = ws.getRow(1);
      header.eachCell(cell => { cell.font = headerStyle.font; cell.fill = headerStyle.fill; cell.alignment = headerStyle.alignment; });
      ws.views = [{ state: 'frozen', ySplit: 1 }];
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="cari_dokumu.xlsx"');
    await workbook.xlsx.write(res);
    res.end();
  } catch (e) {
    res.status(500).json({ hata: t(getLang(req), 'excelProcessError') + ': ' + e.message });
  }
});

// tüm tedarikçilerin cari ekstrelerini tek PDF'te üret
app.get('/api/cari/tedarikci-ekstre-pdf', async (req, res) => {
  try {
    const kartlar = db.prepare(`
      SELECT c.* FROM cari_kartlar c
      WHERE c.hesap_turu = 'Tedarikçi'
        AND EXISTS (SELECT 1 FROM cari_hareketler h WHERE h.cari_id = c.id)
      ORDER BY c.cari_adi ASC
    `).all();

    const fonts = resolveCariPdfFonts();
    const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 36 });
    if (fonts) { doc.registerFont('Regular', fonts.regular); doc.registerFont('Bold', fonts.bold); }
    const F = fonts ? 'Regular' : 'Helvetica';
    const FB = fonts ? 'Bold' : 'Helvetica-Bold';

    const margin = 36;
    const rightX = doc.page.width - margin;
    const now = new Date();
    const tarihSaat = now.toLocaleDateString('tr-TR', { day: '2-digit', month: 'long', year: 'numeric' }) + ', ' + now.toLocaleDateString('tr-TR', { weekday: 'long' });

    // başlık (ilk sayfa)
    doc.font(FB).fontSize(10).text(COMPANY_NAME, rightX - 220, margin, { width: 220, align: 'right' });
    doc.font(F).fontSize(9).text('Sayfa No : 1', rightX - 220, margin + 14, { width: 220, align: 'right' });
    doc.font(F).fontSize(9);
    doc.text('Rapor Adı           : Tedarikçi Cari Hesap Ekstreleri', margin, margin);
    doc.text('Rapor Tarih / Saat  : ' + tarihSaat, margin, margin + 14);
    doc.text('Kullanıcı Adı       : ' + (req.user ? req.user.username : '-'), margin, margin + 28);
    doc.moveTo(margin, margin + 46).lineTo(rightX, margin + 46).stroke();

    kartlar.forEach((k, i) => {
      if (i > 0) doc.addPage();
      const hareketler = db.prepare('SELECT * FROM cari_hareketler WHERE cari_id = ? ORDER BY tarih ASC, id ASC').all(k.id);
      const startY = (i === 0) ? (margin + 56) : (margin + 16);
      cariEkstreTablo(doc, k, hareketler, F, FB, startY);
    });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="Tedarikci_Cari_Ekstreleri.pdf"');
    doc.pipe(res);
    doc.end();
  } catch (e) {
    res.status(500).json({ hata: t(getLang(req), 'excelProcessError') + ': ' + e.message });
  }
});

// tek cari kartı getir
app.get('/api/cari/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM cari_kartlar WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ hata: t(getLang(req), 'cariNotFound') });
  res.json(row);
});

// yeni cari kartı ekle
app.post('/api/cari', (req, res) => {
  const { hesap_turu, cari_adi, vergi_no, telefon, email, vade, adres, notlar } = req.body;
  if (!cari_adi) return res.status(400).json({ hata: t(getLang(req), 'cariAdiRequired') });
  const info = db.prepare(`
    INSERT INTO cari_kartlar (hesap_turu, cari_adi, vergi_no, telefon, email, vade, adres, notlar)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
  `).run(hesap_turu || 'Tedarikçi', cari_adi, vergi_no || '', telefon || '', email || '', vade || '', adres || '', notlar || '');
  res.status(201).json({ id: info.lastInsertRowid, mesaj: t(getLang(req), 'cariCreated') });
});

// cari kartı güncelle
app.put('/api/cari/:id', (req, res) => {
  const { hesap_turu, cari_adi, vergi_no, telefon, email, vade, adres, notlar } = req.body;
  const eski = db.prepare('SELECT * FROM cari_kartlar WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'cariNotFound') });
  db.prepare(`
    UPDATE cari_kartlar SET hesap_turu=?, cari_adi=?, vergi_no=?, telefon=?, email=?, vade=?, adres=?, notlar=?
    WHERE id=?
  `).run(hesap_turu || 'Tedarikçi', cari_adi || eski.cari_adi, vergi_no || '', telefon || '', email || '', vade || '', adres || '', notlar || '', req.params.id);
  res.json({ mesaj: t(getLang(req), 'cariUpdated') });
});

// cari kartı ve bağlı hareketlerini sil
app.delete('/api/cari/:id', (req, res) => {
  const eski = db.prepare('SELECT * FROM cari_kartlar WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'cariNotFound') });
  db.prepare('DELETE FROM cari_hareketler WHERE cari_id = ?').run(req.params.id);
  db.prepare('DELETE FROM cari_kartlar WHERE id = ?').run(req.params.id);
  res.json({ mesaj: t(getLang(req), 'cariDeleted') });
});

// cari ekstre PDF'i için Türkçe destekli font çözümü (Windows: Arial, Linux: DejaVu)
function resolveCariPdfFonts() {
  const candidates = [
    { regular: 'C:/Windows/Fonts/arial.ttf', bold: 'C:/Windows/Fonts/arialbd.ttf' },
    { regular: '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', bold: '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf' }
  ];
  for (const c of candidates) {
    if (fs.existsSync(c.regular) && fs.existsSync(c.bold)) return c;
  }
  for (const c of candidates) {
    if (fs.existsSync(c.regular)) return { regular: c.regular, bold: c.regular };
  }
  return null;
}
const cariPdfPara = new Intl.NumberFormat('tr-TR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
function cariPdfTarih(s) {
  if (!s) return '';
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[3] + '.' + m[2] + '.' + m[1] : String(s);
}
function cariPdfFisTuru(h) {
  if (h.tur === 'borc') return 'Mal Alım Faturası';
  if (h.tur === 'alacak') return 'Mal Satış Faturası';
  if (h.tur === 'odeme') {
    if (h.odeme_sekli === 'Çek') return 'Kendi Çekimiz Çıkış';
    if (h.odeme_sekli === 'Nakit') return 'Nakit Ödeme';
    return 'Ödeme';
  }
  if (h.tur === 'tahsilat') return 'Tahsilat';
  return h.tur || '';
}

function cariPdfCols() {
  return {
    tarih: { x: 36, w: 74 },
    tur: { x: 110, w: 150 },
    belge: { x: 260, w: 112 },
    aciklama: { x: 372, w: 156 },
    borc: { x: 528, w: 100 },
    alacak: { x: 628, w: 88 },
    bakiye: { x: 716, w: 90 }
  };
}

// Tek cari için ekstre bölümünü çizer (startY'den başlar, taşarsa yeni sayfa açar); bitiş Y döner
function cariEkstreTablo(doc, kart, hareketler, F, FB, startY) {
  const margin = 36;
  const rightX = doc.page.width - margin;
  const pageH = doc.page.height;
  const cols = cariPdfCols();

  const cizBaslik = (y) => {
    doc.rect(margin, y, rightX - margin, 18).fill('#2563eb');
    doc.font(FB).fontSize(9).fillColor('#ffffff');
    doc.text('FİŞ TARİH', cols.tarih.x, y + 5, { width: cols.tarih.w });
    doc.text('FİŞ TÜRÜ', cols.tur.x, y + 5, { width: cols.tur.w });
    doc.text('BELGE NO', cols.belge.x, y + 5, { width: cols.belge.w });
    doc.text('AÇIKLAMA', cols.aciklama.x, y + 5, { width: cols.aciklama.w });
    doc.text('YAPILAN ÖDEME', cols.borc.x, y + 5, { width: cols.borc.w, align: 'right' });
    doc.text('BORÇ', cols.alacak.x, y + 5, { width: cols.alacak.w, align: 'right' });
    doc.text('BAKİYE', cols.bakiye.x, y + 5, { width: cols.bakiye.w, align: 'right' });
    doc.fillColor('#000000');
  };

  const cariKod = String(kart.id).padStart(10, '0');
  doc.font(FB).fontSize(11).text(cariKod + '   ' + (kart.hesap_turu || 'Tedarikçi'), margin, startY);
  doc.font(FB).fontSize(14).text(kart.cari_adi || '', margin, startY + 18);

  let y = startY + 44;
  cizBaslik(y);
  y += 18;

  const yeniSayfa = () => {
    if (y > pageH - margin - 24) {
      doc.addPage();
      y = margin;
      cizBaslik(y);
      y += 18;
    }
  };

  let running = 0, topBorc = 0, topAlacak = 0;

  if (hareketler.length === 0) {
    doc.font(F).fontSize(9).fillColor('#64748b');
    doc.text('Hareket yok.', margin, y + 4);
    y += 22;
  } else {
    hareketler.forEach(h => {
      yeniSayfa();
      const borcVal = (h.tur === 'alacak' || h.tur === 'odeme') ? (parseFloat(h.tutar) || 0) : 0;
      const alacakVal = (h.tur === 'borc' || h.tur === 'tahsilat') ? (parseFloat(h.tutar) || 0) : 0;
      running += borcVal - alacakVal;
      topBorc += borcVal;
      topAlacak += alacakVal;
      const isaret = running < 0 ? '(A)' : (running > 0 ? '(B)' : '(.)');
      doc.font(F).fontSize(9).fillColor('#000000');
      doc.text(cariPdfTarih(h.tarih), cols.tarih.x, y + 4, { width: cols.tarih.w, ellipsis: true });
      doc.text(cariPdfFisTuru(h), cols.tur.x, y + 4, { width: cols.tur.w, ellipsis: true });
      doc.text(h.belge_no || '', cols.belge.x, y + 4, { width: cols.belge.w, ellipsis: true });
      doc.text(h.aciklama || '', cols.aciklama.x, y + 4, { width: cols.aciklama.w, ellipsis: true });
      doc.text(borcVal ? cariPdfPara.format(borcVal) : '', cols.borc.x, y + 4, { width: cols.borc.w, align: 'right' });
      doc.text(alacakVal ? cariPdfPara.format(alacakVal) : '', cols.alacak.x, y + 4, { width: cols.alacak.w, align: 'right' });
      doc.text(cariPdfPara.format(running) + ' ' + isaret, cols.bakiye.x, y + 4, { width: cols.bakiye.w, align: 'right' });
      y += 18;
    });
  }

  yeniSayfa();
  doc.moveTo(margin, y).lineTo(rightX, y).stroke();
  doc.font(FB).fontSize(9).fillColor('#000000');
  doc.text('TOPLAMLAR :', cols.tarih.x, y + 6, { width: cols.borc.x - cols.tarih.x, align: 'right' });
  doc.text(cariPdfPara.format(topBorc), cols.borc.x, y + 6, { width: cols.borc.w, align: 'right' });
  doc.text(cariPdfPara.format(topAlacak), cols.alacak.x, y + 6, { width: cols.alacak.w, align: 'right' });
  doc.font(FB).fontSize(9);
  doc.text(cariPdfPara.format(topBorc - topAlacak) + ' ' + ((topBorc - topAlacak) < 0 ? '(A)' : (topBorc - topAlacak) > 0 ? '(B)' : '(.)'), cols.bakiye.x, y + 6, { width: cols.bakiye.w, align: 'right' });
  y += 26;
  return y;
}

// cari ekstre raporu (PDF) — örnek "Cari Hesap Ekstresi" düzeninde
app.get('/api/cari/:id/ekstre-pdf', async (req, res) => {
  try {
    const kart = db.prepare('SELECT * FROM cari_kartlar WHERE id = ?').get(req.params.id);
    if (!kart) return res.status(404).json({ hata: t(getLang(req), 'cariNotFound') });
    const hareketler = db.prepare('SELECT * FROM cari_hareketler WHERE cari_id = ? ORDER BY tarih ASC, id ASC').all(req.params.id);

    const fonts = resolveCariPdfFonts();
    const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 36 });
    if (fonts) { doc.registerFont('Regular', fonts.regular); doc.registerFont('Bold', fonts.bold); }
    const F = fonts ? 'Regular' : 'Helvetica';
    const FB = fonts ? 'Bold' : 'Helvetica-Bold';

    const pageW = doc.page.width; // 841.89
    const margin = 36;
    const rightX = pageW - margin;
    const now = new Date();
    const tarihSaat = now.toLocaleDateString('tr-TR', { day: '2-digit', month: 'long', year: 'numeric' }) + ', ' + now.toLocaleDateString('tr-TR', { weekday: 'long' });

    // --- Başlık ---
    doc.font(FB).fontSize(10).text(COMPANY_NAME, rightX - 220, margin, { width: 220, align: 'right' });
    doc.font(F).fontSize(9).text('Sayfa No : 1', rightX - 220, margin + 14, { width: 220, align: 'right' });
    doc.font(F).fontSize(9);
    doc.text('Rapor Adı           : Cari Hesap Ekstresi', margin, margin);
    doc.text('Rapor Tarih / Saat  : ' + tarihSaat, margin, margin + 14);
    const kullanici = req.user ? req.user.username : '';
    doc.text('Kullanıcı Adı       : ' + (kullanici || '-'), margin, margin + 28);

    doc.moveTo(margin, margin + 46).lineTo(rightX, margin + 46).stroke();

    // --- Cari bilgileri + tablo ---
    cariEkstreTablo(doc, kart, hareketler, F, FB, margin + 56);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="Cari_Ekstre_' + String(kart.id) + '.pdf"');
    doc.pipe(res);
    doc.end();
  } catch (e) {
    res.status(500).json({ hata: t(getLang(req), 'excelProcessError') + ': ' + e.message });
  }
});

// cari kartına ait hareketleri listele (ekstre)
app.get('/api/cari/:id/hareketler', (req, res) => {
  const rows = db.prepare('SELECT * FROM cari_hareketler WHERE cari_id = ? ORDER BY tarih ASC, id ASC').all(req.params.id);
  res.json(rows);
});

// yeni hareket ekle
app.post('/api/cari/:id/hareketler', (req, res) => {
  const { tarih, tur, tutar, para_birimi, aciklama, belge_no, odeme_sekli, vade } = req.body;
  const kart = db.prepare('SELECT id FROM cari_kartlar WHERE id = ?').get(req.params.id);
  if (!kart) return res.status(404).json({ hata: t(getLang(req), 'cariNotFound') });
  if (!tarih || !tur || tutar == null) return res.status(400).json({ hata: t(getLang(req), 'cariHareketRequired') });
  const info = db.prepare(`
    INSERT INTO cari_hareketler (cari_id, tarih, tur, tutar, para_birimi, aciklama, belge_no, odeme_sekli, vade)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(req.params.id, tarih, tur, tutar || 0, para_birimi || 'TL', aciklama || '', belge_no || '', odeme_sekli || '', vade || '');
  res.status(201).json({ id: info.lastInsertRowid, mesaj: t(getLang(req), 'cariHareketCreated') });
});

// hareket güncelle
app.put('/api/cari/hareketler/:id', (req, res) => {
  const { tarih, tur, tutar, para_birimi, aciklama, belge_no, odeme_sekli, vade } = req.body;
  const eski = db.prepare('SELECT * FROM cari_hareketler WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'cariHareketNotFound') });
  if (eski.kaynak) return res.status(400).json({ hata: t(getLang(req), 'cariHareketAutoEdit') });
  db.prepare(`
    UPDATE cari_hareketler SET tarih=?, tur=?, tutar=?, para_birimi=?, aciklama=?, belge_no=?, odeme_sekli=?, vade=?
    WHERE id=?
  `).run(tarih || eski.tarih, tur || eski.tur, tutar || 0, para_birimi || 'TL', aciklama || '', belge_no || '', odeme_sekli || '', vade || '', req.params.id);
  res.json({ mesaj: t(getLang(req), 'cariHareketUpdated') });
});

// hareket sil
app.delete('/api/cari/hareketler/:id', (req, res) => {
  const eski = db.prepare('SELECT * FROM cari_hareketler WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'cariHareketNotFound') });
  if (eski.kaynak) return res.status(400).json({ hata: t(getLang(req), 'cariHareketAutoDelete') });
  db.prepare('DELETE FROM cari_hareketler WHERE id = ?').run(req.params.id);
  res.json({ mesaj: t(getLang(req), 'cariHareketDeleted') });
});

// ==============================
// FATURALAB (alınan ödemeler + çek iskontosu)
// ==============================
app.use('/api/faturalab', requirePermission('faturalab'));

// alınan ödemeleri listele (vade + iskonto durumuyla)
app.get('/api/faturalab', (req, res) => {
  const rows = db.prepare(`
    SELECT h.*, c.cari_adi
    FROM cari_hareketler h
    LEFT JOIN cari_kartlar c ON c.id = h.cari_id
    WHERE h.tur = 'tahsilat' AND COALESCE(h.kaynak,'') NOT LIKE 'iskonto:%'
    ORDER BY COALESCE(h.vade, h.tarih) ASC, h.id ASC
  `).all();
  const bugun = new Date(); bugun.setHours(0, 0, 0, 0);
  rows.forEach(h => {
    let iskonto = null;
    try { iskonto = h.iskonto_detay ? JSON.parse(h.iskonto_detay) : null; } catch(e) { iskonto = null; }
    h._iskonto = iskonto;
    if (h.vade) {
      const v = new Date(h.vade + 'T00:00:00');
      h._kalanGun = Math.round((v - bugun) / 86400000);
    } else {
      h._kalanGun = null;
    }
  });
  res.json(rows);
});

// çeki iskonto et (nakite çevir)
app.post('/api/faturalab/iskonto/:id', (req, res) => {
  const { net_tutar, faiz, bsmv, komisyon, kalan_tutar, tarih, banka } = req.body;
  const eski = db.prepare('SELECT * FROM cari_hareketler WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'cariHareketNotFound') });
  if (eski.tur !== 'tahsilat') return res.status(400).json({ hata: t(getLang(req), 'faturalabOnlyAlinan') });
  if (eski.durum === 'iskonto_edildi') return res.status(400).json({ hata: t(getLang(req), 'faturalabAlreadyIskonto') });
  const net = parseFloat(net_tutar) || 0;
  const detay = JSON.stringify({
    net_tutar: net,
    faiz: parseFloat(faiz) || 0,
    bsmv: parseFloat(bsmv) || 0,
    komisyon: parseFloat(komisyon) || 0,
    kalan_tutar: parseFloat(kalan_tutar) || 0,
    tarih: tarih || '',
    banka: banka || ''
  });
  db.prepare("UPDATE cari_hareketler SET durum='iskonto_edildi', iskonto_detay=? WHERE id=?").run(detay, req.params.id);
  res.json({ mesaj: 'ok' });
});

// iskonto edilen faturaların aylık raporu
app.get('/api/faturalab/iskonto-aylik', (req, res) => {
  const rows = db.prepare(`
    SELECT h.id, c.cari_adi, h.tutar, h.iskonto_detay, h.tarih
    FROM cari_hareketler h
    LEFT JOIN cari_kartlar c ON c.id = h.cari_id
    WHERE h.tur = 'tahsilat' AND h.durum = 'iskonto_edildi'
    ORDER BY h.tarih ASC
  `).all();

  const aylik = {};
  rows.forEach(r => {
    let d = {};
    try { d = r.iskonto_detay ? JSON.parse(r.iskonto_detay) : {}; } catch(e) { d = {}; }
    const tarih = d.tarih || r.tarih;
    const key = (tarih || '').substring(0, 7);
    if (!key || key.length < 7) return;
    if (!aylik[key]) aylik[key] = { ay: key, adet: 0, tutar: 0, faiz: 0, bsmv: 0, komisyon: 0, net: 0, bankalar: [] };
    const g = aylik[key];
    g.adet++;
    g.tutar += (parseFloat(r.tutar) || 0);
    g.faiz += (parseFloat(d.faiz) || 0);
    g.bsmv += (parseFloat(d.bsmv) || 0);
    g.komisyon += (parseFloat(d.komisyon) || 0);
    g.net += (parseFloat(d.net_tutar) || 0);
    if (d.banka) g.bankalar.push(d.banka);
  });

  const sonuc = Object.values(aylik).sort((a, b) => a.ay.localeCompare(b.ay));
  sonuc.forEach(g => { g.banka = g.bankalar.length ? [...new Set(g.bankalar)].join(', ') : ''; });
  res.json(sonuc);
});

// krediler (ödeme listesindeki kredi tipleri) — gelir takip için
// Aynı kredinin taksitleri (kredi_grup_id) tek satırda toplanır; kredi yalnız alındığı tarihte hesaba geçer.
app.get('/api/krediler', requirePermission('gelir_takip'), (req, res) => {
  const rows = db.prepare(`
    SELECT id, firma_adi, odeme_turu, tutar, ana_para, faiz_tutari, bsmv_tutari, vade_tarihi, durum, kredi_grup_id, tarih
    FROM odemeler
    WHERE odeme_turu IN ('Taksitli Kredi', 'Spot Kredi', 'Rotatif Kredi') AND COALESCE(tarih, '') != ''
    ORDER BY vade_tarihi ASC
  `).all();
  const groups = {};
  const order = [];
  rows.forEach(r => {
    const key = r.kredi_grup_id ? 'g:' + r.kredi_grup_id : 'id:' + r.id;
    if (!groups[key]) {
      groups[key] = { id: r.id, firma_adi: r.firma_adi, odeme_turu: r.odeme_turu, ana_para: 0, faiz_tutari: 0, bsmv_tutari: 0, tutar: 0, tarih: r.tarih || r.vade_tarihi, durum: r.durum, taksit: 0 };
      order.push(key);
    }
    const g = groups[key];
    g.ana_para += (parseFloat(r.ana_para) || 0);
    g.faiz_tutari += (parseFloat(r.faiz_tutari) || 0);
    g.bsmv_tutari += (parseFloat(r.bsmv_tutari) || 0);
    g.tutar += (parseFloat(r.tutar) || 0);
    g.taksit++;
    const t = r.tarih || r.vade_tarihi;
    if (t && t < g.tarih) g.tarih = t;
  });
  res.json(order.map(key => groups[key]));
});

// ==============================
// PAKETLEME MALZEMELERİ (stok + tüketim)
// ==============================
app.use('/api/paketleme', requirePermission('gerceklesen'));

// bir malzemenin birim fiyatı = en son alımın birim fiyatı
function paketlemeBirimFiyat(malzemeId) {
  const row = db.prepare('SELECT birim_fiyat FROM paketleme_alimlar WHERE malzeme_id = ? AND adet > 0 ORDER BY tarih DESC, id DESC LIMIT 1').get(malzemeId);
  return row ? (parseFloat(row.birim_fiyat) || 0) : 0;
}

// tüm gerçekleşen kayıtlarındaki paketleme tüketimini topla (malzeme_id → adet)
function paketlemeTuketimMap() {
  const map = {};
  const rows = db.prepare("SELECT paketleme_malzeme_detay FROM gerceklesen WHERE paketleme_malzeme_detay IS NOT NULL AND paketleme_malzeme_detay != '' AND paketleme_malzeme_detay != '[]'").all();
  rows.forEach(r => {
    let arr = [];
    try { arr = typeof r.paketleme_malzeme_detay === 'string' ? JSON.parse(r.paketleme_malzeme_detay) : (r.paketleme_malzeme_detay || []); } catch(e) { arr = []; }
    if (!Array.isArray(arr)) arr = [];
    arr.forEach(d => {
      const mid = d.malzeme_id;
      if (mid == null) return;
      map[mid] = (map[mid] || 0) + (parseFloat(d.adet) || 0);
    });
  });
  return map;
}

// malzemeleri özetle listele (alınan / tüketilen / kalan / birim fiyat)
app.get('/api/paketleme/malzemeler', (req, res) => {
  const malzemeler = db.prepare('SELECT * FROM paketleme_malzemeleri ORDER BY adi ASC').all();
  const tuketim = paketlemeTuketimMap();
  const ozet = malzemeler.map(m => {
    const alim = db.prepare('SELECT COALESCE(SUM(adet),0) AS adet, COALESCE(SUM(toplam_tutar),0) AS tutar FROM paketleme_alimlar WHERE malzeme_id = ?').get(m.id);
    const alinan = parseFloat(alim.adet) || 0;
    const tutar = parseFloat(alim.tutar) || 0;
    const birim = paketlemeBirimFiyat(m.id); // en son alımın birim fiyatı
    const tuketilen = parseFloat(tuketim[m.id]) || 0;
    return { id: m.id, adi: m.adi, alinan, tutar, birim_fiyat: birim, tuketilen, kalan: alinan - tuketilen };
  });
  res.json(ozet);
});

// malzeme ekle
app.post('/api/paketleme/malzemeler', (req, res) => {
  const { adi } = req.body;
  if (!adi || !String(adi).trim()) return res.status(400).json({ hata: t(getLang(req), 'paketlemeMalzemeAdiRequired') });
  const info = db.prepare('INSERT INTO paketleme_malzemeleri (adi) VALUES (?)').run(String(adi).trim());
  res.status(201).json({ id: info.lastInsertRowid, mesaj: 'ok' });
});

// malzeme güncelle
app.put('/api/paketleme/malzemeler/:id', (req, res) => {
  const { adi } = req.body;
  const eski = db.prepare('SELECT * FROM paketleme_malzemeleri WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'paketlemeMalzemeNotFound') });
  db.prepare('UPDATE paketleme_malzemeleri SET adi = ? WHERE id = ?').run(adi ? String(adi).trim() : eski.adi, req.params.id);
  res.json({ mesaj: 'ok' });
});

// malzeme sil (bağlı alımları ve cari hareketlerini de sil)
app.delete('/api/paketleme/malzemeler/:id', (req, res) => {
  const eski = db.prepare('SELECT * FROM paketleme_malzemeleri WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'paketlemeMalzemeNotFound') });
  const alimlar = db.prepare('SELECT id FROM paketleme_alimlar WHERE malzeme_id = ?').all(req.params.id);
  alimlar.forEach(a => cariSenkron('paketleme_alim:' + a.id + ':', []));
  db.prepare('DELETE FROM paketleme_alimlar WHERE malzeme_id = ?').run(req.params.id);
  db.prepare('DELETE FROM paketleme_malzemeleri WHERE id = ?').run(req.params.id);
  res.json({ mesaj: 'ok' });
});

// toplu alımları listele
app.get('/api/paketleme/alimlar', (req, res) => {
  const rows = db.prepare(`
    SELECT a.*, m.adi AS malzeme_adi, c.cari_adi
    FROM paketleme_alimlar a
    LEFT JOIN paketleme_malzemeleri m ON m.id = a.malzeme_id
    LEFT JOIN cari_kartlar c ON c.id = a.cari_id
    ORDER BY a.tarih DESC, a.id DESC
  `).all();
  res.json(rows);
});

// toplu alım ekle (cari seçilirse → borç)
app.post('/api/paketleme/alimlar', (req, res) => {
  const { malzeme_id, adet, toplam_tutar, cari_id, tarih, belge_no, notlar } = req.body;
  if (!malzeme_id) return res.status(400).json({ hata: t(getLang(req), 'paketlemeMalzemeSeciniz') });
  const ad = parseFloat(adet) || 0;
  const tutar = parseFloat(toplam_tutar) || 0;
  const birim = ad > 0 ? Math.round((tutar / ad) * 10000) / 10000 : 0;
  const info = db.prepare(`
    INSERT INTO paketleme_alimlar (malzeme_id, adet, toplam_tutar, birim_fiyat, cari_id, tarih, belge_no, notlar)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
  `).run(malzeme_id, ad, tutar, birim, cari_id ? parseInt(cari_id) : null, tarih || new Date().toISOString().split('T')[0], belge_no || '', notlar || '');

  const malzemeAdi = db.prepare('SELECT adi FROM paketleme_malzemeleri WHERE id = ?').get(malzeme_id);
  cariSenkron('paketleme_alim:' + info.lastInsertRowid + ':', [{
    kaynak: 'paketleme_alim:' + info.lastInsertRowid + ':',
    cari_id,
    tutar,
    aciklama: 'Paketleme Alımı' + (malzemeAdi ? ' - ' + malzemeAdi.adi : ''),
    belge_no: belge_no || ''
  }]);

  res.status(201).json({ id: info.lastInsertRowid, mesaj: 'ok' });
});

// toplu alım güncelle
app.put('/api/paketleme/alimlar/:id', (req, res) => {
  const { malzeme_id, adet, toplam_tutar, cari_id, tarih, belge_no, notlar } = req.body;
  const eski = db.prepare('SELECT * FROM paketleme_alimlar WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'paketlemeAlimNotFound') });
  const mid = malzeme_id || eski.malzeme_id;
  const ad = parseFloat(adet) || 0;
  const tutar = parseFloat(toplam_tutar) || 0;
  const birim = ad > 0 ? Math.round((tutar / ad) * 10000) / 10000 : 0;
  db.prepare(`
    UPDATE paketleme_alimlar SET malzeme_id=?, adet=?, toplam_tutar=?, birim_fiyat=?, cari_id=?, tarih=?, belge_no=?, notlar=?
    WHERE id=?
  `).run(mid, ad, tutar, birim, cari_id ? parseInt(cari_id) : null, tarih || eski.tarih, belge_no || '', notlar || '', req.params.id);

  const malzemeAdi = db.prepare('SELECT adi FROM paketleme_malzemeleri WHERE id = ?').get(mid);
  cariSenkron('paketleme_alim:' + req.params.id + ':', [{
    kaynak: 'paketleme_alim:' + req.params.id + ':',
    cari_id,
    tutar,
    aciklama: 'Paketleme Alımı' + (malzemeAdi ? ' - ' + malzemeAdi.adi : ''),
    belge_no: belge_no || ''
  }]);

  res.json({ mesaj: 'ok' });
});

// toplu alım sil
app.delete('/api/paketleme/alimlar/:id', (req, res) => {
  const eski = db.prepare('SELECT * FROM paketleme_alimlar WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'paketlemeAlimNotFound') });
  db.prepare('DELETE FROM paketleme_alimlar WHERE id = ?').run(req.params.id);
  cariSenkron('paketleme_alim:' + req.params.id + ':', []);
  res.json({ mesaj: 'ok' });
});

// --- yedekleme ---

// Tam yedek: veritabanı + uploads + bant verileri (tar.gz)
app.get('/api/yedekle', (req, res) => {
  const tarih = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
  res.setHeader('Content-Type', 'application/gzip');
  res.setHeader('Content-Disposition', `attachment; filename="${COMPANY_SLUG}-tam-yedek-${tarih}.tar.gz"`);

  // Veritabanını yedekle (WAL checkpoint)
  db.pragma('wal_checkpoint(TRUNCATE)');

  const tar = spawn('tar', ['-czf', '-', '-C', dataDir, '.']);
  tar.stdout.pipe(res);
  tar.stderr.on('data', () => {});
  tar.on('error', () => {
    if (!res.headersSent) res.status(500).json({ hata: t(getLang(req), 'backupFailed') });
  });
});

// Sadece veritabanı dosyası (.db)
app.get('/api/yedekle/db', (req, res) => {
  const tarih = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
  // WAL checkpoint - tüm değişiklikleri ana db dosyasına yaz
  db.pragma('wal_checkpoint(TRUNCATE)');
  const dbFile = path.join(dataDir, 'fiyatlar.db');
  if (!fs.existsSync(dbFile)) return res.status(404).json({ hata: t(getLang(req), 'fiyatNotFound') });
  res.download(dbFile, `fiyatlar-${tarih}.db`);
});

// Excel yedeği (tüm tablolar tek Excel'de)
app.get('/api/yedekle/excel', async (req, res) => {
  try {
    const tarih = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
    const workbook = new ExcelJS.Workbook();

    const headerStyle = { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } }, alignment: { horizontal: 'center', vertical: 'middle' } };

    const tablolar = [
      { name: 'Fiyatlar', query: 'SELECT * FROM fiyatlar ORDER BY created_at DESC' },
      { name: 'Siparişler', query: 'SELECT * FROM siparisler ORDER BY created_at DESC' },
      { name: 'Gerçekleşen', query: 'SELECT * FROM gerceklesen ORDER BY created_at DESC' },
      { name: 'Ödemeler', query: 'SELECT * FROM odemeler ORDER BY created_at DESC' },
      { name: 'Aksesuar Takip', query: 'SELECT * FROM aksesuar_takip ORDER BY siparis_id, created_at ASC' },
      { name: 'Kumaş Takip', query: 'SELECT * FROM kumas_takip ORDER BY siparis_id, created_at ASC' },
      { name: 'Kullanıcılar', query: 'SELECT id, username, rol, permissions, created_at FROM kullanicilar ORDER BY id ASC' },
      { name: 'Oturumlar', query: 'SELECT * FROM sessions ORDER BY created_at DESC' },
      { name: 'Gelir Takip', query: 'SELECT * FROM gelir_takip ORDER BY created_at DESC' },
    ];

    for (const tablo of tablolar) {
      const ws = workbook.addWorksheet(tablo.name);
      const rows = db.prepare(tablo.query).all();
      if (rows.length === 0) { ws.addRow([`${tablo.name} tablosunda kayıt yok`]); continue; }

      const cols = Object.keys(rows[0]);
      // Başlık satırı
      const baslikRow = ws.addRow(cols);
      baslikRow.font = headerStyle.font;
      baslikRow.fill = headerStyle.fill;
      baslikRow.alignment = headerStyle.alignment;

      // Veri satırları
      for (const row of rows) {
        ws.addRow(cols.map(c => row[c] !== null && row[c] !== undefined ? row[c] : ''));
      }

      // Sütun genişliklerini otomatik ayarla
      cols.forEach((c, i) => {
        const maxLen = Math.max(c.length, ...rows.slice(0, 100).map(r => String(r[c] || '').length));
        ws.getColumn(i + 1).width = Math.min(maxLen + 4, 40);
      });

      ws.views = [{ state: 'frozen', ySplit: 1 }];
      const sonKolon = String.fromCharCode(64 + Math.min(cols.length, 26));
      ws.autoFilter = `A1:${sonKolon}${rows.length + 1}`;
    }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${COMPANY_SLUG}-excel-yedek-${tarih}.xlsx"`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (e) {
    res.status(500).json({ hata: t(getLang(req), 'excelProcessError') + ': ' + e.message });
  }
});

// --- başlat ---
app.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ ${COMPANY_NAME} - Fiyat Takip Sistemi çalışıyor → http://localhost:${PORT}`);
  console.log(`📡 Ağdaki diğer bilgisayarlar için → http://${getLocalIP()}:${PORT}`);
});

function getLocalIP() {
  const nets = os.networkInterfaces();
  for (const name of Object.keys(nets)) {
    for (const net of nets[name]) {
      if (net.family === 'IPv4' && !net.internal) {
        return net.address;
      }
    }
  }
  return 'localhost';
}
