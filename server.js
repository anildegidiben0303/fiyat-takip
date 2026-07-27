const express = require('express');
const Database = require('better-sqlite3');
const path = require('path');
const multer = require('multer');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { spawn } = require('child_process');
const os = require('os');
const crypto = require('crypto');
const { t } = require('./server-i18n');
const swaggerUi = require('swagger-ui-express');

const app = express();
const PORT = process.env.PORT || 3000;

// dil tespiti
function getLang(req) {
  return req.query.lang || ((req.headers['accept-language'] || '').startsWith('en') ? 'en' : 'tr');
}

// --- çok kullanıcılı auth sistemi ---

const ROLE_PERMISSIONS = {
  admin: ['fiyatlar','siparisler','gerceklesen','odemeler','bant','kullanicilar'],
  muhasebe: ['fiyatlar','gerceklesen','odemeler'],
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
  const token = req.headers['x-auth-token'] || '';
  const user = getSession(token);
  if (user) { req.user = user; return next(); }
  res.status(401).json({ hata: t(getLang(req), 'sessionInvalid') });
}

function requirePermission(permission) {
  return (req, res, next) => {
    const token = req.headers['x-auth-token'] || '';
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
    .run('admin', hashPassword('admin123'), 'admin', allPerms);
  console.log('✅ Varsayılan admin kullanıcısı oluşturuldu: admin / admin123');
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

app.use(express.static(path.join(__dirname, 'public')));
app.use('/uploads', express.static(uploadsDir));

// --- API Dökümantasyonu (Swagger) ---
const swaggerDoc = {
  openapi: '3.0.0',
  info: {
    title: 'ESA MODA API',
    version: '1.0.0',
    description: 'Fiyat Takip Sistemi - Tüm endpoint\'ler **x-auth-token** header\'ı gerektirir (login hariç).'
  },
  servers: [{ url: 'https://esamoda.fly.dev', description: 'Production' }],
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
      get: { tags: ['Sistem'], summary: 'Veritabanını yedekle (tar.gz)', responses: { '200': { description: 'Yedek dosyası' } } }
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
  const { fiyat_id, termin, renk, miktar, renk_detay, fiyat } = req.body;
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
    INSERT INTO siparisler (musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, tarih, notlar, gorsel, excel_dosya, termin, renk, renk_detay, termin_detay)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
    terminDetay
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
  const { musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, tarih, notlar, durum, termin, renk, renk_detay, kumas_tamam, fit, ld_test, bulk_test, gold_test, termin_detay } = req.body;

  if (!musteri_adi || !urun_aciklamasi || fiyat == null) {
    return res.status(400).json({ hata: t(getLang(req), 'fiyatRequired') });
  }

  const eski = db.prepare('SELECT gorsel, excel_dosya FROM siparisler WHERE id = ?').get(req.params.id);
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
    SET musteri_adi=?, firma_adi=?, urun_aciklamasi=?, fiyat=?, para_birimi=?, miktar=?, tarih=?, notlar=?, durum=?, gorsel=?, termin=?, renk=?, renk_detay=?, excel_dosya=?, kumas_tamam=?, fit=?, ld_test=?, bulk_test=?, gold_test=?, termin_detay=?
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
    req.params.id
  );

  res.json({ mesaj: t(getLang(req), 'siparisUpdated') });
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
  db.prepare('UPDATE siparisler SET durum=? WHERE id=?').run(durum, req.params.id);
  res.json({ mesaj: t(getLang(req), 'siparisStatusUpdated') });
});

// siparişi geri al (Tamamlandı → Üretimde)
app.post('/api/siparisler/:id/geri-al', (req, res) => {
  const siparis = db.prepare('SELECT * FROM siparisler WHERE id = ?').get(req.params.id);
  if (!siparis) return res.status(404).json({ hata: t(getLang(req), 'siparisNotFound') });

  db.prepare('UPDATE siparisler SET durum = ? WHERE id = ?').run('Üretimde', req.params.id);

  // varsa bağlı gerçekleşen kaydını sil
  const gerceklesen = db.prepare('SELECT id FROM gerceklesen WHERE siparis_id = ?').get(req.params.id);
  if (gerceklesen) {
    db.prepare('DELETE FROM gerceklesen WHERE siparis_id = ?').run(req.params.id);
  }

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
  const rows = db.prepare('SELECT * FROM aksesuar_takip WHERE siparis_id = ? ORDER BY created_at ASC').all(req.params.id);
  res.json(rows);
});

// yeni aksesuar ekle
app.post('/api/siparisler/:id/aksesuar', (req, res) => {
  const { aksesuar_adi, siparis_adet, gelen_adet, notlar, fatura } = req.body;
  if (!aksesuar_adi) return res.status(400).json({ hata: t(getLang(req), 'aksesuarRequired') });

  const info = db.prepare(`
    INSERT INTO aksesuar_takip (siparis_id, aksesuar_adi, siparis_adet, gelen_adet, notlar, fatura, gelen_detay)
    VALUES (?, ?, ?, ?, ?, ?, ?)
  `).run(req.params.id, aksesuar_adi, siparis_adet || 0, gelen_adet || 0, notlar || '', fatura || '', '[]');

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
  const { aksesuar_adi, siparis_adet, notlar, gelen_detay } = req.body;
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
    UPDATE aksesuar_takip SET aksesuar_adi=?, siparis_adet=?, gelen_adet=?, notlar=?, gelen_detay=? WHERE id=?
  `).run(aksesuar_adi, siparis_adet || 0, gelenAdet, notlar || '', detayJson, req.params.id);

  res.json({ mesaj: t(getLang(req), 'aksesuarUpdated'), gelen_adet: gelenAdet });
});

// aksesuar sil
app.delete('/api/aksesuar/:id', (req, res) => {
  const eski = db.prepare('SELECT * FROM aksesuar_takip WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'aksesuarNotFound') });
  db.prepare('DELETE FROM aksesuar_takip WHERE id = ?').run(req.params.id);
  res.json({ mesaj: t(getLang(req), 'aksesuarDeleted') });
});

// --- KUMAŞ TAKİP API ---

// siparişe ait kumaşları listele
app.get('/api/siparisler/:id/kumas', (req, res) => {
  const rows = db.prepare('SELECT * FROM kumas_takip WHERE siparis_id = ? ORDER BY created_at ASC').all(req.params.id);
  res.json(rows);
});

// yeni kumaş ekle
app.post('/api/siparisler/:id/kumas', (req, res) => {
  const { kumas_adi, renk, adet, pastal_boyu, kumas_en, kumas_gramaj, fire, asorti, biye, kumas_fiyati, kesim_adedi, kesim_yuzdesi, notlar } = req.body;
  const info = db.prepare(`
    INSERT INTO kumas_takip (siparis_id, kumas_adi, renk, adet, pastal_boyu, kumas_en, kumas_gramaj, fire, asorti, biye, birim_gider, kumas_fiyati, kesim_adedi, kesim_yuzdesi, gerceklesen_birim, notlar)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(req.params.id, kumas_adi || '', renk || '', adet || 0, pastal_boyu || 0, kumas_en || 0, kumas_gramaj || 0, fire || 0, asorti || 1, biye || 0, 0, kumas_fiyati || 0, kesim_adedi || 0, kesim_yuzdesi || 0, 0, notlar || '');
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
  const { kumas_adi, renk, adet, pastal_boyu, kumas_en, kumas_gramaj, fire, asorti, biye, kumas_fiyati, kesim_adedi, kesim_yuzdesi, notlar, gelen_detay } = req.body;
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
  const gerceklesenBirim = ka > 0 ? Math.round((toplamGelenKumas / ka) * 1000) / 1000 : 0;

  db.prepare(`
    UPDATE kumas_takip SET kumas_adi=?, renk=?, adet=?, pastal_boyu=?, kumas_en=?, kumas_gramaj=?, fire=?, asorti=?, biye=?, birim_gider=?, kumas_fiyati=?, kesim_adedi=?, kesim_yuzdesi=?, gerceklesen_birim=?, notlar=?, gelen_detay=?
    WHERE id=?
  `).run(kumas_adi || '', renk || '', ad, pastal_boyu || 0, kumas_en || 0, kumas_gramaj || 0, fire || 0, asorti || 1, biye || 0, birimGider, kumas_fiyati || 0, ka, hesaplananKesimYuzdesi, gerceklesenBirim, notlar || '', detayJson, req.params.id);

  // gerçekleşen varsa kesim adedini güncelle
  const siparisId = eski.siparis_id;
  const gerceklesen = db.prepare('SELECT id FROM gerceklesen WHERE siparis_id = ?').get(siparisId);
  if (gerceklesen) {
    const toplamKesim = db.prepare('SELECT COALESCE(SUM(kesim_adedi), 0) as toplam FROM kumas_takip WHERE siparis_id = ?').get(siparisId);
    db.prepare('UPDATE gerceklesen SET kesim_adedi = ? WHERE id = ?').run(toplamKesim.toplam, gerceklesen.id);
  }

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
  res.json({ mesaj: t(getLang(req), 'kumasDeleted') });
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
          yuklenen_adet_detay, satis_toplam_detay, kumas_bedeli_detay, aksesuar_bedeli_detay, iscilik_bedeli_detay, baski_nakis_yikama, baski_nakis_yikama_detay } = req.body;

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

  db.prepare(`
    UPDATE gerceklesen
    SET musteri_adi=?, firma_adi=?, urun_aciklamasi=?, fiyat=?, para_birimi=?, miktar=?,
        kesim_adedi=?, yuklenen_adet=?, ikinci_kalite=?, kumas_bedeli=?, aksesuar_bedeli=?, iscilik_bedeli=?, satis_toplam=?, termin=?, excel_dosya=?,
        yuklenen_adet_detay=?, satis_toplam_detay=?, kumas_bedeli_detay=?, aksesuar_bedeli_detay=?, iscilik_bedeli_detay=?, baski_nakis_yikama=?, baski_nakis_yikama_detay=?
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
    req.params.id
  );

  res.json({ mesaj: t(getLang(req), 'gerceklesenUpdated') });
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
  res.json({ mesaj: t(getLang(req), 'fiyatDeleted') });
});

// --- ÖDEME PLANI API ---

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

// haftalık gruplandırılmış ödeme planı
app.get('/api/odemeler/haftalik', (req, res) => {
  const bugun = new Date();
  const bugunStr = bugun.toISOString().split('T')[0];

  // bekleyen tüm ödemeleri getir
  const rows = db.prepare(
    "SELECT * FROM odemeler WHERE durum = 'Bekliyor' ORDER BY vade_tarihi ASC"
  ).all();

  // en uzak vadeye göre hafta sayısını hesapla
  const enUzakVade = rows.reduce((max, r) => {
    if (r.vade_tarihi > max) return r.vade_tarihi;
    return max;
  }, bugunStr);
  const sonTarih = new Date(enUzakVade);
  const haftaFarki = Math.ceil((sonTarih - bugun) / (7 * 24 * 60 * 60 * 1000));
  const haftaSayisi = Math.max(haftaFarki + 2, 8); // en az 8 hafta, en uzak vadeye kadar

  const haftalar = [];
  for (let i = 0; i < haftaSayisi; i++) {
    const baslangic = new Date(bugun);
    baslangic.setDate(bugun.getDate() - bugun.getDay() + 1 + (i * 7)); // pazartesi
    const bitis = new Date(baslangic);
    bitis.setDate(baslangic.getDate() + 6); // pazar

    const bStr = baslangic.toISOString().split('T')[0];
    const btStr = bitis.toISOString().split('T')[0];

    haftalar.push({ baslangic: bStr, bitis: btStr, odemeler: [], toplam_tutar: 0 });
  }

  const gecikmis = [];

  rows.forEach(row => {
    if (row.vade_tarihi < bugunStr) {
      gecikmis.push(row);
      return;
    }

    for (const h of haftalar) {
      if (row.vade_tarihi >= h.baslangic && row.vade_tarihi <= h.bitis) {
        h.odemeler.push(row);
        h.toplam_tutar += row.tutar;
        break;
      }
    }
  });

  // gelecek toplam (bugünden sonraki tüm bekleyenler)
  const gelecek_toplam = rows
    .filter(r => r.vade_tarihi >= bugunStr)
    .reduce((sum, r) => sum + r.tutar, 0);

  res.json({
    haftalar,
    gecikmis,
    gecikmis_toplam: gecikmis.reduce((sum, r) => sum + r.tutar, 0),
    gelecek_toplam,
    bugun: bugunStr
  });
});

// tek ödeme getir
app.get('/api/odemeler/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM odemeler WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ hata: t(getLang(req), 'fiyatNotFound') });
  res.json(row);
});

// yeni ödeme ekle
app.post('/api/odemeler', (req, res) => {
  const { musteri_adi, firma_adi, odeme_turu, tutar, para_birimi, vade_tarihi, aciklama, cek_no, banka, cek_sahibi, kredi_grup_id } = req.body;

  if (!tutar || !vade_tarihi) {
    return res.status(400).json({ hata: t(getLang(req), 'odemeRequired') });
  }

  const stmt = db.prepare(`
    INSERT INTO odemeler (musteri_adi, firma_adi, odeme_turu, tutar, para_birimi, vade_tarihi, aciklama, cek_no, banka, cek_sahibi, kredi_grup_id)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const info = stmt.run(
    '',
    firma_adi || '',
    odeme_turu || 'Rotatif Kredi',
    tutar,
    para_birimi || 'TL',
    vade_tarihi,
    aciklama || '',
    cek_no || '',
    banka || '',
    cek_sahibi || '',
    kredi_grup_id || ''
  );

  res.status(201).json({ id: info.lastInsertRowid, mesaj: t(getLang(req), 'odemeSaved') });
});

// ödeme güncelle
app.put('/api/odemeler/:id', (req, res) => {
  const { musteri_adi, firma_adi, odeme_turu, tutar, para_birimi, vade_tarihi, aciklama, durum, cek_no, banka, cek_sahibi, kredi_grup_id } = req.body;

  if (!tutar || !vade_tarihi) {
    return res.status(400).json({ hata: t(getLang(req), 'odemeRequired') });
  }

  const eski = db.prepare('SELECT * FROM odemeler WHERE id = ?').get(req.params.id);
  if (!eski) return res.status(404).json({ hata: t(getLang(req), 'fiyatNotFound') });

  db.prepare(`
    UPDATE odemeler
    SET musteri_adi='', firma_adi=?, odeme_turu=?, tutar=?, para_birimi=?, vade_tarihi=?, aciklama=?, durum=?, cek_no=?, banka=?, cek_sahibi=?, kredi_grup_id=?
    WHERE id=?
  `).run(
    firma_adi || '',
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
    req.params.id
  );

  res.json({ mesaj: t(getLang(req), 'odemeUpdated') });
});

// ödeme sil
app.delete('/api/odemeler/:id', (req, res) => {
  db.prepare('DELETE FROM odemeler WHERE id = ?').run(req.params.id);
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

// --- yedekleme ---
app.get('/api/yedekle', (req, res) => {
  const tarih = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
  res.setHeader('Content-Type', 'application/gzip');
  res.setHeader('Content-Disposition', `attachment; filename="fiyat-takip-yedek-${tarih}.tar.gz"`);
  const tar = spawn('tar', ['-czf', '-', '-C', '/app/data', '.']);
  tar.stdout.pipe(res);
  tar.stderr.on('data', () => {});
  tar.on('error', () => res.status(500).json({ hata: t(getLang(req), 'backupFailed') }));
});

// --- başlat ---
app.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ Fiyat Takip Sistemi çalışıyor → http://localhost:${PORT}`);
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
