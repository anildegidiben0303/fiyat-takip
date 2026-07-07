// Mevcut Banka Kredisi kayıtlarını gruplandır
const Database = require('better-sqlite3');
const path = require('path');

const dbPath = process.env.DB_PATH || path.join(__dirname, 'fiyatlar.db');
const db = new Database(dbPath);

// kredi_grup_id'si boş olan Banka Kredisi kayıtlarını bul
const rows = db.prepare(`
  SELECT * FROM odemeler
  WHERE odeme_turu = 'Banka Kredisi' AND (kredi_grup_id IS NULL OR kredi_grup_id = '')
  ORDER BY created_at ASC
`).all();

if (rows.length === 0) {
  console.log('Güncellenecek kayıt yok.');
  db.close();
  process.exit(0);
}

// Aynı anda oluşturulanları (±2 sn) aynı gruba ata
let grupId = '';
let sonTarih = '';
let gruplanan = 0;

for (const row of rows) {
  const created = row.created_at || '';

  // Yeni grup başlat (aynı saniyede oluşturulanlar aynı gruba)
  if (!grupId || Math.abs(new Date(created) - new Date(sonTarih)) > 3000) {
    grupId = 'kredi_' + Date.now() + '_' + Math.random().toString(36).substr(2, 4);
  }

  db.prepare('UPDATE odemeler SET kredi_grup_id = ? WHERE id = ?').run(grupId, row.id);
  console.log('✅ ID ' + row.id + ' → grup: ' + grupId);
  sonTarih = created;
  gruplanan++;
}

db.close();
console.log(gruplanan + ' kayıt gruplandı.');
