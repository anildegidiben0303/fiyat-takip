// Kayıp siparişleri gerçekleşen kayıtlarından geri oluştur
const Database = require('better-sqlite3');
const path = require('path');

const dbPath = process.env.DB_PATH || path.join(__dirname, 'fiyatlar.db');
const db = new Database(dbPath);

// Gerçekleşendeki kayıtları tara, siparişte olmayanları bul
const gerceklesenRows = db.prepare('SELECT * FROM gerceklesen').all();
const siparisIds = new Set(db.prepare('SELECT id FROM siparisler').all().map(r => r.id));

let geriYuklenen = 0;

for (const g of gerceklesenRows) {
  if (!g.siparis_id) continue;
  if (siparisIds.has(g.siparis_id)) continue; // zaten var

  // Siparişi geri oluştur
  const info = db.prepare(`
    INSERT INTO siparisler (id, musteri_adi, firma_adi, urun_aciklamasi, fiyat, para_birimi, miktar, gorsel, durum, created_at, siparis_tarihi)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Tamamlandı', ?, ?)
  `).run(
    g.siparis_id,
    g.musteri_adi,
    g.firma_adi || '',
    g.urun_aciklamasi,
    g.fiyat,
    g.para_birimi || 'TL',
    g.miktar || 1,
    g.gorsel || '',
    g.created_at,
    g.created_at
  );

  console.log('✅ Geri yüklendi: siparis ID ' + g.siparis_id + ' | ' + g.musteri_adi + ' | ' + g.urun_aciklamasi);
  geriYuklenen++;
}

if (geriYuklenen === 0) {
  console.log('ℹ️ Geri yüklenecek kayıp sipariş bulunamadı.');
}

db.close();
console.log('Tamamlandı. ' + geriYuklenen + ' sipariş geri yüklendi.');
