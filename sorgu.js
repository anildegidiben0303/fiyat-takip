const Database = require('better-sqlite3');
const db = new Database('/app/data/fiyatlar.db', { readonly: true });

console.log('=== GERCEKLESEN ===');
const g = db.prepare('SELECT * FROM gerceklesen ORDER BY created_at DESC').all();
console.log('Kayit sayisi:', g.length);
g.forEach(x => console.log('ID:', x.id, '| Müsteri:', x.musteri_adi, '| Ürün:', x.urun_aciklamasi, '| Fiyat:', x.fiyat, '| Tarih:', x.created_at));

console.log('');
console.log('=== SIPARISLER ===');
const s = db.prepare('SELECT * FROM siparisler ORDER BY created_at DESC').all();
console.log('Kayit sayisi:', s.length);

console.log('');
console.log('=== ODEMELER ===');
const o = db.prepare('SELECT count(*) as cnt FROM odemeler').get();
console.log('Kayit sayisi:', o.cnt);

db.close();
