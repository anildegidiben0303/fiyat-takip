const initSqlJs = require('sql.js');
const fs = require('fs');

(async () => {
  const SQL = await initSqlJs();
  const buf = fs.readFileSync('fiyatlar.db');
  const db = new SQL.Database(buf);

  const rows = db.exec("SELECT * FROM fiyatlar");
  if (rows.length) {
    const data = rows[0];
    console.log('Fiyatlar:', data.values.length, 'kayıt bulundu');
    data.values.forEach(r => console.log('  ID:', r[0], '| Müşteri:', r[1], '| Model:', r[4], '| Fiyat:', r[5]));
  } else {
    console.log('Fiyatlar: boş');
  }

  const sip = db.exec("SELECT * FROM siparisler");
  console.log('Siparişler:', sip.length ? sip[0].values.length : 0, 'kayıt');

  const ger = db.exec("SELECT * FROM gerceklesen");
  console.log('Gerçekleşen:', ger.length ? ger[0].values.length : 0, 'kayıt');
})();
