const initSqlJs = require('sql.js');
const fs = require('fs');

let SQL = null;
let db = null;
let dbPath = '';

// better-sqlite3 uyumlu wrapper
class Database {
  constructor(path) {
    dbPath = path;
    if (fs.existsSync(path)) {
      const buffer = fs.readFileSync(path);
      db = new SQL.Database(buffer);
    } else {
      db = new SQL.Database();
    }
  }

  pragma(sql) { db.run(sql); return this; }

  exec(sql) { db.run(sql); this._save(); return this; }

  prepare(sql) {
    return new Statement(db, sql, dbPath);
  }

  _save() {
    if (dbPath) {
      const data = db.export();
      fs.writeFileSync(dbPath, Buffer.from(data));
    }
  }
}

class Statement {
  constructor(db, sql, dbPath) {
    this.db = db;
    this.sql = sql;
    this.dbPath = dbPath;
  }

  all(...params) {
    const stmt = this.db.prepare(this.sql);
    if (params.length) stmt.bind(params);
    const rows = [];
    while (stmt.step()) {
      rows.push(stmt.getAsObject());
    }
    stmt.free();
    return rows;
  }

  get(...params) {
    const stmt = this.db.prepare(this.sql);
    if (params.length) stmt.bind(params);
    let row = null;
    if (stmt.step()) {
      row = stmt.getAsObject();
    }
    stmt.free();
    return row || undefined;
  }

  run(...params) {
    const stmt = this.db.prepare(this.sql);
    if (params.length) {
      const flatParams = params.length === 1 && typeof params[0] === 'object' ? params[0] : params;
      if (Array.isArray(flatParams)) {
        stmt.bind(flatParams);
      } else {
        stmt.bind(flatParams);
      }
    }
    try {
      stmt.step();
      stmt.free();
      const lastId = this.db.exec('SELECT last_insert_rowid() as id');
      const id = lastId.length ? lastId[0].values[0][0] : 0;
      const changes = this.db.getRowsModified();
      // kaydet
      if (this.dbPath) {
        const data = this.db.export();
        fs.writeFileSync(this.dbPath, Buffer.from(data));
      }
      return { lastInsertRowid: id, changes };
    } catch (e) {
      stmt.free();
      throw e;
    }
  }
}

async function createDatabase(path) {
  SQL = await initSqlJs();
  return new Database(path);
}

module.exports = { createDatabase };
