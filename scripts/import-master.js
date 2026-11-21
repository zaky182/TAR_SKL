#!/usr/bin/env node
'use strict';
/**
 * Impor data master dari CSV (header wajib ada, UTF-8). Aman dijalankan berulang: tidak menghapus data.
 *   node scripts/import-master.js units      seed/units.csv       (jobsite,code_unit,merk_unit,model_unit,machine_sn)
 *   node scripts/import-master.js components seed/components.csv  (nama)
 *   node scripts/import-master.js users      seed/users.csv       (nama_lengkap,username,password,role,email,wa_number,jobsite)
 * units : upsert per (jobsite, code_unit) — code unit yang sama boleh ada di jobsite berbeda.
 * users : user yang sudah ada TIDAK diubah (password/WA/email aman), hanya jobsite diisi bila masih kosong;
 *         user baru wajib berpassword.
 */
const fs = require('fs');
const bcrypt = require('bcrypt');
const { parse } = require('csv-parse/sync');
const pool = require('./_db');
const ROLES = ['Mekanik', 'Group Leader', 'Super admin'];
const [kind, file] = process.argv.slice(2);

async function main() {
  if (!kind || !file) throw new Error('Pemakaian: import-master.js <units|components|users> <file.csv>');
  const rows = parse(fs.readFileSync(file), { columns: true, bom: true, skip_empty_lines: true, trim: true });
  let n = 0;
  for (const r of rows) {
    if (kind === 'units') {
      if (!r.jobsite || !r.code_unit || !r.merk_unit || !r.model_unit) { console.warn('Baris dilewati (jobsite/code_unit/merk/model wajib):', r); continue; }
      await pool.query(
        `INSERT INTO units(jobsite,code_unit,merk_unit,model_unit,machine_sn) VALUES($1,$2,$3,$4,$5)
         ON CONFLICT (jobsite, code_unit) DO UPDATE SET merk_unit=EXCLUDED.merk_unit, model_unit=EXCLUDED.model_unit, machine_sn=EXCLUDED.machine_sn`,
        [r.jobsite.toUpperCase(), r.code_unit, r.merk_unit, r.model_unit, r.machine_sn || null]);
    } else if (kind === 'components') {
      if (!r.nama) continue;
      await pool.query('INSERT INTO components(nama) VALUES($1) ON CONFLICT DO NOTHING', [r.nama]);
    } else if (kind === 'users') {
      const js = (r.jobsite || '').toUpperCase() || null;
      const { rows: [ex] } = await pool.query('SELECT id FROM users WHERE username=$1', [r.username]);
      if (ex) { if (js) await pool.query('UPDATE users SET jobsite=$1 WHERE id=$2 AND jobsite IS NULL', [js, ex.id]); }
      else {
        if (!r.username || !r.password || !ROLES.includes(r.role)) { console.warn('Baris dilewati (username/password/role):', r.username); continue; }
        if (r.role !== 'Super admin' && !js) { console.warn('Baris dilewati (jobsite wajib untuk Mekanik/Group Leader):', r.username); continue; }
        await pool.query('INSERT INTO users(username,password_hash,nama_lengkap,role,email,wa_number,jobsite) VALUES($1,$2,$3,$4,$5,$6,$7)',
          [r.username, await bcrypt.hash(String(r.password), 10), r.nama_lengkap, r.role, r.email || null, r.wa_number || null, js]);
      }
    } else throw new Error('Jenis harus units, components, atau users');
    n++;
  }
  console.log(`${n} baris diproses untuk ${kind}.`);
}
main().catch((e) => { console.error('Error:', e.message); process.exitCode = 1; }).finally(() => pool.end());
