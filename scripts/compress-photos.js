#!/usr/bin/env node
'use strict';
/**
 * Kompres foto laporan LAMA yang sudah ada di server (foto baru otomatis dikompres saat upload).
 *   node scripts/compress-photos.js --dry-run     (hanya hitung perkiraan penghematan)
 *   node scripts/compress-photos.js               (jalankan sungguhan)
 *   node scripts/compress-photos.js --min=200     (lewati foto JPEG di bawah 200 KB; bawaan 300)
 * Hasil: sisi terpanjang maks 1600 px, JPEG kualitas 80. Alamat foto (/api/photos/:id) tidak berubah.
 * Jalankan di Terminal container aplikasi (Dokploy), setelah backup volume foto bila perlu.
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const pool = require('./_db');
let sharp;
try { sharp = require('sharp'); } catch { console.error('Modul sharp tidak tersedia.'); process.exit(1); }

const UPLOAD_DIR = process.env.UPLOAD_DIR || path.join(__dirname, '..', 'uploads');
const dry = process.argv.includes('--dry-run');
const minArg = process.argv.find((a) => a.startsWith('--min='));
const MIN = (minArg ? Number(minArg.split('=')[1]) : 300) * 1024;
const rm = (p) => [p, `${p}.thumb.jpg`, `${p}.sm.jpg`].forEach((f) => { try { fs.unlinkSync(f); } catch { /* abaikan */ } });
const mb = (n) => (n / 1048576).toFixed(2) + ' MB';

async function main() {
  const { rows } = await pool.query('SELECT id, filename, mime, size_bytes FROM report_photos ORDER BY id');
  let before = 0, after = 0, done = 0, skipped = 0, failed = 0;
  for (const r of rows) {
    const file = path.join(UPLOAD_DIR, r.filename);
    let st;
    try { st = fs.statSync(file); } catch { failed++; console.warn(`#${r.id} file tidak ditemukan: ${r.filename}`); continue; }
    if (st.size < MIN && r.mime === 'image/jpeg') { skipped++; continue; }
    try {
      const out = await sharp(fs.readFileSync(file), { failOn: 'none' }).rotate()
        .resize({ width: 1600, height: 1600, fit: 'inside', withoutEnlargement: true }).jpeg({ quality: 80, mozjpeg: true }).toBuffer();
      if (r.mime === 'image/jpeg' && out.length >= st.size) { skipped++; continue; }
      before += st.size; after += out.length; done++;
      if (!dry) {
        const name = `${crypto.randomUUID()}.jpg`;
        fs.writeFileSync(path.join(UPLOAD_DIR, name), out, { flag: 'wx' });
        await pool.query('UPDATE report_photos SET filename=$1, mime=$2, size_bytes=$3 WHERE id=$4', [name, 'image/jpeg', out.length, r.id]);
        rm(file);
      }
      console.log(`#${r.id}: ${mb(st.size)} -> ${mb(out.length)}`);
    } catch (e) { failed++; console.warn(`#${r.id} gagal: ${e.message}`); }
  }
  console.log(`\n${dry ? '[DRY-RUN] ' : ''}${done} foto ${dry ? 'akan dikompres' : 'dikompres'}, ${skipped} dilewati, ${failed} gagal.`);
  console.log(`Ukuran: ${mb(before)} -> ${mb(after)} (hemat ${mb(before - after)}).`);
}
main().catch((e) => { console.error('Error:', e.message); process.exitCode = 1; }).finally(() => pool.end());
