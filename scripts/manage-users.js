#!/usr/bin/env node
'use strict';
/**
 * CLI kelola user. Jalankan setelah aplikasi pernah start (tabel sudah ada).
 *   node scripts/manage-users.js list
 *   node scripts/manage-users.js add <username> "<Nama Lengkap>" "<Mekanik|Group Leader|Super admin>" [password] [jobsite]
 *   node scripts/manage-users.js passwd <username> [password]
 *   node scripts/manage-users.js role <username> "<role>"
 *   node scripts/manage-users.js deactivate|activate <username>
 *   node scripts/manage-users.js delete <username>
 * Bila password tidak diberikan, password acak dibuat dan ditampilkan sekali.
 */
const crypto = require('crypto');
const bcrypt = require('bcrypt');
const pool = require('./_db');
const ROLES = ['Mekanik', 'Group Leader', 'Super admin'];
const [cmd, ...a] = process.argv.slice(2);

const randPw = () => crypto.randomBytes(9).toString('base64url');

async function main() {
  switch (cmd) {
    case 'list': {
      const { rows } = await pool.query('SELECT username, nama_lengkap, role, jobsite, active FROM users ORDER BY role, jobsite, nama_lengkap');
      console.table(rows);
      break;
    }
    case 'add': {
      const [username, nama, role, pw, jobsiteArg] = a;
      if (!username || !nama || !ROLES.includes(role)) throw new Error(`Pemakaian: add <username> "<nama>" "<${ROLES.join('|')}>" [password] [jobsite]`);
      const jobsite = (jobsiteArg || '').toUpperCase() || null;
      if (role !== 'Super admin' && !jobsite) throw new Error('Jobsite wajib untuk Mekanik/Group Leader (argumen ke-5, mis. BBM atau DBK)');
      const password = pw || randPw();
      if (password.length < 6 || password.length > 72) throw new Error('Password 6–72 karakter');
      await pool.query('INSERT INTO users(username,password_hash,nama_lengkap,role,jobsite) VALUES($1,$2,$3,$4,$5)',
        [username, await bcrypt.hash(password, 10), nama, role, jobsite]);
      console.log(`User ${username} dibuat.` + (pw ? '' : ` Password: ${password}`));
      break;
    }
    case 'passwd': {
      const [username, pw] = a;
      if (!username) throw new Error('Pemakaian: passwd <username> [password]');
      const password = pw || randPw();
      if (password.length < 6 || password.length > 72) throw new Error('Password 6–72 karakter');
      const r = await pool.query('UPDATE users SET password_hash=$1, token_version=token_version+1 WHERE username=$2',
        [await bcrypt.hash(password, 10), username]);
      if (!r.rowCount) throw new Error('User tidak ditemukan');
      console.log('Password diubah.' + (pw ? '' : ` Password baru: ${password}`));
      break;
    }
    case 'role': {
      const [username, role] = a;
      if (!username || !ROLES.includes(role)) throw new Error(`Role harus salah satu: ${ROLES.join(', ')}`);
      const r = await pool.query('UPDATE users SET role=$1, token_version=token_version+1 WHERE username=$2', [role, username]);
      if (!r.rowCount) throw new Error('User tidak ditemukan');
      console.log('Role diubah.');
      break;
    }
    case 'activate':
    case 'deactivate': {
      const r = await pool.query('UPDATE users SET active=$1, token_version=token_version+1 WHERE username=$2', [cmd === 'activate', a[0]]);
      if (!r.rowCount) throw new Error('User tidak ditemukan');
      console.log(cmd === 'activate' ? 'User diaktifkan.' : 'User dinonaktifkan.');
      break;
    }
    case 'delete': {
      try {
        const r = await pool.query('DELETE FROM users WHERE username=$1', [a[0]]);
        if (!r.rowCount) throw new Error('User tidak ditemukan');
        console.log('User dihapus.');
      } catch (e) {
        if (e.code === '23503') throw new Error('User punya data laporan; gunakan deactivate.');
        throw e;
      }
      break;
    }
    default:
      console.log('Perintah: list | add | passwd | role | activate | deactivate | delete (lihat komentar di atas file)');
  }
}
main().catch((e) => { console.error('Error:', e.message); process.exitCode = 1; }).finally(() => pool.end());
