/** APLIKASI NILAI SISWA - SMK NEGERI 2 PATI | Database: Google Spreadsheet (script terikat) */
const KKM = 70;
const JUR = {B:['DPIB','Desain Pemodelan dan Informasi Bangunan'],E:['TAV','Teknik Audio Video'],L:['TOI','Teknik Otomasi Industri'],M:['TP','Teknik Pemesinan'],O:['TKRO','Teknik Kendaraan Ringan Otomotif'],D:['DKV','Desain Komunikasi Visual']};
const MAPEL = {
  X:['A. Pendidikan Agama dan Budi Pekerti','A. Pendidikan Pancasila','A. Bahasa Indonesia','A. Pendidikan Jasmani, Olahraga, dan Kesehatan','A. Sejarah','A. Seni','B. Matematika','B. Bahasa Inggris','B. Informatika','B. Proyek Ilmu Pengetahuan Alam dan Sosial','B. Dasar-Dasar Program Keahlian','Mulok Bahasa Jawa'],
  XI:['A. Pendidikan Agama dan Budi Pekerti','A. Pendidikan Pancasila','A. Bahasa Indonesia','A. Pendidikan Jasmani, Olahraga, dan Kesehatan','A. Sejarah','B. Matematika','B. Bahasa Inggris','Mata Pelajaran Konsentrasi Keahlian','Projek Kreatif dan Kewirausahaan','Mata Pelajaran Pilihan','Mulok Bahasa Jawa'],
  XII:['A. Pendidikan Agama dan Budi Pekerti','A. Pendidikan Pancasila','A. Bahasa Indonesia','B. Matematika','B. Bahasa Inggris','Mata Pelajaran Konsentrasi Keahlian','Projek Kreatif dan Kewirausahaan','Mata Pelajaran Pilihan','Mulok Bahasa Jawa']
};
const SS = () => SpreadsheetApp.getActiveSpreadsheet();
let _M = null; // daftar mata pelajaran dibaca dari sheet 'Mapel' (bisa diubah admin)
function mapelList() { if (!_M) { const s = SS().getSheetByName('Mapel'); _M = s && s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 3).getValues() : []; } return _M; }
const sn = t => mapelList().filter(r => r[0] == t).map(r => r[2]);
const ALL = () => mapelList().map(r => r[2]);
const mapelAll = () => ({ X: sn('X'), XI: sn('XI'), XII: sn('XII') });
const chk = s => { if (s && ALL().indexOf(s) < 0) throw new Error('Mata pelajaran tidak valid'); };
const A = t => { if (!CacheService.getScriptCache().get('t_' + t)) throw new Error('Sesi admin habis, silakan login ulang'); };
function data(n) { const s = SS().getSheetByName(n); return s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).getValues() : []; }

function doGet(e) {
  if (e && e.parameter && e.parameter.ping) return ContentService.createTextOutput('OK'); // uji koneksi API
  if (!SS().getSheetByName('Siswa')) setup();
  else if (sn('X')[0] && SS().getSheetByName(sn('X')[0]).getRange(1, 3).getValue() === 'Kelas') sinkron(); // migrasi otomatis
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Nilai Siswa SMKN 2 Pati')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Jalankan sekali: membuat seluruh sheet (Jurusan, Kelas, WaliKelas, Siswa, Mapel + 32 sheet mapel) */
function setup() {
  const ss = SS();
  const mk = (n, h) => {
    let s = ss.getSheetByName(n);
    if (!s) { s = ss.insertSheet(n); s.appendRow(h); s.setFrozenRows(1);
      s.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#fff');
      if (h[0] === 'NIS') s.getRange('A:A').setNumberFormat('@'); }
    return s;
  };
  const j = mk('Jurusan', ['Kode', 'Singkatan', 'Kepanjangan']);
  if (j.getLastRow() < 2) Object.keys(JUR).forEach(k => j.appendRow([k, JUR[k][0], JUR[k][1]]));
  const k = mk('Kelas', ['Kelas', 'Tingkat', 'KodeJurusan', 'Jurusan', 'Rombel']);
  const w = mk('WaliKelas', ['Kelas', 'Nama', 'NIP']);
  if (k.getLastRow() < 2) {
    const r = [], wr = [];
    ['X', 'XI', 'XII'].forEach(t => Object.keys(JUR).forEach(c => [1, 2, 3].forEach(n => { const nm = t + ' ' + c + n; r.push([nm, t, c, JUR[c][0], n]); wr.push([nm, '', '']); })));
    k.getRange(2, 1, r.length, 5).setValues(r);
    w.getRange(2, 1, wr.length, 3).setValues(wr);
    // contoh data dari legger X-D3
    const nama = ['AGIL DWI SAPUTRA','AKBAR YAHYA MAULANA PRAYITNO PUTRA','ANANDITA LINTANG SARY','ANGGUN ISNA FARANTI','ARGONTARA HARUN AL MISKI KAMANURU','AVDINDA WURIZKA','DAVIN PRAKAS ELVIANO','DELLA OKTAVIA','DELVIN RADITYA PUTRA','DENITA DYAN PRAMUSTI','DHIAR RAHMAN','DIANDRA SEKARTAJI','ERLANGGA TRESNA VANO','FAKHRIDHO AZRIL MAULANA','GIO CANDRA WIBOWO','HILMATUL FADHILAH','ISA LEFIANA DEWI','KEVIN ILHAM SUSANTO','LUTHFIA SARI RAHMA NINGRUM','MARLINA YULIA ALIANSA','MILADI JOLAN KADAFI','MOHAMMAD ILHAM SAPUTRO','MOZA VALENTINA VEBIOLA','MUHAMMAD FARHAN NUR HUDA','MUHAMMAD IQBAL MUZAKKI','MUHAMMAD RADITYA FADHIL ANANDA','MUHAMMAD REYHAN PRASETYO','MUTIA HUSNUL HOTIMAH','NATHAN AFRILIAN','RADHITYA TEGAR ANDARWAN','RAIHAN FA\'IQ FEBRIYANTO','SYAHRUDIN AMIN SULAIMAN','VIKO ARJUNA ADITIYA','YULIA EKA WIJAYANTI','ZULFA PRISCA MAULIDA'];
    const s = mk('Siswa', ['NIS', 'Nama', 'Kelas']);
    s.getRange(2, 1, nama.length, 3).setValues(nama.map((n, i) => [String(18268 + i), n, 'X D3']));
    const wi = w.getRange(2, 1, 54, 1).getValues().findIndex(x => x[0] === 'X D3');
    w.getRange(wi + 2, 2, 1, 2).setValues([['Irwan Sumarjianto, S.Kom.', '198701272022211007']]);
  }
  mk('Siswa', ['NIS', 'Nama', 'Kelas']);
  const m = mk('Mapel', ['Tingkat', 'Mapel', 'NamaSheet']);
  if (m.getLastRow() < 2) ['X', 'XI', 'XII'].forEach(t => MAPEL[t].forEach(x => m.appendRow([t, x, t + ' ' + x])));
  _M = null;
  ALL().forEach(n => mk(n, ['NIS', 'Nama', 'Gasal', 'Genap']));
  sinkron();
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('ADMIN_PASS')) p.setProperty('ADMIN_PASS', 'nDalem'); // GANTI di Project Settings > Script Properties
}

function init() {
  const wali = {}; data('WaliKelas').forEach(r => wali[r[0]] = [r[1], r[2]]);
  return { kelas: data('Kelas').map(r => ({ k: r[0], t: r[1], c: r[2], j: r[3] })), jur: JUR, mapel: mapelAll(), wali, kkm: KKM };
}
function login(p) {
  if (p !== PropertiesService.getScriptProperties().getProperty('ADMIN_PASS')) throw new Error('Kata sandi salah');
  const t = Utilities.getUuid(); CacheService.getScriptCache().put('t_' + t, '1', 21600); return t;
}

function getKelas(kelas, sheet) {
  const nm = {}; if (sheet) { chk(sheet); data(sheet).forEach(r => nm[String(r[0])] = r); }
  return data('Siswa').filter(r => r[2] == kelas).map(r => { const x = nm[String(r[0])] || []; return { nis: String(r[0]), nama: r[1], kelas: r[2], g: x[2] ?? '', e: x[3] ?? '' }; })
    .sort((a, b) => a.nama < b.nama ? -1 : 1);
}

function statistik(sheet, jur, kelas, sem) {
  chk(sheet); const t = sheet.split(' ')[0], col = sem == 'Genap' ? 3 : 2, sis = {}, nm = {}, ada = {};
  data('Siswa').forEach(r => { const k = r[2]; if (k.split(' ')[0] != t) return; if (kelas && k != kelas) return; if (jur && k.split(' ')[1][0] != jur) return; sis[String(r[0])] = k; nm[String(r[0])] = r[1]; });
  const bins = [0, 0, 0, 0, 0], ks = {}, all = [];
  data(sheet).forEach(r => { const k = sis[String(r[0])], v = r[col]; if (!k || v === '' || v === null) return; ada[String(r[0])] = 1; const n = Number(v); all.push(n); bins[n < 60 ? 0 : n < 70 ? 1 : n < 80 ? 2 : n < 90 ? 3 : 4]++; (ks[k] = ks[k] || []).push(n); });
  const avg = a => a.length ? Math.round(a.reduce((x, y) => x + y, 0) / a.length * 10) / 10 : 0;
  return { siswa: Object.keys(sis).length, n: all.length, avg: avg(all), max: all.length ? Math.max.apply(null, all) : 0, min: all.length ? Math.min.apply(null, all) : 0,
    tuntas: all.length ? Math.round(all.filter(x => x >= KKM).length * 100 / all.length) : 0, bins, belum: Object.keys(sis).filter(x => !ada[x]).map(x => ({ nis: x, nama: nm[x], kelas: sis[x] })).sort((a, b) => (a.kelas + a.nama) < (b.kelas + b.nama) ? -1 : 1), kelas: Object.keys(ks).sort().map(k => ({ k, avg: avg(ks[k]), n: ks[k].length })) };
}

function cari(q) {
  q = String(q).trim().toLowerCase(); if (q.length < 2) return [];
  const m = data('Siswa').filter(r => String(r[0]).toLowerCase() == q || String(r[1]).toLowerCase().indexOf(q) >= 0).slice(0, 8), c = {};
  return m.map(r => {
    const t = r[2].split(' ')[0];
    const nl = sn(t).map(s => { if (!c[s]) { c[s] = {}; data(s).forEach(x => c[s][String(x[0])] = x); } const x = c[s][String(r[0])] || []; return { mapel: s.substring(t.length + 1), g: x[2] ?? '', e: x[3] ?? '' }; });
    return { nis: String(r[0]), nama: r[1], kelas: r[2], nilai: nl };
  });
}

// ===== ADMIN =====
function saveNilai(tok, sheet, rows) {
  A(tok); chk(sheet); const t = sheet.split(' ')[0], idx = {};
  const rd = () => { data(sheet).forEach((r, i) => idx[String(r[0])] = i + 2); };
  rd(); if (rows.some(r => !idx[r.nis])) { sinkron([t]); rd(); }
  const s = SS().getSheetByName(sheet), cl = x => x === '' || x == null ? '' : Number(x);
  rows.forEach(r => { if (idx[r.nis]) s.getRange(idx[r.nis], 3, 1, 2).setValues([[cl(r.g), cl(r.e)]]); });
  return 'Nilai tersimpan';
}
/** Menghubungkan semua sheet mata pelajaran ke sheet Siswa: kolom Nama & Kelas berupa rumus VLOOKUP dari NIS. Nilai tetap aman. */
function sinkron(tl) {
  const sis = data('Siswa');
  (tl || ['X', 'XI', 'XII']).forEach(t => {
    const st = sis.filter(r => String(r[2]).split(' ')[0] == t).sort((a, b) => (a[2] + a[1]) < (b[2] + b[1]) ? -1 : 1);
    sn(t).forEach(n => {
      const s = SS().getSheetByName(n), old = {};
      const lama = s.getRange(1, 3).getValue() === 'Kelas'; // migrasi dari format lama (5 kolom)
      data(n).forEach(r => old[String(r[0])] = lama ? [r[3], r[4]] : [r[2], r[3]]);
      if (lama) s.deleteColumn(3);
      if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 4).clearContent();
      s.getRange(1, 1, 1, 4).setValues([['NIS', 'Nama', 'Gasal', 'Genap']]);
      if (!st.length) return;
      const out = st.map((r, i) => { const o = old[String(r[0])] || ['', '']; return [String(r[0]), '=IFERROR(VLOOKUP($A' + (i + 2) + ',Siswa!$A:$C,2,FALSE),"")', o[0], o[1]]; });
      s.getRange(2, 1, out.length, 4).setValues(out);
    });
  });
}
function onOpen() { SpreadsheetApp.getUi().createMenu('Nilai Siswa').addItem('Sinkronkan NIS/Nama ke semua mapel', 'sinkron').addToUi(); }

function listSiswa(tok, kelas) { A(tok); return getKelas(kelas, ''); }
function saveSiswa(tok, old, nis, nama, kelas) {
  A(tok); nis = String(nis).trim(); nama = String(nama).trim().toUpperCase(); if (!nis || !nama) throw new Error('NIS dan nama wajib diisi');
  const s = SS().getSheetByName('Siswa'), v = data('Siswa').map(r => String(r[0]));
  if (nis != old && v.indexOf(nis) >= 0) throw new Error('NIS sudah terdaftar');
  if (!old) { s.appendRow([nis, nama, kelas]); sinkron([kelas.split(' ')[0]]); return 'Siswa ditambahkan'; }
  const i = v.indexOf(String(old)); if (i < 0) throw new Error('Data tidak ditemukan');
  const t = s.getRange(i + 2, 3).getValue().split(' ')[0];
  s.getRange(i + 2, 1, 1, 3).setValues([[nis, nama, kelas]]);
  sn(t).forEach(n => { const sh = SS().getSheetByName(n), d = data(n).findIndex(r => String(r[0]) == String(old)); if (d >= 0) sh.getRange(d + 2, 1).setValue(nis); });
  sinkron(Array.from(new Set([t, kelas.split(' ')[0]])));
  return 'Siswa diperbarui';
}
function delSiswa(tok, nis) {
  A(tok); const s = SS().getSheetByName('Siswa'), i = data('Siswa').findIndex(r => String(r[0]) == String(nis)); if (i < 0) return 'Tidak ditemukan';
  const t = s.getRange(i + 2, 3).getValue().split(' ')[0]; s.deleteRow(i + 2);
  sn(t).forEach(n => { const d = data(n).findIndex(r => String(r[0]) == String(nis)); if (d >= 0) SS().getSheetByName(n).deleteRow(d + 2); });
  return 'Siswa & nilainya dihapus';
}
function importSiswa(tok, kelas, text) {
  A(tok); const s = SS().getSheetByName('Siswa'), ex = data('Siswa').map(r => String(r[0])), add = [];
  String(text).split('\n').forEach(l => { const p = l.split(/\t|;|,/).map(x => x.trim()).filter(Boolean); if (p.length < 2 || ex.indexOf(p[0]) >= 0) return; ex.push(p[0]); add.push([p[0], p.slice(1).join(' ').toUpperCase(), kelas]); });
  if (add.length) { s.getRange(s.getLastRow() + 1, 1, add.length, 3).setValues(add); sinkron([kelas.split(' ')[0]]); }
  return add.length + ' siswa diimpor';
}
function saveWali(tok, rows) { A(tok); SS().getSheetByName('WaliKelas').getRange(2, 1, rows.length, 3).setValues(rows); return 'Data wali kelas tersimpan'; }

// ===== ADMIN: MATA PELAJARAN (tambah / edit / hapus) =====
function cekNama(t, nama, skip) {
  nama = String(nama).trim(); if (!nama) throw new Error('Nama mata pelajaran wajib diisi');
  if (/[\[\]*?:\/\\]/.test(nama)) throw new Error('Nama tidak boleh mengandung [ ] * ? : / \\');
  const sh = t + ' ' + nama; if (sh.length > 100) throw new Error('Nama terlalu panjang');
  if (ALL().indexOf(sh) >= 0 && sh !== skip) throw new Error('Mata pelajaran sudah ada'); return nama;
}
function addMapel(tok, t, nama) {
  A(tok); if (['X', 'XI', 'XII'].indexOf(t) < 0) throw new Error('Kelas tidak valid');
  nama = cekNama(t, nama); const sh = t + ' ' + nama;
  SS().getSheetByName('Mapel').appendRow([t, nama, sh]); _M = null;
  if (!SS().getSheetByName(sh)) { const s = SS().insertSheet(sh); s.getRange('A:A').setNumberFormat('@'); s.appendRow(['NIS', 'Nama', 'Gasal', 'Genap']); s.setFrozenRows(1); s.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#fff'); }
  sinkron([t]); return { msg: 'Mata pelajaran ditambahkan', mapel: mapelAll() };
}
function editMapel(tok, old, nama) {
  A(tok); chk(old); const t = old.split(' ')[0]; nama = cekNama(t, nama, old); const sh = t + ' ' + nama;
  const i = mapelList().findIndex(r => r[2] === old); if (i < 0) throw new Error('Data tidak ditemukan');
  if (sh !== old) SS().getSheetByName(old).setName(sh);
  SS().getSheetByName('Mapel').getRange(i + 2, 2, 1, 2).setValues([[nama, sh]]); _M = null;
  return { msg: 'Mata pelajaran diperbarui', mapel: mapelAll() };
}
function delMapel(tok, sh) {
  A(tok); chk(sh); const i = mapelList().findIndex(r => r[2] === sh); if (i < 0) throw new Error('Data tidak ditemukan');
  SS().getSheetByName('Mapel').deleteRow(i + 2); const s = SS().getSheetByName(sh); if (s) SS().deleteSheet(s); _M = null;
  return { msg: 'Mata pelajaran & nilainya dihapus', mapel: mapelAll() };
}

// ===== SUSULAN: rekap siswa belum mengerjakan (baris: Kelas > Nama, kolom: tiap mata pelajaran) =====
function susulan(t, jur, rom, sem) {
  const col = sem == 'Genap' ? 3 : 2, sh = sn(t);
  const sis = data('Siswa').filter(r => { const k = String(r[2]).split(' '); return k[0] == t && (!jur || k[1][0] == jur) && (!rom || k[1].slice(1) == rom); })
    .sort((a, b) => (a[2] + a[1]) < (b[2] + b[1]) ? -1 : 1);
  const m = sh.map(n => { const o = {}; data(n).forEach(r => o[String(r[0])] = r[col]); return o; });
  return { sheets: sh, mapel: sh.map(n => n.substring(t.length + 1)), rows: sis.map(r => ({ nis: String(r[0]), nama: r[1], kelas: r[2], v: m.map(o => o[String(r[0])] ?? '') })) };
}
function saveSusulan(tok, sem, ch) {
  A(tok); const col = sem == 'Genap' ? 4 : 3, by = {};
  ch.forEach(c => { if (!c.sheet) throw new Error('Mata pelajaran tidak valid'); chk(c.sheet); (by[c.sheet] = by[c.sheet] || []).push(c); });
  Object.keys(by).forEach(n => { const s = SS().getSheetByName(n), idx = {}; data(n).forEach((r, i) => idx[String(r[0])] = i + 2);
    by[n].forEach(c => { if (idx[c.nis]) s.getRange(idx[c.nis], col).setValue(c.v === '' ? '' : Number(c.v)); }); });
  return ch.length + ' nilai tersimpan';
}
// ===== PERINGKAT: lingkup kelas / jurusan / angkatan; mapel tertentu atau seluruh mapel (rata-rata) =====
function peringkat(t, ling, jur, kelas, mapel, sem) {
  chk(mapel); const col = sem == 'Genap' ? 3 : 2;
  const sis = data('Siswa').filter(r => { const k = String(r[2]); if (k.split(' ')[0] != t) return false; return ling == 'kelas' ? k == kelas : ling == 'jurusan' ? k.split(' ')[1][0] == jur : true; });
  const sh = mapel ? [mapel] : sn(t), sum = {}, cnt = {};
  sh.forEach(n => data(n).forEach(r => { const v = r[col]; if (v === '' || v === null) return; const k = String(r[0]); sum[k] = (sum[k] || 0) + Number(v); cnt[k] = (cnt[k] || 0) + 1; }));
  const rows = sis.filter(r => cnt[String(r[0])]).map(r => { const k = String(r[0]); return { nis: k, nama: r[1], kelas: r[2], n: cnt[k], jml: sum[k], avg: Math.round(sum[k] / cnt[k] * 100) / 100 }; });
  rows.sort((a, b) => b.avg - a.avg || (a.nama < b.nama ? -1 : 1));
  let rk = 0, pv = null; rows.forEach((r, i) => { if (r.avg !== pv) { rk = i + 1; pv = r.avg; } r.rank = rk; });
  const g = {}; rows.forEach(r => (g[r.kelas] = g[r.kelas] || []).push(r));
  Object.keys(g).forEach(k => { let x = 0, p = null; g[k].forEach((r, i) => { if (r.avg !== p) { x = i + 1; p = r.avg; } r.rk = x; }); });
  return { rows, tot: sh.length };
}

// ===== API JSON untuk frontend di Netlify (fetch POST) =====
const API = { init, login, getKelas, statistik, cari, saveNilai, listSiswa, saveSiswa, delSiswa, importSiswa, saveWali, addMapel, editMapel, delMapel, susulan, saveSusulan, peringkat };
function doPost(e) {
  let out;
  try {
    const q = JSON.parse(e.postData.contents);
    if (!API[q.fn]) throw new Error('Fungsi tidak dikenal');
    out = { ok: true, data: API[q.fn].apply(null, q.args || []) };
  } catch (err) { out = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
