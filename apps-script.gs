/**
 * PFS Projesi - Google Apps Script (Web uygulaması)
 * Dağıtım: Dağıt > Dağıtımları yönet > (mevcut dağıtım) > Düzenle > Yeni sürüm
 *   Çalıştıran: Ben   |   Erişim: Herkes
 * (Aynı dağıtımı yeni sürüme güncellerseniz web uygulaması URL'si değişmez.)
 *
 * Veri sayfası: "Veri" (yoksa otomatik oluşturulur, başlıklar otomatik yazılır).
 * Her katılımcı = 1 satır.
 */
const SHEET_NAME = 'Veri';
const HEADERS = [
  'zaman_damgasi', 'katilimci_id', 'surum', 'onam_surumu', 'onam_zamani',
  'cinsiyet', 'yas_araligi', 'ehliyet', 'kaza',
  'cihaz', 'girdi', 'yenileme_hz', 'olcek_sirasi', 'hava_sirasi',
  'kuru_klasik', 'kuru_mesafe', 'kuru_pfs',
  'islak_klasik', 'islak_mesafe', 'islak_pfs',
  'buzlu_klasik', 'buzlu_mesafe', 'buzlu_pfs',
  'toplam_sure_sn'
];
const SPEED_COLS = HEADERS.slice(14, 23);
const ALLOWED = {
  cinsiyet: ['Kadın', 'Erkek'],
  yas_araligi: ['18–25', '26–45', '46+'],
  ehliyet: ['Var', 'Yok'],
  kaza: ['Hayır', 'Evet'],
  cihaz: ['mobil', 'masaustu']
};

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function validate_(d) {
  if (!d || d.action !== 'submit') return 'Geçersiz istek';
  if (!/^PFS-[A-Z0-9]{8}$/.test(String(d.katilimci_id || ''))) return 'Geçersiz katılımcı kodu';
  if (!d.onam_zamani || !d.onam_surumu) return 'Onam bilgisi eksik';
  for (const k in ALLOWED) if (ALLOWED[k].indexOf(d[k]) < 0) return 'Geçersiz alan: ' + k;
  for (const c of SPEED_COLS) {
    const v = Number(d[c]);
    if (!isFinite(v) || v < 40 || v > 145) return 'Geçersiz hız: ' + c;
  }
  return null;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const err = validate_(d);
    if (err) return out_({ ok: false, error: err });
    const sh = getSheet_();
    const last = sh.getLastRow();
    if (last > 1) {
      const ids = sh.getRange(2, 2, last - 1, 1).getValues().flat();
      if (ids.indexOf(d.katilimci_id) >= 0) return out_({ ok: true, duplicate: true });
    }
    const row = HEADERS.map(h => h === 'zaman_damgasi' ? new Date() : (d[h] === undefined ? '' : d[h]));
    sh.appendRow(row);
    return out_({ ok: true });
  } catch (ex) {
    return out_({ ok: false, error: String(ex) });
  } finally {
    lock.releaseLock();
  }
}

// Web sayfası bu uç noktadan anonim satırları okur (katılımcı kodu hariç).
function doGet(e) {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return out_({ ok: true, rows: [] });
  const vals = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const rows = vals.map(r => {
    const o = {};
    HEADERS.forEach((h, i) => {
      if (h === 'katilimci_id') return;
      o[h] = (r[i] instanceof Date) ? r[i].toISOString() : r[i];
    });
    return o;
  });
  return out_({ ok: true, rows: rows });
}
