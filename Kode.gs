/**
 * OKEGAS CRM / SheetBlast WA Backend
 * PRD v1.2 Implementation
 */

var SPREADSHEET_ID = '1a-6aEajSj519aJqv41jQunHEQThxg0J4_rnrltST_gI';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('OKEGAS CRM - SheetBlast WA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Endpoint doPost untuk Integrasi n8n (BR-10, F-23, F-26)
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var content = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    var action = content.action || (e && e.parameter ? e.parameter.action : '');
    var token = content.token || (e && e.parameter ? e.parameter.token : '');

    var validToken = getSettingValue('webhook_token', 'SECRET_DEFAULT_KEY');
    if (token !== validToken) {
      return jsonResponse({ ok: false, message: 'Unauthorized webhook token' });
    }

    var result;
    if (action === 'ambil_antrean') {
      result = claimQueueItems(content.id_job, content.batas || 10);
    } else if (action === 'lapor_kirim') {
      result = reportSendResult(content);
    } else if (action === 'lapor_pengirim') {
      result = reportSenderStatus(content.id_pengirim, content.status, content.alasan);
    } else {
      result = { ok: false, message: 'Aksi webhook tidak dikenali: ' + action };
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ ok: false, message: 'Webhook error: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(serializeForClient(obj)))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * SETUP DATABASE (Idempotent: update schema & preserve existing rows)
 */
function setupDatabase() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  var schemas = [
    { name: 'Kontak', headers: ['id', 'nama', 'no_hp', 'email', 'kota', 'sumber', 'tahap', 'opt_out', 'catatan', 'terakhir_dikirim_pada', 'terakhir_pengirim_id', 'terakhir_kampanye_id', 'dibuat_pada', 'diubah_pada'] },
    { name: 'Kontak_Tag', headers: ['id_kontak', 'id_tag', 'ditambahkan_pada'] },
    { name: 'Tag', headers: ['id', 'nama', 'keterangan', 'dibuat_pada'] },
    { name: 'Tag_Aturan', headers: ['id', 'id_tag_pemicu', 'aksi', 'id_tag_target'] },
    { name: 'Segmen', headers: ['id', 'nama', 'kondisi_json', 'dibuat_pada'] },
    { name: 'Template', headers: ['id', 'nama', 'isi_pesan', 'variabel', 'aktif', 'dibuat_pada'] },
    { name: 'Pengirim', headers: ['id', 'label', 'no_hp', 'device_id', 'status', 'kuota_harian', 'terpakai_hari_ini', 'tanggal_reset', 'terakhir_dipakai_pada', 'catatan'] },
    { name: 'Broadcast_Job', headers: ['id', 'kampanye_id', 'id_job_induk', 'nama', 'id_template', 'id_pengirim', 'jeda_min', 'jeda_max', 'status', 'jumlah_target', 'dibuat_pada'] },
    { name: 'Broadcast_Job_Aksi', headers: ['id_job', 'id_tag', 'aksi'] },
    { name: 'Broadcast_Kirim', headers: ['id', 'id_job', 'id_kontak', 'no_hp', 'isi_pesan', 'status', 'alasan', 'id_klaim', 'waktu_klaim', 'percobaan', 'waktu_update'] },
    { name: 'Log', headers: ['waktu', 'aksi', 'id_data', 'detail'] },
    { name: 'Pengaturan', headers: ['kunci', 'nilai'] }
  ];

  schemas.forEach(function(s) {
    var sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
    }
    // Update Header Baris 1
    sheet.getRange(1, 1, 1, s.headers.length).setValues([s.headers]);
  });

  // Seed data default jika sheet Pengaturan kosong
  var setSheet = ss.getSheetByName('Pengaturan');
  if (setSheet.getLastRow() <= 1) {
    var defaultSettings = [
      ['cooldown_hari', '7'],
      ['jeda_default_min', '1'],
      ['jeda_default_max', '10'],
      ['kuota_default', '50'],
      ['jam_kirim_mulai', '08:00'],
      ['jam_kirim_selesai', '21:00'],
      ['webhook_token', 'OKEGAS_' + Utilities.getUuid().substring(0, 8)]
    ];
    setSheet.getRange(2, 1, defaultSettings.length, 2).setValues(defaultSettings);
  }

  // Seed default Tag & Rule jika kosong (AC-05 demo)
  var tagSheet = ss.getSheetByName('Tag');
  if (tagSheet.getLastRow() <= 1) {
    var now = new Date().toISOString();
    tagSheet.appendRow(['tag_tertarik', 'Tertarik', 'Lead menunjukkan minat', now]);
    tagSheet.appendRow(['tag_dingin', 'Dingin', 'Lead tidak merespon', now]);
    
    var ruleSheet = ss.getSheetByName('Tag_Aturan');
    ruleSheet.appendRow(['rule_1', 'tag_tertarik', 'hapus', 'tag_dingin']);
  }

  writeLog('SETUP_DATABASE', 'SYSTEM', 'Database initialized or refreshed headers.');
  return { ok: true, message: 'Database schema successfully updated.' };
}

/**
 * SERIALIZATION HELPER (NO Date objects returned)
 */
function serializeForClient(value) {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (Array.isArray(value)) {
    return value.map(function(item) { return serializeForClient(item); });
  }
  if (typeof value === 'object') {
    var out = {};
    for (var k in value) {
      if (Object.prototype.hasOwnProperty.call(value, k)) {
        out[k] = serializeForClient(value[k]);
      }
    }
    return out;
  }
  return value;
}

/**
 * PHONE NUMBER NORMALIZATION (BR-12, AC-12)
 * Removes non-digits. 08xx -> 628xx. 8xx -> 628xx. 628xx unchanged.
 */
function normalizePhoneNumber(raw) {
  if (!raw) return '';
  var clean = ('' + raw).replace(/\D/g, '');
  if (clean.indexOf('0') === 0) {
    clean = '62' + clean.substring(1);
  } else if (clean.indexOf('8') === 0) {
    clean = '62' + clean;
  }
  return clean;
}

/**
 * DASHBOARD METRICS (F-29)
 */
function getDashboardData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    checkAndResetDailyQuotas();
    checkAndResetStalledClaims();

    var kontakSheet = ss.getSheetByName('Kontak');
    var kData = kontakSheet.getDataRange().getValues();
    var totalKontak = Math.max(0, kData.length - 1);
    
    var tahapStats = { baru: 0, dihubungi: 0, tertarik: 0, closing: 0, pelanggan: 0, tidak_tertarik: 0 };
    for (var i = 1; i < kData.length; i++) {
      var stg = String(kData[i][6]).toLowerCase() || 'baru';
      if (tahapStats[stg] !== undefined) {
        tahapStats[stg]++;
      } else {
        tahapStats[stg] = 1;
      }
    }

    var pengirimSheet = ss.getSheetByName('Pengirim');
    var pData = pengirimSheet.getDataRange().getValues();
    var pengirimList = [];
    var alertPengirim = [];
    var terkirimHariIni = 0;
    var pengirimAktifCount = 0;

    for (var j = 1; j < pData.length; j++) {
      var p = {
        id: String(pData[j][0]),
        label: String(pData[j][1]),
        no_hp: String(pData[j][2]),
        device_id: String(pData[j][3]),
        status: String(pData[j][4]),
        kuota_harian: Number(pData[j][5]) || 0,
        terpakai_hari_ini: Number(pData[j][6]) || 0,
        terakhir_dipakai_pada: pData[j][8] ? String(pData[j][8]) : ''
      };
      terkirimHariIni += p.terpakai_hari_ini;
      if (p.status === 'aktif') pengirimAktifCount++;
      if (p.status === 'banned' || p.status === 'limit' || (p.kuota_harian > 0 && p.terpakai_hari_ini >= p.kuota_harian * 0.8)) {
        alertPengirim.push(p);
      }
      pengirimList.push(p);
    }

    var jobSheet = ss.getSheetByName('Broadcast_Job');
    var jData = jobSheet.getDataRange().getValues();
    var recentBroadcasts = [];
    var jobAktifCount = 0;

    for (var k = jData.length - 1; k >= 1 && recentBroadcasts.length < 5; k--) {
      var jStatus = String(jData[k][8]);
      if (jStatus === 'berjalan') jobAktifCount++;
      recentBroadcasts.push({
        id: String(jData[k][0]),
        nama: String(jData[k][3]),
        pengirim_label: getSenderLabelById(String(jData[k][5])),
        jumlah_target: Number(jData[k][9]) || 0,
        status: jStatus,
        dibuat_pada: jData[k][10] ? String(jData[k][10]) : ''
      });
    }

    return {
      ok: true,
      data: serializeForClient({
        totalKontak: totalKontak,
        terkirimHariIni: terkirimHariIni,
        pengirimAktifCount: pengirimAktifCount,
        jobAktifCount: jobAktifCount,
        tahapStats: tahapStats,
        pengirimList: pengirimList,
        alertPengirim: alertPengirim,
        recentBroadcasts: recentBroadcasts
      })
    };
  } catch (err) {
    return { ok: false, message: 'Gagal memuat dashboard: ' + err.message };
  }
}

/**
 * CONTACT MANAGEMENT (F-01 s/d F-06, BR-01, AC-01)
 */
function getKontakList(filter) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Kontak');
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { ok: true, data: [] };

    var tagMap = getKontakTagMap();
    var senderMap = getSenderLabelMap();
    var results = [];

    var search = (filter && filter.search) ? String(filter.search).toLowerCase() : '';
    var tahap = (filter && filter.tahap) ? String(filter.tahap).toLowerCase() : '';
    var tagFilter = (filter && filter.tag) ? String(filter.tag) : '';
    var optOutFilter = (filter && filter.optOut !== undefined && filter.optOut !== '') ? String(filter.optOut) === 'true' : null;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var kId = String(row[0]);
      var nama = String(row[1]);
      var noHp = String(row[2]);
      var email = String(row[3]);
      var kota = String(row[4]);
      var sumber = String(row[5]);
      var kTahap = String(row[6]);
      var optOut = !!row[7];
      var catatan = String(row[8]);
      var tglKirim = row[9] ? String(row[9]) : '';
      var pengirimId = String(row[10]);

      if (search) {
        var match = nama.toLowerCase().indexOf(search) > -1 ||
                    noHp.indexOf(search) > -1 ||
                    email.toLowerCase().indexOf(search) > -1 ||
                    kota.toLowerCase().indexOf(search) > -1;
        if (!match) continue;
      }

      if (tahap && kTahap.toLowerCase() !== tahap) continue;
      if (optOutFilter !== null && optOut !== optOutFilter) continue;

      var tags = tagMap[kId] || [];
      if (tagFilter) {
        var hasTag = tags.some(function(t) { return t.id === tagFilter; });
        if (!hasTag) continue;
      }

      results.push({
        id: kId,
        nama: nama,
        no_hp: noHp,
        email: email,
        kota: kota,
        sumber: sumber,
        tahap: kTahap,
        opt_out: optOut,
        catatan: catatan,
        terakhir_dikirim_pada: tglKirim,
        terakhir_pengirim_id: pengirimId,
        terakhir_pengirim_label: senderMap[pengirimId] || '',
        tags: tags
      });
    }

    return { ok: true, data: serializeForClient(results.slice(0, 100)) }; // Pagination 50-100 baris
  } catch (err) {
    return { ok: false, message: 'Gagal mengambil data kontak: ' + err.message };
  }
}

function saveKontak(kontakData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Kontak');
    var normalizedPhone = normalizePhoneNumber(kontakData.no_hp);

    if (!normalizedPhone || normalizedPhone.length < 9) {
      return { ok: false, message: 'Format nomor WhatsApp tidak valid setelah normalisasi: ' + kontakData.no_hp };
    }

    var data = sheet.getDataRange().getValues();
    var isUpdate = !!kontakData.id;
    var targetRow = -1;

    for (var i = 1; i < data.length; i++) {
      var currentId = String(data[i][0]);
      var currentPhone = String(data[i][2]);
      
      // Duplication Check (BR-01)
      if (currentPhone === normalizedPhone && (!isUpdate || currentId !== kontakData.id)) {
        return { ok: false, message: 'Nomor ' + normalizedPhone + ' sudah terdaftar pada kontak: ' + data[i][1] };
      }

      if (isUpdate && currentId === kontakData.id) {
        targetRow = i + 1;
      }
    }

    var now = new Date().toISOString();
    var idKontak = isUpdate ? kontakData.id : 'ktk_' + Utilities.getUuid().substring(0, 8);

    if (isUpdate && targetRow > 0) {
      sheet.getRange(targetRow, 2, 1, 8).setValues([[
        kontakData.nama,
        normalizedPhone,
        kontakData.email || '',
        kontakData.kota || '',
        kontakData.sumber || '',
        kontakData.tahap || 'baru',
        kontakData.opt_out ? true : false,
        kontakData.catatan || ''
      ]]);
      sheet.getRange(targetRow, 14).setValue(now);
      writeLog('UBAH_KONTAK', idKontak, 'Mengubah data kontak ' + kontakData.nama);
    } else {
      var newRow = [
        idKontak,
        kontakData.nama,
        normalizedPhone,
        kontakData.email || '',
        kontakData.kota || '',
        kontakData.sumber || '',
        kontakData.tahap || 'baru',
        kontakData.opt_out ? true : false,
        kontakData.catatan || '',
        '', '', '', // Terakhir dikirim
        now, now
      ];
      sheet.appendRow(newRow);
      writeLog('TAMBAH_KONTAK', idKontak, 'Menambah kontak baru ' + kontakData.nama);
    }

    // Update Tag Relasional
    if (kontakData.tagIds && Array.isArray(kontakData.tagIds)) {
      syncKontakTags(idKontak, kontakData.tagIds);
    }

    return { ok: true, data: { id: idKontak } };
  } catch (err) {
    return { ok: false, message: 'Gagal menyimpan kontak: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteKontak(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Kontak');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        removeKontakTags(id);
        writeLog('HAPUS_KONTAK', id, 'Menghapus kontak');
        return { ok: true };
      }
    }
    return { ok: false, message: 'Kontak tidak ditemukan.' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function importKontakCsv(csvString) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Kontak');
    var existingPhones = {};
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      existingPhones[String(data[i][2])] = true;
    }

    var lines = csvString.split('\n');
    var importedCount = 0;
    var rejected = [];
    var now = new Date().toISOString();
    var rowsToAppend = [];

    lines.forEach(function(line, idx) {
      var trimmed = line.trim();
      if (!trimmed) return;
      var parts = trimmed.split(',');
      if (parts.length < 2) {
        rejected.push('Baris ' + (idx + 1) + ': Format kolom tidak lengkap.');
        return;
      }
      var nama = parts[0].trim();
      var noHp = normalizePhoneNumber(parts[1].trim());

      if (!nama || !noHp || noHp.length < 9) {
        rejected.push('Baris ' + (idx + 1) + ': Nama atau format nomor HP tidak valid (' + parts[1] + ')');
        return;
      }

      if (existingPhones[noHp]) {
        rejected.push('Baris ' + (idx + 1) + ': Nomor ' + noHp + ' duplikat/sudah ada di database (AC-01).');
        return;
      }

      existingPhones[noHp] = true;
      var id = 'ktk_' + Utilities.getUuid().substring(0, 8);
      rowsToAppend.push([
        id,
        nama,
        noHp,
        (parts[2] || '').trim(),
        (parts[3] || '').trim(),
        (parts[4] || '').trim(),
        (parts[5] || 'baru').trim().toLowerCase(),
        false,
        'Impor CSV ' + now.substring(0, 10),
        '', '', '',
        now, now
      ]);
      importedCount++;
    });

    if (rowsToAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
      writeLog('IMPOR_CSV', 'SYSTEM', 'Berhasil mengimpor ' + importedCount + ' kontak.');
    }

    return {
      ok: true,
      data: {
        imported: importedCount,
        rejected: rejected
      }
    };
  } catch (err) {
    return { ok: false, message: 'Gagal impor kontak: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * TAG & ACTION RULES (F-07, F-08, AC-05)
 */
function getTagList() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Tag');
    var data = sheet.getDataRange().getValues();
    var counts = getTagUsageCounts();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var id = String(data[i][0]);
      list.push({
        id: id,
        nama: String(data[i][1]),
        keterangan: String(data[i][2]),
        dibuat_pada: String(data[i][3]),
        kontakCount: counts[id] || 0
      });
    }
    return { ok: true, data: serializeForClient(list) };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function getTagAturanList() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Tag_Aturan');
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      list.push({
        id: String(data[i][0]),
        id_tag_pemicu: String(data[i][1]),
        aksi: String(data[i][2]),
        id_tag_target: String(data[i][3])
      });
    }
    return { ok: true, data: serializeForClient(list) };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function deleteTag(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Tag');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    // Hapus juga asosiasi di Kontak_Tag
    var ktSheet = ss.getSheetByName('Kontak_Tag');
    var ktData = ktSheet.getDataRange().getValues();
    for (var j = ktData.length - 1; j >= 1; j--) {
      if (String(ktData[j][1]) === String(id)) {
        ktSheet.deleteRow(j + 1);
      }
    }
    writeLog('HAPUS_TAG', id, 'Menghapus tag');
    return { ok: true };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * Non-cascading Tag Rule Execution (AC-05, BR-03)
 */
function applyTagToContact(contactId, tagId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var ktSheet = ss.getSheetByName('Kontak_Tag');
  var ktData = ktSheet.getDataRange().getValues();

  var alreadyHas = false;
  for (var i = 1; i < ktData.length; i++) {
    if (String(ktData[i][0]) === String(contactId) && String(ktData[i][1]) === String(tagId)) {
      alreadyHas = true;
      break;
    }
  }
  if (!alreadyHas) {
    ktSheet.appendRow([contactId, tagId, new Date().toISOString()]);
  }

  // Cari Tag_Aturan untuk pemicu tagId (TIDAK BERANTAI)
  var ruleSheet = ss.getSheetByName('Tag_Aturan');
  var rData = ruleSheet.getDataRange().getValues();
  for (var r = 1; r < rData.length; r++) {
    if (String(rData[r][1]) === String(tagId)) {
      var aksi = String(rData[r][2]).toLowerCase();
      var targetTagId = String(rData[r][3]);

      if (aksi === 'hapus') {
        removeTagFromContactDirect(contactId, targetTagId);
      } else if (aksi === 'tambah') {
        addTagToContactDirect(contactId, targetTagId);
      }
    }
  }
}

function removeTagFromContactDirect(contactId, tagId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var ktSheet = ss.getSheetByName('Kontak_Tag');
  var ktData = ktSheet.getDataRange().getValues();
  for (var i = ktData.length - 1; i >= 1; i--) {
    if (String(ktData[i][0]) === String(contactId) && String(ktData[i][1]) === String(tagId)) {
      ktSheet.deleteRow(i + 1);
    }
  }
}

function addTagToContactDirect(contactId, tagId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var ktSheet = ss.getSheetByName('Kontak_Tag');
  var ktData = ktSheet.getDataRange().getValues();
  for (var i = 1; i < ktData.length; i++) {
    if (String(ktData[i][0]) === String(contactId) && String(ktData[i][1]) === String(tagId)) {
      return;
    }
  }
  ktSheet.appendRow([contactId, tagId, new Date().toISOString()]);
}

/**
 * PENGIRIM MANAGEMENT (F-11, F-12, BR-07, BR-09)
 */
function getPengirimList() {
  try {
    checkAndResetDailyQuotas();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Pengirim');
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      list.push({
        id: String(data[i][0]),
        label: String(data[i][1]),
        no_hp: String(data[i][2]),
        device_id: String(data[i][3]),
        status: String(data[i][4]),
        kuota_harian: Number(data[i][5]) || 0,
        terpakai_hari_ini: Number(data[i][6]) || 0,
        terakhir_dipakai_pada: data[i][8] ? String(data[i][8]) : '',
        catatan: String(data[i][9])
      });
    }
    return { ok: true, data: serializeForClient(list) };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function updatePengirimStatus(id, newStatus) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Pengirim');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 5).setValue(newStatus);
        writeLog('STATUS_PENGIRIM', id, 'Status diubah ke ' + newStatus);
        return { ok: true };
      }
    }
    return { ok: false, message: 'Pengirim tidak ditemukan.' };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function hentikanSemuaPengirim() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var jobSheet = ss.getSheetByName('Broadcast_Job');
    var data = jobSheet.getDataRange().getValues();
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][8]) === 'berjalan' || String(data[i][8]) === 'draft') {
        jobSheet.getRange(i + 1, 9).setValue('dihentikan');
        count++;
      }
    }
    writeLog('HENTIKAN_SEMUA', 'SYSTEM', 'Menghentikan darurat ' + count + ' broadcast job.');
    return { ok: true, count: count };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * TEMPLATES (F-10)
 */
function getTemplateList() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Template');
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      list.push({
        id: String(data[i][0]),
        nama: String(data[i][1]),
        isi_pesan: String(data[i][2]),
        variabel: String(data[i][3]),
        aktif: !!data[i][4],
        dibuat_pada: String(data[i][5])
      });
    }
    return { ok: true, data: serializeForClient(list) };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * BROADCAST ENGINE & PRE-FLIGHT (F-13 s/d F-24, AC-02, AC-09, AC-11)
 */
function calculateBroadcastPreflight(params) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var kontakSheet = ss.getSheetByName('Kontak');
    var kData = kontakSheet.getDataRange().getValues();

    var activeContactIds = getActiveJobContactIds();
    var cooldownDays = Number(getSettingValue('cooldown_hari', 7));
    var nowMs = new Date().getTime();
    var cooldownThresholdMs = nowMs - (cooldownDays * 24 * 60 * 60 * 1000);

    var totalEligible = 0;
    var inCooldown = 0;
    var inActiveJob = 0;
    var optedOut = 0;

    for (var i = 1; i < kData.length; i++) {
      var kId = String(kData[i][0]);
      var isOptOut = !!kData[i][7];
      var lastSentStr = kData[i][9];

      if (isOptOut) {
        optedOut++;
        continue;
      }

      if (activeContactIds[kId]) {
        inActiveJob++; // AC-09: dilewati karena di job aktif lain
        continue;
      }

      if (lastSentStr) {
        var lastSentDate = new Date(lastSentStr);
        if (!isNaN(lastSentDate.getTime()) && lastSentDate.getTime() > cooldownThresholdMs) {
          inCooldown++; // AC-11: ditandai cooldown
          continue;
        }
      }

      totalEligible++;
    }

    return {
      ok: true,
      data: {
        totalEligible: totalEligible,
        inCooldown: inCooldown,
        inActiveJob: inActiveJob,
        optedOut: optedOut
      }
    };
  } catch (err) {
    return { ok: false, message: 'Preflight calculation failed: ' + err.message };
  }
}

function createBroadcastJob(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Periksa kuota dan status pengirim
    var pengirimSheet = ss.getSheetByName('Pengirim');
    var pData = pengirimSheet.getDataRange().getValues();
    var senderRow = -1;
    var senderObj = null;

    for (var p = 1; p < pData.length; p++) {
      if (String(pData[p][0]) === String(payload.idPengirim)) {
        senderRow = p + 1;
        senderObj = {
          status: String(pData[p][4]),
          kuota: Number(pData[p][5]) || 0,
          terpakai: Number(pData[p][6]) || 0
        };
        break;
      }
    }

    if (!senderObj || senderObj.status !== 'aktif') {
      return { ok: false, message: 'Nomor pengirim tidak aktif atau tidak ditemukan.' };
    }
    if (senderObj.terpakai >= senderObj.kuota) {
      return { ok: false, message: 'Pengirim telah mencapai batas kuota harian (limit).' };
    }

    // BR-08 / F-16: Periksa apakah ada job aktif pada kampanye
    var kampanyeId = payload.kampanyeId || ('camp_' + Utilities.getUuid().substring(0, 8));
    var jobSheet = ss.getSheetByName('Broadcast_Job');
    var jData = jobSheet.getDataRange().getValues();
    for (var j = 1; j < jData.length; j++) {
      if (String(jData[j][1]) === kampanyeId && String(jData[j][8]) === 'berjalan') {
        return { ok: false, message: 'Masih ada job aktif dalam kampanye ini (BR-08 / AC-06).' };
      }
    }

    var kontakSheet = ss.getSheetByName('Kontak');
    var kData = kontakSheet.getDataRange().getValues();
    var activeContactIds = getActiveJobContactIds();
    var cooldownDays = Number(getSettingValue('cooldown_hari', 7));
    var now = new Date();
    var cooldownThresholdMs = now.getTime() - (cooldownDays * 24 * 60 * 60 * 1000);

    var idJob = 'job_' + Utilities.getUuid().substring(0, 8);
    var recipients = [];
    var skippedDueToOptOut = 0;
    var skippedDueToActive = 0;

    for (var i = 1; i < kData.length; i++) {
      var kId = String(kData[i][0]);
      var nama = String(kData[i][1]);
      var noHp = String(kData[i][2]);
      var isOptOut = !!kData[i][7];
      var lastSentStr = kData[i][9];

      // F-14 & AC-02: Kontak opt-out dilewati
      if (isOptOut) {
        skippedDueToOptOut++;
        continue;
      }

      // F-19 & AC-09: Lewati jika di job aktif lain
      if (activeContactIds[kId]) {
        skippedDueToActive++;
        continue;
      }

      // Cooldown check (F-20, AC-11)
      if (!payload.overrideCooldown && lastSentStr) {
        var lastSentDate = new Date(lastSentStr);
        if (!isNaN(lastSentDate.getTime()) && lastSentDate.getTime() > cooldownThresholdMs) {
          continue;
        }
      }

      var msgBody = (payload.isiPesan || '')
        .replace(/{nama}/g, nama)
        .replace(/{kota}/g, String(kData[i][4]) || '')
        .replace(/{sumber}/g, String(kData[i][5]) || '');

      recipients.push([
        'bkirim_' + Utilities.getUuid().substring(0, 8),
        idJob,
        kId,
        noHp,
        msgBody,
        'antre', // BR-02: Status awal antre
        '',
        '',
        '',
        0,
        now.toISOString()
      ]);
    }

    if (recipients.length === 0) {
      return { ok: false, message: 'Tidak ada kontak yang memenuhi syarat untuk dikirim broadcast.' };
    }

    // Simpan Job
    var newJobRow = [
      idJob,
      kampanyeId,
      payload.idJobInduk || '',
      payload.namaJob || ('Broadcast ' + now.toLocaleDateString()),
      payload.idTemplate,
      payload.idPengirim,
      payload.jedaMin || 1,
      payload.jedaMax || 10,
      'berjalan',
      recipients.length,
      now.toISOString()
    ];
    jobSheet.appendRow(newJobRow);

    // Simpan Antrean Kirim secara batch
    var kirimSheet = ss.getSheetByName('Broadcast_Kirim');
    kirimSheet.getRange(kirimSheet.getLastRow() + 1, 1, recipients.length, recipients[0].length).setValues(recipients);

    // Simpan Aksi Tag Job
    var aksiSheet = ss.getSheetByName('Broadcast_Job_Aksi');
    if (payload.tagTambahIds && payload.tagTambahIds.length > 0) {
      payload.tagTambahIds.forEach(function(tid) {
        aksiSheet.appendRow([idJob, tid, 'tambah']);
      });
    }
    if (payload.tagHapusIds && payload.tagHapusIds.length > 0) {
      payload.tagHapusIds.forEach(function(tid) {
        aksiSheet.appendRow([idJob, tid, 'hapus']);
      });
    }

    writeLog('BUAT_BROADCAST', idJob, 'Membuat broadcast job ke ' + recipients.length + ' penerima.');
    return { ok: true, data: { id_job: idJob, count: recipients.length } };
  } catch (err) {
    return { ok: false, message: 'Gagal membuat job broadcast: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Buat Ulang Sisa Job (F-24, AC-04)
 */
function retryFailedBroadcastJob(parentJobId, newPengirimId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var jobSheet = ss.getSheetByName('Broadcast_Job');
    var jData = jobSheet.getDataRange().getValues();
    var parentJob = null;

    for (var j = 1; j < jData.length; j++) {
      if (String(jData[j][0]) === String(parentJobId)) {
        parentJob = {
          kampanye_id: String(jData[j][1]),
          nama: String(jData[j][3]),
          id_template: String(jData[j][4]),
          jeda_min: Number(jData[j][6]),
          jeda_max: Number(jData[j][7])
        };
        break;
      }
    }

    if (!parentJob) return { ok: false, message: 'Job induk tidak ditemukan.' };

    // Dapatkan semua kontak yang sudah 'terkirim' di kampanye ini (AC-04)
    var alreadySentInCampaign = {};
    var kirimSheet = ss.getSheetByName('Broadcast_Kirim');
    var kData = kirimSheet.getDataRange().getValues();

    for (var k = 1; k < kData.length; k++) {
      var rowJobId = String(kData[k][1]);
      var rowStatus = String(kData[k][5]);
      var contactId = String(kData[k][2]);

      if (rowStatus === 'terkirim' && isJobInCampaign(rowJobId, parentJob.kampanye_id)) {
        alreadySentInCampaign[contactId] = true;
      }
    }

    // Ambil penerima dari parentJobId yang gagal atau belum terkirim
    var newRecipients = [];
    var now = new Date();
    var newJobId = 'job_' + Utilities.getUuid().substring(0, 8);

    for (var m = 1; m < kData.length; m++) {
      if (String(kData[m][1]) === String(parentJobId)) {
        var status = String(kData[m][5]);
        var cId = String(kData[m][2]);

        if (status !== 'terkirim' && !alreadySentInCampaign[cId]) {
          newRecipients.push([
            'bkirim_' + Utilities.getUuid().substring(0, 8),
            newJobId,
            cId,
            kData[m][3], // no_hp
            kData[m][4], // isi_pesan
            'antre',
            '', '', '', 0,
            now.toISOString()
          ]);
        }
      }
    }

    if (newRecipients.length === 0) {
      return { ok: false, message: 'Tidak ada sisa penerima yang belum terkirim untuk dibuat ulang.' };
    }

    // Simpan Child Job
    jobSheet.appendRow([
      newJobId,
      parentJob.kampanye_id,
      parentJobId,
      parentJob.nama + ' (Sisa Rerun)',
      parentJob.id_template,
      newPengirimId,
      parentJob.jeda_min,
      parentJob.jeda_max,
      'berjalan',
      newRecipients.length,
      now.toISOString()
    ]);

    kirimSheet.getRange(kirimSheet.getLastRow() + 1, 1, newRecipients.length, newRecipients[0].length).setValues(newRecipients);
    writeLog('RERUN_JOB', newJobId, 'Membuat job rerun sisa dari ' + parentJobId);

    return { ok: true, data: { count: newRecipients.length, newJobId: newJobId } };
  } catch (err) {
    return { ok: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function getBroadcastJobList() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Broadcast_Job');
    var data = sheet.getDataRange().getValues();
    var senderMap = getSenderLabelMap();
    var list = [];
    for (var i = data.length - 1; i >= 1; i--) {
      list.push({
        id: String(data[i][0]),
        kampanye_id: String(data[i][1]),
        id_job_induk: String(data[i][2]),
        nama: String(data[i][3]),
        id_pengirim: String(data[i][5]),
        pengirim_label: senderMap[String(data[i][5])] || '',
        jeda_min: Number(data[i][6]),
        jeda_max: Number(data[i][7]),
        status: String(data[i][8]),
        jumlah_target: Number(data[i][9]),
        dibuat_pada: String(data[i][10])
      });
    }
    return { ok: true, data: serializeForClient(list) };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function getBroadcastJobDetail(jobId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var jobSheet = ss.getSheetByName('Broadcast_Job');
    var jData = jobSheet.getDataRange().getValues();
    var jobInfo = null;

    for (var i = 1; i < jData.length; i++) {
      if (String(jData[i][0]) === String(jobId)) {
        jobInfo = {
          id: String(jData[i][0]),
          nama: String(jData[i][3]),
          status: String(jData[i][8]),
          recipients: []
        };
        break;
      }
    }

    if (!jobInfo) return { ok: false, message: 'Job tidak ditemukan' };

    var kirimSheet = ss.getSheetByName('Broadcast_Kirim');
    var kData = kirimSheet.getDataRange().getValues();

    for (var j = 1; j < kData.length; j++) {
      if (String(kData[j][1]) === String(jobId)) {
        jobInfo.recipients.push({
          id: String(kData[j][0]),
          no_hp: String(kData[j][3]),
          isi_pesan: String(kData[j][4]),
          status: String(kData[j][5]),
          alasan: String(kData[j][6])
        });
      }
    }

    return { ok: true, data: serializeForClient(jobInfo) };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/**
 * ATOMIC QUEUE CLAIM FOR N8N (BR-10, US-4.1, US-4.2)
 */
function claimQueueItems(idJob, limit) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var kirimSheet = ss.getSheetByName('Broadcast_Kirim');
  var data = kirimSheet.getDataRange().getValues();

  var claimed = [];
  var claimId = 'claim_' + Utilities.getUuid().substring(0, 8);
  var now = new Date().toISOString();
  limit = limit || 10;

  for (var i = 1; i < data.length && claimed.length < limit; i++) {
    if (String(data[i][1]) === String(idJob) && String(data[i][5]) === 'antre') {
      var rowIdx = i + 1;
      kirimSheet.getRange(rowIdx, 6).setValue('diproses');
      kirimSheet.getRange(rowIdx, 8).setValue(claimId);
      kirimSheet.getRange(rowIdx, 9).setValue(now);
      kirimSheet.getRange(rowIdx, 10).setValue((Number(data[i][9]) || 0) + 1);

      claimed.push({
        id_kirim: String(data[i][0]),
        id_kontak: String(data[i][2]),
        no_hp: String(data[i][3]),
        isi_pesan: String(data[i][4]),
        id_klaim: claimId
      });
    }
  }

  // Ambil detail jeda dan device
  var jobSheet = ss.getSheetByName('Broadcast_Job');
  var jData = jobSheet.getDataRange().getValues();
  var jedaMin = 1, jedaMax = 10, senderId = '';

  for (var j = 1; j < jData.length; j++) {
    if (String(jData[j][0]) === String(idJob)) {
      jedaMin = Number(jData[j][6]);
      jedaMax = Number(jData[j][7]);
      senderId = String(jData[j][5]);
      break;
    }
  }

  return {
    ok: true,
    data: {
      claim_id: claimId,
      claimed_count: claimed.length,
      jeda_min: jedaMin,
      jeda_max: jedaMax,
      sender_id: senderId,
      items: claimed
    }
  };
}

/**
 * REPORT SEND RESULT FROM N8N (BR-02, BR-11, AC-03)
 */
function reportSendResult(payload) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var kirimSheet = ss.getSheetByName('Broadcast_Kirim');
  var data = kirimSheet.getDataRange().getValues();

  var rowFound = -1;
  var jobRowId = '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(payload.id_kirim) || (String(data[i][1]) === String(payload.id_job) && String(data[i][2]) === String(payload.id_kontak))) {
      rowFound = i + 1;
      jobRowId = String(data[i][1]);
      break;
    }
  }

  if (rowFound > 0) {
    var status = payload.status; // 'terkirim' | 'gagal_permanen' | 'gagal_pengirim_banned'
    var nowStr = new Date().toISOString();
    kirimSheet.getRange(rowFound, 6).setValue(status);
    kirimSheet.getRange(rowFound, 7).setValue(payload.alasan || '');
    kirimSheet.getRange(rowFound, 11).setValue(nowStr);

    if (status === 'terkirim') {
      // 1. Perbarui ringkasan di sheet Kontak (BR-11)
      updateKontakSendSummary(payload.id_kontak, nowStr, payload.id_pengirim, payload.kampanye_id);
      
      // 2. Tambah kuota terpakai pengirim
      incrementPengirimQuota(payload.id_pengirim);

      // 3. Jalankan aksi tag job (BR-03, F-22)
      executeJobTagActions(jobRowId, payload.id_kontak);
    }
  }

  return { ok: true };
}

/**
 * REPORT SENDER ISSUE (F-23, AC-03)
 * Sender banned or limit -> halt job & mark unsent as gagal_pengirim_banned
 */
function reportSenderStatus(idPengirim, status, alasan) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var pSheet = ss.getSheetByName('Pengirim');
  var pData = pSheet.getDataRange().getValues();

  for (var i = 1; i < pData.length; i++) {
    if (String(pData[i][0]) === String(idPengirim)) {
      pSheet.getRange(i + 1, 5).setValue(status);
      pSheet.getRange(i + 1, 10).setValue(alasan || '');
      break;
    }
  }

  if (status === 'banned' || status === 'limit') {
    // Hentikan semua job aktif milik pengirim ini
    var jobSheet = ss.getSheetByName('Broadcast_Job');
    var jData = jobSheet.getDataRange().getValues();
    var affectedJobIds = [];

    for (var j = 1; j < jData.length; j++) {
      if (String(jData[j][5]) === String(idPengirim) && String(jData[j][8]) === 'berjalan') {
        jobSheet.getRange(j + 1, 9).setValue('dihentikan');
        affectedJobIds.push(String(jData[j][0]));
      }
    }

    // Tandai sisa antrean menjadi gagal_pengirim_banned
    var kirimSheet = ss.getSheetByName('Broadcast_Kirim');
    var kData = kirimSheet.getDataRange().getValues();
    for (var k = 1; k < kData.length; k++) {
      var rowJob = String(kData[k][1]);
      var rowStatus = String(kData[k][5]);
      if (affectedJobIds.indexOf(rowJob) > -1 && (rowStatus === 'antre' || rowStatus === 'diproses')) {
        kirimSheet.getRange(k + 1, 6).setValue('gagal_pengirim_banned');
        kirimSheet.getRange(k + 1, 7).setValue('Pengirim ' + status + ': ' + alasan);
      }
    }
  }

  writeLog('LAPOR_PENGIRIM', idPengirim, 'Pengirim dilaporkan ' + status + ': ' + alasan);
  return { ok: true };
}

/**
 * SELF-HEALING: Reset Stalled Claims > 10 Minutes (BR-10, AC-10)
 */
function checkAndResetStalledClaims() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Broadcast_Kirim');
    var data = sheet.getDataRange().getValues();
    var tenMinutesAgo = new Date(new Date().getTime() - (10 * 60 * 1000));
    var resetCount = 0;

    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][5]);
      var waktuKlaimStr = data[i][8];

      if (status === 'diproses' && waktuKlaimStr) {
        var claimDate = new Date(waktuKlaimStr);
        if (!isNaN(claimDate.getTime()) && claimDate < tenMinutesAgo) {
          sheet.getRange(i + 1, 6).setValue('antre');
          sheet.getRange(i + 1, 8).setValue(''); // hapus id_klaim
          sheet.getRange(i + 1, 9).setValue(''); // hapus waktu_klaim
          resetCount++;
        }
      }
    }

    if (resetCount > 0) {
      writeLog('RESET_STALLED_CLAIM', 'SYSTEM', 'Mengembalikan ' + resetCount + ' klaim stalled ke antre.');
    }
  } catch (e) {
    // ignore
  }
}

/**
 * QUOTA RESET (BR-09: Asia/Jakarta Midnight)
 */
function checkAndResetDailyQuotas() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Pengirim');
    var data = sheet.getDataRange().getValues();
    var todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');

    for (var i = 1; i < data.length; i++) {
      var resetDate = data[i][7] ? String(data[i][7]).substring(0, 10) : '';
      if (resetDate !== todayStr) {
        sheet.getRange(i + 1, 7).setValue(0); // reset terpakai
        sheet.getRange(i + 1, 8).setValue(todayStr);
        if (String(data[i][4]) === 'limit') {
          sheet.getRange(i + 1, 5).setValue('aktif');
        }
      }
    }
  } catch (e) {
    // ignore
  }
}

/**
 * HELPER UTILITIES
 */
function updateKontakSendSummary(contactId, waktuIso, pengirimId, kampanyeId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Kontak');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(contactId)) {
      sheet.getRange(i + 1, 10).setValue(waktuIso);
      sheet.getRange(i + 1, 11).setValue(pengirimId || '');
      sheet.getRange(i + 1, 12).setValue(kampanyeId || '');
      break;
    }
  }
}

function incrementPengirimQuota(pengirimId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Pengirim');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(pengirimId)) {
      var used = Number(data[i][6]) || 0;
      var quota = Number(data[i][5]) || 0;
      var newUsed = used + 1;
      sheet.getRange(i + 1, 7).setValue(newUsed);
      sheet.getRange(i + 1, 9).setValue(new Date().toISOString());

      if (quota > 0 && newUsed >= quota) {
        sheet.getRange(i + 1, 5).setValue('limit');
      }
      break;
    }
  }
}

function executeJobTagActions(jobId, contactId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var aksiSheet = ss.getSheetByName('Broadcast_Job_Aksi');
  var aData = aksiSheet.getDataRange().getValues();
  for (var i = 1; i < aData.length; i++) {
    if (String(aData[i][0]) === String(jobId)) {
      var tagId = String(aData[i][1]);
      var aksi = String(aData[i][2]);
      if (aksi === 'tambah') {
        applyTagToContact(contactId, tagId);
      } else if (aksi === 'hapus') {
        removeTagFromContactDirect(contactId, tagId);
      }
    }
  }
}

function getActiveJobContactIds() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var jobSheet = ss.getSheetByName('Broadcast_Job');
  var jData = jobSheet.getDataRange().getValues();
  var activeJobs = {};

  for (var j = 1; j < jData.length; j++) {
    if (String(jData[j][8]) === 'berjalan') {
      activeJobs[String(jData[j][0])] = true;
    }
  }

  var kirimSheet = ss.getSheetByName('Broadcast_Kirim');
  var kData = kirimSheet.getDataRange().getValues();
  var activeContacts = {};

  for (var k = 1; k < kData.length; k++) {
    var jobId = String(kData[k][1]);
    var status = String(kData[k][5]);
    if (activeJobs[jobId] && (status === 'antre' || status === 'diproses')) {
      activeContacts[String(kData[k][2])] = true;
    }
  }
  return activeContacts;
}

function isJobInCampaign(jobId, campaignId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var jobSheet = ss.getSheetByName('Broadcast_Job');
  var data = jobSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(jobId) && String(data[i][1]) === String(campaignId)) {
      return true;
    }
  }
  return false;
}

function getKontakTagMap() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var tagSheet = ss.getSheetByName('Tag');
  var tData = tagSheet.getDataRange().getValues();
  var tagsById = {};
  for (var t = 1; t < tData.length; t++) {
    tagsById[String(tData[t][0])] = String(tData[t][1]);
  }

  var ktSheet = ss.getSheetByName('Kontak_Tag');
  var ktData = ktSheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < ktData.length; i++) {
    var cId = String(ktData[i][0]);
    var tgId = String(ktData[i][1]);
    if (!map[cId]) map[cId] = [];
    map[cId].push({ id: tgId, nama: tagsById[tgId] || tgId });
  }
  return map;
}

function getTagUsageCounts() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var ktSheet = ss.getSheetByName('Kontak_Tag');
  var ktData = ktSheet.getDataRange().getValues();
  var counts = {};
  for (var i = 1; i < ktData.length; i++) {
    var tgId = String(ktData[i][1]);
    counts[tgId] = (counts[tgId] || 0) + 1;
  }
  return counts;
}

function getSenderLabelMap() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Pengirim');
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    map[String(data[i][0])] = String(data[i][1]) + ' (' + String(data[i][2]) + ')';
  }
  return map;
}

function getSenderLabelById(id) {
  var map = getSenderLabelMap();
  return map[id] || id;
}

function syncKontakTags(contactId, newTagIds) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Kontak_Tag');
  var data = sheet.getDataRange().getValues();
  
  // Hapus tag lama
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(contactId)) {
      sheet.deleteRow(i + 1);
    }
  }
  // Tambah tag baru via rule engine
  newTagIds.forEach(function(tid) {
    applyTagToContact(contactId, tid);
  });
}

function removeKontakTags(contactId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Kontak_Tag');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(contactId)) {
      sheet.deleteRow(i + 1);
    }
  }
}

function getPengaturan() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Pengaturan');
  var data = sheet.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < data.length; i++) {
    out[String(data[i][0])] = data[i][1];
  }
  return { ok: true, data: serializeForClient(out) };
}

function getSettingValue(key, fallback) {
  var p = getPengaturan();
  if (p && p.ok && p.data && p.data[key] !== undefined) {
    return p.data[key];
  }
  return fallback;
}

function savePengaturan(settingsObj) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Pengaturan');
  var data = sheet.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < data.length; i++) {
    existing[String(data[i][0])] = i + 1;
  }

  for (var k in settingsObj) {
    if (existing[k]) {
      sheet.getRange(existing[k], 2).setValue(settingsObj[k]);
    } else {
      sheet.appendRow([k, settingsObj[k]]);
    }
  }
  writeLog('PENGATURAN', 'CONFIG', 'Memperbarui pengaturan sistem');
  return { ok: true };
}

function writeLog(aksi, idData, detail) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Log');
    sheet.appendRow([new Date().toISOString(), aksi, idData, detail]);
  } catch (e) {
    // ignore
  }
}

function getLogs(limit) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Log');
    var data = sheet.getDataRange().getValues();
    var list = [];
    limit = limit || 50;
    for (var i = data.length - 1; i >= 1 && list.length < limit; i--) {
      list.push({
        waktu: String(data[i][0]),
        aksi: String(data[i][1]),
        id_data: String(data[i][2]),
        detail: String(data[i][3])
      });
    }
    return { ok: true, data: serializeForClient(list) };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}
