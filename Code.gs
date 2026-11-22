/**
 * CBH Go - Full Admin Backend (Google Apps Script)
 * รองรับการเก็บข้อมูล:
 * 1. Clinics (รายชื่อหน่วยงาน/คลินิก/หอผู้ป่วย)
 * 2. Buildings (พิกัด ขนาด สี และข้อมูลอาคาร 3D)
 * 3. Nodes & Edges (จุดทางแยกและเส้นทางเดิน/รถ One-Way, Two-Way)
 * 4. QR_Anchors & Search_Logs (สถิติการสแกน QR และค้นหา)
 */

const SHEETS = {
  CONFIG: 'System_Config',
  CLINICS: 'Clinics',
  QR: 'QR_Anchors',
  LOGS: 'Search_Logs'
};

// รันฟังก์ชันนี้ 1 ครั้งเพื่อสร้างฐานข้อมูลใน Google Sheets
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. System_Config (เก็บ JSON ของ Buildings, Nodes, Edges เพื่อให้โหลดเร็วและซิงค์ 3D ได้ 100%)
  let cfgSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet(SHEETS.CONFIG);
    cfgSheet.appendRow(['key', 'json_value', 'updated_at', 'updated_by']);
    cfgSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    cfgSheet.appendRow(['map_layout_json', '', new Date(), 'system_init']);
  }

  // 2. Clinics (ตารางแก้ไขคลินิกแบบอ่านง่ายใน Sheets หรือผ่านหน้าเว็บ Admin)
  let clinicSheet = ss.getSheetByName(SHEETS.CLINICS);
  if (!clinicSheet) {
    clinicSheet = ss.insertSheet(SHEETS.CLINICS);
    clinicSheet.appendRow(['id', 'name', 'building_id', 'floor', 'open_hours', 'phone', 'keywords', 'status']);
    clinicSheet.getRange('A1:H1').setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');
  }

  // 3. QR_Anchors
  let qrSheet = ss.getSheetByName(SHEETS.QR);
  if (!qrSheet) {
    qrSheet = ss.insertSheet(SHEETS.QR);
    qrSheet.appendRow(['qr_code_id', 'location_name', 'node_id', 'scan_count']);
    qrSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#059669').setFontColor('#ffffff');
    qrSheet.appendRow(['gate_main', 'ประตูทางเข้าหลัก ถนนสุขุมวิท', 'gate_main', 0]);
    qrSheet.appendRow(['skywalk_er', 'ทางลงสะพานลอย ตึกชลายุวัฒน์ (ER)', 'skywalk_er', 0]);
    qrSheet.appendRow(['parking7_entrance', 'อาคารจอดรถ 7 ชั้น', 'parking7_entrance', 0]);
    qrSheet.appendRow(['opd14_entrance', 'โถงตึก 14 ชั้น', 'opd14_entrance', 0]);
  }

  // 4. Search_Logs
  let logSheet = ss.getSheetByName(SHEETS.LOGS);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEETS.LOGS);
    logSheet.appendRow(['timestamp', 'start_node', 'destination_name', 'travel_mode']);
    logSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#475569').setFontColor('#ffffff');
  }
}

// GET: ดึงข้อมูลผังตึก 3D + เส้นทาง + รายชื่อคลินิกล่าสุด
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfgSheet = ss.getSheetByName(SHEETS.CONFIG);
  const clinicSheet = ss.getSheetByName(SHEETS.CLINICS);
  const qrSheet = ss.getSheetByName(SHEETS.QR);

  if (e && e.parameter && e.parameter.qr && qrSheet) {
    incrementQrScan(qrSheet, e.parameter.qr);
  }

  let mapLayout = null;
  if (cfgSheet && cfgSheet.getLastRow() >= 2) {
    const rawJson = cfgSheet.getRange(2, 2).getValue();
    if (rawJson) {
      try { mapLayout = JSON.parse(rawJson); } catch (err) {}
    }
  }

  const clinics = sheetToObjects(clinicSheet);

  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'success',
      updated_at: new Date().toISOString(),
      map_layout: mapLayout,
      clinics: clinics
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST: รองรับทั้งการบันทึกสถิติ (log) และการบันทึกผัง 3D จากหน้า Full Admin (save_layout)
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const payload = JSON.parse(e.postData.contents);

    if (payload.action === 'save_layout') {
      // ตรวจสอบรหัสผ่าน Admin ฝั่ง Server (เปลี่ยนได้ตามต้องการ)
      if (payload.admin_pin !== 'cbh2026') {
        return jsonResponse({ status: 'error', message: 'รหัสผ่านผู้ดูแลระบบไม่ถูกต้อง' });
      }

      let cfgSheet = ss.getSheetByName(SHEETS.CONFIG);
      if (!cfgSheet) { setupDatabase(); cfgSheet = ss.getSheetByName(SHEETS.CONFIG); }

      const layoutObj = {
        buildings: payload.buildings || [],
        nodes: payload.nodes || {},
        edges: payload.edges || []
      };
      cfgSheet.getRange(2, 1, 1, 4).setValues([
        ['map_layout_json', JSON.stringify(layoutObj), new Date(), payload.updated_by || 'admin']
      ]);

      // อัปเดตตาราง Clinics ให้ตรงกัน
      if (Array.isArray(payload.clinics)) {
        let clinicSheet = ss.getSheetByName(SHEETS.CLINICS);
        clinicSheet.clearContents();
        clinicSheet.appendRow(['id', 'name', 'building_id', 'floor', 'open_hours', 'phone', 'keywords', 'status']);
        clinicSheet.getRange('A1:H1').setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');
        if (payload.clinics.length > 0) {
          const rows = payload.clinics.map(c => [
            c.id, c.name, c.building_id, c.floor, c.open_hours || '08:00 - 16:00 น.',
            c.phone || '-', c.keywords || '', c.status || 'เปิดปกติ'
          ]);
          clinicSheet.getRange(2, 1, rows.length, 8).setValues(rows);
        }
      }

      return jsonResponse({ status: 'saved', message: 'บันทึกผังตึก เส้นทาง และคลินิกขึ้น Google Sheets เรียบร้อยแล้ว' });
    }

    // Default: บันทึกสถิติการค้นหาเส้นทาง
    const logSheet = ss.getSheetByName(SHEETS.LOGS);
    if (logSheet) {
      logSheet.appendRow([
        new Date(),
        payload.start_node || 'unknown',
        payload.destination_name || 'unknown',
        payload.travel_mode || 'walk'
      ]);
    }
    return jsonResponse({ status: 'logged' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function incrementQrScan(sheet, qrId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === qrId) {
      sheet.getRange(i + 1, 4).setValue((Number(data[i][3]) || 0) + 1);
      break;
    }
  }
}
