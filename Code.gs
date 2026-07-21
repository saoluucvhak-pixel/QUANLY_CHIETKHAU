const SHEETS = {
  THUONGHIEU: 'DM_THUONGHIEU', NHOMKH: 'DM_NHOMKH', LOAISP: 'DM_LOAISP',
  DACTINH: 'DM_DACTINH', CONGDUNG: 'DM_CONGDUNG', QUYCACH: 'DM_QUYCACH',
  MHCK: 'DM_MHCK', GIACONGBO: 'DM_GIACONGBO',
  PROGRAMS: 'CHUONGTRINHCHIETKHAU', PURCHASES: 'SO_CHI_TIET_MUA_HANG',
  LUYKE: 'CHUONGTRINHCHIETKHAU_LUYKE', REPORT_CKTH: 'REPORT_CKTH', REPORT_CKCT: 'REPORT_CKCT',
  // THEO MÔ HÌNH 3 TẦNG MỚI: Tầng 1 (Danh mục chương trình) + Tầng 3 (Bậc điều kiện con của mỗi mã
  // chiết khấu). Tầng 2 (mã chiết khấu) vẫn dùng đúng sheet PROGRAMS cũ, chỉ bổ sung thêm cột liên kết.
  CHUONGTRINH: 'DM_CHUONGTRINH', DIEUKIEN: 'DM_DIEUKIEN'
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Quản lý Chiết khấu NCC')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ===== CHẨN ĐOÁN KẾT NỐI: cho biết chính xác script đang đọc file/sheet nào, để phát hiện
// trường hợp Apps Script không gắn đúng vào file Sheet có dữ liệu, hoặc tên tab trong Sheet không
// khớp với tên script đang tìm (2 nguyên nhân phổ biến nhất khiến app hiện "0 dữ liệu" dù Sheet có data). =====
function getDiagnosticInfo() {
  const ss = SpreadsheetApp.getActive();
  const info = {
    spreadsheetName: ss.getName(),
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    actualSheetNames: ss.getSheets().map(s => s.getName()),
    expected: []
  };
  Object.keys(SHEETS).forEach(key => {
    const name = SHEETS[key];
    const sh = ss.getSheetByName(name);
    info.expected.push({
      key: key,
      expectedName: name,
      found: !!sh,
      rowCount: sh ? Math.max(0, sh.getLastRow() - 1) : 0
    });
  });
  return info;
}

// ===== Quản lý phiên bản dữ liệu (để Frontend cache, tránh load lại Sheet khi không cần) =====
// Dùng THỜI ĐIỂM SỬA ĐỔI CUỐI CÙNG thật của chính file Google Sheet (lấy từ Drive) làm "phiên bản".
// Cách này phản ánh ĐÚNG mọi thay đổi — kể cả khi người dùng gõ tay trực tiếp vào Sheet (không qua app)
// — khác với cách đếm số cũ (chỉ tăng khi bấm Lưu trong app) khiến sửa tay trực tiếp trên Sheet không
// được app nhận ra, dẫn đến app tiếp tục dùng bản nhớ tạm (cache) cũ dù Sheet đã có dữ liệu mới.
function getDataVersion() {
  try {
    const ss = SpreadsheetApp.getActive();
    const file = DriveApp.getFileById(ss.getId());
    return String(file.getLastUpdated().getTime());
  } catch (e) {
    // Phòng trường hợp thiếu quyền Drive -> fallback về cách đếm cũ (vẫn còn tác dụng khi lưu qua app)
    const props = PropertiesService.getScriptProperties();
    return props.getProperty('DATA_VERSION') || '0';
  }
}
function bumpDataVersion_() {
  // SpreadsheetApp có thể trì hoãn ghi tới cuối lần thực thi; flush() để đảm bảo ghi xong ngay,
  // rồi đọc lại "lần sửa cuối" mới nhất từ Drive để trả về đúng phiên bản sau khi lưu.
  try { SpreadsheetApp.flush(); } catch (e) {}
  return getDataVersion();
}

// Chuyển 1 giá trị ô về dạng an toàn để gửi qua google.script.run
// (Sheet có thể trả về đối tượng Date "invalid" khi ô định dạng ngày nhưng để trống -> gây lỗi
// "Cannot return an invalid date" khi Apps Script serialize JSON để gửi về trình duyệt)
function sanitizeCellValue_(v) {
  // SỬA LỖI: đối tượng Date lồng trong mảng lớn có thể bị lỗi/hỏng khi truyền qua google.script.run
  // (lỗi đã biết của Apps Script). Nay luôn chuyển Date thành CHUỖI TEXT (yyyy-MM-dd) trước khi gửi,
  // áp dụng thống nhất ở đây cho mọi hàm đọc dữ liệu — tránh hoàn toàn kiểu Date trong gói tin RPC.
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd');
  }
  return v;
}

// Hàm lấy dữ liệu an toàn
// SỬA LỖI (ổn định hoá tải dữ liệu): trước đây loadAllData() đọc TOÀN BỘ 13 sheet (hơn 3.300 dòng,
// trong đó riêng DM_GIACONGBO + SO_CHI_TIET_MUA_HANG đã hơn 2.800 dòng) trong DUY NHẤT 1 lượt gọi
// google.script.run. Với dữ liệu lớn, 1 lượt gọi to như vậy đôi khi bị timeout/lỗi giữa chừng ở phía
// máy chủ Google (không báo lỗi rõ ràng về client) và trả về kết quả rỗng dù dữ liệu vẫn còn nguyên
// trên Sheet — đúng hiện tượng "chập chờn" quan sát được (chẩn đoán ra đủ dữ liệu, làm mới lại ra 0).
// Nay tách thành 3 lượt gọi NHỎ hơn để giảm rủi ro timeout, và cho phép Frontend tự thử lại riêng
// từng phần nếu phần đó bị rỗng bất thường, thay vì phải tải lại toàn bộ.
function readSheetRows_(sheetName) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(r => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = sanitizeCellValue_(r[i]));
    return obj;
  });
}
function loadCoreData() {
  const result = { catalog: {}, programs: [], programsLuyKe: [], reportCKTH: [], reportCKCT: [], chuongtrinh: [], dieukien: [] };
  ['THUONGHIEU','NHOMKH','LOAISP','DACTINH','CONGDUNG','QUYCACH','MHCK'].forEach(key => {
    result.catalog[key.toLowerCase()] = readSheetRows_(SHEETS[key]);
  });
  result.programs = readSheetRows_(SHEETS.PROGRAMS);
  result.programsLuyKe = readSheetRows_(SHEETS.LUYKE);
  result.reportCKTH = readSheetRows_(SHEETS.REPORT_CKTH);
  result.reportCKCT = readSheetRows_(SHEETS.REPORT_CKCT);
  result.chuongtrinh = readSheetRows_(SHEETS.CHUONGTRINH);
  result.dieukien = readSheetRows_(SHEETS.DIEUKIEN);
  result.version = getDataVersion();
  return result;
}

// SỬA LỖI (chia nhỏ gói truyền): dù đã dùng định dạng gọn, gói 1.435 dòng vẫn có thể bị hỏng/rỗng khi
// truyền qua kênh giao tiếp iframe của Apps Script (hiện tượng đã biết, không báo lỗi rõ ràng). Nay
// cho phép tải THEO TỪNG GÓI NHỎ (offset/limit) — Frontend gọi nhiều lần liên tiếp, mỗi lần chỉ vài
// trăm dòng, ghép lại — đảm bảo mỗi gói luôn đủ nhỏ để truyền an toàn.
function getSheetRowCount_(sheetKey) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEETS[sheetKey]);
  if (!sh) sh = ss.insertSheet(SHEETS[sheetKey]);
  return { total: Math.max(0, sh.getLastRow() - 1), version: getDataVersion() };
}
function getPurchasesRowCount() { return getSheetRowCount_('PURCHASES'); }
function getGiaCongBoRowCount() { return getSheetRowCount_('GIACONGBO'); }
function readSheetChunk_(sheetName, offset, limit) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { headers: [], rows: [] };
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const startRow = 2 + offset; // dòng 1 là header, dữ liệu bắt đầu dòng 2
  if (startRow > lastRow) return { headers: headers, rows: [] };
  const numRows = Math.min(limit, lastRow - startRow + 1);
  const data = sh.getRange(startRow, 1, numRows, lastCol).getValues();
  const rows = data
    .filter(r => r.some(c => String(c).trim() !== ''))
    .map(r => r.map(c => sanitizeCellValue_(c)));
  return { headers: headers, rows: rows };
}
function loadPurchasesChunk(offset, limit) { return readSheetChunk_(SHEETS.PURCHASES, offset, limit); }
function loadGiaCongBoChunk(offset, limit) { return readSheetChunk_(SHEETS.GIACONGBO, offset, limit); }
// SỬA LỖI (giảm dung lượng truyền): readSheetRows_() trả về mảng OBJECT đầy đủ — với sheet lớn
// (1.400+ dòng x 16 cột), tên cột bị LẶP LẠI trong JSON ở MỖI dòng, làm gói tin nặng hơn nhiều so với
// cần thiết, dễ bị ngắt/timeout trên đường truyền chậm. Nay dùng định dạng GỌN: gửi tên cột 1 LẦN DUY
// NHẤT (headers), dữ liệu chỉ là mảng giá trị thuần (rows) — Frontend tự ráp lại thành object.
function loadAllData() {
  const core = loadCoreData();
  core.catalog.giacongbo = readSheetRows_(SHEETS.GIACONGBO);
  core.purchases = readSheetRows_(SHEETS.PURCHASES);
  return core;
}

// Hàm lưu dữ liệu tổng quát
function saveSheetData(sheetName, data) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  // SỬA LỖI NGHIÊM TRỌNG: trước đây hàm này LUÔN xóa sạch (clearContents) tab rồi ghi lại, kể cả khi
  // "data" truyền vào là mảng RỖNG. saveCatalogData() gửi lên TOÀN BỘ object catalog (mọi danh mục)
  // mỗi khi lưu — kể cả khi người dùng chỉ vừa import 1 loại danh mục (VD chỉ Giá công bố). Nếu tại
  // thời điểm đó trình duyệt chưa tải đầy đủ các danh mục KHÁC (VD Thương hiệu, Nhóm KH...) từ Sheet
  // lên bộ nhớ tạm, các mảng đó sẽ RỖNG trong bộ nhớ, khiến lệnh lưu XÓA SẠCH dữ liệu thật đang có
  // trên các tab đó dù người dùng không hề có ý định xóa. Nay: nếu dữ liệu gửi lên rỗng NHƯNG tab đích
  // đang có sẵn dữ liệu thật, KHÔNG xóa (bỏ qua, giữ nguyên dữ liệu cũ) — để tránh mất dữ liệu ngoài ý
  // muốn; chỉ cho phép xóa sạch khi tab đích hiện đang trống hoặc dữ liệu gửi lên cũng có nội dung.
  const existingRowCount = Math.max(0, sh.getLastRow() - 1);
  if ((!data || data.length === 0) && existingRowCount > 0) {
    return; // Bỏ qua để bảo vệ dữ liệu đã có, không ghi đè bằng mảng rỗng
  }
  sh.clearContents();
  if (data && data.length > 0) {
    const headers = Object.keys(data[0]);
    sh.appendRow(headers);
    const values = data.map(r => headers.map(h => r[h] ?? ''));
    sh.getRange(2, 1, values.length, headers.length).setValues(values);
  }
}

// Các hàm cầu nối gọi từ Frontend
// Mỗi hàm save trả về phiên bản dữ liệu mới nhất để Frontend cập nhật cache cục bộ.
function saveCatalogData(catalog) {
  const ss = SpreadsheetApp.getActive();
  
  // Lưu bảng ánh xạ mã hàng (MHCK)
  saveSheetData(SHEETS.MHCK, catalog.mhck);
  
  // Lưu tất cả các bảng danh mục phụ (thuonghieu, nhomkh, loaisp, dactinh, congdung, quycach)
  Object.keys(catalog).forEach(key => {
    if (key !== 'mhck') {
      const sheetName = SHEETS[key.toUpperCase()];
      if (sheetName) {
        saveSheetData(sheetName, catalog[key]);
      }
    }
  });
  return bumpDataVersion_();
}
function saveProgramsData(data) { saveSheetData(SHEETS.PROGRAMS, data); return bumpDataVersion_(); }
function savePurchasesData(data) { saveSheetData(SHEETS.PURCHASES, data); return bumpDataVersion_(); }
function saveLuyKeData(data) { saveSheetData(SHEETS.LUYKE, data); return bumpDataVersion_(); }
// THEO MÔ HÌNH 3 TẦNG: Tầng 1 (Danh mục chương trình) và Tầng 3 (Bậc điều kiện con của mã chiết khấu)
function saveChuongTrinhData(data) { saveSheetData(SHEETS.CHUONGTRINH, data); return bumpDataVersion_(); }
function saveDieuKienData(data) { saveSheetData(SHEETS.DIEUKIEN, data); return bumpDataVersion_(); }

// Lưu báo cáo chiết khấu (REPORT_CKTH / REPORT_CKCT) theo kiểu CỘNG DỒN + CẬP NHẬT ĐÈ nếu trùng khóa,
// KHÔNG xóa dữ liệu các kỳ đã lưu trước đó — vì "Tính chiết khấu bổ sung" cần dữ liệu nhiều tháng cộng dồn
// (VD lưu báo cáo tháng 1, rồi tháng 2, rồi tháng 3 — cả 3 bản ghi phải còn nguyên trong Sheet).
// Chuẩn hóa 1 giá trị dùng làm khóa so khớp: Google Sheets có thể TỰ ĐỘNG chuyển chuỗi dạng
// "2026-01-01" thành kiểu Date thật khi lưu; lần đọc lại sau đó giá trị trở thành đối tượng Date
// (không còn là chuỗi text gốc) khiến khóa so khớp bị lệch định dạng -> tưởng là bản ghi MỚI dù
// thực chất là CÙNG 1 kỳ đã lưu trước đó -> sinh dòng trùng thay vì ghi đè. Nay ép mọi giá trị Date
// về đúng dạng chuỗi yyyy-MM-dd (theo giờ địa phương) trước khi so khớp, đảm bảo luôn khớp đúng.
function normKeyVal_(v) {
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return String(v == null ? '' : v).trim();
}
function appendOrUpsertRows_(sheetName, rows, keyFields) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  const data = sh.getDataRange().getValues();
  let headers, existingRows;
  if (data.length > 0 && data[0].length > 0 && String(data[0][0]) !== '') {
    headers = data[0];
    existingRows = data.slice(1).map(r => {
      let o = {};
      headers.forEach((h, i) => o[h] = sanitizeCellValue_(r[i]));
      return o;
    });
  } else {
    headers = rows.length ? Object.keys(rows[0]) : [];
    existingRows = [];
  }
  // Đảm bảo đủ cột nếu dữ liệu mới có thêm trường chưa từng lưu trước đây
  if (rows.length) {
    Object.keys(rows[0]).forEach(k => { if (headers.indexOf(k) === -1) headers.push(k); });
  }
  const keyOf = (r) => keyFields.map(k => normKeyVal_(r[k])).join('||');
  const idx = {};
  existingRows.forEach((r, i) => { idx[keyOf(r)] = i; });
  rows.forEach(r => {
    const k = keyOf(r);
    if (idx.hasOwnProperty(k)) existingRows[idx[k]] = Object.assign({}, existingRows[idx[k]], r);
    else { existingRows.push(r); idx[k] = existingRows.length - 1; }
  });
  // Chỉ chuẩn hóa CÁC CỘT KHÓA (keyFields) về chuỗi cố định trước khi ghi lại — các cột số liệu khác
  // (sản lượng, doanh số, tiền chiết khấu...) vẫn giữ nguyên kiểu dữ liệu gốc để không bị lỗi định dạng
  // số trong Google Sheet. Việc so khớp key ở lần lưu SAU vẫn luôn đúng vì keyOf() đã tự chuẩn hóa lại
  // giá trị đọc lên (kể cả khi Sheets tự ý chuyển ô ngày thành kiểu Date).
  existingRows.forEach(r => { keyFields.forEach(k => { r[k] = normKeyVal_(r[k]); }); });
  sh.clearContents();
  if (headers.length) {
    sh.appendRow(headers);
    if (existingRows.length) {
      const values = existingRows.map(r => headers.map(h => r[h] ?? ''));
      sh.getRange(2, 1, values.length, headers.length).setValues(values);
    }
  }
}
function saveReportCKTH(rows) {
  appendOrUpsertRows_(SHEETS.REPORT_CKTH, rows, ['program_id', 'tu_ngay', 'den_ngay']);
  return bumpDataVersion_();
}
function saveReportCKCT(rows) {
  appendOrUpsertRows_(SHEETS.REPORT_CKCT, rows, ['program_id', 'tu_ngay', 'den_ngay', 'mahang']);
  return bumpDataVersion_();
}

// ===== TEST KẾT NỐI TỐI GIẢN: không đọc/ghi Sheet, chỉ trả về thời gian máy chủ. Dùng để tách bạch
// xem vấn đề nằm ở kết nối cơ bản (mạng/hạ tầng) hay ở việc đọc dữ liệu Sheet. =====
function ping() {
  return { ok: true, serverTime: new Date().toString() };
}
// THEO GÓP Ý: lấy email người dùng hiện tại để ghi vết "người sửa gần nhất" cho Chương trình chiết
// khấu và Giá công bố. LƯU Ý: chỉ trả về đúng email nếu Web App được deploy với "Execute as: User
// accessing the web app" — nếu deploy "Execute as: Me" (mặc định phổ biến), hàm này trả về rỗng vì
// không có quyền biết danh tính người dùng cuối; khi đó Frontend sẽ tự hiện "Không xác định".
function getCurrentUserEmail() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email || '';
  } catch (e) {
    return '';
  }
}

// =====================================================================================
// HÀM KHÔI PHỤC DỮ LIỆU (CHẠY 1 LẦN DUY NHẤT, THỦ CÔNG TỪ APPS SCRIPT EDITOR)
// -----------------------------------------------------------------------------------
// Bối cảnh: một số tab trong Google Sheet hiện đang chứa dữ liệu THÔ (dán trực tiếp từ file
// Excel gốc / phần mềm kế toán, tiêu đề cột tiếng Việt) hoặc nằm SAI TÊN TAB so với những gì
// SHEETS{} ở đầu file này yêu cầu (VD: dữ liệu Chương trình chiết khấu nằm ở tab khác thay vì
// đúng tab "CHUONGTRINHCHIETKHAU"). Vì loadAllData() chỉ đọc đúng tên tab + đúng tên cột nội bộ
// (mamh_ncc, gia_congbo, dim_thuonghieu...), các tab sai định dạng này KHÔNG BAO GIỜ được đọc lên
// Webapp, dù dữ liệu thật vẫn còn nguyên trong Sheet.
//
// Hàm dưới đây quét toàn bộ tab hiện có trong file, TỰ NHẬN DIỆN loại dữ liệu qua tiêu đề cột,
// CHUYỂN ĐỔI sang đúng cấu trúc nội bộ, rồi ghi vào đúng tab mà loadAllData() cần đọc.
// Tab gốc KHÔNG bị xóa — chỉ đổi tên thêm hậu tố "_BACKUP_<timestamp>" để giữ lại đối chiếu.
//
// CÁCH CHẠY: mở Apps Script editor -> chọn hàm "migrateLegacyData" ở dropdown trên cùng -> bấm ▶ Run.
// Xem kết quả (bao nhiêu dòng đã chuyển cho từng loại) ở View -> Logs (hoặc Executions).
// =====================================================================================
function migrateLegacyData() {
  const ss = SpreadsheetApp.getActive();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyyyMMdd_HHmmss');
  const report = [];

  // Đọc toàn bộ 1 tab thành mảng object {header: value}, dò tiêu đề bất kể viết hoa/thường/khoảng trắng
  function readSheetAsObjects_(sheet) {
    const data = sheet.getDataRange().getValues();
    if (data.length < 1) return { headers: [], rows: [] };
    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1)
      .filter(r => r.some(c => String(c).trim() !== ''))
      .map(r => { const o = {}; headers.forEach((h, i) => o[h] = sanitizeCellValue_(r[i])); return o; });
    return { headers, rows };
  }
  // Lấy giá trị theo danh sách tên cột khả dĩ (không phân biệt hoa/thường, khoảng trắng, gạch dưới)
  function pick_(row, candidates) {
    const normalize = s => String(s).toLowerCase().replace(/[\s_]/g, '');
    const keys = Object.keys(row);
    for (const c of candidates) {
      const k = keys.find(k => normalize(k) === normalize(c));
      if (k != null && row[k] !== '' && row[k] != null) return row[k];
    }
    return '';
  }
  function excelDateToISO_(v) {
    if (v === '' || v == null) return '';
    if (v instanceof Date) {
      if (isNaN(v.getTime())) return '';
      return Utilities.formatDate(v, Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd');
    }
    // dd/mm/yyyy
    const m = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
    return String(v);
  }
  function cleanNumber_(v) {
    if (v === '' || v == null) return 0;
    let s = String(v).replace(/[^0-9.,-]/g, '').trim();
    if (!s) return 0;
    if (s.includes('.') && s.includes(',')) {
      s = s.lastIndexOf('.') > s.lastIndexOf(',') ? s.replace(/,/g, '') : s.replace(/\./g, '').replace(',', '.');
    } else if (s.includes(',')) {
      const parts = s.split(',');
      s = (parts[1] && parts[1].length === 3) ? s.replace(/,/g, '') : s.replace(',', '.');
    }
    return Number(s) || 0;
  }
  function writeTarget_(sheetName, objRows) {
    if (!objRows.length) return;
    let sh = ss.getSheetByName(sheetName);
    if (!sh) sh = ss.insertSheet(sheetName);
    const existing = readSheetAsObjects_(sh).rows;
    const merged = existing.concat(objRows); // cộng dồn, không đè mất dữ liệu đã có sẵn đúng chuẩn
    const headers = Object.keys(merged[0]);
    merged.forEach(r => Object.keys(r).forEach(k => { if (headers.indexOf(k) === -1) headers.push(k); }));
    sh.clearContents();
    sh.appendRow(headers);
    sh.getRange(2, 1, merged.length, headers.length).setValues(merged.map(r => headers.map(h => r[h] ?? '')));
  }
  function backupOriginal_(sheet) {
    try { sheet.setName(sheet.getName() + '_BACKUP_' + stamp); } catch (e) {}
  }

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name.indexOf('_BACKUP_') > -1) return;

    const { headers, rows } = readSheetAsObjects_(sheet);
    if (!rows.length) return;
    const first = rows[0];

    // LỖI ĐÃ SỬA: trước đây bỏ qua MỌI tab có TÊN trùng với tên tab đích (VD "DM_GIACONGBO",
    // "SO_CHI_TIET_MUA_HANG") dù nội dung bên trong vẫn là tiêu đề THÔ tiếng Việt chưa chuyển đổi
    // -> 2 tab quan trọng nhất bị bỏ qua hoàn toàn, không tab nào được xử lý. Nay chỉ bỏ qua khi
    // tab đó THẬT SỰ đã đúng định dạng nội bộ (có đúng tên cột app cần), bất kể tên tab là gì.
    const alreadyCorrectFormat =
      (name === SHEETS.GIACONGBO && pick_(first, ['mamh_ncc']) !== '') ||
      (name === SHEETS.PURCHASES && pick_(first, ['mahang']) !== '' && pick_(first, ['ngay']) !== '') ||
      (name === SHEETS.PROGRAMS && pick_(first, ['ma_chietkhau']) !== '' && pick_(first, ['dim_thuonghieu']) !== '') ||
      (Object.keys(SHEETS).some(k => SHEETS[k] === name && name.indexOf('DM_') === 0) && pick_(first, ['ma']) !== '' && pick_(first, ['ten']) !== '');
    if (alreadyCorrectFormat) return;

    // 1) Danh mục Thương hiệu (Ma_ThuongHieu, Tên thương hiệu)
    if (pick_(first, ['Ma_ThuongHieu', 'Ma ThuongHieu']) !== '') {
      const out = rows.map(r => ({ ma: String(pick_(r, ['Ma_ThuongHieu', 'Ma ThuongHieu'])).trim(), ten: String(pick_(r, ['Tên thương hiệu']) || '').trim() })).filter(r => r.ma);
      backupOriginal_(sheet);
      writeTarget_(SHEETS.THUONGHIEU, out);
      report.push(name + ' -> ' + SHEETS.THUONGHIEU + ' (' + out.length + ' dòng)');
      return;
    }
    // 2) Giá công bố NCC (Mã hàng, Tên hàng, Giá công bố, Hiệu lực từ/đến ngày)
    if (pick_(first, ['Giá công bố', 'GIA_CONGBO', 'Đơn giá công bố']) !== '') {
      const out = rows.map(r => ({
        mamh_ncc: String(pick_(r, ['Mã MH', 'MaMH', 'Mã hàng'])).trim(),
        tenhang: String(pick_(r, ['Tên hàng']) || '').trim(),
        gia_congbo: cleanNumber_(pick_(r, ['Giá công bố', 'GIA_CONGBO', 'Đơn giá công bố'])),
        hieuluctu: excelDateToISO_(pick_(r, ['Hiệu lực từ ngày', 'Ngày hiệu lực'])),
        hieulucden: excelDateToISO_(pick_(r, ['Hiệu lực đến ngày']))
      })).filter(r => r.mamh_ncc);
      backupOriginal_(sheet);
      writeTarget_(SHEETS.GIACONGBO, out);
      report.push(name + ' -> ' + SHEETS.GIACONGBO + ' (' + out.length + ' dòng)');
      return;
    }
    // 3) Sổ chi tiết mua hàng (Ngày chứng từ, Số hóa đơn, Mã hàng, Số lượng mua...)
    if (pick_(first, ['Ngày chứng từ', 'Ngày CT']) !== '' && pick_(first, ['Mã hàng', 'Mã MH']) !== '') {
      const out = rows.map(r => {
        const ngay = excelDateToISO_(pick_(r, ['Ngày chứng từ', 'Ngày CT', 'Ngày']));
        const sohd = pick_(r, ['Số hóa đơn', 'Số HĐ']);
        const mahang = String(pick_(r, ['Mã hàng', 'Mã MH'])).trim();
        const soluong = cleanNumber_(pick_(r, ['Số lượng mua', 'Số lượng', 'SL']));
        const giatri = cleanNumber_(pick_(r, ['Giá trị mua', 'Giá trị', 'Thành tiền']));
        return {
          id: [ngay, sohd, mahang, soluong, giatri].join('|'),
          ngay, sohd,
          diengiai: pick_(r, ['Diễn giải']),
          makho: pick_(r, ['Mã kho']),
          mahang,
          tenhang: pick_(r, ['Tên hàng']),
          dvt: pick_(r, ['ĐVT']),
          soluong, dongia: cleanNumber_(pick_(r, ['Đơn giá', 'Đơn giá mua', 'DG'])), giatri,
          nhacungcap: pick_(r, ['Tên nhà cung cấp', 'Nhà cung cấp']),
          diachi: pick_(r, ['Địa chỉ']),
          chietkhau_goc: cleanNumber_(pick_(r, ['Chiết khấu'])),
          tkno: pick_(r, ['TK Nợ']), tkco: pick_(r, ['TK Có']), mathongke: pick_(r, ['Mã thống kê'])
        };
      }).filter(r => r.mahang && r.ngay);
      backupOriginal_(sheet);
      writeTarget_(SHEETS.PURCHASES, out);
      report.push(name + ' -> ' + SHEETS.PURCHASES + ' (' + out.length + ' dòng)');
      return;
    }
    // 4) Chương trình chiết khấu - định dạng THÔ theo bậc thang (Ma_Chietkhau, SL_MIN, TL_CK...)
    if (pick_(first, ['Ma_Chietkhau']) !== '' && pick_(first, ['dim_thuonghieu']) === '') {
      const DIM_KEYS_ = ['thuonghieu', 'nhomkh', 'loaisp', 'dactinh', 'congdung', 'quycach'];
      function guessHT_(t) {
        t = String(t || '').toUpperCase();
        if (t.indexOf('LŨY KẾ') > -1 || t.indexOf('LUY KE') > -1) return 'LUY_KE';
        if (t.indexOf('CÔNG BỐ') > -1 || t.indexOf('CONG BO') > -1) return 'PERCENT_GIACONGBO';
        if (t.indexOf('CỐ ĐỊNH') > -1 || t.indexOf('CO DINH') > -1) return 'DONGIA_CODINH';
        return 'PERCENT_GIAMUA';
      }
      function guessTTCK_(t) {
        t = String(t || '').toUpperCase();
        if (t.indexOf('GIÁ BÁN') > -1 || t.indexOf('GIA BAN') > -1) return 'GIAM_GIABAN';
        return 'GIAM_HOADON';
      }
      const out = rows.map(r => {
        const maCK = String(pick_(r, ['Ma_Chietkhau']) || '').trim();
        const parts = maCK.split('_');
        const dims = {};
        DIM_KEYS_.forEach((k, i) => { dims['dim_' + k] = parts[i] || 'NO'; });
        let tl = pick_(r, ['TL_CK']);
        tl = tl === '' ? '' : (Number(tl) <= 1 ? Number(tl) * 100 : Number(tl));
        return Object.assign({
          id: Utilities.getUuid(),
          ngaytao: excelDateToISO_(pick_(r, ['Ngày tạo'])) || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd'),
          hieuluctu: excelDateToISO_(pick_(r, ['Hiệu lực từ ngày'])),
          hieulucden: excelDateToISO_(pick_(r, ['Hiệu lực đến ngày'])),
          ma_chietkhau: maCK,
          sl_min: pick_(r, ['SL_MIN']), sl_max: pick_(r, ['SL_MAX']),
          dt_min: pick_(r, ['DT_MIN']), dt_max: pick_(r, ['DT_MAX']),
          thoidiem_ck: pick_(r, ['THOIDIEM_CK']) || 'Quý',
          ht_chietkhau: guessHT_(pick_(r, ['HT_CHIETKHAU'])),
          tt_ck: guessTTCK_(pick_(r, ['TT_CK'])),
          dt_kehoach: pick_(r, ['DT_KEHOACH']), tl_ck: tl,
          dongia_codinh: pick_(r, ['DONGIA_CODINH', 'DONGIA_APDUNG']),
          pct_kh_min: pick_(r, ['PCT_KH_MIN', '%KH Min']),
          pct_kh_max: pick_(r, ['PCT_KH_MAX', '%KH Max']),
          dieukien_sldt: String(pick_(r, ['DIEUKIEN_SLDT', 'Điều kiện SL/DT'])).toUpperCase() === 'OR' ? 'OR' : 'AND',
          ghichu: pick_(r, ['ID_CHIETKHAU']) || ''
        }, dims);
      }).filter(r => r.ma_chietkhau);
      backupOriginal_(sheet);
      writeTarget_(SHEETS.PROGRAMS, out);
      report.push(name + ' -> ' + SHEETS.PROGRAMS + ' (' + out.length + ' dòng, định dạng bậc thang)');
      return;
    }
    // 5) Chương trình chiết khấu - ĐÃ ĐÚNG cấu trúc nội bộ (ma_chietkhau, dim_thuonghieu, id...)
    //    chỉ SAI TÊN TAB -> chỉ cần chuyển nguyên trạng sang đúng tab.
    if (pick_(first, ['ma_chietkhau']) !== '' && pick_(first, ['dim_thuonghieu']) !== '') {
      backupOriginal_(sheet);
      writeTarget_(SHEETS.PROGRAMS, rows);
      report.push(name + ' -> ' + SHEETS.PROGRAMS + ' (' + rows.length + ' dòng, đã đúng cấu trúc, chỉ đổi tên tab)');
      return;
    }
  });

  bumpDataVersion_();
  const summary = report.length ? report.join('\n') : '(Không tìm thấy tab dữ liệu thô nào cần chuyển đổi)';
  Logger.log(summary);
  return summary;
}
