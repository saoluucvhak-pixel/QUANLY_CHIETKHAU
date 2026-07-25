const SHEETS = {
  THUONGHIEU: 'DM_THUONGHIEU', NHOMKH: 'DM_NHOMKH', LOAISP: 'DM_LOAISP',
  DACTINH: 'DM_DACTINH', CONGDUNG: 'DM_CONGDUNG', QUYCACH: 'DM_QUYCACH',
  MHCK: 'DM_MHCK', GIACONGBO: 'DM_GIACONGBO',
  PROGRAMS: 'CHUONGTRINHCHIETKHAU', PURCHASES: 'SO_CHI_TIET_MUA_HANG',
  LUYKE: 'CHUONGTRINHCHIETKHAU_LUYKE', REPORT_CKTH: 'REPORT_CKTH', REPORT_CKCT: 'REPORT_CKCT',
  // THEO MÔ HÌNH 3 TẦNG MỚI: Tầng 1 (Danh mục chương trình) + Tầng 3 (Bậc điều kiện con của mỗi mã
  // chiết khấu). Tầng 2 (mã chiết khấu) vẫn dùng đúng sheet PROGRAMS cũ, chỉ bổ sung thêm cột liên kết.
  CHUONGTRINH: 'DM_CHUONGTRINH', DIEUKIEN: 'DM_DIEUKIEN',
  // MỚI: Danh mục tổng hợp doanh thu sản lượng (Mã doanh thu) + Danh mục kế hoạch doanh thu
  DOANHTHU: 'DM_DOANHTHU', KEHOACH: 'DM_KEHOACH', KEHOACH_CHITIET: 'DM_KEHOACH_CHITIET'
};

function doGet() {
  checkAccess_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Quản lý Chiết khấu NCC')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ===== KIỂM SOÁT QUYỀN TRUY CẬP (TUỲ CHỌN, TẮT MẶC ĐỊNH) =====
// Mặc định: KHÔNG giới hạn gì (giữ đúng hành vi cũ, không phá vỡ ứng dụng đang chạy).
// Để BẬT giới hạn theo email: mở Apps Script editor -> Project Settings -> Script Properties
// -> thêm property "ALLOWED_EMAILS" với giá trị là danh sách email được phép, phân tách bởi dấu phẩy
// (VD: "an@congty.com, binh@congty.com"). Khi property này có giá trị, mọi lượt mở app (doGet) và mọi
// lượt LƯU dữ liệu (qua withLock_) đều bị chặn nếu người dùng không có trong danh sách.
// LƯU Ý QUAN TRỌNG (giới hạn của chính nền tảng Apps Script, không thể vượt qua bằng code):
// Session.getActiveUser().getEmail() CHỈ trả về đúng email khi Web App được deploy với
// "Execute as: User accessing the web app" VÀ người dùng thuộc cùng Google Workspace domain với người
// deploy (hoặc file được chia sẻ trực tiếp cho họ). Nếu deploy "Execute as: Me" (mặc định phổ biến),
// hàm này luôn trả về chuỗi rỗng — khi đó, để an toàn, checkAccess_() sẽ TỪ CHỐI truy cập luôn (thay vì
// âm thầm cho qua) một khi đã bật ALLOWED_EMAILS, kèm thông báo rõ nguyên nhân để người quản trị biết
// cách xử lý (đổi chế độ deploy), tránh ảo tưởng là "đã có bảo mật" trong khi thực chất không kiểm
// tra được danh tính.
function checkAccess_() {
  const allowedRaw = PropertiesService.getScriptProperties().getProperty('ALLOWED_EMAILS');
  if (!allowedRaw) return; // Chưa cấu hình -> không giới hạn, giữ nguyên hành vi mặc định
  const allowed = allowedRaw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
  if (!allowed.length) return;
  let email = '';
  try { email = (Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  if (!email) {
    throw new Error('Không xác định được danh tính người dùng để kiểm tra quyền truy cập (ALLOWED_EMAILS đã bật). ' +
      'Hãy deploy Web App với chế độ "Execute as: User accessing the web app" để tính năng này hoạt động đúng.');
  }
  if (allowed.indexOf(email) === -1) {
    throw new Error('Tài khoản "' + email + '" chưa được cấp quyền truy cập ứng dụng này. Liên hệ quản trị viên nếu cần được cấp quyền.');
  }
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
// SỬA LỖI QUAN TRỌNG: Google Sheet đôi khi TỰ ĐỘNG định dạng nhầm 1 cột SỐ (VD cột "Giá công bố")
// thành định dạng NGÀY THÁNG — thường xảy ra khi dán dữ liệu, hoặc Sheet "đoán" định dạng theo 1 vài ô
// trông giống ngày tháng trong cùng cột. Khi đó sh.getRange(...).getValues() trả về 1 đối tượng Date()
// thay vì con số thật, dù giá trị ô KHÔNG hề thay đổi. Nếu vẫn convert Date -> chuỗi "yyyy-MM-dd" như
// trước đây một cách MÙ QUÁNG, cột Giá công bố sẽ nhận được chuỗi kiểu "2026-07-24" thay vì con số giá
// -> khi Frontend Number("2026-07-24") sẽ ra NaN -> hiển thị/tính toán ra 0đ, sai hoàn toàn.
// Nay sanitizeCellValue_ nhận biết luôn TÊN CỘT (header): nếu tên cột KHÔNG nằm trong danh sách các cột
// ngày tháng THẬT SỰ của app (DATE_FIELD_NAMES_ bên dưới), một giá trị Date() nhận được sẽ bị coi là LỖI
// ĐỊNH DẠNG và được CHUYỂN NGƯỢC LẠI thành đúng con số gốc (số ngày kể từ mốc 30/12/1899 — công thức
// serial date chuẩn của Google Sheets/Excel), giữ đúng giá trị số ban đầu người dùng đã nhập.
const DATE_FIELD_NAMES_ = new Set([
  'ngay', 'tu_ngay', 'den_ngay', 'hieuluctu', 'hieulucden', 'ngaylap', 'ngay_tinh', 'ngaytao'
]);
function serialFromDate_(d) {
  // Mốc gốc serial date của Google Sheets/Excel là 30/12/1899 — dùng đúng constructor Date() cục bộ
  // (không phải UTC) để khớp với cách Apps Script quy đổi Date() từ serial number khi đọc cell.
  const epoch = new Date(1899, 11, 30);
  return Math.round((d.getTime() - epoch.getTime()) / 86400000 * 1e6) / 1e6; // làm tròn nhẹ, tránh sai số dấu phẩy động
}
function sanitizeCellValue_(v, header) {
  // SỬA LỖI: đối tượng Date lồng trong mảng lớn có thể bị lỗi/hỏng khi truyền qua google.script.run
  // (lỗi đã biết của Apps Script). Nay luôn chuyển Date thành CHUỖI TEXT (yyyy-MM-dd) trước khi gửi,
  // áp dụng thống nhất ở đây cho mọi hàm đọc dữ liệu — tránh hoàn toàn kiểu Date trong gói tin RPC.
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    const key = String(header || '').trim().toLowerCase();
    if (!DATE_FIELD_NAMES_.has(key)) {
      // Cột này KHÔNG phải cột ngày tháng thật -> đây là lỗi định dạng ô, trả về ĐÚNG CON SỐ gốc.
      return serialFromDate_(v);
    }
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
// TỐI ƯU: gộp chung đoạn "lấy sheet theo tên, tự tạo mới nếu chưa có" — lặp lại giống hệt nhau ở
// nhiều hàm đọc dữ liệu khác nhau.
function getOrCreateSheet_(ss, sheetName) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  return sh;
}

// ===== KHOÁ ĐỒNG THỜI (LockService) =====
// SỬA LỖI NGHIÊM TRỌNG: trước đây mọi hàm lưu (saveSheetData, appendOrUpsertRows_...) đều ĐỌC toàn
// sheet -> XOÁ -> GHI LẠI mà không có khoá. Nếu 2 người dùng bấm lưu gần như đồng thời (VD cùng sửa
// Danh mục hoặc cùng lưu Báo cáo), người ghi SAU sẽ ghi đè mất hoàn toàn thay đổi của người ghi TRƯỚC
// — không có lỗi, không cảnh báo, mất dữ liệu âm thầm. Nay mọi điểm vào (entry point) có thể được gọi
// từ Frontend để LƯU dữ liệu đều được bọc qua withLock_(): dùng LockService.getScriptLock() để đảm bảo
// tại một thời điểm chỉ có DUY NHẤT 1 lượt lưu được thực thi, các lượt gọi khác phải đợi tới lượt.
// Đồng thời kiểm tra quyền truy cập (nếu ALLOWED_EMAILS đã bật) ngay tại đây để bảo vệ luôn đường ghi.
function withLock_(fn, waitMs) {
  checkAccess_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(waitMs == null ? 30000 : waitMs);
  } catch (e) {
    throw new Error('Hệ thống đang bận xử lý một lượt lưu khác từ người dùng khác, vui lòng thử lại sau vài giây.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ===== GHI SHEET TỐI ƯU (giảm rủi ro timeout khi sheet đã lớn và tăng dần theo thời gian) =====
// TRƯỚC ĐÂY: mỗi lần lưu đều clearContents() TOÀN BỘ sheet rồi appendRow(headers) + setValues() lại
// từ đầu — với các sheet lớn và ngày càng phình to (Mua hàng, Giá công bố), thao tác "xoá sạch rồi ghi
// lại từ đầu" tốn nhiều thời gian máy chủ hơn cần thiết, đúng nguyên nhân timeout đã được vá tạm bằng
// cách tải/ghi theo chunk ở các nơi khác. NAY: ghi ĐÈ TRỰC TIẾP lên đúng vùng dữ liệu mới (setValues),
// chỉ dọn phần "thừa" còn sót lại từ dữ liệu CŨ nếu dữ liệu mới ít dòng/cột hơn dữ liệu cũ trước đó —
// với trường hợp phổ biến nhất (dữ liệu ngày càng nhiều lên, không co lại), hoàn toàn không cần bước
// dọn thừa nào, giảm đáng kể số thao tác đọc/ghi so với cách "xoá sạch rồi ghi lại" trước đây.
function writeRowsEfficient_(sh, headers, values) {
  const oldLastRow = sh.getLastRow();
  const oldLastCol = sh.getLastColumn();
  const newRowCount = values.length;
  const newColCount = headers.length;
  if (newColCount > 0) {
    sh.getRange(1, 1, 1, newColCount).setValues([headers]);
    if (newRowCount > 0) {
      sh.getRange(2, 1, newRowCount, newColCount).setValues(values);
    }
  } else {
    sh.clearContents();
    return;
  }
  // Dọn phần dòng thừa còn sót lại (dữ liệu mới ít dòng hơn dữ liệu cũ)
  if (oldLastRow > newRowCount + 1) {
    sh.getRange(newRowCount + 2, 1, oldLastRow - (newRowCount + 1), Math.max(oldLastCol, newColCount)).clearContent();
  }
  // Dọn phần cột thừa còn sót lại (dữ liệu mới ít cột hơn dữ liệu cũ)
  if (oldLastCol > newColCount) {
    sh.getRange(1, newColCount + 1, Math.max(newRowCount + 1, 1), oldLastCol - newColCount).clearContent();
  }
}
function readSheetRows_(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, sheetName);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(r => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = sanitizeCellValue_(r[i], h));
    return obj;
  });
}
function loadCoreData() {
  const result = { catalog: {}, programs: [], programsLuyKe: [], reportCKTH: [], reportCKCT: [], chuongtrinh: [], dieukien: [], doanhthu: [], kehoach: [], kehoachChiTiet: [] };
  ['THUONGHIEU','NHOMKH','LOAISP','DACTINH','CONGDUNG','QUYCACH','MHCK'].forEach(key => {
    result.catalog[key.toLowerCase()] = readSheetRows_(SHEETS[key]);
  });
  result.programs = readSheetRows_(SHEETS.PROGRAMS);
  result.programsLuyKe = readSheetRows_(SHEETS.LUYKE);
  result.reportCKTH = readSheetRows_(SHEETS.REPORT_CKTH);
  result.reportCKCT = readSheetRows_(SHEETS.REPORT_CKCT);
  result.chuongtrinh = readSheetRows_(SHEETS.CHUONGTRINH);
  // DM_DIEUKIEN KHÔNG đọc nguyên khối ở đây nữa — Frontend tự tải riêng theo chunk qua
  // getDieuKienRowCount()/loadDieuKienChunk() ngay sau khi gọi loadCoreData(), cùng cơ chế phòng
  // ngừa timeout như Giá công bố/Mua hàng (xem ghi chú tại 2 hàm đó).
  result.doanhthu = readSheetRows_(SHEETS.DOANHTHU);
  result.kehoach = readSheetRows_(SHEETS.KEHOACH);
  result.kehoachChiTiet = readSheetRows_(SHEETS.KEHOACH_CHITIET);
  result.version = getDataVersion();
  return result;
}

// SỬA LỖI (chia nhỏ gói truyền): dù đã dùng định dạng gọn, gói 1.435 dòng vẫn có thể bị hỏng/rỗng khi
// truyền qua kênh giao tiếp iframe của Apps Script (hiện tượng đã biết, không báo lỗi rõ ràng). Nay
// cho phép tải THEO TỪNG GÓI NHỎ (offset/limit) — Frontend gọi nhiều lần liên tiếp, mỗi lần chỉ vài
// trăm dòng, ghép lại — đảm bảo mỗi gói luôn đủ nhỏ để truyền an toàn.
function getSheetRowCount_(sheetKey) {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, SHEETS[sheetKey]);
  return { total: Math.max(0, sh.getLastRow() - 1), version: getDataVersion() };
}
function getPurchasesRowCount() { return getSheetRowCount_('PURCHASES'); }
function getGiaCongBoRowCount() { return getSheetRowCount_('GIACONGBO'); }
// TỐI ƯU (phòng ngừa): DM_DIEUKIEN (bậc điều kiện con của mã chiết khấu, Tầng 3) hiện còn nhỏ nên
// vẫn đọc trong loadCoreData(), nhưng bổ sung sẵn cặp hàm đếm dòng + tải chunk cùng khuôn mẫu với
// Giá công bố/Mua hàng — để Frontend có thể chuyển sang tải theo gói nhỏ ngay khi cần, không phải
// chờ tới lúc sheet này phình to rồi mới gặp lại đúng lỗi "load 0 dữ liệu chập chờn" đã từng vá.
function getDieuKienRowCount() { return getSheetRowCount_('DIEUKIEN'); }
function loadDieuKienChunk(offset, limit) { return readSheetChunk_(SHEETS.DIEUKIEN, offset, limit); }
function readSheetChunk_(sheetName, offset, limit) {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, sheetName);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { headers: [], rows: [] };
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const startRow = 2 + offset; // dòng 1 là header, dữ liệu bắt đầu dòng 2
  if (startRow > lastRow) return { headers: headers, rows: [] };
  const numRows = Math.min(limit, lastRow - startRow + 1);
  const data = sh.getRange(startRow, 1, numRows, lastCol).getValues();
  const rows = data
    .filter(r => r.some(c => String(c).trim() !== ''))
    .map(r => r.map((c, i) => sanitizeCellValue_(c, headers[i])));
  return { headers: headers, rows: rows };
}
function loadPurchasesChunk(offset, limit) { return readSheetChunk_(SHEETS.PURCHASES, offset, limit); }
function loadGiaCongBoChunk(offset, limit) { return readSheetChunk_(SHEETS.GIACONGBO, offset, limit); }

// THEO YÊU CẦU: cho phép Frontend chỉ tải dữ liệu Mua hàng của MỘT SỐ Nhà cung cấp được chọn (thay vì
// luôn tải toàn bộ ~1.400+ dòng), giảm dung lượng truyền và bộ nhớ trình duyệt phải giữ khi người dùng
// chỉ cần xem/lọc theo 1 vài NCC cụ thể.

// Lấy danh sách Nhà cung cấp DUY NHẤT — CHỈ đọc 1 cột (nhacungcap), KHÔNG đọc toàn bộ các cột khác —
// dùng để hiển thị ô chọn NCC cho người dùng NGAY CẢ KHI chưa tải dữ liệu Mua hàng nào (tránh vòng lặp
// "phải tải hết mới biết có NCC nào để chọn tải theo NCC").
function getDistinctPurchaseSuppliers() {
  const sh = getOrCreateSheet_(SpreadsheetApp.getActive(), SHEETS.PURCHASES);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf('nhacungcap');
  if (idx < 0) return [];
  const values = sh.getRange(2, idx + 1, lastRow - 1, 1).getValues();
  const seen = {};
  values.forEach(r => { const v = String(r[0] || '').trim(); if (v) seen[v] = true; });
  return Object.keys(seen).sort();
}
// Đọc TOÀN BỘ dòng Mua hàng rồi LỌC theo danh sách Nhà cung cấp được chọn (so khớp không phân biệt hoa
// thường/khoảng trắng thừa). Nếu suppliers rỗng/không truyền -> trả về TOÀN BỘ (an toàn, không lọc
// nhầm mất dữ liệu nếu Frontend gọi thiếu tham số).
function readFilteredPurchaseRows_(suppliers) {
  const sh = getOrCreateSheet_(SpreadsheetApp.getActive(), SHEETS.PURCHASES);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { headers: [], rows: [] };
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf('nhacungcap');
  const wanted = (Array.isArray(suppliers) ? suppliers : []).map(s => String(s).trim().toLowerCase()).filter(Boolean);
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows = data
    .filter(r => r.some(c => String(c).trim() !== '')) // bỏ dòng trắng
    .filter(r => {
      if (!wanted.length) return true; // không lọc gì -> giữ hết (an toàn, không mất dữ liệu)
      if (idx < 0) return true; // sheet không có cột nhacungcap -> không lọc được, trả hết
      return wanted.indexOf(String(r[idx]).trim().toLowerCase()) !== -1;
    })
    .map(r => r.map((c, i) => sanitizeCellValue_(c, headers[i])));
  return { headers, rows };
}
function getPurchasesRowCountBySupplier(suppliers) {
  return { total: readFilteredPurchaseRows_(suppliers).rows.length };
}
function loadPurchasesChunkBySupplier(suppliers, offset, limit) {
  const all = readFilteredPurchaseRows_(suppliers);
  return { headers: all.headers, rows: all.rows.slice(offset, offset + limit) };
}
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
  const sh = getOrCreateSheet_(ss, sheetName);
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
  if (!data || data.length === 0) {
    sh.clearContents();
    return;
  }
  const headers = Object.keys(data[0]);
  const values = data.map(r => headers.map(h => r[h] ?? ''));
  writeRowsEfficient_(sh, headers, values);
}

// Các hàm cầu nối gọi từ Frontend
// Mỗi hàm save trả về phiên bản dữ liệu mới nhất để Frontend cập nhật cache cục bộ.
// TỐI ƯU (SỬA LỖI HIỆU NĂNG): trước đây saveCatalogData() nhận NGUYÊN object catalog (gồm cả MHCK
// và Giá công bố — có thể hơn 1.300 dòng) rồi ghi đè lại TẤT CẢ các sheet đó, dù người dùng chỉ vừa
// sửa 1 dòng ở 1 danh mục nhỏ (VD Thương hiệu). Mỗi lần lưu như vậy tốn thời gian + quota Apps Script
// để clearContents() + ghi lại toàn bộ sheet lớn không hề thay đổi, tăng rủi ro timeout không cần
// thiết. Nay saveCatalogData() CHỈ lưu 6 danh mục nhỏ (kích thước); MHCK và Giá công bố có hàm lưu
// RIÊNG (saveMhckData/saveGiaCongBoData) để Frontend chỉ gọi đúng phần thực sự thay đổi.
const DIM_CATALOG_KEYS_ = ['thuonghieu', 'nhomkh', 'loaisp', 'dactinh', 'congdung', 'quycach'];
function saveCatalogData(catalog) {
  return withLock_(() => {
    if (!catalog) return bumpDataVersion_();
    DIM_CATALOG_KEYS_.forEach(key => {
      if (catalog[key] != null) saveSheetData(SHEETS[key.toUpperCase()], catalog[key]);
    });
    return bumpDataVersion_();
  });
}
function saveMhckData(data) { return simpleSave_('MHCK', data); }
function saveGiaCongBoData(data) { return simpleSave_('GIACONGBO', data); }
// TỐI ƯU (gộp code trùng lặp): các hàm save*Data() dưới đây trước đây là các dòng lặp lại gần như
// giống hệt nhau (saveSheetData(SHEETS.X, data); return bumpDataVersion_();) — nay dùng chung 1 hàm
// simpleSave_(sheetKey, data) để giảm trùng lặp, tránh sai sót khi copy-paste thêm hàm mới sau này.
// Vẫn giữ mỗi hàm là 1 "function" khai báo riêng ở phạm vi toàn cục (không dùng const/arrow) để đảm
// bảo google.script.run từ Frontend luôn gọi được đúng tên hàm.
// NAY: bọc thêm withLock_() để chống 2 người dùng lưu cùng lúc ghi đè mất dữ liệu của nhau (xem chi
// tiết giải thích tại định nghĩa withLock_ phía trên).
function simpleSave_(sheetKey, data) {
  return withLock_(() => { saveSheetData(SHEETS[sheetKey], data); return bumpDataVersion_(); });
}
function saveProgramsData(data) { return simpleSave_('PROGRAMS', data); }
function savePurchasesData(data) { return simpleSave_('PURCHASES', data); }
function saveLuyKeData(data) { return simpleSave_('LUYKE', data); }
// THEO MÔ HÌNH 3 TẦNG: Tầng 1 (Danh mục chương trình) và Tầng 3 (Bậc điều kiện con của mã chiết khấu)
function saveChuongTrinhData(data) { return simpleSave_('CHUONGTRINH', data); }
function saveDieuKienData(data) { return simpleSave_('DIEUKIEN', data); }
// MỚI: Danh mục tổng hợp doanh thu sản lượng (Mã doanh thu) + Danh mục kế hoạch doanh thu
function saveDoanhThuData(data) { return simpleSave_('DOANHTHU', data); }
function saveKeHoachData(data) { return simpleSave_('KEHOACH', data); }
function saveKeHoachChiTietData(data) { return simpleSave_('KEHOACH_CHITIET', data); }

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
  const sh = getOrCreateSheet_(ss, sheetName);
  const data = sh.getDataRange().getValues();
  let headers, existingRows;
  if (data.length > 0 && data[0].length > 0 && String(data[0][0]) !== '') {
    headers = data[0];
    existingRows = data.slice(1).map(r => {
      let o = {};
      headers.forEach((h, i) => o[h] = sanitizeCellValue_(r[i], h));
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
  // TỐI ƯU: dùng Map thay vì object thường {} để lập chỉ mục khóa chống trùng — tránh trường hợp
  // hiếm nhưng có thật: nếu 1 khóa ghép (VD program_id||tu_ngay||den_ngay) vô tình trùng tên thuộc
  // tính có sẵn trên Object (như "__proto__"), việc gán qua ngoặc vuông trên object thường có thể cư
  // xử khác thường; Map không có rủi ro này và tra cứu/khớp trùng vẫn ở độ phức tạp O(1) như cũ dù
  // dữ liệu báo cáo tích lũy nhiều kỳ (tháng/quý/năm) theo thời gian.
  const keyOf = (r) => keyFields.map(k => normKeyVal_(r[k])).join('||');
  const idx = new Map();
  existingRows.forEach((r, i) => { idx.set(keyOf(r), i); });
  rows.forEach(r => {
    const k = keyOf(r);
    if (idx.has(k)) existingRows[idx.get(k)] = Object.assign({}, existingRows[idx.get(k)], r);
    else { existingRows.push(r); idx.set(k, existingRows.length - 1); }
  });
  // Chỉ chuẩn hóa CÁC CỘT KHÓA (keyFields) về chuỗi cố định trước khi ghi lại — các cột số liệu khác
  // (sản lượng, doanh số, tiền chiết khấu...) vẫn giữ nguyên kiểu dữ liệu gốc để không bị lỗi định dạng
  // số trong Google Sheet. Việc so khớp key ở lần lưu SAU vẫn luôn đúng vì keyOf() đã tự chuẩn hóa lại
  // giá trị đọc lên (kể cả khi Sheets tự ý chuyển ô ngày thành kiểu Date).
  existingRows.forEach(r => { keyFields.forEach(k => { r[k] = normKeyVal_(r[k]); }); });
  if (!headers.length) { sh.clearContents(); return; }
  const values = existingRows.map(r => headers.map(h => r[h] ?? ''));
  writeRowsEfficient_(sh, headers, values);
}
// NAY: bọc withLock_() — trước đây 2 người dùng cùng lưu báo cáo gần như đồng thời có thể đọc cùng 1
// bản "existing rows" rồi ghi đè lên nhau (bản lưu sau xoá mất bản ghi upsert của bản lưu trước dù
// khác khóa). Khoá đảm bảo đọc-sửa-đổi-ghi (read-modify-write) diễn ra TRỌN VẸN, không bị xen ngang.
function saveReportCKTH(rows) {
  return withLock_(() => {
    appendOrUpsertRows_(SHEETS.REPORT_CKTH, rows, ['program_id', 'tu_ngay', 'den_ngay']);
    return bumpDataVersion_();
  });
}
function saveReportCKCT(rows) {
  return withLock_(() => {
    appendOrUpsertRows_(SHEETS.REPORT_CKCT, rows, ['program_id', 'tu_ngay', 'den_ngay', 'mahang']);
    return bumpDataVersion_();
  });
}
// THEO YÊU CẦU (nút "Dọn dẹp trùng lặp" ở Frontend): GHI ĐÈ TOÀN BỘ REPORT_CKTH/REPORT_CKCT bằng đúng
// danh sách đã được Frontend lọc bỏ các dòng trùng — khác hẳn saveReportCKTH/CKCT ở trên (vốn chỉ UPSERT,
// không có khả năng XOÁ bớt dòng). Đây là thao tác GHI ĐÈ THẬT SỰ nên có rủi ro cao hơn nếu gọi nhầm với
// mảng rỗng/thiếu — saveSheetData() đã có sẵn lớp bảo vệ "không ghi đè bằng mảng rỗng nếu sheet đang có
// dữ liệu", và withLock_() đảm bảo không bị người khác ghi đè chồng lên trong lúc dọn dẹp.
function overwriteReportCKTH(rows) {
  return withLock_(() => { saveSheetData(SHEETS.REPORT_CKTH, rows); return bumpDataVersion_(); });
}
function overwriteReportCKCT(rows) {
  return withLock_(() => { saveSheetData(SHEETS.REPORT_CKCT, rows); return bumpDataVersion_(); });
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
  return withLock_(() => migrateLegacyData_impl_(), 60000);
}
function migrateLegacyData_impl_() {
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
      .map(r => { const o = {}; headers.forEach((h, i) => o[h] = sanitizeCellValue_(r[i], h)); return o; });
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
    const sh = getOrCreateSheet_(ss, sheetName);
    const existing = readSheetAsObjects_(sh).rows;
    const merged = existing.concat(objRows); // cộng dồn, không đè mất dữ liệu đã có sẵn đúng chuẩn
    const headers = Object.keys(merged[0]);
    merged.forEach(r => Object.keys(r).forEach(k => { if (headers.indexOf(k) === -1) headers.push(k); }));
    writeRowsEfficient_(sh, headers, merged.map(r => headers.map(h => r[h] ?? '')));
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
