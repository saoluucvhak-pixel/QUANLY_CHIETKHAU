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
// SỬA LỖI (rà soát theo phản ánh người dùng — "nhập lại không ghi vào Sheet, làm mới về dữ liệu cũ"):
// TRƯỚC ĐÂY "phiên bản" dữ liệu CHỈ dựa vào DriveApp.getFileById(...).getLastUpdated() — thời điểm Drive
// ghi nhận "lần sửa cuối" của FILE. Đây là chỉ số CÓ ĐỘ TRỄ LAN TRUYỀN đã biết trên nền tảng Apps
// Script/Drive (metadata Drive có thể cập nhật CHẬM hơn vài giây tới vài phút so với thời điểm
// SpreadsheetApp thực sự ghi xong dữ liệu, đặc biệt khi lưu liên tục nhiều lần trong thời gian ngắn).
// Hậu quả: ngay sau khi lưu thành công (dữ liệu ĐÃ ghi đúng lên Sheet), bumpDataVersion_() có thể trả về
// CÙNG 1 "phiên bản" như trước khi lưu (do Drive chưa kịp cập nhật) — Frontend vẫn nhận đúng dữ liệu mới
// trong chính phiên làm việc đó (vì ghi thẳng vào cache cục bộ), nhưng ở lượt tải lại SAU (mở lại app,
// máy khác, hoặc sau khi có thao tác khác làm mất cache cục bộ), loadAll() so sánh "phiên bản" cache với
// "phiên bản" server, thấy TRÙNG (do Drive vẫn chưa cập nhật) -> lầm tưởng dữ liệu chưa đổi -> tiếp tục
// dùng cache CŨ thay vì tải lại đúng dữ liệu mới từ Sheet — nhìn như vừa lưu xong nhưng "làm mới lại mất".
// Ngoài ra fallback cũ (PropertiesService 'DATA_VERSION') KHÔNG BAO GIỜ được ghi ở đâu cả trong code —
// nên khi DriveApp lỗi/thiếu quyền, hàm LUÔN trả về hằng số '0' cố định, càng dễ gây "khớp version giả".
// NAY: kết hợp CẢ 2 tín hiệu trong 1 chuỗi phiên bản — (a) bộ đếm tăng dần lưu trong PropertiesService,
// TĂNG NGAY LẬP TỨC (không có độ trễ) mỗi khi lưu qua app (bumpDataVersion_ luôn set lại giá trị thật),
// và (b) thời điểm Drive lastUpdated như cũ để VẪN phát hiện được khi người dùng sửa tay trực tiếp trên
// Sheet (không qua app — trường hợp bộ đếm (a) không tự biết). Chỉ cần 1 trong 2 tín hiệu đổi là chuỗi
// phiên bản đổi theo, và (a) đảm bảo LUÔN đổi ngay sau mỗi lần lưu qua app dù Drive có trễ hay không.
function getDataVersion() {
  const props = PropertiesService.getScriptProperties();
  const counter = props.getProperty('DATA_VERSION') || '0';
  let driveTs = '0';
  try {
    const ss = SpreadsheetApp.getActive();
    driveTs = String(DriveApp.getFileById(ss.getId()).getLastUpdated().getTime());
  } catch (e) {
    // Thiếu quyền Drive hoặc lỗi tạm thời -> bỏ qua tín hiệu này, vẫn còn bộ đếm (a) đảm bảo phát hiện
    // đúng mọi thay đổi lưu QUA APP (chỉ không phát hiện được sửa tay trực tiếp trên Sheet trong TH này).
  }
  return driveTs + '.' + counter;
}
function bumpDataVersion_() {
  // SpreadsheetApp có thể trì hoãn ghi tới cuối lần thực thi; flush() để đảm bảo ghi xong ngay trước khi
  // tính phiên bản mới.
  try { SpreadsheetApp.flush(); } catch (e) {}
  const props = PropertiesService.getScriptProperties();
  const next = String((Number(props.getProperty('DATA_VERSION')) || 0) + 1);
  props.setProperty('DATA_VERSION', next);
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
  'ngay', 'tu_ngay', 'den_ngay', 'hieuluctu', 'hieulucden', 'ngaylap', 'ngay_tinh', 'ngaytao',
  // SỬA LỖI (báo lỗi thực tế: "nhập Khoảng ngày xét đạt KH xong, lưu rồi làm mới lại là mất"): 2 cột
  // này (Mã chiết khấu / PROGRAMS, thêm ngày 28/08/2026) LÀ cột ngày tháng thật nhưng trước đây THIẾU
  // trong danh sách này. Google Sheets tự nhận diện chuỗi "yyyy-mm-dd" ghi vào 2 cột này và lưu thành ô
  // Date thật -> khi đọc lại, vì tên cột không có ở đây nên bị coi NHẦM là "lỗi định dạng" và bị chuyển
  // ngược thành số serial thô (VD 46123) thay vì đúng chuỗi ngày -> ô <input type="date"> ở Frontend
  // không hiểu số serial này nên hiển thị RỖNG, đúng như hiện tượng "nhập xong, lưu, làm mới là mất".
  'kh_tu_ngay', 'kh_den_ngay'
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
// TỐI ƯU TỐC ĐỘ TẢI (phát hiện khi rà soát báo cáo "tải rất chậm"): Google Sheets giữ lại rất nhiều
// "hàng trống ma" ở cuối sheet (từng có nội dung/định dạng, đã xoá nhưng Sheet vẫn tính vào
// getDataRange()/getLastRow()) — kiểm chứng thực tế trên dữ liệu người dùng: CHUONGTRINHCHIETKHAU có
// 695 dòng nhưng chỉ 11 dòng có dữ liệu thật, DM_MHCK 999 dòng nhưng chỉ 391 dòng thật, DM_QUYCACH 999
// dòng nhưng chỉ ~9 dòng thật. loadCoreData() gọi hàm này cho TOÀN BỘ danh mục NGAY KHI MỞ APP (không
// tải lười/phân trang như Giá công bố/Mua hàng), nên trước đây các hàng trống ma này vẫn bị đọc,
// sanitize, đóng gói JSON và truyền qua kênh RPC mỗi lần mở app — góp phần đáng kể vào cảm giác "tải
// rất chậm". Nay bỏ qua hẳn các hàng trống hoàn toàn trước khi dựng đối tượng, giống đúng cách
// readSheetChunk_()/readFilteredPurchaseRows_() bên dưới đã làm từ trước cho Giá công bố/Mua hàng.
function readSheetRows_(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, sheetName);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(r => r.some(c => String(c).trim() !== ''))
    .map(r => {
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
  // TỐI ƯU: dùng Set thay vì mảng để tra "có trong danh sách NCC được chọn hay không" — tránh quét lại
  // toàn bộ mảng wanted (indexOf) cho MỖI dòng dữ liệu (có thể 1.400+ dòng), giảm từ O(n×m) xuống O(n).
  const wanted = new Set((Array.isArray(suppliers) ? suppliers : []).map(s => String(s).trim().toLowerCase()).filter(Boolean));
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows = data
    .filter(r => r.some(c => String(c).trim() !== '')) // bỏ dòng trắng
    .filter(r => {
      if (!wanted.size) return true; // không lọc gì -> giữ hết (an toàn, không mất dữ liệu)
      if (idx < 0) return true; // sheet không có cột nhacungcap -> không lọc được, trả hết
      return wanted.has(String(r[idx]).trim().toLowerCase());
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

// =====================================================================================
// TRUY VẤN TRỰC TIẾP TỪ SHEET (theo yêu cầu 25/08/2026 — "sao mỗi lần load lại phải tải lên cache,
// sao mỗi lần load không đọc thẳng google sheet"): 2 màn hình DUYỆT/DANH SÁCH — Danh mục Giá công bố
// và Sổ chi tiết mua hàng — không còn cần Frontend tải NGUYÊN khối dữ liệu (~1.300-1.400 dòng mỗi bảng)
// vào bộ nhớ trình duyệt rồi lọc/phân trang tại chỗ. Thay vào đó, MỖI LẦN người dùng lọc/chuyển trang,
// Frontend gọi 1 trong các hàm dưới đây — hàm ĐỌC THẲNG Sheet (luôn là dữ liệu MỚI NHẤT, không có khái
// niệm "cache" nào ở đây), lọc + phân trang ngay tại máy chủ, chỉ trả về ĐÚNG phần nhỏ (mặc định 20
// dòng) cần hiển thị. Việc lọc vẫn phải đọc toàn bộ dữ liệu VÀO BỘ NHỚ MÁY CHỦ (Apps Script) để so khớp
// điều kiện — nhưng đây là thao tác ĐỌC nội bộ trong 1 lượt gọi (nhanh, không tốn kênh truyền iframe
// tới trình duyệt), khác hẳn với việc TRUYỀN NGUYÊN 1.300+ dòng qua network cho Frontend như trước đây.
// Cơ chế TẢI TOÀN BỘ cũ (ensureGiaCongBoLoaded/ensurePurchasesLoaded/ensurePurchasesLoadedFiltered ở
// Frontend, gọi readSheetChunk_ ở trên) VẪN GIỮ NGUYÊN KHÔNG ĐỔI — các màn hình cần dữ liệu ĐẦY ĐỦ để
// tính toán (Tính chiết khấu, Tổng quan, Đối chiếu giá, Báo cáo) tiếp tục dùng cơ chế đó như cũ.
const DIM_KEYS_ = ['thuonghieu', 'nhomkh', 'loaisp', 'dactinh', 'congdung', 'quycach'];

// Đọc TOÀN BỘ Giá công bố vào bộ nhớ máy chủ (không gửi qua Frontend), lọc theo từ khóa tìm kiếm (Mã MH
// / Tên hàng) và theo "đang hiệu lực tại 1 ngày" (mặc định hôm nay, theo múi giờ máy chủ) nếu
// filters.onlyActive !== false, rồi CHỈ trả về đúng 1 trang (offset/limit) — kèm _row (số dòng thật trên
// Sheet, 2-based) để Frontend dùng khi Sửa/Xóa đúng dòng đó mà không cần tải cả bảng.
function queryGiaCongBoPage(filters, offset, limit) {
  filters = filters || {};
  offset = Number(offset) || 0;
  limit = Number(limit) || 20;
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, SHEETS.GIACONGBO);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  const emptyResult = { headers: [], rows: [], total: 0, version: getDataVersion() };
  if (lastRow < 2 || lastCol < 1) return emptyResult;
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const iMamh = headers.indexOf('mamh_ncc');
  const iTen = headers.indexOf('tenhang');
  const iTu = headers.indexOf('hieuluctu');
  const iDen = headers.indexOf('hieulucden');
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd');
  const asOf = filters.asOfDate || today;
  const onlyActive = filters.onlyActive !== false;
  const q = String(filters.q || '').trim().toLowerCase();

  const matched = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (!r.some(c => String(c).trim() !== '')) continue; // bỏ dòng trắng
    if (onlyActive) {
      const tu = iTu > -1 ? sanitizeCellValue_(r[iTu], 'hieuluctu') : '';
      const den = iDen > -1 ? sanitizeCellValue_(r[iDen], 'hieulucden') : '';
      if (tu && String(tu) > asOf) continue; // chưa tới ngày hiệu lực
      if (den && String(den) < asOf) continue; // đã hết hiệu lực
    }
    if (q) {
      const mamh = iMamh > -1 ? String(sanitizeCellValue_(r[iMamh], 'mamh_ncc') || '').toLowerCase() : '';
      const ten = iTen > -1 ? String(sanitizeCellValue_(r[iTen], 'tenhang') || '').toLowerCase() : '';
      if (mamh.indexOf(q) === -1 && ten.indexOf(q) === -1) continue;
    }
    matched.push(i);
  }
  const total = matched.length;
  const pageIdx = matched.slice(offset, offset + limit);
  const rows = pageIdx.map(i => {
    const r = data[i];
    const obj = {};
    headers.forEach((h, c) => obj[h] = sanitizeCellValue_(r[c], h));
    obj._row = i + 2; // dòng thật trên Sheet (1 = header)
    return obj;
  });
  return { headers, rows, total, version: getDataVersion() };
}

// Ghi 1 dòng Giá công bố vào ĐÚNG vị trí dòng thật trên Sheet (rowNum, lấy từ obj._row do
// queryGiaCongBoPage trả về) — dùng cho "Sửa". Nếu rowNum rỗng/null -> ghi thêm dòng MỚI ở cuối (dùng
// cho "Thêm"). Không cần Frontend phải tải cả bảng để làm việc này.
function upsertGiaCongBoRowAt(rowNum, rec) {
  return withLock_(() => {
    const ss = SpreadsheetApp.getActive();
    const sh = getOrCreateSheet_(ss, SHEETS.GIACONGBO);
    const lastCol = sh.getLastColumn();
    let headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    let headersChanged = headers.length === 0;
    Object.keys(rec || {}).forEach(k => { if (headers.indexOf(k) === -1) { headers.push(k); headersChanged = true; } });
    if (headersChanged) sh.getRange(1, 1, 1, headers.length).setValues([headers]);

    // CHỐNG TRÙNG (theo yêu cầu 25/08/2026): trước đây khi nhập tay từng dòng qua form "Thêm/Sửa giá
    // công bố", hệ thống KHÔNG kiểm tra trùng — có thể vô tình tạo 2 dòng cùng Mã MH + cùng khoảng hiệu
    // lực (trong khi nhập file Excel đã luôn tự gộp trùng theo đúng bộ khóa này — xem bulkUpsertGiaCongBo
    // / appendOrUpsertRows_). Nay áp dụng CÙNG bộ khóa (Mã MH + Hiệu lực từ + Hiệu lực đến) để kiểm tra
    // trước khi ghi:
    //  - Nếu đang "Thêm" (rowNum rỗng) mà trùng 1 dòng đã có -> GỘP vào đúng dòng đó (giống hệt cách
    //    nhập file Excel xử lý trùng) thay vì tạo thêm 1 dòng trùng mới.
    //  - Nếu đang "Sửa" 1 dòng mà nội dung sửa khiến nó trùng khóa với 1 dòng KHÁC đã có sẵn -> CHẶN lại
    //    (báo lỗi rõ ràng) thay vì âm thầm ghi đè nhầm dòng khác hoặc tạo trùng.
    const isAdd = !rowNum;
    const iMamh = headers.indexOf('mamh_ncc');
    const iTu = headers.indexOf('hieuluctu');
    const iDen = headers.indexOf('hieulucden');
    let dupRow = null;
    if (iMamh > -1) {
      const lastRow = sh.getLastRow();
      if (lastRow >= 2) {
        const data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
        const wantKey = normKeyVal_(rec.mamh_ncc) + '||' + normKeyVal_(rec.hieuluctu) + '||' + normKeyVal_(rec.hieulucden);
        for (let i = 0; i < data.length; i++) {
          const r = data[i];
          if (!r.some(c => String(c).trim() !== '')) continue;
          const physicalRow = i + 2;
          if (rowNum && physicalRow === rowNum) continue; // chính dòng đang sửa, bỏ qua
          const rowKey = normKeyVal_(iMamh > -1 ? r[iMamh] : '') + '||' + normKeyVal_(iTu > -1 ? r[iTu] : '') + '||' + normKeyVal_(iDen > -1 ? r[iDen] : '');
          if (rowKey === wantKey) { dupRow = physicalRow; break; }
        }
      }
    }
    let merged = false;
    if (dupRow) {
      if (isAdd) { rowNum = dupRow; merged = true; }
      else {
        throw new Error('Đã có 1 bản ghi khác với đúng Mã MH "' + rec.mamh_ncc + '" và cùng khoảng hiệu lực (dòng ' + dupRow + ' trên Sheet). Vui lòng sửa/xóa bản ghi đó trước, hoặc đổi khoảng hiệu lực của bản ghi đang sửa để tránh trùng.');
      }
    }

    const values = headers.map(h => (rec && rec[h] != null) ? rec[h] : '');
    if (rowNum) {
      sh.getRange(rowNum, 1, 1, headers.length).setValues([values]);
    } else {
      sh.getRange(sh.getLastRow() + 1, 1, 1, headers.length).setValues([values]);
    }
    return { version: bumpDataVersion_(), merged: merged, row: rowNum };
  });
}
// Xóa ĐÚNG 1 dòng thật trên Sheet theo rowNum (obj._row) — không cần tải/ghi lại cả bảng.
function deleteGiaCongBoRowAt(rowNum) {
  return withLock_(() => {
    const ss = SpreadsheetApp.getActive();
    const sh = getOrCreateSheet_(ss, SHEETS.GIACONGBO);
    if (rowNum >= 2 && rowNum <= sh.getLastRow()) sh.deleteRow(rowNum);
    return bumpDataVersion_();
  });
}
// Nạp file Excel Giá công bố: hợp nhất (cập nhật đè nếu trùng Mã MH + khoảng hiệu lực, thêm mới nếu
// chưa có) NGAY TẠI MÁY CHỦ — Frontend chỉ cần gửi lên đúng các dòng đọc được từ file Excel, KHÔNG cần
// tải trước toàn bộ Giá công bố hiện có trên Sheet để tự so khớp như trước đây.
function bulkUpsertGiaCongBo(newRecords) {
  return withLock_(() => {
    appendOrUpsertRows_(SHEETS.GIACONGBO, newRecords || [], ['mamh_ncc', 'hieuluctu', 'hieulucden']);
    return bumpDataVersion_();
  });
}

// XEM TRƯỚC (KHÔNG GHI GÌ) — theo yêu cầu 25/08/2026: trước khi thực sự ghi file Giá công bố vừa tải
// lên, so từng dòng với dữ liệu ĐANG CÓ trên Sheet (đúng bộ khóa Mã MH + Hiệu lực từ + Hiệu lực đến đã
// dùng để hợp nhất ở bulkUpsertGiaCongBo) và trả về trạng thái từng dòng để Frontend hiện "bản nháp" cho
// người dùng tự chọn dòng nào sẽ thực sự ghi/cập nhật, tránh ghi nhầm/ghi sai lên Sheet chính:
//  - 'moi': chưa có bản ghi nào trùng khóa -> sẽ THÊM MỚI nếu được chọn ghi.
//  - 'capnhat': đã có bản ghi trùng khóa nhưng Giá công bố hoặc Tên hàng khác -> sẽ GHI ĐÈ nếu được chọn.
//  - 'khongdoi': đã có bản ghi trùng khóa và giống hệt -> không cần ghi lại (mặc định KHÔNG tick chọn).
//  - 'loi': dữ liệu dòng đó không hợp lệ (thiếu Mã MH, Giá công bố không hợp lệ, hoặc Hiệu lực đến sớm
//    hơn Hiệu lực từ) -> không cho chọn ghi.
function previewGiaCongBoImport(records) {
  const recs = records || [];
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, SHEETS.GIACONGBO);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  const existingByKey = {};
  if (lastRow >= 2 && lastCol >= 1) {
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    data.forEach(r => {
      if (!r.some(c => String(c).trim() !== '')) return;
      const obj = {};
      headers.forEach((h, c) => obj[h] = sanitizeCellValue_(r[c], h));
      const key = normKeyVal_(obj.mamh_ncc) + '||' + normKeyVal_(obj.hieuluctu) + '||' + normKeyVal_(obj.hieulucden);
      existingByKey[key] = obj;
    });
  }
  return recs.map(rec => {
    const errors = [];
    const gia = Number(rec.gia_congbo);
    if (!rec.mamh_ncc) errors.push('Thiếu Mã MH');
    if (rec.gia_congbo === '' || rec.gia_congbo == null || isNaN(gia) || gia <= 0) errors.push('Giá công bố không hợp lệ');
    if (rec.hieuluctu && rec.hieulucden && String(rec.hieuluctu) > String(rec.hieulucden)) errors.push('Hiệu lực đến ngày sớm hơn Hiệu lực từ ngày');
    if (errors.length) return { status: 'loi', rec: rec, existing: null, reason: errors.join('; ') };
    const key = normKeyVal_(rec.mamh_ncc) + '||' + normKeyVal_(rec.hieuluctu) + '||' + normKeyVal_(rec.hieulucden);
    const existing = existingByKey[key];
    if (!existing) return { status: 'moi', rec: rec, existing: null };
    const giaChanged = Number(existing.gia_congbo) !== Number(rec.gia_congbo);
    const tenChanged = String(existing.tenhang || '').trim() !== String(rec.tenhang || '').trim();
    return { status: (giaChanged || tenChanged) ? 'capnhat' : 'khongdoi', rec: rec, existing: existing };
  });
}

// Đọc TOÀN BỘ Sổ chi tiết mua hàng vào bộ nhớ máy chủ, lọc theo ngày/NCC/từ khóa/6 kích thước phân loại
// (tra qua DM_MHCK), rồi CHỈ trả về đúng 1 trang (offset/limit, mặc định 20 dòng/trang) — kèm tổng số
// dòng khớp, tổng sản lượng, tổng giá trị và số NCC trong TOÀN BỘ kết quả khớp (không chỉ trang đang
// xem), để Frontend hiện đúng KPI tổng quan dù chỉ tải 1 trang nhỏ.
function queryPurchasesPage(filters, offset, limit) {
  filters = filters || {};
  offset = Number(offset) || 0;
  limit = Number(limit) || 20;
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, SHEETS.PURCHASES);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  const emptyResult = { headers: [], rows: [], total: 0, totalSL: 0, totalGiaTri: 0, supplierCount: 0, version: getDataVersion() };
  if (lastRow < 2 || lastCol < 1) return emptyResult;
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const anyDimFilter = DIM_KEYS_.some(k => filters[k]);
  let mhckMap = null;
  if (anyDimFilter) {
    mhckMap = {};
    readSheetRows_(SHEETS.MHCK).forEach(r => { if (r.mamh) mhckMap[String(r.mamh).trim().toLowerCase()] = r; });
  }

  const q = String(filters.q || '').trim().toLowerCase();
  const supplierQ = String(filters.supplier || '').trim().toLowerCase();
  const from = filters.from || '';
  const to = filters.to || '';

  const matched = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (!r.some(c => String(c).trim() !== '')) continue;
    const ngay = idx['ngay'] > -1 ? String(sanitizeCellValue_(r[idx['ngay']], 'ngay') || '') : '';
    if (from && ngay < from) continue;
    if (to && ngay > to) continue;
    if (supplierQ) {
      const ncc = idx['nhacungcap'] > -1 ? String(sanitizeCellValue_(r[idx['nhacungcap']], 'nhacungcap') || '').toLowerCase() : '';
      if (ncc.indexOf(supplierQ) === -1) continue;
    }
    if (q) {
      const mahang = idx['mahang'] > -1 ? String(sanitizeCellValue_(r[idx['mahang']], 'mahang') || '').toLowerCase() : '';
      const tenhang = idx['tenhang'] > -1 ? String(sanitizeCellValue_(r[idx['tenhang']], 'tenhang') || '').toLowerCase() : '';
      if (mahang.indexOf(q) === -1 && tenhang.indexOf(q) === -1) continue;
    }
    if (anyDimFilter) {
      const mahangKey = idx['mahang'] > -1 ? String(sanitizeCellValue_(r[idx['mahang']], 'mahang') || '').toLowerCase() : '';
      const dims = mhckMap[mahangKey] || {};
      let dimOk = true;
      DIM_KEYS_.forEach(k => { if (filters[k] && String(dims[k] || '').toLowerCase() !== String(filters[k]).toLowerCase()) dimOk = false; });
      if (!dimOk) continue;
    }
    matched.push(i);
  }
  const total = matched.length;
  let totalSL = 0, totalGiaTri = 0;
  const suppliersSet = {};
  matched.forEach(i => {
    const r = data[i];
    totalSL += (idx['soluong'] > -1 ? Number(sanitizeCellValue_(r[idx['soluong']], 'soluong')) : 0) || 0;
    totalGiaTri += (idx['giatri'] > -1 ? Number(sanitizeCellValue_(r[idx['giatri']], 'giatri')) : 0) || 0;
    const ncc = idx['nhacungcap'] > -1 ? sanitizeCellValue_(r[idx['nhacungcap']], 'nhacungcap') : '';
    if (ncc) suppliersSet[ncc] = true;
  });
  const pageIdx = matched.slice(offset, offset + limit);
  const rows = pageIdx.map(i => {
    const r = data[i];
    const obj = {};
    headers.forEach((h, c) => obj[h] = sanitizeCellValue_(r[c], h));
    obj._row = i + 2;
    return obj;
  });
  return { headers, rows, total, totalSL, totalGiaTri, supplierCount: Object.keys(suppliersSet).length, version: getDataVersion() };
}

// Nạp file Excel Sổ chi tiết mua hàng: CHỈ THÊM dòng thật sự MỚI (so khớp trùng theo cột "id" — khóa
// ghép [ngày, số HĐ, mã hàng, số lượng, giá trị] Frontend đã tự ghép sẵn trước khi gửi lên, giữ đúng quy
// tắc chống trùng đã dùng từ trước) — thực hiện NGAY TẠI MÁY CHỦ, chỉ đọc 1 CỘT "id" hiện có để so khớp
// (không cần đọc/truyền cả bảng), rồi CHỈ ghi thêm đúng các dòng mới vào cuối Sheet — không cần Frontend
// phải tải trước toàn bộ ~1.400+ dòng Mua hàng hiện có để tự so khớp trùng như trước đây.
function bulkAppendPurchases(newRecords) {
  return withLock_(() => {
    const recs = newRecords || [];
    if (!recs.length) return { added: 0, skipped: 0, version: getDataVersion() };
    const ss = SpreadsheetApp.getActive();
    const sh = getOrCreateSheet_(ss, SHEETS.PURCHASES);
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    let headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    let headersChanged = headers.length === 0;
    recs.forEach(r => Object.keys(r).forEach(k => { if (headers.indexOf(k) === -1) { headers.push(k); headersChanged = true; } }));
    const idIdx = headers.indexOf('id');
    const existingIds = new Set();
    if (idIdx > -1 && lastRow >= 2) {
      sh.getRange(2, idIdx + 1, lastRow - 1, 1).getValues().forEach(r => { if (r[0] !== '') existingIds.add(String(r[0])); });
    }
    if (headersChanged) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    const toAppend = [];
    let skipped = 0;
    recs.forEach(rec => {
      const id = String(rec.id || '');
      if (id && existingIds.has(id)) { skipped++; return; }
      if (id) existingIds.add(id);
      toAppend.push(headers.map(h => (rec[h] != null) ? rec[h] : ''));
    });
    if (toAppend.length) {
      const startRow = Math.max(sh.getLastRow() + 1, 2);
      sh.getRange(startRow, 1, toAppend.length, headers.length).setValues(toAppend);
    }
    return { added: toAppend.length, skipped: skipped, version: bumpDataVersion_() };
  });
}

// XEM TRƯỚC (KHÔNG GHI GÌ) — theo yêu cầu 25/08/2026: trước khi thực sự ghi file Sổ chi tiết mua hàng
// vừa tải lên, so từng dòng với dữ liệu ĐANG CÓ trên Sheet (đúng cột "id" đã dùng để chống trùng ở
// bulkAppendPurchases) và trả về trạng thái từng dòng để Frontend hiện "bản nháp" cho người dùng tự chọn
// dòng nào sẽ thực sự ghi thêm, tránh ghi trùng/ghi sai lên Sheet chính:
//  - 'moi': chưa có dòng nào trùng "id" (kể cả so với các dòng KHÁC trong CÙNG file đang xem trước) ->
//    sẽ THÊM MỚI nếu được chọn ghi.
//  - 'trunglap': đã có 1 dòng giống hệt (trùng Ngày + Số HĐ + Mã hàng + Số lượng + Giá trị) trên Sheet
//    hoặc trùng với 1 dòng khác đứng trước trong CHÍNH file này -> mặc định KHÔNG tick chọn (ghi lại chỉ
//    tạo ra dòng trùng vô nghĩa vì Sổ mua hàng không có khái niệm "cập nhật đè" như Giá công bố).
//  - 'loi': thiếu Mã hàng hoặc Ngày chứng từ không đọc được -> không cho chọn ghi.
function previewPurchasesImport(records) {
  const recs = records || [];
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, SHEETS.PURCHASES);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  const existingIds = new Set();
  if (lastRow >= 2 && lastCol >= 1) {
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const idIdx = headers.indexOf('id');
    if (idIdx > -1) {
      sh.getRange(2, idIdx + 1, lastRow - 1, 1).getValues().forEach(r => { if (r[0] !== '') existingIds.add(String(r[0])); });
    }
  }
  const seenInThisFile = new Set();
  return recs.map(rec => {
    const errors = [];
    if (!rec.mahang) errors.push('Thiếu Mã hàng');
    if (!rec.ngay) errors.push('Ngày chứng từ không đọc được');
    if (errors.length) return { status: 'loi', rec: rec, reason: errors.join('; ') };
    const id = String(rec.id || '');
    if (id && (existingIds.has(id) || seenInThisFile.has(id))) return { status: 'trunglap', rec: rec };
    if (id) seenInThisFile.add(id);
    return { status: 'moi', rec: rec };
  });
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
// SỬA LỖI (25/08/2026 — theo báo lỗi thực tế "xóa toàn bộ Giá công bố mà Google Sheet vẫn còn nguyên"):
// tham số forceEmpty (mặc định false/undefined) — khi TRUE, cho phép ghi mảng RỖNG đè lên dữ liệu đang
// có (bỏ qua bước bảo vệ bên dưới). Trước đây hàm này KHÔNG có cách nào phân biệt "mảng rỗng vì lỗi tải
// dữ liệu" (trường hợp gốc mà bước bảo vệ dưới đây được thêm vào để chặn) với "mảng rỗng vì người dùng
// CHỦ ĐỘNG bấm nút Xóa toàn bộ" (đã có confirm() riêng ở Frontend xác nhận ý định) — cả 2 đều bị chặn
// như nhau, khiến nút "Xóa toàn bộ" không có tác dụng thật trên Sheet dù Frontend báo "đã lưu" (vì
// bumpDataVersion_() vẫn chạy bất kể có ghi hay không). Nay Frontend truyền forceEmpty=true CHỈ khi gọi
// từ đúng nút "Xóa toàn bộ" (đã qua bước confirm() của người dùng) — mọi lượt gọi khác giữ nguyên hành
// vi bảo vệ cũ, không đổi.
function saveSheetData(sheetName, data, forceEmpty) {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ss, sheetName);
  // SỬA LỖI NGHIÊM TRỌNG: trước đây hàm này LUÔN xóa sạch (clearContents) tab rồi ghi lại, kể cả khi
  // "data" truyền vào là mảng RỖNG. saveCatalogData() gửi lên TOÀN BỘ object catalog (mọi danh mục)
  // mỗi khi lưu — kể cả khi người dùng chỉ vừa import 1 loại danh mục (VD chỉ Giá công bố). Nếu tại
  // thời điểm đó trình duyệt chưa tải đầy đủ các danh mục KHÁC (VD Thương hiệu, Nhóm KH...) từ Sheet
  // lên bộ nhớ tạm, các mảng đó sẽ RỖNG trong bộ nhớ, khiến lệnh lưu XÓA SẠCH dữ liệu thật đang có
  // trên các tab đó dù người dùng không hề có ý định xóa. Nay: nếu dữ liệu gửi lên rỗng NHƯNG tab đích
  // đang có sẵn dữ liệu thật, KHÔNG xóa (bỏ qua, giữ nguyên dữ liệu cũ) — để tránh mất dữ liệu ngoài ý
  // muốn; chỉ cho phép xóa sạch khi tab đích hiện đang trống, dữ liệu gửi lên cũng có nội dung, hoặc
  // forceEmpty=true (người dùng đã chủ động xác nhận xóa toàn bộ qua đúng nút bấm dành cho việc đó).
  const existingRowCount = Math.max(0, sh.getLastRow() - 1);
  if ((!data || data.length === 0) && existingRowCount > 0 && !forceEmpty) {
    return; // Bỏ qua để bảo vệ dữ liệu đã có, không ghi đè bằng mảng rỗng NGOÀI Ý MUỐN
  }
  if (!data || data.length === 0) {
    sh.clearContents();
    return;
  }
  // SỬA LỖI NGHIÊM TRỌNG (28/08/2026 — báo lỗi thực tế: "Mục 7 nhập Khoảng ngày xét đạt KH xong làm
  // mới bị mất"): TRƯỚC ĐÂY headers chỉ lấy Object.keys(data[0]) — DUY NHẤT dòng ĐẦU TIÊN trong mảng
  // "data" gửi lên. "data" ở đây thường là TOÀN BỘ mảng trong bộ nhớ (VD DB.programs), trộn lẫn cả các
  // dòng CŨ (tải nguyên từ Sheet, chỉ có đúng các cột ĐÃ TỪNG tồn tại lúc đó) và dòng VỪA SỬA/THÊM (có
  // đủ trường MỚI, VD kh_tu_ngay/kh_den_ngay thêm ngày 28/08/2026). Nếu dòng đứng ĐẦU mảng (thường là
  // chương trình cũ nhất) là 1 bản ghi CŨ chưa từng có trường mới đó, headers tính ra sẽ THIẾU HẲN cột
  // này — khiến MỌI dòng khác (kể cả dòng vừa nhập có giá trị thật) đều bị ghi thiếu cột đó lên Sheet.
  // Mất dữ liệu ÂM THẦM: số dòng vẫn khớp (chỉ thiếu CỘT, không thiếu DÒNG) nên qua lọt bước xác minh
  // verifyThenCache_ ở Frontend (chỉ đối chiếu SỐ DÒNG) mà không có cảnh báo gì. Nay lấy HỘI (union) của
  // TOÀN BỘ khoá xuất hiện ở BẤT KỲ dòng nào trong "data", cộng thêm giữ nguyên thứ tự các cột ĐÃ CÓ SẴN
  // trên Sheet trước đó (nếu có) ở đầu danh sách — vừa không mất trường nào, vừa không xáo trộn thứ tự
  // cột hiện có mỗi lần lưu. Áp dụng đúng cách làm đã dùng ở appendOrUpsertRows_()/writeTarget_() phía
  // dưới, nay đồng bộ luôn cho saveSheetData() — đường lưu chính của hầu hết mọi bảng trong app.
  const existingHeaders = existingRowCount > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  const headers = unionHeaders_(existingHeaders, data);
  // SỬA LỖI (mất dữ liệu ÂM THẦM ở cột "mồ côi"): unionHeaders_ ở trên chỉ đảm bảo TÊN cột không bị mất
  // (VD 1 cột được thêm trực tiếp trên Sheet, ngoài app, SAU khi trình duyệt đã tải "data" vào bộ nhớ —
  // nên KHÔNG dòng nào trong "data" có cột đó làm khoá riêng). Trước đây bước ghép "values" bên dưới vẫn
  // dùng `r[h] ?? ''` cho MỌI cột, kể cả cột "mồ côi" này — tức là GIỮ ĐÚNG TÊN cột nhưng XOÁ SẠCH toàn bộ
  // giá trị của cột đó trên Sheet (ghi '' cho mọi dòng), dù người dùng không hề có ý định đó. Nay: với các
  // cột đã có sẵn trên Sheet (existingHeaders) mà KHÔNG dòng nào trong "data" sở hữu (hasOwnProperty) làm
  // khoá riêng, giữ NGUYÊN giá trị cũ theo đúng vị trí dòng hiện có trên Sheet thay vì ghi đè bằng chuỗi
  // rỗng — chỉ những dòng thực sự MỚI (vượt quá số dòng cũ) mới chấp nhận để trống ở cột đó (không có dữ
  // liệu lịch sử nào để giữ).
  const orphanCols = {};
  headers.forEach(h => {
    if (existingHeaders.indexOf(h) === -1) return; // cột hoàn toàn mới, không có dữ liệu cũ cần giữ
    const ownedByAny = data.some(r => r && Object.prototype.hasOwnProperty.call(r, h));
    if (!ownedByAny) orphanCols[h] = true;
  });
  // SỬA LỖI (rà soát code review): bước giữ giá trị cột "mồ côi" TRƯỚC ĐÂY đối chiếu dòng CŨ↔MỚI THEO VỊ
  // TRÍ (oldValuesByRow[i] ứng với data[i]) — SAI khi "data" (mảng trong bộ nhớ) đã bị xoá/thêm dòng Ở
  // GIỮA trước khi lưu (VD deleteProgram() gọi DB.programs.splice(idx,1)): mọi dòng SAU điểm xoá bị lệch
  // vị trí đi 1 so với thứ tự THẬT trên Sheet (chưa ghi), khiến orphan-column của các dòng đó bị gán
  // NHẦM sang dòng liền kề — corrupt âm thầm đúng cột đang cố bảo vệ. Nay đối chiếu theo CỘT NHẬN DẠNG ỔN
  // ĐỊNH (thử lần lượt 'id' rồi 'mamh' — đủ dùng cho mọi sheet có orphan-column thực tế: PROGRAMS/
  // DIEUKIEN/DOANHTHU/KEHOACH/KEHOACH_CHITIET/CHUONGTRINH/LUYKE dùng 'id', riêng MHCK dùng 'mamh') thay
  // vì vị trí — chỉ khi KHÔNG có cột nhận dạng nào dùng được mới rơi về đối chiếu theo vị trí như cũ.
  const KEY_CANDIDATES_ = ['id', 'mamh'];
  let keyColIdx = -1, keyName = null;
  for (const k of KEY_CANDIDATES_) {
    const ci = existingHeaders.indexOf(k);
    if (ci > -1 && data.every(r => r && r[k] !== '' && r[k] != null)) { keyColIdx = ci; keyName = k; break; }
  }
  let oldValuesByRow = null, oldValuesByKey = null;
  if (Object.keys(orphanCols).length && existingRowCount > 0) {
    oldValuesByRow = sh.getRange(2, 1, existingRowCount, existingHeaders.length).getValues();
    if (keyName) {
      oldValuesByKey = {};
      oldValuesByRow.forEach(row => { const kv = row[keyColIdx]; if (kv !== '' && kv != null) oldValuesByKey[kv] = row; });
    }
  }
  const values = data.map((r, i) => headers.map(h => {
    if (orphanCols[h]) {
      const ci = existingHeaders.indexOf(h);
      if (ci === -1) return '';
      if (keyName && oldValuesByKey && Object.prototype.hasOwnProperty.call(oldValuesByKey, r[keyName])) {
        return oldValuesByKey[r[keyName]][ci];
      }
      if (oldValuesByRow && i < oldValuesByRow.length) return oldValuesByRow[i][ci];
      return '';
    }
    return r[h] ?? '';
  }));
  writeRowsEfficient_(sh, headers, values);
}
// Dùng chung cho saveSheetData(): hội (union) toàn bộ tên cột — ưu tiên giữ đúng thứ tự cột ĐÃ CÓ SẴN
// trên Sheet trước (existingHeaders, nếu có) rồi mới nối thêm các cột MỚI phát hiện được ở bất kỳ dòng
// nào trong "rows" mà chưa từng là cột — đảm bảo không dòng/trường nào bị âm thầm bỏ sót khi ghi.
function unionHeaders_(existingHeaders, rows) {
  const headers = [];
  const seen = new Set();
  (existingHeaders || []).forEach(h => { if (h !== '' && !seen.has(h)) { seen.add(h); headers.push(h); } });
  (rows || []).forEach(r => {
    if (!r) return;
    Object.keys(r).forEach(k => { if (!seen.has(k)) { seen.add(k); headers.push(k); } });
  });
  return headers;
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
function saveGiaCongBoData(data, forceEmpty) { return simpleSave_('GIACONGBO', data, forceEmpty); }
// TỐI ƯU (gộp code trùng lặp): các hàm save*Data() dưới đây trước đây là các dòng lặp lại gần như
// giống hệt nhau (saveSheetData(SHEETS.X, data); return bumpDataVersion_();) — nay dùng chung 1 hàm
// simpleSave_(sheetKey, data) để giảm trùng lặp, tránh sai sót khi copy-paste thêm hàm mới sau này.
// Vẫn giữ mỗi hàm là 1 "function" khai báo riêng ở phạm vi toàn cục (không dùng const/arrow) để đảm
// bảo google.script.run từ Frontend luôn gọi được đúng tên hàm.
// NAY: bọc thêm withLock_() để chống 2 người dùng lưu cùng lúc ghi đè mất dữ liệu của nhau (xem chi
// tiết giải thích tại định nghĩa withLock_ phía trên).
function simpleSave_(sheetKey, data, forceEmpty) {
  return withLock_(() => { saveSheetData(SHEETS[sheetKey], data, forceEmpty); return bumpDataVersion_(); });
}
function saveProgramsData(data) { return simpleSave_('PROGRAMS', data); }
function savePurchasesData(data, forceEmpty) { return simpleSave_('PURCHASES', data, forceEmpty); }
function saveLuyKeData(data) { return simpleSave_('LUYKE', data); }
// THEO MÔ HÌNH 3 TẦNG: Tầng 1 (Danh mục chương trình) và Tầng 3 (Bậc điều kiện con của mã chiết khấu)
function saveChuongTrinhData(data) { return simpleSave_('CHUONGTRINH', data); }
function saveDieuKienData(data) { return simpleSave_('DIEUKIEN', data); }
// MỚI: Danh mục tổng hợp doanh thu sản lượng (Mã doanh thu) + Danh mục kế hoạch doanh thu
function saveDoanhThuData(data) { return simpleSave_('DOANHTHU', data); }
function saveKeHoachData(data) { return simpleSave_('KEHOACH', data); }
function saveKeHoachChiTietData(data) { return simpleSave_('KEHOACH_CHITIET', data); }

// ===== GHI THEO GÓI (chunked save) — SỬA LỖI (rà soát theo phản ánh người dùng "nhập lại không ghi vào
// Sheet, làm mới về dữ liệu cũ"): các bảng LỚN và TĂNG DẦN theo thời gian (Sổ mua hàng — mỗi tháng nạp
// thêm hàng trăm hóa đơn; Giá công bố NCC) trước đây được lưu bằng savePurchasesData()/saveGiaCongBoData()
// — gửi TOÀN BỘ mảng dữ liệu trong 1 lượt google.script.run DUY NHẤT. Đây ĐÚNG "hiện tượng đã biết" mà
// chính code này đã ghi chú khi vá cho việc ĐỌC (xem loadInChunks_/callChunkFnRetry_ ở Frontend, và ghi
// chú tại readSheetChunk_ phía trên): 1 gói dữ liệu ĐỦ LỚN truyền qua kênh giao tiếp iframe của Apps
// Script có thể bị hỏng/rỗng/không tới nơi mà KHÔNG báo lỗi rõ ràng. Trước đây chỉ vá cho ĐỌC, chưa vá
// cho GHI — khi 2 sheet này đã tích lũy đủ lớn theo thời gian sử dụng thực tế, việc LƯU (không phải chỉ
// tải) cũng có thể gặp lại đúng hiện tượng này: Frontend tưởng đã lưu xong (hoặc thậm chí báo lỗi mơ hồ),
// nhưng Sheet thực ra KHÔNG nhận đủ dữ liệu — "làm mới" (tải lại thẳng từ Sheet, bỏ qua cache) sẽ lộ ra
// đúng hiện tượng "về lại dữ liệu cũ" người dùng phản ánh.
// NAY: chia thành 3 bước gọi từ Frontend (xem saveToSheetChunked_ ở Index.html) — mỗi gói dữ liệu luôn
// nhỏ (mặc định 300 dòng, giống hệt kích thước gói đã dùng ổn định cho việc ĐỌC), nên không gặp lại hiện
// tượng trên:
//  1) beginChunkedSave_(sheetKey, headers, totalRows): ghi đúng dòng tiêu đề, dọn phần THỪA còn sót lại
//     từ dữ liệu CŨ nếu dữ liệu mới ít dòng/cột hơn (không clearContents() toàn bộ — giữ đúng triết lý
//     "ghi đè trực tiếp, không xoá sạch rồi ghi lại" đã áp dụng cho writeRowsEfficient_ ở trên).
//  2) writeChunkedSaveChunk_(sheetKey, offset, values): ghi ĐÚNG 1 gói nhỏ vào đúng vị trí dòng của nó.
//  3) finishChunkedSave_(): tăng phiên bản dữ liệu sau khi TẤT CẢ gói đã ghi xong.
// Mỗi bước đều bọc withLock_() như các hàm lưu khác — LƯU Ý: khoá chỉ giữ trong PHẠM VI TỪNG LƯỢT GỌI
// (không giữ xuyên suốt cả 3 bước, vì LockService không cho giữ khoá qua nhiều lượt thực thi tách biệt).
// Với đặc thù ứng dụng hiện chỉ 1 kế toán NPP sử dụng, rủi ro ghi xen kẽ giữa các bước là rất thấp.
function beginChunkedSave_(sheetKey, headers, totalRows) {
  return withLock_(() => {
    const ss = SpreadsheetApp.getActive();
    const sh = getOrCreateSheet_(ss, SHEETS[sheetKey]);
    if (!headers || !headers.length) { sh.clearContents(); return 'ok'; }
    const oldLastRow = sh.getLastRow();
    const oldLastCol = sh.getLastColumn();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    // Dọn phần dòng thừa còn sót lại từ dữ liệu CŨ nếu dữ liệu mới ít dòng hơn (như writeRowsEfficient_)
    if (oldLastRow > totalRows + 1) {
      sh.getRange(totalRows + 2, 1, oldLastRow - (totalRows + 1), Math.max(oldLastCol, headers.length)).clearContent();
    }
    // Dọn phần cột thừa còn sót lại nếu dữ liệu mới ít cột hơn
    if (oldLastCol > headers.length) {
      sh.getRange(1, headers.length + 1, Math.max(totalRows + 1, 1), oldLastCol - headers.length).clearContent();
    }
    return 'ok';
  });
}
function writeChunkedSaveChunk_(sheetKey, offset, values) {
  return withLock_(() => {
    if (!values || !values.length) return 'ok';
    const ss = SpreadsheetApp.getActive();
    const sh = getOrCreateSheet_(ss, SHEETS[sheetKey]);
    sh.getRange(2 + offset, 1, values.length, values[0].length).setValues(values);
    return 'ok';
  });
}
function finishChunkedSave_() {
  return withLock_(() => bumpDataVersion_());
}
// Dùng để XÁC MINH lại (từ Frontend, cho MỌI kiểu lưu — cả gói lẫn thường) rằng số dòng THỰC SỰ có trên
// Sheet sau khi lưu đúng bằng số dòng Frontend định gửi — nếu lệch, Frontend sẽ báo lỗi rõ ràng thay vì
// im lặng coi là đã lưu thành công (xem giải thích đầy đủ tại saveToSheet_ ở Index.html).
function getSheetRowCountForVerify_(sheetKey) {
  return getSheetRowCount_(sheetKey).total;
}

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
    headers = [];
    existingRows = [];
  }
  // SỬA LỖI (28/08/2026 — cùng lớp lỗi vừa vá ở saveSheetData()/unionHeaders_): trước đây chỉ soi
  // Object.keys(rows[0]) — DÒNG ĐẦU TIÊN của lô dữ liệu MỚI đang thêm — để phát hiện cột mới. Nếu lô
  // đang thêm có NHIỀU dòng và dòng đầu tiên (rows[0]) lại thiếu 1 trường mà các dòng SAU nó có, trường
  // đó vẫn bị bỏ sót. Nay hội (union) toàn bộ khoá xuất hiện ở BẤT KỲ dòng nào trong "rows".
  headers = unionHeaders_(headers, rows);
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
// ============================ BỘ TEST CÔNG THỨC (chuyển ra khỏi Index.html, 06/09/2026) ============================
// TỐI ƯU (rà soát tốc độ tải): runFormulaTests() (~50KB mã nguồn, ~690 dòng) trước đây nhúng thẳng trong
// Index.html — bị trình duyệt tải/parse MỖI LẦN MỞ APP dù chỉ thực sự chạy khi người dùng chủ động bấm
// nút debug "🧪 Test công thức chiết khấu" ở Trợ giúp. Đây đúng lớp vấn đề đã từng gây "trắng màn hình
// webapp" khi Index.html vượt ngưỡng ~1MiB (xem ghi chú tại getMauNhapLieuXlsxB64 phía trên) — tại thời
// điểm rà soát này Index.html đã LẠI vượt đúng ngưỡng đó (~1,09MB). Nay chuyển mã nguồn bộ test sang đây
// dưới dạng base64 tĩnh; Frontend chỉ gọi getFormulaTestsSrc() tải về + chạy (new Function(...)) đúng lúc
// người dùng bấm nút test lần đầu (có cache lại, các lần bấm sau không cần tải lại) — không còn tính vào
// gói tải ban đầu của mọi người dùng khi mở app.
const FORMULA_TESTS_SRC_B64_ = "ICBjb25zdCByZXN1bHRzID0gW107CiAgZnVuY3Rpb24gdGVzdChuYW1lLCBmbil7CiAgICB0cnl7IGZuKCk7IHJlc3VsdHMucHVzaCh7IG5hbWUsIG9rOnRydWUgfSk7IH0KICAgIGNhdGNoKGUpeyByZXN1bHRzLnB1c2goeyBuYW1lLCBvazpmYWxzZSwgZXJyOiBlLm1lc3NhZ2V8fFN0cmluZyhlKSB9KTsgfQogIH0KICBmdW5jdGlvbiBhc3NlcnRFcXVhbChhY3R1YWwsIGV4cGVjdGVkLCBsYWJlbCl7CiAgICBpZihhY3R1YWwgIT09IGV4cGVjdGVkKSB0aHJvdyBuZXcgRXJyb3IoYCR7bGFiZWx8fCcnfSBr4buzIHbhu41uZyAke0pTT04uc3RyaW5naWZ5KGV4cGVjdGVkKX0sIHRo4buxYyB04bq/ICR7SlNPTi5zdHJpbmdpZnkoYWN0dWFsKX1gKTsKICB9CiAgZnVuY3Rpb24gYXNzZXJ0Q2xvc2UoYWN0dWFsLCBleHBlY3RlZCwgbGFiZWwsIGVwcyl7CiAgICBlcHMgPSBlcHM9PW51bGw/MC4wMDE6ZXBzOwogICAgaWYoTWF0aC5hYnMoYWN0dWFsLWV4cGVjdGVkKSA+IGVwcykgdGhyb3cgbmV3IEVycm9yKGAke2xhYmVsfHwnJ30ga+G7syB24buNbmcgfiR7ZXhwZWN0ZWR9LCB0aOG7sWMgdOG6vyAke2FjdHVhbH1gKTsKICB9CiAgZnVuY3Rpb24gYXNzZXJ0VHJ1ZShjb25kLCBsYWJlbCl7IGlmKCFjb25kKSB0aHJvdyBuZXcgRXJyb3IobGFiZWx8fCfEkGnhu4F1IGtp4buHbiBwaOG6o2kgxJHDum5nICh0cnVlKScpOyB9CgogIC8qIC0tLS0gZGltc01hdGNoIC0tLS0gKi8KICBjb25zdCBhbGxObyA9IHt0aHVvbmdoaWV1OidOTycsbmhvbWtoOidOTycsbG9haXNwOidOTycsZGFjdGluaDonTk8nLGNvbmdkdW5nOidOTycscXV5Y2FjaDonTk8nfTsKICB0ZXN0KCdkaW1zTWF0Y2g6IGNoxrDGoW5nIHRyw6xuaCBjaOG7jW4gTk8g4bufIDEgY+G7mXQgLT4gYuG7jyBxdWEsIGx1w7RuIGto4bubcCBj4buZdCDEkcOzJywgKCk9PnsKICAgIGFzc2VydFRydWUoZGltc01hdGNoKHsuLi5hbGxObywgdGh1b25naGlldTonQ0FTVSd9LCBhbGxObykpOwogIH0pOwogIHRlc3QoJ2RpbXNNYXRjaDoga2jhu5twIMSRw7puZyBnacOhIHRy4buLIGPhu6UgdGjhu4MnLCAoKT0+ewogICAgYXNzZXJ0VHJ1ZShkaW1zTWF0Y2goey4uLmFsbE5vLCB0aHVvbmdoaWV1OidDQVNVJ30sIHsuLi5hbGxObywgdGh1b25naGlldTonQ0FTVSd9KSk7CiAgfSk7CiAgdGVzdCgnZGltc01hdGNoOiBraMO0bmcga2jhu5twIGtoaSBnacOhIHRy4buLIGtow6FjIG5oYXUnLCAoKT0+ewogICAgYXNzZXJ0VHJ1ZSghZGltc01hdGNoKHsuLi5hbGxObywgdGh1b25naGlldTonQ0FTVSd9LCB7Li4uYWxsTm8sIHRodW9uZ2hpZXU6J0JSSURHRVNUT05FJ30pKTsKICB9KTsKICB0ZXN0KCdkaW1zTWF0Y2g6IG3DoyBo4buXbiBo4bujcCAiVEJSLEJJQVMiIGto4bubcCBu4bq/dSBz4bqjbiBwaOG6qW0gdGh14buZYyAxIHRyb25nIDIgbcOjJywgKCk9PnsKICAgIGFzc2VydFRydWUoZGltc01hdGNoKHsuLi5hbGxObywgbG9haXNwOidUQlInfSwgey4uLmFsbE5vLCBsb2Fpc3A6J1RCUixCSUFTJ30pKTsKICAgIGFzc2VydFRydWUoZGltc01hdGNoKHsuLi5hbGxObywgbG9haXNwOidCSUFTJ30sIHsuLi5hbGxObywgbG9haXNwOidUQlIsQklBUyd9KSk7CiAgICBhc3NlcnRUcnVlKCFkaW1zTWF0Y2goey4uLmFsbE5vLCBsb2Fpc3A6J1BDUid9LCB7Li4uYWxsTm8sIGxvYWlzcDonVEJSLEJJQVMnfSkpOwogIH0pOwoKICAvKiAtLS0tIHBlcmlvZE92ZXJsYXAgLS0tLSAqLwogIHRlc3QoJ3BlcmlvZE92ZXJsYXA6IDIga2hv4bqjbmcgY8OzIGdpYW8gbmhhdSAtPiB0cnVlJywgKCk9PnsKICAgIGFzc2VydFRydWUocGVyaW9kT3ZlcmxhcCgnMjAyNi0wMS0wMScsJzIwMjYtMDEtMzEnLCcyMDI2LTAxLTE1JywnMjAyNi0wMi0xNScpKTsKICB9KTsKICB0ZXN0KCdwZXJpb2RPdmVybGFwOiAyIGtob+G6o25nIHTDoWNoIHLhu51pIC0+IGZhbHNlJywgKCk9PnsKICAgIGFzc2VydFRydWUoIXBlcmlvZE92ZXJsYXAoJzIwMjYtMDEtMDEnLCcyMDI2LTAxLTMxJywnMjAyNi0wMi0wMScsJzIwMjYtMDItMjgnKSk7CiAgfSk7CiAgdGVzdCgncGVyaW9kT3ZlcmxhcDogaGnhu4d1IGzhu7FjIGNoxrDGoW5nIHRyw6xuaCDEkeG7gyB0cuG7kW5nIC0+IGx1w7RuIGNvaSBsw6AgY8OybiBoaeG7h3UgbOG7sWMnLCAoKT0+ewogICAgYXNzZXJ0VHJ1ZShwZXJpb2RPdmVybGFwKCcyMDI2LTAxLTAxJywnMjAyNi0wMS0zMScsJycsJycpKTsKICB9KTsKCiAgLyogLS0tLSBpbkJyYWNrZXQgLS0tLSAqLwogIHRlc3QoJ2luQnJhY2tldDogbWF4PTAgbmdoxKlhIGzDoCBLSMOUTkcgZ2nhu5tpIGjhuqFuIHRyw6puIChxdXkgxrDhu5tjIHJpw6puZyBj4bunYSBhcHApJywgKCk9PnsKICAgIGFzc2VydFRydWUoaW5CcmFja2V0KDk5OTk5OTk5OSwgMTAwLCAwKSk7CiAgfSk7CiAgdGVzdCgnaW5CcmFja2V0OiBtaW4vbWF4IMSR4buDIHRy4buRbmcgLT4ga2jDtG5nIGdp4bubaSBo4bqhbiBj4bqjIDIgcGjDrWEnLCAoKT0+ewogICAgYXNzZXJ0VHJ1ZShpbkJyYWNrZXQoLTUwMCwgJycsICcnKSk7CiAgfSk7CiAgdGVzdCgnaW5CcmFja2V0OiDEkcO6bmcgdOG6oWkgMiBiacOqbiBtaW4vbWF4LCBzYWkgbmdheSBuZ2/DoGkgYmnDqm4nLCAoKT0+ewogICAgYXNzZXJ0VHJ1ZShpbkJyYWNrZXQoMTAwLCAxMDAsIDIwMCkgJiYgaW5CcmFja2V0KDIwMCwgMTAwLCAyMDApKTsKICAgIGFzc2VydFRydWUoIWluQnJhY2tldCg5OS45OSwgMTAwLCAyMDApICYmICFpbkJyYWNrZXQoMjAwLjAxLCAxMDAsIDIwMCkpOwogIH0pOwoKICAvKiAtLS0tIGNoZWNrU2xEdEJyYWNrZXQ6IGxvZ2ljIEFORC9PUiBnaeG7r2EgYuG6rWMgU+G6o24gbMaw4bujbmcgdsOgIERvYW5oIHRodSAtLS0tICovCiAgdGVzdCgnY2hlY2tTbER0QnJhY2tldDogQU5EIC0+IHBo4bqjaSDEkeG6oXQgQ+G6oiAyIMSRaeG7gXUga2nhu4duJywgKCk9PnsKICAgIGNvbnN0IHByb2cgPSB7c2xfbWluOjEwMCwgc2xfbWF4OjAsIGR0X21pbjoxMDAwMDAwLCBkdF9tYXg6MCwgZGlldWtpZW5fc2xkdDonQU5EJ307CiAgICBhc3NlcnRUcnVlKGNoZWNrU2xEdEJyYWNrZXQocHJvZywgMTUwLCAyMDAwMDAwKS5kYXQpOwogICAgYXNzZXJ0VHJ1ZSghY2hlY2tTbER0QnJhY2tldChwcm9nLCAxNTAsIDUwMDAwMCkuZGF0LCAnRFQgY2jGsGEgxJHhuqF0IHRow6wgQU5EIHBo4bqjaSBLSMOUTkcgxJHhuqF0Jyk7CiAgfSk7CiAgdGVzdCgnY2hlY2tTbER0QnJhY2tldDogT1IgLT4gY2jhu4kgY+G6p24gxJHhuqF0IDEgVFJPTkcgMicsICgpPT57CiAgICBjb25zdCBwcm9nID0ge3NsX21pbjoxMDAsIHNsX21heDowLCBkdF9taW46MTAwMDAwMCwgZHRfbWF4OjAsIGRpZXVraWVuX3NsZHQ6J09SJ307CiAgICBhc3NlcnRUcnVlKGNoZWNrU2xEdEJyYWNrZXQocHJvZywgMTUwLCA1MDAwMDApLmRhdCwgJ1NMIMSR4bqhdCB0aMOsIE9SIHBo4bqjaSDEkeG6oXQgZMO5IERUIGNoxrBhIMSR4bqhdCcpOwogIH0pOwoKICAvKiAtLS0tIGNhbGNEaXNjb3VudEZvclByb2dyYW06IDMgaMOsbmggdGjhu6ljIGNoaeG6v3Qga2jhuqV1IC0tLS0gKi8KICB0ZXN0KCdjYWxjRGlzY291bnRGb3JQcm9ncmFtOiAlIHRoZW8gZ2nDoSBtdWEgKFBFUkNFTlRfR0lBTVVBKScsICgpPT57CiAgICBhc3NlcnRDbG9zZShjYWxjRGlzY291bnRGb3JQcm9ncmFtKHtodF9jaGlldGtoYXU6J1BFUkNFTlRfR0lBTVVBJywgdGxfY2s6NX0sIDEwMCwgMTAwMDAwMCwgMCksIDUwMDAwKTsKICB9KTsKICB0ZXN0KCdjYWxjRGlzY291bnRGb3JQcm9ncmFtOiAlIHRoZW8gZ2nDoSBjw7RuZyBi4buRIChQRVJDRU5UX0dJQUNPTkdCTyknLCAoKT0+ewogICAgYXNzZXJ0Q2xvc2UoY2FsY0Rpc2NvdW50Rm9yUHJvZ3JhbSh7aHRfY2hpZXRraGF1OidQRVJDRU5UX0dJQUNPTkdCTycsIHRsX2NrOjEwfSwgMTAwLCAwLCAyMDAwMCksIDIwMDAwMCk7CiAgfSk7CiAgdGVzdCgnY2FsY0Rpc2NvdW50Rm9yUHJvZ3JhbTogxJHGoW4gZ2nDoSBj4buRIMSR4buLbmggKERPTkdJQV9DT0RJTkgpJywgKCk9PnsKICAgIGFzc2VydENsb3NlKGNhbGNEaXNjb3VudEZvclByb2dyYW0oe2h0X2NoaWV0a2hhdTonRE9OR0lBX0NPRElOSCcsIGRvbmdpYV9jb2Rpbmg6MTUwMH0sIDEwMCwgMCwgMCksIDE1MDAwMCk7CiAgfSk7CiAgdGVzdCgnY2FsY0Rpc2NvdW50Rm9yUHJvZ3JhbTogaMOsbmggdGjhu6ljIGtow7RuZyB4w6FjIMSR4buLbmggLT4gdHLhuqMgMMSRIChraMO0bmcgbOG7l2kpJywgKCk9PnsKICAgIGFzc2VydEVxdWFsKGNhbGNEaXNjb3VudEZvclByb2dyYW0oe2h0X2NoaWV0a2hhdTonS0hPTkdfVE9OX1RBSSd9LCAxMDAsIDEwMDAwMDAsIDApLCAwKTsKICB9KTsKICAvKiAtLS0tIGNhbGNBbW91bnRCeUh0XzogbMO1aSBkw7luZyBjaHVuZyBjaG8gQ0sgY2jDrW5oIEzhuqpOICJDaGnhur90IGto4bqldSBi4buVIHN1bmcgKGzFqXkga+G6vykiIC0tLS0gKi8KICB0ZXN0KCdjYWxjQW1vdW50QnlIdF86IFBFUkNFTlRfR0lBTVVBIMOhcCDEkcO6bmcgdOG7tyBs4buHIHTDuXkgw70gKGtow7RuZyBuaOG6pXQgdGhp4bq/dCB0bF9jaykgbMOqbiBnacOhIG11YScsICgpPT57CiAgICBhc3NlcnRDbG9zZShjYWxjQW1vdW50QnlIdF8oJ1BFUkNFTlRfR0lBTVVBJywgMTAwLCAxMDAwMDAwLCAwLCAwLjUvMTAwKSwgNTAwMCk7CiAgfSk7CiAgdGVzdCgnY2FsY0Ftb3VudEJ5SHRfOiBQRVJDRU5UX0dJQUNPTkdCTyDDoXAgxJHDum5nIHThu7cgbOG7hyB0w7l5IMO9IGzDqm4gU0zDl2dpw6EgY8O0bmcgYuG7kSBiw6xuaCBxdcOibiwgS0jDlE5HIGTDuW5nIGdpw6EgbXVhJywgKCk9PnsKICAgIGFzc2VydENsb3NlKGNhbGNBbW91bnRCeUh0XygnUEVSQ0VOVF9HSUFDT05HQk8nLCAxMDAsIDk5OTk5OTk5OSwgMjAwMDAsIDAuNS8xMDApLCAxMDAwMCk7CiAgfSk7CgogIC8qIC0tLS0gZ2lhTXVhQmluaFF1YW4gLS0tLSAqLwogIHRlc3QoJ2dpYU11YUJpbmhRdWFuOiBTTD0wIC0+IHRy4bqjIDAgKGtow7RuZyBjaGlhIGNobyAwIHJhIGzhu5dpL05hTiknLCAoKT0+ewogICAgYXNzZXJ0RXF1YWwoZ2lhTXVhQmluaFF1YW4oMCwgNTAwMDAwKSwgMCk7CiAgfSk7CiAgdGVzdCgnZ2lhTXVhQmluaFF1YW46IHTDrW5oIMSRw7puZyBnacOhIGLDrG5oIHF1w6JuID0gRFQvU0wnLCAoKT0+ewogICAgYXNzZXJ0Q2xvc2UoZ2lhTXVhQmluaFF1YW4oMTAwLCAxMDAwMDAwKSwgMTAwMDApOwogIH0pOwoKICAvKiAtLS0tIHJlc29sdmVFZmZlY3RpdmVCYWNfOiBtw7QgaMOsbmggMyB04bqnbmcsIGLhuq1jIMSRaeG7gXUga2nhu4duIGNvbiAoREIuZGlldWtpZW4pIC0tLS0gKi8KICBjb25zdCBkaWV1a2llbkJhY2t1cCA9IERCLmRpZXVraWVuOwogIHRyeXsKICAgIHRlc3QoJ3Jlc29sdmVFZmZlY3RpdmVCYWNfOiBraMO0bmcgY8OzIGLhuq1jIGNvbiAtPiBkw7luZyB0aOG6s25nIGNow61uaCBDVCAodMawxqFuZyB0aMOtY2ggbmfGsOG7o2MpJywgKCk9PnsKICAgICAgREIuZGlldWtpZW4gPSBbXTsKICAgICAgY29uc3QgcHJvZyA9IHtpZDonUDEnLCBzbF9taW46MCwgc2xfbWF4OjAsIGR0X21pbjowLCBkdF9tYXg6MCwgdGxfY2s6NX07CiAgICAgIGNvbnN0IHIgPSByZXNvbHZlRWZmZWN0aXZlQmFjXyhwcm9nLCAxMDAsIDEwMCk7CiAgICAgIGFzc2VydFRydWUoIXIubm9CYWNNYXRjaGVkKTsKICAgICAgYXNzZXJ0RXF1YWwoci5lZmZlY3RpdmUsIHByb2cpOwogICAgfSk7CiAgICB0ZXN0KCdyZXNvbHZlRWZmZWN0aXZlQmFjXzogY8OzIGLhuq1jIGNvbiwgY2jhu41uIMSRw7puZyBi4bqtYyBraOG7m3AgU0wgdGjhu7FjIHThur8nLCAoKT0+ewogICAgICBEQi5kaWV1a2llbiA9IFsKICAgICAgICB7IGlkOidEMScsIHByb2dyYW1faWQ6J1AxJywgc2xfbWluOjAsIHNsX21heDo5OSwgZHRfbWluOjAsIGR0X21heDowLCB0bF9jazozIH0sCiAgICAgICAgeyBpZDonRDInLCBwcm9ncmFtX2lkOidQMScsIHNsX21pbjoxMDAsIHNsX21heDowLCBkdF9taW46MCwgZHRfbWF4OjAsIHRsX2NrOjcgfQogICAgICBdOwogICAgICBjb25zdCByID0gcmVzb2x2ZUVmZmVjdGl2ZUJhY18oe2lkOidQMScsIHRsX2NrOjB9LCAxNTAsIDApOwogICAgICBhc3NlcnRFcXVhbChyLm1hdGNoZWRCYWMuaWQsICdEMicpOwogICAgICBhc3NlcnRFcXVhbChyLmVmZmVjdGl2ZS50bF9jaywgNyk7CiAgICB9KTsKICAgIHRlc3QoJ3Jlc29sdmVFZmZlY3RpdmVCYWNfOiBuaGnhu4F1IGLhuq1jIGNo4buTbmcgbOG6pW4gLT4gxrB1IHRpw6puIHThu7cgbOG7hyBDSyBDQU8gTkjhuqRUJywgKCk9PnsKICAgICAgREIuZGlldWtpZW4gPSBbCiAgICAgICAgeyBpZDonRDEnLCBwcm9ncmFtX2lkOidQMScsIHNsX21pbjowLCBzbF9tYXg6MCwgZHRfbWluOjAsIGR0X21heDowLCB0bF9jazozIH0sCiAgICAgICAgeyBpZDonRDInLCBwcm9ncmFtX2lkOidQMScsIHNsX21pbjowLCBzbF9tYXg6MCwgZHRfbWluOjAsIGR0X21heDowLCB0bF9jazo5IH0KICAgICAgXTsKICAgICAgYXNzZXJ0RXF1YWwocmVzb2x2ZUVmZmVjdGl2ZUJhY18oe2lkOidQMSd9LCAxMCwgMTApLm1hdGNoZWRCYWMuaWQsICdEMicpOwogICAgfSk7CiAgICB0ZXN0KCdyZXNvbHZlRWZmZWN0aXZlQmFjXzogY8OzIGLhuq1jIGNvbiBuaMawbmcgS0jDlE5HIGLhuq1jIG7DoG8ga2jhu5twIC0+IG5vQmFjTWF0Y2hlZD10cnVlJywgKCk9PnsKICAgICAgREIuZGlldWtpZW4gPSBbeyBpZDonRDEnLCBwcm9ncmFtX2lkOidQMScsIHNsX21pbjoxMDAwLCBzbF9tYXg6MCwgZHRfbWluOjAsIGR0X21heDowLCB0bF9jazo1IH1dOwogICAgICBhc3NlcnRUcnVlKHJlc29sdmVFZmZlY3RpdmVCYWNfKHtpZDonUDEnfSwgMTAsIDEwKS5ub0JhY01hdGNoZWQpOwogICAgfSk7CgogICAgLyogLS0tLSBT4busQSBM4buWSSAyNC8wOC8yMDI2OiBi4bqtYyBjaGlhIFRIRU8gJUtIIChTTC9EVCBraMO0bmcgZ2nhu5tpIGjhuqFuIOG7nyBt4buNaSBi4bqtYykgcGjhuqNpIGNo4buNbiDEkMOaTkcgYuG6rWMKICAgICAgIGto4bubcCAlS0ggdGjhu7FjIHThur8g4oCUIEtIw5RORyDEkcaw4bujYyBsdcO0biBuaOG6o3kgbMOqbiBi4bqtYyAlY2FvIG5o4bqldCBy4buTaSBt4bubaSBraeG7g20gdHJhIG5nxrDhu6NjIChs4buXaSBjxakga2hp4bq/bgogICAgICAgTlBQIMSR4bqhdCAlS0gg4bufIGLhuq1jIGdp4buvYSBi4buLIGLDoW8gc2FpIGzDoCAiY2jGsGEgxJHhuqF0IikuIEThu68gbGnhu4d1IGLhuq1jIGzhuqV5IMSRw7puZyB0aGVvIE3hu6VjIDMgVEIgMDEwMSB0aOG7sWMgdOG6vwogICAgICAgxJFhbmcga2hhaSBiw6FvIHRyb25nIFNoZWV0ICg5MOKAkzEwMCXihpIyJTsgMTAw4oCTMTA1JeKGkjIsNSU7IDEwNeKAkzEwOSXihpIzJTsg4omlMTEwJeKGkjMsNSUpLiAtLS0tICovCiAgICB0ZXN0KCdyZXNvbHZlRWZmZWN0aXZlQmFjXzogYuG6rWMgY2hpYSBUSEVPICVLSCDigJQgJUtIPTk1JSAoYuG6rWMgZ2nhu69hLCBLSMOUTkcgcGjhuqNpIGNhbyBuaOG6pXQpIHBo4bqjaSBjaOG7jW4gxJHDum5nIELhuq1jIDEgPSAyJSwga2jDtG5nIG5o4bqjeSBsw6puIGLhuq1jIDMsNSUnLCAoKT0+ewogICAgICBEQi5kaWV1a2llbiA9IFsKICAgICAgICB7IGlkOidCMScsIHByb2dyYW1faWQ6J1AxJywgc2xfbWluOjAsIHNsX21heDowLCBkdF9taW46MCwgZHRfbWF4OjAsIHBjdF9raF9taW46OTAsIHBjdF9raF9tYXg6MTAwLCB0bF9jazoyLjAgfSwKICAgICAgICB7IGlkOidCMicsIHByb2dyYW1faWQ6J1AxJywgc2xfbWluOjAsIHNsX21heDowLCBkdF9taW46MCwgZHRfbWF4OjAsIHBjdF9raF9taW46MTAwLjAwMDAwMDEsIHBjdF9raF9tYXg6MTA1LCB0bF9jazoyLjUgfSwKICAgICAgICB7IGlkOidCMycsIHByb2dyYW1faWQ6J1AxJywgc2xfbWluOjAsIHNsX21heDowLCBkdF9taW46MCwgZHRfbWF4OjAsIHBjdF9raF9taW46MTA1LjAwMDAwMDEsIHBjdF9raF9tYXg6MTA5LCB0bF9jazozLjAgfSwKICAgICAgICB7IGlkOidCNCcsIHByb2dyYW1faWQ6J1AxJywgc2xfbWluOjAsIHNsX21heDowLCBkdF9taW46MCwgZHRfbWF4OjAsIHBjdF9raF9taW46MTEwLCBwY3Rfa2hfbWF4Om51bGwsIHRsX2NrOjMuNSB9LAogICAgICBdOwogICAgICBjb25zdCByID0gcmVzb2x2ZUVmZmVjdGl2ZUJhY18oe2lkOidQMSd9LCA1MDAsIDk1MDAwMDAwMCwgOTUpOyAvLyAlS0ggdGjhu7FjIHThur8gPSA5NSwgdHJ1eeG7gW4gcXVhIHRoYW0gc+G7kSB0aOG7qSA0CiAgICAgIGFzc2VydFRydWUoIXIubm9CYWNNYXRjaGVkLCAncGjhuqNpIGto4bubcCDEkcaw4bujYyAxIGLhuq1jICg5NSUgbuG6sW0gdHJvbmcgQuG6rWMgMTogOTAtMTAwJSknKTsKICAgICAgYXNzZXJ0RXF1YWwoci5tYXRjaGVkQmFjLmlkLCAnQjEnLCAncGjhuqNpIGNo4buNbiBC4bqtYyAxICg5MC0xMDAlKSwgS0jDlE5HIHBo4bqjaSBC4bqtYyA0IChi4bqtYyAlY2FvIG5o4bqldCknKTsKICAgICAgYXNzZXJ0RXF1YWwoci5lZmZlY3RpdmUudGxfY2ssIDIuMCk7CiAgICB9KTsKICAgIHRlc3QoJ3Jlc29sdmVFZmZlY3RpdmVCYWNfOiBi4bqtYyBjaGlhIFRIRU8gJUtIIOKAlCAlS0g9NDY3JSAodsaw4bujdCB4YSkgduG6q24gxJHDum5nIGNo4buNbiBC4bqtYyA0IGNhbyBuaOG6pXQnLCAoKT0+ewogICAgICBjb25zdCByID0gcmVzb2x2ZUVmZmVjdGl2ZUJhY18oe2lkOidQMSd9LCA1MDAsIDk1MDAwMDAwMCwgNDY3Ljc4KTsKICAgICAgYXNzZXJ0RXF1YWwoci5tYXRjaGVkQmFjLmlkLCAnQjQnKTsKICAgICAgYXNzZXJ0RXF1YWwoci5lZmZlY3RpdmUudGxfY2ssIDMuNSk7CiAgICB9KTsKICAgIHRlc3QoJ3Jlc29sdmVFZmZlY3RpdmVCYWNfOiBi4bqtYyBjaGlhIFRIRU8gJUtIIOKAlCBraMO0bmcgdHJ1eeG7gW4gJUtIIChob2FudGhhbmhGb3JTZWxlY3Rpb249bnVsbCkgLT4gR0nhu64gaMOgbmggdmkgY8WpIChsdcO0biBjaOG7jW4gYuG6rWMgJWNhbyBuaOG6pXQgdHLGsOG7m2MpJywgKCk9PnsKICAgICAgY29uc3QgciA9IHJlc29sdmVFZmZlY3RpdmVCYWNfKHtpZDonUDEnfSwgNTAwLCA5NTAwMDAwMDApOyAvLyBraMO0bmcgdHJ1eeG7gW4gdGhhbSBz4buRIDQsIHkgaOG7h3QgY29kZSBjxakgZ+G7jWkKICAgICAgYXNzZXJ0RXF1YWwoci5tYXRjaGVkQmFjLmlkLCAnQjQnLCAna2jDtG5nIGPDsyAlS0ggxJHhu4Mgc28gLT4gduG6q24gY2jhu41uIGLhuq1jIHRsX2NrIGNhbyBuaOG6pXQgbmjGsCBow6BuaCB2aSBjxakgKGtow7RuZyBwaMOhIHbhu6EgY2jGsMahbmcgdHLDrG5oIHRoaeG6v3UgS+G6vyBob+G6oWNoKScpOwogICAgfSk7CiAgICB0ZXN0KCdyZXNvbHZlRWZmZWN0aXZlQmFjXzogYuG6rWMgZOG6oW5nIE1BIFRS4bqsTiAyIENISeG7gFUgKHF1eSBtw7QgRFQgdGjhuq10ICsgJUtIIHRo4bqtdCwga2nhu4N1IE3hu6VjIDcgQ0sgbsSDbSkg4oCUIHBo4bqjaSBraOG7m3AgxJDDmk5HIGPhuqMgMiBjaGnhu4F1JywgKCk9PnsKICAgICAgREIuZGlldWtpZW4gPSBbCiAgICAgICAgeyBpZDonTV9BXzEwMCcsIHByb2dyYW1faWQ6J1AyJywgc2xfbWluOjAsIHNsX21heDowLCBkdF9taW46MCwgZHRfbWF4OjExOTk5OTk5OTk5LCBwY3Rfa2hfbWluOjk1LCBwY3Rfa2hfbWF4OjEwMCwgdGxfY2s6MS4wNSB9LAogICAgICAgIHsgaWQ6J01fQV8xMTAnLCBwcm9ncmFtX2lkOidQMicsIHNsX21pbjowLCBzbF9tYXg6MCwgZHRfbWluOjAsIGR0X21heDoxMTk5OTk5OTk5OSwgcGN0X2toX21pbjoxMDAuMDAwMDAwMSwgcGN0X2toX21heDoxMTAsIHRsX2NrOjEuMjAgfSwKICAgICAgICB7IGlkOidNX0JfMTAwJywgcHJvZ3JhbV9pZDonUDInLCBzbF9taW46MCwgc2xfbWF4OjAsIGR0X21pbjoxMjAwMDAwMDAwMCwgZHRfbWF4OjAsIHBjdF9raF9taW46OTUsIHBjdF9raF9tYXg6MTAwLCB0bF9jazoxLjM1IH0sCiAgICAgICAgeyBpZDonTV9CXzExMCcsIHByb2dyYW1faWQ6J1AyJywgc2xfbWluOjAsIHNsX21heDowLCBkdF9taW46MTIwMDAwMDAwMDAsIGR0X21heDowLCBwY3Rfa2hfbWluOjEwMC4wMDAwMDAxLCBwY3Rfa2hfbWF4OjExMCwgdGxfY2s6MS42NSB9LAogICAgICBdOwogICAgICAvLyBRdXkgbcO0IGRvYW5oIHPhu5EgbsSDbSA9IDIwIHThu7cgKHLGoWkgdsOgbyBuaMOzbSAiQiIsIGR0X21pbj0xMiB04bu3KSB2w6AgJUtIPTEwNSUgKHLGoWkgdsOgbyBi4bqtYyAiMTAwLjAwMDAwMDEtMTEwIikKICAgICAgLy8gPT4gcGjhuqNpIGNo4buNbiDEkcO6bmcgTV9CXzExMCAoMSw2NSUpLCBLSMOUTkcgxJHGsOG7o2MgY2jhu41uIE1fQV8xMTAvTV9CXzEwMCBkw7kgbeG7l2kgY8OhaSBraOG7m3AgxJDDmk5HIDEgdHJvbmcgMiBjaGnhu4F1LgogICAgICBjb25zdCByID0gcmVzb2x2ZUVmZmVjdGl2ZUJhY18oe2lkOidQMid9LCAxMDAwLCAyMDAwMDAwMDAwMCwgMTA1KTsKICAgICAgYXNzZXJ0RXF1YWwoci5tYXRjaGVkQmFjLmlkLCAnTV9CXzExMCcsICdwaOG6o2kga2jhu5twIMSRw7puZyBD4bqiIHF1eSBtw7QgKDIwIHThu7cg4omlIDEyIHThu7cpIEzhuqpOICVLSCAoMTA1JSB0cm9uZyAxMDAtMTEwJSknKTsKICAgICAgYXNzZXJ0RXF1YWwoci5lZmZlY3RpdmUudGxfY2ssIDEuNjUpOwogICAgfSk7CgogICAgLyogLS0tLSB0aW5oSG9hblRoYW5oS0hfOiAlS0ggcGjhuqNpIMawdSB0acOqbiB0w61uaCB0csOqbiBHScOBIEPDlE5HIELhu5AsIGtow7RuZyBwaOG6o2kgZ2nDoSBtdWEgdGjhu7FjIHThur8gLS0tLSAqLwogICAgdGVzdCgndGluaEhvYW5UaGFuaEtIXzogxJHDoyBjw7MgxJHhu6cgZ2nDoSBjw7RuZyBi4buRIGNobyB0b8OgbiBi4buZIFNMIC0+IGTDuW5nIGRvYW5oIHPhu5EgR0nDgSBDw5RORyBC4buQIChjYW8gaMahbiBnacOhIG11YSkgxJHhu4MgdMOtbmggJUtIJywgKCk9PnsKICAgICAgLy8gR2nDoSBtdWEgdGjhu7FjIHThur8gOTUwdHIgKMSRw6MgZ+G7k20gQ0sgY8ahIGLhuqNuIDEwJSkgbmjGsG5nIGdpw6EgY8O0bmcgYuG7kSB0xrDGoW5nIOG7qW5nIH4xLjA1NXRyICg5NTB0ci8wLDkpCiAgICAgIGNvbnN0IHIgPSB0aW5oSG9hblRoYW5oS0hfKDEwMDAwMDAwMDAsIDk1MDAwMDAwMCwgMTA1NTU1NTU1NiwgNTAwLCA1MDApOwogICAgICBhc3NlcnRDbG9zZShyLmhvYW50aGFuaCwgMTA1LjU1NTU1NTYsICclS0ggcGjhuqNpIHTDrW5oIHRyw6puIDEuMDU1LjU1NS41NTYgKGdpw6EgY8O0bmcgYuG7kSksIGtow7RuZyBwaOG6o2kgOTUwLjAwMC4wMDAgKGdpw6EgbXVhKScsIDAuMDEpOwogICAgICBhc3NlcnRUcnVlKCFyLmNhbmhCYW9HaWFDb25nQm8sICfEkcOjIMSR4bunIGdpw6EgY8O0bmcgYuG7kSBjaG8gdG/DoG4gYuG7mSBTTCAtPiBraMO0bmcgY+G6p24gY+G6o25oIGLDoW8nKTsKICAgIH0pOwogICAgdGVzdCgndGluaEhvYW5UaGFuaEtIXzogVEhJ4bq+VSBnacOhIGPDtG5nIGLhu5EgY2hvIDEgcGjhuqduIFNMIC0+IHThuqFtIGTDuW5nIGdpw6EgbXVhIHRo4buxYyB04bq/IFbDgCBiw6FvIGNhbmhCYW9HaWFDb25nQm89dHJ1ZScsICgpPT57CiAgICAgIGNvbnN0IHIgPSB0aW5oSG9hblRoYW5oS0hfKDEwMDAwMDAwMDAsIDk1MDAwMDAwMCwgNTAwMDAwMDAwLCAyMDAsIDUwMCk7IC8vIGNiU0w9MjAwIDwgdG90YWxTbD01MDAgLT4gdGhp4bq/dQogICAgICBhc3NlcnRDbG9zZShyLmhvYW50aGFuaCwgOTUsICclS0ggcGjhuqNpIHThuqFtIGTDuW5nIGRvYW5oIHPhu5EgZ2nDoSBtdWEgdGjhu7FjIHThur8gKDk1MHRyKSBraGkgZ2nDoSBjw7RuZyBi4buRIGNoxrBhIHBo4bunIGjhur90JywgMC4wMSk7CiAgICAgIGFzc2VydFRydWUoci5jYW5oQmFvR2lhQ29uZ0JvLCAncGjhuqNpIMSRw6FuaCBk4bqldSBj4bqjbmggYsOhbyB2w6wgJUtIIMSRYW5nIGzDoCBz4buRIGfhuqduIMSRw7puZyAodGhp4bq/dSBnacOhIGPDtG5nIGLhu5EpJyk7CiAgICB9KTsKICB9IGZpbmFsbHkgewogICAgREIuZGlldWtpZW4gPSBkaWV1a2llbkJhY2t1cDsKICB9CgogIC8qIC0tLS0gcmVzb2x2ZURvYW5oU29fL2RvYW5oU29CYXNpc186ICJDw6FjaCB0w61uaCBkb2FuaCBz4buRIiAoZ2nDoSBjw7RuZyBi4buRL2dpw6EgaMOzYSDEkcahbikga2hhaSB0csOqbiBNw6MKICAgICBkb2FuaCB0aHUsIGTDuW5nIENIVU5HIGNobyBj4bqjIHNvIGLhuq1jIFNML0RUIGzhuqtuIHTDrW5oICVLSCBj4bunYSBt4buNaSBNw6MgY2hp4bq/dCBraOG6pXUgZ+G6r24gTcOjIGRvYW5oIHRodSDEkcOzCiAgICAgKGLhu5Ugc3VuZyAyNC8wOC8yMDI2KSAtLS0tICovCiAgY29uc3QgZG9hbmh0aHVCYWNrdXAgPSBEQi5kb2FuaHRodTsKICB0cnl7CiAgICBEQi5kb2FuaHRodSA9IFsKICAgICAgeyBpZDonRFRfQ0InLCBtYTonRFRfQ0InLCBjYWNoX3RpbmhfZG9hbmhzbzonR0lBX0NPTkdCTycgfSwKICAgICAgeyBpZDonRFRfSEQnLCBtYTonRFRfSEQnLCBjYWNoX3RpbmhfZG9hbmhzbzonR0lBX0hPQURPTicgfSwKICAgICAgeyBpZDonRFRfQ1UnLCBtYTonRFRfQ1UnIH0gLy8gTcOjIGRvYW5oIHRodSBDxagsIGNoxrBhIGPDsyBmaWVsZCBjYWNoX3RpbmhfZG9hbmhzbyAoZOG7ryBsaeG7h3UgdHLGsOG7m2Mga2hpIGPDsyB0w61uaCBuxINuZyBuw6B5KQogICAgXTsKICAgIHRlc3QoJ2RvYW5oU29CYXNpc186IE3DoyBkb2FuaCB0aHUga2hhaSAiR0lBX0NPTkdCTyIgLT4gdHLhuqMgduG7gSBHSUFfQ09OR0JPJywgKCk9PnsKICAgICAgYXNzZXJ0RXF1YWwoZG9hbmhTb0Jhc2lzXygnRFRfQ0InKSwgJ0dJQV9DT05HQk8nKTsKICAgIH0pOwogICAgdGVzdCgnZG9hbmhTb0Jhc2lzXzogTcOjIGRvYW5oIHRodSBraGFpICJHSUFfSE9BRE9OIiAtPiB0cuG6oyB24buBIEdJQV9IT0FET04nLCAoKT0+ewogICAgICBhc3NlcnRFcXVhbChkb2FuaFNvQmFzaXNfKCdEVF9IRCcpLCAnR0lBX0hPQURPTicpOwogICAgfSk7CiAgICB0ZXN0KCdkb2FuaFNvQmFzaXNfOiBNw6MgZG9hbmggdGh1IEPFqCAoY2jGsGEgY8OzIGZpZWxkKSBob+G6t2Mga2jDtG5nIHTDrG0gdGjhuqV5IC0+IG3hurdjIMSR4buLbmggR0lBX0hPQURPTiAoS0jDlE5HIMOibSB0aOG6p20gxJHhu5VpIGjDoG5oIHZpIHTDrW5oIGPFqSknLCAoKT0+ewogICAgICBhc3NlcnRFcXVhbChkb2FuaFNvQmFzaXNfKCdEVF9DVScpLCAnR0lBX0hPQURPTicpOwogICAgICBhc3NlcnRFcXVhbChkb2FuaFNvQmFzaXNfKCdLSE9OR19UT05fVEFJJyksICdHSUFfSE9BRE9OJyk7CiAgICAgIGFzc2VydEVxdWFsKGRvYW5oU29CYXNpc18odW5kZWZpbmVkKSwgJ0dJQV9IT0FET04nKTsKICAgIH0pOwogICAgdGVzdCgncmVzb2x2ZURvYW5oU29fOiBNw6MgZG9hbmggdGh1ID0gR0lBX0NPTkdCTyArIMSR4bunIGdpw6EgY8O0bmcgYuG7kSB0b8OgbiBi4buZIFNMIC0+IGRvYW5oc28gbOG6pXkgdGhlbyBnacOhIGPDtG5nIGLhu5EnLCAoKT0+ewogICAgICBjb25zdCByID0gcmVzb2x2ZURvYW5oU29fKCdEVF9DQicsIDk1MDAwMDAwMCwgMTA1NTU1NTU1NiwgNTAwLCA1MDApOwogICAgICBhc3NlcnRDbG9zZShyLmRvYW5oc28sIDEwNTU1NTU1NTYsICdwaOG6o2kgbOG6pXkgY2JHVHN1bSAoZ2nDoSBjw7RuZyBi4buRKSwga2jDtG5nIHBo4bqjaSB0b3RhbER0IChnacOhIG11YSknKTsKICAgICAgYXNzZXJ0VHJ1ZSghci5jYW5oQmFvR2lhQ29uZ0JvKTsKICAgIH0pOwogICAgdGVzdCgncmVzb2x2ZURvYW5oU29fOiBNw6MgZG9hbmggdGh1ID0gR0lBX0NPTkdCTyBuaMawbmcgVEhJ4bq+VSBnacOhIGPDtG5nIGLhu5EgMSBwaOG6p24gU0wgLT4gdOG6oW0gZMO5bmcgdG90YWxEdCArIGPhuqNuaCBiw6FvJywgKCk9PnsKICAgICAgY29uc3QgciA9IHJlc29sdmVEb2FuaFNvXygnRFRfQ0InLCA5NTAwMDAwMDAsIDUwMDAwMDAwMCwgMjAwLCA1MDApOwogICAgICBhc3NlcnRDbG9zZShyLmRvYW5oc28sIDk1MDAwMDAwMCk7CiAgICAgIGFzc2VydFRydWUoci5jYW5oQmFvR2lhQ29uZ0JvKTsKICAgIH0pOwogICAgdGVzdCgncmVzb2x2ZURvYW5oU29fOiBNw6MgZG9hbmggdGh1ID0gR0lBX0hPQURPTiAtPiBMVcOUTiBkw7luZyB0b3RhbER0IChnacOhIG11YSB0aOG7sWMgdOG6vyksIGTDuSDEkcOjIGPDsyDEkeG7pyBnacOhIGPDtG5nIGLhu5EnLCAoKT0+ewogICAgICBjb25zdCByID0gcmVzb2x2ZURvYW5oU29fKCdEVF9IRCcsIDk1MDAwMDAwMCwgMTA1NTU1NTU1NiwgNTAwLCA1MDApOwogICAgICBhc3NlcnRDbG9zZShyLmRvYW5oc28sIDk1MDAwMDAwMCwgJ2toYWkgR0lBX0hPQURPTiB0aMOsIGtow7RuZyDEkcaw4bujYyBs4bqleSBnacOhIGPDtG5nIGLhu5EgZMO5IMSRw6MgxJHhu6cgZOG7ryBsaeG7h3UnKTsKICAgICAgYXNzZXJ0VHJ1ZSghci5jYW5oQmFvR2lhQ29uZ0JvLCAnR0lBX0hPQURPTiBraMO0bmcgY+G6p24gY+G6o25oIGLDoW8gdGhp4bq/dSBnacOhIGPDtG5nIGLhu5EgKGtow7RuZyBkw7luZyDEkeG6v24gZ2nDoSBjw7RuZyBi4buRKScpOwogICAgfSk7CiAgfSBmaW5hbGx5IHsKICAgIERCLmRvYW5odGh1ID0gZG9hbmh0aHVCYWNrdXA7CiAgfQoKICAvKiAtLS0tIHNuYXBQZXJpb2RGcm9tRGF0ZTogY2h14bqpbiBob8OhIHJhbmggZ2nhu5tpIGvhu7MgKG7hu4FuIHThuqNuZyBj4bunYSBjxqEgY2jhur8gY2jhu5FuZyB0csO5bmcgYsOhbyBjw6FvKSAtLS0tICovCiAgdGVzdCgnc25hcFBlcmlvZEZyb21EYXRlOiBUaMOhbmcgLT4gxJHDum5nIG5nw6B5IDEgdsOgIG5nw6B5IGN14buRaSB0aMOhbmcnLCAoKT0+ewogICAgY29uc3QgciA9IHNuYXBQZXJpb2RGcm9tRGF0ZSgnMjAyNi0wNy0xNScsICdUaMOhbmcnKTsKICAgIGFzc2VydEVxdWFsKHIuZnJvbSwgJzIwMjYtMDctMDEnKTsgYXNzZXJ0RXF1YWwoci50bywgJzIwMjYtMDctMzEnKTsKICB9KTsKICB0ZXN0KCdzbmFwUGVyaW9kRnJvbURhdGU6IFF1w70gLT4gxJHDum5nIHF1w70gY2jhu6lhIG5nw6B5IMSRw7MnLCAoKT0+ewogICAgY29uc3QgciA9IHNuYXBQZXJpb2RGcm9tRGF0ZSgnMjAyNi0wOC0xMCcsICdRdcO9Jyk7IC8vIHRow6FuZyA4IC0+IFF1w70gMyAoNyw4LDkpCiAgICBhc3NlcnRFcXVhbChyLmZyb20sICcyMDI2LTA3LTAxJyk7IGFzc2VydEVxdWFsKHIudG8sICcyMDI2LTA5LTMwJyk7CiAgfSk7CiAgdGVzdCgnc25hcFBlcmlvZEZyb21EYXRlOiAiVMO5eSBjaOG7jW4iIC0+IGtow7RuZyB04buxIHF1eSDEkeG7lWkgKHRy4bqjIG51bGwpJywgKCk9PnsKICAgIGFzc2VydEVxdWFsKHNuYXBQZXJpb2RGcm9tRGF0ZSgnMjAyNi0wNy0xNScsICcnKSwgbnVsbCk7CiAgfSk7CiAgdGVzdCgnc25hcFBlcmlvZEZyb21EYXRlOiBuZ8OgeSDEkeG6p3UgbsSDbS/EkeG6p3UgdGjDoW5nIChiacOqbiBuaOG6oXkgY+G6o20gbcO6aSBnaeG7nSkgduG6q24gxJHDum5nIG7Eg20vdGjDoW5nIMSRYW5nIHjDqXQnLCAoKT0+ewogICAgLy8gVHLGsOG7m2Mga2hpIHPhu61hLCBow6BtIHBhcnNlICIyMDI2LTAxLTAxIiBi4bqxbmcgbmV3IERhdGUoKSBy4buTaSDEkeG7jWMgbOG6oWkgZ2V0RnVsbFllYXIoKS9nZXRNb250aCgpIHRoZW8KICAgIC8vIGdp4budIMSQ4buKQSBQSMavxqBORyB0csOsbmggZHV54buHdCDigJQgduG7m2kgbcO6aSBnaeG7nSBsw7lpIHNhdSBVVEMsIG3hu5FjIFVUQyAwMDowMCBuw6B5IGPDsyB0aOG7gyBi4buLIMSR4buNYyBsw7lpIHbhu4EgbmfDoHkKICAgIC8vIDMxLzEyIG7Eg20gdHLGsOG7m2MsIGtoaeG6v24gQuG6rWMgVGjDoW5nL07Eg20gc25hcCBzYWkuIE5heSBkw7luZyB5ZWFyT2ZJU09fL21vbnRoT2ZJU09fIChj4bqvdCBjaHXhu5dpIHRy4buxYwogICAgLy8gdGnhur9wKSwga2jDtG5nIHBo4bulIHRodeG7mWMgbcO6aSBnaeG7nSB0csOsbmggZHV54buHdCDigJQgdGVzdCBuw6B5IGtob8OhIMSRw7puZyBow6BuaCB2aSB04bqhaSBjw6FjIG3hu5FjIGThu4UgbOG7mSBs4buXaSBuaOG6pXQuCiAgICBhc3NlcnRFcXVhbChzbmFwUGVyaW9kRnJvbURhdGUoJzIwMjYtMDEtMDEnLCAnVGjDoW5nJykuZnJvbSwgJzIwMjYtMDEtMDEnKTsKICAgIGFzc2VydEVxdWFsKHNuYXBQZXJpb2RGcm9tRGF0ZSgnMjAyNi0wMS0wMScsICdUaMOhbmcnKS50bywgJzIwMjYtMDEtMzEnKTsKICAgIGFzc2VydEVxdWFsKHNuYXBQZXJpb2RGcm9tRGF0ZSgnMjAyNi0wMS0wMScsICdOxINtJykuZnJvbSwgJzIwMjYtMDEtMDEnKTsKICAgIGFzc2VydEVxdWFsKHNuYXBQZXJpb2RGcm9tRGF0ZSgnMjAyNi0wMS0wMScsICdOxINtJykudG8sICcyMDI2LTEyLTMxJyk7CiAgfSk7CiAgdGVzdCgnc25hcFBlcmlvZEZyb21EYXRlOiBuZ8OgeSBy4buXbmcva2jDtG5nIGjhu6NwIGzhu4cgLT4gdHLhuqMgbnVsbCAoa2jDtG5nIE5hTi9jcmFzaCknLCAoKT0+ewogICAgYXNzZXJ0RXF1YWwoc25hcFBlcmlvZEZyb21EYXRlKCcnLCAnVGjDoW5nJyksIG51bGwpOwogICAgYXNzZXJ0RXF1YWwoc25hcFBlcmlvZEZyb21EYXRlKCdraG9uZy1waGFpLW5nYXknLCAnVGjDoW5nJyksIG51bGwpOwogIH0pOwoKICAvKiAtLS0tIGlzU2FuZVJlcG9ydFllYXJfOiBjaOG6t24gaOG6rXUgcXXhuqMgcXVpcmsgInPhu61hIG7Eg20gdHLDqm4gbOG7i2NoIHThu7EgbmjhuqN5IHJhIDE5MDYiIC0tLS0gKi8KICB0ZXN0KCdpc1NhbmVSZXBvcnRZZWFyXzogbsSDbSBo4bujcCBsw70gKDIwMTUtMjAzNSkgLT4gdHJ1ZScsICgpPT57CiAgICBhc3NlcnRUcnVlKGlzU2FuZVJlcG9ydFllYXJfKDIwMjYpKTsKICAgIGFzc2VydFRydWUoaXNTYW5lUmVwb3J0WWVhcl8oMjAxNSkpOwogICAgYXNzZXJ0VHJ1ZShpc1NhbmVSZXBvcnRZZWFyXygyMDM1KSk7CiAgfSk7CiAgdGVzdCgnaXNTYW5lUmVwb3J0WWVhcl86IG7Eg20gdsO0IGzDvSAoVkQgcXVpcmsgd2lkZ2V0IGzhu4tjaCByYSAxOTA2KSAtPiBmYWxzZScsICgpPT57CiAgICBhc3NlcnRUcnVlKCFpc1NhbmVSZXBvcnRZZWFyXygxOTA2KSk7CiAgICBhc3NlcnRUcnVlKCFpc1NhbmVSZXBvcnRZZWFyXygyMDE0KSk7CiAgICBhc3NlcnRUcnVlKCFpc1NhbmVSZXBvcnRZZWFyXygyMDM2KSk7CiAgICBhc3NlcnRUcnVlKCFpc1NhbmVSZXBvcnRZZWFyXyhOYU4pKTsKICB9KTsKCiAgLyogLS0tLSB1cHNlcnRMb2NhbDogZMO5bmcgY2hvIGPhu5luZyBk4buTbiBEQi5yZXBvcnRDS1RIL0NLQ1QgcXVhIG5oaeG7gXUgbOG6p24gbMawdSBiw6FvIGPDoW8gLS0tLSAqLwogIHRlc3QoJ3Vwc2VydExvY2FsOiBraG/DoSDEkcOjIGPDsyAtPiBHSMOJUCAobWVyZ2UpIHbDoG8gxJHDum5nIGTDsm5nIGPFqSwga2jDtG5nIG5ow6JuIMSRw7RpJywgKCk9PnsKICAgIGNvbnN0IGFyciA9IFt7IGs6J0EnLCB2OjEsIGdpdTonY3UnIH0sIHsgazonQicsIHY6MiB9XTsKICAgIHVwc2VydExvY2FsKGFyciwgW3sgazonQScsIHY6OTkgfV0sIHI9PnIuayk7CiAgICBhc3NlcnRFcXVhbChhcnIubGVuZ3RoLCAyLCAna2jDtG5nIMSRxrDhu6NjIHRow6ptIGTDsm5nIG3hu5tpIGtoaSBraG/DoSDEkcOjIHThu5NuIHThuqFpJyk7CiAgICBhc3NlcnRFcXVhbChhcnJbMF0udiwgOTksICdwaOG6o2kgbOG6pXkgZ2nDoSB0cuG7iyBN4buaSSBjaG8gdHLGsOG7nW5nIMSRw6MgZ+G7rWkgbMOqbicpOwogICAgYXNzZXJ0RXF1YWwoYXJyWzBdLmdpdSwgJ2N1JywgJ3Bo4bqjaSBHSeG7riBOR1VZw4pOIHRyxrDhu51uZyBraMO0bmcgbuG6sW0gdHJvbmcgZOG7ryBsaeG7h3UgbeG7m2kgKE9iamVjdC5hc3NpZ24gbWVyZ2UpJyk7CiAgfSk7CiAgdGVzdCgndXBzZXJ0TG9jYWw6IGtob8OhIGNoxrBhIGPDsyAtPiBUSMOKTSBN4buaSSB2w6BvIGN14buRaSBt4bqjbmcnLCAoKT0+ewogICAgY29uc3QgYXJyID0gW3sgazonQScsIHY6MSB9XTsKICAgIHVwc2VydExvY2FsKGFyciwgW3sgazonQycsIHY6MyB9XSwgcj0+ci5rKTsKICAgIGFzc2VydEVxdWFsKGFyci5sZW5ndGgsIDIpOwogICAgYXNzZXJ0RXF1YWwoYXJyWzFdLmssICdDJyk7CiAgfSk7CiAgdGVzdCgndXBzZXJ0TG9jYWw6IG5oaeG7gXUgZMOybmcgbeG7m2kgdHJvbmcgQ8OZTkcgMSBsxrDhu6N0IGfhu41pLCBjw7MgY+G6oyB0csO5bmcgbOG6q24gbeG7m2kgLT4geOG7rSBsw70gxJHDum5nIHThu6tuZyBkw7JuZycsICgpPT57CiAgICBjb25zdCBhcnIgPSBbeyBrOidBJywgdjoxIH1dOwogICAgdXBzZXJ0TG9jYWwoYXJyLCBbeyBrOidBJywgdjoxMCB9LCB7IGs6J0QnLCB2OjQgfSwgeyBrOidBJywgdjoyMCB9XSwgcj0+ci5rKTsKICAgIGFzc2VydEVxdWFsKGFyci5sZW5ndGgsIDIsICdBIGNo4buJIDEgZMOybmcgKGdow6lwIDIgbOG6p24gbGnDqm4gdGnhur9wKSwgRCB0aMOqbSBt4bubaSAtPiB04buVbmcgMiBkw7JuZycpOwogICAgYXNzZXJ0RXF1YWwoYXJyWzBdLnYsIDIwLCAnbMaw4bujdCBnaMOpcCBTQVUgKHY6MjApIHBo4bqjaSB0aOG6r25nIGzGsOG7o3QgZ2jDqXAgVFLGr+G7mkMgKHY6MTApIHRyb25nIGPDuW5nIDEgbOG6p24gZ+G7jWknKTsKICB9KTsKCiAgLyogLS0tLSBidWlsZEV4Y2x1ZGVkRnJvbVRvdGFsU2V0Xy9pc0V4Y2x1ZGVkRnJvbVRvdGFsXzogbcOjIGNoaeG6v3Qga2jhuqV1IMSRw6MgbuG6sW0gc+G6tW4gdHJvbmcgZ2nDoSAoVkQgIkNLIGPGoQogICAgIGLhuqNuIDEwJSIpIHBo4bqjaSBi4buLIGxv4bqhaSBraOG7j2kgY8OhYyBz4buRIFThu5RORyBDS1RNIG5oxrBuZyB24bqrbiB0cmEgY+G7qXUgcmnDqm5nIMSRxrDhu6NjIC0tLS0gKi8KICBjb25zdCBwcm9ncmFtc0JhY2t1cEV4Y2wgPSBEQi5wcm9ncmFtczsKICB0cnl7CiAgICBEQi5wcm9ncmFtcyA9IFsKICAgICAgeyBpZDonUFgxJywgbWFfY2hpZXRraGF1OidDS19DT19CQU5fMTBQVCcsIGtob25nX2dvcF90b25nOiB0cnVlIH0sCiAgICAgIHsgaWQ6J1BYMicsIG1hX2NoaWV0a2hhdTonQ0tfVEhVT05HJywga2hvbmdfZ29wX3Rvbmc6IGZhbHNlIH0sCiAgICAgIHsgaWQ6J1BYMycsIG1hX2NoaWV0a2hhdTonQ0tfS0hBQycgfSAvLyBraMO0bmcga2hhaSBj4budIG7DoHkgLT4gbeG6t2MgxJHhu4tuaCBLSMOUTkcgbG/huqFpCiAgICBdOwogICAgdGVzdCgnYnVpbGRFeGNsdWRlZEZyb21Ub3RhbFNldF86IGNo4buJIGNo4bupYSDEkcO6bmcgY8OhYyBNw6MgY2hp4bq/dCBraOG6pXUgY8OzIGtob25nX2dvcF90b25nPXRydWUnLCAoKT0+ewogICAgICBjb25zdCBzID0gYnVpbGRFeGNsdWRlZEZyb21Ub3RhbFNldF8oKTsKICAgICAgYXNzZXJ0VHJ1ZShzLmhhcygnQ0tfQ09fQkFOXzEwUFQnKSk7CiAgICAgIGFzc2VydFRydWUoIXMuaGFzKCdDS19USFVPTkcnKSk7CiAgICAgIGFzc2VydFRydWUoIXMuaGFzKCdDS19LSEFDJykpOwogICAgfSk7CiAgICB0ZXN0KCdpc0V4Y2x1ZGVkRnJvbVRvdGFsXzoga2jhu5twIMSRw7puZyBr4bq/dCBxdeG6oyB24bubaSBidWlsZEV4Y2x1ZGVkRnJvbVRvdGFsU2V0XyAoMiDEkcaw4budbmcgcGjhuqNpIGx1w7RuIMSR4buTbmcgbmjhuqV0KScsICgpPT57CiAgICAgIGFzc2VydFRydWUoaXNFeGNsdWRlZEZyb21Ub3RhbF8oJ0NLX0NPX0JBTl8xMFBUJykpOwogICAgICBhc3NlcnRUcnVlKCFpc0V4Y2x1ZGVkRnJvbVRvdGFsXygnQ0tfVEhVT05HJykpOwogICAgICBhc3NlcnRUcnVlKCFpc0V4Y2x1ZGVkRnJvbVRvdGFsXygnS0hPTkdfVE9OX1RBSScpKTsKICAgIH0pOwogIH0gZmluYWxseSB7CiAgICBEQi5wcm9ncmFtcyA9IHByb2dyYW1zQmFja3VwRXhjbDsKICB9CgogIC8qIC0tLS0gQsOgaSB0ZXN0IHTDrWNoIGjhu6NwOiBjb21wdXRlUmVwb3J0KCkgdHLDqm4gZOG7ryBsaeG7h3UgbXVhIGjDoG5nICsgY2jGsMahbmcgdHLDrG5oIGdp4bqjIGzhuq1wIC0tLS0gKi8KICBjb25zdCBkYkJhY2t1cCA9IERCOwogIHRyeXsKICAgIERCID0gewogICAgICBjYXRhbG9nOiB7IG1oY2s6IFsgeyBtYW1oOidTUDAwMScsIHRlbmhhbmc6J0zhu5FwIENhc3UgMDAxJywgdGh1b25naGlldTonQ0FTVScsIG5ob21raDonTk8nLCBsb2Fpc3A6J05PJywgZGFjdGluaDonTk8nLCBjb25nZHVuZzonTk8nLCBxdXljYWNoOidOTycgfSBdLCBnaWFjb25nYm86IFtdIH0sCiAgICAgIHB1cmNoYXNlczogWwogICAgICAgIHsgbmdheTonMjAyNi0wNy0wNScsIG1haGFuZzonU1AwMDEnLCBzb2x1b25nOjYwLCBnaWF0cmk6NjAwMDAwLCBuaGFjdW5nY2FwOidOQ0MgQScgfSwKICAgICAgICB7IG5nYXk6JzIwMjYtMDctMjAnLCBtYWhhbmc6J1NQMDAxJywgc29sdW9uZzo2MCwgZ2lhdHJpOjYwMDAwMCwgbmhhY3VuZ2NhcDonTkNDIEEnIH0KICAgICAgXSwKICAgICAgcHJvZ3JhbXM6IFsgeyBpZDonUFQxJywgbWFfY2hpZXRraGF1OidDVF9URVNUJywgZGltX3RodW9uZ2hpZXU6J0NBU1UnLCBkaW1fbmhvbWtoOidOTycsIGRpbV9sb2Fpc3A6J05PJywgZGltX2RhY3Rpbmg6J05PJywgZGltX2NvbmdkdW5nOidOTycsIGRpbV9xdXljYWNoOidOTycsCiAgICAgICAgICAgICAgICAgICAgc2xfbWluOjEwMCwgc2xfbWF4OjAsIGR0X21pbjowLCBkdF9tYXg6MCwgZGlldWtpZW5fc2xkdDonQU5EJywgaHRfY2hpZXRraGF1OidQRVJDRU5UX0dJQU1VQScsIHRsX2NrOjUsIG1hX2RvYW5odGh1OicnLCBkdF9rZWhvYWNoOjAgfSBdLAogICAgICBkaWV1a2llbjogW10sIGRvYW5odGh1OiBbXSwga2Vob2FjaDogW10sIGtlaG9hY2hDaGlUaWV0OiBbXQogICAgfTsKICAgIHRlc3QoJ2NvbXB1dGVSZXBvcnQgKHTDrWNoIGjhu6NwKTogU0w9MTIwIMSR4bqhdCBi4bqtYyDiiaUxMDAsIHTDrW5oIMSRw7puZyA1JSB4IGRvYW5oIHRodScsICgpPT57CiAgICAgIGNvbnN0IHJlcCA9IGNvbXB1dGVSZXBvcnQoJzIwMjYtMDctMDEnLCcyMDI2LTA3LTMxJywnJywgZmFsc2UsICcnLCAnJyk7CiAgICAgIGNvbnN0IHJvdyA9IHJlcC5maW5kKHI9PnIua2V5PT09J3Byb2dfUFQxJyk7CiAgICAgIGFzc2VydFRydWUoISFyb3csICdwaOG6o2kgY8OzIGTDsm5nIGNobyBjaMawxqFuZyB0csOsbmggQ1RfVEVTVCcpOwogICAgICBhc3NlcnRUcnVlKHJvdy5kYXQsICdwaOG6o2kgxJDhuqBUIGLhuq1jIHbDrCB04buVbmcgU0w9MTIwIOKJpSAxMDAnKTsKICAgICAgYXNzZXJ0RXF1YWwocm93LnNsLCAxMjApOyBhc3NlcnRFcXVhbChyb3cuZHQsIDEyMDAwMDApOwogICAgICBhc3NlcnRDbG9zZShyb3cudGllbmNrLCA2MDAwMCwgJ1Phu5EgdGnhu4FuIENLIHBo4bqjaSA9IDUlIHggMS4yMDAuMDAwJyk7CiAgICB9KTsKICAgIHRlc3QoJ2NvbXB1dGVSZXBvcnQgKHTDrWNoIGjhu6NwKTogY2jGsGEgxJHhuqF0IGLhuq1jIGtoaSBTTCBkxrDhu5tpIG5nxrDhu6FuZycsICgpPT57CiAgICAgIERCLnB1cmNoYXNlcyA9IFt7IG5nYXk6JzIwMjYtMDctMDUnLCBtYWhhbmc6J1NQMDAxJywgc29sdW9uZzozMCwgZ2lhdHJpOjMwMDAwMCwgbmhhY3VuZ2NhcDonTkNDIEEnIH1dOwogICAgICBjb25zdCByZXAgPSBjb21wdXRlUmVwb3J0KCcyMDI2LTA3LTAxJywnMjAyNi0wNy0zMScsJycsIGZhbHNlLCAnJywgJycpOwogICAgICBjb25zdCByb3cgPSByZXAuZmluZChyPT5yLmtleT09PSdwcm9nX1BUMScpOwogICAgICBhc3NlcnRUcnVlKHJvdyAmJiByb3cuZGF0PT09ZmFsc2UsICdTTD0zMCA8IDEwMCAtPiBwaOG6o2kgQ0jGr0EgxJHhuqF0IGLhuq1jJyk7CiAgICB9KTsKICAgIC8vIFPhu6xBIEzhu5ZJIChyw6Agc2/DoXQgdGhlbyBwaOG6o24gw6FuaCBuZ8aw4budaSBkw7luZyB0aOG7sWMgdOG6vyDigJQgImRvYW5oIHPhu5EgxJHDum5nIG5oxrBuZyB0aeG7gW4gY2hp4bq/dCBraOG6pXUgc2FpL3Ro4bqlcAogICAgLy8gaMahbiBo4bqzbiIpOiB0csaw4bubYyDEkcOieSBj4budIGNvR2lhQ29uZ0JvIChkw7luZyDEkeG7gyBj4bqjbmggYsOhbyAi4pqgIGNoxrBhIGtoYWkgYsOhbyDEkeG7pyBnacOhIGPDtG5nIGLhu5EiIGtoaSBow6xuaAogICAgLy8gdGjhu6ljIENLIGzDoCAlIHRoZW8gZ2nDoSBjw7RuZyBi4buRKSBjaOG7iSDEkcOyaSBo4buPaSBDw5MgTeG7mFQgUEjhuqZOIGdpw6EgY8O0bmcgYuG7kSBraOG7m3AgxJHGsOG7o2MgKGNiU0xzdW0+MCksIGtow7RuZyDEkcOyaQogICAgLy8gaOG7j2kga2jhu5twIMSQ4bumIGNobyB0b8OgbiBi4buZIHPhuqNuIGzGsOG7o25nIOKAlCBuw6puIGtoaSBwaOG6p24gbOG7m24gbcOjIGjDoG5nIHRyb25nIG5ow7NtIENIxq9BIGPDsyBnacOhIGPDtG5nIGLhu5EgKGNo4buJIDEKICAgIC8vIHbDoGkgbcOjIGPDsyksIGPhu50gduG6q24gYsOhbyAixJHDoyDEkeG7pyIsIEtIw5RORyBj4bqjbmggYsOhbyBnw6wsIHRyb25nIGtoaSB0aeG7gW4gQ0sgdGjhu7FjIHThur8gY2jhu4kgdMOtbmggxJHGsOG7o2MgdHLDqm4gxJHDum5nCiAgICAvLyBwaOG6p24gw610IOG7j2kgxJHDoyBraOG7m3Ag4oCUIHBo4bqnbiBjw7JuIGzhuqFpIG5n4bqnbSByYSAwxJEuIFRlc3QgbsOgeSBraG/DoSDEkcO6bmcgaMOgbmggdmkgxJDDgyBT4busQTogcGjhuqNpIHBow6F0IGhp4buHbgogICAgLy8gxJHDum5nIHRyxrDhu51uZyBo4bujcCB0aGnhur91IGdpw6EgY8O0bmcgYuG7kSBN4buYVCBQSOG6pk4gKGtow7RuZyBwaOG6o2kgY2jhu4kga2hpIHRoaeG6v3UgSE/DgE4gVE/DgE4pLgogICAgdGVzdCgnY29tcHV0ZVJlcG9ydCAodMOtY2ggaOG7o3ApOiBDSyB0aGVvICVnacOhIGPDtG5nIGLhu5EsIENI4buIIDEvMiBtw6MgaMOgbmcgY8OzIGdpw6EgY8O0bmcgYuG7kSAtPiBjb0dpYUNvbmdCbz1mYWxzZSAocGjhuqNpIGPhuqNuaCBiw6FvKSwgdGnhu4FuIENLIGNo4buJIHTDrW5oIMSRw7puZyBwaOG6p24gxJHDoyBraOG7m3AnLCAoKT0+ewogICAgICAvLyBGaXh0dXJlIFJJw4pORyBjaG8gdGVzdCBuw6B5IOKAlCBzYW8gbMawdS9raMO0aSBwaOG7pWMgY2F0YWxvZytwcm9ncmFtcytwdXJjaGFzZXMgZ+G7kWMgKFNQMDAxL1BUMSkgxJHhu4MKICAgICAgLy8ga2jDtG5nIOG6o25oIGjGsOG7n25nIGPDoWMgdGVzdCBQSMONQSBTQVUgdHJvbmcgY8O5bmcga2jhu5FpIHRyeSAoVkQgIkNoaeG6v3Qga2jhuqV1IGzFqXkga+G6vyIgbmdheSBiw6puIGTGsOG7m2kgduG6q24KICAgICAgLy8gY+G6p24gxJHDum5nIERCLnByb2dyYW1zIGPDsyBQVDEgdsOgIERCLmNhdGFsb2cgY8OzIFNQMDAxKS4KICAgICAgY29uc3QgY2F0YWxvZ0JhayA9IERCLmNhdGFsb2csIHByb2dyYW1zQmFrID0gREIucHJvZ3JhbXMsIHB1cmNoYXNlc0JhayA9IERCLnB1cmNoYXNlczsKICAgICAgdHJ5ewogICAgICAgIERCLmNhdGFsb2cgPSB7CiAgICAgICAgICBtaGNrOiBbCiAgICAgICAgICAgIHsgbWFtaDonU1BfQScsIG1hbWhfbmNjOidNQU1IX0EnLCB0ZW5oYW5nOidM4buRcCBBJywgdGh1b25naGlldTonQ0FTVScsIG5ob21raDonTk8nLCBsb2Fpc3A6J05PJywgZGFjdGluaDonTk8nLCBjb25nZHVuZzonTk8nLCBxdXljYWNoOidOTycgfSwKICAgICAgICAgICAgeyBtYW1oOidTUF9CJywgbWFtaF9uY2M6J01BTUhfQicsIHRlbmhhbmc6J0zhu5FwIEInLCB0aHVvbmdoaWV1OidDQVNVJywgbmhvbWtoOidOTycsIGxvYWlzcDonTk8nLCBkYWN0aW5oOidOTycsIGNvbmdkdW5nOidOTycsIHF1eWNhY2g6J05PJyB9CiAgICAgICAgICBdLAogICAgICAgICAgZ2lhY29uZ2JvOiBbIHsgbWFtaF9uY2M6J01BTUhfQScsIGdpYV9jb25nYm86MTAwMDAsIGhpZXVsdWN0dTonJywgaGlldWx1Y2RlbjonJyB9IF0gLy8gQ0jhu4gga2hhaSBnacOhIGPDtG5nIGLhu5EgY2hvIFNQX0EsIFRISeG6vlUgY2hvIFNQX0IKICAgICAgICB9OwogICAgICAgIERCLnB1cmNoYXNlcyA9IFsKICAgICAgICAgIHsgbmdheTonMjAyNi0wNy0wNScsIG1haGFuZzonU1BfQScsIHNvbHVvbmc6ODAsIGdpYXRyaTo3MDAwMDAsIG5oYWN1bmdjYXA6J05DQyBBJyB9LAogICAgICAgICAgeyBuZ2F5OicyMDI2LTA3LTEwJywgbWFoYW5nOidTUF9CJywgc29sdW9uZzo4MCwgZ2lhdHJpOjcwMDAwMCwgbmhhY3VuZ2NhcDonTkNDIEEnIH0gLy8gdOG7lW5nIFNMPTE2MCDiiaUgc2xfbWluPTEwMCwgbmjGsG5nIFNQX0IgY2jGsGEgY8OzIGdpw6EgY8O0bmcgYuG7kQogICAgICAgIF07CiAgICAgICAgREIucHJvZ3JhbXMgPSBbIHsgaWQ6J1BHQ0InLCBtYV9jaGlldGtoYXU6J0NUX0dDQicsIGRpbV90aHVvbmdoaWV1OidDQVNVJywgZGltX25ob21raDonTk8nLCBkaW1fbG9haXNwOidOTycsIGRpbV9kYWN0aW5oOidOTycsIGRpbV9jb25nZHVuZzonTk8nLCBkaW1fcXV5Y2FjaDonTk8nLAogICAgICAgICAgc2xfbWluOjEwMCwgc2xfbWF4OjAsIGR0X21pbjowLCBkdF9tYXg6MCwgZGlldWtpZW5fc2xkdDonQU5EJywgaHRfY2hpZXRraGF1OidQRVJDRU5UX0dJQUNPTkdCTycsIHRsX2NrOjMsIG1hX2RvYW5odGh1OicnLCBkdF9rZWhvYWNoOjAgfSBdOwogICAgICAgIGNvbnN0IHJlcCA9IGNvbXB1dGVSZXBvcnQoJzIwMjYtMDctMDEnLCcyMDI2LTA3LTMxJywnJywgZmFsc2UsICcnLCAnJyk7CiAgICAgICAgY29uc3Qgcm93ID0gcmVwLmZpbmQocj0+ci5rZXk9PT0ncHJvZ19QR0NCJyk7CiAgICAgICAgYXNzZXJ0VHJ1ZSghIXJvdyAmJiByb3cuZGF0LCAncGjhuqNpIMSQ4bqgVCBi4bqtYyAodOG7lW5nIFNMPTE2MCDiiaUgMTAwKScpOwogICAgICAgIGFzc2VydFRydWUoIXJvdy5jb0dpYUNvbmdCbywgJ2No4buJIDEvMiBtw6MgaMOgbmcgY8OzIGdpw6EgY8O0bmcgYuG7kSAtPiBLSMOUTkcgxJHGsOG7o2MgY29pIGzDoCDEkcOjIMSR4bunIChidWcgY8WpOiBjYlNMc3VtPjAgY29pIG5o4bqnbSBsw6AgxJHhu6cpJyk7CiAgICAgICAgYXNzZXJ0Q2xvc2Uocm93LnRpZW5jaywgODAqMTAwMDAqMC4wMywgJ3Rp4buBbiBDSyBjaOG7iSB0w61uaCDEkcaw4bujYyB0csOqbiDEkcO6bmcgcGjhuqduIFNQX0EgxJHDoyBraOG7m3AgZ2nDoSBjw7RuZyBi4buRIChTUF9CIGNoxrBhIGPDsyBnacOhIGPDtG5nIGLhu5EgLT4gZ8OzcCAwxJEpJyk7CiAgICAgIH0gZmluYWxseSB7CiAgICAgICAgREIuY2F0YWxvZyA9IGNhdGFsb2dCYWs7IERCLnByb2dyYW1zID0gcHJvZ3JhbXNCYWs7IERCLnB1cmNoYXNlcyA9IHB1cmNoYXNlc0JhazsKICAgICAgfQogICAgfSk7CiAgICAvLyBLSMOTQSBIw4BOSCBWSSBzYXUga2hpIGfhu5lwIGNvbXB1dGVSZXBvcnQoKS9ldmFsdWF0ZVByb2dyYW1Gb3JQZXJpb2QoKSB2w6BvIHJlc29sdmVQcm9ncmFtRXZhbHVhdGlvbl8oKQogICAgLy8gKHThu5FpIMawdSBjb2RlIOKAlCByw6Agc2/DoXQgbOG6p24gbuG7r2EpOiBjYSBiacOqbiAiY8OzIELhuq1jIMSRaeG7gXUga2nhu4duIG5oxrBuZyBLSMOUTkcgYuG6rWMgbsOgbyBraOG7m3AsIMSQ4buSTkcgVEjhu5xJIGPDswogICAgLy8ga2hhaSBN4bulYyA3IGdhdGUgKGtoX3R1X25nYXkva2hfZGVuX25nYXkpIiDigJQgdHLGsOG7m2Mga2hpIGfhu5lwLCAyIGjDoG0gdHLhuqMgaG9hbnRoYW5oIEtIw4FDIE5IQVUg4bufIMSRw7puZyBjYQogICAgLy8gbsOgeSAoZXZhbHVhdGVQcm9ncmFtRm9yUGVyaW9kIHRy4bqjICVLSCB04burIGdhdGUgZMO5IGLhuq1jIGNoxrBhIGto4bubcDsgY29tcHV0ZVJlcG9ydCB0cuG6oyBudWxsKS4gQuG6o24gZ+G7mXAKICAgIC8vIHRoZW8gcXV5IMaw4bubYyBjb21wdXRlUmVwb3J0IChudWxsIOKAlCAiY2jGsGEga2jhu5twIGLhuq1jIHRow6wga2jDtG5nIGhp4buHbiAlS0ggZOG7hSBnw6J5IGhp4buDdSBuaOG6p20iKS4gVGVzdCBuw6B5CiAgICAvLyBraG/DoSDEkcO6bmcgaMOgbmggdmkgxJDDgyBUSOG7kE5HIE5I4bqkVCDEkcOzLCB0csOhbmggMSBs4bqnbiBz4butYSBzYXUgbsOgeSB2w7QgdMOsbmggbMOgbSAyIG7GoWkgbOG7h2NoIG5oYXUgdHLhu58gbOG6oWkuCiAgICB0ZXN0KCdjb21wdXRlUmVwb3J0ICh0w61jaCBo4bujcCk6IELhuq1jIMSRaeG7gXUga2nhu4duIEtIw5RORyBraOG7m3AgKyBjw7Mga2hhaSBLaG/huqNuZyBuZ8OgeSB4w6l0IMSR4bqhdCBLSCAoZ2F0ZSBN4bulYyA3KSAtPiBDSMavQSDEkeG6oXQsICVob8OgbiB0aMOgbmggxJHhu4MgdHLhu5FuZyAoa2jDtG5nIGzhu5kgJUtIIHRow7QgdOG7qyBnYXRlKScsICgpPT57CiAgICAgIGNvbnN0IGNhdGFsb2dCYWsgPSBEQi5jYXRhbG9nLCBwcm9ncmFtc0JhayA9IERCLnByb2dyYW1zLCBwdXJjaGFzZXNCYWsgPSBEQi5wdXJjaGFzZXMsCiAgICAgICAgICAgIGRpZXVraWVuQmFrID0gREIuZGlldWtpZW4sIGRvYW5odGh1QmFrID0gREIuZG9hbmh0aHUsIGtlaG9hY2hCYWsgPSBEQi5rZWhvYWNoLCBrZWhvYWNoQ1RCYWsgPSBEQi5rZWhvYWNoQ2hpVGlldDsKICAgICAgdHJ5ewogICAgICAgIERCLmNhdGFsb2cgPSB7IG1oY2s6IFsgeyBtYW1oOidTUF9VJywgdGVuaGFuZzonTOG7kXAgVScsIHRodW9uZ2hpZXU6J1UnLCBuaG9ta2g6J05PJywgbG9haXNwOidOTycsIGRhY3Rpbmg6J05PJywgY29uZ2R1bmc6J05PJywgcXV5Y2FjaDonTk8nIH0gXSwgZ2lhY29uZ2JvOiBbXSB9OwogICAgICAgIERCLnB1cmNoYXNlcyA9IFsgeyBuZ2F5OicyMDI2LTAyLTEwJywgbWFoYW5nOidTUF9VJywgc29sdW9uZzo1LCBnaWF0cmk6NTAwMDAwLCBuaGFjdW5nY2FwOidOQ0MnIH0gXTsgLy8gU0w9NSwga2jDtG5nIGLhuq1jIG7DoG8gYmFvIGfhu5NtIDUKICAgICAgICBEQi5kb2FuaHRodSA9IFsgeyBpZDonRFRfVScsIHRlbjonVScsIGRpbV90aHVvbmdoaWV1OidVJywgZGltX25ob21raDonTk8nLCBkaW1fbG9haXNwOidOTycsIGRpbV9kYWN0aW5oOidOTycsIGRpbV9jb25nZHVuZzonTk8nLCBkaW1fcXV5Y2FjaDonTk8nLCB0aG9pZGllbV9jazonVGhlbyBoaeG7h3UgbOG7sWMgQ1RLTScsIGNhY2hfdGluaF9kb2FuaHNvOidHSUFfSE9BRE9OJyB9IF07CiAgICAgICAgREIua2Vob2FjaCA9IFsgeyBpZDonS0hVJywgbWFfZG9hbmh0aHU6J0FMTCcsIGR0X2tlaG9hY2g6MTAwMDAwMCwgc2xfa2Vob2FjaDowLCBoaWV1bHVjdHU6JzIwMjYtMDEtMDEnLCBoaWV1bHVjZGVuOicyMDI2LTEyLTMxJyB9IF07CiAgICAgICAgREIua2Vob2FjaENoaVRpZXQgPSBbIHsga2Vob2FjaF9pZDonS0hVJywgdHVfbmdheTonMjAyNi0wMS0wMScsIGRlbl9uZ2F5OicyMDI2LTEyLTMxJywgZHRfa2Vob2FjaDoxMDAwMDAwLCBzbF9rZWhvYWNoOjAsIGt5Om51bGwgfSBdOwogICAgICAgIERCLnByb2dyYW1zID0gWyB7IGlkOidQNicsIG1hX2NoaWV0a2hhdTonUDYnLCBkaW1fdGh1b25naGlldTonVScsIGRpbV9uaG9ta2g6J05PJywgZGltX2xvYWlzcDonTk8nLCBkaW1fZGFjdGluaDonTk8nLCBkaW1fY29uZ2R1bmc6J05PJywgZGltX3F1eWNhY2g6J05PJywKICAgICAgICAgIHNsX21pbjowLCBzbF9tYXg6MCwgZHRfbWluOjAsIGR0X21heDowLCBkaWV1a2llbl9zbGR0OidBTkQnLCBodF9jaGlldGtoYXU6J1BFUkNFTlRfR0lBTVVBJywgdGxfY2s6MCwgbWFfZG9hbmh0aHU6J0RUX1UnLCBkdF9rZWhvYWNoOjAsCiAgICAgICAgICBoaWV1bHVjdHU6JzIwMjYtMDEtMDEnLCBoaWV1bHVjZGVuOicyMDI2LTEyLTMxJywga2hfdHVfbmdheTonMjAyNi0wMS0wMScsIGtoX2Rlbl9uZ2F5OicyMDI2LTEyLTMxJyB9IF07CiAgICAgICAgREIuZGlldWtpZW4gPSBbIHsgaWQ6J0Q2QScsIHByb2dyYW1faWQ6J1A2Jywgc2xfbWluOjEwMDAsIHNsX21heDowLCBkdF9taW46MCwgZHRfbWF4OjAsIGRpZXVraWVuX3NsZHQ6J0FORCcsIHBjdF9raF9taW46JycsIHBjdF9raF9tYXg6JycsIHRsX2NrOjIsIGR0X2tlaG9hY2g6JycgfSBdOwogICAgICAgIGNvbnN0IHJlcCA9IGNvbXB1dGVSZXBvcnQoJzIwMjYtMDEtMDEnLCcyMDI2LTEyLTMxJywnJywgZmFsc2UsICcnLCAnJyk7CiAgICAgICAgY29uc3Qgcm93ID0gcmVwLmZpbmQocj0+ci5rZXk9PT0ncHJvZ19QNicpOwogICAgICAgIGFzc2VydFRydWUoISFyb3csICdwaOG6o2kgY8OzIGTDsm5nIGNobyBjaMawxqFuZyB0csOsbmggUDYgKGPDsyBtw6MgaMOgbmcgcGjDoXQgc2luaCBtdWEpJyk7CiAgICAgICAgYXNzZXJ0VHJ1ZSghcm93LmRhdCAmJiByb3cubm9CYWNNYXRjaGVkLCAnU0w9NSA8IG5nxrDhu6FuZyBi4bqtYyBENkEgKDEwMDApIC0+IHBo4bqjaSBDSMavQSDEkeG6oXQsIG5vQmFjTWF0Y2hlZD10cnVlJyk7CiAgICAgICAgYXNzZXJ0VHJ1ZShyb3cuaG9hbnRoYW5oPT1udWxsLCAnJWhvw6BuIHRow6BuaCBwaOG6o2kgxJHhu4MgVFLhu5BORyBraGkgY2jGsGEga2jhu5twIGLhuq1jIG7DoG8sIGvhu4MgY+G6oyBraGkgY8OzIGtoYWkgZ2F0ZSByacOqbmcgKGtow7RuZyBs4buZICVLSCB0aMO0IGfDonkgaGnhu4N1IG5o4bqnbSknKTsKICAgICAgfSBmaW5hbGx5IHsKICAgICAgICBEQi5jYXRhbG9nID0gY2F0YWxvZ0JhazsgREIucHJvZ3JhbXMgPSBwcm9ncmFtc0JhazsgREIucHVyY2hhc2VzID0gcHVyY2hhc2VzQmFrOwogICAgICAgIERCLmRpZXVraWVuID0gZGlldWtpZW5CYWs7IERCLmRvYW5odGh1ID0gZG9hbmh0aHVCYWs7IERCLmtlaG9hY2ggPSBrZWhvYWNoQmFrOyBEQi5rZWhvYWNoQ2hpVGlldCA9IGtlaG9hY2hDVEJhazsKICAgICAgfQogICAgfSk7CiAgICAvLyBLSMOTQSBIw4BOSCBWSSAocGjDoXQgaGnhu4duIHThu6sgZOG7ryBsaeG7h3UgdGjhuq10IGPhu6dhIG5nxrDhu51pIGTDuW5nIOKAlCBmaWxlIEV4Y2VsIGV4cG9ydCBTaGVldCwgbcOjIGNoaeG6v3Qga2jhuqV1CiAgICAvLyAiQ1NNTl8yMzEwMTNfTk9fTk9fTk9fTk9fMDAxIiBiw6FvIDDEkSBjaGnhur90IGto4bqldSBkw7kgcsO1IHLDoG5nIGPDsyBkb2FuaCBz4buRKTogZGltX25ob21raCB0csOqbgogICAgLy8gQ0hVT05HVFJJTkhDSElFVEtIQVUgY+G7p2EgOCBtw6MgKHThuqFvIDI0LTI3LzA4LzIwMjYpIGLhu4sgR29vZ2xlIFNoZWV0IGzGsHUgTOG7hkNIIMSQ4buYIENIw41OSCBYw4FDIFPhu5AKICAgIC8vICgiMjMxMDEzLjAwOTAyOCIvIjIzMTAxMy4wMDQ1MTQiIHRoYXkgdsOsICIyMzEwMTMiKSwgdHJvbmcga2hpIE3DoyBow6BuZyAoRE1fTUhDSykgduG6q24gbMawdSDEkcO6bmcKICAgIC8vICIyMzEwMTMiIOKAlCBkaW1zTWF0Y2goKSB0csaw4bubYyDEkcOieSBzbyBLSOG7mlAgQ0hV4buWSSB0dXnhu4d0IMSR4buRaSBuw6puIGxv4bqhaSBi4buPIFRPw4BOIELhu5ggbcOjIGjDoG5nLCByYSAwxJEgQ0sKICAgIC8vIGtow7RuZyBj4bqjbmggYsOhby4gVGVzdCBuw6B5IGtob8OhIMSRw7puZyBow6BuaCB2aSDEkMODIFPhu6xBOiBkaW1zIGThuqFuZyBz4buRIGzhu4djaCDEkeG7mSBjaMOtbmggeMOhYyBuaOG7jyB24bqrbiBwaOG6o2kga2jhu5twLgogICAgdGVzdCgnZGltc01hdGNoL2NvbXB1dGVSZXBvcnQgKHTDrWNoIGjhu6NwKTogZGltX25ob21raCBs4buHY2ggxJHhu5kgY2jDrW5oIHjDoWMgc+G7kSAoMjMxMDEzLjAwOTAyOCB2cyAyMzEwMTMpIHbhuqtuIHBo4bqjaSBraOG7m3AgbcOjIGjDoG5nJywgKCk9PnsKICAgICAgY29uc3QgY2F0YWxvZ0JhayA9IERCLmNhdGFsb2csIHByb2dyYW1zQmFrID0gREIucHJvZ3JhbXMsIHB1cmNoYXNlc0JhayA9IERCLnB1cmNoYXNlcywgZG9hbmh0aHVCYWsgPSBEQi5kb2FuaHRodTsKICAgICAgdHJ5ewogICAgICAgIERCLmNhdGFsb2cgPSB7CiAgICAgICAgICBtaGNrOiBbIHsgbWFtaDonWUVNMDEnLCBtYW1oX25jYzonTkNDX1lFTTAxJywgdGVuaGFuZzonWeG6v20gQ2FzdW1pbmEnLCB0aHVvbmdoaWV1OidDU01OJywgbmhvbWtoOjIzMTAxMy4wLCBsb2Fpc3A6J05PJywgZGFjdGluaDonTk8nLCBjb25nZHVuZzonTk8nLCBxdXljYWNoOidOTycgfSBdLAogICAgICAgICAgZ2lhY29uZ2JvOiBbXQogICAgICAgIH07CiAgICAgICAgREIucHVyY2hhc2VzID0gWyB7IG5nYXk6JzIwMjYtMDMtMDUnLCBtYWhhbmc6J1lFTTAxJywgc29sdW9uZzoxMDAsIGdpYXRyaToxMzAwMDAwMCwgbmhhY3VuZ2NhcDonTkNDIEEnIH0gXTsKICAgICAgICBEQi5kb2FuaHRodSA9IFtdOwogICAgICAgIC8vIGRpbV9uaG9ta2ggQ+G7kCDDnSBs4buHY2ggxJHhu5kgY2jDrW5oIHjDoWMgc+G7kSBnaeG7kW5nIGjhu4d0IGThu68gbGnhu4d1IHRo4bqtdCBi4buLIGzhu5dpIHRyw6puIFNoZWV0CiAgICAgICAgREIucHJvZ3JhbXMgPSBbIHsgaWQ6J1BDU01OJywgbWFfY2hpZXRraGF1OidDU01OXzIzMTAxM19OT19OT19OT19OT18wMDEnLCBkaW1fdGh1b25naGlldTonQ1NNTicsIGRpbV9uaG9ta2g6MjMxMDEzLjAwOTAyOCwKICAgICAgICAgIGRpbV9sb2Fpc3A6J05PJywgZGltX2RhY3Rpbmg6J05PJywgZGltX2NvbmdkdW5nOidOTycsIGRpbV9xdXljYWNoOidOTycsCiAgICAgICAgICBzbF9taW46MCwgc2xfbWF4OjAsIGR0X21pbjowLCBkdF9tYXg6MCwgZGlldWtpZW5fc2xkdDonQU5EJywgaHRfY2hpZXRraGF1OidQRVJDRU5UX0dJQU1VQScsIHRsX2NrOjEwLCBtYV9kb2FuaHRodTonJywgZHRfa2Vob2FjaDowLAogICAgICAgICAgaGlldWx1Y3R1OicyMDI2LTAxLTAxJywgaGlldWx1Y2RlbjonMjAyNi0xMi0zMScgfSBdOwogICAgICAgIGNvbnN0IHJlcCA9IGNvbXB1dGVSZXBvcnQoJzIwMjYtMDMtMDEnLCcyMDI2LTAzLTMxJywnJywgZmFsc2UsICcnLCAnJyk7CiAgICAgICAgY29uc3Qgcm93ID0gcmVwLmZpbmQocj0+ci5rZXk9PT0ncHJvZ19QQ1NNTicpOwogICAgICAgIGFzc2VydFRydWUoISFyb3csICdwaOG6o2kgY8OzIGTDsm5nIGNobyBtw6MgQ1NNTl8yMzEwMTNfTk9fTk9fTk9fTk9fMDAxIChraMO0bmcgxJHGsOG7o2MgxJHhu4MgMCBtw6MgaMOgbmcga2jhu5twIHbDrCBs4buHY2ggxJHhu5kgY2jDrW5oIHjDoWMgc+G7kSknKTsKICAgICAgICBhc3NlcnRFcXVhbChyb3cuc29NYUhhbmcsIDEsICdwaOG6o2kga2jhu5twIMSRw7puZyAxIG3DoyBow6BuZyB0aHXhu5ljIE5ow7NtIGjDoG5nIDIzMTAxMycpOwogICAgICAgIGFzc2VydENsb3NlKHJvdy50aWVuY2ssIDEzMDAwMDAwKjAuMSwgJ3Rp4buBbiBDSyBwaOG6o2kgdMOtbmggxJHDum5nIHRyw6puIGRvYW5oIHPhu5EgxJHDoyBraOG7m3AsIGtow7RuZyBwaOG6o2kgMMSRJyk7CiAgICAgIH0gZmluYWxseSB7CiAgICAgICAgREIuY2F0YWxvZyA9IGNhdGFsb2dCYWs7IERCLnByb2dyYW1zID0gcHJvZ3JhbXNCYWs7IERCLnB1cmNoYXNlcyA9IHB1cmNoYXNlc0JhazsgREIuZG9hbmh0aHUgPSBkb2FuaHRodUJhazsKICAgICAgfQogICAgfSk7CgogICAgLyogLS0tLSBDaGnhur90IGto4bqldSBsxal5IGvhur8gKGNoaeG6v3Qga2jhuqV1IGLhu5Ugc3VuZyk6IGFsbF9vcl9yZWR1Y2VkIC0tLS0gKi8KICAgIERCLnB1cmNoYXNlcyA9IFsKICAgICAgeyBuZ2F5OicyMDI2LTAxLTEwJywgbWFoYW5nOidTUDAwMScsIHNvbHVvbmc6MTUwLCBnaWF0cmk6MTUwMDAwMCwgbmhhY3VuZ2NhcDonTkNDIEEnIH0sCiAgICAgIHsgbmdheTonMjAyNi0wMi0xMCcsIG1haGFuZzonU1AwMDEnLCBzb2x1b25nOjEwMCwgZ2lhdHJpOjEwMDAwMDAsIG5oYWN1bmdjYXA6J05DQyBBJyB9CiAgICBdOwogICAgREIucmVwb3J0Q0tUSCA9IFsKICAgICAgeyBwcm9ncmFtX2lkOidQVDEnLCB0aG9pZGllbV9jaGlldGtoYXU6J1Row6FuZycsIHRoYW5nX2NoaWV0a2hhdToxLCBuYW1fY2hpZXRraGF1OjIwMjYsIHRyYW5ndGhhaTonREFUJywgcGN0X2h0X2toX2R0Ojk1IH0sCiAgICAgIHsgcHJvZ3JhbV9pZDonUFQxJywgdGhvaWRpZW1fY2hpZXRraGF1OidUaMOhbmcnLCB0aGFuZ19jaGlldGtoYXU6MiwgbmFtX2NoaWV0a2hhdToyMDI2LCB0cmFuZ3RoYWk6J0RBVCcsIHBjdF9odF9raF9kdDo5NSB9CiAgICBdOwogICAgdGVzdCgnY29tcHV0ZUJvbnVzRGlzY291bnQgKGFsbF9vcl9yZWR1Y2VkKTogdOG6pXQgY+G6oyB0aMOhbmcgxJHhuqF0ICsgYuG6rWMgcmnDqm5nIMSR4bqhdCAtPiBoxrDhu59uZyDEkOG7piAlIGLhu5Ugc3VuZycsICgpPT57CiAgICAgIGNvbnN0IGxrID0geyBpZDonTEsxJywgbWFfY2tfY2hpbmg6J1BUMScsIGtpZXVfdGluaDonYWxsX29yX3JlZHVjZWQnLCBzbF9taW46MjAwLCBzbF9tYXg6MCwgZHRfbWluOjAsIGR0X21heDowLAogICAgICAgIGNrX2Jvc3VuZzoyLCBja19yZWR1Y2VkX3BjdDo1MCwgaGlldWx1Y3R1OicyMDI2LTAxLTAxJywgaGlldWx1Y2RlbjonMjAyNi0wMi0yOCcgfTsKICAgICAgY29uc3QgciA9IGNvbXB1dGVCb251c0Rpc2NvdW50KGxrLCBbMSwyXSwgJ1Row6FuZycsICcyMDI2LTAxLTAxJywgJzIwMjYtMDItMjgnKTsKICAgICAgYXNzZXJ0VHJ1ZShyLm9rICYmIHIuZmluYWxEYXQsICdwaOG6o2kgxJDhuqBUIChTTCBj4bqjIGvhu7MgMjUwIOKJpSAyMDAgdsOgIGPhuqMgMiB0aMOhbmcgxJHhu4F1IMSR4bqhdCknKTsKICAgICAgYXNzZXJ0VHJ1ZShyLmFsbE1vbnRoc09rLCAnY+G6oyAyIHRow6FuZyDEkeG7gXUgxJHDoyBsxrB1IMSQ4bqgVCcpOwogICAgICBhc3NlcnRFcXVhbChyLmVmZmVjdGl2ZVJhdGVQY3QsIDIsICdraMO0bmcgYuG7iyBnaeG6o20gdHLhu6sgLT4gaMaw4bufbmcgxJHhu6cgMiUnKTsKICAgICAgYXNzZXJ0Q2xvc2Uoci50aWVuY2ssIDI1MDAwMDAqMC4wMiwgJ1Rp4buBbiBDSyBi4buVIHN1bmcgPSAyJSB4IHThu5VuZyBEVCBj4bqjIGvhu7MnKTsKICAgIH0pOwogICAgdGVzdCgnY29tcHV0ZUJvbnVzRGlzY291bnQgKGFsbF9vcl9yZWR1Y2VkKTogMSB0aMOhbmcgQ0jGr0EgxJHhuqF0IG5oxrBuZyBj4bqjIGvhu7MgduG6q24gxJHhuqF0IGLhuq1jIHJpw6puZyAtPiBHSeG6ok0gVFLhu6ogdGhlbyBja19yZWR1Y2VkX3BjdCcsICgpPT57CiAgICAgIERCLnJlcG9ydENLVEhbMV0udHJhbmd0aGFpID0gJ0NIVUFfREFUJzsgLy8gdGjDoW5nIDIgY2jGsGEgxJHhuqF0CiAgICAgIGNvbnN0IGxrID0geyBpZDonTEsxJywgbWFfY2tfY2hpbmg6J1BUMScsIGtpZXVfdGluaDonYWxsX29yX3JlZHVjZWQnLCBzbF9taW46MjAwLCBzbF9tYXg6MCwgZHRfbWluOjAsIGR0X21heDowLAogICAgICAgIGNrX2Jvc3VuZzoyLCBja19yZWR1Y2VkX3BjdDo1MCwgaGlldWx1Y3R1OicyMDI2LTAxLTAxJywgaGlldWx1Y2RlbjonMjAyNi0wMi0yOCcgfTsKICAgICAgY29uc3QgciA9IGNvbXB1dGVCb251c0Rpc2NvdW50KGxrLCBbMSwyXSwgJ1Row6FuZycsICcyMDI2LTAxLTAxJywgJzIwMjYtMDItMjgnKTsKICAgICAgYXNzZXJ0VHJ1ZShyLmZpbmFsRGF0ICYmIHIuaXNSZWR1Y2VkLCAnduG6q24gxJHhuqF0IG5oxrBuZyDhu58gbeG7qWMgZ2nhuqNtIHRy4burJyk7CiAgICAgIGFzc2VydEVxdWFsKHIuZWZmZWN0aXZlUmF0ZVBjdCwgMSwgJ2dp4bqjbSB0cuG7qyA1MCUgY+G7p2EgMiUgPSAxJScpOwogICAgICBEQi5yZXBvcnRDS1RIWzFdLnRyYW5ndGhhaSA9ICdEQVQnOyAvLyBraMO0aSBwaOG7pWMgY2hvIGPDoWMgdGVzdCBzYXUKICAgIH0pOwoKICAgIC8qIC0tLS0gQ2hp4bq/dCBraOG6pXUgbMWpeSBr4bq/OiBkdWFsX2NvbmRpdGlvbiAoMiDEkWnhu4F1IGtp4buHbiDEkeG7mWMgbOG6rXAsIGPhu5luZyBk4buTbiAlKSAtLS0tICovCiAgICB0ZXN0KCdjb21wdXRlRHVhbENvbmRpdGlvbkJvbnVzOiBjaOG7iSDEkeG6oXQgxJBLMSAodGhlbyB0aMOhbmcpLCBDSMavQSDEkeG6oXQgxJBLMiAodGhlbyBr4buzKSAtPiBjaOG7iSBj4buZbmcgZGsxX3JhdGUnLCAoKT0+ewogICAgICBjb25zdCBsayA9IHsgaWQ6J0xLMicsIG1hX2NrX2NoaW5oOidQVDEnLCBraWV1X3Rpbmg6J2R1YWxfY29uZGl0aW9uJywgZGsxX3BjdF90aGFuZ19taW46OTAsIGRrMV9yYXRlOjEsCiAgICAgICAgZGsyX3BjdF9reV9taW46MTEwLCBkazJfcmF0ZToxLjUsIHNsX21pbjowLCBzbF9tYXg6MCwgZHRfbWluOjAsIGR0X21heDowLCBkdF9rZWhvYWNoOjMwMDAwMDAsCiAgICAgICAgaGlldWx1Y3R1OicyMDI2LTAxLTAxJywgaGlldWx1Y2RlbjonMjAyNi0wMi0yOCcgfTsKICAgICAgLy8gVOG7lW5nIERUIGPhuqMga+G7syA9IDIuNTAwLjAwMCAvIEtIIDMuMDAwLjAwMCA9IDgzJSA8IDExMCUgLT4gxJBLMiBLSMOUTkcgxJHhuqF0CiAgICAgIGNvbnN0IHIgPSBjb21wdXRlQm9udXNEaXNjb3VudChsaywgWzEsMl0sICdUaMOhbmcnLCAnMjAyNi0wMS0wMScsICcyMDI2LTAyLTI4Jyk7CiAgICAgIGFzc2VydEVxdWFsKHIua2lldV90aW5oLCAnZHVhbF9jb25kaXRpb24nKTsKICAgICAgYXNzZXJ0VHJ1ZShyLmRrMU1ldD09PXRydWUsICdj4bqjIDIgdGjDoW5nIMSR4buBdSDiiaU5MCUgS0ggdGjDoW5nICg5NSUpIC0+IMSQSzEgxJHhuqF0Jyk7CiAgICAgIGFzc2VydFRydWUoci5kazJNZXQ9PT1mYWxzZSwgJ3Thu5VuZyBr4buzIDgzJSA8IDExMCUgLT4gxJBLMiBjaMawYSDEkeG6oXQnKTsKICAgICAgYXNzZXJ0RXF1YWwoci5lZmZlY3RpdmVSYXRlUGN0LCAxLCAnY2jhu4kgY+G7mW5nIGRrMV9yYXRlPTElLCBraMO0bmcgY+G7mW5nIGRrMl9yYXRlJyk7CiAgICB9KTsKICAgIHRlc3QoJ2NvbXB1dGVEdWFsQ29uZGl0aW9uQm9udXM6IMSR4bqhdCBD4bqiIDIgxJFp4buBdSBraeG7h24gLT4gQ+G7mE5HIEThu5JOIGPhuqMgMiB04bu3IGzhu4cnLCAoKT0+ewogICAgICBjb25zdCBsayA9IHsgaWQ6J0xLMicsIG1hX2NrX2NoaW5oOidQVDEnLCBraWV1X3Rpbmg6J2R1YWxfY29uZGl0aW9uJywgZGsxX3BjdF90aGFuZ19taW46OTAsIGRrMV9yYXRlOjEsCiAgICAgICAgZGsyX3BjdF9reV9taW46MTEwLCBkazJfcmF0ZToxLjUsIHNsX21pbjowLCBzbF9tYXg6MCwgZHRfbWluOjAsIGR0X21heDowLCBkdF9rZWhvYWNoOjIwMDAwMDAsCiAgICAgICAgaGlldWx1Y3R1OicyMDI2LTAxLTAxJywgaGlldWx1Y2RlbjonMjAyNi0wMi0yOCcgfTsKICAgICAgLy8gVOG7lW5nIERUIGPhuqMga+G7syA9IDIuNTAwLjAwMCAvIEtIIDIuMDAwLjAwMCA9IDEyNSUg4omlIDExMCUgLT4gxJBLMiDEkeG6oXQKICAgICAgY29uc3QgciA9IGNvbXB1dGVCb251c0Rpc2NvdW50KGxrLCBbMSwyXSwgJ1Row6FuZycsICcyMDI2LTAxLTAxJywgJzIwMjYtMDItMjgnKTsKICAgICAgYXNzZXJ0VHJ1ZShyLmRrMU1ldCAmJiByLmRrMk1ldCwgJ3Bo4bqjaSDEkeG6oXQgY+G6oyAyIMSRaeG7gXUga2nhu4duJyk7CiAgICAgIGFzc2VydEVxdWFsKHIuZWZmZWN0aXZlUmF0ZVBjdCwgMi41LCAnY+G7mW5nIGThu5NuIDElICsgMS41JSA9IDIuNSUsIEtIw5RORyBwaOG6o2kgY2jhu41uIG1heCcpOwogICAgfSk7CgogICAgLyogLS0tLSBmaW5kTW9udGhSZXBvcnRfOiBEUkFGVCBwaOG6o2kgxJHGsOG7o2MgxrB1IHRpw6puIGjGoW4gYsOhbyBjw6FvIENIw41OSCBUSOG7qEMgxJHDoyBsxrB1IC0tLS0gKi8KICAgIHRlc3QoJ2ZpbmRNb250aFJlcG9ydF86IMawdSB0acOqbiBEUkFGVCAobeG7m2kgdMOtbmgpIGjGoW4gREIucmVwb3J0Q0tUSCAoxJHDoyBsxrB1IGPFqSknLCAoKT0+ewogICAgICBEUkFGVF9SRVBPUlRfQ0tUSF8gPSBbeyBwcm9ncmFtX2lkOidQVDEnLCB0aG9pZGllbV9jaGlldGtoYXU6J1Row6FuZycsIHRoYW5nX2NoaWV0a2hhdToxLCBuYW1fY2hpZXRraGF1OjIwMjYsIHRyYW5ndGhhaTonQ0hVQV9EQVQnIH1dOwogICAgICBjb25zdCByZWMgPSBmaW5kTW9udGhSZXBvcnRfKCdQVDEnLCdUaMOhbmcnLDEsMjAyNik7CiAgICAgIGFzc2VydEVxdWFsKHJlYy50cmFuZ3RoYWksICdDSFVBX0RBVCcsICdwaOG6o2kgbOG6pXkgdOG7qyBkcmFmdCAoQ0hVQV9EQVQpLCBraMO0bmcgcGjhuqNpIERCLnJlcG9ydENLVEggKMSRYW5nIGzDoCBEQVQpJyk7CiAgICAgIERSQUZUX1JFUE9SVF9DS1RIXyA9IFtdOwogICAgICBjb25zdCByZWMyID0gZmluZE1vbnRoUmVwb3J0XygnUFQxJywnVGjDoW5nJywxLDIwMjYpOwogICAgICBhc3NlcnRFcXVhbChyZWMyLnRyYW5ndGhhaSwgJ0RBVCcsICdzYXUga2hpIHhvw6EgZHJhZnQgcGjhuqNpIHLGoWkgduG7gSDEkcO6bmcgYsOhbyBjw6FvIGNow61uaCB0aOG7qWMgxJHDoyBsxrB1Jyk7CiAgICB9KTsKICB9IGZpbmFsbHkgewogICAgREIgPSBkYkJhY2t1cDsKICAgIERSQUZUX1JFUE9SVF9DS1RIXyA9IFtdOwogIH0KCiAgLyogLS0tLSBlbnVtZXJhdGVQZXJpb2RzXzogdMOhY2gga2hv4bqjbmcgaGnhu4d1IGzhu7FjIHRow6BuaCB04burbmcga+G7syBjaHXhuqluIC0tLS0gKi8KICB0ZXN0KCdlbnVtZXJhdGVQZXJpb2RzXzogVGjDoW5nIC0+IHTDoWNoIMSRw7puZyB04burbmcgdGjDoW5nIHRyb25nIGtob+G6o25nIGhp4buHdSBs4buxYycsICgpPT57CiAgICBjb25zdCBwID0gZW51bWVyYXRlUGVyaW9kc18oJzIwMjYtMDUtMTAnLCAnMjAyNi0wNy0zMScsICdUaMOhbmcnKTsKICAgIGFzc2VydEVxdWFsKHAubGVuZ3RoLCAzKTsKICAgIGFzc2VydEVxdWFsKHBbMF0uZnJvbSwgJzIwMjYtMDUtMDEnKTsgYXNzZXJ0RXF1YWwocFswXS50bywgJzIwMjYtMDUtMzEnKTsKICAgIGFzc2VydEVxdWFsKHBbMl0uZnJvbSwgJzIwMjYtMDctMDEnKTsgYXNzZXJ0RXF1YWwocFsyXS50bywgJzIwMjYtMDctMzEnKTsKICB9KTsKICB0ZXN0KCdlbnVtZXJhdGVQZXJpb2RzXzogUXXDvSAtPiB0w6FjaCDEkcO6bmcgdOG7q25nIHF1w70nLCAoKT0+ewogICAgY29uc3QgcCA9IGVudW1lcmF0ZVBlcmlvZHNfKCcyMDI2LTAxLTAxJywgJzIwMjYtMDktMzAnLCAnUXXDvScpOwogICAgYXNzZXJ0RXF1YWwocC5sZW5ndGgsIDMpOwogICAgYXNzZXJ0RXF1YWwocFsxXS5mcm9tLCAnMjAyNi0wNC0wMScpOyBhc3NlcnRFcXVhbChwWzFdLnRvLCAnMjAyNi0wNi0zMCcpOwogIH0pOwogIHRlc3QoJ2VudW1lcmF0ZVBlcmlvZHNfOiAiVMO5eSBjaOG7jW4iIChy4buXbmcpIC0+IEtIw5RORyB0w6FjaCwgY29pIGPhuqMga2hv4bqjbmcgbMOgIDEga+G7sycsICgpPT57CiAgICBjb25zdCBwID0gZW51bWVyYXRlUGVyaW9kc18oJzIwMjYtMDEtMDEnLCAnMjAyNi0wNi0zMCcsICcnKTsKICAgIGFzc2VydEVxdWFsKHAubGVuZ3RoLCAxKTsKICAgIGFzc2VydEVxdWFsKHBbMF0uZnJvbSwgJzIwMjYtMDEtMDEnKTsgYXNzZXJ0RXF1YWwocFswXS50bywgJzIwMjYtMDYtMzAnKTsKICB9KTsKICB0ZXN0KCdlbnVtZXJhdGVQZXJpb2RzXzogY2jGsGEga2hhaSBIaeG7h3UgbOG7sWMgdOG7qyBuZ8OgeSAtPiB0cuG6oyBt4bqjbmcgcuG7l25nJywgKCk9PnsKICAgIGFzc2VydEVxdWFsKGVudW1lcmF0ZVBlcmlvZHNfKCcnLCAnMjAyNi0wNi0zMCcsICdUaMOhbmcnKS5sZW5ndGgsIDApOwogIH0pOwoKICAvKiAtLS0tIGJ1aWxkRGVkdXBQbGFuXzogY2jhu41uIMSRw7puZyBraG/huqNuZyBuZ8OgeSBo4bujcCBs4buHLCB4b8OhIMSRw7puZyBz4buRIGTDsm5nIHRyw7luZyAtLS0tICovCiAgdGVzdCgnYnVpbGREZWR1cFBsYW5fOiBnaeG7ryBraG/huqNuZyBraOG7m3AgcmFuaCBnaeG7m2kga+G7syBjaHXhuqluLCBsb+G6oWkga2hv4bqjbmcgYuG7iyBs4buHY2ggbmfDoHknLCAoKT0+ewogICAgY29uc3QgZGJCYWNrdXAyID0gREI7CiAgICB0cnl7CiAgICAgIERCID0gT2JqZWN0LmFzc2lnbih7fSwgREIsIHsKICAgICAgICByZXBvcnRDS1RIOiBbCiAgICAgICAgICB7IHByb2dyYW1faWQ6J1BYJywgdHVfbmdheTonMjAyNi0wNy0wMScsIGRlbl9uZ2F5OicyMDI2LTA3LTMwJywgdGhvaWRpZW1fY2hpZXRraGF1OidUaMOhbmcnLCB0aGFuZ19jaGlldGtoYXU6NywgbmFtX2NoaWV0a2hhdToyMDI2LCB0cmFuZ3RoYWk6J0RBVCcsIHNvX3RpZW5fY2hpZXRraGF1OjEwMDAgfSwKICAgICAgICAgIHsgcHJvZ3JhbV9pZDonUFgnLCB0dV9uZ2F5OicyMDI2LTA3LTAxJywgZGVuX25nYXk6JzIwMjYtMDctMzEnLCB0aG9pZGllbV9jaGlldGtoYXU6J1Row6FuZycsIHRoYW5nX2NoaWV0a2hhdTo3LCBuYW1fY2hpZXRraGF1OjIwMjYsIHRyYW5ndGhhaTonREFUJywgc29fdGllbl9jaGlldGtoYXU6MjAwMCB9CiAgICAgICAgXSwKICAgICAgICByZXBvcnRDS0NUOiBbXQogICAgICB9KTsKICAgICAgY29uc3QgcGxhbiA9IGJ1aWxkRGVkdXBQbGFuXygpOwogICAgICBhc3NlcnRFcXVhbChwbGFuLnJlbW92ZWRDa3RoLCAxLCAncGjhuqNpIGxv4bqhaSDEkcO6bmcgMSBkw7JuZyB0csO5bmcgKGtob+G6o25nIG5nw6B5IGzhu4djaDogMDctMzAgdGhheSB2w6wgMDctMzEpJyk7CiAgICAgIGFzc2VydEVxdWFsKHBsYW4uZmluYWxDa3RoLmxlbmd0aCwgMSk7CiAgICAgIGFzc2VydEVxdWFsKHBsYW4uZmluYWxDa3RoWzBdLmRlbl9uZ2F5LCAnMjAyNi0wNy0zMScsICdwaOG6o2kgR0nhu64gxJHDum5nIGtob+G6o25nIGto4bubcCByYW5oIGdp4bubaSBr4buzIGNodeG6qW4gKDAxLzA3LTMxLzA3KSwgbG/huqFpIGtob+G6o25nIGzhu4djaCAoMDEvMDctMzAvMDcpJyk7CiAgICB9IGZpbmFsbHkgewogICAgICBEQiA9IGRiQmFja3VwMjsKICAgIH0KICB9KTsKCiAgLyogLS0tLSByZXNvbHZlVGFtVGluaFBlcmlvZF86IGNo4buNbiBr4buzIHRoYW0gY2hp4bq/dSBjaG8gIkLhuqNuZyBjaGnhur90IGto4bqldSB04bqhbSB0w61uaCB0aGVvIHRow6FuZyIgLS0tLSAqLwogIHRlc3QoJ3Jlc29sdmVUYW1UaW5oUGVyaW9kXzogVGjDoW5nIC0+IExVw5ROIGzhuqV5IHRow6FuZyBoaeG7h24gaMOgbmggKGFzT2ZJU08pJywgKCk9PnsKICAgIGNvbnN0IHIgPSByZXNvbHZlVGFtVGluaFBlcmlvZF8oJ1Row6FuZycsICcyMDI2LTA5LTE1Jyk7CiAgICBhc3NlcnRFcXVhbChyLmZyb20sICcyMDI2LTA5LTAxJyk7IGFzc2VydEVxdWFsKHIudG8sICcyMDI2LTA5LTMwJyk7CiAgfSk7CiAgdGVzdCgncmVzb2x2ZVRhbVRpbmhQZXJpb2RfOiBRdcO9IOKAlCB0aMOhbmcgaGnhu4duIGjDoG5oIEdJ4buuQSBxdcO9ICh0aMOhbmcgOCkgLT4gbMO5aSB24buBIHF1w70gVFLGr+G7mkMgKFEyKScsICgpPT57CiAgICBjb25zdCByID0gcmVzb2x2ZVRhbVRpbmhQZXJpb2RfKCdRdcO9JywgJzIwMjYtMDgtMjAnKTsKICAgIGFzc2VydEVxdWFsKHIuZnJvbSwgJzIwMjYtMDQtMDEnKTsgYXNzZXJ0RXF1YWwoci50bywgJzIwMjYtMDYtMzAnKTsKICB9KTsKICB0ZXN0KCdyZXNvbHZlVGFtVGluaFBlcmlvZF86IFF1w70g4oCUIHRow6FuZyBoaeG7h24gaMOgbmggTMOAIHRow6FuZyBjdeG7kWkgcXXDvSAodGjDoW5nIDkpIC0+IGTDuW5nIHF1w70gSEnhu4ZOIEjDgE5IIChRMyknLCAoKT0+ewogICAgY29uc3QgciA9IHJlc29sdmVUYW1UaW5oUGVyaW9kXygnUXXDvScsICcyMDI2LTA5LTA1Jyk7CiAgICBhc3NlcnRFcXVhbChyLmZyb20sICcyMDI2LTA3LTAxJyk7IGFzc2VydEVxdWFsKHIudG8sICcyMDI2LTA5LTMwJyk7CiAgfSk7CiAgdGVzdCgncmVzb2x2ZVRhbVRpbmhQZXJpb2RfOiBRdcO9IOKAlCB0aMOhbmcgMSAoxJHhuqd1IFExKSAtPiBsw7lpIHbhu4EgUTQgTsSCTSBUUsav4buaQyAoduG6r3QgcXVhIG7Eg20pJywgKCk9PnsKICAgIGNvbnN0IHIgPSByZXNvbHZlVGFtVGluaFBlcmlvZF8oJ1F1w70nLCAnMjAyNi0wMS0xMCcpOwogICAgYXNzZXJ0RXF1YWwoci5mcm9tLCAnMjAyNS0xMC0wMScpOyBhc3NlcnRFcXVhbChyLnRvLCAnMjAyNS0xMi0zMScpOwogIH0pOwogIHRlc3QoJ3Jlc29sdmVUYW1UaW5oUGVyaW9kXzogNiBUaMOhbmcg4oCUIHRow6FuZyA5IChnaeG7r2EgbuG7rWEgY3Xhu5FpIG7Eg20pIC0+IGzDuWkgduG7gSBu4butYSBuxINtIMSQ4bqmVSAoSDEpJywgKCk9PnsKICAgIGNvbnN0IHIgPSByZXNvbHZlVGFtVGluaFBlcmlvZF8oJzYgVGjDoW5nJywgJzIwMjYtMDktMDUnKTsKICAgIGFzc2VydEVxdWFsKHIuZnJvbSwgJzIwMjYtMDEtMDEnKTsgYXNzZXJ0RXF1YWwoci50bywgJzIwMjYtMDYtMzAnKTsKICB9KTsKICB0ZXN0KCdyZXNvbHZlVGFtVGluaFBlcmlvZF86IDYgVGjDoW5nIOKAlCB0aMOhbmcgNiAoY3Xhu5FpIEgxKSAtPiBkw7luZyBIMSBISeG7hk4gSMOATkgnLCAoKT0+ewogICAgY29uc3QgciA9IHJlc29sdmVUYW1UaW5oUGVyaW9kXygnNiBUaMOhbmcnLCAnMjAyNi0wNi0xMCcpOwogICAgYXNzZXJ0RXF1YWwoci5mcm9tLCAnMjAyNi0wMS0wMScpOyBhc3NlcnRFcXVhbChyLnRvLCAnMjAyNi0wNi0zMCcpOwogIH0pOwogIHRlc3QoJ3Jlc29sdmVUYW1UaW5oUGVyaW9kXzogNiBUaMOhbmcg4oCUIHRow6FuZyAxMiAoY3Xhu5FpIEgyKSAtPiBkw7luZyBIMiBISeG7hk4gSMOATkgnLCAoKT0+ewogICAgY29uc3QgciA9IHJlc29sdmVUYW1UaW5oUGVyaW9kXygnNiBUaMOhbmcnLCAnMjAyNi0xMi0xMCcpOwogICAgYXNzZXJ0RXF1YWwoci5mcm9tLCAnMjAyNi0wNy0wMScpOyBhc3NlcnRFcXVhbChyLnRvLCAnMjAyNi0xMi0zMScpOwogIH0pOwogIHRlc3QoJ3Jlc29sdmVUYW1UaW5oUGVyaW9kXzogNiBUaMOhbmcg4oCUIHRow6FuZyAzICjEkeG6p3UgSDEpIC0+IGzDuWkgduG7gSBIMiBOxIJNIFRSxq/hu5pDICh24bqvdCBxdWEgbsSDbSknLCAoKT0+ewogICAgY29uc3QgciA9IHJlc29sdmVUYW1UaW5oUGVyaW9kXygnNiBUaMOhbmcnLCAnMjAyNi0wMy0xMCcpOwogICAgYXNzZXJ0RXF1YWwoci5mcm9tLCAnMjAyNS0wNy0wMScpOyBhc3NlcnRFcXVhbChyLnRvLCAnMjAyNS0xMi0zMScpOwogIH0pOwogIHRlc3QoJ3Jlc29sdmVUYW1UaW5oUGVyaW9kXzogTsSDbSDigJQgdGjDoW5nIGtow7RuZyBwaOG6o2kgMTIgLT4gbOG6pXkgTsSCTSBUUsav4buaQycsICgpPT57CiAgICBjb25zdCByID0gcmVzb2x2ZVRhbVRpbmhQZXJpb2RfKCdOxINtJywgJzIwMjYtMDktMDUnKTsKICAgIGFzc2VydEVxdWFsKHIuZnJvbSwgJzIwMjUtMDEtMDEnKTsgYXNzZXJ0RXF1YWwoci50bywgJzIwMjUtMTItMzEnKTsKICB9KTsKICB0ZXN0KCdyZXNvbHZlVGFtVGluaFBlcmlvZF86IE7Eg20g4oCUIHRow6FuZyAxMiAtPiBkw7luZyBOxIJNIEhJ4buGTiBIw4BOSCcsICgpPT57CiAgICBjb25zdCByID0gcmVzb2x2ZVRhbVRpbmhQZXJpb2RfKCdOxINtJywgJzIwMjYtMTItMjUnKTsKICAgIGFzc2VydEVxdWFsKHIuZnJvbSwgJzIwMjYtMDEtMDEnKTsgYXNzZXJ0RXF1YWwoci50bywgJzIwMjYtMTItMzEnKTsKICB9KTsKICB0ZXN0KCdyZXNvbHZlVGFtVGluaFBlcmlvZF86IHRo4budaSDEkWnhu4NtIGtow7RuZyB4w6FjIMSR4buLbmggaG/hurdjIG5nw6B5IHLhu5duZyAtPiBudWxsJywgKCk9PnsKICAgIGFzc2VydEVxdWFsKHJlc29sdmVUYW1UaW5oUGVyaW9kXygnVGhlbyBoaeG7h3UgbOG7sWMgQ1RLTScsICcyMDI2LTA5LTA1JyksIG51bGwpOwogICAgYXNzZXJ0RXF1YWwocmVzb2x2ZVRhbVRpbmhQZXJpb2RfKCdUaMOhbmcnLCAnJyksIG51bGwpOwogIH0pOwoKICAvKiAtLS0tIELDoGkgdGVzdCB0w61jaCBo4bujcDogY29tcHV0ZVRhbVRpbmhDaGlldEtoYXVfL2J1aWxkVGFtVGluaFRvbmdIb3BfIOKAlCBraeG7g20gdHJhIMSQw5pORyDEkWnhu4NtIGPhu5F0IGzDtWkKICAgICBj4bunYSB0w61uaCBuxINuZzogJSBs4bqleSB04burIEvhu7IgVEhBTSBDSEnhur5VICjEkcOjIGhvw6BuIHThuqV0KSBuaMawbmcgw4FQIEThu6RORyBsw6puIGRvYW5oIHPhu5EgVEjDgU5HIEhJ4buGTiBIw4BOSC4gLS0tLSAqLwogIGNvbnN0IGRiQmFja3VwVGFtVGluaCA9IERCOwogIHRyeXsKICAgIERCID0gewogICAgICBjYXRhbG9nOiB7IG1oY2s6IFsgeyBtYW1oOidTUDAwMScsIHRlbmhhbmc6J0zhu5FwIENhc3UgMDAxJywgdGh1b25naGlldTonQ0FTVScsIG5ob21raDonR1JQMScsIGxvYWlzcDonTk8nLCBkYWN0aW5oOidOTycsIGNvbmdkdW5nOidOTycsIHF1eWNhY2g6J05PJyB9IF0sIGdpYWNvbmdibzogW10gfSwKICAgICAgcHVyY2hhc2VzOiBbCiAgICAgICAgLy8gVGjDoW5nIGhp4buHbiBow6BuaCAoYXNPZiA9IDIwMjYtMDgtMjApIOKAlCBkw7luZyDEkeG7gyDDgVAgROG7pE5HIG3hu41pIHThu7cgbOG7hyBDSyBsw6puCiAgICAgICAgeyBuZ2F5OicyMDI2LTA4LTEwJywgbWFoYW5nOidTUDAwMScsIHNvbHVvbmc6MTAwLCBnaWF0cmk6MTAwMDAwMCwgbmhhY3VuZ2NhcDonTkNDIEEnIH0sCiAgICAgICAgLy8gS+G7syB0aGFtIGNoaeG6v3UgUXXDvSAoUTI6IDA0LTA2LzIwMjYpIFbDgCA2IFRow6FuZyAoSDE6IDAxLTA2LzIwMjYsIHRyw7luZyBsdcO0bikg4oCUIFNMPTIwMCDEkeG7pyDEkeG6oXQgYuG6rWMKICAgICAgICB7IG5nYXk6JzIwMjYtMDUtMTUnLCBtYWhhbmc6J1NQMDAxJywgc29sdW9uZzoyMDAsIGdpYXRyaToyMDAwMDAwLCBuaGFjdW5nY2FwOidOQ0MgQScgfSwKICAgICAgICAvLyBL4buzIHRoYW0gY2hp4bq/dSBOxINtIChuxINtIHRyxrDhu5tjLCAyMDI1KSDigJQgU0w9ODAgxJHhu6cgxJHhuqF0IGLhuq1jIFBfTkFNIChzbF9taW49NTApCiAgICAgICAgeyBuZ2F5OicyMDI1LTA2LTAxJywgbWFoYW5nOidTUDAwMScsIHNvbHVvbmc6ODAsIGdpYXRyaTo4MDAwMDAsIG5oYWN1bmdjYXA6J05DQyBBJyB9LAogICAgICBdLAogICAgICBwcm9ncmFtczogWwogICAgICAgIHsgaWQ6J1BfVEhBTkcnLCBtYV9jaGlldGtoYXU6J0NUX1RIQU5HJywgdGhvaWRpZW1fY2s6J1Row6FuZycsIGRpbV90aHVvbmdoaWV1OidDQVNVJywgZGltX25ob21raDonTk8nLCBkaW1fbG9haXNwOidOTycsIGRpbV9kYWN0aW5oOidOTycsIGRpbV9jb25nZHVuZzonTk8nLCBkaW1fcXV5Y2FjaDonTk8nLAogICAgICAgICAgc2xfbWluOjUwLCBzbF9tYXg6MCwgZHRfbWluOjAsIGR0X21heDowLCBkaWV1a2llbl9zbGR0OidBTkQnLCBodF9jaGlldGtoYXU6J1BFUkNFTlRfR0lBTVVBJywgdGxfY2s6NSwgbWFfZG9hbmh0aHU6JycsIGR0X2tlaG9hY2g6MCB9LAogICAgICAgIHsgaWQ6J1BfUVVZJywgbWFfY2hpZXRraGF1OidDVF9RVVknLCB0aG9pZGllbV9jazonUXXDvScsIGRpbV90aHVvbmdoaWV1OidDQVNVJywgZGltX25ob21raDonTk8nLCBkaW1fbG9haXNwOidOTycsIGRpbV9kYWN0aW5oOidOTycsIGRpbV9jb25nZHVuZzonTk8nLCBkaW1fcXV5Y2FjaDonTk8nLAogICAgICAgICAgc2xfbWluOjE1MCwgc2xfbWF4OjAsIGR0X21pbjowLCBkdF9tYXg6MCwgZGlldWtpZW5fc2xkdDonQU5EJywgaHRfY2hpZXRraGF1OidQRVJDRU5UX0dJQU1VQScsIHRsX2NrOjgsIG1hX2RvYW5odGh1OicnLCBkdF9rZWhvYWNoOjAgfSwKICAgICAgICB7IGlkOidQXzZUSEFORycsIG1hX2NoaWV0a2hhdTonQ1RfNlRIQU5HJywgdGhvaWRpZW1fY2s6JzYgVGjDoW5nJywgZGltX3RodW9uZ2hpZXU6J0NBU1UnLCBkaW1fbmhvbWtoOidOTycsIGRpbV9sb2Fpc3A6J05PJywgZGltX2RhY3Rpbmg6J05PJywgZGltX2NvbmdkdW5nOidOTycsIGRpbV9xdXljYWNoOidOTycsCiAgICAgICAgICBzbF9taW46MTUwLCBzbF9tYXg6MCwgZHRfbWluOjAsIGR0X21heDowLCBkaWV1a2llbl9zbGR0OidBTkQnLCBodF9jaGlldGtoYXU6J1BFUkNFTlRfR0lBTVVBJywgdGxfY2s6MTAsIG1hX2RvYW5odGh1OicnLCBkdF9rZWhvYWNoOjAgfSwKICAgICAgICB7IGlkOidQX05BTScsIG1hX2NoaWV0a2hhdTonQ1RfTkFNJywgdGhvaWRpZW1fY2s6J07Eg20nLCBkaW1fdGh1b25naGlldTonQ0FTVScsIGRpbV9uaG9ta2g6J05PJywgZGltX2xvYWlzcDonTk8nLCBkaW1fZGFjdGluaDonTk8nLCBkaW1fY29uZ2R1bmc6J05PJywgZGltX3F1eWNhY2g6J05PJywKICAgICAgICAgIHNsX21pbjo1MCwgc2xfbWF4OjAsIGR0X21pbjowLCBkdF9tYXg6MCwgZGlldWtpZW5fc2xkdDonQU5EJywgaHRfY2hpZXRraGF1OidQRVJDRU5UX0dJQU1VQScsIHRsX2NrOjMsIG1hX2RvYW5odGh1OicnLCBkdF9rZWhvYWNoOjAgfSwKICAgICAgXSwKICAgICAgZGlldWtpZW46IFtdLCBkb2FuaHRodTogW10sIGtlaG9hY2g6IFtdLCBrZWhvYWNoQ2hpVGlldDogW10sIHByb2dyYW1zTHV5S2U6IFtdLCByZXBvcnRDS1RIOiBbXQogICAgfTsKICAgIHRlc3QoJ2NvbXB1dGVUYW1UaW5oQ2hpZXRLaGF1XzogbeG7l2kgbG/huqFpIENLIGzhuqV5ICUgdOG7qyDEkMOaTkcga+G7syB0aGFtIGNoaeG6v3UgbmjGsG5nIMOhcCBsw6puIGRvYW5oIHPhu5EgVEjDgU5HIEhJ4buGTiBIw4BOSCcsICgpPT57CiAgICAgIGNvbnN0IHJlcyA9IGNvbXB1dGVUYW1UaW5oQ2hpZXRLaGF1XygnMjAyNi0wOC0yMCcpOwogICAgICBjb25zdCByb3cgPSByZXMucm93cy5maW5kKHI9PnIubWFoYW5nPT09J1NQMDAxJyk7CiAgICAgIGFzc2VydFRydWUoISFyb3csICdwaOG6o2kgY8OzIGTDsm5nIGNobyBTUDAwMSAoY8OzIHBow6F0IHNpbmggbXVhIHRow6FuZyA4LzIwMjYpJyk7CiAgICAgIGFzc2VydEVxdWFsKHJvdy5kdCwgMTAwMDAwMCwgJ0RvYW5oIHPhu5EgbXVhIHBo4bqjaSBM4bqkWSBUSMOBTkcgSEnhu4ZOIEjDgE5IICg4LzIwMjYgPSAxLjAwMC4wMDApLCBraMO0bmcgY+G7mW5nIGThu5NuIGPDoWMga+G7syBraMOhYycpOwogICAgICBhc3NlcnRDbG9zZShyb3cuY2tfdGhhbmcsIDUwMDAwLCAnQ0sgdGjDoW5nID0gNSUgeCAxLjAwMC4wMDAgKFRow6FuZyBsdcO0biBs4bqleSBjaMOtbmggdGjDoW5nIGhp4buHbiBow6BuaCBsw6BtIGvhu7MgdGhhbSBjaGnhur91KScpOwogICAgICBhc3NlcnRDbG9zZShyb3cuY2tfcXV5LCA4MDAwMCwgJ0NLIHF1w70gPSA4JSAoxJHhuqF0IGLhuq1jIOG7nyBRMiBuaOG7nSBTTD0yMDApIHggMS4wMDAuMDAwICjDgVAgbMOqbiBEVCB0aMOhbmcgOCwga2jDtG5nIHBo4bqjaSBEVCBj4bunYSBRMiknKTsKICAgICAgYXNzZXJ0Q2xvc2Uocm93LmNrXzZ0aGFuZywgMTAwMDAwLCAnQ0sgMS8yIG7Eg20gPSAxMCUgKMSR4bqhdCBi4bqtYyDhu58gSDEgbmjhu50gU0w9MjAwKSB4IDEuMDAwLjAwMCcpOwogICAgICBhc3NlcnRDbG9zZShyb3cuY2tfbmFtLCAzMDAwMCwgJ0NLIG7Eg20gPSAzJSAoxJHhuqF0IGLhuq1jIOG7nyBuxINtIDIwMjUgbmjhu50gU0w9ODApIHggMS4wMDAuMDAwJyk7CiAgICAgIGFzc2VydEVxdWFsKHJvdy5ja19raGFjLCAwLCAna2jDtG5nIGPhuqV1IGjDrG5oIENoaeG6v3Qga2jhuqV1IGzFqXkga+G6vyBuw6BvIHRyb25nIHRlc3QgbsOgeSAtPiBDSyBraMOhYyA9IDAnKTsKICAgICAgYXNzZXJ0Q2xvc2Uocm93LnRvbmdUaWVuQ0ssIDI2MDAwMCwgJ1Thu5VuZyA9IDUwLjAwMCs4MC4wMDArMTAwLjAwMCszMC4wMDAnKTsKICAgICAgYXNzZXJ0Q2xvc2Uocm93LnBjdEdpYU11YSwgMjYsICdU4buVbmcgJUNLL2dpw6EgbXVhID0gMjYwLjAwMC8xLjAwMC4wMDAgeCAxMDAnKTsKICAgIH0pOwogICAgdGVzdCgnY29tcHV0ZVRhbVRpbmhDaGlldEtoYXVfOiBr4buzIHRoYW0gY2hp4bq/dSB0cuG6oyB24buBIMSRw7puZyBjaG8gdOG7q25nIGxv4bqhaSAoZMO5bmcgxJHhu4MgaGnhu4NuIHRo4buLL2Now7ogdGjDrWNoKScsICgpPT57CiAgICAgIGNvbnN0IHJlcyA9IGNvbXB1dGVUYW1UaW5oQ2hpZXRLaGF1XygnMjAyNi0wOC0yMCcpOwogICAgICBjb25zdCBieUtleSA9IHt9OyByZXMuY2F0UmVmUGVyaW9kcy5mb3JFYWNoKGM9PmJ5S2V5W2Mua2V5XT1jLnJlZik7CiAgICAgIGFzc2VydEVxdWFsKGJ5S2V5LmNrX3RoYW5nLmZyb20sICcyMDI2LTA4LTAxJyk7IGFzc2VydEVxdWFsKGJ5S2V5LmNrX3RoYW5nLnRvLCAnMjAyNi0wOC0zMScpOwogICAgICBhc3NlcnRFcXVhbChieUtleS5ja19xdXkuZnJvbSwgJzIwMjYtMDQtMDEnKTsgYXNzZXJ0RXF1YWwoYnlLZXkuY2tfcXV5LnRvLCAnMjAyNi0wNi0zMCcpOwogICAgICBhc3NlcnRFcXVhbChieUtleS5ja182dGhhbmcuZnJvbSwgJzIwMjYtMDEtMDEnKTsgYXNzZXJ0RXF1YWwoYnlLZXkuY2tfNnRoYW5nLnRvLCAnMjAyNi0wNi0zMCcpOwogICAgICBhc3NlcnRFcXVhbChieUtleS5ja19uYW0uZnJvbSwgJzIwMjUtMDEtMDEnKTsgYXNzZXJ0RXF1YWwoYnlLZXkuY2tfbmFtLnRvLCAnMjAyNS0xMi0zMScpOwogICAgfSk7CiAgICB0ZXN0KCdidWlsZFRhbVRpbmhUb25nSG9wXzogZ+G7mXAgxJHDum5nIHRoZW8gTmjDs20gaMOgbmcgKyBUaMawxqFuZyBoaeG7h3UsIHThu5VuZyBraOG7m3AgduG7m2kgZMOybmcgY2hpIHRp4bq/dCcsICgpPT57CiAgICAgIGNvbnN0IHJlcyA9IGNvbXB1dGVUYW1UaW5oQ2hpZXRLaGF1XygnMjAyNi0wOC0yMCcpOwogICAgICBjb25zdCBncm91cHMgPSBidWlsZFRhbVRpbmhUb25nSG9wXyhyZXMpOwogICAgICBhc3NlcnRFcXVhbChncm91cHMubGVuZ3RoLCAxLCAnY2jhu4kgMSBtw6MgaMOgbmcgLT4gMSBuaMOzbSAoR1JQMSArIENBU1UpJyk7CiAgICAgIGFzc2VydEVxdWFsKGdyb3Vwc1swXS5uaG9ta2gsICdHUlAxJyk7IGFzc2VydEVxdWFsKGdyb3Vwc1swXS50aHVvbmdoaWV1LCAnQ0FTVScpOwogICAgICBhc3NlcnRFcXVhbChncm91cHNbMF0uc29NYUhhbmcsIDEpOwogICAgICBhc3NlcnRDbG9zZShncm91cHNbMF0udG9uZ1RpZW5DSywgMjYwMDAwLCAndOG7lW5nIG5ow7NtIHBo4bqjaSBraOG7m3AgdOG7lW5nIGTDsm5nIGNoaSB0aeG6v3QgZHV5IG5o4bqldCB0aHXhu5ljIG5ow7NtIMSRw7MnKTsKICAgIH0pOwoKICAgIC8qIC0tLS0gQsOhbyBjw6FvIDQ6IENLIGPGoSBi4bqjbiAoZ2nDoSBjw7RuZyBi4buRIC0gZ2nDoSBtdWEpICsgYnVpbGRUYW1UaW5oVGhlb0dpYUNvbmdCb18gLS0tLSAqLwogICAgdGVzdCgnY29tcHV0ZVRhbVRpbmhDaGlldEtoYXVfOiBja19jb2JhbiA9IERvYW5oIHPhu5EgZ2nDoSBjw7RuZyBi4buRIC0gRG9hbmggc+G7kSBnacOhIG11YSwgdG9uZ1RpZW5DS0Z1bGwgPSBja19jb2JhbiArIHRvbmdUaWVuQ0snLCAoKT0+ewogICAgICBEQi5jYXRhbG9nID0gewogICAgICAgIG1oY2s6IFsgeyBtYW1oOidTUDAwMScsIG1hbWhfbmNjOidNQU1IMDEnLCB0ZW5oYW5nOidM4buRcCBDYXN1IDAwMScsIHRodW9uZ2hpZXU6J0NBU1UnLCBuaG9ta2g6J0dSUDEnLCBsb2Fpc3A6J05PJywgZGFjdGluaDonTk8nLCBjb25nZHVuZzonTk8nLCBxdXljYWNoOidOTycgfSBdLAogICAgICAgIGdpYWNvbmdibzogWyB7IG1hbWhfbmNjOidNQU1IMDEnLCBnaWFfY29uZ2JvOjEyMDAwLCBoaWV1bHVjdHU6JycsIGhpZXVsdWNkZW46JycgfSBdCiAgICAgIH07CiAgICAgIERCLnB1cmNoYXNlcyA9IFsgeyBuZ2F5OicyMDI2LTA4LTEwJywgbWFoYW5nOidTUDAwMScsIHNvbHVvbmc6MTAwLCBnaWF0cmk6MTAwMDAwMCwgbmhhY3VuZ2NhcDonTkNDIEEnIH0gXTsKICAgICAgREIucHJvZ3JhbXMgPSBbXTsgLy8ga2jDtG5nIGPhuqV1IGjDrG5oIGNoxrDGoW5nIHRyw6xuaCBuw6BvIC0+IGPDtCBs4bqtcCDEkcO6bmcgcGjhuqduIENLIGPGoSBi4bqjbgogICAgICBjb25zdCByZXMgPSBjb21wdXRlVGFtVGluaENoaWV0S2hhdV8oJzIwMjYtMDgtMjAnKTsKICAgICAgY29uc3Qgcm93ID0gcmVzLnJvd3MuZmluZChyPT5yLm1haGFuZz09PSdTUDAwMScpOwogICAgICBhc3NlcnRUcnVlKCEhcm93KTsKICAgICAgYXNzZXJ0RXF1YWwocm93LmR0LCAxMDAwMDAwKTsgYXNzZXJ0RXF1YWwocm93LmR0R2lhQ29uZ0JvLCAxMjAwMDAwLCAnRG9hbmggc+G7kSBnacOhIGPDtG5nIGLhu5EgPSAxMDAgeCAxMi4wMDAnKTsKICAgICAgYXNzZXJ0Q2xvc2Uocm93LmNrX2NvYmFuLCAyMDAwMDAsICdDSyBjxqEgYuG6o24gPSAxLjIwMC4wMDAgLSAxLjAwMC4wMDAnKTsKICAgICAgYXNzZXJ0Q2xvc2Uocm93LnBjdENvQmFuLCAyMDAwMDAvMTIwMDAwMCoxMDApOwogICAgICBhc3NlcnRFcXVhbChyb3cudG9uZ1RpZW5DSywgMCwgJ2tow7RuZyBjw7MgY2jGsMahbmcgdHLDrG5oIG7DoG8gLT4gQ0sgdGhlbyBjaMawxqFuZyB0csOsbmggPSAwJyk7CiAgICAgIGFzc2VydENsb3NlKHJvdy50b25nVGllbkNLRnVsbCwgMjAwMDAwLCAndG9uZ1RpZW5DS0Z1bGwgPSBja19jb2JhbiAoMjAwLjAwMCkgKyB0b25nVGllbkNLICgwKScpOwogICAgICBhc3NlcnRDbG9zZShyb3cucGN0R2lhQ29uZ0JvRnVsbCwgMjAwMDAwLzEyMDAwMDAqMTAwKTsKCiAgICAgIGNvbnN0IGdyb3VwczQgPSBidWlsZFRhbVRpbmhUaGVvR2lhQ29uZ0JvXyhyZXMpOwogICAgICBhc3NlcnRFcXVhbChncm91cHM0Lmxlbmd0aCwgMSk7CiAgICAgIGFzc2VydEVxdWFsKGdyb3VwczRbMF0uaXRlbXMubGVuZ3RoLCAxKTsKICAgICAgYXNzZXJ0Q2xvc2UoZ3JvdXBzNFswXS5ja19jb2JhbiwgMjAwMDAwLCAndOG7lW5nIG5ow7NtIHBo4bqjaSBraOG7m3AgQ0sgY8ahIGLhuqNuIGPhu6dhIGTDsm5nIGNoaSB0aeG6v3QgZHV5IG5o4bqldCcpOwogICAgICBhc3NlcnRDbG9zZShncm91cHM0WzBdLnRvbmdUaWVuQ0tGdWxsLCAyMDAwMDApOwogICAgfSk7CiAgICB0ZXN0KCdjb21wdXRlVGFtVGluaENoaWV0S2hhdV86IG3hurd0IGjDoG5nIFRISeG6vlUgZ2nDoSBjw7RuZyBi4buRIC0+IHThuqFtIGTDuW5nIGdpw6EgbXVhIGLDrG5oIHF1w6JuICsgxJHDoW5oIGThuqV1IGNhbmhCYW9HaWFDb25nQm8vZ2hpQ2h1R2lhQ29uZ0JvJywgKCk9PnsKICAgICAgREIuY2F0YWxvZyA9IHsKICAgICAgICBtaGNrOiBbIHsgbWFtaDonU1AwMDInLCBtYW1oX25jYzonTUFNSDAyJywgdGVuaGFuZzonTOG7kXAgQ2FzdSAwMDInLCB0aHVvbmdoaWV1OidDQVNVJywgbmhvbWtoOidHUlAxJywgbG9haXNwOidOTycsIGRhY3Rpbmg6J05PJywgY29uZ2R1bmc6J05PJywgcXV5Y2FjaDonTk8nIH0gXSwKICAgICAgICBnaWFjb25nYm86IFtdIC8vIEtIw5RORyBraGFpIGdpw6EgY8O0bmcgYuG7kSBjaG8gTUFNSDAyIC0+IHMuY2JTTCBz4bq9ID0gMCBjaG8gU1AwMDIKICAgICAgfTsKICAgICAgREIucHVyY2hhc2VzID0gWyB7IG5nYXk6JzIwMjYtMDgtMTAnLCBtYWhhbmc6J1NQMDAyJywgc29sdW9uZzo1MCwgZ2lhdHJpOjUwMDAwMCwgbmhhY3VuZ2NhcDonTkNDIEEnIH0gXTsKICAgICAgREIucHJvZ3JhbXMgPSBbXTsKICAgICAgY29uc3QgcmVzID0gY29tcHV0ZVRhbVRpbmhDaGlldEtoYXVfKCcyMDI2LTA4LTIwJyk7CiAgICAgIGNvbnN0IHJvdyA9IHJlcy5yb3dzLmZpbmQocj0+ci5tYWhhbmc9PT0nU1AwMDInKTsKICAgICAgYXNzZXJ0VHJ1ZSghIXJvdyk7CiAgICAgIGFzc2VydFRydWUocm93LmNhbmhCYW9HaWFDb25nQm8sICdwaOG6o2kgxJHDoW5oIGThuqV1IGPhuqNuaCBiw6FvIHbDrCBraMO0bmcgY8OzIGdpw6EgY8O0bmcgYuG7kSBuw6BvIGto4bubcCDEkcaw4bujYycpOwogICAgICBhc3NlcnRUcnVlKCEhcm93LmdoaUNodUdpYUNvbmdCbywgJ3Bo4bqjaSBjw7MgZ2hpIGNow7ogZ2nhuqNpIHRow61jaCBsw70gZG8nKTsKICAgICAgYXNzZXJ0RXF1YWwocm93LmR0R2lhQ29uZ0JvLCByb3cuZHQsICd0aGnhur91IGdpw6EgY8O0bmcgYuG7kSAtPiBEb2FuaCBz4buRIGdpw6EgY8O0bmcgYuG7kSBwaOG6o2kgVOG6oE0gbOG6pXkgxJHDum5nIGLhurFuZyBEb2FuaCBz4buRIGdpw6EgbXVhIChnacOhIHRydW5nIGLDrG5oIGjDs2EgxJHGoW4pJyk7CiAgICAgIGFzc2VydEVxdWFsKHJvdy5ja19jb2JhbiwgMCwgJ2tow7RuZyB04buxIHN1eSBkaeG7hW4gQ0sgY8ahIGLhuqNuIGtoaSB0aGnhur91IGThu68gbGnhu4d1IHRo4bqtdCAtPiBwaOG6o2kgPSAwLCBraMO0bmcgcGjhuqNpIHPhu5Egw6JtJyk7CiAgICB9KTsKICB9IGZpbmFsbHkgewogICAgREIgPSBkYkJhY2t1cFRhbVRpbmg7CiAgfQoKICBjb25zdCBva0NvdW50ID0gcmVzdWx0cy5maWx0ZXIocj0+ci5vaykubGVuZ3RoOwogIG9wZW5Nb2RhbChgCiAgICA8aDM+S+G6v3QgcXXhuqMgdOG7sSBraeG7g20gdGjhu60gY8O0bmcgdGjhu6ljIGNoaeG6v3Qga2jhuqV1PC9oMz4KICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tbm90ZSIgc3R5bGU9ImJvcmRlci1jb2xvcjoke29rQ291bnQ9PT1yZXN1bHRzLmxlbmd0aD8ndmFyKC0tZW1lcmFsZCknOid2YXIoLS1yZWQpJ307Y29sb3I6JHtva0NvdW50PT09cmVzdWx0cy5sZW5ndGg/J3ZhcigtLWVtZXJhbGQtZGFyayknOid2YXIoLS1yZWQpJ30iPgogICAgICAke29rQ291bnQ9PT1yZXN1bHRzLmxlbmd0aCA/IGDinJMgVG/DoG4gYuG7mSAke3Jlc3VsdHMubGVuZ3RofS8ke3Jlc3VsdHMubGVuZ3RofSB0ZXN0IFBBU1MuIEPDoWMgY8O0bmcgdGjhu6ljIGPhu5F0IGzDtWkgxJFhbmcgaG/huqF0IMSR4buZbmcgxJHDum5nIG5oxrAgdGhp4bq/dCBr4bq/LmAgOiBg4pqgICR7cmVzdWx0cy5sZW5ndGgtb2tDb3VudH0vJHtyZXN1bHRzLmxlbmd0aH0gdGVzdCBGQUlMIOKAlCBjw7MgY8O0bmcgdGjhu6ljIMSRYW5nIHTDrW5oIFNBSSwgeGVtIGNoaSB0aeG6v3QgYsOqbiBkxrDhu5tpIHRyxrDhu5tjIGtoaSBkw7luZyBiw6FvIGPDoW8uYH0KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0ibW9ubyIgc3R5bGU9ImZvbnQtc2l6ZToxMi41cHg7cGFkZGluZzoxMHB4O2JhY2tncm91bmQ6dmFyKC0tc2xhdGUtNTApO2JvcmRlci1yYWRpdXM6OHB4O2JvcmRlcjoxcHggc29saWQgdmFyKC0tc2xhdGUtMjAwKTttYXgtaGVpZ2h0OjM2MHB4O292ZXJmbG93OmF1dG8iPgogICAgICAke3Jlc3VsdHMubWFwKHI9PmA8ZGl2IHN0eWxlPSJwYWRkaW5nOjNweCAwO2NvbG9yOiR7ci5vaz8ndmFyKC0tZW1lcmFsZC1kYXJrKSc6J3ZhcigtLXJlZCknfSI+JHtyLm9rPyfinJMnOifinJcnfSAke2VzYyhyLm5hbWUpfSR7ci5vaz8nJzonIOKAlCAnK2VzYyhyLmVycil9PC9kaXY+YCkuam9pbignJyl9CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImJ0bi1yb3ciIHN0eWxlPSJqdXN0aWZ5LWNvbnRlbnQ6ZmxleC1lbmQ7bWFyZ2luLXRvcDo4cHgiPjxidXR0b24gY2xhc3M9ImJ0biBidG4tcHJpbWFyeSIgb25jbGljaz0iY2xvc2VNb2RhbCgpIj7EkMOzbmc8L2J1dHRvbj48L2Rpdj4KICBgKTsK";
function getFormulaTestsSrc() {
  return FORMULA_TESTS_SRC_B64_;
}

// ============================ MẪU IMPORT EXCEL — TỪ GOOGLE SHEET TEMPLATE (26/08/2026) ============================
// LỊCH SỬ: (25/08/2026) file mẫu vốn dựng bằng openpyxl (giữ màu/chú thích mà XLSX.js không ghi được),
// nhúng thẳng base64 vào Index.html → góp phần đẩy Index.html vượt ngưỡng ~1MiB, gây trắng màn hình
// webapp (xem phương án mục 18). (26/08/2026, bước 1) chuyển base64 sang Code.gs — hết trắng màn hình,
// nhưng vẫn là 1 chuỗi TĨNH: mỗi lần cần sửa mẫu (thêm cột, đổi ghi chú...) phải nhờ Claude dựng lại file
// rồi dán lại base64 mới, không ai tự sửa trực tiếp được.
// (26/08/2026, bước 2 — THEO YÊU CẦU "sao không tạo từ Google Sheet mẫu"): đổi hẳn sang lấy nguồn từ 1
// GOOGLE SHEET THẬT — tự tạo 1 lần (convert từ file .xlsx gốc, giữ nguyên màu/chú thích/đóng băng dòng
// đầu), lưu ID trong Script Properties. Từ nay: mở đúng Sheet đó (nút "Xem/Sửa file mẫu trên Google
// Sheet" ở màn Trợ giúp → Mẫu import Excel) để tự sửa trực tiếp — lần tải file mẫu (.xlsx) tiếp theo TỰ
// ĐỘNG xuất đúng bản mới nhất, không cần nhờ Claude regenerate base64 nữa.
// YÊU CẦU KỸ THUẬT: cần Advanced Service "Drive API" (đã khai ở appsscript.json,
// dependencies.enabledAdvancedServices) — lần đầu deploy lại sau thay đổi này, Apps Script sẽ yêu cầu
// cấp quyền Drive mới (bình thường, do đổi sang thao tác tạo/đọc file trên Drive thay vì chỉ hằng số).
const MAU_TEMPLATE_PROP_KEY_ = 'MAU_TEMPLATE_SHEET_ID';

function getMauNhapLieuXlsxB64() {
  const sheetId = ensureMauTemplateSheetId_();
  const token = ScriptApp.getOAuthToken();
  const url = 'https://docs.google.com/spreadsheets/d/' + sheetId + '/export?format=xlsx';
  const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error('Không xuất được file mẫu từ Google Sheet template (mã lỗi ' + resp.getResponseCode() + '). Kiểm tra lại quyền truy cập Sheet ID ' + sheetId + ', hoặc mở Apps Script chạy lại sau khi cấp quyền Drive.');
  }
  return Utilities.base64Encode(resp.getBlob().getBytes());
}

// Trả về link Google Sheet template để người dùng tự mở và chỉnh sửa trực tiếp (thêm/xóa cột, đổi màu,
// sửa ghi chú...) — không cần qua Claude nữa. Dùng cho nút "Xem/Sửa file mẫu trên Google Sheet".
function getMauNhapLieuTemplateSheetUrl() {
  const sheetId = ensureMauTemplateSheetId_();
  return 'https://docs.google.com/spreadsheets/d/' + sheetId + '/edit';
}

// Tạo (nếu chưa có) hoặc trả về ID của Google Sheet template. Sheet này SỐNG LÂU DÀI trong Drive của tài
// khoản đang triển khai Web App (executeAs: USER_DEPLOYING ở appsscript.json → luôn cùng 1 tài khoản dù
// ai mở webapp), không phải file tạm — sửa trực tiếp trên Sheet này sẽ phản ánh ngay ở lần tải kế tiếp.
// Nếu ID đã lưu nhưng Sheet bị xóa/mất quyền truy cập, tự tạo lại từ bản gốc (seed) bên dưới.
function ensureMauTemplateSheetId_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty(MAU_TEMPLATE_PROP_KEY_);
  if (savedId) {
    try {
      const f = DriveApp.getFileById(savedId);
      if (!f.isTrashed()) return savedId;
    } catch (e) {
      // Sheet không còn tồn tại hoặc mất quyền truy cập — tạo lại bên dưới, không chặn người dùng.
    }
  }
  const newId = buildMauTemplateSheetFromXlsx_();
  props.setProperty(MAU_TEMPLATE_PROP_KEY_, newId);
  Logger.log('Đã tạo Google Sheet template mẫu import mới — mở tại: https://docs.google.com/spreadsheets/d/' + newId + '/edit');
  return newId;
}

// Chuyển đúng file .xlsx mẫu gốc (dựng sẵn bằng openpyxl, có tô màu cột bắt buộc, tô vàng dòng mẫu kèm
// chú thích, đóng băng dòng tiêu đề) thành 1 Google Sheet thật trong Drive, dùng Advanced Drive Service
// (v2) với convert:true để Google tự chuyển đổi định dạng sang đúng kiểu Google Sheets tương ứng — chỉ
// chạy đúng 1 lần (kết quả được cache theo ID ở Script Properties, xem ensureMauTemplateSheetId_).
function buildMauTemplateSheetFromXlsx_() {
  const bytes = Utilities.base64Decode(MAU_NHAP_LIEU_XLSX_B64_SEED_);
  const blob = Utilities.newBlob(bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'MAU_NHAP_LIEU_CHIET_KHAU_CASUMINA.xlsx');
  const file = Drive.Files.insert(
    { title: 'MAU_NHAP_LIEU_CHIET_KHAU_CASUMINA (TEMPLATE — sửa trực tiếp tại đây)', mimeType: MimeType.GOOGLE_SHEETS },
    blob,
    { convert: true }
  );
  return file.id;
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

// ============================ HƯỚNG DẪN SỬ DỤNG — NỘI DUNG TẢI QUA google.script.run (28/08/2026) ============================
// LỊCH SỬ: renderHuongDan() ở Index.html (client) từng dựng thẳng ~46,5KB HTML tĩnh (2 tab "Quy trình" +
// "Nghiệp vụ") ngay trong bundle gửi trình duyệt — cùng cơ chế từng đẩy Index.html gần chạm ngưỡng
// ~1.048.576 byte khiến webapp trắng màn hình (xem phương án mục 18, đã áp dụng y hệt hướng xử lý cho
// file mẫu Excel ở mục 19). Nay chuyển toàn bộ nội dung tĩnh này sang Code.gs, chỉ tải về đúng lúc người
// dùng mở tab Hướng dẫn sử dụng — Index.html chỉ còn giữ lại phần khung (tab bar, khung "đang tải...",
// gắn lại sự kiện click) và gọi google.script.run.getHuongDanContentHtml(sub) để lấy HTML đầy đủ.
// LƯU Ý TRIỂN KHAI: hàm flowDiag_()/dữ liệu FLOW_DATA_ (dùng cho sơ đồ quy trình ở Bước 13) đã được RENDER
// SẴN thành HTML tĩnh ngay trong HD_CONTENT_NGHIEPVU_ bên dưới (không cần porting flowDiag_/esc sang
// Code.gs) — dữ liệu 10 sơ đồ này không đổi theo người dùng/thời gian nên render tĩnh 1 lần là đủ; muốn
// sửa nội dung sơ đồ thì sửa trực tiếp đoạn HTML tương ứng trong HD_CONTENT_NGHIEPVU_.
function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function(c) {
    return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c];
  });
}

// ============================ BẢO TRÌ HỆ THỐNG ============================
// THEO YÊU CẦU: bộ 3 công cụ bảo trì, xây dựng trực tiếp từ các lớp lỗi ÂM THẦM đã thực tế phát hiện khi
// rà soát dữ liệu thật của người dùng trong phiên làm việc này (dim_nhomkh của Mã chiết khấu bị lệch độ
// chính xác số so với Mã doanh thu liên kết — nghi do cell từng bị Google Sheets tự đổi định dạng
// ngày/giờ rồi sanitizeCellValue_() quy đổi ngược qua serialFromDate_() sinh ra phần thập phân dư, xem
// giải thích tại 2 hàm đó — khiến dimsMatch() so KHỚP CHUỖI tuyệt đối trả về "không khớp" cho TOÀN BỘ mã
// hàng, ra 0đ chiết khấu không cảnh báo):
//  (1) Kiểm tra sức khỏe dữ liệu — quét toàn bộ PROGRAMS/DOANHTHU/CHUONGTRINH/MHCK/PURCHASES, phát hiện
//      SỚM đúng lớp lỗi trên (và vài lớp lỗi cấu hình phổ biến khác) thay vì phải đợi báo cáo ra 0đ.
//  (2) Sao lưu các sheet cấu hình quan trọng (bản sao kèm ngày giờ) + tự dọn bớt bản backup cũ.
//  (3) Trigger chạy hằng ngày gọi cả 2 việc trên, gửi email tóm tắt cho người triển khai nếu có vấn đề.
// Cả 3 đều CHỈ ĐỌC (trừ việc tạo/xoá sheet backup ở mục 2) — không đụng tới dữ liệu nghiệp vụ đang có.

const MAINTENANCE_BACKUP_KEEP_ = 10; // giữ tối đa 10 bản sao lưu gần nhất CHO MỖI sheet cấu hình
// Các sheet cấu hình "ít dòng, dễ sửa tay, hậu quả nặng nếu hỏng" — KHÔNG gồm PURCHASES/GIACONGBO/
// REPORT_CKTH/REPORT_CKCT (quá nhiều dòng, đã có "Xuất Excel" riêng, sao lưu ở đây tốn thời gian không
// cần thiết cho mục tiêu "cấu hình chính sách chiết khấu").
const MAINTENANCE_BACKUP_SHEET_KEYS_ = ['PROGRAMS', 'DIEUKIEN', 'DOANHTHU', 'KEHOACH', 'KEHOACH_CHITIET', 'CHUONGTRINH', 'LUYKE', 'MHCK'];
const MAINTENANCE_STALE_DAYS_ = 90; // "không phát sinh mua hàng khớp phạm vi trong X ngày gần nhất" khi xét mã chiết khấu nghi ngờ

function pad2_(n) { return String(n).length < 2 ? '0' + n : String(n); }
function backupTimestamp_(d) {
  d = d || new Date();
  return d.getFullYear() + pad2_(d.getMonth() + 1) + pad2_(d.getDate()) + '_' + pad2_(d.getHours()) + pad2_(d.getMinutes());
}

// ---- (2) Sao lưu + dọn dẹp ----
// Copy từng sheet cấu hình sang 1 sheet mới "<TÊN>_BACKUP_<yyyyMMdd_HHmm>" (định dạng timestamp cố định
// độ rộng nên so chuỗi = so thời gian, dùng được để sắp xếp mới→cũ mà không cần parse ngày), rồi xoá bớt
// các bản backup CŨ NHẤT của CHÍNH sheet đó nếu vượt quá MAINTENANCE_BACKUP_KEEP_ bản.
function backupKeyConfigSheets_() {
  const ss = SpreadsheetApp.getActive();
  const ts = backupTimestamp_();
  const created = [];
  MAINTENANCE_BACKUP_SHEET_KEYS_.forEach(key => {
    const name = SHEETS[key];
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() <= 1) return; // sheet trống hoặc chưa tồn tại — không cần sao lưu
    const backupName = name + '_BACKUP_' + ts;
    const copy = sh.copyTo(ss);
    copy.setName(backupName);
    copy.hideSheet();
    created.push(backupName);
    pruneOldBackups_(ss, name);
  });
  return created;
}
function pruneOldBackups_(ss, baseName) {
  const prefix = baseName + '_BACKUP_';
  const candidates = ss.getSheets()
    .filter(sh => sh.getName().indexOf(prefix) === 0)
    .sort((a, b) => a.getName() < b.getName() ? 1 : -1); // tên có timestamp tăng dần -> mới nhất đứng trước
  candidates.slice(MAINTENANCE_BACKUP_KEEP_).forEach(sh => ss.deleteSheet(sh));
}
// Gọi được trực tiếp từ nút "Sao lưu ngay" trên web.
function runBackupNow() {
  const created = backupKeyConfigSheets_();
  return { created, count: created.length };
}

// ---- (1) Kiểm tra sức khỏe dữ liệu ----
// Đọc THẲNG từ Sheet qua readSheetRows_() (không phụ thuộc dữ liệu Frontend đã tải) — chạy được cả từ
// trigger (không có trình duyệt) lẫn khi người dùng bấm nút trên web.
const HEALTHCHECK_DIM_KEYS_ = ['thuonghieu', 'nhomkh', 'loaisp', 'dactinh', 'congdung', 'quycach'];
function healthCheckDimsMatchLoose_(itemDims, ruleDims) {
  return HEALTHCHECK_DIM_KEYS_.every(k => {
    const rv = String(ruleDims[k] == null ? '' : ruleDims[k]).trim().toUpperCase();
    const iv = String(itemDims[k] == null ? '' : itemDims[k]).trim().toUpperCase();
    if (rv === 'NO' || rv === '') return true;
    if (rv.indexOf(',') > -1) return rv.split(',').map(x => x.trim()).indexOf(iv) > -1;
    if (iv === rv) return true;
    const rn = Number(rv), inum = Number(iv);
    return Number.isFinite(rn) && Number.isFinite(inum) && Math.round(rn) === Math.round(inum);
  });
}
function runDataHealthCheck_() {
  const programs = readSheetRows_(SHEETS.PROGRAMS);
  const doanhthu = readSheetRows_(SHEETS.DOANHTHU);
  const chuongtrinh = readSheetRows_(SHEETS.CHUONGTRINH);
  const mhck = readSheetRows_(SHEETS.MHCK);
  const purchases = readSheetRows_(SHEETS.PURCHASES);

  const issues = [];
  function addIssue(muc, mucdo, doituong, chitiet) {
    issues.push({ muc: muc, mucdo: mucdo, doituong: doituong, chitiet: chitiet });
  }

  const doanhthuById = {}; doanhthu.forEach(d => { if (d.id) doanhthuById[d.id] = d; });
  const chuongtrinhById = {}; chuongtrinh.forEach(c => { if (c.id) chuongtrinhById[c.id] = c; });

  // (a) Trùng Mã chiết khấu / (b) liên kết Chương trình cha bị đứt / (c) dim_nhomkh lệch số
  const seenMaChietKhau = {};
  programs.forEach(p => {
    if (p.ma_chietkhau) {
      if (seenMaChietKhau[p.ma_chietkhau]) {
        addIssue('Trùng Mã chiết khấu', 'error', p.ma_chietkhau,
          'Có nhiều hơn 1 dòng (id="' + seenMaChietKhau[p.ma_chietkhau] + '" và id="' + p.id + '") cùng dùng Mã chiết khấu "' + p.ma_chietkhau + '" — dễ nhầm lẫn khi chọn ở màn Tính chiết khấu.');
      } else {
        seenMaChietKhau[p.ma_chietkhau] = p.id;
      }
    }
    if (p.chuongtrinh_id && !chuongtrinhById[p.chuongtrinh_id]) {
      addIssue('Liên kết Chương trình bị đứt', 'error', p.ma_chietkhau || p.id,
        'chuongtrinh_id="' + p.chuongtrinh_id + '" không khớp Chương trình nào trong Danh mục chương trình — hiệu lực Từ ngày/Đến ngày có thể không tự cập nhật đúng theo chương trình cha.');
    }
    const dt = p.ma_doanhthu ? doanhthuById[p.ma_doanhthu] : null;
    if (dt) {
      HEALTHCHECK_DIM_KEYS_.forEach(k => {
        const pv = p['dim_' + k], dv = dt['dim_' + k];
        if (pv == null || dv == null || String(pv) === String(dv)) return; // trống hoặc khớp hệt — không có vấn đề
        const pn = Number(pv), dn = Number(dv);
        const numericClose = Number.isFinite(pn) && Number.isFinite(dn) && Math.round(pn) === Math.round(dn);
        addIssue(numericClose ? 'Dim lệch độ chính xác số (đã tự khắc phục lúc chạy)' : 'Dim KHÔNG khớp Mã doanh thu liên kết',
          numericClose ? 'warn' : 'error', p.ma_chietkhau || p.id,
          'dim_' + k + ' = "' + pv + '" trên Mã chiết khấu ' + (numericClose ? 'khác' : 'KHÁC HẲN') + ' "' + dv + '" trên Mã doanh thu "' + p.ma_doanhthu + '"' +
          (numericClose ? ' (chỉ lệch phần thập phân, nghi do cell từng bị đổi định dạng ngày/giờ trên Sheet — app đã tự nhận đúng nhóm nên KHÔNG ảnh hưởng số tiền, nhưng nên sửa lại cho sạch dữ liệu: mở Mã chiết khấu, chọn lại đúng Mã doanh thu rồi lưu).'
            : ' — có thể khiến mã này KHÔNG khớp được mã hàng nào (0đ chiết khấu mà không cảnh báo gì).'));
      });
      // dim_nhomkh vốn LUÔN là mã số nguyên (VD 231013) — có phần thập phân là dấu hiệu rõ ràng của lỗi
      // định dạng ô, kiểm tra ĐỘC LẬP (không cần đối chiếu với Mã doanh thu) để bắt được cả trường hợp
      // dim_nhomkh của CHÍNH Mã doanh thu cũng bị lỗi (khi đó phép so ở trên vẫn "khớp" vì cả 2 cùng sai).
      [['dim_nhomkh (Mã chiết khấu)', p.dim_nhomkh], ['dim_nhomkh (Mã doanh thu ' + p.ma_doanhthu + ')', dt.dim_nhomkh]].forEach(pair => {
        const v = Number(pair[1]);
        if (Number.isFinite(v) && Math.abs(v - Math.round(v)) > 1e-6) {
          addIssue('Mã Nhóm hàng có phần thập phân bất thường', 'warn', p.ma_chietkhau || p.id,
            pair[0] + ' = ' + pair[1] + ' — mã Nhóm hàng luôn phải là số nguyên, phần thập phân dư (' + (v - Math.round(v)).toFixed(6) + ') nghi do lỗi định dạng ô trên Sheet.');
        }
      });
    }
    // (d) SL/DT Min-Max/Tỷ lệ CK không parse được thành số hợp lệ (VD lỗi "#VALUE!"/"#REF!" từ Excel) —
    // app đã tự coi các giá trị này là "không giới hạn" (xem inBracket()) nên KHÔNG chặn tính toán, chỉ
    // là dấu hiệu dữ liệu bẩn nên cảnh báo để người dùng xác nhận lại đúng ý (0 thật hay không giới hạn).
    ['sl_min', 'sl_max', 'dt_min', 'dt_max', 'tl_ck'].forEach(f => {
      const v = p[f];
      if (v === '' || v == null) return;
      if (!Number.isFinite(Number(v))) {
        addIssue('Giá trị số bị lưu sai định dạng', 'warn', p.ma_chietkhau || p.id,
          'Trường "' + f + '" = "' + v + '" không phải số hợp lệ — app tạm coi như "không giới hạn"/0, nên vào Mã chiết khấu sửa lại đúng giá trị mong muốn.');
      }
    });
  });

  // (e) Mã chiết khấu ĐANG HIỆU LỰC nhưng KHÔNG khớp mã hàng nào phát sinh mua trong N ngày gần nhất —
  // dấu hiệu cấu hình dims sai/lệch (đúng lớp lỗi đã gặp thực tế), hoặc đơn giản mặt hàng đã ngừng bán.
  const todayMs = Date.now();
  const staleMs = MAINTENANCE_STALE_DAYS_ * 24 * 3600 * 1000;
  const mhckByCode = {}; mhck.forEach(m => { if (m.mamh) mhckByCode[String(m.mamh).trim().toUpperCase()] = m; });
  const recentDimsSeen = [];
  purchases.forEach(p => {
    if (!p.ngay || !p.mahang) return;
    const d = new Date(p.ngay);
    if (isNaN(d.getTime()) || (todayMs - d.getTime()) > staleMs || d.getTime() > todayMs) return;
    const rec = mhckByCode[String(p.mahang).trim().toUpperCase()];
    if (!rec) return;
    recentDimsSeen.push({ thuonghieu: rec.thuonghieu, nhomkh: rec.nhomkh, loaisp: rec.loaisp, dactinh: rec.dactinh, congdung: rec.congdung, quycach: rec.quycach });
  });
  programs.forEach(p => {
    const hieuluctu = p.hieuluctu ? new Date(p.hieuluctu) : null;
    const hieulucden = p.hieulucden ? new Date(p.hieulucden) : null;
    const active = (!hieulucden || isNaN(hieulucden.getTime()) || hieulucden.getTime() >= todayMs) &&
                   (!hieuluctu || isNaN(hieuluctu.getTime()) || hieuluctu.getTime() <= todayMs);
    if (!active) return;
    const pd = { thuonghieu: p.dim_thuonghieu, nhomkh: p.dim_nhomkh, loaisp: p.dim_loaisp, dactinh: p.dim_dactinh, congdung: p.dim_congdung, quycach: p.dim_quycach };
    const matched = recentDimsSeen.some(id => healthCheckDimsMatchLoose_(id, pd));
    if (!matched) {
      addIssue('Mã chiết khấu không khớp mã hàng nào gần đây', 'warn', p.ma_chietkhau || p.id,
        'Đang trong hiệu lực nhưng không có mã hàng nào khớp phạm vi (Thương hiệu/Nhóm hàng/Loại SP/Đặc tính/Công dụng/Quy cách) phát sinh mua trong ' + MAINTENANCE_STALE_DAYS_ + ' ngày gần nhất — kiểm tra lại cấu hình dims, hoặc xác nhận mặt hàng đã ngừng mua.');
    }
  });

  // (f) Cột "Giá công bố" (DM_GIACONGBO/gia_congbo) bị Google Sheets TỰ ĐỘNG định dạng nhầm thành Ngày
  // tháng — đúng lớp lỗi định dạng ô đã gặp thực tế với dim_nhomkh (xem sanitizeCellValue_/
  // serialFromDate_ ở đầu file), nhưng ở đây readSheetRows_()/sanitizeCellValue_ đã TỰ KHÔI PHỤC đúng số
  // tiền khi đọc nên KHÔNG có dấu hiệu gì bất thường trong dữ liệu ĐÃ ĐỌC (không như dim_nhomkh, phần
  // thập phân lẻ lộ ra sau khôi phục) — phải đọc TRỰC TIẾP ô gốc trên Sheet (chưa sanitize) để phát hiện.
  // Chỉ là cảnh báo "vệ sinh dữ liệu": không ảnh hưởng số tiền chiết khấu tính ra, nhưng định dạng kiểu
  // này dễ vỡ khi mở/xuất bằng Excel thật (Excel có giới hạn năm hợp lệ, sẽ báo lỗi hoặc mất dữ liệu với
  // các ô có serial ngày quá lớn như thế này).
  {
    const shGcb = getOrCreateSheet_(SpreadsheetApp.getActive(), SHEETS.GIACONGBO);
    const lastRowGcb = shGcb.getLastRow(), lastColGcb = shGcb.getLastColumn();
    if (lastRowGcb >= 2 && lastColGcb >= 1) {
      const headersGcb = shGcb.getRange(1, 1, 1, lastColGcb).getValues()[0];
      const iGia = headersGcb.indexOf('gia_congbo');
      const iMaNcc = headersGcb.indexOf('mamh_ncc');
      if (iGia > -1) {
        const rawGcb = shGcb.getRange(2, 1, lastRowGcb - 1, lastColGcb).getValues();
        const sample = [];
        let dateFormatCount = 0;
        rawGcb.forEach((r, idx) => {
          if (r[iGia] instanceof Date) {
            dateFormatCount++;
            if (sample.length < 5) sample.push((iMaNcc > -1 && r[iMaNcc] ? r[iMaNcc] : ('dòng ' + (idx + 2))));
          }
        });
        if (dateFormatCount > 0) {
          addIssue('Ô Giá công bố đang bị định dạng nhầm thành Ngày tháng', 'warn', 'DM_GIACONGBO',
            dateFormatCount + ' dòng có cột "gia_congbo" đang bị Sheet tự định dạng thành Ngày tháng (VD: ' + sample.join(', ') + (dateFormatCount > sample.length ? ', ...' : '') +
            ') — app đã tự khôi phục đúng số tiền khi đọc (KHÔNG ảnh hưởng số tiền chiết khấu tính ra), nhưng nên bôi đen các ô này trên Sheet rồi Format > Số để làm sạch định dạng, tránh lỗi khi mở/xuất bằng Excel thật.');
        }
      }
    }
  }

  const summary = {
    total: issues.length,
    error: issues.filter(x => x.mucdo === 'error').length,
    warn: issues.filter(x => x.mucdo === 'warn').length,
    checkedAt: new Date().toISOString()
  };
  return { summary: summary, issues: issues };
}
// Gọi được trực tiếp từ nút "Quét sức khỏe dữ liệu" trên web.
function getDataHealthReport() {
  return runDataHealthCheck_();
}

// ---- (3) Trigger hằng ngày + email báo cáo ----
function dailyMaintenanceJob_() {
  const backedUp = backupKeyConfigSheets_();
  const report = runDataHealthCheck_();
  if (report.summary.total > 0) sendMaintenanceEmail_(report, backedUp);
  return report;
}
function sendMaintenanceEmail_(report, backedUp) {
  let email = '';
  try { email = Session.getEffectiveUser().getEmail(); } catch (e) { email = ''; }
  if (!email) return;
  const lines = [];
  lines.push('Kiểm tra sức khỏe dữ liệu tự động — ' + new Date().toLocaleString('vi-VN'));
  lines.push('Tổng: ' + report.summary.total + ' vấn đề (' + report.summary.error + ' lỗi, ' + report.summary.warn + ' cảnh báo)');
  lines.push('');
  report.issues.slice(0, 50).forEach(x => {
    lines.push('[' + (x.mucdo === 'error' ? 'LỖI' : 'CẢNH BÁO') + '] ' + x.muc + ' — ' + x.doituong);
    lines.push('  ' + x.chitiet);
  });
  if (report.issues.length > 50) lines.push('... và ' + (report.issues.length - 50) + ' vấn đề khác — xem đầy đủ trên web, mục Bảo trì hệ thống.');
  lines.push('');
  lines.push('Đã sao lưu ' + backedUp.length + ' sheet: ' + (backedUp.join(', ') || '(không có sheet nào cần sao lưu)'));
  MailApp.sendEmail(email, '[Chiết khấu NCC] Báo cáo bảo trì tự động — ' + report.summary.total + ' vấn đề cần xem', lines.join('\n'));
}
// Cài đặt/gỡ trigger chạy hằng ngày lúc 2h sáng — gọi từ nút bấm trên web (yêu cầu cấp quyền 1 lần).
function installDailyMaintenanceTrigger() {
  uninstallDailyMaintenanceTrigger(); // tránh tạo trùng nếu bấm nhiều lần
  ScriptApp.newTrigger('dailyMaintenanceJob_').timeBased().everyDays(1).atHour(2).create();
  return true;
}
function uninstallDailyMaintenanceTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyMaintenanceJob_') ScriptApp.deleteTrigger(t);
  });
  return true;
}
function isDailyMaintenanceTriggerInstalled() {
  return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'dailyMaintenanceJob_');
}

const HD_TOC_QUYTRINH_ = [
  { lbl: 'I. Dữ liệu nền', items: [
    ['hd-b1', '1. Danh mục nhóm chiết khấu'], ['hd-b2', '2. Giá công bố NCC'], ['hd-b3', '3. Sổ chi tiết mua hàng']
  ]},
  { lbl: 'II. Doanh số & kế hoạch', items: [
    ['hd-b4', '4. Danh mục doanh số'], ['hd-b5', '5. Danh mục kế hoạch']
  ]},
  { lbl: 'III. Cấu hình chiết khấu', items: [
    ['hd-b6', '6. Danh mục chương trình'], ['hd-b7', '7. Mã chiết khấu & bậc'], ['hd-b8', '8. Chiết khấu lũy kế']
  ]},
  { lbl: 'IV. Vận hành hằng kỳ', items: [
    ['hd-b9', '9. Tính chiết khấu'], ['hd-b9b', '9b. Chiết khấu tạm tính theo tháng'], ['hd-b10', '10. Báo cáo & đối chiếu']
  ]},
  { lbl: '', items: [ ['hd-faq', 'Mẹo & câu hỏi thường gặp'] ] },
];
const HD_TOC_NGHIEPVU_ = [
  { lbl: 'Chi tiết từng loại chiết khấu', items: [
    ['hd-b11', '11. Mục 6 & Mục 7 — ví dụ cụ thể'],
    ['hd-b12', '12. CK 4 tháng liên tục — ví dụ cụ thể'],
    ['hd-b13', '13. Sơ đồ quy trình từng loại chiết khấu'],
    ['hd-b14', '14. Mục 7 CK năm — hướng dẫn từng bước đầy đủ (2 khoảng ngày)']
  ]},
  { lbl: '', items: [ ['hd-faq', 'Mẹo & câu hỏi thường gặp'] ] },
];

const HD_CONTENT_QUYTRINH_ = `
    <div class="section-note">Ứng dụng không "hard-code" chính sách Casumina — mọi chương trình chiết khấu là <b>dữ liệu khai báo</b>. Trang này đi theo đúng thứ tự thao tác thực tế: khai dữ liệu nền trước, cấu hình chương trình sau, rồi lặp lại 2 bước cuối mỗi kỳ.</div>

    <div class="card">
      <div class="card-head"><div><h2>Quy trình tổng quan</h2><div class="hint">10 bước, đi từ trái sang phải</div></div></div>
      <div class="hd-flow">
        <span class="tag">1. Danh mục nhóm CK</span><span class="arrow">→</span>
        <span class="tag">2. Giá công bố</span><span class="arrow">→</span>
        <span class="tag">3. Sổ mua hàng</span><span class="arrow">→</span>
        <span class="tag tag-gold">4. Mã doanh thu</span><span class="arrow">→</span>
        <span class="tag">5. Kế hoạch</span><span class="arrow">→</span>
        <span class="tag">6. Chương trình</span><span class="arrow">→</span>
        <span class="tag">7. Mã chiết khấu &amp; bậc</span><span class="arrow">→</span>
        <span class="tag">8. Lũy kế</span><span class="arrow">→</span>
        <span class="tag tag-ok">9. Tính chiết khấu</span><span class="arrow">→</span>
        <span class="tag tag-ok">10. Báo cáo</span>
      </div>
    </div>

    <div class="card">
      <div class="card-head"><div><h2>Mô hình 3 tầng cấu hình</h2><div class="hint">Hiểu đúng mô hình này quan trọng hơn nhớ từng nút bấm</div></div></div>
      <div class="hd-tier">
        <div class="side t1">TẦNG 1<br>Chương trình</div>
        <div class="body"><b>Danh mục chương trình</b> — 1 khung thời gian dùng chung (VD "CSMN_2026_Q1"). Không tự tính tiền, chỉ nhóm các Mã chiết khấu con và cấp hiệu lực chung cho chúng. <span class="hint">Menu: Chiết khấu → Danh mục chương trình</span></div>
      </div>
      <div class="hd-tier">
        <div class="side t2">TẦNG 2<br>Mã chiết khấu</div>
        <div class="body"><b>Chương trình chiết khấu</b> — đơn vị tính tiền thực sự: chọn phạm vi mặt hàng (qua Mã doanh thu), bậc Sản lượng/Doanh thu, hình thức tính (% giá mua / % giá công bố / số tiền cố định), tỷ lệ. <span class="hint">Menu: Chiết khấu → Mã chiết khấu</span></div>
      </div>
      <div class="hd-tier">
        <div class="side t3">TẦNG 3<br>Bậc điều kiện</div>
        <div class="body"><b>Bậc con (tùy chọn)</b> — khi 1 Mã chiết khấu có NHIỀU mức % theo từng khoảng SL/DT/%KH khác nhau. Không có bậc con thì dùng thẳng số liệu trên chính Mã chiết khấu. <span class="hint">Trong Mã chiết khấu, bấm "Quản lý bậc"</span></div>
      </div>
      <div class="section-note" style="border-color:var(--gold);color:#8a6a1a;background:var(--gold-soft);margin-top:14px">
        <b>Mã doanh thu</b> là mảnh ghép thứ 4, nằm song song với 3 tầng trên: nơi khai "chiết khấu này áp dụng cho nhóm mặt hàng nào" — và từ 24/08/2026, cả "doanh số dùng để so bậc/%KH lấy theo giá công bố hay giá hóa đơn". Mỗi Mã chiết khấu đều phải chọn 1 Mã doanh thu (xem Bước 4).
      </div>
    </div>

    <div class="card hd-step" id="hd-b1">
      <div class="card-head"><div><h2>Bước 1 — Danh mục nhóm chiết khấu</h2><div class="hint">Menu: Danh mục → Danh mục nhóm chiết khấu</div></div></div>
      <p style="font-size:13px;color:var(--slate-600)">Mọi chương trình chiết khấu lọc mặt hàng theo <b>6 kích thước</b>: Thương hiệu, Nhóm hàng, Loại sản phẩm, Đặc tính, Công dụng, Quy cách (có thêm khoảng Từ–Đến). Khai danh mục từng kích thước trước, sau đó gắn từng <b>Mã hàng</b> vào đúng tổ hợp ở tab con "Mã hàng → Nhóm CK". Chọn <code>NO</code> ở kích thước nào nghĩa là "áp dụng cho tất cả".</p>
      <div class="section-note">Cột <b>Mã MH</b> khi gắn mã hàng là mã theo bảng giá công bố NCC (khác "Mã hàng_misa" dùng trong Sổ mua hàng) — khai đúng cột này để tra được giá công bố ở Bước 2.</div>
      <div class="section-note" style="border-color:var(--emerald);color:var(--emerald-dark)">Mã hàng mới xuất hiện trong Sổ mua hàng nhưng chưa gắn nhóm sẽ được app <b>tự động thêm</b> (6 kích thước = NO) và đánh dấu "cần rà soát" — xem banner nhắc ở trang Tổng quan.</div>
      <p style="font-size:13px;color:var(--slate-600)">Có thể tải lên/xuất Excel (<code>DM_NHOM_CHIETKHAU.xlsx</code>, <code>DM_MHCK</code>) để nhập hàng loạt.</p>
    </div>

    <div class="card hd-step" id="hd-b2">
      <div class="card-head"><div><h2>Bước 2 — Giá công bố NCC</h2><div class="hint">Menu: Quản lý dữ liệu → Giá công bố NCC</div></div></div>
      <p style="font-size:13px;color:var(--slate-600)">Khai giá niêm yết Casumina theo từng <b>Mã MH</b>, có thể khai nhiều mức theo từng giai đoạn hiệu lực. Chuỗi tra cứu: <code>Mã hàng_misa (Sổ mua hàng) → Mã MH (Danh mục nhóm CK) → Giá công bố</code>. Dữ liệu này phục vụ hình thức "% theo giá công bố" và tính năng "doanh số theo giá công bố" ở Bước 4.</p>
      <div class="section-note" style="border-color:var(--emerald);color:var(--emerald-dark)">Màn hình này <b>đọc thẳng Google Sheet theo trang</b> (20 dòng/trang, không tự tải/cache cả bảng) — mặc định chỉ hiện các bản ghi <b>đang hiệu lực</b> hôm nay; bỏ tick ô "Chỉ đang hiệu lực" để xem cả bản ghi cũ/tương lai. Gõ ô tìm kiếm hoặc bấm Trang trước/sau đều tự động truy vấn lại — không cần bấm gì thêm.</div>
      <h4 style="margin:16px 0 6px;font-size:13.5px">Chống trùng dữ liệu</h4>
      <p style="font-size:13px;color:var(--slate-600)">Bộ khóa xác định "1 bản ghi" là <b>Mã MH + Hiệu lực từ ngày + Hiệu lực đến ngày</b>. Bấm <b>"Thêm giá công bố"</b> mà trùng khóa với 1 dòng đã có → hệ thống tự <b>gộp vào đúng dòng đó</b> (không tạo dòng trùng). Bấm <b>"Sửa"</b> mà nội dung sửa khiến trùng khóa với 1 dòng KHÁC → bị <b>chặn lại</b>, báo rõ dòng nào đang trùng để tự xử lý.</p>
      <h4 style="margin:16px 0 6px;font-size:13.5px">Tải file Excel — qua "Bản nháp" trước khi ghi</h4>
      <p style="font-size:13px;color:var(--slate-600)">Bấm <b>"Tải file Giá công bố (.xlsx)"</b> KHÔNG ghi ngay lên Sheet — hệ thống so khớp từng dòng đọc được với dữ liệu hiện có rồi mở màn hình <b>Bản nháp</b>, gắn trạng thái từng dòng:</p>
      <div class="table-wrap"><table>
        <thead><tr><th>Trạng thái</th><th>Ý nghĩa</th></tr></thead>
        <tbody>
          <tr><td><span class="tag tag-ok">Mới</span></td><td>Chưa có bản ghi nào trùng Mã MH + khoảng hiệu lực — sẽ thêm mới nếu chọn ghi</td></tr>
          <tr><td><span class="tag tag-gold">Cập nhật</span></td><td>Đã có bản ghi trùng khóa nhưng Giá công bố hoặc Tên hàng khác (hiện "giá cũ → giá mới") — sẽ ghi đè nếu chọn ghi</td></tr>
          <tr><td><span class="tag">Không đổi</span></td><td>Đã có bản ghi trùng khóa và giống hệt — mặc định không tick, không cần ghi lại</td></tr>
          <tr><td><span class="tag tag-no">Lỗi</span></td><td>Thiếu Mã MH, Giá công bố không hợp lệ, hoặc Hiệu lực đến sớm hơn Hiệu lực từ — không tick chọn được</td></tr>
        </tbody>
      </table></div>
      <p style="font-size:13px;color:var(--slate-600)">Tick chọn đúng dòng muốn ghi (có nút chọn nhanh "Mới + Cập nhật" / "Chọn tất cả" / "Bỏ chọn tất cả"), bấm <b>"Ghi vào bảng chính"</b> — chỉ đúng các dòng đã chọn mới thực sự ghi lên Sheet. Bấm "Hủy" thì không đổi gì.</p>
      <p style="font-size:13px;color:var(--slate-600)"><b>"Xuất Excel"</b> tải toàn bộ dữ liệu (không giới hạn trang) để xuất file — chỉ tải khi bấm nút này, không tự động khi mở tab.</p>
    </div>

    <div class="card hd-step" id="hd-b3">
      <div class="card-head"><div><h2>Bước 3 — Sổ chi tiết mua hàng</h2><div class="hint">Menu: Quản lý dữ liệu → Sổ chi tiết mua hàng</div></div></div>
      <p style="font-size:13px;color:var(--slate-600)">Tải lên Excel hóa đơn mua hàng thực tế — nguồn số liệu gốc cho mọi phép tính chiết khấu. Bộ lọc hỗ trợ theo khoảng ngày, nhà cung cấp, từ khóa và cả 6 kích thước phân loại.</p>
      <div class="section-note" style="border-color:var(--emerald);color:var(--emerald-dark)">Màn hình này <b>không tự tải gì khi mở tab</b>. Chỉnh bộ lọc rồi bấm <b>"Lọc"</b> (hoặc Enter) mới đọc thẳng Google Sheet — trả về 20 dòng khớp/trang (có Trang trước/sau) kèm tổng SL/giá trị/số NCC tính trên <b>toàn bộ</b> kết quả khớp, không chỉ trang đang xem.</div>
      <h4 style="margin:16px 0 6px;font-size:13.5px">Tải file Excel — qua "Bản nháp" trước khi ghi</h4>
      <p style="font-size:13px;color:var(--slate-600)">Bấm <b>"Tải file Sổ chi tiết mua hàng (.xlsx)"</b> cũng KHÔNG ghi ngay — mở màn hình <b>Bản nháp</b> gắn trạng thái từng dòng theo khóa <b>Ngày + Số HĐ + Mã hàng + Số lượng + Giá trị</b>:</p>
      <div class="table-wrap"><table>
        <thead><tr><th>Trạng thái</th><th>Ý nghĩa</th></tr></thead>
        <tbody>
          <tr><td><span class="tag tag-ok">Mới</span></td><td>Chưa có trên Sheet — mặc định tick, sẽ thêm mới nếu chọn ghi</td></tr>
          <tr><td><span class="tag">Trùng lặp</span></td><td>Đã có sẵn trên Sheet hoặc trùng với 1 dòng khác trong chính file đang tải — không tick chọn được (Sổ mua hàng chỉ thêm mới, không có khái niệm "cập nhật đè")</td></tr>
          <tr><td><span class="tag tag-no">Lỗi</span></td><td>Thiếu Mã hàng hoặc Ngày chứng từ không đọc được — không tick chọn được</td></tr>
        </tbody>
      </table></div>
      <p style="font-size:13px;color:var(--slate-600)">Tick chọn dòng muốn ghi rồi bấm <b>"Ghi vào bảng chính"</b>. Mã hàng mới xuất hiện trong các dòng <b>đã thực sự ghi</b> nhưng chưa có trong Danh mục Mã hàng→Nhóm CK sẽ được tự động thêm (6 kích thước = NO, đánh dấu "cần rà soát").</p>
    </div>

    <div class="card hd-step" id="hd-b4">
      <div class="card-head"><div><h2>Bước 4 — Danh mục doanh số <span class="tag tag-gold">mới 24/08/2026</span></h2><div class="hint">Menu: Danh mục → Danh mục doanh số</div></div></div>
      <p style="font-size:13px;color:var(--slate-600)">Mỗi <b>"Mã doanh thu"</b> = 6 kích thước phân loại + Thời điểm chiết khấu (Tháng/Quý/6 Tháng/Năm) + Cách tính doanh số. Khai 1 lần, dùng lại cho nhiều Mã chiết khấu — cả làm điều kiện xét bậc lẫn làm doanh thu tính tiền.</p>
      <div class="table-wrap"><table>
        <thead><tr><th>Trường</th><th>Ý nghĩa</th></tr></thead>
        <tbody>
          <tr><td class="mono">Mã / Tên gợi nhớ</td><td>Định danh ngắn để chọn lại ở Mã chiết khấu</td></tr>
          <tr><td class="mono">6 kích thước</td><td>Phạm vi mặt hàng — mã hàng nào được gộp khi tính SL/DT. NO = không giới hạn</td></tr>
          <tr><td class="mono">Thời điểm CK</td><td>Tháng/Quý/6 Tháng/Năm — đối chiếu đúng kỳ kế hoạch</td></tr>
          <tr><td class="mono">Cách tính doanh số</td><td>Giá hóa đơn (mặc định) hoặc Giá công bố</td></tr>
        </tbody>
      </table></div>
      <h4 style="margin:16px 0 6px;font-size:13.5px">Ô tick "Cách tính doanh số"</h4>
      <p style="font-size:13px;color:var(--slate-600)">Quyết định app dùng <b>giá mua thực tế</b> hay <b>giá công bố</b> làm "doanh số" khi: (1) so bậc Sản lượng/Doanh thu, và (2) tính %KH để so %KH Min/Max. Mặc định là <b>Giá hóa đơn</b> cho mọi Mã doanh thu mới — chỉ đổi sang Giá công bố khi văn bản chính sách ghi rõ (VD Mục 3 TB 0101 — CK tháng).</p>
      <div class="section-note" style="border-color:var(--emerald);color:var(--emerald-dark)">Nếu tick "Giá công bố" nhưng thiếu giá công bố cho 1 số mã hàng, hệ thống tự dùng tạm giá hóa đơn cho phần thiếu VÀ gắn cảnh báo ở báo cáo — không bao giờ báo sai mà không cảnh báo.</div>
      <div class="section-note" style="border-color:var(--gold);color:#8a6a1a;background:var(--gold-soft)">Muốn thêm 1 chương trình cho Thương hiệu/nhóm hàng mới (VD PCR)? Không cần sửa code — chỉ tạo 1 Mã doanh thu mới chọn đúng Thương hiệu ở đây, rồi qua Bước 7 tạo Mã chiết khấu tham chiếu đúng Mã doanh thu đó.</div>
    </div>

    <div class="card hd-step" id="hd-b5">
      <div class="card-head"><div><h2>Bước 5 — Danh mục kế hoạch</h2><div class="hint">Menu: Danh mục → Danh mục kế hoạch</div></div></div>
      <p style="font-size:13px;color:var(--slate-600)">Khai chỉ tiêu Doanh thu/Sản lượng kế hoạch — mẫu số để tính %KH ở mọi nơi. Mỗi kế hoạch có khung hiệu lực tổng thể, áp dụng cho "Tất cả (ALL)" hoặc chỉ những Mã doanh thu được chọn.</p>
      <p style="font-size:13px;color:var(--slate-600)">Nếu chỉ tiêu khác nhau theo từng kỳ con, bấm biểu tượng "Chi tiết kế hoạch" ở mỗi dòng để khai riêng theo Tháng/Quý/6 Tháng/Năm — khi có Chi tiết kế hoạch, app luôn ưu tiên dùng số của kỳ con thay vì số tổng.</p>
    </div>

    <div class="card hd-step" id="hd-b6">
      <div class="card-head"><div><h2>Bước 6 — Danh mục chương trình (Tầng 1)</h2><div class="hint">Menu: Chiết khấu → Danh mục chương trình</div></div></div>
      <p style="font-size:13px;color:var(--slate-600)">Tạo 1 dòng cho mỗi đợt chính sách (VD <code>CSMN_2026_Q1</code>): Mã chương trình, Nội dung, khung Từ ngày–Đến ngày. Khung ngày tự động áp xuống toàn bộ Mã chiết khấu con liên kết — sửa ở đây, mọi mã con cập nhật theo. Dùng nút <b>Sao chép</b> khi mở đợt chính sách mới cho kỳ sau.</p>
    </div>

    <div class="card hd-step" id="hd-b7">
      <div class="card-head"><div><h2>Bước 7 — Mã chiết khấu & bậc điều kiện (Tầng 2 & 3)</h2><div class="hint">Menu: Chiết khấu → Mã chiết khấu</div></div></div>
      <ol style="font-size:13px;color:var(--slate-600);padding-left:18px;margin:0 0 10px">
        <li style="margin-bottom:6px">Liên kết vào <b>Chương trình (Tầng 1)</b> — hiệu lực tự khóa theo chương trình cha.</li>
        <li style="margin-bottom:6px">Chọn <b>Mã doanh thu (điều kiện)</b> đã tạo ở Bước 4 — quyết định phạm vi mặt hàng và cách tính doanh số. Có thể chọn thêm <b>Mã doanh thu tính chiết khấu</b> riêng nếu muốn xét đạt bậc trên phạm vi rộng nhưng chỉ trả tiền trên 1 nhóm hẹp hơn.</li>
        <li style="margin-bottom:6px"><b>Hình thức chiết khấu</b>: % theo giá mua, % theo giá công bố, số tiền cố định/đơn vị, hoặc "Lũy kế" (làm chương trình gốc cho Bước 8).</li>
        <li>Bậc <b>Sản lượng/Doanh thu</b> (Min–Max, AND/OR) và % Kế hoạch Min–Max nếu chỉ áp dụng trong 1 khoảng %KH.</li>
      </ol>
      <div class="section-note" style="border-color:var(--red);color:var(--red)"><b>Cờ "Không cộng vào Tổng CKTM"</b> — bật cho khoản đã nằm sẵn trong giá hóa đơn (VD "CK cơ bản 10%"). Vẫn tính/hiển thị riêng đầy đủ, nhưng loại khỏi mọi số "Tổng" ở báo cáo và Excel xuất.</div>
      <h4 style="margin:16px 0 6px;font-size:13.5px">Nhiều mức % trong 1 mã — dùng Bậc điều kiện (Tầng 3)</h4>
      <p style="font-size:13px;color:var(--slate-600)">Bấm <b>"Quản lý bậc"</b> để khai từng mức thành 1 bậc riêng (SL/DT Min–Max, %KH Min–Max, Tỷ lệ CK). Hệ thống tự so khớp %KH/SL/DT thực tế với đúng bậc phù hợp. Không có bậc con thì dùng thẳng số liệu trên chính Mã chiết khấu — tương thích ngược hoàn toàn.</p>
    </div>

    <div class="card hd-step" id="hd-b8">
      <div class="card-head"><div><h2>Bước 8 — Chiết khấu lũy kế</h2><div class="hint">Chiết khấu → Mã chiết khấu → tab con "Chiết khấu lũy kế"</div></div></div>
      <p style="font-size:13px;color:var(--slate-600)">Các khoản CK bổ sung phụ thuộc đạt liên tục nhiều tháng/kỳ. Mỗi dòng gắn 1 <b>Mã chiết khấu chính</b> và chọn 1 trong 3 kiểu tính:</p>
      <div class="grid grid-3" style="margin-bottom:10px">
        <div class="hd-kv"><div class="k">1. Đạt tất cả các tháng</div><div class="v">Đủ tất cả tháng mới hưởng ĐẦY ĐỦ; lỡ 1 tháng vẫn có thể hưởng mức % giảm trừ nếu cấu hình</div></div>
        <div class="hd-kv"><div class="k">2. 2 điều kiện độc lập</div><div class="v">ĐK1 (mỗi tháng ≥X%) và ĐK2 (cả kỳ ≥Y%) — đạt 1 hoặc cả 2 thì cộng % tương ứng</div></div>
        <div class="hd-kv"><div class="k">3. Bù CK tháng theo Quý <span class="tag tag-gold" style="font-size:10px">mới</span></div><div class="v">Tháng lẻ không đạt vẫn có thể được bù, nếu Quý chứa tháng đó đạt kế hoạch — xem sơ đồ bên dưới</div></div>
      </div>
      <p style="font-size:13px;color:var(--slate-600)">Kiểu 1 và 2 đọc lại báo cáo THÁNG <b>đã lưu</b> (REPORT_CKTH) — cần tính và lưu đủ các tháng liên quan (Bước 9) trước, hoặc dùng nút <b>"Tự động tính & lưu cho các tháng đã chọn"</b> để làm 1 lần cho tất cả.</p>
      <div class="section-note" style="border-color:var(--gold);color:#8a6a1a;background:var(--gold-soft)">Cần ngưỡng "đạt tháng" khác mã CK tháng đang trả tiền thật? Tạo 1 "mã cổng" (Tỷ lệ CK = 0, %KH Min = ngưỡng cần) rồi liên kết dòng lũy kế vào mã cổng đó thay vì mã đang trả tiền.</div>

      <div class="section-note" style="margin-top:14px;border-color:var(--gold);color:#8a6a1a;background:var(--gold-soft)"><b>Kiểu 3 — Bù CK tháng theo Quý</b> hoạt động thế nào (theo QĐ hỗ trợ Casumina cho NPP/ĐL lốp ô tô):</div>
      <div class="flow-diag" style="margin:8px 0 10px">
        <div class="flow-box">Bấm "Tính chiết khấu bổ sung" → từng tháng đã chọn được tính LẠI ngay từ Sổ chi tiết mua hàng (không cần đã lưu báo cáo tháng trước, không "quay lui" sửa số cũ)</div>
        <div class="flow-down">↓ Tháng đó có TỰ đạt kế hoạch riêng của nó không?</div>
        <div class="flow-branch">
          <div class="col"><div class="lbl">Tháng TỰ đạt</div><div class="flow-box emerald">Không cần bù — CK tháng vẫn tính như bình thường</div></div>
          <div class="col"><div class="lbl">Tháng KHÔNG đạt, Quý chứa tháng đó ĐẠT ≥100% KH</div><div class="flow-box gold">ĐƯỢC BÙ = % của bậc QUÝ đã đạt × Doanh số THẬT của chính THÁNG đó</div></div>
          <div class="col"><div class="lbl">Tháng KHÔNG đạt, Quý cũng CHƯA đạt 100% KH</div><div class="flow-box red">0đ — không đủ điều kiện bù</div></div>
        </div>
      </div>
      <p style="font-size:13px;color:var(--slate-600)">Cấu hình (chỉ 2 bước):</p>
      <ol style="font-size:13px;color:var(--slate-600);padding-left:18px;margin:0 0 10px">
        <li style="margin-bottom:6px">Chọn <b>"Mã chiết khấu Quý liên kết"</b> — 1 Mã chiết khấu khác, cùng phạm vi mã hàng, Thời điểm chiết khấu = Quý (hệ thống không tự dò, phải chọn đúng tay).</li>
        <li>Nếu chương trình bao phủ cả Nhóm hàng KHÔNG thuộc diện hỗ trợ (VD Casumina nhóm 12/NPP PCR Advenza-Milestar), khai vào <b>"Nhóm hàng LOẠI TRỪ khỏi chính sách bù"</b> — hệ thống tự bỏ các mã hàng thuộc nhóm đó ra khỏi số tiền bù, mã chiết khấu chính/Quý không cần tách riêng gì cả.</li>
      </ol>
    </div>

    <div class="card hd-step" id="hd-b9">
      <div class="card-head"><div><h2>Bước 9 — Tính chiết khấu</h2><div class="hint">Menu: Chiết khấu → Tính chiết khấu</div></div></div>
      <ol style="font-size:13px;color:var(--slate-600);padding-left:18px;margin:0 0 10px">
        <li style="margin-bottom:6px">Chọn <b>Thời điểm chiết khấu</b> — app tự quy đổi Từ/Đến ngày về đúng ranh giới kỳ chuẩn và chỉ lọc Mã chiết khấu cùng Thời điểm tính CK.</li>
        <li style="margin-bottom:6px">Có thể lọc thêm Nhà cung cấp / 1 Mã chiết khấu cụ thể, bấm <b>Tính chiết khấu</b>.</li>
        <li>Bấm <b>Lưu báo cáo</b> để ghi vào REPORT_CKTH/CKCT (cộng dồn, không mất kỳ trước), hoặc <b>Xuất Excel báo cáo</b>.</li>
      </ol>
      <div class="section-note" style="border-color:var(--red);color:var(--red)">Khung "Tính chiết khấu" và khung "Chiết khấu bổ sung" (lũy kế) bên dưới là <b>2 khu vực chọn kỳ độc lập</b>, không tự đồng bộ ngày — xem lại Bước 8 để tính nhanh nhiều tháng cùng lúc.</div>
    </div>

    <div class="card hd-step" id="hd-b9b">
      <div class="card-head"><div><h2>Bước 9b — Chiết khấu tạm tính theo tháng <span class="tag tag-gold">mới</span></h2><div class="hint">Menu: Chiết khấu → Chiết khấu tạm tính theo tháng</div></div></div>
      <div class="section-note" style="border-color:var(--red);color:var(--red)">Đây là công cụ ƯỚC TÍNH để có cơ sở BÁN HÀNG trong tháng đang chạy, dùng SONG SONG chứ KHÔNG THAY THẾ Bước 9 "Tính chiết khấu". Kết quả ở đây <b>không ghi vào REPORT_CKTH/CKCT</b> và không xuất hiện ở Bước 10 "Báo cáo &amp; đối chiếu" — số liệu chính thức để quyết toán/trả tiền vẫn luôn phải chạy và Lưu báo cáo ở Bước 9 cho đúng kỳ đã khép (tháng/quý/6 tháng/năm đã kết thúc).</div>
      <p style="font-size:13px;color:var(--slate-600)">Lý do cần công cụ riêng: nhiều Mã chiết khấu tính theo Quý/6 Tháng/Năm — nếu đợi kỳ đó kết thúc mới biết % thì NPP/ĐL không có cơ sở % chiết khấu nào để chào giá bán ngay trong các tháng giữa kỳ. Trang này lấy % của <b>kỳ tham chiếu gần nhất ĐÃ HOÀN TẤT</b> làm tạm tính cho tháng hiện hành, theo đúng quy tắc chọn kỳ dưới đây (không tự suy diễn khác cho từng Mã chiết khấu):</p>
      <div class="table-wrap"><table>
        <thead><tr><th>Thời điểm CK của Mã chiết khấu</th><th>Kỳ tham chiếu lấy % để tạm tính</th></tr></thead>
        <tbody>
          <tr><td>Tháng</td><td>Chính tháng hiện hành đang chạy</td></tr>
          <tr><td>Quý</td><td>Quý hiện hành nếu tháng đang chạy là <b>tháng cuối quý</b> (3/6/9/12); các tháng khác trong quý → lấy % của <b>quý liền trước</b></td></tr>
          <tr><td>6 Tháng</td><td>Nửa năm hiện hành nếu tháng đang chạy là <b>tháng 6 hoặc tháng 12</b>; các tháng khác → lấy % của <b>nửa năm liền trước</b></td></tr>
          <tr><td>Năm</td><td>Năm hiện hành nếu tháng đang chạy là <b>tháng 12</b>; các tháng khác → lấy % của <b>năm liền trước</b></td></tr>
        </tbody>
      </table></div>
      <p style="font-size:13px;color:var(--slate-600)">Doanh số/sản lượng dùng để so bậc và tính tiền tạm tính luôn lấy đúng mua hàng của <b>kỳ tham chiếu</b> ở bảng trên (không phải doanh số tháng hiện hành) — riêng cột Doanh số/Sản lượng hiển thị KPI đầu trang là của tháng hiện hành, chỉ để đối chiếu quy mô. Cơ chế chọn bậc, %KH, Mã doanh thu, "Không cộng vào Tổng CKTM" dùng lại NGUYÊN VẸN engine của Bước 9 (không có công thức tính riêng) — chỉ khác ở chỗ TỰ ĐỘNG chọn kỳ tham chiếu theo bảng trên thay vì để người dùng chọn Từ ngày–Đến ngày thủ công.</p>
      <p style="font-size:13px;color:var(--slate-600)">Mặt hàng phát sinh mua trong tháng nhưng <b>chưa đủ Giá công bố NCC</b> cho toàn bộ sản lượng tháng đó: cột "CK cơ bản" (chênh lệch Giá công bố − Giá mua ở Báo cáo 4) tạm coi bằng 0đ (không suy diễn có lợi/bất lợi) và được đánh dấu rõ ở cột <b>Ghi chú</b> — bổ sung Giá công bố NCC (Bước 2) rồi tính lại để có số chính xác.</p>
      <p style="font-size:13px;color:var(--slate-600)">2 báo cáo (Tổng hợp theo Nhóm hàng + Thương hiệu, Chi tiết theo mặt hàng) và Báo cáo 4 (Tổng hợp theo giá công bố) đều có thể lọc theo Nhóm hàng/Thương hiệu và xuất Excel.</p>
    </div>

    <div class="card hd-step" id="hd-b10">
      <div class="card-head"><div><h2>Bước 10 — Báo cáo & đối chiếu</h2><div class="hint">Menu: Báo cáo → Báo cáo chiết khấu / Đối chiếu giá mua - giá công bố</div></div></div>
      <p style="font-size:13px;color:var(--slate-600)">Các báo cáo đọc dữ liệu <b>đã lưu</b> ở REPORT_CKTH/CKCT (không tự tính lại từ Sổ mua hàng) — cần hoàn tất Bước 9 trước.</p>
      <div class="table-wrap"><table>
        <thead><tr><th>Chế độ xem</th><th>Dùng khi cần</th></tr></thead>
        <tbody>
          <tr><td>Theo mặt hàng &amp; hình thức CK</td><td>Chi tiết từng mã hàng đã tính CK, lọc theo hình thức/mã chiết khấu</td></tr>
          <tr><td>Tổng hợp theo Mã chiết khấu</td><td>Tổng tiền mỗi chương trình đã trả trong kỳ, kèm giá mua/giá công bố bình quân</td></tr>
          <tr><td>Bảng hạch toán theo mặt hàng</td><td>Chuẩn bị số liệu hạch toán kế toán</td></tr>
          <tr><td>Đối chiếu kỳ liền kề</td><td>So sánh kỳ này với kỳ trước, phát hiện chênh lệch bất thường</td></tr>
        </tbody>
      </table></div>
      <p style="font-size:13px;color:var(--slate-600)">Ngoài ra: <b>Đối chiếu giá mua – giá công bố</b> giúp phát hiện mã hàng thiếu dữ liệu giá công bố (ảnh hưởng đến tính năng "doanh số theo giá công bố" ở Bước 4); nút <b>Kiểm tra chiết khấu trùng</b> và <b>Bảng theo dõi chiết khấu</b> ở đầu trang Báo cáo hỗ trợ rà soát trước khi chốt số liệu.</p>
    </div>

    <div class="card hd-step" id="hd-faq">
      <div class="card-head"><div><h2>Mẹo & câu hỏi thường gặp</h2></div></div>
      <details class="hd-faq" open><summary>Có cần sao lưu dữ liệu không?</summary>
        <div class="a">Nên sao lưu định kỳ Google Sheet phía sau app (File → Tạo bản sao, hoặc Lịch sử phiên bản có sẵn) — đặc biệt trước khi nhập Excel hàng loạt hoặc xóa toàn bộ 1 danh mục.</div>
      </details>
      <details class="hd-faq"><summary>Excel nhập/xuất có ở những màn hình nào?</summary>
        <div class="a">Danh mục nhóm chiết khấu, Giá công bố NCC, Sổ chi tiết mua hàng, Danh mục doanh số, Danh mục kế hoạch, Danh mục chương trình, Chiết khấu lũy kế, và Báo cáo chiết khấu đều hỗ trợ tải lên/xuất Excel. Xem mục <b>Trợ giúp → Mẫu import Excel</b> để tải sẵn 1 file mẫu dùng chung cho mọi màn hình trên.</div>
      </details>
      <details class="hd-faq"><summary>Tải file Giá công bố / Sổ mua hàng lên xong sao chưa thấy trên Sheet?</summary>
        <div class="a">Từ 25/08/2026, 2 màn hình này KHÔNG ghi thẳng lên Sheet khi tải file nữa — sẽ mở ra màn hình <b>"Bản nháp"</b> để bạn xem trạng thái từng dòng (Mới/Cập nhật/Không đổi/Trùng lặp/Lỗi) và tick chọn đúng dòng muốn ghi. Phải bấm <b>"Ghi vào bảng chính"</b> ở cuối bản nháp thì dữ liệu mới thực sự lên Sheet — xem lại Bước 2/3 phía trên.</div>
      </details>
      <details class="hd-faq"><summary>Vì sao Danh mục Giá công bố không hiện đủ tất cả các dòng?</summary>
        <div class="a">Mặc định màn hình chỉ hiện các bản ghi <b>đang hiệu lực hôm nay</b> (đọc thẳng Sheet mỗi lần, không phải "thiếu dữ liệu"). Bỏ tick ô "Chỉ đang hiệu lực" ở đầu trang để xem cả các bản ghi đã hết hạn hoặc chưa tới ngày hiệu lực.</div>
      </details>
    </div>
`;

const HD_CONTENT_NGHIEPVU_ = `
    <div class="section-note">Chi tiết cách tính từng Mục chiết khấu — chọn đúng Bước tương ứng với chính sách đang cấu hình. Xem tab <b>"Hướng dẫn quy trình"</b> nếu cần ôn lại quy trình khai báo Danh mục/Mã chiết khấu/tải Excel.</div>

    <div class="card hd-step" id="hd-b11">
      <div class="card-head"><div><h2>Bước 11 — Ví dụ: Mục 6 &amp; Mục 7 TB 0101 <span class="tag tag-gold">mới 26/08/2026</span></h2><div class="hint">Kết hợp Bước 4–7 cho 2 mục phức tạp nhất</div></div></div>
      <p style="font-size:13px;color:var(--slate-600)"><b>Mục 6 (CK 6 tháng)</b>: tạo 2 Mã chiết khấu song song — <code>0101_TB_CSMN_006_CUNGKY</code> (Mã doanh thu <code>DT_6THANG_CUNGKY</code>, Kế hoạch = DS thực tế 6 tháng đầu 2025 = <b>7.978.328.956đ</b>) và <code>0101_TB_CSMN_006_KH</code> (Mã doanh thu <code>DT_6THANG_KH</code>, Kế hoạch = KH Quý I+II 2026 = <b>8.280.000.000đ</b>). Cả 2 dùng Thời điểm CK = "6 Tháng", Bậc: 100–&lt;110%KH → 0,5%; ≥110%KH → 1,0%. VD: DS 6 tháng 2026 thực tế = 11,17 tỷ → cả 2 phương án đều ra %KH ≥110% → CK 1,0% ≈ 111.700.000đ. Chọn đúng 1 phương án khi chốt báo cáo, dùng "Không cộng vào Tổng CKTM" cho phương án còn lại.</p>
      <p style="font-size:13px;color:var(--slate-600)"><b>Mục 7 (CK năm)</b>: 1 Mã chiết khấu, 24 Bậc điều kiện, VÀ (từ 28/08/2026) 2 khoảng ngày tách riêng — khoảng xét đạt %KH năm khác khoảng tính doanh số để trả thưởng. Xem hướng dẫn từng bước đầy đủ, kèm nguyên bảng 25 bậc điền sẵn, ở <b>Bước 14</b> bên dưới.</p>
    </div>

    <div class="card hd-step" id="hd-b12">
      <div class="card-head"><div><h2>Bước 12 — Ví dụ: CK 0,5% cho NPP/DL đặt hàng liên tục 4 tháng <span class="tag tag-gold">mới 27/08/2026</span></h2><div class="hint">Kết hợp Bước 8 (Chiết khấu lũy kế) + mẫu "mã cổng"</div></div></div>
      <p style="font-size:13px;color:var(--slate-600)">Chính sách: CK 0,5% trên doanh số giá công bố cho NPP/DL đặt hàng liên tục <b>4 tháng</b>, không tháng nào &lt;100%KH tháng, chia <b>2 giai đoạn</b> (T1–4/2026 và T7–10/2026); lỡ 1 tháng nhưng tổng 4 tháng vẫn đạt KH thì chỉ hưởng <b>80%</b> mức 0,5%. Cấu hình: tạo <b>1 "mã cổng"</b> (Tỷ lệ CK = 0, %KH Min = 100, Thời điểm = Tháng, Mã doanh thu giá công bố riêng) làm Mã chiết khấu chính — vì ngưỡng 100% ở đây khác ngưỡng của mã CK tháng đang trả tiền thật. Sau đó tạo <b>2 dòng Chiết khấu lũy kế</b> (Bước 8), Kiểu tính = "Đạt tất cả các tháng": dòng 1 tick tháng 1,2,3,4 + Hiệu lực 01/01–30/04/2026; dòng 2 tick tháng 7,8,9,10 + Hiệu lực 01/07–31/10/2026 — cả 2 dòng đều gán Mã chiết khấu chính = mã cổng trên, <b>%KH Min (cả kỳ) = 100</b> (bắt buộc sửa từ mặc định 0), Chiết khấu đầy đủ = 0,5%, % hưởng nếu lỡ tháng = 80.</p>
      <div class="section-note" style="border-color:var(--gold);color:#8a6a1a;background:var(--gold-soft)">Ở Danh mục kế hoạch (Bước 5), ngoài kỳ THÁNG cho cả 8 tháng liên quan, phải khai thêm <b>2 kỳ tùy chỉnh gộp cả giai đoạn</b> (VD "GD1_2026" 01/01–30/04, "GD2_2026" 01/07–31/10, Doanh thu KH = tổng 4 tháng) trong cùng 1 kế hoạch — thiếu 2 kỳ gộp này, app có thể tự lấy nhầm kế hoạch của 1 tháng lẻ làm mẫu số xét "đạt cả giai đoạn" mà không cảnh báo. Vận hành: tính + lưu báo cáo tháng cho mã cổng đủ cả 8 tháng trước, rồi mới chạy "Tính chiết khấu bổ sung" theo từng giai đoạn (Từ ngày–Đến ngày = đúng khung 4 tháng). Xem tài liệu <code>huong-dan-ck-4thang-lientuc-2026-08-27</code> trong dự án để có hướng dẫn từng bước đầy đủ kèm mọi tên trường chính xác.</div>
      <div class="section-note" style="border-color:var(--red,#c0392b);color:var(--red,#c0392b)">⚠ Sửa 28/08/2026 — phân biệt 2 loại "tháng không đạt": (1) tháng <b>CÓ đặt hàng</b> nhưng dưới 100%KH tháng → vẫn có thể hưởng mức <b>giảm còn 80%</b> nếu cả giai đoạn vẫn đạt KH — đúng câu chữ "có tháng không đạt sẽ chỉ hưởng 80%"; (2) tháng <b>HOÀN TOÀN KHÔNG phát sinh doanh số</b> (không đặt hàng) → làm gián đoạn "đặt hàng liên tục" ngay từ gốc → <b>mất toàn bộ</b> 0,5% (0đ), không phải chỉ giảm còn 80%, dù tổng 4 tháng vẫn đạt KH cả giai đoạn. Màn "Chiết khấu bổ sung" nay tự phát hiện đúng trường hợp (2) bằng cách đối chiếu trực tiếp dữ liệu mua hàng thô của từng tháng (không chỉ dựa vào báo cáo đã lưu) và hiện rõ tag "✗ Không có doanh số" + cảnh báo đỏ riêng.</div>
    </div>

    <div class="card hd-step" id="hd-b13">
      <div class="card-head"><div><h2>Bước 13 — Sơ đồ quy trình từng loại chiết khấu <span class="tag tag-gold">mới 27/08/2026</span></h2><div class="hint">Nhìn nhanh cách mỗi loại chiết khấu đang áp dụng được tính — hộp xanh = có tiền, hộp vàng = mức giảm/đặc biệt, hộp đỏ = không đạt</div></div></div>
      <div class="flow-grid"><div class="flow-panel">
          <div class="ttl">Mục 1 — CK cơ bản 10%</div>
          <div class="hint">Nhóm 1 Casumina — đã nằm sẵn trong giá hóa đơn</div>
          <div class="flow-diag"><div class="flow-box ">Đơn giá hóa đơn Casumina (đã trừ sẵn 10%)</div><div class="flow-down">↓</div><div class="flow-box gold">Mã chiết khấu: Tỷ lệ 10%, cờ &quot;Không cộng vào Tổng CKTM&quot; = BẬT</div><div class="flow-down">↓</div><div class="flow-box emerald">Chỉ ghi nhận riêng để đối chiếu — KHÔNG cộng vào Tổng CKTM báo cáo</div></div>
        </div><div class="flow-panel">
          <div class="ttl">Mục 2 — CK quý</div>
          <div class="hint">Theo %Hoàn thành KH quý, giá mua thực tế</div>
          <div class="flow-diag"><div class="flow-box ">Doanh số quý (giá mua thực tế) → so KH quý (Danh mục kế hoạch) → %Hoàn thành KH</div><div class="flow-down">↓</div><div class="flow-branch"><div class="col"><div class="lbl">Khớp 1 bậc</div><div class="flow-box emerald">CK = DS quý × Tỷ lệ bậc đó</div></div><div class="col"><div class="lbl">Không khớp bậc nào</div><div class="flow-box red">CK = 0</div></div></div></div>
        </div><div class="flow-panel">
          <div class="ttl">Mục 3 — CK tháng theo %KH doanh thu</div>
          <div class="hint">Giá công bố, nhiều bậc theo %KH</div>
          <div class="flow-diag"><div class="flow-box ">Sổ mua hàng tháng × Giá công bố NCC → Doanh số giá công bố</div><div class="flow-down">↓</div><div class="flow-box ">So với KH tháng (Danh mục kế hoạch) → %Hoàn thành KH tháng</div><div class="flow-down">↓</div><div class="flow-branch"><div class="col"><div class="lbl">Khớp 1 trong nhiều bậc (Tầng 3)</div><div class="flow-box emerald">CK = DS × Tỷ lệ bậc đó</div></div><div class="col"><div class="lbl">Không khớp bậc nào</div><div class="flow-box red">CK = 0</div></div></div></div>
        </div><div class="flow-panel">
          <div class="ttl">Mục 4 — CK quý tải nhẹ / tải nặng</div>
          <div class="hint">Tải nặng luôn hưởng mức CAO GẤP ĐÔI tải nhẹ ở cùng 1 bậc</div>
          <div class="flow-diag"><div class="flow-box ">Mua hàng quý → phân loại theo Công dụng (Tải nhẹ DT005 / Tải nặng DT004)</div><div class="flow-down">↓</div><div class="flow-box ">So với KH quý của đúng nhóm đó → %Hoàn thành KH quý</div><div class="flow-down">↓</div><div class="flow-branch"><div class="col"><div class="lbl">≤ 90% KH quý</div><div class="flow-box red">CK = 0 (cả 2 nhóm)</div></div><div class="col"><div class="lbl">90% – 100% KH quý</div><div class="flow-box gold">Tải nhẹ +0,5% · Tải nặng +1%</div></div><div class="col"><div class="lbl">&gt; 100% KH quý</div><div class="flow-box emerald">Tải nhẹ +1% · Tải nặng +2%</div></div></div></div>
        </div><div class="flow-panel">
          <div class="ttl">Mục 5 — CK 4 tháng liên tục</div>
          <div class="hint">Chi tiết đầy đủ ở Bước 12</div>
          <div class="flow-diag"><div class="flow-box ">&quot;Mã cổng&quot;: Tỷ lệ = 0, %KH Min = 100 → chỉ ghi ĐẠT/CHƯA ĐẠT từng tháng</div><div class="flow-down">↓</div><div class="flow-box ">Chiết khấu lũy kế: tick đủ 4 tháng của giai đoạn — đối chiếu báo cáo tháng ĐÃ LƯU + kiểm tra riêng doanh số thô từng tháng có = 0 không</div><div class="flow-down">↓</div><div class="flow-branch"><div class="col"><div class="lbl">Đủ cả 4 tháng + đạt KH cả kỳ</div><div class="flow-box emerald">CK ĐẦY ĐỦ 0,5%</div></div><div class="col"><div class="lbl">CÓ đặt hàng nhưng lỡ ≥1 tháng dưới KH, vẫn đạt KH cả kỳ</div><div class="flow-box gold">CK GIẢM còn 80%×0,5%</div></div><div class="col"><div class="lbl">≥1 tháng HOÀN TOÀN không đặt hàng (DS=0)</div><div class="flow-box red">CK = 0 (mất toàn bộ, dù cả kỳ vẫn đạt KH)</div></div><div class="col"><div class="lbl">Không đạt KH cả kỳ</div><div class="flow-box red">CK = 0</div></div></div></div>
        </div><div class="flow-panel">
          <div class="ttl">Mục 6 — CK 6 tháng đầu năm</div>
          <div class="hint">Chi tiết đầy đủ ở Bước 11</div>
          <div class="flow-diag"><div class="flow-box ">Doanh số 6 tháng đầu năm (giá mua thực tế)</div><div class="flow-down">↓</div><div class="flow-box ">So với mốc đã chọn: DS cùng kỳ năm trước HOẶC KH quý I+II (1 trong 2 phương án)</div><div class="flow-down">↓</div><div class="flow-branch"><div class="col"><div class="lbl">100% – &lt;110% mốc</div><div class="flow-box gold">CK 0,5%</div></div><div class="col"><div class="lbl">≥ 110% mốc</div><div class="flow-box emerald">CK 1,0%</div></div></div></div>
        </div><div class="flow-panel">
          <div class="ttl">Mục 7 — CK năm (24 bậc, 2 khoảng ngày)</div>
          <div class="hint">Chi tiết đầy đủ ở Bước 14</div>
          <div class="flow-diag"><div class="flow-box ">%Hoàn thành KH năm: tính RIÊNG trên &quot;Khoảng ngày xét đạt KH&quot; (VD 09/01–31/12) — so với KH năm</div><div class="flow-down">↓</div><div class="flow-box ">Doanh số dùng để so bậc Doanh thu (8 mức) VÀ tính tiền: lấy trên &quot;Hiệu lực từ/đến ngày&quot; (khoảng tính thưởng, VD 13 tháng — có thể KHÁC khoảng xét đạt KH ở trên)</div><div class="flow-down">↓</div><div class="flow-box ">Khớp ĐỒNG THỜI Bậc Doanh thu (8 mức, theo DS khoảng tính thưởng) × Bậc %KH (3 mức, theo %KH khoảng xét đạt KH) = 24 bậc</div><div class="flow-down">↓</div><div class="flow-box gold">CK = Doanh số khoảng tính thưởng × Tỷ lệ bậc khớp (1,05% – 1,65%)</div></div>
        </div><div class="flow-panel">
          <div class="ttl">CK thanh toán &amp; đặt hàng sớm</div>
          <div class="hint">Mục 10 + đặt hàng trước ngày 20 hằng tháng</div>
          <div class="flow-diag"><div class="flow-box ">Điều kiện thanh toán đúng hạn / đặt hàng trước ngày 20 hằng tháng</div><div class="flow-down">↓</div><div class="flow-branch"><div class="col"><div class="lbl">CK thanh toán (Mục 10)</div><div class="flow-box emerald">Tối đa 1,9%</div></div><div class="col"><div class="lbl">CK đặt hàng sớm</div><div class="flow-box emerald">0,5%</div></div></div><div class="flow-down">↓</div><div class="flow-box ">Cộng thẳng vào Tổng CKTM cùng các mục khác</div></div>
        </div><div class="flow-panel">
          <div class="ttl">Nhóm 2 PCR (Advenza &amp; Milestar)</div>
          <div class="hint">Quy định 42 — khác cơ chế Nhóm 1</div>
          <div class="flow-diag"><div class="flow-box ">Hóa đơn ký hiệu 1C26TTT → gộp SẢN LƯỢNG (chiếc), theo giá hóa đơn thực tế</div><div class="flow-down">↓</div><div class="flow-box ">So bậc thang SẢN LƯỢNG (không dùng giá công bố — khác Nhóm 1 Casumina)</div><div class="flow-down">↓</div><div class="flow-branch"><div class="col"><div class="lbl">Đạt 1 bậc sản lượng</div><div class="flow-box emerald">CK theo tỷ lệ bậc đó</div></div><div class="col"><div class="lbl">Chưa đạt bậc nào</div><div class="flow-box red">CK = 0</div></div></div></div>
        </div><div class="flow-panel">
          <div class="ttl">Cơ chế lũy kế &quot;2 điều kiện độc lập&quot;</div>
          <div class="hint">Dùng khi 1 chính sách cần cả 2 điều kiện tháng + cả kỳ</div>
          <div class="flow-diag"><div class="flow-box ">ĐK1: mỗi tháng đạt ≥X% KH tháng (đọc báo cáo THÁNG đã lưu)</div><div class="flow-down">↓</div><div class="flow-box ">ĐK2: cả kỳ đạt ≥Y% KH cả kỳ (tính trên toàn bộ khoảng đã chọn)</div><div class="flow-down">↓</div><div class="flow-box gold">Đạt ĐK nào cộng % của ĐK đó — đạt CẢ 2 thì CỘNG DỒN cả 2 tỷ lệ</div></div>
        </div></div>
    </div>

    <div class="card hd-step" id="hd-b14">
      <div class="card-head"><div><h2>Bước 14 — Mục 7 CK năm: hướng dẫn từng bước đầy đủ (2 khoảng ngày) <span class="tag tag-gold">mới 28/08/2026</span></h2><div class="hint">Kết hợp Bước 4–7 + 2 ô mới "Khoảng ngày xét đạt KH" trên Mã chiết khấu</div></div></div>
      <p style="font-size:13px;color:var(--slate-600)">Chính sách (TB 0101, mục 3.7 + bổ sung 28/08/2026): %hoàn thành KH năm xét trên khoảng <b>09/01–31/12/2026</b> so với KH năm; nếu đạt, tiền thưởng lại tính trên doanh số khoảng <b>01/12/2025–31/12/2026 (13 tháng)</b> nếu NPP/ĐL ký hợp đồng từ/trước 01/12/2025 (nếu không, dùng khoảng 01/01–31/12/2026 dương lịch thường) — 2 khoảng ngày khác nhau cho 2 việc khác nhau trên CÙNG 1 chương trình. Các mốc ngày/số KH dưới đây là ví dụ mẫu theo hồ sơ hiện có — khi cấu hình cho từng NPP cụ thể, thay đúng ngày ký hợp đồng và số KH năm thực tế của NPP đó.</p>

      <div class="section-note"><b>Bước 1 — Danh mục doanh số</b>: tạo Mã doanh thu riêng cho Mục 7, VD <code>CSMN_DT_NAM</code> — 6 bộ lọc mặt hàng để "NO" (toàn bộ nhóm Casumina), Thời điểm chiết khấu = <b>"Theo hiệu lực CTKM"</b>, Cách tính doanh số = <b>"Giá hóa đơn (giá mua thực tế)"</b> (giống Mục 6).</div>

      <div class="section-note"><b>Bước 2 — Danh mục kế hoạch</b>: tạo 1 Kế hoạch riêng (VD <code>KH_NAM_2026_MUC7</code>, Mã doanh thu áp dụng = đúng <code>CSMN_DT_NAM</code>), rồi vào "Chi tiết kế hoạch" → "Thêm kỳ tùy chỉnh" → 1 kỳ ĐÚNG khoảng xét đạt KH: Từ ngày <b>09/01/2026</b>, Đến ngày <b>31/12/2026</b>, Doanh thu kế hoạch = KH năm (VD 18.000.000.000đ). Không cần thêm kỳ nào khác — khoảng tính thưởng (13 tháng) chỉ cộng doanh số thật, không so với KH nào.</div>

      <div class="section-note"><b>Bước 3 — Mã chiết khấu (Tầng 2)</b>: tạo 1 mã, chọn "Không liên kết (tự nhập hiệu lực riêng)". Điền <b>Hiệu lực từ/đến ngày = 01/12/2025 → 31/12/2026</b> (khoảng TÍNH THƯỞNG, 13 tháng — đổi thành 01/01→31/12/2026 nếu NPP không thuộc diện 13 tháng), rồi điền tiếp 2 ô MỚI <b>Khoảng ngày xét đạt KH = 09/01/2026 → 31/12/2026</b> (khoảng XÉT ĐẠT KH — khác hẳn khoảng ở trên). Mã doanh thu (điều kiện) = <code>CSMN_DT_NAM</code>; Hình thức chiết khấu = "% theo Giá mua"; Thời điểm tính CK = "Theo hiệu lực CTKM"; Sản lượng/Doanh thu Min-Max = 0/0 (để 24 Bậc điều kiện quyết định); % Kế hoạch Min/Max để trống (bậc con sẽ ghi đè).</div>

      <div class="section-note"><b>Bước 4 — Bậc điều kiện (Tầng 3)</b>: nhập đủ 24 bậc theo đúng bảng TB 0101 (Sản lượng Min/Max = 0/0, Điều kiện SL/DT = AND, Doanh thu kế hoạch để trống ở mọi bậc):
        <div class="table-wrap" style="margin-top:8px"><table>
          <thead><tr><th>Bậc</th><th class="num">Doanh thu Min (đ)</th><th class="num">Doanh thu Max (đ, 0=∞)</th><th class="num">%KH Min</th><th class="num">%KH Max (0=∞)</th><th class="num">Tỷ lệ CK</th></tr></thead>
          <tbody>
            <tr style="color:var(--slate-500)"><td>0 (khuyến nghị thêm)</td><td class="num mono">0</td><td class="num mono">0</td><td class="num mono">0</td><td class="num mono">94,9999999</td><td class="num mono">0%</td></tr>
            <tr><td>1</td><td class="num mono">0</td><td class="num mono">11.999.999.999</td><td class="num mono">95</td><td class="num mono">99,9999999</td><td class="num mono">1,05%</td></tr>
            <tr><td>2</td><td class="num mono">0</td><td class="num mono">11.999.999.999</td><td class="num mono">100</td><td class="num mono">109,9999999</td><td class="num mono">1,15%</td></tr>
            <tr><td>3</td><td class="num mono">0</td><td class="num mono">11.999.999.999</td><td class="num mono">110</td><td class="num mono">0</td><td class="num mono">1,25%</td></tr>
            <tr><td>4</td><td class="num mono">12.000.000.000</td><td class="num mono">17.999.999.999</td><td class="num mono">95</td><td class="num mono">99,9999999</td><td class="num mono">1,10%</td></tr>
            <tr><td>5</td><td class="num mono">12.000.000.000</td><td class="num mono">17.999.999.999</td><td class="num mono">100</td><td class="num mono">109,9999999</td><td class="num mono">1,20%</td></tr>
            <tr><td>6</td><td class="num mono">12.000.000.000</td><td class="num mono">17.999.999.999</td><td class="num mono">110</td><td class="num mono">0</td><td class="num mono">1,30%</td></tr>
            <tr><td>7</td><td class="num mono">18.000.000.000</td><td class="num mono">23.999.999.999</td><td class="num mono">95</td><td class="num mono">99,9999999</td><td class="num mono">1,15%</td></tr>
            <tr><td>8</td><td class="num mono">18.000.000.000</td><td class="num mono">23.999.999.999</td><td class="num mono">100</td><td class="num mono">109,9999999</td><td class="num mono">1,25%</td></tr>
            <tr><td>9</td><td class="num mono">18.000.000.000</td><td class="num mono">23.999.999.999</td><td class="num mono">110</td><td class="num mono">0</td><td class="num mono">1,35%</td></tr>
            <tr><td>10</td><td class="num mono">24.000.000.000</td><td class="num mono">29.999.999.999</td><td class="num mono">95</td><td class="num mono">99,9999999</td><td class="num mono">1,20%</td></tr>
            <tr><td>11</td><td class="num mono">24.000.000.000</td><td class="num mono">29.999.999.999</td><td class="num mono">100</td><td class="num mono">109,9999999</td><td class="num mono">1,30%</td></tr>
            <tr><td>12</td><td class="num mono">24.000.000.000</td><td class="num mono">29.999.999.999</td><td class="num mono">110</td><td class="num mono">0</td><td class="num mono">1,40%</td></tr>
            <tr><td>13</td><td class="num mono">30.000.000.000</td><td class="num mono">35.999.999.999</td><td class="num mono">95</td><td class="num mono">99,9999999</td><td class="num mono">1,25%</td></tr>
            <tr><td>14</td><td class="num mono">30.000.000.000</td><td class="num mono">35.999.999.999</td><td class="num mono">100</td><td class="num mono">109,9999999</td><td class="num mono">1,35%</td></tr>
            <tr><td>15</td><td class="num mono">30.000.000.000</td><td class="num mono">35.999.999.999</td><td class="num mono">110</td><td class="num mono">0</td><td class="num mono">1,45%</td></tr>
            <tr><td>16</td><td class="num mono">36.000.000.000</td><td class="num mono">44.999.999.999</td><td class="num mono">95</td><td class="num mono">99,9999999</td><td class="num mono">1,35%</td></tr>
            <tr><td>17</td><td class="num mono">36.000.000.000</td><td class="num mono">44.999.999.999</td><td class="num mono">100</td><td class="num mono">109,9999999</td><td class="num mono">1,45%</td></tr>
            <tr><td>18</td><td class="num mono">36.000.000.000</td><td class="num mono">44.999.999.999</td><td class="num mono">110</td><td class="num mono">0</td><td class="num mono">1,55%</td></tr>
            <tr><td>19</td><td class="num mono">45.000.000.000</td><td class="num mono">59.999.999.999</td><td class="num mono">95</td><td class="num mono">99,9999999</td><td class="num mono">1,40%</td></tr>
            <tr><td>20</td><td class="num mono">45.000.000.000</td><td class="num mono">59.999.999.999</td><td class="num mono">100</td><td class="num mono">109,9999999</td><td class="num mono">1,50%</td></tr>
            <tr><td>21</td><td class="num mono">45.000.000.000</td><td class="num mono">59.999.999.999</td><td class="num mono">110</td><td class="num mono">0</td><td class="num mono">1,60%</td></tr>
            <tr><td>22</td><td class="num mono">60.000.000.000</td><td class="num mono">0</td><td class="num mono">95</td><td class="num mono">99,9999999</td><td class="num mono">1,45%</td></tr>
            <tr><td>23</td><td class="num mono">60.000.000.000</td><td class="num mono">0</td><td class="num mono">100</td><td class="num mono">109,9999999</td><td class="num mono">1,55%</td></tr>
            <tr><td>24</td><td class="num mono">60.000.000.000</td><td class="num mono">0</td><td class="num mono">110</td><td class="num mono">0</td><td class="num mono">1,65%</td></tr>
          </tbody>
        </table></div>
        <div class="hint" style="margin-top:6px">Bậc 0 (dưới 95%KH, tỷ lệ 0%) không có trong văn bản gốc — thêm để mọi mức %KH từ 0% trở lên luôn có đúng 1 bậc khớp, giúp "% Hoàn thành KH" hiển thị rõ ràng khi rà soát thay vì để trống ở dải dưới 95%. Không ảnh hưởng số tiền chiết khấu (vẫn 0đ dù có hay không có bậc này).</div>
      </div>

      <div class="section-note" style="border-color:var(--gold);color:#8a6a1a;background:var(--gold-soft)"><b>Bước 5 — Tính thử để đối chiếu</b>: vào Tính chiết khấu → Thời điểm chiết khấu = "Theo hiệu lực CTKM" (tự nhảy đúng khoảng Hiệu lực đã khai) → chọn đúng Mã chiết khấu Mục 7 → bấm Tính chiết khấu. Kiểm tra: Doanh số hiển thị phải là TOÀN BỘ khoảng tính thưởng (13 tháng), không phải khoảng 09/01–31/12; %Hoàn thành KH hiện kèm dòng chú thích tím "⏱ %KH tính trên khoảng riêng 09/01/2026→31/12/2026 (khác khoảng đang tính thưởng)" — đây là dấu hiệu cơ chế tách khoảng ngày đang chạy đúng; Bậc khớp phải đúng CẢ 2 chiều (quy mô DT theo khoảng tính thưởng, %KH theo khoảng xét đạt).</div>

      <div class="section-note" style="border-color:var(--red,#c0392b);color:var(--red,#c0392b)">⚠ Nếu áp dụng kỳ 13 tháng, Sổ chi tiết mua hàng phải có đủ hóa đơn Tháng 12/2025 — thiếu tháng này, Doanh số khoảng tính thưởng sẽ bị thiếu và ra số CK thấp hơn thực tế. Cơ chế 2 khoảng ngày hiện chỉ áp dụng cho Mã chiết khấu dùng trực tiếp (Tầng 2), chưa áp dụng cho Chiết khấu lũy kế (Mục 5 dùng cơ chế khác, không cần tới đây).</div>
    </div>

    <div class="card hd-step" id="hd-faq">
      <div class="card-head"><div><h2>Mẹo & câu hỏi thường gặp</h2></div></div>
      <details class="hd-faq" open><summary>Vì sao 1 Mã chiết khấu không cho nhập Thương hiệu/Nhóm hàng/... trực tiếp?</summary>
        <div class="a">Phạm vi mặt hàng được <b>kế thừa tự động</b> từ Mã doanh thu (điều kiện) đã chọn ở Bước 7 — để 1 bộ lọc sản phẩm dùng lại được cho nhiều chương trình. Muốn đổi phạm vi, sửa/tạo mới Mã doanh thu ở Bước 4 rồi chọn lại.</div>
      </details>
      <details class="hd-faq"><summary>Vừa tick "Giá công bố" cho 1 Mã doanh thu — kết quả có đổi ngay không?</summary>
        <div class="a">Có — mọi Mã chiết khấu tham chiếu Mã doanh thu đó dùng ngay cơ sở mới ở lần "Tính chiết khấu" tiếp theo. Báo cáo <b>đã lưu</b> trước đó không tự đổi lại — cần tính và lưu lại nếu muốn cập nhật số liệu kỳ cũ.</div>
      </details>
      <details class="hd-faq"><summary>Vì sao 1 mã chiết khấu báo "CHƯA ĐẠT" dù số liệu nhìn có vẻ ổn?</summary>
        <div class="a">Kiểm tra theo thứ tự: (1) mặt hàng có nằm đúng phạm vi Mã doanh thu đã chọn; (2) nếu có Bậc điều kiện (Tầng 3), có bậc nào khớp đúng khoảng SL/DT/%KH thực tế; (3) %KH Min/Max trên Mã chiết khấu hoặc bậc đang chọn có đang giới hạn quá chặt.</div>
      </details>
    </div>
`;

// Trả về HTML đầy đủ (tab bar + mục lục + nội dung) cho tab "Hướng dẫn sử dụng", theo đúng sub-tab đang
// chọn ('quytrinh' hoặc 'nghiepvu') — client gọi qua google.script.run mỗi lần vào tab hoặc đổi sub-tab.
function getHuongDanContentHtml(sub) {
  const isNV = sub === 'nghiepvu';
  const tocGroups = isNV ? HD_TOC_NGHIEPVU_ : HD_TOC_QUYTRINH_;
  const tabBar = '\n    <div class="tabs2">\n      <div class="tab2 ' + (!isNV ? 'active' : '') + '" data-sub="quytrinh">Hướng dẫn quy trình</div>\n      <div class="tab2 ' + (isNV ? 'active' : '') + '" data-sub="nghiepvu">Hướng dẫn nghiệp vụ chiết khấu</div>\n    </div>\n  ';
  const tocCard = '\n    <div class="card">\n      <div class="card-head"><div><h2>Mục lục</h2></div></div>\n      <div class="hd-toc">\n        ' +
    tocGroups.map(function(g) {
      return '<div class="hd-toc-grp">' + (g.lbl ? '<div class="lbl">' + esc_(g.lbl) + '</div>' : '') +
        g.items.map(function(it) { return '<a href="#' + it[0] + '">' + esc_(it[1]) + '</a>'; }).join('') +
        '</div>';
    }).join('') +
    '\n      </div>\n    </div>\n  ';
  return tabBar + tocCard + (isNV ? HD_CONTENT_NGHIEPVU_ : HD_CONTENT_QUYTRINH_);
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

// =====================================================================================
// HÀM KHỞI TẠO CHIẾT KHẤU NHÓM 2 — PCR ADVENZA & MILESTAR (CHẠY 1 LẦN, THỦ CÔNG)
// -----------------------------------------------------------------------------------
// Bối cảnh (rà soát 24/08/2026): Quy định .../QĐ-BH-MAR quy định CKTM cho dòng PCR Advenza &
// Milestar (Advenza + Milestar GỘP CHUNG 1 bậc thang, không tách nhãn hiệu) theo 3 mốc thời
// gian Tháng / Quý / Năm, tính trên SỐ LƯỢNG (chiếc) mua trong kỳ, % trên giá bán NPP (giá hóa
// đơn, chưa VAT theo văn bản). Tại thời điểm viết hàm này, DM_CHUONGTRINH/CHUONGTRINHCHIETKHAU
// CHƯA có bất kỳ dòng nào cho nhóm PCR — dù Sổ chi tiết mua hàng đã có ~129 dòng hóa đơn PCR
// thực tế (~1,43 tỷ đồng) hoàn toàn CHƯA được tính chiết khấu gì. Đây thuần túy là THIẾU DỮ
// LIỆU CẤU HÌNH (engine đã hỗ trợ sẵn, generic theo Tầng 1/2/3), không phải lỗi code — hàm này
// nhập đủ dữ liệu đó 1 lần, đúng số liệu trong văn bản.
//
// Lọc đúng sản phẩm PCR: dùng dim_thuonghieu = "PCR" (THEO XÁC NHẬN TRỰC TIẾP của người dùng —
// chỉ cần tạo mã chiết khấu và chọn Thương hiệu = PCR là đủ, không cần gì thêm). Hàm tự thêm giá
// trị "PCR" vào danh mục DM_THUONGHIEU nếu chưa có (trước đây chỉ có NO/CSMN/PNC). LƯU Ý: các mã
// hàng PCR trong DM_MHCK cần được gắn đúng Thương hiệu = "PCR" thì bộ lọc này mới khớp được sản
// phẩm — dữ liệu snapshot cũ (24/08/2026) mới chỉ gắn ở cột Đặc tính (dactinh=PCR), CHƯA gắn ở cột
// Thương hiệu, nên cần cập nhật lại cột Thương hiệu cho 135 mã hàng PCR trong DM_MHCK trước khi
// tính, nếu không toàn bộ 4 mã chiết khấu PCR sẽ khớp 0 sản phẩm và luôn ra 0đ. Các dim khác để
// "NO" (không phân biệt) — dimsMatch() trong Index.html coi "NO" là ký tự đại diện (khớp mọi giá
// trị), nên bộ lọc chỉ phụ thuộc đúng 1 cột Thương hiệu này.
//
// AN TOÀN: hàm tự kiểm tra — nếu chương trình PCR (ma_chuongtrinh) đã tồn tại thì KHÔNG tạo lại
// (không tạo trùng dữ liệu) — chạy lại nhiều lần vô hại, chỉ log lại thông báo đã tồn tại.
//
// NHỮNG ĐIỂM CẦN NGƯỜI DÙNG XÁC NHẬN LẠI TRƯỚC KHI DÙNG SỐ LIỆU ĐỂ CHI TRẢ THẬT:
//  (1) "Giá bán NPP (hóa đơn), CHƯA VAT" — cần xác nhận cột "giatri" (Giá trị mua) trong Sổ chi
//      tiết mua hàng hiện lưu giá trị CHƯA VAT hay ĐÃ GỒM VAT. Nếu đã gồm VAT, % chiết khấu tính
//      ra theo cấu hình PERCENT_GIAMUA bên dưới sẽ CAO HƠN thực tế một chút (chênh đúng % VAT).
//  (2) "Vinh danh >12.000 chiếc/năm: +0,5%" — được tạo ở đây thành 1 mã CHIẾT KHẤU RIÊNG, cộng
//      dồn thêm vào mã "Năm" (không thay thế) — đúng tinh thần "cộng thêm" của văn bản.
//  (3) "Hình ảnh TTDV 100% chuẩn (Advenza Tire Spa): +1,0%" — là tiêu chí ĐỊNH TÍNH, không tính
//      được từ dữ liệu mua hàng — KHÔNG được tạo tự động ở đây, vẫn cần cộng thủ công khi đủ
//      điều kiện.
//  (4) Cách CỘNG DỒN Tháng+Quý+Năm với nhau: văn bản ghi "cộng dồn với nhau (không loại trừ lẫn
//      nhau) — nhưng cần Casumina xác nhận chính thức cách cộng dồn nếu chưa có văn bản nêu rõ".
//      Hàm này tạo 4 mã ĐỘC LẬP (Tháng/Quý/Năm/Vinh danh) — app sẽ tự cộng dồn cả 4 vào "Tổng
//      CKTM" như mọi chương trình khác đang hoạt động cùng lúc. Nếu Casumina xác nhận KHÔNG được
//      cộng dồn (chỉ lấy 1 mức), cần sửa lại cách xử lý (VD đánh dấu "Không cộng vào Tổng CKTM"
//      cho các mã không nên cộng, giống cơ chế đã làm cho Mục 1).
//  (5) Sau khi chạy hàm này, mở lại app (F5) MỘT LẦN để app tự gán "Mã doanh thu" cho 4 mã chiết
//      khấu mới (cơ chế migrateProgramsToDoanhThu_ ở Index.html tự tạo khi phát hiện mã chiết
//      khấu chưa có ma_doanhthu) — không cần thao tác gì thêm cho bước này.
//
// CÁCH CHẠY: mở Apps Script editor -> chọn hàm "seedPCRProgram" ở dropdown trên cùng -> bấm ▶
// Run. Xem kết quả ở View -> Logs (hoặc Executions).
// =====================================================================================
function seedPCRProgram() {
  return withLock_(() => seedPCRProgram_impl_(), 60000);
}
function seedPCRProgram_impl_() {
  const ss = SpreadsheetApp.getActive();
  const MA_CT = 'PCR_ADVENZA_MILESTAR_2026';

  function readAsObjects_(sheetName) {
    const sh = getOrCreateSheet_(ss, sheetName);
    const data = sh.getDataRange().getValues();
    if (data.length < 1) return { rows: [] };
    const headers = data[0];
    const rows = data.slice(1)
      .filter(r => r.some(c => String(c).trim() !== ''))
      .map(r => { const o = {}; headers.forEach((h, i) => o[h] = sanitizeCellValue_(r[i], h)); return o; });
    return { rows };
  }
  function appendRows_(sheetName, existingRows, newRows) {
    if (!newRows.length) return;
    const merged = existingRows.concat(newRows);
    const headers = Object.keys(merged[0]);
    merged.forEach(r => Object.keys(r).forEach(k => { if (headers.indexOf(k) === -1) headers.push(k); }));
    const sh = getOrCreateSheet_(ss, sheetName);
    writeRowsEfficient_(sh, headers, merged.map(r => headers.map(h => r[h] ?? '')));
  }

  // ----- Danh mục Thương hiệu: đảm bảo có sẵn giá trị "PCR" để dùng làm bộ lọc (dim_thuonghieu) -----
  // THEO YÊU CẦU (rà soát 24/08/2026, người dùng xác nhận trực tiếp): lọc sản phẩm PCR bằng cột
  // "Thương hiệu" (dim_thuonghieu = 'PCR'), KHÔNG dùng "Đặc tính" (dactinh) như bản seed trước — chỉ
  // cần tạo mã chiết khấu và chọn thương hiệu PCR là đủ, không cần thêm bước nào khác. DM_THUONGHIEU
  // hiện chỉ có NO/CSMN/PNC — hàm tự thêm dòng "PCR" nếu chưa có, để giá trị này hợp lệ khi lọc.
  const thData = readAsObjects_(SHEETS.THUONGHIEU);
  if (!thData.rows.some(r => String(r.ma).trim().toUpperCase() === 'PCR')) {
    appendRows_(SHEETS.THUONGHIEU, thData.rows, [{ ma: 'PCR', ten: 'PCR Advenza & Milestar' }]);
  }

  // ----- Tầng 1: Chương trình -----
  const ctData = readAsObjects_(SHEETS.CHUONGTRINH);
  const already = ctData.rows.find(r => r.ma_chuongtrinh === MA_CT);
  if (already) {
    const msg = 'Chương trình "' + MA_CT + '" đã tồn tại (id=' + already.id + ') — KHÔNG tạo lại để tránh trùng dữ liệu. ' +
      'Nếu muốn seed lại từ đầu, xoá thủ công dòng đó trong DM_CHUONGTRINH và các mã chiết khấu con liên quan trong CHUONGTRINHCHIETKHAU/DM_DIEUKIEN rồi chạy lại.';
    Logger.log(msg);
    return msg;
  }
  const ctId = Utilities.getUuid();
  const todayISO = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd');
  appendRows_(SHEETS.CHUONGTRINH, ctData.rows, [{
    id: ctId, ma_chuongtrinh: MA_CT,
    noidung: 'Chiết khấu PCR Advenza & Milestar (gộp chung, theo Quy định .../QĐ-BH-MAR) — tạo tự động bởi seedPCRProgram()',
    tu_ngay: '2026-01-01', den_ngay: '2026-12-31'
  }]);

  // ----- Tầng 2 (mã chiết khấu) + Tầng 3 (bậc SL) -----
  const pData = readAsObjects_(SHEETS.PROGRAMS);
  const dData = readAsObjects_(SHEETS.DIEUKIEN);
  const newPrograms = [];
  const newDieuKien = [];
  // NO = không phân biệt (dimsMatch() coi "NO" là ký tự đại diện); dim_thuonghieu='PCR' = lọc đúng
  // toàn bộ mã hàng PCR Advenza/Milestar theo cột Thương hiệu (theo đúng cách người dùng xác nhận).
  // LƯU Ý: cần đảm bảo các mã hàng PCR trong DM_MHCK đã được gắn Thương hiệu = "PCR" (không phải chỉ
  // gắn ở cột Đặc tính như dữ liệu snapshot cũ) — nếu chưa, hãy cập nhật lại cột Thương hiệu cho các
  // mã hàng PCR trong DM_MHCK trước khi tính chiết khấu, nếu không bộ lọc này sẽ không khớp được sản
  // phẩm nào và toàn bộ 4 mã chiết khấu PCR sẽ luôn ra 0đ.
  const DIM_DEFAULTS = { dim_thuonghieu: 'PCR', dim_nhomkh: 'NO', dim_loaisp: 'NO', dim_dactinh: 'NO', dim_congdung: 'NO', dim_quycach: 'NO' };

  function addProgram_(maChietkhau, thoidiem, ghiChuCT, bacList) {
    const progId = Utilities.getUuid();
    newPrograms.push(Object.assign({
      id: progId, chuongtrinh_id: ctId, ma_chietkhau: maChietkhau,
      hieuluctu: '2026-01-01', hieulucden: '2026-12-31',
      // PERCENT_GIAMUA = % trên giá trị hóa đơn thực tế (giá bán NPP) — đúng cơ sở văn bản yêu cầu,
      // KHÁC với Mục 1/Mục 3 Nhóm 1 (dùng PERCENT_GIACONGBO/giá công bố) — xem lưu ý (1) ở đầu hàm
      // về VAT trước khi dùng số liệu để chi trả thật.
      ht_chietkhau: 'PERCENT_GIAMUA', tt_ck: 'GIAM_HOADON',
      thoidiem_ck: thoidiem, dieukien_sldt: 'AND',
      ma_doanhthu: '', ma_doanhthu_tinhck: '', // để trống — app tự tạo/gán khi mở lại (xem lưu ý (5))
      pct_kh_min: '', pct_kh_max: '', dt_kehoach: '',
      sl_min: 0, sl_max: 0, dt_min: 0, dt_max: 0, tl_ck: '', dongia_codinh: '',
      ngaytao: todayISO,
      ghichu: ghiChuCT + ' — tạo tự động bởi seedPCRProgram()'
    }, DIM_DEFAULTS));
    bacList.forEach(b => {
      newDieuKien.push({
        id: Utilities.getUuid(), program_id: progId,
        sl_min: b.slMin, sl_max: b.slMax, dt_min: 0, dt_max: 0, dieukien_sldt: 'AND',
        pct_kh_min: '', pct_kh_max: '', dt_kehoach: '',
        tl_ck: b.tl, dongia_codinh: '', ghichu: b.ghichu
      });
    });
  }

  addProgram_('PCR_ADVENZA_MILESTAR_THANG', 'Tháng', 'CKTM PCR theo tháng (số chiếc)', [
    { slMin: 1,    slMax: 99,   tl: 7.0,  ghichu: 'Bậc 1 (1-99 chiếc)' },
    { slMin: 100,  slMax: 199,  tl: 7.5,  ghichu: 'Bậc 2 (100-199 chiếc)' },
    { slMin: 200,  slMax: 299,  tl: 8.0,  ghichu: 'Bậc 3 (200-299 chiếc)' },
    { slMin: 300,  slMax: 499,  tl: 8.5,  ghichu: 'Bậc 4 (300-499 chiếc)' },
    { slMin: 500,  slMax: 699,  tl: 9.0,  ghichu: 'Bậc 5 (500-699 chiếc)' },
    { slMin: 700,  slMax: 999,  tl: 9.5,  ghichu: 'Bậc 6 (700-999 chiếc)' },
    { slMin: 1000, slMax: 0,    tl: 10.0, ghichu: 'Bậc 7 (từ 1000 chiếc)' }
  ]);
  addProgram_('PCR_ADVENZA_MILESTAR_QUY', 'Quý', 'CKTM PCR theo quý (số chiếc)', [
    { slMin: 100,  slMax: 299,  tl: 1.0, ghichu: 'Bậc 1 (100-299 chiếc)' },
    { slMin: 300,  slMax: 599,  tl: 1.5, ghichu: 'Bậc 2 (300-599 chiếc)' },
    { slMin: 600,  slMax: 899,  tl: 2.0, ghichu: 'Bậc 3 (600-899 chiếc)' },
    { slMin: 900,  slMax: 1499, tl: 2.5, ghichu: 'Bậc 4 (900-1499 chiếc)' },
    { slMin: 1500, slMax: 2099, tl: 3.0, ghichu: 'Bậc 5 (1500-2099 chiếc)' },
    { slMin: 2100, slMax: 2999, tl: 3.5, ghichu: 'Bậc 6 (2100-2999 chiếc)' },
    { slMin: 3000, slMax: 0,    tl: 4.0, ghichu: 'Bậc 7 (từ 3000 chiếc)' }
  ]);
  addProgram_('PCR_ADVENZA_MILESTAR_NAM', 'Năm', 'CKTM PCR theo năm (số chiếc)', [
    { slMin: 400,   slMax: 1199,  tl: 1.0, ghichu: 'Bậc 1 (400-1199 chiếc)' },
    { slMin: 1200,  slMax: 2399,  tl: 1.5, ghichu: 'Bậc 2 (1200-2399 chiếc)' },
    { slMin: 2400,  slMax: 3599,  tl: 2.0, ghichu: 'Bậc 3 (2400-3599 chiếc)' },
    { slMin: 3600,  slMax: 5999,  tl: 2.5, ghichu: 'Bậc 4 (3600-5999 chiếc)' },
    { slMin: 6000,  slMax: 8399,  tl: 3.0, ghichu: 'Bậc 5 (6000-8399 chiếc)' },
    { slMin: 8400,  slMax: 11999, tl: 3.5, ghichu: 'Bậc 6 (8400-11999 chiếc)' },
    { slMin: 12000, slMax: 0,     tl: 4.0, ghichu: 'Bậc 7 (từ 12000 chiếc)' }
  ]);
  // Vinh danh — mã RIÊNG, cộng dồn THÊM vào mã Năm (không thay thế), theo đúng lưu ý (2) ở đầu hàm.
  addProgram_('PCR_ADVENZA_MILESTAR_VINHDANH', 'Năm', 'CK Vinh danh PCR (>12.000 chiếc/năm, cộng dồn thêm mã Năm)', [
    { slMin: 12001, slMax: 0, tl: 0.5, ghichu: 'Vinh danh (>12.000 chiếc/năm)' }
  ]);

  appendRows_(SHEETS.PROGRAMS, pData.rows, newPrograms);
  appendRows_(SHEETS.DIEUKIEN, dData.rows, newDieuKien);

  bumpDataVersion_();
  const msg = 'Đã tạo chương trình "' + MA_CT + '" + ' + newPrograms.length + ' mã chiết khấu (Tháng/Quý/Năm/Vinh danh PCR Advenza-Milestar, lọc theo Thương hiệu=PCR), tổng ' +
    newDieuKien.length + ' bậc. NHỚ: (a) đảm bảo các mã hàng PCR trong DM_MHCK đã gắn Thương hiệu=PCR (nếu chỉ có Đặc tính=PCR thì cần cập nhật lại), ' +
    '(b) mở lại app (F5) 1 lần để app tự gán Mã doanh thu cho các mã mới, (c) đọc kỹ 5 lưu ý ở đầu hàm (đặc biệt điểm (1) về VAT và điểm (4) về cách cộng dồn) TRƯỚC KHI dùng số liệu để chi trả thật.';
  Logger.log(msg);
  return msg;
}

// Chuỗi base64 GỐC (SEED) của file mẫu import Excel (MAU_NHAP_LIEU_CHIET_KHAU_CASUMINA.xlsx) — CHỈ
// dùng đúng 1 lần để tạo Google Sheet template ở buildMauTemplateSheetFromXlsx_() phía trên, không còn
// được trả trực tiếp cho trình duyệt nữa. Sau khi Sheet template đã tạo, mọi chỉnh sửa về sau nên làm
// TRỰC TIẾP trên Sheet đó (xem getMauNhapLieuTemplateSheetUrl()), không cần sửa lại chuỗi này. Đặt ở
// cuối file cho gọn, không xen giữa các hàm nghiệp vụ.
const MAU_NHAP_LIEU_XLSX_B64_SEED_ = "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";
