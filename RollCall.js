function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const enabled =
    PropertiesService.getScriptProperties()
      .getProperty("AUTO_ATTENDANCE") === "true";
  ui.createMenu("📅 Công cụ điểm danh")
    .addItem("Chọn tháng/năm", "showDatePicker")
    .addItem("Cập nhật thông tin học viên", "updateInfoData")
    .addItem('Chọn lịch học cho học viên', 'openCalenderPopupForSelectedCell')
    .addItem('Điểm danh học viên', 'openAttendancePopupForSelectedCell')
    .addSeparator()
    .addItem(
      enabled
        ? "Điểm danh tự động ✅"
        : "Điểm danh tự động",
      "toggleAutoAttendance"
    )
    .addToUi();
}

function onEdit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const editedSheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const lastCol = editedSheet.getLastColumn();

  // Chỉ chạy nếu sheet được chỉnh sửa tên bắt đầu bằng "ĐIỂM DANH"
  if (editedSheet.getName().startsWith('ĐIỂM DANH')) {
    // Chỉ chạy tổng hợp điểm danh khi cột điểm thay đổi (cột từ G trở đi)
    const firstDataCol = 7; // cột G
    if (col >= firstDataCol && col < lastCol) {
      tongHopDiemDanh(row, col, editedSheet);
    }
  }
  if (editedSheet.getName() == 'TỔNG ĐIỂM DANH') {
    // --- Cột Điểm danh năm ngoái ---
    if (col === 3 && row >= 2) {
      tongHopDiemDanh();
    }
  }
  if (editedSheet.getName() == 'THÔNG TIN HỌC VIÊN') {
    // --- Cột Giờ học (F=6) ---
    if ((col === 6 && row >= 2) || (col === 7 && row >= 2)) {
      capNhatTKB();
      taoLessonInstance();
      hienThiLichHomNayNgayMai();
    }
    // --- Cột Trạng thái (G=7) ---
    if (col === 7 && row >= 2) {
      capNhatTrangThaiHocVien();
    }
    // --- Cột Tên ---
    if (col === 2 && row >=2) {
      kiemTraTrungLapTenHocVien(e, editedSheet);
    }
    sumConfigMatches();
  }
}

function onSelectionChange(e) {
  const sheet = e.range.getSheet();

  if (sheet.getName() !== "THỜI KHÓA BIỂU") return;

  const state = {
    sheet: sheet.getName(),
    cell: e.range.getA1Notation(),
    value: e.range.getDisplayValue(),
    ts: Date.now()
  };

  PropertiesService
      .getDocumentProperties()
      .setProperty(
        "ACTIVE_TKB_CELL",
        JSON.stringify(state)
      );
}

function updateInfoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    if (!sheet.getName().startsWith("ĐIỂM DANH")) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) return;

    // Lấy toàn bộ dữ liệu trong 1 lần
    const data = sheet.getRange(3, 1, lastRow - 1, lastCol).getValues();
    const displayB = data.map(r => r[1]); // cột B hiển thị

    const formulas = [];
    const clearData = [];

    // Tạo mảng công thức và mảng xóa dữ liệu
    for (let i = 0; i < data.length; i++) {
      if (displayB[i] !== "") {

        formulas.push([
          `=SUM(G${i + 3}:INDEX(${i + 3}:${i + 3};6+DAY(EOMONTH(DATEVALUE("1/"&$F$2);0))))`
        ]);

        clearData.push(null);

      } else {

        formulas.push([null]);
        clearData.push(Array(lastCol).fill("")); // xóa toàn bộ row

      }
    }

    // Xóa dữ liệu thừa 1 lần
    const rowsToClear = clearData
      .map((r, idx) => r ? r : null)
      .filter(r => r);

    if (rowsToClear.length > 0) {
      sheet.getRange(3, 1, lastRow - 1, lastCol)
           .setValues(
             data.map((row, i) =>
               displayB[i] !== "" ? row : Array(lastCol).fill("")
             )
           );
    }
    
    // Ghi công thức 1 lần
    sheet.getRange(3, 1, sheet.getLastRow(), 5).clearContent();
    sheet.getRange(3, lastCol, formulas.length, 1).setFormulas(formulas);
  });

  // 5️⃣ Cập nhật tổng điểm danh
  tongHopDiemDanh();
}

function showDatePicker() {
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  // chỉ cho phép popup ở sheet ĐIỂM DANH
  if (!sheet.getName().startsWith("ĐIỂM DANH")) {
    SpreadsheetApp.getUi().alert("Vui lòng chuyển sang sheet ĐIỂM DANH");
    return;
  }
  const html = HtmlService.createHtmlOutputFromFile("DatePicker")
    .setWidth(250)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, "Chọn tháng và năm");
}

function setDate(value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet.getName().startsWith("ĐIỂM DANH")) return;
  const cell = sheet.getRange("F2");
  cell.setValue("'" + value);
  cell.setNumberFormat("@");
  
  fillDaysOfMonth(sheet);
}

function fillDaysOfMonth(sheet) {
  const value = sheet.getRange("F2").getValue(); // "MM/YYYY"
  if (!value) return;

  const [month, year] = value.split("/").map(Number);
  if (!month || !year) return;

  const lastDay = new Date(year, month, 0).getDate(); // số ngày trong tháng
  const startCol = 7; // cột G
  const preserveCols = 1; // giữ 1 cột cuối

  // Tạo mảng ngày
  const days = [];
  for (let i = 1; i <= lastDay; i++) {
    if (i < 10)
      days.push(`0${i}/${month}`);
    else
      days.push(`${i}/${month}`);
  }

  // Xác định cột cuối cùng chứa ngày mới
  const lastDayCol = startCol + lastDay - 1;

  // Xác định cột hiện tại trước cột giữ lại
  const lastColBeforePreserve = sheet.getLastColumn() - preserveCols;

  // 1️⃣ Thêm cột nếu cần
  if (lastDayCol > lastColBeforePreserve) {
    sheet.insertColumnsAfter(lastColBeforePreserve, lastDayCol - lastColBeforePreserve);
  }

  // 2️⃣ Xóa cột thừa nếu cần
  if (lastDayCol < lastColBeforePreserve) {
    const numColsToDelete = lastColBeforePreserve - lastDayCol;
    sheet.deleteColumns(lastDayCol + 1, numColsToDelete);
  }

  // 3️⃣ Điền ngày mới vào hàng 1
  sheet.getRange(1, startCol, 1, lastDay).setValues([days]);

  // 4️⃣ Xóa dữ liệu bên dưới hàng ngày (hàng 2 đến hết)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const numCols = lastDayCol - startCol + 1; // số cột ngày thực tế
    const range = sheet.getRange(2, startCol, lastRow - 1, numCols);
    range.clearContent();
    range.setFontColor(null);
  }
}

function tongHopDiemDanh(row = null, col = null, sheet = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const targetSheet = ss.getSheetByName("TỔNG ĐIỂM DANH");
  const infoSheet = ss.getSheetByName("THÔNG TIN HỌC VIÊN");
  if (!targetSheet) return;

  // Lấy dữ liệu thông tin học viên để biết số buổi còn lại
  const infoValues = infoSheet.getRange(2, 1, infoSheet.getLastRow() - 1, 10).getValues();
  const soBuoiMap = {}; // { "Tên học viên": số buổi còn lại }

  infoValues.forEach(row => {
    const ten = row[1];     // Cột B: Tên học viên
    const soBuoi = row[7];  // Cột H: Số buổi còn lại
    if (ten) {
      soBuoiMap[ten] = Number(soBuoi) || 0;
    }
  });
  if (sheet){
    const tenHocVien = sheet.getRange(row, 2).getValue(); 
    // ❌ Nếu hết buổi → bỏ qua buổi điểm danh
    if (soBuoiMap[tenHocVien] < 1) {
      sheet.getRange(row, col).clearContent();
      SpreadsheetApp.getUi().alert(`Khóa học hiện tại của học viên ${tenHocVien} đã hết!`);
      return;
    }
  }

  // Lấy dữ liệu cũ (nếu có) để lấy tổng năm trước
  const existingData = targetSheet.getLastRow() > 1
    ? targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 3).getValues()
    : [];

  const previousMap = {}; // { "Tên học viên": tổng năm trước }
  existingData.forEach(row => {
    const name = row[0];
    const prevTotal = Number(row[2]) || 0; // cột C: tổng năm trước
    if (name) previousMap[name] = prevTotal;
  });

  const summaryMap = {}; // { "Tên học viên": tổng điểm năm nay }

  // Duyệt tất cả sheet "ĐIỂM DANH"
  sheets.forEach(sheet => {
    if (!sheet.getName().startsWith("ĐIỂM DANH")) return;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return; // sheet trống

    const lastCol = data[0].length - 1; // cột cuối (cột Tổng)
    const rows = data.slice(1); // bỏ tiêu đề

    rows.forEach(r => {
      const hocVien = r[1]; // cột B
      const diem = Number(r[lastCol]) || 0;
      if (hocVien) summaryMap[hocVien] = (summaryMap[hocVien] || 0) + diem;
    });
  });

  // Chuẩn bị mảng 2D để ghi ra sheet, **cộng tổng năm trước vào tổng năm nay**
  const allData = Object.entries(summaryMap).map(([hocVien, tongNamNay]) => {
    const prev = previousMap[hocVien] || 0;
    const tongCong = tongNamNay + prev; // tổng = năm nay + năm trước
    return [hocVien, tongCong, prev]; // cột A: tên, B: tổng cộng, C: tổng năm trước
  });

  // Xóa dữ liệu từ hàng 2 xuống (giữ tiêu đề)
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 3).clearContent();
  }

  // Ghi dữ liệu
  if (allData.length > 0) {
    targetSheet.getRange(2, 1, allData.length, 3).setValues(allData);
  }
}

function sumConfigMatches() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); // Sheet đang mở
  const configSheet = ss.getSheetByName("CONFIG");

  // Lấy dữ liệu cấu hình
  const configData = configSheet.getRange("A2:C").getValues().filter(r => r[0] !== "");
  const configMap = {};
  configData.forEach(r => configMap[r[0]] = Number(r[2]) || 0);

  // Giả sử bạn muốn quét từ cột D đến O (tùy chỉnh tại đây)
  const startCol = 4; // D
  const endCol = 26;  // Z
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, startCol, lastRow - 1, endCol - startCol + 1);
  const values = range.getValues();

  // Đọc giá trị hiện tại của cột E
  const currentE = sheet.getRange(2, 5, lastRow - 1, 1).getValues();

  const output = values.map((row, i) => {

    // Nếu cột D = "Piano Gia sư" thì giữ nguyên E
    if (row[0] === "Piano Gia sư") {
      return [currentE[i][0]];
    }

    let sum = 0;
    row.forEach(v => {
      if (v && configMap[v] !== undefined) {
        sum += configMap[v];
      }
    
    // Khớp "Buổi tặng X"
    const match = String(v).match(/^Buổi tặng\s+(\d+)$/i);
    if (match) {
      sum += Number(match[1]);
    }
    });

    return [sum === 0 ? "" : sum];
  });

  // Ghi kết quả vào cột bạn muốn)
  sheet.getRange(2, startCol + 1, output.length, 1).setValues(output);
}

function capNhatTrangThaiHocVien() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const infoSheet = ss.getSheetByName("THÔNG TIN HỌC VIÊN");
  const sheets = ss.getSheets();

  // Đọc danh sách học viên và trạng thái
  const data = infoSheet.getRange("B2:G").getValues(); // B: Tên, C: ?, G: Trạng thái (sửa theo vị trí cột thật)
  const students = data.filter(r => r[0] !== ""); // loại dòng trống

  // Lặp qua tất cả các sheet có tên bắt đầu bằng "ĐIỂM DANH"
  sheets.forEach(sh => {
    const name = sh.getName();
    if (name.startsWith("ĐIỂM DANH")) {
      const names = sh.getRange("B2:B").getValues().map(r => r[0]); // cột chứa tên học viên ở sheet điểm danh
      const lastCol = sh.getLastColumn(); // cột cuối (cột "Tổng")
      const tongValues = sh.getRange(2, lastCol, names.length, 1).getValues(); // lấy giá trị cột tổng
      sh.showRows(1, sh.getMaxRows()); // hiện tất cả trước (để đảm bảo không bị ẩn nhầm)

      // Lấy tháng trong ô F2 (ví dụ: 10, 11, 12)
      const thangTrongSheet = parseInt(sh.getRange("F2").getValue(), 10);
      const thangHienTai = new Date().getMonth() + 1; // tháng hiện tại (1-12)
      const studentMap = new Map();
      students.forEach(r => studentMap.set(r[0], r));

      names.forEach((ten, i) => {
        const info = studentMap.get(ten);
        const tong = tongValues[i][0] || 0; // giá trị tổng (nếu trống thì = 0)
        if (info) {
          const trangThai = (info[5] || "").toString().trim();
          // Điều kiện 1: nghỉ hoặc tạm nghỉ và tổng = 0
          const dieuKienNghi = (trangThai === "Đã nghỉ" || trangThai === "Tạm nghỉ") && tong === 0;

          // Điều kiện 2: đang học, tổng = 0, tháng hiện tại > tháng trong F2
          const dieuKienHetThang = (trangThai === "Đang học") && tong === 0 && thangHienTai > thangTrongSheet;

          if (dieuKienNghi || dieuKienHetThang) {
            sh.hideRows(i + 2); // i+2 vì dòng đầu là tiêu đề
          }
        }
      });
    }
  });
}

function capNhatTKB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const infoSheet = ss.getSheetByName("THÔNG TIN HỌC VIÊN");
  const tkbSheet = ss.getSheetByName("THỜI KHÓA BIỂU");

  if (!infoSheet || !tkbSheet) throw new Error("Không tìm thấy sheet THÔNG TIN HỌC VIÊN hoặc THỜI KHÓA BIỂU.");

  // --- Lấy toàn bộ dữ liệu học viên 1 lần ---
  const infoValues = infoSheet.getRange(2, 2, infoSheet.getLastRow() - 1, 6).getValues(); 
  // Cột B:G (Tên, ?, Khoa, ?, Giờ học, Trạng thái)
  
  // --- Lấy số dòng TKB ---
  const tkbLastRow = tkbSheet.getLastRow();

  // --- Tạo mảng output để ghi tất cả dữ liệu 1 lần ---
  const output = Array(tkbLastRow - 1).fill(null).map(() => Array(7).fill("")); // 7 cột từ T2 -> CN

  const thuMapping = {3: "T2", 4: "T3", 5: "T4", 6: "T5", 7: "T6", 8: "T7", 9: "CN"};

  // --- Xử lý từng dòng TKB ---
  for (let r = 2; r <= tkbLastRow; r++) {
    const gio = getMergedValue(tkbSheet, r, 1); // cột A
    const mon = getMergedValue(tkbSheet, r, 2); // cột B
    if (!gio || !mon) continue;

    for (let col = 3; col <= 9; col++) { // cột T2 -> CN
      const thu = thuMapping[col];
      if (!thu) continue;

      const slotString = thu + " " + gio;

      // Lọc danh sách học viên
      const danhSachHV = infoValues
        .filter(row => {
          const ten = row[0];
          const khoa = (row[2] || "").toLowerCase();
          const gioHoc = (row[4] || "").toString();
          const trangThai = (row[5] || "").toString().trim();

          if (!ten || trangThai === "Đã nghỉ" || trangThai === "Tạm nghỉ") return false;

          let monHV = "";
          if (khoa.includes("guitar")) monHV = "Guitar";
          else if (khoa.includes("piano")) monHV = "Piano";

          if (monHV !== mon) return false;

          return gioHoc.includes(slotString);
        })
        .map(row => "• " + row[0])
        .join("\n");

      output[r - 2][col - 3] = danhSachHV;
    }
  }

  // --- Ghi tất cả dữ liệu 1 lần ---
  tkbSheet.getRange(2, 3, output.length, 7).setValues(output);
}

// --- Hàm lấy giá trị của ô merged ---
function getMergedValue(sheet, row, col) {
  while (row > 1) {
    const v = sheet.getRange(row, col).getValue();
    if (v !== "") return v; // tìm dòng đầu của ô merged
    row--;
  }
  return "";
}

function kiemTraTrungLapTenHocVien(e, editedSheet) {
  const newName = e.value;

  if (!newName) return;

  // Lấy toàn bộ tên
  const names = editedSheet
    .getRange(2, 2, editedSheet.getLastRow()-1, 1)
    .getValues()
    .flat();

  // Đếm số lần xuất hiện
  const count = names.filter(name => name === newName).length;

  if (count > 1) {
    e.range.clearContent();
    SpreadsheetApp.getUi().alert(
      "Tên học viên đã tồn tại: " + newName
    );}
}

function openCalenderPopupForSelectedCell() {
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  const cell = sheet.getActiveCell();

  // chỉ cho phép popup ở cột Giờ học (VD: cột F = 6)
  if (cell.getColumn() !== 6 || sheet.getName() !== "THÔNG TIN HỌC VIÊN") {
    SpreadsheetApp.getUi().alert("Vui lòng chọn ô trong cột Giờ học của học viên trong sheet THÔNG TIN HỌC VIÊN trước khi nhập lịch.");
    return;
  }

  showCalenderPopup(cell);
}

function showCalenderPopup(cell) {
  const template = HtmlService.createTemplateFromFile('CalenderPicker');

  template.row = cell.getRow();
  template.col = cell.getColumn();
  template.currentValue = cell.getValue();   // <-- dữ liệu giờ học đang có

  const html = template.evaluate()
    .setWidth(450)
    .setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(html, "Chọn giờ học");
}

function saveToCellAndUpdateTKB(row, col, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("THÔNG TIN HỌC VIÊN");
  sheet.getRange(row, col).setValue(value);

  // Gọi trực tiếp capNhatTKB() server-side
  capNhatTKB();
}

function openAttendancePopupForSelectedCell() {
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  const cell = sheet.getActiveCell();

  // chỉ cho phép popup ở cột các ô ca học
  if (sheet.getName() !== "THỜI KHÓA BIỂU") {
    SpreadsheetApp.getUi().alert("Vui lòng chuyển sang sheet THỜI KHÓA BIỂU");
    return;
  }
  if (cell.getColumn() < 3 || cell.getColumn() > 9 || cell.getRow() < 2 || cell.getRow() > 15) {
    SpreadsheetApp.getUi().alert("Vui lòng chọn ca học");
    return;
  }

  showAttendancePopup()
}

function showAttendancePopup() {

  const range = SpreadsheetApp.getActiveRange();

  const value = range.getDisplayValue();

  const students = value
    .split('\n')
    .map(s => s.replace(/^•\s*/, '').trim())
    .filter(Boolean);

  const template =
    HtmlService.createTemplateFromFile('AttendancePopup');

  template.attendanceData =
  getAttendanceData();
  template.initialTs =
  JSON.parse(
    PropertiesService
      .getDocumentProperties()
      .getProperty("ACTIVE_TKB_CELL")
  )?.ts || 0;

  const html = template
    .evaluate()
    .setTitle('Điểm danh');

  SpreadsheetApp.getUi().showSidebar(html);
}

function getCurrentSelectionState() {

  const data = PropertiesService
    .getDocumentProperties()
    .getProperty("ACTIVE_TKB_CELL");

  return data
    ? JSON.parse(data)
    : null;
}

const SLOT_MAP = {
  "09:30": "8h",
  "11:00": "9h30",
  "14:30": "13h",
  "16:00": "14h30",
  "17:30": "16h",
  "19:00": "17h30",
  "20:30": "19h"
};

const SUBJECT_TEACHERS = {
  "Piano": "Thúy",
  "Guitar": "Dương"
};

const TEACHER_COLORS = {
  "Thúy": "#000000",
  "Dương": "#0000ff",
  "Duy": "#ff0000"
};

function getTeacherByColor(color) {

  color = (color || "").toLowerCase();

  if (color === "#ff0000") return "Duy";
  if (color === "#0000ff") return "Dương";

  return "Thúy";
}

function getAttendanceContext() {

  const state = JSON.parse(
    PropertiesService
      .getDocumentProperties()
      .getProperty("ACTIVE_TKB_CELL")
  );

  const column = columnToNumber(state.cell);

  const weekdayMap = {
    3: 2, // T2
    4: 3, // T3
    5: 4, // T4
    6: 5, // T5
    7: 6, // T6
    8: 7, // T7
    9: 8  // CN
  };

  const targetWeekday = weekdayMap[column];

  const today = new Date();

  const currentJsDay = today.getDay(); // CN=0

  let currentVNDay =
    currentJsDay === 0 ? 8 : currentJsDay + 1;

  let diff = targetWeekday - currentVNDay;

  const attendanceDate = new Date(today);

  attendanceDate.setDate(
    attendanceDate.getDate() + diff
  );

  const month = attendanceDate.getMonth() + 1;
  const day = attendanceDate.getDate();

  return {
    date: attendanceDate,
    month,
    day,
    state
  };
}

function columnToNumber(a1) {

  const letters = a1.match(/[A-Z]+/)[0];

  let result = 0;

  for (let i = 0; i < letters.length; i++) {
    result = result * 26 +
      letters.charCodeAt(i) - 64;
  }

  return result;
}

function getAttendanceData() {

  const context = getAttendanceContext();

  const attendanceSheet =
    SpreadsheetApp.getActive()
      .getSheetByName(`ĐIỂM DANH T${context.month}`);

  if (!attendanceSheet) return [];

  const dayColumn = 6 + context.day;

  // Danh sách HV trong ô đang chọn
  const names = context.state.value
    .split('\n')
    .map(s => s.replace(/^•\s*/, '').trim())
    .filter(Boolean);

  // Lấy môn học từ cột B của TKB
  const tkbSheet =
    SpreadsheetApp.getActive()
      .getSheetByName("THỜI KHÓA BIỂU");

  const selectedRow =
    Number(
      context.state.cell.match(/\d+/)[0]
    );

  const subject =
    tkbSheet
      .getRange(selectedRow, 2)
      .getDisplayValue()
      .trim()
      .toLowerCase();

  let defaultTeacher = "";

  if (subject.includes("piano")) {
    defaultTeacher = "Thúy";
  }
  else if (subject.includes("guitar")) {
    defaultTeacher = "Dương";
  }

  const lastRow =
    attendanceSheet.getLastRow();

  const sheetStudents =
    attendanceSheet
      .getRange(2, 2, lastRow - 1, 1)
      .getValues()
      .flat();

  const studentMap = {};

  sheetStudents.forEach((name, index) => {
    studentMap[name] = index + 2;
  });

  return names.map(name => {

    const row = studentMap[name];

    if (!row) {

      return {
        student: name,
        present: false,
        teacher: defaultTeacher
      };
    }

    const cell =
      attendanceSheet.getRange(
        row,
        dayColumn
      );

    const value =
      cell.getDisplayValue();

    const color =
      cell.getFontColor();

    let teacher = defaultTeacher;

    // Đã điểm danh -> lấy GV theo màu chữ
    if (value == "1") {
      teacher = getTeacherByColor(color);
    }

    return {
      student: name,
      present: value == "1",
      teacher: teacher
    };

  });

}

function saveAttendance(data) {

  const context = getAttendanceContext();

  const attendanceSheet =
    SpreadsheetApp.getActive()
      .getSheetByName(`ĐIỂM DANH T${context.month}`);

  if (!attendanceSheet) return;

  const dayColumn = 6 + context.day;

  const lastRow =
    attendanceSheet.getLastRow();

  const sheetStudents =
    attendanceSheet
      .getRange(2, 2, lastRow - 1, 1)
      .getValues()
      .flat();

  const studentMap = {};

  sheetStudents.forEach((name, index) => {
    studentMap[name] = index + 2;
  });

  data.forEach(item => {

    const row = studentMap[item.student];

    if (!row) return;

    const cell =
      attendanceSheet.getRange(
        row,
        dayColumn
      );

    if (!item.present) {

      cell.clearContent();
      cell.setFontColor("#000000");

    return;
}

    cell.setValue(1);

    cell.setFontColor(
      TEACHER_COLORS[item.teacher] ||
      "#000000"
    );

  });

}

function getAttendanceHeader() {

  const context = getAttendanceContext();

  const date = context.date;

  const formattedDate =
    Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );

  return {
    date: formattedDate
  };
}

function taoLessonInstance() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const tkbSheet =
    ss.getSheetByName("THỜI KHÓA BIỂU");

  const insSheet =
    ss.getSheetByName("LESSON_INSTANCE");


  if (!tkbSheet || !insSheet) return;



  const values =
    tkbSheet.getDataRange().getValues();



  // ============================
  // Cache dữ liệu cũ
  // ============================

  const statusCache = {};


  if (insSheet.getLastRow() > 1) {

    const oldData =
      insSheet.getRange(
        2,
        1,
        insSheet.getLastRow()-1,
        8
      ).getValues();



    oldData.forEach(r=>{

      const id = r[0];

      if(id){

        statusCache[id] = {

          status: r[6],
          source: r[7]

        };

      }

    });

  }




  // ============================
  // Từ hôm nay -> 30 ngày
  // ============================

  const start = new Date();

  start.setHours(0,0,0,0);


  const end = new Date(start);

  end.setDate(
    end.getDate()+29
  );




  const thuMapping = {

    3:"T2",
    4:"T3",
    5:"T4",
    6:"T5",
    7:"T6",
    8:"T7",
    9:"CN"

  };



  let output=[];



  let currentGio="";
  let currentMon="";



  for(let r=2;r<=values.length;r++){


    // A merge
    const gio =
      getMergedValue(
        tkbSheet,
        r,
        1
      );


    // B không merge
    const mon =
      values[r-1][1];



    if(gio)
      currentGio = gio;


    if(mon)
      currentMon = mon;




    for(let c=3;c<=9;c++){


      const hvText =
        values[r-1][c-1];



      if(!hvText)
        continue;



      const thu =
        thuMapping[c];



      const hocviens =
        hvText
        .toString()
        .split("\n")
        .map(x =>
          x.replace("•","").trim()
        )
        .filter(Boolean);




      const dates =
        getDatesByWeekday(
          start,
          end,
          thu
        );




      dates.forEach(date=>{


        const dateKey =
          Utilities.formatDate(
            date,
            "GMT+7",
            "yyyyMMdd"
          );



        hocviens.forEach(hv=>{


          /*
            InstanceID:
            ngày + thứ + giờ + môn + học viên
          */

          const id =
            dateKey
            +"_"+thu
            +"_"+currentGio
            +"_"+currentMon
            +"_"+hv;



          const cached =
            statusCache[id];



          let status = "ACTIVE";
          let source = "TKB";


          // Nếu đã có phiếu
          if(cached){

            status =
              cached.status;

            source =
              cached.source;

          }



          output.push([

            id,

            date,

            thu,

            currentGio,

            currentMon,

            hv,

            status,

            source

          ]);



        });



      });



    }


  }




  // ============================
  // Ghi lại
  // ============================


  insSheet.clearContents();



  insSheet
  .getRange(
    1,
    1,
    1,
    8
  )
  .setValues([[

    "InstanceID",
    "Ngày",
    "Thứ",
    "Giờ",
    "Môn",
    "Học viên",
    "Schedule_Status",
    "Nguồn"

  ]]);



  if(output.length){


    insSheet
    .getRange(
      2,
      1,
      output.length,
      8
    )
    .setValues(output);


  }


}

function getDatesByWeekday(start,end,thu){

 const map={
  "CN":0,
  "T2":1,
  "T3":2,
  "T4":3,
  "T5":4,
  "T6":5,
  "T7":6
 };


 let arr=[];

 let d=new Date(start);


 while(d<=end){

  if(d.getDay()==map[thu]){
    arr.push(new Date(d));
  }

  d.setDate(
    d.getDate()+1
  );

 }

 return arr;

}

function hienThiLichHomNayNgayMai() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const tkbSheet =
    ss.getSheetByName("THỜI KHÓA BIỂU");

  const insSheet =
    ss.getSheetByName("LESSON_INSTANCE");


  if (!tkbSheet || !insSheet) return;


  // ====================
  // tạo cột Hôm nay / Mai
  // ====================

  const startCol = 11; // J

  tkbSheet
    .getRange(1,startCol)
    .setValue("HÔM NAY");

  tkbSheet
    .getRange(1,startCol+1)
    .setValue("NGÀY MAI");



  const today =
    new Date();

  today.setHours(0,0,0,0);


  const tomorrow =
    new Date(today);

  tomorrow.setDate(
    today.getDate()+1
  );



  const data =
    insSheet
      .getRange(
        2,
        1,
        insSheet.getLastRow()-1,
        8
      )
      .getValues();



  const result = {
    today:{},
    tomorrow:{}
  };



  data.forEach(r=>{


    const [
      id,
      date,
      thu,
      gio,
      mon,
      hv,
      status
    ] = r;


    // chỉ lấy lịch thực tế
    if(status !== "ACTIVE") return;


    const d =
      new Date(date);

    d.setHours(0,0,0,0);



    let target = null;


    if(
      d.getTime() === today.getTime()
    ){
      target="today";
    }


    if(
      d.getTime() === tomorrow.getTime()
    ){
      target="tomorrow";
    }


    if(!target) return;



    const key =
      gio+"|"+mon;



    if(!result[target][key]){
      result[target][key]=[];
    }


    result[target][key]
      .push("• "+hv);


  });



  const lastRow =
    tkbSheet.getLastRow();



  // ====================
  // đi theo dòng TKB hiện tại
  // ====================

  for(let r=2;r<=lastRow;r++){


    const gio =
      getMergedValue(
        tkbSheet,
        r,
        1
      );


    const mon =
      tkbSheet
        .getRange(r,2)
        .getValue();


    if(!gio || !mon) continue;



    const key =
      gio+"|"+mon;



    tkbSheet
      .getRange(r,startCol)
      .setValue(
        result.today[key]
        ? result.today[key].join("\n")
        : ""
      );


    tkbSheet
      .getRange(r,startCol+1)
      .setValue(
        result.tomorrow[key]
        ? result.tomorrow[key].join("\n")
        : ""
      );

  }



  // copy format giống cột CN
  tkbSheet
    .getRange(1,9,lastRow,1)
    .copyTo(
      tkbSheet.getRange(1,11,lastRow,2),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );

}

function updateLessonStatus(instanceId, status, source = "PHIEU") {

  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("LESSON_INSTANCE");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  for (let i = 0; i < data.length; i++) {

    if (data[i][0] === instanceId) {

      sheet.getRange(i + 2, 7).setValue(status);
      sheet.getRange(i + 2, 8).setValue(source);
      hienThiLichHomNayNgayMai();
      return true;
    }
  }

  return false;
}

function doPost(e) {
  try {

    const body = JSON.parse(e.postData.contents || "{}");

    if (body.action === "processPhieu") {

      const result = processPhieu(body);

      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: "Unknown action"
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {

    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: err.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


function processPhieu(payload) {

  const ss = SpreadsheetApp.openById("1goQl0iqKg7gnKDw1TKlBmRRJVscp02j9DyGiKnP9LOA");

  const phieuSheet = ss.getSheetByName("PHIẾU LỊCH HỌC");
  const insSheet = ss.getSheetByName("LESSON_INSTANCE");

  if (!payload?.phieu_id) {
    return { success: false, message: "Missing phieu_id" };
  }

  const phieuData = phieuSheet.getDataRange().getValues();
  const insData = insSheet.getDataRange().getValues();
  const col = getHeaderMap(insSheet);

  const rows = phieuData.slice(1);
  const statusCol = 9; // I

  const phieuIndex = rows.findIndex(r => r[0] == payload.phieu_id);

  if (phieuIndex === -1) {
    return { success: false, message: "Không tìm thấy phiếu" };
  }

  const p = rows[phieuIndex];

  if (p[8] !== "MỚI") {
    return { success: false, message: "Phiếu không hợp lệ hoặc đã xử lý" };
  }

  const hv = p[2];
  const loai = p[3];
  const ngayCu = new Date(p[4]);
  const ngayMoi = p[5] ? new Date(p[5]) : null;
  const gioMoi = p[6] ? String(p[6]).trim() : "";

  let updated = false;

  for (let i = 1; i < insData.length; i++) {

    const row = insData[i];
    if (row[5] !== hv) continue;

    if (!sameDate(new Date(row[1]), ngayCu)) continue;

    const instanceId = row[0];

    if (loai === "OFF") {
      updated = updateLessonStatus(instanceId, "OFF", "PHIEU") || updated;
    }

    if (loai === "ĐỔI LỊCH" && ngayMoi) {

  // OFF buổi cũ
  updated = updateLessonStatus(instanceId, "OFF", "PHIEU") || updated;

  // Clone đúng dòng đang duyệt
  const newRow = [...row];

  const newDate = new Date(ngayMoi);

  newRow[col["Ngày"]] = newDate;

  const weekdayMap = [
    "CN",
    "T2",
    "T3",
    "T4",
    "T5",
    "T6",
    "T7"
  ];

  newRow[col["Thứ"]] = weekdayMap[newDate.getDay()];

  const dateKey = Utilities.formatDate(
    newDate,
    Session.getScriptTimeZone(),
    "yyyyMMdd"
  );

  newRow[col["Giờ"]] = gioMoi;
  const subject = newRow[col["Môn"]];
  const student = newRow[col["Học viên"]];

  newRow[col["InstanceID"]] =
    `${dateKey}_${weekdayMap[newDate.getDay()]}_${gioMoi}_${subject}_${student}`;

  newRow[col["Schedule_Status"]] = "ACTIVE";
  newRow[col["Nguồn"]] = "PHIEU";

  insSheet.appendRow(newRow);
    }
  }

  const resultStatus = updated ? "ĐÃ XỬ LÝ" : "LỖI";

  phieuSheet.getRange(phieuIndex + 2, statusCol).setValue(resultStatus);
  hienThiLichHomNayNgayMai();
  return {
    success: updated,
    phieu_id: payload.phieu_id,
    message: resultStatus
  };
}


// helper
function sameDate(d1, d2) {
  return d1 && d2 &&
    d1.getFullYear() === d2.getFullYear() &&
    d1.getMonth() === d2.getMonth() &&
    d1.getDate() === d2.getDate();
}

function toggleAutoAttendance() {

  const prop = PropertiesService.getScriptProperties();

  const enabled =
    prop.getProperty("AUTO_ATTENDANCE") === "true";

  if (enabled) {

    removeAttendanceTriggers();

    prop.setProperty("AUTO_ATTENDANCE", "false");

  } else {

    createAttendanceTriggers();

    prop.setProperty("AUTO_ATTENDANCE", "true");

  }

  // Cập nhật lại menu
  onOpen();

  SpreadsheetApp.getUi().alert(
    enabled
      ? "Đã tắt điểm danh tự động."
      : "Đã bật điểm danh tự động."
  );

}

function removeAttendanceTriggers() {

  const handlers = [
    "scheduleTodayAttendance",
    "autoAttendance"
  ];

  ScriptApp.getProjectTriggers().forEach(trigger => {

    if (handlers.includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }

  });

}

function createAttendanceTriggers() {

  removeAttendanceTriggers();

  ScriptApp.newTrigger("scheduleTodayAttendance")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .nearMinute(5)
    .create();

  // Tạo luôn trigger của hôm nay
  scheduleTodayAttendance();

}

function scheduleTodayAttendance() {

  // Xóa toàn bộ trigger autoAttendance cũ
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "autoAttendance") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  const today = new Date();

  const slots = [
    { hour: 8, minute: 0 },
    { hour: 9, minute: 30 },
    { hour: 13, minute: 0 },
    { hour: 14, minute: 30 },
    { hour: 17, minute: 30 },
    { hour: 19, minute: 0 }
  ];

  slots.forEach(slot => {

    const triggerTime = new Date(today);

    triggerTime.setHours(
      slot.hour,
      slot.minute + 90,
      0,
      0
    );

    if (triggerTime <= new Date()) return;

    ScriptApp.newTrigger("autoAttendance")
      .timeBased()
      .at(triggerTime)
      .create();

  });
  hienThiLichHomNayNgayMai();

}

function autoAttendance(testDate = null) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const lessonSheet = ss.getSheetByName("LESSON_INSTANCE");
  if (!lessonSheet) return;

  const col = getHeaderMap(lessonSheet);

  //--------------------------------------------------
  // Xác định ca học từ thời điểm trigger
  //--------------------------------------------------

  const now = testDate || new Date();

  const timeKey = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "HH:mm"
  );

  const targetSlot = SLOT_MAP[timeKey];

  if (!targetSlot) {
    Logger.log("Không xác định được ca học.");
    return;
  }

  //--------------------------------------------------
  // Sheet điểm danh tháng hiện tại
  //--------------------------------------------------

  const month = now.getMonth() + 1;

  const attendanceSheet =
    ss.getSheetByName(`ĐIỂM DANH T${month}`);

  if (!attendanceSheet) return;

  const dayColumn = 6 + now.getDate();

  //--------------------------------------------------
  // Map học viên -> dòng
  //--------------------------------------------------

  const lastRow = attendanceSheet.getLastRow();

  const names = attendanceSheet
    .getRange(2, 2, lastRow - 1, 1)
    .getValues()
    .flat();

  const rowMap = {};

  names.forEach((name, i) => {
    if (name) rowMap[name] = i + 2;
  });

  //--------------------------------------------------
  // Ngày hôm nay
  //--------------------------------------------------

  const todayKey = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "yyyyMMdd"
  );

  //--------------------------------------------------
  // Đọc LESSON_INSTANCE
  //--------------------------------------------------

  const values = lessonSheet.getDataRange().getValues();

  let count = 0;

  for (let i = 1; i < values.length; i++) {

    const row = values[i];

    const lessonDate = Utilities.formatDate(
      new Date(row[col["Ngày"]]),
      Session.getScriptTimeZone(),
      "yyyyMMdd"
    );

    if (lessonDate !== todayKey) continue;

    if (row[col["Giờ"]] !== targetSlot) continue;

    if (row[col["Schedule_Status"]] !== "ACTIVE") continue;

    const student = row[col["Học viên"]];

    const subject = row[col["Môn"]];

    const teacher = SUBJECT_TEACHERS[String(subject).toLowerCase()] || SUBJECT_TEACHERS.piano;

    const attendanceRow = rowMap[student];

    if (!attendanceRow) continue;

    const cell = attendanceSheet.getRange(
      attendanceRow,
      dayColumn
    );

    // Giáo viên đã điểm danh trước thì bỏ qua
    if (cell.getDisplayValue() == "1") continue;

    cell
      .setValue(1)
      .setFontColor(
        TEACHER_COLORS[teacher] || "#000000"
      );

    count++;

  }

  Logger.log(
    `AutoAttendance: ${targetSlot} - ${count} học viên`
  );

}

function getHeaderMap(sheet) {

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  const map = {};

  headers.forEach((header, index) => {
    map[header] = index;
  });

  return map;

}

// function test16h() {

//   autoAttendance(
//     new Date(2026, 5, 30, 17, 30)
//   );

// }












