function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const enabled =
    PropertiesService.getScriptProperties()
      .getProperty("AUTO_ATTENDANCE") === "true";
  ui.createMenu("📅 Công cụ điểm danh")
    .addItem("Chọn tháng/năm", "showDatePicker")
    .addItem("Cập nhật thông tin học viên", "updateInfoData")
    .addItem("Cập nhật TKB", "capNhatTKB")
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
    const data = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
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
      sheet.getRange(3, 1, lastRow - 2, lastCol)
           .setValues(
             data.map((row, i) =>
               displayB[i] !== "" ? row : Array(lastCol).fill("")
             )
           );
    }
    
    // Ghi công thức 1 lần
    sheet.getRange(3, 1, sheet.getLastRow(), 5).clearContent();   // 👈 XÓA CỘT A:E
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

  //--------------------------------------------------
  // Map Mã HV -> Số buổi
  //--------------------------------------------------

  const infoValues = infoSheet
    .getRange(2, 1, infoSheet.getLastRow() - 1, 10)
    .getValues();

  const soBuoiMap = {};

  infoValues.forEach(r => {

    const ma = r[0];
    const soBuoi = Number(r[7]) || 0;

    if (ma) {
      soBuoiMap[ma] = soBuoi;
    }

  });

  //--------------------------------------------------
  // Kiểm tra còn buổi học
  //--------------------------------------------------

  if (sheet) {

    const maHV = sheet.getRange(row, 1).getValue();
    const tenHocVien = sheet.getRange(row, 2).getValue();

    if ((soBuoiMap[maHV] || 0) < 1) {

      sheet.getRange(row, col).clearContent();

      SpreadsheetApp.getUi().alert(
        `Khóa học hiện tại của học viên ${tenHocVien} đã hết!`
      );

      return;

    }

  }

  //--------------------------------------------------
  // Đọc dữ liệu cũ
  //--------------------------------------------------

  const existingData =
    targetSheet.getLastRow() > 1
      ? targetSheet
          .getRange(
            2,
            1,
            targetSheet.getLastRow() - 1,
            4
          )
          .getValues()
      : [];

  const previousMap = {};

  existingData.forEach(r => {

    const maHV = r[0];
    const tongNamTruoc = Number(r[3]) || 0;

    if (maHV) {
      previousMap[maHV] = tongNamTruoc;
    }

  });

  //--------------------------------------------------
  // Tổng hợp năm nay
  //--------------------------------------------------

  const summaryMap = {};
  const tenMap = {};

  sheets.forEach(sheet => {

    if (!sheet.getName().startsWith("ĐIỂM DANH"))
      return;

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1)
      return;

    const lastCol = data[0].length - 1;

    data.slice(1).forEach(r => {

      const maHV = r[0];
      const ten = r[1];
      const diem = Number(r[lastCol]) || 0;

      if (!maHV)
        return;

      tenMap[maHV] = ten;

      summaryMap[maHV] =
        (summaryMap[maHV] || 0) + diem;

    });

  });

  //--------------------------------------------------
  // Chuẩn bị dữ liệu ghi
  //--------------------------------------------------

  const allData = Object.entries(summaryMap).map(
    ([maHV, tongNamNay]) => {

      const prev =
        previousMap[maHV] || 0;

      return [

        maHV,
        tenMap[maHV],
        tongNamNay + prev,
        prev

      ];

    });

  //--------------------------------------------------
  // Ghi dữ liệu
  //--------------------------------------------------

  if (targetSheet.getLastRow() > 1) {

    targetSheet
      .getRange(
        2,
        1,
        targetSheet.getLastRow() - 1,
        4
      )
      .clearContent();

  }

  if (allData.length) {

    targetSheet
      .getRange(
        2,
        1,
        allData.length,
        4
      )
      .setValues(allData);

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

  // Đọc danh sách học viên: A=Mã HV, B=Tên, ..., G=Trạng thái
  const data = infoSheet.getRange("A2:G").getValues();
  const students = data.filter(r => r[0] !== ""); // loại dòng trống (theo Mã HV)

  const studentMap = new Map();
  students.forEach(r => studentMap.set(String(r[0]).trim(), r));

  // Lặp qua tất cả các sheet có tên bắt đầu bằng "ĐIỂM DANH"
  sheets.forEach(sh => {
    const name = sh.getName();
    if (!name.startsWith("ĐIỂM DANH")) return;

    const ids = sh.getRange("A2:A")
      .getValues()
      .map(r => String(r[0]).trim()); // cột Mã HV ở sheet điểm danh

    const lastCol = sh.getLastColumn(); // cột cuối (cột "Tổng")
    const tongValues = sh.getRange(2, lastCol, ids.length, 1).getValues();

    sh.showRows(1, sh.getMaxRows()); // hiện tất cả trước

    const thangTrongSheet = parseInt(sh.getRange("F2").getValue(), 10);
    const thangHienTai = new Date().getMonth() + 1;

    ids.forEach((maHV, i) => {

      if (!maHV) return;

      const info = studentMap.get(maHV);
      const tong = tongValues[i][0] || 0;

      if (info) {

        const trangThai = (info[6] || "").toString().trim(); // G = Trạng thái

        const dieuKienNghi =
          (trangThai === "Đã nghỉ" || trangThai === "Tạm nghỉ") && tong === 0;

        const dieuKienHetThang =
          (trangThai === "Đang học") && tong === 0 && thangHienTai > thangTrongSheet;

        if (dieuKienNghi || dieuKienHetThang) {
          sh.hideRows(i + 2);
        }
      }
    });
  });

  taoLessonInstance();
  hienThiLichHomNayNgayMai();
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

  taoLessonInstance();
  hienThiLichHomNayNgayMai();
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
  if (cell.getColumn() < 3 || cell.getColumn() > 9 || cell.getRow() < 2) {
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

function getTargetSlot(now) {

  const SLOT_MAP = [
    { trigger: "09:30", slot: "8h" },
    { trigger: "11:00", slot: "9h30" },
    { trigger: "14:30", slot: "13h" },
    { trigger: "16:00", slot: "14h30" },
    { trigger: "17:30", slot: "16h" },
    { trigger: "19:00", slot: "17h30" },
    { trigger: "20:30", slot: "19h" }
  ];

  const current = now.getHours() * 60 + now.getMinutes();

  for (const s of SLOT_MAP) {

    const [h, m] = s.trigger.split(":").map(Number);
    const trigger = h * 60 + m;

    if (Math.abs(current - trigger) <= 20) {
      return s.slot;
    }

  }

  return null;

}

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

  // Danh sách HV trong ô đang chọn (TKB vẫn hiển thị theo Tên)
  const names = context.state.value
    .split('\n')
    .map(s => s.replace(/^•\s*/, '').trim())
    .filter(Boolean);

  //--------------------------------------------------
  // Map Tên -> Mã HV (từ THÔNG TIN HỌC VIÊN)
  //--------------------------------------------------

  const infoSheet =
    SpreadsheetApp.getActive()
      .getSheetByName("THÔNG TIN HỌC VIÊN");

  const infoData = infoSheet
    .getRange(2, 1, infoSheet.getLastRow() - 1, 2)
    .getValues();

  const maHVMap = {};

  infoData.forEach(r => {
    const ma = String(r[0]).trim();
    const ten = String(r[1]).trim();
    if (ten) maHVMap[ten] = ma;
  });

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

  //--------------------------------------------------
  // Map Mã HV -> dòng trong sheet điểm danh (cột A)
  //--------------------------------------------------

  const lastRow =
    attendanceSheet.getLastRow();

  const sheetIds =
    attendanceSheet
      .getRange(2, 1, lastRow - 1, 1)
      .getValues()
      .flat();

  const rowMap = {};

  sheetIds.forEach((id, index) => {
    if (id) rowMap[String(id).trim()] = index + 2;
  });

  return names.map(name => {

    const maHV = maHVMap[name];

    const row = maHV ? rowMap[maHV] : null;

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

  //--------------------------------------------------
  // Map Tên -> Mã HV (từ THÔNG TIN HỌC VIÊN)
  //--------------------------------------------------

  const infoSheet =
    SpreadsheetApp.getActive()
      .getSheetByName("THÔNG TIN HỌC VIÊN");

  const infoData = infoSheet
    .getRange(2, 1, infoSheet.getLastRow() - 1, 2)
    .getValues();

  const maHVMap = {};

  infoData.forEach(r => {
    const ma = String(r[0]).trim();
    const ten = String(r[1]).trim();
    if (ten) maHVMap[ten] = ma;
  });

  //--------------------------------------------------
  // Map Mã HV -> dòng trong sheet điểm danh (cột A)
  //--------------------------------------------------

  const lastRow =
    attendanceSheet.getLastRow();

  const sheetIds =
    attendanceSheet
      .getRange(2, 1, lastRow - 1, 1)
      .getValues()
      .flat();

  const rowMap = {};

  sheetIds.forEach((id, index) => {
    if (id) rowMap[String(id).trim()] = index + 2;
  });

  //--------------------------------------------------
  // Ghi dữ liệu điểm danh
  //--------------------------------------------------

  data.forEach(item => {

    const maHV = maHVMap[item.student];

    if (!maHV) return;

    const row = rowMap[maHV];

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

  tongHopDiemDanh();

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

  const tkbSheet = ss.getSheetByName("THỜI KHÓA BIỂU");
  const insSheet = ss.getSheetByName("LESSON_INSTANCE");
  const infoSheet = ss.getSheetByName("THÔNG TIN HỌC VIÊN");

  if (!tkbSheet || !insSheet || !infoSheet) return;

  const values = tkbSheet.getDataRange().getValues();

  //--------------------------------------------------
  // Map Tên -> Mã HV
  //--------------------------------------------------

  const infoData = infoSheet
    .getRange(2, 1, infoSheet.getLastRow() - 1, 2)
    .getValues();

  const maHVMap = {};

  infoData.forEach(r => {

    const ma = String(r[0]).trim();
    const ten = String(r[1]).trim();

    if (ten) maHVMap[ten] = ma;

  });

  //--------------------------------------------------
  // Cache dữ liệu cũ
  //--------------------------------------------------

  const statusCache = {};
  const customInstances = [];

  if (insSheet.getLastRow() > 1) {

    const oldData = insSheet
      .getRange(2, 1, insSheet.getLastRow() - 1, 9)
      .getValues();
  
    const today = new Date();
    today.setHours(0,0,0,0);

    
    oldData.forEach(r => {

      const id = r[0];

      if (!id) return;

      statusCache[id] = {
        status: r[7],
        source: r[8]
      };
      const lessonDate = new Date(r[1]);
      if (r[8] !== "TKB" && lessonDate >= today) {
        customInstances.push(r);
      }

    });

  }

  //--------------------------------------------------
  // Khoảng ngày tạo
  //--------------------------------------------------

  const start = new Date();
  start.setHours(0, 0, 0, 0);

  const end = new Date(start);
  end.setDate(end.getDate() + 29);

  const thuMapping = {
    3: "T2",
    4: "T3",
    5: "T4",
    6: "T5",
    7: "T6",
    8: "T7",
    9: "CN"
  };

  let output = [];
  let currentGio = "";
  let currentMon = "";

  //--------------------------------------------------
  // Sinh Lesson Instance
  //--------------------------------------------------

  for (let r = 2; r <= values.length; r++) {

    const gio = getMergedValue(tkbSheet, r, 1);
    const mon = values[r - 1][1];

    if (gio) currentGio = gio;
    if (mon) currentMon = mon;

    for (let c = 3; c <= 9; c++) {

      const hvText = values[r - 1][c - 1];

      if (!hvText) continue;

      const thu = thuMapping[c];

      const hocviens = hvText
        .toString()
        .split("\n")
        .map(x => x.replace("•", "").trim())
        .filter(Boolean);

      const dates = getDatesByWeekday(start, end, thu);

      dates.forEach(date => {

        const dateKey = Utilities.formatDate(
          date,
          Session.getScriptTimeZone(),
          "yyyyMMdd"
        );

        hocviens.forEach(tenHV => {

          const maHV = maHVMap[tenHV];

          if (!maHV) {
            Logger.log("Không tìm thấy Mã HV: " + tenHV);
            return;
          }

          const id =
            `${dateKey}_${thu}_${currentGio}_${currentMon}_${maHV}`;

          const cached = statusCache[id];

          let status = "ACTIVE";
          let source = "TKB";

          if (cached) {
            status = cached.status;
            source = cached.source;
          }

          output.push([
            id,
            date,
            thu,
            currentGio,
            currentMon,
            maHV,
            tenHV,
            status,
            source
          ]);

        });

      });

    }

  }

  //--------------------------------------------------
  // Giữ các instance tạo bởi phiếu
  //--------------------------------------------------

  const outputIds = new Set(output.map(r => r[0]));

  customInstances.forEach(r => {

    if (!outputIds.has(r[0])) {
      output.push(r);
    }

  });

  //--------------------------------------------------
  // Ghi dữ liệu
  //--------------------------------------------------

  insSheet.clearContents();

  insSheet.getRange(1, 1, 1, 9).setValues([[
    "InstanceID",
    "Ngày",
    "Thứ",
    "Giờ",
    "Môn",
    "Mã HV",
    "Học viên",
    "Schedule_Status",
    "Nguồn"
  ]]);

  if (output.length) {

    insSheet
      .getRange(2, 1, output.length, 9)
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
  const tkbSheet = ss.getSheetByName("THỜI KHÓA BIỂU");
  const insSheet = ss.getSheetByName("LESSON_INSTANCE");

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
        9
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
      maHV,
      hv,
      status,
      nguon
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

function updateLessonStatus(instanceId, status, source = "TKB") {

  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("LESSON_INSTANCE");

  const col = getHeaderMap(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    if (data[i][col["InstanceID"]] !== instanceId)
      continue;

    sheet.getRange(i + 1, col["Schedule_Status"] + 1)
      .setValue(status);

    sheet.getRange(i + 1, col["Nguồn"] + 1)
      .setValue(source);

    return true;
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
      if (body.action === "huyPhieu") {

      const result = huyPhieu(body);

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
    return {
      success: false,
      message: "Missing phieu_id"
    };
  }

  const phieuData = phieuSheet.getDataRange().getValues();
  const insData = insSheet.getDataRange().getValues();

  const phieuCol = getHeaderMap(phieuSheet);
  const instanceCol = getHeaderMap(insSheet);

  const rows = phieuData.slice(1);
  const statusCol = phieuCol["Trạng thái"] + 1;

  const phieuIndex = rows.findIndex(
    r => r[phieuCol["Phiếu ID"]] == payload.phieu_id
  );

  if (phieuIndex == -1) {
    return {
      success: false,
      message: "Không tìm thấy phiếu"
    };
  }

  const p = rows[phieuIndex];

  if (p[phieuCol["Trạng thái"]] != "MỚI") {
    return {
      success: false,
      message: "Phiếu đã xử lý"
    };
  }

  const instanceId = p[phieuCol["Buổi học"]];
  const maHV = String(p[phieuCol["Mã HV"]] || "").trim();

  const maHVList = String(p[phieuCol["Danh sách mã HV"]] || "")
    .split(" , ")
    .map(s => s.trim())
    .filter(Boolean);

  const loai = p[phieuCol["Loại phiếu"]];
  const phamVi = p[phieuCol["Phạm vi"]];
  const ngayCu = new Date(p[phieuCol["Ngày học"]]);
  const gioHoc = String(p[phieuCol["Giờ học"]] || "").trim();
  const ngayMoi = p[phieuCol["Ngày mới"]]
    ? new Date(p[phieuCol["Ngày mới"]])
    : null;
  const gioMoi = String(p[phieuCol["Giờ mới"]] || "").trim();
  const monHoc = p[phieuCol["Môn"]];
  const toiNgay = p[phieuCol["Tới ngày"]]
    ? new Date(p[phieuCol["Tới ngày"]])
    : null;

  let updated = false;

  //--------------------------------------------------
  // OFF CÁ NHÂN / ĐỔI LỊCH
  //--------------------------------------------------

  let lessonRow = null;

  if (phamVi == "CÁ NHÂN") {

    // Có chọn buổi học -> lấy đúng instance
    if (instanceId) {

      const lessonIndex = insData.findIndex(
        r => r[instanceCol["InstanceID"]] == instanceId
      );

      if (lessonIndex == -1) {

        phieuSheet
          .getRange(phieuIndex + 2, statusCol)
          .setValue("LỖI");

        return {
          success: false,
          message: "Không tìm thấy buổi học"
        };

      }

      lessonRow = insData[lessonIndex];

    }

    // Không chọn buổi học -> THÊM BUỔI
    else {

      const lessonIndex = insData.findIndex(r =>
        r[instanceCol["Mã HV"]] == maHV &&
        r[instanceCol["Schedule_Status"]] == "ACTIVE"
      );

      if (lessonIndex == -1) {

        phieuSheet
          .getRange(phieuIndex + 2, statusCol)
          .setValue("LỖI");

        return {
          success: false,
          message: "Không tìm thấy lịch của học viên"
        };

      }

      lessonRow = insData[lessonIndex];

    }

  }

  //--------------------------------------------------
  // OFF
  //--------------------------------------------------

  if (loai == "OFF") {

    //---------------- OFF CÁ NHÂN ----------------

    if (phamVi == "CÁ NHÂN") {

      // OFF 1 BUỔI
      if (!toiNgay) {

        updated =
          updateLessonStatus(
            lessonRow[instanceCol["InstanceID"]],
            "OFF",
            payload.phieu_id
          ) || updated;

      }

      // OFF DÀI
      else {

        const studentMaHV = lessonRow[instanceCol["Mã HV"]];
        const startDate = new Date(
          lessonRow[instanceCol["Ngày"]]
        );

        for (let i = 1; i < insData.length; i++) {

          const row = insData[i];

          if (row[instanceCol["Mã HV"]] != studentMaHV)
            continue;

          if (row[instanceCol["Môn"]] != monHoc)
            continue;

          if (row[instanceCol["Schedule_Status"]] != "ACTIVE")
            continue;

          const lessonDate =
            new Date(row[instanceCol["Ngày"]]);

          if (lessonDate < startDate)
            continue;

          if (lessonDate > toiNgay)
            continue;

          updated =
            updateLessonStatus(
              row[instanceCol["InstanceID"]],
              "OFF",
              payload.phieu_id
            ) || updated;

        }

      }

    }

    //---------------- OFF CA HỌC ----------------

    else if (phamVi == "CA HỌC") {

      for (let i = 1; i < insData.length; i++) {

        const row = insData[i];

        if (!sameDate(
          new Date(row[instanceCol["Ngày"]]),
          ngayCu
        ))
          continue;

        if (row[instanceCol["Giờ"]] != gioHoc)
          continue;
        if (row[instanceCol["Môn"]] != monHoc)
          continue;

        if (row[instanceCol["Schedule_Status"]] != "ACTIVE")
          continue;

        if (maHVList.length && !maHVList.includes(String(row[instanceCol["Mã HV"]]).trim())) {
          continue;
        }

        updated =
          updateLessonStatus(
            row[instanceCol["InstanceID"]],
            "OFF",
            payload.phieu_id
          ) || updated;

      }

    }

  }

  //--------------------------------------------------
  // ĐỔI LỊCH
  //--------------------------------------------------

  if (loai == "ĐỔI LỊCH" && ngayMoi && gioMoi) {

    if (instanceId) {

      updateLessonStatus(
        instanceId,
        "OFF",
        payload.phieu_id
      );

    }

    const newRow = [...lessonRow];

    const newDate = new Date(ngayMoi);

    const weekdayMap = [
      "CN",
      "T2",
      "T3",
      "T4",
      "T5",
      "T6",
      "T7"
    ];

    newRow[instanceCol["Ngày"]] = newDate;

    newRow[instanceCol["Thứ"]] =
      weekdayMap[newDate.getDay()];

    newRow[instanceCol["Giờ"]] =
      gioMoi;

    const dateKey = Utilities.formatDate(
      newDate,
      Session.getScriptTimeZone(),
      "yyyyMMdd"
    );

    newRow[instanceCol["InstanceID"]] =
      `${dateKey}_${weekdayMap[newDate.getDay()]}_${gioMoi}_${newRow[instanceCol["Môn"]]}_${newRow[instanceCol["Mã HV"]]}`;

    newRow[instanceCol["Schedule_Status"]] =
      "ACTIVE";

    newRow[instanceCol["Nguồn"]] =
      payload.phieu_id;

    insSheet.appendRow(newRow);
    updated = true;
  }

  //--------------------------------------------------

  const resultStatus =
    updated ? "ĐÃ XỬ LÝ" : "LỖI";

  phieuSheet
    .getRange(phieuIndex + 2, statusCol)
    .setValue(resultStatus);

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
    "autoAttendanceBySlot",
    "autoAttendanceDaily",
    "hienThiLichHomNayNgayMai"
  ];

  ScriptApp.getProjectTriggers().forEach(trigger => {

    if (handlers.includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }

  });

}

function createAttendanceTriggers() {

  removeAttendanceTriggers();

  // Điểm danh cuối ngày
  ScriptApp.newTrigger("autoAttendanceDaily")
    .timeBased()
    .everyDays(1)
    .atHour(22)
    .nearMinute(0)
    .create();

  // Refresh lịch Hôm nay/Ngày mai đầu ngày
  ScriptApp.newTrigger("hienThiLichHomNayNgayMai")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .nearMinute(5)
    .create();

  // Refresh ngay lần đầu khi bật tính năng
  hienThiLichHomNayNgayMai();

}

function createAttendanceTriggersBySlot() {

  removeAttendanceTriggers();

  ScriptApp.newTrigger("scheduleTodayAttendance")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .nearMinute(5)
    .create();

  scheduleTodayAttendance();

}

function scheduleTodayAttendance() {

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "autoAttendanceBySlot") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  const today = new Date();

  const slots = [
    { hour: 8, minute: 0 },
    { hour: 9, minute: 30 },
    { hour: 13, minute: 0 },
    { hour: 14, minute: 30 },
    { hour: 16, minute: 0 },
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

    ScriptApp.newTrigger("autoAttendanceBySlot")
      .timeBased()
      .at(triggerTime)
      .create();

  });

  hienThiLichHomNayNgayMai();

}

function autoAttendanceBySlot() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const lessonSheet = ss.getSheetByName("LESSON_INSTANCE");
  if (!lessonSheet) return;

  const col = getHeaderMap(lessonSheet);

  //--------------------------------------------------
  // Xác định ca học
  //--------------------------------------------------

  const now = new Date();

  const targetSlot = getTargetSlot(now);

  if (!targetSlot) {
    Logger.log("Không nằm trong khoảng trigger.");
    return;
  }

  //--------------------------------------------------
  // Sheet điểm danh
  //--------------------------------------------------

  const month = now.getMonth() + 1;

  const attendanceSheet =
    ss.getSheetByName(`ĐIỂM DANH T${month}`);

  if (!attendanceSheet) return;

  const dayColumn = 6 + now.getDate();

  //--------------------------------------------------
  // Map Mã HV -> dòng
  //--------------------------------------------------

  const lastRow = attendanceSheet.getLastRow();

  const ids = attendanceSheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat();

  const rowMap = {};

  ids.forEach((id, i) => {
    if (id) rowMap[String(id).trim()] = i + 2;
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

    const maHV = String(row[col["Mã HV"]]).trim();

    const subject = row[col["Môn"]];

    const teacher =
      SUBJECT_TEACHERS[String(subject)]
      || SUBJECT_TEACHERS.Piano;

    const attendanceRow = rowMap[maHV];

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

  tongHopDiemDanh();

}

function autoAttendanceDaily() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const lessonSheet = ss.getSheetByName("LESSON_INSTANCE");
  if (!lessonSheet) return;

  const col = getHeaderMap(lessonSheet);

  const now = new Date();

  const month = now.getMonth() + 1;

  const attendanceSheet =
    ss.getSheetByName(`ĐIỂM DANH T${month}`);

  if (!attendanceSheet) return;

  const dayColumn = 6 + now.getDate();

  const lastRow = attendanceSheet.getLastRow();
  if (lastRow < 2) return;

  const ids = attendanceSheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat();

  const rowMap = {};

  ids.forEach((id, i) => {
    if (id) rowMap[String(id).trim()] = i + 2;
  });

  const todayKey = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "yyyyMMdd"
  );

  if (lessonSheet.getLastRow() < 2) return;

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
    if (row[col["Schedule_Status"]] !== "ACTIVE") continue;

    const maHV = String(row[col["Mã HV"]]).trim();
    const subject = row[col["Môn"]];

    const teacher =
      SUBJECT_TEACHERS[String(subject)]
      || SUBJECT_TEACHERS.Piano;

    const attendanceRow = rowMap[maHV];

    if (!attendanceRow) continue;

    const cell = attendanceSheet.getRange(
      attendanceRow,
      dayColumn
    );

    if (cell.getDisplayValue() == "1") continue;

    cell
      .setValue(1)
      .setFontColor(
        TEACHER_COLORS[teacher] || "#000000"
      );

    count++;

  }

  Logger.log(
    `AutoAttendance cuối ngày: ${count} học viên`
  );

  tongHopDiemDanh();

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

function huyPhieu(payload) {

  const ss = SpreadsheetApp.openById("1goQl0iqKg7gnKDw1TKlBmRRJVscp02j9DyGiKnP9LOA");

  const lessonSheet = ss.getSheetByName("LESSON_INSTANCE");
  const phieuSheet = ss.getSheetByName("PHIẾU LỊCH HỌC");

  if (!payload?.phieu_id) {
    return {
      success: false,
      message: "Missing phieu_id"
    };
  }

  const lessonCol = getHeaderMap(lessonSheet);
  const phieuCol = getHeaderMap(phieuSheet);

  const lessonData = lessonSheet.getDataRange().getValues();
  const phieuData = phieuSheet.getDataRange().getValues();

  //--------------------------------------------------
  // Tìm phiếu
  //--------------------------------------------------

  let phieuRow = -1;
  let loai = "";
  let p;
  for (let i = 1; i < phieuData.length; i++) {

    if (phieuData[i][phieuCol["Phiếu ID"]] == payload.phieu_id) {

      phieuRow = i + 1;
      p = phieuData[i]
      loai = p[phieuCol["Loại phiếu"]];
      break;

    }
  }

  if (phieuRow === -1) {
    return {
      success: false,
      message: "Không tìm thấy phiếu"
    };
  }
  
  //--------------------------------------------------
  // ĐỔI LỊCH
  // Xóa instance mới
  // ACTIVE lại instance cũ
  //--------------------------------------------------

  if (loai === "ĐỔI LỊCH") {

    for (let i = lessonData.length - 1; i >= 1; i--) {

      if (lessonData[i][lessonCol["Nguồn"]] != payload.phieu_id)
        continue;

      const oldInstanceId = p[phieuCol["Buổi học"]];
      const currentInstanceId = lessonData[i][lessonCol["InstanceID"]];

      if (currentInstanceId === oldInstanceId && lessonData[i][lessonCol["Schedule_Status"]] === "OFF") {
        lessonSheet.getRange(
          i + 1,
          lessonCol["Schedule_Status"] + 1
        ).setValue("ACTIVE");

        lessonSheet.getRange(
          i + 1,
          lessonCol["Nguồn"] + 1
        ).setValue("TKB");
      } else {
        lessonSheet.deleteRow(i + 1);
      }
    }
  }

  //--------------------------------------------------
  // OFF
  //--------------------------------------------------

  else {
    for (let i = lessonData.length - 1; i >= 1; i--) {

      if (lessonData[i][lessonCol["Nguồn"]] != payload.phieu_id)
        continue;

      lessonSheet.getRange(
        i + 1,
        lessonCol["Schedule_Status"] + 1
      ).setValue("ACTIVE");

      lessonSheet.getRange(
        i + 1,
        lessonCol["Nguồn"] + 1
      ).setValue("TKB");

    }
  }

  //--------------------------------------------------
  // Cập nhật trạng thái phiếu
  //--------------------------------------------------

  phieuSheet.getRange(
    phieuRow,
    phieuCol["Trạng thái"] + 1
  ).setValue("ĐÃ HỦY");

  hienThiLichHomNayNgayMai();

  return {
    success: true,
    message: "Đã hủy phiếu"
  };

}












