var SHEET_ID = '1g3TaxSKJEbLmPIDLsX_32_FcdR8xBl6arEFXvPGMq6Y';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Lịch Phân Ca Tài Xế Van')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function getDriversList() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheetTX = ss.getSheetByName('danhsachtaixe');
  var dataTX = sheetTX.getRange(2, 1, sheetTX.getLastRow() - 1, 2).getValues();
  return dataTX.map(function(r) { return { name: r[0].toString().trim(), id: r[1].toString().trim() }; });
}

function checkAndGetShiftData(driverId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheetLich = ss.getSheetByName('lichphanca');
  var lastRow = sheetLich.getLastRow();
  var dataLich = lastRow > 1 ? sheetLich.getRange(2, 1, lastRow - 1, 6).getValues() : [];
  
  var isAlreadyRegistered = false;
  var shiftCounts = {};
  
  // Lấy mốc ngày mai để so sánh (Chỉ khóa nếu đã đăng ký lịch cho tương lai)
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0,0,0,0);

  for (var i = 0; i < dataLich.length; i++) {
    var rowId = dataLich[i][2].toString().trim(); 
    var rawDate = dataLich[i][4];                 
    var shift = dataLich[i][5] ? dataLich[i][5].toString().trim() : ""; 
    
    var dateStr = "";
    var rowDateObj = null;

    if (rawDate instanceof Date) {
       dateStr = Utilities.formatDate(rawDate, "GMT+7", "dd/MM/yyyy");
       rowDateObj = new Date(rawDate);
    } else {
       dateStr = rawDate.toString().trim();
       var parts = dateStr.split('/');
       if(parts.length === 3) {
         dateStr = ("0" + parts[0]).slice(-2) + "/" + ("0" + parts[1]).slice(-2) + "/" + parts[2];
         rowDateObj = new Date(parts[2], parts[1] - 1, parts[0]);
       }
    }

    // Nếu ID khớp và lịch đã đăng ký là từ ngày mai trở đi -> Khóa tài khoản
    if (rowId === driverId.toString().trim() && rowDateObj && rowDateObj >= tomorrow) {
       isAlreadyRegistered = true; 
    }

    if (!shiftCounts[dateStr]) shiftCounts[dateStr] = { '06:00-15:00': 0, '08:00-17:00': 0, '13:00-22:00': 0 };
    if (shiftCounts[dateStr][shift] !== undefined) {
        shiftCounts[dateStr][shift]++;
    }
  }

  return { isAlreadyRegistered: isAlreadyRegistered, shiftCounts: shiftCounts };
}

function submitFinalSchedule(data) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('lichphanca');
  var timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
  
  data.schedules.forEach(function(item) {
    var finalArea = (item.shift === 'OFF') ? 'Nghỉ' : item.area;
    sheet.appendRow([
      timestamp,       
      data.driverName, 
      data.driverId,   
      finalArea,       
      item.date,       
      item.shift       
    ]);
  });
  return "SUCCESS";
}