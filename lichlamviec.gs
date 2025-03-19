// ======== FUNCTIONS QUẢN LÝ LỊCH LÀM VIỆC ========

// Lấy dữ liệu đăng ký lịch làm việc từ Sheet
// Lấy dữ liệu đăng ký lịch làm việc từ Sheet
function getScheduleData(startDate, endDate, sheetName = 'Kỹ thuật') {
  try {
    // ID của Google Sheet chứa dữ liệu đăng ký lịch làm việc
    const SCHEDULE_SHEET_ID = '1XUedszSYOzuW-jcZPPGCqUUwzuqdpuMFSl6gwWBQEEs';
    const spreadsheet = SpreadsheetApp.openById(SCHEDULE_SHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log(`Không tìm thấy sheet "${sheetName}" trong file đăng ký lịch`);
      return {
        success: false,
        message: `Không tìm thấy sheet "${sheetName}" trong file đăng ký lịch`,
        data: []
      };
    }
    
    // Lấy tất cả dữ liệu từ sheet
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // Ghi log header cho debug
    Logger.log("Headers sheet " + sheetName + ": " + values[0].join(", "));
    
    // Bỏ qua dòng tiêu đề (dòng 1)
    const data = values.slice(1).map((row, index) => {
      const rowDate = row[2]; // Ngày đăng ký ở cột C (index 2)
      let formattedDate = '';
      
      // Kiểm tra và định dạng ngày
      if (rowDate instanceof Date && !isNaN(rowDate)) {
        const day = String(rowDate.getDate()).padStart(2, '0');
        const month = String(rowDate.getMonth() + 1).padStart(2, '0');
        const year = rowDate.getFullYear();
        formattedDate = `${day}/${month}/${year}`;
      } else if (typeof rowDate === 'string') {
        formattedDate = rowDate;
      }
      
      // Trả về đúng cấu trúc dữ liệu theo các cột của sheet
      return {
        rowIndex: index + 2, // +2 vì dòng 1 là tiêu đề và index bắt đầu từ 0
        staffName: row[0] || '', // Tên nhân viên ở cột A
        staffId: row[1] || '',   // ID ở cột B
        registerDate: formattedDate, // Ngày đăng ký ở cột C
        shift: row[3] || '',     // Ca làm việc ở cột D
        status: row[4] || '',    // Trạng thái ở cột E
        notes: row[5] || ''      // Ghi chú ở cột F
      };
    });
    
    // Lọc những ca "Off"
    let filteredData = data.filter(item => 
      !(item.shift === 'Off' || 
        (typeof item.shift === 'string' && item.shift.toLowerCase().includes('off')))
    );
    
    Logger.log(`Đã lọc được ${filteredData.length} bản ghi sau khi loại bỏ ca "Off"`);
    
    // Nếu có ngày bắt đầu và kết thúc, lọc dữ liệu theo khoảng thời gian
    if (startDate && endDate) {
      const startDateObj = parseDate(startDate);
      const endDateObj = parseDate(endDate);
      
      if (startDateObj && endDateObj) {
        filteredData = filteredData.filter(item => {
          const itemDate = parseDate(item.registerDate);
          return itemDate && itemDate >= startDateObj && itemDate <= endDateObj;
        });
      }
    }
    
    Logger.log(`Đã tìm thấy ${filteredData.length} bản ghi lịch làm việc cho ${sheetName}`);
    
    return {
      success: true,
      message: `Tải thành công ${filteredData.length} bản ghi`,
      data: filteredData
    };
  } catch (error) {
    Logger.log('Lỗi trong getScheduleData: ' + error.toString());
    return {
      success: false,
      message: 'Lỗi khi tải dữ liệu lịch làm việc: ' + error.toString(),
      data: []
    };
  }
}


// Hàm lấy danh sách nhân viên từ sheet nhân viên
// Sửa function getStaffListFromSheet() trong file lichlamviec.gs.txt
function getStaffListFromSheet(type = 'technician') {
  try {
    // ID của Google Sheet chứa danh sách nhân viên
    const STAFF_SHEET_ID = '11MArx0UJ2YkHsyGVKmN3b8sg_Lwgl-k44NXmwkbGBJU';
    const spreadsheet = SpreadsheetApp.openById(STAFF_SHEET_ID);
    const sheet = spreadsheet.getSheetByName('Nhân Viên');
    
    if (!sheet) {
      Logger.log('Không tìm thấy sheet "Nhân Viên"');
      // Trả về dữ liệu mẫu nếu không tìm thấy sheet
      return {
        success: false,
        message: 'Không tìm thấy sheet "Nhân Viên"',
        data: type === 'technician' ? ["Nhân viên A", "Nhân viên B", "Nhân viên C"] : ["Lễ tân A", "Lễ tân B", "Lễ tân C"]
      };
    }
    
    // Lấy dữ liệu từ cột phù hợp với loại nhân viên
    // Kỹ thuật: cột B (B2:B20), Lễ tân: cột A (A2:A20)
    const column = type === 'technician' ? 'B' : 'A';
    const dataRange = sheet.getRange(`${column}2:${column}20`);
    const values = dataRange.getValues();
    
    // Lọc các giá trị không rỗng
    const staffList = values
      .flat()
      .filter(name => name && typeof name === 'string' && name.trim() !== '');
    
    Logger.log(`Đã tìm thấy ${staffList.length} ${type === 'technician' ? 'nhân viên kỹ thuật' : 'nhân viên lễ tân'} từ sheet`);
    
    return {
      success: true,
      message: `Tải thành công ${staffList.length} ${type === 'technician' ? 'nhân viên kỹ thuật' : 'nhân viên lễ tân'}`,
      data: staffList
    };
  } catch (error) {
    Logger.log(`Lỗi trong getStaffListFromSheet (${type}): ` + error.toString());
    // Trả về dữ liệu mẫu trong trường hợp lỗi
    return {
      success: false,
      message: `Lỗi khi tải danh sách ${type === 'technician' ? 'nhân viên kỹ thuật' : 'nhân viên lễ tân'}: ` + error.toString(),
      data: type === 'technician' ? ["Nhân viên A", "Nhân viên B", "Nhân viên C"] : ["Lễ tân A", "Lễ tân B", "Lễ tân C"]
    };
  }
}

// Thêm function mới để gọi lấy danh sách nhân viên lễ tân
function getReceptionStaffList() {
  return getStaffListFromSheet('reception');
}

// Hàm lưu thông tin phân công lịch làm việc
function saveScheduleAssignment(assignmentData, sheetName = 'Kỹ thuật') {
  try {
    const ASSIGNMENT_SHEET_ID = '1tWsNM9nU0Vd6RXv0ygq7CwZSATArLIn9crpjrisweqM';
    const spreadsheet = SpreadsheetApp.openById(ASSIGNMENT_SHEET_ID);
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.getRange('A1:G1').setValues([
        ['Tên nhân viên', 'Người phân công', 'Ngày làm việc', 'Ca làm việc', 'Phân công bởi', 'Ghi chú', 'Cửa hàng']
      ]);
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
    }
    
    let validEntries = 0;
    
    // Lấy toàn bộ dữ liệu hiện tại
    const existingData = sheet.getDataRange().getValues().slice(1); // Bỏ qua dòng tiêu đề
    
    assignmentData.forEach(entry => {
      if (entry.staffName && entry.registerDate) {
        // Kiểm tra xem dữ liệu đã tồn tại chưa
        const isDuplicate = existingData.some(row => 
          row[0] === entry.staffName && 
          row[2] === entry.registerDate && 
          row[3] === entry.shift && 
          row[6] === entry.store
        );
        
        // Chỉ thêm nếu chưa tồn tại
        if (!isDuplicate) {
          const lastRow = sheet.getLastRow();
          sheet.getRange(lastRow + 1, 1, 1, 7).setValues([
            [
              entry.staffName, 
              entry.assignedBy || 'System', 
              entry.registerDate, 
              entry.shift, 
              entry.assignedBy || 'System', 
              entry.notes || '', 
              entry.store || ''
            ]
          ]);
          validEntries++;
        }
      }
    });
    
    Logger.log(`Đã lưu ${validEntries} phân công lịch làm việc vào sheet ${sheetName}`);
    
    return {
      success: true,
      message: `Đã lưu ${validEntries} phân công lịch làm việc thành công`,
      savedCount: validEntries
    };
  } catch (error) {
    Logger.log('Lỗi trong saveScheduleAssignment: ' + error.toString());
    return {
      success: false,
      message: 'Lỗi khi lưu phân công lịch làm việc: ' + error.toString(),
      savedCount: 0
    };
  }
}

// Hàm hỗ trợ tìm dòng theo uniqueKey
function findRowByUniqueKey(sheet, uniqueKey) {
  try {
    const lastRow = sheet.getLastRow();
    
    // Tìm kiếm uniqueKey trong toàn bộ sheet
    for (let i = 2; i <= lastRow; i++) {
      // So sánh kết hợp nhiều trường để xác định duy nhất
      const staffName = sheet.getRange(i, 1).getValue();
      const registerDate = sheet.getRange(i, 3).getValue();
      const shift = sheet.getRange(i, 4).getValue();
      const store = sheet.getRange(i, 7).getValue();
      
      const currentUniqueKey = `${staffName}-${registerDate}-${shift}-${store}`;
      
      if (currentUniqueKey === uniqueKey) {
        return i;
      }
    }
    
    return null; // Không tìm thấy
  } catch (error) {
    Logger.log('Lỗi khi tìm dòng theo uniqueKey: ' + error.toString());
    return null;
  }
}
// Hàm hỗ trợ chuyển đổi chuỗi ngày thành đối tượng Date
function parseDate(dateString) {
  if (!dateString) return null;
  
  try {
    // Hỗ trợ định dạng dd/mm/yyyy
    const parts = dateString.split('/');
    if (parts.length === 3) {
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1; // Tháng từ 0-11 trong JavaScript
      const year = parseInt(parts[2], 10);
      
      // Kiểm tra tính hợp lệ của ngày và đảm bảo định dạng đúng là dd/mm/yyyy
      if (!isNaN(day) && !isNaN(month) && !isNaN(year) && day <= 31 && month < 12) {
        // Log đầy đủ thông tin để debug
        const createdDate = new Date(year, month, day);
        Logger.log(`parseDate: ${dateString} -> ${createdDate} (timestamp: ${createdDate.getTime()})`);
        return createdDate;
      }
      
      // Thử định dạng mm/dd/yyyy nếu không phải dd/mm/yyyy
      const possibleMonth = parseInt(parts[0], 10) - 1;
      const possibleDay = parseInt(parts[1], 10);
      
      if (!isNaN(possibleDay) && !isNaN(possibleMonth) && !isNaN(year) && 
          possibleDay <= 31 && possibleMonth < 12) {
        Logger.log(`Thử đảo ngày/tháng: ${dateString} -> ${possibleDay}/${possibleMonth+1}/${year}`);
        const alternateDate = new Date(year, possibleMonth, possibleDay);
        
        // Ghi rõ thông tin để debug
        Logger.log(`parseDate (định dạng thay thế): ${dateString} -> ${alternateDate} (timestamp: ${alternateDate.getTime()})`);
        return alternateDate;
      }
    }
  } catch (e) {
    Logger.log('Lỗi parseDate: ' + e);
  }
  
  Logger.log(`Không thể parse ngày: ${dateString}`);
  return null;
}

// Hàm định dạng đối tượng Date thành chuỗi ngày dd/mm/yyyy
function formatDate(date) {
  if (!date || !(date instanceof Date) || isNaN(date)) return '';
  
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  
  return `${day}/${month}/${year}`;
}

// Hàm lấy ngày đầu tiên và cuối cùng của tuần
function getWeekRange(dateString) {
  try {
    // Nếu không có ngày, sử dụng ngày hiện tại
    let currentDate;
    if (!dateString) {
      currentDate = new Date();
    } else {
      currentDate = parseDate(dateString);
      if (!currentDate) currentDate = new Date();
    }
    
    // Lấy ngày trong tuần (0 = Chủ nhật, 1 = Thứ 2, ..., 6 = Thứ 7)
    const dayOfWeek = currentDate.getDay();
    
    // Tính ngày thứ Hai của tuần
    const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek; // Nếu là Chủ nhật, quay lại 6 ngày, ngược lại tính từ thứ 2
    const firstDay = new Date(currentDate);
    firstDay.setDate(currentDate.getDate() + mondayOffset);
    firstDay.setHours(0, 0, 0, 0);
    
    // Tính ngày Chủ nhật của tuần (thêm 6 ngày từ thứ 2)
    const lastDay = new Date(firstDay);
    lastDay.setDate(firstDay.getDate() + 6);
    lastDay.setHours(23, 59, 59, 999);
    
    const result = {
      startDate: formatDate(firstDay),
      endDate: formatDate(lastDay)
    };
    
    Logger.log(`Tính tuần cho ngày ${dateString}: ${result.startDate} - ${result.endDate}`);
    return result;
  } catch (error) {
    Logger.log('Lỗi trong getWeekRange: ' + error.toString());
    // Trả về tuần hiện tại trong trường hợp lỗi
    const today = new Date();
    return {
      startDate: formatDate(today),
      endDate: formatDate(today)
    };
  }
}
// Hàm kiểm tra các sheet cần thiết
function checkScheduleSheets() {
  try {
    // Định nghĩa IDs của các sheet
    const SCHEDULE_SHEET_ID = '1XUedszSYOzuW-jcZPPGCqUUwzuqdpuMFSl6gwWBQEEs';
    const STAFF_SHEET_ID = '11MArx0UJ2YkHsyGVKmN3b8sg_Lwgl-k44NXmwkbGBJU';
    const ASSIGNMENT_SHEET_ID = '1tWsNM9nU0Vd6RXv0ygq7CwZSATArLIn9crpjrisweqM';
    
    const result = {
      scheduleSheet: false,
      staffSheet: false,
      assignmentSheet: false,
      scheduleSheetId: SCHEDULE_SHEET_ID,
      staffSheetId: STAFF_SHEET_ID,
      assignmentSheetId: ASSIGNMENT_SHEET_ID
    };
    
    // Kiểm tra sheet đăng ký lịch
    try {
      const scheduleSpreadsheet = SpreadsheetApp.openById(SCHEDULE_SHEET_ID);
      const scheduleSheet = scheduleSpreadsheet.getSheetByName('Kỹ thuật');
      result.scheduleSheet = scheduleSheet != null;
    } catch (e) {
      Logger.log('Không thể mở sheet đăng ký lịch: ' + e.toString());
    }
    
    // Kiểm tra sheet nhân viên
    try {
      const staffSpreadsheet = SpreadsheetApp.openById(STAFF_SHEET_ID);
      const staffSheet = staffSpreadsheet.getSheetByName('Nhân Viên');
      result.staffSheet = staffSheet != null;
    } catch (e) {
      Logger.log('Không thể mở sheet nhân viên: ' + e.toString());
    }
    
    // Kiểm tra sheet phân công
    try {
      const assignmentSpreadsheet = SpreadsheetApp.openById(ASSIGNMENT_SHEET_ID);
      const assignmentSheet = assignmentSpreadsheet.getSheetByName('Kỹ thuật');
      result.assignmentSheet = assignmentSheet != null;
    } catch (e) {
      Logger.log('Không thể mở sheet phân công: ' + e.toString());
    }
    
    return result;
  } catch (error) {
    Logger.log('Lỗi trong checkScheduleSheets: ' + error.toString());
    return {
      success: false,
      message: 'Lỗi khi kiểm tra sheets: ' + error.toString()
    };
  }
}

// Hàm tạo dữ liệu mẫu cho trường hợp không kết nối được
function getTestScheduleData() {
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  
  const dayAfterTomorrow = new Date(today);
  dayAfterTomorrow.setDate(today.getDate() + 2);
  
  return {
    success: true,
    message: 'Dữ liệu mẫu được tạo thành công',
    data: [
      {
        rowIndex: 2,
        staffName: 'Nhân viên A',
        registerDate: formatDate(today),
        shift: 'Sáng',
        status: 'Đăng ký',
        notes: 'Dữ liệu mẫu'
      },
      {
        rowIndex: 3,
        staffName: 'Nhân viên B',
        registerDate: formatDate(tomorrow),
        shift: 'Chiều',
        status: 'Đăng ký',
        notes: 'Dữ liệu mẫu'
      },
      {
        rowIndex: 4,
        staffName: 'Nhân viên C',
        registerDate: formatDate(dayAfterTomorrow),
        shift: 'Tối',
        status: 'Đăng ký',
        notes: 'Dữ liệu mẫu'
      }
    ]
  };
}
// Hàm tạo dữ liệu mẫu cho lịch làm việc lễ tân
function getTestReceptionData() {
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  
  const dayAfterTomorrow = new Date(today);
  dayAfterTomorrow.setDate(today.getDate() + 2);
  
  return {
    success: true,
    message: 'Dữ liệu mẫu cho lễ tân được tạo thành công',
    data: [
      {
        rowIndex: 2,
        staffName: 'Lễ tân A',
        registerDate: formatDate(today),
        shift: 'Sáng',
        status: 'Đăng ký',
        notes: 'Dữ liệu mẫu',
        store: 'Ba Đình'
      },
      {
        rowIndex: 3,
        staffName: 'Lễ tân B',
        registerDate: formatDate(today),
        shift: 'Chiều',
        status: 'Đăng ký',
        notes: 'Dữ liệu mẫu',
        store: 'Ba Đình'
      },
      {
        rowIndex: 4,
        staffName: 'Lễ tân C',
        registerDate: formatDate(tomorrow),
        shift: 'Sáng',
        status: 'Đăng ký',
        notes: 'Dữ liệu mẫu',
        store: 'Phan Thanh'
      },
      {
        rowIndex: 5,
        staffName: 'Lễ tân D',
        registerDate: formatDate(tomorrow),
        shift: 'Tối',
        status: 'Đăng ký',
        notes: 'Dữ liệu mẫu',
        store: 'Phan Thanh'
      },
      {
        rowIndex: 6,
        staffName: 'Lễ tân E',
        registerDate: formatDate(dayAfterTomorrow),
        shift: 'Chiều',
        status: 'Đăng ký',
        notes: 'Dữ liệu mẫu',
        store: 'Ba Đình'
      }
    ]
  };
}

function getAssignedScheduleData(startDate, endDate, sheetName = 'Lễ tân') {
  try {
    // Log khoảng thời gian yêu cầu để debug
    Logger.log(`Đang lấy dữ liệu từ ${startDate} đến ${endDate}`);
    
    // ĐÃ SỬA: Đảm bảo luôn lấy dữ liệu từ đầu tuần (thứ 2)
    let effectiveStartDate = startDate;
    if (startDate) {
      // Phân tích chuỗi ngày
      const parts = startDate.split('/');
      if (parts.length === 3) {
        // Tạo đối tượng Date từ chuỗi ngày
        const startDateObj = new Date(
          parseInt(parts[2]), // năm
          parseInt(parts[1]) - 1, // tháng (0-11)
          parseInt(parts[0]) // ngày
        );
        
        // Lấy thứ trong tuần (0 = Chủ nhật, 1 = Thứ 2, ..., 6 = Thứ 7)
        const dayOfWeek = startDateObj.getDay();
        
        // Nếu không phải thứ 2, tính toán lại ngày thứ 2 của tuần này
        if (dayOfWeek !== 1) {
          // Tính offset để lấy ngày thứ 2 (đặc biệt xử lý trường hợp Chủ nhật)
          const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
          startDateObj.setDate(startDateObj.getDate() + mondayOffset);
          
          // Định dạng lại ngày thứ 2 thành dd/mm/yyyy
          const day = String(startDateObj.getDate()).padStart(2, '0');
          const month = String(startDateObj.getMonth() + 1).padStart(2, '0');
          const year = startDateObj.getFullYear();
          effectiveStartDate = `${day}/${month}/${year}`;
          
          Logger.log(`Đã điều chỉnh khoảng thời gian để bắt đầu từ thứ 2: ${effectiveStartDate}`);
        }
      }
    }
    
    // ID của Google Sheet chứa dữ liệu phân công đã lưu
    const ASSIGNMENT_SHEET_ID = '1tWsNM9nU0Vd6RXv0ygq7CwZSATArLIn9crpjrisweqM';
    const spreadsheet = SpreadsheetApp.openById(ASSIGNMENT_SHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log(`Không tìm thấy sheet "${sheetName}" trong file chia lịch`);
      return {
        success: false,
        message: `Không tìm thấy sheet "${sheetName}" trong file chia lịch`,
        data: []
      };
    }
    
    // Lấy tất cả dữ liệu từ sheet
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // Bỏ qua dòng tiêu đề (dòng 1)
    const data = values.slice(1).map((row, index) => {
      const rowDate = row[2]; // Ngày đăng ký ở cột C (index 2)
      const shift = row[3] || ''; // Ca làm việc ở cột D
      
      // Kiểm tra nếu ca làm có chứa "Off" hoặc trống, bỏ qua dòng này
      if (!shift || shift.toLowerCase().includes('off')) {
        return null; // Bỏ qua dòng này
      }
      
      let formattedDate = '';
      
      // Kiểm tra và định dạng ngày
      if (rowDate instanceof Date && !isNaN(rowDate)) {
        const day = String(rowDate.getDate()).padStart(2, '0');
        const month = String(rowDate.getMonth() + 1).padStart(2, '0');
        const year = rowDate.getFullYear();
        formattedDate = `${day}/${month}/${year}`;
      } else if (typeof rowDate === 'string') {
        formattedDate = rowDate;
      }
      
      return {
        rowIndex: index + 2, // +2 vì dòng 1 là tiêu đề và index bắt đầu từ 0
        staffName: row[0] || '', // Tên nhân viên ở cột A
        registerDate: formattedDate, // Ngày đăng ký đã định dạng
        shift: shift, // Ca làm việc ở cột D
        status: row[4] || '', // Trạng thái ở cột E
        notes: row[5] || '', // Ghi chú ở cột F
        store: row[6] || '' // Cửa hàng ở cột G (chỉ áp dụng cho sheet Lễ tân)
      };
    }).filter(item => item !== null); // Lọc bỏ các mục null (Off hoặc trống)
    
    // Tiền xử lý dữ liệu để chuẩn hóa ngày tháng
    data.forEach(item => {
      if (item.registerDate) {
        // Kiểm tra xem có phải định dạng MM/DD/YYYY không
        const parts = item.registerDate.split('/');
        if (parts.length === 3) {
          const firstPart = parseInt(parts[0], 10);
          const secondPart = parseInt(parts[1], 10);
          
          // Nếu phần đầu là tháng (1-12) và phần hai là ngày (1-31)
          if (firstPart <= 12 && secondPart <= 31 && firstPart < secondPart) {
            // Khả năng cao là định dạng bị đảo từ MM/DD/YYYY thành DD/MM/YYYY
            const correctedDate = `${secondPart.toString().padStart(2, '0')}/${firstPart.toString().padStart(2, '0')}/${parts[2]}`;
            Logger.log(`Đã phát hiện và sửa định dạng ngày: ${item.registerDate} -> ${correctedDate}`);
            item.registerDate = correctedDate;
          }
        }
      }
    });
    
    // Nếu có ngày bắt đầu và kết thúc, lọc dữ liệu theo khoảng thời gian
    let filteredData = data;
    if (startDate && endDate) {
      const startDateObj = parseDate(effectiveStartDate);
      const endDateObj = parseDate(endDate);
      
      if (startDateObj && endDateObj) {
        // ĐÃ SỬA: Log để debug việc lọc dữ liệu
        Logger.log(`Lọc dữ liệu từ ${effectiveStartDate} (${startDateObj.getTime()}) đến ${endDate} (${endDateObj.getTime()})`);
        
        filteredData = data.filter(item => {
          // Thử cả hai định dạng cho registerDate (dd/mm/yyyy và mm/dd/yyyy)
          let itemDate = parseDate(item.registerDate);
          
          // Nếu không parse được, thử đảo ngày và tháng
          if (!itemDate && item.registerDate) {
            const parts = item.registerDate.split('/');
            if (parts.length === 3) {
              const alternativeDate = `${parts[1]}/${parts[0]}/${parts[2]}`;
              itemDate = parseDate(alternativeDate);
              if (itemDate) {
                Logger.log(`Đã parse thành công sau khi đảo ngày/tháng: ${item.registerDate} -> ${alternativeDate}`);
                // Cập nhật lại định dạng ngày chính xác để hiển thị đúng
                item.registerDate = `${parts[1]}/${parts[0]}/${parts[2]}`;
              }
            }
          }
          
          const result = itemDate && itemDate >= startDateObj && itemDate <= endDateObj;
          
          // ĐÃ THÊM: Debug log cho việc lọc từng ngày
          if (itemDate) {
            Logger.log(`Kiểm tra ngày ${item.registerDate} (${itemDate.getTime()}): ${result ? "Phù hợp" : "Không phù hợp"}`);
          } else {
            Logger.log(`Không thể parse ngày: ${item.registerDate}`);
          }
          
          return result;
        });
      }
    }
    
    Logger.log(`Đã tìm thấy ${filteredData.length} bản ghi lịch đã phân công cho ${sheetName}`);
    
    return {
      success: true,
      message: `Tải thành công ${filteredData.length} bản ghi đã phân công`,
      data: filteredData
    };
  } catch (error) {
    Logger.log('Lỗi trong getAssignedScheduleData: ' + error.toString());
    return {
      success: false,
      message: 'Lỗi khi tải dữ liệu lịch đã phân công: ' + error.toString(),
      data: []
    };
  }
}
// Hàm riêng cho tab Kỹ thuật
function getScheduleDataForTechnician(startDate, endDate) {
  try {
    // ID của Google Sheet chứa dữ liệu đăng ký lịch làm việc
    const SCHEDULE_SHEET_ID = '1XUedszSYOzuW-jcZPPGCqUUwzuqdpuMFSl6gwWBQEEs';
    const spreadsheet = SpreadsheetApp.openById(SCHEDULE_SHEET_ID);
    const sheet = spreadsheet.getSheetByName('Kỹ thuật');
    
    if (!sheet) {
      Logger.log(`Không tìm thấy sheet "Kỹ thuật" trong file đăng ký lịch`);
      return {
        success: false,
        message: `Không tìm thấy sheet "Kỹ thuật" trong file đăng ký lịch`,
        data: []
      };
    }
    
    // Lấy tất cả dữ liệu từ sheet
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // Ghi log header để debug
    Logger.log("Headers sheet Kỹ thuật: " + values[0].join(", "));
    
    // Bỏ qua dòng tiêu đề (dòng 1)
    const data = values.slice(1).map((row, index) => {
      // Lấy ngày từ cột C (index 2)
      const rowDate = row[2]; 
      let formattedDate = '';
      
      // Kiểm tra và định dạng ngày
      if (rowDate instanceof Date && !isNaN(rowDate)) {
        const day = String(rowDate.getDate()).padStart(2, '0');
        const month = String(rowDate.getMonth() + 1).padStart(2, '0');
        const year = rowDate.getFullYear();
        formattedDate = `${day}/${month}/${year}`;
      } else if (typeof rowDate === 'string') {
        formattedDate = rowDate;
      }
      
      // Dữ liệu trả về theo cấu trúc của sheet như trong hình ảnh
      return {
        rowIndex: index + 2, // +2 vì dòng 1 là tiêu đề và index bắt đầu từ 0
        staffName: row[0] || '', // Tên nhân viên ở cột A
        staffId: row[1] || '',   // ID nhân viên ở cột B
        registerDate: formattedDate, // Ngày đăng ký ở cột C
        shift: row[3] || '',     // Ca làm việc ở cột D
        status: row[4] || '',    // Trạng thái ở cột E
        notes: row[5] || ''      // Ghi chú ở cột F
      };
    });
    
    // Lọc bỏ các ca "Off"
    const filteredByShift = data.filter(item => {
      // Nếu ca làm là "Off" hoặc chứa "off", bỏ qua
      return !(item.shift === 'Off' || 
               (typeof item.shift === 'string' && 
                item.shift.toLowerCase().includes('off')));
    });
    
    // Ghi log số lượng bản ghi sau khi lọc "Off"
    Logger.log(`Đã lọc được ${filteredByShift.length} bản ghi sau khi loại bỏ ca "Off"`);
    
    // Nếu có ngày bắt đầu và kết thúc, lọc dữ liệu theo khoảng thời gian
    let filteredData = filteredByShift;
    if (startDate && endDate) {
      const startDateObj = parseDate(startDate);
      const endDateObj = parseDate(endDate);
      
      if (startDateObj && endDateObj) {
        filteredData = filteredByShift.filter(item => {
          const itemDate = parseDate(item.registerDate);
          return itemDate && itemDate >= startDateObj && itemDate <= endDateObj;
        });
      }
    }
    
    Logger.log(`Đã tìm thấy ${filteredData.length} bản ghi lịch làm việc cho Kỹ thuật`);
    
    return {
      success: true,
      message: `Tải thành công ${filteredData.length} bản ghi`,
      data: filteredData
    };
  } catch (error) {
    Logger.log('Lỗi trong getScheduleDataForTechnician: ' + error.toString());
    return {
      success: false,
      message: 'Lỗi khi tải dữ liệu lịch làm việc: ' + error.toString(),
      data: []
    };
  }
}
// Hàm kiểm tra kết nối và cấu trúc dữ liệu
function testTechnicianSheetConnection() {
  try {
    const SCHEDULE_SHEET_ID = '1XUedszSYOzuW-jcZPPGCqUUwzuqdpuMFSl6gwWBQEEs';
    const spreadsheet = SpreadsheetApp.openById(SCHEDULE_SHEET_ID);
    const sheet = spreadsheet.getSheetByName('Kỹ thuật');
    
    if (!sheet) {
      return {
        success: false,
        message: 'Không tìm thấy sheet "Kỹ thuật"',
        availableSheets: spreadsheet.getSheets().map(s => s.getName())
      };
    }
    
    // Lấy header
    const headers = sheet.getRange(1, 1, 1, 10).getValues()[0];
    
    // Lấy 5 dòng đầu tiên để kiểm tra
    const sampleData = sheet.getRange(2, 1, 5, 10).getValues();
    
    return {
      success: true,
      message: 'Kết nối thành công',
      sheetName: sheet.getName(),
      headers: headers,
      sampleData: sampleData
    };
  } catch (error) {
    return {
      success: false,
      message: 'Lỗi kết nối: ' + error.toString()
    };
  }
}
