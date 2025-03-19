// ======== FUNCTIONS CHÍNH ========

// Cấu hình đường dẫn đến file HTML
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Hệ Thống Phân Đơn')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Hàm kiểm tra kết nối
function testConnection() {
  try {
    // Thử kết nối với Google Sheet để kiểm tra
    const SHEET_ID = '1EDxG3ZXtuc_k1wENp5_M7lWf3pOGVmmEDOIpWXnb08E';
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetNames = ss.getSheets().map(s => s.getName());
    
    return {
      success: true, 
      message: "Kết nối thành công!", 
      sheetNames: sheetNames, 
      timeStamp: new Date().toLocaleString()
    };
  } catch (e) {
    return {
      success: false, 
      message: "Lỗi kết nối: " + e.toString(),
      timeStamp: new Date().toLocaleString()
    };
  }
}

// ======== FUNCTIONS LẤY DỮ LIỆU ========

// Lấy dữ liệu mẫu cho các bài kiểm tra ban đầu
function getTestData() {
  // Tạo dữ liệu mẫu với nhiều mục hơn để kiểm tra phân trang
  const testData = [];
  
  // Tạo 15 mục dữ liệu mẫu
for (let i = 1; i <= 15; i++) {
    const isEven = i % 2 === 0;
    testData.push({
        rowIndex: i + 3,
        reviewStatus: isEven,
        qualityCheck: isEven,
        orderStatus: isEven ? "Hoàn thành" : "Đang xử lý",
        orderId: `DH00${i}`,
        closedBy: `Nhân viên ${isEven ? 'A' : 'B'}`,
        deliveryDate: `${(5 + i % 5).toString().padStart(2, '0')}/03/2025`,
        productScore: isEven ? "95" : "90",
        imageLink: "https://drive.google.com/uc?id=1zma9FQq9ACpEyAhz8a4_ZfIK5uZp0lPZ",
        techNote: isEven ? "Đảm bảo kỹ thuật" : "Kiểm tra kỹ chất lượng", // Thêm trường mới
        orderDate: `${(1 + i % 7).toString().padStart(2, '0')}/03/2025`,
        returnDate: `${(10 + i % 5).toString().padStart(2, '0')}/03/2025`,
        services: isEven ? "Vệ sinh giày|Sửa đế giày" : "Thay đế giày|Làm mới giày",
        customerNote: isEven ? "Làm sạch kỹ phần đế" : "",
        receptionistNote: isEven ? "Khách hàng cần gấp" : "",
        customerName: `Khách hàng ${String.fromCharCode(65 + i % 10)}`,
        qcNote: isEven ? "Đã kiểm tra" : "",
        assignmentDate: isEven ? `${(3 + i % 5).toString().padStart(2, '0')}/03/2025` : "",
        service1: isEven ? "Vệ sinh giày" : "",
        assignee1: isEven ? "Nhân viên X" : "",
        service2: isEven ? "Sửa đế giày" : "",
        assignee2: isEven ? "Nhân viên Y" : "",
        service3: "",
        assignee3: "",
        completionDate: isEven ? `${(8 + i % 3).toString().padStart(2, '0')}/03/2025` : ""
    });
}
  
  return testData;
}

// Kiểm tra xem Google Sheet có tồn tại không
function checkSheetExists() {
  try {
    const SHEET_ID = '1EDxG3ZXtuc_k1wENp5_M7lWf3pOGVmmEDOIpWXnb08E';
    const ss = SpreadsheetApp.openById(SHEET_ID);
    return {
      success: true,
      name: ss.getName(),
      sheets: ss.getSheets().map(s => s.getName()),
      timeStamp: new Date().toLocaleString()
    };
  } catch (e) {
    Logger.log('Lỗi kiểm tra Sheet: ' + e.toString());
    return {
      success: false,
      error: e.toString(),
      timeStamp: new Date().toLocaleString()
    };
  }
}

// Lấy dữ liệu từ Google Sheet
function getSheetData() {
  try {
    // Trước tiên, kiểm tra xem có thể truy cập vào sheet không
    const checkResult = checkSheetExists();
    if (!checkResult.success) {
      Logger.log('Không thể truy cập Google Sheet, sử dụng dữ liệu mẫu: ' + checkResult.error);
      return getTestData(); // Trả về dữ liệu mẫu nếu sheet không thể truy cập
    }
    
    const SHEET_ID = '1EDxG3ZXtuc_k1wENp5_M7lWf3pOGVmmEDOIpWXnb08E';
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    
    // Thử lấy từ Sheet1 trước
    let sheet = spreadsheet.getSheetByName('Sheet1');
    
    // Nếu không tìm thấy Sheet1, thử lấy sheet đầu tiên
    if (!sheet) {
      Logger.log('Sheet "Sheet1" không tồn tại, thử lấy sheet đầu tiên');
      const sheets = spreadsheet.getSheets();
      if (sheets && sheets.length > 0) {
        sheet = sheets[0];
        Logger.log('Sử dụng sheet: ' + sheet.getName());
      } else {
        Logger.log('Không tìm thấy sheet nào trong file');
        return getTestData(); // Trả về dữ liệu mẫu nếu không có sheet nào
      }
    }
    
    // Lấy dữ liệu từ dòng 4
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) {
      Logger.log('Không có dữ liệu trong sheet (lastRow < 4)');
      return getTestData(); // Trả về dữ liệu mẫu nếu không có dữ liệu
    }
    
    // Giới hạn số lượng dòng để tránh quá tải
    const maxRows = Math.min(lastRow, 100);
    const dataRange = sheet.getRange('A4:AG' + maxRows);
    const data = dataRange.getValues();
    
    Logger.log(`Đã tìm thấy ${data.length} dòng dữ liệu từ sheet`);
    
    // Chuyển đổi dữ liệu thành mảng các đối tượng
    const orders = data
      .map((row, index) => {
        // Chỉ lọc các đơn hàng có mã đơn và thông tin khách hàng
        if (!row[4] || row[4].toString().trim() === '') {
          return null; // Bỏ qua dòng không có mã đơn
        }
        
return {
    rowIndex: index + 4, // Bắt đầu từ dòng 4
    reviewStatus: row[0] || false,
    qualityCheck: row[1] || false,
    orderStatus: row[2] || 'Đang xử lý', // Mặc định là "Đang xử lý" thay vì "Chờ xử lý"
    orderId: row[4] || '',
    closedBy: row[5] || '',
    deliveryDate: row[6] || '',
    productScore: row[7] || '',
    imageLink: row[8] ? row[8].toString().split('|')[0].trim() : '', // Lấy URL hình ảnh từ cột I
    techNote: row[17] || '', // Ghi chú kỹ thuật từ cột R (index 17)
    orderDate: formatDateIfDate(row[14]),
    returnDate: formatDateIfDate(row[15]),
    services: row[16] || '',
    customerNote: row[17] || '', // Ghi chú khách hàng cũng từ cột R?
    receptionistNote: row[18] || '', // Ghi chú lễ tân từ cột S (index 18)
    customerName: row[19] || '',
    qcNote: row[20] || '',
    assignmentDate: formatDateIfDate(row[21]),
    service1: row[22] || '',
    assignee1: row[23] || '',
    service2: row[24] || '',
    assignee2: row[25] || '',
    service3: row[26] || '',
    assignee3: row[27] || '',
    completionDate: formatDateIfDate(row[32])
};
      })
      .filter(order => order !== null); // Lọc bỏ các giá trị null
    
    return orders;
  } catch (error) {
    Logger.log('Lỗi trong getSheetData: ' + error.toString());
    // Trả về dữ liệu mẫu trong trường hợp lỗi
    return getTestData();
  }
}

// Hàm hỗ trợ để định dạng các giá trị ngày tháng
function formatDateIfDate(value) {
  if (!value) return '';
  
  // Nếu giá trị đã là string
  if (typeof value === 'string') return value;
  
  // Kiểm tra nếu là đối tượng Date
  if (value instanceof Date && !isNaN(value)) {
    const day = String(value.getDate()).padStart(2, '0');
    const month = String(value.getMonth() + 1).padStart(2, '0');
    const year = value.getFullYear();
    return `${day}/${month}/${year}`;
  }
  
  return String(value);
}

// Lấy danh sách nhân viên
function getStaffList() {
  try {
    // Dữ liệu nhân viên mẫu
    const testStaff = ["Nhân viên A", "Nhân viên B", "Nhân viên C", "Nhân viên X", "Nhân viên Y", "Nhân viên Z"];
    
    // Kiểm tra xem sheet có thể truy cập không
    const checkResult = checkSheetExists();
    if (!checkResult.success) {
      Logger.log('Không thể truy cập Google Sheet để lấy danh sách nhân viên: ' + checkResult.error);
      return testStaff;
    }
    
    const SHEET_ID = '1EDxG3ZXtuc_k1wENp5_M7lWf3pOGVmmEDOIpWXnb08E';
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    
    // Thử tìm sheet "Data" trước
    let sheet = spreadsheet.getSheetByName('Data');
    
    // Nếu không tìm thấy "Data", tìm sheet nào có tên phù hợp ("Staff", "Nhân viên", v.v.)
    if (!sheet) {
      const sheets = spreadsheet.getSheets();
      const dataSheetNames = ['Data', 'Staff', 'Nhân viên', 'NhanVien', 'Personnel'];
      
      for (const name of dataSheetNames) {
        const found = sheets.find(s => s.getName().toUpperCase().includes(name.toUpperCase()));
        if (found) {
          sheet = found;
          Logger.log('Sử dụng sheet cho nhân viên: ' + sheet.getName());
          break;
        }
      }
      
      // Nếu vẫn không tìm thấy, trả về dữ liệu mẫu
      if (!sheet) {
        Logger.log('Không tìm thấy sheet nhân viên phù hợp');
        return testStaff;
      }
    }
    
    // Lấy dữ liệu nhân viên (thử nhiều phạm vi)
    let staffData = [];
    const ranges = ['A2:A20', 'B2:B20', 'A1:A20'];
    
    for (const range of ranges) {
      try {
        const data = sheet.getRange(range).getValues();
        const tempStaff = data.map(row => row[0]).filter(Boolean);
        
        if (tempStaff.length > 0) {
          staffData = tempStaff;
          Logger.log(`Tìm thấy ${staffData.length} nhân viên từ phạm vi ${range}`);
          break;
        }
      } catch (e) {
        Logger.log(`Lỗi khi thử lấy dữ liệu từ phạm vi ${range}: ${e}`);
        continue;
      }
    }
    
    // Nếu không tìm thấy dữ liệu, trả về dữ liệu mẫu
    if (staffData.length === 0) {
      Logger.log('Không tìm thấy dữ liệu nhân viên, sử dụng dữ liệu mẫu');
      return testStaff;
    }
    
    return staffData;
  } catch (error) {
    Logger.log('Lỗi khi lấy danh sách nhân viên: ' + error.toString());
    return ["Nhân viên A", "Nhân viên B", "Nhân viên C", "Nhân viên X", "Nhân viên Y", "Nhân viên Z"];
  }
}
// Hàm lấy dữ liệu từ Google Sheet với bộ lọc
function getFilteredSheetData(filterParams) {
  try {
    // Trước tiên, kiểm tra xem có thể truy cập vào sheet không
    const checkResult = checkSheetExists();
    if (!checkResult.success) {
      Logger.log('Không thể truy cập Google Sheet, sử dụng dữ liệu mẫu: ' + checkResult.error);
      return getFilteredTestData(filterParams); // Trả về dữ liệu mẫu đã lọc
    }
    
    const SHEET_ID = '1EDxG3ZXtuc_k1wENp5_M7lWf3pOGVmmEDOIpWXnb08E';
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    
    // Thử lấy từ Sheet1 trước
    let sheet = spreadsheet.getSheetByName('Sheet1');
    
    // Nếu không tìm thấy Sheet1, thử lấy sheet đầu tiên
    if (!sheet) {
      Logger.log('Sheet "Sheet1" không tồn tại, thử lấy sheet đầu tiên');
      const sheets = spreadsheet.getSheets();
      if (sheets && sheets.length > 0) {
        sheet = sheets[0];
        Logger.log('Sử dụng sheet: ' + sheet.getName());
      } else {
        Logger.log('Không tìm thấy sheet nào trong file');
        return getFilteredTestData(filterParams); // Trả về dữ liệu mẫu đã lọc
      }
    }
    
    // Lấy tất cả dữ liệu từ dòng 4
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) {
      Logger.log('Không có dữ liệu trong sheet (lastRow < 4)');
      return getFilteredTestData(filterParams);
    }
    
    // Trích xuất các giá trị bộ lọc
    const orderDateFrom = filterParams ? filterParams.orderDateFrom : null;
    const orderDateTo = filterParams ? filterParams.orderDateTo : null;
    const returnDateFrom = filterParams ? filterParams.returnDateFrom : null;
    const returnDateTo = filterParams ? filterParams.returnDateTo : null;
    const serviceFilter = filterParams ? filterParams.serviceFilter : null;
    const orderIdFilter = filterParams ? filterParams.orderIdFilter : null;
    const statusFilter = filterParams ? filterParams.statusFilter : null;
    const staffFilter = filterParams ? filterParams.staffFilter : null;

    // Giới hạn số lượng dòng để tránh timeout
    const maxRows = 500; // Giới hạn số dòng trả về để tránh quá tải
    const actualRows = Math.min(lastRow, maxRows + 3); // +3 vì bắt đầu từ dòng 4
    
    // Lấy dữ liệu
    const dataRange = sheet.getRange('A4:AG' + actualRows);
    const data = dataRange.getValues();
    
    Logger.log(`Đã tìm thấy ${data.length} dòng dữ liệu từ sheet, bắt đầu lọc theo ${JSON.stringify(filterParams)}`);
    
    // Chuyển đổi dữ liệu thành mảng các đối tượng và lọc ngay tại đây
    const filteredOrders = [];
    
    for (let index = 0; index < data.length; index++) {
      const row = data[index];
      
      // Bỏ qua dòng không có mã đơn
      if (!row[4] || row[4].toString().trim() === '') {
        continue;
      }
      
      // Tạo đối tượng đơn hàng
      const order = {
        rowIndex: index + 4, // Bắt đầu từ dòng 4
        reviewStatus: row[0] || false,
        qualityCheck: row[1] || false,
        orderStatus: row[2] || 'Đang xử lý',
        orderId: row[4] || '',
        closedBy: row[5] || '',
        deliveryDate: row[6] || '',
        productScore: row[7] || '',
        imageLink: row[8] ? row[8].toString().split('|')[0].trim() : '',
        techNote: row[17] || '',
        orderDate: formatDateIfDate(row[14]),
        returnDate: formatDateIfDate(row[15]),
        services: row[16] || '',
        customerNote: row[17] || '',
        receptionistNote: row[18] || '',
        customerName: row[19] || '',
        qcNote: row[20] || '',
        assignmentDate: formatDateIfDate(row[21]),
        service1: row[22] || '',
        assignee1: row[23] || '',
        service2: row[24] || '',
        assignee2: row[25] || '',
        service3: row[26] || '',
        assignee3: row[27] || '',
        completionDate: formatDateIfDate(row[32])
      };
      
      // Áp dụng bộ lọc
      if (!shouldFilterOut(order, filterParams)) {
        filteredOrders.push(order);
      }
      
      // Nếu đã đạt đến giới hạn kết quả, dừng
      if (filteredOrders.length >= maxRows) {
        Logger.log(`Đã đạt giới hạn ${maxRows} kết quả, dừng lọc`);
        break;
      }
    }
    
    Logger.log(`Đã lọc được ${filteredOrders.length} đơn hàng phù hợp từ ${data.length} dòng dữ liệu`);
    
    return filteredOrders;
  } catch (error) {
    Logger.log('Lỗi trong getFilteredSheetData: ' + error.toString());
    // Trả về dữ liệu mẫu đã lọc trong trường hợp lỗi
    return getFilteredTestData(filterParams);
  }
}

// Hàm kiểm tra xem đơn hàng có nên bị lọc ra không
function shouldFilterOut(order, filterParams) {
  if (!filterParams) return false; // Nếu không có bộ lọc, không lọc ra
  
  // Lọc theo ngày nhận
  if (filterParams.orderDateFrom && !isDateAfterOrEqual(order.orderDate, filterParams.orderDateFrom)) {
    return true;
  }
  if (filterParams.orderDateTo && !isDateBeforeOrEqual(order.orderDate, filterParams.orderDateTo)) {
    return true;
  }
  
  // Lọc theo ngày trả
  if (filterParams.returnDateFrom && !isDateAfterOrEqual(order.returnDate, filterParams.returnDateFrom)) {
    return true;
  }
  if (filterParams.returnDateTo && !isDateBeforeOrEqual(order.returnDate, filterParams.returnDateTo)) {
    return true;
  }
  
  // Lọc theo tình trạng đơn
  if (filterParams.statusFilter) {
    let matchStatus = false;
    
    if (filterParams.statusFilter === "QC Duyệt Đơn") {
      matchStatus = order.orderStatus === "Đơn Chưa Duyệt" || order.orderStatus === "QC Duyệt Đơn";
    } else if (filterParams.statusFilter === "Đang xử lý") {
      matchStatus = order.orderStatus === "Đơn đang được xử lý" || 
                   order.orderStatus === "Hoàn Thành,Chờ QC" ||
                   order.orderStatus === "Đang xử lý";
    } else if (filterParams.statusFilter === "Hoàn thành") {
      matchStatus = order.orderStatus === "Qc Xong, Báo khách nhận sản phẩm" || 
                   order.orderStatus === "Hoàn thành";
    } else {
      matchStatus = order.orderStatus === filterParams.statusFilter;
    }
    
    if (!matchStatus) {
      return true;
    }
  }
  
  // Lọc theo dịch vụ
  if (filterParams.serviceFilter && (!order.services || !order.services.includes(filterParams.serviceFilter))) {
    return true;
  }
  
  // Lọc theo mã đơn hàng
  if (filterParams.orderIdFilter && !order.orderId.toLowerCase().includes(filterParams.orderIdFilter.toLowerCase())) {
    return true;
  }
  
  // Lọc theo nhân viên
  if (filterParams.staffFilter) {
    let staffFound = false;
    
    if (order.assignee1 && order.assignee1 === filterParams.staffFilter) {
      staffFound = true;
    } else if (order.assignee2 && order.assignee2 === filterParams.staffFilter) {
      staffFound = true;
    } else if (order.assignee3 && order.assignee3 === filterParams.staffFilter) {
      staffFound = true;
    }
    
    if (!staffFound) {
      return true;
    }
  }
  
  return false; // Không lọc ra
}

// Hàm kiểm tra ngày
function isDateAfterOrEqual(dateStr1, dateStr2) {
  if (!dateStr1 || !dateStr2) return true;
  
  const date1 = parseDateStr(dateStr1);
  const date2 = parseDateStr(dateStr2);
  
  if (!date1 || !date2) return true;
  
  return date1 >= date2;
}

function isDateBeforeOrEqual(dateStr1, dateStr2) {
  if (!dateStr1 || !dateStr2) return true;
  
  const date1 = parseDateStr(dateStr1);
  const date2 = parseDateStr(dateStr2);
  
  if (!date1 || !date2) return true;
  
  return date1 <= date2;
}

// Hàm parse date string
function parseDateStr(dateStr) {
  try {
    const parts = dateStr.split('/');
    if (parts.length === 3) {
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1; // Tháng từ 0-11
      const year = parseInt(parts[2], 10);
      return new Date(year, month, day).getTime();
    }
  } catch (e) {
    Logger.log('Lỗi parse date: ' + e);
  }
  return null;
}

// Hàm lọc dữ liệu mẫu
function getFilteredTestData(filterParams) {
  // Lấy dữ liệu mẫu
  const testData = getTestData();
  
  // Nếu không có bộ lọc, trả về tất cả dữ liệu mẫu
  if (!filterParams) return testData;
  
  // Lọc dữ liệu mẫu theo bộ lọc
  return testData.filter(order => !shouldFilterOut(order, filterParams));
}
// ======== FUNCTIONS CẬP NHẬT ========

// Cập nhật thông tin phân công
function updateAssignment(rowIndex, data) {
  try {
    // Ghi log dữ liệu nhận được để debug
    Logger.log('updateAssignment - rowIndex: ' + rowIndex);
    Logger.log('updateAssignment - data: ' + JSON.stringify(data));
    
    // Kiểm tra xem có thể cập nhật dữ liệu thực không
    const checkResult = checkSheetExists();
    if (!checkResult.success) {
      // Nếu sheet không thể truy cập, giả lập cập nhật thành công
      Logger.log('Không thể truy cập Google Sheet để cập nhật, trả về thành công giả');
      return {success: true, message: "Cập nhật thành công (Chế độ thử nghiệm)", timeStamp: new Date().toLocaleString()};
    }
    
    const SHEET_ID = '1EDxG3ZXtuc_k1wENp5_M7lWf3pOGVmmEDOIpWXnb08E';
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet = spreadsheet.getSheetByName('Sheet1');
    
    if (!sheet) {
      Logger.log('Không tìm thấy Sheet1 để cập nhật');
      return {success: false, message: "Không tìm thấy Sheet1", timeStamp: new Date().toLocaleString()};
    }
    
    // Cập nhật thông tin QC Note (cột U)
    if (data.qcNote !== undefined) {
      sheet.getRange(`U${rowIndex}`).setValue(data.qcNote);
    }
    
    // Cập nhật ngày phân công (cột V)
    if (data.assignmentDate) {
      sheet.getRange(`V${rowIndex}`).setValue(data.assignmentDate);
    }
    
    // Cập nhật dịch vụ 1 (cột W)
    if (data.service1 !== undefined) {
      sheet.getRange(`W${rowIndex}`).setValue(data.service1);
    }
    
    // Cập nhật người làm 1 (cột X)
    if (data.assignee1 !== undefined) {
      sheet.getRange(`X${rowIndex}`).setValue(data.assignee1);
    }
    // Cập nhật trạng thái Quality Check (cột B)
    if (data.qualityCheck !== undefined) {
    sheet.getRange(`B${rowIndex}`).setValue(data.qualityCheck);
}
    // Cập nhật dịch vụ 2 (cột Y)
    if (data.service2 !== undefined) {
      sheet.getRange(`Y${rowIndex}`).setValue(data.service2);
    }
    
    // Cập nhật người làm 2 (cột Z)
    if (data.assignee2 !== undefined) {
      sheet.getRange(`Z${rowIndex}`).setValue(data.assignee2);
    }
    
    // Cập nhật dịch vụ 3 (cột AA)
    if (data.service3 !== undefined) {
      sheet.getRange(`AA${rowIndex}`).setValue(data.service3);
    }
    
    // Cập nhật người làm 3 (cột AB)
    if (data.assignee3 !== undefined) {
      sheet.getRange(`AB${rowIndex}`).setValue(data.assignee3);
    }
    
    // Cập nhật ngày hoàn thành (cột AG)
    if (data.completionDate !== undefined) {
      sheet.getRange(`AG${rowIndex}`).setValue(data.completionDate);
    }
    
    Logger.log('Cập nhật thành công đơn hàng tại dòng ' + rowIndex);
    return {
      success: true, 
      message: "Cập nhật thành công",
      timeStamp: new Date().toLocaleString(),
      updatedData: data
    };
  } catch (error) {
    Logger.log('Lỗi trong updateAssignment: ' + error.toString());
    return {
      success: false, 
      message: "Lỗi: " + error.toString(),
      timeStamp: new Date().toLocaleString()
    };
  }
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
