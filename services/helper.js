const xlsx = require("xlsx");
const fs = require('fs');

function readExcelFile(filePath) {
  const workbook = xlsx.readFile(filePath); // Đọc file Excel
  const allData = []; // Mảng để lưu dữ liệu từ tất cả các sheet

  workbook.SheetNames.forEach((sheetName, index) => {
    const sheet = workbook.Sheets[sheetName]; // Lấy dữ liệu sheet
    const sheetData = xlsx.utils.sheet_to_json(sheet, { header: 1 }); // Chuyển sheet thành mảng các hàng

    // Lọc bỏ các dòng trống
    const filteredData = sheetData.filter((row) =>
      row.some((cell) => cell !== null && cell !== undefined && cell !== "")
    );

    // Sheet đầu tiên lấy toàn bộ dữ liệu
    if (index === 0) {
      allData.push(...filteredData);
    } else {
      // Các sheet khác chỉ lấy từ hàng thứ 2 trở đi
      allData.push(...filteredData.slice(1));
    }
  });

  return allData;
}

// Hàm lưu dữ liệu vào file Excel
function saveToExcelFile(
  data,
  filePath,
  nameSheet,
  isNormalPresentation = false
) {
  let workbook;
  try {
    workbook = xlsx.readFile(filePath);
  } catch (error) {
    console.log(`File not found, creating new file: ${filePath}`);
    workbook = xlsx.utils.book_new();
  }

  // Kiểm tra nếu sheetName đã tồn tại, xóa nó
  if (workbook.SheetNames.includes(nameSheet)) {
    console.log(`Sheet ${nameSheet} already exists. Removing old content.`);
    const sheetIndex = workbook.SheetNames.indexOf(nameSheet);
    workbook.SheetNames.splice(sheetIndex, 1);
    delete workbook.Sheets[nameSheet];
  }

  const worksheetData = [];

  if (isNormalPresentation) {
    // Trình bày như bình thường
    const headers = Object.keys(data[0] || {});
    worksheetData.push(headers);

    data.forEach((row) => {
      worksheetData.push(Object.values(row));
    });

    const worksheet = xlsx.utils.aoa_to_sheet(worksheetData);

    // Điều chỉnh kích thước cột
    const maxWidths = worksheetData[0].map((_, colIndex) =>
      Math.max(
        ...worksheetData.map((row) =>
          row[colIndex] ? row[colIndex].toString().length : 10
        )
      )
    );

    worksheet["!cols"] = maxWidths.map((width) => ({
      wch: Math.max(10, width + 2),
    }));

    xlsx.utils.book_append_sheet(workbook, worksheet, nameSheet);
    xlsx.writeFile(workbook, filePath);
    return;
  }

  // Trình bày với merge cột
  if (data.length > 0 && data[0].danhSachDeTai.length > 0) {
    const headers = ["Hội đồng", ...Object.keys(data[0].danhSachDeTai[0])];
    worksheetData.push(headers);
  }

  const mergeRanges = [];

  data.forEach((hoidong) => {
    const tenHoiDong = hoidong.tenHoiDong;
    const rowCount = hoidong.danhSachDeTai.length;

    hoidong.danhSachDeTai.forEach((detai) => {
      const row = [tenHoiDong, ...Object.values(detai)];
      worksheetData.push(row);
    });

    if (rowCount > 1) {
      const startRow = worksheetData.length - rowCount;
      const endRow = worksheetData.length;
      mergeRanges.push({
        s: { r: startRow, c: 0 },
        e: { r: endRow - 1, c: 0 },
      });
    }
  });

  const worksheet = xlsx.utils.aoa_to_sheet(worksheetData);

  if (mergeRanges.length > 0) {
    worksheet["!merges"] = mergeRanges;
  }

  const headers = worksheetData[0];
  headers.forEach((header, colIndex) => {
    const cellAddress = xlsx.utils.encode_cell({ r: 0, c: colIndex });
    if (!worksheet[cellAddress]) worksheet[cellAddress] = {};
    worksheet[cellAddress].s = {
      font: { bold: true },
      alignment: { horizontal: "center", vertical: "center" },
    };
  });

  const maxWidths = worksheetData[0].map((_, colIndex) =>
    Math.max(
      ...worksheetData.map((row) =>
        row[colIndex] ? row[colIndex].toString().length : 10
      )
    )
  );

  worksheet["!cols"] = maxWidths.map((width) => ({
    wch: Math.max(10, width + 2),
  }));

  xlsx.utils.book_append_sheet(workbook, worksheet, nameSheet);
  xlsx.writeFile(workbook, filePath);
}

function filter(data, bodyRequest) {
  const formattedData = [];
  const attributes = bodyRequest.attributes;
  const criteriaArray = bodyRequest.criteria;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    // Kiểm tra nếu row[0] chứa từ 'Hội đồng'
    if (
      typeof row[0] == "string" &&
      row[0].toLowerCase().includes("hội đồng")
    ) {
      const hoidong = {
        tenHoiDong: row[0],
        danhSachDeTai: [],
      };

      // Lấy danh sách đề tài từ các dòng tiếp theo
      for (let j = i + 1; j < data.length; j++) {
        const currentRow = data[j];

        // Nếu dòng tiếp theo là 'Hội đồng' khác, dừng vòng lặp
        if (
          typeof currentRow[0] === "string" &&
          currentRow[0].toLowerCase().includes("hội đồng")
        ) {
          i = j - 1;
          break;
        }

        // Kiểm tra nếu thỏa 1 trong các tiêu chí trong criteriaArray
        const isCriteriaMet = criteriaArray.some((criteria) => {
          return (
            criteria.index >= 0 &&
            currentRow[criteria.index]?.includes(criteria.value)
          );
        });

        // Nếu không thỏa mãn ít nhất 1 tiêu chuẩn, tiếp tục vòng lặp
        if (!isCriteriaMet) continue;

        // Đối tượng đề tài
        const detai = attributes.reduce((acc, attribute) => {
          acc[attribute.key] = currentRow[attribute.index] || "";
          return acc;
        }, {});

        // Kiểm tra xem đề tài đã tồn tại trong danh sách hay chưa
        const isDuplicate = hoidong.danhSachDeTai.some((existingDeTai) =>
          Object.keys(detai).every((key) => existingDeTai[key] === detai[key])
        );

        // Chỉ thêm đề tài nếu không trùng lặp
        if (!isDuplicate) {
          hoidong.danhSachDeTai.push(detai);
        }
      }

      if (hoidong.danhSachDeTai.length > 0) formattedData.push(hoidong);
    }
  }

  return formattedData;
}

function filterLvtn(data, bodyRequest) {
  const formattedData = []; // Kết quả chứa các đề tài sau khi lọc
  const attributes = bodyRequest.attributes; // Các thuộc tính cần lấy từ dữ liệu
  const criteriaArray = bodyRequest.criteria; // Các tiêu chí lọc

  for (let i = 1; i < data.length; i++) {
    const currentRow = data[i]; // Dòng dữ liệu hiện tại

    // Kiểm tra nếu thỏa mãn ít nhất một tiêu chí
    const isCriteriaMet = criteriaArray.some((criteria) => {
      return (
        criteria.index >= 0 &&
        currentRow[criteria.index]?.toString().includes(criteria.value)
      );
    });

    // Nếu không thỏa mãn tiêu chí nào, tiếp tục với dòng tiếp theo
    if (!isCriteriaMet) continue;

    // Tạo đối tượng đề tài mới từ các thuộc tính
    const detai = attributes.reduce((acc, attribute) => {
      acc[attribute.key] = currentRow[attribute.index] || ""; // Lấy giá trị tại vị trí index
      return acc;
    }, {});

    // Kiểm tra trùng lặp
    const isDuplicate = formattedData.some((existingDeTai) =>
      Object.keys(detai).every((key) => existingDeTai[key] === detai[key])
    );

    // Nếu không trùng lặp, thêm vào danh sách kết quả
    if (!isDuplicate) {
      formattedData.push(detai);
    }
  }

  return formattedData; // Trả về danh sách đề tài đã lọc
}

function createFolderIfNotExits(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
    console.log(`Directory created: ${dirPath}`);
  } else {
    console.log(`Directory already exists: ${dirPath}`);
  }
}

module.exports = { readExcelFile, saveToExcelFile, filter, filterLvtn, createFolderIfNotExits };
