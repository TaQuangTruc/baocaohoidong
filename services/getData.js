const xlsx = require("xlsx");
const { filter, filterLvtn } = require("./helper.js");

function getData(name, filePath, bodyRequest, isLvtn = false) {
  const workbook = xlsx.readFile(filePath);
  const sheetNames = workbook.SheetNames;
  let combinedData = []; // Kết quả tổng hợp từ tất cả các sheet

  sheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    // Kiểm tra xem các cột index trong attributes có trùng key không
    const isValidSheet = bodyRequest.attributes.every((attr) => {
      const headerValue = data[0][attr.index];
      if (headerValue != attr.key) {
        console.log(
          `Sheet "${sheetName}" không hợp lệ: Cột ${attr.index} không khớp với key "${attr.key}". Bỏ qua sheet này.`
        );
        return false;
      }
      return true;
    });

    if (!isValidSheet) return; // Bỏ qua sheet này nếu không hợp lệ

    // Lọc dữ liệu theo tiêu chí
    let formattedData;
    if (isLvtn) formattedData = filterLvtn(data, bodyRequest);
    else formattedData = filter(data, bodyRequest);
    
  
    combinedData = combinedData.concat(formattedData); // Gộp dữ liệu lại
  });

  return combinedData;
}


module.exports = {getData};
