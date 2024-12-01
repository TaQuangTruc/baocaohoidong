const fs = require("fs");
const {
  listInfos,
  resultDir,
  pathDacn,
  pathDatn,
  pathLvtn,
  nameFileResult,
  titleEmail,
  bodyEmail,
} = require("./info.js");

const { getData } = require("./services/getData.js");
const {
  saveToExcelFile,
  createFolderIfNotExits,
} = require("./services/helper.js");
const authorize = require("./services/googleApiAuthService.js");
const { sendEmail } = require("./services/gmailApiService.js");

// Sử dụng vòng lặp để tạo file Excel
async function main() {
  // Tạo thư mục lưu kết quả nếu chưa tồn tại
  createFolderIfNotExits(resultDir);

  // Đăng nhập bằng google để lấy quyền gửi email
  const authClient = await authorize().then().catch(console.error);

  // Xử lý thông tin từng cán bộ
  for (let i = 0; i < listInfos.length; i++) {
    const name = listInfos[i].name;
    const email = listInfos[i].email;
    const resultPath = `${resultDir}/${name}.xlsx`;

    console.log("***Bắt đầu xử lý Đồ án chuyên ngành");
    const bodyRequestOfDacn = {
      attributes: [
        { index: 5, key: "Chủ tịch" },
        { index: 6, key: "Thư ký" },
        { index: 7, key: "Uỷ viên" },
        { index: 9, key: "Việt - Tên ĐT" },
      ],
      criteria: [
        { index: 5, key: "Chủ tịch", value: name },
        { index: 6, key: "Thư ký", value: name },
        { index: 7, key: "Uỷ viên", value: name },
      ],
    };
    const dataOfDacn = getData(name, pathDacn, bodyRequestOfDacn);
    if (dataOfDacn.length == 0) {
      console.log(
        `Không tìm thấy dữ liệu của ${name} trong báo cáo hội đồng Đồ án chuyên ngành`
      );
    } else {
      saveToExcelFile(dataOfDacn, resultPath, "Đồ án chuyên ngành");
    }

    console.log("***Bắt đầu xử lý Đồ án tốt nghiệp");

    const bodyRequestOfDatn = {
      attributes: [
        { index: 5, key: "Phản biện" },
        { index: 6, key: "Chủ tịch" },
        { index: 7, key: "Thư ký" },
        { index: 8, key: "Ủy viên" },
        { index: 10, key: "Việt - Tên ĐT" },
      ],
      criteria: [
        { index: 5, key: "Phản biện", value: name },
        { index: 6, key: "Chủ tịch", value: name },
        { index: 7, key: "Thư ký", value: name },
        { index: 8, key: "Ủy viên", value: name },
      ],
    };
    const dataOfDatn = getData(name, pathDatn, bodyRequestOfDatn);
    if (dataOfDatn.length == 0) {
      console.log(
        `Không tìm thấy dữ liệu của ${name} trong báo cáo hội đồng Đồ án tốt nghiệp`
      );
    } else {
      saveToExcelFile(dataOfDatn, resultPath, "Đồ án tốt nghiệp");
    }

    console.log("***Bắt đầu xử lý Luận văn tốt nghiệp");

    const bodyRequestOfLvtn = {
      attributes: [
        { index: 5, key: "Phản biện" },
        { index: 7, key: "Việt - Tên ĐT" },
      ],
      criteria: [{ index: 5, key: "Phản biện", value: name }],
    };
    const dataOfLvtn = getData(name, pathLvtn, bodyRequestOfLvtn, true);
    if (dataOfLvtn.length == 0) {
      console.log(
        `Không tìm thấy dữ liệu của ${name} trong báo cáo hội đồng Luận văn tốt nghiệp`
      );
    } else {
      saveToExcelFile(dataOfLvtn, resultPath, "Luận văn tốt nghiệp", true);
    }

    const filePath = `result/${name}.xlsx`;

    // Đọc file và mã hóa Base64
    const fileContent = fs.readFileSync(filePath);
    const fileBase64 = fileContent.toString("base64");

    // Tạo nội dung email (MIME message)
    const message = [
      `To: ${email}`,
      `Subject: =?UTF-8?B?${Buffer.from(titleEmail).toString('base64')}?=`,
      'Content-Type: multipart/mixed; boundary="boundary"',
      "",
      "--boundary",
      'Content-Type: text/plain; charset="UTF-8"',
      "",
      `${bodyEmail}`,
      "",
      "--boundary",
      `Content-Type: application/pdf; name="${nameFileResult}"`,
      `Content-Disposition: attachment; filename="${nameFileResult}.xlsx"`,
      "Content-Transfer-Encoding: base64",
      "",
      fileBase64,
      "",
      "--boundary--",
    ].join("\n");

    await sendEmail(authClient, message).catch(console.error);
  }
}

main();
