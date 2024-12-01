// Thông tin giảng viên
const listInfos = [
  {
    name: "Lê Bình Đẳng",
    email: "truc.taquang23904@hcmut.edu.vn",
  },
];

// Đường dẫn đến thực mục kết quả
const resultDir = "result";

// Đường dẫn tới file đồ án chuyên ngành
const pathDacn = "database/dacn.xlsx";

// Đường dẫn tới file đồ án tốt nghiệp
const pathDatn = "database/datn.xlsx";

// Đường dẫn tới file luận văn tốt nghiệp
const pathLvtn = "database/lvtn.xlsx";

// Tên file kết quả muốn gửi
const nameFileResult = "Lịch báo cáo hội đồng";

// Tiêu đề email
const titleEmail = `LỊCH BÁO CÁO HỘI ĐỒNG`;

// Nội dung email
const bodyEmail = `
    Kính gửi thầy/cô,
    Xin chuyển đến thầy/cô đề xuất về phản biện và lịch hội đồng bảo vệ Luận văn TN/ Đồ án TN- Đồ án chuyên ngành (KHMT)- Đồ án môn học KTMT HK241, link: Lịch Hội đồng 241

    Thầy/cô vui lòng phản hồi nếu có thay đổi về lịch tham dự. Hạn phản hồi: 25/11/2024.

    1, Luận văn TN/ Đồ án TN:

    - Ngày SV nộp báo cáo: 09/12/2024

    - Tuần phản biện: tuần 09/12 và tuần 16/12/2024

    - Ngày hội Kỹ thuật: tuần 16/12/2024 (chi tiết thông báo sau).

    - Tuần tổ chức hội đồng: tuần 23/12/2024

    2, Đồ án chuyên ngành/ Đồ án môn học Kỹ thuật Máy tính:

    - Ngày SV nộp báo cáo: 16/12/2024

    - Tuần tổ chức hội đồng: tuần 30/12/2024

    Một số lưu ý:
    - Thư ký Khoa sẽ chuẩn bị hồ sơ cho hội đồng (QĐ thành lập HĐ, Biên bản, Phiếu nhận xét sử dụng tại hội đồng). Biểu mẫu dạng file được cung cấp theo link: Biểu mẫu Hội đồng 241
    - Thư ký Hội đồng lưu ý hạn nộp hồ sơ sau bảo vệ về Khoa: tuần 23/12/2024 (đối với ĐATN), tuần 30/12/2024 (đối với ĐACN/ĐAMHKTMT)
    -  Địa điểm tổ chức hội đồng: 
    + chương trình tiêu chuẩn, tài năng: ĐH Bách Khoa, cơ sở Dĩ An, Bình Dương.
    + chương trình giảng dạy bằng tiếng Anh, định hướng Nhật Bản: ĐH Bách Khoa, cơ sở Lý Thường Kiệt, quận 10.
    Trân trọng,
    Như Vi
`;

module.exports = {
  listInfos,
  resultDir,
  pathDacn,
  pathDatn,
  pathLvtn,
  nameFileResult,
  titleEmail,
  bodyEmail,
};
