# Bài 2: Hồ sơ dự án PTIT-Learning (module Học trực tuyến)

## Các file nộp
| File | Nội dung |
|---|---|
| `PTIT_SRS_LopHocAo.docx` | SRS tính năng Lớp học ảo, 10 Heading 2, mục lục đã Update entire table |
| `PTIT_DanhSachDangKy.xlsx` | 200 bản ghi, 1 Header Row duy nhất, AutoFilter (Ctrl+Shift+L) đã lọc CNTT và chuyên cần < 7 |
| `PTIT_Slide_TinhNang.pptx` | Slide nền Navy, chữ trắng, accent cam, flowchart 5 bước, bố cục 3 vùng |

## Phần 1. Sơ đồ cây Heading của SRS
```
Tài liệu SRS (Title)
├── 1. Giới thiệu (Heading 1)
│   ├── 1.1 Mục đích (Heading 2)
│   └── 1.2 Phạm vi (Heading 2)
├── 2. Mô tả tổng quan (Heading 1)
│   ├── 2.1 Đối tượng người dùng (Heading 2)
│   └── 2.2 Môi trường vận hành (Heading 2)
├── 3. Yêu cầu chức năng (Heading 1)
│   ├── 3.1 Tạo và quản lý lớp học ảo, FR-01 (Heading 2)
│   ├── 3.2 Tham gia lớp học trực tuyến, FR-02 (Heading 2)
│   ├── 3.3 Chia sẻ màn hình và tài liệu, FR-03 (Heading 2)
│   ├── 3.4 Chat và hỏi đáp, FR-04 (Heading 2)
│   └── 3.5 Điểm danh và theo dõi chuyên cần, FR-05 (Heading 2)
└── 4. Yêu cầu phi chức năng (Heading 1)
    └── 4.1 Hiệu năng và bảo mật (Heading 2)
```
Tổng cộng 4 Heading 1 và đúng 10 Heading 2. Dùng Navigation Pane (View > Navigation Pane) để kiểm soát và kéo thả các mục.

## Phần 2. Triển khai
- Word: mục lục tự động đã được cập nhật toàn bộ trước khi lưu.
- Excel: dòng 1 là header duy nhất, không có dòng tổng cộng trong vùng dữ liệu. Thống kê nằm ở sheet `ThongKe`. Kết quả lọc có 46 sinh viên. Dữ liệu là dữ liệu giả định.
- PowerPoint: Vùng 1 là tiêu đề khẳng định, Vùng 2 là flowchart, Vùng 3 là thông điệp chốt. Chữ trắng và chữ navy trên nền cam/trắng đều có độ tương phản cao.

## Phần 3. Kiểm thử: lọc theo tên không ra kết quả dù thấy tên
**Nguyên nhân:** tên có khoảng trắng thừa ở đầu hoặc cuối (ví dụ `"Nguyễn Văn An "`), nên khác với chuỗi bạn gõ vào ô tìm.

**Cách xử lý bằng TRIM:**
1. Chèn cột phụ bên cạnh, ví dụ ô `G2`: `=TRIM(B2)`, kéo xuống hết dữ liệu. `TRIM` xóa khoảng trắng đầu/cuối và rút các khoảng trắng giữa từ về một dấu cách.
2. Kiểm tra: `=LEN(B2)<>LEN(G2)` cho TRUE tức ô đó có khoảng trắng thừa.
3. Sao chép cột `G`, dán đè vào cột `B` bằng Paste Special > Values để thay dữ liệu gốc bằng dữ liệu đã sạch.
4. Xóa cột phụ, bật lại bộ lọc (Ctrl+Shift+L) và lọc theo tên, kết quả sẽ hiện đúng.

Trong file Excel đã có sẵn cột `HoTen_Sach` dùng `TRIM`, và `ThongKe` cho biết 36 tên có khoảng trắng thừa.
