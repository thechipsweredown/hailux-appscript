# HaiLux - Hệ thống quản lý công việc sửa đồ da

Web app nội bộ xây dựng trên **Google Apps Script** + **Google Sheets**, phục vụ quản lý đơn sửa chữa đồ da tại tiệm HaiLux Luxury & Spa.

## Tính năng

- Quản lý đơn hàng (job): tạo, sửa, xóa, lọc theo trạng thái / ngày nhận / danh mục
- Quản lý công đoạn (task) trong từng đơn: phân công thợ, theo dõi tiến độ, tải ảnh nghiệm thu
- Giao diện thợ: xem danh sách việc được phân công, cập nhật trạng thái, upload ảnh bằng điện thoại
- Dashboard tổng quan: doanh thu, trạng thái đơn, hiệu suất từng thợ
- Phân quyền: thợ (employee) / quản lý (manager, xác thực bằng mật khẩu)
- Lưu ảnh sản phẩm và ảnh nghiệm thu trên Google Drive, phục vụ qua CDN

## Cấu trúc file

| File | Nội dung |
|------|----------|
| `Code.gs` | Entry point — `doGet()` |
| `Utils.gs` | Hàm tiện ích dùng chung |
| `Setup.gs` | Khởi tạo sheet, seed data, debug |
| `Auth.gs` | Xác thực mật khẩu quản lý |
| `Users.gs` | CRUD nhân viên |
| `Statuses.gs` | CRUD trạng thái job / task |
| `Settings.gs` | Cài đặt hệ thống |
| `Templates.gs` | Mẫu công đoạn (task template) |
| `Jobs.gs` | CRUD đơn hàng |
| `Tasks.gs` | CRUD công đoạn + logic tuần tự |
| `Images.gs` | Upload / xóa ảnh Drive |
| `Dashboard.gs` | Thống kê + dữ liệu khởi tạo |
| `Migration.gs` | Import dữ liệu từ file Excel cũ |
| `index.html` | Toàn bộ giao diện (SPA) |

## Triển khai

1. Tạo Google Spreadsheet mới
2. Vào **Extensions → Apps Script**, copy toàn bộ file `.gs` và `index.html` vào project
3. Chạy `setupSheets()` để tạo cấu trúc sheet
4. Deploy → **New deployment** → Web app → Execute as: Me → Who has access: Anyone
5. Lấy URL web app và chia sẻ cho nhân viên

## Import dữ liệu cũ từ Excel

1. Upload file Excel vào Google Drive, mở bằng Google Sheets
2. Vào Apps Script của file đó, copy toàn bộ `.gs` và `index.html`
3. Chạy `migrateFromXlsx()` — tự động đọc sheet và tạo dữ liệu
4. Nếu file Excel > 50MB, ảnh sẽ không được import tự động.  
   Dùng script `extract_images.py` để trích xuất ảnh cục bộ, upload lên Drive,  
   rồi chạy `migrateImages('FOLDER_ID')` để gán ảnh vào từng đơn.

## Yêu cầu

- Tài khoản Google Workspace hoặc Google cá nhân
- Quyền tạo Google Spreadsheet và Google Drive folder
