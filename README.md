# Hệ thống phân chia phòng ký túc xá

## Mô tả
Hệ thống web giúp tự động phân chia phòng ký túc xá cho sinh viên dựa trên giới tính và chỗ trống của phòng.

## Tính năng
1. **Upload file Excel**: Upload danh sách phòng và danh sách sinh viên
2. **Tải file mẫu**: Tải các file mẫu để chuẩn bị dữ liệu
3. **Phân chia tự động**: Tự động phân chia phòng theo giới tính
4. **Xuất kết quả**: Tải file Excel chứa kết quả phân chia

## Cấu trúc file dữ liệu

### File danh sách phòng
| Cột | Mô tả |
|-----|-------|
| Khu | Tên khu (A, B, C...) |
| Tên phòng | Mã phòng (A101, B102...) |
| Giới tính | Nam hoặc Nữ |
| Chỗ ở thực tế | Số chỗ trống hiện tại |

### File danh sách sinh viên
| Cột | Mô tả |
|-----|-------|
| Mã sinh viên | Mã số sinh viên |
| Họ và tên | Họ tên đầy đủ |
| Giới tính | Nam hoặc Nữ |
| Lớp | Lớp học |
| Tên phòng | Để trống nếu chưa phân chia, điền nếu đã phân chia thủ công |
| Khu | Để trống nếu chưa phân chia, điền nếu đã phân chia thủ công |

## Cài đặt và chạy

### Yêu cầu
- Node.js (version 14+)
- npm

### Cài đặt
```bash
npm install
```

### Tạo file mẫu
```bash
node create-templates.js
```

### Chạy ứng dụng
```bash
# Chế độ development
npm run dev

# Chế độ production
npm start
```

Ứng dụng sẽ chạy tại: http://localhost:3000

## Hướng dẫn sử dụng

### Cách 1: Sử dụng file template có sẵn
1. **Kiểm tra file template**: Ứng dụng sẽ đọc dữ liệu từ 2 file có sẵn:
   - `uploads/danh_sach_phong.xlsx`
   - `uploads/danh_sach_sinh_vien.xlsx`

2. **Tải template**: Click "Tải template" để download file, chỉnh sửa dữ liệu theo nhu cầu

3. **Cập nhật file**: Copy file đã chỉnh sửa vào thư mục `uploads/` (ghi đè file cũ)

4. **Làm mới**: Click "Làm mới thông tin" để kiểm tra dữ liệu mới

### Cách 2: Upload file mới qua giao diện web
1. **Upload file**: Sử dụng tính năng drag & drop hoặc click để chọn file Excel mới

2. **Cập nhật**: Click "Upload và Cập Nhật Files" để ghi đè file hiện tại

3. **Kiểm tra**: Hệ thống tự động làm mới thông tin file sau khi upload

### Xử lý dữ liệu
5. **Phân chia**: Click "Thực hiện phân chia phòng" để xử lý

6. **Tải kết quả**: Click "Tải kết quả phân chia" để download file Excel kết quả

## Quy tắc phân chia

- Sinh viên đã có thông tin phòng trong file sẽ được giữ nguyên
- Sinh viên chưa có phòng sẽ được phân chia tự động:
  - Nam vào phòng nam có chỗ trống
  - Nữ vào phòng nữ có chỗ trống
- Mỗi lần phân chia sẽ giảm số "Chỗ ở thực tế" của phòng đi 1

## Cấu trúc thư mục
```
├── app.js              # File chính của ứng dụng
├── package.json        # Thông tin project và dependencies
├── create-templates.js # Script tạo file mẫu
├── views/              # Template EJS
│   └── index.ejs       # Giao diện chính
├── public/             # Static files
├── templates/          # File mẫu Excel
├── uploads/            # File upload từ user
└── data/               # File kết quả sau xử lý
```

## API Endpoints

- `GET /` - Trang chủ
- `POST /upload` - Upload file mới (ghi đè file hiện tại)
- `GET /download/:type` - Tải file từ thư mục uploads (template hiện tại)
  - `/download/template-phong` - Tải file phòng hiện tại
  - `/download/template-sinh-vien` - Tải file sinh viên hiện tại
  - `/download/result` - Tải file kết quả
- `GET /template/:type` - Tải file template gốc từ thư mục templates
  - `/template/phong` - Tải file mẫu danh sách phòng
  - `/template/sinh-vien` - Tải file mẫu danh sách sinh viên
- `GET /api/check-templates` - Kiểm tra trạng thái file template
- `GET /api/list-templates` - Liệt kê tất cả file template có sẵn
- `POST /api/phan-chia-phong` - Xử lý phân chia phòng

## Lưu ý

- File upload phải có định dạng .xlsx hoặc .xls
- Tên cột trong file Excel phải khớp chính xác với mẫu
- Kiểm tra encoding UTF-8 cho tiếng Việt
- Backup dữ liệu trước khi xử lý
