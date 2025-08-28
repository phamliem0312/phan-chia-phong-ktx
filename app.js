const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 4000;

// Cấu hình EJS làm view engine
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Middleware
app.use(express.static('public'));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Cấu hình multer để upload file
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    // Tạo thư mục temp để lưu file tạm thời
    const tempDir = path.join(__dirname, 'temp');
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }
    cb(null, tempDir);
  },
  filename: function (req, file, cb) {
    // Tạo tên file tạm thời với timestamp
    const timestamp = Date.now();
    const originalName = Buffer.from(file.originalname, 'latin1').toString('utf8');
    cb(null, `${timestamp}_${originalName}`);
  }
});

const upload = multer({ 
  storage: storage,
  fileFilter: function (req, file, cb) {
    if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
        file.mimetype === 'application/vnd.ms-excel') {
      cb(null, true);
    } else {
      cb(new Error('Chỉ chấp nhận file Excel (.xlsx, .xls)'), false);
    }
  }
});

// Routes
app.get('/', (req, res) => {
  res.render('index');
});

// Upload files - cập nhật file trong thư mục uploads
app.post('/upload', upload.fields([
  { name: 'danhSachPhong', maxCount: 1 },
  { name: 'danhSachSinhVien', maxCount: 1 }
]), (req, res) => {
  try {
    let message = 'Upload thành công: ';
    let uploadedFiles = [];
    
    if (req.files.danhSachPhong) {
      const originalName = 'danh_sach_phong.xlsx';
      const uploadedFile = req.files.danhSachPhong[0];
      const targetPath = path.join(__dirname, 'uploads', originalName);
      
      // Di chuyển file upload đè lên file cũ
      fs.renameSync(uploadedFile.path, targetPath);
      uploadedFiles.push('danh sách phòng');
      message += `Danh sách phòng `;
    }
    
    if (req.files.danhSachSinhVien) {
      const originalName = 'danh_sach_sinh_vien.xlsx';
      const uploadedFile = req.files.danhSachSinhVien[0];
      const targetPath = path.join(__dirname, 'uploads', originalName);
      
      // Di chuyển file upload đè lên file cũ
      fs.renameSync(uploadedFile.path, targetPath);
      uploadedFiles.push('danh sách sinh viên');
      message += `Danh sách sinh viên `;
    }
    
    res.json({ 
      success: true, 
      message: message,
      uploadedFiles: uploadedFiles
    });
  } catch (error) {
    console.error('Lỗi upload:', error);
    res.status(500).json({ 
      success: false, 
      message: 'Lỗi upload file: ' + error.message 
    });
  }
});

// API tải file template từ thư mục uploads
app.get('/download/:type', (req, res) => {
  const type = req.params.type;
  let filePath;
  
  switch (type) {
    case 'template-phong':
      filePath = path.join(__dirname, 'uploads', 'danh_sach_phong.xlsx');
      break;
    case 'template-sinh-vien':
      filePath = path.join(__dirname, 'uploads', 'danh_sach_sinh_vien.xlsx');
      break;
    case 'result':
      filePath = path.join(__dirname, 'data', 'du_lieu_chia_phong.xlsx');
      break;
    default:
      return res.status(404).json({ message: 'File không tồn tại' });
  }
  
  if (fs.existsSync(filePath)) {
    res.download(filePath);
  } else {
    res.status(404).json({ message: 'File không tồn tại' });
  }
});

// API tải file template gốc từ thư mục templates
app.get('/template/:type', (req, res) => {
  const type = req.params.type;
  let filePath;
  let fileName;
  
  switch (type) {
    case 'phong':
      filePath = path.join(__dirname, 'templates', 'danh_sach_phong.xlsx');
      fileName = 'mau_danh_sach_phong.xlsx';
      break;
    case 'sinh-vien':
      filePath = path.join(__dirname, 'templates', 'danh_sach_sinh_vien.xlsx');
      fileName = 'mau_danh_sach_sinh_vien.xlsx';
      break;
    default:
      return res.status(404).json({ message: 'File template không tồn tại' });
  }
  
  if (fs.existsSync(filePath)) {
    res.download(filePath, fileName, (err) => {
      if (err) {
        console.error('Lỗi tải file:', err);
        res.status(500).json({ message: 'Lỗi tải file template' });
      }
    });
  } else {
    res.status(404).json({ message: 'File template không tồn tại' });
  }
});

// API liệt kê file template có sẵn
app.get('/api/list-templates', (req, res) => {
  try {
    const templatesPath = path.join(__dirname, 'templates');
    const uploadsPath = path.join(__dirname, 'uploads');
    
    let templates = {
      templates: [],
      uploads: []
    };
    
    // Kiểm tra thư mục templates
    if (fs.existsSync(templatesPath)) {
      const templateFiles = fs.readdirSync(templatesPath).filter(file => 
        file.endsWith('.xlsx') || file.endsWith('.xls')
      );
      
      templates.templates = templateFiles.map(file => {
        const filePath = path.join(templatesPath, file);
        const stats = fs.statSync(filePath);
        return {
          name: file,
          size: stats.size,
          modified: stats.mtime,
          path: `/template/${file.includes('phong') ? 'phong' : 'sinh-vien'}`
        };
      });
    }
    
    // Kiểm tra thư mục uploads
    if (fs.existsSync(uploadsPath)) {
      const uploadFiles = fs.readdirSync(uploadsPath).filter(file => 
        file.endsWith('.xlsx') || file.endsWith('.xls')
      );
      
      templates.uploads = uploadFiles.map(file => {
        const filePath = path.join(uploadsPath, file);
        const stats = fs.statSync(filePath);
        return {
          name: file,
          size: stats.size,
          modified: stats.mtime,
          path: `/download/${file.includes('phong') ? 'template-phong' : 'template-sinh-vien'}`
        };
      });
    }
    
    res.json({ success: true, data: templates });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

// API kiểm tra file template
app.get('/api/check-templates', (req, res) => {
  try {
    const phongFilePath = path.join(__dirname, 'uploads', 'danh_sach_phong.xlsx');
    const sinhVienFilePath = path.join(__dirname, 'uploads', 'danh_sach_sinh_vien.xlsx');
    
    const phongExists = fs.existsSync(phongFilePath);
    const sinhVienExists = fs.existsSync(sinhVienFilePath);
    
    let info = {
      phongFile: { exists: phongExists, count: 0 },
      sinhVienFile: { exists: sinhVienExists, count: 0 }
    };
    
    if (phongExists) {
      try {
        const phongWorkbook = XLSX.readFile(phongFilePath);
        const phongSheet = phongWorkbook.Sheets[phongWorkbook.SheetNames[0]];
        const danhSachPhong = XLSX.utils.sheet_to_json(phongSheet);
        info.phongFile.count = danhSachPhong.length;
      } catch (error) {
        info.phongFile.error = 'Lỗi đọc file: ' + error.message;
      }
    }
    
    if (sinhVienExists) {
      try {
        const sinhVienWorkbook = XLSX.readFile(sinhVienFilePath);
        const sinhVienSheet = sinhVienWorkbook.Sheets[sinhVienWorkbook.SheetNames[0]];
        const danhSachSinhVien = XLSX.utils.sheet_to_json(sinhVienSheet);
        info.sinhVienFile.count = danhSachSinhVien.length;
      } catch (error) {
        info.sinhVienFile.error = 'Lỗi đọc file: ' + error.message;
      }
    }
    
    res.json({ success: true, data: info });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

// API xử lý phân chia phòng
app.post('/api/phan-chia-phong', async (req, res) => {
  try {
    // Đọc file template từ thư mục uploads (file có sẵn)
    const phongFilePath = path.join(__dirname, 'uploads', 'danh_sach_phong.xlsx');
    const sinhVienFilePath = path.join(__dirname, 'uploads', 'danh_sach_sinh_vien.xlsx');
    
    if (!fs.existsSync(phongFilePath) || !fs.existsSync(sinhVienFilePath)) {
      return res.status(400).json({ 
        success: false, 
        message: 'Không tìm thấy file template. Vui lòng kiểm tra file danh_sach_phong.xlsx và danh_sach_sinh_vien.xlsx trong thư mục uploads' 
      });
    }
    
    // Đọc dữ liệu từ file Excel
    const phongWorkbook = XLSX.readFile(phongFilePath);
    const sinhVienWorkbook = XLSX.readFile(sinhVienFilePath);
    
    const phongSheet = phongWorkbook.Sheets[phongWorkbook.SheetNames[0]];
    const sinhVienSheet = sinhVienWorkbook.Sheets[sinhVienWorkbook.SheetNames[0]];
    
    const danhSachPhong = XLSX.utils.sheet_to_json(phongSheet);
    const danhSachSinhVien = XLSX.utils.sheet_to_json(sinhVienSheet);
    
    // Xử lý phân chia phòng
    const ketQuaPhanChia = await phanChiaPhong(danhSachPhong, danhSachSinhVien);
    
    // Tạo file kết quả
    const outputPath = path.join(__dirname, 'data', 'du_lieu_chia_phong.xlsx');
    await taoFileKetQua(ketQuaPhanChia, outputPath);
    
    res.json({ 
      success: true, 
      message: 'Phân chia phòng thành công!',
      data: ketQuaPhanChia
    });
    
  } catch (error) {
    console.error('Lỗi phân chia phòng:', error);
    res.status(500).json({ 
      success: false, 
      message: 'Lỗi xử lý: ' + error.message 
    });
  }
});

// Hàm xử lý phân chia phòng
async function phanChiaPhong(danhSachPhong, danhSachSinhVien) {
  // Tạo bản sao để không thay đổi dữ liệu gốc
  const phongData = JSON.parse(JSON.stringify(danhSachPhong)).map(phong => {
    let phongMoi = {};
    for (let key in phong) {
      phongMoi[key.trim()] = phong[key];
    }
    return phongMoi;
  });
  const sinhVienData = JSON.parse(JSON.stringify(danhSachSinhVien)).map(sv => {
    let sinhVien = {};

    for (let key in sv) {
      sinhVien[key.trim()] = sv[key];
    }

    return sinhVien;
  });

  // Phân loại phòng theo giới tính
  const phongNam = phongData.filter(phong => 
    phong['Giới tính'] && phong['Giới tính'].trim().toLowerCase() === 'nam'
  );
  const phongNu = phongData.filter(phong => 
    phong['Giới tính'] && phong['Giới tính'].trim().toLowerCase() === 'nữ'
  );
  
  // Lọc sinh viên chưa được phân phòng
  const sinhVienChuaPhanPhong = sinhVienData.filter(sv => 
    !sv['Phòng'] || sv['Phòng'].trim() === '' || sv['Phòng'] === null
  );

  // Lọc sinh viên đã được phân phòng
  const sinhVienDaPhanPhong = sinhVienData.filter(sv => 
    sv['Phòng'] && sv['Phòng'].trim() !== '' && sv['Phòng'] !== null
  );
  
  // Phân chia sinh viên nam
  const sinhVienNam = sinhVienChuaPhanPhong.filter(sv => 
    sv['Giới tính'] && sv['Giới tính'].trim().toLowerCase() === 'nam'
  );
  
  // Phân chia sinh viên nữ  
  const sinhVienNu = sinhVienChuaPhanPhong.filter(sv => 
    sv['Giới tính'] && sv['Giới tính'].trim().toLowerCase() === 'nữ'
  );

  let students = [];

  // Phân phòng cho sinh viên nam
  for (let sinhVien of sinhVienNam) {
    const phongTrongIndex = phongNam.findIndex(phong => 
      phong['Số lượng thực tế'] && parseInt(phong['Số lượng thực tế']) > 0
    );
    
    if (phongTrongIndex !== -1) {
      sinhVien['KTX'] = phongNam[phongTrongIndex]['KTX'];
      sinhVien['Phòng'] = sinhVien['Phòng'] ?? phongNam[phongTrongIndex]['Phòng'];
      sinhVien['Khu'] = phongNam[phongTrongIndex]['Khu'];
      phongNam[phongTrongIndex]['Số lượng thực tế'] = parseInt(phongNam[phongTrongIndex]['Số lượng thực tế']) - 1;
    }

    students.push(sinhVien);
  }
  
  // Phân phòng cho sinh viên nữ
  for (let sinhVien of sinhVienNu) {
    const phongTrongIndex = phongNu.findIndex(phong => 
      phong['Số lượng thực tế'] && parseInt(phong['Số lượng thực tế']) > 0
    );
    
    if (phongTrongIndex !== -1) {
      sinhVien['KTX'] = phongNu[phongTrongIndex]['KTX'];
      sinhVien['Phòng'] = sinhVien['Phòng'] ?? phongNu[phongTrongIndex]['Phòng'];
      sinhVien['Khu'] = phongNu[phongTrongIndex]['Khu'];
      phongNu[phongTrongIndex]['Số lượng thực tế'] = parseInt(phongNu[phongTrongIndex]['Số lượng thực tế']) - 1;
    }
    
    students.push(sinhVien);
  }

  for (const sinhVien of sinhVienDaPhanPhong) {
    sinhVien['Chia thủ công'] = 'X';

    students.push(sinhVien);
  }
  
  return {
    danhSachPhong: phongData,
    danhSachSinhVien: sinhVienData,
    danhSachSinhVienDaPhanPhong: students,
    thongKe: {
      tongSinhVien: sinhVienData.length,
      sinhVienDaPhanPhong: sinhVienData.filter(sv => sv['Phòng']).length,
      sinhVienChuaPhanPhong: sinhVienData.filter(sv => !sv['Phòng']).length
    }
  };
}

// Hàm tạo file kết quả
async function taoFileKetQua(ketQua, outputPath) {
  const workbook = XLSX.utils.book_new();
  
  // Sheet danh sách sinh viên
  const sinhVienSheet = XLSX.utils.json_to_sheet(ketQua.danhSachSinhVien);
  XLSX.utils.book_append_sheet(workbook, sinhVienSheet, 'Danh sách sinh viên');
  
  // Sheet danh sách phòng
  const phongSheet = XLSX.utils.json_to_sheet(ketQua.danhSachPhong);
  XLSX.utils.book_append_sheet(workbook, phongSheet, 'Danh sách phòng');
  
  // Sheet thống kê
  const thongKeSheet = XLSX.utils.json_to_sheet([ketQua.thongKe]);
  XLSX.utils.book_append_sheet(workbook, thongKeSheet, 'Thống kê');
  
  // Ghi file
  XLSX.writeFile(workbook, outputPath);
}

// Khởi động server
app.listen(PORT, () => {
  console.log(`Server đang chạy tại http://localhost:${PORT}`);
});

module.exports = app;
