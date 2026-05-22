# 📁 Hướng dẫn Deploy Add-in qua Shared Folder (Word 2026)

## ✅ Giải pháp: Ép Word nhận Add-in qua Thư mục chia sẻ

Khi Word 2026 bị lỗi tài khoản hoặc khóa tính năng "Upload My Add-in", cách duy nhất là dùng **Shared Folder** (Thư mục chia sẻ).

---

## 🔧 Bước 1: Tạo thư mục chia sẻ

### 1.1 Tạo thư mục
```cmd
mkdir C:\WordAddins
```

### 1.2 Copy manifest.xml vào thư mục
```cmd
copy manifest.xml C:\WordAddins\manifest.xml
```

### 1.3 Share thư mục
1. Mở **File Explorer**
2. Chuột phải vào `C:\WordAddins` → **Properties**
3. Chuyển sang tab **Sharing**
4. Nhấn **Share...**
5. Chọn tài khoản của bạn (hoặc `Everyone`)
6. Nhấn **Add** → **Share**
7. **Copy Network Path** (ví dụ: `\\DOANH-PC\WordAddins`)

---

## 📋 Bước 2: Thêm vào Word Trust Center

### 2.1 Mở Word Options
- **File** → **Options**

### 2.2 Mở Trust Center Settings
- Chọn **Trust Center** (bên trái)
- Nhấn **Trust Center Settings...**

### 2.3 Thêm Shared Folder
- Chọn **Trusted Add-in Catalogs**
- Trong ô **Catalog URL**, dán Network Path:
  ```
  \\DOANH-PC\WordAddins
  ```
- Nhấn **Add catalog**
- Tick ✓ **Show in Menu**
- Nhấn **OK**

### 2.4 Khởi động lại Word
- **Đóng Word hoàn toàn**
- **Mở Word lại**

---

## 🎯 Bước 3: Sử dụng Add-in

1. Mở Word
2. Vào tab **Home** (Trang chủ)
3. Nhấn **Add-ins** (hoặc **Get Add-ins**)
4. Bạn sẽ thấy tab **SHARED FOLDER** xuất hiện
5. Click vào **SHARED FOLDER**
6. Tìm **"Soạn thảo văn bản Việt Nam"**
7. Nhấn **Add**

---

## 🚀 Bước 4: Phát triển & Cập nhật

Mỗi khi bạn cập nhật `manifest.xml`:

1. **Cập nhật file** trong `C:\WordAddins\manifest.xml`
2. **Khởi động lại Word**
3. Add-in sẽ tự động load phiên bản mới

---

## 📝 Manifest cho Local Development

Đảm bảo manifest.xml dùng HTTPS localhost:

```xml
<IconUrl DefaultValue="https://localhost:3000/assets/icon-32.png"/>
<HighResolutionIconUrl DefaultValue="https://localhost:3000/assets/icon-80.png"/>
<SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
```

---

## 🔗 Chạy Local HTTPS Server

```bash
node https-server.js
```

Server sẽ chạy tại: `https://localhost:3000`

---

## ⚠️ Troubleshooting

| Vấn đề | Giải pháp |
|-------|----------|
| Shared Folder không xuất hiện | Khởi động lại Word, kiểm tra Network Path |
| Add-in không load | Kiểm tra HTTPS server đang chạy không |
| Lỗi certificate | Chấp nhận certificate tự ký khi Word yêu cầu |
| Manifest không hợp lệ | Kiểm tra XML syntax, icon URLs |

---

## 💡 Mẹo cho Developer

Nếu dùng Office Add-in CLI:

```bash
npx office-addin-debugging start manifest.xml word
```

Lệnh này sẽ tự động:
- Tạo Shared Folder
- Thêm vào Trust Center
- Khởi động Word
- Load add-in

---

## 🎓 Kết luận

**Shared Folder** là cách chính thức để deploy add-in khi:
- Gặp lỗi tài khoản Word
- Không có Microsoft 365 Admin
- Không muốn đăng ký AppSource
- Chỉ cần test/phát triển locally

Đây là phương pháp được Microsoft khuyến nghị cho developer!
