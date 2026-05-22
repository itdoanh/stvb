# 📋 Hướng dẫn Upload Add-in lên Word 365 (2026)

## ⚠️ Vấn đề hiện tại
- Word 365 phiên bản mới nhất (2026) **không có nút "Upload My Add-ins"** trong Developer tab
- Cần phải host add-in trên web server công cộng
- Đăng ký qua **Trusted Catalogs** hoặc **Centralized Deployment**

---

## 🔧 Giải pháp 1: Dùng Trusted Catalogs (Nhanh nhất)

### Bước 1: Host Add-in lên Web Server
Chọn một trong các tùy chọn:

#### A. Dùng GitHub Pages (Miễn phí)
```bash
# 1. Tạo repo trên GitHub
# 2. Push các file lên:
#    - manifest.xml
#    - taskpane.html
#    - taskpane.css
#    - taskpane.js
#    - assets/icon-*.png

# 3. Enable GitHub Pages trong Settings
# 4. URL sẽ là: https://yourusername.github.io/repo-name
```

#### B. Dùng Netlify (Miễn phí)
```bash
# 1. Kết nối GitHub repo
# 2. Deploy tự động
# 3. URL: https://your-site.netlify.app
```

#### C. Dùng Azure Static Web Apps (Miễn phí tier)
```bash
# 1. Tạo Static Web App trên Azure
# 2. Deploy từ GitHub
# 3. URL: https://your-app.azurestaticapps.net
```

#### D. Dùng Vercel (Miễn phí)
```bash
# 1. Import GitHub repo
# 2. Deploy tự động
# 3. URL: https://your-project.vercel.app
```

### Bước 2: Cập nhật manifest.xml
Thay thế `https://localhost:3000` bằng URL công cộng:

```xml
<!-- Trước -->
<IconUrl DefaultValue="https://localhost:3000/assets/icon-32.png" />
<HighResolutionIconUrl DefaultValue="https://localhost:3000/assets/icon-80.png" />
<SupportUrl DefaultValue="https://localhost:3000/help.html" />
<SourceLocation DefaultValue="https://localhost:3000/taskpane.html" />

<!-- Sau -->
<IconUrl DefaultValue="https://yourdomain.com/assets/icon-32.png" />
<HighResolutionIconUrl DefaultValue="https://yourdomain.com/assets/icon-80.png" />
<SupportUrl DefaultValue="https://yourdomain.com/help.html" />
<SourceLocation DefaultValue="https://yourdomain.com/taskpane.html" />
```

Cập nhật trong Resources section:
```xml
<bt:Image id="Icon.16x16" DefaultValue="https://yourdomain.com/assets/icon-16.png"/>
<bt:Image id="Icon.32x32" DefaultValue="https://yourdomain.com/assets/icon-32.png"/>
<bt:Image id="Icon.80x80" DefaultValue="https://yourdomain.com/assets/icon-80.png"/>
<bt:Url id="Taskpane.Url" DefaultValue="https://yourdomain.com/taskpane.html" />
```

### Bước 3: Đăng ký Trusted Catalog trong Word
1. Mở **Word** → **File** → **Options**
2. Chọn **Trust Center** → **Trust Center Settings**
3. Chọn **Trusted Add-in Catalogs**
4. Nhập URL manifest:
   ```
   https://yourdomain.com/manifest.xml
   ```
5. Tick **"Show in Menu"**
6. Nhấn **Add catalog** → **OK**
7. **Khởi động lại Word**

### Bước 4: Sử dụng Add-in
- Mở Word
- Vào **Insert** → **Get Add-ins** → **My Add-ins** → **My Organization**
- Add-in sẽ xuất hiện trong danh sách
- Nhấn **Add** để cài đặt

---

## 🏢 Giải pháp 2: Centralized Deployment (Cho doanh nghiệp)

### Yêu cầu
- Tài khoản Microsoft 365 Admin
- Quyền quản trị

### Các bước
1. Đăng nhập vào **Microsoft 365 Admin Center**
2. Vào **Settings** → **Integrated apps** → **Upload custom apps**
3. Upload file `manifest.xml`
4. Chọn người dùng/nhóm để deploy
5. Chờ deployment hoàn tất (có thể mất 24 giờ)

---

## 📦 Giải pháp 3: AppSource (Cho phân phối rộng)

### Yêu cầu
- Tài khoản Microsoft Partner Center
- Phí đăng ký: ~$99/năm
- Quá trình review: 3-5 ngày

### Các bước
1. Đăng ký tại [Partner Center](https://partner.microsoft.com)
2. Tạo app submission
3. Upload manifest.xml + icon + mô tả
4. Chờ Microsoft review
5. Sau khi phê duyệt, add-in sẽ xuất hiện trên AppSource

---

## ✅ Kiểm tra Add-in hoạt động

### Test trên máy local trước
```bash
# 1. Khởi động web server
npx serve -l 3000

# 2. Mở Word
# 3. File → Options → Trust Center → Trusted Add-in Catalogs
# 4. Thêm: http://localhost:3000/manifest.xml
# 5. Khởi động lại Word
# 6. Kiểm tra xem add-in có xuất hiện không
```

### Troubleshooting
| Vấn đề | Giải pháp |
|-------|----------|
| Add-in không xuất hiện | Khởi động lại Word, xóa cache |
| Lỗi SSL/HTTPS | Dùng `http://` cho local, `https://` cho production |
| Manifest không load | Kiểm tra URL có đúng không, file có tồn tại không |
| Icon không hiển thị | Kiểm tra đường dẫn icon trong manifest |

---

## 🎯 Khuyến nghị

**Nhanh nhất (5 phút):**
1. Push code lên GitHub
2. Enable GitHub Pages
3. Cập nhật manifest.xml
4. Thêm Trusted Catalog trong Word

**An toàn nhất (cho doanh nghiệp):**
- Dùng Centralized Deployment qua Microsoft 365 Admin Center

**Phân phối rộng nhất:**
- Đăng ký trên AppSource (mất thời gian nhưng có độ tin cậy cao)

---

## 📝 Ghi chú
- Manifest.xml phải có HTTPS (trừ localhost)
- Icon phải là PNG, kích thước: 16x16, 32x32, 80x80
- Add-in cần quyền ReadWriteDocument để chỉnh sửa tài liệu
- Mỗi lần cập nhật manifest, cần khởi động lại Word
