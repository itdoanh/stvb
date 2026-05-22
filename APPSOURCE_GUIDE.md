# 📱 Hướng dẫn Đăng ký Add-in trên AppSource (Word 2026)

## ⚠️ Vấn đề
Word 2026 không hỗ trợ Trusted Catalogs qua UI hoặc Registry. Cách duy nhất để deploy add-in là:
1. **AppSource** (Microsoft Store cho Office)
2. **Centralized Deployment** (yêu cầu Microsoft 365 Admin)

## ✅ Giải pháp: Đăng ký trên AppSource

### Bước 1: Tạo tài khoản Microsoft Partner Center
1. Vào https://partner.microsoft.com
2. Nhấn **Sign in** → Đăng nhập bằng Microsoft Account
3. Nếu chưa có, tạo Microsoft Account mới
4. Hoàn tất profile

### Bước 2: Đăng ký Developer
1. Vào **Dashboard** → **Account settings**
2. Chọn **Developer** → **Enroll**
3. Chọn **Individual** hoặc **Company**
4. Điền thông tin cá nhân/công ty
5. Thanh toán phí đăng ký (~$99 USD/năm)

### Bước 3: Tạo App Submission
1. Vào **Office Store** → **Create new app**
2. Chọn **Word Add-in**
3. Điền thông tin:
   - **App name**: Soạn thảo văn bản Việt Nam
   - **Description**: Công cụ soạn thảo văn bản chuẩn Việt Nam theo Nghị định 30/2020/NĐ-CP
   - **Category**: Productivity
   - **Support URL**: https://github.com/itdoanh/stvb

### Bước 4: Upload Manifest
1. Upload file `manifest.xml` (cập nhật URL GitHub Pages)
2. Upload icon (16x16, 32x32, 80x80 PNG)
3. Điền mô tả chi tiết

### Bước 5: Submit for Review
1. Nhấn **Submit**
2. Microsoft sẽ review trong 3-5 ngày
3. Nếu pass, add-in sẽ xuất hiện trên AppSource

### Bước 6: Người dùng cài đặt
1. Mở Word
2. **Insert** → **Get Add-ins** → **Search**
3. Tìm "Soạn thảo văn bản Việt Nam"
4. Nhấn **Add**

---

## 💰 Chi phí
- Đăng ký Developer: ~$99 USD/năm
- Đăng ký app: Miễn phí
- Review: Miễn phí

---

## 🔗 Link hữu ích
- Partner Center: https://partner.microsoft.com
- Office Store: https://appsource.microsoft.com/office
- Submission Guide: https://docs.microsoft.com/office/dev/store/submit-to-appsource-via-partner-center

---

## ⚡ Giải pháp tạm thời (Local Testing)
Nếu bạn chỉ muốn test locally:
1. Chạy local server: `http-server -p 3000`
2. Cập nhật manifest.xml dùng `http://localhost:3000`
3. Thêm Trusted Catalog: `http://localhost:3000/manifest.xml`
4. Khởi động lại Word

**Lưu ý**: Cách này chỉ hoạt động khi server đang chạy, không thể chia sẻ với người khác.

---

## 📝 Manifest cho GitHub Pages (Production)
```xml
<IconUrl DefaultValue="https://itdoanh.github.io/stvb/assets/icon-32.png"/>
<HighResolutionIconUrl DefaultValue="https://itdoanh.github.io/stvb/assets/icon-80.png"/>
<SourceLocation DefaultValue="https://itdoanh.github.io/stvb/taskpane.html"/>
```

---

## 🎯 Khuyến nghị
1. **Ngắn hạn**: Dùng local server để test
2. **Dài hạn**: Đăng ký trên AppSource để phân phối rộng
3. **Doanh nghiệp**: Dùng Centralized Deployment (yêu cầu Microsoft 365 Admin)
