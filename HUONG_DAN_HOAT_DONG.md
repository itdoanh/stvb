# ✅ Hướng dẫn Hoàn tất: Thêm Add-in vào Word 365

## 📍 Bước 1: Enable GitHub Pages

1. Vào **https://github.com/itdoanh/stvb**
2. Nhấn **Settings** (tab cài đặt)
3. Chọn **Pages** (bên trái)
4. Trong **Source**, chọn:
   - Branch: `main`
   - Folder: `/ (root)`
5. Nhấn **Save**
6. Chờ 1-2 phút để GitHub Pages build xong
7. Bạn sẽ thấy thông báo: "Your site is live at https://itdoanh.github.io/stvb/"

---

## 🔗 Bước 2: Kiểm tra URL hoạt động

Mở trình duyệt và kiểm tra các URL này:
- https://itdoanh.github.io/stvb/manifest.xml ✅
- https://itdoanh.github.io/stvb/taskpane.html ✅
- https://itdoanh.github.io/stvb/assets/icon-32.png ✅

Nếu tất cả đều hiển thị, bạn có thể tiếp tục.

---

## 📋 Bước 3: Thêm Trusted Catalog trong Word

### Trên Windows:

1. **Mở Word 365**
2. Nhấn **File** → **Options**
3. Chọn **Trust Center** (bên trái)
4. Nhấn **Trust Center Settings**
5. Chọn **Trusted Add-in Catalogs** (bên trái)
6. Trong ô **Catalog Url**, nhập:
   ```
   https://itdoanh.github.io/stvb/manifest.xml
   ```
7. Tick ✓ **"Show in Menu"**
8. Nhấn **Add catalog**
9. Nhấn **OK**
10. **Khởi động lại Word hoàn toàn** (đóng và mở lại)

### Trên Mac:

1. **Mở Word**
2. Nhấn **Word** → **Preferences**
3. Chọn **Trust Center**
4. Nhấn **Trusted Add-in Catalogs**
5. Nhập URL manifest:
   ```
   https://itdoanh.github.io/stvb/manifest.xml
   ```
6. Tick ✓ **"Show in Menu"**
7. Nhấn **Add catalog**
8. **Khởi động lại Word**

---

## 🎯 Bước 4: Sử dụng Add-in

**Cách 1: Tự động xuất hiện (Sau khi thêm Trusted Catalog)**
1. **Mở Word** (hoặc tạo tài liệu mới)
2. Vào tab **Home** (Trang chủ)
3. Tìm nhóm **"Văn bản VN"** trong ribbon
4. Nhấn nút **"Soạn thảo văn bản Việt Nam"**
5. Task Pane sẽ mở bên phải màn hình

**Cách 2: Nếu không thấy trong Home ribbon**
1. Vào **File** → **Options** → **Trust Center** → **Trust Center Settings**
2. Chọn **Trusted Add-in Catalogs**
3. Kiểm tra URL manifest đã được thêm chưa
4. Nếu chưa, thêm: `https://itdoanh.github.io/stvb/manifest.xml`
5. Khởi động lại Word
6. Add-in sẽ tự động xuất hiện trong ribbon **Home**

**Lưu ý:** Word 2026 không có nút "Get Add-ins" trong Insert tab. Add-in sẽ tự động xuất hiện trong Home ribbon sau khi thêm Trusted Catalog.

---

## ⚠️ Troubleshooting

| Vấn đề | Giải pháp |
|-------|----------|
| Add-in không xuất hiện | Khởi động lại Word hoàn toàn, xóa cache |
| Lỗi "Cannot reach the add-in" | Kiểm tra URL manifest có đúng không |
| Icon không hiển thị | Kiểm tra đường dẫn icon trong manifest.xml |
| Manifest không load | Mở URL manifest trong trình duyệt để kiểm tra |
| "This add-in is no longer available" | Xóa add-in và cài lại |

---

## 🔄 Cập nhật Add-in

Mỗi khi bạn cập nhật code:

1. **Commit và push lên GitHub**
   ```bash
   git add .
   git commit -m "Update: [mô tả thay đổi]"
   git push origin main
   ```

2. **Chờ GitHub Pages build** (1-2 phút)

3. **Trong Word**, nhấn **F9** hoặc:
   - **File** → **Options** → **Trust Center** → **Trusted Add-in Catalogs**
   - Xóa catalog cũ
   - Thêm lại URL manifest
   - Khởi động lại Word

---

## 📞 Liên hệ hỗ trợ

- **GitHub Issues**: https://github.com/itdoanh/stvb/issues
- **Microsoft Support**: https://support.microsoft.com/office

---

## ✨ Chúc mừng!

Add-in của bạn đã sẵn sàng sử dụng trên Word 365. Bạn có thể:
- ✅ Thiết lập trang chuẩn Việt Nam
- ✅ Tạo styles đề mục
- ✅ Đánh số trang
- ✅ Tạo caption hình ảnh/bảng
- ✅ Tạo mục lục tự động
- ✅ Sử dụng các công cụ nâng cao

Thưởng thức công cụ soạn thảo văn bản chuẩn Việt Nam! 🇻🇳
