# 📁 Hướng dẫn Deploy Add-in qua Shared Folder (Word 2026)

## ✅ Giải pháp: Ép Word nhận Add-in qua Thư mục chia sẻ + GitHub Pages

Khi Word 2026 bị lỗi tài khoản hoặc khóa tính năng "Upload My Add-in", cách duy nhất là dùng **Shared Folder** (Thư mục chia sẻ) với manifest.xml trỏ đến **GitHub Pages**.

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

**Lưu ý:** Manifest.xml phải có nội dung sau (dùng GitHub Pages URLs):
```xml
<IconUrl DefaultValue="https://itdoanh.github.io/stvb/assets/icon-32.png"/>
<HighResolutionIconUrl DefaultValue="https://itdoanh.github.io/stvb/assets/icon-80.png"/>
<SourceLocation DefaultValue="https://itdoanh.github.io/stvb/taskpane.html"/>
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

Mỗi khi bạn cập nhật code:

1. **Push lên GitHub**
   ```bash
   git add .
   git commit -m "Update: [mô tả thay đổi]"
   git push origin main
   ```

2. **Chờ GitHub Pages build** (1-2 phút)

3. **Trong Word**, nhấn **Refresh** trong tab SHARED FOLDER

4. Add-in sẽ tự động load phiên bản mới từ GitHub Pages

---

## 📝 Manifest cho GitHub Pages

Đảm bảo manifest.xml dùng GitHub Pages URLs:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides/1.1"
           xsi:type="TaskPaneApp">

  <Id>f3e8a1b2-c4d5-4e6f-8a9b-0c1d2e3f4a5b</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>VietNam Document Editor</ProviderName>
  <DefaultLocale>vi-VN</DefaultLocale>
  <DisplayName DefaultValue="Soạn thảo văn bản Việt Nam"/>
  <Description DefaultValue="Công cụ soạn thảo văn bản chuẩn Việt Nam"/>
  <IconUrl DefaultValue="https://itdoanh.github.io/stvb/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://itdoanh.github.io/stvb/assets/icon-80.png"/>
  <SupportUrl DefaultValue="https://itdoanh.github.io/stvb/"/>

  <Hosts>
    <Host Name="Document"/>
  </Hosts>

  <DefaultSettings>
    <SourceLocation DefaultValue="https://itdoanh.github.io/stvb/taskpane.html"/>
  </DefaultSettings>

  <Permissions>ReadWriteDocument</Permissions>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <!-- ... rest of manifest ... -->
  </VersionOverrides>
</OfficeApp>
```

---

## ⚠️ Troubleshooting

| Vấn đề | Giải pháp |
|-------|----------|
| Shared Folder không xuất hiện | Khởi động lại Word, kiểm tra Network Path |
| "Cannot connect to catalog" | Kiểm tra manifest.xml có trong `C:\WordAddins` không |
| Add-in không load | Kiểm tra GitHub Pages URLs có đúng không, refresh tab |
| Lỗi certificate | Chấp nhận certificate khi Word yêu cầu |
| Manifest không hợp lệ | Kiểm tra XML syntax, icon URLs |

---

## 💡 Ưu điểm của cách này

✅ **Không cần local server** - Dùng GitHub Pages  
✅ **Không cần AppSource** - Deploy trực tiếp  
✅ **Không cần Microsoft 365 Admin** - Chỉ cần Shared Folder  
✅ **Dễ cập nhật** - Push lên GitHub, Word tự load phiên bản mới  
✅ **Hoạt động 100% với Word 2026** - Phương pháp chính thức của Microsoft  

---

## 🎓 Kết luận

**Shared Folder + GitHub Pages** là cách tốt nhất để:
- Deploy add-in mà không cần local server
- Phát triển và test add-in
- Chia sẻ add-in với team
- Cập nhật add-in dễ dàng

Đây là phương pháp được Microsoft khuyến nghị cho developer!
