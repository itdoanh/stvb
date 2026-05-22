# 🇻🇳 Soạn thảo văn bản Việt Nam - Word Add-in

Add-in Microsoft Word hỗ trợ soạn thảo văn bản theo chuẩn **Nghị định 30/2020/NĐ-CP** về công tác văn thư và các quy chuẩn trình bày văn bản phổ biến tại Việt Nam.

---

## ✨ Tính năng chính

| Nhóm | Tính năng | Mô tả |
|------|-----------|-------|
| **📄 Thiết lập trang** | Khổ giấy A4/A3/Letter | Tự động căn chỉnh kích thước trang |
| | Lề chuẩn | Lề trái 3cm, lề phải 2cm, lề trên/dưới 2cm (tùy chỉnh) |
| | Font & dãn dòng | Times New Roman 13pt, dãn 1.5 dòng |
| **📑 Styles đề mục** | 5 cấp Heading | Heading 1-5 theo đúng quy định phân cấp |
| | Heading 1 | CHỮ IN HOA, ĐẬM, canh giữa |
| | Heading 2 | In hoa thường, đậm, đứng |
| | Heading 3 | In hoa thường, đậm, nghiêng |
| | Heading 4 | Thường, đứng |
| | Heading 5 | Thường, nghiêng |
| **🔢 Đánh số trang** | Vị trí | Canh giữa / trái / phải |
| | Bắt đầu từ số bất kỳ | Tùy chọn số trang khởi đầu |
| | Bỏ qua trang đầu | Không đánh số trang bìa |
| **🖼️ Caption** | Hình ảnh | Caption dưới hình, đánh số theo chương |
| | Bảng biểu | Caption trên bảng, đánh số theo chương |
| | Công thức | Số thứ tự căn phải trong ngoặc đơn |
| **📋 Mục lục** | Tạo/Tái lập | Tự động tạo mục lục từ Heading 1-5 |
| | Danh sách hình | Tạo danh sách hình ảnh |
| | Danh sách bảng | Tạo danh sách bảng biểu |
| | Cập nhật Fields | F9 toàn bộ văn bản |
| **🛠️ Công cụ nâng cao** | Sửa orphan headings | Đảm bảo heading không nằm một mình cuối trang |
| | Xóa khoảng trắng thừa | Chuẩn hóa space trong văn bản |
| | Chuẩn hóa paragraph spacing | Thay Enter bằng Before/After |
| | Dấu ngoặc kép chuẩn VN | Chuyển "..." thành "..." |
| | Khối ký tên | Chèn khối ký tên hành chính |
| | Đầu trang văn bản | Chèn header chuẩn cơ quan VN |

---

## 📁 Cấu trúc dự án

```
vn-word-addin/
├── manifest.xml          # Manifest đăng ký add-in với Office
├── taskpane.html         # Giao diện Task Pane
├── taskpane.css          # CSS styling
├── taskpane.js           # Logic xử lý JavaScript (Office.js)
├── assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
└── README.md
```

---

## 🚀 Hướng dẫn cài đặt

### Yêu cầu

- Microsoft Word 2016 trở lên (khuyến nghị Office 365/Microsoft 365)
- Quyền cài đặt Add-in (có thể cần quản trị viên trong môi trường doanh nghiệp)

### Cách 1: Cài đặt qua Web Server + Trusted Catalogs (Duy nhất hoạt động với Word 2026)

**Bước 1:** Khởi động web server
```bash
cd c:/code/stvb
npx serve -l 3000
```

**Bước 2:** Mở Word → **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Trusted Add-in Catalogs**

**Bước 3:** Thêm URL vào danh sách:
```
http://localhost:3000
```

**Bước 4:** Tick **"Show in Menu"** → Nhấn **Add catalog** → **OK**

**Bước 5:** Khởi động lại Word

**Bước 6:** Sau khi khởi động lại, add-in sẽ tự động xuất hiện trong **Insert** ribbon hoặc **Developer** tab

**Lưu ý:**
- Dùng `http://` không phải `https://` để tránh lỗi SSL
- Web server phải chạy liên tục khi sử dụng add-in
- Đây là cách duy nhất hoạt động với Word 2026 vì không có nút Upload trong Developer tab
- Sau khi add catalog, add-in sẽ tự động load mà không cần tìm nút Get Add-ins

### Cách 2: Cài đặt qua Production Server (Cho doanh nghiệp)

1. **Host các file** lên web server (IIS, Apache, Nginx, AWS, Azure, etc.)

2. **Cập nhật URL** trong `manifest.xml`:
   ```xml
   <SourceLocation DefaultValue="https://your-domain.com/taskpane.html" />
   <IconUrl DefaultValue="https://your-domain.com/assets/icon-32.png" />
   ```

3. **Deploy manifest** qua:
   - **Microsoft 365 Admin Center** → Settings → Integrated apps → Upload custom apps
   - Hoặc **Exchange Admin Center** → Organization → Add-ins

### Cách 3: Upload My Add-in (Word Office 365 2026)

**Bước 1:** Mở Word

**Bước 2:** Nhấn tab **Developer** (ở ribbon)

**Bước 3:** Tìm nút **"Add-ins"** hoặc **"Get Add-ins"** trong Developer tab

**Bước 4:** Chọn **"Upload Custom Add-in"** hoặc **"My Add-ins"**

**Bước 5:** Chọn file `manifest.xml` từ `C:\code\stvb\`

**Bước 6:** Nhấn **Upload**

**Bước 7:** Add-in sẽ xuất hiện trong **My Add-ins** mỗi lần mở Word

---

**Nếu không tìm thấy nút Upload trong Developer tab:**
- Thử **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Trusted Add-in Catalogs**
- Thêm: `http://localhost:3000`
- Khởi động lại Word

**Lưu ý:** Trong Word 2026, add-ins được quản lý qua **Developer tab**, không phải Insert tab.

---

## 📖 Hướng dẫn sử dụng

### 1. Thiết lập trang mới

1. Mở add-in từ tab **Home** → nhóm **Văn bản VN**
2. Mở rộng panel **"Thiết lập trang & Lề"**
3. Chọn khổ giấy, nhập lề, chọn font và dãn dòng
4. Nhấn **"Áp dụng thiết lập trang"**

### 2. Tạo đề mục chuẩn

1. Viết tiêu đề chương/mục
2. Bôi đen dòng tiêu đề
3. Nhấn **"Áp dụng Heading X"** tương ứng
4. Hoặc nhấn **"Tạo/Cập nhật Styles"** để định dạng toàn bộ heading trong văn bản

### 3. Chèn Caption

1. Đặt con trỏ ngay trên bảng / dưới hình / bên cạnh công thức
2. Chọn loại Caption (Hình/Bảng/Công thức)
3. Nhập nội dung mô tả
4. Nhấn **"Chèn Caption"**
5. Sau khi chèn xong toàn bộ, nhấn **"Cập nhật toàn bộ Fields"**

### 4. Tạo Mục lục

1. Đặt con trỏ vào vị trí muốn chèn mục lục (thường là đầu văn bản)
2. Nhấn **"Tạo/Tái lập Mục lục"**
3. Nhấn **"Cập nhật toàn bộ Fields"** để cập nhật số trang

### 5. Đánh số trang

1. Chọn vị trí (canh giữa/trái/phải)
2. Chọn số trang bắt đầu
3. Tick "Không đánh số trang đầu" nếu là trang bìa
4. Nhấn **"Chèn số trang"**

---

## ⚠️ Lưu ý quan trọng

1. **Cập nhật Fields**: Sau mỗi lần chỉnh sửa heading, caption, hoặc số trang, hãy nhấn **"Cập nhật toàn bộ Fields"** để số thứ tự và số trang được cập nhật.

2. **Lưu văn bản**: Add-in không tự động lưu. Nhớ nhấn `Ctrl + S` thường xuyên.

3. **Macro/Script**: Add-in này sử dụng Office.js API, hoàn toàn an toàn và không cần bật macro.

4. **Tương thích**: Được thiết kế cho Word trên Windows. Word for Mac và Word Online có thể hỗ trợ hạn chế một số tính năng.

---

## 🔧 Troubleshooting

| Vấn đề | Giải pháp |
|--------|-----------|
| Add-in không hiển thị | Kiểm tra lại đường dẫn trong manifest.xml. Đảm bảo file HTML có thể truy cập được. |
| Lỗi "Add-in không tải được" | Kiểm tra kết nối mạng (nếu host trên server). Đảm bảo URL đúng. |
| Fields không cập nhật | Nhấn chuột phải vào field → Update Field. Hoặc dùng nút "Cập nhật toàn bộ Fields" trong add-in. |
| Styles không áp dụng | Đảm bảo đã chọn đoạn văn bản trước khi nhấn nút. |

---

## 📜 Giấy phép

MIT License - Tự do sử dụng, chỉnh sửa và phân phối.

---

## 🤝 Đóng góp

Mọi đóng góp, báo lỗi hoặc góp ý cải tiến đều được hoan nghênh!

---

**Được xây dựng với ❤️ cho cộng đồng văn phòng Việt Nam.**
