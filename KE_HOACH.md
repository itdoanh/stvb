# KẾ HOẠCH CHI TIẾT DỰ ÁN
# ADD-IN WORD "SOẠN THẢO VĂN BẢN VIỆT NAM AI"

---

## 1. TỔNG QUAN DỰ ÁN

### 1.1. Tên dự án
**Soạn thảo văn bản Việt Nam AI** — Word Add-in cho Microsoft Office 365

### 1.2. Mục tiêu
Xây dựng add-in Word tích hợp AI Agentic giúp người dùng soạn thảo văn bản hành chính, báo cáo, luận văn theo đúng chuẩn **Nghị định 30/2020/NĐ-CP** về công tác văn thư, đồng thời tận dụng AI để tự động hóa các tác vụ phức tạp.

### 1.3. Đối tượng sử dụng
- Cán bộ, công chức, viên chức nhà nước
- Sinh viên, nghiên cứu sinh viết luận văn, báo cáo
- Doanh nghiệp cần soạn thảo văn bản chuẩn hóa

### 1.4. Phạm vi
- Hỗ trợ Microsoft Word 2016+ và Office 365 (Desktop)
- Task Pane add-in (HTML/CSS/JS + Office.js API)
- Tích hợp AI qua API OpenRouter hoặc Local LM (Ollama, LM Studio)

---

## 2. YÊU CẦU CHỨC NĂNG

### 2.1. Tab AI Agent (Trung tâm)
| STT | Chức năng | Mô tả | Trạng thái |
|-----|-----------|-------|------------|
| 1 | Chat interface | Nhập yêu cầu tự nhiên, AI trả lời | Đang làm |
| 2 | Document analysis | AI đọc toàn bộ nội dung Word hiện tại | Đang làm |
| 3 | Action planning | AI lập kế hoạch các bước sửa tài liệu | Đang làm |
| 4 | Auto-apply styles | AI tự động áp Heading 1-5 đúng cấp | Đang làm |
| 5 | Auto TOC | AI tạo/tái lập mục lục | Đang làm |
| 6 | Auto captions | AI đánh số và chèn caption hình/bảng | Đang làm |
| 7 | Image insertion | AI chèn ảnh từ URL/base64 vào Word | Đang làm |
| 8 | Text rewriting | AI viết lại đoạn văn giữ nguyên ý | Đang làm |
| 9 | Model selection | Chọn model AI cho từng lần chạy | Đang làm |
| 10 | Agent logging | Ghi nhật ký các bước AI thực hiện | Đang làm |

### 2.2. Tab Format (Định dạng)
| STT | Chức năng | Mô tả | Trạng thái |
|-----|-----------|-------|------------|
| 1 | Page setup | Khổ A4/A3, lề chuẩn VN | Hoàn thành |
| 2 | Font settings | Times New Roman, cỡ chữ, dãn dòng | Hoàn thành |
| 3 | Heading 1 | CHỮ IN HOA, ĐẬM, canh giữa | Hoàn thành |
| 4 | Heading 2 | In hoa thường, đậm, đứng | Hoàn thành |
| 5 | Heading 3 | In hoa thường, đậm, nghiêng | Hoàn thành |
| 6 | Heading 4 | Thường, đứng | Hoàn thành |
| 7 | Heading 5 | Thường, nghiêng | Hoàn thành |
| 8 | Batch update | Cập nhật style toàn văn bản | Hoàn thành |

### 2.3. Tab Chèn (Insert)
| STT | Chức năng | Mô tả | Trạng thái |
|-----|-----------|-------|------------|
| 1 | Caption hình | Dưới hình, đánh số theo chương | Hoàn thành |
| 2 | Caption bảng | Trên bảng, đánh số theo chương | Hoàn thành |
| 3 | Caption công thức | Căn phải, số trong ngoặc đơn | Hoàn thành |
| 4 | Page numbers | Canh giữa/trái/phải, tùy chọn bắt đầu | Hoàn thành |
| 5 | Official header | Đầu trang văn bản hành chính VN | Hoàn thành |
| 6 | Signature block | Khối ký tên chuẩn | Hoàn thành |

### 2.4. Tab Mục lục (TOC)
| STT | Chức năng | Mô tả | Trạng thái |
|-----|-----------|-------|------------|
| 1 | Table of Contents | Tạo từ Heading 1-5 | Hoàn thành |
| 2 | List of Figures | Danh sách hình ảnh | Hoàn thành |
| 3 | List of Tables | Danh sách bảng biểu | Hoàn thành |
| 4 | Update all fields | Cập nhật toàn bộ fields | Hoàn thành |

### 2.5. Tab Tools
| STT | Chức năng | Mô tả | Trạng thái |
|-----|-----------|-------|------------|
| 1 | Fix orphan headings | Sửa heading mồ côi cuối trang | Hoàn thành |
| 2 | Remove double spaces | Xóa khoảng trắng thừa | Hoàn thành |
| 3 | Fix paragraph spacing | Chuẩn hóa Before/After | Hoàn thành |
| 4 | Vietnamese quotes | Chuyển "..." thành "..." | Hoàn thành |

### 2.6. Tab Settings
| STT | Chức năng | Mô tả | Trạng thái |
|-----|-----------|-------|------------|
| 1 | AI Provider | OpenRouter / Local LM | Đang làm |
| 2 | Base URL | Tùy chỉnh endpoint API | Đang làm |
| 3 | API Key | Lưu key an toàn (localStorage) | Đang làm |
| 4 | Default model | Model mặc định | Đang làm |
| 5 | Load models | Tự động load danh sách model | Đang làm |
| 6 | System prompt | Tùy chỉnh prompt hệ thống | Đang làm |

---

## 3. KIẾN TRÚC KỸ THUẬT

### 3.1. Stack công nghệ
- **Frontend**: HTML5, CSS3, Vanilla JavaScript
- **Office Integration**: Office.js API (Word 1.1+)
- **AI Integration**: Fetch API → OpenRouter / Local OpenAI-compatible endpoint
- **Storage**: localStorage (settings, API key)
- **Deployment**: Static web server + manifest.xml sideload

### 3.2. Luồng dữ liệu AI Agent
```
Người dùng nhập prompt
    ↓
Đọc toàn bộ nội dung Word (Office.js)
    ↓
Gửi đến AI API kèm system prompt + document context
    ↓
AI trả về JSON action plan
    ↓
Thực thi từng action qua Office.js
    ↓
Ghi log và hiển thị kết quả
```

### 3.3. Cấu trúc file
```
c:/code/stvb/
├── manifest.xml              # Office Add-in manifest
├── taskpane.html             # UI chính (6 tabs)
├── taskpane.css              # Stylesystem
├── taskpane.js               # Logic + Office.js + AI Agent
├── assets/
│   └── icon.svg              # Icon add-in
├── KE_HOACH.md               # Kế hoạch này
├── TIEN_TRINH.md             # Tiến trình thực hiện
└── README.md                 # Hướng dẫn người dùng
```

---

## 4. LỊCH TRÌNH THỰC HIỆN

### Giai đoạn 1: Nền tảng (Đã hoàn thành)
- [x] Tạo manifest.xml
- [x] Thiết kế UI taskpane dọc với tab ngang
- [x] Viết CSS modern design system
- [x] Tích hợp Office.js cơ bản

### Giai đoạn 2: Chức năng Word cơ bản (Đã hoàn thành)
- [x] Thiết lập trang & lề
- [x] Styles Heading 1-5
- [x] Caption hình/bảng/công thức
- [x] Đánh số trang
- [x] Mục lục & danh sách
- [x] Công cụ chuẩn hóa

### Giai đoạn 3: AI Agent (Đang thực hiện)
- [ ] Cấu hình AI Provider (OpenRouter + Local)
- [ ] Load danh sách model động
- [ ] Document context extraction
- [ ] Action plan JSON schema
- [ ] Thực thi actions qua Office.js
- [ ] Image insertion từ URL/base64
- [ ] Agent logging UI

### Giai đoạn 4: Hoàn thiện (Chưa bắt đầu)
- [ ] Test toàn bộ chức năng
- [ ] Xử lý lỗi & edge cases
- [ ] Tối ưu performance
- [ ] Viết tài liệu hướng dẫn
- [ ] Đóng gói & deploy

---

## 5. TIÊU CHÍ THÀNH CÔNG

1. **Chức năng**: Tất cả 6 tab hoạt động ổn định trên Word Desktop
2. **AI**: Agent có thể đọc, phân tích và sửa tài liệu với độ chính xác > 80%
3. **UI/UX**: Giao diện rõ ràng, responsive, dễ sử dụng trên taskpane hẹp
4. **Hiệu năng**: Thao tác Word không quá 3 giây, AI response không quá 30 giây
5. **Bảo mật**: API key chỉ lưu local, không gửi đến server bên thứ 3 ngoài API AI

---

## 6. RỦI RO & GIẢI PHÁP

| Rủi ro | Ảnh hưởng | Giải pháp |
|--------|-----------|-----------|
| Office.js API hạn chế | Không thể tạo style mới | Dùng paragraph formatting thay thế |
| AI API chậm/timeout | Trải nghiệm kém | Thêm loading indicator, retry logic |
| Local LM không khả dụng | Không chạy offline | Fallback sang OpenRouter hoặc thông báo rõ |
| Taskpane quá hẹp | UI bị vỡ | Thiết kế mobile-first, scroll ngang tab |

---

*Lập kế hoạch ngày: 22/05/2026*
*Người lập: AI Assistant*
