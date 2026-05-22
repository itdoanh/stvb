/**
 * Soạn thảo văn bản Việt Nam - Word Add-in
 * Theo Nghị định 30/2020/NĐ-CP và chuẩn trình bày văn bản Việt Nam
 */

// =====================
// GLOBAL CONSTANTS (Must be at top)
// =====================
const AI_SETTINGS_KEY = 'vn_word_ai_settings';

// =====================
// Office Initialization
// =====================
Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) {
        loadAiSettingsToUI();
        showStatus("Add-in đã sẵn sàng!", "success");
    }
});

// =====================
// UI: Tab Switching
// =====================
function switchTab(tabId) {
    // Update nav items
    document.querySelectorAll('.nav-item').forEach(item => {
        item.classList.remove('active');
        if (item.dataset.tab === tabId) {
            item.classList.add('active');
        }
    });

    // Update tab panels
    document.querySelectorAll('.tab-panel').forEach(panel => {
        panel.classList.remove('active');
    });
    const targetPanel = document.getElementById('tab-' + tabId);
    if (targetPanel) {
        targetPanel.classList.add('active');
    }
}

// =====================
// UI: Status Bar
// =====================
function showStatus(message, type) {
    const statusBar = document.getElementById('status-bar');
    const statusMessage = document.getElementById('status-message');
    const statusIcon = document.getElementById('status-icon');

    statusMessage.textContent = message;
    statusBar.classList.remove('hidden', 'success', 'error', 'warning');
    statusBar.classList.add(type);

    if (type === 'success') statusIcon.textContent = '✓';
    else if (type === 'error') statusIcon.textContent = '✕';
    else if (type === 'warning') statusIcon.textContent = '!';
    else statusIcon.textContent = 'ℹ';

    setTimeout(() => {
        statusBar.classList.add('hidden');
    }, 4000);
}

function cmToPoints(cm) {
    return cm * 28.3465;
}

// =====================
// QUICK SETUP (Home Tab)
// =====================
function quickSetup() {
    Word.run(function (context) {
        const sections = context.document.sections;
        sections.load("items");

        return context.sync().then(function () {
            sections.items.forEach(function (section) {
                section.pageWidth = cmToPoints(21.0);
                section.pageHeight = cmToPoints(29.7);
                section.leftMargin = cmToPoints(3.0);
                section.rightMargin = cmToPoints(2.0);
                section.topMargin = cmToPoints(2.0);
                section.bottomMargin = cmToPoints(2.0);
            });

            const body = context.document.body;
            body.font.name = "Times New Roman";
            body.font.size = 13;

            const paragraphs = body.paragraphs;
            paragraphs.load("items");

            return context.sync().then(function () {
                paragraphs.items.forEach(function (para) {
                    para.lineSpacing = 1.5;
                    para.spaceAfter = 6;
                    para.spaceBefore = 0;
                });
                return context.sync();
            });
        });
    })
    .then(function () {
        showStatus("Đã áp dụng thiết lập nhanh: A4, lề chuẩn VN, Times New Roman 13pt, 1.5 dòng!", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

// =====================
// 1. Thiết lập trang & Lề (Format Tab)
// =====================
function applyPageSetup() {
    const paperSize = document.getElementById("paperSize").value;
    const marginLeft = parseFloat(document.getElementById("marginLeft").value);
    const marginRight = parseFloat(document.getElementById("marginRight").value);
    const marginTop = parseFloat(document.getElementById("marginTop").value);
    const marginBottom = parseFloat(document.getElementById("marginBottom").value);
    const fontName = document.getElementById("defaultFont").value;
    const fontSize = parseFloat(document.getElementById("defaultFontSize").value);
    const lineSpacing = parseFloat(document.getElementById("lineSpacing").value);

    Word.run(function (context) {
        const sections = context.document.sections;
        sections.load("items");

        return context.sync().then(function () {
            sections.items.forEach(function (section) {
                if (paperSize === "A4") {
                    section.pageWidth = cmToPoints(21.0);
                    section.pageHeight = cmToPoints(29.7);
                } else if (paperSize === "A3") {
                    section.pageWidth = cmToPoints(29.7);
                    section.pageHeight = cmToPoints(42.0);
                } else {
                    section.pageWidth = cmToPoints(21.59);
                    section.pageHeight = cmToPoints(27.94);
                }

                section.leftMargin = cmToPoints(marginLeft);
                section.rightMargin = cmToPoints(marginRight);
                section.topMargin = cmToPoints(marginTop);
                section.bottomMargin = cmToPoints(marginBottom);
            });

            const body = context.document.body;
            body.font.name = fontName;
            body.font.size = fontSize;

            const paragraphs = body.paragraphs;
            paragraphs.load("items");

            return context.sync().then(function () {
                paragraphs.items.forEach(function (para) {
                    para.lineSpacing = lineSpacing;
                    para.spaceAfter = 6;
                    para.spaceBefore = 0;
                });
                return context.sync();
            });
        });
    })
    .then(function () {
        showStatus("Đã áp dụng thiết lập trang thành công!", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

// =====================
// 2. Styles đề mục chuẩn Việt Nam
// =====================
const HEADING_CONFIG = {
    1: {
        styleName: "Heading 1",
        fontSize: 16,
        bold: true,
        italic: false,
        alignment: Word.Alignment.centered,
        spaceBefore: 24,
        spaceAfter: 12,
        keepWithNext: true,
        keepLinesTogether: true,
        outlineLevel: 0
    },
    2: {
        styleName: "Heading 2",
        fontSize: 14,
        bold: true,
        italic: false,
        alignment: Word.Alignment.left,
        spaceBefore: 18,
        spaceAfter: 6,
        keepWithNext: true,
        keepLinesTogether: true,
        outlineLevel: 1
    },
    3: {
        styleName: "Heading 3",
        fontSize: 13,
        bold: true,
        italic: true,
        alignment: Word.Alignment.left,
        spaceBefore: 12,
        spaceAfter: 6,
        keepWithNext: true,
        keepLinesTogether: true,
        outlineLevel: 2
    },
    4: {
        styleName: "Heading 4",
        fontSize: 13,
        bold: false,
        italic: false,
        alignment: Word.Alignment.left,
        spaceBefore: 6,
        spaceAfter: 6,
        keepWithNext: true,
        keepLinesTogether: true,
        outlineLevel: 3
    },
    5: {
        styleName: "Heading 5",
        fontSize: 13,
        bold: false,
        italic: true,
        alignment: Word.Alignment.left,
        spaceBefore: 6,
        spaceAfter: 6,
        keepWithNext: true,
        keepLinesTogether: true,
        outlineLevel: 4
    }
};

function createVietnameseStyles() {
    const fontName = document.getElementById("defaultFont")?.value || "Times New Roman";

    Word.run(function (context) {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load("items, style, text");

        return context.sync().then(function () {
            paragraphs.items.forEach(function (para) {
                const styleName = para.style;
                for (let level = 1; level <= 5; level++) {
                    const cfg = HEADING_CONFIG[level];
                    if (styleName === cfg.styleName || styleName === cfg.styleName.replace("Heading ", "heading ")) {
                        applyHeadingFormatting(para, cfg, fontName);
                    }
                }
            });
            return context.sync();
        });
    })
    .then(function () {
        showStatus("Đã cập nhật định dạng Heading 1-5 theo chuẩn VN!", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

function applyHeadingFormatting(paragraph, config, fontName) {
    paragraph.font.name = fontName;
    paragraph.font.size = config.fontSize;
    paragraph.font.bold = config.bold;
    paragraph.font.italic = config.italic;
    paragraph.alignment = config.alignment;
    paragraph.spaceBefore = config.spaceBefore;
    paragraph.spaceAfter = config.spaceAfter;
    paragraph.keepWithNext = config.keepWithNext;
    paragraph.keepLinesTogether = config.keepLinesTogether;
}

function applyHeadingStyle(level) {
    const fontName = document.getElementById("defaultFont")?.value || "Times New Roman";
    const cfg = HEADING_CONFIG[level];

    Word.run(function (context) {
        const selection = context.document.getSelection();
        const paragraphs = selection.paragraphs;
        paragraphs.load("items");

        return context.sync().then(function () {
            paragraphs.items.forEach(function (para) {
                para.style = cfg.styleName;
                applyHeadingFormatting(para, cfg, fontName);
            });
            return context.sync();
        });
    })
    .then(function () {
        showStatus("Đã áp dụng Heading " + level + "!", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

// =====================
// 3. Đánh số trang
// =====================
function insertPageNumbers() {
    const position = document.getElementById("pageNumberPosition").value;
    const startNumber = parseInt(document.getElementById("startPageNumber").value) || 1;
    const skipFirst = document.getElementById("skipFirstPage").checked;
    const startFromPart1 = document.getElementById("startFromPart1").checked;

    Word.run(function (context) {
        const sections = context.document.sections;
        sections.load("items");

        return context.sync().then(function () {
            sections.items.forEach(function (section, index) {
                const footer = section.getFooter("primary");
                footer.load("paragraphs");

                return context.sync().then(function () {
                    footer.clear();

                    const para = footer.insertParagraph("", Word.InsertLocation.end);

                    if (position === "center") {
                        para.alignment = Word.Alignment.centered;
                    } else if (position === "right") {
                        para.alignment = Word.Alignment.right;
                    } else {
                        para.alignment = Word.Alignment.left;
                    }

                    const range = para.getRange('End');
                    range.insertField(Word.FieldType.page, true);

                    if (skipFirst && index === 0) {
                        section.differentFirstPage = true;
                        const firstFooter = section.getFooter("firstPage");
                        firstFooter.clear();
                    }

                    if (startNumber > 1 || startFromPart1) {
                        section.pageStart = startNumber;
                    }

                    return context.sync();
                });
            });
        });
    })
    .then(function () {
        showStatus("Đã chèn số trang thành công!", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

function removePageNumbers() {
    Word.run(function (context) {
        const sections = context.document.sections;
        sections.load("items");

        return context.sync().then(function () {
            sections.items.forEach(function (section) {
                const footer = section.getFooter("primary");
                footer.clear();
                section.differentFirstPage = false;
            });
            return context.sync();
        });
    })
    .then(function () {
        showStatus("Đã xóa số trang!", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

// =====================
// 4. Caption Hình, Bảng, Công thức
// =====================
function insertCaption() {
    const type = document.getElementById("captionType").value;
    const text = document.getElementById("captionText").value || "";
    const autoNumber = document.getElementById("autoNumber").checked;

    Word.run(function (context) {
        const selection = context.document.getSelection();
        const paragraphs = selection.paragraphs;
        paragraphs.load("items");

        return context.sync().then(function () {
            if (paragraphs.items.length === 0) {
                showStatus("Vui lòng đặt con trỏ vào vị trí cần chèn caption!", "error");
                return;
            }

            const currentPara = paragraphs.items[0];
            let captionText = "";

            if (autoNumber) {
                const seqFieldName = type === "Hình" ? "Figure" : (type === "Bảng" ? "Table" : "Equation");
                captionText = type + " ";
                
                const captionPara = currentPara.insertParagraph(captionText, Word.InsertLocation.after);
                captionPara.font.italic = true;
                captionPara.font.size = 12;
                captionPara.alignment = Word.Alignment.centered;
                
                const range = captionPara.getRange('End');
                const seqField = range.insertField(Word.FieldType.sequence, true);
                captionPara.insertText(": " + text, Word.InsertLocation.end);
            } else {
                captionText = type + ": " + text;
                const captionPara = currentPara.insertParagraph(captionText, Word.InsertLocation.after);
                captionPara.font.italic = true;
                captionPara.font.size = 12;
                captionPara.alignment = Word.Alignment.centered;
            }

            return context.sync();
        });
    })
    .then(function () {
        showStatus("Đã chèn caption! Nhớ nhấn 'Cập nhật toàn bộ Fields' để đánh số tự động.", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

function createCaptionStyles() {
    Word.run(function (context) {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load("items, text, font");

        return context.sync().then(function () {
            paragraphs.items.forEach(function (para) {
                const text = para.text.trim();
                if (text.startsWith("Hình ") || text.startsWith("Bảng ") || text.startsWith("Công thức ")) {
                    para.font.italic = true;
                    para.font.size = 12;
                    para.alignment = Word.Alignment.centered;
                    para.spaceBefore = 6;
                    para.spaceAfter = 6;
                }
            });
            return context.sync();
        });
    })
    .then(function () {
        showStatus("Đã định dạng caption chuẩn!", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

// =====================
// 5. Mục lục & Danh sách
// =====================
function insertTableOfContents() {
    Word.run(function (context) {
        const selection = context.document.getSelection();
        const tocRange = selection.insertTableOfContents("\\o \"1-5\" \\h \\z \\u", true);
        return context.sync();
    })
    .then(function () {
        showStatus("Đã chèn mục lục! Nhấn 'Cập nhật toàn bộ Fields' nếu cần.", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

function insertListOfFigures() {
    Word.run(function (context) {
        const selection = context.document.getSelection();
        const range = selection.insertParagraph("DANH SÁCH HÌNH ẢNH", Word.InsertLocation.before);
        range.font.bold = true;
        range.font.size = 14;
        range.alignment = Word.Alignment.centered;
        
        const tocPara = range.insertParagraph("", Word.InsertLocation.after);
        const tocRange = tocPara.getRange('End');
        tocRange.insertField(Word.FieldType.toc, true);
        
        return context.sync();
    })
    .then(function () {
        showStatus("Đã chèn khung Danh sách hình ảnh! Cập nhật fields để hiển thị.", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

function insertListOfTables() {
    Word.run(function (context) {
        const selection = context.document.getSelection();
        const range = selection.insertParagraph("DANH SÁCH BẢNG BIỂU", Word.InsertLocation.before);
        range.font.bold = true;
        range.font.size = 14;
        range.alignment = Word.Alignment.centered;
        
        const tocPara = range.insertParagraph("", Word.InsertLocation.after);
        const tocRange = tocPara.getRange('End');
        tocRange.insertField(Word.FieldType.toc, true);
        
        return context.sync();
    })
    .then(function () {
        showStatus("Đã chèn khung Danh sách bảng biểu! Cập nhật fields để hiển thị.", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

function updateAllFields() {
    Word.run(function (context) {
        const body = context.document.body;
        const fields = body.fields;
        fields.load("items");

        return context.sync().then(function () {
            fields.items.forEach(function (field) {
                try {
                    field.update();
                } catch (e) {
                    console.warn("Không thể cập nhật field:", e);
                }
            });
            return context.sync();
        });
    })
    .then(function () {
        showStatus("Đã cập nhật toàn bộ fields!", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

// =====================
// 6. Công cụ nâng cao
// =====================
function fixOrphanHeadings() {
    Word.run(function (context) {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load("items, style, text");

        return context.sync().then(function () {
            let fixedCount = 0;
            const items = paragraphs.items;

            for (let i = 0; i < items.length; i++) {
                const para = items[i];
                const style = para.style;
                
                if (style && (style.includes("Heading") || style.includes("heading"))) {
                    if (i + 1 < items.length) {
                        const nextPara = items[i + 1];
                        const nextStyle = nextPara.style;
                        if (nextStyle && (nextStyle.includes("Heading") || nextStyle.includes("heading"))) {
                            para.insertBreak(Word.BreakType.page, Word.InsertLocation.before);
                            fixedCount++;
                        }
                    }
                    
                    para.keepWithNext = true;
                    para.keepLinesTogether = true;
                }
            }

            return context.sync().then(function () {
                return fixedCount;
            });
        });
    })
    .then(function (count) {
        showStatus("Đã sửa " + count + " heading bị mồ côi!", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

function removeDoubleSpaces() {
    Word.run(function (context) {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load("items, text");

        return context.sync().then(function () {
            paragraphs.items.forEach(function (para) {
                const originalText = para.text;
                const cleanedText = originalText.replace(/\s{2,}/g, " ");
                if (originalText !== cleanedText) {
                    para.insertText(cleanedText, Word.InsertLocation.replace);
                }
            });
            return context.sync();
        });
    })
    .then(function () {
        showStatus("Đã xóa khoảng trắng thừa!", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

function fixParagraphSpacing() {
    Word.run(function (context) {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load("items, style, text");

        return context.sync().then(function () {
            paragraphs.items.forEach(function (para) {
                const style = para.style;
                
                if (style && (style.includes("Heading") || style.includes("heading"))) {
                    para.spaceBefore = 12;
                    para.spaceAfter = 6;
                } else if (para.text.trim() === "") {
                    para.spaceBefore = 0;
                    para.spaceAfter = 0;
                } else {
                    para.spaceBefore = 0;
                    para.spaceAfter = 6;
                }
            });
            return context.sync();
        });
    })
    .then(function () {
        showStatus("Đã chuẩn hóa khoảng cách đoạn!", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

function applyVietnameseQuotes() {
    Word.run(function (context) {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load("items, text");

        return context.sync().then(function () {
            paragraphs.items.forEach(function (para) {
                let text = para.text;
                text = text.replace(/"([^"]*)"/g, "\u201C$1\u201D");
                para.insertText(text, Word.InsertLocation.replace);
            });
            return context.sync();
        });
    })
    .then(function () {
        showStatus("Đã chuyển dấu ngoặc kép chuẩn VN!", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

function insertSignatureBlock() {
    Word.run(function (context) {
        const selection = context.document.getSelection();
        
        const para1 = selection.insertParagraph("", Word.InsertLocation.after);
        para1.alignment = Word.Alignment.right;
        para1.font.bold = true;
        para1.font.size = 13;
        para1.insertText("GIÁM ĐỐC", Word.InsertLocation.end);
        
        const para2 = para1.insertParagraph("(Ký, ghi rõ họ tên, đóng dấu)", Word.InsertLocation.after);
        para2.alignment = Word.Alignment.right;
        para2.font.italic = true;
        para2.font.size = 12;
        
        const para3 = para2.insertParagraph("", Word.InsertLocation.after);
        para3.alignment = Word.Alignment.right;
        para3.insertText("____________________", Word.InsertLocation.end);
        
        return context.sync();
    })
    .then(function () {
        showStatus("Đã chèn khối ký tên!", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

function insertOfficialHeader() {
    Word.run(function (context) {
        const selection = context.document.getSelection();
        
        const line1 = selection.insertParagraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", Word.InsertLocation.after);
        line1.alignment = Word.Alignment.centered;
        line1.font.bold = true;
        line1.font.size = 13;
        
        const line2 = line1.insertParagraph("Độc lập - Tự do - Hạnh phúc", Word.InsertLocation.after);
        line2.alignment = Word.Alignment.centered;
        line2.font.bold = true;
        line2.font.size = 13;
        
        const line3 = line2.insertParagraph("---------------", Word.InsertLocation.after);
        line3.alignment = Word.Alignment.centered;
        
        const line4 = line3.insertParagraph("TÊN CƠ QUAN/TỔ CHỨC", Word.InsertLocation.after);
        line4.alignment = Word.Alignment.centered;
        line4.font.bold = true;
        line4.font.size = 14;
        
        const line5 = line4.insertParagraph("Số: ……/……/……", Word.InsertLocation.after);
        line5.alignment = Word.Alignment.left;
        line5.font.size = 13;
        
        return context.sync();
    })
    .then(function () {
        showStatus("Đã chèn đầu trang văn bản hành chính!", "success");
    })
    .catch(function (error) {
        showStatus("Lỗi: " + error.message, "error");
        console.error(error);
    });
}

// =====================
// 7. AI AGENTIC INTEGRATION
// =====================

function getAiSettings() {
    try {
        const raw = localStorage.getItem(AI_SETTINGS_KEY);
        if (raw) return JSON.parse(raw);
    } catch (e) { console.warn('Lỗi đọc AI settings:', e); }
    return {
        provider: 'openrouter',
        baseUrl: 'https://openrouter.ai/api/v1',
        apiKey: '',
        defaultModel: 'openai/gpt-4o-mini',
        systemPrompt: 'Bạn là AI Agent soạn thảo Word chuẩn Việt Nam. Luôn ưu tiên Nghị định 30/2020/NĐ-CP, văn phong hành chính rõ ràng, heading đúng cấp, caption hình/bảng/công thức thống nhất, mục lục cập nhật. Khi cần sửa tài liệu, trả về JSON actions hợp lệ.'
    };
}

function saveAiSettings() {
    const settings = {
        provider: document.getElementById('aiProvider').value,
        baseUrl: document.getElementById('aiBaseUrl').value.trim(),
        apiKey: document.getElementById('aiApiKey').value.trim(),
        defaultModel: document.getElementById('aiDefaultModel').value.trim(),
        systemPrompt: document.getElementById('aiSystemPrompt').value
    };
    localStorage.setItem(AI_SETTINGS_KEY, JSON.stringify(settings));
    showStatus('Đã lưu cài đặt AI!', 'success');
}

function updateProviderFields() {
    const provider = document.getElementById('aiProvider').value;
    const baseUrlInput = document.getElementById('aiBaseUrl');
    if (provider === 'openrouter') {
        baseUrlInput.placeholder = 'https://openrouter.ai/api/v1';
        if (!baseUrlInput.value) baseUrlInput.value = 'https://openrouter.ai/api/v1';
    } else {
        baseUrlInput.placeholder = 'http://localhost:11434/v1 hoặc http://localhost:1234/v1';
    }
}

function loadAiSettingsToUI() {
    const s = getAiSettings();
    document.getElementById('aiProvider').value = s.provider;
    document.getElementById('aiBaseUrl').value = s.baseUrl;
    document.getElementById('aiApiKey').value = s.apiKey;
    document.getElementById('aiDefaultModel').value = s.defaultModel;
    document.getElementById('aiSystemPrompt').value = s.systemPrompt;
    updateProviderFields();
}

async function loadAiModels() {
    const s = getAiSettings();
    const select = document.getElementById('agentModel');
    select.innerHTML = '<option>Đang tải...</option>';
    logAgent('Đang tải danh sách model từ ' + s.baseUrl + '...');

    try {
        const res = await fetch(s.baseUrl + '/models', {
            method: 'GET',
            headers: {
                'Authorization': 'Bearer ' + (s.apiKey || ''),
                'Content-Type': 'application/json'
            }
        });
        if (!res.ok) throw new Error('HTTP ' + res.status);
        const data = await res.json();
        const models = data.data || data.models || [];
        select.innerHTML = '';
        if (models.length === 0) {
            select.innerHTML = '<option>Không có model</option>';
            logAgent('Không tìm thấy model nào. Kiểm tra API key và base URL.');
            return;
        }
        models.forEach(function(m) {
            const id = m.id || m.name || m.model;
            const opt = document.createElement('option');
            opt.value = id;
            opt.textContent = id;
            if (id === s.defaultModel) opt.selected = true;
            select.appendChild(opt);
        });
        logAgent('Đã load ' + models.length + ' model.');
        showStatus('Đã load ' + models.length + ' model!', 'success');
    } catch (err) {
        select.innerHTML = '<option>Lỗi load model</option>';
        logAgent('Lỗi load model: ' + err.message);
        showStatus('Lỗi load model: ' + err.message, 'error');
    }
}

function filterAiModels(searchText) {
    const select = document.getElementById('agentModel');
    const options = select.querySelectorAll('option');
    let visibleCount = 0;
    
    options.forEach(function(opt) {
        if (opt.value === '' || opt.textContent.toLowerCase().includes(searchText.toLowerCase())) {
            opt.style.display = '';
            visibleCount++;
        } else {
            opt.style.display = 'none';
        }
    });
    
    if (visibleCount === 0) {
        logAgent('Không tìm thấy model phù hợp với: ' + searchText);
    }
}

function setAgentPrompt(text) {
    document.getElementById('agentPrompt').value = text;
}

function logAgent(text) {
    const el = document.getElementById('agentLog');
    const time = new Date().toLocaleTimeString('vi-VN');
    el.textContent += '\n[' + time + '] ' + text;
    el.scrollTop = el.scrollHeight;
}

function updateAgentProgress(percent) {
    const progressBar = document.getElementById('agentProgress');
    const progressFill = document.getElementById('agentProgressFill');
    const progressText = document.getElementById('agentProgressText');
    if (percent > 0) {
        progressBar.classList.remove('hidden');
        progressFill.style.width = percent + '%';
        progressText.textContent = percent + '%';
    } else {
        progressBar.classList.add('hidden');
    }
}

async function runAiAgent() {
    const prompt = document.getElementById('agentPrompt').value.trim();
    const model = document.getElementById('agentModel').value;
    const canEdit = document.getElementById('agentCanEdit').checked && !document.getElementById('agentAnalysisOnly').checked;
    const canInsertImages = document.getElementById('agentCanInsertImages').checked;
    const useSmartContext = document.getElementById('agentSmartContext').checked;
    const s = getAiSettings();

    if (!prompt) {
        showStatus('Vui lòng nhập yêu cầu cho AI!', 'error');
        return;
    }
    if (!s.apiKey && s.provider === 'openrouter') {
        showStatus('Vui lòng nhập API Key ở tab Setting!', 'error');
        switchTab('settings');
        return;
    }

    logAgent('=== BẮT ĐẦU AGENT ===');
    logAgent('Model: ' + (model || s.defaultModel));
    logAgent('Yêu cầu: ' + prompt);

    // Step 1: Read document content
    updateAgentProgress(10);
    let docText = '';
    try {
        docText = await getDocumentText();
        logAgent('Đã đọc tài liệu (' + docText.length + ' ký tự, ~' + estimateTokens(docText) + ' tokens).');
    } catch (e) {
        logAgent('Lỗi đọc tài liệu: ' + e.message);
    }

    // Step 1.5: Smart context extraction
    updateAgentProgress(20);
    let contextText = docText;
    if (useSmartContext) {
        try {
            contextText = await extractSmartContext(docText, prompt);
            logAgent('Context trích xuất: ' + contextText.length + ' ký tự.');
        } catch (e) {
            logAgent('Lỗi trích xuất context: ' + e.message);
        }
    }

    // Step 2: Build messages
    const system = s.systemPrompt +
        '\n\nBạn có thể thực hiện các hành động sau bằng cách trả về JSON array "actions":' +
        '\n- {"type":"rewrite","paragraphIndex":N,"newText":"..."} — viết lại đoạn' +
        '\n- {"type":"applyHeading","paragraphIndex":N,"level":1-5} — áp heading' +
        '\n- {"type":"insertCaption","afterParagraph":N,"kind":"Hình|Bảng|Công thức","text":"..."} — chèn caption' +
        '\n- {"type":"insertImage","afterParagraph":N,"url":"https://...","width":400} — chèn ảnh từ URL' +
        '\n- {"type":"insertPageBreak","beforeParagraph":N} — ngắt trang' +
        '\n- {"type":"insertSectionBreak","beforeParagraph":N} — ngắt section' +
        '\n- {"type":"insertText","afterParagraph":N,"text":"..."} — chèn đoạn văn' +
        '\n- {"type":"findReplace","find":"...","replace":"..."} — tìm và thay thế' +
        '\n- {"type":"setFont","paragraphIndex":N,"name":"Times New Roman","size":13,"bold":true,"italic":false,"color":"#000000"} — đặt font' +
        '\n- {"type":"setAlignment","paragraphIndex":N,"alignment":"left|center|right|justify"} — căn chỉnh' +
        '\n- {"type":"insertBullet","afterParagraph":N,"text":"..."} — chèn bullet point' +
        '\n- {"type":"insertNumberedList","afterParagraph":N,"number":1,"text":"..."} — chèn danh sách số' +
        '\n- {"type":"insertHyperlink","paragraphIndex":N,"url":"https://...","text":"..."} — chèn hyperlink' +
        '\n- {"type":"insertBookmark","paragraphIndex":N,"name":"bookmark_name"} — chèn bookmark' +
        '\n- {"type":"insertComment","paragraphIndex":N,"text":"..."} — chèn comment' +
        '\n- {"type":"highlight","paragraphIndex":N,"color":"yellow|green|blue|red"} — tô sáng' +
        '\n- {"type":"clearFormatting","paragraphIndex":N} — xóa định dạng' +
        '\n- {"type":"updateFields"} — cập nhật fields' +
        '\n- {"type":"createTOC"} — tạo mục lục' +
        (canInsertImages ? '' : '\nKHÔNG được chèn ảnh.') +
        (canEdit ? '' : '\nCHỈ được phân tích, KHÔNG được sửa tài liệu.');

    const messages = [
        { role: 'system', content: system },
        { role: 'user', content: 'NỘI DUNG TÀI LIỆU WORD HIỆN TẠI:\n---\n' + contextText + '\n---\n\nYÊU CẦU: ' + prompt + '\n\nHãy phân tích và trả về JSON với "analysis" (phân tích) và "actions" (mảng hành động).' }
    ];

    // Step 3: Call AI API
    updateAgentProgress(40);
    logAgent('Đang gọi AI API...');
    let responseText = '';
    try {
        responseText = await callAiApi(messages, model || s.defaultModel, s);
        logAgent('AI đã phản hồi.');
    } catch (err) {
        logAgent('Lỗi gọi AI: ' + err.message);
        showStatus('Lỗi gọi AI: ' + err.message, 'error');
        updateAgentProgress(0);
        return;
    }

    // Step 4: Parse response
    updateAgentProgress(60);
    let plan = null;
    try {
        const jsonMatch = responseText.match(/```json\s*([\s\S]*?)```/) || responseText.match(/```\s*([\s\S]*?)```/) || [null, responseText];
        plan = JSON.parse(jsonMatch[1] || responseText);
    } catch (e) {
        logAgent('Không parse được JSON. Hiển thị phản hồi thô.');
        logAgent(responseText.substring(0, 800));
        showStatus('AI phản hồi (không có actions JSON). Xem log.', 'warning');
        updateAgentProgress(0);
        return;
    }

    if (plan.analysis) {
        logAgent('PHÂN TÍCH: ' + plan.analysis);
    }

    if (!plan.actions || !plan.actions.length) {
        logAgent('AI không đề xuất hành động nào.');
        showStatus('AI phân tích xong. Không có hành động cần thực hiện.', 'success');
        updateAgentProgress(100);
        setTimeout(() => updateAgentProgress(0), 1500);
        return;
    }

    logAgent('AI đề xuất ' + plan.actions.length + ' hành động.');

    // Step 5: Execute actions
    updateAgentProgress(75);
    if (canEdit) {
        await executeAiActions(plan.actions);
    } else {
        logAgent('Chế độ phân tích — không thực thi actions.');
        showStatus('Chế độ phân tích. AI đã phân tích nhưng chưa sửa.', 'warning');
    }

    updateAgentProgress(100);
    logAgent('=== HOÀN TẤT AGENT ===');
    setTimeout(() => updateAgentProgress(0), 1500);
}

async function getDocumentText() {
    return new Promise(function(resolve, reject) {
        Word.run(function (context) {
            const body = context.document.body;
            body.load('text');
            return context.sync().then(function () {
                resolve(body.text);
            });
        }).catch(reject);
    });
}

async function callAiApi(messages, model, settings) {
    const res = await fetch(settings.baseUrl + '/chat/completions', {
        method: 'POST',
        headers: {
            'Authorization': 'Bearer ' + (settings.apiKey || ''),
            'Content-Type': 'application/json',
            'HTTP-Referer': 'https://localhost:3000',
            'X-Title': 'VietNam Word AI'
        },
        body: JSON.stringify({
            model: model,
            messages: messages,
            temperature: 0.3,
            max_tokens: 4000
        })
    });
    if (!res.ok) {
        const errText = await res.text();
        throw new Error('API ' + res.status + ': ' + errText.substring(0, 200));
    }
    const data = await res.json();
    return data.choices[0].message.content;
}

async function executeAiActions(actions) {
    for (let i = 0; i < actions.length; i++) {
        const act = actions[i];
        logAgent('Thực thi [' + (i + 1) + '/' + actions.length + ']: ' + act.type);
        try {
            await executeSingleAction(act);
        } catch (e) {
            logAgent('  LỖI: ' + e.message);
        }
    }
    showStatus('AI đã thực hiện ' + actions.length + ' hành động!', 'success');
}

async function executeSingleAction(action) {
    return new Promise(function(resolve, reject) {
        Word.run(function (context) {
            const body = context.document.body;
            const paragraphs = body.paragraphs;
            paragraphs.load('items');

            return context.sync().then(function () {
                const items = paragraphs.items;

                switch (action.type) {
                    case 'applyHeading': {
                        const idx = action.paragraphIndex || 0;
                        if (idx < items.length) {
                            const para = items[idx];
                            const level = Math.max(1, Math.min(5, action.level || 1));
                            const cfg = HEADING_CONFIG[level];
                            const fontName = document.getElementById('defaultFont')?.value || 'Times New Roman';
                            para.style = cfg.styleName;
                            applyHeadingFormatting(para, cfg, fontName);
                        }
                        break;
                    }
                    case 'insertText': {
                        const idx = action.afterParagraph || 0;
                        if (idx < items.length) {
                            const newPara = items[idx].insertParagraph(action.text || '', Word.InsertLocation.after);
                            newPara.font.name = document.getElementById('defaultFont')?.value || 'Times New Roman';
                            newPara.font.size = 13;
                        }
                        break;
                    }
                    case 'insertCaption': {
                        const idx = action.afterParagraph || 0;
                        if (idx < items.length) {
                            const captionPara = items[idx].insertParagraph((action.kind || 'Hình') + ': ' + (action.text || ''), Word.InsertLocation.after);
                            captionPara.font.italic = true;
                            captionPara.font.size = 12;
                            captionPara.alignment = Word.Alignment.centered;
                        }
                        break;
                    }
                    case 'insertPageBreak': {
                        const idx = action.beforeParagraph || 0;
                        if (idx < items.length) {
                            items[idx].insertBreak(Word.BreakType.page, Word.InsertLocation.before);
                        }
                        break;
                    }
                    case 'insertImage': {
                        const idx = action.afterParagraph || 0;
                        if (idx < items.length && action.url) {
                            const imgPara = items[idx].insertParagraph('', Word.InsertLocation.after);
                            const img = imgPara.insertInlinePictureFromBase64('', Word.InsertLocation.replace);
                            // Office.js không hỗ trợ insert từ URL trực tiếp; cần fetch base64 trước
                            logAgent('  Chèn ảnh từ URL cần fetch base64 trước (chưa hỗ trợ trực tiếp).');
                        }
                        break;
                    }
                    case 'createTOC': {
                        const selection = context.document.getSelection();
                        selection.insertTableOfContents('\\o \"1-5\" \\h \\z \\u', true);
                        break;
                    }
                    case 'updateFields': {
                        const fields = body.fields;
                        fields.load('items');
                        return context.sync().then(function () {
                            fields.items.forEach(function (f) { try { f.update(); } catch (e) {} });
                            return context.sync();
                        });
                    }
                    default:
                        logAgent('  Action chưa hỗ trợ: ' + action.type);
                }

                return context.sync();
            });
        })
        .then(resolve)
        .catch(reject);
    });
}

// =====================
// 8. SMART CONTEXT & CHUNKING
// =====================

function estimateTokens(text) {
    return Math.ceil(text.length / 4);
}

async function extractSmartContext(fullText, userRequest) {
    const maxTokens = 8000;
    const requestTokens = estimateTokens(userRequest);
    const availableTokens = maxTokens - requestTokens - 500;
    
    if (estimateTokens(fullText) <= availableTokens) {
        logAgent('Văn bản đủ nhỏ (' + estimateTokens(fullText) + ' tokens), sử dụng toàn bộ.');
        return fullText;
    }
    
    logAgent('Văn bản quá dài (' + estimateTokens(fullText) + ' tokens), trích xuất context thông minh...');
    
    const keywords = extractKeywords(userRequest);
    const lines = fullText.split('\n');
    const relevantLines = [];
    let tokenCount = 0;
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const lineTokens = estimateTokens(line);
        if (tokenCount + lineTokens > availableTokens) break;
        
        const isRelevant = keywords.some(kw => line.toLowerCase().includes(kw.toLowerCase())) ||
                          line.match(/^(CHƯƠNG|Chương|\d+\.|Heading|#)/i);
        
        if (isRelevant || i < 5 || i > lines.length - 5) {
            relevantLines.push(line);
            tokenCount += lineTokens;
        }
    }
    
    const context = relevantLines.join('\n');
    logAgent('Trích xuất ' + relevantLines.length + ' dòng (' + estimateTokens(context) + ' tokens).');
    return context;
}

function extractKeywords(text) {
    const stopwords = ['là', 'và', 'của', 'để', 'trong', 'có', 'được', 'từ', 'với', 'này', 'đó', 'cái', 'những'];
    const words = text.toLowerCase().match(/\b\w+\b/g) || [];
    return words.filter(w => w.length > 3 && !stopwords.includes(w)).slice(0, 10);
}

function chunkDocument(text, chunkSize = 3000) {
    const chunks = [];
    let start = 0;
    while (start < text.length) {
        const end = Math.min(start + chunkSize, text.length);
        const lastNewline = text.lastIndexOf('\n', end);
        const chunkEnd = lastNewline > start ? lastNewline : end;
        chunks.push(text.substring(start, chunkEnd));
        start = chunkEnd + 1;
    }
    return chunks;
}

// =====================
// 9. ENHANCED TOOL ACTIONS
// =====================

async function executeSingleAction(action) {
    return new Promise(function(resolve, reject) {
        Word.run(function (context) {
            const body = context.document.body;
            const paragraphs = body.paragraphs;
            paragraphs.load('items');

            return context.sync().then(function () {
                const items = paragraphs.items;

                switch (action.type) {
                    case 'applyHeading': {
                        const idx = action.paragraphIndex || 0;
                        if (idx < items.length) {
                            const para = items[idx];
                            const level = Math.max(1, Math.min(5, action.level || 1));
                            const cfg = HEADING_CONFIG[level];
                            const fontName = document.getElementById('defaultFont')?.value || 'Times New Roman';
                            para.style = cfg.styleName;
                            applyHeadingFormatting(para, cfg, fontName);
                        }
                        break;
                    }
                    case 'insertText': {
                        const idx = action.afterParagraph || 0;
                        if (idx < items.length) {
                            const newPara = items[idx].insertParagraph(action.text || '', Word.InsertLocation.after);
                            newPara.font.name = document.getElementById('defaultFont')?.value || 'Times New Roman';
                            newPara.font.size = 13;
                        }
                        break;
                    }
                    case 'insertCaption': {
                        const idx = action.afterParagraph || 0;
                        if (idx < items.length) {
                            const captionPara = items[idx].insertParagraph((action.kind || 'Hình') + ': ' + (action.text || ''), Word.InsertLocation.after);
                            captionPara.font.italic = true;
                            captionPara.font.size = 12;
                            captionPara.alignment = Word.Alignment.centered;
                        }
                        break;
                    }
                    case 'insertPageBreak': {
                        const idx = action.beforeParagraph || 0;
                        if (idx < items.length) {
                            items[idx].insertBreak(Word.BreakType.page, Word.InsertLocation.before);
                        }
                        break;
                    }
                    case 'insertSectionBreak': {
                        const idx = action.beforeParagraph || 0;
                        if (idx < items.length) {
                            items[idx].insertBreak(Word.BreakType.sectionContinuous, Word.InsertLocation.before);
                        }
                        break;
                    }
                    case 'findReplace': {
                        const searchResults = body.getRange('Start').getRange('End').search(action.find || '', { matchCase: false, matchWholeWord: false });
                        searchResults.load('items');
                        return context.sync().then(function () {
                            searchResults.items.forEach(function (item) {
                                item.insertText(action.replace || '', Word.InsertLocation.replace);
                            });
                            return context.sync();
                        });
                    }
                    case 'setFont': {
                        const idx = action.paragraphIndex || 0;
                        if (idx < items.length) {
                            const para = items[idx];
                            if (action.name) para.font.name = action.name;
                            if (action.size) para.font.size = action.size;
                            if (action.bold !== undefined) para.font.bold = action.bold;
                            if (action.italic !== undefined) para.font.italic = action.italic;
                            if (action.color) para.font.color = action.color;
                        }
                        break;
                    }
                    case 'setAlignment': {
                        const idx = action.paragraphIndex || 0;
                        if (idx < items.length) {
                            const alignMap = { 'left': Word.Alignment.left, 'center': Word.Alignment.centered, 'right': Word.Alignment.right, 'justify': Word.Alignment.justified };
                            items[idx].alignment = alignMap[action.alignment] || Word.Alignment.left;
                        }
                        break;
                    }
                    case 'insertBullet': {
                        const idx = action.afterParagraph || 0;
                        if (idx < items.length) {
                            const bulletPara = items[idx].insertParagraph('• ' + (action.text || ''), Word.InsertLocation.after);
                            bulletPara.leftIndent = 36;
                        }
                        break;
                    }
                    case 'insertNumberedList': {
                        const idx = action.afterParagraph || 0;
                        if (idx < items.length) {
                            const numPara = items[idx].insertParagraph((action.number || 1) + '. ' + (action.text || ''), Word.InsertLocation.after);
                            numPara.leftIndent = 36;
                        }
                        break;
                    }
                    case 'insertHyperlink': {
                        const idx = action.paragraphIndex || 0;
                        if (idx < items.length) {
                            const para = items[idx];
                            const range = para.getRange('Start');
                            range.insertHyperlink(action.url || '', action.text || 'Link');
                        }
                        break;
                    }
                    case 'insertBookmark': {
                        const idx = action.paragraphIndex || 0;
                        if (idx < items.length) {
                            const para = items[idx];
                            para.insertBookmark(action.name || 'bookmark_' + Date.now());
                        }
                        break;
                    }
                    case 'insertComment': {
                        const idx = action.paragraphIndex || 0;
                        if (idx < items.length) {
                            const para = items[idx];
                            const range = para.getRange('Start');
                            range.insertComment(action.text || 'Comment');
                        }
                        break;
                    }
                    case 'highlight': {
                        const idx = action.paragraphIndex || 0;
                        if (idx < items.length) {
                            const para = items[idx];
                            const colorMap = { 'yellow': '#FFFF00', 'green': '#00FF00', 'blue': '#0000FF', 'red': '#FF0000' };
                            para.font.highlightColor = colorMap[action.color] || '#FFFF00';
                        }
                        break;
                    }
                    case 'clearFormatting': {
                        const idx = action.paragraphIndex || 0;
                        if (idx < items.length) {
                            const para = items[idx];
                            para.font.bold = false;
                            para.font.italic = false;
                            para.font.underline = Word.UnderlineType.none;
                            para.font.highlightColor = '#FFFFFF';
                        }
                        break;
                    }
                    case 'createTOC': {
                        const selection = context.document.getSelection();
                        selection.insertTableOfContents('\\o \"1-5\" \\h \\z \\u', true);
                        break;
                    }
                    case 'updateFields': {
                        const fields = body.fields;
                        fields.load('items');
                        return context.sync().then(function () {
                            fields.items.forEach(function (f) { try { f.update(); } catch (e) {} });
                            return context.sync();
                        });
                    }
                    default:
                        logAgent('  Action chưa hỗ trợ: ' + action.type);
                }

                return context.sync();
            });
        })
        .then(resolve)
        .catch(reject);
    });
}

