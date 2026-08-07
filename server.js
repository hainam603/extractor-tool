const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const AdmZip = require('adm-zip');
const { XMLParser } = require('fast-xml-parser');
const pdfParse = require('pdf-parse');
const ExcelJS = require('exceljs');
const { exec } = require('child_process');

const appDir = process.pkg ? path.dirname(process.execPath) : __dirname;
const exportsDir = path.join(appDir, 'exports');

const app = express();
const PORT = process.env.PORT || 3000;

// Cấu hình multer để lưu trữ file tạm thời trong bộ nhớ
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json());
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));
app.use('/exports', express.static(exportsDir));

// Hàm helper split khoảng thời gian thông minh hỗ trợ định dạng MM-YYYY và MM/YYYY
function splitDateRange(timeStr) {
    if (!timeStr) return [];
    const cleanStr = timeStr.trim();
    // Match các mốc thời gian dạng MM/YYYY, MM-YYYY, DD/MM/YYYY, DD-MM-YYYY hoặc hiện tại
    const dateRegex = /(?:(?:\d{1,2})[\/-])?\d{1,2}[\/-]\d{4}|hiện\s*tại/ig;
    const matches = cleanStr.match(dateRegex);
    if (matches && matches.length >= 2) {
        return [matches[0], matches[1]];
    }
    // Fallback split thông thường
    const cleanNoSpace = cleanStr.replace(/\s+/g, '');
    return cleanNoSpace.split(/[-–]|đến/i);
}

// Hàm tính số tháng công tác từ chuỗi thời gian
function calculateMonths(timeStr) {
    if (!timeStr) return 0;
    
    const parts = splitDateRange(timeStr);
    if (parts.length === 0) return 0;
    
    const startStr = parts[0];
    const endStr = parts.length >= 2 ? parts[1] : startStr;
    
    function parseDate(str) {
        // Hỗ trợ cả dấu gạch chéo / và gạch ngang -
        const regex = /(?:(\d{1,2})[\/-])?(\d{1,2})[\/-](\d{4})/;
        const match = str.match(regex);
        if (match) {
            return {
                month: parseInt(match[2], 10),
                year: parseInt(match[3], 10)
            };
        }
        if (/hiện\s*tại/i.test(str)) {
            const now = new Date();
            return {
                month: now.getMonth() + 1,
                year: now.getFullYear()
            };
        }
        return null;
    }
    
    const start = parseDate(startStr);
    const end = parseDate(endStr);
    
    if (!start || !end) return 0;
    
    const months = (end.year - start.year) * 12 + (end.month - start.month) + 1;
    return months > 0 ? months : 0;
}

// Hàm trả về danh sách các tháng dạng số nguyên duy nhất đại diện cho khoảng thời gian
function getMonthsList(timeStr) {
    if (!timeStr) return [];
    
    const parts = splitDateRange(timeStr);
    if (parts.length === 0) return [];
    
    const startStr = parts[0];
    const endStr = parts.length >= 2 ? parts[1] : startStr;
    
    function parseDate(str) {
        const regex = /(?:(\d{1,2})[\/-])?(\d{1,2})[\/-](\d{4})/;
        const match = str.match(regex);
        if (match) {
            return {
                month: parseInt(match[2], 10),
                year: parseInt(match[3], 10)
            };
        }
        if (/hiện\s*tại/i.test(str)) {
            const now = new Date();
            return {
                month: now.getMonth() + 1,
                year: now.getFullYear()
            };
        }
        return null;
    }
    
    const start = parseDate(startStr);
    const end = parseDate(endStr);
    
    if (!start || !end) return [];
    
    const startVal = start.year * 12 + start.month;
    const endVal = end.year * 12 + end.month;
    
    const list = [];
    for (let val = Math.min(startVal, endVal); val <= Math.max(startVal, endVal); val++) {
        list.push(val);
    }
    return list;
}

// Logic parse DOCX
function extractDataFromDocx(fileBuffer) {
    const zip = new AdmZip(fileBuffer);
    const docXmlEntry = zip.getEntry('word/document.xml');
    if (!docXmlEntry) {
        throw new Error('Không tìm thấy tệp word/document.xml trong file docx.');
    }
    const xmlText = docXmlEntry.getData().toString('utf8');
    
    const parser = new XMLParser({
        ignoreAttributes: false,
        attributeNamePrefix: "@_",
        parseAttributeValue: false,
        parseTagValue: false
    });
    const jsonObj = parser.parse(xmlText);
    
    const body = jsonObj['w:document']?.['w:body'];
    if (!body) return [];
    
    const tables = [];
    function findTables(obj) {
        if (!obj || typeof obj !== 'object') return;
        if (obj['w:tbl']) {
            if (Array.isArray(obj['w:tbl'])) {
                tables.push(...obj['w:tbl']);
            } else {
                tables.push(obj['w:tbl']);
            }
        }
        for (const key in obj) {
            if (key !== 'w:tbl') {
                findTables(obj[key]);
            }
        }
    }
    findTables(body);
    
    const parsedData = [];
    
    for (const table of tables) {
        const rows = table['w:tr'];
        if (!rows) continue;
        const rowsArr = Array.isArray(rows) ? rows : [rows];
        
        // Mảng theo dõi giá trị merge dọc của từng cột
        const lastMergedValues = [];
        let currentAutoStt = 1;
        
        for (const row of rowsArr) {
            const cells = row['w:tc'];
            if (!cells) continue;
            const cellsArr = Array.isArray(cells) ? cells : [cells];
            
            // Parse text của từng cell, có xử lý vMerge
            const rowText = cellsArr.map((cell, cIdx) => {
                const tList = [];
                function findTexts(obj) {
                    if (!obj || typeof obj !== 'object') return;
                    if (obj['w:t'] !== undefined) {
                        const val = obj['w:t'];
                        if (typeof val === 'string' || typeof val === 'number') {
                            tList.push(val.toString());
                        } else if (val && val['#text']) {
                            tList.push(val['#text'].toString());
                        }
                    }
                    for (const key in obj) {
                        findTexts(obj[key]);
                    }
                }
                findTexts(cell);
                let text = tList.join('').trim();
                
                // Xử lý vMerge (gộp ô dọc)
                const tcPr = cell['w:tcPr'];
                if (tcPr && tcPr['w:vMerge'] !== undefined) {
                    const vMerge = tcPr['w:vMerge'];
                    // Nếu là restart, ghi nhận giá trị merge mới
                    if (vMerge && vMerge['@_w:val'] === 'restart') {
                        lastMergedValues[cIdx] = text;
                    } else {
                        // Nếu là continue hoặc thẻ rỗng, thừa hưởng giá trị cũ
                        text = lastMergedValues[cIdx] || '';
                    }
                } else {
                    // Không có vMerge, xóa trạng thái merge của cột này
                    lastMergedValues[cIdx] = null;
                }
                
                return text;
            });
            
            if (rowText.length >= 2) {
                const sttRaw = rowText[0].trim();
                const cleanStt = sttRaw.replace(/\.$/, ''); // Bỏ dấu chấm ở cuối nếu có
                const isRoman = /^[IVXLCDM]+$/i.test(cleanStt) && cleanStt.length > 0;
                
                // Kiểm tra xem dòng này có chứa thời gian hay không
                const hasTime = rowText.some(cell => cell && /[\d]{1,2}[\/-][\d]{4}/.test(cell));
                
                if (isRoman && !hasTime) {
                    parsedData.push({
                        stt: cleanStt,
                        project: rowText[1] ? rowText[1].trim() : '',
                        content: '',
                        position: '',
                        time: '',
                        months: 0,
                        isHeader: true,
                        isSubHeader: true
                    });
                    continue;
                }
            }

            if (rowText.length >= 5) {
                const sttRaw = rowText[0].trim();
                const project = rowText[1] ? rowText[1].trim() : '';
                const content = rowText[2] ? rowText[2].trim() : '';
                const position = rowText[3] ? rowText[3].trim() : '';
                const time = rowText[4] ? rowText[4].trim() : '';
                
                // Bỏ qua các dòng tiêu đề
                if (sttRaw.toLowerCase() === 'stt' || project.toLowerCase() === 'tên đề án, dự án, thiết kế kỹ thuật - dự toán nhiệm vụ đo đạcvà bản đồ' || project.toLowerCase().includes('tên đề án') || project.toLowerCase().includes('tên đề án, dự án')) {
                    continue;
                }
                
                const cleanStt = sttRaw.replace(/\.$/, ''); // Bỏ dấu chấm ở cuối nếu có
                const isNumber = /^\d+$/.test(cleanStt);
                const hasTime = /[\d]{1,2}[\/-][\d]{4}/.test(time);
                
                // Dòng hợp lệ: có thời gian và (có project hoặc content)
                if (hasTime && (project || content)) {
                    let sttVal = currentAutoStt;
                    if (isNumber) {
                        sttVal = parseInt(cleanStt, 10);
                        currentAutoStt = sttVal + 1;
                    } else {
                        currentAutoStt++;
                    }
                    
                    const months = calculateMonths(time);
                    parsedData.push({
                        stt: sttVal,
                        project,
                        content,
                        position,
                        time,
                        months
                    });
                }
            }
        }
    }
    return parsedData;
}

// Hàm trích xuất thông tin cá nhân từ file Word
function extractPersonalInfo(fileBuffer) {
    try {
        const zip = new AdmZip(fileBuffer);
        const docXmlEntry = zip.getEntry('word/document.xml');
        if (!docXmlEntry) return null;
        
        const xmlText = docXmlEntry.getData().toString('utf8');
        // Loại bỏ thẻ XML để lấy văn bản thuần
        const cleanText = xmlText.replace(/<[^>]+>/g, ' ');
        
        const info = {
            fullName: '',
            dob: '',
            cccd: '',
            qualification: ''
        };
        
        const nameMatch = cleanText.match(/Họ\s+và\s+tên\s*:\s*([^\r\n;.]+)/i) || cleanText.match(/Họ\s+tên\s*:\s*([^\r\n;.]+)/i);
        const dobMatch = cleanText.match(/Ngày\s+tháng\s+năm\s+sinh\s*:\s*([\d/.-]+)/i) || cleanText.match(/Ngày\s+sinh\s*:\s*([\d/.-]+)/i);
        const cccdMatch = cleanText.match(/Số\s+thẻ\s+căn\s+cước\s+công\s+dân\s*:\s*([0-9]+)/i) || cleanText.match(/Căn\s+cước\s+công\s+dân\s*:\s*([0-9]+)/i) || cleanText.match(/Số\s+CCCD\s*:\s*([0-9]+)/i);
        const qualMatch = cleanText.match(/Trình\s+độ\s+chuyên\s+môn\s*:\s*([^\r\n.]+)/i);
        
        if (nameMatch) info.fullName = nameMatch[1].replace(/\s+/g, ' ').trim();
        if (dobMatch) info.dob = dobMatch[1].trim();
        if (cccdMatch) info.cccd = cccdMatch[1].trim();
        if (qualMatch) info.qualification = qualMatch[1].replace(/\s+/g, ' ').trim();
        
        return info;
    } catch (err) {
        console.error('Lỗi khi trích xuất thông tin cá nhân:', err);
        return null;
    }
}

// API endpoint để upload và phân tích file
app.post('/api/upload', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'Vui lòng chọn một file.' });
        }
        
        const ext = path.extname(req.file.originalname).toLowerCase();
        let data = [];
        
        if (ext === '.docx') {
            data = extractDataFromDocx(req.file.buffer);
            const personalInfo = extractPersonalInfo(req.file.buffer);
            res.json({ data, personalInfo });
            return;
        } else if (ext === '.doc') {
            // Tạo tên file tạm ngẫu nhiên trong thư mục exports
            const tempId = Date.now() + '_' + Math.floor(Math.random() * 1000);
            if (!fs.existsSync(exportsDir)) {
                fs.mkdirSync(exportsDir);
            }
            const tempDocPath = path.join(exportsDir, `temp_${tempId}.doc`);
            const tempDocxPath = path.join(exportsDir, `temp_${tempId}.docx`);
            
            // Ghi buffer vào file tạm .doc
            fs.writeFileSync(tempDocPath, req.file.buffer);
            
            const docFileAbs = path.resolve(tempDocPath);
            const docxFileAbs = path.resolve(tempDocxPath);
            
            // Chạy lệnh PowerShell convert sang .docx dùng MS Word COM Object
            const psCommand = `powershell -Command "$word = New-Object -ComObject Word.Application; $word.Visible = $false; $doc = $word.Documents.Open('${docFileAbs}'); $doc.SaveAs([ref] '${docxFileAbs}', [ref] 16); $doc.Close(); $word.Quit();"`;
            
            try {
                await new Promise((resolve, reject) => {
                    exec(psCommand, (err, stdout, stderr) => {
                        if (err) {
                            reject(err);
                        } else {
                            resolve();
                        }
                    });
                });
                
                if (!fs.existsSync(tempDocxPath)) {
                    throw new Error('Chuyển đổi file .doc sang .docx không thành công.');
                }
                
                const docxBuffer = fs.readFileSync(tempDocxPath);
                data = extractDataFromDocx(docxBuffer);
                const personalInfo = extractPersonalInfo(docxBuffer);
                
                res.json({ data, personalInfo });
            } finally {
                // Đảm bảo luôn dọn dẹp các file tạm
                try {
                    if (fs.existsSync(tempDocPath)) fs.unlinkSync(tempDocPath);
                    if (fs.existsSync(tempDocxPath)) fs.unlinkSync(tempDocxPath);
                } catch (unlinkErr) {
                    console.error('Lỗi khi xóa file tạm:', unlinkErr);
                }
            }
            return;
        } else if (ext === '.pdf') {
            // Thử trích xuất text từ file PDF
            const pdfData = await pdfParse(req.file.buffer);
            const text = pdfData.text.trim();
            
            if (text.length === 0) {
                return res.json({ 
                    warning: 'File PDF này không chứa lớp văn bản (quét từ ảnh). Bạn có thể tự điền tay dữ liệu vào bảng dưới đây hoặc chuyển đổi PDF sang DOCX trước.',
                    data: [] 
                });
            }
            
            res.json({ 
                warning: 'Để có kết quả phân tích bảng chính xác nhất, vui lòng sử dụng file định dạng DOCX. Đối với file PDF, bạn có thể tự nhập dữ liệu bên dưới.',
                data: [] 
            });
            return;
        } else {
            return res.status(400).json({ error: 'Định dạng file không hỗ trợ. Vui lòng chọn file .docx, .doc hoặc .pdf.' });
        }
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Đã xảy ra lỗi khi xử lý file: ' + err.message });
    }
});

// API endpoint để xuất file Excel
app.post('/api/export', async (req, res) => {
    try {
        let items = req.body.items;
        if (req.body.itemsJson) {
            try {
                items = JSON.parse(req.body.itemsJson);
            } catch (jsonErr) {
                return res.status(400).json({ error: 'Dữ liệu JSON không hợp lệ.' });
            }
        }
        
        if (!items || !Array.isArray(items)) {
            return res.status(400).json({ error: 'Dữ liệu không hợp lệ.' });
        }
        
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Quá trình công tác');
        
        // Định nghĩa các cột
        worksheet.columns = [
            { header: 'STT', key: 'stt', width: 8 },
            { header: 'Tên đề án, dự án, thiết kế kỹ thuật - dự toán nhiệm vụ', key: 'project', width: 45 },
            { header: 'Nội dung công việc đã tham gia', key: 'content', width: 45 },
            { header: 'Vị trí đảm nhiệm', key: 'position', width: 20 },
            { header: 'Thời gian tham gia', key: 'time', width: 20 },
            { header: 'Số tháng', key: 'originalMonths', width: 15 },
            { header: 'Số tháng thực tính', key: 'actualMonths', width: 20 }
        ];
        
        // Định dạng tiêu đề cột (Header styling)
        const headerRow = worksheet.getRow(1);
        headerRow.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF1F497D' } // Màu xanh navy đậm sang trọng
        };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        headerRow.height = 35;
        
        // Thêm dữ liệu
        items.forEach((item) => {
            const isHeader = !!item.isHeader;
            const row = worksheet.addRow({
                stt: item.stt,
                project: item.project,
                content: isHeader ? '' : item.content,
                position: isHeader ? '' : item.position,
                time: isHeader ? '' : item.time,
                originalMonths: isHeader ? '' : (parseInt(item.originalMonths, 10) || 0),
                actualMonths: isHeader ? (item.months !== undefined ? item.months + ' tháng' : '0 tháng') : (parseInt(item.months, 10) || 0)
            });
            
            const rowIndex = row.number;
            
            if (isHeader) {
                // Gộp các cột Tên dự án, Nội dung, Vị trí, Thời gian (B đến E)
                worksheet.mergeCells(`B${rowIndex}:E${rowIndex}`);
                
                row.font = { name: 'Arial', size: 10, bold: true };
                row.alignment = { vertical: 'middle', wrapText: true };
                row.getCell('stt').alignment = { vertical: 'middle', horizontal: 'center' };
                row.getCell('project').alignment = { vertical: 'middle', horizontal: 'left' };
                row.getCell('actualMonths').alignment = { vertical: 'middle', horizontal: 'center' };
                
                row.eachCell(cell => {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFF2F2F2' }
                    };
                });
            } else {
                // Định dạng dòng thường
                row.font = { name: 'Arial', size: 10 };
                row.alignment = { vertical: 'middle', wrapText: true };
                
                row.getCell('stt').alignment = { vertical: 'middle', horizontal: 'center' };
                row.getCell('time').alignment = { vertical: 'middle', horizontal: 'center' };
                row.getCell('originalMonths').alignment = { vertical: 'middle', horizontal: 'center' };
                row.getCell('actualMonths').alignment = { vertical: 'middle', horizontal: 'center' };
            }
        });
        
        // Vẽ viền (border) cho tiêu đề và dữ liệu
        const borderStyle = {
            top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
            left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
            bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
            right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
        };
        
        worksheet.eachRow((row, rowNumber) => {
            row.eachCell(cell => {
                cell.border = borderStyle;
            });
        });
        
        // Thêm dòng tổng cộng
        const lastRowIndex = items.length + 2; // +1 cho header, +1 cho dòng mới
        const totalRow = worksheet.addRow({
            stt: '',
            project: '',
            content: '',
            position: '',
            time: 'Tổng cộng thời gian thực tính:',
            originalMonths: '',
            actualMonths: { formula: `=SUM(G2:G${lastRowIndex - 1})` }
        });
        
        totalRow.font = { name: 'Arial', size: 10, bold: true };
        totalRow.getCell('time').alignment = { vertical: 'middle', horizontal: 'right' };
        totalRow.getCell('actualMonths').alignment = { vertical: 'middle', horizontal: 'center' };
        
        // Định dạng dòng tổng cộng nổi bật hơn
        totalRow.eachCell(cell => {
            cell.border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'double', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
                right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
            };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFF2F2F2' }
            };
        });
        
        // Ghi file ra thư mục exports
        const now = new Date();
        const timestamp = now.getFullYear() + 
                          String(now.getMonth() + 1).padStart(2, '0') + 
                          String(now.getDate()).padStart(2, '0') + '_' +
                          String(now.getHours()).padStart(2, '0') + 
                          String(now.getMinutes()).padStart(2, '0') + 
                          String(now.getSeconds()).padStart(2, '0');
        const fileName = `Qua_trinh_cong_tac_${timestamp}.xlsx`;
        
        if (!fs.existsSync(exportsDir)) {
            fs.mkdirSync(exportsDir);
        }
        const filePath = path.join(exportsDir, fileName);
        await workbook.xlsx.writeFile(filePath);
        
        res.json({ success: true, fileName: fileName, url: '/exports/' + fileName });
        
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'Đã xảy ra lỗi khi xuất file Excel: ' + err.message });
    }
});

// API mở file Excel trên local
app.get('/api/open-file', (req, res) => {
    const fileName = req.query.file;
    if (!fileName) {
        return res.status(400).json({ error: 'Thiếu tên file.' });
    }
    const safeName = path.basename(fileName);
    const filePath = path.join(exportsDir, safeName);
    
    if (!fs.existsSync(filePath)) {
        return res.status(404).json({ error: 'File không tồn tại.' });
    }
    
    exec(`start "" "${filePath}"`, (err) => {
        if (err) {
            console.error('Lỗi khi mở file:', err);
            return res.status(500).json({ error: 'Không thể mở file: ' + err.message });
        }
        res.json({ success: true });
    });
});

// Khởi động server
app.listen(PORT, () => {
    console.log(`Server đang chạy tại http://localhost:${PORT}`);
});
