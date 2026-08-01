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
    if (parts.length < 2) return 0;
    
    const startStr = parts[0];
    const endStr = parts[1];
    
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
    if (parts.length < 2) return [];
    
    const startStr = parts[0];
    const endStr = parts[1];
    
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
        
        for (const row of rowsArr) {
            const cells = row['w:tc'];
            if (!cells) continue;
            const cellsArr = Array.isArray(cells) ? cells : [cells];
            
            const rowText = cellsArr.map(cell => {
                const tList = [];
                function findTexts(obj) {
                    if (!obj || typeof obj !== 'object') return;
                    if (obj['w:t'] !== undefined) {
                        const val = obj['w:t'];
                        if (typeof val === 'string' || typeof val === 'number') {
                            tList.push(val.toString());
                        } else if (val && val['#text']) {
                            tList.push(val['#text'].toString());
                        } else if (typeof val === 'object') {
                            tList.push('');
                        }
                    }
                    for (const key in obj) {
                        findTexts(obj[key]);
                    }
                }
                findTexts(cell);
                return tList.join('').trim();
            });
            
            if (rowText.length >= 2) {
                const stt = rowText[0].trim();
                const project = rowText[1] ? rowText[1].trim() : '';
                const content = rowText[2] ? rowText[2].trim() : '';
                const position = rowText[3] ? rowText[3].trim() : '';
                const time = rowText[4] ? rowText[4].trim() : '';
                
                const cleanStt = stt.trim().replace(/\.$/, ''); // Bỏ dấu chấm ở cuối nếu có
                const isNumber = /^\d+$/.test(stt);
                const isRoman = /^[IVXLCDM]+$/i.test(cleanStt) && cleanStt.length > 0;
                const hasTime = /[\d]{1,2}[\/-][\d]{4}/.test(time);
                
                if (isNumber && hasTime && rowText.length >= 5) {
                    const months = calculateMonths(time);
                    parsedData.push({
                        stt: parseInt(stt, 10),
                        project,
                        content,
                        position,
                        time,
                        months
                    });
                } else if (isRoman && project) {
                    parsedData.push({
                        stt: cleanStt,
                        project,
                        content: '',
                        position: '',
                        time: '',
                        months: 0,
                        isHeader: true,
                        isSubHeader: true
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
            return res.status(400).json({ error: 'Định dạng file không hỗ trợ. Vui lòng chọn file .docx hoặc .pdf.' });
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
            { header: 'Thời gian công tác (tháng)', key: 'months', width: 25 }
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
                months: isHeader ? (item.months !== undefined ? item.months + ' tháng' : '0 tháng') : (parseInt(item.months, 10) || 0)
            });
            
            const rowIndex = row.number;
            
            if (isHeader) {
                // Gộp các cột Tên dự án, Nội dung, Vị trí, Thời gian (B đến E)
                worksheet.mergeCells(`B${rowIndex}:E${rowIndex}`);
                
                row.font = { name: 'Arial', size: 10, bold: true };
                row.alignment = { vertical: 'middle', wrapText: true };
                row.getCell('stt').alignment = { vertical: 'middle', horizontal: 'center' };
                row.getCell('project').alignment = { vertical: 'middle', horizontal: 'left' };
                row.getCell('months').alignment = { vertical: 'middle', horizontal: 'center' };
                
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
                row.getCell('months').alignment = { vertical: 'middle', horizontal: 'center' };
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
            time: 'Tổng cộng thời gian:',
            months: { formula: `=SUM(F2:F${lastRowIndex - 1})` }
        });
        
        totalRow.font = { name: 'Arial', size: 10, bold: true };
        totalRow.getCell('time').alignment = { vertical: 'middle', horizontal: 'right' };
        totalRow.getCell('months').alignment = { vertical: 'middle', horizontal: 'center' };
        
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
    // Tự động mở trình duyệt trên Windows
    if (process.platform === 'win32') {
        exec(`start http://localhost:${PORT}`);
    }
});
