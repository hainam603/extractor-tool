const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const workspaceCard = document.getElementById('workspaceCard');
const tableBody = document.getElementById('tableBody');
const exportBtn = document.getElementById('exportBtn');
const loadingOverlay = document.getElementById('loadingOverlay');
const loadingText = document.getElementById('loadingText');
const alertBox = document.getElementById('alertBox');
const alertMessage = document.getElementById('alertMessage');
const summaryCard = document.getElementById('summaryCard');
const summaryList = document.getElementById('summaryList');
const summaryTotal = document.getElementById('summaryTotal');
const gopHoSoCheckbox = document.getElementById('gopHoSoCheckbox');
let tableItems = [];
let isGopHoSo = false;

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

// Hàm helper tính số tháng
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

// Show loading overlay
function showLoading(text) {
    loadingText.innerText = text || 'Đang xử lý dữ liệu...';
    loadingOverlay.style.display = 'flex';
}

// Hide loading overlay
function hideLoading() {
    loadingOverlay.style.display = 'none';
}

// Show alert
function showAlert(message, type = 'warning') {
    alertBox.className = `alert alert-${type}`;
    alertMessage.innerHTML = message;
    alertBox.style.display = 'flex';
    // Tự cuộn lên đầu để xem thông báo
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// Hide alert
function hideAlert() {
    alertBox.style.display = 'none';
}

// Sự kiện Drag & Drop
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFiles(files);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFiles(e.target.files);
    }
});

// Sửa lỗi nhận diện nhầm chữ O/o thành số 0 trong các chuỗi ngày tháng
function cleanOcrDates(str) {
    return str.replace(/([Oo0-9]{1,2})\/([Oo0-9]{4})/g, (match, p1, p2) => {
        const cleanMonth = p1.replace(/[Oo]/g, '0');
        const cleanYear = p2.replace(/[Oo]/g, '0');
        return `${cleanMonth}/${cleanYear}`;
    });
}

// Tiền xử lý văn bản OCR: ghép các khoảng thời gian bị ngắt dòng
function preprocessOcrText(text) {
    let cleanText = cleanOcrDates(text);
    
    const lines = cleanText.split('\n');
    const mergedLines = [];
    
    for (let i = 0; i < lines.length; i++) {
        let line = lines[i].trim();
        if (!line) continue;
        
        // TH1: Dòng kết thúc bằng dấu gạch ngang và dòng tiếp theo bắt đầu bằng ngày tháng
        if (line.endsWith('-') && i + 1 < lines.length) {
            const nextLine = lines[i+1].trim();
            if (/^\d{1,2}\/\d{4}/.test(nextLine)) {
                line += nextLine;
                i++; // Bỏ qua dòng tiếp theo
            }
        }
        // TH2: Dòng kết thúc bằng ngày tháng, dòng dưới bắt đầu bằng dấu gạch ngang và ngày tháng
        else if (/\d{1,2}\/\d{4}$/.test(line) && i + 1 < lines.length) {
            const nextLine = lines[i+1].trim();
            if (nextLine.startsWith('-') && /^\s*-\s*\d{1,2}\/\d{4}/.test(nextLine)) {
                line += nextLine;
                i++; // Bỏ qua dòng tiếp theo
            }
        }
        // TH3: Dòng kết thúc bằng ngày tháng, dòng dưới chỉ là ngày tháng (mất dấu gạch ngang do OCR)
        else if (/\d{1,2}\/\d{4}$/.test(line) && i + 1 < lines.length) {
            const nextLine = lines[i+1].trim();
            if (/^\d{1,2}\/\d{4}$/.test(nextLine)) {
                line += "-" + nextLine;
                i++; // Bỏ qua dòng tiếp theo
            }
        }
        
        mergedLines.push(line);
    }
    
    return mergedLines.join('\n');
}

// Hàm parse văn bản OCR sang bảng
function parseOcrText(text) {
    // Chạy tiền xử lý để gộp các dòng bị ngắt và sửa lỗi OCR ngày tháng
    const processedText = preprocessOcrText(text);
    
    const lines = processedText.split('\n');
    const parsedData = [];
    const dateRangeRegex = /(?:(\d{1,2})[\/-])?(\d{2})[\/-](\d{4})\s*(?:[-–đ]|đến)\s*(?:(?:(\d{1,2})[\/-])?(\d{2})[\/-](\d{4})|hiện\s*tại)/i;
    let sttCounter = 1;
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;
        
        const match = line.match(dateRangeRegex);
        if (match) {
            const timeStr = match[0].trim();
            let leftText = line.substring(0, line.indexOf(timeStr)).trim();
            
            let stt = sttCounter;
            const sttMatch = leftText.match(/^(\d+)[\s.,/-]+/);
            if (sttMatch) {
                stt = parseInt(sttMatch[1], 10);
                leftText = leftText.substring(sttMatch[0].length).trim();
            }
            
            let position = "Nhân viên kỹ thuật";
            const posRegex = /(nhân\s*viên(?:\s*kỹ\s*thuật|\s*đo\s*đạc)?|kỹ\s*sư|trưởng\s*nhóm|chuyên\s*viên)/i;
            const posMatch = leftText.match(posRegex);
            if (posMatch) {
                position = posMatch[0].trim();
                leftText = leftText.substring(0, leftText.indexOf(position)).trim();
            }
            
            let project = leftText;
            let content = "";
            const contentKeywords = /(trích\s*đo|lập\s*lưới|khảo\s*sát|đo\s*vẽ|đo\s*đạc|thi\s*công|thu\s*thập)/i;
            const contentIdx = leftText.search(contentKeywords);
            if (contentIdx > 0) {
                project = leftText.substring(0, contentIdx).trim();
                content = leftText.substring(contentIdx).trim();
            }
            
            if (project.endsWith('.')) project = project.slice(0, -1).trim();
            if (project.endsWith(',')) project = project.slice(0, -1).trim();
            
            parsedData.push({
                stt: stt,
                project: project,
                content: content,
                position: position,
                time: timeStr,
                months: calculateMonths(timeStr)
            });
            sttCounter = stt + 1;
        }
    }
    return parsedData;
}

// Chuyển đổi số sang số La Mã để phân nhóm file
function romanize(num) {
    const roman = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV', 'XV'];
    return roman[num - 1] || num.toString();
}

// Xử lý upload nhiều file và phân tích hàng loạt
async function handleFiles(files) {
    hideAlert();
    let allItems = [];
    let hasWarning = false;
    let warningMsg = '';
    
    // Lọc danh sách file Word (.docx, .doc) hợp lệ
    const validFiles = Array.from(files).filter(file => {
        const ext = file.name.split('.').pop().toLowerCase();
        return ext === 'docx' || ext === 'doc';
    });
    
    if (validFiles.length === 0) {
        showAlert('Không tìm thấy file Word (.docx, .doc) hợp lệ nào. Vui lòng chọn các file có định dạng .docx hoặc .doc.', 'error');
        return;
    }

    showLoading(`Đang phân tích dữ liệu từ ${validFiles.length} file Word...`);
    
    try {
        for (let i = 0; i < validFiles.length; i++) {
            const file = validFiles[i];
            const formData = new FormData();
            formData.append('file', file);
            
            const response = await fetch('/api/upload', {
                method: 'POST',
                body: formData
            });
            const result = await response.json();
            
            if (!response.ok) {
                throw new Error(result.error || `Lỗi khi xử lý file ${file.name}`);
            }
            
            if (result.warning) {
                hasWarning = true;
                warningMsg += `<strong>${file.name}</strong>: ${result.warning}<br>`;
            }
            
            const data = result.data || [];
            if (data.length > 0) {
                // Tạo dòng phân nhóm bằng tên file và thông tin cá nhân
                const personalInfo = result.personalInfo;
                let infoStr = '';
                if (personalInfo) {
                    const infoParts = [];
                    if (personalInfo.fullName) infoParts.push(`Họ tên: ${personalInfo.fullName}`);
                    if (personalInfo.dob) infoParts.push(`Ngày sinh: ${personalInfo.dob}`);
                    if (personalInfo.cccd) infoParts.push(`CCCD: ${personalInfo.cccd}`);
                    if (personalInfo.qualification) infoParts.push(`Trình độ: ${personalInfo.qualification}`);
                    if (infoParts.length > 0) {
                        infoStr = ` - [${infoParts.join(' | ')}]`;
                    }
                }

                const headerIndex = allItems.filter(item => item.isParentHeader).length + 1;
                allItems.push({
                    stt: romanize(headerIndex) + '.',
                    project: `HỒ SƠ: ${file.name.replace(/\.docx$/i, '')}${infoStr}`,
                    isHeader: true,
                    isParentHeader: true
                });
                
                // Thêm các dòng dữ liệu của file
                data.forEach(item => {
                    const parts = (item.time || '').split(/[-–]|đến/i);
                    allItems.push({
                        ...item,
                        startTime: parts[0] ? parts[0].trim() : '',
                        endTime: parts[1] ? parts[1].trim() : '',
                        isSubHeader: item.isSubHeader || (item.isHeader && !item.isParentHeader)
                    });
                });
            }
        }
        
        if (allItems.length > 0) {
            // Đánh lại số thứ tự (stt) cho các dòng thường liên tục
            let normalRowCounter = 1;
            allItems.forEach(item => {
                if (!item.isHeader) {
                    item.stt = normalRowCounter++;
                }
            });
            
            tableItems = allItems;
            renderTable();
            workspaceCard.style.display = 'block';
            
            if (hasWarning) {
                showAlert(warningMsg, 'warning');
            }
        } else {
            showAlert('Không trích xuất được dữ liệu từ các file đã chọn.', 'error');
        }
    } catch (err) {
        console.error(err);
        showAlert(err.message, 'error');
    } finally {
        hideLoading();
    }
}

// Hàm helper tự động giãn chiều cao cho textarea để ẩn thanh cuộn thô
function autoResizeTextarea(textarea) {
    textarea.style.height = 'auto';
    textarea.style.height = textarea.scrollHeight + 'px';
}

// Render dữ liệu bảng lên HTML
function renderTable() {
    tableBody.innerHTML = '';
    
    if (tableItems.length === 0) {
        summaryCard.style.display = 'none';
        return;
    }
    
    const globalMonthsSet = new Set();
    let currentHoSoMonthsSet = new Set();
    
    const stats = [];
    let currentHoSo = null;
    let currentHoSoActualMonths = 0;
    let currentSubGroup = null;
    let currentSubGroupIsLocTrung = true;
    let normalRowCounter = 1;

    // Phân tích trùng lặp trước cho tất cả các dòng
    tableItems.forEach((item, index) => {
        const isHeader = !!item.isHeader;
        const isParentHeader = !!item.isParentHeader;
        const isSubHeader = !!item.isSubHeader;
        
        if (isParentHeader) {
            if (currentHoSo) {
                currentHoSo.months = currentHoSoActualMonths; // Gán ngược lại số tháng của Parent Header
                stats.push({
                    name: currentHoSo.project,
                    months: currentHoSoActualMonths,
                    subGroups: currentHoSo.subGroups || []
                });
            }
            currentHoSo = item;
            currentHoSo.subGroups = [];
            currentHoSo.months = 0;
            currentHoSoActualMonths = 0;
            currentSubGroup = null;
            currentSubGroupIsLocTrung = true;
            normalRowCounter = 1; // Reset STT dòng thường về 1 khi bắt đầu hồ sơ mới
            
            if (!isGopHoSo) {
                currentHoSoMonthsSet = new Set(); // Reset set nếu không gộp
            }
            
            item.actualMonths = 0;
            item.isOverlapped = false;
            item.isPartialOverlapped = false;
            return;
        }
        
        if (isSubHeader) {
            // Không reset normalRowCounter, không reset currentHoSo
            item.actualMonths = 0;
            item.isOverlapped = false;
            item.isPartialOverlapped = false;
            item.months = 0; // Khởi tạo số tháng của Sub Header
            
            currentSubGroup = item;
            currentSubGroupIsLocTrung = (item.isLocTrung !== false); // Lấy trạng thái lọc trùng của nhóm nhỏ
            
            const subStat = {
                stt: item.stt,
                name: item.project,
                months: 0,
                itemIndex: index,
                itemRef: item
            };
            item.subStatRef = subStat;
            
            if (currentHoSo) {
                if (!currentHoSo.subGroups) {
                    currentHoSo.subGroups = [];
                }
                currentHoSo.subGroups.push(subStat);
            }
            return;
        }

        const monthsList = getMonthsList(item.time);
        const originalMonths = monthsList.length;
        
        let newMonthsCount = originalMonths;
        let isOverlapped = false;
        let isPartialOverlapped = false;

        const activeSet = isGopHoSo ? globalMonthsSet : currentHoSoMonthsSet;

        if (currentSubGroupIsLocTrung) {
            newMonthsCount = 0;
            monthsList.forEach(m => {
                if (!activeSet.has(m)) {
                    newMonthsCount++;
                }
            });

            isOverlapped = originalMonths > 0 && newMonthsCount === 0;
            isPartialOverlapped = originalMonths > 0 && newMonthsCount > 0 && newMonthsCount < originalMonths;
        }

        // Luôn add vào activeSet để làm mốc trùng lặp cho các dòng khác so sánh
        monthsList.forEach(m => {
            activeSet.add(m);
            if (isGopHoSo) {
                globalMonthsSet.add(m);
            }
        });

        currentHoSoActualMonths += newMonthsCount;
        
        if (currentSubGroup) {
            currentSubGroup.months = (currentSubGroup.months || 0) + newMonthsCount;
            if (currentSubGroup.subStatRef) {
                currentSubGroup.subStatRef.months = currentSubGroup.months;
            }
        }

        item.stt = normalRowCounter++;
        item.originalMonths = originalMonths;
        item.actualMonths = newMonthsCount;
        item.months = newMonthsCount; // Đồng bộ ghi đè để xuất Excel chính xác
        item.isOverlapped = isOverlapped;
        item.isPartialOverlapped = isPartialOverlapped;
    });

    // Lưu hồ sơ cuối cùng
    if (currentHoSo) {
        currentHoSo.months = currentHoSoActualMonths; // Gán ngược lại số tháng của Parent Header cuối cùng
        stats.push({
            name: currentHoSo.project,
            months: currentHoSoActualMonths,
            subGroups: currentHoSo.subGroups || []
        });
    }

    // Tính tổng số tháng thực tế
    let totalActualMonths = 0;
    if (isGopHoSo) {
        totalActualMonths = globalMonthsSet.size;
    } else {
        totalActualMonths = stats.reduce((sum, hoso) => sum + hoso.months, 0);
    }

    // Cập nhật vùng thống kê lên trên đầu
    if (stats.length > 0) {
        summaryCard.style.display = 'block';
        summaryList.innerHTML = stats.map(h => {
            const displayName = h.name.replace(/^HỒ\s+SƠ:\s*/i, '');
            const hYears = Math.floor(h.months / 12);
            const hRem = h.months % 12;
            const hQuyDoi = hYears > 0 ? ` (${hYears} năm ${hRem} tháng)` : ` (${hRem} tháng)`;
            
            let subGroupsHtml = '';
            if (h.subGroups && h.subGroups.length > 0) {
                subGroupsHtml = h.subGroups.map(sg => {
                    const sgYears = Math.floor(sg.months / 12);
                    const sgRem = sg.months % 12;
                    const sgQuyDoi = sgYears > 0 ? ` (${sgYears} năm ${sgRem} tháng)` : ` (${sgRem} tháng)`;
                    return `<div style="display: flex; justify-content: space-between; padding: 2px 0 2px 24px; font-size: 0.85rem; color: #475569; align-items: center;">
                        <span style="display: inline-flex; align-items: center; gap: 6px;">
                            <input type="checkbox" class="subgroup-loc-trung-checkbox" data-index="${sg.itemIndex}" ${sg.itemRef.isLocTrung !== false ? 'checked' : ''} style="width: 14px; height: 14px; cursor: pointer;">
                            🔹 ${sg.stt}. ${sg.name}
                        </span>
                        <span style="font-weight: 500; color: #0f172a;">${sg.months} tháng${sgQuyDoi}</span>
                    </div>`;
                }).join('');
            }

            return `<div style="padding: 8px 0; border-bottom: 1px dashed #f1f5f9;">
                <div style="display: flex; justify-content: space-between; font-weight: 600; color: #1e293b; margin-bottom: 4px;">
                    <span>📁 ${displayName}</span>
                    <span style="color: #1e40af;">${h.months} tháng${hQuyDoi}</span>
                </div>
                ${subGroupsHtml}
            </div>`;
        }).join('');
        
        const totalYears = Math.floor(totalActualMonths / 12);
        const totalRem = totalActualMonths % 12;
        const totalQuyDoi = totalYears > 0 ? ` (${totalYears} năm ${totalRem} tháng)` : ` (${totalRem} tháng)`;
        
        const labelText = isGopHoSo ? 'TỔNG CỘNG THỜI GIAN (Gộp trùng lặp liên hồ sơ):' : 'TỔNG CỘNG THỜI GIAN (Cộng tổng các hồ sơ):';
        
        summaryTotal.innerHTML = `<div style="display: flex; justify-content: space-between; align-items: center;">
            <span>${labelText}</span>
            <span style="font-size: 1.25rem; color: #1d4ed8; text-decoration: underline;">${totalActualMonths} tháng ${totalQuyDoi}</span>
        </div>`;

        // Lắng nghe checkbox lọc trùng của các nhóm nhỏ
        const subGroupCheckboxes = summaryCard.querySelectorAll('.subgroup-loc-trung-checkbox');
        subGroupCheckboxes.forEach(cb => {
            cb.addEventListener('change', (e) => {
                const itemIndex = parseInt(e.target.dataset.index, 10);
                tableItems[itemIndex].isLocTrung = e.target.checked;
                renderTable();
            });
        });
    } else {
        summaryCard.style.display = 'none';
    }

    tableItems.forEach((item, index) => {
        const row = document.createElement('tr');
        const isHeader = !!item.isHeader;

        if (isHeader) {
            row.style.backgroundColor = '#f1f5f9';
            const monthsText = item.months !== undefined ? `${item.months} th` : '-';
            row.innerHTML = `
                <td class="text-center font-bold" style="font-weight: 700; color: #334155; border-bottom: 2px solid #cbd5e1;">${item.stt}</td>
                <td colspan="5" style="border-bottom: 2px solid #cbd5e1;">
                    <textarea class="table-input" data-field="project" rows="1" style="font-weight: 700; background-color: #f1f5f9; color: #0f172a; border: none; padding: 4px 8px;">${item.project || ''}</textarea>
                    <input type="hidden" class="table-input" data-field="content" value="">
                    <input type="hidden" class="table-input" data-field="position" value="">
                    <input type="hidden" class="table-input" data-field="startTime" value="">
                    <input type="hidden" class="table-input" data-field="endTime" value="">
                </td>
                <td class="text-center font-bold" style="border-bottom: 2px solid #cbd5e1; color: #64748b;">-</td>
                <td class="text-center col-months font-bold" style="border-bottom: 2px solid #cbd5e1; color: #1e40af;">${monthsText}</td>
            `;
        } else {
            let rowStyle = '';
            let monthsBadge = '';
            
            if (item.isOverlapped) {
                rowStyle = 'background-color: #fef2f2; border-left: 4px solid #ef4444;';
                monthsBadge = `<span style="color: #ef4444; font-weight: bold;">0 th</span>`;
            } else if (item.isPartialOverlapped) {
                rowStyle = 'background-color: #fffbeb; border-left: 4px solid #f59e0b;';
                monthsBadge = `<span style="color: #b45309; font-weight: bold;">${item.actualMonths} th</span>`;
            } else {
                monthsBadge = `<span style="color: #1d4ed8; font-weight: bold;">${item.actualMonths} th</span>`;
            }

            row.style.cssText = rowStyle;
            row.innerHTML = `
                <td class="text-center">${item.stt}</td>
                <td class="col-project"><textarea class="table-input" data-field="project" rows="1">${item.project || ''}</textarea></td>
                <td class="col-content"><textarea class="table-input" data-field="content" rows="1">${item.content || ''}</textarea></td>
                <td class="col-position"><input type="text" class="table-input" data-field="position" value="${item.position || ''}"></td>
                <td class="col-time-start"><input type="text" class="table-input" data-field="startTime" value="${item.startTime || ''}" placeholder="MM/YYYY" style="text-align: center;"></td>
                <td class="col-time-end"><input type="text" class="table-input" data-field="endTime" value="${item.endTime || ''}" placeholder="MM/YYYY" style="text-align: center;"></td>
                <td class="text-center" style="font-weight: 500; color: #475569;">${item.originalMonths || 0} th</td>
                <td class="text-center col-months">${monthsBadge}</td>
            `;
        }

        // Lắng nghe sự kiện thay đổi dữ liệu trên input
        const inputs = row.querySelectorAll('.table-input');
        inputs.forEach(input => {
            if (input.tagName.toLowerCase() === 'textarea') {
                setTimeout(() => autoResizeTextarea(input), 0);
                input.addEventListener('input', (e) => {
                    autoResizeTextarea(e.target);
                });
            }

            input.addEventListener('change', (e) => {
                const field = e.target.dataset.field;
                const value = e.target.value;
                
                // Cập nhật giá trị vào mảng gốc
                tableItems[index][field] = value;
                
                // Nếu sửa startTime hoặc endTime, ghép lại time
                if (field === 'startTime' || field === 'endTime') {
                    const st = tableItems[index].startTime || '';
                    const et = tableItems[index].endTime || '';
                    tableItems[index].time = st && et ? `${st}-${et}` : (st || et);
                }
                
                // Render lại bảng để tính toán lại trùng lặp và tổng số tháng
                renderTable();
            });
        });

        tableBody.appendChild(row);
    });
}

// Bắt sự kiện checkbox Gộp hồ sơ thay đổi
gopHoSoCheckbox.addEventListener('change', (e) => {
    isGopHoSo = e.target.checked;
    renderTable();
});

// Xuất file Excel
exportBtn.addEventListener('click', async () => {
    if (tableItems.length === 0) {
        showAlert('Bảng dữ liệu đang trống. Vui lòng import file trước khi xuất Excel.', 'error');
        return;
    }

    showLoading('Đang chuẩn bị file Excel tải xuống...');

    try {
        const cleanItems = tableItems.map(item => {
            const cleanItem = { ...item };
            delete cleanItem.subStatRef;
            delete cleanItem.subGroups;
            return cleanItem;
        });

        const response = await fetch('/api/export', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ items: cleanItems })
        });

        if (!response.ok) {
            const result = await response.json();
            throw new Error(result.error || 'Không thể xuất file Excel.');
        }

        const result = await response.json();
        const fileName = result.fileName;
        const fileUrl = result.url;

        // Vẫn tải file về máy thông qua link tĩnh
        const a = document.createElement('a');
        a.href = fileUrl;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        a.remove();


    } catch (err) {
        console.error(err);
        showAlert('Lỗi khi xuất Excel: ' + err.message, 'error');
    } finally {
        hideLoading();
    }
});
