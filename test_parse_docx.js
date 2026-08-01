const fs = require('fs');
const AdmZip = require('adm-zip');
const { XMLParser } = require('fast-xml-parser');

const buffer = fs.readFileSync('d:\\project\\New folder\\Bản khai QTCT-Thiệp (Đạt Phương).docx');
const zip = new AdmZip(buffer);
const docXmlEntry = zip.getEntry('word/document.xml');
const xmlText = docXmlEntry.getData().toString('utf8');

const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    parseAttributeValue: false
});
const jsonObj = parser.parse(xmlText);
const body = jsonObj['w:document']?.['w:body'];

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

console.log("Total tables found:", tables.length);

tables.forEach((table, tIdx) => {
    console.log(`\n--- TABLE ${tIdx + 1} ---`);
    const rows = table['w:tr'];
    if (!rows) return;
    const rowsArr = Array.isArray(rows) ? rows : [rows];
    rowsArr.forEach((row, rIdx) => {
        const cells = row['w:tc'];
        if (!cells) return;
        const cellsArr = Array.isArray(cells) ? cells : [cells];
        const rowText = cellsArr.map(cell => {
            const tList = [];
            function findTexts(obj) {
                if (!obj || typeof obj !== 'object') return;
                if (obj['w:t'] !== undefined) {
                    const val = obj['w:t'];
                    if (typeof val === 'string') {
                        tList.push(val);
                    } else if (val && val['#text']) {
                        tList.push(val['#text']);
                    }
                }
                for (const key in obj) {
                    findTexts(obj[key]);
                }
            }
            findTexts(cell);
            return tList.join('').trim();
        });
        console.log(`Row ${rIdx + 1} (cells: ${rowText.length}):`, JSON.stringify(rowText));
    });
});
