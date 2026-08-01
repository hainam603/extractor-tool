const fs = require('fs');

async function testUpload() {
    const filePath = 'd:\\project\\New folder\\Bản khai QTCT-Thiệp (Đạt Phương).docx';
    const fileBuffer = fs.readFileSync(filePath);
    const blob = new Blob([fileBuffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    
    const formData = new FormData();
    formData.append('file', blob, 'Bản khai QTCT-Thiệp (Đạt Phương).docx');

    try {
        const response = await fetch('http://localhost:3000/api/upload', {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        console.log("API UPLOAD TEST:");
        console.log("Status Code:", response.status);
        if (result.data) {
            console.log("Successfully extracted rows count:", result.data.length);
            console.log("First 3 rows:", JSON.stringify(result.data.slice(0, 3), null, 2));
        } else {
            console.log("Result:", JSON.stringify(result, null, 2));
        }
    } catch (err) {
        console.error("Fetch Error:", err);
    }
}

testUpload();
