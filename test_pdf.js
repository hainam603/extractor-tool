const fs = require('fs');
const pdf = require('pdf-parse');

const dataBuffer = fs.readFileSync('d:\\project\\New folder\\1.pdf');

pdf(dataBuffer).then(function(data) {
    console.log("PDF TEXT LEN:", data.text.trim().length);
    console.log("PDF TEXT PREVIEW:", data.text.trim().substring(0, 500));
}).catch(err => {
    console.error("PDF ERROR:", err);
});
