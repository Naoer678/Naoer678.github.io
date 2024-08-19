const express = require('express');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();

app.use(express.urlencoded({ extended: true }));

app.post('/submit', (req, res) => {
    const { name, email, age, 'frist-survey': firstSurvey, referrer, refer, comment } = req.body;

    // 处理 refer，如果不是数组则将其转换为数组
    let referArray = Array.isArray(refer)? refer : [refer];

    const data = [
        [name, email, age, firstSurvey, referrer, referArray.join(','), comment]
    ];

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    XLSX.writeFile(workbook, 'formData.xlsx');

    res.send('表单数据已成功写入 Excel 文件！');
});

app.listen(3000, () => {
    console.log('服务器运行在 3000 端口');
});