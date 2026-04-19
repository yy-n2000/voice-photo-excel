const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// 中间件
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

// 上传图片临时存储
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// ============================
// 接口：生成Excel文件
// ============================
app.post('/generate-excel', async (req, res) => {
  try {
    const { rows } = req.body;
    // rows 格式: [ { text: "语音文字", images: ["base64...", "base64..."] }, ... ]

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('记录');

    // 设置表头
    worksheet.columns = [
      { header: '语音文字', key: 'text', width: 40 },
      { header: '图片1', key: 'img1', width: 20 },
      { header: '图片2', key: 'img2', width: 20 },
      { header: '图片3', key: 'img3', width: 20 },
    ];

    // 设置表头样式
    worksheet.getRow(1).font = { bold: true, size: 12 };
    worksheet.getRow(1).fill = {
      type: 'pattern', pattern: 'solid',
      fgColor: { argb: 'FF4472C4' }
    };
    worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };

    // 填充数据行
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowIndex = i + 2; // 从第2行开始（第1行是表头）
      const ROW_HEIGHT = 80; // 行高（像素）

      // 设置行高
      worksheet.getRow(rowIndex).height = ROW_HEIGHT;

      // 填入文字（第1列）
      const textCell = worksheet.getCell(rowIndex, 1);
      textCell.value = row.text || '';
      textCell.alignment = { vertical: 'middle', wrapText: true };

      // 填入图片（第2列开始，每张图片占一列）
      if (row.images && row.images.length > 0) {
        for (let j = 0; j < row.images.length && j < 3; j++) {
          const base64Data = row.images[j];
          // 去掉 data:image/xxx;base64, 前缀
          const base64Clean = base64Data.replace(/^data:image\/\w+;base64,/, '');
          const extension = base64Data.includes('data:image/png') ? 'png' : 'jpeg';

          const imageId = workbook.addImage({
            base64: base64Clean,
            extension: extension,
          });

          const col = j + 2; // 从第2列开始
          const colLetter = getColumnLetter(col);

          // 设置列宽
          worksheet.getColumn(col).width = 18;

          // 嵌入图片到单元格
          worksheet.addImage(imageId, {
            tl: { col: col - 1 + 0.1, row: rowIndex - 1 + 0.1 },
            br: { col: col - 1 + 0.9, row: rowIndex - 1 + 0.9 },
          });
        }
      }
    }

    // 生成文件并返回
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="voice_photo_record.xlsx"');
    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// 辅助函数：数字列号转字母（1→A, 2→B ...）
function getColumnLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

app.listen(PORT, () => {
  console.log(`服务器运行在端口 ${PORT}`);
});