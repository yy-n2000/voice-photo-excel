const express = require('express');
const ExcelJS = require('exceljs');
const cors    = require('cors');
const path    = require('path');

const app  = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

app.post('/generate-excel', async (req, res) => {
  try {
    const { rows } = req.body;

    const workbook  = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('记录');

    // 表头
    worksheet.getRow(1).values = ['语音文字', '图片1', '图片2', '图片3'];
    worksheet.getRow(1).font   = { bold: true, color: { argb: 'FFFFFFFF' } };
    worksheet.getRow(1).fill   = {
      type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' }
    };
    worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getColumn(1).width  = 40;

    for (let i = 0; i < rows.length; i++) {
      const row      = rows[i];
      const rowIndex = i + 2;
      const ROW_PX   = 80;

      worksheet.getRow(rowIndex).height = ROW_PX;

      // 文字
      const cell = worksheet.getCell(rowIndex, 1);
      cell.value     = row.text || '';
      cell.alignment = { vertical: 'middle', wrapText: true };

      // 图片
      if (row.images && row.images.length > 0) {
        for (let j = 0; j < Math.min(row.images.length, 3); j++) {
          const base64Raw = row.images[j];
          const base64    = base64Raw.replace(/^data:image\/\w+;base64,/, '');
          const ext       = base64Raw.includes('data:image/png') ? 'png' : 'jpeg';
          const col       = j + 2;

          worksheet.getColumn(col).width = 18;

          const imageId = workbook.addImage({ base64, extension: ext });
          worksheet.addImage(imageId, {
            tl: { col: col - 1 + 0.05, row: rowIndex - 1 + 0.05 },
            br: { col: col - 1 + 0.95, row: rowIndex - 1 + 0.95 },
          });
        }
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',
      'attachment; filename="voice_photo_record.xlsx"');
    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

function getColumnLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col    = Math.floor((col - 1) / 26);
  }
  return letter;
}

app.listen(PORT, () => console.log(`服务器运行在端口 ${PORT}`));