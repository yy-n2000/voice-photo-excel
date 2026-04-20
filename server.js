const express = require('express');
const ExcelJS = require('exceljs');
const cors    = require('cors');

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

    worksheet.getColumn(1).width = 35;
    worksheet.getColumn(2).width = 22;
    worksheet.getColumn(3).width = 22;
    worksheet.getColumn(4).width = 22;

    worksheet.getRow(1).values    = ['语音文字', '图片1', '图片2', '图片3'];
    worksheet.getRow(1).height    = 22;
    worksheet.getRow(1).font      = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
    worksheet.getRow(1).fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
    worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

    for (let i = 0; i < rows.length; i++) {
      const row      = rows[i];
      const rowIndex = i + 2;

      worksheet.getRow(rowIndex).height = 90;

      const cell     = worksheet.getCell(rowIndex, 1);
      cell.value     = row.text || '';
      cell.alignment = { vertical: 'middle', wrapText: true };
      cell.font      = { size: 11 };

      if (row.images && row.images.length > 0) {
        for (let j = 0; j < Math.min(row.images.length, 3); j++) {
          const base64Raw = row.images[j];
          const base64    = base64Raw.replace(/^data:image\/\w+;base64,/, '');
          const ext       = base64Raw.startsWith('data:image/png') ? 'png' : 'jpeg';
          const col       = j + 2;

          const imageId = workbook.addImage({ base64, extension: ext });
          worksheet.addImage(imageId, {
            tl: { col: col - 1 + 0.08, row: rowIndex - 1 + 0.08 },
            br: { col: col - 1 + 0.92, row: rowIndex - 1 + 0.92 },
            editAs: 'oneCell'
          });
        }
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="record.xlsx"');
    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`服务器运行在端口 ${PORT}`));