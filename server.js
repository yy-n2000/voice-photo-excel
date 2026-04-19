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

    // 列宽设置（适当加宽一点，方便放图）
    worksheet.getColumn(1).width = 35; // 文字列
    worksheet.getColumn(2).width = 25;
    worksheet.getColumn(3).width = 25;
    worksheet.getColumn(4).width = 25;

    // 表头
    const headerRow = worksheet.getRow(1);
    headerRow.values = ['语音文字', '图片1', '图片2', '图片3'];
    headerRow.height = 22;
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4472C4' }
    };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

    // 图片尺寸（核心参数）
    const IMG_WIDTH  = 150; // 像素
    const IMG_HEIGHT = 100; // 像素

    for (let i = 0; i < rows.length; i++) {
      const row      = rows[i];
      const rowIndex = i + 2;

      // 行高（稍微大于图片高度，留空间）
      worksheet.getRow(rowIndex).height = 110;

      // ===== 文字 =====
      const cell = worksheet.getCell(rowIndex, 1);
      cell.value = row.text || '';
      cell.alignment = { vertical: 'middle', wrapText: true };
      cell.font = { size: 11 };

      // ===== 图片 =====
      if (row.images && row.images.length > 0) {
        for (let j = 0; j < Math.min(row.images.length, 3); j++) {
          const base64Raw = row.images[j];
          const base64    = base64Raw.replace(/^data:image\/\w+;base64,/, '');
          const ext       = base64Raw.startsWith('data:image/png') ? 'png' : 'jpeg';

          const col = j + 2;

          const imageId = workbook.addImage({
            base64,
            extension: ext
          });

          // ✅ 使用像素控制 + 留边距
          worksheet.addImage(imageId, {
            tl: {
              col: col - 1 + 0.1,   // 左边距
              row: rowIndex - 1 + 0.15 // 上边距
            },
            ext: {
              width: IMG_WIDTH,
              height: IMG_HEIGHT
            }
          });
        }
      }
    }

    // 输出
    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename="record.xlsx"'
    );

    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`服务器运行在端口 ${PORT}`);
});