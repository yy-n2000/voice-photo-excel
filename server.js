const express = require('express');
const ExcelJS = require('exceljs');
const cors    = require('cors');
const sizeOf  = require('image-size'); // 引入图片尺寸计算库

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

    // 列宽设置（Excel单位：字符宽度）
    worksheet.getColumn(1).width = 35; // 文字列
    worksheet.getColumn(2).width = 22; // 图片列1
    worksheet.getColumn(3).width = 22; // 图片列2
    worksheet.getColumn(4).width = 22; // 图片列3

    // 表头
    worksheet.getRow(1).values    = ['语音文字', '图片1', '图片2', '图片3'];
    worksheet.getRow(1).height    = 22;
    worksheet.getRow(1).font      = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
    worksheet.getRow(1).fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
    worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

    for (let i = 0; i < rows.length; i++) {
      const row      = rows[i];
      const rowIndex = i + 2;

      // 行高：90（Excel行高单位是磅，1磅≈1.33像素，90磅约120像素）
      worksheet.getRow(rowIndex).height = 90;

      // 文字单元格
      const cell     = worksheet.getCell(rowIndex, 1);
      cell.value     = row.text || '';
      cell.alignment = { vertical: 'middle', wrapText: true };
      cell.font      = { size: 11 };

      // 图片处理逻辑
      if (row.images && row.images.length > 0) {
        for (let j = 0; j < Math.min(row.images.length, 3); j++) {
          const base64Raw = row.images[j];
          const base64    = base64Raw.replace(/^data:image\/\w+;base64,/, '');
          const ext       = base64Raw.startsWith('data:image/png') ? 'png' : 'jpeg';
          const col       = j + 2;

          // 将图片注册到 workbook
          const imageId = workbook.addImage({ base64, extension: ext });

          // --- 核心修复：计算图片的等比例缩放大小 ---
          const imageBuffer = Buffer.from(base64, 'base64');
          const dimensions = sizeOf(imageBuffer); // 获取原图的宽高
          const imgWidth = dimensions.width;
          const imgHeight = dimensions.height;

          // 定义单元格内允许的图片最大像素（留出内边距）
          // 对应上方设置：列宽22(约154px)，行高90(约120px)
          const maxW = 140; 
          const maxH = 105;

          // 计算缩放比例 (取宽、高的最小比例值以确保图片完整放入)
          const ratio = Math.min(maxW / imgWidth, maxH / imgHeight);
          const finalWidth = Math.round(imgWidth * ratio);
          const finalHeight = Math.round(imgHeight * ratio);
          // ----------------------------------------

          // 插入图片：使用 tl 确定左上角起点，使用 ext 确定绝对大小
          worksheet.addImage(imageId, {
            tl: { col: col - 1 + 0.05, row: rowIndex - 1 + 0.1 }, // 起点略微偏移，不贴边
            ext: { width: finalWidth, height: finalHeight },      // 强制明确的像素大小
            editAs: 'oneCell'                                     // 图片随单元格移动，但不会被强制拉伸变形
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