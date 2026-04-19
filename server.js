const express = require('express');
const ExcelJS = require('exceljs');
const cors    = require('cors');
const sizeOf  = require('image-size');

const app  = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

// ===== Excel单位换算 =====
const colWidthToPx = (width) => width * 7;       // 列宽 → 像素（近似）
const rowHeightToPx = (height) => height * 1.33; // 行高 → 像素

// 留白（像素）
const PADDING = 8;

app.post('/generate-excel', async (req, res) => {
  try {
    const { rows } = req.body;

    const workbook  = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('记录');

    // ===== 列宽 =====
    worksheet.getColumn(1).width = 35; // 文本
    worksheet.getColumn(2).width = 22;
    worksheet.getColumn(3).width = 22;
    worksheet.getColumn(4).width = 22;

    // ===== 表头 =====
    const header = worksheet.getRow(1);
    header.values    = ['语音文字', '图片1', '图片2', '图片3'];
    header.height    = 22;
    header.font      = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
    header.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
    header.alignment = { vertical: 'middle', horizontal: 'center' };

    // ===== 数据 =====
    for (let i = 0; i < rows.length; i++) {
      const rowData  = rows[i];
      const rowIndex = i + 2;

      const excelRow = worksheet.getRow(rowIndex);
      excelRow.height = 90; // ≈120px

      // ===== 文本 =====
      const textCell = worksheet.getCell(rowIndex, 1);
      textCell.value = rowData.text || '';
      textCell.alignment = { vertical: 'middle', wrapText: true };
      textCell.font = { size: 11 };

      // ===== 图片 =====
      if (rowData.images && rowData.images.length > 0) {
        for (let j = 0; j < Math.min(rowData.images.length, 3); j++) {
          const base64Raw = rowData.images[j];
          const base64    = base64Raw.replace(/^data:image\/\w+;base64,/, '');
          const ext       = base64Raw.startsWith('data:image/png') ? 'png' : 'jpeg';
          const col       = j + 2;

          const imageId = workbook.addImage({
            base64,
            extension: ext
          });

          // ===== 获取图片原始尺寸 =====
          const buffer = Buffer.from(base64, 'base64');
          const { width: imgW, height: imgH } = sizeOf(buffer);

          // ===== 单元格尺寸 =====
          const colWidth  = worksheet.getColumn(col).width || 20;
          const rowHeight = excelRow.height || 80;

          const cellW = colWidthToPx(colWidth);
          const cellH = rowHeightToPx(rowHeight);

          // ===== 可用区域（扣掉留白）=====
          const maxW = cellW - PADDING * 2;
          const maxH = cellH - PADDING * 2;

          // ===== 等比例缩放 =====
          const ratio = Math.min(maxW / imgW, maxH / imgH, 1); // 不放大

          const finalW = imgW * ratio;
          const finalH = imgH * ratio;

          // ===== 居中偏移 =====
          const offsetX = (cellW - finalW) / 2;
          const offsetY = (cellH - finalH) / 2;

          // ===== 转换为Excel坐标比例 =====
          const colOffset = offsetX / cellW;
          const rowOffset = offsetY / cellH;

          // ===== 插入图片 =====
          worksheet.addImage(imageId, {
            tl: {
              col: col - 1 + colOffset,
              row: rowIndex - 1 + rowOffset
            },
            ext: {
              width: finalW,
              height: finalH
            },
            editAs: 'oneCell'
          });
        }
      }
    }

    // ===== 导出 =====
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