const express = require('express');
const ExcelJS = require('exceljs');
const cors    = require('cors');
const sizeOf  = require('image-size');

const app  = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

// ===== 单位换算 =====
const colWidthToPx = (width) => width * 7;
const rowHeightToPx = (height) => height * 1.33;

// 留白
const PADDING = 8;

app.post('/generate-excel', async (req, res) => {
  try {
    const { rows } = req.body;

    if (!rows || !Array.isArray(rows)) {
      return res.status(400).json({ error: 'rows 数据格式错误' });
    }

    const workbook  = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('记录');

    // ===== 列宽 =====
    worksheet.getColumn(1).width = 35;
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
      excelRow.height = 90;

      // ===== 文本 =====
      const textCell = worksheet.getCell(rowIndex, 1);
      textCell.value = rowData.text || '';
      textCell.alignment = { vertical: 'middle', wrapText: true };
      textCell.font = { size: 11 };

      // ===== 图片处理 =====
      if (Array.isArray(rowData.images)) {
        for (let j = 0; j < Math.min(rowData.images.length, 3); j++) {
          try {
            const base64Raw = rowData.images[j];
            if (!base64Raw || typeof base64Raw !== 'string') continue;

            // ===== 校验base64 =====
            const match = base64Raw.match(/^data:image\/(\w+);base64,/);
            if (!match) {
              console.warn(`第${rowIndex}行图片格式非法`);
              continue;
            }

            const ext = match[1] === 'png' ? 'png' : 'jpeg';
            const base64 = base64Raw.replace(/^data:image\/\w+;base64,/, '');

            const col = j + 2;

            const imageId = workbook.addImage({
              base64,
              extension: ext
            });

            // ===== 获取图片尺寸 =====
            let imgW = 100;
            let imgH = 100;

            try {
              const buffer = Buffer.from(base64, 'base64');
              const size   = sizeOf(buffer);
              imgW = size.width;
              imgH = size.height;
            } catch (e) {
              console.warn(`第${rowIndex}行图片尺寸解析失败`);
            }

            // ===== 单元格尺寸 =====
            const colWidth  = worksheet.getColumn(col).width || 20;
            const rowHeight = excelRow.height || 80;

            const cellW = colWidthToPx(colWidth);
            const cellH = rowHeightToPx(rowHeight);

            // ===== 计算缩放 =====
            const maxW = cellW - PADDING * 2;
            const maxH = cellH - PADDING * 2;

            const ratio = Math.min(maxW / imgW, maxH / imgH, 1);

            const finalW = imgW * ratio;
            const finalH = imgH * ratio;

            // ===== 居中 =====
            const offsetX = (cellW - finalW) / 2;
            const offsetY = (cellH - finalH) / 2;

            worksheet.addImage(imageId, {
              tl: {
                col: col - 1 + offsetX / cellW,
                row: rowIndex - 1 + offsetY / cellH
              },
              ext: {
                width: finalW,
                height: finalH
              },
              editAs: 'oneCell'
            });

          } catch (imgErr) {
            // 单张图片失败不影响整体
            console.error(`第${rowIndex}行图片插入失败`, imgErr);
          }
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
    console.error('导出失败：', err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`服务器运行在端口 ${PORT}`);
});