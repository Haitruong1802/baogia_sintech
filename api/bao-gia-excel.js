import fs from 'fs';
import path from 'path';
import Excel from 'exceljs';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method Not Allowed' });

  try {
    const { items, customer } = req.body;

    const filePath = path.join(process.cwd(), 'form_bao_gia_cau_hinh_sintech.xlsx');
    const buffer = fs.readFileSync(filePath);

    const workbook = new Excel.Workbook();
    await workbook.xlsx.load(buffer);
    const ws = workbook.getWorksheet('CH1');
    if (!ws) return res.status(500).json({ error: 'Không tìm thấy worksheet!' });

    // Update thông tin khách hàng
    ws.getCell('C7').value = customer?.name || '';
    ws.getCell('C8').value = customer?.phone || '';
    ws.getCell('C9').value = customer?.address || '';

    // Xác định số dòng trống có sẵn trong template
    let startRow = 12; // Dòng đầu tiên để ghi dữ liệu
    const templateRows = 15; // Sửa đúng bằng số dòng trống có sẵn dưới template
    const neededRows = items.length;

    // Nếu thiếu dòng, phát sinh thêm dòng mới và copy style từ dòng cuối template
    function copyRowStyle(ws, fromRowNum, toRowNum) {
      const fromRow = ws.getRow(fromRowNum);
      const toRow = ws.getRow(toRowNum);
      fromRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        toRow.getCell(colNumber).style = { ...cell.style };
        // Copy border/align nếu cần:
        toRow.getCell(colNumber).border = cell.border ? { ...cell.border } : undefined;
        toRow.getCell(colNumber).alignment = cell.alignment ? { ...cell.alignment } : undefined;
        toRow.getCell(colNumber).font = cell.font ? { ...cell.font } : undefined;
        toRow.getCell(colNumber).numFmt = cell.numFmt || undefined;
        toRow.getCell(colNumber).fill = cell.fill || undefined;
      });
    }

    if (neededRows > templateRows) {
      const lastTemplateRow = startRow + templateRows - 1;
      for (let i = 0; i < neededRows - templateRows; i++) {
        ws.insertRow(lastTemplateRow + 1 + i, []);
        copyRowStyle(ws, lastTemplateRow, lastTemplateRow + 1 + i);
      }
    }

    // Ghi dữ liệu vào sheet
    let writeRow = startRow, stt = 1;
    for (const item of items) {
      ws.getCell(`A${writeRow}`).value = stt;
      ws.getCell(`B${writeRow}`).value = item.name || '';
      ws.getCell(`C${writeRow}`).value = item.brand || '';
      ws.getCell(`D${writeRow}`).value = item.qty || '';
      ws.getCell(`E${writeRow}`).value = item.price || '';
      ws.getCell(`F${writeRow}`).value = item.total || '';
      ws.getCell(`G${writeRow}`).value = item.warranty || '';
      writeRow++; stt++;
    }

    const excelBuffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="bao_gia_sintech.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(excelBuffer));
  } catch (e) {
    res.status(500).json({ error: 'Export failed!', detail: e.message });
  }
}
