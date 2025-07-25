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

    // --- Ghi sản phẩm & Xóa dòng thừa ---
    const startRow = 12;  // Dòng đầu tiên để ghi dữ liệu
    const maxRows = 16;   // Số dòng tối đa (linh kiện tối đa)
    const writeCount = Math.min(items.length, maxRows);

    // Ghi dữ liệu sản phẩm
    let writeRow = startRow, stt = 1;
    for (let i = 0; i < writeCount; i++) {
      const item = items[i];
      ws.getCell(`A${writeRow}`).value = stt;
      ws.getCell(`B${writeRow}`).value = item.name || '';
      ws.getCell(`C${writeRow}`).value = item.brand || '';
      ws.getCell(`D${writeRow}`).value = item.qty || '';
      ws.getCell(`E${writeRow}`).value = item.price || '';
      ws.getCell(`F${writeRow}`).value = item.total || '';
      ws.getCell(`G${writeRow}`).value = item.warranty || '';
      writeRow++; stt++;
    }

    // Xóa các dòng dư nếu có
    if (writeCount < maxRows) {
      // Chú ý: spliceRows xóa theo index, mỗi lần xóa dòng phía dưới sẽ dồn lên
      const firstRowToDelete = startRow + writeCount;
      const rowsToDelete = maxRows - writeCount;
      ws.spliceRows(firstRowToDelete, rowsToDelete);
    }

    const excelBuffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="bao_gia_sintech.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(excelBuffer));
  } catch (e) {
    res.status(500).json({ error: 'Export failed!', detail: e.message });
  }
}
