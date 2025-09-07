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
    const { items = [], customer = {} } = req.body;

    const filePath = path.join(process.cwd(), 'form_bao_gia_cau_hinh_sintech.xlsx');
    const buffer = fs.readFileSync(filePath);

    const workbook = new Excel.Workbook();
    await workbook.xlsx.load(buffer);
    const ws = workbook.getWorksheet('CH1');
    if (!ws) return res.status(500).json({ error: 'Không tìm thấy worksheet!' });

    // Fill customer info
    ws.getCell('B7').value = customer.name || '';
    ws.getCell('B8').value = customer.phone || '';
    ws.getCell('B9').value = customer.address || '';

    // ===== Vùng dữ liệu =====
    const DATA_START = 12;                 // dòng mẫu có sẵn trong file
    const count = Math.max(items.length, 0);

    // Nếu có nhiều hơn 1 item, nhân bản dòng mẫu (giữ style) và CHÈN xuống dưới
    // Ví dụ có 5 item → cần thêm 4 dòng nữa (vì dòng 12 đã là mẫu)
    if (count > 1) {
      ws.duplicateRow(DATA_START, count - 1, true); // true = insert
    }

    // Ghi dữ liệu
    let rowIndex = DATA_START;
    let stt = 1;
    for (const it of items) {
      const row = ws.getRow(rowIndex);
      row.getCell(1).value = stt;                          // A - STT
      row.getCell(2).value = it.name || '';                // B - Tên SP
      row.getCell(3).value = Number(it.qty) || 0;          // C - SL
      row.getCell(4).value = Number(it.price) || 0;        // D - Đơn giá
      // E - Thành tiền: để công thức trong file (C*D) hoặc đặt lại tại đây:
      row.getCell(5).value = { formula: `C${rowIndex}*D${rowIndex}` };
      row.getCell(6).value = it.warranty || '';            // F - Bảo hành
      row.commit();
      rowIndex++; stt++;
    }

    // ===== Cập nhật lại ô TỔNG (nếu cần) =====
    // Tìm dòng có chữ "TỔNG:" ở cột D để đặt lại công thức tổng cột E
    let totalRow = null;
    for (let r = rowIndex; r <= rowIndex + 50; r++) {
      const v = (ws.getCell(`D${r}`).value || '').toString().trim();
      if (v && v.toUpperCase().includes('TỔNG')) { totalRow = r; break; }
    }
    if (totalRow && items.length > 0) {
      ws.getCell(`E${totalRow}`).value = { formula: `SUM(E${DATA_START}:E${rowIndex - 1})` };
    }

    const excelBuffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="bao_gia_sintech.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(excelBuffer));
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Export failed!' });
  }
}
