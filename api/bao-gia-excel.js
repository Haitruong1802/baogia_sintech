import fs from 'fs';
import path from 'path';
import Excel from 'exceljs';

const toNumber = (v) => {
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  if (v == null) return 0;
  const s = String(v).replace(/[^\d.-]/g, '');
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : 0;
};

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

    const wb = new Excel.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.getWorksheet('CH1');
    if (!ws) return res.status(500).json({ error: 'Không tìm thấy worksheet!' });

    // Điền thông tin KH
    ws.getCell('B7').value = customer.name || '';
    ws.getCell('B8').value = customer.phone || '';
    ws.getCell('B9').value = customer.address || '';

    const DATA_START = 12;

    // Nhân bản dòng mẫu để giữ style/border/công thức
    if (items.length > 1) {
      ws.duplicateRow(DATA_START, items.length - 1, true); // insert
    }

    // Ghi dữ liệu (KHÔNG đụng cột E để giữ công thức từ dòng mẫu)
    let rowIdx = DATA_START;
    let stt = 1;
    for (const it of items) {
      const qty = toNumber(it.qty);
      const price = toNumber(it.price);

      const row = ws.getRow(rowIdx);
      row.getCell(1).value = stt;                  // A: STT
      row.getCell(2).value = it.name || '';        // B: Tên SP
      row.getCell(3).value = qty;                  // C: SL (số)
      row.getCell(4).value = price;                // D: Đơn giá (số)
      row.getCell(5).value = { formula: `C${rowIndex}*D${rowIndex}` };
      row.getCell(6).value = it.warranty || '';    // F: Bảo hành
      row.commit();

      rowIdx++; stt++;
    }

    // Cập nhật ô TỔNG (nếu tổng ở cột E, nhãn "TỔNG:" ở cột D)
    // Tìm dòng nhãn "TỔNG:" phía dưới
    let totalRow = null;
    for (let r = rowIdx; r <= rowIdx + 50; r++) {
      const v = (ws.getCell(`D${r}`).value || '').toString().trim().toUpperCase();
      if (v.includes('TỔNG')) { totalRow = r; break; }
    }
    if (totalRow && items.length > 0) {
      ws.getCell(`E${totalRow}`).value = { formula: `SUM(E${DATA_START}:E${rowIdx - 1})`, result: 0 };
    }

    // Tên file theo yyyyMMdd_HHmmss
    const now = new Date();
    const pad = (n) => n.toString().padStart(2, '0');
    const dateStr = `${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;

    const excelBuffer = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', `attachment; filename="bao_gia_sintech_${dateStr}.xlsx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(excelBuffer));
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Export failed!' });
  }
}
