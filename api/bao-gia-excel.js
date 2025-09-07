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

    const wb = new Excel.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.getWorksheet('CH1');
    if (!ws) return res.status(500).json({ error: 'Không tìm thấy worksheet!' });

    // Thông tin KH
    ws.getCell('C7').value = customer.name || '';
    ws.getCell('C8').value = customer.phone || '';
    ws.getCell('C9').value = customer.address || '';

    const DATA_START = 12;
    const n = items.length;

    // Nhân bản dòng mẫu để giữ border/format/công thức
    if (n > 1) ws.duplicateRow(DATA_START, n - 1, true);

    // (tuỳ chọn) hàm copy style từ dòng mẫu phòng khi duplicateRow không giữ đủ border
    const templateRow = ws.getRow(DATA_START);
    const copyStyleFromTemplate = (row) => {
      row.eachCell({ includeEmpty: true }, (cell, c) => {
        const s = templateRow.getCell(c).style || {};
        // deep clone để tránh tham chiếu chung
        cell.style = JSON.parse(JSON.stringify(s));
      });
    };

    let r = DATA_START, stt = 1;
    for (const it of items) {
      const qty = toNumber(it.qty);
      const price = toNumber(it.price);

      const row = ws.getRow(r);
      // Nếu thấy mất kẻ, bật dòng dưới:
      // copyStyleFromTemplate(row);

      row.getCell(1).value = stt;               // A: STT
      row.getCell(2).value = it.name || '';     // B: Tên SP
      row.getCell(3).value = qty;               // C: SL (số)
      row.getCell(4).value = price;             // D: Đơn giá (số)
      // E: GIỮ công thức từ dòng mẫu (=C*D). KHÔNG gán lại bằng code.
      row.getCell(6).value = it.warranty || ''; // F: Bảo hành
      row.commit();

      r++; stt++;
    }

    // Cập nhật ô TỔNG (nếu có nhãn "TỔNG:" ở cột D)
    let totalRow = null;
    for (let i = r; i <= r + 50; i++) {
      const v = (ws.getCell(`D${i}`).value || '').toString().trim().toUpperCase();
      if (v.includes('TỔNG')) { totalRow = i; break; }
    }
    if (totalRow && n > 0) {
      ws.getCell(`E${totalRow}`).value = { formula: `SUM(E${DATA_START}:E${r - 1})`, result: 0 };
    }

    const buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="bao_gia_sintech.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Export failed!' });
  }
}
