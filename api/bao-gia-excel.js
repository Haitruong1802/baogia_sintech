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

    // Gán thông tin khách hàng
    ws.getCell('C7').value = customer?.name || '';
    ws.getCell('C8').value = customer?.phone || '';
    ws.getCell('C9').value = customer?.address || '';

    // Thông số vùng sản phẩm (dòng bắt đầu và số dòng tối đa)
    const startRow = 12; // dòng bắt đầu ghi sản phẩm
    const maxRows = 16;  // tối đa số dòng sản phẩm cho phép (template)
    const n = items.length;

    // Ghi dữ liệu sản phẩm vào sheet
    for (let i = 0; i < n; i++) {
      const row = ws.getRow(startRow + i);
      row.getCell(1).value = i + 1;
      row.getCell(2).value = items[i].name || '';
      row.getCell(3).value = items[i].brand || '';
      row.getCell(4).value = items[i].qty || '';
      row.getCell(5).value = items[i].price || '';
      // row.getCell(6).value = items[i].total || '';
      row.getCell(7).value = items[i].warranty || '';
    }

    // Xóa dòng dư nếu số sản phẩm < maxRows
    if (n < maxRows) {
      // Chỉ xóa row vùng sản phẩm, không ảnh hưởng vùng phía dưới
      ws.spliceRows(startRow + n, maxRows - n);
    }

    // Xuất file
    const excelBuffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="bao_gia_sintech.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(excelBuffer));
  } catch (e) {
    res.status(500).json({ error: 'Export failed!', detail: e.message });
  }
}
