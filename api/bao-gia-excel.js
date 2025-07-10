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

    ws.getCell('C7').value = customer?.name || '';
    ws.getCell('C8').value = customer?.phone || '';
    ws.getCell('C9').value = customer?.address || '';

    let startRow = 12, stt = 1;
    for (const item of items) {
      ws.getCell(`A${startRow}`).value = stt;
      ws.getCell(`B${startRow}`).value = item.group;
      ws.getCell(`C${startRow}`).value = item.name;
      ws.getCell(`D${startRow}`).value = item.brand;
      ws.getCell(`E${startRow}`).value = item.qty;
      ws.getCell(`F${startRow}`).value = item.price;
      ws.getCell(`G${startRow}`).value = item.total;
      ws.getCell(`H${startRow}`).value = item.warranty;
      startRow++; stt++;
    }

    const excelBuffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="bao_gia_sintech.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(excelBuffer));
  } catch (e) {
    res.status(500).json({ error: 'Export failed!' });
  }
}
