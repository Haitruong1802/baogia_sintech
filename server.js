const Excel = require('exceljs');
const fs = require('fs');
const path = require('path');

// Vercel API route xuất excel cho báo giá
export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method Not Allowed' });
    return;
  }
  try {
    const { items = [], customer = {} } = req.body;

    // --- Đọc file mẫu dưới dạng buffer (phải nằm trong repo, cùng cấp file này hoặc thư mục con) ---
    const filePath = path.join(process.cwd(), 'form_bao_gia_cau_hinh_sintech.xlsx');
    const buffer = fs.readFileSync(filePath); // Đảm bảo file này push lên github repo luôn!

    const workbook = new Excel.Workbook();
    await workbook.xlsx.load(buffer);

    const ws = workbook.getWorksheet('CH1');
    if (!ws) {
      res.status(500).json({ error: 'Không tìm thấy worksheet!' });
      return;
    }

    // Ghi thông tin khách hàng
    ws.getCell('B7').value = customer.name || '';
    ws.getCell('B8').value = customer.phone || '';
    ws.getCell('B9').value = customer.address || '';

    // Ghi danh sách sản phẩm từ dòng 12
    let startRow = 12;
    let stt = 1;
    items.forEach(item => {
      ws.getCell(`A${startRow}`).value = stt;
      ws.getCell(`B${startRow}`).value = item.group;
      ws.getCell(`C${startRow}`).value = item.name;
      ws.getCell(`D${startRow}`).value = item.brand;
      ws.getCell(`E${startRow}`).value = item.qty;
      ws.getCell(`F${startRow}`).value = item.price;
      ws.getCell(`G${startRow}`).value = item.total;
      ws.getCell(`H${startRow}`).value = item.warranty;
      startRow++;
      stt++;
    });

    // Xuất buffer trả về cho client
    const outBuffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="bao_gia_sintech.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(outBuffer);

  } catch (err) {
    // console.error(err);
    res.status(500).send('Có lỗi xảy ra khi xuất báo giá, vui lòng thử lại sau!');
  }
}
