const express = require('express');
const Excel = require('exceljs');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json()); // <<< THÊM DÒNG NÀY để nhận JSON từ client

app.post('/bao-gia-excel', async (req, res) => {
  try {
    const items = req.body.items;

    // Nếu có truyền info khách hàng, bạn lấy luôn:
    const customer = req.body.customer || {}; // {name, phone, address}

    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile('form_bao_gia_cau_hinh_sintech.xlsx');

    const ws = workbook.getWorksheet('CH1');
    if (!ws) return res.status(500).send('Không tìm thấy worksheet!');

    // --- GHI THÔNG TIN KHÁCH HÀNG (nếu có) ---
    ws.getCell('B7').value = customer.name || '';
    ws.getCell('B8').value = customer.phone || '';
    ws.getCell('B9').value = customer.address || '';

    // --- GHI DANH SÁCH SẢN PHẨM ---
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

    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="bao_gia_sintech.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);

  } catch (err) {
    // console.error(err);
    res.status(500).send('Có lỗi xảy ra khi xuất báo giá, vui lòng thử lại sau!');
  }
});

app.listen(3000, () => console.log('Server started!'));
