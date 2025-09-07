import fs from 'fs';
import path from 'path';
import Excel from 'exceljs';

const toNumber = (v) => {
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  if (v == null) return 0;
  const s = String(v).replace(/[^\d.-]/g, ''); // bỏ , . khoảng trắng, ₫
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : 0;
};

const sanitize = (s) =>
  String(s ?? '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');

const ymdHMS = () => {
  const now = new Date();
  const pad = (n) => n.toString().padStart(2, '0');
  return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
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

    // Thông tin khách hàng
    ws.getCell('B7').value = sanitize(customer.name);
    ws.getCell('B8').value = sanitize(customer.phone);
    ws.getCell('B9').value = sanitize(customer.address);

    const DATA_START = 12;
    const n = items.length;

    // Nhân bản dòng mẫu 12 để giữ nguyên border/numFmt/công thức
    if (n > 1) ws.duplicateRow(DATA_START, n - 1, true);

    // Kiểm tra dòng mẫu E12 đã có công thức (=C*D) chưa
    const tplVal = ws.getCell(`E${DATA_START}`).value;
    const templateHasFormula = tplVal && typeof tplVal === 'object' && 'formula' in tplVal;

    // Ghi dữ liệu từng dòng
    let r = DATA_START;
    let stt = 1;
    for (const it of items) {
      const qty = toNumber(it.qty);
      const price = toNumber(it.price);

      const row = ws.getRow(r);
      row.getCell(1).value = stt;                  // A: STT
      row.getCell(2).value = sanitize(it.name);    // B: Tên SP
      row.getCell(3).value = qty;                  // C: SL (số)
      row.getCell(4).value = price;                // D: Đơn giá (số)

      // E: Giữ công thức từ dòng mẫu; nếu mẫu KHÔNG có công thức thì set kèm result
      if (!templateHasFormula) {
        row.getCell(5).value = { formula: `C${r}*D${r}`, result: qty * price };
      }

      row.getCell(6).value = sanitize(it.warranty); // F: Bảo hành
      row.commit();

      r++; stt++;
    }

    // Tìm dòng "TỔNG:" ở cột D và đặt công thức tổng cột E
    let totalRow = null;
    for (let i = r; i <= r + 60; i++) {
      const v = (ws.getCell(`D${i}`).value || '').toString().trim().toUpperCase();
      if (v.includes('TỔNG')) { totalRow = i; break; }
    }
    if (totalRow && n > 0) {
      ws.getCell(`E${totalRow}`).value = { formula: `SUM(E${DATA_START}:E${r - 1})`, result: 0 };
    }

    // (Tuỳ chọn) chốt viền dưới cho dòng sản phẩm cuối nếu cần
    // const last = ws.getRow(r - 1);
    // for (let c = 1; c <= 6; c++) {
    //   const prevBorder = last.getCell(c).border || {};
    //   last.getCell(c).border = { ...prevBorder, bottom: { style: 'thin', color: { argb: 'FF000000' } } };
    // }
    // last.commit();

    const buf = await wb.xlsx.writeBuffer();

    // ===== Sau khi ghi xong các dòng sản phẩm =====

    // last data row của bảng
    const lastDataRow = items.length ? (DATA_START + items.length - 1) : (DATA_START - 1);

    // Helper: tìm hàng theo nhãn ở cột D
    const findRowByLabel = (regex) => {
      const maxR = ws.rowCount;
      for (let r = DATA_START; r <= maxR; r++) {
        const txt = (ws.getCell(`D${r}`).value || '').toString().trim();
        if (regex.test(txt)) return r;
      }
      return null;
    };

    // Xác định các dòng cần tính
    const rTong       = findRowByLabel(/^TỔNG\b/i);
    const rVat        = findRowByLabel(/THUẾ\s*GTGT/i);
    const rGiamGia    = findRowByLabel(/GIẢM\s*GIÁ/i);
    const rThanhTien  = findRowByLabel(/THÀNH\s*TIỀN/i);

    // TỔNG = SUM(E từ dòng dữ liệu đầu đến cuối)
    if (rTong && items.length) {
      ws.getCell(`E${rTong}`).value = { formula: `SUM(E${DATA_START}:E${lastDataRow})`, result: 0 };
    }

    // VAT = TỔNG * % lấy từ nhãn (mặc định 8% nếu không parse được)
    if (rVat && rTong) {
      const dText = (ws.getCell(`D${rVat}`).value || '').toString();
      const m = dText.match(/(\d+(?:\.\d+)?)\s*%/);
      const rate = m ? Number(m[1]) / 100 : 0.08;
      ws.getCell(`E${rVat}`).value = { formula: `E${rTong}*${rate}`, result: 0 };
    }

    // THÀNH TIỀN = TỔNG + VAT - GIẢM GIÁ
    if (rThanhTien && rTong) {
      const parts = [`E${rTong}`];
      if (rVat)     parts.push(`E${rVat}`);
      if (rGiamGia) parts.push(`-E${rGiamGia}`);
      ws.getCell(`E${rThanhTien}`).value = { formula: parts.join('+'), result: 0 };
    }


    // Tên file theo yyyyMMdd_HHmmss (có thể kèm tên KH đã bỏ dấu)
    const nameSlug = sanitize(customer.name || '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-zA-Z0-9]+/g, '_').replace(/^_+|_+$/g, '')
      .toLowerCase();
    const dateStr = ymdHMS();

    res.setHeader(
      'Content-Disposition',
      `attachment; filename="bao_gia_sintech${nameSlug ? '_' + nameSlug : ''}_${dateStr}.xlsx"`
    );
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Export failed!' });
  }
}
