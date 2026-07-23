const ExcelJS = require('exceljs');

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { mediaName = 'MEDIAKOMUNIKASI.ID', period = 'MEI - JULI 2026', items = [] } = req.body || {};

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Lamp1. URL Berita Cetak');

  worksheet.pageSetup = { paperSize: 9, orientation: 'landscape' };

  const headerTitleFont = { name: 'Calibri', size: 11, bold: true };
  const redTitleFont = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFC00000' } };
  const tableHeaderFont = { name: 'Calibri', size: 10, bold: true };
  const cellFont = { name: 'Calibri', size: 9 };
  const urlFont = { name: 'Calibri', size: 9, color: { argb: 'FF0000FF' }, underline: true };

  const thinBorder = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  const row1 = worksheet.getRow(1);
  row1.getCell(1).value = 'Lampiran 1.';
  row1.getCell(1).font = headerTitleFont;

  const row2 = worksheet.getRow(2);
  row2.getCell(1).value = `Rangkuman Berita Online ${mediaName.toUpperCase()} bulan ${period.toUpperCase()}`;
  row2.getCell(1).font = redTitleFont;

  const headers = ['NO', 'BULAN', 'JUDUL BERITA', 'TANGGAL PUBLISH', 'ALAMAT URL'];
  const headerRow = worksheet.getRow(4);

  headers.forEach((h, idx) => {
    const cell = headerRow.getCell(idx + 1);
    cell.value = h;
    cell.font = tableHeaderFont;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = thinBorder;
  });

  worksheet.getColumn(1).width = 6;
  worksheet.getColumn(2).width = 12;
  worksheet.getColumn(3).width = 65;
  worksheet.getColumn(4).width = 20;
  worksheet.getColumn(5).width = 75;

  items.forEach((item, index) => {
    const rowNumber = 5 + index;
    const row = worksheet.getRow(rowNumber);

    const cellNo = row.getCell(1);
    cellNo.value = index + 1;
    cellNo.font = cellFont;
    cellNo.alignment = { vertical: 'middle', horizontal: 'center' };
    cellNo.border = thinBorder;

    const cellBulan = row.getCell(2);
    cellBulan.value = (item.bulan || 'MEI').toUpperCase();
    cellBulan.font = cellFont;
    cellBulan.alignment = { vertical: 'middle', horizontal: 'center' };
    cellBulan.border = thinBorder;

    const cellJudul = row.getCell(3);
    cellJudul.value = item.title || '';
    cellJudul.font = cellFont;
    cellJudul.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    cellJudul.border = thinBorder;

    const cellTanggal = row.getCell(4);
    cellTanggal.value = item.tanggal || '';
    cellTanggal.font = cellFont;
    cellTanggal.alignment = { vertical: 'middle', horizontal: 'center' };
    cellTanggal.border = thinBorder;

    const cellUrl = row.getCell(5);
    cellUrl.value = { text: item.url || '', hyperlink: item.url || '' };
    cellUrl.font = urlFont;
    cellUrl.alignment = { vertical: 'middle', horizontal: 'left' };
    cellUrl.border = thinBorder;
  });

  const filename = `REKAP_BERITA_${mediaName.replace(/[^a-zA-Z0-9]/g, '_')}_${Date.now()}.xlsx`;
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

  const buffer = await workbook.xlsx.writeBuffer();
  res.send(buffer);
};
