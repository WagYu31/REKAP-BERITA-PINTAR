const express = require('express');
const axios = require('axios');
const { parse } = require('node-html-parser');
const ExcelJS = require('exceljs');
const cors = require('cors');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(express.json({ limit: '10mb' }));

// Serve static files from root and public
app.use(express.static(__dirname));
app.use(express.static(path.join(__dirname, 'public')));

// Bulan Bahasa Indonesia
const BULAN_INDONESIA = [
  'JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI',
  'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'
];

const BULAN_MAP = {
  'jan': 'JANUARI', 'januari': 'JANUARI',
  'feb': 'FEBRUARI', 'februari': 'FEBRUARI',
  'mar': 'MARET', 'maret': 'MARET',
  'apr': 'APRIL', 'april': 'APRIL',
  'mei': 'MEI', 'may': 'MEI',
  'jun': 'JUNI', 'juni': 'JUNI',
  'jul': 'JULI', 'juli': 'JULI',
  'agu': 'AGUSTUS', 'agustus': 'AGUSTUS', 'aug': 'AGUSTUS',
  'sep': 'SEPTEMBER', 'september': 'SEPTEMBER',
  'okt': 'OKTOBER', 'oktober': 'OKTOBER', 'oct': 'OKTOBER',
  'nov': 'NOVEMBER', 'november': 'NOVEMBER',
  'des': 'DESEMBER', 'desember': 'DESEMBER', 'dec': 'DESEMBER'
};

// Helper: Normalize Date & Extract Month
function parseDateAndMonth(dateString, rawHtml = '', url = '') {
  let parsedDate = new Date();
  let foundValidDate = false;
  let dateFormatted = '';
  let monthName = '';

  if (dateString) {
    const d = new Date(dateString);
    if (!isNaN(d.getTime()) && d.getFullYear() > 2000 && d.getFullYear() < 2100) {
      parsedDate = d;
      foundValidDate = true;
    }
  }

  if (!foundValidDate) {
    const combinedText = `${rawHtml} ${url}`;
    const numericMatch = combinedText.match(/\b(\d{1,2})[\/\.-](\d{1,2})[\/\.-](\d{2,4})\b/) ||
                         combinedText.match(/\b(\d{4})[\/\.-](\d{1,2})[\/\.-](\d{1,2})\b/);
    if (numericMatch) {
      let day, month, year;
      if (numericMatch[1].length === 4) {
        year = parseInt(numericMatch[1]);
        month = parseInt(numericMatch[2]) - 1;
        day = parseInt(numericMatch[3]);
      } else {
        day = parseInt(numericMatch[1]);
        month = parseInt(numericMatch[2]) - 1;
        year = parseInt(numericMatch[3]);
        if (year < 100) year += 2000;
      }
      if (month >= 0 && month < 12 && day > 0 && day <= 31) {
        parsedDate = new Date(year, month, day);
        foundValidDate = true;
      }
    }

    if (!foundValidDate) {
      const textMatch = combinedText.match(/\b(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})\b/);
      if (textMatch) {
        const day = parseInt(textMatch[1]);
        const monthStr = textMatch[2].toLowerCase();
        const year = parseInt(textMatch[3]);
        if (BULAN_MAP[monthStr]) {
          monthName = BULAN_MAP[monthStr];
          const monthIdx = BULAN_INDONESIA.indexOf(monthName);
          parsedDate = new Date(year, monthIdx, day);
          foundValidDate = true;
        }
      }
    }
  }

  if (foundValidDate) {
    const day = parsedDate.getDate();
    const month = parsedDate.getMonth() + 1;
    const year = parsedDate.getFullYear();
    dateFormatted = `${day}/${month}/${year}`;
    if (!monthName) {
      monthName = BULAN_INDONESIA[parsedDate.getMonth()] || 'MEI';
    }
  } else {
    const today = new Date();
    dateFormatted = `${today.getDate()}/${today.getMonth() + 1}/${today.getFullYear()}`;
    monthName = BULAN_INDONESIA[today.getMonth()];
  }

  return { dateFormatted, monthName };
}

function getMetaContent(root, selector) {
  const el = root.querySelector(selector);
  return el ? el.getAttribute('content') : null;
}

async function scrapeUrlMetadata(urlStr) {
  let targetUrl = urlStr.trim();
  if (!targetUrl.startsWith('http://') && !targetUrl.startsWith('https://')) {
    targetUrl = 'https://' + targetUrl;
  }

  try {
    const response = await axios.get(targetUrl, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7'
      },
      timeout: 10000
    });

    const html = response.data;
    const root = parse(html);

    let title = getMetaContent(root, 'meta[property="og:title"]') ||
                getMetaContent(root, 'meta[name="twitter:title"]') ||
                getMetaContent(root, 'meta[name="title"]') ||
                root.querySelector('title')?.textContent ||
                root.querySelector('h1')?.textContent ||
                '';

    title = title.trim()
      .replace(/\s+/g, ' ')
      .replace(/\s*[\-\|]\s*(MediaKomunikasi|Kompas|Detik|Antara|Liputan6|Tribun|Cnn|Tempo).*$/i, '');

    let rawDate = getMetaContent(root, 'meta[property="article:published_time"]') ||
                  getMetaContent(root, 'meta[name="pubdate"]') ||
                  getMetaContent(root, 'meta[name="publishdate"]') ||
                  getMetaContent(root, 'meta[itemprop="datePublished"]') ||
                  root.querySelector('time')?.getAttribute('datetime') ||
                  root.querySelector('time')?.textContent ||
                  '';

    if (!rawDate) {
      const scripts = root.querySelectorAll('script[type="application/ld+json"]');
      scripts.forEach(script => {
        try {
          const json = JSON.parse(script.textContent);
          if (json.datePublished) rawDate = json.datePublished;
          if (json['@graph']) {
            json['@graph'].forEach(item => {
              if (item.datePublished) rawDate = item.datePublished;
            });
          }
        } catch (e) {}
      });
    }

    const { dateFormatted, monthName } = parseDateAndMonth(rawDate, html, targetUrl);

    return {
      success: true,
      url: targetUrl,
      title: title || 'Judul Berita Tidak Ditemukan',
      tanggal: dateFormatted,
      bulan: monthName
    };
  } catch (err) {
    let fallbackTitle = 'Berita Online';
    try {
      const urlObj = new URL(targetUrl);
      const pathname = urlObj.pathname.replace(/\/$/, '');
      const slug = pathname.split('/').pop() || '';
      if (slug) {
        fallbackTitle = slug
          .replace(/[\-_]/g, ' ')
          .replace(/\.html?$/i, '')
          .replace(/\b\w/g, c => c.toUpperCase());
      }
    } catch (e) {}

    const { dateFormatted, monthName } = parseDateAndMonth('', '', targetUrl);

    return {
      success: false,
      url: targetUrl,
      title: fallbackTitle,
      tanggal: dateFormatted,
      bulan: monthName,
      error: err.message
    };
  }
}

// API Endpoint: Extract Metadata
app.post('/api/extract', async (req, res) => {
  const { urls } = req.body;
  if (!urls || !Array.isArray(urls) || urls.length === 0) {
    return res.status(400).json({ error: 'Atribut urls (array) wajib diisi' });
  }

  const results = await Promise.all(urls.map(url => scrapeUrlMetadata(url)));
  res.json({ data: results });
});

// API Endpoint: Export to Excel (.xlsx)
app.post('/api/export-excel', async (req, res) => {
  const { mediaName = 'MEDIAKOMUNIKASI.ID', period = 'MEI - JULI 2026', items = [] } = req.body;

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

  await workbook.xlsx.write(res);
  res.end();
});

// Fallback HTML route
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`====================================================`);
    console.log(`🚀 Aplikasi Rekap Berita Pintar Siap Digunakan!`);
    console.log(`👉 Buka di browser: http://localhost:${PORT}`);
    console.log(`====================================================`);
  });
}

module.exports = app;
