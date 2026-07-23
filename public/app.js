document.addEventListener('DOMContentLoaded', () => {
  // State
  let items = [];

  // Default Webhook URL provided by user
  const DEFAULT_WEBHOOK = "https://script.google.com/macros/s/AKfycby9evU_ZpfOwAgzn2YZiwtQGA3vPg12EHDZTLy0CghB7qGzQjaC3VtOmbHFe5HxB1My/exec";

  // DOM Elements
  const tabBtns = document.querySelectorAll('.tab-btn');
  const tabContents = document.querySelectorAll('.tab-content');
  const singleUrlInput = document.getElementById('singleUrlInput');
  const addSingleBtn = document.getElementById('addSingleBtn');
  const batchUrlInput = document.getElementById('batchUrlInput');
  const processBatchBtn = document.getElementById('processBatchBtn');
  const mediaNameInput = document.getElementById('mediaNameInput');
  const periodInput = document.getElementById('periodInput');

  const webhookUrlInput = document.getElementById('webhookUrlInput');
  const saveWebhookBtn = document.getElementById('saveWebhookBtn');
  const syncSheetsBtn = document.getElementById('syncSheetsBtn');

  const progressContainer = document.getElementById('progressContainer');
  const progressText = document.getElementById('progressText');
  const progressPercent = document.getElementById('progressPercent');
  const progressBarFill = document.getElementById('progressBarFill');

  const tableBody = document.getElementById('tableBody');
  const emptyState = document.getElementById('emptyState');
  const totalCountBadge = document.getElementById('totalCountBadge');

  const addManualBtn = document.getElementById('addManualBtn');
  const copyClipboardBtn = document.getElementById('copyClipboardBtn');
  const exportExcelBtn = document.getElementById('exportExcelBtn');
  const clearAllBtn = document.getElementById('clearAllBtn');
  const toastNotification = document.getElementById('toastNotification');

  // Load Saved Webhook URL or set Default
  const savedWebhook = localStorage.getItem('googleSheetWebhookUrl');
  if (savedWebhook) {
    webhookUrlInput.value = savedWebhook;
  } else {
    webhookUrlInput.value = DEFAULT_WEBHOOK;
    localStorage.setItem('googleSheetWebhookUrl', DEFAULT_WEBHOOK);
  }

  saveWebhookBtn.addEventListener('click', () => {
    const val = webhookUrlInput.value.trim();
    if (val) {
      localStorage.setItem('googleSheetWebhookUrl', val);
      showToast('URL Google Sheets Webhook berhasil disimpan! 💾', 'success');
    } else {
      localStorage.removeItem('googleSheetWebhookUrl');
      showToast('Webhook dihapus.', 'info');
    }
  });

  // Tab Switcher
  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      tabBtns.forEach(b => b.classList.remove('active'));
      tabContents.forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      const tabId = btn.getAttribute('data-tab') + 'Tab';
      document.getElementById(tabId).classList.add('active');
    });
  });

  // Toast Helper
  function showToast(message, type = 'info') {
    toastNotification.textContent = message;
    toastNotification.className = `toast ${type}`;
    toastNotification.classList.remove('hidden');
    setTimeout(() => {
      toastNotification.classList.add('hidden');
    }, 3500);
  }

  // Render Table
  function renderTable() {
    tableBody.innerHTML = '';

    if (items.length === 0) {
      emptyState.classList.remove('hidden');
      totalCountBadge.textContent = '0 Berita';
      return;
    }

    emptyState.classList.add('hidden');
    totalCountBadge.textContent = `${items.length} Berita`;

    items.forEach((item, index) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="cell-center" style="font-weight: 600;">${index + 1}</td>
        <td>
          <input type="text" value="${item.bulan || ''}" data-field="bulan" data-index="${index}" style="text-align: center; text-transform: uppercase; font-weight: 600;">
        </td>
        <td>
          <input type="text" value="${escapeHtml(item.title || '')}" data-field="title" data-index="${index}">
        </td>
        <td>
          <input type="text" value="${item.tanggal || ''}" data-field="tanggal" data-index="${index}" style="text-align: center;">
        </td>
        <td>
          <a href="${item.url}" target="_blank" class="cell-url" title="${item.url}">${item.url}</a>
        </td>
        <td class="cell-center">
          <button class="btn-delete-row" data-index="${index}" title="Hapus baris ini">&times;</button>
        </td>
      `;
      tableBody.appendChild(tr);
    });

    tableBody.querySelectorAll('input').forEach(input => {
      input.addEventListener('change', (e) => {
        const idx = parseInt(e.target.getAttribute('data-index'));
        const field = e.target.getAttribute('data-field');
        if (items[idx]) {
          items[idx][field] = e.target.value;
        }
      });
    });

    tableBody.querySelectorAll('.btn-delete-row').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const idx = parseInt(e.target.getAttribute('data-index'));
        items.splice(idx, 1);
        renderTable();
        showToast('Baris berhasil dihapus', 'info');
      });
    });
  }

  function escapeHtml(str) {
    return str.replace(/"/g, '&quot;');
  }

  // Fetch Metadata Single/Batch
  async function extractUrls(urlList) {
    if (urlList.length === 0) return;

    progressContainer.classList.remove('hidden');
    progressText.textContent = `Memproses 1 dari ${urlList.length} link...`;
    progressPercent.textContent = `0%`;
    progressBarFill.style.width = `0%`;

    try {
      const response = await fetch('/api/extract', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ urls: urlList })
      });

      const json = await response.json();
      if (json.data && Array.isArray(json.data)) {
        json.data.forEach(res => {
          items.push({
            bulan: res.bulan || 'MEI',
            title: res.title || 'Judul Berita',
            tanggal: res.tanggal || '',
            url: res.url
          });
        });

        renderTable();
        showToast(`Berhasil mengekstrak ${json.data.length} link berita!`, 'success');

        // Auto Sync if webhook is set
        const webhookUrl = webhookUrlInput.value.trim() || DEFAULT_WEBHOOK;
        if (webhookUrl) {
          syncToGoogleSheets(webhookUrl, true);
        }
      }
    } catch (err) {
      showToast('Gagal mengambil data berita: ' + err.message, 'danger');
    } finally {
      progressBarFill.style.width = `100%`;
      progressPercent.textContent = `100%`;
      progressText.textContent = `Selesai memproses ${urlList.length} link.`;
      setTimeout(() => {
        progressContainer.classList.add('hidden');
      }, 2000);
    }
  }

  // Sync to Google Sheets Function
  async function syncToGoogleSheets(webhookUrl, isAuto = false) {
    const targetUrl = webhookUrl || webhookUrlInput.value.trim() || DEFAULT_WEBHOOK;

    if (items.length === 0) {
      showToast('Tabel masih kosong! Tambahkan link berita dulu.', 'warning');
      return;
    }

    const payload = {
      mediaName: mediaNameInput.value.trim() || 'MEDIAKOMUNIKASI.ID',
      period: periodInput.value.trim() || 'MEI - JULI 2026',
      items: items
    };

    try {
      showToast(isAuto ? 'Auto Sync ke Google Spreadsheet...' : 'Mengirim data ke Google Spreadsheet...', 'info');
      
      await fetch(targetUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload),
        mode: 'no-cors'
      });

      showToast('✅ Data berhasil dikirim ke Google Spreadsheet Anda! 🚀', 'success');
    } catch (err) {
      showToast('Gagal sync ke Google Sheets: ' + err.message, 'danger');
    }
  }

  syncSheetsBtn.addEventListener('click', () => {
    const webhookUrl = webhookUrlInput.value.trim() || DEFAULT_WEBHOOK;
    syncToGoogleSheets(webhookUrl, false);
  });

  // Add Single URL
  addSingleBtn.addEventListener('click', () => {
    const url = singleUrlInput.value.trim();
    if (!url) {
      showToast('Harap masukkan URL/Link berita terlebih dahulu', 'warning');
      return;
    }
    extractUrls([url]);
    singleUrlInput.value = '';
  });

  singleUrlInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      addSingleBtn.click();
    }
  });

  // Batch Extract
  processBatchBtn.addEventListener('click', () => {
    const rawText = batchUrlInput.value.trim();
    if (!rawText) {
      showToast('Harap tempel beberapa link berita di kolom teks', 'warning');
      return;
    }

    const lines = rawText.split('\n')
      .map(l => l.trim())
      .filter(l => l.startsWith('http://') || l.startsWith('https://') || l.includes('.'));

    if (lines.length === 0) {
      showToast('Tidak ada URL valid yang ditemukan', 'warning');
      return;
    }

    extractUrls(lines);
    batchUrlInput.value = '';
  });

  // Add Manual Row
  addManualBtn.addEventListener('click', () => {
    const today = new Date();
    const formattedDate = `${today.getDate()}/${today.getMonth() + 1}/${today.getFullYear()}`;
    items.push({
      bulan: 'MEI',
      title: 'Judul Berita Baru',
      tanggal: formattedDate,
      url: 'https://mediakomunikasi.id/'
    });
    renderTable();
    showToast('Baris manual ditambahkan', 'info');
  });

  // Export to Excel (.xlsx)
  exportExcelBtn.addEventListener('click', async () => {
    if (items.length === 0) {
      showToast('Tabel masih kosong! Tambahkan link berita terlebih dahulu.', 'warning');
      return;
    }

    const mediaName = mediaNameInput.value.trim() || 'MEDIAKOMUNIKASI.ID';
    const period = periodInput.value.trim() || 'MEI - JULI 2026';

    try {
      showToast('Sedang membuat file Excel...', 'info');
      const response = await fetch('/api/export-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ mediaName, period, items })
      });

      if (!response.ok) throw new Error('Gagal mengekspor file Excel');

      const blob = await response.blob();
      const downloadUrl = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = downloadUrl;
      a.download = `REKAP_BERITA_${mediaName.replace(/[^a-zA-Z0-9]/g, '_')}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(downloadUrl);

      showToast('File Excel (.xlsx) berhasil didownload! 🚀', 'success');
    } catch (err) {
      showToast('Error: ' + err.message, 'danger');
    }
  });

  // Copy Table for Excel (TSV format)
  copyClipboardBtn.addEventListener('click', () => {
    if (items.length === 0) {
      showToast('Tabel masih kosong!', 'warning');
      return;
    }

    let tsv = 'NO\tBULAN\tJUDUL BERITA\tTANGGAL PUBLISH\tALAMAT URL\n';
    items.forEach((item, index) => {
      tsv += `${index + 1}\t${item.bulan || ''}\t${item.title || ''}\t${item.tanggal || ''}\t${item.url || ''}\n`;
    });

    navigator.clipboard.writeText(tsv).then(() => {
      showToast('Tabel berhasil disalin! Anda bisa langsung Paste (Ctrl+V) di Excel.', 'success');
    }).catch(err => {
      showToast('Gagal menyalin ke clipboard', 'danger');
    });
  });

  // Clear All
  clearAllBtn.addEventListener('click', () => {
    if (items.length === 0) return;
    if (confirm('Apakah Anda yakin ingin menghapus semua baris rekap?')) {
      items = [];
      renderTable();
      showToast('Semua baris telah dihapus', 'info');
    }
  });

  // Initial render
  renderTable();
});
