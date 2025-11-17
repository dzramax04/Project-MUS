// Global variable to store current sampling results
let currentSampledData = [];
let currentNamaHeader = '';
let currentNamaAkun = '';
let currentAuditInfo = {};

// Set today's date as default
document.getElementById('tanggalDibuat').valueAsDate = new Date();
document.getElementById('tanggalDireview').valueAsDate = new Date();

function formatCurrency(amount) {
    return new Intl.NumberFormat('id-ID', {
        style: 'currency',
        currency: 'IDR',
        minimumFractionDigits: 0
    }).format(amount);
}

// Safe getElementById with null check
function safeGetElement(id) {
    const element = document.getElementById(id);
    if (!element) {
        console.warn(`Element with id '${id}' not found`);
    }
    return element;
}

// Safe set textContent with null check
function safeSetTextContent(id, text) {
    const element = safeGetElement(id);
    if (element) {
        element.textContent = text;
    }
}

// Show error modal with detailed information
function showErrorModal(message, counts = null) {
    const errorMessage = safeGetElement('errorMessage');
    const errorDetails = safeGetElement('errorDetails');
    const errorRows = safeGetElement('errorRows');
    const errorSummary = safeGetElement('errorSummary');
    if (errorMessage) {
        errorMessage.innerHTML = message.replace(/\n/g, '<br>');
    }
    if (errorDetails && counts) {
        // Show detailed error information
        errorDetails.style.display = 'block';
        if (errorRows) {
            // Create detailed row information
            const columnNames = {
                'tanggal': 'Tanggal',
                'voucher': 'Voucher', 
                'keterangan': 'Keterangan',
                'nominal': 'Nominal'
            };
            let rowsHTML = '';
            const countsArray = Object.entries(counts);
            const maxCount = Math.max(...countsArray.map(([_, count]) => count));
            const minCount = Math.min(...countsArray.map(([_, count]) => count));
            countsArray.forEach(([key, count]) => {
                const columnName = columnNames[key] || key;
                const isMismatched = count !== maxCount;
                rowsHTML += `
                    <div class="error-row">
                        <span class="error-column">${columnName}</span>
                        <span class="error-count ${isMismatched ? 'text-danger' : 'text-success'}">${count} baris</span>
                    </div>
                `;
            });
            errorRows.innerHTML = rowsHTML;
        }
        if (errorSummary) {
            const maxCount = Math.max(...Object.values(counts));
            const minCount = Math.min(...Object.values(counts));
            const difference = maxCount - minCount;
            // Find which columns have the maximum and minimum counts
            const maxColumns = [];
            const minColumns = [];
            for (const [key, count] of Object.entries(counts)) {
                const columnNames = {
                    'tanggal': 'Tanggal',
                    'voucher': 'Voucher', 
                    'keterangan': 'Keterangan',
                    'nominal': 'Nominal'
                };
                const columnName = columnNames[key] || key;
                if (count === maxCount) {
                    maxColumns.push(columnName);
                }
                if (count === minCount) {
                    minColumns.push(columnName);
                }
            }
            let summaryHTML = `<strong>Analisis Perbedaan:</strong><br>`;
            summaryHTML += `• Kolom dengan jumlah baris terbanyak (${maxCount}): <strong>${maxColumns.join(', ')}</strong><br>`;
            summaryHTML += `• Kolom dengan jumlah baris paling sedikit (${minCount}): <strong>${minColumns.join(', ')}</strong><br>`;
            summaryHTML += `• Selisih jumlah baris: <strong>${difference}</strong> baris`;
            errorSummary.innerHTML = summaryHTML;
        }
    } else {
        errorDetails.style.display = 'none';
    }
    const modal = new bootstrap.Modal(document.getElementById('errorModal'));
    modal.show();
}

// Enhanced date parsing function
function parseDate(dateStr) {
    if (!dateStr) return null;
    const input = dateStr.trim();
    // Try ISO format first (yyyy-mm-dd)
    if (/^\d{4}-\d{2}-\d{2}$/.test(input)) {
        return input;
    }
    // Try dd/mm/yyyy format
    const ddMmYyyyMatch = input.match(/^(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})$/);
    if (ddMmYyyyMatch) {
        let day = ddMmYyyyMatch[1].padStart(2, '0');
        let month = ddMmYyyyMatch[2].padStart(2, '0');
        let year = ddMmYyyyMatch[3];
        if (year.length === 2) {
            year = parseInt(year) >= 70 ? '19' + year : '20' + year;
        }
        return `${year}-${month}-${day}`;
    }
    // Try "18 Des 2024" or "18 Dec 2024" format
    const monthNames = {
        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04', 'may': '05', 'jun': '06',
        'jul': '07', 'aug': '08', 'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12',
        'januari': '01', 'februari': '02', 'maret': '03', 'april': '04', 'mei': '05', 'juni': '06',
        'juli': '07', 'agustus': '08', 'september': '09', 'oktober': '10', 'november': '11', 'desember': '12'
    };
    const dateParts = input.split(/\s+/);
    if (dateParts.length >= 2) {
        const day = dateParts[0];
        const monthAbbr = dateParts[1].toLowerCase().substring(0, 3);
        const year = dateParts[2] || new Date().getFullYear().toString();
        if (day && monthAbbr && monthNames[monthAbbr]) {
            const formattedDay = day.padStart(2, '0');
            const formattedMonth = monthNames[monthAbbr];
            const formattedYear = year.length === 2 ? (parseInt(year) >= 70 ? '19' + year : '20' + year) : year;
            return `${formattedYear}-${formattedMonth}-${formattedDay}`;
        }
    }
    // If all parsing fails, return original string (will be handled as invalid)
    return input;
}

// Enhanced nominal parsing function
function parseNominal(nominalStr) {
    if (!nominalStr) return 0;
    let cleanStr = nominalStr.trim().replace(/[^0-9.,-]/g, '');
    // Handle cases like "1.000.000,00" (Indonesian format)
    if (cleanStr.includes(',') && cleanStr.lastIndexOf(',') > cleanStr.lastIndexOf('.')) {
        // Comma is decimal separator
        cleanStr = cleanStr.replace(/\./g, '').replace(',', '.');
    } 
    // Handle cases like "1,000,000.00" (US format)
    else if (cleanStr.includes('.') && cleanStr.lastIndexOf('.') > cleanStr.lastIndexOf(',')) {
        // Dot is decimal separator
        cleanStr = cleanStr.replace(/,/g, '');
    }
    // Handle cases with only commas or only dots as thousand separators
    else {
        cleanStr = cleanStr.replace(/[.,]/g, '');
    }
    const num = parseFloat(cleanStr);
    return isNaN(num) ? 0 : Math.round(num);
}

// Generate dummy data
function generateDummyData(count = 50) {
    const keteranganList = [
        'Pembelian barang kantor', 'Jasa konsultasi', 'Transportasi', 'Perjalanan dinas',
        'Biaya telepon', 'Listrik dan air', 'Sewa gedung', 'Gaji karyawan',
        'Biaya perawatan', 'Pembelian peralatan', 'Jasa pemasaran', 'Biaya pelatihan',
        'Langganan software', 'Biaya legal', 'Pajak dan retribusi', 'Biaya administrasi',
        'Pembelian bahan baku', 'Biaya pengiriman', 'Jasa IT', 'Perbaikan fasilitas'
    ];
    const tanggalData = [];
    const voucherData = [];
    const keteranganData = [];
    const nominalData = [];
    // Generate random dates within the last year
    const today = new Date();
    const oneYearAgo = new Date();
    oneYearAgo.setFullYear(today.getFullYear() - 1);
    for (let i = 0; i < count; i++) {
        // Random date
        const randomDate = new Date(oneYearAgo.getTime() + Math.random() * (today.getTime() - oneYearAgo.getTime()));
        const day = String(randomDate.getDate()).padStart(2, '0');
        const month = String(randomDate.getMonth() + 1).padStart(2, '0');
        const year = randomDate.getFullYear();
        // Use dd/mm/yyyy format for dummy data
        tanggalData.push(`${day}/${month}/${year}`);
        // Voucher
        voucherData.push(`VOU-${String(i + 1).padStart(3, '0')}`);
        // Keterangan
        const keterangan = keteranganList[Math.floor(Math.random() * keteranganList.length)];
        keteranganData.push(keterangan);
        // Nominal (random between 500,000 and 10,000,000) - use Indonesian format
        const nominal = Math.floor(Math.random() * 9500000) + 500000;
        const formattedNominal = new Intl.NumberFormat('id-ID', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(nominal);
        nominalData.push(formattedNominal);
    }
    return { tanggalData, voucherData, keteranganData, nominalData };
}

function fillDummyData() {
    const { tanggalData, voucherData, keteranganData, nominalData } = generateDummyData(50);
    safeGetElement('tanggalInput').value = tanggalData.join('\n');
    safeGetElement('voucherInput').value = voucherData.join('\n');
    safeGetElement('keteranganInput').value = keteranganData.join('\n');
    safeGetElement('nominalInput').value = nominalData.join('\n');
    // Auto-fill audit info with dummy data
    safeGetElement('namaKlien').value = 'PT Contoh Perusahaan';
    safeGetElement('namaHeader').value = 'Biaya Operasional';
    safeGetElement('namaAkun').value = 'Akun 12345';
    safeGetElement('dibuatOleh').value = 'Auditor Dummy';
    safeGetElement('direviewOleh').value = 'Reviewer Dummy';
    safeGetElement('schedule').value = 'Audit Q4 2025';
    // Show success message
    const successModal = document.createElement('div');
    successModal.innerHTML = `
        <div class="modal fade" id="successModal" tabindex="-1">
            <div class="modal-dialog modal-dialog-centered">
                <div class="modal-content">
                    <div class="modal-header bg-success text-white">
                        <h5 class="modal-title"><i class="fas fa-check-circle me-2"></i>Data Berhasil Diisi</h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body text-center">
                        <i class="fas fa-database fa-3x text-success mb-3"></i>
                        <p>Data dummy berhasil diisi dengan 50 baris data contoh!</p>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-success" data-bs-dismiss="modal">OK</button>
                    </div>
                </div>
            </div>
        </div>
    `;
    document.body.appendChild(successModal);
    const modal = new bootstrap.Modal(document.getElementById('successModal'));
    modal.show();
    // Clean up after modal closes
    document.getElementById('successModal').addEventListener('hidden.bs.modal', function() {
        document.body.removeChild(successModal);
    });
}

function validateDataColumns() {
    const tanggalText = safeGetElement('tanggalInput').value.trim();
    const voucherText = safeGetElement('voucherInput').value.trim();
    const keteranganText = safeGetElement('keteranganInput').value.trim();
    const nominalText = safeGetElement('nominalInput').value.trim();
    // Handle empty inputs
    const tanggalLines = tanggalText ? tanggalText.split('\n').filter(line => line.trim() !== '') : [];
    const voucherLines = voucherText ? voucherText.split('\n').filter(line => line.trim() !== '') : [];
    const keteranganLines = keteranganText ? keteranganText.split('\n').filter(line => line.trim() !== '') : [];
    const nominalLines = nominalText ? nominalText.split('\n').filter(line => line.trim() !== '') : [];

    const counts = {
        tanggal: tanggalLines.length,
        voucher: voucherLines.length,
        keterangan: keteranganLines.length,
        nominal: nominalLines.length
    };

    // Check if all counts are equal and greater than 0
    const allEqual = counts.tanggal === counts.voucher && 
                     counts.voucher === counts.keterangan && 
                     counts.keterangan === counts.nominal;
    const allNonZero = counts.tanggal > 0 && counts.voucher > 0 && counts.keterangan > 0 && counts.nominal > 0;

    if (!allNonZero) {
        if (counts.tanggal === 0 && counts.voucher === 0 && counts.keterangan === 0 && counts.nominal === 0) {
            throw new Error('Semua kolom data harus diisi!');
        } else {
            const missingColumns = [];
            if (counts.tanggal === 0) missingColumns.push('Tanggal');
            if (counts.voucher === 0) missingColumns.push('Voucher');
            if (counts.keterangan === 0) missingColumns.push('Keterangan');
            if (counts.nominal === 0) missingColumns.push('Nominal');
            throw new Error(`Kolom berikut kosong: ${missingColumns.join(', ')}\nPastikan semua kolom memiliki data!`);
        }
    }
    if (!allEqual) {
        throw {
            message: 'Jumlah baris tidak sama di semua kolom!\nPastikan setiap kolom memiliki jumlah baris yang sama.',
            counts: counts
        };
    }
    return {
        tanggalLines,
        voucherLines,
        keteranganLines,
        nominalLines
    };
}

function parseDataPerColumn() {
    const validation = validateDataColumns();
    const data = [];
    for (let i = 0; i < validation.tanggalLines.length; i++) {
        const tanggal = parseDate(validation.tanggalLines[i]);
        const voucher = validation.voucherLines[i].trim();
        const keterangan = validation.keteranganLines[i].trim();
        const nominal = parseNominal(validation.nominalLines[i]);
        data.push({
            tanggal,
            voucher,
            keterangan,
            nominal
        });
    }
    return data;
}

function getRandomItems(array, count) {
    const shuffled = [...array].sort(() => 0.5 - Math.random());
    return shuffled.slice(0, Math.min(count, shuffled.length));
}

function performSampling(data, method) {
    // Sort by nominal descending
    const sortedData = [...data].sort((a, b) => b.nominal - a.nominal);
    let sampledData = [];
    switch(method) {
        case 'rendah':
            // Top 10 by nominal
            sampledData = sortedData.slice(0, Math.min(10, sortedData.length));
            break;
        case 'moderate':
            // Top 10 + 10 random
            const top10 = sortedData.slice(0, Math.min(10, sortedData.length));
            const remainingData = sortedData.length > 10 ? sortedData.slice(10) : [];
            const random10 = remainingData.length > 0 ? getRandomItems(remainingData, 10) : [];
            sampledData = [...top10, ...random10];
            break;
        case 'tinggi':
            // Top 15 + 15 random
            const top15 = sortedData.slice(0, Math.min(15, sortedData.length));
            const remainingDataHigh = sortedData.length > 15 ? sortedData.slice(15) : [];
            const random15 = remainingDataHigh.length > 0 ? getRandomItems(remainingDataHigh, 15) : [];
            sampledData = [...top15, ...random15];
            break;
        default:
            sampledData = sortedData.slice(0, Math.min(10, sortedData.length));
    }
    // Remove duplicates (in case random selection picks from top)
    const uniqueData = [];
    const seen = new Set();
    for (const item of sampledData) {
        const key = `${item.tanggal}|${item.voucher}|${item.keterangan}|${item.nominal}`;
        if (!seen.has(key)) {
            seen.add(key);
            uniqueData.push(item);
        }
    }
    return uniqueData;
}

function renderResults(data) {
    const tbody = safeGetElement('resultsTable');
    if (!tbody) return;
    if (data.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="4" class="text-center py-4">
                    <i class="fas fa-exclamation-triangle fa-2x text-warning mb-2"></i>
                    <div class="text-muted">Tidak ada data untuk ditampilkan</div>
                </td>
            </tr>
        `;
        return;
    }
    // Sort sampled data by nominal descending (highest to lowest)
    const sortedData = data.sort((a, b) => b.nominal - a.nominal);
    tbody.innerHTML = sortedData.map(item => `
        <tr>
            <td>${item.tanggal}</td>
            <td><code>${item.voucher}</code></td>
            <td>${item.keterangan}</td>
            <td class="text-end fw-medium">${formatCurrency(item.nominal)}</td>
        </tr>
    `).join('');
    // Enable export button
    const exportBtn = safeGetElement('exportBtn');
    if (exportBtn) {
        exportBtn.disabled = false;
    }
}

function updateSummaryPanels(totalData, totalNominal, method, sampleCount, sampledData, originalData) {
    // Update main summary metrics
    safeSetTextContent('totalData', totalData.toLocaleString('id-ID'));
    safeSetTextContent('totalNominal', formatCurrency(totalNominal));
    const methodLabels = {
        'rendah': 'Rendah (Top 10)',
        'moderate': 'Moderate (Top 10 + 10 Acak)',
        'tinggi': 'Tinggi (Top 15 + 15 Acak)'
    };
    const methodLabel = methodLabels[method] || methodLabels['moderate'];
    safeSetTextContent('metodeSampling', methodLabel);
    safeSetTextContent('jumlahSampel', sampleCount.toLocaleString('id-ID'));
    // Update detailed summary tab
    safeSetTextContent('summaryTotalData', totalData.toLocaleString('id-ID'));
    safeSetTextContent('summaryTotalNominal', formatCurrency(totalNominal));
    safeSetTextContent('summaryMetodeSampling', methodLabel);
    safeSetTextContent('summaryJumlahSampel', sampleCount.toLocaleString('id-ID'));
    // Update sampling analysis
    if (sampledData.length > 0) {
        const minNominal = Math.min(...sampledData.map(item => item.nominal));
        const maxNominal = Math.max(...sampledData.map(item => item.nominal));
        const avgNominal = sampledData.reduce((sum, item) => sum + item.nominal, 0) / sampledData.length;
        const coverage = totalNominal > 0 ? (sampledData.reduce((sum, item) => sum + item.nominal, 0) / totalNominal) * 100 : 0;
        safeSetTextContent('summaryMinNominal', formatCurrency(minNominal));
        safeSetTextContent('summaryMaxNominal', formatCurrency(maxNominal));
        safeSetTextContent('summaryAvgNominal', formatCurrency(Math.round(avgNominal)));
        safeSetTextContent('summaryCoverage', coverage.toFixed(2) + '%');
    } else {
        safeSetTextContent('summaryMinNominal', 'Rp 0');
        safeSetTextContent('summaryMaxNominal', 'Rp 0');
        safeSetTextContent('summaryAvgNominal', 'Rp 0');
        safeSetTextContent('summaryCoverage', '0%');
    }
    // Update audit info
    safeSetTextContent('summaryNamaKlien', safeGetElement('namaKlien').value || '-');
    safeSetTextContent('summaryNamaHeader', safeGetElement('namaHeader').value || '-');
    safeSetTextContent('summaryNamaAkun', safeGetElement('namaAkun').value || '-');
    safeSetTextContent('summaryDibuatOleh', safeGetElement('dibuatOleh').value || '-');
    safeSetTextContent('summaryDireviewOleh', safeGetElement('direviewOleh').value || '-');
    safeSetTextContent('summarySchedule', safeGetElement('schedule').value || '-');
}

function processData() {
    const namaKlien = safeGetElement('namaKlien').value;
    currentNamaHeader = safeGetElement('namaHeader').value || 'Sampling Results';
    currentNamaAkun = safeGetElement('namaAkun').value || '';
    // Store audit info for export
    currentAuditInfo = {
        dibuatOleh: safeGetElement('dibuatOleh').value || '',
        tanggalDibuat: safeGetElement('tanggalDibuat').value || '',
        direviewOleh: safeGetElement('direviewOleh').value || '',
        tanggalDireview: safeGetElement('tanggalDireview').value || '',
        namaKlien: safeGetElement('namaKlien').value || '',
        schedule: safeGetElement('schedule').value || ''
    };
    if (!namaKlien.trim()) {
        showErrorModal('Nama Klien wajib diisi!');
        return;
    }
    try {
        const parsedData = parseDataPerColumn();
        if (parsedData.length === 0) {
            showErrorModal('Tidak ada data yang valid ditemukan!');
            return;
        }
        const totalNominal = parsedData.reduce((sum, item) => sum + item.nominal, 0);
        const samplingMethod = safeGetElement('samplingMethod').value;
        const sampledData = performSampling(parsedData, samplingMethod);
        currentSampledData = sampledData;
        renderResults(sampledData);
        updateSummaryPanels(parsedData.length, totalNominal, samplingMethod, sampledData.length, sampledData, parsedData);
    } catch (error) {
        console.error('Error processing data:', error);
        if (error.counts) {
            // This is a validation error with counts
            showErrorModal(error.message, error.counts);
        } else {
            // This is a regular error
            showErrorModal(error.message);
        }
    }
}

    // ExcelJS Export function with all requested specifications
    async function exportToExcelJasa() {
        if (currentSampledData.length === 0) {
            alert('Tidak ada data sampling untuk diekspor!');
            return;
        }

        try {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Uji Detil');

            // === 1. Header besar A1:O5 (tanpa logo dulu, bisa ditambahkan nanti via base64)
            worksheet.mergeCells('A1:O5');
            const titleCell = worksheet.getCell('A1');
            titleCell.value = currentNamaHeader || 'Sampling Results';
            titleCell.font = { bold: true, size: 16 };
            titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
            titleCell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };

            // === 2. UJI DETIL DOKUMEN (P1:W5)
            worksheet.mergeCells('P1:W5');
            worksheet.getCell('P1').value = 'UJI DETIL DOKUMEN (Test Of Detail)';
            worksheet.getCell('P1').font = { bold: true, size: 14 };
            worksheet.getCell('P1').alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getCell('P1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6E6E6' } };
            worksheet.getCell('P1').border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };

            // === 3. INDEKS (X1:Y1 dan X2:Y5)
            worksheet.mergeCells('X1:Y1');
            worksheet.getCell('X1').value = 'INDEKS :';
            worksheet.getCell('X1').font = { bold: true };
            worksheet.getCell('X1').alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getCell('X1').border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };

            worksheet.mergeCells('X2:Y5');
            worksheet.getCell('X2').value = 'XX';
            worksheet.getCell('X2').alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getCell('X2').border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };

            // === 4. Info Klien & Periode
            worksheet.getCell('A7').value = 'Klien :';
            worksheet.getCell('A7').font = { bold: true };
            worksheet.getCell('B7').value = currentAuditInfo.namaKlien || '';

            worksheet.getCell('A8').value = 'Periode :';
            worksheet.getCell('A8').font = { bold: true };
            worksheet.getCell('B8').value = currentAuditInfo.schedule || '';  // <-- ini kuncinya

            // === 5. Auditor & Reviewer
            worksheet.getCell('U7').value = 'Dibuat oleh :';
            worksheet.getCell('U7').font = { bold: true };
            worksheet.getCell('V7').value = currentAuditInfo.dibuatOleh || '';

            worksheet.getCell('U8').value = 'Tanggal :';
            worksheet.getCell('U8').font = { bold: true };
            worksheet.getCell('V8').value = currentAuditInfo.tanggalDibuat || '';

            worksheet.getCell('U9').value = 'Paraf :';
            worksheet.getCell('U9').font = { bold: true };

            worksheet.getCell('X7').value = 'Direview oleh :';
            worksheet.getCell('X7').font = { bold: true };
            worksheet.getCell('Y7').value = currentAuditInfo.direviewOleh || '';

            worksheet.getCell('X8').value = 'Tanggal :';
            worksheet.getCell('X8').font = { bold: true };
            worksheet.getCell('Y8').value = currentAuditInfo.tanggalDireview || '';

            worksheet.getCell('X9').value = 'Paraf :';
            worksheet.getCell('X9').font = { bold: true };

            // === 6. Header Tabel (baris 11–12)
            const headerConfigs = [
                { range: 'A11:A12', text: 'No', width: 9 },
                { range: 'B11:B12', text: 'Tgl', width: 12 },
                { range: 'C11:C12', text: 'No. Voucher', width: 18 },
                { range: 'D11:D12', text: 'Nama Transaksi', width: 25 },
                { range: 'E11:E12', text: 'Jumlah Menurut GL', width: 18 },
                { range: 'F11:H11', text: 'Perhitungan Pajak' },
                { cell: 'F12', text: 'DPP' },
                { cell: 'G12', text: 'PPN' },
                { cell: 'H12', text: 'PPh' },
                { range: 'I11:I12', text: 'Nett Amount\nBilled (Nilai Tagihan Bersih)', width: 20 },
                { range: 'J11:L11', text: 'Purchase Order' },
                { cell: 'J12', text: 'No PO', width: 16, noWrap: true },
                { cell: 'K12', text: 'Tgl PO', width: 12, noWrap: true },
                { cell: 'L12', text: 'Nominal PO', width: 16, noWrap: true },
                { range: 'M11:O11', text: 'Invoice' },
                { cell: 'M12', text: 'No Invoice', width: 16, noWrap: true },
                { cell: 'N12', text: 'Tgl Invoice', width: 12, noWrap: true },
                { cell: 'O12', text: 'Nominal Invoice', width: 16, noWrap: true },
                { range: 'P11:R11', text: 'Faktur' },
                { cell: 'P12', text: 'No Faktur', width: 16, noWrap: true },
                { cell: 'Q12', text: 'Tgl Faktur', width: 12, noWrap: true },
                { cell: 'R12', text: 'Nominal Faktur', width: 16, noWrap: true },
                { range: 'S11:U11', text: 'Bukti Bayar' },
                { cell: 'S12', text: 'No Bukti', width: 16, noWrap: true },
                { cell: 'T12', text: 'Tgl Bukti', width: 12, noWrap: true },
                { cell: 'U12', text: 'Nominal Bukti', width: 16, noWrap: true },
                // Asersi
                { range: 'V11:Y11', text: 'Asersi', width: 16 },
                { cell: 'V12', text: 'Keterjadian', width: 14, noWrap: true },
                { cell: 'W12', text: 'Keakurasian', width: 14, noWrap: true },
                { cell: 'X12', text: 'Pisah Batas', width: 14, noWrap: true },
                { cell: 'Y12', text: 'Klasifikasi', width: 14, noWrap: true },
                // Kelengkapan Dokumen
                { range: 'Z11:Z12', text: 'Kelengkapan Dokumen', width: 20 },
                // Temuan (AA–AD)
                { range: 'AA11:AA12', text: 'Selisih', width: 14 },
                { range: 'AB11:AB12', text: 'Ket.', width: 14 },
                { range: 'AC11:AC12', text: 'Otorisasi\nBukti Bayar', width: 16 },
                { range: 'AD11:AD12', text: 'Deskripsi Temuan', width: 25 },
                // Baru: File Eksternal
                { range: 'AE11:AE12', text: 'File Eksternal', width: 20 }
            ];

            headerConfigs.forEach(item => {
                if (item.range) {
                    worksheet.mergeCells(item.range);
                    const cell = worksheet.getCell(item.range.split(':')[0]);
                    cell.value = item.text;
                    cell.font = { bold: true };
                    const hasNewline = (typeof item.text === 'string') && item.text.includes('\n');
                    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: hasNewline };
                    if (item.width) {
                        const colLetter = item.range.split(':')[0].replace(/[0-9]/g, '');
                        worksheet.getColumn(colLetter).width = item.width;
                    }
                }
                if (item.cell) {
                    const cell = worksheet.getCell(item.cell);
                    cell.value = item.text;
                    cell.font = { bold: true };
                    const hasNewline = (typeof item.text === 'string') && item.text.includes('\n');
                    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: hasNewline };
                    if (item.width) {
                        const colLetter = item.cell.replace(/[0-9]/g, '');
                        worksheet.getColumn(colLetter).width = item.width;
                    }
                }
            });

            // Border header baris 11–12
            [11, 12].forEach(rowNum => {
                const row = worksheet.getRow(rowNum);
                row.eachCell({ includeEmpty: true }, cell => {
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                    if (!cell.isMerged) {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
                    }
                });
            });

            // === 7. Isi Data Sampling (mulai baris 13)
            const startRow = 13;
            currentSampledData.forEach((item, idx) => {
                const row = startRow + idx;

                // Kolom A: No
                worksheet.getCell(`A${row}`).value = idx + 1;

                // Kolom B: Tgl → format dd/mm/yyyy
                worksheet.getCell(`B${row}`).value = new Date(item.tanggal);
                worksheet.getCell(`B${row}`).numFmt = 'dd/mm/yyyy';

                // Kolom C: Voucher
                worksheet.getCell(`C${row}`).value = item.voucher;

                // Kolom D: Nama Transaksi
                worksheet.getCell(`D${row}`).value = item.keterangan;

                // Kolom E: Jumlah Menurut GL
                worksheet.getCell(`E${row}`).value = item.nominal;
                worksheet.getCell(`E${row}`).numFmt = '#,##0';

                // Kolom F–H: DPP, PPN, PPh
                ['F', 'G', 'H'].forEach(col => {
                    worksheet.getCell(`${col}${row}`).value = 0;
                    worksheet.getCell(`${col}${row}`).numFmt = '#,##0';
                });

                // Kolom I: Nett Amount Billed
                worksheet.getCell(`I${row}`).value = { formula: `F${row}+G${row}-H${row}`, result: 0 };
                worksheet.getCell(`I${row}`).numFmt = '#,##0';

                // Kolom J–U: Dokumen pendukung
                worksheet.getCell(`J${row}`).value = '';
                worksheet.getCell(`K${row}`).value = '';
                worksheet.getCell(`K${row}`).numFmt = 'dd/mm/yyyy';
                worksheet.getCell(`L${row}`).value = 0;
                worksheet.getCell(`L${row}`).numFmt = '#,##0';

                worksheet.getCell(`M${row}`).value = '';
                worksheet.getCell(`N${row}`).value = '';
                worksheet.getCell(`N${row}`).numFmt = 'dd/mm/yyyy';
                worksheet.getCell(`O${row}`).value = 0;
                worksheet.getCell(`O${row}`).numFmt = '#,##0';

                worksheet.getCell(`P${row}`).value = '';
                worksheet.getCell(`Q${row}`).value = '';
                worksheet.getCell(`Q${row}`).numFmt = 'dd/mm/yyyy';
                worksheet.getCell(`R${row}`).value = 0;
                worksheet.getCell(`R${row}`).numFmt = '#,##0';

                worksheet.getCell(`S${row}`).value = '';
                worksheet.getCell(`T${row}`).value = '';
                worksheet.getCell(`T${row}`).numFmt = 'dd/mm/yyyy';
                worksheet.getCell(`U${row}`).value = 0;
                worksheet.getCell(`U${row}`).numFmt = '#,##0';

                // Kolom V–Y: Asersi
                ['V', 'W', 'X', 'Y'].forEach(col => {
                    worksheet.getCell(`${col}${row}`).value = '';
                    worksheet.getCell(`${col}${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                });

                // Kolom Z: Kelengkapan Dokumen
                worksheet.getCell(`Z${row}`).value = '';
                worksheet.getCell(`Z${row}`).alignment = { vertical: 'middle', horizontal: 'center' };

                // Kolom AA: Selisih
                worksheet.getCell(`AA${row}`).value = { formula: `I${row}-U${row}`, result: 0 };
                worksheet.getCell(`AA${row}`).numFmt = '#,##0';

                // Kolom AB–AD: Ket, Otorisasi, Deskripsi Temuan
                worksheet.getCell(`AB${row}`).value = '';
                worksheet.getCell(`AC${row}`).value = '';
                worksheet.getCell(`AD${row}`).value = '';

                // Kolom AE: File Eksternal (baru)
                worksheet.getCell(`AE${row}`).value = '';
                worksheet.getCell(`AE${row}`).alignment = { vertical: 'middle', horizontal: 'center' };

                // Border semua kolom A–AE
                const allCols = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE'];
                allCols.forEach(col => {
                    const cell = worksheet.getCell(`${col}${row}`);
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                    if (!['D', 'AD', 'AE'].includes(col)) {
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    }
                });
            });

            const lastDataRow = startRow + currentSampledData.length - 1;

            // === 7b. Dropdown: Asersi (V–Y) dan Kelengkapan (Z)
            if (currentSampledData.length > 0) {
                const firstDataRow = 13;
                const addDropdown = (colLetter, options) => {
                    for (let row = firstDataRow; row <= lastDataRow; row++) {
                        worksheet.getCell(`${colLetter}${row}`).dataValidation = {
                            type: 'list',
                            allowBlank: true,
                            formulae: [`"${options}"`],
                            showErrorMessage: true,
                            errorTitle: 'Input Tidak Valid',
                            error: 'Silakan pilih dari daftar yang tersedia.'
                        };
                    }
                };

                ['V', 'W', 'X', 'Y'].forEach(col => addDropdown(col, 'V,X'));
                addDropdown('Z', 'Lengkap,Tidak Lengkap');
            }

            // === 8. Keterangan
            const keteranganRow = lastDataRow + 3;
            worksheet.getCell(`A${keteranganRow}`).value = 'Keterangan :';
            worksheet.getCell(`A${keteranganRow}`).font = { underline: true, bold: true };
            worksheet.getCell(`A${keteranganRow}`).alignment = { vertical: 'middle' };

            // === 9. Daftar Kode
            const kodeList = ['PBC', 'AFR', 'V', '‹', 'G/L', 'N/A', 'Sp', 'Im', 'Ts'];
            const penjelasanList = [
                'Disiapkan Oleh Klien',
                'Sesuai Dengan Laporan Keuangan Perusahaan',
                'Voucher',
                'Penjumlahan Kebawah dan Kesamping Adalah Benar',
                'Sesuai Dengan Buku Besar Perusahaan',
                'Tidak Terdapat (Non Aplikable)',
                'Error Material / Beda Material',
                'Beda Tidak Material',
                'Tidak Selisih'
            ];

            const startKodeRow = lastDataRow + 4;
            for (let i = 0; i < kodeList.length; i++) {
                const r = startKodeRow + i;
                worksheet.getCell(`B${r}`).value = kodeList[i];
                worksheet.getCell(`C${r}`).value = penjelasanList[i];
                // Border dihapus — tidak ada penugasan `.border`
            }

            // === 10. Simpulan (sekarang merge sampai AE)
            const simpulanRow = startKodeRow + kodeList.length + 2;
            worksheet.getCell(`A${simpulanRow}`).value = 'Simpulan :';
            worksheet.getCell(`A${simpulanRow}`).font = { bold: true };
            worksheet.mergeCells(`B${simpulanRow}:AE${simpulanRow + 2}`); // diperluas ke AE
            const simpulanCell = worksheet.getCell(`B${simpulanRow}`);
            simpulanCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5E6D3' } };
            simpulanCell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };

            // === 11. Simpan file
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Uji_Detil_Sampling.xlsx';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

        } catch (error) {
            console.error('Error:', error);
            alert('Gagal ekspor: ' + (error.message || 'Error tidak dikenal'));
        }
    }

// Auto-update sampling method in summary when changed
const samplingMethodElement = safeGetElement('samplingMethod');
if (samplingMethodElement) {
    samplingMethodElement.addEventListener('change', function() {
        const methodLabels = {
            'rendah': 'Rendah (Top 10)',
            'moderate': 'Moderate (Top 10 + 10 Acak)',
            'tinggi': 'Tinggi (Top 15 + 15 Acak)'
        };
        safeSetTextContent('summaryMetodeSampling', methodLabels[this.value] || methodLabels['moderate']);
    });
}

//Membuat Alert konfirmasi sebelum mulai sampling
function showConfirmSamplingModal() {
    const confirmModal = new bootstrap.Modal(document.getElementById('confirmSamplingModal'));
    confirmModal.show();

    // Tambahkan event listener untuk tombol "Ya, Mulai Sekarang"
    document.getElementById('confirmProceedBtn').onclick = function() {
        confirmModal.hide();
        processData(); // Jalankan proses sampling
    };
}

async function exportToExcelBLUD() {
    if (currentSampledData.length === 0) {
        alert('Tidak ada data sampling untuk diekspor!');
        return;
    }

    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Uji Detil');

        // === 1. Header besar A1:E5
        worksheet.mergeCells('A1:E5');
        const titleCell = worksheet.getCell('A1');
        titleCell.value = currentNamaHeader || 'Sampling Results';
        titleCell.font = { bold: true, size: 16 };
        titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
        titleCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // === 2. UJI DETIL DOKUMEN (F1:J5)
        worksheet.mergeCells('F1:J5');
        worksheet.getCell('F1').value = 'UJI DETIL DOKUMEN (Test Of Detail)';
        worksheet.getCell('F1').font = { bold: true, size: 14 };
        worksheet.getCell('F1').alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getCell('F1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6E6E6' } };
        worksheet.getCell('F1').border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // === 3. INDEKS (K1:O1 dan K2:O5)
        worksheet.mergeCells('K1:O1');
        worksheet.getCell('K1').value = 'INDEKS :';
        worksheet.getCell('K1').font = { bold: true };
        worksheet.getCell('K1').alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getCell('K1').border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('K2:O5');
        worksheet.getCell('K2').value = 'XX';
        worksheet.getCell('K2').alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getCell('K2').border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // === 4. Info Klien & Periode (A7–B8)
        worksheet.getCell('A7').value = 'Klien :';
        worksheet.getCell('A7').font = { bold: true };
        worksheet.getCell('B7').value = currentAuditInfo.namaKlien || '';

        worksheet.getCell('A8').value = 'Periode :';
        worksheet.getCell('A8').font = { bold: true };
        worksheet.getCell('B8').value = currentAuditInfo.schedule || '';

        // === 5. Auditor & Reviewer (K7–O9)
        worksheet.getCell('K7').value = 'Dibuat oleh :';
        worksheet.getCell('K7').font = { bold: true };
        worksheet.getCell('L7').value = currentAuditInfo.dibuatOleh || '';

        worksheet.getCell('K8').value = 'Tanggal :';
        worksheet.getCell('K8').font = { bold: true };
        worksheet.getCell('L8').value = currentAuditInfo.tanggalDibuat || '';

        worksheet.getCell('K9').value = 'Paraf :';
        worksheet.getCell('K9').font = { bold: true };

        worksheet.getCell('N7').value = 'Direview oleh :';
        worksheet.getCell('N7').font = { bold: true };
        worksheet.getCell('O7').value = currentAuditInfo.direviewOleh || '';

        worksheet.getCell('N8').value = 'Tanggal :';
        worksheet.getCell('N8').font = { bold: true };
        worksheet.getCell('O8').value = currentAuditInfo.tanggalDireview || '';

        worksheet.getCell('N9').value = 'Paraf :';
        worksheet.getCell('N9').font = { bold: true };

        // === 6. Header Tabel Baru (baris 11–12)
        const headerConfigs = [
            { range: 'A11:A12', text: 'No', width: 8 },
            { range: 'B11:B12', text: 'Tgl', width: 12 },
            { range: 'C11:C12', text: 'Nomor Voucher', width: 18 },
            { range: 'D11:D12', text: 'Nama Transaksi', width: 25 },
            { range: 'E11:F11', text: 'Jurnal' },
            { cell: 'E12', text: 'D', width: 10 },
            { cell: 'F12', text: 'K', width: 10 },
            { range: 'G11:G12', text: 'Jumlah Menurut GL', width: 18 },
            { range: 'H11:J11', text: 'Bukti Bayar (BB)' },
            { cell: 'H12', text: 'No Bukti', width: 16 },
            { cell: 'I12', text: 'Tgl Bukti', width: 12 },
            { cell: 'J12', text: 'Nominal Bukti', width: 16 },
            { range: 'K11:K12', text: 'Selisih', width: 14 },
            { range: 'L11:L12', text: 'Ket', width: 14 },
            { range: 'M11:M12', text: 'Otorisasi dan Pejabat Otorisasi', width: 22 },
            { range: 'N11:N12', text: 'Deskripsi Temuan', width: 28 },
            { range: 'O11:O12', text: 'File Eksternal', width: 20 }
        ];

        headerConfigs.forEach(item => {
            if (item.range) {
                worksheet.mergeCells(item.range);
                const cell = worksheet.getCell(item.range.split(':')[0]);
                cell.value = item.text;
                cell.font = { bold: true };
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                const colLetter = item.range.split(':')[0].replace(/[0-9]/g, '');
                if (item.width) worksheet.getColumn(colLetter).width = item.width;
            }
            if (item.cell) {
                const cell = worksheet.getCell(item.cell);
                cell.value = item.text;
                cell.font = { bold: true };
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
                const colLetter = item.cell.replace(/[0-9]/g, '');
                if (item.width) worksheet.getColumn(colLetter).width = item.width;
            }
        });

        // === Border dan background header (baris 11–12)
        [11, 12].forEach(rowNum => {
            const row = worksheet.getRow(rowNum);
            row.eachCell({ includeEmpty: true }, cell => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                if (!cell.isMerged) {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
                }
            });
        });

        // === 7. Isi Data Sampling (mulai baris 13)
        const startRow = 13;
        currentSampledData.forEach((item, idx) => {
            const row = startRow + idx;

            worksheet.getCell(`A${row}`).value = idx + 1;
            worksheet.getCell(`B${row}`).value = new Date(item.tanggal);
            worksheet.getCell(`B${row}`).numFmt = 'dd/mm/yyyy';
            worksheet.getCell(`C${row}`).value = item.voucher;
            worksheet.getCell(`D${row}`).value = item.keterangan;
            worksheet.getCell(`E${row}`).value = 0;
            worksheet.getCell(`F${row}`).value = 0;
            worksheet.getCell(`E${row}`).numFmt = '#,##0';
            worksheet.getCell(`F${row}`).numFmt = '#,##0';
            worksheet.getCell(`G${row}`).value = item.nominal;
            worksheet.getCell(`G${row}`).numFmt = '#,##0';
            worksheet.getCell(`H${row}`).value = '';
            worksheet.getCell(`I${row}`).value = '';
            worksheet.getCell(`I${row}`).numFmt = 'dd/mm/yyyy';
            worksheet.getCell(`J${row}`).value = 0;
            worksheet.getCell(`J${row}`).numFmt = '#,##0';
            worksheet.getCell(`K${row}`).value = { formula: `G${row}-J${row}`, result: item.nominal };
            worksheet.getCell(`K${row}`).numFmt = '#,##0';
            worksheet.getCell(`L${row}`).value = '';
            worksheet.getCell(`M${row}`).value = '';
            worksheet.getCell(`N${row}`).value = '';
            worksheet.getCell(`O${row}`).value = '';

            const cols = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O'];
            cols.forEach(col => {
                const cell = worksheet.getCell(`${col}${row}`);
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                if (!['D', 'N', 'O'].includes(col)) {
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                }
            });
        });

        const lastDataRow = startRow + currentSampledData.length - 1;

        // === 8. Buat sheet Lists untuk dropdown "Ket"
        const listSheet = workbook.addWorksheet('Lists');
        listSheet.state = 'hidden';
        listSheet.getCell('A1').value = 'Ts';
        listSheet.getCell('A2').value = 'Im';
        listSheet.getCell('A3').value = 'Sp';
        listSheet.getColumn('A').width = 12;

        // === Dropdown di kolom "Ket" (L) → Ts, Im, Sp sebagai 3 opsi terpisah
        if (currentSampledData.length > 0) {
            for (let r = startRow; r <= lastDataRow; r++) {
                worksheet.getCell(`L${r}`).dataValidation = {
                    type: 'list',
                    allowBlank: true,
                    formulae: ['Lists!$A$1:$A$3'],
                    showErrorMessage: true,
                    errorTitle: 'Input Tidak Valid',
                    error: 'Silakan pilih dari daftar: Ts, Im, atau Sp.'
                };
                worksheet.getCell(`L${r}`).alignment = { vertical: 'middle', horizontal: 'center' };
            }
        }

        // === 9. Keterangan & Kode
        const keteranganRow = lastDataRow + 3;
        worksheet.getCell(`A${keteranganRow}`).value = 'Keterangan :';
        worksheet.getCell(`A${keteranganRow}`).font = { underline: true, bold: true };

        const kodeList = ['PBC', 'AFR', 'V', '‹', 'G/L', 'N/A', 'Sp', 'Im', 'Ts'];
        const penjelasanList = [
            'Disiapkan Oleh Klien',
            'Sesuai Dengan Laporan Keuangan Perusahaan',
            'Voucher',
            'Penjumlahan Kebawah dan Kesamping Adalah Benar',
            'Sesuai Dengan Buku Besar Perusahaan',
            'Tidak Terdapat (Non Aplikable)',
            'Error Material / Beda Material',
            'Beda Tidak Material',
            'Tidak Selisih'
        ];

        const startKodeRow = lastDataRow + 4;
        for (let i = 0; i < kodeList.length; i++) {
            const r = startKodeRow + i;
            worksheet.getCell(`B${r}`).value = kodeList[i];
            worksheet.getCell(`C${r}`).value = penjelasanList[i];
        }

        // === 10. Simpulan
        const simpulanRow = startKodeRow + kodeList.length + 2;
        worksheet.getCell(`A${simpulanRow}`).value = 'Simpulan :';
        worksheet.getCell(`A${simpulanRow}`).font = { bold: true };
        worksheet.mergeCells(`B${simpulanRow}:O${simpulanRow + 2}`);
        const simpulanCell = worksheet.getCell(`B${simpulanRow}`);
        simpulanCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5E6D3' } };
        simpulanCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        // === 11. Simpan file
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Uji_Detil_Sampling.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

    } catch (error) {
        console.error('Error:', error);
        alert('Gagal ekspor: ' + (error.message || 'Error tidak dikenal'));
    }
}

// Dark Mode Toggle
document.addEventListener('DOMContentLoaded', () => {
    const toggleButton = document.getElementById('darkModeToggle');
    if (!toggleButton) return;

    // Cek preferensi dari localStorage atau sistem
    const savedTheme = localStorage.getItem('theme');
    const systemPrefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;

    let currentTheme = 'light';
    if (savedTheme === 'dark' || (!savedTheme && systemPrefersDark)) {
        currentTheme = 'dark';
    }

    // Terapkan tema
    document.documentElement.setAttribute('data-theme', currentTheme);
    updateToggleIcon(currentTheme);

    // Event listener toggle
    toggleButton.addEventListener('click', () => {
        currentTheme = currentTheme === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', currentTheme);
        localStorage.setItem('theme', currentTheme);
        updateToggleIcon(currentTheme);
    });

    function updateToggleIcon(theme) {
        const icon = toggleButton.querySelector('i');
        if (theme === 'dark') {
            icon.className = 'fas fa-sun';
        } else {
            icon.className = 'fas fa-moon';
        }
    }
});