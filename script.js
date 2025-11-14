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
        errorDetails.style.display = 'block';
        if (errorRows) {
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
            const columnNames = {
                'tanggal': 'Tanggal',
                'voucher': 'Voucher', 
                'keterangan': 'Keterangan',
                'nominal': 'Nominal'
            };
            const maxColumns = [];
            const minColumns = [];
            for (const [key, count] of Object.entries(counts)) {
                const columnName = columnNames[key] || key;
                if (count === maxCount) maxColumns.push(columnName);
                if (count === minCount) minColumns.push(columnName);
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
    if (/^\d{4}-\d{2}-\d{2}$/.test(input)) {
        return input;
    }
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
    return input;
}

// Enhanced nominal parsing function
function parseNominal(nominalStr) {
    if (!nominalStr) return 0;
    let cleanStr = nominalStr.trim().replace(/[^0-9.,-]/g, '');
    if (cleanStr.includes(',') && cleanStr.lastIndexOf(',') > cleanStr.lastIndexOf('.')) {
        cleanStr = cleanStr.replace(/\./g, '').replace(',', '.');
    } else if (cleanStr.includes('.') && cleanStr.lastIndexOf('.') > cleanStr.lastIndexOf(',')) {
        cleanStr = cleanStr.replace(/,/g, '');
    } else {
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
    const today = new Date();
    const oneYearAgo = new Date();
    oneYearAgo.setFullYear(today.getFullYear() - 1);
    for (let i = 0; i < count; i++) {
        const randomDate = new Date(oneYearAgo.getTime() + Math.random() * (today.getTime() - oneYearAgo.getTime()));
        const day = String(randomDate.getDate()).padStart(2, '0');
        const month = String(randomDate.getMonth() + 1).padStart(2, '0');
        const year = randomDate.getFullYear();
        tanggalData.push(`${day}/${month}/${year}`);
        voucherData.push(`VOU-${String(i + 1).padStart(3, '0')}`);
        const keterangan = keteranganList[Math.floor(Math.random() * keteranganList.length)];
        keteranganData.push(keterangan);
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
    safeGetElement('namaKlien').value = 'PT Contoh Perusahaan';
    safeGetElement('namaHeader').value = 'Biaya Operasional';
    safeGetElement('namaAkun').value = 'Akun 12345';
    safeGetElement('dibuatOleh').value = 'Auditor Dummy';
    safeGetElement('direviewOleh').value = 'Reviewer Dummy';
    safeGetElement('schedule').value = 'Audit Q4 2025';

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
    document.getElementById('successModal').addEventListener('hidden.bs.modal', function() {
        document.body.removeChild(successModal);
    });
}

function validateDataColumns() {
    const tanggalText = safeGetElement('tanggalInput').value.trim();
    const voucherText = safeGetElement('voucherInput').value.trim();
    const keteranganText = safeGetElement('keteranganInput').value.trim();
    const nominalText = safeGetElement('nominalInput').value.trim();

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
    const sortedData = [...data].sort((a, b) => b.nominal - a.nominal);
    let sampledData = [];
    switch(method) {
        case 'rendah':
            sampledData = sortedData.slice(0, Math.min(10, sortedData.length));
            break;
        case 'moderate':
            const top10 = sortedData.slice(0, Math.min(10, sortedData.length));
            const remainingData = sortedData.length > 10 ? sortedData.slice(10) : [];
            const random10 = remainingData.length > 0 ? getRandomItems(remainingData, 10) : [];
            sampledData = [...top10, ...random10];
            break;
        case 'tinggi':
            const top15 = sortedData.slice(0, Math.min(15, sortedData.length));
            const remainingDataHigh = sortedData.length > 15 ? sortedData.slice(15) : [];
            const random15 = remainingDataHigh.length > 0 ? getRandomItems(remainingDataHigh, 15) : [];
            sampledData = [...top15, ...random15];
            break;
        default:
            sampledData = sortedData.slice(0, Math.min(10, sortedData.length));
    }
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
    const sortedData = data.sort((a, b) => b.nominal - a.nominal);
    tbody.innerHTML = sortedData.map(item => `
        <tr>
            <td>${item.tanggal}</td>
            <td><code>${item.voucher}</code></td>
            <td>${item.keterangan}</td>
            <td class="text-end fw-medium">${formatCurrency(item.nominal)}</td>
        </tr>
    `).join('');
    const exportBtn = safeGetElement('exportBtn');
    if (exportBtn) exportBtn.disabled = false;
}

function updateSummaryPanels(totalData, totalNominal, method, sampleCount, sampledData, originalData) {
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
    safeSetTextContent('summaryTotalData', totalData.toLocaleString('id-ID'));
    safeSetTextContent('summaryTotalNominal', formatCurrency(totalNominal));
    safeSetTextContent('summaryMetodeSampling', methodLabel);
    safeSetTextContent('summaryJumlahSampel', sampleCount.toLocaleString('id-ID'));
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
            showErrorModal(error.message, error.counts);
        } else {
            showErrorModal(error.message);
        }
    }
}

// === FUNGSI UTAMA EKSPOR YANG DIPERBAIKI ===

// Helper: ubah struktur data untuk Excel
function prepareSamplingDataForExport(sampledData) {
    return sampledData.map(item => ({
        Tanggal: item.tanggal,
        Voucher: item.voucher,
        Keterangan: item.keterangan,
        Nominal: item.nominal
    }));
}

// Fungsi ekspor ke Excel
async function exportToExcel(samplingData, metadata) {
    if (!samplingData || samplingData.length === 0) {
        alert("Tidak ada data sampling untuk diekspor.");
        return;
    }

    const {
        namaAkun = "N/A",
        namaKlien = "Nama Klien Belum Diisi",
        dibuatOleh = "",
        direviewOleh = "",
        tanggal = new Date()
    } = metadata;

    try {
        const response = await fetch('./Template_TOD.xlsx');
        if (!response.ok) throw new Error(
            'Gagal memuat Template_TOD.xlsx.\n' +
            'Pastikan file ada di folder yang sama dan web dijalankan via server (bukan file://).'
        );

        const arrayBuffer = await response.arrayBuffer();
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);

        const worksheet = workbook.getWorksheet('result');
        if (!worksheet) throw new Error("Sheet 'result' tidak ditemukan di template.");

        const tanggalStr = new Date(tanggal).toLocaleDateString('id-ID', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        });

        // Isi metadata
        worksheet.getCell('B7').value = namaKlien;
        worksheet.getCell('AB7').value = dibuatOleh;
        worksheet.getCell('AB8').value = tanggalStr;
        worksheet.getCell('AE7').value = direviewOleh;
        worksheet.getCell('AE8').value = tanggalStr;

        // Isi data sampling mulai baris 13
        const startRow = 13;
        for (let i = 0; i < samplingData.length; i++) {
            const rowNumber = startRow + i;
            const newRow = worksheet.getRow(rowNumber);

            newRow.getCell('B').value = namaAkun;                          // COA
            newRow.getCell('C').value = samplingData[i].Tanggal || '';     // Tanggal
            newRow.getCell('D').value = samplingData[i].Voucher || '';     // Voucher
            newRow.getCell('E').value = samplingData[i].Keterangan || '';  // Keterangan
            newRow.getCell('F').value = samplingData[i].Nominal || 0;      // Nominal

            // Salin style dari baris template (baris 12)
            const templateRow = worksheet.getRow(startRow - 1);
            ['B', 'C', 'D', 'E', 'F'].forEach(col => {
                const cell = newRow.getCell(col);
                const ref = templateRow.getCell(col);
                if (ref) {
                    cell.font = { ...ref.font };
                    cell.fill = { ...ref.fill };
                    cell.border = { ...ref.border };
                    cell.alignment = { ...ref.alignment };
                }
            });
            newRow.commit();
        }

        // Unduh file
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Hasil_Sampling_TOD.xlsx';
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 100);

    } catch (error) {
        console.error('Error export Excel:', error);
        alert('❌ Gagal mengekspor:\n' + error.message);
    }
}

// === EVENT HANDLER UNTUK TOMBOL EKSPOR ===
document.addEventListener('DOMContentLoaded', () => {
    const exportBtn = safeGetElement('exportBtn');
    if (exportBtn) {
        exportBtn.addEventListener('click', function() {
            if (!currentSampledData || currentSampledData.length === 0) {
                alert('❌ Belum ada data hasil sampling!\nSilakan klik "Proses Sampling" terlebih dahulu.');
                return;
            }

            const metadata = {
                namaAkun: currentNamaAkun || 'N/A',
                namaKlien: currentAuditInfo.namaKlien || 'N/A',
                dibuatOleh: currentAuditInfo.dibuatOleh || '',
                direviewOleh: currentAuditInfo.direviewOleh || '',
                tanggal: currentAuditInfo.tanggalDibuat || new Date()
            };

            const exportData = prepareSamplingDataForExport(currentSampledData);
            exportToExcel(exportData, metadata);
        });
    }
});

// Auto-update sampling method in summary
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

//fungsi alert konfirmasi sampling
function showConfirmSamplingModal() {
    const confirmModal = new bootstrap.Modal(document.getElementById('confirmSamplingModal'));
    confirmModal.show();

    // Tambahkan event listener untuk tombol "Ya, Mulai Sekarang"
    document.getElementById('confirmProceedBtn').onclick = function() {
        confirmModal.hide();
        processData(); // Jalankan proses sampling
    };
}