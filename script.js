// Mendapatkan elemen DOM
const excelFileInput = document.getElementById('excelFile');
const sheetSelect = document.getElementById('sheetSelect');
const searchInput = document.getElementById('searchInput');
const tableContainer = document.getElementById('table-container');
const loadingMessage = document.getElementById('loadingMessage');
const controlsDiv = document.getElementById('controls');
const dropZone = document.getElementById('dropZone');
const exportCsvBtn = document.getElementById('exportCsvBtn');
const exportPdfBtn = document.getElementById('exportPdfBtn'); // Tombol Ekspor PDF baru
const dataSummaryDiv = document.getElementById('dataSummary');
const summaryContentDiv = document.getElementById('summaryContent');
const manageColumnsBtn = document.getElementById('manageColumnsBtn');
const columnManagerModal = document.getElementById('columnManagerModal');
const columnCheckboxesDiv = document.getElementById('columnCheckboxes');
const applyColumnChangesBtn = document.getElementById('applyColumnChangesBtn');
const cancelColumnChangesBtn = document.getElementById('cancelColumnChangesBtn');
const resetAppBtn = document.getElementById('resetAppBtn');

// Elemen Paginasi
const paginationControls = document.getElementById('paginationControls');
const prevPageBtn = document.getElementById('prevPageBtn');
const nextPageBtn = document.getElementById('nextPageBtn');
const pageInfo = document.getElementById('pageInfo');
const rowsPerPageSelect = document.getElementById('rowsPerPageSelect');

// Elemen Chart
const chartControlsDiv = document.getElementById('chartControls');
const chartTypeSelect = document.getElementById('chartType');
const xAxisSelect = document.getElementById('xAxisSelect');
const yAxisSelect = document.getElementById('yAxisSelect');
const myChartCanvas = document.getElementById('myChart');
let chartInstance = null; // Variabel untuk menyimpan instance Chart.js

// Variabel global untuk menyimpan data
let currentWorkbook = null;
let currentSheetData = []; // Data mentah dari sheet yang dipilih (tanpa header)
let currentHeaders = [];
let filteredData = []; // Data yang sudah difilter
let sortColumnIndex = -1; // Indeks kolom yang sedang diurutkan
let sortDirection = 'asc'; // Arah pengurutan: 'asc' atau 'desc'
let columnDataTypes = {}; // Menyimpan tipe data yang terdeteksi per kolom
let columnFilters = {}; // Menyimpan filter aktif per kolom
let visibleColumns = []; // Menyimpan kolom yang saat ini terlihat

// Variabel Paginasi
let currentPage = 1;
let rowsPerPage = parseInt(rowsPerPageSelect.value);
const defaultRowsPerPage = 10; // Nilai default untuk baris per halaman

// --- Event Listeners ---
excelFileInput.addEventListener('change', (e) => handleFile(e.target.files[0]));
sheetSelect.addEventListener('change', (event) => {
    loadSheetData(event.target.value);
    searchInput.value = ''; // Reset pencarian saat sheet berubah
    currentPage = 1; // Reset halaman ke 1
    savePreferences(); // Simpan preferensi sheet
});
searchInput.addEventListener('input', () => {
    applyAllFilters(); // Panggil fungsi filter utama
    currentPage = 1; // Reset halaman ke 1 setelah pencarian
});
exportCsvBtn.addEventListener('click', exportTableToCsv);
exportPdfBtn.addEventListener('click', exportTableToPdf); // Event listener untuk tombol Ekspor PDF baru
manageColumnsBtn.addEventListener('click', openColumnManagerModal);
applyColumnChangesBtn.addEventListener('click', applyColumnVisibilityChanges);
cancelColumnChangesBtn.addEventListener('click', () => columnManagerModal.classList.add('hidden'));
resetAppBtn.addEventListener('click', resetAll);

// Event Listeners Paginasi
prevPageBtn.addEventListener('click', () => {
    if (currentPage > 1) {
        currentPage--;
        renderTable(filteredData, currentHeaders);
    }
});
nextPageBtn.addEventListener('click', () => {
    const totalPages = Math.ceil(filteredData.length / rowsPerPage);
    if (currentPage < totalPages) {
        currentPage++;
        renderTable(filteredData, currentHeaders);
    }
});
rowsPerPageSelect.addEventListener('change', (event) => {
    rowsPerPage = parseInt(event.target.value);
    currentPage = 1; // Reset halaman ke 1 saat mengganti jumlah baris per halaman
    renderTable(filteredData, currentHeaders);
    savePreferences(); // Simpan preferensi rows per page
});

// Event Listeners Chart
chartTypeSelect.addEventListener('change', renderChart);
xAxisSelect.addEventListener('change', renderChart);
yAxisSelect.addEventListener('change', renderChart);

// --- Drag and Drop Event Listeners ---
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault(); // Mencegah perilaku default (membuka file di browser)
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault(); // Mencegah perilaku default
    dropZone.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
});

// --- Inisialisasi: Muat preferensi saat halaman dimuat ---
document.addEventListener('DOMContentLoaded', loadPreferences);

/**
 * Fungsi untuk menangani pemilihan file Excel (baik dari input maupun drag-drop).
 * @param {File} file - Objek File yang dipilih.
 */
function handleFile(file) {
    if (!file) {
        tableContainer.innerHTML = '<p class="text-center text-gray-600">Tidak ada file yang dipilih.</p>';
        controlsDiv.classList.add('hidden');
        paginationControls.classList.add('hidden');
        dataSummaryDiv.classList.add('hidden'); // Sembunyikan ringkasan data
        chartControlsDiv.classList.add('hidden'); // Sembunyikan kontrol chart
        return;
    }

    loadingMessage.style.display = 'block';
    tableContainer.innerHTML = '';
    controlsDiv.classList.add('hidden');
    paginationControls.classList.add('hidden');
    dataSummaryDiv.classList.add('hidden'); // Sembunyikan ringkasan data
    chartControlsDiv.classList.add('hidden'); // Sembunyikan kontrol chart

    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            currentWorkbook = XLSX.read(data, { type: 'array' });

            loadingMessage.style.display = 'none';
            populateSheetSelect(currentWorkbook.SheetNames);

            // Coba muat sheet yang terakhir dilihat dari preferensi
            const savedSheet = localStorage.getItem('lastSheet');
            if (savedSheet && currentWorkbook.SheetNames.includes(savedSheet)) {
                sheetSelect.value = savedSheet;
                loadSheetData(savedSheet);
            } else {
                loadSheetData(currentWorkbook.SheetNames[0]);
            }

            controlsDiv.classList.remove('hidden');
            paginationControls.classList.remove('hidden');
            dataSummaryDiv.classList.remove('hidden'); // Tampilkan ringkasan data
            chartControlsDiv.classList.remove('hidden'); // Tampilkan kontrol chart
        } catch (error) {
            loadingMessage.style.display = 'none';
            console.error('Error membaca atau memparsing file Excel:', error);
            tableContainer.innerHTML = `<p class="text-center text-red-500">Terjadi kesalahan saat memproses file: ${error.message}. Pastikan ini adalah file Excel yang valid.</p>`;
            controlsDiv.classList.add('hidden');
            paginationControls.classList.add('hidden');
            dataSummaryDiv.classList.add('hidden'); // Sembunyikan ringkasan data
            chartControlsDiv.classList.add('hidden'); // Sembunyikan kontrol chart
        }
    };

    reader.onerror = function() {
        loadingMessage.style.display = 'none';
        console.error('Error saat membaca file:', reader.error);
        tableContainer.innerHTML = '<p class="text-center text-red-500">Gagal membaca file.</p>';
        controlsDiv.classList.add('hidden');
        paginationControls.classList.add('hidden');
        dataSummaryDiv.classList.add('hidden'); // Sembunyikan ringkasan data
        chartControlsDiv.classList.add('hidden'); // Sembunyikan kontrol chart
    };

    reader.readAsArrayBuffer(file);
}

/**
 * Mempopulasi dropdown pemilihan sheet.
 * @param {string[]} sheetNames - Array berisi nama-nama sheet.
 */
function populateSheetSelect(sheetNames) {
    sheetSelect.innerHTML = '';
    sheetNames.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        sheetSelect.appendChild(option);
    });
}

/**
 * Memuat data dari sheet yang dipilih.
 * @param {string} sheetName - Nama sheet yang akan dimuat.
 */
function loadSheetData(sheetName) {
    if (!currentWorkbook) return;

    const worksheet = currentWorkbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (jsonData.length === 0) {
        tableContainer.innerHTML = '<p class="text-center text-red-500">Sheet ini kosong atau tidak dapat dibaca.</p>';
        currentHeaders = [];
        currentSheetData = [];
        filteredData = [];
        dataSummaryDiv.classList.add('hidden'); // Sembunyikan ringkasan data jika kosong
        chartControlsDiv.classList.add('hidden'); // Sembunyikan kontrol chart jika kosong
        return;
    }

    currentHeaders = jsonData[0];
    currentSheetData = jsonData.slice(1);
    columnFilters = {}; // Reset filters when new sheet loads
    detectColumnTypes(currentSheetData, currentHeaders); // Deteksi tipe data

    // Inisialisasi visibleColumns dengan semua kolom terlihat
    visibleColumns = [...currentHeaders];
    const savedVisibleColumns = localStorage.getItem('visibleColumns');
    if (savedVisibleColumns) {
        try {
            const parsed = JSON.parse(savedVisibleColumns);
            // Pastikan kolom yang disimpan masih ada di header saat ini
            visibleColumns = currentHeaders.filter(header => parsed.includes(header));
            // Tambahkan kembali kolom baru yang mungkin belum ada di preferensi yang disimpan
            currentHeaders.forEach(header => {
                if (!visibleColumns.includes(header)) {
                    visibleColumns.push(header);
                }
            });
        } catch (e) {
            console.error("Error parsing saved visible columns:", e);
            visibleColumns = [...currentHeaders]; // Fallback ke semua kolom
        }
    }


    applyAllFilters(); // Terapkan filter awal (kosong) untuk mengisi filteredData
    sortColumnIndex = -1;
    sortDirection = 'asc';
    currentPage = 1; // Pastikan halaman direset saat data sheet baru dimuat

    updateChartControls(); // Perbarui dropdown chart
    renderChart(); // Render chart awal
    calculateStatistics(filteredData, currentHeaders); // Hitung statistik setelah data dimuat
}

/**
 * Mendeteksi tipe data untuk setiap kolom.
 * @param {Array<Array<any>>} data - Data tabel (tanpa header).
 * @param {string[]} headers - Header kolom.
 */
function detectColumnTypes(data, headers) {
    columnDataTypes = {};
    if (data.length === 0) return;

    headers.forEach((header, colIndex) => {
        let isNumeric = true;
        let isDate = true;
        let hasNonEmpty = false;

        for (let i = 0; i < data.length; i++) {
            const cellValue = data[i][colIndex];
            if (cellValue !== null && cellValue !== undefined && String(cellValue).trim() !== '') {
                hasNonEmpty = true;
                // Cek numerik
                if (isNaN(parseFloat(cellValue))) {
                    isNumeric = false;
                }
                // Cek tanggal (format YYYY-MM-DD atau MM/DD/YYYY)
                if (isNaN(new Date(cellValue).getTime()) || !/\d{4}-\d{2}-\d{2}|\d{2}\/\d{2}\/\d{4}/.test(String(cellValue))) {
                    isDate = false;
                }
            }
        }

        if (!hasNonEmpty) {
            columnDataTypes[header] = 'empty';
        } else if (isNumeric && !isDate) { // Prioritaskan angka daripada tanggal jika keduanya mungkin
            columnDataTypes[header] = 'number';
        } else if (isDate) {
            columnDataTypes[header] = 'date';
        } else {
            columnDataTypes[header] = 'string';
        }
    });
}

/**
 * Merender tabel HTML ke DOM.
 * @param {Array<Array<any>>} dataToDisplay - Data yang akan ditampilkan di tabel (tanpa header).
 * @param {string[]} headers - Array header kolom.
 */
function renderTable(dataToDisplay, headers) {
    tableContainer.innerHTML = '';

    if (dataToDisplay.length === 0 && (searchInput.value !== '' || Object.keys(columnFilters).length > 0)) {
        tableContainer.innerHTML = '<p class="text-center text-gray-600">Tidak ada hasil yang ditemukan untuk pencarian atau filter Anda.</p>';
        paginationControls.classList.add('hidden');
        return;
    } else if (dataToDisplay.length === 0 && searchInput.value === '' && Object.keys(columnFilters).length === 0) {
         tableContainer.innerHTML = '<p class="text-center text-gray-600">Tidak ada data untuk ditampilkan. Silakan unggah file Excel.</p>';
         paginationControls.classList.add('hidden');
         return;
    }

    // Hitung data untuk halaman saat ini
    const startIndex = (currentPage - 1) * rowsPerPage;
    const endIndex = startIndex + rowsPerPage;
    const dataForCurrentPage = dataToDisplay.slice(startIndex, endIndex);

    const table = document.createElement('table');
    table.classList.add('min-w-full', 'bg-white', 'rounded-lg', 'shadow-md');

    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');

    // Membuat header tabel
    const headerRow = document.createElement('tr');
    headers.forEach((headerText, index) => {
        // Hanya render kolom yang terlihat (dari panel kelola kolom)
        if (!visibleColumns.includes(headerText)) {
            return;
        }

        const th = document.createElement('th');
        th.textContent = headerText;
        th.classList.add('py-3', 'px-4', 'border-b', 'border-gray-200', 'bg-gray-50', 'text-left', 'text-xs', 'font-semibold', 'text-gray-700', 'uppercase', 'tracking-wider', 'rounded-tl-lg', 'rounded-tr-lg');
        th.setAttribute('data-column-index', index);

        // Ikon Sort
        const sortIcon = document.createElement('span');
        sortIcon.classList.add('sort-icon');
        if (index === sortColumnIndex) {
            sortIcon.classList.add('active');
            sortIcon.textContent = sortDirection === 'asc' ? '▲' : '▼';
        } else {
            sortIcon.textContent = '◆';
        }
        th.appendChild(sortIcon);
        th.addEventListener('click', (e) => {
            // Hanya sort jika klik bukan pada ikon filter
            if (!e.target.classList.contains('filter-icon')) {
                handleSort(index);
            }
        });

        // Ikon Filter
        const filterIcon = document.createElement('span');
        filterIcon.classList.add('filter-icon');
        filterIcon.textContent = '▼'; // Simbol umum untuk filter
        if (columnFilters[headerText] && Object.keys(columnFilters[headerText]).length > 0) {
            filterIcon.classList.add('active'); // Tanda jika filter aktif
        }
        filterIcon.addEventListener('click', (e) => {
            e.stopPropagation(); // Mencegah event click bubbling ke TH (sort)
            toggleFilterModal(th, headerText, index);
        });
        th.appendChild(filterIcon);

        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Membuat baris data untuk halaman saat ini
    dataForCurrentPage.forEach((rowData, rowIndex) => {
        const dataRow = document.createElement('tr');
        rowData.forEach((cellData, colIndex) => {
            // Hanya render kolom yang terlihat (dari panel kelola kolom)
            if (!visibleColumns.includes(currentHeaders[colIndex])) {
                return;
            }

            const td = document.createElement('td');
            td.textContent = cellData;
            td.classList.add('py-3', 'px-4', 'border-b', 'border-gray-200');
            td.setAttribute('data-row-index', rowIndex + startIndex); // Indeks baris asli
            td.setAttribute('data-col-index', colIndex); // Indeks kolom asli

            // Validasi Data & Penyorotan
            if (!isValidData(cellData, columnDataTypes[currentHeaders[colIndex]])) {
                td.classList.add('invalid-data');
            }

            // Mode Edit Sederhana
            td.contentEditable = "true"; // Membuat sel dapat diedit
            td.addEventListener('blur', (e) => handleCellEdit(e, rowIndex + startIndex, colIndex));
            td.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault(); // Mencegah newline di sel
                    e.target.blur(); // Mengakhiri editing
                }
            });

            dataRow.appendChild(td);
        });
        tbody.appendChild(dataRow);
    });
    table.appendChild(tbody);

    tableContainer.appendChild(table);

    // Perbarui kontrol paginasi
    updatePaginationControls(dataToDisplay.length);
}

/**
 * Memvalidasi data sel berdasarkan tipe kolom yang terdeteksi.
 * @param {*} value - Nilai sel.
 * @param {string} type - Tipe data kolom ('string', 'number', 'date', 'empty').
 * @returns {boolean} True jika data valid, false jika tidak.
 */
function isValidData(value, type) {
    if (value === null || value === undefined || String(value).trim() === '') {
        return true; // Sel kosong selalu valid
    }
    const strValue = String(value).trim();

    switch (type) {
        case 'number':
            return !isNaN(parseFloat(strValue));
        case 'date':
            // Cek format YYYY-MM-DD atau MM/DD/YYYY dan validitas tanggal
            return !isNaN(new Date(strValue).getTime()) && (/\d{4}-\d{2}-\d{2}|\d{2}\/\d{2}\/\d{4}/.test(strValue));
        case 'string':
            return typeof value === 'string';
        default:
            return true; // Tipe lain dianggap valid
    }
}

/**
 * Menangani pengeditan sel.
 * @param {Event} event - Objek event dari 'blur'.
 * @param {number} rowIndex - Indeks baris asli dalam filteredData.
 * @param {number} colIndex - Indeks kolom asli.
 */
function handleCellEdit(event, originalRowIndex, colIndex) {
    const newValue = event.target.textContent;

    // Temukan baris yang diedit di filteredData
    const actualRow = filteredData[originalRowIndex];
    if (actualRow) {
        actualRow[colIndex] = newValue; // Perbarui data di filteredData

        // Perbarui juga di currentSheetData agar perubahan persisten saat filter/sheet berubah
        // Ini memerlukan pencarian baris yang sesuai di currentSheetData
        // Untuk kesederhanaan, kita asumsikan rowIndex di filteredData sama dengan di currentSheetData
        const originalCurrentSheetRowIndex = currentSheetData.findIndex(row => row === actualRow);
        if (originalCurrentSheetRowIndex !== -1) {
             currentSheetData[originalCurrentSheetRowIndex][colIndex] = newValue;
        }


        // Validasi ulang sel setelah diedit
        if (!isValidData(newValue, columnDataTypes[currentHeaders[colIndex]])) {
            event.target.classList.add('invalid-data');
        } else {
            event.target.classList.remove('invalid-data');
        }

        // Perbarui statistik dan chart jika kolom yang diedit adalah numerik
        if (columnDataTypes[currentHeaders[colIndex]] === 'number') {
            calculateStatistics(filteredData, currentHeaders);
            renderChart();
        }
    }
}


/**
 * Toggle dan isi modal filter untuk kolom tertentu.
 * @param {HTMLElement} thElement - Elemen TH yang diklik.
 * @param {string} headerText - Teks header kolom.
 * @param {number} columnIndex - Indeks kolom.
 */
function toggleFilterModal(thElement, headerText, columnIndex) {
    // Hapus modal filter yang sudah ada
    document.querySelectorAll('.filter-modal').forEach(modal => modal.remove());

    const modal = document.createElement('div');
    modal.classList.add('filter-modal');
    thElement.appendChild(modal); // Tambahkan modal sebagai anak dari TH

    // Dapatkan nilai unik untuk kolom ini
    const uniqueValues = [...new Set(currentSheetData.map(row => String(row[columnIndex]).trim()))].sort();

    // Dapatkan filter aktif untuk kolom ini
    const activeFilters = columnFilters[headerText] || {}; // Gunakan objek untuk filter angka

    // Isi modal berdasarkan tipe data kolom
    const columnType = columnDataTypes[headerText];
    if (columnType === 'string') {
        uniqueValues.forEach(value => {
            const label = document.createElement('label');
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = value;
            // Jika activeFilters adalah array string, periksa apakah nilai ada di dalamnya
            checkbox.checked = Array.isArray(activeFilters) ? (activeFilters.includes(value) || activeFilters.length === 0) : true;
            checkbox.classList.add('mr-2');
            label.appendChild(checkbox);
            label.append(value === '' ? '(Kosong)' : value); // Tampilkan (Kosong) untuk string kosong
            modal.appendChild(label);
        });

        // Tombol Select All / Clear All
        const selectAllBtn = document.createElement('button');
        selectAllBtn.textContent = 'Pilih Semua';
        selectAllBtn.classList.add('bg-gray-200', 'hover:bg-gray-300', 'text-gray-800', 'py-1', 'px-2', 'rounded-md', 'text-xs', 'mr-2');
        selectAllBtn.addEventListener('click', () => {
            modal.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
        });
        modal.appendChild(selectAllBtn);

        const clearAllBtn = document.createElement('button');
        clearAllBtn.textContent = 'Bersihkan Pilihan';
        clearAllBtn.classList.add('bg-gray-200', 'hover:bg-gray-300', 'text-gray-800', 'py-1', 'px-2', 'rounded-md', 'text-xs');
        clearAllBtn.addEventListener('click', () => {
            modal.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
        });
        modal.appendChild(clearAllBtn);


    } else if (columnType === 'number') {
        // Placeholder untuk filter angka (min/max)
        modal.innerHTML = `
            <label>Min: <input type="number" id="filterMinVal" class="border rounded px-2 py-1 w-full mb-2" value="${activeFilters.min !== undefined && activeFilters.min !== null ? activeFilters.min : ''}"></label>
            <label>Max: <input type="number" id="filterMaxVal" class="border rounded px-2 py-1 w-full" value="${activeFilters.max !== undefined && activeFilters.max !== null ? activeFilters.max : ''}"></label>
            <p class="text-xs text-gray-500 mt-2">Filter angka belum sepenuhnya diimplementasikan.</p>
        `;
    } else {
        modal.innerHTML = `<p class="text-xs text-gray-500">Filter untuk tipe data ini belum tersedia.</p>`;
    }

    // Tombol Apply dan Clear Filter
    const buttonContainer = document.createElement('div');
    buttonContainer.classList.add('filter-buttons');

    const applyBtn = document.createElement('button');
    applyBtn.textContent = 'Terapkan';
    applyBtn.classList.add('bg-blue-500', 'hover:bg-blue-600', 'text-white', 'py-1', 'px-3', 'rounded-md', 'text-sm');
    applyBtn.addEventListener('click', () => {
        if (columnType === 'string') {
            const selectedValues = Array.from(modal.querySelectorAll('input[type="checkbox"]:checked'))
                                        .map(cb => cb.value);
            if (selectedValues.length === uniqueValues.length) { // Jika semua dipilih, berarti tidak ada filter
                delete columnFilters[headerText];
            } else if (selectedValues.length === 0) { // Jika tidak ada yang dipilih, filter kosong
                columnFilters[headerText] = ['__NO_MATCH__']; // Gunakan nilai yang tidak mungkin cocok
            } else {
                columnFilters[headerText] = selectedValues;
            }
        } else if (columnType === 'number') {
            const minVal = modal.querySelector('#filterMinVal').value;
            const maxVal = modal.querySelector('#filterMaxVal').value;
            const parsedMin = minVal === '' ? null : parseFloat(minVal);
            const parsedMax = maxVal === '' ? null : parseFloat(maxVal);

            if (parsedMin !== null || parsedMax !== null) {
                columnFilters[headerText] = { min: parsedMin, max: parsedMax };
            } else {
                delete columnFilters[headerText];
            }
        }
        applyAllFilters();
        modal.remove(); // Tutup modal setelah menerapkan
    });
    buttonContainer.appendChild(applyBtn);

    const clearFilterBtn = document.createElement('button');
    clearFilterBtn.textContent = 'Bersihkan Filter';
    clearFilterBtn.classList.add('bg-red-500', 'hover:bg-red-600', 'text-white', 'py-1', 'px-3', 'rounded-md', 'text-sm');
    clearFilterBtn.addEventListener('click', () => {
        delete columnFilters[headerText]; // Hapus filter untuk kolom ini
        applyAllFilters();
        modal.remove(); // Tutup modal setelah membersihkan
    });
    buttonContainer.appendChild(clearFilterBtn);

    modal.appendChild(buttonContainer);

    // Klik di luar modal untuk menutupnya
    document.addEventListener('click', function closeFilterModal(event) {
        if (!modal.contains(event.target) && !thElement.contains(event.target)) {
            modal.remove();
            document.removeEventListener('click', closeFilterModal);
        }
    });
}


/**
 * Fungsi untuk menangani pengurutan kolom.
 * @param {number} columnIndex - Indeks kolom yang akan diurutkan.
 */
function handleSort(columnIndex) {
    if (sortColumnIndex === columnIndex) {
        sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        sortColumnIndex = columnIndex;
        sortDirection = 'asc';
    }
    applyAllFilters(); // Panggil fungsi filter utama setelah pengurutan
}

/**
 * Fungsi utama untuk menerapkan semua filter dan pengurutan.
 */
function applyAllFilters() {
    let tempFilteredData = [...currentSheetData];

    // 1. Terapkan Pencarian Global
    const globalQuery = searchInput.value.toLowerCase().trim();
    if (globalQuery !== '') {
        tempFilteredData = tempFilteredData.filter(row =>
            row.some(cell =>
                String(cell).toLowerCase().includes(globalQuery)
            )
        );
    }

    // 2. Terapkan Filter Kolom Spesifik
    for (const header in columnFilters) {
        const filterValues = columnFilters[header];
        const colIndex = currentHeaders.indexOf(header);

        if (colIndex !== -1 && filterValues) {
            const columnType = columnDataTypes[header];
            if (columnType === 'string') {
                if (filterValues[0] === '__NO_MATCH__') {
                    tempFilteredData = [];
                    break;
                }
                tempFilteredData = tempFilteredData.filter(row =>
                    filterValues.includes(String(row[colIndex]).trim())
                );
            } else if (columnType === 'number' && (filterValues.min !== null || filterValues.max !== null)) {
                tempFilteredData = tempFilteredData.filter(row => {
                    const cellNum = parseFloat(row[colIndex]);
                    return (filterValues.min === null || cellNum >= filterValues.min) &&
                           (filterValues.max === null || cellNum <= filterValues.max);
                });
            }
        }
    }

    // 3. Terapkan Pengurutan
    if (sortColumnIndex !== -1) {
        tempFilteredData.sort((a, b) => {
            const valA = a[sortColumnIndex];
            const valB = b[sortColumnIndex];

            if (valA === null || valA === undefined) return sortDirection === 'asc' ? 1 : -1;
            if (valB === null || valB === undefined) return sortDirection === 'asc' ? -1 : 1;

            const numA = parseFloat(valA);
            const numB = parseFloat(valB);

            if (!isNaN(numA) && !isNaN(numB)) {
                return sortDirection === 'asc' ? numA - numB : numB - numA;
            }

            const strA = String(valA).toLowerCase();
            const strB = String(valB).toLowerCase();
            if (strA < strB) return sortDirection === 'asc' ? -1 : 1;
            if (strA > strB) return sortDirection === 'asc' ? 1 : -1;
            return 0;
        });
    }

    filteredData = tempFilteredData;
    currentPage = 1; // Reset halaman ke 1 setelah filter/sort
    renderTable(filteredData, currentHeaders);
    calculateStatistics(filteredData, currentHeaders); // Perbarui statistik
    renderChart(); // Perbarui chart
}

/**
 * Memperbarui kontrol paginasi (nomor halaman, tombol aktif/nonaktif).
 * @param {number} totalRows - Total baris data yang difilter.
 */
function updatePaginationControls(totalRows) {
    const totalPages = Math.ceil(totalRows / rowsPerPage);
    pageInfo.textContent = `Halaman ${currentPage} dari ${totalPages}`;

    prevPageBtn.disabled = currentPage === 1;
    nextPageBtn.disabled = currentPage === totalPages || totalPages === 0;

    if (totalRows > rowsPerPage) {
        paginationControls.classList.remove('hidden');
    } else {
        paginationControls.classList.add('hidden');
    }
}

/**
 * Mengekspor data tabel yang sedang ditampilkan (difilter/diurutkan) ke file CSV.
 */
function exportTableToCsv() {
    if (filteredData.length === 0 || currentHeaders.length === 0) {
        displayCustomMessage('Tidak ada data untuk diekspor.', 'error');
        return;
    }

    // Hanya ekspor kolom yang terlihat
    const exportHeaders = currentHeaders.filter(header => visibleColumns.includes(header));
    const exportColIndices = exportHeaders.map(header => currentHeaders.indexOf(header));

    let csvContent = exportHeaders.map(header => `"${header}"`).join(',') + '\n'; // Header CSV

    filteredData.forEach(row => {
        const exportRow = exportColIndices.map(colIndex => row[colIndex]);
        const rowString = exportRow.map(cell => {
            const cellStr = String(cell);
            if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
                return `"${cellStr.replace(/"/g, '""')}"`;
            }
            return cellStr;
        }).join(',');
        csvContent += rowString + '\n';
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'data_excel_export.csv';
    link.click();
    URL.revokeObjectURL(link.href);
    displayCustomMessage('Data berhasil diekspor ke CSV!', 'info');
}

/**
 * Mengekspor tabel yang sedang ditampilkan ke format PDF.
 */
async function exportTableToPdf() {
    if (filteredData.length === 0 || currentHeaders.length === 0) {
        displayCustomMessage('Tidak ada data untuk diekspor ke PDF.', 'error');
        return;
    }

    displayCustomMessage('Mempersiapkan PDF...', 'info');

    // Buat elemen div sementara untuk menampung tabel yang akan diekspor
    // Ini penting agar html2canvas mengambil tabel yang sudah difilter/diurutkan
    const tempContainer = document.createElement('div');
    tempContainer.style.position = 'absolute';
    tempContainer.style.left = '-9999px'; // Sembunyikan dari tampilan
    document.body.appendChild(tempContainer);

    // Render tabel ke kontainer sementara
    const tempTable = document.createElement('table');
    tempTable.classList.add('min-w-full', 'bg-white', 'rounded-lg', 'shadow-md'); // Gunakan kelas yang sama untuk styling

    const tempThead = document.createElement('thead');
    const tempTbody = document.createElement('tbody');

    // Buat header tabel untuk PDF (hanya kolom yang terlihat)
    const pdfHeaders = currentHeaders.filter(header => visibleColumns.includes(header));
    const headerRow = document.createElement('tr');
    pdfHeaders.forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText;
        th.style.padding = '8px'; // Styling dasar untuk PDF
        th.style.border = '1px solid #ddd';
        th.style.backgroundColor = '#f2f2f2';
        th.style.textAlign = 'left';
        headerRow.appendChild(th);
    });
    tempThead.appendChild(headerRow);
    tempTable.appendChild(tempThead);

    // Buat baris data untuk PDF (hanya data yang difilter/diurutkan dan kolom yang terlihat)
    filteredData.forEach(rowData => {
        const dataRow = document.createElement('tr');
        pdfHeaders.forEach(headerText => {
            const colIndex = currentHeaders.indexOf(headerText);
            const td = document.createElement('td');
            td.textContent = rowData[colIndex];
            td.style.padding = '8px'; // Styling dasar untuk PDF
            td.style.border = '1px solid #ddd';
            td.style.textAlign = 'left';
            dataRow.appendChild(td);
        });
        tempTbody.appendChild(dataRow);
    });
    tempTable.appendChild(tempTbody);
    tempContainer.appendChild(tempTable);


    try {
        // Gunakan html2canvas untuk mengambil screenshot tabel
        const canvas = await html2canvas(tempTable, {
            scale: 2, // Meningkatkan skala untuk kualitas yang lebih baik
            useCORS: true, // Penting jika ada gambar dari domain lain
            logging: false, // Nonaktifkan logging konsol html2canvas
        });

        // Hapus kontainer sementara setelah canvas dibuat
        document.body.removeChild(tempContainer);

        const imgData = canvas.toDataURL('image/png');
        const pdf = new window.jspdf.jsPDF('p', 'mm', 'a4'); // 'p' for portrait, 'mm' for millimeters, 'a4' size

        const imgWidth = 210; // A4 width in mm
        const pageHeight = 295; // A4 height in mm
        const imgHeight = canvas.height * imgWidth / canvas.width;
        let heightLeft = imgHeight;

        let position = 0;

        pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
        heightLeft -= pageHeight;

        while (heightLeft >= 0) {
            position = heightLeft - imgHeight;
            pdf.addPage();
            pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
            heightLeft -= pageHeight;
        }

        pdf.save('data_excel_export.pdf');
        displayCustomMessage('Data berhasil diekspor ke PDF!', 'info');

    } catch (error) {
        console.error('Error exporting to PDF:', error);
        displayCustomMessage(`Gagal mengekspor ke PDF: ${error.message}`, 'error');
        // Pastikan kontainer sementara dihapus jika terjadi error
        if (document.body.contains(tempContainer)) {
            document.body.removeChild(tempContainer);
        }
    }
}


/**
 * Menampilkan pesan kustom (pengganti alert()).
 * @param {string} message - Pesan yang akan ditampilkan.
 * @param {string} type - Tipe pesan (misal: 'info', 'error').
 */
function displayCustomMessage(message, type = 'info') {
    const messageBox = document.createElement('div');
    messageBox.classList.add('fixed', 'top-4', 'right-4', 'p-4', 'rounded-md', 'shadow-lg', 'text-white', 'z-50');
    if (type === 'error') {
        messageBox.classList.add('bg-red-500');
    } else {
        messageBox.classList.add('bg-blue-500');
    }
    messageBox.textContent = message;
    document.body.appendChild(messageBox);

    setTimeout(() => {
        messageBox.remove();
    }, 3000);
}

/**
 * Menghitung dan menampilkan statistik dasar untuk kolom numerik.
 * @param {Array<Array<any>>} data - Data tabel (tanpa header).
 * @param {string[]} headers - Header kolom.
 */
function calculateStatistics(data, headers) {
    summaryContentDiv.innerHTML = '';

    if (data.length === 0) {
        dataSummaryDiv.classList.add('hidden');
        return;
    }

    const numericColumns = [];
    headers.forEach((header, colIndex) => {
        // Hanya hitung statistik untuk kolom yang terlihat
        if (visibleColumns.includes(header) && columnDataTypes[header] === 'number') {
            numericColumns.push({ header, colIndex });
        }
    });

    if (numericColumns.length === 0) {
        summaryContentDiv.innerHTML = '<p>Tidak ada kolom numerik yang terdeteksi untuk ringkasan statistik.</p>';
        return;
    }

    numericColumns.forEach(col => {
        const values = data.map(row => parseFloat(row[col.colIndex])).filter(val => !isNaN(val));

        if (values.length > 0) {
            const sum = values.reduce((acc, val) => acc + val, 0);
            const avg = sum / values.length;
            const min = Math.min(...values);
            const max = Math.max(...values);
            const count = values.length;

            const formatNumber = (num) => {
                return num.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
            };

            const columnSummary = document.createElement('div');
            columnSummary.classList.add('bg-white', 'p-3', 'rounded-md', 'shadow-sm', 'border', 'border-blue-100');
            columnSummary.innerHTML = `
                <h3 class="font-semibold text-blue-700 mb-1">${col.header}</h3>
                <p><strong>Jumlah Data:</strong> ${count}</p>
                <p><strong>Total:</strong> ${formatNumber(sum)}</p>
                <p><strong>Rata-rata:</strong> ${formatNumber(avg)}</p>
                <p><strong>Min:</strong> ${formatNumber(min)}</p>
                <p><strong>Max:</strong> ${formatNumber(max)}</p>
            `;
            summaryContentDiv.appendChild(columnSummary);
        }
    });
}

/**
 * Memperbarui dropdown untuk pemilihan sumbu X dan Y pada grafik.
 */
function updateChartControls() {
    xAxisSelect.innerHTML = '';
    yAxisSelect.innerHTML = '';

    const emptyOption = document.createElement('option');
    emptyOption.value = '';
    emptyOption.textContent = 'Pilih Kolom...';
    xAxisSelect.appendChild(emptyOption.cloneNode(true));
    yAxisSelect.appendChild(emptyOption.cloneNode(true));

    currentHeaders.forEach(header => {
        // Hanya tampilkan kolom yang terlihat di dropdown chart
        if (!visibleColumns.includes(header)) {
            return;
        }

        const optionX = document.createElement('option');
        optionX.value = header;
        optionX.textContent = header;
        xAxisSelect.appendChild(optionX);

        // Hanya tambahkan kolom numerik ke sumbu Y
        if (columnDataTypes[header] === 'number') {
            const optionY = document.createElement('option');
            optionY.value = header;
            optionY.textContent = header;
            yAxisSelect.appendChild(optionY);
        }
    });
}

/**
 * Merender grafik menggunakan Chart.js.
 */
function renderChart() {
    if (chartInstance) {
        chartInstance.destroy(); // Hancurkan instance chart sebelumnya
    }

    const ctx = myChartCanvas.getContext('2d');
    const chartType = chartTypeSelect.value;
    const xAxisHeader = xAxisSelect.value;
    const yAxisHeader = yAxisSelect.value;

    if (!xAxisHeader || !yAxisHeader || filteredData.length === 0) {
        // Sembunyikan chart jika tidak ada data atau kolom yang dipilih
        myChartCanvas.style.display = 'none';
        return;
    } else {
        myChartCanvas.style.display = 'block';
    }

    const xAxisIndex = currentHeaders.indexOf(xAxisHeader);
    const yAxisIndex = currentHeaders.indexOf(yAxisHeader);

    if (xAxisIndex === -1 || yAxisIndex === -1) {
        return; // Kolom tidak ditemukan
    }

    // Agregasi data untuk chart (misal: sum Y-axis values by X-axis category)
    const aggregatedData = {};
    filteredData.forEach(row => {
        const xValue = String(row[xAxisIndex]).trim();
        const yValue = parseFloat(row[yAxisIndex]);

        if (!isNaN(yValue)) {
            if (aggregatedData[xValue]) {
                aggregatedData[xValue] += yValue;
            } else {
                aggregatedData[xValue] = yValue;
            }
        }
    });

    const labels = Object.keys(aggregatedData);
    const dataValues = Object.values(aggregatedData);

    chartInstance = new Chart(ctx, {
        type: chartType,
        data: {
            labels: labels,
            datasets: [{
                label: yAxisHeader,
                data: dataValues,
                backgroundColor: chartType === 'bar' ? 'rgba(75, 192, 192, 0.6)' : 'rgba(153, 102, 255, 0.6)',
                borderColor: chartType === 'bar' ? 'rgba(75, 192, 192, 1)' : 'rgba(153, 102, 255, 1)',
                borderWidth: 1,
                fill: chartType === 'line' ? false : true,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}

/**
 * Membuka modal untuk mengelola visibilitas kolom.
 */
function openColumnManagerModal() {
    columnCheckboxesDiv.innerHTML = ''; // Bersihkan konten sebelumnya

    currentHeaders.forEach(header => {
        const label = document.createElement('label');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = header;
        checkbox.checked = visibleColumns.includes(header); // Cek berdasarkan status visibleColumns
        label.appendChild(checkbox);
        label.append(header);
        columnCheckboxesDiv.appendChild(label);
    });

    columnManagerModal.classList.remove('hidden');
}

/**
 * Menerapkan perubahan visibilitas kolom dari modal.
 */
function applyColumnVisibilityChanges() {
    const newVisibleColumns = [];
    columnCheckboxesDiv.querySelectorAll('input[type="checkbox"]:checked').forEach(checkbox => {
        newVisibleColumns.push(checkbox.value);
    });

    if (newVisibleColumns.length === 0) {
        displayCustomMessage('Setidaknya satu kolom harus terlihat.', 'error');
        return;
    }

    visibleColumns = newVisibleColumns;
    localStorage.setItem('visibleColumns', JSON.stringify(visibleColumns)); // Simpan preferensi
    columnManagerModal.classList.add('hidden');
    applyAllFilters(); // Render ulang tabel dengan kolom yang terlihat
    updateChartControls(); // Perbarui dropdown chart
    renderChart(); // Render ulang chart
    calculateStatistics(filteredData, currentHeaders); // Perbarui statistik
}

/**
 * Menyimpan preferensi pengguna ke localStorage.
 */
function savePreferences() {
    localStorage.setItem('rowsPerPage', rowsPerPageSelect.value);
    localStorage.setItem('lastSheet', sheetSelect.value);
    localStorage.setItem('visibleColumns', JSON.stringify(visibleColumns)); // Simpan kolom yang terlihat
}

/**
 * Memuat preferensi pengguna dari localStorage.
 */
function loadPreferences() {
    const savedRowsPerPage = localStorage.getItem('rowsPerPage');
    const savedSheet = localStorage.getItem('lastSheet');
    const savedVisibleColumns = localStorage.getItem('visibleColumns');

    if (savedRowsPerPage) {
        rowsPerPageSelect.value = savedRowsPerPage;
        rowsPerPage = parseInt(savedRowsPerPage);
    }

    // Sheet dan visibleColumns akan dimuat setelah workbook di-parse di handleFile
    // Logika pemuatan ada di handleFile.
}

/**
 * Mereset semua filter, pengurutan, visibilitas kolom, dan data yang dimuat.
 */
function resetAll() {
    // Konfirmasi reset
    if (!confirm('Apakah Anda yakin ingin mereset semua filter, pengurutan, dan visibilitas kolom? Data yang telah diunggah akan tetap ada, tetapi perubahan sel akan hilang jika Anda mengunggah ulang file yang sama.')) {
        return;
    }

    // Reset variabel internal
    searchInput.value = '';
    columnFilters = {};
    sortColumnIndex = -1;
    sortDirection = 'asc';
    currentPage = 1;
    rowsPerPageSelect.value = defaultRowsPerPage;
    rowsPerPage = defaultRowsPerPage;

    // Reset preferensi di localStorage
    localStorage.removeItem('rowsPerPage');
    localStorage.removeItem('lastSheet');
    localStorage.removeItem('visibleColumns');

    // Hapus semua modal filter yang mungkin terbuka
    document.querySelectorAll('.filter-modal').forEach(modal => modal.remove());

    // Muat ulang data dari currentWorkbook untuk mereset perubahan sel lokal
    if (currentWorkbook) {
        loadSheetData(sheetSelect.value); // Muat ulang sheet yang aktif
        displayCustomMessage('Semua pengaturan telah direset!', 'info');
    } else {
        // Jika tidak ada workbook yang dimuat, bersihkan UI sepenuhnya
        tableContainer.innerHTML = '<p class="text-center text-gray-600">Tidak ada data untuk ditampilkan. Silakan unggah file Excel.</p>';
        controlsDiv.classList.add('hidden');
        paginationControls.classList.add('hidden');
        dataSummaryDiv.classList.add('hidden');
        chartControlsDiv.classList.add('hidden');
        loadingMessage.style.display = 'none';
        displayCustomMessage('Aplikasi telah direset ke kondisi awal.', 'info');
    }
}
