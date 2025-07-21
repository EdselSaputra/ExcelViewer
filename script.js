// Mendapatkan elemen DOM
const excelFileInput = document.getElementById('excelFile');
const sheetSelect = document.getElementById('sheetSelect');
const searchInput = document.getElementById('searchInput');
const tableContainer = document.getElementById('table-container');
const loadingMessage = document.getElementById('loadingMessage');
const controlsDiv = document.getElementById('controls');
const dropZone = document.getElementById('dropZone');
const exportCsvBtn = document.getElementById('exportCsvBtn');
const exportPdfBtn = document.getElementById('exportPdfBtn');
const copyToClipboardBtn = document.getElementById('copyToClipboardBtn'); // Tombol Salin Tabel baru
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
const standardChartAxes = document.getElementById('standardChartAxes');
const yAxisContainer = document.getElementById('yAxisContainer');
const scatterChartAxes = document.getElementById('scatterChartAxes');
const scatterXAxisSelect = document.getElementById('scatterXAxisSelect');
const scatterYAxisSelect = document.getElementById('scatterYAxisSelect');
let chartInstance = null; // Variabel untuk menyimpan instance Chart.js

// Elemen Pengelompokan Data
const groupingControlsDiv = document.getElementById('groupingControls');
const groupBySelect = document.getElementById('groupBySelect');
const aggregateBySelect = document.getElementById('aggregateBySelect');
const groupedSummaryContentDiv = document.getElementById('groupedSummaryContent');


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

// Variabel untuk Drag & Drop Kolom
let draggedTh = null;
let dragOverTh = null;

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
exportPdfBtn.addEventListener('click', exportTableToPdf);
copyToClipboardBtn.addEventListener('click', copyTableToClipboard); // Event listener untuk tombol Salin Tabel
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
chartTypeSelect.addEventListener('change', () => {
    updateChartControlsVisibility();
    renderChart();
});
xAxisSelect.addEventListener('change', renderChart);
yAxisSelect.addEventListener('change', renderChart);
scatterXAxisSelect.addEventListener('change', renderChart);
scatterYAxisSelect.addEventListener('change', renderChart);


// Event Listeners Pengelompokan Data
groupBySelect.addEventListener('change', calculateGroupedSummary);
aggregateBySelect.addEventListener('change', calculateGroupedSummary);


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
        groupingControlsDiv.classList.add('hidden'); // Sembunyikan kontrol pengelompokan
        return;
    }

    loadingMessage.style.display = 'block';
    tableContainer.innerHTML = '';
    controlsDiv.classList.add('hidden');
    paginationControls.classList.add('hidden');
    dataSummaryDiv.classList.add('hidden'); // Sembunyikan ringkasan data
    chartControlsDiv.classList.add('hidden'); // Sembunyikan kontrol chart
    groupingControlsDiv.classList.add('hidden'); // Sembunyikan kontrol pengelompokan

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
            groupingControlsDiv.classList.remove('hidden'); // Tampilkan kontrol pengelompokan
        } catch (error) {
            loadingMessage.style.display = 'none';
            console.error('Error membaca atau memparsing file Excel:', error);
            tableContainer.innerHTML = `<p class="text-center text-red-500">Terjadi kesalahan saat memproses file: ${error.message}. Pastikan ini adalah file Excel yang valid.</p>`;
            controlsDiv.classList.add('hidden');
            paginationControls.classList.add('hidden');
            dataSummaryDiv.classList.add('hidden'); // Sembunyikan ringkasan data
            chartControlsDiv.classList.add('hidden'); // Sembunyikan kontrol chart
            groupingControlsDiv.classList.add('hidden'); // Sembunyikan kontrol pengelompokan
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
        groupingControlsDiv.classList.add('hidden'); // Sembunyikan kontrol pengelompokan
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
        groupingControlsDiv.classList.add('hidden'); // Sembunyikan kontrol pengelompokan jika kosong
        return;
    }

    currentHeaders = jsonData[0].map(header => header || ""); // Ganti header null/undefined dengan string kosong
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
    populateGroupingControls(); // Perbarui dropdown pengelompokan
    renderChart(); // Render chart awal
    calculateStatistics(filteredData, currentHeaders); // Hitung statistik setelah data dimuat
    calculateGroupedSummary(); // Hitung ringkasan pengelompokan awal
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
                if (isNaN(new Date(cellValue).getTime()) || !/(\d{4}-\d{2}-\d{2})|(\d{2}\/\d{2}\/\d{4})/.test(String(cellValue))) {
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
    // Urutkan header berdasarkan visibleColumns
    const orderedHeaders = visibleColumns.filter(header => headers.includes(header));

    orderedHeaders.forEach((headerText, index) => {
        const originalIndex = headers.indexOf(headerText); // Indeks asli di currentHeaders

        const th = document.createElement('th');
        th.textContent = headerText;
        th.classList.add('py-3', 'px-4', 'border-b', 'border-gray-200', 'bg-gray-50', 'text-left', 'text-xs', 'font-semibold', 'text-gray-700', 'uppercase', 'tracking-wider', 'rounded-tl-lg', 'rounded-tr-lg');
        th.setAttribute('data-column-index', originalIndex); // Gunakan indeks asli untuk sort/filter
        th.setAttribute('draggable', 'true'); // Membuat kolom bisa diseret
        th.classList.add('relative'); // Untuk posisi ikon filter/sort

        // Tambahkan kelas sticky untuk kolom pertama yang terlihat
        if (index === 0) {
            th.classList.add('sticky', 'left-0', 'z-10');
            th.style.backgroundColor = '#f8fafc'; // Pastikan warna latar belakang sticky
            th.style.boxShadow = '2px 0 5px rgba(0,0,0,0.1)'; // Shadow untuk efek sticky
        }


        // Ikon Sort
        const sortIcon = document.createElement('span');
        sortIcon.classList.add('sort-icon');
        if (originalIndex === sortColumnIndex) {
            sortIcon.classList.add('active');
            sortIcon.textContent = sortDirection === 'asc' ? '▲' : '▼';
        } else {
            sortIcon.textContent = '◆';
        }
        th.appendChild(sortIcon);
        th.addEventListener('click', (e) => {
            if (!e.target.classList.contains('filter-icon')) {
                handleSort(originalIndex); // Sort berdasarkan indeks asli
            }
        });

        // Ikon Filter
        const filterIcon = document.createElement('span');
        filterIcon.classList.add('filter-icon');
        filterIcon.textContent = '▼';
        if (columnFilters[headerText] && Object.keys(columnFilters[headerText]).length > 0) {
            filterIcon.classList.add('active');
        }
        filterIcon.addEventListener('click', (e) => {
            e.stopPropagation();
            toggleFilterModal(th, headerText, originalIndex); // Filter berdasarkan indeks asli
        });
        th.appendChild(filterIcon);

        // Event listener Drag & Drop
        th.addEventListener('dragstart', handleColumnReorderStart);
        th.addEventListener('dragover', handleColumnReorderOver);
        th.addEventListener('dragleave', handleColumnReorderLeave);
        th.addEventListener('drop', handleColumnReorderDrop);
        th.addEventListener('dragend', handleColumnReorderEnd);


        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Membuat baris data untuk halaman saat ini
    dataForCurrentPage.forEach((rowData, rowIndex) => {
        const dataRow = document.createElement('tr');
        orderedHeaders.forEach((headerText, colIdxInOrderedHeaders) => {
            const originalIndex = headers.indexOf(headerText); // Indeks asli di currentHeaders
            const cellData = rowData[originalIndex];

            const td = document.createElement('td');
            td.textContent = cellData;
            td.classList.add('py-3', 'px-4', 'border-b', 'border-gray-200');
            td.setAttribute('data-row-index', rowIndex + startIndex); // Indeks baris asli
            td.setAttribute('data-col-index', originalIndex); // Indeks kolom asli

            // Tambahkan kelas sticky untuk kolom pertama yang terlihat
            if (colIdxInOrderedHeaders === 0) {
                td.classList.add('sticky', 'left-0', 'z-10');
                td.style.backgroundColor = (rowIndex + startIndex) % 2 === 0 ? '#ffffff' : '#f7fafc'; // Sesuaikan warna latar belakang baris genap/ganjil
                td.style.boxShadow = '2px 0 5px rgba(0,0,0,0.05)'; // Shadow untuk efek sticky
            }

            // Validasi Data & Penyorotan
            if (!isValidData(cellData, columnDataTypes[currentHeaders[originalIndex]])) {
                td.classList.add('invalid-data');
            }

            // Mode Edit Sederhana
            td.contentEditable = "true"; // Membuat sel dapat diedit
            td.addEventListener('blur', (e) => handleCellEdit(e, rowIndex + startIndex, originalIndex));
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
            return typeof value === 'string'; // Seharusnya selalu benar
        default:
            return true; // Tipe lain dianggap valid
    }
}

/**
 * Menangani pengeditan sel.
 * @param {Event} event - Objek event dari 'blur'.
 * @param {number} originalRowIndex - Indeks baris asli dalam filteredData.
 * @param {number} colIndex - Indeks kolom asli.
 */
function handleCellEdit(event, originalRowIndex, colIndex) {
    const newValue = event.target.textContent;

    const actualRow = filteredData[originalRowIndex];
    if (actualRow) {
        const oldValue = actualRow[colIndex];
        actualRow[colIndex] = newValue; // Perbarui data di filteredData

        // Temukan baris yang sesuai di currentSheetData dan perbarui juga
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

        // Perbarui statistik dan chart jika ada perubahan pada data
        calculateStatistics(filteredData, currentHeaders);
        renderChart();
        calculateGroupedSummary();
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
    modal.style.display = 'block'; // Tampilkan modal
    thElement.appendChild(modal);

    const uniqueValues = [...new Set(currentSheetData.map(row => String(row[columnIndex]).trim()))].sort();
    const activeFilters = columnFilters[headerText] || {};
    const columnType = columnDataTypes[headerText];

    if (columnType === 'string' || columnType === 'date') {
        uniqueValues.forEach(value => {
            const label = document.createElement('label');
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = value;
            checkbox.checked = !Array.isArray(activeFilters) || activeFilters.length === 0 || activeFilters.includes(value);
            label.appendChild(checkbox);
            label.append(` ${value === '' ? '(Kosong)' : value}`);
            modal.appendChild(label);
        });
    } else if (columnType === 'number') {
        modal.innerHTML = `
            <label>Min: <input type="number" id="filterMinVal" class="border rounded px-2 py-1 w-full mb-2" value="${activeFilters.min ?? ''}"></label>
            <label>Max: <input type="number" id="filterMaxVal" class="border rounded px-2 py-1 w-full" value="${activeFilters.max ?? ''}"></label>
        `;
    } else {
        modal.innerHTML = `<p class="text-xs text-gray-500">Filter tidak tersedia.</p>`;
    }

    const buttonContainer = document.createElement('div');
    buttonContainer.classList.add('filter-buttons');

    const applyBtn = document.createElement('button');
    applyBtn.textContent = 'Terapkan';
    applyBtn.classList.add('bg-blue-500', 'hover:bg-blue-600', 'text-white', 'py-1', 'px-3', 'rounded-md', 'text-sm');
    applyBtn.addEventListener('click', () => {
        if (columnType === 'string' || columnType === 'date') {
            const selectedValues = Array.from(modal.querySelectorAll('input:checked')).map(cb => cb.value);
            if (selectedValues.length === uniqueValues.length || selectedValues.length === 0) {
                delete columnFilters[headerText];
            } else {
                columnFilters[headerText] = selectedValues;
            }
        } else if (columnType === 'number') {
            const min = modal.querySelector('#filterMinVal').value;
            const max = modal.querySelector('#filterMaxVal').value;
            if (min === '' && max === '') {
                delete columnFilters[headerText];
            } else {
                columnFilters[headerText] = { min: min === '' ? null : parseFloat(min), max: max === '' ? null : parseFloat(max) };
            }
        }
        applyAllFilters();
        modal.remove();
    });
    buttonContainer.appendChild(applyBtn);

    const clearBtn = document.createElement('button');
    clearBtn.textContent = 'Bersihkan';
    clearBtn.classList.add('text-sm');
    clearBtn.addEventListener('click', () => {
        delete columnFilters[headerText];
        applyAllFilters();
        modal.remove();
    });
    buttonContainer.appendChild(clearBtn);
    modal.appendChild(buttonContainer);

    // Event listener untuk menutup modal
    setTimeout(() => {
        document.addEventListener('click', function closeOnClickOutside(event) {
            if (!modal.contains(event.target)) {
                modal.remove();
                document.removeEventListener('click', closeOnClickOutside);
            }
        }, { once: true });
    }, 0);
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
    Object.keys(columnFilters).forEach(header => {
        const colIndex = currentHeaders.indexOf(header);
        if (colIndex === -1) return;

        const filter = columnFilters[header];
        const columnType = columnDataTypes[header];

        if (Array.isArray(filter)) { // Filter untuk string/date
            tempFilteredData = tempFilteredData.filter(row => filter.includes(String(row[colIndex]).trim()));
        } else if (typeof filter === 'object' && filter !== null) { // Filter untuk number
            tempFilteredData = tempFilteredData.filter(row => {
                const val = parseFloat(row[colIndex]);
                if (isNaN(val)) return false;
                const minOk = filter.min === null || val >= filter.min;
                const maxOk = filter.max === null || val <= filter.max;
                return minOk && maxOk;
            });
        }
    });

    // 3. Terapkan Pengurutan
    if (sortColumnIndex !== -1) {
        const columnType = columnDataTypes[currentHeaders[sortColumnIndex]];
        tempFilteredData.sort((a, b) => {
            let valA = a[sortColumnIndex];
            let valB = b[sortColumnIndex];

            if (columnType === 'number') {
                valA = parseFloat(valA) || 0;
                valB = parseFloat(valB) || 0;
                return sortDirection === 'asc' ? valA - valB : valB - valA;
            } else {
                valA = String(valA).toLowerCase();
                valB = String(valB).toLowerCase();
                if (valA < valB) return sortDirection === 'asc' ? -1 : 1;
                if (valA > valB) return sortDirection === 'asc' ? 1 : -1;
                return 0;
            }
        });
    }

    filteredData = tempFilteredData;
    currentPage = 1; // Reset halaman ke 1 setelah filter/sort
    renderTable(filteredData, currentHeaders);
    calculateStatistics(filteredData, currentHeaders); // Perbarui statistik
    calculateGroupedSummary(); // Perbarui ringkasan pengelompokan
    renderChart(); // Perbarui chart
}

/**
 * Memperbarui kontrol paginasi (nomor halaman, tombol aktif/nonaktif).
 * @param {number} totalRows - Total baris data yang difilter.
 */
function updatePaginationControls(totalRows) {
    if (totalRows <= 0) {
        paginationControls.classList.add('hidden');
        return;
    }
    paginationControls.classList.remove('hidden');
    const totalPages = Math.ceil(totalRows / rowsPerPage);
    pageInfo.textContent = `Halaman ${currentPage} dari ${totalPages}`;
    prevPageBtn.disabled = currentPage === 1;
    nextPageBtn.disabled = currentPage >= totalPages;
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
    const exportHeaders = visibleColumns;
    const exportColIndices = exportHeaders.map(header => currentHeaders.indexOf(header));

    let csvContent = exportHeaders.map(header => `"${header}"`).join(',') + '\n'; // Header CSV

    filteredData.forEach(row => {
        const exportRow = exportColIndices.map(colIndex => {
            const cell = row[colIndex];
            const cellStr = String(cell === null || cell === undefined ? "" : cell);
            // Escape quotes by doubling them
            return `"${cellStr.replace(/"/g, '""')}"`;
        });
        csvContent += exportRow.join(',') + '\n';
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    const sheetName = sheetSelect.value.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    link.download = `${sheetName}_export.csv`;
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
    
    // Gunakan pustaka jsPDF dengan plugin autoTable untuk hasil yang lebih baik
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    const head = [visibleColumns];
    const body = filteredData.map(row => {
        return visibleColumns.map(header => {
            const colIndex = currentHeaders.indexOf(header);
            return row[colIndex] ?? "";
        });
    });

    doc.autoTable({
        head: head,
        body: body,
        styles: { fontSize: 8 },
        headStyles: { fillColor: [22, 160, 133] },
        alternateRowStyles: { fillColor: [245, 245, 245] },
        margin: { top: 10 }
    });
    
    const sheetName = sheetSelect.value.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    doc.save(`${sheetName}_export.pdf`);
    displayCustomMessage('Data berhasil diekspor ke PDF!', 'info');
}

/**
 * Menampilkan pesan kustom (pengganti alert()).
 * @param {string} message - Pesan yang akan ditampilkan.
 * @param {string} type - Tipe pesan (misal: 'info', 'error').
 */
function displayCustomMessage(message, type = 'info') {
    const messageBox = document.createElement('div');
    messageBox.className = `fixed top-4 right-4 p-4 rounded-md shadow-lg text-white z-50 transition-opacity duration-300 ${type === 'error' ? 'bg-red-500' : 'bg-blue-500'}`;
    messageBox.textContent = message;
    document.body.appendChild(messageBox);

    setTimeout(() => {
        messageBox.style.opacity = '0';
        setTimeout(() => messageBox.remove(), 300);
    }, 3000);
}

/**
 * Menghitung dan menampilkan statistik dasar untuk kolom numerik.
 * @param {Array<Array<any>>} data - Data tabel (tanpa header).
 * @param {string[]} headers - Header kolom.
 */
function calculateStatistics(data, headers) {
    summaryContentDiv.innerHTML = '';
    dataSummaryDiv.classList.add('hidden');

    const numericColumns = headers.filter((header, i) => visibleColumns.includes(header) && columnDataTypes[header] === 'number');

    if (numericColumns.length === 0) {
        return;
    }
    
    dataSummaryDiv.classList.remove('hidden');

    numericColumns.forEach(header => {
        const colIndex = headers.indexOf(header);
        const values = data.map(row => parseFloat(row[colIndex])).filter(v => !isNaN(v));
        if (values.length === 0) return;

        const sum = values.reduce((acc, val) => acc + val, 0);
        const avg = sum / values.length;
        const min = Math.min(...values);
        const max = Math.max(...values);

        const format = (num) => num.toLocaleString(undefined, { maximumFractionDigits: 2 });

        const summaryEl = document.createElement('div');
        summaryEl.className = 'bg-white p-3 rounded-md shadow-sm border border-blue-100';
        summaryEl.innerHTML = `
            <h3 class="font-semibold text-blue-700 mb-1">${header}</h3>
            <p><strong>Total:</strong> ${format(sum)}</p>
            <p><strong>Rata-rata:</strong> ${format(avg)}</p>
            <p><strong>Min:</strong> ${format(min)}</p>
            <p><strong>Max:</strong> ${format(max)}</p>
        `;
        summaryContentDiv.appendChild(summaryEl);
    });
}

/**
 * Memperbarui dropdown untuk pemilihan sumbu X dan Y pada grafik.
 */
function updateChartControls() {
    const selects = [xAxisSelect, yAxisSelect, scatterXAxisSelect, scatterYAxisSelect];
    selects.forEach(s => s.innerHTML = '');

    const emptyOption = document.createElement('option');
    emptyOption.value = '';
    emptyOption.textContent = 'Pilih Kolom...';
    selects.forEach(s => s.appendChild(emptyOption.cloneNode(true)));

    const numericHeaders = [];
    const allHeaders = [];

    currentHeaders.forEach(header => {
        if (!visibleColumns.includes(header)) return;
        
        allHeaders.push(header);
        if (columnDataTypes[header] === 'number') {
            numericHeaders.push(header);
        }
    });

    allHeaders.forEach(header => {
        const option = document.createElement('option');
        option.value = header;
        option.textContent = header;
        xAxisSelect.appendChild(option.cloneNode(true));
    });

    numericHeaders.forEach(header => {
        const option = document.createElement('option');
        option.value = header;
        option.textContent = header;
        yAxisSelect.appendChild(option.cloneNode(true));
        scatterXAxisSelect.appendChild(option.cloneNode(true));
        scatterYAxisSelect.appendChild(option.cloneNode(true));
    });
    
    updateChartControlsVisibility();
}


/**
 * Mengatur visibilitas kontrol sumbu chart berdasarkan tipe chart yang dipilih.
 */
function updateChartControlsVisibility() {
    const selectedType = chartTypeSelect.value;

    if (selectedType === 'scatter') {
        standardChartAxes.classList.add('hidden');
        yAxisContainer.classList.add('hidden');
        scatterChartAxes.classList.remove('hidden');
        scatterChartAxes.style.display = 'grid'; // Pastikan display grid
    } else {
        standardChartAxes.classList.remove('hidden');
        yAxisContainer.classList.remove('hidden');
        scatterChartAxes.classList.add('hidden');
        scatterChartAxes.style.display = 'none';

        yAxisContainer.querySelector('label').textContent = selectedType === 'pie' ? 'Nilai (Value):' : 'Sumbu Y (Nilai):';
    }
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
    
    let config;
    if (chartType === 'bar' || chartType === 'line' || chartType === 'pie') {
        config = getStandardChartConfig(chartType);
    } else if (chartType === 'scatter') {
        config = getScatterChartConfig();
    }

    if (!config) {
        myChartCanvas.style.display = 'none';
        return;
    }
    
    myChartCanvas.style.display = 'block';
    chartInstance = new Chart(ctx, config);
}

/**
 * Membuat konfigurasi untuk Bar, Line, dan Pie chart.
 */
function getStandardChartConfig(chartType) {
    const xAxisHeader = xAxisSelect.value;
    const yAxisHeader = yAxisSelect.value;

    if (!xAxisHeader || !yAxisHeader || filteredData.length === 0) return null;

    const xAxisIndex = currentHeaders.indexOf(xAxisHeader);
    const yAxisIndex = currentHeaders.indexOf(yAxisHeader);

    if (xAxisIndex === -1 || yAxisIndex === -1) return null;

    const aggregatedData = {};
    filteredData.forEach(row => {
        const xValue = String(row[xAxisIndex] || "N/A").trim();
        const yValue = parseFloat(row[yAxisIndex]);
        if (!isNaN(yValue)) {
            aggregatedData[xValue] = (aggregatedData[xValue] || 0) + yValue;
        }
    });

    const labels = Object.keys(aggregatedData);
    const dataValues = Object.values(aggregatedData);

    if (labels.length === 0) return null;

    const pieColors = ['#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF', '#FF9F40'];

    return {
        type: chartType,
        data: {
            labels: labels,
            datasets: [{
                label: yAxisHeader,
                data: dataValues,
                backgroundColor: chartType === 'pie' ? pieColors : 'rgba(75, 192, 192, 0.6)',
                borderColor: chartType === 'pie' ? '#fff' : 'rgba(75, 192, 192, 1)',
                borderWidth: 1,
                fill: chartType === 'line' ? false : undefined,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: chartType === 'pie' ? undefined : { y: { beginAtZero: true } },
            plugins: {
                title: { display: true, text: `${yAxisHeader} by ${xAxisHeader}` }
            }
        }
    };
}


/**
 * Membuat konfigurasi untuk Scatter chart.
 */
function getScatterChartConfig() {
    const xAxisHeader = scatterXAxisSelect.value;
    const yAxisHeader = scatterYAxisSelect.value;

    if (!xAxisHeader || !yAxisHeader || filteredData.length === 0) return null;

    const xAxisIndex = currentHeaders.indexOf(xAxisHeader);
    const yAxisIndex = currentHeaders.indexOf(yAxisHeader);

    if (xAxisIndex === -1 || yAxisIndex === -1) return null;

    const scatterData = filteredData.map(row => ({
        x: parseFloat(row[xAxisIndex]),
        y: parseFloat(row[yAxisIndex])
    })).filter(p => !isNaN(p.x) && !isNaN(p.y));

    if (scatterData.length === 0) return null;
    
    return {
        type: 'scatter',
        data: {
            datasets: [{
                label: `${yAxisHeader} vs ${xAxisHeader}`,
                data: scatterData,
                backgroundColor: 'rgba(255, 99, 132, 0.6)',
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { type: 'linear', position: 'bottom', title: { display: true, text: xAxisHeader } },
                y: { title: { display: true, text: yAxisHeader } }
            },
            plugins: {
                 title: { display: true, text: `Scatter Plot of ${yAxisHeader} vs ${xAxisHeader}` }
            }
        }
    };
}


/**
 * Membuka modal untuk mengelola visibilitas kolom.
 */
function openColumnManagerModal() {
    columnCheckboxesDiv.innerHTML = ''; // Bersihkan konten sebelumnya

    currentHeaders.forEach(header => {
        const label = document.createElement('label');
        label.className = 'flex items-center space-x-2';
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = header;
        checkbox.checked = visibleColumns.includes(header);
        checkbox.className = 'h-4 w-4 rounded border-gray-300 text-indigo-600 focus:ring-indigo-500';
        label.appendChild(checkbox);
        const text = document.createElement('span');
        text.textContent = header;
        label.appendChild(text);
        columnCheckboxesDiv.appendChild(label);
    });

    columnManagerModal.classList.remove('hidden');
}

/**
 * Menerapkan perubahan visibilitas kolom dari modal.
 */
function applyColumnVisibilityChanges() {
    const newVisibleColumns = Array.from(columnCheckboxesDiv.querySelectorAll('input:checked')).map(cb => cb.value);

    if (newVisibleColumns.length === 0) {
        displayCustomMessage('Setidaknya satu kolom harus terlihat.', 'error');
        return;
    }

    // Pertahankan urutan asli dari currentHeaders
    visibleColumns = currentHeaders.filter(header => newVisibleColumns.includes(header));
    
    localStorage.setItem('visibleColumns', JSON.stringify(visibleColumns));
    columnManagerModal.classList.add('hidden');
    applyAllFilters(); // Render ulang tabel dengan kolom yang terlihat
    updateChartControls(); // Perbarui dropdown chart
    populateGroupingControls(); // Perbarui grup
    renderChart(); // Render ulang chart
    calculateStatistics(filteredData, currentHeaders); // Perbarui statistik
}

/**
 * Menyimpan preferensi pengguna ke localStorage.
 */
function savePreferences() {
    if (sheetSelect.value) {
        localStorage.setItem('lastSheet', sheetSelect.value);
    }
    localStorage.setItem('rowsPerPage', rowsPerPageSelect.value);
}

/**
 * Memuat preferensi pengguna dari localStorage.
 */
function loadPreferences() {
    const savedRowsPerPage = localStorage.getItem('rowsPerPage');
    if (savedRowsPerPage) {
        rowsPerPageSelect.value = savedRowsPerPage;
        rowsPerPage = parseInt(savedRowsPerPage, 10);
    }
}

/**
 * Mereset semua filter, pengurutan, visibilitas kolom, dan data yang dimuat.
 */
function resetAll() {
    if (!confirm('Apakah Anda yakin ingin mereset aplikasi? Semua filter, urutan, dan perubahan akan hilang.')) {
        return;
    }
    
    // Hapus preferensi
    localStorage.removeItem('rowsPerPage');
    localStorage.removeItem('lastSheet');
    localStorage.removeItem('visibleColumns');
    
    // Reset variabel
    currentWorkbook = null;
    currentSheetData = [];
    currentHeaders = [];
    filteredData = [];
    sortColumnIndex = -1;
    sortDirection = 'asc';
    columnFilters = {};
    visibleColumns = [];
    currentPage = 1;
    rowsPerPage = defaultRowsPerPage;
    
    // Reset UI
    excelFileInput.value = '';
    searchInput.value = '';
    sheetSelect.innerHTML = '';
    tableContainer.innerHTML = '<p class="text-center text-gray-600">Silakan unggah file Excel untuk memulai.</p>';
    controlsDiv.classList.add('hidden');
    paginationControls.classList.add('hidden');
    dataSummaryDiv.classList.add('hidden');
    chartControlsDiv.classList.add('hidden');
    groupingControlsDiv.classList.add('hidden');
    if (chartInstance) chartInstance.destroy();
    
    displayCustomMessage('Aplikasi telah direset.', 'info');
}

/**
 * Menyalin data tabel yang difilter dan diurutkan ke papan klip.
 */
function copyTableToClipboard() {
    if (filteredData.length === 0) {
        displayCustomMessage('Tidak ada data untuk disalin.', 'error');
        return;
    }

    const copyHeaders = visibleColumns;
    const copyColIndices = copyHeaders.map(header => currentHeaders.indexOf(header));

    let clipboardContent = copyHeaders.join('\t') + '\n';
    clipboardContent += filteredData.map(row => 
        copyColIndices.map(index => row[index] ?? "").join('\t')
    ).join('\n');

    navigator.clipboard.writeText(clipboardContent).then(() => {
        displayCustomMessage('Data tabel berhasil disalin!', 'info');
    }, (err) => {
        displayCustomMessage('Gagal menyalin data.', 'error');
        console.error('Could not copy text: ', err);
    });
}

/**
 * Mempopulasi dropdown untuk kontrol pengelompokan data.
 */
function populateGroupingControls() {
    const selects = [groupBySelect, aggregateBySelect];
    selects.forEach(s => s.innerHTML = '');

    const emptyOption = document.createElement('option');
    emptyOption.value = '';
    emptyOption.textContent = 'Pilih Kolom...';
    selects.forEach(s => s.appendChild(emptyOption.cloneNode(true)));
    
    const numericHeaders = [];
    const allHeaders = [];

    currentHeaders.forEach(header => {
        if (!visibleColumns.includes(header)) return;
        allHeaders.push(header);
        if (columnDataTypes[header] === 'number') {
            numericHeaders.push(header);
        }
    });

    allHeaders.forEach(header => {
        const option = document.createElement('option');
        option.value = header;
        option.textContent = header;
        groupBySelect.appendChild(option);
    });
    
    numericHeaders.forEach(header => {
        const option = document.createElement('option');
        option.value = header;
        option.textContent = header;
        aggregateBySelect.appendChild(option);
    });
}

/**
 * Menghitung dan menampilkan ringkasan data yang dikelompokkan.
 */
function calculateGroupedSummary() {
    groupedSummaryContentDiv.innerHTML = '';
    const groupByHeader = groupBySelect.value;
    const aggregateByHeader = aggregateBySelect.value;

    if (!groupByHeader || !aggregateByHeader || filteredData.length === 0) {
        groupedSummaryContentDiv.innerHTML = '<p class="text-gray-600">Pilih kolom untuk pengelompokan dan agregasi.</p>';
        return;
    }

    const groupByColIndex = currentHeaders.indexOf(groupByHeader);
    const aggregateByColIndex = currentHeaders.indexOf(aggregateByHeader);

    const grouped = {};
    filteredData.forEach(row => {
        const key = row[groupByColIndex] || "N/A";
        const value = parseFloat(row[aggregateByColIndex]);
        if (isNaN(value)) return;
        
        if (!grouped[key]) {
            grouped[key] = { sum: 0, count: 0, values: [] };
        }
        grouped[key].sum += value;
        grouped[key].count++;
        grouped[key].values.push(value);
    });
    
    if(Object.keys(grouped).length === 0) {
        groupedSummaryContentDiv.innerHTML = '<p class="text-gray-600">Tidak ada data untuk dikelompokkan.</p>';
        return;
    }

    for (const key in grouped) {
        const group = grouped[key];
        const avg = group.sum / group.count;
        const format = (num) => num.toLocaleString(undefined, { maximumFractionDigits: 2 });

        const card = document.createElement('div');
        card.className = 'bg-white p-3 rounded-md shadow-sm border border-yellow-100';
        card.innerHTML = `
            <h3 class="font-semibold text-yellow-700 mb-1">${groupByHeader}: ${key}</h3>
            <p><strong>Total ${aggregateByHeader}:</strong> ${format(group.sum)}</p>
            <p><strong>Rata-rata ${aggregateByHeader}:</strong> ${format(avg)}</p>
            <p><strong>Jumlah Data:</strong> ${group.count}</p>
        `;
        groupedSummaryContentDiv.appendChild(card);
    }
}

// --- Drag and Drop Kolom ---
function handleColumnReorderStart(e) {
    draggedTh = e.target;
    e.dataTransfer.effectAllowed = 'move';
    setTimeout(() => {
        draggedTh.classList.add('opacity-50');
    }, 0);
}

function handleColumnReorderOver(e) {
    e.preventDefault();
    const targetTh = e.target.closest('th');
    if (targetTh && targetTh !== draggedTh) {
        if (dragOverTh) dragOverTh.classList.remove('border-l-2', 'border-blue-500');
        dragOverTh = targetTh;
        dragOverTh.classList.add('border-l-2', 'border-blue-500');
    }
}

function handleColumnReorderLeave(e) {
    const targetTh = e.target.closest('th');
    if (targetTh && targetTh === dragOverTh) {
        dragOverTh.classList.remove('border-l-2', 'border-blue-500');
        dragOverTh = null;
    }
}

function handleColumnReorderDrop(e) {
    e.preventDefault();
    const targetTh = e.target.closest('th');
    if (targetTh && targetTh !== draggedTh) {
        const fromHeader = draggedTh.textContent.trim();
        const toHeader = targetTh.textContent.trim();
        
        const fromIndex = visibleColumns.indexOf(fromHeader);
        const toIndex = visibleColumns.indexOf(toHeader);

        if (fromIndex !== -1 && toIndex !== -1) {
            const [movedItem] = visibleColumns.splice(fromIndex, 1);
            visibleColumns.splice(toIndex, 0, movedItem);
            
            localStorage.setItem('visibleColumns', JSON.stringify(visibleColumns));
            applyAllFilters();
        }
    }
}

function handleColumnReorderEnd(e) {
    draggedTh.classList.remove('opacity-50');
    if (dragOverTh) {
        dragOverTh.classList.remove('border-l-2', 'border-blue-500');
    }
    draggedTh = null;
    dragOverTh = null;
}
