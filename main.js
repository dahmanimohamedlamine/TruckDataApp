const dataWorker = new Worker('dataWorker.js'); // Initialize the Web Worker

let initialData = [];
let currentData = [];
let filteredData = [];
let isExpanded = false;

// File Upload Handling
document.getElementById('fileUpload').addEventListener('change', function (event) {
    const reader = new FileReader();
    reader.onload = function (event) {
        const dataArray = new Uint8Array(event.target.result);
        const workbook = XLSX.read(dataArray, { type: 'array' });
        dataWorker.postMessage({ type: 'PROCESS_FILE', workbook: JSON.stringify(workbook) });
    };
    reader.readAsArrayBuffer(event.target.files[0]);
});

// Worker Message Handling
dataWorker.addEventListener('message', (event) => {
    const { type, payload } = event.data;

    if (type === 'DATA_PROCESSED') {
        initialData = payload.initialData;
        currentData = [...initialData];
        populateFilters(payload.filters);
        displayData();
    } else if (type === 'FILTERED_DATA') {
        filteredData = payload.filteredData;
        displayData();
    } else if (type === 'EXPANDED_DATA') {
        currentData = payload.expandedData;
        isExpanded = true;
        applyFilters();
    }
});

// Populate Filters
function populateFilters(filters) {
    populateDropdown('marca', filters.marca);
    populateDropdown('condizione', filters.condizione);
    populateDropdown('acquisto', filters.acquisto);
    populateDropdown('statoAttuale', filters.statoAttuale);
}

// Dropdown Utility
function populateDropdown(id, values) {
    const select = document.getElementById(id);
    select.innerHTML = '';
    values.forEach(value => {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = value;
        select.appendChild(option);
    });
}

// Apply Filters
function applyFilters() {
    const criteria = {
        marca: getSelectedValues('marca'),
        condizione: getSelectedValues('condizione'),
        acquisto: getSelectedValues('acquisto'),
        statoAttuale: getSelectedValues('statoAttuale'),
    };
    dataWorker.postMessage({ type: 'APPLY_FILTERS', criteria, data: currentData });
}

// Expand Data
function expandTable() {
    const endDate = document.getElementById('dataFineRivalutazione').value;
    dataWorker.postMessage({ type: 'EXPAND_TABLE', endDate, data: currentData });
}

// Display Data
function displayData() {
    const tableHeader = document.getElementById('tableHeader');
    const tableBody = document.getElementById('tableBody');

    tableHeader.innerHTML = '';
    tableBody.innerHTML = '';

    if (filteredData.length === 0) return;

    // Build Table Header
    Object.keys(filteredData[0]).forEach(key => {
        const th = document.createElement('th');
        th.textContent = key;
        tableHeader.appendChild(th);
    });

    // Build Table Body
    filteredData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(value => {
            const td = document.createElement('td');
            td.textContent = value;
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
}

// Get Selected Filter Values
function getSelectedValues(id) {
    return Array.from(document.getElementById(id).selectedOptions).map(option => option.value);
}