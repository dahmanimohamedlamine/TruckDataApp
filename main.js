const dataWorker = new Worker('dataWorker.js'); // Create a Web Worker

let initialData = []; // Original data
let currentData = []; // Filtered data
let filteredData = [];
let currentPage = 1;
const rowsPerPage = 15;

// Handle file upload
document.getElementById('fileUpload').addEventListener('change', function(event) {
    const reader = new FileReader();
    reader.onload = function(event) {
        const dataArray = new Uint8Array(event.target.result);
        const workbook = XLSX.read(dataArray, { type: 'array' });

        // Send workbook to Web Worker for processing
        dataWorker.postMessage({ type: 'PROCESS_FILE', workbook: JSON.stringify(workbook) });
    };
    reader.readAsArrayBuffer(event.target.files[0]);
});

// Listen for messages from the worker
dataWorker.addEventListener('message', (event) => {
    const { type, payload } = event.data;

    switch (type) {
        case 'DATA_PROCESSED':
            initialData = payload.data;
            currentData = [...initialData];
            populateFilters(payload.filters);
            displayData();
            break;
        case 'FILTERED_DATA':
            filteredData = payload.filteredData;
            displayData();
            break;
        case 'EXPANDED_DATA':
            currentData = payload.expandedData;
            filteredData = [...currentData];
            displayData();
            break;
        default:
            console.error('Unknown message type:', type);
    }
});

// Populate filter dropdowns
function populateFilters(filters) {
    populateDropdown('marca', filters.marca);
    populateDropdown('condizione', filters.condizione);
    populateDropdown('acquisto', filters.acquisto);
    populateDropdown('statoAttuale', filters.statoAttuale);
}

function populateDropdown(id, values) {
    const select = document.getElementById(id);
    select.innerHTML = values.map(value => `<option value="${value}">${value}</option>`).join('');
    select.addEventListener('change', applyFilters);
}

// Apply filters
function applyFilters() {
    const criteria = {
        marca: getSelectedValues('marca'),
        condizione: getSelectedValues('condizione'),
        acquisto: getSelectedValues('acquisto'),
        statoAttuale: getSelectedValues('statoAttuale'),
    };
    dataWorker.postMessage({ type: 'APPLY_FILTERS', criteria, data: currentData });
}

// Expand table
function expandTable() {
    const endDate = document.getElementById('dataFineRivalutazione').value;
    dataWorker.postMessage({ type: 'EXPAND_TABLE', endDate, data: initialData });
}

// Get selected values from dropdown
function getSelectedValues(id) {
    return Array.from(document.getElementById(id).selectedOptions).map(option => option.value);
}

// Display data in the table
function displayData() {
    const tableHeader = document.getElementById('tableHeader');
    const tableBody = document.getElementById('tableBody');
    tableHeader.innerHTML = '';
    tableBody.innerHTML = '';

    if (filteredData.length === 0) return;

    Object.keys(filteredData[0]).forEach(key => {
        const th = document.createElement('th');
        th.textContent = key;
        tableHeader.appendChild(th);
    });

    const startIdx = (currentPage - 1) * rowsPerPage;
    const endIdx = startIdx + rowsPerPage;
    const pageData = filteredData.slice(startIdx, endIdx);

    pageData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(value => {
            const td = document.createElement('td');
            td.textContent = value;
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
}