const dataWorker = new Worker('dataWorker.js');
let dataCache = {};
let currentData = [];
let filteredData = [];
let currentPage = 1;
const rowsPerPage = 15;

// File Upload Handler
document.getElementById('fileUpload').addEventListener('change', function (event) {
    const reader = new FileReader();
    reader.onload = function (event) {
        const dataArray = new Uint8Array(event.target.result);
        const workbook = XLSX.read(dataArray, { type: 'array' });
        dataWorker.postMessage({ type: 'PROCESS_FILE', workbook });
    };
    reader.readAsArrayBuffer(event.target.files[0]);
});

// Listen to Worker Messages
dataWorker.addEventListener('message', (event) => {
    const { type, payload } = event.data;

    if (type === 'DATA_PROCESSED') {
        dataCache = payload;
        populateCausaDropdown(Object.keys(dataCache));
        loadSheetData(Object.keys(dataCache)[0]);
    } else if (type === 'FILTERED_DATA') {
        filteredData = payload.filteredData;
        displayData();
    } else if (type === 'EXPANDED_DATA') {
        currentData = payload;
        applyFilters();
    }
});

// Utility Functions
function applyFilters() {
    dataWorker.postMessage({
        type: 'APPLY_FILTERS',
        criteria: getFilterCriteria(),
        data: currentData
    });
}

function expandTable() {
    dataWorker.postMessage({
        type: 'EXPAND_TABLE',
        endDate: document.getElementById('dataFineRivalutazione').value,
        data: currentData
    });
}