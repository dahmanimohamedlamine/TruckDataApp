importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js');

self.addEventListener('message', (event) => {
    const { type, workbook, criteria, endDate, data } = event.data;

    if (type === 'PROCESS_FILE') {
        processFile(workbook);
    } else if (type === 'APPLY_FILTERS') {
        filterData(criteria, data);
    } else if (type === 'EXPAND_TABLE') {
        expandData(endDate, data);
    }
});

function processFile(workbook) {
    let dataCache = {};
    workbook.SheetNames.forEach(sheet => {
        dataCache[sheet] = XLSX.utils.sheet_to_json(workbook.Sheets[sheet]);
    });
    self.postMessage({ type: 'DATA_PROCESSED', payload: dataCache });
}

function filterData(criteria, data) {
    const filteredData = data.filter(item => {
        // Apply filtering logic
    });
    self.postMessage({ type: 'FILTERED_DATA', payload: { filteredData } });
}

function expandData(endDate, data) {
    const expandedData = [];
    data.forEach(row => {
        // Expansion logic here
    });
    self.postMessage({ type: 'EXPANDED_DATA', payload: expandedData });
}