importScripts('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js');

// Utility Functions
function processData(workbook) {
    const sheets = JSON.parse(workbook).SheetNames;
    const initialData = {};
    sheets.forEach(sheet => {
        initialData[sheet] = XLSX.utils.sheet_to_json(JSON.parse(workbook).Sheets[sheet]);
    });
    return initialData;
}

function applyFilters(criteria, data) {
    return data.filter(row => {
        return (
            (!criteria.marca.length || criteria.marca.includes(row.marca)) &&
            (!criteria.condizione.length || criteria.condizione.includes(row.nuovousato)) &&
            (!criteria.acquisto.length || criteria.acquisto.includes(row.acquistoleasing)) &&
            (!criteria.statoAttuale.length || criteria.statoAttuale.includes(row.revendita))
        );
    });
}

function expandData(data, endDate) {
    const expanded = [];
    data.forEach(row => {
        const startDate = moment(row.mese_acquisto, "DD/MM/YYYY");
        while (startDate.isBefore(endDate)) {
            const newRow = { ...row, mese: startDate.format("MM/YYYY") };
            expanded.push(newRow);
            startDate.add(1, 'month');
        }
    });
    return expanded;
}

// Worker Message Handling
self.addEventListener('message', (event) => {
    const { type, workbook, criteria, data, endDate } = event.data;

    if (type === 'PROCESS_FILE') {
        const processedData = processData(workbook);
        self.postMessage({ type: 'DATA_PROCESSED', payload: { initialData: processedData, filters: getFilters(processedData) } });
    } else if (type === 'APPLY_FILTERS') {
        const filtered = applyFilters(criteria, data);
        self.postMessage({ type: 'FILTERED_DATA', payload: { filteredData: filtered } });
    } else if (type === 'EXPAND_TABLE') {
        const expanded = expandData(data, endDate);
        self.postMessage({ type: 'EXPANDED_DATA', payload: { expandedData: expanded } });
    }
});

function getFilters(data) {
    return {
        marca: [...new Set(data.map(row => row.marca))],
        condizione: [...new Set(data.map(row => row.nuovousato))],
        acquisto: [...new Set(data.map(row => row.acquistoleasing))],
        statoAttuale: [...new Set(data.map(row => row.revendita))],
    };
}