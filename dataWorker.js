importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js');

// Message handler for the Web Worker
self.addEventListener('message', (event) => {
    const { type, workbook, criteria, data, endDate } = event.data;

    switch (type) {
        case 'PROCESS_FILE':
            const parsedWorkbook = JSON.parse(workbook);
            const processedData = processWorkbook(parsedWorkbook);
            self.postMessage({ type: 'DATA_PROCESSED', payload: processedData });
            break;
        case 'APPLY_FILTERS':
            const filtered = applyFilters(criteria, data);
            self.postMessage({ type: 'FILTERED_DATA', payload: { filteredData: filtered } });
            break;
        case 'EXPAND_TABLE':
            const expanded = expandData(data, endDate);
            self.postMessage({ type: 'EXPANDED_DATA', payload: { expandedData: expanded } });
            break;
        default:
            console.error('Unknown message type:', type);
    }
});

// Process workbook data
function processWorkbook(workbook) {
    const sheetNames = workbook.SheetNames;
    const data = workbook.SheetNames.reduce((acc, sheet) => {
        acc[sheet] = XLSX.utils.sheet_to_json(workbook.Sheets[sheet]);
        return acc;
    }, {});

    const filters = {
        marca: getUniqueValues(data, 'marca'),
        condizione: getUniqueValues(data, 'nuovousato'),
        acquisto: getUniqueValues(data, 'acquistoleasing'),
        statoAttuale: getUniqueValues(data, 'revendita'),
    };

    return { data, filters };
}

// Apply filters to the data
function applyFilters(criteria, data) {
    return data.filter(row =>
        (!criteria.marca.length || criteria.marca.includes(row.marca)) &&
        (!criteria.condizione.length || criteria.condizione.includes(row.nuovousato)) &&
        (!criteria.acquisto.length || criteria.acquisto.includes(row.acquistoleasing)) &&
        (!criteria.statoAttuale.length || criteria.statoAttuale.includes(row.revendita))
    );
}

// Expand data
function expandData(data, endDate) {
    const expanded = [];
    const endMoment = moment(endDate, 'YYYY-MM-DD');

    data.forEach(row => {
        const startMoment = moment(row.dataacquisto, 'DD/MM/YYYY');
        while (startMoment.isBefore(endMoment)) {
            expanded.push({ ...row, mese: startMoment.format('MM/YYYY') });
            startMoment.add(1, 'month');
        }
    });

    return expanded;
}

// Get unique values for filters
function getUniqueValues(data, key) {
    return [...new Set(data.map(row => row[key]))];
}