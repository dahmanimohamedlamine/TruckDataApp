let data = {}; 
let initialData = [];  // Store the original unexpanded data
let currentData = []; 
let filteredData = []; 
let currentPage = 1; 
const rowsPerPage = 15; 
let isExpanded = false;  // Track if the data has already been expanded

document.getElementById('fileUpload').addEventListener('change', function(event) {
    const reader = new FileReader();
    reader.onload = function(event) {
        const dataArray = new Uint8Array(event.target.result);
        const workbook = XLSX.read(dataArray, { type: 'array' });
        data = workbook.SheetNames.reduce((acc, sheetName) => {
            acc[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });
            return acc;
        }, {});
        
        populateCausaDropdown(workbook.SheetNames);
        isExpanded = false;
        loadSheetData(workbook.SheetNames[0]); 
    };
    reader.readAsArrayBuffer(event.target.files[0]);
});

function populateCausaDropdown(sheetNames) {
    const causaSelect = document.getElementById('causa');
    causaSelect.innerHTML = '';
    sheetNames
        .filter(name => name !== 'TEGM') // Filter out "TEGM"
        .forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            causaSelect.appendChild(option);
        });
    causaSelect.addEventListener('change', function() {
        loadSheetData(this.value); 
    });
}

function loadSheetData(sheetName) {
    let tegmData = []; // Define tegmData outside the if block for wider scope

    // Check if TEGM sheet is present and load it if so
    if (data['TEGM']) {
        // Process TEGM sheet data
        tegmData = data['TEGM'].map(row => ({ ...row })); // Make a copy of TEGM data

        // Transform 'qtr' column in TEGM data using transformDateColumns function
        transformDateColumns(tegmData);

        // Add quarter_acquisto and anno_acquisto based on transformed 'qtr'
        tegmData = tegmData.map(row => ({
            ...row,
            quarter_acquisto: row.qtr ? moment(row.qtr, "DD/MM/YYYY").quarter() : '',
            anno_acquisto: row.qtr ? moment(row.qtr, "DD/MM/YYYY").year() : ''
        }));
    }

    // Load initial data
    initialData = data[sheetName].map(row => {
        delete row.anno_acquisto; // Remove existing 'anno_acquisto' if it exists
        return {
            ...row,
            revendita: row.datavendita ? 'VENDUTI' : 'NO'
        };
    });

    transformDateColumns(initialData);

    // Add `quarter_acquisto` and `anno_acquisto` in initialData from `dataacquisto`
    initialData = initialData.map(row => ({
        ...row,
        quarter_acquisto: row.dataacquisto ? moment(row.dataacquisto, "DD/MM/YYYY").quarter() : '',
        anno_acquisto: row.dataacquisto ? moment(row.dataacquisto, "DD/MM/YYYY").year() : ''
    }));

    // Check if tegmData has been loaded properly
    if (tegmData.length === 0) {
        console.warn("TEGM data not loaded. Check if TEGM sheet exists in the file.");
    } else {
        // Match `initialData` with `tegmData` based on `quarter_acquisto` and `anno_acquisto`
        initialData = initialData.map(row => {
            // Find matching row in tegmData
            const match = tegmData.find(tegmRow => 
                tegmRow.quarter_acquisto === row.quarter_acquisto && 
                tegmRow.anno_acquisto === row.anno_acquisto
            );

            // Merge match data if found
            return match ? { ...row, ...match } : row;
        });
    }
    // Clear `tegmData` as it is no longer needed
    tegmData = null;
    // Define `tegm` based on `Prezzo Netto` ranges
    initialData = initialData.map(row => {
        let tegmValue;

        // Determine tegm based on Prezzo Netto ranges
        if (parseFloat(row['prezzo_netto']) <= 5000) {
            tegmValue = row.tegm0_5000/100;
        } else if (parseFloat(row['prezzo_netto'])  > 5000 && parseFloat(row['prezzo_netto'])  <= 25000) {
            tegmValue = row.tegm5000_25000/100;
        } else if (parseFloat(row['prezzo_netto']) > 25000 && parseFloat(row['prezzo_netto'])  <= 50000) {
            tegmValue = row.tegm25000_50000/100;
        } else if (parseFloat(row['prezzo_netto']) > 50000) {
            tegmValue = row.tegmoltre_50000/100;
        } else {
            tegmValue = null; // Default to null if no range matches
        }

        // Add `tegm` column to the row
        return { ...row, tegm: tegmValue };
    });

                // Create TEGM Mensile: Monthly TEGM rate
    initialData = initialData.map(row => ({
        ...row,
        'TEGM Mensile': row.tegm !== null ? (Math.pow(1 + row.tegm, 1 / 12) - 1) : null  // Calculate the monthly TEGM
    }));

    // Drop the TEGM-related columns (like tegm0_5000, tegm5000_25000, etc.)
    initialData = initialData.map(row => {
        // Remove all columns starting with 'tegm'
        Object.keys(row).forEach(key => {
            if (key.startsWith('tegm') || key === 'qtr'|| key === 'quarter') {
                delete row[key];
            }
        });
        return row;
    });

    isExpanded = false;
    currentData = [...initialData];  // Set currentData to the processed initialData
    populateFilters(currentData);
    // Expand the table and populate filters
    expandTable();
}



function transformDateColumns(data) {
    data.forEach(row => {
        for (const key in row) {
            if (/data|date|mese|qtr/i.test(key) && row[key]) {
                row[key] = formatExcelDate(row[key]);
            }
        }
    });
}

function formatExcelDate(value) {
    if (typeof value === 'number') {
        const date = new Date((value - 25569) * 86400 * 1000);
        return moment(date).format("DD/MM/YYYY");
    } else if (moment(value, ["DD/MM/YYYY", "YYYY-MM-DD", "MM/DD/YYYY"], true).isValid()) {
        return moment(value).format("DD/MM/YYYY");
    }
    return value;
}

function populateFilters(data) {
    populateDropdown('marca', getUniqueValues(data, 'marca'));
    populateDropdown('condizione', getUniqueValues(data, 'nuovousato'));
    populateDropdown('acquisto', getUniqueValues(data, 'acquistoleasing'));
    populateDropdown('statoAttuale', getUniqueValues(data, 'revendita'));
    const prezzoColumns = Object.keys(data[0]).filter(col => col.toLowerCase().includes('price') || col.toLowerCase().includes('prezzo'));
    populateDropdown('prezzo', prezzoColumns, 'prezzo_netto');
}

function getUniqueValues(data, key) {
    return [...new Set(data.map(item => item[key]))];
}

function populateDropdown(elementId, values, defaultValue = null) {
    const select = document.getElementById(elementId);
    select.innerHTML = '';
    values.forEach(value => {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = value;
        select.appendChild(option);
    });
    if (defaultValue) select.value = defaultValue;
    select.addEventListener('change', filterAndDisplayData);
}

function expandTable() {
    if (isExpanded) {
        currentData = [...initialData]; // Reset to unexpanded data
        isExpanded = false;
        applyFilters(); // Reapply filters on unexpanded data
        return;
    }

    const endDate = moment(document.getElementById('dataFineRivalutazione').value, "YYYY-MM-DD");
    let expandedData = [];
    const batchSize = 1000; // Process 1000 rows per batch
    let currentIndex = 0;

    function processBatch() {
        const batchEndIndex = Math.min(currentIndex + batchSize, currentData.length);

        for (let i = currentIndex; i < batchEndIndex; i++) {
            const row = currentData[i];
            const startDate = moment(row.mese_acquisto, "DD/MM/YYYY");
            const dataFineDate = moment(row.data_fine, "DD/MM/YYYY");
            let firstMonth = true;

            if (startDate.isValid() && endDate.isValid() && startDate.isBefore(endDate)) {
                let currentDate = startDate.clone();

                while (currentDate.isSameOrBefore(endDate, 'month')) {
                    let expandedRow = { ...row };
                    expandedRow.mese = currentDate.format("MM/YYYY");

                    if (row.acquistoleasing !== "LEASING" && !firstMonth) {
                        Object.keys(expandedRow).forEach(key => {
                            if (key.toLowerCase().includes("prezzo") || key.toLowerCase().includes("price")) {
                                expandedRow[key] = "";
                            }
                        });
                    }

                    if (row.acquistoleasing === "LEASING" && dataFineDate.isValid() && currentDate.isAfter(dataFineDate, 'month')) {
                        Object.keys(expandedRow).forEach(key => {
                            if (key.toLowerCase().includes("prezzo") || key.toLowerCase().includes("price") || key.toLowerCase().includes("riscatto")) {
                                expandedRow[key] = "";
                            }
                        });
                    }

                    expandedData.push(expandedRow);
                    currentDate.add(1, 'month');
                    firstMonth = false;
                }
            } else {
                expandedData.push(row);
            }
        }

        currentIndex = batchEndIndex;

        if (currentIndex < currentData.length) {
            setTimeout(processBatch, 0); // Schedule next batch
        } else {
            currentData = expandedData; // Update currentData
            isExpanded = true; // Mark as expanded
            ProcessData();
        }
    }

    processBatch();
}


function ProcessData() {
    const selectedPrezzo = document.getElementById('prezzo').value;
    const sovrapprezzoCartello = parseFloat(document.getElementById('sovrapprezzoCartello').value) || 0;
    const sovrapprezzoLingering = parseFloat(document.getElementById('sovrapprezzoLingering').value) || 0;

    const cumulativeQuotaPerVehicle = {};

    // Map and process the filtered data
    currentData= currentData.map(item => {
        let row = { ...item };
        row['Prezzo Netto'] = selectedPrezzo && row[selectedPrezzo] ? row[selectedPrezzo] : '';

        const dataAcquisto = moment(row.dataacquisto, "DD/MM/YYYY");
        const cartelloDate = moment("18/01/2011", "DD/MM/YYYY");
        const percentage = dataAcquisto.isValid() && dataAcquisto.isSameOrBefore(cartelloDate, 'day') ? sovrapprezzoCartello : sovrapprezzoLingering;

        const startDate = moment(row.mese_acquisto, "DD/MM/YYYY").format("MM/YYYY");
        const currentDate = moment(row.mese, "MM/YYYY");
        const dataFineDate = moment(row.data_fine, "DD/MM/YYYY");
        const durationPassed = currentDate.diff(moment(startDate, "MM/YYYY"), 'months');
        const durataResidua = row.durata ? Math.max(parseInt(row.durata) - durationPassed, 0) : null;

        if (row['Prezzo Netto']) {
            row['Danno Overcharge'] = (1 - (1 / (1 + (percentage / 100)))) * parseFloat(row['Prezzo Netto']);
        } else {
            row['Danno Overcharge'] = '';
        }

        if (row.acquistoleasing != "LEASING") {
            row['Danno Sovrapprezzo'] = (1 - (1 / (1 + (percentage / 100)))) * parseFloat(row['Prezzo Netto']);
        } else {
            row['Danno Sovrapprezzo'] = '';
        }

        row['Durata Residua'] = durataResidua;

        if (durataResidua === 0) {
            row['riscatto'] = '';
        }

        if (row['riscatto']) {
            row['Danno Riscatto'] = (1 - (1 / (1 + (percentage / 100)))) * parseFloat(row['riscatto']);
        } else {
            row['Danno Riscatto'] = 0;
        }

        if (row['Danno Overcharge'] && row['TEGM Mensile'] && row['durata']) {
            const rataCostante = ((parseFloat(row['Danno Overcharge']) - parseFloat(row['Danno Riscatto'])) /
                (1 - (1 / Math.pow(1 + parseFloat(row['TEGM Mensile']), parseInt(row['durata']))))) *
                parseFloat(row['TEGM Mensile']);

            const rataRiscatto = rataCostante + (parseFloat(row['Danno Riscatto']) * parseFloat(row['TEGM Mensile']));

            const quotaCapitaleIniziale = durataResidua !== null && rataCostante
                ? rataCostante / Math.pow(1 + parseFloat(row['TEGM Mensile']), durataResidua)
                : '';

            const quotaInteressi = quotaCapitaleIniziale && rataRiscatto
                ? rataRiscatto - quotaCapitaleIniziale
                : '';

            const vehicleKey = `${row.impresa}-${row.targa}-${row.nuovousato}`;

            if (!cumulativeQuotaPerVehicle[vehicleKey]) {
                cumulativeQuotaPerVehicle[vehicleKey] = 0;
            }

            cumulativeQuotaPerVehicle[vehicleKey] += quotaCapitaleIniziale ? parseFloat(quotaCapitaleIniziale) : 0;

            let capitaleResiduo = row['Danno Overcharge'] && quotaCapitaleIniziale
                ? row['Danno Overcharge'] - cumulativeQuotaPerVehicle[vehicleKey]
                : '';

            row['Quota Capitale'] = (durataResidua === 1 || currentDate.isSame(dataFineDate, 'month'))
                ? (quotaCapitaleIniziale ? parseFloat(quotaCapitaleIniziale) : 0) + (capitaleResiduo ? parseFloat(capitaleResiduo) : 0)
                : quotaCapitaleIniziale;


            if (row.acquistoleasing === "LEASING") {
                row['Danno Sovrapprezzo'] = (row['Quota Capitale'] ? parseFloat(row['Quota Capitale']) : 0) +
                                            (quotaInteressi ? parseFloat(quotaInteressi) : 0);
            }

            if (row.acquistoleasing === "LEASING" && durataResidua === 0)  {
                row['Danno Sovrapprezzo'] = '';
            }

            return {
                ...row,
                'Rata Costante': rataCostante,
                'Rata Riscatto': rataRiscatto,
                'Quota Capitale Iniziale': quotaCapitaleIniziale,
                'Quota Interessi': quotaInteressi,
                'Capitale Residuo': capitaleResiduo,
                'Quota Capitale': row['Quota Capitale'],
            };
        }

        return {
            ...row,
            'Rata Costante': '',
            'Rata Riscatto': '',
            'Quota Capitale Iniziale': '',
            'Quota Interessi': '',
            'Capitale Residuo': '',
            'Quota Capitale': '',
        };
    });
    filterAndDisplayData();
}



function filterAndDisplayData() {
    const selectedMarca = getSelectedValues('marca');
    const selectedCondizione = getSelectedValues('condizione');
    const selectedAcquisto = getSelectedValues('acquisto');
    const selectedStatoAttuale = getSelectedValues('statoAttuale');
    const selectedPrezzo = document.getElementById('prezzo').value;

    // Filter the data based on selected criteria
    filteredData = currentData.filter(item => {
        return (selectedMarca.length === 0 || selectedMarca.includes(item.marca)) &&
               (selectedCondizione.length === 0 || selectedCondizione.includes(item.nuovousato)) &&
               (selectedAcquisto.length === 0 || selectedAcquisto.includes(item.acquistoleasing)) &&
               (selectedStatoAttuale.length === 0 || selectedStatoAttuale.includes(item.revendita)) &&
               (selectedPrezzo === '' || item[selectedPrezzo] != null);
    });

    // Display filtered data
    displayData();

    // Call calculateFatturatoAndConteggio with the selected price
    calculateFatturatoAndConteggio(selectedPrezzo);
}



function calculateFatturatoAndConteggio(prezzoColumn) {
if (!prezzoColumn) {
document.getElementById('fatturato').textContent = '€0,00';
document.getElementById('conteggio').textContent = '0';
document.getElementById('dannoCartello').textContent = '€0,00';
document.getElementById('dannoLingering').textContent = '€0,00';
return;
}

let fatturato = 0;
let dannoCartelloTotal = 0;
let dannoLingeringTotal = 0;
const uniqueVehicles = new Map();  // Use Map for faster lookups
const cartelloDate = moment("18/01/2011", "DD/MM/YYYY");

// Loop through filteredData only once
filteredData.forEach(row => {
const uniqueKey = `${row.impresa}-${row.targa}-${row.nuovousato}`;

// Update uniqueVehicles and calculate fatturato
if (!uniqueVehicles.has(uniqueKey)) {
    uniqueVehicles.set(uniqueKey, true);
    const prezzoValue = parseFloat(row[prezzoColumn]) || 0;
    fatturato += prezzoValue;
}

// Calculate Danno Sovrapprezzo
const dannoSovrapprezzo = parseFloat(row['Danno Sovrapprezzo']) || 0;

// Check date condition once and accumulate totals
const dataAcquisto = moment(row.dataacquisto, "DD/MM/YYYY");
if (dataAcquisto.isSameOrBefore(cartelloDate, 'day')) {
    dannoCartelloTotal += dannoSovrapprezzo;
} else {
    dannoLingeringTotal += dannoSovrapprezzo;
}
});

// Update DOM after all calculations to minimize reflows
document.getElementById('fatturato').textContent = `€${fatturato.toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
document.getElementById('conteggio').textContent = uniqueVehicles.size;  // Count unique vehicles
document.getElementById('dannoCartello').textContent = `Danno Sovrapprezzo (Periodo Cartello): €${dannoCartelloTotal.toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
document.getElementById('dannoLingering').textContent = `Danno Sovrapprezzo (Periodo Lingering): €${dannoLingeringTotal.toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}



function getSelectedValues(elementId) {
    const select = document.getElementById(elementId);
    return Array.from(select.selectedOptions).map(option => option.value);
}

function displayData() {
    const tableHeader = document.getElementById('tableHeader');
    const tableBody = document.getElementById('tableBody');
    tableHeader.innerHTML = '';
    tableBody.innerHTML = '';

    if (filteredData.length === 0) return;

    // Build table headers
    Object.keys(filteredData[0]).forEach(key => {
        const th = document.createElement('th');
        th.textContent = key;
        tableHeader.appendChild(th);
    });

    // Calculate total pages and update pagination
    const totalPages = Math.ceil(filteredData.length / rowsPerPage);
    document.getElementById('totalPages').textContent = totalPages;
    document.getElementById('pageInput').value = currentPage;

    document.getElementById('prevPage').disabled = currentPage === 1;
    document.getElementById('nextPage').disabled = currentPage === totalPages;

    // Get the current page data
    const startIdx = (currentPage - 1) * rowsPerPage;
    const endIdx = startIdx + rowsPerPage;
    const pageData = filteredData.slice(startIdx, endIdx);

    // Use DocumentFragment for efficient DOM manipulation
    const fragment = document.createDocumentFragment();
    pageData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(value => {
            const td = document.createElement('td');
            td.textContent = value;
            tr.appendChild(td);
        });
        fragment.appendChild(tr);
    });

    // Append all rows at once
    tableBody.appendChild(fragment);
}


document.getElementById('prevPage').addEventListener('click', () => {
    if (currentPage > 1) {
        currentPage--;
        displayData();
        document.getElementById('pageInput').value = currentPage;  // Update input field
    }
});

document.getElementById('nextPage').addEventListener('click', () => {
    const totalPages = Math.ceil(filteredData.length / rowsPerPage);
    if (currentPage < totalPages) {
        currentPage++;
        displayData();
        document.getElementById('pageInput').value = currentPage;  // Update input field
    }
});

function goToPage() {
    const pageInput = document.getElementById('pageInput');
    const totalPages = Math.ceil(filteredData.length / rowsPerPage);
    let requestedPage = parseInt(pageInput.value);

    if (requestedPage < 1 || requestedPage > totalPages || isNaN(requestedPage)) {
        pageInput.value = currentPage;  // Reset to current page if invalid input
    } else {
        currentPage = requestedPage;
        displayData();
    }
}

function exportTableToExcel() {
    // Ensure filteredData exists and is not empty
    if (!filteredData || filteredData.length === 0) {
        alert('No data to export!');
        return;
    }

    // Step 1: Extract relevant columns
    const extractedData = filteredData.map(row => ({
        impresa: row['impresa'],
        targa: row['targa'],
        nuovousato: row['nuovousato'],
        DannoSovrapprezzo: parseFloat(row['Danno Sovrapprezzo']) || 0 // Ensure numeric value
    }));

    // Step 2: Collapse rows by summing 'Danno Sovrapprezzo'
    const collapsedData = {};
    extractedData.forEach(row => {
        const key = `${row.impresa}-${row.targa}-${row.nuovousato}`;
        if (!collapsedData[key]) {
            collapsedData[key] = { ...row }; // Initialize group
        } else {
            collapsedData[key].DannoSovrapprezzo += row.DannoSovrapprezzo; // Sum the values
        }
    });

    // Convert collapsedData back to an array
    const outputData = Object.values(collapsedData);

    // Step 3: Export to Excel
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(outputData);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');
    XLSX.writeFile(workbook, 'TruckDataset_Collapsed.xlsx');
}