let data = {}; 
let initialData = [];  // Store the original unexpanded data
let currentData = []; 
let filteredData = []; 
let currentPage = 1; 
const rowsPerPage = 10; 
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
        .filter(name => name !== 'TEGM' && name !== 'Tasso Legale' && name !== 'FOI') // Exclude these sheets
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

document.addEventListener('DOMContentLoaded', () => {
    // Add event listener to the Periodo Lingering dropdown
    const periodoLingeringDropdown = document.getElementById('periodoLingering');
    if (periodoLingeringDropdown) {
        periodoLingeringDropdown.addEventListener('change', () => {
            const selectedSheetName = document.getElementById('causa').value || "DefaultSheetName"; // Fallback to default if no sheet is selected
            loadSheetData(selectedSheetName); // Call loadSheetData dynamically
        });
    }
});


function loadSheetData(sheetName) {
    // Display the loading bar
    const loadingBarContainer = document.getElementById('loadingBarContainer');
    const loadingLabel = document.getElementById('loadingLabel');
    const loadingBar = document.getElementById('loadingBar');
    loadingBarContainer.style.display = 'block';
    loadingLabel.innerText = "Loading data...";
    loadingBar.style.width = "0%"; // Initialize progress bar

    let progress = 0; // Initialize progress percentage

    // Function to update progress
    const updateProgress = (stepProgress, message) => {
        progress += stepProgress;
        if (progress > 100) progress = 100; // Cap at 100%
        loadingBar.style.width = `${progress}%`;
        loadingLabel.innerText = message;
    };

    let tegmData = []; // Define tegmData outside the if block for wider scope

    // Load and transform the FOI data once
    if (data['FOI']) {
        updateProgress(10, "Transforming FOI data...");
        console.log("FOI Data Before Transformation:", data['FOI']);
        foiData = data['FOI'].map(row => ({ ...row })); // Copy FOI data
        transformDateColumns(foiData); // Transform date columns

        // Normalize 'mese' column in FOI data
        foiData = foiData.map(row => ({
            ...row,
            mese: row.mese ? moment(row.mese, "DD/MM/YYYY").format("MM/YYYY").trim().toLowerCase() : '' // Normalize FOI 'mese'
        }));

        console.log("FOI Data After Transformation:", foiData);
    }

    // Load and transform the Tasso Legale data once
    if (data['Tasso Legale']) {
        updateProgress(10, "Processing Tasso Legale data...");
        console.log("Tasso Legale Data Before Transformation:", data['Tasso Legale']);
        tassoLegaleData = data['Tasso Legale'].map(row => ({ ...row })); // Copy Tasso Legale data
    }

    // Check if TEGM sheet is present and load it if so
    if (data['TEGM']) {
        updateProgress(10, "Processing TEGM data...");
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
    updateProgress(20, "Loading and transforming initial data...");
    initialData = data[sheetName].map(row => {
        delete row.anno_acquisto; // Remove existing 'anno_acquisto' if it exists
        return {
            ...row,
            marca: row.marca ? row.marca.toUpperCase() : '', // Transform `marca` to uppercase
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
        updateProgress(20, "Matching initial data with TEGM data...");
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

    // Apply filtering based on 'Periodo Lingering (Anni)'
    updateProgress(10, "Applying filtering...");
    const periodoLingering = parseInt(document.getElementById('periodoLingering').value, 10);

    if (periodoLingering !== 3) {
        const cutoffDates = {
            0: "18/01/2011",
            1: "31/12/2011",
            2: "31/12/2012"
        };

        const cutoffDate = moment(cutoffDates[periodoLingering], "DD/MM/YYYY");

        // Filter out rows with dataacquisto beyond the cutoff date
        initialData = initialData.filter(row => {
            const dataAcquistoDate = row.dataacquisto ? moment(row.dataacquisto, "DD/MM/YYYY") : null;
            return dataAcquistoDate && dataAcquistoDate.isSameOrBefore(cutoffDate);
        });
    }

    // Clear `tegmData` as it is no longer needed
    tegmData = null;

    updateProgress(20, "Calculating TEGM rates...");
    // Define `tegm` based on `Prezzo Netto` ranges
    initialData = initialData.map(row => {
        let tegmValue;

        // Determine tegm based on Prezzo Netto ranges
        if (parseFloat(row['prezzo_netto']) <= 5000) {
            tegmValue = row.tegm0_5000 / 100;
        } else if (parseFloat(row['prezzo_netto']) > 5000 && parseFloat(row['prezzo_netto']) <= 25000) {
            tegmValue = row.tegm5000_25000 / 100;
        } else if (parseFloat(row['prezzo_netto']) > 25000 && parseFloat(row['prezzo_netto']) <= 50000) {
            tegmValue = row.tegm25000_50000 / 100;
        } else if (parseFloat(row['prezzo_netto']) > 50000) {
            tegmValue = row.tegmoltre_50000 / 100;
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
            if (key.startsWith('tegm') || key === 'qtr' || key === 'quarter') {
                delete row[key];
            }
        });
        return row;
    });

    isExpanded = false;
    currentData = [...initialData]; // Set currentData to the processed initialData
    populateFilters(currentData);

    updateProgress(30, "Processing Data...");
    expandTable();

    // Finalize progress and hide the loading bar
    setTimeout(() => {
        loadingBar.style.width = "100%";
        loadingLabel.innerText = "Danno aggiornato!";
        setTimeout(() => {
            loadingBarContainer.style.display = 'none';
        }, 500); // Small delay to show success message
    }, 500);
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
    console.log("Initial Data:", initialData);
    // Reset currentData to initialData at the beginning
    currentData = [...initialData];
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
            expandedData=null;
        }
    }

    processBatch();
}



function ProcessData() {
    const selectedPrezzo = document.getElementById('prezzo').value;
    const sovrapprezzoCartello = parseFloat(document.getElementById('sovrapprezzoCartello').value) || 0;
    const sovrapprezzoLingering = parseFloat(document.getElementById('sovrapprezzoLingering').value) || 0;

    const cumulativeQuotaPerVehicle = {};

        // Match FOI data with currentData
    currentData = currentData.map(row => {
        const monthKey = row.mese ? row.mese.trim().toLowerCase() : '';
        const foiMatch = foiData.find(foiRow => foiRow.mese === monthKey);

        return foiMatch ? { ...row, ...foiMatch } : row; // Merge FOI data
    });

        // Add `anno` column to currentData
    currentData = currentData.map(row => ({
        ...row,
        anno: row.mese ? moment(row.mese, "MM/YYYY").year() : null // Extract year from `mese`
    }));

    // Match Tasso Legale data with currentData
    currentData = currentData.map(row => {
        const annoKey = row.anno; // Use the `anno` column for matching
        const tassoLegaleMatch = tassoLegaleData.find(tlRow => tlRow.anno === annoKey);

        return tassoLegaleMatch ? { ...row, ...tassoLegaleMatch } : row; // Merge Tasso Legale data
    });

    // Map and process the filtered data
    currentData = currentData.map(item => {
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

        if (row.acquistoleasing !== "LEASING") {
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

            if (row.acquistoleasing === "LEASING" && durataResidua === 0) {
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
    // Compute Danno Totale for all rows in currentData
    currentData = currentData.map(row => {
        // Directly assign Danno Totale to Danno Sovrapprezzo
        row['Danno Totale'] = row['Danno Sovrapprezzo'] || 0;
        return row;
    });
        // Add Cumulative Sum for Danno Totale
    const groupedData = {}; // Temporary storage for grouping
    currentData.forEach(row => {
        const groupKey = `${row.impresa}-${row.targa}-${row.nuovousato}`;
        if (!groupedData[groupKey]) groupedData[groupKey] = 0;
        groupedData[groupKey] += row['Danno Totale'] || 0;
        row['danno_cumulato'] = groupedData[groupKey];
    });

    // Add Cumulative Inflation and Danno Rivalutato
    const inflationData = {}; // Temporary storage for cumulative inflation
    currentData.forEach(row => {
        const groupKey = `${row.impresa}-${row.targa}-${row.nuovousato}`;
        if (!inflationData[groupKey]) inflationData[groupKey] = 1; // Start inflation factor at 1

        // Calculate cumulative inflation
        inflationData[groupKey] *= row['foi'] ? parseFloat(row['foi']) : 1;
        row['cumulative_inflation'] = inflationData[groupKey];

        // Calculate Danno Rivalutato
        row['Danno rivalutato'] = row['danno_cumulato'] * row['cumulative_inflation'];
    });


// Add calculations for `Danno rivalutato`, `Interessi legali`, and WACC-based variables
    const cumulativeWACC = {}; // To store cumulative sum of `Interessi legali WACC` for each group

    currentData = currentData.map(row => {
        // Calculate `Interessi legali`
        const interessiLegali = row['Danno rivalutato'] && row['tassolegale_mens']
            ? row['Danno rivalutato'] * parseFloat(row['tassolegale_mens'])
            : 0;

        // Calculate `Interessi legali WACC`
        const interessiLegaliWACC = row['danno_cumulato'] && row['Wacc (tutti mensile)']
            ? row['danno_cumulato'] * parseFloat(row['Wacc (tutti mensile)'])
            : 0;

        // Track cumulative WACC interest by group
        const groupKey = `${row.impresa}-${row.targa}-${row.nuovousato}`;
        if (!cumulativeWACC[groupKey]) cumulativeWACC[groupKey] = 0;
        cumulativeWACC[groupKey] += interessiLegaliWACC;

        // Calculate `Danno rivalutato WACC`
        const dannoRivalutatoWACC = row['danno_cumulato']
            ? row['danno_cumulato'] + cumulativeWACC[groupKey]
            : row['danno_cumulato'];

        return {
            ...row,
            'Interessi legali': interessiLegali,
            'Interessi legali WACC': interessiLegaliWACC,
            'Danno rivalutato WACC': dannoRivalutatoWACC
        };
    });
    // Reapply filters and update the UI
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
        document.getElementById('dannoTotale').textContent = '€0,00';
        document.getElementById('interessiLegaliTotale').textContent = '€0,00';
        document.getElementById('interessiLegaliWACCTotale').textContent = '€0,00';
        document.getElementById('dannoRivalutatoTotale').textContent = '€0,00';
        document.getElementById('dannoRivalutatoWACCTotale').textContent = '€0,00';
        return;
    }

    let fatturato = 0;
    let dannoCartelloTotal = 0;
    let dannoLingeringTotal = 0;
    let interessiLegaliSum = 0;
    let interessiLegaliWACCSum = 0;
    let dannoRivalutatoSum = 0;
    let dannoRivalutatoWACCSum = 0;

    const uniqueVehicles = new Map(); // To track unique vehicles
    const groupedLastRows = {}; // To store last rows per group
    const cartelloDate = moment("18/01/2011", "DD/MM/YYYY");

    // Loop through filteredData
    filteredData.forEach(row => {
        const uniqueKey = `${row.impresa}-${row.targa}-${row.nuovousato}`;
        const mese = moment(row.mese, "MM/YYYY");

        // Update uniqueVehicles and calculate fatturato
        if (!uniqueVehicles.has(uniqueKey)) {
            uniqueVehicles.set(uniqueKey, true);
            const prezzoValue = parseFloat(row[prezzoColumn]) || 0;
            fatturato += prezzoValue;
        }
        // Accumulate `Interessi legali`
        if ('Interessi legali' in row) {
            interessiLegaliSum += parseFloat(row['Interessi legali']) || 0;
        }

        // Accumulate `Interessi legali WACC`
        if ('Interessi legali WACC' in row) {
            interessiLegaliWACCSum += parseFloat(row['Interessi legali WACC']) || 0;
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

        // Track the last row by group
        if (!groupedLastRows[uniqueKey] || mese.isAfter(moment(groupedLastRows[uniqueKey].mese, "MM/YYYY"))) {
            groupedLastRows[uniqueKey] = row;
        }
    });

    // Calculate Danno Sovrapprezzo Totale
    const dannoSovrapprezzoTotale = dannoCartelloTotal + dannoLingeringTotal;

    // Extract last rows and calculate totals for other metrics
    Object.values(groupedLastRows).forEach(row => {
        // Calculate Danno Rivalutato
        if ('Danno rivalutato' in row) {
            dannoRivalutatoSum += parseFloat(row['Danno rivalutato']) || 0;
        }

        // Calculate Danno Rivalutato WACC
        if ('Danno rivalutato WACC' in row) {
            dannoRivalutatoWACCSum += parseFloat(row['Danno rivalutato WACC']) || 0;
        }


    });
     // Add Interessi Legali to Danno Rivalutato Totale
    dannoRivalutatoSum += interessiLegaliSum;
    // Update DOM with calculated values
    document.getElementById('fatturato').textContent = `€${fatturato.toLocaleString('it-IT', { minimumFractionDigits: 2 })}`;
    document.getElementById('conteggio').textContent = uniqueVehicles.size; // Count unique vehicles
    document.getElementById('dannoCartello').textContent = `€${dannoCartelloTotal.toLocaleString('it-IT', { minimumFractionDigits: 2 })}`;
    document.getElementById('dannoLingering').textContent = `€${dannoLingeringTotal.toLocaleString('it-IT', { minimumFractionDigits: 2 })}`;
    document.getElementById('dannoTotale').textContent = `€${dannoSovrapprezzoTotale.toLocaleString('it-IT', { minimumFractionDigits: 2 })}`;
    document.getElementById('interessiLegaliTotale').textContent = `€${interessiLegaliSum.toLocaleString('it-IT', { minimumFractionDigits: 2 })}`;
    document.getElementById('interessiLegaliWACCTotale').textContent = `€${interessiLegaliWACCSum.toLocaleString('it-IT', { minimumFractionDigits: 2 })}`;
    document.getElementById('dannoRivalutatoTotale').textContent = `€${dannoRivalutatoSum.toLocaleString('it-IT', { minimumFractionDigits: 2 })}`;
    document.getElementById('dannoRivalutatoWACCTotale').textContent = `€${dannoRivalutatoWACCSum.toLocaleString('it-IT', { minimumFractionDigits: 2 })}`;
}


function getSelectedValues(elementId) {
    const select = document.getElementById(elementId);
    return Array.from(select.selectedOptions).map(option => option.value);
}

function getGroupedExportTableData() {
    const groupedData = {};

    filteredData.forEach(row => {
        const uniqueKey = `${row['impresa']}-${row['targa']}-${row['nuovousato']}`;
        if (!groupedData[uniqueKey]) groupedData[uniqueKey] = [];

        groupedData[uniqueKey].push({
            ...row,
            mese: moment(row['mese'], "MM/YYYY")
        });
    });

    return Object.keys(groupedData).map(key => {
        const rows = groupedData[key];
        rows.sort((a, b) => a.mese - b.mese);

        const summedRow = rows.reduce((acc, row) => {
            acc["Danno Sovrapprezzo"] += parseFloat(row['Danno Sovrapprezzo']) || 0;
            acc["Danno Totale"] += parseFloat(row['Danno Totale']) || 0;
            acc["Interessi legali"] += parseFloat(row['Interessi legali']) || 0;
            acc["Interessi legali WACC"] += parseFloat(row['Interessi legali WACC']) || 0;
            acc.maxPrezzoNetto = Math.max(acc.maxPrezzoNetto, parseFloat(row['prezzo_netto']) || 0);
            return acc;
        }, {
            impresa: rows[0]['impresa'] || '',
            targa: rows[0]['targa'] || '',
            acquistoleasing: rows[0]['acquistoleasing'] || '',
            nuovousato: rows[0]['nuovousato'] || '',
            dataacquisto: rows[0]['dataacquisto'] || '',
            statoattuale: rows[0]['statoattuale'] || '',
            "Danno Sovrapprezzo": 0,
            "Danno Totale": 0,
            "Interessi legali": 0,
            "Interessi legali WACC": 0,
            maxPrezzoNetto: 0
        });

        const lastRow = rows[rows.length - 1];
        summedRow.prezzo_netto = summedRow.maxPrezzoNetto;
        summedRow["Danno rivalutato"] = parseFloat(lastRow['Danno rivalutato']) || 0;
        summedRow["Danno rivalutato WACC"] = parseFloat(lastRow['Danno rivalutato WACC']) || 0;
        delete summedRow.maxPrezzoNetto;

        return {
            impresa: summedRow.impresa,
            targa: summedRow.targa,
            acquistoleasing: summedRow.acquistoleasing,
            nuovousato: summedRow.nuovousato,
            dataacquisto: summedRow.dataacquisto,
            statoattuale: summedRow.statoattuale,
            prezzo_netto: summedRow.prezzo_netto,
            "Danno Sovrapprezzo": summedRow["Danno Sovrapprezzo"],
            "Danno Totale": summedRow["Danno Totale"],
            "Interessi legali": summedRow["Interessi legali"],
            "Interessi legali WACC": summedRow["Interessi legali WACC"],
            "Danno rivalutato": summedRow["Danno rivalutato"],
            "Danno rivalutato WACC": summedRow["Danno rivalutato WACC"]
        };
    });
}

const COLUMN_LABELS = {
    impresa: "Impresa",
    targa: "Targa",
    acquistoleasing: "Tipo Acquisto",
    nuovousato: "Condizione",
    dataacquisto: "Data Acquisto",
    statoattuale: "Stato Attuale",
    prezzo_netto: "Prezzo Netto (€)",
    "Danno Sovrapprezzo": "Danno Sovrapprezzo (€)",
    "Danno Totale": "Danno Totale (€)",
    "Interessi legali": "Interessi Legali (€)",
    "Interessi legali WACC": "Interessi Legali WACC (€)",
    "Danno rivalutato": "Danno Rivalutato (€)",
    "Danno rivalutato WACC": "Danno Rivalutato WACC (€)"
};

const EURO_COLUMNS = [
    'prezzo_netto',
    'Danno Sovrapprezzo',
    'Danno Totale',
    'Interessi legali',
    'Interessi legali WACC',
    'Danno rivalutato',
    'Danno rivalutato WACC'
];

function displayData() {
    const tableHeader = document.getElementById('tableHeader');
    const tableBody = document.getElementById('tableBody');
    tableHeader.innerHTML = '';
    tableBody.innerHTML = '';

    const groupedExportData = getGroupedExportTableData();
    const searchValue = document.getElementById('tableSearchInput')?.value?.toLowerCase() || '';

    if (groupedExportData.length === 0) return;

    const headers = Object.keys(groupedExportData[0]);

    // Filter rows using search value
    const filteredTableData = groupedExportData.filter(row => {
        return Object.values(row).some(val =>
            String(val).toLowerCase().includes(searchValue)
        );
    });

    // Render column headers with friendly labels
    headers.forEach(key => {
        const th = document.createElement('th');
        th.textContent = COLUMN_LABELS[key] || key;
        tableHeader.appendChild(th);
    });

    const totalPages = Math.ceil(filteredTableData.length / rowsPerPage);
    document.getElementById('totalPages').textContent = totalPages;
    document.getElementById('pageInput').value = currentPage;

    document.getElementById('prevPage').disabled = currentPage === 1;
    document.getElementById('nextPage').disabled = currentPage === totalPages;

    const startIdx = (currentPage - 1) * rowsPerPage;
    const endIdx = startIdx + rowsPerPage;
    const pageData = filteredTableData.slice(startIdx, endIdx);

    const fragment = document.createDocumentFragment();

    // Render table rows
    pageData.forEach(row => {
        const tr = document.createElement('tr');

        headers.forEach(key => {
            const td = document.createElement('td');
            let value = row[key];

            if (EURO_COLUMNS.includes(key) && typeof value === 'number') {
                value = `€ ${value.toLocaleString('it-IT', {
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2
                })}`;
                td.classList.add('euro'); // Optional: apply right alignment
            }

            td.textContent = value !== undefined ? value : '';
            tr.appendChild(td);
        });

        fragment.appendChild(tr);
    });

    tableBody.appendChild(fragment);
}

document.getElementById('tableSearchInput')?.addEventListener('input', () => {
    currentPage = 1;
    displayData();
});


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

    // Track grouped data
    const groupedData = {};

    // Iterate through filteredData to group by unique key
    filteredData.forEach(row => {
        const uniqueKey = `${row['impresa']}-${row['targa']}-${row['nuovousato']}`;

        // Initialize the group if it doesn't exist
        if (!groupedData[uniqueKey]) {
            groupedData[uniqueKey] = [];
        }

        // Add row to the group
        groupedData[uniqueKey].push({
            ...row,
            mese: moment(row['mese'], "MM/YYYY") // Convert mese to a Moment.js object
        });
    });

    // Process each group
    const outputData = Object.keys(groupedData).map(key => {
        const rows = groupedData[key];

        // Sort rows by `mese`
        rows.sort((a, b) => a.mese - b.mese);

        // Sum relevant numeric fields and find max prezzo_netto
        const summedRow = rows.reduce(
            (acc, row) => {
                acc["Danno Sovrapprezzo"] += parseFloat(row['Danno Sovrapprezzo']) || 0;
                acc["Danno Totale"] += parseFloat(row['Danno Totale']) || 0;
                acc["Interessi legali"] += parseFloat(row['Interessi legali']) || 0;
                acc["Interessi legali WACC"] += parseFloat(row['Interessi legali WACC']) || 0;
                acc.maxPrezzoNetto = Math.max(acc.maxPrezzoNetto, parseFloat(row['prezzo_netto']) || 0);
                return acc;
            },
            {
                impresa: rows[0]['impresa'] || '',
                targa: rows[0]['targa'] || '',
                acquistoleasing: rows[0]['acquistoleasing'] || '',
                nuovousato: rows[0]['nuovousato'] || '',
                dataacquisto: rows[0]['dataacquisto'] || '',
                statoattuale: rows[0]['statoattuale'] || '',
                "Danno Sovrapprezzo": 0,
                "Danno Totale": 0,
                "Interessi legali": 0,
                "Interessi legali WACC": 0,
                maxPrezzoNetto: 0
            }
        );

        // Get the last row for specific fields
        const lastRow = rows[rows.length - 1];
        summedRow.prezzo_netto = summedRow.maxPrezzoNetto; // Use max prezzo_netto
        summedRow["Danno rivalutato"] = parseFloat(lastRow['Danno rivalutato']) || 0;
        summedRow["Danno rivalutato WACC"] = parseFloat(lastRow['Danno rivalutato WACC']) || 0;

        // Remove intermediate maxPrezzoNetto
        delete summedRow.maxPrezzoNetto;

        return {
            impresa: summedRow.impresa,
            targa: summedRow.targa,
            acquistoleasing: summedRow.acquistoleasing,
            nuovousato: summedRow.nuovousato,
            dataacquisto: summedRow.dataacquisto,
            statoattuale: summedRow.statoattuale,
            prezzo_netto: summedRow.prezzo_netto, // Preceding Danno Sovrapprezzo
            "Danno Sovrapprezzo": summedRow["Danno Sovrapprezzo"],
            "Danno Totale": summedRow["Danno Totale"],
            "Interessi legali": summedRow["Interessi legali"],
            "Interessi legali WACC": summedRow["Interessi legali WACC"],
            "Danno rivalutato": summedRow["Danno rivalutato"],
            "Danno rivalutato WACC": summedRow["Danno rivalutato WACC"]
        };
    });

    // Step 3: Export to Excel
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(outputData);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Dati');
    XLSX.writeFile(workbook, 'Danno_Camion.xlsx');
}


