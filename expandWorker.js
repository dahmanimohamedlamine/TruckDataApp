// Import moment.js for the worker
importScripts('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.4/moment.min.js');

onmessage = function(event) {
    const { data, endDateStr } = event.data;
    const endDate = moment(endDateStr, "YYYY-MM-DD");
    let expandedData = [];

    try {
        data.forEach(row => {
            const startDate = moment(row.mese_acquisto, "DD/MM/YYYY");
            const dataFineDate = moment(row.data_fine, "DD/MM/YYYY");
            let firstMonth = true;

            if (startDate.isValid() && endDate.isValid() && startDate.isBefore(endDate)) {
                let currentDate = startDate.clone();

                while (currentDate.isSameOrBefore(endDate, 'month')) {
                    let expandedRow = { ...row };
                    expandedRow.mese = currentDate.format("MM/YYYY");

                    // Clear price fields for non-leasing contracts after the first month
                    if (row.acquistoleasing !== "LEASING" && !firstMonth) {
                        Object.keys(expandedRow).forEach(key => {
                            if (key.toLowerCase().includes("prezzo") || key.toLowerCase().includes("price")) {
                                expandedRow[key] = "";
                            }
                        });
                    }

                    // Clear price fields for leasing contracts after data_fine date
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
        });

        postMessage(expandedData);
    } catch (error) {
        postMessage({ error: error.message });
    }
};