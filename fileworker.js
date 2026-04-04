importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js');

self.onmessage = function(event) {
    try {
        const { file, productType } = event.data;
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                let aggregatedData = [];
                workbook.SheetNames.forEach(sheetName => {
                    const sheet = workbook.Sheets[sheetName];
                    const parsedData = XLSX.utils.sheet_to_json(sheet, { defval: '' });
                    aggregatedData = aggregatedData.concat(parsedData);
                });

                const stageOrder = getStageOrder(productType);
                const summaryData = computeTestingTimes(aggregatedData, stageOrder);

                self.postMessage({ fileName: file.name, summaryData });
            } catch (error) {
                self.postMessage({ error: error.message, fileName: file.name });
            }
        };

        reader.onerror = function() {
            self.postMessage({ error: 'Failed to read file', fileName: file.name });
        };

        reader.readAsArrayBuffer(file);
    } catch (error) {
        self.postMessage({ error: error.message });
    }
};

function getStageOrder(productType) {
    const stageOrders = {
        "PG548": ["INIT", "PRE-CHECK", "AST", "FLA", "IOT", "FCT", "FC2", "IST", "FPF", "NVL"],
        "Umbriel": ["INIT", "FLT", "FLA", "FLB", "FCT", "FLC", "FTS", "RIN", "DCC", "IOT"],
        "Vulcan FC2": ["INIT", "FLT", "FLB", "FCT", "DCC", "DCT", "RIN", "EBT"],
        "Umbriel VikingFru": ["VikingFruINIT", "INIT", "FLA", "FTS", "FCT", "FCT-1", "FCT1", "RIN", "DCC", "FCT-2", "FCT2"],
        "Viking": ["INIT", "PT1", "PT2", "FLA", "FL1", "FLB", "FCT", "RIN", "DCC", "FIN", "INSTALLATION"],
                "Gaines": ["INIT", "FLA", "FLC", "FCT", "FINT", "RIN", "NVL", "INSTALLATION"],
                "Skywalker": ["INIT", "FLA", "FLB", "FLC", "FCT", "FINT", "NVL", "INSTALLATION"],
    };
    return stageOrders[productType] || [];
}

function parseExcelDates(excelDate) {
    if (!isNaN(excelDate) && excelDate > 59) {
        const utc_days = Math.floor(excelDate - 25569);
        const utc_value = utc_days * 86400;
        return new Date(utc_value * 1000);
    } else {
        const parsedDate = moment(excelDate, [
            'YYYY-MM-DD HH:mm:ss', 'M/D/YYYY h:mm:ss A', 'M/D/YYYY hh:mm:ss A',
            'M/D/YYYY h:mm A', 'M/D/YYYY hh:mm A', 'M/D/YYYY HH:mm',
            'M/D/YYYY HH:mm:ss', 'M-D-YYYY HH:mm', 'M-D-YYYY HH:mm:ss',
            'YYYY/MM/DD HH:mm:ss', 'YYYY.MM.DD HH:mm:ss'
        ], true);
        return parsedDate.isValid() ? parsedDate.toDate() : null;
    }
}

function computeTestingTimes(data, stageOrder) {
    const skuStageTimes = {};
    data.forEach(row => {
        const sku = row["SkuName"];
        const stage = row["Stage"];
        if (stageOrder.includes(stage)) {
            const startTime = parseExcelDates(row.StartTime);
            const endTime = parseExcelDates(row.EndTime);
            if (!startTime || !endTime) return;
            const diff = moment(endTime).diff(moment(startTime));
            if (!skuStageTimes[sku]) {
                skuStageTimes[sku] = { total: 0, counts: {} };
            }
            if (!skuStageTimes[sku][stage]) {
                skuStageTimes[sku][stage] = 0;
                skuStageTimes[sku].counts[stage] = 0;
            }
            skuStageTimes[sku][stage] += diff / 60000; // Convert to minutes
            skuStageTimes[sku].counts[stage]++;
        }
    });

    const result = [];
    for (const sku in skuStageTimes) {
        const row = { sku, Total: 0, "Total Minutes": 0 };
        stageOrder.forEach(stage => {
            if (skuStageTimes[sku][stage] !== undefined) {
                const averageTime = skuStageTimes[sku][stage] / skuStageTimes[sku].counts[stage];
                row[stage] = convertMillisToTime(averageTime * 60000);
                row.Total += averageTime;
                row["Total Minutes"] += averageTime;
            } else {
                row[stage] = null;
            }
        });
        row.Total = convertMillisToTime(row.Total * 60000);
        row["Total Minutes"] = Math.round(row["Total Minutes"]);  // Ensure rounding here
        result.push(row);
    }
    return result;
}

function convertMillisToTime(millis) {
    const hours = Math.floor(millis / (1000 * 60 * 60));
    const minutes = Math.floor((millis % (1000 * 60 * 60)) / (1000 * 60));
    const seconds = Math.floor((millis % (1000 * 60)) / 1000);
    return `${hours}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
}
