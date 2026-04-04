importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
importScripts('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.1/moment.min.js');

self.onmessage = function (e) {
    const { file, productType } = e.data;

    const reader = new FileReaderSync();
    try {
        const data = new Uint8Array(reader.readAsArrayBuffer(file));
        const workbook = XLSX.read(data, { type: 'array' });

        let allData = [];
        workbook.SheetNames.forEach(sheetName => {
            const parsedSheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: null });
            allData = allData.concat(parsedSheet);
        });

        const summaryData = calculateSummary(allData, productType);
        self.postMessage({ fileName: file.name, summaryData });
    } catch (error) {
        self.postMessage({ fileName: file.name, error: error.message });
    }
};

function calculateSummary(data, productType) {
    const stageOrder = getStageOrder(productType);

    const skuTimes = {};
    data.forEach(row => {
        const sku = row['SkuName'];
        const stage = row['Stage'];
        const startTime = parseDate(row['StartTime']);
        const endTime = parseDate(row['EndTime']);

        if (!startTime || !endTime || !stageOrder.includes(stage)) return;

        const duration = moment(endTime).diff(moment(startTime), 'seconds');

        if (!skuTimes[sku]) skuTimes[sku] = {};
        if (!skuTimes[sku][stage]) skuTimes[sku][stage] = 0;

        skuTimes[sku][stage] += duration;
        skuTimes[sku].Total = (skuTimes[sku].Total || 0) + duration;
    });

    const summary = [];
    for (const sku in skuTimes) {
        const row = { SkuName: sku };
        stageOrder.forEach(stage => {
            const time = skuTimes[sku][stage];
            row[stage] = time ? formatSecondsToTime(time) : 'N/A';
        });
        row.Total = formatSecondsToTime(skuTimes[sku].Total);
        summary.push(row);
    }
    return summary;
}

function parseDate(excelDate) {
    return moment(excelDate, [
        moment.ISO_8601,
        'YYYY-MM-DD HH:mm:ss',
        'DD/MM/YYYY HH:mm:ss',
        'M/D/YYYY h:mm A',
    ]).isValid()
        ? moment.utc(excelDate).toDate()
        : null;
}

function formatSecondsToTime(seconds) {
    const h = Math.floor(seconds / 3600);
    const m = Math.floor((seconds % 3600) / 60);
    const s = Math.floor(seconds % 60);
    return `${h}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
}

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
