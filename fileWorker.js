importScripts('https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js');

self.onmessage = function(e) {

  const {buffers} = e.data;
  let results = [];

  buffers.forEach(buffer => {

    const wb = XLSX.read(buffer, {type:'array'});

    let dur = {};
    let cnt = {};

    // ✅ LOOP ALL SHEETS
    wb.SheetNames.forEach(sheetName => {

      const sheet = wb.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet);

      json.forEach(r => {

        const stage = r.Stage;
        const st = new Date(r.StartTime);
        const et = new Date(r.EndTime);

        if (stage && !isNaN(st) && !isNaN(et)) {

          const d = (et - st)/1000;

          if (!dur[stage]) {
            dur[stage] = 0;
            cnt[stage] = 0;
          }

          dur[stage] += d;
          cnt[stage]++;
        }

      });

    });

    // ✅ push combined result (all sheets)
    results.push({dur, cnt});
  });

  self.postMessage({results});
};
