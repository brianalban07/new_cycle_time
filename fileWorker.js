importScripts('https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js');

self.onmessage = function(e) {

  const {buffers} = e.data;
  let results = [];

  buffers.forEach(buffer => {

    const wb = XLSX.read(buffer, {type:'array'});
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);

    let dur={}, cnt={};

    json.forEach(r=>{
      const s=r.Stage;
      const st=new Date(r.StartTime);
      const et=new Date(r.EndTime);

      if(s && !isNaN(st) && !isNaN(et)){
        const d=(et-st)/1000;

        if(!dur[s]){dur[s]=0;cnt[s]=0;}
        dur[s]+=d;
        cnt[s]++;
      }
    });

    results.push({dur,cnt});
  });

  self.postMessage({results});
};
