importScripts("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.16.6/shim.min.js",
              "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.16.6/xlsx.full.min.js");
postMessage({t: "ready"});

onmessage = function (evt) {
    try {
        let v = XLSX.read(evt.data.d, {type: evt.data.b});
        postMessage({t: "xlsx", d: JSON.stringify(v)});
    } catch (e) {
        postMessage({t: "e", d: e.stack || e});
    }
};
