function download(filename, text) {
    let element = document.createElement('a');
    element.setAttribute('href', 'data:csv/plain;charset=utf-8,' + encodeURIComponent(text));
    element.setAttribute('download', filename);
    element.style.display = 'none';

    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
}


var X = XLSX;
var XW = {
	msg: 'xlsx',
	worker: './xlsxworker.js'
};

var global_wb;

function process_wb(wb) {
    if (!wb) return;

    const start_at_row = 6, max_passed = 100;
    const sheet = wb.Sheets[wb.SheetNames[0]];

    let cell = null, next_cell = null, category = null, row_index = start_at_row, passed = 0;
    let list = [["CATEGORIE", "VOLNAAM", "BRUTOPRIJS"]];

    while (passed < max_passed) {
        cell = sheet[X.utils.encode_cell({c:0, r:row_index})];

        row_index++;
        if (!cell) {passed++; continue}
        passed = 0;

        next_cell = sheet[X.utils.encode_cell({c:0, r:row_index})];

        if (next_cell) { // cell is category
            category = cell.v;
            continue
        } else { // cell is a product group
            try {
                while (true) {
                    let product_cell = sheet[X.utils.encode_cell({c:1, r:row_index})];
                    if (!product_cell) break;
                    let price_cell = sheet[X.utils.encode_cell({c:2, r:row_index})];
                    let price =  (price_cell && !isNaN(price_cell.v)) ? price_cell.v : 0;
                    list.push([category, product_cell.v, price.toFixed(2)]);
                    row_index++;
                }
            } catch (err) {
                console.log({loc: row_index, cat: category, cell: cell, next_cell: next_cell,
                    r1: sheet[X.utils.encode_cell({c:1, r:row_index})],
                    r2: sheet[X.utils.encode_cell({c:2, r:row_index})],
                })
            }
        }
    }

    download("price_list_export.csv", X.utils.sheet_to_csv(X.utils.aoa_to_sheet(list), {FS: ';'}))
}

var b64it = window.b64it = (function() {
	var tarea = document.getElementById('b64data');
	return function b64it() {
		if(typeof console !== 'undefined') console.log("onload", new Date());
		var wb = X.read(tarea.value, {type:'base64', WTF:false});
		process_wb(wb);
	};
})();

var do_file = (function() {
	var rABS = typeof FileReader !== "undefined" && (FileReader.prototype||{}).readAsBinaryString;
	var domrabs = document.getElementsByName("userabs")[0];
	if(!rABS) domrabs.disabled = !(domrabs.checked = false);

	var use_worker = typeof Worker !== 'undefined';
	var domwork = document.getElementsByName("useworker")[0];
	if(!use_worker) domwork.disabled = !(domwork.checked = false);

	var xw = function xw(data, cb) {
		var worker = new Worker(XW.worker);
		worker.onmessage = function(e) {
			switch(e.data.t) {
				case 'ready': break;
				case 'e': console.error(e.data.d); break;
				case XW.msg: cb(JSON.parse(e.data.d)); break;
			}
		};
		worker.postMessage({d:data,b:rABS?'binary':'array'});
	};

	return function do_file(files) {
		rABS = true;
		use_worker = true;
		var f = files[0];
		var reader = new FileReader();
		reader.onload = function(e) {
			console.log("onload", new Date(), rABS, use_worker);
			var data = e.target.result;
			if(use_worker) xw(data, process_wb);
			else process_wb(X.read(data, {type: rABS ? 'binary' : 'array'}));
		};
		if(rABS) reader.readAsBinaryString(f);
	};
})();

const drop = document.getElementById('drop');
function handleDrop(e) {
    e.stopPropagation();
    e.preventDefault();
    do_file(e.dataTransfer.files);
}

function handleDragover(e) {
    e.stopPropagation();
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
}

drop.addEventListener('dragenter', handleDragover, false);
drop.addEventListener('dragover', handleDragover, false);
drop.addEventListener('drop', handleDrop, false);

const xlf = document.getElementById('xlf');
xlf.addEventListener('change', (e) => do_file(e.target.files), false);
