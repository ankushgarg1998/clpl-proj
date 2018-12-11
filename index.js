var master = {};

addToMaster = function (master_key, el) {
    if (master_key in master) {
        master[master_key]['QTY'] += el['QTY'];
    } else
        master[master_key] = el;
}

ExcelDateToJSDate = function (serial) {
    var utc_days = Math.floor(serial - 25569);
    var utc_value = utc_days * 86400;
    var date_info = new Date(utc_value * 1000);

    var fractional_day = serial - Math.floor(serial) + 0.0000001;

    var total_seconds = Math.floor(86400 * fractional_day);

    var seconds = total_seconds % 60;

    total_seconds -= seconds;

    var hours = Math.floor(total_seconds / (60 * 60));
    var minutes = Math.floor(total_seconds / 60) % 60;

    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
}

DateToString = function (today) {
    var dd = today.getDate();
    var mm = today.getMonth() + 1; //January is 0!

    var yyyy = today.getFullYear();
    if (dd < 10) {
        dd = '0' + dd;
    }
    if (mm < 10) {
        mm = '0' + mm;
    }
    return (mm + ' / ' + yyyy);
}

findWord = function (word, str) {
    return str.split(' ').some(function (w) { return w === word })
}

$('#input-excel').change(function (e) {
    var reader = new FileReader();
    reader.readAsArrayBuffer(e.target.files[0]);

    reader.onload = function (e) {
        var data = new Uint8Array(reader.result);
        var workbook = XLSX.read(data, { type: 'array' });
        var sheet = workbook.Sheets[workbook.SheetNames[0]];
        var sheet_json = XLSX.utils.sheet_to_json(sheet);

        var barcode_column = 'z';
        var date_column = 'z';
        var price_column = 'z';

        var el = sheet_json[0];
        Object.keys(el).forEach(function (key, index) {
            if (findWord('MRP', el))
                price_column = el;
            if (findWord('EAN', el))
                barcode_column = el;
            if (findWord('MFD', el))
                date_column = el;
        });

        if (barcode_column === 'z')
            console.log(`Barcode Column not found. Should have the work "EAN".`);
        if (date_column === 'z')
            console.log(`Date Column not found. Should have the word "MFD".`);
        if (price_column === 'z')
            console.log(`MRP column not found. Should have the word "MRP".`)

        console.log(`barcode_column: ${barcode_column}`);
        console.log(`date_column: ${date_column}`);
        console.log(`price_column: ${price_column}`);

        sheet_json.forEach(el => {
            // console.log(el);
            const master_key = el[barcode_column];
            el[barcode_column] = el[barcode_column].toString();
            el[date_column] = DateToString(ExcelDateToJSDate(el[date_column]));
            el[price_column] = parseFloat(el[price_column]).toFixed(2);
            addToMaster(master_key, el);
        });

        var new_json = [];
        for (var key in master) {
            new_json.push(master[key]);
        }

        var new_sheet = XLSX.utils.json_to_sheet(new_json);

        var new_workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(new_workbook, new_sheet, "New Sheet");
        var wbout = XLSX.write(new_workbook, { bookType: 'xlsx', type: 'binary' });

        function s2ab(s) {
            var buf = new ArrayBuffer(s.length); //convert s to arrayBuffer
            var view = new Uint8Array(buf);  //create uint8array as viewer
            for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF; //convert to octet
            return buf;
        }
        saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), 'test.xlsx');
    }
});

