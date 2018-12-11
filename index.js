$('#input-excel').change(function (e) {
    var reader = new FileReader();

    reader.readAsArrayBuffer(e.target.files[0]);

    reader.onload = function (e) {
        var data = new Uint8Array(reader.result);
        var workbook = XLSX.read(data, { type: 'array' });
        // console.log(workbook);

        var mySheet = workbook.Sheets[workbook.SheetNames[0]];
        var sheet_json = XLSX.utils.sheet_to_json(mySheet);


        var row_count = sheet_json.length;
        var date_format = 'dd/mm/yyyy';
        var date_column = 'G';
        var price_column = 'F';
        var new_sheet = XLSX.utils.json_to_sheet(sheet_json);

        for (let i = 2; i <= row_count + 1; i++) {
            string_date = ExcelDateToJSDate(new_sheet[date_column + i].v);
            new_sheet[date_column + i].t = 's';
            new_sheet[date_column + i].v = DateToString(string_date);
            // console.log(new_sheet[date_column + i]);
        }

        for (let i = 2; i <= row_count + 1; i++) {
            new_sheet[price_column + i].v = new_sheet[price_column + i].v.toString() + '.00';
            new_sheet[price_column + i].t = 's';
        }

        var new_workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(new_workbook, new_sheet, "New Sheet");

        // console.log(new_sheet);

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
    return (mm + '/' + yyyy);
}


 // ----------------------------------------------------------

// if(typeof require !== 'undefined') XLSX = require('xlsx');
// var workbook = XLSX.readFile('file.xlsx');

// var mySheet = workbook.Sheets[workbook.SheetNames[0]];
// var sheet_json = XLSX.utils.sheet_to_json(mySheet);


// var row_count = sheet_json.length;
// var date_format = 'dd/mm/yyyy';
// var date_column = 'G';
// var price_column = 'F';
// var new_sheet = XLSX.utils.json_to_sheet(sheet_json);

// for(let i=2; i <= row_count+1; i++) {
//     string_date = ExcelDateToJSDate(new_sheet[date_column + i].v);
//     new_sheet[date_column + i].t = 's';
//     new_sheet[date_column + i].v = DateToString(string_date);
//     console.log(new_sheet[date_column + i]);
// }

// for(let i=2; i <= row_count+1; i++) {
//     new_sheet[price_column + i].v = new_sheet[price_column + i].v.toString() + '.00';
//     new_sheet[price_column + i].t = 's';
// }

// var new_workbook = XLSX.utils.book_new();
// XLSX.utils.book_append_sheet(new_workbook, new_sheet, "New Sheet");
// // console.log(new_workbook.Sheets[new_workbook.SheetNames[0]].G4);

// XLSX.writeFile(new_workbook, 'out.xlsx');