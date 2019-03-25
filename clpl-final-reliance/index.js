var master = {};
var xlFileName;

// Cruciol
var barcode_column = 'z';
var date_column = 'z';
var price_column = 'z';
var qty_column = 'z';

// essential
var metsize_column = 'z';
var desc_column = 'z';
var articleno_column = 'z';
var labeltype_column = 'z';

// No Changes
var style_column = 'z';
var color_column = 'z';
var size_column = 'z';
var vendor_column = 'z';
var fashiongradedesc_column = 'z';
var barcolor_column = 'z';

// Hard Coded
var extraqty_column = 'Extra Qty';
var totalqty_column = 'Total Qty';


var collect_column = 'Collection';



// FUNCTIONS
addToMaster = function (master_key, el) {
    if (master_key in master) {
        master[master_key][qty_column] += el[qty_column];
        master[master_key][collect_column] = "TRUE";
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

    var brand = document.getElementById('brand').value;
    if (brand === 'ajio')
        return (mm + ' / ' + yyyy);
    else if (brand === 'trends')
        return (mm + '/' + yyyy);
}

findWord = function (word, str) {
    return str.split(' ').some(function (w) { return w.toUpperCase() === word.toUpperCase() })
}

handleMetsize = function(str) {
    str = str.trim();
    while(isNaN(parseInt(str[0])))
        str = str.substr(1);
    let len = str.length;
    for(let i=len; i>=0; i--) {
        if(!isNaN(parseInt(str[i]))) {
            if(str[i+1] === ' ')
                return str;
            else
                return (str.substr(0, i+1) + ' ' + str.substr(i+1));
        }
    }
}

markColumns = function (el) {
    Object.keys(el).forEach(function (key, index) {
        if (findWord('MRP', key))
            price_column = key;
        if (findWord('EAN', key))
            barcode_column = key;
        if (findWord('MFD', key) || findWord('Yrmonth', key))
            date_column = key;
        if (findWord('LABEL', key) && findWord('QTY', key))
            qty_column = key;

        if (findWord('METSIZE', key))
            metsize_column = key;
        if (findWord('DESC', key) && !findWord('FASHION', key))
            desc_column = key;
        if (findWord('ARTICLE', key))
            articleno_column = key;
        if (findWord('LABEL', key) && findWord('TYPE', key))
            labeltype_column = key;

        if (findWord('STYLE', key))
            style_column = key;
        if (findWord('COLOR', key))
            color_column = key;
        if (findWord('SIZE', key))
            size_column = key;
        if (findWord('VENDOR', key))
            vendor_column = key;
        if (findWord('FASHION', key) && findWord('DESC', key))
            fashiongradedesc_column = key;
        if (findWord('BAR', key) && findWord('COLOUR', key))
            barcolor_column = key;
    });


    if (barcode_column === 'z')
        alert(`Barcode Column not found. Should have the word "EAN".`);
    if (date_column === 'z')
        alert(`Date Column not found. Should have the word "MFD" or "Yrmonth".`);
    if (price_column === 'z')
        alert(`MRP column not found. Should have the word "MRP".`)
    if (qty_column === 'z')
        alert(`Quantity column not found. Should have the word "Label" and "QTY".`)

    if (metsize_column === 'z')
        alert(`Metsize column not found. Should have the word "METSIZE".`)
    if (desc_column === 'z')
        alert(`Desc Column not found. Should have the work "DESC".`);
    if (articleno_column === 'z')
        alert(`Article No Column not found. Should have the word "ARTICLE".`);
    if (labeltype_column === 'z')
        alert(`Label Type Column not found. Should have the word "LABEL" and "TYPE".`);

    if (style_column === 'z')
        alert(`Style Code Column not found. Should have the word "STYLE".`);
    if (color_column === 'z')
        alert(`Color Column not found. Should have the word "COLOR".`);
    if (size_column === 'z')
        alert(`Size column not found. Should have the word "SIZE".`)
    if (vendor_column === 'z')
        alert(`Vendor Column not found. Should have the word "VENDOR".`);
    if (fashiongradedesc_column === 'z')
        alert(`Fashion Grade Description Column not found. Should have the word "FASHION" and "DESC".`);
    if (barcolor_column === 'z')
        alert(`Barcode Color column not found. Should have the word "BAR" and "COLOUR".`)


    if (barcode_column === 'z' || date_column === 'z' || price_column === 'z' || qty_column === 'z')
        return;
    if (metsize_column === 'z' || desc_column === 'z' || articleno_column === 'z' || labeltype_column === 'z')
        return;
    if (style_column === 'z' || color_column === 'z' || size_column === 'z' || vendor_column === 'z' || fashiongradedesc_column === 'z' || barcolor_column === 'z')
        return;

    console.log(`barcode_column: ${barcode_column}`);
    console.log(`date_column: ${date_column}`);
    console.log(`price_column: ${price_column}`);
    console.log(`qty_column: ${qty_column}`);
    console.log(`---------------------`);
    console.log(`metsize_column: ${metsize_column}`);
    console.log(`desc_column: ${desc_column}`);
    console.log(`articleno_column: ${articleno_column}`);
    console.log(`labeltype_column: ${labeltype_column}`);
    console.log(`---------------------`);
    console.log(`style_column: ${style_column}`);
    console.log(`color_column: ${color_column}`);
    console.log(`size_column: ${size_column}`);
    console.log(`vendor_column: ${vendor_column}`);
    console.log(`fashiongradedesc_column: ${fashiongradedesc_column}`);
    console.log(`barcolor_column: ${barcolor_column}`);
}

createAndDownload = function (new_json) {
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
    saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), xlFileName);
}



// WORKING
$('#input-excel').change(function (e) {
    var reader = new FileReader();
    reader.readAsArrayBuffer(e.target.files[0]);
    xlFileName = "NEW " + e.target.files[0].name;
    reader.onload = function (e) {
        var data = new Uint8Array(reader.result);
        var workbook = XLSX.read(data, { type: 'array' });
        var sheet = workbook.Sheets[workbook.SheetNames[0]];
        // console.log(sheet);
        var sheet_json = XLSX.utils.sheet_to_json(sheet);
        console.log(sheet_json);

        var el = sheet_json[0];
        markColumns(el);

        sheet_json.forEach(el => {
            // console.log(el);
            var row = {};
            row[style_column] = el[style_column];
            row[color_column] = el[color_column];
            row[size_column] = el[size_column];
            row[metsize_column] = (el[metsize_column]? handleMetsize(el[metsize_column]) : "");
            row[desc_column] = el[desc_column].toUpperCase();
            row[date_column] = DateToString(ExcelDateToJSDate(el[date_column]));
            row[price_column] = parseFloat(el[price_column]).toFixed(2);
            row[barcode_column] = el[barcode_column].toString();
            row[articleno_column] = el[articleno_column].toString();
            row[qty_column] = parseInt(el[qty_column]);
            row[extraqty_column] = 0;
            row[totalqty_column] = row[qty_column];
            row[vendor_column] = el[vendor_column];
            row[labeltype_column] = el[labeltype_column].toString().trim() + 'N';
            row[fashiongradedesc_column] = el[fashiongradedesc_column];
            row[barcolor_column] = el[barcolor_column];
            row[collect_column] = "";

            // el[barcode_column] = el[barcode_column].toString();
            // el[date_column] = DateToString(ExcelDateToJSDate(el[date_column]));
            // el[price_column] = parseFloat(el[price_column]).toFixed(2);
            // el[collect_column] = "";
            const master_key = row[barcode_column] + row[date_column];
            addToMaster(master_key, row);
        });

        var new_json = [];
        for (var key in master) {

            // Extra Qty and Total Quantity
            var qtybuffer = parseInt(document.getElementById('qtybuffer').value);
            var curval = master[key][qty_column];
            var extraqty = Math.round((qtybuffer / 100) * curval);
            master[key][extraqty_column] = extraqty;
            master[key][totalqty_column] = curval + extraqty;

            new_json.push(master[key]);
        }

        createAndDownload(new_json);
    }
});

