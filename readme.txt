DESCRIPTION

xlsx-to-array parses an uploaded XLSX spreadsheet, creating an array of header
names and another array of the data rows.

Each row is represented by an object whose keys correspond to the header names.

The input XLSX spreadsheet must be formatted with the header names in the
first row and data in the following rows.

Date      Name   Item quantity  Total price   Average
23/7/15   Joe    32             200.13        6.254063
15/2/14   Jim    7              1             0.142856

If a column header is blank, the entire column is ignored.
If a column header is repeated, the rightmost value is used.
Empty rows or cells will appear as undefined.
Formula results are outputted.

USAGE

Include a file input element in your HTML. Then, add the xlsx.core.min.js
and xlsx-to-array.js scripts.

See index.html and ui.js as an example.

var fileInput = document.getElementById("fileInput");
fileInput.addEventListener("change", function(e) {
  xlsxToArray.parse(e, processData);
});

function processData(header, rows) {
  console.log(header);
  for (var i = 0, len = rows.length; i < len; i++) {
    console.log(rows[i]["Name"]);
  }
}
