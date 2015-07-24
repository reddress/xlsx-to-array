DESCRIPTION

xlsx-to-array parses formatted data from an uploaded XLSX spreadsheet,
creating an array of header cells and another array containing data for
each row.

Each row is represented by an object whose keys correspond to the header names.

The input XLSX spreadsheet must be formatted with header names in the
first row and data in the remaining rows.

Date      Name   Item quantity  Total price   Average
23/7/15   Joe    32             200.13        6.254063
15/2/14   Jim    7              1             0.142856

If a column header is blank, row values are not included.
If a column header is repeated, the rightmost value is used.
Formula results are used in the output.

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
