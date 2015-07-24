"use strict";

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
