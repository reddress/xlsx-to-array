"use strict";

var fileInput = document.getElementById("fileInput");
fileInput.addEventListener("change", function(e) { xlsxToArray(e, useArray); });

function useArray(header, arr) {
  console.log(header);
  console.log(arr[0]["Fruits and Vegetables"]);
}
