"use strict";

/**
 * Requirements: xlsx.core.min.js from http://sheetjs.com/
 */

function cellRefToIndices(cellRef) {
  var colLetter = cellRef[0];
  var rowIndex = parseInt(cellRef.slice(1), 10);
  var colIndex = colLetter.charCodeAt(0) - 64;

  return { row: rowIndex, col: colIndex };
}

function cellIndicesToRef(indices) {
  return String.fromCharCode(indices.col + 64) + indices.row.toString();
}

function lastCellIndices(map) {
  var maxRow = 0;
  var maxCol = 0;
  
  for (var cellRef in map) {
    if (map.hasOwnProperty(cellRef)) {
      var cellIndices = cellRefToIndices(cellRef);
      var currentRow = cellIndices.row;
      var currentCol = cellIndices.col;
      
      if (currentRow > maxRow) {
        maxRow = currentRow;
      }
      if (currentCol > maxCol) {
        maxCol = currentCol;
      }
    }
  }

  return { row: maxRow, col: maxCol };
}

function xlsxToArray(e, callback) {
  var file = e.target.files[0];
  var reader = new FileReader();

  reader.onload = function(e) {
    var arr = fixdata(e.target.result);
    var wb = XLSX.read(btoa(arr), { type: "base64" });
    var worksheet = wb.Sheets[wb.SheetNames[0]];

    var map = {};

    for (var z in worksheet) {
      if (z[0] === "!") {
        continue;
      }
      map[z] = worksheet[z].v;
    }

    var lastCell = lastCellIndices(map);

    var header = [];
    var array = [];
    
    // populate header
    for (var c = 1, lastCol = lastCell.col; c <= lastCol; c++) {
      var cell = map[cellIndicesToRef({ row: 1, col: c })];
      header.push(cell);
    }

    // populate array of rows
    for (var r = 2, lastRow = lastCell.row; r <= lastRow; r++) {
      var row = {};
      for (var c = 1, lastCol = lastCell.col; c <= lastCol; c++) {
        row[header[c - 1]] = map[cellIndicesToRef({ row: r, col: c })];
      }
      array.push(row);
    }
      
    callback(header, array);
  }

  reader.readAsArrayBuffer(file);
}

function fixdata(data) {
  var o = "", l = 0, w = 10240;
  for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
  o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
  return o;
}
