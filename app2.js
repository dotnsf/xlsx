// app2.js
var XLSX = require( 'xlsx' );


// ファイル読み込み
var book = XLSX.readFile( './SalesSample.xls' );

// シート
var sheet1 = book.Sheets["Sheet1"];
console.log( sheet1 );

// セル更新
sheet1["C13"] = { v: 1.01, t: 'n', w: '1.01' };

// シート更新
book.Sheets["Sheet1"] = sheet1;

// ファイル書き込み
XLSX.writeFile( book, './SalesSample2.xls' );


