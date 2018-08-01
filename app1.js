// app1.js
var XLSX = require( 'xlsx' );
var filename = 'SalesSample.xls';
if( process.argv.length > 2 ){
  filename = process.argv[2];
}


// ファイル読み込み
var book = XLSX.readFile( './' + filename );

// シート
var sheet1 = book.Sheets["Sheet1"];
console.log( sheet1 );

//var range1 = sheet1["!ref"];
//console.log( range1 );



