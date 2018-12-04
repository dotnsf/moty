//. count.js
var XLSX = require( 'xlsx' );
var Utils = XLSX.utils;

var rows = 100; //. データ数

var book = XLSX.readFile( 'xls/manholemap2018.xlsx' );

//console.log( book.Sheets );
for( sheetname in book.Sheets ){
  if( sheetname == 'データセット1' ){
    var sheet = book.Sheets[sheetname];
    //console.log( sheet );

    //var range = sheet["!ref"];
    //console.log( range );  //. [A1:I991]

    //var decodeRange = Utils.decode_range( range );
    //console.log( decodeRange );  //. { s: { c: 0, r: 0 }, e: { c: 8, r: 990 } }

    //. エクセル読み取り
    var ids = {};
    for( var r = 1; r < 1 + rows; r ++ ){
      var a_address = Utils.encode_cell( { r: r, c: 0 } );
      var a_cell = sheet[a_address];
      var b_address = Utils.encode_cell( { r: r, c: 1 } );
      var b_cell = sheet[b_address];

      var page = a_cell.w;
      var pv = parseInt( b_cell.w );
      //console.log( 'page = ' + page + ', pv = ' + pv );

      var tmp1 = page.split( '?' );
      var tmp2 = page.split( '=' );
      var filename = tmp1[0];
      var id = tmp2[1];

      if( id in ids ){
        var obj = ids[id];
        obj.filenames.push( filename );
        obj.pv += pv;
        ids[id] = obj;
      }else{
        var obj = {
           id: id,
           filenames: [ filename ],
           pv: pv,
        };
        ids[id] = obj;
      }
    }
    //console.log( ids );

    //. 配列化
    var pvs = [];
    Object.keys( ids ).forEach( function( key ){
      pvs.push( ids[key] );
    });

    //. ソート
    pvs.sort( compareByPvRev );
    console.log( pvs );
  }
}

function compareByPvRev( a, b ){
  var r = 0;
  if( a.pv < b.pv ){ r = 1; }
  else if( a.pv > b.pv ){ r = -1; }

  return r;
}
