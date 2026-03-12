'use strict';
var data = '';
process.stdin.on('data', function(d){ data += d; });
process.stdin.on('end', function(){
  var rows = JSON.parse(data);
  var header = rows[0];
  var dataRows = rows.slice(1);
  var last10 = dataRows.slice(Math.max(0, dataRows.length - 10));
  console.log('Last ' + last10.length + ' entries — Activity tab\n');
  last10.forEach(function(row, idx){
    var out = {};
    for (var i = 0; i < header.length; i++) {
      if (row[i] !== '' && row[i] !== undefined && row[i] !== null) {
        out[header[i]] = row[i];
      }
    }
    console.log('[' + (dataRows.length - last10.length + idx + 1) + '] ' + JSON.stringify(out, null, 2));
    console.log('---');
  });
});
