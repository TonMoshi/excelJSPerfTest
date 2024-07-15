/**
 * Copyright (c) 2014-2019 Guyon Roche
 * LICENCE: MIT - please refer to LICENSE file included with this module
 * or https://github.com/exceljs/exceljs/blob/master/LICENSE
 */

if (parseInt(process.versions.node.split('.')[0], 10) < 10) {
  throw new Error(
    'For node versions older than 10, please use the ES5 Import: https://github.com/exceljs/exceljs#es5-imports'
  );
}

const v8 =require('v8');

const totalHeapSize = v8.getHeapStatistics().total_available_size;
const totalHeapSizeinGB = (totalHeapSize /1024/1024/1024).toFixed(2);
console.log(`Total heap size: ${totalHeapSizeinGB} GB`);

const {Writable} = require('node:stream');
const fs = require('fs');

const textDecoder = typeof TextDecoder === 'undefined' ? null : new TextDecoder('utf-8');
const {StringDecoder} = require('string_decoder');

const decoder =typeof StringDecoder ===  'undefined'? null: new StringDecoder('utf8');
function bufferToString(chunk) {
  if (typeof chunk === 'string') {
    return chunk;
  }
  if (decoder) {
    return decoder.write(chunk);
  }
  if (textDecoder) {
    return textDecoder.decode(chunk);
  }
  return chunk.toString();
}

const dataPrev = require('./generated.json');

let data = [...dataPrev,
];

console.log(`Data lenght: ${  data.length}`);

const start = new Date().getTime();
console.log(`Start: ${start}`);

const exceljs = require('./lib/exceljs.nodejs.js');

 const file = [];
 const archivo = fs.createWriteStream('./big.xlsx');


const writable = new Writable({encoding: 'utf8',write(chunk, encoding, callback){
    
    archivo.write(chunk);
    file.push(chunk);
    callback();
    },
  });


// const workbook = new exceljs.stream.xlsx.WorkbookWriter({filename: 'testFilename.xlsx'});
const workbook = new exceljs.stream.xlsx.WorkbookWriter({stream: writable});

const worksheet = workbook.addWorksheet('shit');
worksheet.columns = Object.keys(data[0]).map(k => ({
  header: k, key: k, width: 50,
}));

const usedHeapSize = v8.getHeapStatistics().used_heap_size;
const usedHeapSizeGB = (usedHeapSize /1024/1024/1024).toFixed(2);
console.log(`Used heap size PRE Add Rows: ${usedHeapSizeGB} GB`);

data.forEach(row => worksheet.addRow(row).commit());
data = null;
// worksheet.addRows(data);

console.log('-----------------');

const usedHeapSizeRows = v8.getHeapStatistics().used_heap_size;
const usedHeapSizeRowsGB = (usedHeapSizeRows /1024/1024/1024).toFixed(2);
console.log(`Used heap size POST Add Rows: ${usedHeapSizeRowsGB} GB`);

const addRws = new Date().getTime();
// console.log(`AddRows: ${addRws / 1000}`);
console.log(`AddRows Diff: ${(addRws - start) / 1000}`);
console.log('-----------------');

worksheet.commit();
workbook.commit()
.then(x => console.log('OK'))
.catch(x => console.log(`KO: ${x}`))
.finally(() => {

  console.log(`length: ${file.length}`);
  fs.writeFileSync('testMax.xlsx', Buffer.concat(file), function(err) {
      if(err) {
          return console.log(err);
      }
      console.log('The file was saved!');
  });

  archivo.end();

  console.log('FINALLY');
  const end = new Date().getTime();

  // console.log(`End: ${end / 1000}`);
  console.log(`End Diff: ${(end - start) / 1000}`);

  const endGb = (v8.getHeapStatistics().used_heap_size /1024/1024/1024).toFixed(2);
  console.log(`Used heap size On END: ${endGb} GB`);
});



// workbook.xlsx.writeFile('test.xlsx')
// workbook.xlsx.writeBuffer()
// .then(() => console.log('OK'))
// .catch(x => console.log(`KO: ${  x}`))
// .finally(() => {
//   const end = new Date().getTime();

//   console.log(`End: ${end / 1000}`);
//   console.log(`Diff: ${(end - start) / 1000}`);

//   const endGb = (v8.getHeapStatistics().used_heap_size /1024/1024/1024).toFixed(2);
//   console.log(`Used heap size On END: ${endGb} GB`);
// });

module.exports = exceljs;

