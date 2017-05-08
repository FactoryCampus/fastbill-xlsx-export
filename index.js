const XLSX = require('xlsx');
const getSheet = require('./functions/generator');
require('dotenv').config({});

getSheet({
  email: process.env.FASTBILL_LOGIN,
  apikey: process.env.FASTBILL_APIKEY
})
.then(wb => {
  XLSX.writeFile(wb, 'Auswertung_fastbill.xlsx');
})
.catch(console.error);
