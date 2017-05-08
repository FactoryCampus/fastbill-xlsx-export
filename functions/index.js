const functions = require('firebase-functions');
const getSheet = require('./generator');
const XLSX = require('xlsx');

exports.getSheet = functions.https.onRequest((request, response) => {
  console.log(request.body);
  getSheet({
    email: request.body.email,
    apikey: request.body.apikey
  })
  .then(wb => {
    const wbbuf = XLSX.write(wb, { type: 'buffer' });
    response.set('Content-Type',  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    response.status(200).send(wbbuf);
  })
  .catch(err => response.status(500).send(err));
});
