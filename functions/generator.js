const moment = require('moment');
const fastbill = require('fastbill-api');
const XLSX = require('xlsx');

const getAggregate = function(api, filter, former, offset) {
  return api.get(filter, offset, 100)
    .then(result => {
      if(result.length == 100) {
        return getAggregate(api, filter, former.concat(result), offset+100);
      }
      return former.concat(result);
    });
};

const transformDate = function(input) {
  const date = moment(input, 'YYYY-MM-DD');
  if(!date.isValid()) {
    return '';
  }
  return date.toDate();
}

const transformNumber = function(num) {
  return Number(num);
}

function getSheet(creds) {
  fastbill.bootstrap(creds.email, creds.apikey);
  const wb = { SheetNames:[], Sheets:{} };

  return Promise.all([
    getAggregate(fastbill.api.invoice, { TYPE: 'outgoing' }, [], 0)
      .then(result => {
        return result.reduce((sum, invoice) =>
          sum.concat(invoice.ITEMS.map(item => [
            transformDate(invoice.INVOICE_DATE),
            invoice.CUSTOMER_NUMBER,
            invoice.INVOICE_NUMBER,
            item.ARTICLE_NUMBER,
            item.DESCRIPTION,
            transformNumber(item.QUANTITY),
            transformNumber(item.UNIT_PRICE),
            item.COMPLETE_NET,
            invoice.SUB_TOTAL,
            invoice.VAT_TOTAL,
            invoice.TOTAL
          ])),
        [])
      })
      .then(lists => {
        lists.unshift(['Rechnungdatum','Kunden-Nr.','Rechnungsnummer','Artikelnummer','Artikelbeschreibung','Menge','Einzelpreis netto','Gesamtpreis netto','Gesamtsumme netto','Mehrwertsteuer','gesamt Brutto']);
        const ws = XLSX.utils.aoa_to_sheet(lists);
        wb.SheetNames.push("Rechnungen");
        wb.Sheets["Rechnungen"] = ws;
      }),
    getAggregate(fastbill.api.estimate, {}, [], 0)
      .then(result => {
        return result.reduce((sum, estimate) =>
          sum.concat(estimate.ITEMS.map(item => [
            transformDate(estimate.ESTIMATE_DATE),
            estimate.CUSTOMER_NUMBER,
            estimate.ESTIMATE_NUMBER,
            item.ARTICLE_NUMBER,
            item.DESCRIPTION,
            transformNumber(item.QUANTITY),
            transformNumber(item.UNIT_PRICE),
            item.COMPLETE_NET,
            estimate.SUB_TOTAL,
            estimate.VAT_TOTAL,
            estimate.TOTAL,
            estimate.STATE
          ])),
        [])
      })
      .then(lists => {
        lists.unshift(['Angebotsdatum','Kunden-Nr.','Rechnungsnummer','Artikelnummer','Artikelbeschreibung','Menge','Einzelpreis netto','Gesamtpreis netto','Gesamtsumme netto','Mehrwertsteuer','gesamt Brutto','Status']);
        const ws = XLSX.utils.aoa_to_sheet(lists);
        wb.SheetNames.push("Angebote");
        wb.Sheets["Angebote"] = ws;
      })
  ])
  .then(() => {
    return wb;
  });
}

module.exports = getSheet;
