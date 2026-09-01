/**
 * PFT Quality Audit — backend API
 *
 * Deploy: Extensions > Apps Script in a Google Sheet, paste this whole file
 * in as Code.gs, then Deploy > New deployment > type "Web app",
 * Execute as: Me, Who has access: Anyone. Copy the resulting /exec URL —
 * that's the API_URL to paste into the dashboard's index.html.
 */

var HEADERS = ['id','date','auditDate','auditor','agent','ticketId','score','fatal',
  'summary','improvement','fatalFeedback','submittedAt','ratings'];

function getSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Audits');
  if (!sheet) {
    sheet = ss.insertSheet('Audits');
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e){
  var sheet = getSheet();
  var values = sheet.getDataRange().getValues();
  var headers = values.shift();
  var audits = values
    .filter(function(row){ return row[0]; })
    .map(function(row){
      var obj = {};
      headers.forEach(function(h, i){ obj[h] = row[i]; });
      try { obj.ratings = JSON.parse(obj.ratings || '[]'); } catch (err) { obj.ratings = []; }
      obj.fatal = (obj.fatal === true || obj.fatal === 'true' || obj.fatal === 'TRUE');
      obj.score = Number(obj.score) || 0;
      return obj;
    });
  return jsonOut({audits: audits});
}

function doPost(e){
  var sheet = getSheet();
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ok:false, error:'bad json'});
  }

  if (body.action === 'delete') {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === body.id) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return jsonOut({ok:true});
  }

  // Reject duplicate ids (idempotent import / accidental double-submit)
  var existing = sheet.getDataRange().getValues();
  for (var j = 1; j < existing.length; j++) {
    if (existing[j][0] === body.id) {
      return jsonOut({ok:true, skipped:true});
    }
  }

  var row = HEADERS.map(function(h){
    if (h === 'ratings') return JSON.stringify(body.ratings || []);
    return body[h] !== undefined ? body[h] : '';
  });
  sheet.appendRow(row);
  return jsonOut({ok:true});
}
