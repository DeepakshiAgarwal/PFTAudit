/**
 * PFT Quality Audit — backend API
 *
 * Deploy: Extensions > Apps Script in a Google Sheet, paste this whole file
 * in as Code.gs, then Deploy > New deployment > type "Web app",
 * Execute as: Me, Who has access: Anyone. Copy the resulting /exec URL —
 * that's the API_URL to paste into the dashboard's index.html.
 *
 * Slack notifications: create an Incoming Webhook at https://api.slack.com/apps
 * (New app > From scratch > Incoming Webhooks > Add New Webhook to Workspace,
 * pick the channel). Then in this Apps Script project go to Project Settings
 * (gear icon) > Script Properties > Add script property, name it
 * SLACK_WEBHOOK_URL, and paste the webhook URL as the value. Keeping it there
 * (rather than in this source file) means it never ends up in the GitHub repo.
 */

var HEADERS = ['id','date','auditDate','auditor','agent','ticketId','score','fatal',
  'summary','improvement','fatalFeedback','submittedAt','ratings'];

var PASS_THRESHOLD = 85;
var MID_THRESHOLD = 70;

function getSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Audits');
  if (!sheet) {
    sheet = ss.insertSheet('Audits');
    sheet.appendRow(HEADERS);
  }
  // Force every data column to Plain Text so Sheets never auto-converts
  // date-like strings (e.g. "2026-08-31") into real Date cells, which
  // would shift the value by the sheet's timezone offset.
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2000), HEADERS.length).setNumberFormat('@');
  return sheet;
}

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function asPlainDate(v){
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v;
}

function rowToAudit(headers, row){
  var obj = {};
  headers.forEach(function(h, i){ obj[h] = row[i]; });
  obj.date = asPlainDate(obj.date);
  obj.auditDate = asPlainDate(obj.auditDate);
  try { obj.ratings = JSON.parse(obj.ratings || '[]'); } catch (err) { obj.ratings = []; }
  obj.fatal = (obj.fatal === true || obj.fatal === 'true' || obj.fatal === 'TRUE');
  obj.score = Number(obj.score) || 0;
  return obj;
}

function doGet(e){
  if (e && e.parameter && e.parameter.test === 'slack') {
    return jsonOut(testSlackWebhook());
  }
  var sheet = getSheet();
  var values = sheet.getDataRange().getValues();
  var headers = values.shift();
  var audits = values
    .filter(function(row){ return row[0]; })
    .map(function(row){ return rowToAudit(headers, row); });
  return jsonOut({audits: audits});
}

function testSlackWebhook(){
  var url = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!url) {
    return {configured: false, message: 'SLACK_WEBHOOK_URL script property is not set. Go to Project Settings (gear icon) > Script Properties and add it.'};
  }
  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({text: '✅ Test message from PFT Quality Audit — if you see this in Slack, the webhook works!'}),
      muteHttpExceptions: true
    });
    return {
      configured: true,
      urlPrefix: url.substring(0, 40) + '...',
      responseCode: resp.getResponseCode(),
      responseBody: resp.getContentText()
    };
  } catch (err) {
    return {configured: true, urlPrefix: url.substring(0, 40) + '...', error: String(err)};
  }
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
    var delData = sheet.getDataRange().getValues();
    for (var i = 1; i < delData.length; i++) {
      if (delData[i][0] === body.id) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return jsonOut({ok:true});
  }

  if (body.action === 'notify') {
    var notifyData = sheet.getDataRange().getValues();
    var notifyHeaders = notifyData.shift();
    for (var k = 0; k < notifyData.length; k++) {
      if (notifyData[k][0] === body.id) {
        var audit = rowToAudit(notifyHeaders, notifyData[k]);
        return jsonOut({ok:true, slack: notifySlack(audit)});
      }
    }
    return jsonOut({ok:false, error:'not found'});
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

  var slackResult = body.silent ? null : notifySlack(body);
  return jsonOut({ok:true, slack: slackResult});
}

// ---------- Slack ----------

function statusFor(score, fatal){
  if (fatal) return {label:'Fatal', color:'#B23A32'};
  if (score >= PASS_THRESHOLD) return {label:'Pass', color:'#2E7D46'};
  if (score >= MID_THRESHOLD) return {label:'Needs Improvement', color:'#B8790A'};
  return {label:'Fail', color:'#B23A32'};
}

function buildSlackPayload(audit){
  var st = statusFor(audit.score, audit.fatal);
  var ratings = audit.ratings || [];
  var flagged = ratings.filter(function(r){ return r.rating === 'No' || r.rating === 'Fatal'; });

  var blocks = [];
  blocks.push({
    type: 'header',
    text: {type: 'plain_text', text: '📋 Call Quality Audit — ' + audit.agent, emoji: true}
  });
  blocks.push({
    type: 'section',
    fields: [
      {type: 'mrkdwn', text: '*Score:*\n' + audit.score + '/100'},
      {type: 'mrkdwn', text: '*Status:*\n' + st.label},
      {type: 'mrkdwn', text: '*Auditor:*\n' + audit.auditor},
      {type: 'mrkdwn', text: '*Call date:*\n' + audit.date + (audit.ticketId ? ('  (Ticket: ' + audit.ticketId + ')') : '')}
    ]
  });
  if (flagged.length) {
    var lines = flagged.map(function(r){
      return '• *[' + r.rating + ']* ' + r.question + (r.reason ? (': ' + r.reason) : '');
    }).join('\n');
    blocks.push({type: 'section', text: {type: 'mrkdwn', text: '*Areas flagged:*\n' + lines}});
  }
  if (audit.improvement) {
    blocks.push({type: 'section', text: {type: 'mrkdwn', text: '*Areas of improvement:*\n' + audit.improvement}});
  }
  if (audit.fatal && audit.fatalFeedback) {
    blocks.push({type: 'section', text: {type: 'mrkdwn', text: ':rotating_light: *Fatal feedback:*\n' + audit.fatalFeedback}});
  }

  return {
    text: 'New audit for ' + audit.agent + ' — ' + audit.score + '/100 (' + st.label + ')',
    attachments: [{color: st.color, blocks: blocks}]
  };
}

function notifySlack(audit){
  var webhook = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!webhook) return false;
  try {
    var resp = UrlFetchApp.fetch(webhook, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(buildSlackPayload(audit)),
      muteHttpExceptions: true
    });
    return resp.getResponseCode() === 200;
  } catch (err) {
    return false;
  }
}
