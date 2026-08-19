/**
 * YouTube video bulk edit (bound spreadsheet): pull metadata from the API, edit, push back.
 *
 * Setup (once):
 * 1) Sheet → Extensions → Apps Script. Copy Code.gs, Videos.gs, and appsscript.json into the project.
 * 2) Run any function once → Review permissions → Allow.
 * 3) GCP project for the script → enable YouTube Data API v3.
 *
 * Sheet row 1: Video ID | Title | Description | Tags | Category ID | Privacy | Made for kids | Result
 */

var HEADER_ROW = 1;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('YouTube video bulk')
    .addItem('Prepare video sheet', 'prepareVideosSheet')
    .addItem('Load all videos from channel…', 'loadAllVideosFromChannel')
    .addItem('Fetch rows from YouTube (pull)', 'fetchVideoRowsFromYouTube')
    .addItem('Push updates to YouTube', 'pushVideoUpdatesFromSheet')
    .addToUi();
}

/** Pulls clearer text from YouTube advanced-service errors when present. */
function formatYoutubeApiError_(err) {
  var msg = err && err.message ? String(err.message) : String(err);
  try {
    var idx = msg.indexOf('{');
    if (idx !== -1) {
      var j = JSON.parse(msg.substring(idx));
      if (j.error) {
        var parts = [];
        if (j.error.message) parts.push(j.error.message);
        if (j.error.errors && j.error.errors.length) {
          var e0 = j.error.errors[0];
          if (e0.reason) parts.push('reason: ' + e0.reason);
          if (e0.message && j.error.errors.length > 1) parts.push(e0.message);
        }
        if (parts.length) msg = parts.join(' — ');
      }
    }
  } catch (parseErr) {
    // keep msg
  }
  if (/forbidden|^403\b/i.test(msg)) {
    msg += ' — Use the Google account that owns or manages this video’s channel.';
  }
  return msg;
}

function getDataRowRange_(sheet) {
  var active = sheet.getActiveRange();
  if (active && active.getRow() > HEADER_ROW) {
    var start = active.getRow();
    var end = start + active.getNumRows() - 1;
    return { start: start, end: end };
  }
  return { start: HEADER_ROW + 1, end: sheet.getLastRow() };
}

function cellOrNull_(sheet, row, colZero) {
  var v = sheet.getRange(row, colZero + 1).getValue();
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'string') {
    var t = v.trim();
    return t === '' ? null : t;
  }
  return String(v);
}

function trim_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}
