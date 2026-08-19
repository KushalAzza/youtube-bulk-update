/**
 * Video rows: pull snippet + status from YouTube, edit in-sheet, push updates.
 * Sheet row 1 must match VIDEO_HEADERS (menu: Prepare video sheet).
 */

var VIDEO_HEADERS = [
  'Video ID',
  'Title',
  'Description',
  'Tags',
  'Category ID',
  'Privacy',
  'Made for kids',
  'Result'
];

var VC = {
  VIDEO_ID: 0,
  TITLE: 1,
  DESCRIPTION: 2,
  TAGS: 3,
  CATEGORY_ID: 4,
  PRIVACY: 5,
  MADE_FOR_KIDS: 6,
  RESULT: 7
};

function prepareVideosSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(HEADER_ROW, 1, 1, VIDEO_HEADERS.length).setValues([VIDEO_HEADERS]);
  sheet.getRange(HEADER_ROW, 1, 1, VIDEO_HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, VIDEO_HEADERS.length);
  SpreadsheetApp.getUi().alert(
    'Video sheet ready. Use “Load all videos from channel…” or paste Video IDs in column A, then Fetch / Push from the menu.'
  );
}

/** Prompts for UC… ID, clears data rows, fills one row per upload on the channel. */
function loadAllVideosFromChannel() {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt(
    'Load videos from channel',
    'Enter the channel ID (UC…). All videos from that channel’s uploads playlist will be loaded (may take a while for large channels).',
    ui.ButtonSet.OK_CANCEL
  );
  if (prompt.getSelectedButton() !== ui.Button.OK) return;

  var channelId = trim_(prompt.getResponseText());
  if (!channelId) {
    ui.alert('Channel ID is required.');
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureVideoHeaders_(sheet);

  var uploadsId = getUploadsPlaylistId_(channelId);
  var videoIds = listAllVideoIdsInUploadsPlaylist_(uploadsId);
  if (videoIds.length === 0) {
    ui.alert('No videos found in the uploads playlist.');
    return;
  }

  var oldLast = sheet.getLastRow();
  var clearThrough = Math.max(oldLast, HEADER_ROW + videoIds.length);
  if (clearThrough > HEADER_ROW) {
    var clearNumRows = clearThrough - HEADER_ROW;
    sheet.getRange(HEADER_ROW + 1, 1, clearNumRows, VIDEO_HEADERS.length).clearContent();
  }

  var rows = buildVideoRowsFromIds_(videoIds);
  if (rows.length > 0) {
    sheet.getRange(HEADER_ROW + 1, 1, rows.length, VIDEO_HEADERS.length).setValues(rows);
  }

  ui.alert('Loaded ' + rows.length + ' video(s). Edit rows, then use “Videos: Push updates”.');
}

/** Re-fetch Title, Description, Tags, Category, Privacy, Made for kids from YouTube for each row (by Video ID). */
function fetchVideoRowsFromYouTube() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureVideoHeaders_(sheet);

  var range = getDataRowRange_(sheet);
  if (range.end < range.start) {
    SpreadsheetApp.getUi().alert('No data rows to process.');
    return;
  }

  for (var r = range.start; r <= range.end; r++) {
    var videoId = trim_(sheet.getRange(r, VC.VIDEO_ID + 1).getValue());
    if (!videoId) {
      sheet.getRange(r, VC.RESULT + 1).setValue('Skipped (no Video ID)');
      continue;
    }

    try {
      writeVideoRowFromApi_(sheet, r, videoId);
    } catch (err) {
      sheet.getRange(r, VC.RESULT + 1).setValue('Error: ' + formatYoutubeApiError_(err));
    }
  }

  SpreadsheetApp.getUi().alert('Video fetch finished. Edit cells, then push updates.');
}

/** Push non-empty cells to YouTube per row (blank cell = leave that field unchanged on YouTube). */
function pushVideoUpdatesFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureVideoHeaders_(sheet);

  var range = getDataRowRange_(sheet);
  if (range.end < range.start) {
    SpreadsheetApp.getUi().alert('No data rows to process.');
    return;
  }

  for (var r = range.start; r <= range.end; r++) {
    var videoId = trim_(sheet.getRange(r, VC.VIDEO_ID + 1).getValue());
    if (!videoId) {
      sheet.getRange(r, VC.RESULT + 1).setValue('Skipped (no Video ID)');
      continue;
    }

    try {
      applySheetRowToVideo_(sheet, r, videoId);
      sheet.getRange(r, VC.RESULT + 1).setValue('OK ' + new Date().toISOString());
    } catch (err) {
      sheet.getRange(r, VC.RESULT + 1).setValue('Error: ' + formatYoutubeApiError_(err));
    }
  }

  SpreadsheetApp.getUi().alert('Video updates finished. Check the Result column.');
}

function ensureVideoHeaders_(sheet) {
  var row = sheet.getRange(HEADER_ROW, 1, 1, VIDEO_HEADERS.length).getValues()[0];
  for (var i = 0; i < VIDEO_HEADERS.length; i++) {
    if (trim_(row[i]) !== VIDEO_HEADERS[i]) {
      throw new Error(
        'Row 1 must be the video headers. Run “Prepare video sheet” on this tab (expected: ' +
          VIDEO_HEADERS.join(' | ') +
          ').'
      );
    }
  }
}

function getUploadsPlaylistId_(channelId) {
  var res = YouTube.Channels.list('contentDetails', { id: channelId });
  if (!res.items || res.items.length === 0) {
    throw new Error('Channel not found: ' + channelId);
  }
  var uploads = res.items[0].contentDetails && res.items[0].contentDetails.relatedPlaylists
    ? res.items[0].contentDetails.relatedPlaylists.uploads
    : null;
  if (!uploads) {
    throw new Error('No uploads playlist for channel: ' + channelId);
  }
  return uploads;
}

function listAllVideoIdsInUploadsPlaylist_(playlistId) {
  var ids = [];
  var pageToken;
  do {
    var resp = YouTube.PlaylistItems.list('contentDetails', {
      playlistId: playlistId,
      maxResults: 50,
      pageToken: pageToken
    });
    if (resp.items) {
      for (var i = 0; i < resp.items.length; i++) {
        var vid = resp.items[i].contentDetails && resp.items[i].contentDetails.videoId;
        if (vid) ids.push(vid);
      }
    }
    pageToken = resp.nextPageToken;
  } while (pageToken);
  return ids;
}

function buildVideoRowsFromIds_(videoIds) {
  var byId = {};
  var batchSize = 50;
  for (var b = 0; b < videoIds.length; b += batchSize) {
    var slice = videoIds.slice(b, b + batchSize);
    var resp = YouTube.Videos.list('snippet,status', {
      id: slice.join(','),
      maxResults: 50
    });
    if (resp.items) {
      for (var j = 0; j < resp.items.length; j++) {
        var v = resp.items[j];
        byId[v.id] = v;
      }
    }
  }

  var rows = [];
  for (var k = 0; k < videoIds.length; k++) {
    var vid = videoIds[k];
    var v = byId[vid];
    if (!v) {
      rows.push([vid, '', '', '', '', '', '', 'Error: not returned by API']);
      continue;
    }
    rows.push(videoToSheetRow_(v));
  }
  return rows;
}

function videoToSheetRow_(v) {
  var sn = v.snippet || {};
  var st = v.status || {};
  var tags = sn.tags && sn.tags.length ? sn.tags.join(', ') : '';
  var cat = sn.categoryId != null ? String(sn.categoryId) : '';
  var privacy = st.privacyStatus != null ? String(st.privacyStatus) : '';
  var kids = '';
  if (st.selfDeclaredMadeForKids === true) kids = 'TRUE';
  else if (st.selfDeclaredMadeForKids === false) kids = 'FALSE';

  return [
    v.id,
    sn.title != null ? String(sn.title) : '',
    sn.description != null ? String(sn.description) : '',
    tags,
    cat,
    privacy,
    kids,
    'Fetched ' + new Date().toISOString()
  ];
}

function writeVideoRowFromApi_(sheet, row, videoId) {
  var resp = YouTube.Videos.list('snippet,status', { id: videoId, maxResults: 1 });
  if (!resp.items || resp.items.length === 0) {
    throw new Error('Video not found: ' + videoId);
  }
  var cells = videoToSheetRow_(resp.items[0]);
  sheet.getRange(row, VC.VIDEO_ID + 1, 1, VC.RESULT - VC.VIDEO_ID).setValues([cells.slice(0, VC.RESULT)]);
  sheet.getRange(row, VC.RESULT + 1).setValue(cells[VC.RESULT]);
}

function applySheetRowToVideo_(sheet, row, videoId) {
  var title = cellOrNull_(sheet, row, VC.TITLE);
  var description = cellOrNull_(sheet, row, VC.DESCRIPTION);
  var tagsCell = cellOrNull_(sheet, row, VC.TAGS);
  var categoryId = cellOrNull_(sheet, row, VC.CATEGORY_ID);
  var privacy = cellOrNull_(sheet, row, VC.PRIVACY);
  var kidsCell = cellOrNull_(sheet, row, VC.MADE_FOR_KIDS);

  var tagsParsed = tagsCell === null ? null : parseTagsFromCell_(tagsCell);
  var kidsParsed = kidsCell === null ? null : parseBoolLoose_(kidsCell);

  if (
    title === null &&
    description === null &&
    tagsParsed === null &&
    categoryId === null &&
    privacy === null &&
    kidsParsed === null
  ) {
    throw new Error('Nothing to update (all editable cells are empty).');
  }

  var list = YouTube.Videos.list('snippet,status', { id: videoId, maxResults: 1 });
  if (!list.items || list.items.length === 0) {
    throw new Error('Video not found or not editable by this account: ' + videoId);
  }

  var video = list.items[0];
  var parts = [];

  if (title !== null || description !== null || tagsParsed !== null || categoryId !== null) {
    if (!video.snippet) video.snippet = {};
    if (title !== null) video.snippet.title = title;
    if (description !== null) video.snippet.description = description;
    if (categoryId !== null) video.snippet.categoryId = categoryId;
    if (tagsParsed !== null) video.snippet.tags = tagsParsed;
    parts.push('snippet');
  }

  if (privacy !== null || kidsParsed !== null) {
    if (!video.status) video.status = {};
    if (privacy !== null) video.status.privacyStatus = privacy.toLowerCase();
    if (kidsParsed !== null) video.status.selfDeclaredMadeForKids = kidsParsed;
    parts.push('status');
  }

  var partStr = parts.join(',');
  var payload = { id: video.id };
  if (parts.indexOf('snippet') >= 0) {
    payload.snippet = pickWritableSnippetForUpdate_(video.snippet);
  }
  if (parts.indexOf('status') >= 0) {
    payload.status = pickWritableStatusForUpdate_(video.status);
  }

  YouTube.Videos.update(payload, partStr);
}

/**
 * videos.list returns read-only snippet fields; sending them back on update can fail.
 */
function pickWritableSnippetForUpdate_(sn) {
  if (!sn) return {};
  var o = {};
  if (sn.title != null) o.title = sn.title;
  if (sn.description != null) o.description = sn.description;
  if (sn.tags != null) o.tags = sn.tags;
  if (sn.categoryId != null) o.categoryId = String(sn.categoryId);
  if (sn.defaultLanguage != null) o.defaultLanguage = sn.defaultLanguage;
  if (sn.defaultAudioLanguage != null) o.defaultAudioLanguage = sn.defaultAudioLanguage;
  return o;
}

/**
 * videos.list returns read-only status fields (e.g. uploadStatus); never send those on update.
 */
function pickWritableStatusForUpdate_(st) {
  if (!st) return {};
  var o = {};
  if (st.privacyStatus != null) o.privacyStatus = String(st.privacyStatus);
  if (st.publishAt != null) o.publishAt = st.publishAt;
  if (st.license != null) o.license = st.license;
  if (st.embeddable !== undefined && st.embeddable !== null) o.embeddable = st.embeddable;
  if (st.publicStatsViewable !== undefined && st.publicStatsViewable !== null) {
    o.publicStatsViewable = st.publicStatsViewable;
  }
  if (st.selfDeclaredMadeForKids !== undefined && st.selfDeclaredMadeForKids !== null) {
    o.selfDeclaredMadeForKids = st.selfDeclaredMadeForKids;
  }
  return o;
}

function parseTagsFromCell_(s) {
  var raw = String(s).split(',');
  var out = [];
  for (var i = 0; i < raw.length; i++) {
    var t = trim_(raw[i]);
    if (t) out.push(t);
  }
  return out;
}

function parseBoolLoose_(s) {
  var u = trim_(String(s)).toLowerCase();
  if (u === 'true' || u === 'yes' || u === '1' || u === 'y') return true;
  if (u === 'false' || u === 'no' || u === '0' || u === 'n') return false;
  throw new Error('Made for kids must be TRUE or FALSE (got: ' + s + ')');
}
