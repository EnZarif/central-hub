/**
 * Gmail -> Central Hub
 * ---------------------------------------------------------------------------
 * Picks up spreadsheet attachments that a Gmail filter has labelled, and
 * pushes them into the Supabase "imports" bucket. The Hub then offers them
 * on the Allocation Validator and Stock Checks pages.
 *
 * This script is a dumb pipe on purpose. It does not try to work out which
 * page a file belongs to - the Hub does that, because it already parses
 * xlsx and csv with SheetJS and one export often feeds both pages.
 *
 * SETUP (once)
 *   1. script.google.com -> New project -> paste this in.
 *   2. Project Settings -> Script Properties, add three:
 *        SUPABASE_URL  https://qwceyzswwtyqiozxnlah.supabase.co
 *        SERVICE_KEY   <service_role key: Supabase -> Project Settings -> API>
 *        USER_ID       <your uid: Supabase -> Authentication -> Users>
 *   3. In Gmail, make a filter on the report sender that applies the label
 *      "hub-import". Create the label first if it does not exist.
 *   4. Run installTrigger() once, and approve the permission prompt.
 *   5. Run ingest() by hand to test. Check the Execution log.
 *
 * The SERVICE_KEY bypasses row level security. It lives here, in your own
 * Google account, and must never go anywhere near index.html.
 */

var PROP          = PropertiesService.getScriptProperties();
var SEARCH        = 'label:hub-import has:attachment newer_than:7d';
var DONE_LABEL    = 'hub-done';
var MAX_BYTES     = 15 * 1024 * 1024;
var KEEP_SEEN     = 300;
var OK_EXTENSION  = /\.(csv|xlsx|xls)$/i;

/** Run this once to schedule the pickup. */
function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'ingest') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('ingest').timeBased().everyMinutes(5).create();
  Logger.log('Trigger installed: ingest() every 5 minutes.');
}

function ingest() {
  var url = PROP.getProperty('SUPABASE_URL');
  var key = PROP.getProperty('SERVICE_KEY');
  var uid = PROP.getProperty('USER_ID');
  if (!url || !key || !uid) {
    throw new Error('Missing Script Properties. Need SUPABASE_URL, SERVICE_KEY and USER_ID.');
  }

  var seen = loadSeen();
  var label = getOrCreateLabel(DONE_LABEL);
  var threads = GmailApp.search(SEARCH, 0, 50);
  var pushed = 0, skipped = 0;

  threads.forEach(function (thread) {
    // Per MESSAGE, not per thread. A daily report with a constant subject
    // gets threaded by Gmail, so labelling the thread "done" would make
    // every day after the first invisible.
    thread.getMessages().forEach(function (msg) {
      var id = msg.getId();
      if (seen[id]) return;

      var any = false;
      msg.getAttachments({ includeInlineImages: false }).forEach(function (att) {
        var name = att.getName();
        if (!OK_EXTENSION.test(name)) { skipped++; return; }
        if (att.getSize() > MAX_BYTES) {
          Logger.log('Too big, skipped: ' + name + ' (' + att.getSize() + ' bytes)');
          skipped++;
          return;
        }
        try {
          upload(url, key, uid, att, msg.getDate());
          pushed++;
          any = true;
        } catch (e) {
          // Leave it unseen so the next run retries it.
          Logger.log('FAILED ' + name + ': ' + e.message);
        }
      });

      if (any || msg.getAttachments().length === 0) seen[id] = Date.now();
    });
    thread.addLabel(label);
  });

  saveSeen(seen);
  Logger.log('Pushed ' + pushed + ' file(s), skipped ' + skipped + '.');
}

function upload(url, key, uid, att, when) {
  var path = uid + '/inbox/' + stamp(when) + '__' + safeName(att.getName());
  var endpoint = url.replace(/\/+$/, '') + '/storage/v1/object/imports/' + encodeURI(path);

  var res = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: att.getContentType() || 'application/octet-stream',
    payload: att.getBytes(),
    headers: {
      Authorization: 'Bearer ' + key,
      'x-upsert': 'true'
    },
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Supabase said ' + code + ': ' + res.getContentText().slice(0, 300));
  }
  Logger.log('Uploaded ' + path);
}

/** 2026-11-13T06-04-11 - sorts correctly as a string, and is filename safe. */
function stamp(d) {
  var tz = Session.getScriptTimeZone();
  return Utilities.formatDate(d || new Date(), tz, "yyyy-MM-dd'T'HH-mm-ss");
}

function safeName(n) {
  return String(n).replace(/[^A-Za-z0-9.\-_ ]+/g, '_').replace(/\s+/g, ' ').trim();
}

function getOrCreateLabel(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function loadSeen() {
  try { return JSON.parse(PROP.getProperty('SEEN') || '{}'); }
  catch (e) { return {}; }
}

/** Keep the most recent KEEP_SEEN ids so the property never grows unbounded. */
function saveSeen(seen) {
  var ids = Object.keys(seen).sort(function (a, b) { return seen[b] - seen[a]; });
  var trimmed = {};
  ids.slice(0, KEEP_SEEN).forEach(function (id) { trimmed[id] = seen[id]; });
  PROP.setProperty('SEEN', JSON.stringify(trimmed));
}
