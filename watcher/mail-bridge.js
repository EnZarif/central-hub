// Local mail bridge: lets the Hub hand finished emails to desktop Outlook.
//
// The Hub builds each operator email completely - address, CC, subject and the
// HTML table - but a browser cannot send mail. The old answer was a mailto:
// per operator, so fifteen emails meant fifteen windows and fifteen clicks.
//
// This is the smallest thing that closes that gap. It listens on loopback
// only, takes a batch of finished messages, and hands them to a PowerShell
// script that drives the Outlook profile already signed in on this machine.
// No app registration, no admin consent, no tenant permission - it is your
// own desktop mail client, automated as you.
//
// ---- why it is safe to run -------------------------------------------------
// A local HTTP server that sends email as you is a real capability, and any
// page in your browser can make requests to localhost. Three things stop
// that being a problem:
//
//   1. It binds 127.0.0.1, never 0.0.0.0 - nothing off this machine can
//      reach it, including anything else on your network.
//   2. Origin is checked against an allowlist. A request from any other
//      site is refused before the body is even read.
//   3. It requires a custom header. A browser will not send one
//      cross-origin without first passing a CORS preflight, and the
//      preflight only ever succeeds for an allowed origin. This is what
//      stops a hostile page simply POSTing a form at it.
//
// It also never takes an address from the request alone without logging it,
// so a run that went somewhere unexpected is visible afterwards.

const http = require('http');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFile } = require('child_process');

const PORT = 8765;
const HOST = '127.0.0.1';
const PS_SCRIPT = path.join(__dirname, '..', 'automation', 'send-outlook.ps1');

// Where the Hub actually runs. The Vercel address is the real one - the
// Hub opens at login as an Edge app window pointing there - and the
// localhost pair only matter when it is served from start.bat instead.
//
// Because the Vercel page is https and this is loopback, the browser will
// ask permission the first time ("allow local network access"). That is a
// one-off prompt, not a block, and there is nothing to send for it: the
// old Access-Control-Allow-Private-Network header is deprecated.
const ALLOWED_ORIGINS = new Set([
  'https://central-hub-rosy.vercel.app',
  'http://localhost:8000',
  'http://127.0.0.1:8000',
  ...String(process.env.HUB_MAIL_ORIGINS || '')
    .split(',').map(s => s.trim()).filter(Boolean)
]);

const GUARD_HEADER = 'x-hub-mail';
const MAX_BODY = 8 * 1024 * 1024;     // 15 rich emails is nowhere near this
// A working answer is stable, so it can be held for a while. A failed one
// is not: the common reason COM refuses is that this ran during login,
// while Outlook was still coming up. Caching that for five minutes would
// mean the Hub showed the fallback buttons through the whole morning for a
// problem that cured itself in seconds.
const HEALTH_TTL_MS = 5 * 60 * 1000;
const HEALTH_FAIL_TTL_MS = 20 * 1000;

let healthCache = { at: 0, value: null };

function log(...a) { console.log('[mail-bridge]', ...a); }

/** Run the sender. `whatIf` builds everything and sends nothing. */
function runPowerShell(messages, whatIf) {
  return new Promise((resolve) => {
    let tmp, resultPath;
    try {
      const stem = 'hub-mail-' + Date.now() + '-' + Math.random().toString(36).slice(2);
      tmp = path.join(os.tmpdir(), stem + '.json');
      resultPath = path.join(os.tmpdir(), stem + '.result.json');
      fs.writeFileSync(tmp, JSON.stringify({ messages }), 'utf8');
    } catch (e) {
      return resolve({ ok: false, sent: 0, failed: [], error: 'Could not stage payload: ' + e.message });
    }

    const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                  '-File', PS_SCRIPT, '-JsonPath', tmp, '-ResultPath', resultPath];
    if (whatIf) args.push('-WhatIf');

    const cleanup = () => {
      // The payload holds real addresses, so it does not linger in temp
      // any longer than the run that needed it.
      try { fs.unlinkSync(tmp); } catch { /* nothing to do */ }
      try { fs.unlinkSync(resultPath); } catch { /* nothing to do */ }
    };

    execFile('powershell', args, { timeout: 300000, maxBuffer: 4 * 1024 * 1024 },
      (err, stdout, stderr) => {
        // stdout first, then the file the script wrote the moment the mail
        // left. The file is what stops a killed process being mistaken for
        // mail that never went - which is what caused a duplicate before.
        let out = null;
        try { out = JSON.parse(String(stdout).trim()); } catch { /* fall through */ }
        if (!out) {
          try { out = JSON.parse(fs.readFileSync(resultPath, 'utf8')); }
          catch { /* genuinely nothing to report */ }
          if (out) log('recovered the result from disk after the sender was cut short');
        }
        cleanup();

        if (out) return resolve(out);
        if (err) {
          const timedOut = err.killed || err.signal;
          return resolve({ ok: false, sent: 0, failed: [],
            // Never say "not sent" when that is not known.
            error: timedOut
              ? 'The sender was cut short and left no record. Check Sent Items before trying again - it may already have gone.'
              : 'PowerShell failed: ' + (stderr || err.message).slice(0, 400) });
        }
        resolve({ ok: false, sent: 0, failed: [],
                  error: 'Unreadable reply from the sender: ' + String(stdout).slice(0, 300) });
      });
  });
}

// ---- duplicate guard -------------------------------------------------------
// Belt and braces for the same failure. Even with the result now surviving
// a kill, a double-click or an impatient retry should not be able to put
// the same mail out twice - so the same recipient and subject inside this
// window is refused and said so, rather than sent again.
const RECENT_TTL_MS = 15 * 60 * 1000;
const recentSends = new Map();

function sendKey(m) {
  return String(m.to || '').toLowerCase() + '|' + String(m.subject || '');
}
function partitionDuplicates(messages) {
  const now = Date.now();
  for (const [k, t] of recentSends) if (now - t > RECENT_TTL_MS) recentSends.delete(k);
  const fresh = [], duplicates = [];
  for (const m of messages) {
    if (recentSends.has(sendKey(m))) duplicates.push(m); else fresh.push(m);
  }
  return { fresh, duplicates };
}
function rememberSends(messages, failed) {
  const bad = new Set((failed || []).map(f => String(f.to || '').toLowerCase()));
  messages.forEach(m => {
    if (!bad.has(String(m.to || '').toLowerCase())) recentSends.set(sendKey(m), Date.now());
  });
}

async function health(force) {
  const now = Date.now();
  const ttl = (healthCache.value && healthCache.value.ready) ? HEALTH_TTL_MS : HEALTH_FAIL_TTL_MS;
  if (!force && healthCache.value && now - healthCache.at < ttl) return healthCache.value;
  // One message, never sent - just enough for the script to reach its
  // account check and report what it found.
  const r = await runPowerShell(
    [{ to: 'preflight@invalid', cc: '', subject: 'preflight', html: '<p></p>' }], true);
  const value = { ready: !!r.ok, account: r.account || null, error: r.error || null };
  healthCache = { at: now, value };
  return value;
}

function send(res, code, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(code, {
    'Content-Type': 'application/json',
    'Content-Length': Buffer.byteLength(body),
    // /health is a GET whose answer changes from minute to minute. Without
    // this the browser is entitled to keep the first reply it saw and
    // serve it back for the life of the page - which would show Outlook as
    // unavailable long after it came back, and no amount of re-asking on
    // this side would ever be seen.
    'Cache-Control': 'no-store, no-cache, must-revalidate',
    'Pragma': 'no-cache'
  });
  res.end(body);
}

function corsHeaders(origin) {
  return {
    'Access-Control-Allow-Origin': origin,
    'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, X-Hub-Mail',
    'Access-Control-Max-Age': '600',
    'Vary': 'Origin'
  };
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let size = 0; const chunks = [];
    req.on('data', c => {
      size += c.length;
      if (size > MAX_BODY) { reject(new Error('Payload too large')); req.destroy(); return; }
      chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

const server = http.createServer(async (req, res) => {
  const origin = req.headers.origin || '';
  const allowed = ALLOWED_ORIGINS.has(origin);

  if (req.method === 'OPTIONS') {
    // Refusing the preflight for an unknown origin is what makes the guard
    // header unreachable to a hostile page.
    if (!allowed) { res.writeHead(403); return res.end(); }
    res.writeHead(204, corsHeaders(origin));
    return res.end();
  }

  if (!allowed) {
    log('refused request from origin:', origin || '(none)');
    return send(res, 403, { error: 'Origin not allowed' });
  }
  Object.entries(corsHeaders(origin)).forEach(([k, v]) => res.setHeader(k, v));

  if (!req.headers[GUARD_HEADER]) {
    return send(res, 403, { error: 'Missing guard header' });
  }

  try {
    if (req.method === 'GET' && req.url.startsWith('/health')) {
      return send(res, 200, await health(req.url.includes('force=1')));
    }

    if (req.method === 'POST' && req.url.startsWith('/send')) {
      const raw = await readBody(req);
      let parsed;
      try { parsed = JSON.parse(raw); }
      catch { return send(res, 400, { error: 'Body is not valid JSON' }); }

      const messages = Array.isArray(parsed.messages) ? parsed.messages : null;
      if (!messages || !messages.length) return send(res, 400, { error: 'No messages' });
      if (messages.some(m => !m || !m.to || !m.subject)) {
        return send(res, 400, { error: 'Every message needs at least a recipient and a subject' });
      }

      const { fresh, duplicates } = partitionDuplicates(messages);
      if (duplicates.length) {
        log('refused', duplicates.length, 'duplicate(s) within 15 min:', duplicates.map(m => m.to).join(', '));
      }
      if (!fresh.length) {
        return send(res, 200, { ok: true, sent: 0, failed: [], skippedDuplicates: duplicates.length,
          note: 'Already sent in the last 15 minutes - nothing sent again.' });
      }

      log('sending', fresh.length, 'message(s) to:', fresh.map(m => m.to).join(', '));
      const result = await runPowerShell(fresh, false);
      log('result:', JSON.stringify(result).slice(0, 300));
      if (result.sent > 0 || result.ok) rememberSends(fresh, result.failed);
      result.skippedDuplicates = duplicates.length;
      healthCache = { at: 0, value: null };   // a real run is better evidence than a probe
      return send(res, 200, result);
    }

    return send(res, 404, { error: 'Not found' });
  } catch (e) {
    return send(res, 500, { error: e.message });
  }
});

/** Ask Outlook whether it can send, giving it time to finish starting. */
async function warmUp(attemptsLeft) {
  const h = await health(true).catch(e => ({ ready: false, error: e.message }));
  if (h.ready) { log('outlook: ready as', h.account); return; }
  if (attemptsLeft <= 1) {
    log('outlook: not ready -', h.error);
    // Do not let a stale failure linger: the next question re-asks.
    healthCache = { at: 0, value: null };
    return;
  }
  log('outlook: not ready yet, retrying in 20s -', h.error);
  setTimeout(() => warmUp(attemptsLeft - 1), 20000);
}

function start() {
  server.on('error', (e) => {
    if (e.code === 'EADDRINUSE') log('port', PORT, 'already in use - bridge not started');
    else log('server error:', e.message);
  });
  server.listen(PORT, HOST, () => {
    log('listening on http://' + HOST + ':' + PORT);
    log('allowed origins:', [...ALLOWED_ORIGINS].join(', '));
    // Warm the answer now rather than on the first question. Checking
    // Outlook means starting it over COM, which takes the better part of
    // ten seconds - long enough that a page asking on load would give up
    // and report the helper as unavailable when it is merely thinking.
    //
    // It retries because this runs at login, when it is competing with
    // everything else Windows starts. COM refusing at second three and
    // working fine at second forty is normal, and a single attempt would
    // record the wrong answer about a working setup.
    warmUp(4);
  });
}

module.exports = { start };
