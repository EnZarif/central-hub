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
const HEALTH_TTL_MS = 5 * 60 * 1000;

let healthCache = { at: 0, value: null };

function log(...a) { console.log('[mail-bridge]', ...a); }

/** Run the sender. `whatIf` builds everything and sends nothing. */
function runPowerShell(messages, whatIf) {
  return new Promise((resolve) => {
    let tmp;
    try {
      tmp = path.join(os.tmpdir(), 'hub-mail-' + Date.now() + '-' + Math.random().toString(36).slice(2) + '.json');
      fs.writeFileSync(tmp, JSON.stringify({ messages }), 'utf8');
    } catch (e) {
      return resolve({ ok: false, sent: 0, failed: [], error: 'Could not stage payload: ' + e.message });
    }

    const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                  '-File', PS_SCRIPT, '-JsonPath', tmp];
    if (whatIf) args.push('-WhatIf');

    execFile('powershell', args, { timeout: 180000, maxBuffer: 4 * 1024 * 1024 },
      (err, stdout, stderr) => {
        // Best effort: the payload holds real addresses, so it does not
        // linger in temp any longer than the run that needed it.
        try { fs.unlinkSync(tmp); } catch { /* nothing to do */ }
        if (err && !stdout) {
          return resolve({ ok: false, sent: 0, failed: [],
                           error: 'PowerShell failed: ' + (stderr || err.message).slice(0, 400) });
        }
        try {
          resolve(JSON.parse(String(stdout).trim()));
        } catch {
          resolve({ ok: false, sent: 0, failed: [],
                    error: 'Unreadable reply from the sender: ' + String(stdout).slice(0, 300) });
        }
      });
  });
}

async function health(force) {
  const now = Date.now();
  if (!force && healthCache.value && now - healthCache.at < HEALTH_TTL_MS) return healthCache.value;
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
  res.writeHead(code, { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) });
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

      log('sending', messages.length, 'message(s) to:', messages.map(m => m.to).join(', '));
      const result = await runPowerShell(messages, false);
      log('result:', JSON.stringify(result).slice(0, 300));
      healthCache = { at: 0, value: null };   // a real run is better evidence than a probe
      return send(res, 200, result);
    }

    return send(res, 404, { error: 'Not found' });
  } catch (e) {
    return send(res, 500, { error: e.message });
  }
});

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
    health(true)
      .then(h => log('outlook:', h.ready ? ('ready as ' + h.account) : ('not ready - ' + h.error)))
      .catch(e => log('outlook preflight failed:', e.message));
  });
}

module.exports = { start };
