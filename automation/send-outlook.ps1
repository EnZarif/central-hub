<#
.SYNOPSIS
  Sends the Hub's operator emails through the desktop Outlook profile.

.DESCRIPTION
  The Hub builds each email in the browser - address, CC, subject and a rich
  HTML table - but a browser cannot send mail. Until now that gap was filled
  by a mailto: link per operator, which meant opening fifteen windows and
  pressing Send fifteen times.

  This closes the gap the only way that needs no tenant permission, no app
  registration and no admin consent: it drives the copy of Outlook already
  installed and signed in on this machine, through COM, as you.

  IMPORTANT - which Outlook. COM is exposed by CLASSIC Outlook
  (OUTLOOK.EXE, Office 16). The New Outlook store app has no automation
  interface at all, so if your mailbox only exists there, the preflight
  below will tell you so rather than failing halfway through a run.

.PARAMETER JsonPath
  A UTF-8 JSON file: { "messages": [ { to, cc, subject, html }, ... ] }.
  Passed as a file rather than an argument because the HTML bodies are far
  past any sane command-line length.

.PARAMETER WhatIf
  Build every message and report, but call nothing. Used by the Hub's
  preflight so a dry run can never put mail in the outbox.

.OUTPUTS
  One JSON object on stdout: { ok, sent, failed:[{to,error}], error }.
  Everything else goes to stderr so stdout stays machine-readable.
#>
param(
  [Parameter(Mandatory = $true)][string]$JsonPath,
  [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

function Emit($obj) {
  # Depth matters: the failed[] entries are nested objects.
  $obj | ConvertTo-Json -Depth 6 -Compress
  exit 0
}

if (-not (Test-Path $JsonPath)) {
  Emit @{ ok = $false; sent = 0; failed = @(); error = "Payload not found: $JsonPath" }
}

try {
  $payload = Get-Content -Raw -Encoding UTF8 $JsonPath | ConvertFrom-Json
} catch {
  Emit @{ ok = $false; sent = 0; failed = @(); error = "Payload is not valid JSON: $($_.Exception.Message)" }
}

$messages = @($payload.messages)
if ($messages.Count -eq 0) {
  Emit @{ ok = $false; sent = 0; failed = @(); error = 'No messages in payload.' }
}

# ---- preflight -------------------------------------------------------------
# Every reason sending cannot work is cheaper to find now than after the
# first message has already gone.
try {
  $ol = New-Object -ComObject Outlook.Application
} catch {
  Emit @{ ok = $false; sent = 0; failed = @(); error = 'Classic Outlook is not available over COM on this machine. The New Outlook app cannot be automated.' }
}

try {
  $ns = $ol.GetNamespace('MAPI')
  $ns.Logon('', '', $false, $false) | Out-Null
} catch {
  Emit @{ ok = $false; sent = 0; failed = @(); error = "Could not sign in to the Outlook profile: $($_.Exception.Message)" }
}

if ($ns.Accounts.Count -eq 0) {
  Emit @{ ok = $false; sent = 0; failed = @()
          error = 'Classic Outlook has no mail account configured, so there is nothing to send from. Open Outlook (classic) once and sign in with your work account, then try again.' }
}

$account = $ns.Accounts.Item(1)

if ($WhatIf) {
  Emit @{ ok = $true; sent = 0; failed = @(); whatIf = $true
          account = $account.SmtpAddress; built = $messages.Count
          error = $null }
}

# ---- send ------------------------------------------------------------------
$sent = 0
$failed = @()

foreach ($m in $messages) {
  try {
    if (-not $m.to) { throw 'No recipient address.' }
    $item = $ol.CreateItem(0)
    $item.To = [string]$m.to
    if ($m.cc) { $item.CC = [string]$m.cc }
    $item.Subject = [string]$m.subject
    # HTMLBody, not Body: the whole point of these emails is the table.
    $item.HTMLBody = [string]$m.html
    # Bind the message to the account explicitly rather than relying on
    # whichever profile default happens to be first.
    $item.SendUsingAccount = $account
    $item.Send()
    $sent++
  } catch {
    $failed += @{ to = [string]$m.to; error = $_.Exception.Message }
  }
}

# Send() queues into the outbox; this pushes it out without waiting for
# Outlook's own send/receive cycle, which can be minutes.
try { $ns.SendAndReceive($false) } catch { }

Emit @{ ok = ($failed.Count -eq 0); sent = $sent; failed = $failed
        account = $account.SmtpAddress; error = $null }
