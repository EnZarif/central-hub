// Central Hub local helper.
//
// This process exists for one reason: a browser cannot send email, and the
// Hub needs to hand finished operator emails to desktop Outlook. It runs on
// this machine, started at login by start-helper.vbs, and listens on
// loopback for the Hub to call.
//
// It used to do a second job - tailing Claude Code session transcripts and
// pushing live agent status to Supabase for a "Pixel Office" widget. That
// is gone, and deliberately. The table it wrote to was readable and
// writable by anyone holding the publishable key, and that key is in the
// Hub's page source on the public web. Anyone who looked could read the
// local filesystem paths of every project, which included the account name.
// The widget was decorative; the exposure was not.
//
// Removing it also took the Supabase client out of this process, so the
// helper now has no network dependency and no credentials of any kind.

try {
  require('./mail-bridge').start();
} catch (err) {
  console.error('[mail-bridge] failed to start:', err.message);
  process.exit(1);
}
