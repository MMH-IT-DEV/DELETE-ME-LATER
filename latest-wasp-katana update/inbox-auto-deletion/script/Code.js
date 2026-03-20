// ─── DRY RUN MODE ───────────────────────────────────────────────
// Set to true to DELETE emails. Set to false to only LOG what would be deleted.
var DRY_RUN = true;

var PROTECT_STARRED = true;

function purgeOldMail30d() {
  var query = 'in:anywhere older_than:30d -in:trash -in:spam -in:chats';
  if (PROTECT_STARRED) query += ' -label:starred';

  var BATCH = 100;
  var SLEEP_MS = 400;
  var processed = 0;

  while (true) {
    var threads = GmailApp.search(query, 0, BATCH);
    if (threads.length === 0) break;

    if (DRY_RUN) {
      for (var i = 0; i < threads.length; i++) {
        var msg = threads[i].getMessages()[0];
        console.log('[DRY RUN] Would delete: ' + msg.getSubject() + ' | From: ' + msg.getFrom() + ' | Date: ' + msg.getDate());
      }
    } else {
      GmailApp.moveThreadsToTrash(threads);
    }

    processed += threads.length;
    console.log((DRY_RUN ? '[DRY RUN] ' : '') + 'Processed ' + processed + ' threads...');

    Utilities.sleep(SLEEP_MS);

    if (processed >= 3000) break;
  }

  console.log('Done. Total threads ' + (DRY_RUN ? 'that would be deleted: ' : 'moved to Trash: ') + processed);
}
