/**
 * Bugs — Apps Script backend bridge
 * ---------------------------------------------------
 * One Web App that the front-end talks to for everything:
 *   - reading and writing the two issue sheets (Course Errors, Tech Issues)
 *   - reading the Instructors sheet
 *   - uploading screenshots to Drive
 *   - running the Anthropic AI extraction server-side (key stays hidden)
 *   - posting high-priority issues to Slack
 *
 * Issues are split across two tabs by their "category":
 *   course_error -> Course Errors tab
 *   tech_issue   -> Tech Issues tab
 * Editing an issue and changing its category moves the row to the other tab.
 *
 * Before this works, do three things (see the Deployment Guide):
 *   1. Run setup() once to build the sheets and headers.
 *   2. Add your Anthropic key under Project Settings > Script Properties
 *      as ANTHROPIC_API_KEY.
 *   3. Deploy > New deployment > Web app, "Execute as: Me",
 *      "Who has access: Anyone", then copy the /exec URL.
 */

// ---- Config ---------------------------------------------------------------

var SHEET_ID = '1neDxEVI-VHu7PCNt59-fdyZAyHVng7D8FaRESjM3fjk';
var DRIVE_FOLDER_ID = '1E9LB6phKF3VmJR-VsbTzeW7jorqZdrvd';
// The Slack webhook lives ONLY in Project Settings > Script Properties
// (SLACK_WEBHOOK_URL), like the Anthropic key. No fallback constant: Code.gs
// is committed to the public GitHub repo now (GitHub's push protection caught
// the old hardcoded URL on the way in), and the old webhook had leaked into
// planning docs anyway, so rotate it in Slack and keep the property current.
// ---- What actually goes to Slack -------------------------------------------
// Edd, 18 August 2026: "Only high priority new issues (or old issues becoming
// high priority). So it is just a warning of a new major issue." Eleven senders
// were posting into the one channel, and the volume was burying the single
// message that needs somebody to move. Nothing is lost by muting them: every
// one of these notices still exists in the tracker, on the Actions tab or in
// its own queue. This only decides what gets PUSHED at people.
//
// Turn one back on by flipping it to true. Nothing else needs changing.
var SLACK_NOTICES = {
  // on:  is this notice allowed out at all
  // to:  the script property holding ITS channel's webhook. Empty, or a property
  //      that has not been set yet, falls back to the current channel, so a
  //      notice keeps working from the day it is switched on and moves the day
  //      the property is filled in. No code change and no deploy to move one.
  high_priority:     { on: true,  to: '' },                           // stays put
  shipping_chase:    { on: true,  to: 'SLACK_INSTRUCTING_DAILY' },
  notify_student:    { on: true,  to: 'SLACK_INSTRUCTING_DAILY' },
  query_raised:      { on: true,  to: 'SLACK_INSTRUCTING_DAILY' },
  query_answered:    { on: true,  to: 'SLACK_AUREUS_TECH' },
  shared_workaround: { on: true,  to: 'SLACK_INSTRUCTING_UPDATES' },
  // Off for good (Edd, 19 Aug 2026). All five still exist in the tracker.
  feedback:          { on: false, to: '' },
  weekly_digest:     { on: false, to: '' },
  scan_summary:      { on: false, to: '' },
  monthly_checklist: { on: false, to: '' },
  recheck:           { on: false, to: '' }
};
function slackOn_(kind) { return !!(SLACK_NOTICES[kind] && SLACK_NOTICES[kind].on); }
// The channel this notice belongs in, or the current one until that channel
// exists. Never throws and never blocks the caller: a Slack failure has never
// been allowed to stop a save.
function slackUrlFor_(kind) {
  var n = SLACK_NOTICES[kind];
  if (n && n.to) {
    var url = PropertiesService.getScriptProperties().getProperty(n.to);
    if (url) return url;
  }
  return slackWebhook_();
}
// Every Slack message in the app goes through here. One choke point, so what
// the channel carries is decided in the map above and nowhere else.
function slackPost_(kind, text) {
  if (!slackOn_(kind)) return;
  var url = slackUrlFor_(kind);
  if (!url) return;
  try {
    UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ text: text })
    });
  } catch (e) {}
}

function slackWebhook_() {
  return PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL') || '';
}

var COURSE_SHEET = 'Course Errors';
var TECH_SHEET = 'Tech Issues';
var SHIPPING_SHEET = 'Shipping Issues';
var INSTRUCTORS_SHEET = 'Instructors';
// Shipping gets its own tab like the other two: same columns, different
// lifecycle (chased on a date rather than fixed). setup() creates it.
var ISSUE_SHEETS = [COURSE_SHEET, TECH_SHEET, SHIPPING_SHEET];

// Feedback on the tracker itself (bugs/improvements suggested by users).
var FEEDBACK_SHEET = 'Feedback';
// New columns go on the END of this list, never inserted, so existing rows keep
// their data where the sheet already has it.
var FEEDBACK_HEADERS = ['id', 'created_at', 'user_email', 'user_name', 'message', 'image_urls', 'status', 'context', 'ref', 'kind', 'urgency'];
// context = JSON snapshot of where the Feedback button was pressed (view, open
// issue, filters, viewport, browser, recent JS errors, recent API calls, and
// the last few clicks), so a report carries its own crime scene.
// ref     = short human handle, FB-0042, so people can say one out loud.
// kind    = bug | idea | question. urgency = blocking | normal.

// The Feedback sheet gains columns as the tracker grows, and setup() isn't
// always run the moment a new one appears. This tops the header row up in
// place (appending only) so a fresh column starts working straight away.
function ensureFeedbackHeaders_(sheet) {
  var width = Math.max(1, sheet.getLastColumn());
  var head = sheet.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h || ''); });
  var missing = FEEDBACK_HEADERS.filter(function (h) { return head.indexOf(h) < 0; });
  if (!missing.length) return;
  var start = head.length + 1;
  sheet.getRange(1, start, 1, missing.length).setValues([missing]).setFontWeight('bold');
}

// Anthropic models. Extraction and the small matching/dedupe/troubleshooting
// calls all run on Haiku: it is fast, cheap, and returns clean, complete JSON.
// We tried Sonnet 5 for extraction to split long threads better, but it reasons
// by default and the thinking ate the token budget, so the JSON came back
// truncated mid-object and the parse failed. If a stronger extraction model is
// used later, give it a much larger max_tokens to leave room for the JSON after
// the thinking, and confirm it still returns a single clean JSON object.
// Upgraded from Haiku 4.5 (25 Jul 2026) after a couple of poor fix
// suggestions: Sonnet is markedly better at the judgement calls (does this
// really match a past fix? is this one issue or three?) for a few extra
// seconds per call. Volume is small, so the cost difference is pennies.
var ANTHROPIC_MODEL = 'claude-sonnet-5';
var EXTRACTION_MODEL = 'claude-sonnet-5';
// Student-facing drafts use the strongest model: they carry an instructor's
// name and voice, so quality beats cost (Edd, 8 Aug).
var DRAFT_MODEL = 'claude-opus-5';

// Deploys: pushing a change to Code.gs on main now auto-deploys the backend
// via .github/workflows/deploy-backend.yml (GitHub Action -> deployBackend).
// No browser needed. index.html is served by GitHub Pages on the same push.
// (CI pipeline set up 27 Jul 2026.)

// A "Resolved - TBC" issue auto-resolves after this many days of silence (no
// further reports or objections). The timer is the issue's updated_at, so any
// new activity resets it. 14 days per Edd (21 Jul); the sweep runs on a daily
// trigger created by ensureTriggers_.
var TBC_AUTO_RESOLVE_DAYS = 7;  // aligned with what the drawer has always promised ("auto-resolves after 7 days"); it sat at 14 while the copy said 7, which is why TBCs looked stuck in Actions (Edd's sweep, 8 Aug)

// One student sorted out by switching browser is a one-off. Three in a week all
// sorted the same way is a fault everyone is hitting, and each one closing
// quietly on its own is exactly how we miss it (31 Jul: six went by in a single
// morning before anyone joined the dots). So we watch the workarounds going out
// and shout once when the same one keeps working.
var SHARED_WORKAROUND_DAYS = 7;   // rolling window we look back over
var SHARED_WORKAROUND_MIN = 3;    // how many before it stops being coincidence

// Column order for both issue sheets (A..V). Both tabs use the same columns
// so the code can treat them the same.
var HEADERS = [
  'issue_id',          // A
  'submitted_at',      // B
  'updated_at',        // C
  'instructor_name',   // D
  'category',          // E  course_error | tech_issue  (WHAT is broken)
  'raw_text',          // F
  'student_name',      // G
  'student_contact',   // H
  'device_info',       // I  device / OS / browser (student's setup)
  'course',            // J
  'module',            // K
  'lesson',            // L
  'lesson_code',       // M
  'issue_type',        // N
  'summary',           // O
  'priority',          // P
  'priority_reason',   // Q
  'image_urls',        // R
  'status',            // S
  'resolved_at',       // T
  'resolution_note',   // U
  'notified_students', // V
  'report_count',      // W  how many separate reports are rolled into this issue
  'reports_json',      // X  JSON array, one entry per report (student + who logged it)
  'dev_passed_at',     // Y  when the issue was handed to developers (blank if not)
  'dev_fixed_at',      // Z  when a developer marked it fixed
  'dev_notes',         // AA developer's notes on the fix
  'checklist_json',    // AB pre-dev troubleshooting checklist state (tech issues), JSON map of item -> done | na | todo
  'request_kind',      // AC fix | improvement  (improvement = a feature/enhancement request, handled as a calmer backlog)
  'assignee',          // AD name of the person this is assigned to fix (course dev or developer), blank if unassigned
  'media_kind',        // AE course issues: video | text | other  (which part of the lesson)
  'double_checked',    // AF course issues: true if the submitter verified it themselves, not just the student's word
  'impact',            // AG improvements: low | medium | high  (rough impact, for backlog prioritisation)
  'section',           // AH which part of the platform: website | instructor_portal | partner_portal | course_player | app | other (mainly for tech and internal issues)
  'dev_query',         // AI open question from the developer/course team to the admins (blank = none open)
  'dev_query_at',      // AJ when that question was raised (blank = none open); cleared once answered
  'dev_query_by',      // AK who raised it
  'dev_query_target',  // AL who the question is for: 'admins' (default) or 'instructor' (the one who logged it)
  'platform',          // AM where the bug shows: 'browser' | 'app' | 'both' | '' (tick boxes on the form)
  'recheck_at',        // AN when to remind the instructor to double-check a Resolved-TBC (blank = no reminder)
  // WHO it hits, kept separate from WHAT is broken (Edd, 26 Jul): an
  // instructor-portal bug is a tech issue AND internal, and the old single
  // category forced a choice between them. 'internal' used to be a third
  // category; migrateAudience() moved those rows across.
  // APPENDED, never inserted - the column order here IS the sheet order.
  'audience',          // AO student | internal
  // Shipping (Edd, 30 Jul). A parcel problem is neither course content nor
  // platform tech, and it needs chasing on a date rather than fixing. The
  // tracking number is the merge key: DHL opens a fresh email thread for the
  // same consignment, so threads two, three and four fold into one issue.
  'courier',           // AP DHL | Royal Mail | ...
  'tracking_number',   // AQ the consignment reference, uppercased, spaces stripped
  'chase_at',          // AR ISO date when this needs poking again (blank = not waiting on anyone)
  // The student is sorted even though the fault is not (Edd, 2 Aug). A missing
  // Extend button that we worked round by extending the account by hand leaves
  // a real bug for the developers, but nobody to chase or update. Kept separate
  // from status so the issue can stay open without the student sitting in the
  // notify and "still on it" queues.
  'student_sorted',    // AS true = no further contact needed with this student
  // How big a job the fix looks (Edd, FB-0164): small | medium | large, set and
  // edited by the developers from their detail pane, so the queue can be sorted
  // and filtered by effort as well as priority. Blank = not sized yet.
  // APPENDED, never inserted - the column order here IS the sheet order.
  'fix_size',          // AT small | medium | large | ''
  // The reasoned next action, cached (Edd, FB-0203). Picking the next unticked
  // checklist box is not thinking: a student who has sent videos, tried three
  // browsers and updated iOS must not be told to try a different network just
  // because that line is untried. nextActionAi_ reads the whole thread instead.
  // Cached here because the tracker bills real API credits, so an unchanged
  // issue must cost nothing to open; the stored signature says what it was
  // read from. Deliberately LEFT OUT of the list projection (issueListRow_) so
  // the board stays as fast as Round 54 made it.
  // APPENDED, never inserted - the column order here IS the sheet order.
  'next_action_json',  // AU {action, why, instructor_side, student_ask, sig, at}
  // Is there a student on the other end of this report at all? (Edd, FB-0207:
  // "I specifically didn't attach this report to a student and yet you gave
  // things for the student to try.") Set from what the form held when it was
  // logged, not re-guessed later, so the checklist, the things-to-try and the
  // next-action reasoner all read the same answer. Blank on every row logged
  // before this column existed, which reads as "work it out from the name and
  // contact" rather than as a no.
  // APPENDED, never inserted - the column order here IS the sheet order.
  'student_involved',  // AV yes | no | '' (unknown, older rows)
  // Round 61. Who logged it, by account rather than by display name. Every
  // ownership column we have (instructor_name, assignee, resolved_by) stores a
  // name someone typed, so renaming a person in the Users sheet quietly orphans
  // their history, and we already have a Charly and a Charlie. The email is the
  // account's own key, stamped server-side from the session, so it cannot be
  // spoofed or drift. Blank on every row logged before this column existed.
  'instructor_email',  // AW
  // Round 61. The Chatwoot contact id for the student on this issue. We were
  // already fetching this on import and throwing it away, then searching our
  // way back to it by email every time we needed it - and that search fell back
  // to the first arbitrary result when nothing matched. This is the one stable
  // student identifier we can actually get hold of, so we keep it.
  'chatwoot_contact_id', // AX
  // FB-0249/0250 (Charlie): "in the tech issue for dev, there is a lot of info
  // there which is great, but this needs to be summarised with a clear ask /
  // action for the developer. they can then look through the report for further
  // info if needed." Written by whoever hands the issue over; the pane falls
  // back to one assembled from the fields when it is blank, so a hand-off never
  // arrives with nothing at the top.
  // APPENDED, never inserted - the column order here IS the sheet order.
  'dev_ask'            // AY one or two plain sentences: what we need doing
];

// The fixed pre-developer troubleshooting checklist for tech issues. Each item
// can end up done (tried), na (not relevant, e.g. an app-only step on a browser
// issue), or todo (relevant but not yet done). The AI pre-fills these from the
// conversation; a person can adjust any of them. Order here is the order shown.
// staff: does the step still apply when one of us hits a fault on our own
// systems, with no student anywhere? A hard refresh does; checking a student's
// email spelling does not. staffLabel rewords the survivors that are phrased at
// a student. Keep this list in step with CHECKLIST in index.html.
var CHECKLIST_ITEMS = [
  { id: 'confirm_error',          group: 'Identify and record', scope: 'both',    label: "Confirmed exactly what's failing / what the student sees (screenshot if useful)", staff: true, staffLabel: "Confirmed exactly what's failing and what you're seeing (screenshot if useful)" , todoLabel: "Confirm exactly what's failing and what the student sees (a screenshot helps)", staffTodoLabel: "Confirm exactly what's failing and what you're seeing (a screenshot helps)" },
  { id: 'noted_device',           group: 'Identify and record', scope: 'both',    label: 'Noted device make, model, OS version, and browser or app', staff: true, staffLabel: 'Noted which browser, device and OS version you were on' , todoLabel: "Get the device make, model, OS version, and browser or app", staffTodoLabel: "Note which browser, device and OS version you are on" },
  { id: 'replicated',             group: 'Identify and record', scope: 'both',    label: 'Tried the same course, lesson and portal yourself, on your own account and device', staff: true, staffLabel: "Asked someone else on the team to try the same thing, to see if it's just you" , todoLabel: "Try the same course, lesson and portal yourself, on your own account and device", staffTodoLabel: "Ask someone else on the team to try the same thing, to see if it is just you" },
  { id: 'right_place',            group: 'Account and login',   scope: 'both',    label: 'Logging in via the right place (correct partner portal vs ardent-training.com)' , todoLabel: "Check they are logging in via the right place (the correct partner portal, or ardent-training.com)" },
  { id: 'email_correct',          group: 'Account and login',   scope: 'both',    label: "Email spelled correctly, and it's the one they registered with" , todoLabel: "Check the email is spelled correctly, and that it is the one they registered with" },
  { id: 'password_reset',         group: 'Account and login',   scope: 'both',    label: 'Tried "forgot password", then typed email and password manually (no copy-paste)' , todoLabel: "Get them to use \"forgot password\", then type the email and password by hand (no copy-paste)" },
  { id: 'social_signin_password', group: 'Account and login',   scope: 'app',     label: 'Social sign-in: created a password via "organisation -> forgot password"' , todoLabel: "Social sign-in: get them to create a password via \"organisation -> forgot password\"" },
  { id: 'refreshed',              group: 'Standard fixes',      scope: 'browser', label: 'Refreshed the page, then a hard refresh (Ctrl+Shift+R, or Cmd+Shift+R on Mac)', staff: true , todoLabel: "Refresh the page, then try a hard refresh (Ctrl+Shift+R, or Cmd+Shift+R on a Mac)" },
  { id: 'logout_login',           group: 'Standard fixes',      scope: 'both',    label: 'Logged out and back in', staff: true , todoLabel: "Log out and back in" },
  { id: 'restart_device',         group: 'Standard fixes',      scope: 'both',    label: 'Restarted the device (or closed and reopened the app/browser)', staff: true , todoLabel: "Restart the device, or close and reopen the app or browser" },
  { id: 'clear_cache',            group: 'Standard fixes',      scope: 'browser', label: 'Cleared cache / tried an incognito or private window', staff: true , todoLabel: "Clear the cache, or try an incognito or private window" },
  { id: 'app_updated',            group: 'Standard fixes',      scope: 'app',     label: 'Checked the app is up to date' , todoLabel: "Check the app is up to date" },
  { id: 'different_browser',      group: 'Standard fixes',      scope: 'browser', label: 'Tried a different browser', staff: true , todoLabel: "Try a different browser" },
  { id: 'different_device',       group: 'Standard fixes',      scope: 'both',    label: 'Tried a different device', staff: true , todoLabel: "Try a different device" },
  { id: 'different_network',      group: 'Standard fixes',      scope: 'both',    label: 'Tried a different network (mobile data, hotspot, or another wifi)', staff: true , todoLabel: "Try a different network (mobile data, a hotspot, or another wifi)" },
  { id: 'vpn_adblock',            group: 'Standard fixes',      scope: 'both',    label: 'Turned off any VPN, ad blocker, or content/parental filter', staff: true , todoLabel: "Turn off any VPN, ad blocker, or content/parental filter" },
  { id: 'storage_space',          group: 'Standard fixes',      scope: 'app',     label: "Checked there's free storage on the device (download/save problems)" , todoLabel: "Check there is free storage on the device (for download or save problems)" }
];
function checklistItemsFor_(staff) {
  return staff ? CHECKLIST_ITEMS.filter(function (it) { return it.staff; }) : CHECKLIST_ITEMS;
}
function checklistLabel_(it, staff) { return (staff && it.staffLabel) ? it.staffLabel : it.label; }
// The same item said as an instruction rather than a record. The checklist is a
// tick-list of what HAS been done, so its labels are past tense; a list of what
// to try next has to be the other way round or it reads as nonsense ("Noted
// device make, model" offered as a thing to go and do). Falls back to the
// checklist wording if an item has no future form yet, which is still better
// than nothing (Edd, 20 Aug 2026).
function checklistTodoLabel_(it, staff) {
  if (staff && it.staffTodoLabel) return it.staffTodoLabel;
  if (it.todoLabel) return it.todoLabel;
  return checklistLabel_(it, staff);
}

var INSTRUCTORS = [
  { name: 'Edd', email: 'ehewett@ardent-training.com' },
  { name: 'Charly', email: 'charly@ardent-training.com' },
  { name: 'Charlie', email: 'charlie@ardent-training.com' },
  { name: 'Stuart', email: 'stuart@ardent-training.com' },
  { name: 'Tom', email: 'tom@ardent-training.com' },
  { name: 'Michelle', email: 'michelle@ardent-training.com' },
  { name: 'Laura', email: 'laura@ardent-training.com' },
  { name: 'Luke', email: 'luke@ardent-training.com' },
  { name: 'Holly', email: 'holly@ardent-training.com' },
  { name: 'Peter', email: 'peter@ardent-training.com' }
];

// ---- Web app entry points -------------------------------------------------

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    var action = p.action || '';
    CURRENT_ACTION_ = action;
    if (action === 'ping') return jsonOut({ ok: true, time: new Date().toISOString(), backend: backendInfo_() });
    if (action === 'getInvite') return jsonOut(getInvite_(p.token));   // public: validate an invite link
    if (action === 'mirror') return jsonOut(mirror_(p));               // read-only, key-gated mirror for the local Cowork sync

    var user = userForToken_(p.token);
    if (!user) return jsonOut({ ok: false, error: 'unauthorized' });
    if (action === 'me') return jsonOut({ ok: true, user: publicUser_(user), backend: backendInfo_() });
    if (!hasPerm_(user, reqPerm_(action))) return jsonOut({ ok: false, error: 'forbidden' });

    if (action === 'bootstrap') return jsonOut(bootstrap_(user));
    if (action === 'getIssuesList') return jsonOut(getIssuesList_());
    if (action === 'getIssue') return jsonOut(getIssueFull_(p));
    if (action === 'getIssues') return jsonOut(getIssues_());
    if (action === 'getInstructors') return jsonOut(getInstructors_());
    if (action === 'listUsers') return jsonOut(listUsers_());
    if (action === 'getPlaybook') return jsonOut(getPlaybookEndpoint_());
    if (action === 'listPlaybookSuggestions') return jsonOut(listPlaybookSuggestions_());
    if (action === 'listKnownFixFlags') return jsonOut(listKnownFixFlags_());
    if (action === 'getFeedback') return jsonOut(getFeedback_());
    if (action === 'getAssignees') return jsonOut(listAssignees_());
    return jsonOut({ ok: false, error: 'Unknown GET action: ' + action });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);
    var action = body.action || '';
    CURRENT_ACTION_ = action;

    // Public auth actions (no session yet).
    // Self-deploy: gated by its own DEPLOY_KEY (script property), not a user
    // session — same pattern as the mirror. See deployBackend_ below.
    if (action === 'ping') return jsonOut({ ok: true, time: new Date().toISOString(), backend: backendInfo_() });
    if (action === 'getManifest') return jsonOut(getManifest_(body));
    if (action === 'deployBackend') return jsonOut(deployBackend_(body));
    if (action === 'setSlackWebhook') return jsonOut(setSlackWebhook_(body));
    if (action === 'setChatwootConfig') return jsonOut(setChatwootConfig_(body));
    if (action === 'runSetup') return jsonOut(runSetup_(body));
    if (action === 'runMigrateAudience') {
      var mk = PropertiesService.getScriptProperties().getProperty('DEPLOY_KEY');
      if (!mk || String(body.key || '') !== mk) return jsonOut({ ok: false, error: 'bad deploy key' });
      return jsonOut({ ok: true, result: migrateAudience() });
    }
    // The backtest harness (r46): key-gated maintenance machinery, not a user
    // feature. Read-only against Chatwoot; writes only to its own sheets
    // (KnownFixes, BacktestLog), never to Issues.
    if (action === 'backtest') return jsonOut(backtest_(body));

    if (action === 'login') return jsonOut(login_(body));
    if (action === 'acceptInvite') return jsonOut(acceptInvite_(body));
    if (action === 'requestPasswordReset') return jsonOut(requestPasswordReset_(body));

    var user = userForToken_(body.token);
    if (!user) return jsonOut({ ok: false, error: 'unauthorized' });
    if (action === 'logout') return jsonOut(logout_(body.token));
    if (!hasPerm_(user, reqPerm_(action))) return jsonOut({ ok: false, error: 'forbidden' });
    body._user = user;

    // Reads are served over POST too, so the front-end can keep the session
    // token in the request body rather than the URL (URLs leak into browser
    // history and logs; bodies don't). The GET versions below still work.
    if (action === 'reportsTicket') return jsonOut(reportsTicket_(user));
    if (action === 'me') return jsonOut({ ok: true, user: publicUser_(user), backend: backendInfo_() });
    if (action === 'bootstrap') return jsonOut(bootstrap_(user));
    if (action === 'getIssuesList') return jsonOut(getIssuesList_());
    if (action === 'getIssue') return jsonOut(getIssueFull_(body));
    if (action === 'getIssues') return jsonOut(getIssues_());
    if (action === 'getInstructors') return jsonOut(getInstructors_());
    if (action === 'listUsers') return jsonOut(listUsers_());
    if (action === 'getPlaybook') return jsonOut(getPlaybookEndpoint_());
    if (action === 'listPlaybookSuggestions') return jsonOut(listPlaybookSuggestions_());
    if (action === 'listKnownFixFlags') return jsonOut(listKnownFixFlags_());
    if (action === 'getFeedback') return jsonOut(getFeedback_());
    if (action === 'getAssignees') return jsonOut(listAssignees_());

    if (action === 'addIssue') return jsonOut(addIssue_(body));
    if (action === 'updateIssue') return jsonOut(updateIssue_(body));
    if (action === 'addUpdate') return jsonOut(addUpdate_(body));
    if (action === 'splitIssue') return jsonOut(splitIssue_(body));
    if (action === 'linkIssues') return jsonOut(linkIssues_(body));
    if (action === 'deleteIssue') return jsonOut(deleteIssue_(body));
    if (action === 'passToDev') return jsonOut(passToDev_(body));
    if (action === 'markDevFixed') return jsonOut(markDevFixed_(body));
    if (action === 'saveDevNotes') return jsonOut(saveDevNotes_(body));
    if (action === 'flagQuery') return jsonOut(flagQuery_(body));
    if (action === 'answerQuery') return jsonOut(answerQuery_(body));
    if (action === 'requestRecheck') return jsonOut(requestRecheck_(body));
    if (action === 'rateExtraction') return jsonOut(rateExtraction_(body));
    if (action === 'chatwootImport') return jsonOut(chatwootImport_(body));
    if (action === 'chatwootList') return jsonOut(chatwootList_(body));
    if (action === 'chatScanList') return jsonOut(chatScanList_());
    if (action === 'chatScanReview') return jsonOut(chatScanReview_(body));
    if (action === 'runChatScan') return jsonOut(runChatScan_(body));
    if (action === 'lessonIssueCounts') return jsonOut(lessonIssueCounts_());
    if (action === 'runChatBackSweep') return jsonOut(runChatBackSweep_(body));
    if (action === 'chatBackSweepState') return jsonOut(chatBackSweepState_());
    if (action === 'saveChecklist') return jsonOut(saveChecklist_(body));
    if (action === 'assignIssue') return jsonOut(assignIssue_(body));
    if (action === 'bulkAssign') return jsonOut(bulkAssign_(body));
    if (action === 'courseReview') return jsonOut(courseReview_(body));
    if (action === 'fetchStudentUpdate') return jsonOut(fetchStudentUpdate_(body));
    if (action === 'listLiveCases') return jsonOut(listLiveCases_(body));
    if (action === 'caseBrief') return jsonOut(caseBrief_(body));
    if (action === 'caseCheckReply') return jsonOut(caseCheckReply_(body));
    if (action === 'caseDraftReply') return jsonOut(caseDraftReply_(body));
    if (action === 'caseCheckpoint') return jsonOut(caseCheckpoint_(body));
    if (action === 'caseClose') return jsonOut(caseClose_(body));
    if (action === 'caseTouch') return jsonOut(caseTouch_(body));
    if (action === 'batchStudentDrafts') return jsonOut(batchStudentDrafts_(body));
    if (action === 'listContentSuggestions') return jsonOut(listContentSuggestions_());
    if (action === 'resolveContentSuggestion') return jsonOut(resolveContentSuggestion_(body));
    if (action === 'runConfusionReview') return jsonOut(runConfusionReview_(body));
    if (action === 'uploadImage') return jsonOut(uploadImage_(body));
    if (action === 'attachImages') return jsonOut(attachImages_(body));
    if (action === 'extract') return jsonOut(extract_(body));
    if (action === 'askIssues') return jsonOut(askIssues_(body));
    if (action === 'suggestFix') return jsonOut(suggestFix_(body));
    if (action === 'troubleshoot') return jsonOut(troubleshoot_(body));
    if (action === 'sameIssue') return jsonOut(sameIssue_(body));
    if (action === 'draftStudentMessage') return jsonOut(draftStudentMessage_(body));
    if (action === 'nextAction') return jsonOut(nextAction_(body));
    if (action === 'chatwootContactUrl') return jsonOut(chatwootContactUrl_(body));
    if (action === 'setVoiceGuide') return jsonOut(setVoiceGuide_(body));
    if (action === 'listVoiceGuides') return jsonOut(listVoiceGuides_());
    if (action === 'matchUpdate') return jsonOut(matchUpdate_(body));
    if (action === 'inviteUser') return jsonOut(inviteUser_(body));
    if (action === 'updateUser') return jsonOut(updateUser_(body));
    if (action === 'changePassword') return jsonOut(changePassword_(body));
    if (action === 'adminResetLink') return jsonOut(adminResetLink_(body));
    if (action === 'savePlaybook') return jsonOut(savePlaybook_(body));
    if (action === 'resolvePlaybookSuggestion') return jsonOut(resolvePlaybookSuggestion_(body));
    if (action === 'suggestPlaybook') return jsonOut(suggestPlaybook_(body));
    if (action === 'flagKnownFix') return jsonOut(flagKnownFix_(body));
    if (action === 'resolveKnownFixFlag') return jsonOut(resolveKnownFixFlag_(body));
    if (action === 'addFeedback') return jsonOut(addFeedback_(body));
    if (action === 'updateFeedback') return jsonOut(updateFeedback_(body));
    if (action === 'deleteFeedback') return jsonOut(deleteFeedback_(body));
    return jsonOut({ ok: false, error: 'Unknown POST action: ' + action });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

// ---- Accounts, sessions, permissions --------------------------------------
//
// Each person has their own account in the Users sheet with a salted, iterated
// password hash (never the plain password) and a set of permission flags that
// decide which pages they see. Requests carry a session token, not a passcode.

var USERS_SHEET = 'Users';
var USER_HEADERS = ['email', 'name', 'status', 'perms_json', 'pass_hash', 'pass_salt', 'invite_token', 'session_token', 'session_expires', 'created_at'];
var PERM_KEYS = ['log', 'manage', 'analytics', 'dev', 'course', 'users'];
var SESSION_DAYS = 30;
var HASH_ROUNDS = 2000;
// Password-reset links expire after this many minutes. The expiry is packed into
// the token itself (token + "." + expiry-millis), so no extra column is needed;
// a plain invite token with no "." is treated as non-expiring, as before.
var RESET_TOKEN_MINUTES = 60;
// Round 61: invite links now expire too. They used to be plain tokens with no
// "." in them, which tokenExpired_ reads as "never expires" - so a months-old
// invite email was still a live way into an account. Two weeks is long enough
// for someone to get round to setting a password, short enough that a forwarded
// or forgotten email stops being a key. An admin can always Resend.
var INVITE_TOKEN_DAYS = 14;
// An invite token in the same "<token>.<expiry-millis>" shape the reset links
// already use, so tokenExpired_ handles both without knowing the difference.
function newInviteToken_() {
  return newToken_() + '.' + (Date.now() + INVITE_TOKEN_DAYS * 24 * 3600 * 1000);
}

function usersSheet_() { return sheetByName_(USERS_SHEET); }

function rowToUser_(row, idx) {
  var u = {};
  USER_HEADERS.forEach(function (k) { u[k] = idx[k] != null ? row[idx[k]] : ''; });
  return u;
}
function findUserByField_(field, value) {
  var sheet = usersSheet_();
  if (!sheet) return null;
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  var head = values[0]; var idx = {}; head.forEach(function (h, i) { idx[h] = i; });
  var target = String(value || '');
  for (var r = 1; r < values.length; r++) {
    var cell = String(values[r][idx[field]] || '');
    if (field === 'email') { if (cell.trim().toLowerCase() !== target.trim().toLowerCase()) continue; }
    else if (cell !== target || !target) continue;
    return { row: r + 1, idx: idx, sheet: sheet, user: rowToUser_(values[r], idx) };
  }
  return null;
}
function findUserByEmail_(email) { return findUserByField_('email', email); }

function userForToken_(token) {
  if (!token) return null;
  var f = findUserByField_('session_token', token);
  if (!f) return null;
  if (String(f.user.status).toLowerCase() !== 'active') return null;
  var exp = new Date(f.user.session_expires);
  if (isNaN(exp.getTime()) || exp.getTime() < Date.now()) return null;
  return f.user;
}
function permsOf_(user) {
  var p = {};
  try { p = user.perms_json ? JSON.parse(user.perms_json) : {}; } catch (e) { p = {}; }
  return p;
}
function hasPerm_(user, req) {
  if (req === 'any') return true;
  // Round 61: the deny sentinel. reqPerm_ returns this for anything it does not
  // recognise, so a new action wired into doPost without a matching case is shut
  // rather than open. If a new action 403s, the fix is a case in reqPerm_.
  if (req === 'none') return false;
  var p = permsOf_(user);
  if (req === 'devcourse') return !!(p.dev || p.course);
  // 'work' = anyone who actually works issues, so they can tick the checklist
  // wherever it shows (Track or the Developers / Course queues).
  if (req === 'work') return !!(p.log || p.manage || p.dev || p.course);
  return !!p[req];
}
function reqPerm_(action) {
  switch (action) {
    case 'addIssue': case 'addUpdate': case 'extract': case 'suggestFix': case 'troubleshoot': case 'matchUpdate': case 'attachImages': case 'draftStudentMessage': case 'nextAction': case 'sameIssue': case 'chatwootContactUrl': return 'log';
    // updateIssue is 'work' so the dev/course team can retune priority from
    // their drawer; anything beyond priority still needs manage (checked
    // inside updateIssue_ itself).
    case 'updateIssue': return 'work';
    case 'splitIssue': case 'linkIssues': case 'askIssues': return 'manage';
    // Handing work to the developers is an admin call (or automatic on
    // submission); instructors log and manage, they don't route.
    case 'passToDev': return 'users';
    case 'markDevFixed': case 'saveDevNotes': return 'devcourse';
    // Anyone who works issues can raise a question (dev/course asking up, or an
    // admin asking the logging instructor for more info). Answering is gated
    // inside answerQuery_ itself: admin-targeted questions need the users perm,
    // instructor-targeted ones are answered by the instructor who logged it.
    case 'flagQuery': case 'answerQuery': return 'work';
    case 'requestRecheck': case 'rateExtraction': return 'log';
    case 'chatwootImport': case 'chatwootList': return 'log';
    // The scan queue is visible to every instructor (Edd, 26 Jul) - the team
    // is small and whoever spots it first should be able to act. Kicking off a
    // manual scan stays with the admins.
    case 'chatScanList': case 'chatScanReview': return 'log';
    case 'runChatScan': case 'runChatBackSweep': case 'chatBackSweepState': return 'users';
    // Read by the Reports page (19 Aug 2026) to set issue counts against
    // lesson traffic. Same permission as Reports itself, and it returns counts
    // only - no summaries, no students, nothing about any one report.
    case 'lessonIssueCounts': return 'analytics';
    case 'saveChecklist': case 'assignIssue': case 'getAssignees': return 'work';
    // The queue tools (Edd, FB-0165): reviewing and bulk-assigning are for
    // anyone who works a fix queue. Fetching a Chatwoot update sits with the
    // other update paths.
    case 'bulkAssign': case 'courseReview': return 'work';
    case 'fetchStudentUpdate': return 'log';
    // The Live Case workspace (Round 45): visible to every instructor-level
    // user, and cases are shared - anyone can pick one up and carry on.
    case 'listLiveCases': case 'caseBrief': case 'caseCheckReply': case 'caseDraftReply':
    case 'caseCheckpoint': case 'caseClose': case 'caseTouch': case 'batchStudentDrafts': return 'log';
    // Saying "this suggestion was wrong here" belongs to whoever was shown it,
    // so it sits at the same tier as the case. Approving the correction, and
    // therefore changing a corpus row, is Edd's alone (FB-0231).
    case 'flagKnownFix': return 'log';
    // Confusion -> content-tweak suggestions sit with anyone who works a queue.
    case 'listContentSuggestions': case 'resolveContentSuggestion': case 'runConfusionReview': return 'work';
    case 'inviteUser': case 'updateUser': case 'adminResetLink': case 'listUsers':
    case 'getPlaybook': case 'savePlaybook': case 'listPlaybookSuggestions': case 'resolvePlaybookSuggestion': case 'suggestPlaybook':
    case 'listKnownFixFlags': case 'resolveKnownFixFlag':
    case 'getFeedback': case 'updateFeedback': case 'deleteFeedback': case 'setVoiceGuide': case 'listVoiceGuides': return 'users';
    // Available to any logged-in user: feedback and its screenshots, and
    // changing your own password. These used to ride on the old open default;
    // they are listed here now so the default can shut.
    case 'uploadImage': case 'addFeedback': case 'changePassword': return 'any';
    // deleteIssue also used to ride on the default. It has always checked
    // ownership inside deleteIssue_ (your own, or the users permission), so this
    // is the outer gate only: you have to be someone who works issues at all.
    case 'deleteIssue': return 'work';
    case 'getIssues': case 'getIssuesList': case 'getIssue': case 'bootstrap':
    case 'getInstructors': case 'me': return 'any';
    // Ardent Reports vouching ticket. Reuses 'analytics' rather than adding a
    // permission key: it is already held by exactly the accounts that should
    // see Reports, and PERM_KEYS/the admin screen stay untouched.
    case 'reportsTicket': return 'analytics';
    // Fail closed. Every action reaching this point is one nobody listed above,
    // which means nobody decided who should be allowed to call it. The public
    // and key-gated actions (login, acceptInvite, requestPasswordReset, ping,
    // mirror, getInvite, getManifest, deployBackend, runSetup, setSlackWebhook,
    // setChatwootConfig, runMigrateAudience, backtest) all return from doGet /
    // doPost before this gate, so they are unaffected.
    default: return 'none';
  }
}
function publicUser_(user) {
  return { email: user.email, name: user.name, perms: permsOf_(user) };
}
function setCell_(f, key, value) { f.sheet.getRange(f.row, f.idx[key] + 1).setValue(value); }

function bytesToHex_(bytes) {
  var s = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = (bytes[i] + 256) % 256;
    s += ('0' + b.toString(16)).slice(-2);
  }
  return s;
}
function hashPassword_(password, salt) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + ':' + String(password));
  for (var i = 0; i < HASH_ROUNDS; i++) {
    digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + ':' + bytesToHex_(digest));
  }
  return bytesToHex_(digest);
}
function newToken_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); }
// A reset token carries its own expiry as "<token>.<expiry-millis>". Invite
// tokens have no "." and never expire, so this returns false for them.
function tokenExpired_(token) {
  var s = String(token || '');
  var dot = s.indexOf('.');
  if (dot < 0) return false;
  var exp = parseInt(s.slice(dot + 1), 10);
  return !isFinite(exp) || Date.now() > exp;
}

function startSession_(f) {
  var token = newToken_();
  setCell_(f, 'session_token', token);
  setCell_(f, 'session_expires', new Date(Date.now() + SESSION_DAYS * 24 * 3600 * 1000).toISOString());
  return token;
}

// Failed logins are throttled per email address, because the web app URL is
// public and without a brake someone could sit and guess passwords all day.
// After LOGIN_MAX_FAILS misses the address is locked out for LOGIN_LOCK_MINUTES
// (a successful login clears the count). A missing account returns the same
// "wrong email or password" as a wrong password, so the login box can't be
// used to probe which emails have accounts (the reset flow is already neutral).
var LOGIN_MAX_FAILS = 8;
var LOGIN_LOCK_MINUTES = 10;

function login_(body) {
  var email = String(body.email || '').trim().toLowerCase();
  var cache = CacheService.getScriptCache();
  var key = 'loginfail:' + email;
  var fails = parseInt(cache.get(key) || '0', 10) || 0;
  if (fails >= LOGIN_MAX_FAILS) {
    return { ok: false, error: 'Too many attempts. Wait ' + LOGIN_LOCK_MINUTES + ' minutes and try again.' };
  }
  var failed = function () {
    cache.put(key, String(fails + 1), LOGIN_LOCK_MINUTES * 60);
    return { ok: false, error: 'Wrong email or password.' };
  };
  var f = findUserByEmail_(email);
  if (!f) return failed();
  if (String(f.user.status).toLowerCase() === 'disabled') return { ok: false, error: 'This account has been disabled.' };
  if (!f.user.pass_hash) return { ok: false, error: 'This account is not set up yet. Use your invite link.' };
  var hash = hashPassword_(body.password || '', f.user.pass_salt);
  if (hash !== String(f.user.pass_hash)) return failed();
  cache.remove(key);
  var token = startSession_(f);
  return { ok: true, token: token, user: publicUser_(f.user) };
}

function getInvite_(token) {
  var f = findUserByField_('invite_token', token);
  if (!f) return { ok: false, error: 'This link is not valid (it may have already been used).' };
  if (String(f.user.status).toLowerCase() === 'disabled') return { ok: false, error: 'This account has been disabled.' };
  if (tokenExpired_(f.user.invite_token)) return { ok: false, error: 'This link has expired. Request a new one from the login page.' };
  return { ok: true, email: f.user.email, name: f.user.name };
}

function acceptInvite_(body) {
  if (!body.password || String(body.password).length < 8) return { ok: false, error: 'Password must be at least 8 characters.' };
  // Note: the invite token comes in as body.invite, not body.token, because the
  // front-end uses body.token for the session token on authed requests.
  var f = findUserByField_('invite_token', body.invite);
  if (!f) return { ok: false, error: 'This link is not valid (it may have already been used).' };
  if (String(f.user.status).toLowerCase() === 'disabled') return { ok: false, error: 'This account has been disabled.' };
  if (tokenExpired_(f.user.invite_token)) return { ok: false, error: 'This link has expired. Request a new one from the login page.' };
  var salt = Utilities.getUuid();
  setCell_(f, 'pass_salt', salt);
  setCell_(f, 'pass_hash', hashPassword_(body.password, salt));
  setCell_(f, 'status', 'active');
  setCell_(f, 'invite_token', '');
  var token = startSession_(f);
  var fresh = findUserByEmail_(f.user.email);
  return { ok: true, token: token, user: publicUser_(fresh.user) };
}

function logout_(token) {
  var f = findUserByField_('session_token', token);
  if (f) { setCell_(f, 'session_token', ''); setCell_(f, 'session_expires', ''); }
  return { ok: true };
}

function sendInviteEmail_(email, name, token) {
  var url = getAppUrl_();
  var link = (url || '(set APP_URL)') + (url && url.indexOf('?') > -1 ? '&' : '?') + 'invite=' + token;
  if (url) {
    var bodyText = 'Hi ' + (name || '') + ',\n\n' +
      'You have been invited to Bugs, the Ardent Training issue log. Open the link below to set your password and get started:\n\n' +
      link + '\n\nThanks,\nArdent Training';
    try { MailApp.sendEmail(email, 'Your Bugs invite', bodyText); } catch (e) {}
  }
  return link;
}

// ---- Password reset and change ------------------------------------------

// Public: a user asks for a reset link. Always returns ok, so we never reveal
// which emails have accounts. If the email matches an account that isn't
// disabled, set a short-lived reset token and email the link.
function requestPasswordReset_(body) {
  var email = String(body.email || '').trim().toLowerCase();
  if (email && email.indexOf('@') > -1) {
    // Throttle per address (any address, so it can't be used to probe which ones
    // exist): at most one reset email every few minutes.
    var cache = CacheService.getScriptCache();
    var key = 'reset:' + email;
    if (!cache.get(key)) {
      cache.put(key, '1', 300);
      var f = findUserByEmail_(email);
      if (f && String(f.user.status).toLowerCase() !== 'disabled') {
        var token = newToken_() + '.' + (Date.now() + RESET_TOKEN_MINUTES * 60 * 1000);
        setCell_(f, 'invite_token', token);
        try { sendResetEmail_(f.user.email, f.user.name, token); } catch (e) {}
      }
    }
  }
  return { ok: true };
}

function sendResetEmail_(email, name, token) {
  var url = getAppUrl_();
  var link = (url || '(set APP_URL)') + (url && url.indexOf('?') > -1 ? '&' : '?') + 'invite=' + token;
  if (url) {
    var bodyText = 'Hi ' + (name || '') + ',\n\n' +
      'Someone (hopefully you) asked to reset your Bugs password. Open the link below to choose a new one. It expires in ' + RESET_TOKEN_MINUTES + ' minutes.\n\n' +
      link + '\n\nIf you did not ask for this, you can ignore this email and your password will not change.\n\nThanks,\nArdent Training';
    try { MailApp.sendEmail(email, 'Reset your Bugs password', bodyText); } catch (e) {}
  }
  return link;
}

// Authed: the logged-in user changes their own password.
function changePassword_(body) {
  var user = body._user || {};
  var f = findUserByEmail_(user.email);
  if (!f) return { ok: false, error: 'Account not found.' };
  if (!f.user.pass_hash) return { ok: false, error: 'Set a password from your invite link first.' };
  var cur = hashPassword_(body.current_password || '', f.user.pass_salt);
  if (cur !== String(f.user.pass_hash)) return { ok: false, error: 'Your current password is not right.' };
  if (!body.new_password || String(body.new_password).length < 8) return { ok: false, error: 'New password must be at least 8 characters.' };
  var salt = Utilities.getUuid();
  setCell_(f, 'pass_salt', salt);
  setCell_(f, 'pass_hash', hashPassword_(body.new_password, salt));
  // Any outstanding reset link is now moot, so retire it.
  setCell_(f, 'invite_token', '');
  return { ok: true };
}

// Admin fallback: generate a reset link for a user (and email it too), for when
// someone cannot receive or use the self-service email. Returns the link so the
// admin can copy it directly.
function adminResetLink_(body) {
  var email = String(body.email || '').trim().toLowerCase();
  var f = email ? findUserByEmail_(email) : null;
  if (!f) return { ok: false, error: 'No account for that email.' };
  if (String(f.user.status).toLowerCase() === 'disabled') return { ok: false, error: 'That account is disabled. Enable it first.' };
  var token = newToken_() + '.' + (Date.now() + RESET_TOKEN_MINUTES * 60 * 1000);
  setCell_(f, 'invite_token', token);
  var link = sendResetEmail_(f.user.email, f.user.name, token);
  return { ok: true, link: link };
}

function inviteUser_(body) {
  var email = String(body.email || '').trim().toLowerCase();
  if (!email || email.indexOf('@') < 0) return { ok: false, error: 'Enter a valid email.' };
  var perms = {};
  PERM_KEYS.forEach(function (k) { perms[k] = !!(body.perms && body.perms[k]); });
  var token = newInviteToken_();
  var existing = findUserByEmail_(email);
  if (existing) {
    setCell_(existing, 'name', body.name || existing.user.name || '');
    setCell_(existing, 'perms_json', JSON.stringify(perms));
    setCell_(existing, 'status', 'invited');
    setCell_(existing, 'invite_token', token);
    setCell_(existing, 'pass_hash', '');
    setCell_(existing, 'session_token', '');
  } else {
    var u = {
      email: email, name: body.name || '', status: 'invited',
      perms_json: JSON.stringify(perms), pass_hash: '', pass_salt: '',
      invite_token: token, session_token: '', session_expires: '',
      created_at: new Date().toISOString()
    };
    usersSheet_().appendRow(USER_HEADERS.map(function (k) { return u[k]; }));
  }
  return { ok: true, invite_url: sendInviteEmail_(email, body.name || '', token) };
}

function listUsers_() {
  var sheet = usersSheet_();
  if (!sheet) return { ok: true, users: [] };
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return { ok: true, users: [] };
  var head = values[0]; var idx = {}; head.forEach(function (h, i) { idx[h] = i; });
  var out = [];
  for (var r = 1; r < values.length; r++) {
    if (!values[r][idx['email']]) continue;
    var u = rowToUser_(values[r], idx);
    out.push({ email: u.email, name: u.name, status: u.status, perms: permsOf_(u) });
  }
  return { ok: true, users: out };
}

function updateUser_(body) {
  var f = findUserByEmail_(body.email);
  if (!f) return { ok: false, error: 'No such user.' };
  if (body.perms) {
    var perms = {}; PERM_KEYS.forEach(function (k) { perms[k] = !!body.perms[k]; });
    setCell_(f, 'perms_json', JSON.stringify(perms));
  }
  if (body.name != null) setCell_(f, 'name', body.name);
  if (body.disable === true) {
    // Kill the session AND any pending invite/reset link, so a disabled person
    // can't walk back in via an outstanding token.
    setCell_(f, 'status', 'disabled');
    setCell_(f, 'session_token', '');
    setCell_(f, 'invite_token', '');
  }
  if (body.enable === true) setCell_(f, 'status', f.user.pass_hash ? 'active' : 'invited');
  if (body.resend === true) {
    var token = newInviteToken_();
    setCell_(f, 'invite_token', token);
    setCell_(f, 'status', 'invited');
    setCell_(f, 'pass_hash', '');
    setCell_(f, 'session_token', '');
    return { ok: true, invite_url: sendInviteEmail_(f.user.email, f.user.name, token) };
  }
  return { ok: true };
}

function saveDevNotes_(body) {
  var found = findRow_(body.issue_id);
  if (!found) return { ok: false, error: 'No issue found with id ' + body.issue_id };
  // Only touch what was actually sent: the dev pane saves notes and the ask
  // from two different boxes, and one must not blank the other.
  if (body.hasOwnProperty('dev_notes')) found.record.dev_notes = body.dev_notes || '';
  if (body.hasOwnProperty('dev_ask')) found.record.dev_ask = body.dev_ask || '';
  found.record.updated_at = new Date().toISOString();
  found.sheet.getRange(found.rowNum, 1, 1, HEADERS.length).setValues([recordToRow_(found.record)]);
  return { ok: true };
}

// Run from the editor to print a user's current invite/setup link (handy for
// the very first admin, or if an email did not arrive). Defaults to Edd.
function adminInviteLink(email) {
  var f = findUserByEmail_(email || 'ehewett@ardent-training.com');
  if (!f) { Logger.log('No such user: ' + email); return; }
  if (!f.user.invite_token) { Logger.log('No pending invite for ' + f.user.email + ' (already set up). Use Resend from the Users page.'); return; }
  var url = getAppUrl_();
  Logger.log('Setup link for ' + f.user.email + ': ' +
    (url || '(set APP_URL first)') + ((url && url.indexOf('?') > -1) ? '&' : '?') + 'invite=' + f.user.invite_token);
}

// The front-end sends POSTs with Content-Type text/plain so the browser
// treats them as "simple" requests and skips the CORS preflight.
function jsonOut(obj) {
  // Round 54: the single exit point for every response, so it is the one
  // reliable place to drop the cached issue list after a write. Doing it here
  // rather than inside each write function means a new action can never
  // forget - see READ_ONLY_ACTIONS.
  maybeInvalidate_();
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- Sheet helpers --------------------------------------------------------

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function sheetByName_(name) { return ss_().getSheetByName(name); }

// course_error -> Course Errors; anything else (incl tech_issue) decided here.
// Couriers write tracking numbers with spaces, dashes and mixed case, and the
// same consignment can arrive looking different in each email.
function normaliseTracking_(v) {
  var s = String(v || '');
  // A long all-digit reference (Evri, Yodel) comes back from the sheet as a
  // number in scientific notation, e.g. 1.234567890123456E+18, which would
  // strip down to nonsense. Pull it back to digits before we compare.
  if (/^[0-9.]+E\+?[0-9]+$/i.test(s)) { var n = Number(s); if (isFinite(n)) s = n.toFixed(0); }
  return s.toUpperCase().replace(/[^A-Z0-9]/g, '');
}
// A date we only ever care about to the day, read back safely. Sheets happily
// turns '2026-08-02' into a Date object at local midnight, which serialises an
// hour behind in summer, so slicing the ISO string lands a day early.
function dayStr_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v).slice(0, 10);
}
// The newest still-live shipping issue on this consignment, if any. A resolved
// one is left alone: a parcel that arrives, then goes missing on a re-send, is
// genuinely a new problem.
function findShippingByTracking_(track) {
  if (!track) return null;
  var sheet = sheetByName_(SHIPPING_SHEET);
  if (!sheet) return null;
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  var idx = {}; values[0].forEach(function (h, i) { idx[h] = i; });
  if (idx.tracking_number == null) return null;
  var best = null, bestAt = 0;
  for (var r = 1; r < values.length; r++) {
    if (!values[r][idx.issue_id]) continue;
    if (normaliseTracking_(values[r][idx.tracking_number]) !== track) continue;
    var st = String(values[r][idx.status] || '').toLowerCase();
    if (st === 'resolved' || st === 'past') continue;
    var at = new Date(values[r][idx.updated_at] || values[r][idx.submitted_at]).getTime() || 0;
    if (at >= bestAt) { bestAt = at; best = values[r][idx.issue_id]; }
  }
  return best;
}

function targetSheetName_(category) {
  var c = String(category).toLowerCase();
  if (c === 'shipping') return SHIPPING_SHEET;
  // FB-0254 (Edd, 20 Aug 2026): friction. Nothing is broken, but the design
  // cost somebody money or time - Ann paid for the Fast Track without using her
  // delivery code because the code box was not prominent enough, which cost a
  // support round trip and nearly a refund. That is a real finding and it is
  // not a fault, so it must not sit in a developer's fix queue. It rides in the
  // Tech Issues sheet kept apart by its category, exactly as internal work does,
  // so there is no fourth tab and no sheet migration.
  if (c === 'friction') return TECH_SHEET;
  // Internal work lives in the Tech Issues sheet, kept apart by its
  // audience value, so we don't need a fourth tab or a sheet migration.
  return (c === 'tech_issue' || c === 'internal') ? TECH_SHEET : COURSE_SHEET;
}

function getInstructorsSheet_() { return sheetByName_(INSTRUCTORS_SHEET); }

// Read-only mirror for the local Cowork sync. Gated by a standalone key held
// in the MIRROR_KEY script property, separate from user accounts and sessions,
// so if the key ever leaks it exposes read-only issue/feedback data and nothing
// more (no writes, no login, no account access). Rotate it by changing the
// property. Returns the most recent issues (newest first) plus all feedback.
// Optional params: key (required), limit (default 50), since (ISO date).
function mirror_(p) {
  var key = PropertiesService.getScriptProperties().getProperty('MIRROR_KEY');
  if (!key || !p || p.key !== key) return { ok: false, error: 'unauthorized' };

  var limit = parseInt(p.limit, 10);
  if (isNaN(limit) || limit < 1) limit = 50;

  var issues = getIssues_().issues || [];
  issues.sort(function (a, b) {
    var ax = new Date(a.updated_at || a.submitted_at || 0);
    var bx = new Date(b.updated_at || b.submitted_at || 0);
    return bx - ax; // newest first
  });
  var total = issues.length;

  if (p.since) {
    var since = new Date(p.since);
    if (!isNaN(since)) {
      issues = issues.filter(function (i) {
        return new Date(i.updated_at || i.submitted_at || 0) >= since;
      });
    }
  }

  var trimmed = issues.slice(0, limit);
  var feedback = getFeedback_().feedback || [];

  return {
    ok: true,
    generated_at: new Date().toISOString(),
    issue_total: total,
    issue_returned: trimmed.length,
    issues: trimmed,
    feedback: feedback
  };
}

function getIssues_() {
  var all = [];
  ISSUE_SHEETS.forEach(function (name) {
    var sheet = sheetByName_(name);
    if (!sheet) return;
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return;
    var head = values[0];
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if (!row[0]) continue; // skip blank rows
      var obj = {};
      for (var c = 0; c < head.length; c++) {
        obj[head[c]] = row[c] === '' ? null : row[c];
      }
      // Sheets reads these back as a Date and a number, which would reach the
      // app a day out and in scientific notation. Hand over what we stored.
      if (obj.chase_at) obj.chase_at = dayStr_(obj.chase_at);
      if (obj.tracking_number) obj.tracking_number = normaliseTracking_(obj.tracking_number);
      all.push(obj);
    }
  });
  return { ok: true, issues: all };
}

// ---- Round 54: list projection, one-shot bootstrap, and the caches --------
// Everything on the site felt slow (Edd, the oldest complaint in the project).
// Three causes, all addressed here:
//   1. getIssues shipped the whole sheet - raw_text and every report's full
//      transcript - about 2.6 MB, when the list views only ever draw badges,
//      summaries and dates from it. issueListRow_ builds a LIST projection:
//      empty cells are left out entirely (900 rows x 46 mostly-blank columns
//      was itself a third of the payload), and the long texts are clipped to
//      a preview so search still works. `clipped:1` marks a row whose full
//      text lives on the server; getIssue fetches it when a pane opens.
//   2. Opening the app made five or six separate 2-4 second round trips. On
//      Apps Script the round trip itself is ~1.8s before any work happens, so
//      the fix is to make one. bootstrap_ returns the lot.
//   3. Repeated reads re-read the whole spreadsheet. The projection is cached
//      in CacheService for a minute, chunked because a cache value caps at
//      100 KB. maybeInvalidate_ (called on the way out of jsonOut) drops it
//      after ANY write action, so nobody ever sees a stale board after their
//      own click. Reads are listed explicitly - an action we forget about
//      invalidates, which costs a re-read and never shows stale data.
// How much of a report the list carries. Measured on the live corpus: at 300
// characters two rows in three needed a follow-up fetch to be read, and that
// fetch is about five seconds. At 900 it is closer to one in five, and the
// long email threads that remain are the ones nobody expects to arrive
// instantly anyway. The extra weight is paid once, into a cache.
var LIST_RAW_CAP = 900;
var REPORTS_PASSTHROUGH = 1200;      // reports_json this small is not worth reworking
var ISSUE_CACHE_KEY = 'ait_issues_list_v1';
// Ten minutes. Any write drops the cache on its way out (maybeInvalidate_),
// so the only thing this window can hold back is a change typed straight into
// the spreadsheet by hand - and ten minutes is a fair wait for that.
var ISSUE_CACHE_SECONDS = 600;
var CACHE_CHUNK = 80 * 1024;         // CacheService caps a value at 100 KB
var CACHE_MAX_CHUNKS = 60;

function clip_(s, n) { s = String(s == null ? '' : s); return s.length > n ? s.slice(0, n) : s; }

// One issue, slimmed for a list view. Keeps every structured field the
// front end filters, sorts, badges or searches on; clips the two big texts.
function issueListRow_(i) {
  var o = {}, clipped = false, k, v;
  for (k in i) {
    if (!Object.prototype.hasOwnProperty.call(i, k)) continue;
    // next_action_json is a per-issue cache that only a detail pane ever asks
    // for. On 900 rows it would put the payload straight back where Round 54
    // found it, so it stays off the list and rides in on the nextAction call.
    if (k === 'raw_text' || k === 'reports_json' || k === 'next_action_json') continue;
    v = i[k];
    if (v === null || v === '' || v === undefined) continue;   // blanks cost bytes and say nothing
    o[k] = v;
  }
  if (i.raw_text != null && i.raw_text !== '') {
    var s = String(i.raw_text);
    o.raw_text = clip_(s, LIST_RAW_CAP);
    if (s.length > LIST_RAW_CAP) clipped = true;
  }
  if (i.reports_json) {
    var rjs = String(i.reports_json);
    // Most reports are short. Parsing and re-serialising 900 of them costs
    // more time on Apps Script than the bytes it saves, so anything already
    // small goes straight through untouched.
    if (rjs.length <= REPORTS_PASSTHROUGH) { o.reports_json = rjs; if (clipped) o.clipped = 1; return o; }
    var arr = null;
    try { arr = JSON.parse(rjs); } catch (e) { arr = null; }
    if (arr && arr.length !== undefined && arr.map) {
      o.reports_json = JSON.stringify(arr.map(function (rp) {
        var q = {}, rk, rv;
        for (rk in rp) {
          if (!Object.prototype.hasOwnProperty.call(rp, rk)) continue;
          if (rk === 'raw_text') continue;
          rv = rp[rk];
          if (rv === null || rv === '' || rv === undefined) continue;
          q[rk] = rv;
        }
        if (rp.raw_text != null && rp.raw_text !== '') {
          var t = String(rp.raw_text);
          q.raw_text = clip_(t, LIST_RAW_CAP);
          if (t.length > LIST_RAW_CAP) clipped = true;
        }
        return q;
      }));
    } else {
      o.reports_json = i.reports_json;
    }
  }
  if (clipped) o.clipped = 1;
  return o;
}

// The payload is well over a megabyte and a cache value caps at 100 KB, so
// it is gzipped first (JSON of this shape squashes to about a seventh) and
// then chunked. Writing twenty 80 KB chunks took NINETEEN SECONDS and then
// read back as a miss - measured, r54, before this. Gzipped it is three or
// four chunks and the write is free.
function packForCache_(str) {
  return Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(str, 'application/json')).getBytes());
}
function unpackFromCache_(b64) {
  return Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip')).getDataAsString();
}
function cachePutChunked_(key, raw, secs) {
  try {
    var str = packForCache_(raw);
    var c = CacheService.getScriptCache();
    var n = Math.ceil(str.length / CACHE_CHUNK);
    if (n > CACHE_MAX_CHUNKS) { c.remove(key + '_n'); return false; }
    var map = {};
    for (var i = 0; i < n; i++) map[key + '_' + i] = str.substring(i * CACHE_CHUNK, (i + 1) * CACHE_CHUNK);
    c.putAll(map, secs);
    // Written last, so a half-written set never reads as complete. Reading
    // one chunk straight back proves the write actually took: silently
    // dropped values are exactly how the first attempt at this failed.
    if (c.get(key + '_0') == null) return false;
    c.put(key + '_n', String(n), secs);
    return true;
  } catch (e) { return false; }
}

function cacheGetChunked_(key) {
  try {
    var c = CacheService.getScriptCache();
    var n = Number(c.get(key + '_n') || 0);
    if (!n) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(key + '_' + i);
    var got = c.getAll(keys) || {};
    var out = '';
    for (var j = 0; j < n; j++) {
      var part = got[key + '_' + j];
      if (part == null) return null;   // a chunk expired: treat the whole thing as a miss
      out += part;
    }
    return unpackFromCache_(out);
  } catch (e) { return null; }
}

function invalidateIssueCache_() {
  try {
    var c = CacheService.getScriptCache();
    var keys = [ISSUE_CACHE_KEY + '_n'];
    for (var i = 0; i < CACHE_MAX_CHUNKS; i++) keys.push(ISSUE_CACHE_KEY + '_' + i);
    c.removeAll(keys);
  } catch (e) {}
}

// Actions that only read. Everything else drops the issue cache on its way
// out. Being on this list is the ONLY way to keep the cache through a call,
// so a new write action is safe by default.
var READ_ONLY_ACTIONS = {
  ping: 1, me: 1, bootstrap: 1, getIssues: 1, getIssuesList: 1, getIssue: 1, getInstructors: 1,
  listUsers: 1, getPlaybook: 1, listPlaybookSuggestions: 1, listKnownFixFlags: 1, getFeedback: 1, getAssignees: 1,
  getInvite: 1, mirror: 1, chatwootList: 1, chatScanList: 1, chatwootContactUrl: 1, chatBackSweepState: 1, lessonIssueCounts: 1,
  listVoiceGuides: 1, listContentSuggestions: 1, getManifest: 1, extract: 1, askIssues: 1,
  suggestFix: 1, troubleshoot: 1, matchUpdate: 1, draftStudentMessage: 1, listLiveCases: 1,
  // Reads open issues and answers a question; writes nothing, so the cached
  // list projection survives it (the Round 54 invalidation rule).
  sameIssue: 1,
  caseDraftReply: 1, batchStudentDrafts: 1, chatwootImport: 1, login: 1, logout: 1,
  // nextAction DOES write one cell (its own cached answer), and it still
  // belongs here. The list projection leaves next_action_json out entirely, so
  // there is no way for a held cache to show a stale next action - and dropping
  // the whole board cache every time somebody opened an issue would undo most
  // of Round 54. If next_action_json ever reaches the list, take this back out.
  nextAction: 1
};
var CURRENT_ACTION_ = '';
function maybeInvalidate_() {
  if (!CURRENT_ACTION_) return;
  if (READ_ONLY_ACTIONS[CURRENT_ACTION_]) return;
  invalidateIssueCache_();
}

// The list payload: cached projection, or built and cached.
function getIssuesList_() {
  var cached = cacheGetChunked_(ISSUE_CACHE_KEY);
  if (cached) {
    try {
      var p = JSON.parse(cached);
      p.from_cache = true;
      return p;
    } catch (e) {}
  }
  // Straight off the sheet into the projection: building 900 full objects
  // and then 900 slim ones was double the work for no reason.
  var out = [];
  ISSUE_SHEETS.forEach(function (name) {
    var sheet = sheetByName_(name);
    if (!sheet) return;
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return;
    var head = values[0];
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if (!row[0]) continue;
      var obj = {};
      for (var c = 0; c < head.length; c++) {
        if (row[c] === '' || row[c] === null) continue;
        obj[head[c]] = row[c];
      }
      if (obj.chase_at) obj.chase_at = dayStr_(obj.chase_at);
      if (obj.tracking_number) obj.tracking_number = normaliseTracking_(obj.tracking_number);
      out.push(issueListRow_(obj));
    }
  });
  var payload = { ok: true, generated_at: new Date().toISOString(), issues: out };
  cachePutChunked_(ISSUE_CACHE_KEY, JSON.stringify(payload), ISSUE_CACHE_SECONDS);
  return payload;
}

// One issue, full fat. Reads the id column first and then just the one row,
// so opening a detail pane never costs a whole-spreadsheet read.
function getIssueFull_(data) {
  var id = String((data && (data.issue_id || data.id)) || '');
  if (!id) return { ok: false, error: 'need an issue_id' };
  // Every call to the Sheets service costs, and this runs the moment a pane
  // opens, so it is kept to two: find the row, read the row. The column names
  // come from HEADERS rather than a third read of row 1 - the sheet order IS
  // HEADERS, which is the rule the whole file already relies on.
  for (var s = 0; s < ISSUE_SHEETS.length; s++) {
    var sheet = sheetByName_(ISSUE_SHEETS[s]);
    if (!sheet) continue;
    var rowNum = 0;
    try {
      var hit = sheet.createTextFinder(id).matchEntireCell(true).findNext();
      if (hit && hit.getColumn() === 1) rowNum = hit.getRow();
    } catch (e) { rowNum = 0; }
    if (!rowNum) {
      var last = sheet.getLastRow();
      if (last < 2) continue;
      var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
      for (var r = 0; r < ids.length; r++) {
        if (String(ids[r][0]) === id) { rowNum = r + 2; break; }
      }
    }
    if (!rowNum) continue;
    var row = sheet.getRange(rowNum, 1, 1, HEADERS.length).getValues()[0];
    var obj = {};
    for (var c = 0; c < HEADERS.length; c++) obj[HEADERS[c]] = row[c] === '' ? null : row[c];
    if (obj.chase_at) obj.chase_at = dayStr_(obj.chase_at);
    if (obj.tracking_number) obj.tracking_number = normaliseTracking_(obj.tracking_number);
    return { ok: true, issue: obj };
  }
  return { ok: false, error: 'not found' };
}

// Everything the app needs to draw itself, in one round trip. Each piece is
// permission-gated the same way its own action is, and each is wrapped so one
// slow or broken corner (Chatwoot, say) cannot take the whole load down with
// it - the app still opens, just without that panel.
function bootstrap_(user) {
  var out = {
    ok: true,
    generated_at: new Date().toISOString(),
    user: publicUser_(user),
    backend: backendInfo_()
  };
  // Rough server-side timings, so a slow open can be diagnosed from the
  // response instead of guessing which piece of it was the hold-up.
  var t0 = Date.now(), ms = {};
  var list = getIssuesList_();
  out.issues = list.issues || [];
  out.issues_from_cache = !!list.from_cache;
  ms.issues = Date.now() - t0;
  if (hasPerm_(user, reqPerm_('getInstructors'))) {
    try { out.instructors = getInstructors_().instructors || []; } catch (e) { out.instructors_error = String(e); }
  }
  ms.instructors = Date.now() - t0;
  if (hasPerm_(user, reqPerm_('getAssignees'))) {
    try { out.assignees = listAssignees_().assignees || []; } catch (e) { out.assignees_error = String(e); }
  }
  ms.assignees = Date.now() - t0;
  if (hasPerm_(user, reqPerm_('listLiveCases'))) {
    try { out.live_cases = listLiveCases_({ _issues: out.issues }).cases || []; } catch (e) { out.live_cases_error = String(e); }
  }
  ms.live_cases = Date.now() - t0;
  if (hasPerm_(user, reqPerm_('chatScanList'))) {
    try { out.scans = chatScanList_().scans || []; } catch (e) { out.scans_error = String(e); }
  }
  ms.scans = Date.now() - t0;
  if (hasPerm_(user, reqPerm_('listPlaybookSuggestions'))) {
    try { out.playbook_suggestions = getSuggestions_() || []; } catch (e) {}
  }
  if (hasPerm_(user, reqPerm_('listKnownFixFlags'))) {
    try { out.knownfix_corrections = getKfCorrections_() || []; } catch (e) {}
  }
  ms.total = Date.now() - t0;
  out.ms = ms;   // cumulative milliseconds at each step
  return out;
}

function getInstructors_() {
  var sheet = getInstructorsSheet_();
  var values = sheet.getDataRange().getValues();
  var list = [];
  for (var r = 1; r < values.length; r++) {
    if (!values[r][0]) continue;
    list.push({ name: values[r][0], email: values[r][1] || null });
  }
  return { ok: true, instructors: list };
}

// Build a full row array in HEADERS order from an issue object.
function recordToRow_(issue) {
  return HEADERS.map(function (key) {
    var v = issue[key];
    return v == null ? '' : v;
  });
}

function addIssue_(data) {
  // The instructor is always the logged-in user, so it cannot be spoofed.
  if (data._user && data._user.name) data.instructor_name = data._user.name;
  // The account behind that name, for the same reason (Round 61). The name is
  // what we show; the email is what we can still join on in a year.
  if (data._user && data._user.email) data.instructor_email = data._user.email;
  var now = new Date().toISOString();
  var category = (data.category || 'course_error').toLowerCase();
  // Legacy callers (and old saved drafts) may still send category 'internal'.
  // That is now an audience, not a category, and it always meant tech.
  var audience = data.audience === 'internal' || category === 'internal' ? 'internal' : 'student';
  if (category === 'internal') category = 'tech_issue';

  // One entry describing this particular report (who hit it, who logged it).
  // We keep its own priority and raw text too, so a wrongly merged report can
  // be split back out cleanly later.
  var report = {
    kind: 'report',
    student_name: data.student_name || '',
    student_contact: data.student_contact || '',
    device_info: data.device_info || '',
    instructor_name: data.instructor_name || '',
    instructor_email: data.instructor_email || '',
    summary: data.summary || '',
    priority: (data.priority || '').toLowerCase(),
    raw_text: data.raw_text || '',
    recommended_steps: data.recommended_steps || null,
    date: now
  };

  // Let the AI decide if this is really the same issue as one already open, and
  // if so roll it into that row instead of creating a duplicate. Tech issues
  // match on the same underlying bug across any lesson. Course errors match only
  // within the same slide (lesson code), and only when it is the same error.
  var thisKind = data.request_kind === 'improvement' ? 'improvement' : 'fix';

  // SHIPPING: the tracking number is the merge key. DHL starts a fresh email
  // thread for the same consignment, so without this one parcel problem
  // becomes four issues (Edd, 30 Jul). Any open shipping issue with the same
  // tracking number wins outright, before the AI matcher gets a look in.
  if (category === 'shipping') {
    var track = normaliseTracking_(data.tracking_number);
    if (track && !data.no_merge) {
      var hit = findShippingByTracking_(track);
      if (hit) {
        var mergedShip = addReportToIssue_(hit, data, report);
        if (mergedShip && mergedShip.ok) return mergedShip;
      }
    }
  }

  // A conversation that was already sorted out in the chat is logged for the
  // record. Don't roll it into an open issue, and don't route it to a queue.
  // The instructor can also decide this on the form (Round 16): merge_into is
  // an explicit "same fault, add my student as another report", and no_merge
  // is an explicit "this is different" that overrides the AI matcher too.
  var matchId = data.resolved ? null
    : data.no_merge ? null
    : (data.merge_into || aiMatchIssue_(data, category));
  if (matchId) {
    // Never roll a fix into an improvement (or vice versa); they are different
    // things even on the same lesson, so keep them as separate entries.
    var matchRow = findRow_(matchId);
    var matchKind = matchRow ? (String(matchRow.record.request_kind || 'fix').toLowerCase()) : 'fix';
    if (matchKind === thisKind) {
      var merged = addReportToIssue_(matchId, data, report);
      if (merged && merged.ok) return merged;
      // if the merge failed for any reason, fall through and log as new
    }
  }

  var issue = {
    issue_id: Utilities.getUuid(),
    submitted_at: now,
    updated_at: now,
    instructor_name: data.instructor_name || '',
    instructor_email: data.instructor_email || '',
    category: category,
    raw_text: data.raw_text || '',
    student_name: data.student_name || '',
    student_contact: data.student_contact || '',
    chatwoot_contact_id: String(data.chatwoot_contact_id || '').trim(),
    device_info: data.device_info || '',
    course: data.course || '',
    module: data.module || '',
    lesson: data.lesson || '',
    lesson_code: data.lesson_code || '',
    issue_type: data.issue_type || '',
    summary: data.summary || '',
    priority: (function () {
      var p = String(data.priority || '').toLowerCase();
      if (data.request_kind !== 'improvement' && String(category).toLowerCase() !== 'friction') return p;
      // FB-0253 (Edd): "If something is more of a feature request then a real
      // bug, we probably want it logged as low priority." An improvement is
      // backlog by definition - nothing is broken, so nobody is blocked. The
      // extraction reads urgency off the words, and an enthusiastic "we really
      // need this" comes back high. Anyone can raise it by hand afterwards;
      // this only decides what it ARRIVES as. Note the report's own priority a
      // few lines up is left alone, because that is a record of how the report
      // read at the time.
      return (p === 'high' || p === 'medium' || !p) ? 'low' : p;
    })(),
    priority_reason: data.priority_reason || '',
    image_urls: normaliseImageUrls_(data.image_urls),
    // If the instructor has already given the student the suggested fix, this
    // lands as "Resolved - TBC" with that fix saved, and will auto-resolve
    // after a quiet spell unless it comes back.
    status: data.resolved ? 'resolved' : (data.parked ? 'parked' : (data.tbc ? 'resolved_tbc' : (data.status || 'open'))),
    resolved_at: data.resolved ? (data.resolved_at || now) : '',
    // Parked keeps its reason here too: a parked issue with no note is just an
    // open one nobody looks at.
    resolution_note: (data.resolved || data.tbc || data.parked) ? (data.resolution_note || '') : '',
    // Resolved in the pasted chat itself: the student was part of that
    // conversation, so there's nobody left to notify (keeps it out of the
    // instructor's Actions list).
    notified_students: data.resolved ? true : false,
    report_count: 1,
    reports_json: JSON.stringify([report]),
    dev_passed_at: '',
    dev_fixed_at: '',
    dev_notes: '',
    checklist_json: data.checklist_json || '',
    request_kind: data.request_kind === 'improvement' ? 'improvement' : 'fix',
    assignee: data.assignee || '',
    media_kind: data.media_kind || '',
    double_checked: data.double_checked === true || data.double_checked === 'true' ? true : '',
    impact: data.impact || '',
    section: data.section || '',
    dev_query: '', dev_query_at: '', dev_query_by: '', dev_query_target: '',
    platform: ['browser', 'app', 'both'].indexOf(String(data.platform || '')) > -1 ? data.platform : '',
    recheck_at: '',
    audience: audience,
    courier: category === 'shipping' ? (data.courier || '') : '',
    tracking_number: category === 'shipping' ? normaliseTracking_(data.tracking_number) : '',
    student_sorted: (data.student_sorted === true || data.student_sorted === 'true') ? true : '',
    // Whether anyone is on the other end. The form answers this; we only fall
    // back to reading it off the report when an older caller says nothing.
    student_involved: data.student_involved === 'no' || data.student_involved === false ? 'no'
      : data.student_involved === 'yes' || data.student_involved === true ? 'yes'
      : (audience === 'internal' ? 'no'
        : (String(data.student_name || '').trim() || String(data.student_contact || '').trim()) ? 'yes' : 'no'),
    // Default to chasing in three working-ish days if nobody said otherwise:
    // an unchased parcel problem is the one that goes quiet for a fortnight.
    chase_at: category === 'shipping'
      ? (data.chase_at || new Date(Date.now() + 3 * 864e5).toISOString().slice(0, 10))
      : ''
  };

  // Auto-route to the right fix queue (unless the instructor already settled it
  // with a suggested fix). Every course error needs a content fix, so it always
  // goes to the course team. A tech issue goes to the developers when it is high
  // priority (something is badly enough broken to warrant it) OR when the AI
  // judges it a genuine code bug rather than a user-side step.
  // Improvements (feature/enhancement requests) are a calmer backlog: they do
  // NOT auto-route to a fix team, they just sit open for the team to review.
  // Shipping never routes to a fix team: nobody here fixes a parcel, we chase
  // the courier. It stays open with a chase date instead.
  // Scan-logged TECH issues don't auto-route (Edd, FB-0184/0185): nobody has
  // tried anything with the student yet, so "high" is not enough to skip the
  // troubleshooting stage - they stay open in the Actions Scanned lane until
  // a human has looked in. A known-account quirk (the partner portal URL, say)
  // would otherwise land on a developer who can't fix a login habit. Course
  // errors still route: there is no student troubleshooting for a wrong diagram.
  var scanLogged = String(data.instructor_name || '') === 'Overnight scan';
  if (!data.tbc && !data.resolved && !data.parked && issue.request_kind !== 'improvement' && category !== 'shipping' && category !== 'friction') {
    if (category === 'course_error') {
      issue.dev_passed_at = new Date().toISOString();
      issue.status = 'with_dev';
    } else if (audience === 'internal') {
      // Internal work is logged deliberately by an admin, so no AI judgement
      // needed: genuine defects and infrastructure work go straight to the dev
      // queue; content, admin, and feature items sit open for triage.
      if (issue.issue_type === 'bug' || issue.issue_type === 'infrastructure') {
        issue.dev_passed_at = new Date().toISOString();
        issue.status = 'with_dev';
      }
    } else if (category === 'tech_issue' && !scanLogged && String(issue.priority).toLowerCase() === 'high') {
      // Edd's rule (21 Jul): tech issues reach the developers automatically
      // only when HIGH priority, or when a repeat report makes it 3+ reports
      // (see addReportToIssue_). The old AI "needs a developer" judgement was
      // routing single medium reports, so it no longer routes on its own.
      issue.dev_passed_at = new Date().toISOString();
      issue.status = 'with_dev';
    }
  }

  var sheet = sheetByName_(targetSheetName_(category));
  sheet.appendRow(recordToRow_(issue));

  // Slack only for a high-priority fix; never let a Slack failure block the save.
  // Improvements never fire an alert, they are backlog, not something to jump on.
  if (String(issue.priority).toLowerCase() === 'high' && issue.request_kind !== 'improvement' &&
      issue.status !== 'resolved' && issue.status !== 'resolved_tbc') {
    try { sendSlack_(issue, data.app_url || getAppUrl_()); } catch (slackErr) {}
  }

  // Closed on a workaround: see whether this is the third one this week going
  // the same way, in which case it isn't a workaround any more, it's a fault.
  if (String(issue.status).toLowerCase() === 'resolved_tbc') {
    try { checkSharedWorkaround_(issue, data.app_url || getAppUrl_()); } catch (e) {}
  }

  // Imported from a live chat: leave an internal note on that conversation so
  // the two systems stay joined up.
  if (data.chatwoot_conversation_id) {
    try { chatwootNote_(data.chatwoot_conversation_id, issue, data.app_url || getAppUrl_()); } catch (e) {}
  }

  return { ok: true, issue: issue, merged: false };
}

// Roll a new report into an existing issue row: add its student to the list,
// bump the priority a level (capped at high), bump the count, and keep an
// audit note in raw_text. Returns the same shape as addIssue_.
function addReportToIssue_(id, data, report) {
  var found = findRow_(id);
  if (!found) return { ok: false, error: 'matched issue not found: ' + id };
  var rec = found.record;
  // What the priority was before this report touched it. Slack is for an issue
  // BECOMING high (Edd, 18 Aug 2026), so an issue already sitting at high must
  // not ping again every time another report lands on it.
  var priorityBefore = String(rec.priority || '').toLowerCase();

  var reports = [];
  try { reports = rec.reports_json ? JSON.parse(rec.reports_json) : []; } catch (e) { reports = []; }
  if (!reports.length) {
    // Older row with no reports list yet: seed it from the row's own student.
    reports.push({
      student_name: rec.student_name || '', student_contact: rec.student_contact || '',
      device_info: rec.device_info || '', instructor_name: rec.instructor_name || '',
      summary: rec.summary || '', priority: String(rec.priority || '').toLowerCase(),
      raw_text: rec.raw_text || '', date: rec.submitted_at || ''
    });
  }
  reports.push(report);
  rec.reports_json = JSON.stringify(reports);
  rec.report_count = reports.length;

  // Each extra report nudges the priority up a level, never below what this
  // report was logged as.
  rec.priority = bumpPriority_(rec.priority, data.priority);

  // TBC handling on a repeat report:
  //  - if the instructor has applied the suggested fix, mark it Resolved - TBC
  //    (and reset the silence timer via updated_at below).
  //  - if it was already Resolved - TBC and someone is reporting it again
  //    without applying a fix, that means it has NOT worked, so reopen it.
  if (data.tbc) {
    rec.status = 'resolved_tbc';
    if (data.resolution_note) rec.resolution_note = data.resolution_note;
  } else if (String(rec.status).toLowerCase() === 'resolved_tbc') {
    rec.status = 'open';
  }

  // A parked issue was left alone "unless more people report it". Someone just
  // did, so wake it up (Edd, 26 Jul).
  if (String(rec.status).toLowerCase() === 'parked') {
    rec.status = 'open';
    rec.raw_text = (rec.raw_text || '') + '\n\n--- unparked: reported again ---';
  }

  // Note: a "Submit and park" that turns out to MERGE into an existing issue
  // deliberately does not park that issue. Park means "this student went quiet
  // so I can't get to the bottom of it", and that is no reason to stop work on
  // a fault other people are still hitting. The report still joins the trail.

  // Repeat reports can tip a tech issue over the routing line (3+ reports, or
  // the priority bump above making it high): hand it to the developers.
  if (String(rec.category).toLowerCase() === 'tech_issue' &&
      String(rec.status).toLowerCase() === 'open' &&
      (reports.length >= 3 || String(rec.priority).toLowerCase() === 'high')) {
    rec.status = 'with_dev';
    if (!rec.dev_passed_at) rec.dev_passed_at = new Date().toISOString();
  }

  var stamp = new Date().toISOString().slice(0, 10);
  rec.raw_text = (rec.raw_text || '') + '\n\n--- also reported ' + stamp + ' by ' +
    (report.instructor_name || 'someone') +
    (report.student_name ? ' (student: ' + report.student_name + ')' : '') + ' ---\n' +
    (data.raw_text || data.summary || '');

  if (data.image_urls) {
    var existing = rec.image_urls ? String(rec.image_urls).split(',') : [];
    var more = normaliseImageUrls_(data.image_urls).split(',').filter(Boolean);
    rec.image_urls = existing.concat(more).filter(Boolean).join(',');
  }

  rec.updated_at = new Date().toISOString();
  found.sheet.getRange(found.rowNum, 1, 1, HEADERS.length).setValues([recordToRow_(rec)]);

  // If the bump just pushed it to high, let Slack know once, but never ping for
  // something that has landed resolved or Resolved - TBC, and never for one that
  // was already high before this report arrived.
  if (String(rec.priority).toLowerCase() === 'high' && priorityBefore !== 'high' &&
      rec.status !== 'resolved' && rec.status !== 'resolved_tbc') {
    try { sendSlack_(rec, data.app_url || getAppUrl_()); } catch (e) {}
  }

  if (String(rec.status).toLowerCase() === 'resolved_tbc') {
    try { checkSharedWorkaround_(rec, data.app_url || getAppUrl_()); } catch (e) {}
  }

  return { ok: true, issue: rec, merged: true, report_count: rec.report_count };
}

// Raise priority one level toward high, but never below the incoming report's
// own priority.
function bumpPriority_(current, incoming) {
  var order = ['low', 'medium', 'high'];
  var ci = order.indexOf(String(current || 'low').toLowerCase()); if (ci < 0) ci = 0;
  var ii = order.indexOf(String(incoming || '').toLowerCase());
  var base = Math.max(ci, ii < 0 ? ci : ii);
  return order[Math.min(base + 1, 2)];
}

// Ask the AI whether a new report is really the same issue as one already open.
// Returns the matching issue_id, or null. Conservative: no key, no candidates,
// or any doubt and it returns null so we log a fresh row.
//
// Tech issues are matched on the same underlying bug across any lesson.
// Course errors are first narrowed to the SAME slide (lesson code), then the AI
// must confirm it is the same actual error, because one slide can carry two
// completely different errors that must stay separate.
// The same matcher the submit already runs, asked BEFORE the submit so the
// instructor sees the merge coming and gets the choice (Edd, FB-0204: he
// logged the same city drop-down fault twice, five minutes apart, and was
// never offered it). Deliberately the same function rather than a second
// opinion, so what the popup offers is exactly what would otherwise have
// happened quietly at submit. The front end only asks when its own word
// scoring already puts a candidate close, so this is one small call on the
// reports where it might matter, not on every extraction.
function sameIssue_(data) {
  var category = String(data.category || 'tech_issue').toLowerCase();
  if (category === 'internal') category = 'tech_issue';
  if (!String(data.summary || '').trim() && !String(data.raw_text || '').trim()) {
    return { ok: true, match_id: null, why: 'nothing to match on yet' };
  }
  var id = null;
  try { id = aiMatchIssue_(data, category); } catch (e) {
    return { ok: false, error: String(e) };
  }
  return { ok: true, match_id: id || null };
}

function aiMatchIssue_(data, category) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return null;

  var isCourse = String(category).toLowerCase() === 'course_error';

  // A course error is only ever "the same" when it is on the same slide, so with
  // no lesson code there is nothing safe to narrow on.
  if (isCourse && !data.lesson_code) return null;

  var sheet = sheetByName_(isCourse ? COURSE_SHEET : TECH_SHEET);
  if (!sheet) return null;
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  var head = values[0];
  var idx = {}; head.forEach(function (h, i) { idx[h] = i; });

  var wantLesson = String(data.lesson_code || '').trim().toLowerCase();
  var candidates = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!row[idx['issue_id']]) continue;
    // Never merge a new report into a resolved issue, or into a "past" row
    // imported from the old spreadsheets — that would resurrect ancient history.
    var candStatus = String(row[idx['status']]).toLowerCase();
    if (candStatus === 'resolved' || candStatus === 'past') continue;
    // Tech issues and internal tasks share a sheet; never match across the two.
    if (idx['category'] != null && String(row[idx['category']] || '').toLowerCase() !== String(category).toLowerCase()) continue;
    // Course errors: only consider other errors on the same slide.
    if (isCourse) {
      var lc = String(row[idx['lesson_code']] || '').trim().toLowerCase();
      if (lc !== wantLesson) continue;
    }
    candidates.push({
      id: row[idx['issue_id']],
      summary: row[idx['summary']],
      lesson_code: row[idx['lesson_code']]
    });
  }
  if (!candidates.length) return null;

  var instruction = isCourse
    ? 'These are content errors in an online sailing course, and they are all on the SAME slide. ' +
      'Decide whether the NEW error is the SAME actual error as one of the EXISTING open errors ' +
      '(the same wrong fact, the same typo, the same mislabelled diagram, the same wrong quiz answer), even if worded differently. ' +
      'Being on the same slide is NOT enough on its own: one slide can have two completely different errors, which must stay separate. ' +
      'Only match when it is clearly the same specific error. If in doubt, do not match.'
    : 'These are technical support issues for an online sailing course platform. ' +
      'Decide whether the NEW report describes the SAME underlying bug as one of the EXISTING open issues ' +
      '(the same broken button, page, video, quiz, login, or behaviour), even if it is a different student or worded differently. ' +
      'To count as the same it must be BOTH the same part of the platform (the same page, feature, or flow) AND the same symptom. ' +
      'A different page or a different feature is NEVER the same issue, even if the symptoms sound similar. ' +
      'Be conservative: only match when it is clearly the same problem. If in doubt, do not match.';

  var prompt = instruction + '\n\n' +
    'NEW report:\n' + JSON.stringify({ summary: data.summary || '', raw_text: data.raw_text || '', lesson_code: data.lesson_code || '', device: data.device_info || '' }) + '\n\n' +
    'EXISTING open issues:\n' + JSON.stringify(candidates) + '\n\n' +
    'Return ONLY JSON: {"match_id": "<id of the matching existing issue, or null>"}. No prose, no markdown fences.';

  var res;
  try {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 120, messages: [{ role: 'user', content: prompt }] })
    });
  } catch (e) { return null; }
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) return null;

  var parsed; try { parsed = JSON.parse(res.getContentText()); } catch (e) { return null; }
  var text = '';
  if (parsed.content) for (var i = 0; i < parsed.content.length; i++) {
    if (parsed.content[i].type === 'text') text += parsed.content[i].text;
  }
  text = text.trim().replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
  var out; try { out = JSON.parse(text); } catch (e) { return null; }
  var mid = out && out.match_id;
  if (!mid || mid === 'null') return null;
  for (var c = 0; c < candidates.length; c++) if (candidates[c].id === mid) return mid;
  return null;
}

// Find a row by issue_id across both issue sheets.
function findRow_(id) {
  for (var s = 0; s < ISSUE_SHEETS.length; s++) {
    var sheet = sheetByName_(ISSUE_SHEETS[s]);
    if (!sheet) continue;
    var values = sheet.getDataRange().getValues();
    for (var r = 1; r < values.length; r++) {
      if (values[r][0] === id) {
        var head = values[0];
        var obj = {};
        for (var c = 0; c < head.length; c++) obj[head[c]] = values[r][c];
        return { sheetName: ISSUE_SHEETS[s], sheet: sheet, rowNum: r + 1, record: obj };
      }
    }
  }
  return null;
}

function updateIssue_(data) {
  var id = data.issue_id;
  if (!id) return { ok: false, error: 'updateIssue needs an issue_id' };

  var found = findRow_(id);
  if (!found) return { ok: false, error: 'No issue found with id ' + id };

  var record = found.record;
  var wasResolved = String(record.status || '').toLowerCase() === 'resolved';

  // Without the manage permission (dev / course team), only the priority and
  // the estimated fix size may be tweaked - the two edits their drawer offers
  // (Edd, 21 Jul; size added FB-0164).
  if (data._user && !hasPerm_(data._user, 'manage')) {
    var allowedKeys = { priority: 1, priority_reason: 1, fix_size: 1 };
    var blocked = HEADERS.filter(function (k) {
      return data.hasOwnProperty(k) && !allowedKeys[k] && k !== 'issue_id';
    });
    if (blocked.length) return { ok: false, error: 'Only priority and fix size can be edited from this queue.' };
  }

  // Moving something INTO the dev queue by hand (including a kanban drag) is
  // the same admin call as passToDev, so hold it to the same permission.
  if (String(data.status || '').toLowerCase() === 'with_dev' &&
      String(record.status || '').toLowerCase() !== 'with_dev' &&
      data._user && !hasPerm_(data._user, 'users')) {
    return { ok: false, error: 'Only an admin can pass an issue to the developers.' };
  }

  // Overlay any provided fields (except identity fields).
  HEADERS.forEach(function (key) {
    if (key === 'issue_id' || key === 'submitted_at') return;
    if (data.hasOwnProperty(key)) {
      record[key] = key === 'image_urls' ? normaliseImageUrls_(data[key]) : data[key];
    }
  });

  // Stamp resolved_at the first time it is resolved.
  if (String(record.status).toLowerCase() === 'resolved' && !record.resolved_at) {
    record.resolved_at = new Date().toISOString();
  }
  // Keep the developer timestamps in step if the status is changed by hand.
  if (String(record.status).toLowerCase() === 'with_dev' && !record.dev_passed_at) {
    record.dev_passed_at = new Date().toISOString();
  }
  if (String(record.status).toLowerCase() === 'dev_fixed' && !record.dev_fixed_at) {
    record.dev_fixed_at = new Date().toISOString();
  }
  record.updated_at = new Date().toISOString();

  var targetName = targetSheetName_(record.category);
  var row = recordToRow_(record);

  if (targetName === found.sheetName) {
    // Same tab: rewrite the row in place.
    found.sheet.getRange(found.rowNum, 1, 1, HEADERS.length).setValues([row]);
  } else {
    // Category changed: move the row to the other tab.
    found.sheet.deleteRow(found.rowNum);
    sheetByName_(targetName).appendRow(row);
  }

  // When a tech issue first reaches resolved with a note, let the AI consider
  // proposing a playbook addition (queued for admin approval).
  if (!wasResolved && String(record.status).toLowerCase() === 'resolved' &&
      record.resolution_note && String(record.category).toLowerCase() === 'tech_issue') {
    try { proposePlaybookUpdate_(record); } catch (e) {}
  }

  return { ok: true, issue_id: id, moved: targetName !== found.sheetName, sheet: targetName };
}

function normaliseImageUrls_(value) {
  if (!value) return '';
  if (Array.isArray(value)) return value.join(',');
  return String(value);
}

// Save just the troubleshooting checklist state on an issue, without touching
// anything else. Lets instructors and the dev/course teams tick items off
// wherever the checklist shows. checklist_json is a JSON string mapping each
// item id to "done" | "na" | "todo".
function saveChecklist_(data) {
  var id = data.issue_id;
  if (!id) return { ok: false, error: 'saveChecklist needs an issue_id' };
  var found = findRow_(id);
  if (!found) return { ok: false, error: 'No issue found with id ' + id };
  var value = data.checklist_json;
  if (value && typeof value !== 'string') value = JSON.stringify(value);
  setCellOnIssue_(found, 'checklist_json', value || '');
  setCellOnIssue_(found, 'updated_at', new Date().toISOString());
  return { ok: true, issue_id: id };
}

// Write a single named column on an issue row found via findRow_().
function setCellOnIssue_(found, key, value) {
  var col = HEADERS.indexOf(key);
  if (col < 0) return;
  found.sheet.getRange(found.rowNum, col + 1).setValue(value);
}

// Assign (or unassign) an issue to a person who will fix it. assignee is a name
// string; pass an empty string to clear it.
function assignIssue_(data) {
  var id = data.issue_id;
  if (!id) return { ok: false, error: 'assignIssue needs an issue_id' };
  var found = findRow_(id);
  if (!found) return { ok: false, error: 'No issue found with id ' + id };
  setCellOnIssue_(found, 'assignee', data.assignee || '');
  setCellOnIssue_(found, 'updated_at', new Date().toISOString());
  return { ok: true, issue_id: id, assignee: data.assignee || '' };
}

// Assign a whole set of issues to one person in a single call (Edd, FB-0165:
// "assign all knowledge checks to one person"). The front end works out which
// issues the chosen kind covers - knowledge checks and module assessments live
// in the lesson-code suffix, which only the front end knows how to read - and
// sends the ids; this end just walks the list. One round trip instead of one
// per issue.
function bulkAssign_(data) {
  var ids = data.issue_ids;
  if (!ids || !ids.length) return { ok: false, error: 'bulkAssign needs a list of issue_ids' };
  if (ids.length > 200) return { ok: false, error: 'That is over 200 issues in one go - narrow it down first.' };
  var assignee = String(data.assignee || '');
  var now = new Date().toISOString();
  var done = 0, missed = [];
  ids.forEach(function (id) {
    var found = findRow_(id);
    if (!found) { missed.push(id); return; }
    setCellOnIssue_(found, 'assignee', assignee);
    setCellOnIssue_(found, 'updated_at', now);
    done++;
  });
  return { ok: true, assigned: done, missed: missed, assignee: assignee };
}

// An AI pass over the open course-fix queue (Edd, FB-0165): does the stored
// priority still look right? Returns suggestions only - nothing is applied
// here. The admin gets a per-suggestion Apply button on the front end, because
// a priority change is a judgement call and the AI only gets a vote, not a say.
function courseReview_(data) {
  var pool = getIssues_().issues.filter(function (i) {
    if (String(i.category || '').toLowerCase() !== 'course_error') return false;
    var s = String(i.status || '').toLowerCase();
    return s === 'open' || s === 'with_dev';
  });
  if (!pool.length) return { ok: true, reviewed: 0, suggestions: [] };

  // Priority first so a capped run still covers the queue that matters most.
  var rank = { high: 0, medium: 1, low: 2 };
  pool.sort(function (a, b) {
    return (rank[String(a.priority || 'low').toLowerCase()] || 2) - (rank[String(b.priority || 'low').toLowerCase()] || 2);
  });
  pool = pool.slice(0, Math.min(Number(data && data.limit) || 90, 90));

  var byId = {};
  pool.forEach(function (i) { byId[i.issue_id] = i; });

  var suggestions = [];
  var BATCH = 15; // small batches: a truncated reply loses one batch, not the run
  for (var b = 0; b < pool.length; b += BATCH) {
    var batch = pool.slice(b, b + BATCH).map(function (i) {
      return {
        issue_id: i.issue_id,
        lesson_code: i.lesson_code || '',
        summary: String(i.summary || '').slice(0, 220),
        priority: String(i.priority || '').toLowerCase(),
        reports: Number(i.report_count) || 1,
        days_open: Math.max(0, Math.round((Date.now() - new Date(i.submitted_at || Date.now())) / 864e5)),
        part: i.media_kind || '',
        kind: i.request_kind || 'fix'
      };
    });
    var prompt = 'You are reviewing the open course-fix queue for Ardent Training, an online RYA sailing theory school. ' +
      'Each item is an error or problem reported in course content, with the priority it was given at the time.\n\n' +
      'Priority rules:\n' +
      '- high: blocks students completely, a factual error that could cause an exam failure, or safety-critical content being wrong.\n' +
      '- medium: a confusing content error that is not blocking anyone, or something students need a workaround for.\n' +
      '- low: a minor typo, a cosmetic issue, or a single confused student where the content is probably fine.\n\n' +
      'Things that should nudge a priority UP: several separate reports of the same fix, or a factual error sitting open a long time. ' +
      'Things that should nudge one DOWN: a lone cosmetic report marked high at logging time.\n\n' +
      'THE QUEUE:\n' + JSON.stringify(batch) + '\n\n' +
      'Return ONLY JSON, no prose, no fences: {"suggestions":[{"issue_id":"...","suggest_priority":"high|medium|low","why":"one plain sentence"}]}. ' +
      'Include ONLY issues whose current priority looks wrong. An empty list is a perfectly good answer - most priorities are usually fine.';
    // 16k, not a snug fit: sonnet-5 thinks before it answers on a prompt this
    // size, and the thinking spends from the same max_tokens budget - at 4000
    // it burned the lot thinking and returned no JSON at all (live, 9 Aug).
    var got = anthropicRaw_(EXTRACTION_MODEL, prompt, 16000);
    var res = got.json;
    if (res === null) {
      // Announce the failure rather than quietly returning a thin result
      // (the silent-fallback lesson from the troubleshoot helper), and say
      // what actually came back so the fault is debuggable from the toast.
      // The billing case gets said in plain words (Edd, FB-0201: the raw
      // HTTP 400 read as "the button doesn't work" when the account was
      // simply out of API credit - and every AI feature was down with it).
      if (/credit balance is too low/i.test(String(got.why || ''))) {
        return { ok: false, error: 'The account that pays for the automatic reading is out of credit, so this check (and extraction, drafts, and the scan with it) cannot run. Top up at console.anthropic.com > Plans & Billing, then press the button again. The button itself is fine.' };
      }
      return { ok: false, error: 'The priority check failed part-way (batch starting at ' + (b + 1) + ' of ' + pool.length + '): ' + got.why + '. Try again in a minute.' };
    }
    (res.suggestions || []).forEach(function (s) {
      var i = byId[s.issue_id];
      if (!i) return;
      var suggest = String(s.suggest_priority || '').toLowerCase();
      if (['high', 'medium', 'low'].indexOf(suggest) < 0) return;
      if (suggest === String(i.priority || '').toLowerCase()) return; // no-op "suggestion"
      suggestions.push({
        issue_id: i.issue_id,
        lesson_code: i.lesson_code || '',
        summary: String(i.summary || '').slice(0, 160),
        current: String(i.priority || '').toLowerCase(),
        suggest: suggest,
        why: String(s.why || '').slice(0, 300)
      });
    });
  }
  return { ok: true, reviewed: pool.length, suggestions: suggestions };
}

// People an issue can be assigned to: active users who can work a fix queue.
// Each entry carries which queues they can take (course fixes and/or developer),
// so the front-end can offer the right people for the right queue.
function listAssignees_() {
  var sheet = usersSheet_();
  if (!sheet) return { ok: true, assignees: [] };
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return { ok: true, assignees: [] };
  var head = values[0]; var idx = {}; head.forEach(function (h, i) { idx[h] = i; });
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var u = rowToUser_(values[r], idx);
    if (String(u.status).toLowerCase() !== 'active') continue;
    var p = permsOf_(u);
    if (!p.course && !p.dev) continue; // only people who can take a fix queue
    out.push({ name: u.name || u.email, email: u.email, course: !!p.course, dev: !!p.dev });
  }
  return { ok: true, assignees: out };
}

// Add a follow-up update to an existing issue (the same ongoing problem), so we
// keep one entry per issue rather than logging a new one each time. Appends the
// new conversation to the history, refreshes the priority (which can rise as the
// usual fixes get tried), and bumps updated_at. Does not add a new reporter.
function addUpdate_(data) {
  var id = data.issue_id;
  if (!id) return { ok: false, error: 'addUpdate needs an issue_id' };
  var found = findRow_(id);
  if (!found) return { ok: false, error: 'No issue found with id ' + id };
  var rec = found.record;
  // Slack fires on the CHANGE to high, not on every update to something already
  // there (Edd, 19 Aug 2026). Read it before this update touches the record.
  var priorityBefore = String(rec.priority || '').toLowerCase();

  var who = (data._user && data._user.name) || data.instructor_name || 'someone';
  var stamp = new Date().toISOString().slice(0, 10);
  var note = data.raw_text || data.summary || '';
  rec.raw_text = (rec.raw_text || '') + '\n\n--- update ' + stamp + ' by ' + who + ' ---\n' + note;

  // Add this update to the timeline (shown as an accordion entry in the detail).
  var reps = [];
  try { reps = rec.reports_json ? JSON.parse(rec.reports_json) : []; } catch (e) { reps = []; }
  reps.push({
    kind: 'update',
    student_name: data.student_name || '', student_contact: data.student_contact || '',
    device_info: data.device_info || '', instructor_name: who,
    summary: data.summary || '', priority: String(data.priority || '').toLowerCase(),
    raw_text: note, recommended_steps: data.recommended_steps || null,
    // waiting: the instructor has replied and the ball is with the student -
    // the Actions lane reads this to show "Waiting" (Edd, FB-0188).
    waiting: data.waiting ? true : undefined,
    date: new Date().toISOString()
  });
  rec.reports_json = JSON.stringify(reps);

  if (data.priority) rec.priority = String(data.priority).toLowerCase();
  if (data.priority_reason) rec.priority_reason = data.priority_reason;
  if (data.device_info) rec.device_info = data.device_info;

  if (data.image_urls) {
    var existing = rec.image_urls ? String(rec.image_urls).split(',') : [];
    var more = normaliseImageUrls_(data.image_urls).split(',').filter(Boolean);
    rec.image_urls = existing.concat(more).filter(Boolean).join(',');
  }

  // If the issue is currently parked as Resolved - TBC and a fresh update comes
  // in, reopen it (the suggested fix evidently did not stick). keep_status
  // opts out: an administrative note (a correction, an "I've replied" marker)
  // is not fresh news from the student and must not reopen anything (r52).
  if (!data.keep_status) {
    if (String(rec.status).toLowerCase() === 'resolved_tbc') rec.status = 'open';
    // Same reasoning for a parked one. It was parked because the student went
    // quiet, and a follow-up means they haven't, so there is something to go on
    // again. The park branch below can set it straight back if the update is
    // itself an "Update and park".
    if (String(rec.status).toLowerCase() === 'parked') rec.status = 'open';
  }

  // If the pasted follow-up shows the problem was sorted out in the chat, close
  // the issue and save the answer, rather than only gluing more text on.
  if (data.resolved) {
    rec.status = 'resolved';
    rec.resolved_at = data.resolved_at || new Date().toISOString();
    if (data.resolution_note) rec.resolution_note = data.resolution_note;
    // Sorted in the chat itself, with the student there: nobody to notify.
    rec.notified_students = true;
  } else if (data.tbc) {
    rec.status = 'resolved_tbc';
    if (data.resolution_note) rec.resolution_note = data.resolution_note;
  } else if (data.park) {
    // "Update and park": stalled rather than finished. The student stopped
    // replying, so there is nothing left to chase and it shouldn't sit in an
    // open queue making the numbers look worse than they are. It wakes up on
    // its own if somebody else reports the same fault (Edd, FB-0134).
    rec.status = 'parked';
    if (data.resolution_note) rec.resolution_note = data.resolution_note;
  }

  rec.updated_at = new Date().toISOString();
  found.sheet.getRange(found.rowNum, 1, 1, HEADERS.length).setValues([recordToRow_(rec)]);

  if (String(rec.priority).toLowerCase() === 'high' && priorityBefore !== 'high' &&
      rec.status !== 'resolved' && rec.status !== 'resolved_tbc') {
    try { sendSlack_(rec, data.app_url || getAppUrl_()); } catch (e) {}
  }
  if (String(rec.status).toLowerCase() === 'resolved_tbc') {
    try { checkSharedWorkaround_(rec, data.app_url || getAppUrl_()); } catch (e) {}
  }
  return { ok: true, issue_id: id, updated: true };
}

// Delete an issue (e.g. logged by accident). Full admins ('users' permission)
// can delete any issue; everyone else can only delete one they logged.
function deleteIssue_(data) {
  var id = data.issue_id;
  if (!id) return { ok: false, error: 'deleteIssue needs an issue_id' };
  var found = findRow_(id);
  if (!found) return { ok: false, error: 'No issue found with id ' + id };
  var user = data._user || {};
  var perms = permsOf_(user);
  var isOwn = found.record.instructor_name && user.name &&
    String(found.record.instructor_name).trim().toLowerCase() === String(user.name).trim().toLowerCase();
  if (!perms.users && !isOwn) return { ok: false, error: 'You can only delete issues you logged.' };
  found.sheet.deleteRow(found.rowNum);
  return { ok: true };
}

// Manually link one issue into another (combine them), the inverse of split.
// The source's reports, text and images move onto the target, and the source
// row is removed, so they become one entry.
function linkIssues_(data) {
  var sourceId = data.source_id, targetId = data.target_id;
  if (!sourceId || !targetId) return { ok: false, error: 'linkIssues needs source_id and target_id' };
  if (sourceId === targetId) return { ok: false, error: 'Cannot link an issue to itself.' };
  var src = findRow_(sourceId);
  var tgt = findRow_(targetId);
  if (!src) return { ok: false, error: 'Source issue not found.' };
  if (!tgt) return { ok: false, error: 'Target issue not found.' };
  var s = src.record, t = tgt.record;

  var sReps = []; try { sReps = s.reports_json ? JSON.parse(s.reports_json) : []; } catch (e) { sReps = []; }
  var tReps = []; try { tReps = t.reports_json ? JSON.parse(t.reports_json) : []; } catch (e) { tReps = []; }
  if (!sReps.length) {
    sReps = [{ kind: 'report', student_name: s.student_name || '', student_contact: s.student_contact || '',
      device_info: s.device_info || '', instructor_name: s.instructor_name || '', summary: s.summary || '',
      raw_text: s.raw_text || '', date: s.submitted_at || '' }];
  }
  t.reports_json = JSON.stringify(tReps.concat(sReps));
  t.raw_text = (t.raw_text || '') + '\n\n--- linked in from another report ---\n' + (s.raw_text || '');

  var ti = t.image_urls ? String(t.image_urls).split(',') : [];
  var si = s.image_urls ? String(s.image_urls).split(',') : [];
  t.image_urls = ti.concat(si).filter(Boolean).join(',');

  t.report_count = (parseInt(t.report_count, 10) || 1) + (parseInt(s.report_count, 10) || 1);

  var order = ['low', 'medium', 'high'];
  var pi = Math.max(order.indexOf(String(t.priority || 'low').toLowerCase()), order.indexOf(String(s.priority || 'low').toLowerCase()), 0);
  t.priority = order[pi];
  t.updated_at = new Date().toISOString();

  tgt.sheet.getRange(tgt.rowNum, 1, 1, HEADERS.length).setValues([recordToRow_(t)]);
  src.sheet.deleteRow(src.rowNum);
  return { ok: true, target_id: targetId };
}

// ---- Developer handoff ----------------------------------------------------

// Manually hand an issue to the developers (from the admin view).
function passToDev_(data) {
  var found = findRow_(data.issue_id);
  if (!found) return { ok: false, error: 'No issue found with id ' + data.issue_id };
  var rec = found.record;
  if (!rec.dev_passed_at) rec.dev_passed_at = new Date().toISOString();
  rec.dev_fixed_at = '';            // if it was previously fixed and is going back
  rec.status = 'with_dev';
  rec.updated_at = new Date().toISOString();
  found.sheet.getRange(found.rowNum, 1, 1, HEADERS.length).setValues([recordToRow_(rec)]);
  return { ok: true };
}

// A developer or course-team member marks an issue fixed. No automatic "please
// verify every fix" ping any more (Edd's call, 18 Jul): Slack only fires when
// the fixer ticks "Notify student?", and the message asks the instructors to
// let the student know rather than asking anyone to re-verify the fix.
function markDevFixed_(data) {
  var found = findRow_(data.issue_id);
  if (!found) return { ok: false, error: 'No issue found with id ' + data.issue_id };
  var rec = found.record;
  rec.dev_fixed_at = new Date().toISOString();
  if (!rec.dev_passed_at) rec.dev_passed_at = rec.dev_fixed_at;
  // A course error marked fixed IS fixed (Edd, FB-0195): the course team just
  // changed the slide themselves, so asking an instructor to go and check it
  // again is make-work. Straight to resolved; notified_students stays as it
  // was, so the "tell the student" lane still fires where students are
  // affected. Tech issues keep the dev_fixed stop - a code fix genuinely
  // needs someone to see it work before the student's told.
  if (String(rec.category || '').toLowerCase() === 'course_error') {
    rec.status = 'resolved';
    rec.resolved_at = rec.dev_fixed_at;
    if (!rec.resolution_note) rec.resolution_note = 'Fixed by the course team' + (data.dev_notes ? ': ' + data.dev_notes : '.');
  } else {
    rec.status = 'dev_fixed';
  }
  if (data.dev_notes != null) rec.dev_notes = data.dev_notes;
  rec.updated_at = new Date().toISOString();
  found.sheet.getRange(found.rowNum, 1, 1, HEADERS.length).setValues([recordToRow_(rec)]);
  if (data.notify_student) {
    try { sendNotifyStudentSlack_(rec, data.app_url || getAppUrl_()); } catch (e) {}
  }
  return { ok: true };
}

function sendNotifyStudentSlack_(issue, appUrl) {
  if (!slackOn_('notify_student')) return;
  var isCourse = String(issue.category).toLowerCase() === 'course_error';
  var student = (issue.student_name || '') + (issue.student_contact ? ' (' + issue.student_contact + ')' : '');
  var text = [
    ':white_check_mark: *' + (isCourse ? 'Course fix done' : 'Fix done') + ' - student to notify*',
    '*Lesson:* ' + (issue.lesson || '-') + ' (' + (issue.lesson_code || '-') + ')',
    '*Summary:* ' + slackSummary_(issue),
    '*Fix notes:* ' + (issue.dev_notes || '-'),
    '*Student:* ' + (student || '-'),
    '*Logged by:* ' + (issue.instructor_name || '-'),
    '',
    'Instructors: please let the student know this is sorted (it\'s also in your Actions tab).',
    'Open this issue: ' + issueLink_(issue, appUrl)
  ].join('\n');
  slackPost_('notify_student', text);
}

// A developer or course-team member is stuck and needs something from the
// admins before they can carry on (more detail, a decision, access to
// something). Raises a flag on the issue (shown as a badge wherever it's
// listed) and pings Slack so it doesn't just sit there. The question is also
// logged in the timeline (reports_json) so the exchange stays on the record
// even after it's answered and the flag is cleared.
function flagQuery_(data) {
  var id = data.issue_id;
  if (!id) return { ok: false, error: 'flagQuery needs an issue_id' };
  var found = findRow_(id);
  if (!found) return { ok: false, error: 'No issue found with id ' + id };
  var rec = found.record;

  var question = String(data.question || '').trim();
  if (!question) return { ok: false, error: 'Add a question before sending it.' };

  var target = data.target === 'instructor' ? 'instructor' : 'admins';
  var who = (data._user && data._user.name) || 'Someone';
  var now = new Date().toISOString();
  rec.dev_query = question;
  rec.dev_query_at = now;
  rec.dev_query_by = who;
  rec.dev_query_target = target;
  rec.updated_at = now;

  var forWhom = target === 'instructor' ? ('Question for ' + (rec.instructor_name || 'the instructor')) : 'Question for the admins';
  var reps = [];
  try { reps = rec.reports_json ? JSON.parse(rec.reports_json) : []; } catch (e) { reps = []; }
  reps.push({ kind: 'question', instructor_name: who, summary: forWhom, raw_text: question, date: now });
  rec.reports_json = JSON.stringify(reps);
  rec.report_count = reps.length;

  found.sheet.getRange(found.rowNum, 1, 1, HEADERS.length).setValues([recordToRow_(rec)]);
  try { sendQueryRaisedSlack_(rec, data.app_url || getAppUrl_()); } catch (e) {}
  return { ok: true };
}

// Deep link straight to one issue: the front end reads ?issue= and opens that
// issue's drawer after login, so a Slack ping lands people on the exact record
// rather than the top of the tracker.
function issueLink_(issue, appUrl) {
  var url = appUrl || getAppUrl_();
  if (!url) return '(app url not set)';
  return url + (url.indexOf('?') > -1 ? '&' : '?') + 'issue=' + issue.issue_id;
}

function sendQueryRaisedSlack_(issue, appUrl) {
  if (!slackOn_('query_raised')) return;
  var isCourse = String(issue.category).toLowerCase() === 'course_error';
  var toInstructor = issue.dev_query_target === 'instructor';
  var text = [
    ':grey_question: *' + (issue.dev_query_by || 'Someone') + ' has a question for ' +
      (toInstructor ? (issue.instructor_name || 'the instructor') : 'the admins') + '*',
    '*Lesson:* ' + (issue.lesson || '-') + ' (' + (issue.lesson_code || '-') + ')',
    '*Summary:* ' + slackSummary_(issue),
    '*Question:* ' + (issue.dev_query || '-'),
    '',
    toInstructor
      ? ((issue.instructor_name || 'The instructor') + ', reply from your Actions tab so the fix can carry on.')
      : ('Reply in the tracker so ' + (isCourse ? 'the course team' : 'the developer') + ' can carry on.'),
    'Open this issue: ' + issueLink_(issue, appUrl)
  ].join('\n');
  slackPost_('query_raised', text);
}

// An admin answers a developer's/course team's question. Logs the reply in
// the timeline, clears the flag (so the badge disappears everywhere it's
// shown), and lets Slack know so the person waiting isn't stuck refreshing
// the tracker to find out.
function answerQuery_(data) {
  var id = data.issue_id;
  if (!id) return { ok: false, error: 'answerQuery needs an issue_id' };
  var found = findRow_(id);
  if (!found) return { ok: false, error: 'No issue found with id ' + id };
  var rec = found.record;

  // Who may answer depends on who the question was for: admin-targeted
  // questions need the users perm (unchanged from Round 11); instructor-
  // targeted ones can be answered by ANY instructor (the whole team sees the
  // Actions list, and whoever knows the answer shouldn't have to wait - Edd,
  // 21 Jul) or by an admin.
  var perms = permsOf_(data._user || {});
  if (rec.dev_query_target === 'instructor') {
    if (!perms.users && !perms.log) return { ok: false, error: 'forbidden' };
  } else if (!perms.users) {
    return { ok: false, error: 'forbidden' };
  }

  var reply = String(data.reply || '').trim();
  if (!reply) return { ok: false, error: 'Add a reply before sending it back.' };

  var who = (data._user && data._user.name) || 'An admin';
  var question = rec.dev_query || '';
  var now = new Date().toISOString();

  var reps = [];
  try { reps = rec.reports_json ? JSON.parse(rec.reports_json) : []; } catch (e) { reps = []; }
  reps.push({ kind: 'answer', instructor_name: who, summary: 'Reply to question', raw_text: reply, date: now });
  rec.reports_json = JSON.stringify(reps);
  rec.report_count = reps.length;

  var asker = rec.dev_query_by || '';
  rec.dev_query = '';
  rec.dev_query_at = '';
  rec.dev_query_by = '';
  rec.dev_query_target = '';
  rec.updated_at = now;

  found.sheet.getRange(found.rowNum, 1, 1, HEADERS.length).setValues([recordToRow_(rec)]);
  try { sendQueryAnsweredSlack_(rec, question, reply, asker, data.app_url || getAppUrl_()); } catch (e) {}
  return { ok: true };
}

function sendQueryAnsweredSlack_(issue, question, reply, asker, appUrl) {
  if (!slackOn_('query_answered')) return;
  var text = [
    ':speech_balloon: *Question answered*' + (asker ? ' - ' + asker + ', this one\'s for you' : ''),
    '*Lesson:* ' + (issue.lesson || '-') + ' (' + (issue.lesson_code || '-') + ')',
    '*Summary:* ' + slackSummary_(issue),
    '*Question:* ' + (question || '-'),
    '*Reply:* ' + (reply || '-'),
    '',
    'Open this issue: ' + issueLink_(issue, appUrl)
  ].join('\n');
  slackPost_('query_answered', text);
}

// The instructor gave the student a workaround and marked it Resolved-TBC, but
// couldn't check at the time whether it was a one-off or a fault for everyone.
// The reminder goes to Slack STRAIGHT AWAY (Edd's call, 18 Jul) so it sits
// waiting in the channel rather than arriving tomorrow; recheck_at is stamped
// purely as a record that a reminder went out.
function requestRecheck_(data) {
  var id = data.issue_id;
  if (!id) return { ok: false, error: 'requestRecheck needs an issue_id' };
  var found = findRow_(id);
  if (!found) return { ok: false, error: 'No issue found with id ' + id };
  var rec = found.record;
  rec.recheck_at = new Date().toISOString();
  rec.updated_at = rec.recheck_at;
  found.sheet.getRange(found.rowNum, 1, 1, HEADERS.length).setValues([recordToRow_(rec)]);
  try { sendRecheckSlack_(rec, data.app_url || getAppUrl_()); } catch (e) {}
  return { ok: true };
}

function sendRecheckSlack_(rec, appUrl) {
  if (!slackOn_('recheck')) return;
  var text = [
    ':alarm_clock: *Workaround needs a double-check*',
    '*Lesson:* ' + (rec.lesson || '-') + ' (' + (rec.lesson_code || '-') + ')',
    '*Summary:* ' + slackSummary_(rec),
    '*Workaround given:* ' + (rec.resolution_note || '-'),
    '',
    (rec.instructor_name || 'Whoever logged this') + " couldn't check at the time: is this a one-off for that student, or a fault for everyone? " +
      "One-off - leave it as Resolved - TBC. Everyone - reopen it so it gets a proper fix.",
    'Open this issue: ' + issueLink_(rec, appUrl)
  ].join('\n');
  slackPost_('recheck', text);
}

// ---- Shared workaround watch ----------------------------------------------

// Sort a workaround note into one of a handful of things we actually tell
// students to try. Keyword matching rather than anything clever, because the
// notes are written by hand and we only need them grouped, not understood.
// Returns null when the note is something one-off (a wrong username, a payment
// chased up), since those genuinely are individual and shouldn't be counted.
// First match wins, so a note that mentions two things lands in the first
// bucket. That is fine here: we are looking for a pattern, not an audit.
// Checked against the 166 resolution notes on the sheet (4 Aug): it picks up
// the 10 real workarounds and leaves the genuine fixes alone.
function workaroundKind_(text) {
  var t = String(text || '').toLowerCase();
  if (!t) return null;
  var kinds = [
    ['incognito',        /incognito|private (window|browsing)|inprivate/],
    ['a different browser', /(different|another|other|alternative) browser|(switch|swap|mov)(ed|ing)? (from [a-z ]{0,25})?to (google |mozilla |microsoft |apple )?(chrome|firefox|safari|edge)|(worked|works|logged in|got in|tried|opened) [a-z ]{0,15}?(in|on|with|using) (google |mozilla |microsoft |apple )?(chrome|firefox|safari|edge)|us(ed|ing) (google |mozilla |microsoft |apple )?(chrome|firefox|safari|edge) instead/],
    ['clearing the cache', /clear(ed|ing)?( the)? (cache|cookies|browsing data)|hard refresh|ctrl ?\+ ?f5/],
    ['turning off extensions', /extension|ad ?block|plug-?in blocking/],
    ['a reinstall',      /re-?install|delete(d)? and (re-?)?install|uninstall/],
    ['a different device', /(different|another|other) (device|computer|laptop|phone|tablet|ipad|machine)|tried it on (his|her|their) (phone|laptop|ipad|tablet)/],
    ['a different network', /(mobile|cellular) data|different (network|wifi|wi-?fi)|off (the )?(work|office|school) (wifi|wi-?fi|network)|hotspot|vpn/],
    ['logging out and back in', /log(ged|ging)? out and (back )?in|sign(ed|ing)? out and (back )?in|fresh login/]
  ];
  for (var i = 0; i < kinds.length; i++) {
    if (kinds[i][1].test(t)) return kinds[i][0];
  }
  return null;
}

// Called whenever an issue lands at Resolved - TBC. If enough recent TBCs share
// the same workaround, ping Slack once so somebody can look at the pattern
// rather than at eight separate closed tickets. Never allowed to break a save,
// so every caller wraps it in a try/catch and it swallows its own errors too.
function checkSharedWorkaround_(rec, appUrl) {
  var kind = workaroundKind_(rec && rec.resolution_note);
  if (!kind) return { ok: true, kind: null };

  var cutoff = Date.now() - SHARED_WORKAROUND_DAYS * 24 * 3600 * 1000;
  var matches = [];
  var seen = {};
  ISSUE_SHEETS.forEach(function (name) {
    var sheet = sheetByName_(name);
    if (!sheet) return;
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return;
    var idx = {}; values[0].forEach(function (h, i) { idx[h] = i; });
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var id = row[idx['issue_id']];
      if (!id || seen[id]) continue;
      if (String(row[idx['status']]).toLowerCase() !== 'resolved_tbc') continue;
      var when = new Date(row[idx['updated_at']] || row[idx['submitted_at']]);
      if (isNaN(when.getTime()) || when.getTime() < cutoff) continue;
      if (workaroundKind_(row[idx['resolution_note']]) !== kind) continue;
      seen[id] = true;
      matches.push({
        issue_id: id,
        lesson: row[idx['lesson']],
        lesson_code: row[idx['lesson_code']],
        summary: row[idx['summary']],
        raw_text: row[idx['raw_text']],
        instructor_name: row[idx['instructor_name']],
        student_name: row[idx['student_name']]
      });
    }
  });

  if (matches.length < SHARED_WORKAROUND_MIN) return { ok: true, kind: kind, count: matches.length };

  // Only shout once per workaround per window. Otherwise every close after the
  // third one pings again and the channel learns to ignore it.
  var props = PropertiesService.getScriptProperties();
  var log = {};
  try { log = JSON.parse(props.getProperty('SHARED_WORKAROUND_PINGED') || '{}'); } catch (e) { log = {}; }
  var last = new Date(log[kind] || 0).getTime();
  if (last && last >= cutoff) return { ok: true, kind: kind, count: matches.length, skipped: 'already flagged' };

  try {
    sendSharedWorkaroundSlack_(kind, matches, appUrl || getAppUrl_());
    log[kind] = new Date().toISOString();
    props.setProperty('SHARED_WORKAROUND_PINGED', JSON.stringify(log));
  } catch (e) {
    return { ok: false, error: String(e) };
  }
  return { ok: true, kind: kind, count: matches.length, flagged: true };
}

function sendSharedWorkaroundSlack_(kind, matches, appUrl) {
  if (!slackOn_('shared_workaround')) return;
  var lines = [
    ':mag: *' + matches.length + ' issues this week all sorted by ' + kind + '*',
    'Each one closed as Resolved - TBC on its own, so nothing looks wrong until you put them side by side. That many students needing the same workaround usually means the underlying fault is still there for everybody else.',
    ''
  ];
  matches.slice(0, 6).forEach(function (m) {
    var where = m.lesson_code ? ' (' + m.lesson_code + ')' : '';
    lines.push('• ' + slackSummary_(m) + where + ' - logged by ' + (m.instructor_name || 'someone') +
      ' - ' + issueLink_(m, appUrl));
  });
  if (matches.length > 6) lines.push('• plus ' + (matches.length - 6) + ' more');
  lines.push('');
  lines.push('Worth one of us reopening the clearest one and sending it to the developers, rather than waiting for the next report.');
  slackPost_('shared_workaround', lines.join('\n'));
}

// Kept only so any leftover daily trigger from the first Round 13 deploy
// doesn't error: recheck reminders now go to Slack immediately (see
// requestRecheck_), so there is nothing for a scheduled pass to do. The
// trigger itself is removed by ensureTriggers_ next time setup() runs.
function sendRecheckReminders() {}

// Daily: shipping issues whose chase date has arrived. A parcel problem goes
// quiet unless somebody pokes the courier, so this is the nudge (Edd, 30 Jul).
// The chase date is cleared as it fires, so one nudge per chase, and setting a
// new date re-arms it.
function chaseShipping() {
  var sheet = sheetByName_(SHIPPING_SHEET);
  if (!sheet) return;
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  var idx = {}; values[0].forEach(function (h, i) { idx[h] = i; });
  if (idx.chase_at == null) return;
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var due = [];
  for (var r = 1; r < values.length; r++) {
    if (!values[r][idx.issue_id]) continue;
    var when = dayStr_(values[r][idx.chase_at]);
    if (!when || when > today) continue;
    var st = String(values[r][idx.status] || '').toLowerCase();
    if (st === 'resolved' || st === 'past' || st === 'parked') { sheet.getRange(r + 1, idx.chase_at + 1).setValue(''); continue; }
    var rec = {}; HEADERS.forEach(function (h) { rec[h] = idx[h] != null ? values[r][idx[h]] : ''; });
    due.push(rec);
    sheet.getRange(r + 1, idx.chase_at + 1).setValue('');
  }
  if (!due.length) return;
  // The chase dates above have been cleared either way: this job's bookkeeping
  // is the point, the Slack nudge was only ever the reminder on top of it.
  if (!slackOn_('shipping_chase')) return;
  var appUrl = getAppUrl_();
  var lines = [':package: *' + due.length + ' shipping issue' + (due.length === 1 ? '' : 's') + ' due a chase*'];
  due.slice(0, 8).forEach(function (i) {
    lines.push('• ' + (i.courier ? i.courier + ' ' : '') + (normaliseTracking_(i.tracking_number) || '(no tracking)') +
      ' - ' + String(i.summary || '').slice(0, 80) + (i.student_name ? ' (' + i.student_name + ')' : '') +
      '\n  ' + issueLink_(i, appUrl));
  });
  slackPost_('shipping_chase', lines.join('\n'));
  Logger.log('chaseShipping: nudged ' + due.length);
}

// Runs monthly (trigger created by setup()). Reads the last month of tech
// issues and asks the AI whether the pre-developer troubleshooting checklist
// is missing anything, posting suggestions (or a quiet all-clear) to Slack.
function monthlyChecklistReview() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return;
  var cutoff = Date.now() - 31 * 24 * 60 * 60 * 1000;
  var recent = [];
  getIssues_().issues.forEach(function (i) {
    if (String(i.category).toLowerCase() !== 'tech_issue') return;
    if (new Date(i.submitted_at).getTime() < cutoff) return;
    recent.push({ summary: i.summary || '', device: i.device_info || '', section: i.section || '', resolution: i.resolution_note || '' });
  });
  if (!recent.length) return;

  var listText = CHECKLIST_ITEMS.map(function (c) { return '- ' + c.label; }).join('\n');
  var prompt = 'Ardent Training (online sailing courses) has a pre-developer troubleshooting checklist that instructors ' +
    'work through with students before a tech issue is escalated. Current checklist:\n' + listText + '\n\n' +
    'Here are the tech issues logged in the last month (summary, device, section, and how they were resolved where known):\n' +
    JSON.stringify(recent.slice(0, 80)) + '\n\n' +
    'Suggest up to 3 NEW checklist steps that would have helped resolve or triage these issues sooner, if any genuinely ' +
    'recur and are missing from the list. Each needs a short label and one line on why (grounded in the issues above). ' +
    'If the current checklist already covers what came up, say so instead of inventing additions. ' +
    'Return ONLY JSON: {"suggestions": [{"label": "...", "why": "..."}]} with an empty array if none are needed.';

  var out;
  try {
    var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 500, messages: [{ role: 'user', content: prompt }] })
    });
    var payload = JSON.parse(res.getContentText());
    out = JSON.parse(payload.content[0].text.replace(/^```json?\s*|\s*```$/g, ''));
  } catch (e) { return; }

  var sugg = (out && out.suggestions) || [];
  var text = sugg.length
    ? [':clipboard: *Monthly checklist review* - ' + recent.length + ' tech issues looked at, ' + sugg.length + ' suggested addition' + (sugg.length > 1 ? 's' : '') + ':']
        .concat(sugg.map(function (s) { return '• *' + s.label + '* - ' + s.why; }))
        .concat(['', 'If one earns its place, tell Claude to add it to the checklist.']).join('\n')
    : ':clipboard: *Monthly checklist review* - ' + recent.length + ' tech issues looked at, the current checklist already covers what came up. Nothing to add.';
  slackPost_('monthly_checklist', text);
}

// Rotate the Slack webhook without opening Project Settings: the URL travels
// in the POST body (never in this file, which is public on GitHub) and lands
// straight in Script Properties. DEPLOY_KEY gated, same as deployBackend.
function setSlackWebhook_(data) {
  var key = PropertiesService.getScriptProperties().getProperty('DEPLOY_KEY');
  if (!key || String(data.key || '') !== key) return { ok: false, error: 'bad deploy key' };
  var url = String(data.webhook || '');
  if (!/^https:\/\/hooks\.slack\.com\/services\//.test(url)) return { ok: false, error: 'not a Slack webhook URL' };
  PropertiesService.getScriptProperties().setProperty('SLACK_WEBHOOK_URL', url);
  return { ok: true };
}

// ---- Chatwoot ------------------------------------------------------------
// Instructors were copy-pasting live-chat transcripts by hand; this pulls them
// straight from Chatwoot instead, with the student's name and email attached.
// Credentials live in Script Properties (never in this file, which is public
// on GitHub): CHATWOOT_TOKEN and CHATWOOT_ACCOUNT_ID, set via
// setChatwootConfig below. Cloud only for now (app.chatwoot.com).
var CHATWOOT_BASE = 'https://app.chatwoot.com';

function setChatwootConfig_(data) {
  var key = PropertiesService.getScriptProperties().getProperty('DEPLOY_KEY');
  if (!key || String(data.key || '') !== key) return { ok: false, error: 'bad deploy key' };
  var props = PropertiesService.getScriptProperties();
  if (data.token) props.setProperty('CHATWOOT_TOKEN', String(data.token));
  if (data.account_id) props.setProperty('CHATWOOT_ACCOUNT_ID', String(data.account_id));
  return { ok: true };
}
function chatwootCfg_() {
  var p = PropertiesService.getScriptProperties();
  return { token: p.getProperty('CHATWOOT_TOKEN') || '', account: p.getProperty('CHATWOOT_ACCOUNT_ID') || '' };
}
function chatwootCall_(path, method, payload) {
  var cfg = chatwootCfg_();
  if (!cfg.token || !cfg.account) throw new Error('Chatwoot is not configured yet.');
  var res = UrlFetchApp.fetch(CHATWOOT_BASE + '/api/v1/accounts/' + cfg.account + path, {
    method: method || 'get', contentType: 'application/json',
    headers: { api_access_token: cfg.token },
    payload: payload ? JSON.stringify(payload) : undefined,
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code < 200 || code >= 300) throw new Error('Chatwoot ' + code + ': ' + body.slice(0, 200));
  return body ? JSON.parse(body) : {};
}
// Accepts a full conversation URL, or just the number.
function chatwootConvId_(input) {
  var s = String(input || '').trim();
  var m = s.match(/conversations\/(\d+)/);
  if (m) return m[1];
  m = s.match(/^\d+$/);
  return m ? s : '';
}
// Strip the noise real chat messages carry: inline image blobs, email
// signatures, and quoted reply chains. The AI reads better without them, and
// a 3,000-character signature is 3,000 characters of nothing.
function cleanChatwootBody_(raw) {
  var s = String(raw || '');
  s = s.replace(/!\[[^\]]*\]\([^)]*\)/g, '[image]');       // markdown images
  s = s.replace(/<img[^>]*>/gi, '[image]');
  s = s.replace(/\\\n/g, '\n');                            // escaped newlines
  s = s.split(/\n-- ?\n/)[0];                              // standard sig delimiter
  s = s.replace(/\n>[^\n]*/g, '');                          // quoted reply lines
  // Common sign-off block: cut everything after it if it looks like a footer.
  s = s.replace(/\n+(Kind regards|Best regards|All the best|Many thanks|Cheers|Regards)[\s\S]{0,400}$/i, function (m) {
    return /\n.*\n/.test(m) ? '' : m; // only if it spans several lines (a footer, not a one-line sign-off)
  });
  return s.replace(/\n{3,}/g, '\n\n').trim();
}

// Pull one conversation and flatten it into the kind of transcript the
// extraction prompt already understands.
function chatwootImport_(data) {
  var id = chatwootConvId_(data.conversation);
  if (!id) return { ok: false, error: 'Paste a Chatwoot conversation link or number.' };
  var conv, msgs;
  try {
    conv = chatwootCall_('/conversations/' + id);
    msgs = chatwootCall_('/conversations/' + id + '/messages');
  } catch (e) { return { ok: false, error: String(e.message || e) }; }

  var list = (msgs && (msgs.payload || msgs.data && msgs.data.payload)) || [];
  var sender = (conv && conv.meta && conv.meta.sender) || {};
  // The conversation's sender block often carries a name and no email, because
  // a web-widget contact only gets one once they hand it over. The contact
  // record itself usually has it, so go and read that rather than sending the
  // instructor off to find it by hand (Edd, 1 Aug).
  if (!sender.email && sender.id) {
    try {
      var contact = chatwootCall_('/contacts/' + sender.id);
      var payload = (contact && (contact.payload || contact)) || {};
      var extra = payload.additional_attributes || {};
      sender.email = payload.email || extra.email || sender.email || '';
      sender.phone_number = sender.phone_number || payload.phone_number || '';
      if (!sender.name) sender.name = payload.name || '';
    } catch (e) {}   // a contact we can't read is no reason to lose the transcript
  }
  var lines = [];
  var attImages = [];
  list.forEach(function (m) {
    // 0 incoming (student), 1 outgoing (agent). Skip activity lines and
    // internal notes - they're noise in a transcript the AI has to read.
    var t = Number(m.message_type);
    if (t !== 0 && t !== 1) return;
    if (m.private) return;
    var body = cleanChatwootBody_(m.content);
    // Image attachments used to vanish entirely: a message that was only a
    // screenshot had no text, so it never even reached the transcript and
    // nobody downstream knew an image existed (Edd, 9 Aug).
    var pics = (m.attachments || []).filter(function (a) { return String(a.file_type) === 'image' && a.data_url; });
    if (!body && !pics.length) return;
    var who = t === 0
      ? (sender.name || 'Student')
      : ((m.sender && (m.sender.name || m.sender.available_name)) || 'Ardent');
    var when = m.created_at ? new Date(Number(m.created_at) * 1000).toISOString().slice(0, 16).replace('T', ' ') : '';
    var note = pics.length ? '[shared ' + pics.length + ' screenshot' + (pics.length > 1 ? 's' : '') + ']' : '';
    lines.push(who + (when ? ' (' + when + ')' : '') + ': ' + (body || note) + (body && note ? '\n' + note : ''));
    pics.forEach(function (a) { attImages.push(a.data_url); });
  });

  // Copy the screenshots into the same Drive folder the form uploads use -
  // the issue then owns a stable link, not a Chatwoot URL that can expire.
  // Capped at 5 (matching the form) and a sane size; a failed copy never
  // costs the transcript. The backtest passes skip_images so a 500-conversation
  // sweep doesn't fill Drive with copies nobody asked for (r46).
  var savedImages = [];
  var cwTok = chatwootCfg_().token;
  (data.skip_images ? [] : attImages.slice(0, 5)).forEach(function (u, n) {
    try {
      var res = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: true, headers: { api_access_token: cwTok } });
      if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) return;
      var blob = res.getBlob();
      var ct = String(blob.getContentType() || '');
      if (ct.indexOf('image/') !== 0) return;
      if (blob.getBytes().length > 6 * 1024 * 1024) return;
      var up = uploadImage_({ base64: Utilities.base64Encode(blob.getBytes()), mimeType: ct,
        filename: 'chatwoot-' + id + '-' + (n + 1) });
      if (up && up.ok && up.url) savedImages.push(up.url);
    } catch (e) {}
  });

  return {
    ok: true,
    conversation_id: id,
    student_name: sender.name || '',
    student_contact: sender.email || sender.phone_number || '',
    // The contact id we just read anyway. Stable across a change of email or
    // display name, which is more than either of the two fields above manage.
    chatwoot_contact_id: sender.id ? String(sender.id) : '',
    transcript: lines.join('\n\n'),
    message_count: lines.length,
    images: savedImages,
    images_seen: attImages.length,
    link: CHATWOOT_BASE + '/app/accounts/' + chatwootCfg_().account + '/conversations/' + id
  };
}
// Recent conversations for the in-app picker.
// FB-0246/0247. Edd: "These MA/Exam Results emails are never bugs/issues which
// need to be here." They are the automated mock-exam result mails from
// quizresults.eu landing in the support inbox: a full transcript of a student's
// exam answers, never a fault report. They were filling the Chats tab, and the
// overnight scan was paying to read them. Filtered at source so the tab, the
// Browse recent picker, and the scan all agree.
// A conversation can still be opened by pasting its link or number, so nothing
// is unreachable, it is only out of the way.
// FB-0257 (Edd): "The voicemail doesn't need to show here. Only tech, course
// issues/bugs." Machine senders have a fixed address, so naming one keeps it
// out for good and costs nothing to check.
//
// The limit is worth stating: this catches MACHINES, never people. A student
// asking about a discount code arrives on an ordinary address and reads exactly
// like a bug report until somebody opens it, so no sender rule will ever filter
// that. Anything human stays visible and gets closed without filing.
function isAutomatedNotice_(name, email) {
  var e = String(email || '').trim().toLowerCase();
  var n = String(name || '').trim().toLowerCase();
  if (/@quizresults\.eu$/.test(e)) return true;          // mock exam results
  if (n === 'ma/exam results') return true;
  // Catch-all for the shape rather than the domain: donotreply@, no-reply@ and
  // friends are never a student getting in touch.
  if (/^(donotreply|do-not-reply|noreply|no-reply|notifications?|mailer-daemon|postmaster)(@|$)/.test(e)) return true;
  if (/^(donotreply|do not reply|no ?reply|notifications?)$/.test(n)) return true;
  return false;
}
function chatwootList_(data) {
  var out;
  try {
    out = chatwootCall_('/conversations?status=' + encodeURIComponent(data.status || 'open') + '&page=1');
  } catch (e) { return { ok: false, error: String(e.message || e) }; }
  var payload = (out && out.data && out.data.payload) || (out && out.payload) || [];
  payload = payload.filter(function (c) {
    var sender = (c.meta && c.meta.sender) || {};
    return !isAutomatedNotice_(sender.name, sender.email);
  });
  var rows = payload.slice(0, 40).map(function (c) {
    var sender = (c.meta && c.meta.sender) || {};
    return {
      id: c.id,
      student_name: sender.name || '',
      student_contact: sender.email || '',
      status: c.status || '',
      last_at: c.last_activity_at ? new Date(Number(c.last_activity_at) * 1000).toISOString() : '',
      snippet: String((c.messages && c.messages.length && c.messages[c.messages.length - 1].content) || '').replace(/\s+/g, ' ').slice(0, 120)
    };
  });
  return { ok: true, conversations: rows };
}
// Private (internal) note back on the conversation, so Chatwoot shows the
// issue was logged and where to follow it. Never visible to the student.
function chatwootNote_(convId, issue, appUrl) {
  if (!convId) return;
  var cfg = chatwootCfg_();
  if (!cfg.token || !cfg.account) return;
  var text = 'Logged in Bugs: ' + (issue.summary || '(no summary)') +
    '\nPriority: ' + (issue.priority || '-') + (issue.lesson_code ? ' · ' + issue.lesson_code : '') +
    '\n' + issueLink_(issue, appUrl);
  try {
    chatwootCall_('/conversations/' + convId + '/messages', 'post', {
      content: text, message_type: 'outgoing', private: true
    });
  } catch (e) {}
}

// ---- nightly chat scan ----------------------------------------------------
// Finds problems students mentioned in live chat that nobody logged. Runs at
// 5am so the morning list is fresh. Two models on purpose (Edd, 25 Jul): a
// finder reads everything, then a DIFFERENT model re-checks each candidate
// before a human ever sees it. Different models fail differently, so the
// second opinion kills most false positives, and it only runs on the few
// things the first one flagged.
var SCAN_SHEET = 'Chat Scan';
// outcome / outcome_note APPENDED (Edd, 19 Aug 2026): the new-issue path read
// the fault and ignored how the conversation ended, so a problem the student
// had already solved themselves was queued exactly like one still hurting.
var SCAN_HEADERS = ['conversation_id', 'scanned_at', 'confidence', 'summary', 'category', 'lesson_code',
                    'student_name', 'student_contact', 'status', 'issue_id', 'reviewed_by', 'reviewed_at', 'verifier_note',
                    'kind', 'verdict', 'outcome', 'outcome_note'];
var FINDER_MODEL = 'claude-sonnet-5';
var VERIFIER_MODEL = 'claude-opus-5';

// Running tally of AI usage inside one execution, so batch jobs (the backtest,
// mainly) can report what they actually spent rather than guessing (r46).
var AI_TALLY = { calls: 0, in_tokens: 0, out_tokens: 0 };
function tallyAi_(parsed) {
  AI_TALLY.calls++;
  if (parsed && parsed.usage) {
    AI_TALLY.in_tokens += Number(parsed.usage.input_tokens) || 0;
    AI_TALLY.out_tokens += Number(parsed.usage.output_tokens) || 0;
  }
}
var SCAN_MAX_CONVERSATIONS = 60;   // per run, keeps us inside the 6-minute trigger limit
// FB-0244. How much history one press of the back sweep works through. Chatwoot
// pages at 25, so four pages is up to 100 conversations looked at and at most
// SCAN_MAX_CONVERSATIONS of them actually read. Deliberately small: this button
// spends money, and a sweep that took an hour would get pressed once and never
// trusted again. Press it repeatedly to keep walking backwards.
var BACKSWEEP_PAGES = 4;
var SCAN_BATCH = 8;                // conversations per finder call
var SCAN_MAX_SUGGESTIONS = 10;     // a bigger night than this means the prompt is wrong
var SCAN_BUDGET_MS = 4 * 60 * 1000;

function scanSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(SCAN_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SCAN_SHEET);
    sh.setFrozenRows(1);
  }
  // Keep row 1 in step with SCAN_HEADERS, so adding a column later doesn't
  // need a migration.
  sh.getRange(1, 1, 1, SCAN_HEADERS.length).setValues([SCAN_HEADERS]);
  sh.getRange(1, 1, 1, SCAN_HEADERS.length).setFontWeight('bold');
  return sh;
}
function scanRows_() {
  var sh = scanSheet_();
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0];
  return values.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    var o = {}; head.forEach(function (h, i) { o[h] = r[i]; }); return o;
  });
}
// Like anthropicJson_ but says WHY it failed instead of a bare null, and
// forgives the two classic JSON sins - a chatty preamble before the object,
// and a truncated tail - by pulling out the outermost {...} it can find.
// Written for the course review (Round 44), where the plain helper's silent
// null made a real fault look like "try again in a minute" forever.
// ---- Prompt caching (FB-0239) --------------------------------------------
// The extraction prompt is roughly 35,000 tokens of course structure and field
// definitions, and it was being read from scratch on every single call. That is
// where the wait came from: a short student transcript still took nearly half a
// minute, because the model was re-reading the whole syllabus first. Marking the
// static half cacheable means the API keeps it for five minutes and later calls
// skip it, so a run of reports costs the full read once instead of once each.
//
// The variable half must come SECOND, because a cache only ever matches from the
// start of the message: put the transcript first and every call is a fresh one.
//
// If the API ever refuses the cache field the plain prompt goes straight back
// out, and it says so in the response rather than quietly halving in speed or
// failing (a fallback that hides itself is the trap from 6 Aug).
function anthropicCachedFetch_(model, staticPart, variablePart, maxTokens) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { res: null, why: 'no API key configured', cached: false };

  function send(useCache) {
    var content = useCache
      ? [{ type: 'text', text: staticPart, cache_control: { type: 'ephemeral' } },
         { type: 'text', text: variablePart }]
      : [{ type: 'text', text: staticPart + '\n' + variablePart }];
    return UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: model, max_tokens: maxTokens || 8192,
        messages: [{ role: 'user', content: content }] })
    });
  }

  var res;
  try { res = send(true); } catch (e) { return { res: null, why: 'request failed (' + String(e).slice(0, 120) + ')', cached: false }; }
  var code = res.getResponseCode();
  if (code === 400 && /cache/i.test(String(res.getContentText() || ''))) {
    Logger.log('Prompt caching refused by the API, falling back to the plain prompt: ' + String(res.getContentText()).slice(0, 300));
    try { res = send(false); } catch (e2) { return { res: null, why: 'request failed (' + String(e2).slice(0, 120) + ')', cached: false }; }
    return { res: res, why: '', cached: false, cache_refused: true };
  }
  return { res: res, why: '', cached: true };
}

function anthropicRaw_(model, prompt, maxTokens) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { json: null, why: 'no API key configured' };
  var res;
  try {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: model, max_tokens: maxTokens || 1500, messages: [{ role: 'user', content: prompt }] })
    });
  } catch (e) { return { json: null, why: 'request failed (' + String(e).slice(0, 120) + ')' }; }
  var code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    return { json: null, why: 'HTTP ' + code + ' from the API (' + String(res.getContentText() || '').slice(0, 160) + ')' };
  }
  var parsed; try { parsed = JSON.parse(res.getContentText()); } catch (e) { return { json: null, why: 'unreadable API envelope' }; }
  tallyAi_(parsed);
  var text = '';
  if (parsed.content) for (var i = 0; i < parsed.content.length; i++) {
    if (parsed.content[i].type === 'text') text += parsed.content[i].text;
  }
  text = text.replace(/^```json?\s*|\s*```$/g, '').trim();
  try { return { json: JSON.parse(text), why: '' }; } catch (e) {}
  // Preamble or trailing prose: take the outermost braces.
  var a = text.indexOf('{'), z = text.lastIndexOf('}');
  if (a > -1 && z > a) {
    try { return { json: JSON.parse(text.slice(a, z + 1)), why: '' }; } catch (e) {}
  }
  var cut = parsed.stop_reason === 'max_tokens';
  var blocks = (parsed.content || []).map(function (c) { return c.type + ':' + String(c.text || c.thinking || '').length; }).join(',');
  return { json: null, why: (cut ? 'the reply hit the token cap mid-JSON' : 'the reply was not valid JSON') +
    ' (blocks ' + (blocks || 'none') + '; out ' + ((parsed.usage && parsed.usage.output_tokens) || '?') + ' tokens; "' + text.slice(0, 120) + '…")' };
}

function anthropicJson_(model, prompt, maxTokens) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return null;
  var res;
  try {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: model, max_tokens: maxTokens || 1500, messages: [{ role: 'user', content: prompt }] })
    });
  } catch (e) { return null; }
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) return null;
  var parsed; try { parsed = JSON.parse(res.getContentText()); } catch (e) { return null; }
  tallyAi_(parsed);
  var text = '';
  if (parsed.content) for (var i = 0; i < parsed.content.length; i++) {
    if (parsed.content[i].type === 'text') text += parsed.content[i].text;
  }
  text = text.replace(/^```json?\s*|\s*```$/g, '').trim();
  try { return JSON.parse(text); } catch (e) { return null; }
}

// Counts from the last run, so a quiet night can be told apart from a broken
// one (and so the finder/verifier ratio can be watched as we tune).
var SCAN_STATS = {};
// opts.back turns this into the back-catalogue sweep (FB-0244). The nightly run
// only ever reads conversations that have moved since the last one, and the very
// first run set that pointer to "now", so everything from before the scan existed
// has never been read at all. A back sweep walks the list from the far end
// instead: it ignores the pointer, never moves it, and works a fixed slice of
// pages per press so one button can't run away with the AI bill. Conversations
// already in the Scans sheet are dropped by the free filter before any AI call,
// so pressing it twice costs almost nothing.
function scanChatwoot(opts) {
  opts = opts || {};
  var back = !!opts.back;
  var started = Date.now();
  SCAN_STATS = { listed: 0, candidates: 0, prepared: 0, flagged: 0, confirmed: 0, queued: 0, note: '',
                 back: back, pagesRead: 0, reachedEnd: false };
  var props = PropertiesService.getScriptProperties();
  var cfg = chatwootCfg_();
  if (!cfg.token || !cfg.account) return;
  var lastScan = Number(props.getProperty('CHATWOOT_LAST_SCAN') || 0);
  // First ever run starts clean from now (Edd's call): no back-catalogue sweep.
  // That is what the back sweep is for, so it is exempt.
  if (!back && !lastScan) { markScanned_(); return; }
  // The pointer belongs to the nightly run. A back sweep reading old history
  // must never move it, or the next night would skip a day.
  function markScanned_() { if (!back) props.setProperty('CHATWOOT_LAST_SCAN', String(Date.now())); }

  var seen = {};
  scanRows_().forEach(function (r) { seen[String(r.conversation_id)] = true; });

  // Chatwoot pages at 25. At 50-70 conversations a day a single page would
  // silently miss half of them, so page back until we reach conversations
  // older than the last scan (or hit the cap).
  var list = [];
  var firstPage = back ? Math.max(1, Number(opts.fromPage || 1)) : 1;
  var lastPage = back ? (firstPage + BACKSWEEP_PAGES - 1) : 6;
  try {
    for (var page = firstPage; page <= lastPage; page++) {
      var out = chatwootCall_('/conversations?status=all&page=' + page);
      var chunk = (out && out.data && out.data.payload) || (out && out.payload) || [];
      // An empty page on a back sweep means we have walked off the end of the
      // history, which is the signal to start again from the top next time.
      if (!chunk.length) { SCAN_STATS.reachedEnd = true; break; }
      list = list.concat(chunk);
      SCAN_STATS.pagesRead++;
      // The nightly run stops as soon as it reaches conversations older than
      // its pointer. A back sweep is reading exactly that old ground on
      // purpose, so it works its whole slice.
      if (!back) {
        var oldest = Math.min.apply(null, chunk.map(function (c) { return Number(c.last_activity_at || 0) * 1000; }));
        if (oldest <= lastScan) break;         // we've gone back far enough
      }
      if (list.length >= SCAN_MAX_CONVERSATIONS) break;
      if (Date.now() - started > SCAN_BUDGET_MS / 2) break;
    }
  } catch (e) { SCAN_STATS.note = 'list failed: ' + e; return; }
  SCAN_STATS.listed = list.length;

  // Stage 1: the free filter. Anything with no student voice, or a single
  // exchange, is not a report worth an AI call.
  var candidates = list.filter(function (c) {
    if (seen[String(c.id)]) return false;
    // Automated exam-result mail is never a fault report (FB-0246), and every
    // one of them is a long transcript, so reading them was the most expensive
    // way possible to conclude nothing.
    var sndr = (c.meta && c.meta.sender) || {};
    if (isAutomatedNotice_(sndr.name, sndr.email)) return false;
    // The date test is the nightly run's "has this moved since I last looked".
    // On a back sweep every conversation is older than the pointer by
    // definition, so applying it would throw the whole slice away.
    if (!back) {
      var last = Number(c.last_activity_at || 0) * 1000;
      if (last <= lastScan) return false;
    }
    return true;
  }).slice(0, SCAN_MAX_CONVERSATIONS);
  SCAN_STATS.candidates = candidates.length;

  var prepared = [];
  candidates.forEach(function (c) {
    if (Date.now() - started > SCAN_BUDGET_MS) return;
    var imp;
    try { imp = chatwootImport_({ conversation: String(c.id) }); } catch (e) { return; }
    if (!imp || !imp.ok || !imp.transcript) return;
    if (imp.message_count < 2) return;
    prepared.push({
      id: String(c.id), name: imp.student_name || '', contact: imp.student_contact || '',
      text: String(imp.transcript).slice(0, 2500)
    });
  });
  SCAN_STATS.prepared = prepared.length;
  if (!prepared.length) { markScanned_(); return; }

  // Who already has something open? Those conversations are followed up as
  // UPDATES rather than new issues - previously they were dropped at the
  // dedupe step, which threw away exactly the "still broken" and "all sorted
  // now" news the tracker most needs (Edd, 25 Jul).
  var openIssuesAll = getIssues_().issues.filter(function (i) {
    var s = String(i.status || 'open').toLowerCase();
    return s !== 'resolved' && s !== 'past';
  });
  var openByEmail = {};
  openIssuesAll.forEach(function (i) {
    var e = String(i.student_contact || '').trim().toLowerCase();
    if (!e) return;
    var prev = openByEmail[e];
    if (!prev || new Date(i.updated_at || i.submitted_at) > new Date(prev.updated_at || prev.submitted_at)) openByEmail[e] = i;
  });

  var updateCandidates = [], newCandidates = [];
  prepared.forEach(function (p) {
    var match = openByEmail[String(p.contact || '').trim().toLowerCase()];
    if (match) updateCandidates.push({ conv: p, issue: match });
    else newCandidates.push(p);
  });
  SCAN_STATS.updateCandidates = updateCandidates.length;

  // ---- update path: what does this conversation say about that open issue? --
  var updates = [];
  updateCandidates.slice(0, 12).forEach(function (u) {
    if (Date.now() - started > SCAN_BUDGET_MS) return;
    var p = 'A student with an OPEN issue on our sailing course platform has been in touch again. ' +
      'Decide what this conversation says about THAT issue, and nothing else.\n\n' +
      'THE OPEN ISSUE:\n' + JSON.stringify({ summary: u.issue.summary, status: u.issue.status, lesson_code: u.issue.lesson_code, logged: u.issue.submitted_at }) + '\n\n' +
      'THE NEW CONVERSATION:\n' + u.conv.text + '\n\n' +
      'Choose one verdict:\n' +
      '- "fixed": the student says it now works, or confirms the fix or workaround did the job.\n' +
      '- "still_broken": the student says it is still happening, happening again, or worse.\n' +
      '- "new_detail": genuinely new information about the same problem (another device, steps to reproduce, when it started, more students affected).\n' +
      '- "nothing_new": the conversation is about something else entirely, or adds nothing. This is the common answer - use it freely.\n\n' +
      'Return ONLY JSON: {"verdict":"fixed|still_broken|new_detail|nothing_new","note":"<one sentence of what the student actually said>"}. No prose, no fences.';
    var res = anthropicJson_(FINDER_MODEL, p, 400);
    if (res && res.verdict && res.verdict !== 'nothing_new') {
      updates.push({ conv: u.conv, issue: u.issue, verdict: res.verdict, note: res.note || '' });
    }
  });
  SCAN_STATS.updatesFlagged = updates.length;

  // Second opinion on updates too - a wrong "it's fixed" is the costly one.
  var updatesConfirmed = [];
  updates.forEach(function (u) {
    if (Date.now() - started > SCAN_BUDGET_MS) return;
    var vp = 'A first-pass AI read a support conversation and concluded it is an update on a known open issue. Disagree if it is wrong.\n\n' +
      'OPEN ISSUE: ' + JSON.stringify({ summary: u.issue.summary, status: u.issue.status }) + '\n' +
      'CLAIMED VERDICT: ' + u.verdict + ' - ' + u.note + '\n\n' +
      'CONVERSATION:\n' + u.conv.text + '\n\n' +
      'Be strict. "fixed" requires the student actually confirming it works now, not an instructor hoping so. ' +
      'If the conversation is really about a different problem, disagree.\n' +
      'Return ONLY JSON: {"agree":true or false,"verdict":"fixed|still_broken|new_detail","note":"<corrected one sentence>"}. No prose, no fences.';
    var v = anthropicJson_(VERIFIER_MODEL, vp, 400);
    if (v && v.agree === true) {
      updatesConfirmed.push({
        id: u.conv.id, name: u.conv.name, contact: u.conv.contact,
        summary: v.note || u.note, verdict: v.verdict || u.verdict, issue: u.issue
      });
    }
  });
  SCAN_STATS.updatesConfirmed = updatesConfirmed.length;

  // Auto-apply confirmed updates (Edd, FB-0155): both AIs agree this is news
  // on a known open issue, so it lands on the issue itself rather than waiting
  // in a review queue. Status never changes automatically - a "fixed" still
  // surfaces in Actions for a human to verify and resolve. Anything that fails
  // to apply stays 'suggested' so the old review path catches it.
  var updApplied = 0;
  var updAppliedList = [];
  if (updatesConfirmed.length) {
    var ush = scanSheet_();
    var unow = new Date().toISOString();
    var VLABEL = { fixed: 'Student says it now works', still_broken: 'Student says it is still happening', new_detail: 'New detail from the student' };
    updatesConfirmed.slice(0, SCAN_MAX_SUGGESTIONS).forEach(function (u) {
      var applied = false;
      try {
        var vLabel = VLABEL[u.verdict] || 'Update from the student';
        var r = addUpdate_({ issue_id: u.issue.issue_id, instructor_name: 'Overnight scan',
          summary: vLabel + ': ' + u.summary,
          raw_text: '[Overnight chat scan, Chatwoot conversation ' + u.id + ']\n' + vLabel + '.\n' + u.summary });
        applied = !!(r && r.ok);
      } catch (e) {}
      if (applied) { updApplied++; updAppliedList.push({ id: u.issue.issue_id, summary: u.summary }); }
      ush.appendRow([u.id, unow, 'high', u.summary, catOf_(u.issue), u.issue.lesson_code || '',
                     u.name, u.contact, applied ? 'logged' : 'suggested', u.issue.issue_id,
                     applied ? 'Overnight scan' : '', applied ? unow : '', '', 'update', u.verdict,
                     // The update path already carries its verdict; the outcome
                     // columns exist so every row is the same width.
                     '', '']);
    });
  }

  // The new-issue path only looks at students with nothing open.
  prepared = newCandidates;
  if (!prepared.length) {
    markScanned_();
    if (updatesConfirmed.length) scanSlack_([], updAppliedList, updatesConfirmed.length - updApplied);
    return;
  }

  // Stage 2: the finder reads batches and proposes candidates.
  var found = [];
  for (var i = 0; i < prepared.length; i += SCAN_BATCH) {
    if (Date.now() - started > SCAN_BUDGET_MS) break;
    var batch = prepared.slice(i, i + SCAN_BATCH);
    var prompt = 'You are reading support chats from Ardent Training, an online RYA sailing theory school, looking ONLY for ' +
      'problems with their course platform or course content that a student described.\n\n' +
      'Flag a conversation when a student says something is BROKEN, WRONG, or genuinely CONFUSING: a video or page that will not ' +
      'load, progress not saving, a login or access failure, an error in a lesson or quiz answer, a mark or assessment behaving wrongly.\n\n' +
      // FB-0244. Edd asked for shipping to be swept alongside tech and course
      // errors. It was never in this prompt, so a parcel that never turned up
      // could sit in Chatwoot unlogged however many times the scan ran.
      'Also flag a SHIPPING problem: a pack or book that has not arrived, arrived damaged, or is the wrong item.\n\n' +
      'Do NOT flag: sales or pricing enquiries, course recommendations, mock exam marking or feedback on a student\'s work, ' +
      'extensions and admin, general sailing questions, praise, chit-chat, or a student simply not knowing how something works ' +
      'when the answer is "here is how". If the instructor answered a question and nothing was actually broken, that is not an issue.\n\n' +
      'Most conversations are NOT issues. Returning an empty list is the common, correct answer.\n\n' +
      'CONVERSATIONS:\n' + batch.map(function (b) { return '### id ' + b.id + '\n' + b.text; }).join('\n\n') + '\n\n' +
      'Return ONLY JSON: {"issues":[{"id":"<conversation id>","summary":"<one sentence, what is broken>",' +
      '"category":"course_error|tech_issue|shipping","lesson_code":"<e.g. DS.09.12 or empty>","confidence":"high|medium|low"}]}. ' +
      'No prose, no markdown fences.';
    var res = anthropicJson_(FINDER_MODEL, prompt, 1500);
    if (!res) SCAN_STATS.note = 'finder returned nothing parseable';
    if (res && res.issues && res.issues.length) {
      res.issues.forEach(function (x) {
        var src = batch.filter(function (b) { return b.id === String(x.id); })[0];
        if (src && String(x.confidence || '').toLowerCase() !== 'low') found.push({ x: x, src: src });
      });
    }
  }
  SCAN_STATS.flagged = found.length;
  if (!found.length) { markScanned_(); return; }

  // Stage 3: a different model checks each candidate, adversarially.
  var confirmed = [];
  found.slice(0, SCAN_MAX_SUGGESTIONS * 2).forEach(function (f) {
    if (Date.now() - started > SCAN_BUDGET_MS) return;
    var vPrompt = 'A first-pass AI flagged this support conversation as containing an unreported problem with an online ' +
      'sailing course platform. Your job is to DISAGREE if it is wrong. Be strict: a genuine issue means the student described ' +
      'something broken, wrong, or seriously confusing about the platform or course content, or a parcel that has not arrived, ' +
      'arrived damaged, or was the wrong item. A question that was simply answered, ' +
      'an enquiry, marking feedback, admin, or a student misunderstanding with no underlying fault is NOT an issue.\n\n' +
      'CLAIM: ' + JSON.stringify(f.x) + '\n\nFULL CONVERSATION:\n' + f.src.text + '\n\n' +
      // Same call, one more field. How the conversation ENDED decides whether
      // anyone needs to act now, and reading only the fault meant a problem the
      // student had already solved was queued like one still hurting.
      'Also say how the conversation ENDED, for this one student:\n' +
      '- "still_broken": nobody solved it, or it was escalated and left open.\n' +
      '- "worked_around": the student got moving again, but only because someone stepped in by hand or gave them a way round it. ' +
      'The underlying fault was NOT fixed. Passwords reset by an agent, a tutor sending a missing image another way, and ' +
      '"use a different browser" all count as worked around.\n' +
      '- "self_resolved": it stopped happening on its own or through something the student changed at their end, ' +
      'like updating their device or refreshing the page. Nobody fixed our platform.\n' +
      '- "fixed": someone actually fixed the underlying fault and said so.\n' +
      'When a student was helped by hand, that is "worked_around", never "fixed".\n\n' +
      'Return ONLY JSON: {"agree": true or false, "why":"<one short sentence>", "summary":"<corrected one-sentence summary if you agree, else empty>", ' +
      '"outcome":"still_broken|worked_around|self_resolved|fixed", "outcome_note":"<one short sentence on how it ended>"}. No prose, no fences.';
    var v = anthropicJson_(VERIFIER_MODEL, vPrompt, 400);
    if (v && v.agree === true) {
      confirmed.push({
        id: f.src.id, name: f.src.name, contact: f.src.contact,
        summary: v.summary || f.x.summary || '', category: f.x.category || 'tech_issue',
        lesson_code: f.x.lesson_code || '', confidence: f.x.confidence || 'medium',
        note: v.why || '',
        outcome: String(v.outcome || 'still_broken').toLowerCase(),
        outcome_note: v.outcome_note || ''
      });
    }
  });

  // Stage 4: write the queue. Students with an open issue already went down
  // the update path above, so anything here is genuinely new. Different
  // students hitting the same fault are kept deliberately: repeat reports are
  // how faults get prioritised.
  var fresh = confirmed.slice(0, SCAN_MAX_SUGGESTIONS);
  SCAN_STATS.confirmed = confirmed.length;
  SCAN_STATS.queued = fresh.length;

  // Auto-log confirmed new issues (Edd, FB-0155): every one of these was
  // reaching the queue and being logged by hand anyway, so log them now. The
  // full transcript goes through the same extraction and addIssue path the
  // form uses, which keeps the dedupe, the Slack ping on high priority, and
  // the split-guard: a transcript the extraction reads as MORE than one issue
  // stays 'suggested', because the split review (Round 35) is a human step.
  var newLogged = 0;
  var newLoggedList = [];
  if (fresh.length) {
    var sh = scanSheet_();
    var now = new Date().toISOString();
    fresh.forEach(function (c) {
      var row = [c.id, now, c.confidence, c.summary, c.category, c.lesson_code,
                 c.name, c.contact, 'suggested', '', '', '', c.note, 'new', '',
                 c.outcome || '', c.outcome_note || ''];
      // Edd, 19 Aug 2026: the scan read the fault and ignored how it ended, so
      // a problem the student had already solved themselves was filed exactly
      // like one still hurting.
      //
      // still_broken and worked_around are BOTH filed. A workaround is how a
      // live fault stays invisible: two students were blocked by a broken
      // password reset link on one morning, both were handed a password by
      // hand, and nobody filed it (issue 6cee6cde). Being helped is not being
      // fixed.
      //
      // self_resolved and fixed are not filed automatically. Nothing is lost:
      // they sit in the queue with the outcome on them, so a human can log one
      // in a click if the fault is worth chasing anyway.
      var fileIt = (c.outcome !== 'self_resolved' && c.outcome !== 'fixed');
      try {
        if (fileIt && Date.now() - started <= SCAN_BUDGET_MS) {
          var imp = chatwootImport_({ conversation: c.id });
          var transcript = (imp && imp.ok && imp.transcript) ? String(imp.transcript) : '';
          if (transcript) {
            var ex = extract_({ raw_text: transcript });
            var f = ex && ex.ok ? ex.fields : null;
            var list = f ? (f.issues && f.issues.length ? f.issues : [f]) : [];
            if (list.length === 1) {
              var one = list[0];
              one.category = one.category || c.category;
              one.lesson_code = one.lesson_code || c.lesson_code;
              one.student_name = one.student_name || c.name;
              one.student_contact = one.student_contact || c.contact;
              one.summary = one.summary || c.summary;
              one.raw_text = transcript;
              if (imp.images && imp.images.length) one.image_urls = imp.images.join(',');
              one.instructor_name = 'Overnight scan';
              // Attach the things-to-try up front (Edd, FB-0185), so whoever
              // picks it out of the Scanned lane starts with steps, not a
              // blank page. Best effort - a failed call never blocks the log.
              if (String(one.category || c.category || '').toLowerCase() === 'tech_issue') {
                try {
                  var ts = troubleshoot_({ raw_text: transcript });
                  if (ts && ts.ok && ts.found && !ts.degraded) {
                    if (ts.steps && ts.steps.length) one.recommended_steps = ts.steps;
                    // The helper also reads what the transcript says was
                    // ALREADY tried - save it, or the checklist arrives blank
                    // and the next-action asks the student for things they
                    // have already sent (Edd, FB-0198: Sergei had supplied
                    // videos, device models and a whole tried-list, and the
                    // issue still said nothing had been tried).
                    if (ts.checklist) one.checklist_json = JSON.stringify(ts.checklist);
                  }
                } catch (e) {}
              }
              var add = addIssue_(one);
              var newId = add && add.ok ? ((add.issue && add.issue.issue_id) || add.issue_id || '') : '';
              if (newId) { newLogged++; newLoggedList.push({ id: newId, summary: one.summary || c.summary }); row[8] = 'logged'; row[9] = newId; row[10] = 'Overnight scan'; row[11] = now; }
            }
          }
        }
      } catch (e) {}
      sh.appendRow(row);
    });
  }
  if (fresh.length || updatesConfirmed.length) scanSlack_(newLoggedList, updAppliedList, (fresh.length - newLogged) + (updatesConfirmed.length - updApplied));
  markScanned_();
  Logger.log('scanChatwoot: ' + prepared.length + ' read, ' + found.length + ' flagged, ' + confirmed.length + ' confirmed, ' + fresh.length + ' queued.');
  // Piggyback on the nightly run: when a lesson has collected three or more
  // open student-confusion reports, draft a content-tweak suggestion for the
  // course team's queue (suggestions only, never applied - Round 45).
  try { confusionReview_(); } catch (e) {}
}

// FB-0244. The back-catalogue sweep. Walks backwards through the Chatwoot
// history a slice at a time, keeping its own page pointer so each press carries
// on from where the last one stopped. It never touches the nightly pointer, and
// anything already in the Scans sheet is dropped before an AI call, so the
// overlap between presses is free. When it walks off the end of the history it
// says so and starts again from the top next time.
function runChatBackSweep_(data) {
  var got = scanLock_();
  if (!got.lock) return { ok: false, error: got.error };
  var lock = got.lock;
  try {
  var props = PropertiesService.getScriptProperties();
  var page = Number(props.getProperty('CHATWOOT_BACKSWEEP_PAGE') || 1);
  if (data && data.restart) page = 1;
  scanChatwoot({ back: true, fromPage: page });
  var stats = SCAN_STATS || {};
  var reachedEnd = !!stats.reachedEnd;
  var nextPage = reachedEnd ? 1 : page + BACKSWEEP_PAGES;
  props.setProperty('CHATWOOT_BACKSWEEP_PAGE', String(nextPage));
  return {
    ok: true,
    from_page: page,
    next_page: nextPage,
    reached_end: reachedEnd,
    read: stats.prepared || 0,
    flagged: stats.flagged || 0,
    confirmed: stats.confirmed || 0,
    queued: chatScanList_().scans.length,
    stats: stats
  };
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

// Only one scan at a time. Both manual buttons and the nightly trigger run the
// same code, and two overlapping runs each read the "already seen" list before
// either has written its rows, so the same conversations get read and queued
// twice - paid for twice, too. Returns null when something else holds it.
function scanLock_() {
  var lock;
  try { lock = LockService.getScriptLock(); }
  catch (e) { return { error: 'Could not reach the lock service: ' + e }; }
  var got;
  try { got = lock.tryLock(20000); }
  catch (e) { return { error: 'Lock request failed: ' + e }; }
  if (!got) return { error: 'A scan is already running. Give it a minute and try again.' };
  return { lock: lock };
}
// How far back the sweep has walked, so the button can say so before it is
// pressed rather than after.
// How many issues we hold against each lesson code. Built for the Reports
// page: visits alone cannot tell a faulty lesson from a busy one, and the two
// halves of that comparison live in different systems. Counts only, so this
// carries nothing about any individual report or student.
// Open issues are counted separately because a lesson with twenty reports that
// are all resolved is a lesson that WAS bad, which is a different story.
function lessonIssueCounts_() {
  var counts = {}, open = {};
  getIssues_().issues.forEach(function (i) {
    var code = String(i.lesson_code || '').trim().toUpperCase();
    if (!/^[A-Z]{2,4}\.\d{2}\.\d{2}$/.test(code)) return;
    counts[code] = (counts[code] || 0) + 1;
    var st = String(i.status || 'open').toLowerCase();
    if (st !== 'resolved' && st !== 'resolved_tbc' && st !== 'past') open[code] = (open[code] || 0) + 1;
  });
  return { ok: true, counts: counts, open: open };
}

function chatBackSweepState_() {
  var props = PropertiesService.getScriptProperties();
  return { ok: true, next_page: Number(props.getProperty('CHATWOOT_BACKSWEEP_PAGE') || 1), pages: BACKSWEEP_PAGES };
}

// Run the scan on demand (admin button), so it can be tried without waiting
// for 5am. Clears the "start clean" pointer back a few hours so there is
// something to look at.
function runChatScan_(data) {
  var got = scanLock_();
  if (!got.lock) return { ok: false, error: got.error };
  var lock = got.lock;
  try {
  var props = PropertiesService.getScriptProperties();
  if (data && data.since_hours) {
    props.setProperty('CHATWOOT_LAST_SCAN', String(Date.now() - Number(data.since_hours) * 3600 * 1000));
  } else if (!props.getProperty('CHATWOOT_LAST_SCAN')) {
    props.setProperty('CHATWOOT_LAST_SCAN', String(Date.now() - 12 * 3600 * 1000));
  }
  scanChatwoot();
  return { ok: true, queued: chatScanList_().scans.length, stats: SCAN_STATS };
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

// One Slack message covering both halves of the scan, with a direct link per
// find - "the ping said it exists but I can't find it" (Edd, 11 Aug) ends here.
// Accepts arrays of {id, summary}; plain numbers still work as counts.
function scanSlack_(newList, updList, waiting) {
  if (!slackOn_('scan_summary')) return;
  if (typeof newList === 'number') newList = new Array(newList);
  if (typeof updList === 'number') updList = new Array(updList);
  newList = newList || []; updList = updList || [];
  if (!newList.length && !updList.length && !waiting) return;
  var app = getAppUrl_() || '';
  var lines = [':mag: *Spotted in yesterday\'s chats:*'];
  var row = function (x) {
    if (!x || !x.id) return null;
    return '• ' + String(x.summary || 'issue').slice(0, 90) + ' - ' + issueLink_({ issue_id: x.id }, app);
  };
  if (newList.length) {
    lines.push(newList.length + ' new issue' + (newList.length === 1 ? '' : 's') + ' logged automatically (in Actions under *Scanned* until someone looks in on the student):');
    newList.slice(0, 6).forEach(function (x) { var r = row(x); if (r) lines.push(r); });
  }
  if (updList.length) {
    lines.push(updList.length + ' update' + (updList.length === 1 ? '' : 's') + ' added to open issues:');
    updList.slice(0, 6).forEach(function (x) { var r = row(x); if (r) lines.push(r); });
  }
  if (waiting) lines.push(waiting + ' left for a human eye in the Actions tab.');
  slackPost_('scan_summary', lines.join('\n'));
}
// catOf for a sheet record (the frontend has its own).
function catOf_(issue) {
  var c = String((issue && issue.category) || '').toLowerCase();
  if (c === 'friction') return 'friction';
  return c === 'course_error' ? 'course_error' : (c === 'shipping' ? 'shipping' : 'tech_issue');
}

// The admin queue: list what is waiting, and record the verdict.
// One row per conversation. Two runs that overlap both build their "already
// seen" list before either writes, so the same conversation can be queued twice
// (nine conversations arrived as seventeen rows on 19 Aug). The lock below stops
// that happening again, but the queue must read cleanly whatever is in the
// sheet, including the rows already written. Newest wins.
function chatScanList_() {
  var suggested = scanRows_().filter(function (r) { return String(r.status) === 'suggested'; });
  var best = {};
  suggested.forEach(function (r) {
    var k = String(r.conversation_id);
    var prev = best[k];
    if (!prev || String(r.scanned_at || '') > String(prev.scanned_at || '')) best[k] = r;
  });
  return { ok: true, scans: Object.keys(best).map(function (k) { return best[k]; }) };
}
function chatScanReview_(data) {
  var id = String(data.conversation_id || '');
  if (!id) return { ok: false, error: 'need a conversation_id' };
  var sh = scanSheet_();
  var values = sh.getDataRange().getValues();
  var head = values[0];
  var idx = {}; head.forEach(function (h, i) { idx[h] = i; });
  // Every row for this conversation, not just the first: a duplicate left
  // behind would come straight back into the queue as if nobody had looked.
  var hit = 0;
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idx.conversation_id]) !== id) continue;
    if (String(values[r][idx.status]) !== 'suggested') continue;
    sh.getRange(r + 1, idx.status + 1).setValue(data.status === 'logged' ? 'logged' : 'dismissed');
    sh.getRange(r + 1, idx.reviewed_by + 1).setValue((data._user && data._user.name) || '');
    sh.getRange(r + 1, idx.reviewed_at + 1).setValue(new Date().toISOString());
    if (data.issue_id) sh.getRange(r + 1, idx.issue_id + 1).setValue(data.issue_id);
    hit++;
  }
  return hit ? { ok: true, rows: hit } : { ok: false, error: 'not found' };
}

// ---- self-deploy ----------------------------------------------------------
// Lets Claude push a new Code.gs without anyone opening the Apps Script
// editor: POST { action:'deployBackend', key:<DEPLOY_KEY>, source:<full file> }
// and this updates the project's code via the Apps Script API, cuts a new
// version, and repoints the existing web-app deployment at it (same /exec URL).
//
// One-time setup (per account): enable the Apps Script API at
// script.google.com/home/usersettings, add the script.projects and
// script.deployments scopes to appsscript.json, set the DEPLOY_KEY script
// property, and do one last manual deploy so all of that takes effect.
//
// Safety: gated by DEPLOY_KEY (a leak means code-deploy rights, so rotate it
// like a password); sanity checks stop a truncated or self-locking file going
// live; the manifest is never touched, only the code file.
// Read-only look at the project manifest, DEPLOY_KEY gated. Added 2 Aug while
// chasing "Permission denied while enabling APIs: drive": when appsscript.json
// lists oauthScopes by hand, Apps Script stops working them out from the code,
// so a scope missing from that list is exactly how DriveApp starts failing.
// Reading it beats guessing, and this never writes anything.
function getManifest_(data) {
  var key = PropertiesService.getScriptProperties().getProperty('DEPLOY_KEY');
  if (!key || String(data.key || '') !== key) return { ok: false, error: 'bad deploy key' };
  try {
    var res = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/' + ScriptApp.getScriptId() + '/content', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
    });
    var files = (JSON.parse(res.getContentText() || '{}').files || []);
    var manifest = files.filter(function (f) { return f.type === 'JSON'; })[0];
    return { ok: true, manifest: manifest ? manifest.source : '(no JSON file found)',
             files: files.map(function (f) { return f.name + '.' + f.type; }) };
  } catch (e) { return { ok: false, error: String(e) }; }
}

function deployBackend_(data) {
  var key = PropertiesService.getScriptProperties().getProperty('DEPLOY_KEY');
  if (!key || String(data.key || '') !== key) return { ok: false, error: 'bad deploy key' };

  var src = String(data.source || '');
  if (src.length < 50000) return { ok: false, error: 'source suspiciously short (' + src.length + ' chars) - refusing' };
  var missing = ['function doGet', 'function doPost', 'function deployBackend_', 'var HEADERS'].filter(function (must) {
    return src.indexOf(must) < 0;
  });
  if (missing.length) return { ok: false, error: 'sanity check failed, missing: ' + missing.join(', ') };

  var scriptId = ScriptApp.getScriptId();
  var base = 'https://script.googleapis.com/v1/projects/' + scriptId;
  var auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  var call = function (url, method, payload) {
    var res = UrlFetchApp.fetch(url, {
      method: method || 'get', contentType: 'application/json', headers: auth,
      payload: payload ? JSON.stringify(payload) : undefined, muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code < 200 || code >= 300) throw new Error((method || 'get') + ' ' + url.replace(base, '') + ' -> HTTP ' + code + ': ' + res.getContentText().slice(0, 300));
    return JSON.parse(res.getContentText() || '{}');
  };

  try {
    // Swap the server code, keeping every other file (manifest included) as is.
    var current = call(base + '/content');
    var files = (current.files || []).map(function (f) {
      var out = { name: f.name, type: f.type, source: f.source };
      if (f.type === 'SERVER_JS') out.source = src; // single server-file project
      return out;
    });
    call(base + '/content', 'put', { files: files });

    // New immutable version of what we just saved.
    var v = call(base + '/versions', 'post', { description: String(data.note || 'Deployed by Claude') });

    // Repoint the live web-app deployment (not @HEAD) at that version, which
    // keeps the /exec URL everyone uses.
    var deps = call(base + '/deployments');
    var target = null;
    (deps.deployments || []).forEach(function (d) {
      (d.entryPoints || []).forEach(function (ep) { if (ep.entryPointType === 'WEB_APP') target = d; });
    });
    if (!target) return { ok: false, error: 'no web app deployment found to update' };
    call(base + '/deployments/' + target.deploymentId, 'put', {
      deploymentConfig: {
        scriptId: scriptId, versionNumber: v.versionNumber,
        manifestFileName: 'appsscript',
        description: String(data.note || ('v' + v.versionNumber + ' by Claude'))
      }
    });
    // Remember what just went live, so the front end can stamp it onto feedback.
    PropertiesService.getScriptProperties().setProperties({
      BACKEND_VERSION: String(v.versionNumber),
      BACKEND_DEPLOYED_AT: new Date().toISOString(),
      BACKEND_NOTE: String(data.note || '')
    });
    return { ok: true, version: v.versionNumber, deployment: target.deploymentId };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

// Keep the checklist-review (monthly) trigger in place, and clear out the old
// daily recheck trigger if one exists (rechecks ping Slack immediately now).
// Called from setup(), safe to run repeatedly.
function ensureTriggers_() {
  var haveMonthly = false, haveTbc = false, haveBackup = false, haveDigest = false, haveScan = false, haveChase = false;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendRecheckReminders') ScriptApp.deleteTrigger(t);
    if (t.getHandlerFunction() === 'monthlyChecklistReview') haveMonthly = true;
    if (t.getHandlerFunction() === 'autoResolveTbc') haveTbc = true;
    if (t.getHandlerFunction() === 'weeklyBackup') haveBackup = true;
    if (t.getHandlerFunction() === 'weeklyDigest') haveDigest = true;
    if (t.getHandlerFunction() === 'scanChatwoot') haveScan = true;
    if (t.getHandlerFunction() === 'chaseShipping') haveChase = true;
  });
  if (!haveMonthly) {
    ScriptApp.newTrigger('monthlyChecklistReview').timeBased().onMonthDay(1).atHour(9).create();
  }
  // The TBC auto-resolve sweep existed but was never given a trigger, which is
  // why an 18-day-old TBC was still sitting there (Edd, 21 Jul).
  if (!haveTbc) {
    ScriptApp.newTrigger('autoResolveTbc').timeBased().everyDays(1).atHour(8).create();
  }
  if (!haveBackup) {
    ScriptApp.newTrigger('weeklyBackup').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  }
  if (!haveDigest) {
    ScriptApp.newTrigger('weeklyDigest').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  }
  if (!haveScan) {
    ScriptApp.newTrigger('scanChatwoot').timeBased().everyDays(1).atHour(5).create();
  }
  if (!haveChase) {
    ScriptApp.newTrigger('chaseShipping').timeBased().everyDays(1).atHour(9).create();
  }
}

// Monday-morning state of the tracker, posted to Slack so the team sees where
// things stand without opening the app. Runs on a weekly trigger (8am, after
// the 7am backup).
function weeklyDigest() {
  if (!slackOn_('weekly_digest')) return;
  var issues = getIssues_().issues || [];
  var openStates = { open: 1, in_progress: 1, with_dev: 1, dev_fixed: 1 };
  var day = 24 * 3600 * 1000, now = Date.now();
  var open = issues.filter(function (i) { return openStates[String(i.status || 'open').toLowerCase()]; });
  var high = open.filter(function (i) { return String(i.priority || '').toLowerCase() === 'high'; });
  var newWeek = issues.filter(function (i) { return now - new Date(i.submitted_at) < 7 * day; });
  var resolvedWeek = issues.filter(function (i) {
    return String(i.status || '').toLowerCase() === 'resolved' && i.resolved_at && (now - new Date(i.resolved_at) < 7 * day);
  });

  var oldest = open.slice().sort(function (a, b) { return new Date(a.submitted_at) - new Date(b.submitted_at); }).slice(0, 3);

  // Lessons carrying the most open weight (report counts included, since
  // repeat reports are how things get prioritised).
  var byLesson = {};
  open.forEach(function (i) {
    var code = String(i.lesson_code || '').trim();
    if (!code) return;
    byLesson[code] = (byLesson[code] || 0) + Math.max(1, Number(i.report_count) || 1);
  });
  var hot = Object.keys(byLesson).filter(function (c) { return byLesson[c] >= 3; })
    .sort(function (a, b) { return byLesson[b] - byLesson[a]; }).slice(0, 5);

  var appUrl = getAppUrl_();
  var lines = [
    ':newspaper: *Bugs - the week ahead*',
    '*Open:* ' + open.length + ' (' + high.length + ' high) · *New last 7 days:* ' + newWeek.length + ' · *Resolved last 7 days:* ' + resolvedWeek.length
  ];
  if (oldest.length) {
    lines.push('');
    lines.push('*Longest waiting:*');
    oldest.forEach(function (i) {
      var days = Math.floor((now - new Date(i.submitted_at)) / day);
      lines.push('• ' + (i.lesson_code ? i.lesson_code + ' - ' : '') + String(i.summary || '').slice(0, 90) + ' (' + days + ' days) ' + issueLink_(i, appUrl));
    });
  }
  if (hot.length) {
    lines.push('');
    lines.push('*Lessons with 3+ open reports:* ' + hot.map(function (c) { return c + ' (' + byLesson[c] + ')'; }).join(', '));
  }
  lines.push('');
  lines.push('Open the tracker: ' + (appUrl || '(app url not set)'));
  slackPost_('weekly_digest', lines.join('\n'));
}

// One tap of feedback on each AI extraction ("Spot on" / "Needed fixing"),
// logged so the extraction prompt can be tuned on real data as usage grows.
var RATINGS_SHEET = 'Extraction Ratings';
var RATINGS_HEADERS = ['rated_at', 'instructor', 'verdict', 'lesson_code', 'summary'];
function rateExtraction_(data) {
  var verdict = data.verdict === 'good' ? 'good' : 'bad';
  var ss = ss_();
  var sheet = ss.getSheetByName(RATINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RATINGS_SHEET);
    sheet.getRange(1, 1, 1, RATINGS_HEADERS.length).setValues([RATINGS_HEADERS]);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    new Date().toISOString(),
    (data._user && data._user.name) || '',
    verdict,
    String(data.lesson_code || '').slice(0, 40),
    String(data.summary || '').slice(0, 300)
  ]);
  return { ok: true };
}

// Weekly safety net: the Sheet IS the database, and until now nothing backed
// it up. Copies the whole spreadsheet into a "Bugs backups" Drive folder every
// Monday morning and keeps the most recent 8 copies.
function weeklyBackup() {
  var ss = ss_();
  var file = DriveApp.getFileById(ss.getId());
  var it = DriveApp.getFoldersByName('Bugs backups');
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder('Bugs backups');
  file.makeCopy('Bugs backup ' + new Date().toISOString().slice(0, 10), folder);
  // Prune to the newest 8.
  var copies = [];
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    if (f.getName().indexOf('Bugs backup ') === 0) copies.push(f);
  }
  copies.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  copies.slice(8).forEach(function (f) { f.setTrashed(true); });
  Logger.log('weeklyBackup: ' + copies.length + ' backup(s) on file.');
}

// One-off (26 Jul): move the old 'internal' CATEGORY onto the new audience
// column. Audience is derived from section rather than blanket-marking every
// old internal row, because "internal" had drifted to mean "we spotted it
// ourselves" as much as "students can't see it" - a video cutting short was
// filed internal but hits students squarely.
function migrateAudience() {
  var internalSections = { instructor_portal: 1, partner_portal: 1, other: 1 };
  var moved = 0, stamped = 0;
  ISSUE_SHEETS.forEach(function (name) {
    var sheet = sheetByName_(name);
    if (!sheet) return;
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return;
    var idx = {}; values[0].forEach(function (h, i) { idx[h] = i; });
    if (idx.audience == null) return;
    for (var r = 1; r < values.length; r++) {
      if (!values[r][idx.issue_id]) continue;
      var cat = String(values[r][idx.category] || '').toLowerCase();
      var aud = String(values[r][idx.audience] || '').toLowerCase();
      if (cat === 'internal') {
        var sec = String(values[r][idx.section] || '').toLowerCase();
        sheet.getRange(r + 1, idx.category + 1).setValue('tech_issue');
        sheet.getRange(r + 1, idx.audience + 1).setValue(internalSections[sec] ? 'internal' : 'student');
        moved++;
      } else if (!aud) {
        sheet.getRange(r + 1, idx.audience + 1).setValue('student');
        stamped++;
      }
    }
  });
  Logger.log('migrateAudience: ' + moved + ' moved off the internal category, ' + stamped + ' stamped student.');
  return { moved: moved, stamped: stamped };
}

// Run setup() remotely (DEPLOY_KEY gated), so schema/trigger changes shipped
// via deployBackend don't need anyone in the editor either.
function runSetup_(data) {
  var key = PropertiesService.getScriptProperties().getProperty('DEPLOY_KEY');
  if (!key || String(data.key || '') !== key) return { ok: false, error: 'bad deploy key' };
  setup();
  // Round 65: touching the corpus sheet is what appends its three new header
  // cells, so the post-deploy runSetup covers KnownFixes as well as Issues.
  try { knownFixesSheet_(false); } catch (e) {}
  return { ok: true };
}

// Ask the AI whether a tech issue genuinely needs a developer (a real code/bug
// fix) rather than something a user-side step would sort first. Conservative:
// only true when a code fix is genuinely required.
function aiNeedsDeveloper_(data) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return false;

  var prompt = 'A technical issue was logged for an online sailing course app and website (Ardent Training). ' +
    'Decide if it genuinely needs a DEVELOPER to fix, meaning a real bug in the code or platform: ' +
    'a broken button, page, or link; a video or audio player that will not work; a quiz that will not submit; ' +
    'progress or data not saving; something that fails no matter what the user does. ' +
    'Answer NO if it is the kind of thing usually sorted on the user side first, and which should be tried before involving a developer: ' +
    'clearing the cache, updating or reinstalling the app, logging out and back in, a browser refresh, or a known simple workaround. ' +
    'Answer NO if it works on some of the user\'s devices or networks but not others (for example fine on a phone but failing on a tablet, or a "check network connection" message), because that points to a local device or network setting rather than a code bug. ' +
    'Answer NO for anything that is not clearly a platform bug. Only answer YES when a code fix is genuinely required.\n\n' +
    'Issue:\n' + JSON.stringify({ summary: data.summary || '', raw_text: data.raw_text || '', issue_type: data.issue_type || '', device: data.device_info || '' }) + '\n\n' +
    'Return ONLY JSON: {"needs_dev": true or false}. No prose, no markdown fences.';

  var res;
  try {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 60, messages: [{ role: 'user', content: prompt }] })
    });
  } catch (e) { return false; }
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) return false;

  var parsed; try { parsed = JSON.parse(res.getContentText()); } catch (e) { return false; }
  var text = '';
  if (parsed.content) for (var i = 0; i < parsed.content.length; i++) {
    if (parsed.content[i].type === 'text') text += parsed.content[i].text;
  }
  text = text.trim().replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
  var out; try { out = JSON.parse(text); } catch (e) { return false; }
  return !!(out && out.needs_dev === true);
}

// Break a merged issue back into separate rows, one per report. Used when the
// AI has rolled together reports that were not actually the same issue. The
// first report stays on the original row; the rest become fresh open rows in
// the same sheet, each carrying its own student, priority, and text.
function splitIssue_(data) {
  var id = data.issue_id;
  if (!id) return { ok: false, error: 'splitIssue needs an issue_id' };
  var found = findRow_(id);
  if (!found) return { ok: false, error: 'No issue found with id ' + id };

  var rec = found.record;
  var reports = [];
  try { reports = rec.reports_json ? JSON.parse(rec.reports_json) : []; } catch (e) { reports = []; }
  if (reports.length <= 1) return { ok: false, error: 'This issue has only one report, nothing to split.' };

  var now = new Date().toISOString();
  var sheet = found.sheet;

  // Rebuild the original row as just the first report.
  var first = reports[0];
  rec.student_name = first.student_name || '';
  rec.student_contact = first.student_contact || '';
  rec.device_info = first.device_info || '';
  rec.summary = first.summary || rec.summary || '';
  rec.priority = (first.priority || rec.priority || 'medium');
  rec.raw_text = first.raw_text || rec.raw_text || '';
  rec.report_count = 1;
  rec.reports_json = JSON.stringify([first]);
  rec.updated_at = now;
  sheet.getRange(found.rowNum, 1, 1, HEADERS.length).setValues([recordToRow_(rec)]);

  // Each remaining report becomes its own open row in the same sheet.
  for (var i = 1; i < reports.length; i++) {
    var rep = reports[i];
    var issue = {
      issue_id: Utilities.getUuid(),
      submitted_at: rep.date || now,
      updated_at: now,
      instructor_name: rep.instructor_name || rec.instructor_name || '',
      category: rec.category,
      raw_text: rep.raw_text || rep.summary || '',
      student_name: rep.student_name || '',
      student_contact: rep.student_contact || '',
      device_info: rep.device_info || '',
      course: rec.course,
      module: rec.module,
      lesson: rec.lesson,
      lesson_code: rec.lesson_code,
      issue_type: rec.issue_type,
      summary: rep.summary || rec.summary || '',
      priority: (rep.priority || 'medium'),
      priority_reason: rec.priority_reason || '',
      image_urls: '',
      status: 'open',
      resolved_at: '',
      resolution_note: '',
      notified_students: false,
      report_count: 1,
      reports_json: JSON.stringify([rep])
    };
    sheet.appendRow(recordToRow_(issue));
  }

  return { ok: true, split_into: reports.length };
}

// ---- Image upload ---------------------------------------------------------

function uploadImage_(data) {
  var base64 = data.base64 || '';
  var comma = base64.indexOf(',');
  if (base64.indexOf('data:') === 0 && comma > -1) {
    base64 = base64.substring(comma + 1);
  }
  if (!base64) return { ok: false, error: 'uploadImage needs a base64 image' };

  var mimeType = data.mimeType || 'image/png';
  var filename = data.filename || ('issue-image-' + new Date().getTime());

  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, mimeType, filename);

  // Two ways in. DriveApp is the simple one, but it leans on the Drive API
  // being enabled in the script's Cloud project, and since about 10 July that
  // has been failing here with "Permission denied while enabling APIs: drive",
  // which quietly killed every screenshot upload in the app. So try DriveApp,
  // and if it throws, go straight at the Drive REST API with the script's own
  // OAuth token: the drive scope is granted either way, and this path doesn't
  // ask the runtime to enable anything.
  var id = '', how = 'driveapp', driveAppError = '';
  try {
    var file = DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(blob);
    // Sharing can be blocked by a Workspace policy even when the upload worked.
    // Keep the file either way; a private link is better than losing the image.
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    id = file.getId();
  } catch (e) {
    driveAppError = String(e);
    var rest = driveRestUpload_(blob, filename);
    if (!rest.ok) {
      return { ok: false, error: 'Could not save the screenshot to Drive. DriveApp said: ' + driveAppError +
        ' The direct upload then said: ' + rest.error +
        ' If both mention enabling APIs, switch the Google Drive API on for the script\'s Cloud project.' };
    }
    id = rest.id;
    how = 'rest';
  }

  return {
    ok: true,
    file_id: id,
    via: how,
    // uc?export=view is the shape we used to store and Google no longer serves
    // it as an image, so new uploads get one that renders. The front end
    // normalises whatever is already in the sheet the same way.
    url: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1600',
    open_url: 'https://drive.google.com/file/d/' + id + '/view'
  };
}

// Upload straight to the Drive REST API with the script's OAuth token. Used
// only when DriveApp has thrown. Multipart so the file and its metadata (name
// and parent folder) go up in one request.
function driveRestUpload_(blob, filename) {
  var boundary = '----aitUpload' + new Date().getTime();
  var meta = JSON.stringify({ name: filename, parents: [DRIVE_FOLDER_ID] });
  var head = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta +
    '\r\n--' + boundary + '\r\nContent-Type: ' + blob.getContentType() + '\r\n\r\n';
  var tail = '\r\n--' + boundary + '--\r\n';
  var body = Utilities.newBlob(head).getBytes()
    .concat(blob.getBytes())
    .concat(Utilities.newBlob(tail).getBytes());
  var token = ScriptApp.getOAuthToken();
  var res;
  try {
    res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id&supportsAllDrives=true', {
      method: 'post', contentType: 'multipart/related; boundary=' + boundary,
      payload: body, headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
    });
  } catch (e) { return { ok: false, error: String(e) }; }
  var code = res.getResponseCode();
  if (code < 200 || code >= 300) return { ok: false, error: 'HTTP ' + code + ': ' + res.getContentText().slice(0, 300) };
  var id = '';
  try { id = JSON.parse(res.getContentText()).id || ''; } catch (e) {}
  if (!id) return { ok: false, error: 'upload returned no file id' };
  // Anyone-with-the-link, so the image renders on a card. A Workspace policy
  // can refuse this; the file still exists, so don't fail the upload over it.
  try {
    UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '/permissions?supportsAllDrives=true', {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ role: 'reader', type: 'anyone' }),
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
    });
  } catch (e) {}
  return { ok: true, id: id };
}

// Run this once from the Apps Script editor (Run > authorizeDrive) to grant
// the Drive permission the web app needs for screenshot uploads. It just
// touches the folder so Google prompts for the missing scope.
function authorizeDrive() {
  var name = DriveApp.getFolderById(DRIVE_FOLDER_ID).getName();
  Logger.log('Drive access OK. Folder: ' + name);
}

// Attach already-uploaded image URLs to an existing issue. Used by the new
// submit flow: the issue row is saved first (fast), then screenshots upload in
// the background and land here, so a slow or failed upload can never block or
// lose the issue itself.
function attachImages_(data) {
  var id = data.issue_id;
  if (!id) return { ok: false, error: 'attachImages needs an issue_id' };
  var urls = normaliseImageUrls_(data.image_urls);
  if (!urls) return { ok: true, attached: 0 };
  var found = findRow_(id);
  if (!found) return { ok: false, error: 'No issue found with id ' + id };
  var rec = found.record;
  var existing = rec.image_urls ? String(rec.image_urls).split(',') : [];
  rec.image_urls = existing.concat(urls.split(',')).filter(Boolean).join(',');
  rec.updated_at = new Date().toISOString();
  found.sheet.getRange(found.rowNum, 1, 1, HEADERS.length).setValues([recordToRow_(rec)]);
  return { ok: true, attached: urls.split(',').length };
}

// ---- Anthropic extraction (server-side) -----------------------------------

function extract_(data) {
  var rawText = data.raw_text || '';
  if (!rawText) return { ok: false, error: 'extract needs raw_text' };

  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    return { ok: false, error: 'ANTHROPIC_API_KEY is not set in Script Properties' };
  }

  // max_tokens 8192: a single thread can split into several full issue objects,
  // so the model needs the room. 1024 truncated the JSON mid-string on long
  // multi-topic threads (exactly the ones this splitting is for) and the parse
  // then failed. A 16-message live-chat import still hit that ceiling on 27 Jul.
  var t0 = Date.now();
  var call = anthropicCachedFetch_(EXTRACTION_MODEL, extractionStaticPrompt_(), rawText + '\n"""', 8192);
  if (!call.res) return { ok: false, error: 'Anthropic call failed: ' + (call.why || 'unknown') };
  var res = call.res;

  var code = res.getResponseCode();
  var bodyText = res.getContentText();
  if (code < 200 || code >= 300) {
    return { ok: false, error: 'Anthropic API error ' + code + ': ' + bodyText };
  }

  var parsed = JSON.parse(bodyText);
  tallyAi_(parsed);
  var text = '';
  if (parsed.content && parsed.content.length) {
    for (var i = 0; i < parsed.content.length; i++) {
      if (parsed.content[i].type === 'text') text += parsed.content[i].text;
    }
  }
  text = text.trim().replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
  // Be tolerant of any stray text or fences around the JSON: keep just the
  // outermost object. This can't rescue a genuinely truncated (cut-off) reply,
  // but it stops a stray preamble or trailing note from failing the parse.
  var firstBrace = text.indexOf('{');
  var lastBrace = text.lastIndexOf('}');
  if (firstBrace > -1 && lastBrace > firstBrace) {
    text = text.slice(firstBrace, lastBrace + 1);
  }

  var fields;
  try { fields = JSON.parse(text); }
  catch (parseErr) {
    // A "max_tokens" stop means the reply was cut off mid-JSON, so no amount of
    // brace-trimming can rescue it. Say what actually happened and how to get
    // past it, rather than the opaque "could not parse". stop_reason is passed
    // back either way so a genuinely garbled (not truncated) reply is diagnosable.
    if (parsed.stop_reason === 'max_tokens') {
      return { ok: false, stop_reason: parsed.stop_reason, raw: text,
        error: 'That report was too long to read in one go, so the extraction got cut off. Try splitting it into a couple of separate pastes, or trim it down a bit.' };
    }
    return { ok: false, stop_reason: parsed.stop_reason, raw: text,
      error: 'Could not parse model output as JSON' };
  }

  // How long it took and whether the cache actually did anything, so a slow
  // extraction can be looked at rather than guessed about.
  var u = parsed.usage || {};
  return { ok: true, fields: fields, diag: {
    ms: Date.now() - t0,
    cache_read: Number(u.cache_read_input_tokens) || 0,
    cache_written: Number(u.cache_creation_input_tokens) || 0,
    input: Number(u.input_tokens) || 0,
    output: Number(u.output_tokens) || 0,
    cache_refused: !!call.cache_refused
  } };
}

// Given a new issue and a set of past RESOLVED issues (with the fix that was
// used), ask the AI whether this is a known, already-solved problem and, if so,
// recommend the fix to the instructor. Returns { found, fix, based_on }.
function suggestFix_(data) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: true, found: false };

  var newIssue = data.new_issue || {};
  var candidates = data.candidates || [];
  // Fold in the KnownFixes corpus (r46): the caller's candidates come from
  // resolved issues, but the corpus remembers fixes that predate the tracker.
  // Callers that came through fixCandidatesFor_ already have them (skip flag).
  if (!data._kf_included) {
    try {
      candidates = candidates.concat(
        fixCandidatesFor_((newIssue.summary || '') + ' ' + String(newIssue.raw_text || '').slice(0, 3000), null, true,
          data.course || newIssue.course || '')
      ).slice(0, 16);
    } catch (e) {}
  }
  if (!candidates.length) return { ok: true, found: false };

  var prompt = 'An instructor has just logged an issue with an online sailing course platform (Ardent Training). ' +
    'Below are PAST issues that were already resolved, each with the fix that was applied. ' +
    'If the new issue clearly matches one of these known, already-solved problems, write a short recommended fix for the instructor: ' +
    '1 to 3 plain, practical sentences they can act on or pass to the student. ' +
    'Base it on how the matching past issue was actually resolved.\n\n' +
    'Before answering, name to yourself the exact thing that FAILED in the new issue, and the exact thing that failed in the ' +
    'past issue you are considering. If those two things are not the same, the answer is found false, however similar the ' +
    'surrounding circumstances (same device type, same student, same lesson, both involving logging in). Two different faults ' +
    'on a tablet are still two different faults.\n\n' +
    'Hold a HIGH bar for "found". A genuine match means the SAME failure mode in the same part of the platform - ' +
    'the same thing failing in the same way, not merely the same lesson, the same device, or a similar-sounding symptom. ' +
    'Never pad a weak match into advice, and never suggest generic steps (restart, reinstall, clear cache, log out and in) ' +
    'unless the matching past issue was genuinely resolved by exactly that step - the logging form already walks instructors ' +
    'through the generic checklist, so repeating it here is noise. A wrong suggestion wastes the student\'s time and the ' +
    'instructor\'s trust; when in doubt, return found false. Most new issues do NOT have a matching past fix.\n\n' +
    // Round 65 (Edd, FB-0231). A past fix can carry a scope, because plenty of
    // them only ever applied to one course or one era of account. Edd's case:
    // a fix that "only applied to really old day skipper accounts, not
    // Yachtmaster ever" was offered on a Yachtmaster conversation. Out-of-scope
    // entries are already dropped before they get here when we know the course;
    // this rule covers the ones where we don't.
    'SCOPE. Some past entries carry "course_scope" (the course or courses the fix applies to) and "applies_when" ' +
    '(a condition: an account age, an old signup route, a particular app version). A scope is a hard limit, not a hint. ' +
    'If the new issue is outside it, the answer is found false however well the words match. ' +
    'If you cannot tell from the new issue whether it falls inside the scope, you may still recommend the fix, but you MUST ' +
    'return the condition in "applies_when" so the instructor can check it before passing anything to the student - ' +
    'and say so plainly in the fix itself. An entry marked "flagged_wrong_before" has already been reported as a wrong ' +
    'suggestion by an instructor, so hold an even higher bar for it.\n\n' +
    'NEW issue:\n' + JSON.stringify(newIssue) + '\n\n' +
    'PAST resolved issues (summary + how it was fixed, most relevant first):\n' + JSON.stringify(candidates) + '\n\n' +
    'Return ONLY JSON: {"found": true or false, "fix": "<the recommendation, or empty string>", "based_on": "<short reference to the matching past issue, or empty string>", "corpus_id": "<the corpus_id of the entry you based it on, if it had one, else empty string>", "applies_when": "<the condition the instructor must check before using this, or empty string when there is none>"}. No prose, no markdown fences.';

  var res;
  try {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 500, messages: [{ role: 'user', content: prompt }] })
    });
  } catch (e) { return { ok: true, found: false }; }
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) return { ok: true, found: false };

  var parsed; try { parsed = JSON.parse(res.getContentText()); } catch (e) { return { ok: true, found: false }; }
  tallyAi_(parsed);
  var text = '';
  if (parsed.content) for (var i = 0; i < parsed.content.length; i++) {
    if (parsed.content[i].type === 'text') text += parsed.content[i].text;
  }
  text = text.trim().replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
  var out; try { out = JSON.parse(text); } catch (e) { return { ok: true, found: false }; }

  if (!out || !out.found || !out.fix) return { ok: true, found: false };
  return { ok: true, found: true, fix: String(out.fix), based_on: String(out.based_on || ''),
    // Which corpus entry this came from, so "this suggestion was wrong here"
    // has something to point at, and the caveat that has to be checked first.
    corpus_id: String(out.corpus_id || ''), applies_when: String(out.applies_when || '') };
}

// Pull a JSON object out of a model reply. Requiring the WHOLE reply to parse
// meant one stray sentence of preamble, or a fence the regex did not match,
// threw the lot away and dropped us into the fallback (Edd, FB-0150). Take the
// outermost braces instead, which survives anything wrapped around them.
function parseModelJson_(text) {
  if (!text) return null;
  var t = String(text).trim().replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
  try { return JSON.parse(t); } catch (e) {}
  var a = t.indexOf('{'), b = t.lastIndexOf('}');
  if (a > -1 && b > a) {
    try { return JSON.parse(t.slice(a, b + 1)); } catch (e2) {}
  }
  return null;
}

// The default tech troubleshooting playbook. The live one is stored as a script
// property (PLAYBOOK_TEXT) so it can be edited from the Admin page and grow as
// we learn from resolved issues; this constant is the fallback / starting point.
var DEFAULT_PLAYBOOK = [
  'ARDENT TECH TROUBLESHOOTING PLAYBOOK (for the instructor helping a student):',
  '',
  'HOW TO USE IT: work through the steps in order, ONE AT A TIME, and confirm the result with the student before moving on to the next. Doing that keeps it clear which step actually made the difference, which is the bit the developers need.',
  '',
  'KNOWN ACCOUNT ISSUES (check these first, each has a specific fix):',
  '- Cannot purchase a course / error at purchase: they probably already have an account from a free trial. Get them to log in first (search their email), then purchase.',
  '- Login not recognising email / login problems: they likely bought through a partner school but are trying to log in at ardent-training.com. Check the instructor portal (services.ardent-training.com) and send them to the correct partner portal.',
  '- Double-check they are not misspelling their email (a screenshot confirms this).',
  '- They may be using a different email than the one they registered with: check the instructor portal.',
  '- If still stuck: get them to use "forgot password" to reset, then type email and password manually (no copy-paste).',
  '- Still cannot log in: reset their password yourself in the students tab of the instructor portal and try logging in as them. If you can log in, it is a device/browser issue on their end, move to the step-by-step list.',
  '- APP login for practical partner students: if they signed up with a social sign-in (Google/Facebook) they may have no password, and social sign-in does not work on the app for partner-school bookings. They must tap "organisation", find their school, then "forgot password" to create a password (or do it for them).',
  '- Missing RYA ebooks / e-pack: ask them to check their junk mail first, the email comes from publications@rya.org.uk. If it is genuinely not there it can be re-sent from trade.rya.org.uk, so pass it to Edd, Charly, or Charlie. Day Skipper students sometimes think the shorebased notes, exercises and Enav slides are missing from the pack when they are meant to be ebooks.',
  '',
  'BROWSER (web) ISSUES, in order, confirming each step with the student before the next:',
  '1. Refresh the page. If that does nothing, get them to do a hard refresh: Ctrl + Shift + R on Windows, Command + Shift + R on Mac.',
  '2. Try the same course, lesson and portal yourself, on your own account and device. If it fails for you too, say so ("Thanks, I have been able to replicate the issue, I am going to work through a few troubleshooting steps from my side") and work the rest of the list on your own machine first, which is far quicker than asking the student to do each one. If it works fine for you, carry on through the list with the student.',
  '3. Log out of Ardent Training and log back in.',
  '4. Try an incognito/private window. If that works, fully quit and reopen the browser; if the problem comes back, clear the cache; if that does not do it, turn off browser extensions.',
  '5. Try a different browser. If that works, update the original browser and note which one was misbehaving and what extensions they have.',
  '6. Try a different device. If that works, note the make, model and OS version of the one with the problem.',
  '7. Try a different internet connection (mobile data, a phone hotspot, or another wifi are all one and the same step). If that works, note their ISP and whether they were on a VPN, a work network, or a school network, since those often block parts of the course.',
  '8. Still stuck: gather the details listed below and log it here so it reaches the team.',
  '',
  'APP (mobile app) ISSUES, in order:',
  '1. Fully close the app and reopen it. It has to be swiped out of the open/recent apps list, not just backing out to the home screen.',
  '2. Try the same course and lesson yourself, on your own account and app. If it fails for you too, work the rest of the list from your side first and tell them you have replicated it.',
  '3. Try the web version of the same lesson. If the website works, that gets them moving again while the app problem is looked at, so say so, but log it as still open because the app fault is still there.',
  '4. Log out and back in.',
  '5. Restart the device.',
  '6. Check the app is up to date (App Store on iPhone or iPad, Google Play on Android). If there is an update, install it, reopen and try again.',
  '7. Try a different device.',
  '8. Try a different network (mobile data, a phone hotspot, or another wifi are all the same step). A "check network connection" message usually means something on that network or device is blocking the app rather than the internet being down.',
  '9. Turn off any VPN, ad blocker, or content/parental filter on the device.',
  '10. If it is about downloading lessons, check there is free space on the device.',
  '11. Read the pattern: if it works on some of their devices or networks but not others (fine on a phone, failing on tablets), it is more likely a setting on the failing device or network than a bug in our app, so rule those out before escalating.',
  '12. Still stuck after genuinely trying the above: gather the details listed below and log it here so it reaches the team.',
  '',
  'WHAT TO GATHER BEFORE IT GOES TO THE DEVELOPERS:',
  '- A screenshot or a short screen recording.',
  '- The course, module and lesson affected.',
  '- The device make and model.',
  '- The operating system and version.',
  '- The browser (or the app and its version).',
  '- The exact wording of any error message.',
  '',
  'AFTER LOGGING IT:',
  '- Tell the student it is with our developers, apologise properly, and let them know we will come back to them.',
  '- Snooze the Chatwoot conversation for 2 days with a private note holding the summary, so somebody sends them an update rather than leaving them wondering.',
  '',
  'OUR OWN SYSTEMS (the instructor portal, internal tools), where the person hitting it is one of us and there is no student:',
  '- The standard fixes still apply to you. Hard refresh, an incognito window or another browser, log out and back in, another device, another network. Staff hit stale caches and bad extensions like anybody else.',
  '- Get a colleague to load the same thing. If it fails for them too it is server-side and the developers need it now; if it is only you, it is your session, cache, or extensions, and that is worth knowing before anyone goes looking.',
  '- A 500 or another server error is ours, not yours, so log it either way. Note the exact time, what you were doing, and whether a hard refresh cleared it, because that is what narrows it down.',
  '',
  'STRAIGHT TO A BOSS, DO NOT WORK THE LIST FIRST:',
  '- A 404 "page not found" on a lesson. Edd or Stu need to know immediately; Edd is happy to be WhatsApped about this one even on his days off.',
  '- Anything that looks like it is hitting every student rather than one (a page, video host, or the site itself down or erroring for everyone). User-side steps cannot fix a server that is down.'
].join('\n');

// ---- The scope gate (Round 63, Edd FB-0226) --------------------------------
// Some faults are not "user side" or "our side" until somebody actually looks,
// and looking takes about ten seconds. A 404 on a lesson, a video that will not
// play, an error page: the instructor can open that same lesson on their own
// account and settle it there and then, and every piece of advice after that
// depends on the answer. Edd, on a 404 report: "the instructor should confirm
// the 404 by loading the same lesson first. They should confirm if this is a
// 404 affecting everyone or isolated to one person. Then the advice given here
// is only true if it affects everyone."
//
// The playbook has carried a replicate-it-yourself line since Round 37 (browser
// step 2), but it sits in the middle of a list the model reads in order, so on
// a fault like this it arrived third or not at all, and the steps above it had
// already assumed a scope nobody had established. So the gate is applied here
// rather than left to the prompt to remember.
var SCOPE_SHAPED_RE = new RegExp(
  // The boundary has to exclude letters and a leading dot, not just digits.
  // A digits-only boundary matched the milliseconds in an ISO timestamp
  // (".404Z") and the middle of a Google Drive file id ("X404W"), which is two
  // of the five it fired on across the open log before this line was tightened.
  '(^|[^0-9a-z.])404([^0-9a-z]|$)' +
  '|page (not found|cannot be found|could not be found|does ?n.?t exist)' +
  '|(error|not found) page' +
  // Bounded on purpose: a bare "500" is a price or a student count far more
  // often than it is a status code, so it only counts next to a word that makes
  // it one.
  '|\\b50[0234]\\b\\s*(error|status|response|page)|(error|status|http)\\s*(code\\s*)?50[0234]\\b' +
  '|server error' +
  '|blank (page|screen)' +
  '|(lesson|video|page|module|slide|quiz|test|assessment|exam|course|player|content|image|diagram|pdf|document)s?\\b[^.!?]{0,60}\\b(w(on|ould\\s?n).?t|will not|does ?n.?t|do ?n.?t|did ?n.?t|fail(s|ed)? to|unable to|cannot|can.?t|never)\\s+(load|play|open|start|display|show|appear|render|come up)' +
  '|(w(on|ould\\s?n).?t|does ?n.?t|will not|fail(s|ed)? to)\\s+load' +
  '|(missing|broken|dead) (page|link|lesson|video)',
  'i');
// Has anybody actually looked yet? Only an explicit statement counts, the same
// strict rule the checklist uses. "The student tried again" is not us looking.
// Deliberately strict, and the asymmetry is the point. Reading "settled" when
// nobody has looked is the exact failure this whole gate exists to stop, so it
// only counts an explicit statement that one of US loaded it. Being too strict
// costs an instructor one sentence telling them to do something they have
// already done; being too loose costs the advice its foundation.
//
// The first draft of this had a "the same thing" alternative and it swallowed
// "I tried it again and got the same thing", which is the student trying, not
// us looking. Nothing vague goes in here.
var SCOPE_SETTLED_RE = new RegExp(
  '\\b(i|we)\\s+(have\\s+|had\\s+|.ve\\s+|just\\s+)*(tried|loaded|opened|checked|tested|looked at|viewed|pulled up)\\s+(it|the|this|that|his|her|their)[^.!?]{0,60}\\b(myself|ourselves|my own|our own|on (my|our) (account|end|side|machine|device|laptop|computer))' +
  '|(works?|loads?|opens?|plays?|runs?)\\s+(fine|ok|okay|normally|perfectly|)\\s*(for|on) (me|us)\\b' +
  '|(fails?|404s?|errors?|breaks?|does ?n.?t work|does ?n.?t load|w(on|ould\\s?n).?t load)\\s+(for|on) (me|us)\\b' +
  '|\\b(i|we)\\s+(can|could|was|were)?\\s*(n.?t|not)?\\s*(replicate|reproduce)d?\\s+(it|the|this)' +
  '|been able to (replicate|reproduce)|able to (replicate|reproduce)|(i|we) (have |had )?replicated' +
  '|(everyone|all students|every student|all of them|all users) (is|are|were|is being) (affected|hit|getting)' +
  '|(it is|it.?s|this is) (happening to|affecting) (everyone|all students|every student)',
  'i');
// The gate is open only while BOTH are true: it is the shape of fault where the
// scope decides the advice, and nothing in front of us says anyone has looked.
function scopeGateOn_(text) {
  var t = String(text || '');
  if (!t) return false;
  return SCOPE_SHAPED_RE.test(t) && !SCOPE_SETTLED_RE.test(t);
}
// Step one, in the two framings. The purpose is stated because a step without
// its reason gets skipped, and this one is only worth doing for its answer.
function scopeStepText_(staff) {
  return staff
    ? 'First, settle the scope before anything else: ask someone else on the team to open the same page in the same portal, and get a colleague to try it too. We are answering one question, is this hitting everyone or only this machine? If it fails for them too it is server or content side and the developers need it now, so do not spend time on your own cache. If it loads fine for them, it is your session, cache, extensions or device, and the steps below are the ones worth working.'
    : 'First, settle the scope before advising anything: open the same course, module and lesson yourself, in the same portal, on your own account and device. We are answering one question, is this hitting everyone or just this one student? Hold the steps below until it is answered.';
}
// The branch. Everything after step one is conditional on its answer, and the
// instructor is told so in plain words rather than being handed a list that
// quietly assumes one of the two answers is already true.
function scopeBranchText_(staff, team) {
  var who = team || 'the developers';
  return staff
    ? 'If it failed for your colleague too, stop here and log it for ' + who + ' with the exact page and the time. If it was only you, work down the list below.'
    : 'If it fails for you too, it is our end, not theirs: log it for ' + who + ' now with the exact lesson and what you saw, and do not send the student round the houses clearing caches for a fault that is not on their machine. If it loads fine for you, then either it is this student\'s session, account, portal or device, or it was a blip that has since cleared, and the user-side steps below are the right ones to work through with them.';
}
// Round 63. The model is told the gate is open, and told plainly that the list
// it usually reaches for is not the answer until the first step has been taken.
function scopeGateBlock_(staff, team) {
  return 'ESTABLISH THE SCOPE FIRST. THIS BEATS EVERY ORDERING BELOW:\n' +
    'This report is the shape of fault where nobody yet knows whether it is happening to everyone or to one person, and nothing in the conversation says anybody has looked. That answer decides which advice is even true, so it is not something to get to later.\n' +
    '- YOUR FIRST STEP IS OURS, NOT THE STUDENT\'S. It is: ' + scopeStepText_(staff) + ' Say it as step one, in those terms, and say what it is for.\n' +
    '- EVERY STEP AFTER IT IS CONDITIONAL ON THE ANSWER, and must say so in the step itself. ' + scopeBranchText_(staff, team) + '\n' +
    '- Do NOT put a hard refresh, a different browser, a different network, an incognito window, clearing a cache or any other user-side step ABOVE that first step, and never give one as unconditional advice. Each of them is only worth the student\'s time once we know the fault is theirs alone.\n' +
    '- Equally, do NOT tell them to escalate it, flag it urgently or hand it on as though it were confirmed to be hitting everyone. That is just as much of an assumption, and it is the one that was made on the report this rule came from.\n' +
    '- Mark the "tried the same course, lesson and portal yourself" checklist item as todo, never na and never done, unless the conversation explicitly says somebody did it.\n\n';
}

// Read the conversation the instructor pasted, work out what has been tried,
// and suggest the next steps from the playbook (pointing out anything missed).
function troubleshoot_(data) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return fallbackSteps_(data && (data.staff === true || data.staff === 'true'));
  var raw = data.raw_text || '';
  if (!raw) return { ok: true, found: false };
  // No student on the other end, so the steps are addressed to whoever is
  // sitting there and the student-account ones drop out. That is our own
  // systems (an internal task, the instructor portal) and, since FB-0207, any
  // report an instructor logged without attaching a student to it.
  var staff = (data.staff === true || data.staff === 'true');
  // Round 63, FB-0226. Decided here rather than left to the model, because the
  // playbook line it comes from sits in the middle of a list and kept losing to
  // whatever the model read first.
  var scopeText = raw + ' ' + String(data.existing_history || '') + ' ' + String(data.summary || '');
  var gate = scopeGateOn_(scopeText);
  var gateTeam = String(data.category || '') === 'course_error' ? 'the course team' : 'the developers';

  var prompt = (staff
    ? 'You are helping an Ardent Training staff member troubleshoot a fault THEY have hit and reported themselves. There is NO student on the other end of this report: nobody has been in touch about it, and there is nobody to relay steps to. It is either one of our own systems (the instructor portal, an internal task) or a bug an instructor spotted while using the platform. Address every step to them directly ("try a hard refresh", "ask someone else on the team to load it"), and NEVER phrase anything as "get the student to", "ask the student", or "send the student" - there is no student to ask, and doing it anyway is the single thing we have been told off for. Skip anything about a student\'s account, their email address, their password, or a partner portal login, because none of it applies. Everything else still counts: staff hit stale caches, bad extensions and flaky networks like anyone else, and knowing whether a hard refresh or another browser clears it tells the developers whether it is everyone or one session. '
    : 'You are helping an Ardent Training instructor troubleshoot a student tech issue. ') +
    (gate ? '\n\n' + scopeGateBlock_(staff, gateTeam) : '') +
    'Below is the troubleshooting playbook, then the conversation or notes the instructor pasted. ' +
    'Work out what has ALREADY been tried in the conversation, then list the NEXT things the instructor should get the student to try, in the playbook order, skipping anything already done. ' +
    'Only count a step as already tried if it is EXPLICITLY described in the conversation or history. Do not assume or infer that a step was tried. In particular, the app working on another device (such as their phone) does NOT mean a different network was tried, so never list "a different network" as already tried unless the student actually says they tried one. ' +
    'Be helpful and thorough: even if a lot has been tried, there are almost always remaining steps, so list every relevant playbook step that has not been explicitly done rather than concluding nothing is left. ' +
    'Pay attention to the pattern: if it works on some of the student\'s devices or networks but not others (for example fine on a phone but failing on tablets, or a "check network connection" message while other things work), that points to a setting on the failing device or network rather than a bug in our app, so suggest ruling those out first (a different network, and turning off any VPN, ad blocker or content filter). ' +
    'Keep the list short, distinct and non-overlapping: aim for 2 to 4 genuinely different next steps and never repeat the same step in different words. Treat mobile data, a phone hotspot and a different wifi as ONE step (trying a different network); if any different network has already been tried, do NOT suggest another network step. ' +
    'Do NOT tell them to escalate to Edd, Charlie, Stu or anyone, or to message Slack; that happens automatically when an issue is high priority. Once the relevant steps have genuinely all been tried, just say to submit it so it reaches the team. ' +
    'Point out any step that seems to have been missed or done out of order. If it matches one of the known account issues, name it and give that specific fix first. ' +
    'Remember the instructor can DO things, not only relay steps: reset the student\'s password from the students tab of the instructor portal, assign a course to their account, mark an exam manually from photos of their answers, post an answer in the course live chat, re-send an ebook, or raise an invoice. When one of those resolves it faster than another student-side step, make THAT the step, phrased as an action the instructor takes ("reset the password for them"), not a request routed through the student. ' +
    'Decide from the conversation whether it is a browser/web issue or a mobile app issue and use the matching list. ' +
    'Focus every step on the thing that is actually FAILING. If another route already works for them (for example the website works, or it works on another device), that is only a temporary workaround, so do NOT suggest troubleshooting the part that already works. A "different network" step means getting them to try the FAILING thing (e.g. the app) on a different network (mobile data, a phone hotspot, or another wifi, which are all one and the same step), never trying the website on a different network. ' +
    'Only suggest checking free storage when the problem is about downloading or saving content, not for a login or "check network connection" problem. ' +
    'Keep each suggestion short and practical, addressed to the instructor. Do not mention filling out any external form.\n\n' +
    'PLAYBOOK:\n' + getPlaybook_() + '\n\n' +
    'CHECKLIST ITEMS (the team ticks these off before an issue reaches the developers). For each one, decide its state from the conversation:\n' +
    checklistItemsForPrompt_(staff) + '\n' +
    'State rules: "done" only if the conversation EXPLICITLY says that step was tried (same strict rule as the steps above; the same network caveat applies). ' +
    '"na" if the step is not relevant to THIS issue, for example app-only steps (app up to date, social sign-in password, free storage) on a browser/web issue, browser-only steps (cleared cache/incognito, a different browser) on a mobile-app issue, the login/account steps when it is not a login problem, or free storage when it is not a download/save problem. ' +
    'Universal steps are NEVER "na": restarting the device, logging out and back in, trying a different device, and trying a different network apply to every platform and every kind of issue, so they can only be "done" or "todo". ' +
    '"todo" if the step is relevant but has not been done yet. When unsure between done and todo, choose todo. When unsure between na and todo, choose todo.\n\n' +
    (data.existing_history ? 'EARLIER HISTORY ON THIS SAME ISSUE (already logged; treat anything here as already tried):\n"""\n' + data.existing_history + '\n"""\n\n' : '') +
    'NEW CONVERSATION / NOTES:\n"""\n' + raw + '\n"""\n\n' +
    'Return ONLY JSON: {"found": true or false, "steps": ["short next step", ...], "escalate": true or false, "note": "<one short line such as an escalation note, or empty string>", "checklist": {"<item id>": "done | na | todo", ...}}. ' +
    'Include every checklist item id in the checklist object. ' +
    'Set found false only if there is genuinely nothing useful to suggest. No prose, no markdown fences.';

  // Anything longer than a couple of thousand characters of transcript was
  // silently failing and dropping into the generic fallback, which is how an
  // instructor got told to log the student out when logging in was the problem
  // (Edd, FB-0150; Charly, FB-0149). 2000 output tokens did not cover 17
  // checklist entries plus the steps once the model started writing at length.
  // Matched to what extract already uses, and the whole call retries once.
  var attempt = function () {
    var res;
    try {
      res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 8000, messages: [{ role: 'user', content: prompt }] })
      });
    } catch (e) { return { why: 'fetch failed: ' + e }; }
    var code = res.getResponseCode();
    if (code < 200 || code >= 300) return { why: 'api ' + code + ': ' + String(res.getContentText()).slice(0, 200) };
    var parsed = parseModelJson_(res.getContentText());
    if (!parsed) return { why: 'api reply was not JSON' };
    var text = '';
    if (parsed.content) for (var i = 0; i < parsed.content.length; i++) {
      if (parsed.content[i].type === 'text') text += parsed.content[i].text;
    }
    var stop = parsed.stop_reason || '';
    var obj = parseModelJson_(text);
    if (!obj) return { why: 'model output unparseable (stop_reason ' + (stop || 'none') + ', ' + text.length + ' chars)' };
    return { obj: obj, stop: stop };
  };

  var got = attempt();
  if (!got.obj) got = attempt();          // one retry: most of these are transient
  if (!got.obj) return fallbackSteps_(staff, got.why, gate, gateTeam);
  var out = got.obj;
  var checklist = normaliseChecklist_(out.checklist, staff);
  var steps = (out && out.found && out.steps && out.steps.length) ? out.steps : stepsFromChecklist_(checklist, staff);
  if (!steps.length) return fallbackSteps_(staff, 'the helper returned no steps', gate, gateTeam);
  var note = (out && out.note) || '';
  // The prompt asks for it; this makes sure of it. A prompt rule that only holds
  // most of the time is no use on the one report where it matters, and the whole
  // point of the gate is that nothing downstream should read as settled advice
  // while the scope is still open.
  if (gate) {
    if (checklist.replicated !== 'done') checklist.replicated = 'todo';
    steps = withScopeStepFirst_(steps, staff, gateTeam);
    note = 'Everything after step one depends on the answer to step one, so hold the rest until the lesson has been loaded ' +
      (staff ? 'by someone else on the team' : 'on your own account') + '.' + (note ? ' ' + note : '');
  }
  return { ok: true, found: true, steps: steps, escalate: !!(out && out.escalate), note: note, checklist: checklist };
}

// Put step one at the top and take out anything the model wrote that says the
// same thing, so the instructor is not told to load it themselves twice, and
// drop any bare escalate-now line, which is the assumption the gate exists to
// stop. Everything else it suggested is kept and is now read as the "if it
// works for you" branch, which the branch line says out loud.
function withScopeStepFirst_(steps, staff, team) {
  var mine = scopeStepText_(staff);
  var kept = (steps || []).filter(function (s) {
    var t = String(s || '');
    if (/\b(yourself|your own account|on your end|from your side|someone else on the team|a colleague)\b/i.test(t) &&
        /\b(load|open|try|log ?in|check|see)\b/i.test(t)) return false;
    // Anything that says "hand it on now" is asserting the answer to the very
    // question step one is asking, so it goes. The live run this was measured
    // against returned "skip the usual troubleshooting list and submit it
    // straight away", which is the assumption Edd caught, worded so that a
    // narrower filter let it through. A step that carries its own "if" is
    // conditional already and is kept.
    if (/\b(escalat|flag (it|this|as)|straight away|straight to|urgent(ly)?|submit it|log it now|raise it|skip the (usual|normal|standard))/i.test(t) &&
        !/\bif (it|they|she|he|you)\b/i.test(t)) return false;
    return true;
  });
  return [mine, scopeBranchText_(staff, team)].concat(kept).slice(0, 5);
}

// The instructor should always come away with something to try where anything
// is genuinely still outstanding (Edd, FB-0143). The model sometimes returns
// found:false while its own checklist still has half a dozen items sitting at
// "todo", which left the form silent on a case where plenty was untried. So
// when that happens, build the steps straight from the todo items instead, in
// playbook order, and hand back the first few.
// What we say when the AI helper cannot answer at all: the key was missing, the
// call failed, or the reply came back as something other than JSON. That used
// to leave the form silent, which is how a completely untried login problem
// reached the developers with nothing tried (Edd, FB-0143). These are the first
// moves from the playbook and they are worth making on almost any tech issue,
// so an unhelpful helper is better than a mute one.
var FALLBACK_STEPS = [
  'Refresh the page, then a hard refresh (Ctrl+Shift+R, or Cmd+Shift+R on a Mac). On the app, swipe it fully out of the recent apps list and reopen it.',
  'Try the same course, lesson and portal yourself, on your own account and device, so we know whether it is just them.',
  'Get them to log out and back in.',
  'Try an incognito or private window, or a different browser, and see whether that gets them going.'
];
// The same idea for one of us hitting a fault on our own systems. Edd's point
// (4 Aug): staff need reminding of the simple fixes too, and a hard refresh on
// a 500 at least tells the developers whether it is everyone or just a stale
// session. What changes is who the steps are addressed to, not whether we give
// any.
var FALLBACK_STEPS_STAFF = [
  'Refresh the page, then a hard refresh (Ctrl+Shift+R, or Cmd+Shift+R on a Mac).',
  'Try it in an incognito or private window, or a different browser, to rule out your cache and extensions.',
  'Log out and back in.',
  'Ask someone else on the team to load the same thing, so we know whether it is everyone or just you.'
];
// IMPORTANT: this is generic advice that has NOT read the conversation, so it
// has to announce itself. Round 33 returned it looking exactly like a considered
// answer, and the result was an instructor being told to log a student out when
// not being able to log in was the whole problem, with everything already tried
// sitting there in the transcript (Edd, FB-0150). Better a visible "I couldn't
// read this one" than confident advice that contradicts the notes.
function fallbackSteps_(staff, why, gate, team) {
  var steps = staff ? FALLBACK_STEPS_STAFF : FALLBACK_STEPS;
  // Round 63. If the scope is open, it is open whether or not the helper
  // answered, so the generic list gets the same treatment: step one first, and
  // the branch said out loud (FB-0226).
  if (gate) steps = withScopeStepFirst_(steps, staff, team);
  return { ok: true, found: true, degraded: true, reason: why || 'the helper did not answer',
           steps: steps, escalate: false,
           note: 'These are the standard first steps, NOT read off this conversation. Check them against what has already been tried before passing any on.',
           checklist: gate ? { replicated: 'todo' } : {} };
}

function stepsFromChecklist_(checklist, staff) {
  if (!checklist) return [];
  var out = [];
  checklistItemsFor_(staff).forEach(function (it) {
    if (out.length >= 4) return;
    if (checklist[it.id] === 'todo') out.push(checklistTodoLabel_(it, staff));
  });
  return out;
}

// A compact, numbered list of the checklist items for the AI prompt, with a
// scope hint so it can mark app/browser-only steps "na" on the wrong platform.
function checklistItemsForPrompt_(staff) {
  return checklistItemsFor_(staff).map(function (it) {
    var scope = it.scope === 'app' ? ' [app only]' : it.scope === 'browser' ? ' [browser/web only]' : '';
    return '- ' + it.id + ': ' + checklistLabel_(it, staff) + scope;
  }).join('\n');
}

// Keep only known item ids and valid states, so a wobbly AI reply cannot put
// junk in the checklist. Anything missing or odd is left out (the front-end
// treats a missing item as "todo").
function normaliseChecklist_(obj, staff) {
  var clean = {};
  if (!obj || typeof obj !== 'object') return clean;
  checklistItemsFor_(staff).forEach(function (it) {
    var v = String(obj[it.id] || '').toLowerCase();
    if (v === 'done' || v === 'na' || v === 'todo') clean[it.id] = v;
  });
  return clean;
}

// ---- Playbook storage and learning ----------------------------------------

function getPlaybook_() {
  return PropertiesService.getScriptProperties().getProperty('PLAYBOOK_TEXT') || DEFAULT_PLAYBOOK;
}
function setPlaybook_(text) {
  PropertiesService.getScriptProperties().setProperty('PLAYBOOK_TEXT', String(text || ''));
}
function appendToPlaybook_(line) {
  var pb = getPlaybook_();
  if (pb.indexOf('LEARNED FIXES') === -1) {
    pb += '\n\nLEARNED FIXES (added from resolved issues, reviewed by an admin):';
  }
  pb += '\n- ' + String(line || '').replace(/^[-\s]+/, '');
  setPlaybook_(pb);
}

function getSuggestions_() {
  var raw = PropertiesService.getScriptProperties().getProperty('PLAYBOOK_SUGGESTIONS');
  if (!raw) return [];
  try { var a = JSON.parse(raw); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}
function saveSuggestions_(arr) {
  PropertiesService.getScriptProperties().setProperty('PLAYBOOK_SUGGESTIONS', JSON.stringify(arr || []));
}
function addSuggestion_(s) {
  var arr = getSuggestions_();
  arr.push(s);
  if (arr.length > 50) arr = arr.slice(arr.length - 50);
  saveSuggestions_(arr);
}

// Endpoints used by the Admin page.
function getPlaybookEndpoint_() { return { ok: true, playbook: getPlaybook_() }; }
function savePlaybook_(body) { setPlaybook_(body.playbook); return { ok: true }; }
function listPlaybookSuggestions_() { return { ok: true, suggestions: getSuggestions_() }; }
// Round 63. The playbook is never edited directly, it is Edd's document, so a
// proposed wording change has to arrive in the same approve-or-reject queue the
// learning path already feeds. Until now only the server could put something in
// there, which meant a change spotted while building had nowhere to go except a
// note in a file. Same admin permission as editing the playbook itself.
function suggestPlaybook_(body) {
  var text = String((body && body.suggestion) || '').trim();
  if (!text) return { ok: false, error: 'Nothing suggested.' };
  addSuggestion_({
    id: Utilities.getUuid(),
    issue_id: String((body && body.issue_id) || ''),
    summary: String((body && body.summary) || ''),
    suggestion: text.slice(0, 4000),
    section: String((body && body.section) || ''),
    created_at: new Date().toISOString()
  });
  return { ok: true };
}
function resolvePlaybookSuggestion_(body) {
  var arr = getSuggestions_();
  var kept = [];
  var matched = null;
  for (var i = 0; i < arr.length; i++) {
    if (arr[i].id === body.id) matched = arr[i];
    else kept.push(arr[i]);
  }
  if (!matched) return { ok: false, error: 'Suggestion not found (it may have already been actioned).' };
  if (body.approve) appendToPlaybook_(body.suggestion || matched.suggestion);
  saveSuggestions_(kept);
  return { ok: true };
}

// When a tech issue is resolved, ask the AI whether the way it was fixed teaches
// a new, reusable troubleshooting step worth adding to the playbook. If so, queue
// a one-line suggestion for an admin to approve. Conservative: usually proposes
// nothing.
function proposePlaybookUpdate_(issue) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return;
  if (String(issue.category).toLowerCase() !== 'tech_issue') return;
  if (!issue.resolution_note) return;

  var prompt = 'Here is our tech troubleshooting playbook:\n' + getPlaybook_() + '\n\n' +
    'A tech issue was just resolved.\nSummary: ' + (issue.summary || '') +
    '\nHow it was fixed: ' + (issue.resolution_note || '') +
    '\nDeveloper notes: ' + (issue.dev_notes || '') + '\n\n' +
    'If the way this was fixed teaches a NEW, reusable troubleshooting step or known fix that is NOT already covered by the playbook, propose a single concise line (in the playbook style) we could add, and say which section it belongs to. ' +
    'If the playbook already covers it, or it is a one-off not worth adding, return found false. Be conservative.\n' +
    'Return ONLY JSON: {"found": true or false, "suggestion": "<the one-line addition>", "section": "<known issues | browser | app>"}. No prose, no fences.';

  var res;
  try {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 200, messages: [{ role: 'user', content: prompt }] })
    });
  } catch (e) { return; }
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) return;

  var parsed; try { parsed = JSON.parse(res.getContentText()); } catch (e) { return; }
  var text = '';
  if (parsed.content) for (var i = 0; i < parsed.content.length; i++) {
    if (parsed.content[i].type === 'text') text += parsed.content[i].text;
  }
  text = text.trim().replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
  var out; try { out = JSON.parse(text); } catch (e) { return; }
  if (out && out.found && out.suggestion) {
    addSuggestion_({
      id: Utilities.getUuid(), issue_id: issue.issue_id, summary: issue.summary || '',
      suggestion: out.suggestion, section: out.section || '', created_at: new Date().toISOString()
    });
  }
}

// When a student has several open issues, decide which one a new conversation is
// most likely an update to (or null for a genuinely new problem).
function matchUpdate_(data) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: true, match_id: null };
  var candidates = data.candidates || [];
  if (!candidates.length) return { ok: true, match_id: null };

  var prompt = 'A student has these open issues (id and summary):\n' + JSON.stringify(candidates) + '\n\n' +
    'Here is a NEW conversation about a problem they are having now:\n"""\n' + (data.raw_text || '') + '\n"""\n\n' +
    'Which existing issue is this new conversation most likely a continuation or update of? ' +
    'If it is clearly a different, separate problem, return null.\n' +
    'Return ONLY JSON: {"match_id": "<id of the matching issue, or null>"}. No prose, no markdown fences.';

  var res;
  try {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 80, messages: [{ role: 'user', content: prompt }] })
    });
  } catch (e) { return { ok: true, match_id: null }; }
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) return { ok: true, match_id: null };

  var parsed; try { parsed = JSON.parse(res.getContentText()); } catch (e) { return { ok: true, match_id: null }; }
  var text = '';
  if (parsed.content) for (var i = 0; i < parsed.content.length; i++) {
    if (parsed.content[i].type === 'text') text += parsed.content[i].text;
  }
  text = text.trim().replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
  var out; try { out = JSON.parse(text); } catch (e) { return { ok: true, match_id: null }; }
  var mid = out && out.match_id;
  if (!mid || mid === 'null') return { ok: true, match_id: null };
  for (var c = 0; c < candidates.length; c++) if (candidates[c].id === mid) return { ok: true, match_id: mid };
  return { ok: true, match_id: null };
}

// ---- Feedback on the tracker itself ---------------------------------------

// The next FB-#### in the sequence. Read off the highest number already in the
// sheet rather than the row count, so deleting a row never hands the same
// reference out twice.
function nextFeedbackRef_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (!values.length) return 'FB-0001';
  var col = values[0].indexOf('ref');
  var top = 0;
  if (col >= 0) {
    for (var r = 1; r < values.length; r++) {
      var m = /(\d+)/.exec(String(values[r][col] || ''));
      if (m) top = Math.max(top, parseInt(m[1], 10));
    }
  }
  // First run after this went in: start the count above the rows already there,
  // so the old unreferenced feedback and the new numbering can't collide.
  top = Math.max(top, values.length - 1);
  var n = String(top + 1);
  while (n.length < 4) n = '0' + n;
  return 'FB-' + n;
}

var FEEDBACK_KINDS = ['bug', 'idea', 'question'];

function addFeedback_(data) {
  var sheet = sheetByName_(FEEDBACK_SHEET);
  if (!sheet) return { ok: false, error: 'Feedback sheet missing. Run setup() once.' };
  if (!data.message && !data.image_urls) return { ok: false, error: 'Add a message first.' };
  ensureFeedbackHeaders_(sheet);
  var u = data._user || {};
  // Nobody should have to classify their own bug report before they can send
  // it, so the box no longer asks and this works it out from what they wrote
  // (Edd, FB-0136). An explicit value still wins if one ever gets sent.
  var judged = (data.kind || data.urgency) ? null : classifyFeedback_(data.message);
  var kind = String(data.kind || (judged && judged.kind) || 'bug').toLowerCase();
  if (FEEDBACK_KINDS.indexOf(kind) < 0) kind = 'bug';
  var urgency = String(data.urgency || (judged && judged.urgency) || 'normal');
  var row = {
    id: Utilities.getUuid(),
    created_at: new Date().toISOString(),
    user_email: u.email || '',
    user_name: u.name || '',
    message: data.message || '',
    image_urls: normaliseImageUrls_(data.image_urls),
    status: 'new',
    context: typeof data.context === 'string' ? data.context : (data.context ? JSON.stringify(data.context) : ''),
    ref: nextFeedbackRef_(sheet),
    kind: kind,
    urgency: urgency === 'blocking' ? 'blocking' : 'normal'
  };
  sheet.appendRow(FEEDBACK_HEADERS.map(function (k) { return row[k]; }));
  // Issues ping Slack, so feedback about the tracker should too - otherwise a
  // blocked instructor waits for somebody to happen to open the Admin tab.
  try { sendFeedbackSlack_(row, getAppUrl_()); } catch (e) {}
  return { ok: true, ref: row.ref };
}

// Read the message and decide what it is and how urgent it sounds. Kept to a
// tiny prompt and a short answer because it sits in the path of pressing Send.
// Anything unreadable falls back to a plain bug at normal urgency, which is
// what the box used to default to anyway, so a bad model day costs nothing.
function classifyFeedback_(message) {
  var msg = String(message || '').replace(/\s+/g, ' ').trim();
  if (!msg) return null;
  var prompt = [
    'Classify this feedback about an internal issue-tracking tool, written by a colleague who works on it.',
    '',
    'kind:',
    '- "bug" = something is broken, wrong, or not doing what it should.',
    '- "idea" = a suggestion, improvement, or request for something new. Wording like "we could do with", "how about", "better to", "it would be nice" is an idea, not a bug.',
    '- "question" = they are asking how something works rather than reporting or proposing anything.',
    '',
    'urgency:',
    '- "blocking" needs a real signal that this person is stopped RIGHT NOW: they say they are stuck, or it is urgent, or they cannot work; they have lost work; or the whole tool is unusable (will not load, cannot log in, cannot save anything).',
    '- "normal" for everything else. One broken button or page with the rest of the tool working is normal, however annoying, and so is anything reported calmly alongside other suggestions. This flag sorts a report to the top of the queue, so it has to mean something. When in doubt use normal.',
    '',
    'Return ONLY {"kind":"...","urgency":"..."} with no other text.',
    '',
    'The feedback:',
    msg.slice(0, 1500)
  ].join('\n');
  var out = anthropicJson_(ANTHROPIC_MODEL, prompt, 60);
  if (!out || FEEDBACK_KINDS.indexOf(String(out.kind || '').toLowerCase()) < 0) return null;
  return {
    kind: String(out.kind).toLowerCase(),
    urgency: String(out.urgency || '').toLowerCase() === 'blocking' ? 'blocking' : 'normal'
  };
}

function sendFeedbackSlack_(fb, appUrl) {
  if (!slackOn_('feedback')) return;
  if (!slackWebhook_()) return;
  var blocking = fb.urgency === 'blocking';
  var kindWord = fb.kind === 'idea' ? 'Idea for the tracker' : fb.kind === 'question' ? 'Question about the tracker' : 'Bug in the tracker';
  var ctx = {};
  try { ctx = fb.context ? JSON.parse(fb.context) : {}; } catch (e) { ctx = {}; }
  var msg = String(fb.message || '').replace(/\s+/g, ' ').trim();
  var lines = [
    (blocking ? ':rotating_light: *' + kindWord + ' - blocking someone right now*' : ':speech_balloon: *' + kindWord + '*') + ' (' + fb.ref + ')',
    '*From:* ' + (fb.user_name || fb.user_email || 'someone'),
    '*What they said:* ' + (msg.length > 400 ? msg.slice(0, 400) + '…' : msg || '(screenshot only)')
  ];
  if (ctx.view) lines.push('*Where:* ' + ctx.view + (ctx.build ? ' · build ' + ctx.build : ''));
  if (ctx.errors && ctx.errors.length) lines.push('*Last error:* ' + String(ctx.errors[ctx.errors.length - 1].message || '').slice(0, 200));
  if (appUrl) lines.push('Open the Feedback tab: ' + appUrl);
  slackPost_('feedback', lines.join('\n'));
}

function getFeedback_() {
  var sheet = sheetByName_(FEEDBACK_SHEET);
  if (!sheet) return { ok: true, feedback: [] };
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return { ok: true, feedback: [] };
  var head = values[0]; var idx = {}; head.forEach(function (h, i) { idx[h] = i; });
  var out = [];
  for (var r = 1; r < values.length; r++) {
    if (!values[r][idx['id']]) continue;
    // A column the sheet hasn't grown yet reads back as empty rather than
    // throwing, so an older Feedback tab keeps working until setup() runs.
    var o = {}; FEEDBACK_HEADERS.forEach(function (k) { o[k] = idx[k] == null ? '' : values[r][idx[k]]; });
    out.push(o);
  }
  out.sort(function (a, b) { return new Date(b.created_at) - new Date(a.created_at); });
  return { ok: true, feedback: out };
}

function deleteFeedback_(data) {
  var sheet = sheetByName_(FEEDBACK_SHEET);
  if (!sheet) return { ok: false, error: 'Feedback sheet missing.' };
  var values = sheet.getDataRange().getValues();
  var head = values[0];
  var idCol = head.indexOf('id');
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === String(data.id)) {
      sheet.deleteRow(r + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Feedback not found.' };
}

function updateFeedback_(data) {
  var sheet = sheetByName_(FEEDBACK_SHEET);
  if (!sheet) return { ok: false, error: 'Feedback sheet missing.' };
  var values = sheet.getDataRange().getValues();
  var head = values[0]; var idx = {}; head.forEach(function (h, i) { idx[h] = i; });
  for (var r = 1; r < values.length; r++) {
    if (values[r][idx['id']] === data.id) {
      if (data.status) sheet.getRange(r + 1, idx['status'] + 1).setValue(data.status);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Feedback not found.' };
}

// The extraction prompt is two halves. The static half carries the entire course
// structure and every field definition, it is byte-identical on every call, and
// it is about 35,000 tokens on its own. Keeping it separate lets it be sent as a
// cached block (anthropicCachedRaw_), which is where nearly all of the wait went.
function buildExtractionPrompt_(rawText) {
  return extractionStaticPrompt_() + '\n' + rawText + '\n"""';
}
function extractionStaticPrompt_() {
  return [
    'You are an assistant helping a sailing training company extract structured information from instructor reports about course issues.',
    '',
    'The instructor has pasted raw text: an email thread, chat message, their own notes, or a mix. Extract the fields below. If a field cannot be determined, return null.',
    '',
    'Return ONLY valid JSON. No preamble, no markdown fences.',
    '',
    'Use this course structure to map references to canonical values:',
    '- Courses: "Essential Navigation", "Day Skipper", "Yachtmaster", "Fast Track", "SRC", "PPR"',
    '- Match module and lesson names loosely (e.g. "tidal streams" -> Module 9: Tidal Streams in Day Skipper)',
    '- If you can identify the lesson, also return the lesson_code (e.g. "DS.09.04")',
    '- Slide/question codes read COURSE.module.lesson.slide, with optional extra parts: "EN.06.03.09" is Essential Navigation, module 6, lesson 3, slide 9. A trailing ".K" means a knowledge check question and a trailing ".M" a module assessment question (e.g. "DS.10.19.09.2.M"). When a report contains such a code, ALWAYS fill course, module, and lesson_code from its first three parts (lesson_code = e.g. "DS.10.19"), even if the code is the only clue.',
    '- MODULE ASSESSMENTS: when the text says the problem is in a module assessment (an end-of-module test) and you know which module, look the module up in the MODULE ASSESSMENTS list below — each module has exactly ONE assessment, so naming the module pins down the lesson. Set lesson to that code with a trailing ".M" (e.g. "DS.09.12.M"; put the question number in first when the text gives one, e.g. "DS.09.12.4.M") and fill course, module, and lesson_code (e.g. "DS.09.12") from it. If the module is not in the list, return the module title with lesson and lesson_code null rather than guessing a lesson number.',
    '',
    COURSE_STRUCTURE_TEXT,
    '',
    MODULE_ASSESSMENTS_TEXT,
    '',
    'Fields:',
    '- category: "shipping" if the problem is with a physical delivery to a student: a student pack, charts, almanac or plotter that has not arrived, arrived damaged or incomplete, was never dispatched, is stuck in customs, went to the wrong address, or was returned to sender. Emails from a courier (DHL, Royal Mail, UPS, FedEv, Parcelforce) about a consignment are shipping. Note that a student who cannot ACCESS online material is NOT shipping, that is a tech issue.',
    '- courier: for shipping only, the carrier name as written (e.g. "DHL", "Royal Mail"), or null.',
    '- tracking_number: for shipping only, the consignment or tracking reference exactly as it appears (couriers quote it in every email, so look for a long alphanumeric code). Return null if none appears. This is how we join several email threads about the same parcel, so it matters more than anything else in a shipping report.',
    '- category: "course_error" if the problem is with the lesson content or teaching material itself (wrong information, a confusing or incorrect explanation, a typo in a lesson, a mislabelled diagram, a quiz answer being wrong, OR a specific lesson that will not open / shows a 404 / page not found, which usually means that lesson was not uploaded properly and the course team needs to re-upload it). "tech_issue" if the problem is with the platform, website or app generally: video or audio not playing, login or access problems, a button that does not work, progress not saving, or anything device or browser specific.',
    '- likely_internal: true if this is NOT a student-facing problem at all but an internal one: instructors talking to each other about the instructor portal, the partner portal, admin tools, or company systems, with no student blocked from learning. A conversation between staff about wrong data shown in the instructor portal is internal. A student unable to watch a video is not. Return false when in doubt.',
    '- section: which part of the platform the problem lives in: one of ["website", "instructor_portal", "partner_portal", "course_player", "app", "other"], or null if unclear. "website" is the public ardent-training.com site, "course_player" is where students take lessons, "app" is the mobile app.',
    '- platform: for tech issues, where the trouble is actually happening, one of ["browser", "app", "both"], or null if the text gives no clue. Go by the words the student uses about what they were looking at. Talk of the "website", the "site", a "web page", a "link", a "tab", a browser by name (Chrome, Safari, Firefox, Edge), or a laptop, PC or Mac means "browser". Talk of "the app", downloading or updating or reinstalling it, or the App Store or Play Store means "app". Choose "both" only when they say they tried it both ways and it failed both ways. A phone or tablet on its own is not enough, because plenty of students use a browser on a tablet, so return null unless they say which they were in.',
    '- student_name: string or null',
    '- student_contact: the student email address if it appears anywhere in the text (prefer an email over a phone number). Return null if no email or contact is found, so the instructor can be asked for it.',
    '- device_info: the students device, operating system and browser if mentioned (e.g. "iPad, iOS 17, Safari" or "Windows 11, Chrome"), or null',
    '- course: one of the canonical course names above, but ONLY when the issue is actually about that course\'s content or a student\'s access to it. If the problem is with the instructor portal, partner portal, website, or another internal system, return null even if a course happens to be mentioned in passing. Do not guess a course from weak clues.',
    '- module: module title string or null',
    '- lesson: the FULL slide/question code exactly as written when one appears in the text (e.g. "EN.06.03.09" or "DS.10.19.09.2.M", one long string, not broken down), otherwise the lesson title if known, or null',
    '- lesson_code: lesson code string (e.g. DS.09.04) or null',
    '- issue_type: one of ["bug", "content_error", "student_confusion", "access_problem", "other"]. For a SHIPPING category report use one of ["not_arrived", "damaged", "wrong_item", "not_dispatched", "customs", "returned", "wrong_address", "other"] instead.',
      '- category: use "friction" when NOTHING is broken but the design cost the student money or time - they paid without spotting a discount code box, missed a deadline because a date was buried, bought the wrong thing because two options read the same. A friction report has a working system and an avoidable loss. If something actually failed, it is not friction.',
      '- request_kind: "improvement" if the report is asking for a NEW feature, an enhancement, or an "it would be nice if" change rather than reporting something broken or wrong (this applies to both course content and the platform, for example "could we add a glossary" or "the player should remember playback speed"); otherwise "fix" for a bug, an error, or something not working or incorrect as it stands. When in doubt, choose "fix". Most reports are "fix".',
    '- media_kind: for a course_error only, which part of the lesson it concerns: "video" if it is about a video or animation, "text" if it is about written text, a diagram, or quiz wording, otherwise "other". Return null for tech_issue.',
    '- impact: for an improvement only, a rough impact rating of "low", "medium", or "high" based on how much it would benefit students. Return null for a fix.',
    '- summary: one or two plain-English sentences summarising the issue. Keep the specific detail someone would need to reproduce it: which page or view, and HOW it is reached when that matters (e.g. "opened via the three-dots menu on the Students page" rather than just "the student profile page"). If the report describes two different symptoms, name both rather than blending them into one vague sentence.',
    '- priority: one of ["high", "medium", "low"]',
    '- priority_reason: one sentence explaining the priority',
    '- resolution_status: "resolved" if the pasted conversation shows this problem was ALREADY sorted out in the chat itself (an instructor gave a definitive answer or fix, the thread says "Conversation was marked resolved by ...", or the student confirms it works now, e.g. "that solved it", "thanks, working now", "all good"). "tbc" if a fix or answer was given but the student has not yet confirmed it worked. "open" if it is still unresolved, or was only logged to hand to the developers. When in doubt, "open".',
    '- IMPORTANT exception to the above: a WORKAROUND is not a fix. If the only thing that got the student going was a way round the fault rather than a correction of it (switching browser, incognito or private mode, clearing cache or cookies, disabling extensions, reinstalling the app, switching device, switching to mobile data), return "tbc" even when the student confirms it works now. The student is unblocked but the fault is still there, and someone needs to decide whether it was a one-off or is hitting everyone. This does NOT apply to a support question that was simply answered, or to a change that genuinely corrected the cause (a corrected username, an account re-enrolled, a lesson republished, a payment taken): those stay "resolved".',
    '- resolution_note: when resolution_status is "resolved" or "tbc", one or two sentences stating the actual answer or fix that was given (what resolved it), otherwise null.',
    '- student_sorted: true when THIS student has everything they need and nobody has to go back to them, even though the fault itself is still there for the developers. Typical shape: the instructor worked around it by hand (extended the account manually because the Extend button was missing, enrolled them themselves, sent the file directly), or the student is happily using another route that works. It is about the person, not the bug, so it can be true while resolution_status is still "open". Return false when the student is still waiting on us, is still blocked, or was promised an update.',
    '- resolved_by: the staff member who resolved it or gave the answer (from "marked resolved by X", or whoever replied with the fix), or null.',
    '- resolved_at: the date the resolution happened if it can be read from the text (ISO 8601 if possible, otherwise the date as written), or null.',
    '- sub_issues: a single pasted thread can hold SEVERAL separate problems raised over time (different pages, features, slides, or topics, each fixed independently, and each possibly at a different date or already resolved). If it holds more than one, return an array with one FULL entry per distinct problem, each carrying the SAME fields as above (category, likely_internal, section, student_name, student_contact, device_info, course, module, lesson, lesson_code, issue_type, request_kind, media_kind, impact, summary, priority, priority_reason, resolution_status, resolution_note, resolved_by, resolved_at, student_sorted). Put the primary or most urgent problem in the top-level fields AND as the FIRST array entry, so the array is the complete set. If it is really one problem (or one problem with knock-on effects), return null. Never split a single problem, and never blend unrelated topics into one entry.',
    '',
    'What counts as a SEPARATE issue for sub_issues, so nothing gets dropped:',
    '- A distinct bug, error, or confusing piece of content counts even when it is only mentioned in passing inside a longer message (for example a login email field that auto-capitalises the first letter). Give it its own entry rather than absorbing it into a bigger one.',
    '- A support QUESTION the student asked and got answered is its own issue too, even though it is not a bug to fix: course access length, extensions or billing, an account or enrolment mix-up, seeing two versions of the same course, "how do I mark a lesson complete", and so on. Log each with category "tech_issue", section "other" (or "course_player" for progress and lesson behaviour), issue_type "access_problem" or "other", request_kind "fix", and resolution_status "resolved" with the answer given as the resolution_note. Do not skip these just because they were only a question.',
    '- Do still group things that are genuinely one problem: the same iOS bug showing up as a crash, then lost progress, then missing discussions is ONE issue, not three.',
    '- The same fault seen in two places is ONE issue, not two. Progress recorded differently on the app vs the website, an assessment answer rejected on one platform and accepted on the other, or the same data wrong in two views are all single sync/recording problems with one root cause: log them as one issue naming both symptoms in the summary. Only split when the problems would clearly be fixed separately (different features, different lessons, unrelated causes). When unsure whether it is one problem or several, prefer ONE.',
    '',
    'Category guidance: content_error and student_confusion are usually course_error; bug and access_problem are usually tech_issue. But judge from what the report actually describes.',
    '',
    'Priority rules. For tech issues, do NOT jump to high until the fixes RELEVANT to this case have actually been tried:',
    '- high: a factual or safety-critical content error; OR a tech issue where the student cannot get into or use the course they have paid for and has no workaround that is working right now. Being locked out IS the test. A student who cannot log in, cannot load their course, cannot reach a lesson, or cannot sit an exam is high from the moment it is reported, so do NOT hold it at medium just because the troubleshooting steps have not been worked through yet. The steps are how we go about fixing it, not a bar the student has to clear before the problem counts.',
    '- high ALSO covers an outage that hits everyone rather than one student: a page, resource, video host, or the site itself down or erroring for all users (e.g. a resource page returning an error). User-side troubleshooting cannot fix a down server, so never hold one of these at medium because steps were not tried.',
    '- medium: a real problem where the student can still get on with their course. They have a workaround that is working (the website while the app misbehaves, a different browser, another device), or the fault is a nuisance rather than a blocker.',
    '- low: minor or cosmetic, a one-off, or something very likely solved by a simple relevant step the student has not tried yet.',
    '',
    'Raw text:',
    '"""'
  ].join('\n');
}

// Full lesson-level structure, generated from Edd's "Course Structures"
// spreadsheet (18 Jul 2026). Lesson titles are included so a loose reference
// like "the lesson about anchoring equipment" can be pinned to a code, and the
// Knowledge Check / Module Assessment positions can be read straight off.
var COURSE_STRUCTURE_TEXT = [
  'COURSE STRUCTURE (each module line reads: module number and title, then its lessons as "number title"):',
  '',
  'Essential Navigation (EN):',
  'M1 Welcome Aboard (EN.01): 1 Course Player and Features; 2 Welcome to Ardent Training; 3 Your Student Pack; 4 What to expect from your EN course; 5 A day out with Cynthia',
  'M2 Cynthia\'s Boat (EN.02): 1 Cynthia\'s Boat; 2 On Deck; 3 Below Decks; 4 Cynthia\'s Boat - Twilight; 5 Knowledge Check',
  'M3 Safety at Sea (EN.03): 1 Safety at Sea; 2 Safety Briefing; 3 Personal Care; 4 Lifejackets; 5 Emergency Procedures; 6 Engine Checks; 7 Cynthia\'s Safety Brief; 8 Knowledge Check',
  'M4 Understanding a Passage Plan (EN.04): 1 Understanding a Passage Plan; 2 Why we need a Passage Plan; 3 Passage Plan Structure; 4 Cynthia\'s Passage Plan; 5 Knowledge Check',
  'M5 Checking the Weather (EN.05): 1 Checking the Weather; 2 Weather Terms; 3 Weather Foecasts; 4 Cynthia\'s Weather Forecast; 5 Knowledge Check',
  'M6 Leaving the Harbour (EN.06): 1 Leaving the Harbour; 2 IALA Bouyage and Lateral Marks; 3 Cardinal Buoys; 4 Other Buoys; 5 Cynthia\'s Departure; 6 Knowledge Check',
  'M7 Setting Course (EN.07): 1 Setting Course; 2 Direction & Relative Directions; 3 Compasses; 4 Logbook; 5 Knots; 6 Cynthia\'s Logbook; 7 Knowledge Check',
  'M8 What Can We See? (EN.08): 1 What Can We See; 2 Charts; 3 Latitude and Longitude; 4 Plotting Latitude and Longitude; 5 What is GNSS; 6 Chartplotters; 7 What can Cynthia See; 8 Knowledge Check',
  'M9 The Groats (EN.09): 1 The Groats; 2 GNSS; 3 Direction; 4 Distance; 5 Position fixing; 6 Range and Bearing; 7 Cynthia Fixes Our Position; 8 Knowledge Check',
  'M10 Going with the Flow (EN.10): 1 Going with the Flow; 2 What are Tidal Streams; 3 Using a Tidal Stream Atlas; 4 Effects of Tidal Streams; 5 Cynthia\'s Tides; 6 Knowledge Check',
  'M11 Avoiding Collision (EN.11): 1 Avoiding Colision; 2 COLREGS; 3 Risk of Colision; 4 Give Way and Stand On; 5 Day Shapes; 6 Cynthia Avoiding Collision; 7 Knowledge Check',
  'M12 Communicating with Other Vessels (EN.12): 1 Communicating With Other Vessels; 2 VHF Radio; 3 Shore Contact; 4 Cynthia and the VHF Call; 5 Knowledge Check',
  'M13 Approaching St Anthony\'s (EN.13): 1 Approaching St Anthony\'s; 2 Choosing an Anchorage; 3 Leading Lines; 4 Visual Techniques; 5 Pilotage Plan; 6 Cynthia\'s Choice; 7 Knowledge Check',
  'M14 A Picnic at Anchor (EN.14): 1 A Picnic at Anchor; 2 Tidal Heights; 3 Springs and Neaps; 4 How to Anchor; 5 At Anchor; 6 Cynthia\'s Tides; 7 Knowledge Check',
  'M15 Returning Home (EN.15): 1 Returning Home; 2 What to expect; 3 Your turn to help Cynthia on your own',
  'M16 Debrief (EN.16): 1 Debrief; 2 Congratulations; 3 The Next Step; 4 Day Skipper Theory; 5 Goodbye',
  '',
  'Day Skipper (DS):',
  'M1 Welcome To Ardent Training (DS.01): 1 Welcome to Ardent Training; 2 Who are the RYA; 3 What To Expect; 4 Opening your student pack and setting up a suitable learning environment',
  'M2 Nautical Terms (DS.02): 1 Nautical Terminology; 2 Relative Direction; 3 Knowledge Check; 4 On Deck; 5 Below the Waterline; 6 Knowledge Check; 7 Below Decks; 8 Vessel Particulars; 9 Knowledge Check; 10 Module Assessment; 11 Module Summary',
  'M3 Introduction to Navigation (DS.03): 1 Navigation; 2 Charts; 3 Knowledge Check; 4 Aids to Navigation - Part 1, Lights; 5 Knowledge Check; 6 Aids to Navigation - Part 2, Buoyage; 7 Knowledge Check; 8 Lat and Long; 9 Knowledge Check; 10 Direction; 11 Measuring Bearings; 12 Knowledge Check; 13 Distance; 14 Measuring Distance on a Chart; 15 Knowledge Check; 16 Giving Position as a Range and Bearing; 17 Knowledge Check; 18 Module Assessment; 19 Module Summary',
  'M4 Navigation Instruments (DS.04): 1 Intro; 2 Compasses; 3 Knowledge Check; 4 Transducers; 5 GNSS; 6 Chartplotters; 7 Knowledge Check; 8 Radar; 9 Knowledge Check; 10 VHF; 11 AIS; 12 Knowledge Check; 13 Log Book; 14 Apps; 15 Knowledge Check; 16 Module Assessment; 17 Module Summary',
  'M5 Position Fixing (DS.05): 1 Position Fixing; 2 GNSS-Derived Positions; 3 Knowledge Check; 4 Bearings; 5 Knowledge Check; 6 Ranges; 7 Knowledge Check; 8 Mixing Methods; 9 Knowledge Check; 10 Module Assessment; 11 Module Summary',
  'M6 Ropework (DS.06): 1 Ropework; 2 Properties of Ropes; 3 Knots; 4 Knowledge Check; 5 Module Assessment; 6 Module Summary',
  'M7 Tidal Heights (DS.07): 1 Tidal Heights; 2 The Theory of Tides; 3 Tidal Terminology; 4 Knowledge Check; 5 Tide Tables; 6 Knowledge Check; 7 Tidal Curves; 8 Knowledge Check; 9 Clearances; 10 Knowledge Check; 11 Module Assessment; 12 Module Summary',
  'M8 Anchorwork (DS.08): 1 Anchorwork; 2 Anchoring Equipment; 3 Knowledge Check; 4 Choosing an Anchorage; 5 At Anchor; 6 Knowledge Check; 7 Module Assessment; 8 Module Summary',
  'M9 Tidal Streams (DS.09): 1 Tidal Streams; 2 Finding Tidal Stream Information; 3 Knowledge Check; 4 Applying Tidal Streams; 5 Knowledge Check; 6 EP; 7 Knowledge Check; 8 Course Shaping; 9 Knowledge Check; 10 Course to Steer; 11 Knowledge Check; 12 Module Assessment; 13 Module Summary',
  'M10 IRPCS (DS.10): 1 IRPCS; 2 Lights, Shapes and Sound Signals; 3 Knowledge Check; 4 Vessel Identification - Part 1; 5 Knowledge Check; 6 Vessel Identification - Part 2; 7 Knowledge Check; 8 Vessel Identification - Part 3; 9 Knowledge Check; 10 Vessel Identification - Part 4; 11 Knowledge Check; 12 Risk of Collision; 13 Give-way and Stand-on; 14 Knowledge Check; 15 TSS and Narrow Channels; 16 Signalling; 17 Restricted Visibility; 18 Knowledge Check; 19 Module Assessment; 20 Module Summary',
  'M11 Meteorology (DS.11): 1 Meteorology; 2 Pressure Systems; 3 Knowledge Check; 4 Land and Sea Breezes; 5 Fog; 6 Knowledge Check; 7 Forecasts; 8 Knowledge Check; 9 Module Assessment; 10 Module Summary',
  'M12 Safety (DS.12): 1 Safety; 2 Personal Care; 3 Lifejackets; 4 Knowledge Check; 5 MOB; 6 MOB Equipment; 7 Knowledge Check; 8 Vessel Safety; 9 Stability; 10 Fire; 11 Knowledge Check; 12 Distress Signals; 13 Knowledge Check; 14 Helicopter Rescue; 15 Abandon Ship; 16 Safety Brief; 17 Knowledge Check; 18 Module Assessment; 19 Module Summary',
  'M13 Pilotage (DS.13): 1 Pilotage; 2 Researching pilotage; 3 Pilotage Techniques; 4 Knowledge Check; 5 Route; 6 Using our Plan; 7 Knowledge Check; 8 Module Assessment; 9 Module Summary',
  'M14 Passage Planning (DS.14): 1 Passage Planning; 2 How to Use a Passage Plan; 3 Objectives; 4 Constraints; 5 Knowledge Check; 6 Route; 7 Safety Considerations; 8 The Example; 9 Knowledge Check; 10 Module Assessment; 11 Module Summary',
  'M15 Marine Environment (DS.15): 1 Marine Environment; 2 Marine Pollution and Prevention; 3 Wildlife and Habitat Protection; 4 Knowledge Check; 5 Module Assessment; 6 Module Summary',
  'M16 Final Exam (DS.16): 1 What to Expect; 2 Common Mistakes; 3 Mock',
  'M17 Final Exam (assessment) (DS.17): 1 Final Exam; 2 Well done!',
  '',
  'Yachtmaster (YM):',
  'M1 Welcome (YM.01): 1 Course Player and Features; 2 Welcome To Ardent Training; 3 Who Are The RYA; 4 What to Expect From Your Yachtmaster Theory; 5 Opening Your Student Pack',
  'M2 Navigation Instruments (YM.02): 1 Navigation Instruments; 2 Compasses; 3 Transducers; 4 Knowledge Check; 5 GNSS; 6 Chartplotters/Apps; 7 Knowledge Check; 8 Radar; 9 Logbook; 10 Knowledge Check; 11 Module Assessment; 12 Module Summary',
  'M3 Tidal Heights (YM.03): 1 Tidal Heights; 2 Tidal Theory; 3 Tidal Terminology; 4 Knowledge Check; 5 Tide Tables; 6 Tidal Curves; 7 Knowledge Check; 8 Secondary Ports; 9 Knowledge Check; 10 Tidal Clearances; 11 Knowledge Check; 12 Module Assessment; 13 Module Summary',
  'M4 Tidal Streams (YM.04): 1 Tidal Streams; 2 Sources of Tidal Stream Information; 3 Tidal Stream Atlas; 4 Computation of Rates; 5 Knowledge Check; 6 Tidal Diamonds and Ladders; 7 Other Tidal Considerations; 8 Knowledge Check; 9 Module Assessment; 10 Module Summary',
  'M5 Position Fixing and Chartwork (YM.05): 1 Position Fixing and Chartwork; 2 Aids to Nav - Lights; 3 Aids to Nav - Buoys; 4 Knowledge Check; 5 Range & Bearing; 6 Transits; 7 Accuracy; 8 Knowledge Check; 9 EP; 10 Running Fix; 11 Knowledge Check; 12 CTS; 13 Knowledge Check; 14 Module Assessment; 15 Module Summary',
  'M6 Meteorology (YM.06): 1 Meteorology; 2 Pressure Systems; 3 Fronts within a depression; 4 Knowledge Check; 5 Weather Forecasts; 6 Effects of the land on the wind; 7 Knowledge Check; 8 Fog; 9 Monitoring Weather at Sea; 10 Knowledge Check; 11 Module Assessment; 12 Module Summary',
  'M7 IRPCS (YM.07): 1 IRPCS; 2 Lights, Shapes & Sounds; 3 Knowledge Check; 4 Vessel Identification Part 1; 5 Vessel Identification Part 2; 6 Knowledge Check; 7 Vessel Identification Part 3; 8 Vessel Identification Part 4; 9 Knowledge Check; 10 Risk of Collision; 11 Give-way and Stand-on; 12 Knowledge Check; 13 TSS and Narrow Channels; 14 Signalling; 15 Restricted Visibility; 16 Knowledge Check; 17 Module Assessment; 18 Module Summary',
  'M8 Pilotage (YM.08): 1 Pilotage; 2 Sources of Information; 3 Knowledge Check; 4 Techniques; 5 Night Pilotage & Blind Navigation; 6 Knowledge Check; 7 Route; 8 Layout; 9 Knowledge Check; 10 More Examples; 11 Module Assessment; 12 Module Summary',
  'M9 Safety (YM.09): 1 Safety; 2 Lifejackets; 3 Knowledge Check; 4 Fire; 5 Stability; 6 Adverse Weather; 7 Knowledge Check; 8 Distress Signals; 9 Helicopter Rescue; 10 Knowledge Check; 11 Abandon Ship; 12 Knowledge Check; 13 Module Assessment; 14 Module Summary',
  'M10 Passage Planning (YM.10): 1 Passage Planning; 2 Appraisal; 3 Planning; 4 Knowledge Check; 5 Execution; 6 Monitoring; 7 Knowledge Check; 8 Example; 9 Module Assessment; 10 Module Summary',
  'M11 Marine Environment (YM.11): 1 Marine Environment; 2 Marine Pollution and Prevention; 3 Wildlife and Habitat Protection; 4 Knowledge Check; 5 Module Assessment; 6 Module Summary',
  'M12 Exams (YM.12): 1 What to Expect; 2 Common Mistakes; 3 Mock',
  'M13 Chartwork and IRPCS (exam) (YM.13): 1 Chartwork and IRPCS',
  'M14 Appraisal (YM.14): 1 Appraisal',
  'M15 Passage Making (YM.15): 1 Passage Making; 2 Well Done',
  '',
  'Fast Track (FT):',
  'M1 Welcome (FT.01): 1 Course Player and Features; 2 Welcome To Ardent Training; 3 Who Are The RYA; 4 What to Expect; 5 Your Student Pack',
  'M2 Foundations of Navigation (FT.02): 1 Navigation; 2 Charts; 3 Knowledge Check; 4 Aids to Navigation - Part 1, Lights; 5 Knowledge Check; 6 Aids to Navigation - Part 2, Buoyage; 7 Knowledge Check; 8 Latitude and Longitude; 9 Knowledge Check; 10 Direction; 11 Distance; 12 Knowledge Check; 13 Module Assessment; 14 Module Summary',
  'M3 Navigation Instruments (FT.03): 1 Navigation Instruments; 2 Compasses; 3 Knowledge Check; 4 Transducers; 5 GNSS; 6 Chartplotters/Apps; 7 Knowledge Check; 8 Radar; 9 VHF Radio, AIS, and Logbook; 10 Knowledge Check; 11 Module Assessment; 12 Module Summary',
  'M4 Position Fixing (FT.04): 1 Position Fixing; 2 Range and Bearing; 3 Knowledge Check; 4 Transits; 5 Three-Point Fixes; 6 Knowledge Check; 7 Accuracy; 8 Module Assessment; 9 Module Summary',
  'M5 Tidal Heights (FT.05): 1 Tidal Heights; 2 Tidal Theory; 3 Tidal Terminology; 4 Knowledge Check; 5 Tide Tables; 6 Tidal Curves; 7 Knowledge Check; 8 Secondary Ports; 9 Knowledge Check; 10 Tidal Clearances; 11 Knowledge Check; 12 Module Assessment; 13 Module Summary',
  'M6 Anchorwork (FT.06): 1 Anchorwork; 2 Anchoring Equipment; 3 Knowledge Check; 4 Choosing an Anchorage; 5 At Anchor; 6 Knowledge Check; 7 Module Assessment; 8 Module Summary',
  'M7 Ropework (FT.07): 1 Ropework; 2 Properties of Ropes; 3 Knots; 4 Knowledge Check; 5 Module Assessment; 6 Module Summary',
  'M8 Tidal Streams (FT.08): 1 Tidal Streams; 2 Sources of Tidal Stream Information; 3 Tidal Stream Atlas; 4 Computation of Rates; 5 Knowledge Check; 6 Tidal Diamonds and Ladders; 7 Other Tidal Considerations; 8 Knowledge Check; 9 Module Assessment; 10 Module Summary',
  'M9 Chartwork (FT.09): 1 Chartwork; 2 EP; 3 Knowledge Check; 4 Running Fix; 5 Knowledge Check; 6 CTS; 7 Knowledge Check; 8 Module Assessment; 9 Module Summary',
  'M10 Meteorology (FT.10): 1 Meteorology; 2 Pressure Systems; 3 Fronts Within a Depression; 4 Knowledge Check; 5 Weather Forecasts; 6 Effects of the Land on the Wind; 7 Knowledge Check; 8 Fog; 9 Monitoring Weather At Sea; 10 Knowledge Check; 11 Module Assessment; 12 Module Summary',
  'M11 IRPCS (FT.11): 1 IRPCS; 2 Lghts and Shapes; 3 Knowledge Check; 4 Vessel Identification Part 1; 5 Knowledge Check; 6 Vessel Identification Part 2; 7 Knowledge Check; 8 Vessel Identification Part 3; 9 Knowledge Check; 10 Vessel Identification Part 4; 11 Knowledge Check; 12 Risk of Collision; 13 Give-way and Stand-on; 14 Knowledge Check; 15 TSS and Narrow Channels; 16 Signalling; 17 Restricted Visibility; 18 Knowledge Check; 19 Module Assessment; 20 Module Summary',
  'M12 Pilotage (FT.12): 1 Pilotage; 2 Sources of Information; 3 Knowledge Check; 4 Techniques; 5 Blind Navigation + night; 6 Knowledge Check; 7 Route; 8 Layout; 9 Knowledge Check; 10 More Examples; 11 Module Assessment; 12 Module Summary',
  'M13 Safety (FT.13): 1 Safety; 2 Lifejackets; 3 Knowledge Check; 4 MOB; 5 MOB Equipment; 6 Knowledge Check; 7 Fire; 8 Stability; 9 Knowledge Check; 10 Adverse Weahter; 11 Knowledge Check; 12 Distress Signals; 13 Knowledge Check; 14 Helicopter Rescue; 15 Abandon Ship; 16 Safety Brief; 17 Knowledge Check; 18 Module Assessment; 19 Module Summary',
  'M14 Passage Planning (FT.14): 1 Passage Planning; 2 Appraisal; 3 Planning; 4 Knowledge Check; 5 Execution; 6 Monitoring; 7 Knowledge Check; 8 Example; 9 Module Assessment; 10 Module Summary',
  'M15 Marine Environment (FT.15): 1 Marine Environment; 2 Marine Pollution and Prevention; 3 Wildlife and Habitat Protection; 4 Knowledge Check; 5 Module Assessment; 6 Module Summary',
  'M16 Exams (FT.16): 1 What to Expect; 2 Common Mistakes; 3 Mock',
  'M17 Final Part 1 (FT.17): 1 Chartwork and IRPCS',
  'M18 Appraisal (FT.18): 1 Appraisal',
  'M19 Final Part 2 (FT.19): 1 Passage Making; 2 Well Done',
  '',
  'Lesson titles are listed so a loose description ("the lesson about anchoring equipment") can be matched to its code — do this whenever the description clearly matches one lesson.',
  '',
  'SRC and PPR: no module/lesson structure. Extract the course name only and leave module, lesson, and lesson_code as null.',
  '',
  'If you can only identify the module (not the exact lesson), return the module title and set lesson and lesson_code to null.'
].join('\n');

// Which lesson holds each module's assessment, from Edd's course structure
// spreadsheet (18 Jul 2026). The assessment usually sits second-to-last (the
// Module Summary follows it), but not always, which is why it's listed
// explicitly. Essential Navigation has knowledge checks only, no assessments.
var MODULE_ASSESSMENTS_TEXT = [
  'MODULE ASSESSMENTS (the one assessment lesson per module — do not guess for modules not listed):',
  'Day Skipper: DS.02.10, DS.03.18, DS.04.16, DS.05.10, DS.06.05, DS.07.11, DS.08.07, DS.09.12, DS.10.19, DS.11.09, DS.12.18, DS.13.08, DS.14.10, DS.15.05',
  'Yachtmaster: YM.02.11, YM.03.12, YM.04.09, YM.05.14, YM.06.11, YM.07.17, YM.08.11, YM.09.13, YM.10.09, YM.11.05',
  'Fast Track: FT.02.13, FT.03.11, FT.04.08, FT.05.12, FT.06.07, FT.07.05, FT.08.09, FT.09.08, FT.10.11, FT.11.19, FT.12.11, FT.13.18, FT.14.09, FT.15.05',
  'Essential Navigation: none — EN modules end on a Knowledge Check, not a module assessment.'
].join('\n');

// ---- Slack ----------------------------------------------------------------

// A report can be submitted without running Extract, which leaves summary
// blank. That is fine on a card (the front end falls back to the raw text),
// but a high priority Slack ping that reads "Summary: -" tells the channel
// nothing at all and the detail is sitting right there in raw_text, so fall
// back to a trimmed slice of what the instructor actually pasted.
function slackSummary_(issue) {
  var s = String((issue && issue.summary) || '').trim();
  if (s) return s;
  var raw = String((issue && issue.raw_text) || '').replace(/\s+/g, ' ').trim();
  if (!raw) return '-';
  return (raw.length > 300 ? raw.slice(0, 300) + '…' : raw) + ' (no summary - report submitted without Extract)';
}

function sendSlack_(issue, appUrl) {
  if (!slackOn_('high_priority')) return;
  var c = String(issue.category).toLowerCase();
  var area = (c === 'tech_issue' ? 'Tech issue' : 'Course error') +
    (String(issue.audience || 'student').toLowerCase() === 'internal' ? ' · internal' : '');
  if (issue.section) area += ' · ' + String(issue.section).replace(/_/g, ' ');
  var text = [
    ':red_circle: *High priority issue logged* (' + area + ')',
    '*Lesson:* ' + (issue.lesson || '-') + ' (' + (issue.lesson_code || '-') + ')',
    '*Type:* ' + (issue.issue_type || '-'),
    '*Summary:* ' + slackSummary_(issue),
    '*Student:* ' + (issue.student_name || '-') + ' (' + (issue.student_contact || '-') + ')',
    '*Device:* ' + (issue.device_info || '-'),
    '*Logged by:* ' + (issue.instructor_name || '-'),
    '*Submitted:* ' + (issue.submitted_at || new Date().toISOString()),
    '',
    'View in Bugs: ' + issueLink_(issue, appUrl)
  ].join('\n');

  slackPost_('high_priority', text);

  return { ok: true };
}

function getAppUrl_() {
  return PropertiesService.getScriptProperties().getProperty('APP_URL') || '';
}

// Which backend is actually live. deployBackend_ records this as it ships, and
// the front end reads it on login so every feedback report says which build it
// came from. Without it, "is this fixed already?" needs guesswork against the
// deploy notes, and a tab left open since Tuesday looks the same as a fresh one.
// Bumped by hand each round, same as BUILD in index.html. The deployment
// number below is more precise but only appears from the first deploy made BY
// this code onwards (the deploy that ships a version is run by the previous
// one), so this stamp is what answers "which round is live" in the meantime.
var CODE_STAMP = 'r81 · 2026-08-20';

// ---- draft a message to the student (Edd, FB-0161) -------------------------
// The Actions "next action" line offers a draft whenever the action is any
// kind of message to the student. Written in the instructor's own voice when
// a guide is on file: the VoiceGuides sheet (tab in the same spreadsheet),
// column A = instructor name exactly as they log in, column B = the style
// guide text. No sheet, no matching row, or no API key all mean the front end
// quietly falls back to the fixed house templates, so the button always works.
function voiceGuideFor_(name) {
  if (!name) return '';
  var sh = sheetByName_('VoiceGuides');
  if (!sh) return '';
  var v = sh.getDataRange().getValues();
  var want = String(name).trim().toLowerCase();
  var wantFirst = want.split(/\s+/)[0];
  var firstHit = '';
  for (var r = 0; r < v.length; r++) {
    var got = String(v[r][0] || '').trim().toLowerCase();
    if (!got) continue;
    if (got === want) return String(v[r][1] || '');
    // "Holly" matches "Holly Vint" and vice versa - guides are stored by
    // first name, but some accounts log in with a full name.
    if (!firstHit && got.split(/\s+/)[0] === wantFirst) firstHit = String(v[r][1] || '');
  }
  return firstHit;
}

// Upsert a voice guide (used by Claude to load the team's guides in bulk).
function setVoiceGuide_(data) {
  var name = String(data.name || '').trim();
  var guide = String(data.guide || '');
  if (!name || !guide) return { ok: false, error: 'need name and guide' };
  var sh = sheetByName_('VoiceGuides');
  if (!sh) { sh = ss_().insertSheet('VoiceGuides'); sh.appendRow(['name', 'guide']); }
  var v = sh.getDataRange().getValues();
  for (var r = 0; r < v.length; r++) {
    if (String(v[r][0] || '').trim().toLowerCase() === name.toLowerCase()) {
      sh.getRange(r + 1, 2).setValue(guide);
      return { ok: true, updated: name };
    }
  }
  sh.appendRow([name, guide]);
  return { ok: true, added: name };
}
function listVoiceGuides_() {
  var sh = sheetByName_('VoiceGuides');
  if (!sh) return { ok: true, guides: [] };
  return { ok: true, guides: sh.getDataRange().getValues().slice(1).map(function (r) {
    return { name: String(r[0] || ''), chars: String(r[1] || '').length };
  }).filter(function (g) { return g.name; }) };
}

// Edd's universal strip-list (11 Aug): no AI tells in anything a student
// reads, whoever the instructor. One constant so every draft prompt carries
// the same rules - r51.2 put them on the batch drafts only; r52 shares them
// with the single-draft and live-case prompts too.
var AI_TELLS_RULES_ = '\nSTYLE - strip the AI tells, every draft, whoever the instructor (Edd, 11 Aug):\n' +
  '- NO em dashes, ever. Commas, full stops, or brackets instead.\n' +
  '- No "isn\'t just X" / "more than just" setup-then-negate constructions. Direct, affirmative statements.\n' +
  '- No throat-clearing openers, no "delve", no "whether you\'re a beginner or an expert" hedges.\n' +
  '- No relentless rule-of-three rhythm (X, Y, and Z in every sentence).\n' +
  '- Never announce that something "matters" - real importance shows in the writing.\n' +
  '- No forward pointers or snappy aphoristic closers that add no information. If the point is made, stop.\n' +
  '- No hollow superlatives (seamless, robust, game-changing, unlock, elevate and their kin).\n' +
  '- No over-tidy symmetrical paragraphs that feel machine-balanced.\n';

function draftStudentMessage_(data) {
  var found = findRow_(data.issue_id);
  if (!found) return { ok: false, error: 'No issue found with that id.' };
  var i = found.record;
  var kind = String(data.kind || 'longopen');
  var who = (data._user && data._user.name) || '';
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'no api key' };

  var reps = [];
  try { reps = i.reports_json ? JSON.parse(i.reports_json) : []; } catch (e) {}
  // Getting the voice right matters more than the pennies here (Edd, 8 Aug):
  // the whole guide goes in, and the draft uses the strongest model we have.
  var last = reps.length ? reps[reps.length - 1] : null;
  var guide = voiceGuideFor_(who);
  // FB-0186: a "still on it" draft should carry the SPECIFIC next thing to
  // try off the pre-dev checklist, not vague reassurance, so the message
  // itself moves the troubleshooting along. First item not yet done or n/a,
  // in checklist order - the same one the detail pane's next action names.
  var nextStep = '', stepSource = '';
  if (kind === 'longopen' && String(i.category || '').toLowerCase() === 'tech_issue') {
    // Round 55 (Edd, FB-0203): ask for the thing the reasoner worked out from
    // the whole thread, not the next unticked box. If it decided the next move
    // is ours to make, there is nothing to ask the student for at all, and the
    // draft goes back to a plain "still on it" rather than inventing a chore.
    var reasoned = null;
    try { reasoned = nextActionCached_(found, {}); } catch (e) { reasoned = null; }
    if (reasoned && reasoned.ok) {
      stepSource = 'reasoned';
      nextStep = reasoned.instructor_side ? '' : String(reasoned.student_ask || reasoned.action || '');
    }
    if (!stepSource) {
      var ckMap = {}; try { ckMap = i.checklist_json ? JSON.parse(i.checklist_json) : {}; } catch (e) { ckMap = {}; }
      for (var ci = 0; ci < CHECKLIST_ITEMS.length; ci++) {
        var stC = ckMap[CHECKLIST_ITEMS[ci].id];
        if (stC !== 'done' && stC !== 'na') { nextStep = CHECKLIST_ITEMS[ci].label; stepSource = 'checklist'; break; }
      }
    }
  }
  var goal = kind === 'fixed'
    ? 'Tell the student the problem they reported has been fixed, and invite them to try again and shout if anything still looks off.'
    : kind === 'followup'
      ? 'Check in on whether the workaround they were given is still doing the job, without pestering.'
      : nextStep
        ? 'Ask the student to try the one specific troubleshooting step given below, phrased in plain language a non-technical sailor can follow, and make clear their report is still being worked on. Do not invent progress that is not in the notes.'
        : 'Reassure the student their report is not forgotten and is still being worked on. Do not invent progress that is not in the notes.';

  var prompt = 'Draft a short, warm email from an instructor at Ardent Training (an online RYA sailing school) to a student.\n\n' +
    'GOAL: ' + goal + '\n\n' +
    'THE ISSUE:\n' + JSON.stringify({
      summary: i.summary, status: i.status,
      lesson: i.lesson || i.lesson_code || '',
      fix_notes: i.dev_notes || i.resolution_note || '',
      student_first_name: String(i.student_name || '').split(' ')[0]
    }) + '\n' +
    (last && (last.summary || last.raw_text) ? 'LATEST UPDATE ON THE ISSUE: ' + String(last.summary || last.raw_text).slice(0, 400) + '\n' : '') +
    (nextStep
      ? (stepSource === 'reasoned'
          ? 'THE ONE THING TO ASK THE STUDENT (worked out from the whole thread - ask for exactly this and nothing else, and do not ask for anything they have already given us):\n"' + nextStep + '"\n'
          : 'THE NEXT TROUBLESHOOTING STEP TO ASK OF THE STUDENT (from our checklist, written for staff - rephrase it as a plain ask):\n"' + nextStep + '"\n')
      : '') +
    (guide ? '\nWrite it in this instructor\'s own voice. Their style guide:\n"""\n' + guide.slice(0, 30000) + '\n"""\n' : '') +
    AI_TELLS_RULES_ +
    '\nRules: plain text only, no subject line, 50-110 words, greet the student by first name. ' +
    'Say the thing in the first sentence rather than warming up to it, skip reassurance the student did not ask for, and never state a fact that is not in the issue or the fix notes. ' +
    (guide
      ? 'Sign off exactly the way this instructor signs off in the style guide (their name: "' + (who || 'The Ardent team') + '"). '
      : 'Sign off as "' + (who || 'The Ardent team') + '". ') +
    'Promise no dates unless the fix notes give one. ' +
    'Return ONLY the email text, nothing else.';

  try {
    var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: DRAFT_MODEL, max_tokens: 1500, messages: [{ role: 'user', content: prompt }] })
    });
    if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) return { ok: false, error: 'API error ' + res.getResponseCode() };
    var parsed = JSON.parse(res.getContentText() || '{}');
    var text = '';
    (parsed.content || []).forEach(function (c) { if (c.type === 'text') text += c.text; });
    text = text.trim();
    if (!text) return { ok: false, error: 'empty draft' };
    // Belt and braces on the one unambiguous tell: even if the model slips an
    // em dash past the prompt, it never reaches a student (Edd, 11 Aug).
    text = text.replace(/\s*—\s*/g, ', ').replace(/,\s*,/g, ',');
    return { ok: true, text: text, voiced: !!guide, next_step: nextStep || '', step_source: stepSource };
  } catch (e) { return { ok: false, error: String(e) }; }
}

// ============================ THE NEXT ACTION ================================
// Edd, FB-0203: "The next action section should be more intelligent. Look at
// the full transcript, troubleshooting list, etc and make a genuinely useful
// suggested next action."
//
// What was wrong. Two thin sources fed that line. nextActionFor() in the front
// end reads the status and then names the first unticked checklist box, and
// briefAi_ returns a single `next` off the conversation alone. Naming the next
// unticked box is not thinking. Sergei Fedorov had sent videos, updated iOS,
// tried other devices and had the developers already looking at an iPhone
// portrait layout bug, and the pane told us to get him to try a different
// network, purely because that line happened to be untried. A network has
// nothing to do with a layout that only breaks in portrait.
//
// What this does instead. One call, everything on the table: the whole thread
// (every report and update, not a 3,000-character slice), the checklist with
// its labels and real states, where the issue stands with the developers or
// the course team, any open question, past fixes that share words with it, and
// who the student is. Out comes one action, one line of why, and a flag for
// whether it is ours to do or theirs to try.
//
// Why it is cached. The tracker bills Edd's Anthropic credits, so opening an
// unchanged issue must cost nothing. nextActionSignature_ fingerprints the
// things that could change the answer; a matching fingerprint returns the
// stored answer with no API call at all.
var NEXT_ACTION_CAP = 24000;   // characters of thread the reasoner reads
// Bump this whenever the reasoning changes. The Round 55 cache keys on a
// fingerprint of the ISSUE, which is right - reopening an unchanged issue must
// cost nothing. But it means a better prompt would go unnoticed on every issue
// already answered, and the wrong answer Edd was looking at in FB-0214 would
// still be sitting there. The revision rides in the fingerprint, so changing it
// retires every stored answer in one go and each issue recomputes when it is
// next opened. Cheap because it is still one call per issue, only once.
var NEXT_ACTION_REV = 'r66';   // Round 66 added rule 4a (test the password you just set, FB-0240), so every stored answer retires
var NEXT_ACTION_MODEL = ANTHROPIC_MODEL;   // sonnet: this runs often, and it is plenty for the job

// The whole conversation, oldest first, attributed and dated, because who said
// a thing and when is half of what makes an action sensible. reports_json holds
// every report AND every update (addUpdate_ writes an entry too), and raw_text
// is those same texts concatenated, so the entries are the fuller record and
// raw_text is only the fallback for an issue that predates them.
function issueTranscript_(rec) {
  var reps = [];
  try { reps = rec.reports_json ? JSON.parse(rec.reports_json) : []; } catch (e) { reps = []; }
  var parts = [];
  if (reps && reps.length) {
    for (var r = 0; r < reps.length; r++) {
      var rp = reps[r] || {};
      // Round 59. Two things used to be missing from every heading, and both
      // are the difference between reading a conversation and guessing at one.
      //
      // The date. It read `rp.submitted_at || rp.at`, and an entry carries
      // NEITHER - the field is `date`. So every line came through as
      // "--- Edd,  ---" and the reasoner had no way to tell the first thing we
      // said from the last, which is exactly the ordering FB-0214 turned on.
      //
      // The kind. A report is what came IN; an update is what WE did about it.
      // Unlabelled they read as one flat pile, so a note saying "student has
      // been asked to try the Sea Regs portal" carried no more weight than the
      // student's original message, and the answer restarted the diagnosis.
      var when = String(rp.date || rp.submitted_at || rp.at || '').slice(0, 16).replace('T', ' ');
      var isUpdate = String(rp.kind || '') === 'update';
      var head = '--- [' +
        (isUpdate ? 'OUR OWN UPDATE, logged by ' : 'REPORT LOGGED BY ') +
        (rp.instructor_name || 'unknown') +
        (when ? ' on ' + when : ' (undated)') + ']' +
        (r === reps.length - 1 ? '  <<< THE MOST RECENT ENTRY, this is where the story is up to' : '') +
        ' ---';
      var body = String(rp.raw_text || '').trim();
      if (!body && rp.summary) body = String(rp.summary);
      if (body) parts.push(head + '\n' + body);
    }
  }
  if (!parts.length && rec.raw_text) parts.push(String(rec.raw_text));
  var all = parts.join('\n\n');
  var truncated = false;
  if (all.length > NEXT_ACTION_CAP) {
    // Keep the END. The oldest part of a long email chain is the part everyone
    // has already acted on; what is unresolved is at the bottom.
    all = '[the earliest part of this thread is not shown]\n\n' + all.slice(all.length - NEXT_ACTION_CAP);
    truncated = true;
  }
  return { text: all, truncated: truncated, messages: (reps && reps.length) || (rec.raw_text ? 1 : 0) };
}

// A fingerprint of everything that could sensibly change the answer. Note what
// is NOT in it: updated_at. Any edit at all moves updated_at, and re-reading a
// thread because somebody retitled a lesson is money for nothing.
function nextActionSignature_(rec) {
  var reps = [];
  try { reps = rec.reports_json ? JSON.parse(rec.reports_json) : []; } catch (e) { reps = []; }
  var ck = {};
  try { ck = rec.checklist_json ? JSON.parse(rec.checklist_json) : {}; } catch (e) { ck = {}; }
  var ckSig = CHECKLIST_ITEMS.map(function (it) {
    var s = String(ck[it.id] || '');
    return s === 'done' ? 'd' : s === 'na' ? 'n' : 't';
  }).join('');
  return [
    NEXT_ACTION_REV,
    'r' + ((reps && reps.length) || 0),
    'x' + String(rec.raw_text || '').length,
    's' + String(rec.status || ''),
    'p' + String(rec.priority || ''),
    'c' + String(rec.category || ''),
    'k' + ckSig,
    'v' + (rec.dev_passed_at ? 1 : 0) + (rec.dev_fixed_at ? 1 : 0) + String(rec.dev_notes || '').length,
    'q' + String(rec.dev_query_at || ''),
    'a' + String(rec.assignee || ''),
    'z' + (String(rec.student_sorted) === 'true' ? 1 : 0),
    'n' + (String(rec.notified_students) === 'true' ? 1 : 0),
    'i' + (issueHasStudent_(rec) ? 1 : 0)
  ].join('|');
}

// The checklist as evidence rather than as a to-do list. The model gets every
// item with its state, and is told plainly that an untried step is only worth
// suggesting if it could actually explain THIS fault.
function checklistEvidence_(rec, staff) {
  var ck = {};
  try { ck = rec.checklist_json ? JSON.parse(rec.checklist_json) : {}; } catch (e) { ck = {}; }
  var items = checklistItemsFor_(staff);
  if (!items.length) return '(no checklist on this issue)';
  return items.map(function (it) {
    var s = String(ck[it.id] || 'todo');
    var word = s === 'done' ? 'ALREADY DONE' : s === 'na' ? 'not relevant' : 'not yet tried';
    return '- [' + word + '] ' + checklistLabel_(it, staff);
  }).join('\n');
}

// Is there a student on the other end of this issue? The flag the form saved
// wins; rows logged before that column existed get read the old way (Edd,
// FB-0207). Kept next to nextActionAi_ because the reasoner is the loudest
// place a wrong answer shows up.
function issueHasStudent_(rec) {
  var flag = String(rec.student_involved || '').toLowerCase();
  if (flag === 'no') return false;
  if (flag === 'yes') return true;
  if (String(rec.audience || '') === 'internal') return false;
  return !!(String(rec.student_name || '').trim() || String(rec.student_contact || '').trim());
}

// Round 59, FB-0214. Everything we have already put to the student, pulled out
// of the thread and handed over as its own block.
//
// The failure it exists for: Ruth could not log in, her email was not
// recognised, and she had bought the course through Sea Regs in Plymouth - a
// partner school, so she logs in at their portal, not ours. Edd had spotted
// that, sent her the right link, and nudged her again. The next action came
// back "reset her password yourself and send her the login link", which is a
// step BACKWARDS: the password was never the fault, and the link had already
// gone. Edd: "the next action here is completely wrong and takes us back a
// step."
//
// The thread held all of that. It was just buried in the middle of a long
// conversation with no dates and no labels, and the model weighed a playbook
// line above a thing a human had already done. So it gets pulled to the front,
// in order, newest last, and the rules below make it the first thing that
// binds.
function alreadySaidTo_(rec) {
  var reps = [];
  try { reps = rec.reports_json ? JSON.parse(rec.reports_json) : []; } catch (e) { reps = []; }
  var ours = [];
  for (var i = 0; i < reps.length; i++) {
    var rp = reps[i] || {};
    if (String(rp.kind || '') !== 'update') continue;
    var body = String(rp.raw_text || rp.summary || '').trim();
    if (!body) continue;
    ours.push({
      on: String(rp.date || rp.submitted_at || '').slice(0, 10),
      by: rp.instructor_name || 'us',
      said: body.slice(0, 1200)
    });
  }
  return ours;
}

// The one symptom that decides which fix is even plausible: the student's own
// words for what is failing. A password reset cannot fix an account the login
// page has never heard of, and "your username is not recognised" is the login
// page saying exactly that. The playbook lists both the partner-portal fix and
// the reset-their-password fix, and left to itself the reasoner picks whichever
// reads as most decisive rather than whichever matches the error on the screen.
function symptomNote_(rec) {
  var hay = (String(rec.summary || '') + ' ' + String(rec.raw_text || '')).toLowerCase();
  var hits = [];
  if (/not recognis|not recogniz|no account|unknown (user|email)|user not found|doesn.t recognise|does not recognise/.test(hay)) {
    hits.push('The student reports the login NOT RECOGNISING them. That is an account-or-portal fault, not a password fault: a reset cannot help an account the login page has never heard of, and a reset email will never arrive at an address that page does not hold. The order matters and it is not negotiable. FIRST find the account and confirm which portal it lives on. Only once it has been found does resetting a password on it mean anything, and that is a SEPARATE, LATER action. Do not put "reset her password" in the same action as "find the account", and do not put it in the action at all while the account has not been found. And a reset only becomes the action once the student has confirmed she is at the RIGHT portal and still cannot get in: until she has answered that, a new password solves a problem nobody has established she has.');
  }
  if (/sea ?regs|searegs|dosa|partner school|bought (it )?(via|through)|via [a-z ]+ in [a-z]/.test(hay)) {
    hits.push('This student came to us through a PARTNER SCHOOL. Partner-school students have no account at ardent-training.com at all; they log in at their school\'s own portal (searegs.ardent-training.com and the like). This exact pattern has caused a wrong answer twice before, so check the thread for whether they have been sent to the right portal yet before suggesting anything else.');
  }
  if (/reset (email|link)( has)? (never|not) (arriv|com)|no (password )?reset email|didn.t get (an )?email|no email came/.test(hay)) {
    hits.push('A reset email that never arrives usually means the address is not on the account at all (wrong address, or the account lives on a partner portal), not that the mail is slow. Treat a missing reset email as evidence about the ACCOUNT.');
  }
  return hits;
}

// Who a fix belongs to. A course error is the course team's, every time: a
// wrong diagram or a clumsy question is written, not coded. Everything else
// that gets handed on goes to the developers.
function owningTeam_(rec) {
  return String(rec.category || '') === 'course_error' ? 'the course team' : 'the developers';
}

function nextActionAi_(rec) {
  var hasStudent = issueHasStudent_(rec);
  var staff = !hasStudent ||
    String(rec.audience || '') === 'internal' ||
    ['instructor_portal', 'partner_portal'].indexOf(String(rec.section || '')) > -1;
  var tr = issueTranscript_(rec);
  var cat = String(rec.category || 'tech_issue');
  var status = String(rec.status || 'open').toLowerCase();

  // Past fixes that share words with this one. Word overlap only, no second AI
  // call: the reasoner is better placed to judge whether a match is real than
  // a shortlist ranker is, and one call is the whole point.
  var pastFixes = [];
  try {
    // The issue knows its own course, so scoped corpus entries for a different
    // one never reach the reasoner (FB-0231).
    pastFixes = fixCandidatesFor_(String(rec.summary || '') + ' ' + tr.text.slice(0, 3000), null, false,
      kfNormaliseScope_(rec.course)).slice(0, 6);
  } catch (e) { pastFixes = []; }

  var state = {
    summary: rec.summary || '',
    status: status,
    priority: rec.priority || '',
    category: cat,
    request_kind: rec.request_kind || 'fix',
    audience: rec.audience || 'student',
    section: rec.section || '',
    platform: rec.platform || '',
    course: rec.course || '',
    lesson: rec.lesson || rec.lesson_code || '',
    student_name: rec.student_name || '',
    student_has_contact: !!rec.student_contact,
    a_student_is_involved: hasStudent,
    student_sorted: String(rec.student_sorted) === 'true',
    students_notified: String(rec.notified_students) === 'true',
    device_info: rec.device_info || '',
    assignee: rec.assignee || '',
    logged: String(rec.submitted_at || '').slice(0, 10),
    // Round 59, FB-0215. These three used to say "developers" whatever the
    // issue was, and the reasoner reads the key names. On Tim's wording tweak -
    // a course improvement, already sitting with the course team - it duly came
    // back "chase the developers for a status update". Edd: "the developers are
    // not needed (except maybe course creators)". A course error belongs to the
    // course team; nothing about it ever goes near a developer, so the record
    // stops calling them that.
    which_team_owns_this: owningTeam_(rec),
    handed_to_that_team_on: rec.dev_passed_at ? String(rec.dev_passed_at).slice(0, 10) : '',
    that_team_marked_it_fixed_on: rec.dev_fixed_at ? String(rec.dev_fixed_at).slice(0, 10) : '',
    notes_from_that_team: String(rec.dev_notes || '').slice(0, 2000),
    open_question_from_that_team: rec.dev_query_at ? String(rec.dev_query || '') : '',
    open_question_is_for: rec.dev_query_at ? String(rec.dev_query_target || 'admins') : '',
    resolution_note: String(rec.resolution_note || '').slice(0, 1000),
    images_attached: String(rec.image_urls || '') ? String(rec.image_urls).split(',').length : 0,
    courier: rec.courier || '',
    tracking_number: rec.tracking_number || '',
    chase_on: rec.chase_at ? String(rec.chase_at).slice(0, 10) : ''
  };

  var already = alreadySaidTo_(rec);
  var symptoms = symptomNote_(rec);
  // Round 63, FB-0226. Same gate as the troubleshooting steps, asked of the
  // whole record and the whole thread. The stored checklist gets a say: if
  // somebody has already ticked "tried it yourself", the scope is settled and
  // the gate stays shut.
  var ckNow = {};
  try { ckNow = rec.checklist_json ? JSON.parse(rec.checklist_json) : {}; } catch (e) { ckNow = {}; }
  // An improvement is somebody asking for something to be better, not a fault,
  // so there is no scope to establish: a clumsy diagram is clumsy for everyone.
  // Same reasoning as the Round 59 fix to the long-open Actions lane.
  var scopeGate = String(rec.request_kind || 'fix') !== 'improvement' &&
    String(ckNow.replicated || '') !== 'done' &&
    scopeGateOn_(String(rec.summary || '') + ' ' + String(rec.raw_text || '') + ' ' + tr.text);

  var prompt = 'You are the most experienced person on the support desk at Ardent Training, an online RYA sailing school. ' +
    'An instructor has this issue open in front of them and wants ONE genuinely useful next action.\n\n' +
    'THE HARD RULES, in order of importance:\n' +
    (hasStudent ? '' :
      '0. THERE IS NO STUDENT ON THIS ONE. It was logged by one of the team about something they hit themselves, and a_student_is_involved is false. So there is nobody to ask, nobody to relay a step to, and nobody waiting on an update. Do NOT return an action that asks the student anything, tells us to contact or update a student, or waits on a student reply, and do NOT ask for information only a student could give. student_ask must be an empty string and instructor_side must be true. The useful action here is one WE take: reproduce it ourselves and pin down exactly when it happens, check whether anyone else on the team sees it, or hand it to the developers with what we already know.\n') +
    '1. START FROM THE LAST THING THAT WAS SAID, AND CARRY ON FROM THERE. The block headed WHAT WE HAVE ALREADY PUT TO THE STUDENT is the state of play. The next action must be the step that follows ON from it. An action that would undo, repeat, or re-ask something already sent is not an action, it is going backwards, and it is the single worst answer this box can give. Read the most recent entry first and treat everything before it as settled.\n' +
    '2. NEVER suggest anything already suggested. Not the same words, not different words, not "as a first step", not "to be sure". If we have sent them a login link, do not send a login link. If we have asked them to try a portal, do not ask them to try a portal. If the thread shows it has been said, it is spent.\n' +
    '1b. THE MOST RECENT UPDATE IS AN INSTRUCTOR SAYING WHERE THEY HAVE GOT TO, AND THEY KNOW THIS CASE. When the last thing on the record is one of us setting out what has been done and what we are waiting for, that is the plan of the person who has actually spoken to this student. Your job is the NEXT step of that plan, not a different plan. You may only depart from it if the thread contains a fact they have plainly missed, and then you must name that fact in the why line. Deciding that a step they chose not to take is the obvious thing to do is not a fact they missed, it is second-guessing somebody with more of the picture than you have. If their stated position is that we are waiting on the student, then we are waiting on the student.\n' +
    '2a. BEFORE YOU ANSWER, READ YOUR OWN ACTION BACK AND STRIKE OUT ANY PART OF IT WE HAVE ALREADY DONE. Actions go wrong by being bundled: one genuinely new step, quietly welded to a repeat of something already sent. "Look her up and then send her the login link" is not a new action, it is a new action with an old one stapled to it, and the student receives the old one. Keep the new part. Delete the rest. Return ONE step, and if the step you are left with is small, that is correct.\n' +
    '2b. IF WE ARE WAITING ON THEM, SAY SO. When the last thing in the thread is a suggestion of ours that they have not answered yet, the honest next action is to wait, or to chase it if it has been long enough, and to say which. Give the date we last wrote and how long it has been quiet. Do NOT invent a new step to fill the silence: a fresh instruction on top of an unanswered one confuses the student and loses the thread of what we were testing. Working out that there is nothing to do yet IS a useful answer, and it is the right one more often than it gets given.\n' +
    '2c. MATCH THE FIX TO THE SYMPTOM THE STUDENT DESCRIBED, not to whatever the playbook or the checklist happens to list first. Their own words for what is going wrong outrank every generic ordering in here. If they say the login does not recognise them, the fault is with the account or the portal, and a password step is beside the point however standard it is. Read the SYMPTOM NOTES below before you choose.\n' +
    // Round 63, FB-0226. It sits here on purpose: below rules 1, 1b and 2, so it
    // can never override a plan an instructor has already put to the student,
    // and above everything that reads off the playbook or the checklist, so a
    // generic user-side step can never outrank it.
    (scopeGate
      ? '2d. WE DO NOT YET KNOW IF THIS IS EVERYONE OR ONE PERSON, AND THAT DECIDES THE ANSWER. This is the shape of fault where the scope is unknown and can be settled in about ten seconds, and nothing on this record says anybody has looked. ' +
        'Unless rule 1 or 1b gives you a step that plainly follows on from something already said, the next action is OURS: ' + scopeStepText_(staff) + ' ' +
        'Say in the why line what the answer would tell us: ' + scopeBranchText_(staff, owningTeam_(rec)) + ' ' +
        'Do not return a user-side step (a hard refresh, another browser, another network, clearing a cache) as the action, because none of them is worth asking for until we know the fault is on their side. Do not return "hand it on" or "flag it urgently" either, because that assumes the other answer. instructor_side must be true and student_ask must be empty, since this is a thing we do and there is nothing to ask them for.\n'
      : '') +
    '3a. NEVER ask the student for something the thread already gives you. If they have told us the device, the iOS version, the browser, or sent a screenshot or video, that question is answered.\n' +
    '3. A troubleshooting step that could not possibly explain THIS fault is not a next action, however untried it is. A layout that breaks only in portrait, or a video that stops at the same second every time, is not going to be fixed by a different network, a hard refresh or a VPN being turned off. Ignore untried steps that do not fit the symptom, and say so in the why line if that is the interesting part.\n' +
    '4. The fastest resolution is often something WE do. The instructor can reset a password from the students tab of the instructor portal, assign a course to an account, extend an account by hand, mark an exam manually from photos, post an answer in the course live chat, re-send an ebook, or raise an invoice. When one of those settles it, THAT is the action, phrased as a thing we do.\n' +
    // FB-0240. Setting a password is not the same as knowing it works, and the
    // instructor is the only person who can tell the difference before the
    // student is told. Costs ten seconds, and saves a reply that says "I have
    // reset your password" to somebody who still cannot get in.
    '4a. A PASSWORD YOU HAVE JUST SET IS NOT A WORKING LOGIN UNTIL SOMEBODY HAS USED IT. Whenever the action is resetting or changing a student\'s password, the same action includes logging in yourself with that email and that password before anything is sent to them. If it lets you in, the reply can say so plainly. If it does not, we have learnt the fault was never the password, which is the more useful answer of the two and it arrives before the student has been told the wrong thing.\n' +
    '5. When the troubleshooting has genuinely gone as far as it can and the fault is real, the action is to move it on, not to keep poking the student: hand it on with the specific evidence that will let them reproduce it, chase whoever already has it if it has sat too long, or answer the question they have asked us. Hand it to the team named in which_team_owns_this and to nobody else. A course error - wrong wording, a wrong diagram, a confusing question - is the COURSE TEAM\'s, always. Never say \"the developers\" about a course error; no developer will ever touch it.\n' +
    '6. If the thread shows that team is already on it and there is nothing new to give them, say plainly that the useful next action is to leave the student alone and chase internally, and name what we would chase for.\n' +
    '7. Notice what the ISSUE RECORD says versus what the thread says. If the conversation says somebody is working on it but the record was never handed to them, saying so IS the useful action, because nothing in our system is tracking it.\n' +
    '8. Read the student. Somebody who has diagnosed the fault themselves and offered to help does not need to be walked through the basics.\n' +
    '9. One action. Concrete, plain English, addressed to the instructor, no jargon, no lists, no hedging between two options.\n\n' +
    // Round 59, FB-0214. Ahead of the record, ahead of the checklist, ahead of
    // the playbook. What a human has already put to this student is the one
    // thing the answer has to be consistent with, so it goes first and it is
    // named for what it is.
    (already.length
      ? 'WHAT WE HAVE ALREADY PUT TO THIS STUDENT (oldest first, so the LAST one is where things stand). Everything in here is spent: it has been said, and saying it again takes us backwards:\n' +
        already.map(function (u) {
          return '- ' + (u.on || 'undated') + ', ' + u.by + ' logged: ' + u.said;
        }).join('\n') + '\n\n'
      : 'WHAT WE HAVE ALREADY PUT TO THIS STUDENT: nothing has been logged since the report came in, so this is still at the start.\n\n') +
    (symptoms.length
      ? 'SYMPTOM NOTES (read these before choosing, they beat the generic ordering below):\n- ' + symptoms.join('\n- ') + '\n\n'
      : '') +
    'THE ISSUE RECORD (our system\'s own view):\n' + JSON.stringify(state, null, 1) + '\n\n' +
    'THE PRE-DEVELOPER CHECKLIST (what the team has ticked off, as evidence - NOT a list to read the next line from):\n' +
    checklistEvidence_(rec, staff) + '\n\n' +
    (pastFixes.length
      ? 'PAST ISSUES WE FIXED THAT SHARE WORDS WITH THIS ONE (only use one if it genuinely matches; word overlap is not a match):\n' +
        JSON.stringify(pastFixes.map(function (c) {
          return { problem: String(c.summary || '').slice(0, 300), fix: String(c.resolution_note || '').slice(0, 300), lesson: c.lesson_code || '' };
        })) + '\n\n'
      : '') +
    'OUR TROUBLESHOOTING PLAYBOOK (for the account faults that have a known specific fix):\n' + getPlaybook_() + '\n\n' +
    'THE WHOLE THREAD, oldest first (every report and update logged on this issue):\n"""\n' + tr.text + '\n"""\n\n' +
    (tr.truncated ? 'NOTE: this thread was long enough that the earliest part is not shown. Do not treat the start as missing information.\n\n' : '') +
    'Return ONLY JSON, no prose, no fences:\n' +
    '{"action": "<the single next action, one or two sentences, addressed to the instructor as something to do>",\n' +
    ' "why": "<one short line saying why this and not the obvious alternative - name the thing in the thread that decided it>",\n' +
    ' "instructor_side": true or false (true when this is something WE do; false when it is something the student is asked to do or tell us),\n' +
    ' "student_ask": "<empty string when instructor_side is true; otherwise the one thing to ask the student, in plain words a non-technical sailor can act on>",\n' +
    ' "waiting_on_student": true or false (true when the honest position is that we have already asked and the ball is in their court)}';

  var got = anthropicRaw_(NEXT_ACTION_MODEL, prompt, 16000);
  if (!got.json || !got.json.action) {
    return { ok: false, why: got.why || 'the reply had no action in it' };
  }
  return {
    ok: true,
    action: String(got.json.action || '').slice(0, 600),
    why: String(got.json.why || '').slice(0, 400),
    instructor_side: !!got.json.instructor_side,
    student_ask: String(got.json.student_ask || '').slice(0, 400),
    waiting_on_student: !!got.json.waiting_on_student,
    messages: tr.messages
  };
}

// Read the cached answer, or work one out and cache it. force:true is the
// Rethink button. compute:false means "give me the cached one if it is still
// good, otherwise nothing" - for callers that must not spend money.
function nextActionCached_(found, opts) {
  opts = opts || {};
  var rec = found.record;
  var sig = nextActionSignature_(rec);
  var stored = null;
  try { stored = rec.next_action_json ? JSON.parse(rec.next_action_json) : null; } catch (e) { stored = null; }
  if (stored && stored.sig === sig && !opts.force) {
    stored.cached = true;
    stored.ok = true;
    return stored;
  }
  if (opts.compute === false) return { ok: false, stale: !!stored, why: 'not computed yet' };

  var out = nextActionAi_(rec);
  if (!out.ok) {
    // A failure must never quietly pose as an answer (the FB-0150 lesson), so
    // nothing is cached and the caller is told exactly what went wrong.
    return { ok: false, why: out.why, had_stale: !!stored };
  }
  var store = {
    action: out.action, why: out.why, instructor_side: out.instructor_side,
    student_ask: out.student_ask, waiting_on_student: !!out.waiting_on_student,
    messages: out.messages,
    sig: sig, at: new Date().toISOString(), model: NEXT_ACTION_MODEL
  };
  try {
    setCellOnIssue_(found, 'next_action_json', JSON.stringify(store));
  } catch (e) {}   // a cache that will not write is a slow feature, not a broken one
  store.ok = true;
  store.cached = false;
  return store;
}

function nextAction_(data) {
  var id = data && data.issue_id;
  if (!id) return { ok: false, error: 'nextAction needs an issue_id' };
  var found = findRow_(id);
  if (!found) return { ok: false, error: 'No issue found with id ' + id };
  var out = nextActionCached_(found, { force: !!data.force });
  if (!out.ok) return { ok: false, issue_id: id, error: out.why || 'could not work out a next action' };
  return {
    ok: true, issue_id: id, action: out.action, why: out.why,
    instructor_side: !!out.instructor_side, student_ask: out.student_ask || '',
    waiting_on_student: !!out.waiting_on_student,
    messages: out.messages || 0, at: out.at, cached: !!out.cached
  };
}

// Find the Chatwoot contact for an email (Round 61).
//
// This used to end in "|| list[0]", so when the search turned up people but
// none of them actually had that email, we opened whoever happened to be first.
// A wrong contact and a right one looked identical from the outside, which is
// the worst way for a lookup to fail. Now: an exact match wins; a single near
// match is used but says so; several near matches with no exact one is a miss,
// because picking between them is a guess and we should not be guessing about
// which student we are looking at.
function chatwootFindContact_(email) {
  var out = chatwootCall_('/contacts/search?q=' + encodeURIComponent(email));
  var list = (out && out.payload) || [];
  var exact = list.filter(function (c) { return String(c.email || '').trim().toLowerCase() === email; })[0];
  if (exact && exact.id) return { hit: exact, match: 'exact' };
  if (list.length === 1 && list[0] && list[0].id) return { hit: list[0], match: 'near' };
  return { hit: null, match: list.length ? 'ambiguous' : 'none' };
}
function contactMissMessage_(email, match) {
  return match === 'ambiguous'
    ? 'Chatwoot found several contacts for ' + email + ' but none with that exact email, so it is not safe to guess which one is the right student. Open Chatwoot and search by hand.'
    : 'No Chatwoot contact matches ' + email + '.';
}

// ---- jump to a contact in Chatwoot (Edd, FB-0162) --------------------------
function chatwootContactUrl_(data) {
  var base = CHATWOOT_BASE + '/app/accounts/' + chatwootCfg_().account + '/contacts/';
  // If the issue already carries the contact id, there is nothing to look up
  // and nothing to get wrong (Round 61).
  var known = String(data.contact_id || '').trim();
  if (known) return { ok: true, url: base + known, match: 'id' };
  var email = String(data.email || '').trim().toLowerCase();
  if (!email) return { ok: false, error: 'need an email or a contact id' };
  try {
    var found = chatwootFindContact_(email);
    if (!found.hit) return { ok: false, error: contactMissMessage_(email, found.match) };
    return { ok: true, url: base + found.hit.id, match: found.match };
  } catch (e) { return { ok: false, error: String(e) }; }
}

// "Fetch update" from the detail pane (Edd, FB-0173): go and look in Chatwoot
// for any conversation with this issue's student that is NEWER than the issue's
// last activity, read it, and if it genuinely says something about this issue,
// add it as an update. Same verdict AI as the overnight scan (finder plus a
// stricter second opinion), because a wrong "it's fixed" is the costly mistake.
// This path NEVER sets a status itself - a "fixed" verdict lands as an update
// for a human to verify, exactly like the scan.
function fetchStudentUpdate_(data) {
  var id = data.issue_id;
  if (!id) return { ok: false, error: 'fetchStudentUpdate needs an issue_id' };
  var found = findRow_(id);
  if (!found) return { ok: false, error: 'No issue found with id ' + id };
  var rec = found.record;

  var email = String(rec.student_contact || '').trim().toLowerCase();
  if (email.indexOf('@') < 0) {
    // Fall back to any report on the issue that carried an email.
    try {
      (JSON.parse(rec.reports_json || '[]') || []).forEach(function (rp) {
        var c = String(rp.student_contact || '').trim().toLowerCase();
        if (email.indexOf('@') < 0 && c.indexOf('@') > 0) email = c;
      });
    } catch (e) {}
  }
  if (email.indexOf('@') < 0) {
    return { ok: true, found: false, message: 'No student email on this issue, so there is nothing to look up in Chatwoot.' };
  }

  // Contact, then their conversations.
  // The stored contact id first: it is exact, and it costs no API call
  // (Round 61). Only fall back to searching by email when we haven't got one.
  var hit = null;
  var stored = String(rec.chatwoot_contact_id || '').trim();
  if (stored) {
    hit = { id: stored };
  } else {
    var found;
    try {
      found = chatwootFindContact_(email);
    } catch (e) { return { ok: false, error: 'Chatwoot contact search failed: ' + String(e.message || e) }; }
    if (!found.hit) return { ok: true, found: false, message: contactMissMessage_(email, found.match) };
    hit = found.hit;
  }

  var convs;
  try {
    var cv = chatwootCall_('/contacts/' + hit.id + '/conversations');
    convs = (cv && cv.payload) || [];
  } catch (e) { return { ok: false, error: 'Could not list the conversations: ' + String(e.message || e) }; }

  // First choice: a conversation we have never read (active since the issue
  // was LOGGED and not stamped into the issue text - both the scan and this
  // path stamp the conversation number, so a re-press can't import the same
  // chat twice as "new"). The gate used to be updated_at, which hid genuinely
  // unread threads the moment anyone touched the issue: Edd ticked Sergei's
  // checklist at 14:59 and the support note from earlier that day instantly
  // read as "old" (FB-0199, live re-test). Dedupe is the stamp's job, not the
  // clock's.
  var since = new Date(rec.submitted_at || rec.updated_at || 0).getTime();
  var already = String(rec.raw_text || '');
  var fresh = convs.filter(function (c) {
    var t = Number(c.last_activity_at || c.timestamp || 0) * 1000;
    if (!(t > since)) return false;
    return already.indexOf('Chatwoot conversation ' + c.id) < 0;
  });
  fresh.sort(function (a, b) { return Number(b.last_activity_at || 0) - Number(a.last_activity_at || 0); });
  var conv = fresh[0] || null;
  var grown = false, prevCount = 0;

  // Second choice (Edd, FB-0199): the conversation we already know about may
  // have GROWN - same thread, more messages. The old code skipped any stamped
  // conversation outright, so a thread that gained a page of detail read as
  // "no new conversations". Time comparisons can't be trusted here (any edit
  // to the issue moves updated_at), so the check is message COUNT against the
  // count recorded when the thread was last read; old stamps without a count
  // fall through to the AI diff below.
  if (!conv) {
    var known = convs.filter(function (c) { return already.indexOf('Chatwoot conversation ' + c.id) > -1; });
    known.sort(function (a, b) { return Number(b.last_activity_at || 0) - Number(a.last_activity_at || 0); });
    if (known.length) {
      conv = known[0];
      grown = true;
      var cm = already.match(new RegExp('Chatwoot conversation ' + conv.id + ' \\((\\d+) messages\\)'));
      prevCount = cm ? Number(cm[1]) : 0;
    }
  }
  if (!conv) {
    return { ok: true, found: false, message: 'No conversations with this student to check - nothing new, and nothing linked.' };
  }

  var imp = chatwootImport_({ conversation: String(conv.id) });
  if (!imp || !imp.ok) return { ok: false, error: 'Could not import conversation ' + conv.id + ': ' + ((imp && imp.error) || 'unknown') };
  if (!imp.message_count) {
    return { ok: true, found: false, message: 'The ' + (grown ? 'linked' : 'newer') + ' conversation (' + conv.id + ') has no readable messages.', conversation: imp.link };
  }
  if (grown && prevCount && Number(imp.message_count) <= prevCount) {
    return { ok: true, found: false, message: 'The linked conversation (' + conv.id + ') has no new messages since it was last read (' + imp.message_count + ' messages).', conversation: imp.link };
  }

  // Verdict pass, same shape as the overnight scan's update path. When the
  // thread is one we've read before (grown), the question changes: is there
  // anything here BEYOND what the issue already holds? (Edd, FB-0199)
  var alreadyTail = grown ? String(rec.raw_text || '').slice(-3000) : '';
  var p = 'A student with an OPEN issue on our sailing course platform has been in touch again. ' +
    'Decide what this conversation says about THAT issue, and nothing else.\n\n' +
    'THE OPEN ISSUE:\n' + JSON.stringify({ summary: rec.summary, status: rec.status, lesson_code: rec.lesson_code, logged: rec.submitted_at }) + '\n\n' +
    (grown
      ? 'ALREADY RECORDED ON THE ISSUE (the end of its running notes - anything in here is NOT new):\n"""\n' + alreadyTail + '\n"""\n\n' +
        'THE FULL CONVERSATION AS IT STANDS NOW (this thread has been read before and has since grown - judge only what is NEW against the notes above):\n'
      : 'THE NEW CONVERSATION:\n') +
    String(imp.transcript || '').slice(0, 12000) + '\n\n' +
    'Choose one verdict:\n' +
    '- "fixed": the student says it now works, or confirms the fix or workaround did the job.\n' +
    '- "still_broken": the student says it is still happening, happening again, or worse.\n' +
    '- "new_detail": genuinely new information about the same problem (another device, steps to reproduce, when it started, more students affected' +
    (grown ? ', or troubleshooting done since the notes above were written' : '') + ').\n' +
    '- "nothing_new": the conversation ' + (grown ? 'adds nothing beyond what the notes above already record' : 'is about something else entirely, or adds nothing') + '. This is the common answer - use it freely.\n\n' +
    'Return ONLY JSON: {"verdict":"fixed|still_broken|new_detail|nothing_new","note":"<one sentence of what the student actually said' + (grown ? ' that is new' : '') + '>"}. No prose, no fences.';
  // 16k, not 400: the grown-thread prompt carries the issue notes AND the
  // full transcript, and a big prompt makes the model think before answering
  // - thinking spends from max_tokens, and at 400 the whole budget burned
  // with no JSON left (the r44 trap, re-hit live on Sergei's thread).
  // anthropicRaw_ also says WHY when it fails instead of a bare null.
  var got1 = anthropicRaw_(FINDER_MODEL, p, 16000);
  var first = got1.json;
  if (first === null) return { ok: false, error: 'Reading the conversation failed (' + (got1.why || 'no answer') + '). Try again in a minute.' };
  if (!first.verdict || first.verdict === 'nothing_new') {
    return { ok: true, found: false,
      message: grown
        ? 'Re-read the linked conversation (' + conv.id + ', now ' + imp.message_count + ' messages) - nothing new on this issue beyond what is already recorded.'
        : 'Found a newer conversation, but it has nothing new on this issue.',
      conversation: imp.link };
  }

  // Second opinion - strict, exactly as the scan runs it.
  var vp = 'A first-pass AI read a support conversation and concluded it is an update on a known open issue. Disagree if it is wrong.\n\n' +
    'OPEN ISSUE: ' + JSON.stringify({ summary: rec.summary, status: rec.status }) + '\n' +
    'CLAIMED VERDICT: ' + first.verdict + ' - ' + (first.note || '') + '\n\n' +
    (grown ? 'ALREADY RECORDED ON THE ISSUE (not new):\n"""\n' + alreadyTail + '\n"""\n\n' : '') +
    'CONVERSATION:\n' + String(imp.transcript || '').slice(0, 12000) + '\n\n' +
    'Be strict. "fixed" requires the student actually confirming it works now, not an instructor hoping so. ' +
    'If the conversation is really about a different problem, disagree.' +
    (grown ? ' This thread was read before: disagree if the claimed news is already in the recorded notes.' : '') + '\n' +
    'Return ONLY JSON: {"agree":true or false,"verdict":"fixed|still_broken|new_detail","note":"<corrected one sentence>"}. No prose, no fences.';
  var got2 = anthropicRaw_(VERIFIER_MODEL, vp, 16000);
  var second = got2.json;
  if (second === null) return { ok: false, error: 'The second-opinion read failed (' + (got2.why || 'no answer') + '). Try again in a minute.' };
  if (second.agree !== true) {
    return { ok: true, found: false, message: 'A newer conversation exists, but the second-opinion check was not convinced it relates to this issue. Worth a human read.', conversation: imp.link };
  }

  var verdict = second.verdict || first.verdict;
  var note = second.note || first.note || '';
  var VLABEL = { fixed: 'Student says it now works', still_broken: 'Student says it is still happening', new_detail: 'New detail from the student' };
  var vLabel = VLABEL[verdict] || 'Update from the student';

  var r = addUpdate_({
    issue_id: id,
    instructor_name: 'Fetch update',
    summary: vLabel + ': ' + note,
    // The message count rides in the stamp so a later press can tell "this
    // thread has grown" from "this thread is as we left it" (FB-0199).
    raw_text: '[Fetch update, Chatwoot conversation ' + conv.id + ' (' + imp.message_count + ' messages)]\n' + vLabel + '.\n' + note +
      '\n\nTranscript:\n' + String(imp.transcript || '').slice(0, 5000),
    student_name: imp.student_name || '',
    student_contact: imp.student_contact || '',
    // Since r43.4 the import copies screenshots to Drive; carry them onto the
    // issue so a fetched update keeps its pictures.
    image_urls: (imp.images && imp.images.length) ? imp.images.join(',') : '',
    app_url: data.app_url || getAppUrl_()
  });
  if (!r || !r.ok) return { ok: false, error: 'Read the conversation fine, but adding the update failed: ' + ((r && r.error) || 'unknown') };

  // The transcript usually says what's been TRIED too, so bring the
  // troubleshooting checklist up to date with it (Edd, FB-0198). Merge only
  // upwards: a human's done/na ticks are never downgraded by the AI read.
  var checklistNote = '';
  if (String(rec.category || '').toLowerCase() === 'tech_issue') {
    try {
      var ts2 = troubleshoot_({ raw_text: String(imp.transcript || '').slice(0, 12000), existing_history: String(rec.raw_text || '').slice(-4000) });
      if (ts2 && ts2.ok && ts2.found && !ts2.degraded && ts2.checklist) {
        var cur = {}; try { cur = rec.checklist_json ? JSON.parse(rec.checklist_json) : {}; } catch (e) { cur = {}; }
        var merged = {}, changed = false;
        Object.keys(ts2.checklist).forEach(function (k) {
          var was = cur[k], now = ts2.checklist[k];
          merged[k] = (was === 'done' || was === 'na') ? was : now;
          if (merged[k] !== was) changed = true;
        });
        Object.keys(cur).forEach(function (k) { if (!(k in merged)) merged[k] = cur[k]; });
        if (changed) {
          var f2 = findRow_(id);  // re-find: addUpdate_ just rewrote the row
          if (f2) { setCellOnIssue_(f2, 'checklist_json', JSON.stringify(merged)); checklistNote = ' The troubleshooting checklist was updated from the transcript.'; }
        }
      }
    } catch (e) {}
  }

  return {
    ok: true, found: true, verdict: verdict,
    message: vLabel + ': ' + note + checklistNote,
    conversation: imp.link,
    images: (imp.images || []).length
  };
}

// ============================ LIVE CASES (Round 45) ==========================
// The instructor-first workspace (Edd, 9 Aug): a student's in touch, what do
// we say? A case is one pulled-in Chatwoot conversation, shared across the
// whole team via the LiveCases sheet tab so anyone can pick it up mid-flow.
// The AI plays the conversation back (summary, what's been tried, the next
// step, any known fix), drafts go out copy-paste, and the bug report falls out
// of the conversation as a "checkpoint" rather than being the way in.

var LIVECASES_SHEET = 'LiveCases';
var LIVECASE_HEADERS = ['conversation_id', 'student_name', 'student_contact', 'opened_by', 'opened_at',
  'last_activity', 'last_touched_by', 'status', 'issue_id', 'summary', 'brief_json', 'draft_count', 'unread'];

function liveCasesSheet_(create) {
  var sh = sheetByName_(LIVECASES_SHEET);
  if (!sh && create) {
    sh = ss_().insertSheet(LIVECASES_SHEET);
    sh.getRange(1, 1, 1, LIVECASE_HEADERS.length).setValues([LIVECASE_HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function liveCaseRows_() {
  var sh = liveCasesSheet_(false);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    if (!values[r][0]) continue;
    var rec = {};
    for (var c = 0; c < head.length; c++) rec[head[c]] = values[r][c];
    rec.conversation_id = String(rec.conversation_id);
    rec._rowNum = r + 1;
    out.push(rec);
  }
  return out;
}
function liveCaseFind_(convId) {
  var id = String(convId || '');
  if (!id) return null;
  var rows = liveCaseRows_();
  for (var i = 0; i < rows.length; i++) if (rows[i].conversation_id === id) return rows[i];
  return null;
}
function liveCaseSave_(rec) {
  var sh = liveCasesSheet_(true);
  var row = LIVECASE_HEADERS.map(function (h) { return rec[h] != null ? rec[h] : ''; });
  if (rec._rowNum) sh.getRange(rec._rowNum, 1, 1, LIVECASE_HEADERS.length).setValues([row]);
  else { sh.appendRow(row); rec._rowNum = sh.getLastRow(); }
  return rec;
}
function caseBriefJson_(rec) {
  try { return rec.brief_json ? JSON.parse(rec.brief_json) : {}; } catch (e) { return {}; }
}
function isTrueLike_(v) { return v === true || String(v).toLowerCase() === 'true'; }

// The open case list, most recent activity first. Also runs the leave rule:
// a case drops off once its issue is reported AND the student is sorted (or
// the issue resolved), cross-checked against Chatwoot's own conversation
// status - still open there and the case stays, flagged "chat still open".
// The peek is a check, never a gate, so any Chatwoot hiccup just skips it.
function listLiveCases_(data) {
  var rows = liveCaseRows_().filter(function (r) { return String(r.status || 'open') === 'open'; });
  // One read of the issue sheets, not one per case.
  var issueById = {};
  if (rows.some(function (r) { return r.issue_id; })) {
    // Round 54: bootstrap has already read the issues, so it hands them over
    // rather than paying for a second whole-spreadsheet read. Only status and
    // student_sorted are used below, and the list projection carries both.
    var pre = data && data._issues;
    (pre || getIssues_().issues || []).forEach(function (i) { issueById[i.issue_id] = i; });
  }
  var peeks = 0;
  rows.forEach(function (r) {
    r._brief = caseBriefJson_(r);
    if (!r.issue_id) return;
    var iss = issueById[r.issue_id];
    if (!iss) return;
    r.issue_status = String(iss.status || 'open').toLowerCase();
    var sorted = r.issue_status === 'resolved' || r.issue_status === 'resolved_tbc' || isTrueLike_(iss.student_sorted);
    if (!sorted || peeks >= 8) return;
    peeks++;
    var cwStatus = '';
    try {
      var conv = chatwootCall_('/conversations/' + r.conversation_id);
      cwStatus = String((conv && conv.status) || (conv && conv.payload && conv.payload.status) || '');
    } catch (e) { cwStatus = ''; }
    if (cwStatus === 'resolved') {
      r.status = 'closed';
      var b = r._brief; b.auto_closed = new Date().toISOString();
      r.brief_json = JSON.stringify(b);
      liveCaseSave_(r);
    } else if (cwStatus) {
      r.chat_still_open = true;
    }
  });
  var open = rows.filter(function (r) { return String(r.status) === 'open'; });
  open.sort(function (a, b) { return new Date(b.last_activity || 0) - new Date(a.last_activity || 0); });
  return { ok: true, cases: open.map(function (r) {
    return {
      conversation_id: r.conversation_id,
      student_name: r.student_name || '',
      student_contact: r.student_contact || '',
      opened_by: r.opened_by || '',
      opened_at: r.opened_at || '',
      last_activity: r.last_activity || '',
      last_touched_by: r.last_touched_by || '',
      issue_id: r.issue_id || '',
      issue_status: r.issue_status || '',
      summary: r.summary || '',
      draft_count: Number(r.draft_count) || 0,
      unread: isTrueLike_(r.unread),
      chat_still_open: !!r.chat_still_open,
      brief: r._brief || {}
    };
  }) };
}

// One AI read of the whole conversation: who and what (so the instructor can
// confirm it's the right chat), what's been tried in plain English, and the
// single next step out of the playbook. Same strictness rules as the
// troubleshoot helper: only count a step as tried if the conversation says so.
// issueId (optional): once a case has an issue filed against it, the issue
// knows things the chat does not - what the team has ticked off, whether the
// developers have it, what they said back. Round 55 folds that in so the case
// brief's next step is reasoned off the same evidence as the detail pane's,
// rather than off the conversation alone. Still ONE call: the brief has to
// read the transcript anyway, and a second call would double the wait on a
// press that is already slow.
function briefAi_(transcript, issueId) {
  var extra = '';
  if (issueId) {
    var f = null; try { f = findRow_(issueId); } catch (e) { f = null; }
    if (f) {
      var ir = f.record;
      var staffB = String(ir.audience || '') === 'internal';
      extra = '\nTHE ISSUE ALREADY FILED FOR THIS STUDENT (our system\'s own view - the chat does not know any of this):\n' +
        JSON.stringify({
          summary: ir.summary || '', status: ir.status || '', priority: ir.priority || '',
          category: ir.category || '', lesson: ir.lesson || ir.lesson_code || '',
          handed_to_developers_on: ir.dev_passed_at ? String(ir.dev_passed_at).slice(0, 10) : '',
          marked_fixed_on: ir.dev_fixed_at ? String(ir.dev_fixed_at).slice(0, 10) : '',
          developer_notes: String(ir.dev_notes || '').slice(0, 1200),
          open_question_from_the_team: ir.dev_query_at ? String(ir.dev_query || '') : '',
          assignee: ir.assignee || ''
        }) + '\n\nWHAT THE TEAM HAS ALREADY TICKED OFF ON IT (evidence, NOT a list to read the next line from):\n' +
        checklistEvidence_(ir, staffB) + '\n';
    }
  }
  var prompt = 'You are helping an Ardent Training instructor (online RYA sailing school) pick up a live student support conversation mid-flow. ' +
    'Read the conversation and play it back, so whoever opens it next knows exactly where things stand without reading the whole thread.\n\n' +
    'OUR TROUBLESHOOTING PLAYBOOK (for choosing the next step):\n' + getPlaybook_() + '\n\n' +
    extra +
    '\nTHE CONVERSATION:\n"""\n' + String(transcript || '').slice(0, 14000) + '\n"""\n\n' +
    'Rules: only count a step as already tried if it is EXPLICITLY described in the conversation - never assume. ' +
    // Round 55 (Edd, FB-0203): the same hard rules the detail pane's reasoner
    // works to, so the two never disagree about what to do next.
    'NEVER ask the student for something the conversation already gives you - if they have told us the device, the OS version or the browser, or sent a screenshot or video, that question is answered. ' +
    'A step that could not possibly explain THIS fault is not a next step however untried it is: a layout that only breaks in portrait, or a video that stops at the same second every time, will not be fixed by a different network or a hard refresh. ' +
    'When the troubleshooting has genuinely gone as far as it can and the fault is real, the next step is to move it on rather than keep poking the student - hand it to the developers or the course team with the evidence that lets them reproduce it, chase whoever already has it, or answer the question they have asked us. ' +
    'If the issue record above shows the developers already have it and there is nothing new to give them, say plainly that the next step is to leave the student in peace and chase internally, and name what for. ' +
    'The app working on another device does NOT mean a different network was tried. ' +
    'The next step is ONE step, the most useful one, chosen from the playbook order and skipping anything already done; ' +
    'if it matches a known account issue in the playbook, that specific fix IS the next step. ' +
    'If the conversation is not a tech problem (a content question, an admin query, a shipping problem), the next step is whatever genuinely helps the student, playbook or not. ' +
    'IMPORTANT (r47): the fastest resolution is often something WE do, not something we ask the student to do. ' +
    'The instructor can reset a student\'s password from the instructor portal, assign a course to their account, ' +
    'mark an exam manually from photos of their answers, post an answer in the course live chat, re-send an ebook, or raise an invoice. ' +
    'When one of those would resolve it, the next step IS that action, phrased as a thing the instructor does ' +
    '("Reset the password for them in the students tab of the instructor portal"), not another question or step routed through the student. ' +
    'Keep everything in plain English an untechnical instructor can act on.\n\n' +
    'One more thing (Edd, FB-0179): screenshots a student shares are often their own COURSEWORK - chartwork photos, ' +
    'exam answers, assessment pages sent for marking - not pictures of a fault. Read the conversation for which it is.\n\n' +
    'Return ONLY JSON, no prose, no fences:\n' +
    '{"summary": "<one or two plain sentences: who the student is, what is going wrong, and on what device or platform if known>",\n' +
    ' "device": "<their device / OS / browser or app if mentioned, else empty string>",\n' +
    // Round 65 (FB-0231). One extra field on a call we were already making, so
    // the known-fix lookup can rule out entries scoped to a different course.
    // Without it every corpus entry reads as applying to everybody.
    ' "course": "<which of our courses this student is on, if the conversation says: Essential Navigation, Day Skipper, Yachtmaster, Fast Track, SRC or PPR. Empty string if it does not say - do not guess>",\n' +
    ' "tried": ["<each thing already tried, one short plain-English entry each - empty list if nothing yet>"],\n' +
    ' "next": "<the single recommended next step, short and practical, addressed to the instructor>",\n' +
    ' "next_why": "<one short line saying why this one and not the obvious alternative - name the thing in the conversation or the issue record that decided it>",\n' +
    ' "instructor_action": true or false (true when the next step is an action the INSTRUCTOR performs - a password reset, a course assignment, manual marking, posting in the chat, an invoice - rather than something the student is asked to try),\n' +
    ' "images_note": "<empty string normally; when screenshots in the chat are likely the student\'s coursework or chartwork for marking rather than fault evidence, one short line saying so>"}';
  return anthropicRaw_(ANTHROPIC_MODEL, prompt, 16000);
}

// Server-side stand-in for the form's candidate ranking: past resolved issues
// that genuinely share words with this conversation, most overlap first, so
// suggestFix_ gets a shortlist rather than an invitation to stretch.
var FIX_STOPWORDS = { 'the': 1, 'and': 1, 'that': 1, 'this': 1, 'with': 1, 'have': 1, 'from': 1, 'your': 1, 'their': 1,
  'they': 1, 'there': 1, 'been': 1, 'were': 1, 'will': 1, 'would': 1, 'could': 1, 'course': 1, 'student': 1,
  'ardent': 1, 'training': 1, 'lesson': 1, 'hello': 1, 'thanks': 1, 'thank': 1, 'please': 1, 'just': 1,
  'what': 1, 'when': 1, 'where': 1, 'which': 1, 'again': 1, 'then': 1, 'them': 1, 'some': 1, 'also': 1,
  'because': 1, 'about': 1, 'into': 1, 'over': 1, 'after': 1, 'before': 1, 'really': 1, 'still': 1,
  'issue': 1, 'problem': 1, 'working': 1, 'works': 1, 'tried': 1, 'trying': 1, 'resolved': 1, 'sorted': 1 };
// cutoffIso (optional) keeps the backtest honest: pass a date and the lookup
// only sees fixes that existed BEFORE it - resolved issues by resolved_at,
// KnownFixes rows by resolved_date. Live use passes nothing and sees it all.
// kfOnly skips the Issues sweep - for callers that already have their own
// issue candidates and only want the KnownFixes corpus folded in.
// course (Round 65, FB-0231) is the course THIS case is on, when we know it. A
// corpus entry scoped to other courses is dropped outright rather than left for
// the AI to talk itself out of, because the scoring below cannot tell the
// difference: a Day Skipper enrolment fix and a Yachtmaster enrolment fault
// share nearly every word, so word overlap will always rank the wrong one high.
// The scoring weights themselves are untouched.
function fixCandidatesFor_(text, cutoffIso, kfOnly, course) {
  var words = {};
  String(text || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).forEach(function (w) {
    if (w.length > 3 && !FIX_STOPWORDS[w]) words[w] = true;
  });
  var keys = Object.keys(words);
  if (!keys.length) return [];
  var cutoff = cutoffIso ? new Date(cutoffIso).getTime() : 0;
  function scoreHay(hay) {
    var score = 0;
    for (var k = 0; k < keys.length; k++) if (hay.indexOf(keys[k]) > -1) score++;
    return score;
  }
  var scored = [];
  (kfOnly ? [] : (getIssues_().issues || [])).forEach(function (i) {
    if (String(i.status || '').toLowerCase() !== 'resolved') return;
    if (!i.resolution_note) return;
    if (cutoff) {
      var when = i.resolved_at ? new Date(i.resolved_at).getTime() : 0;
      if (!when || when >= cutoff) return;
    }
    var hay = ((i.summary || '') + ' ' + (i.resolution_note || '') + ' ' + (i.lesson_code || '')).toLowerCase();
    var score = scoreHay(hay);
    if (score >= 3) scored.push({ score: score, c: { summary: i.summary || '', resolution_note: i.resolution_note || '',
      lesson_code: i.lesson_code || '', section: i.section || '', issue_type: i.issue_type || '' } });
  });
  // The KnownFixes corpus (backfilled from Chatwoot history, r46) plays by the
  // same scoring rules - it is the memory that predates the tracker.
  knownFixRows_().forEach(function (kf) {
    if (kf.dup_of) return;
    if (cutoff) {
      var kd = kf.resolved_date ? new Date(kf.resolved_date).getTime() : 0;
      if (!kd || kd >= cutoff) return;
    }
    // Out of scope for this course: not a candidate at all (FB-0231).
    var verdict = kfScopeVerdict_(kf.course_scope, course);
    if (verdict === 'no') return;
    var hay = ((kf.problem || '') + ' ' + (kf.fix || '') + ' ' + (kf.lesson_code || '')).toLowerCase();
    var score = scoreHay(hay);
    if (score >= 3) scored.push({ score: score, c: { summary: kf.problem || '', resolution_note: kf.fix || '',
      lesson_code: kf.lesson_code || '', section: '', issue_type: kf.category || '',
      // Carried through to suggestFix_ and out to the screen: an entry offered
      // on a case whose course we don't know still has to show its scope, so
      // whoever reads it can see the condition and rule it out themselves.
      corpus_id: String(kf.conversation_id || ''),
      course_scope: String(kf.course_scope || ''),
      applies_when: String(kf.applies_when || ''),
      scope_uncertain: verdict === 'maybe',
      // An instructor has already said this one was wrong somewhere. Not a
      // reason to hide it (they may have been wrong about it being wrong), but
      // the model should hold a higher bar and the caveat rides along.
      flagged_wrong_before: kfFlagCount_(kf) } });
  });
  scored.sort(function (a, b) { return b.score - a.score; });
  return scored.slice(0, 12).map(function (x) { return x.c; });
}

// The shared "read it all back" core: brief + known-fix lookup off one import.
// cutoffIso comes only from the backtest replay - live callers leave it out.
function caseBriefCore_(imp, cutoffIso, issueId) {
  var got = briefAi_(imp.transcript, issueId);
  if (!got.json) return { error: 'Reading the conversation failed: ' + got.why + '. Try again in a minute.' };
  var brief = got.json;
  var fix = { found: false };
  // The course the student is on, read off the same brief call. When the
  // conversation doesn't say, this stays empty and scoped entries come through
  // carrying their caveat instead of being dropped (FB-0231).
  var briefCourse = kfNormaliseScope_(brief.course);
  try {
    var cands = fixCandidatesFor_((brief.summary || '') + ' ' + String(imp.transcript || '').slice(0, 3000), cutoffIso, false, briefCourse);
    if (cands.length) {
      fix = suggestFix_({ new_issue: { summary: brief.summary || '', course: briefCourse, raw_text: String(imp.transcript || '').slice(0, 4000) },
        candidates: cands, course: briefCourse, _kf_included: true }) || { found: false };
    }
  } catch (e) { fix = { found: false }; }
  return { bj: {
    summary: String(brief.summary || ''),
    device: String(brief.device || ''),
    tried: (brief.tried || []).map(function (t) { return String(t); }).slice(0, 12),
    next: String(brief.next || ''),
    next_why: String(brief.next_why || ''),
    instructor_action: !!brief.instructor_action,
    images_note: String(brief.images_note || ''),
    course: briefCourse,
    fix: fix && fix.found ? String(fix.fix) : '',
    fix_based_on: fix && fix.found ? String(fix.based_on || '') : '',
    // The condition to check before the fix goes anywhere near the student, and
    // which corpus entry to point at if it turns out to be wrong (FB-0231).
    fix_applies_when: fix && fix.found ? String(fix.applies_when || '') : '',
    fix_corpus_id: fix && fix.found ? String(fix.corpus_id || '') : '',
    message_count: imp.message_count,
    images: imp.images || [],
    link: imp.link || '',
    transcript: String(imp.transcript || '').slice(0, 15000)
  } };
}

// One composed round-trip (Edd's call - the app is slow enough already,
// FB-0169): import the conversation (screenshots included since r43.4), then
// return the summary, the tried list, the next step, AND the known-fix lookup
// together. Opens the case if it isn't open yet; re-briefs it if it is.
function caseBrief_(data) {
  var id = chatwootConvId_(data.conversation || data.conversation_id);
  if (!id) return { ok: false, error: 'Pass a Chatwoot conversation link or number.' };
  var imp = chatwootImport_({ conversation: id });
  if (!imp || !imp.ok) return { ok: false, error: (imp && imp.error) || 'Could not read that conversation.' };
  if (!imp.message_count) return { ok: false, error: 'That conversation has no readable messages.' };

  var who = (data._user && data._user.name) || '';
  var now = new Date().toISOString();
  var rec = liveCaseFind_(id);
  if (!rec) rec = { conversation_id: id, opened_by: who, opened_at: now, status: 'open', issue_id: '', draft_count: 0 };
  rec.status = 'open';

  var core = caseBriefCore_(imp, null, rec.issue_id || '');
  if (core.error) return { ok: false, error: core.error };
  var prev = caseBriefJson_(rec);
  core.bj.note_posted = !!prev.note_posted;

  rec.student_name = imp.student_name || rec.student_name || '';
  rec.student_contact = imp.student_contact || rec.student_contact || '';
  rec.last_activity = now;
  rec.last_touched_by = who;
  rec.summary = core.bj.summary;
  rec.unread = false;
  rec.brief_json = JSON.stringify(core.bj);
  liveCaseSave_(rec);
  return { ok: true, conversation_id: id, brief: core.bj,
    student_name: rec.student_name || '', student_contact: rec.student_contact || '',
    issue_id: rec.issue_id || '', opened_by: rec.opened_by || '', last_touched_by: who };
}

// "Check for new reply": re-read the case's own conversation. Anything new
// re-runs the whole brief (Edd: freshness beats pennies), marks the case
// unread for everyone else, and - once an issue is linked - lands as an update
// on it through the same strict two-model verdict the scan and Fetch update
// use, so a wrong "it's fixed" can't slip in.
function caseCheckReply_(data) {
  var id = chatwootConvId_(data.conversation_id || data.conversation);
  var rec = liveCaseFind_(id);
  if (!rec) return { ok: false, error: 'No live case for that conversation.' };
  var prev = caseBriefJson_(rec);
  var imp = chatwootImport_({ conversation: id });
  if (!imp || !imp.ok) return { ok: false, error: (imp && imp.error) || 'Could not read the conversation.' };
  var before = Number(prev.message_count) || 0;
  if (imp.message_count <= before) {
    return { ok: true, found: false, message: 'Nothing new from ' + (rec.student_name || 'the student') + ' yet.' };
  }

  var who = (data._user && data._user.name) || '';
  var core = caseBriefCore_(imp, null, rec.issue_id || '');
  if (core.error) return { ok: false, error: 'New messages arrived, but the re-read failed: ' + core.error };
  core.bj.note_posted = !!prev.note_posted;

  var verdict = '', vNote = '';
  if (rec.issue_id) {
    var found = findRow_(rec.issue_id);
    if (found) {
      var iss = found.record;
      var p = 'A student with an OPEN issue on our sailing course platform has been in touch again. ' +
        'Decide what this conversation says about THAT issue, and nothing else.\n\n' +
        'THE OPEN ISSUE:\n' + JSON.stringify({ summary: iss.summary, status: iss.status, lesson_code: iss.lesson_code, logged: iss.submitted_at }) + '\n\n' +
        'THE CONVERSATION:\n' + String(imp.transcript || '').slice(0, 12000) + '\n\n' +
        'Choose one verdict:\n' +
        '- "fixed": the student says it now works, or confirms the fix or workaround did the job.\n' +
        '- "still_broken": the student says it is still happening, happening again, or worse.\n' +
        '- "new_detail": genuinely new information about the same problem (another device, steps to reproduce, when it started, more students affected).\n' +
        '- "nothing_new": the conversation is about something else entirely, or adds nothing. This is the common answer - use it freely.\n\n' +
        'Return ONLY JSON: {"verdict":"fixed|still_broken|new_detail|nothing_new","note":"<one sentence of what the student actually said>"}. No prose, no fences.';
      var first = anthropicJson_(FINDER_MODEL, p, 400);
      if (first && first.verdict && first.verdict !== 'nothing_new') {
        var vp = 'A first-pass AI read a support conversation and concluded it is an update on a known open issue. Disagree if it is wrong.\n\n' +
          'OPEN ISSUE: ' + JSON.stringify({ summary: iss.summary, status: iss.status }) + '\n' +
          'CLAIMED VERDICT: ' + first.verdict + ' - ' + (first.note || '') + '\n\n' +
          'CONVERSATION:\n' + String(imp.transcript || '').slice(0, 12000) + '\n\n' +
          'Be strict. "fixed" requires the student actually confirming it works now, not an instructor hoping so. ' +
          'If the conversation is really about a different problem, disagree.\n' +
          'Return ONLY JSON: {"agree":true or false,"verdict":"fixed|still_broken|new_detail","note":"<corrected one sentence>"}. No prose, no fences.';
        var second = anthropicJson_(VERIFIER_MODEL, vp, 400);
        if (second && second.agree === true) {
          verdict = second.verdict || first.verdict;
          vNote = second.note || first.note || '';
          var VLABEL = { fixed: 'Student says it now works', still_broken: 'Student says it is still happening', new_detail: 'New detail from the student' };
          try {
            addUpdate_({ issue_id: rec.issue_id, instructor_name: 'Live case',
              summary: (VLABEL[verdict] || 'Update from the student') + ': ' + vNote,
              raw_text: '[Live case reply, Chatwoot conversation ' + id + ']\n' + (VLABEL[verdict] || 'Update') + '.\n' + vNote +
                '\n\nLatest transcript:\n' + String(imp.transcript || '').slice(0, 5000),
              student_name: imp.student_name || '', student_contact: imp.student_contact || '',
              image_urls: (imp.images && imp.images.length) ? imp.images.join(',') : '',
              app_url: data.app_url || getAppUrl_() });
          } catch (e) {}
        }
      }
    }
  }

  rec.last_activity = new Date().toISOString();
  rec.last_touched_by = who;
  rec.summary = core.bj.summary;
  rec.unread = true;
  rec.brief_json = JSON.stringify(core.bj);
  liveCaseSave_(rec);
  var newN = imp.message_count - before;
  return { ok: true, found: true, new_messages: newN, verdict: verdict, note: vNote, brief: core.bj,
    student_name: rec.student_name || '', student_contact: rec.student_contact || '', issue_id: rec.issue_id || '',
    message: newN + ' new message' + (newN === 1 ? '' : 's') + ' from ' + (rec.student_name || 'the student') +
      (verdict ? '. ' + (vNote || verdict.replace(/_/g, ' ')) : '.') };
}

// Draft the next reply in the conversation, in the logged-in instructor's own
// voice where a guide is on file. Copy-paste out (Edd's call): the tracker
// never posts drafts into the Chatwoot thread.
// The ONE reply-draft prompt (r47), shared by the live case draft and the
// backtest's replay draft so the benchmark always measures the prompt we
// actually ship. Reworked on the blind judge's repeated verdicts from the
// 10 Aug backtest: the team's real replies are shorter, commit to the answer,
// and skip the padding - and when the resolution is an action only WE can
// take, the real reply performs it ("I've reset your password"). The draft
// can't press the button, but it CAN be written as if the button is being
// pressed, with the instructor told exactly what to do before sending.
function draftReplyPrompt_(ctx, tail, guide, signName) {
  return 'Draft the next reply in a live support conversation, from an instructor at Ardent Training (an online RYA sailing school) to a student.\n\n' +
    'WHERE THE CONVERSATION STANDS:\n' + JSON.stringify(ctx) + '\n\n' +
    'THE MOST RECENT MESSAGES:\n"""\n' + tail + '\n"""\n\n' +
    (guide ? 'Write it in this instructor\'s own voice. Their style guide:\n"""\n' + guide.slice(0, 30000) + '\n"""\n\n' : '') +
    'HOW TO WRITE IT:\n' +
    '- 50 to 110 words, and closer to 50 when the answer is simple. Our team\'s real replies are short; padding is what makes a draft read like a machine.\n' +
    '- If the known fix or the recommended next step answers the student\'s question, COMMIT to the answer in the first sentence. No hedging around a fact we hold, no "it may be worth checking whether", no restating their problem back at them first.\n' +
    '- One step at most, and no second speculative suggestion "in case". If the fix is known, do not pad it with extra troubleshooting.\n' +
    '- No reassurance the student did not ask for, and no apologising for things that have not gone wrong.\n' +
    '- NEVER state a physical-world or account fact you have not been given: what is in their pack, what their error said, where their parcel is, what their account holds. If the right reply depends on a fact you do not hold, ask the ONE checking question that gets it ("have you checked the other side of the sheet?") rather than confidently arranging a fix - a wrong replacement or a wrong promise costs far more than a short question.\n' +
    '- If the recommended next step is an action WE take rather than the student (resetting their password, assigning a course to their account, marking their exam from photos, posting an answer in the course chat, re-sending an ebook, raising an invoice), write the reply as if that action is being done ("I\'ve reset your password - ...") and put the action itself in instructor_action as one short imperative line. The instructor will do it before sending. If no such action is needed, instructor_action is an empty string.\n' +
    // FB-0240. The draft says "I have reset your password", so the reply is only
    // true if the new password actually gets them in. Checking it is part of the
    // action, not a nicety, and it happens before the message goes.
    '- When that action is setting a password for them, instructor_action must also say to log in with that email and password to check it works first. The draft tells the student the problem is sorted, so somebody has to have proved it is.\n' +
    '- Never promise dates, never invent progress.\n' +
    '- If this reads as the first reply in a while, greet the student by first name; otherwise carry the conversation on naturally. Plain text only, no subject line.\n' +
    (guide
      ? '- Sign off the way this instructor signs off in the style guide (their name: "' + (signName || 'The Ardent team') + '").\n'
      : '- Sign off as "' + (signName || 'The Ardent team') + '".\n') +
    AI_TELLS_RULES_ +
    '\nReturn ONLY JSON, no prose, no fences:\n' +
    '{"message": "<the reply text>",\n' +
    ' "instructor_action": "<what the instructor must DO before sending, one short imperative line, or empty string>"}';
}

function caseDraftReply_(data) {
  var id = chatwootConvId_(data.conversation_id || data.conversation);
  var rec = liveCaseFind_(id);
  if (!rec) return { ok: false, error: 'No live case for that conversation.' };
  var who = (data._user && data._user.name) || '';
  var b = caseBriefJson_(rec);
  var guide = voiceGuideFor_(who);
  var tail = String(b.transcript || '').slice(-2500);

  var prompt = draftReplyPrompt_({
    summary: b.summary || rec.summary || '', device: b.device || '',
    already_tried: b.tried || [], recommended_next_step: b.next || '',
    next_step_is_an_instructor_action: !!b.instructor_action,
    known_fix_from_a_past_issue: b.fix || '',
    student_first_name: String(rec.student_name || '').split(' ')[0]
  }, tail, guide, who);

  // anthropicRaw_ rather than a bare fetch: the draft is JSON now (message +
  // instructor_action), and the 8k budget keeps the thinking-eats-max_tokens
  // trap away from a model that reasons before it writes.
  var got = anthropicRaw_(DRAFT_MODEL, prompt, 8000);
  if (!got.json || !String(got.json.message || '').trim()) {
    return { ok: false, error: 'The draft call failed: ' + (got.why || 'empty draft') };
  }
  rec.draft_count = (Number(rec.draft_count) || 0) + 1;
  rec.last_touched_by = who;
  liveCaseSave_(rec);
  return { ok: true, text: String(got.json.message).trim(),
    action_note: String(got.json.instructor_action || '').trim(), voiced: !!guide };
}

// The checkpoint files the issue WITHOUT closing the case - identical
// behaviour to a normal submit (dedupe, routing, Slack on high), and every
// later checkpoint lands as an update on the same issue so the devs get moving
// while the conversation keeps adding detail. When the AI matches an open case
// the team already has, we ask first (needs_confirm) rather than merging
// silently - same pattern as the same-issue popup on the form.
// Round 65 (Edd, FB-0232). Some conversations are already finished by the time
// anyone presses the button - the answer was given in the chat and the student
// said thanks. Filing that as a fresh open high sends it to a fix queue, pings
// Slack, and drops it into somebody's Actions list, all for a job that is done.
// `outcome` says how it goes in: 'resolved' or 'parked' get the same treatment
// as the log form's Submit and resolve / Submit and park, which the live case
// has never had an equivalent of. It still gets filed either way - Edd wants
// the record, he just doesn't want it treated as live work.
// Filing a conversation that is already dealt with is RECORD KEEPING, not new
// work, and it was costing the same 42 seconds as filing live work: the full
// 26-field extraction, with max_tokens 8192 in case one thread had to split
// into several issues. Measured 19 Aug 2026 on a 3,400-character transcript.
//
// This asks the same model, over the same CACHED syllabus block (so the cache
// can still hit and the lesson codes are still mapped against the real course
// structure), for the handful of fields a record actually needs - and caps the
// output at 700 tokens instead of 8192. The instruction to ignore the big
// output format rides in the VARIABLE half, so the cached half stays
// byte-identical to the full extraction's and both share one cache entry.
//
// It also answers the question the instructor should not have to: was this
// actually FIXED, or just sorted for this one student?
function extractLite_(rawText) {
  var instruction = '\n"""\n\n' +
    'IGNORE the output format described above. This conversation is being filed as a RECORD of something already ' +
    'dealt with, not as new work, so only the fields below are wanted and exactly ONE object is ever returned, ' +
    'however many topics the conversation covers.\n\n' +
    'confirmed_fixed is TRUE only when the underlying fault was genuinely fixed and somebody said so. ' +
    'It is FALSE when the student was helped by hand, given a workaround or a way round it, when it stopped ' +
    'happening on its own, or when it was only ever sorted for this one student. Being helped is not being fixed.\n\n' +
    'Return ONLY JSON: {"category":"tech_issue|course_error|shipping","lesson_code":"<e.g. DS.09.12, or empty>",' +
    '"section":"<or empty>","issue_type":"<or empty>","summary":"<one plain sentence, what the problem was>",' +
    '"confirmed_fixed":true or false,"why":"<one short sentence on how it ended>"}. No prose, no fences.';
  var call = anthropicCachedFetch_(EXTRACTION_MODEL, extractionStaticPrompt_(), rawText + instruction, 700);
  if (!call || !call.res) return { ok: false, error: (call && call.why) || 'no response' };
  var res = call.res;
  if (res.getResponseCode() !== 200) return { ok: false, error: 'Anthropic ' + res.getResponseCode() + ': ' + String(res.getContentText() || '').slice(0, 200) };
  var parsed;
  try { parsed = JSON.parse(res.getContentText()); } catch (e) { return { ok: false, error: 'unreadable reply' }; }
  tallyAi_(parsed);
  var text = '';
  if (parsed.content && parsed.content.length) {
    for (var i = 0; i < parsed.content.length; i++) {
      if (parsed.content[i].type === 'text') text += parsed.content[i].text;
    }
  }
  var out = null;
  try { out = JSON.parse(String(text).replace(/^```json?\s*|\s*```$/g, '').trim()); } catch (e) { out = null; }
  if (!out) {
    // Forgive the two classic sins the full extraction already forgives: a
    // chatty preamble, and a truncated tail.
    var m = String(text).match(/\{[\s\S]*\}/);
    if (m) { try { out = JSON.parse(m[0]); } catch (e2) { out = null; } }
  }
  if (!out) return { ok: false, error: 'could not read the reply as JSON' };
  return { ok: true, fields: out, cached: !!call.cached };
}

function caseCheckpoint_(data) {
  var id = chatwootConvId_(data.conversation_id || data.conversation);
  var rec = liveCaseFind_(id);
  if (!rec) return { ok: false, error: 'No live case for that conversation.' };
  var who = (data._user && data._user.name) || '';
  var outcome = String(data.outcome || '').toLowerCase();
  if (outcome !== 'resolved' && outcome !== 'parked') outcome = '';
  var outNote = String(data.resolution_note || '').trim();
  // A resolved issue with no note is a closed one nobody can learn from, and
  // that note is what the next student with the same fault ends up getting.
  if (outcome && !outNote) {
    return { ok: false, error: outcome === 'resolved'
      ? 'Put what actually sorted it in the box first - that answer is what the next student with the same fault gets.'
      : 'Say why it is being parked - a parked issue with no reason is just an open one nobody looks at.' };
  }
  // FB-0239. "It took forever, just spinning and spinning." Forever turned out
  // to be four things in a row, and nobody could see which one was the slow one,
  // so every press now returns how long each part took. A number in the response
  // beats a guess about a spinner somebody has already walked away from.
  var T = { t0: Date.now() }, TT = {};
  function lap(k) { TT[k] = Date.now() - T.t0; T.t0 = Date.now(); }

  var imp = chatwootImport_({ conversation: id });
  lap('read_conversation');
  if (!imp || !imp.ok) return { ok: false, error: (imp && imp.error) || 'Could not read the conversation.' };
  var now = new Date().toISOString();
  var bj = caseBriefJson_(rec);

  // Later checkpoints: the issue exists, so the fresh state lands as an update.
  if (rec.issue_id) {
    var upd = {
      issue_id: rec.issue_id, instructor_name: who,
      summary: (outcome === 'resolved' ? 'Sorted in the live case'
        : outcome === 'parked' ? 'Parked from the live case' : 'Checkpoint from the live case') +
        (bj.summary ? ': ' + String(bj.summary).slice(0, 140) : ''),
      raw_text: '[Live case checkpoint, Chatwoot conversation ' + id + ']\n' + (bj.summary || '') +
        (outNote ? '\n\nOutcome: ' + outNote : '') +
        '\n\nLatest transcript:\n' + String(imp.transcript || '').slice(0, 5000),
      student_name: imp.student_name || '', student_contact: imp.student_contact || '',
      image_urls: (bj.images_note ? '' : ((imp.images && imp.images.length) ? imp.images.join(',') : '')),  // coursework screenshots stay off the report (Edd, FB-0179)
      app_url: data.app_url || getAppUrl_(), _user: data._user };
    if (outcome === 'resolved') { upd.resolved = true; upd.resolution_note = outNote; }
    else if (outcome === 'parked') { upd.park = true; upd.resolution_note = outNote; }
    var r0 = addUpdate_(upd);
    if (!r0 || !r0.ok) return { ok: false, error: 'Could not add the update: ' + ((r0 && r0.error) || 'unknown') };
    bj.message_count = imp.message_count;
    rec.brief_json = JSON.stringify(bj);
    rec.last_activity = now; rec.last_touched_by = who;
    liveCaseSave_(rec);
    lap('add_update');
    return { ok: true, updated: true, issue_id: rec.issue_id,
      resolved: outcome === 'resolved', parked: outcome === 'parked', timings: TT };
  }

  // First checkpoint: extract, check for an open case the team is already on
  // (ask before merging), then file through the normal submit path.
  // Filing something already dealt with is record keeping, so it takes the light
  // path: the same cached syllabus block, a handful of fields, and 700 output
  // tokens instead of 8192. Filing live work still gets the full extraction,
  // because those fields are what route it to the right queue.
  var lite = null, ex = null, f = {};
  if (outcome === 'resolved' || outcome === 'parked') {
    lite = extractLite_(imp.transcript);
    lap('extract_lite');
    if (!lite.ok) {
      // Never lose the filing over the fast path. Fall back to the full one and
      // say so, rather than dropping a conversation somebody has finished with.
      lite = null;
      ex = extract_({ raw_text: imp.transcript });
      lap('extract_full_fallback');
      TT.lite_failed = 1;
    }
  } else {
    ex = extract_({ raw_text: imp.transcript });
    lap('extract');
  }
  if (ex && ex.diag) TT.extract_cache_read = ex.diag.cache_read;
  if (!lite && (!ex || !ex.ok)) return { ok: false, error: 'The extraction failed: ' + ((ex && ex.error) || 'unknown') };
  f = lite ? (lite.fields || {}) : (ex.fields || {});
  var category = String(f.category || 'tech_issue').toLowerCase();
  var payload = {
    _user: data._user,
    category: category,
    raw_text: imp.transcript || '',
    student_name: f.student_name || imp.student_name || '',
    student_contact: f.student_contact || imp.student_contact || '',
    device_info: f.device_info || '',
    course: f.course || '', module: f.module || '', lesson: f.lesson || '',
    lesson_code: f.lesson_code || '',
    issue_type: f.issue_type || '',
    summary: f.summary || bj.summary || '',
    priority: f.priority || 'medium',
    priority_reason: f.priority_reason || '',
    image_urls: (bj.images_note ? '' : (imp.images || []).join(',')),  // coursework screenshots stay off the report (Edd, FB-0179)
    section: f.section || '',
    platform: f.platform || '',
    media_kind: f.media_kind || '',
    request_kind: f.request_kind === 'improvement' ? 'improvement' : 'fix',
    chatwoot_conversation_id: id,
    app_url: data.app_url || getAppUrl_()
  };

  // Going in already sorted, or already stalled. addIssue_ reads these the same
  // way the log form's two buttons do: status and resolved_at set on the row,
  // the routing block skipped (nothing to hand to a fix queue), and no Slack,
  // because the alert exists to make somebody drop what they are doing.
  // Edd, 19 Aug 2026: the instructor picks Resolved or Still open, and parked is
  // worked out rather than chosen. A conversation sorted only for THIS student -
  // a password reset by hand, a workaround, something that stopped on its own -
  // is not a fixed fault, and filing it resolved buries it. Parked keeps it out
  // of the fix queues while leaving it there for the next report to be linked
  // to, which is the whole point of the status.
  var autoParked = false;
  if (outcome === 'resolved') {
    var reallyFixed = lite ? (f.confirmed_fixed === true) : true;   // full path cannot tell, so take the instructor at their word
    if (reallyFixed) {
      payload.resolved = true;
      payload.resolution_note = outNote;
      payload.resolved_by = who;
    } else {
      autoParked = true;
      payload.parked = true;
      payload.resolution_note = outNote +
        '\n\n[Parked rather than resolved: sorted for this student, but the fault itself was not confirmed fixed' +
        (f.why ? ' - ' + String(f.why) : '') + '. It stays here for a later report to be linked to.]';
    }
  } else if (outcome === 'parked') {
    payload.parked = true;
    payload.resolution_note = outNote;
  }
  // The extraction reads urgency off the words in the transcript, and a student
  // who was stuck writes an urgent-sounding message even when the thing got
  // sorted three replies later. High means "drop everything and look at this",
  // so a finished conversation cannot be high - there is nothing left to drop
  // anything for. Down a notch, with the reason kept so the stats still say
  // what it felt like at the time (Edd, FB-0232).
  if (outcome === 'resolved' && String(payload.priority).toLowerCase() === 'high') {
    payload.priority_reason = 'Filed from a conversation that was already sorted, so it is a record rather than live work. ' +
      'It read as high priority while it was happening' + (payload.priority_reason ? ': ' + payload.priority_reason : '.');
    payload.priority = 'medium';
  }

  if (data.merge_into) {
    payload.merge_into = data.merge_into;
  } else if (data.no_merge) {
    payload.no_merge = true;
  } else {
    var matchId = null;
    try { matchId = aiMatchIssue_(payload, category); } catch (e) { matchId = null; }
    lap('duplicate_check');
    if (matchId) {
      var m = findRow_(matchId);
      if (m) {
        return { ok: true, needs_confirm: true, match: {
          issue_id: matchId, summary: m.record.summary || '', status: m.record.status || '',
          lesson_code: m.record.lesson_code || '', priority: m.record.priority || '',
          report_count: Number(m.record.report_count) || 1
        }, fields: { summary: payload.summary, priority: payload.priority }, timings: TT };
      }
    }
    payload.no_merge = true; // the matcher already said no; no point running it twice inside addIssue_
  }

  var r = addIssue_(payload);
  lap('file_the_issue');
  if (!r || !r.ok) return { ok: false, error: 'Could not file the issue: ' + ((r && r.error) || 'unknown') };
  var issueId = (r.issue && r.issue.issue_id) || '';
  rec.issue_id = issueId;
  // A new issue logged from a chat gets the one private note at submit time; a
  // merge doesn't, so the close step knows whether it still owes the note.
  bj.note_posted = !r.merged;
  bj.message_count = imp.message_count;
  rec.summary = payload.summary || rec.summary;
  rec.brief_json = JSON.stringify(bj);
  rec.last_activity = now; rec.last_touched_by = who;
  liveCaseSave_(rec);
  return { ok: true, filed: true, merged: !!r.merged, issue_id: issueId,
    summary: payload.summary, priority: String((r.issue && r.issue.priority) || payload.priority || ''),
    resolved: outcome === 'resolved' && !autoParked && !r.merged,
    parked: (outcome === 'parked' || autoParked) && !r.merged,
    auto_parked: autoParked && !r.merged, auto_parked_why: autoParked ? String(f.why || '') : '',
    // A merge deliberately does NOT resolve or park the issue it joins: that
    // fault is still live for everybody else on it, and one student getting
    // sorted is no reason to stop work (the same rule addReportToIssue_ keeps
    // for a "Submit and park" that merges).
    merged_stays_open: !!r.merged && !!outcome, timings: TT };
}

// Manual close - for the conversations that fizzle out, or once everything is
// wrapped. Close-with-issue leaves the single private note on the Chatwoot
// conversation as the record, unless the checkpoint already posted it (one
// note, never two).
function caseClose_(data) {
  var id = chatwootConvId_(data.conversation_id || data.conversation);
  var rec = liveCaseFind_(id);
  if (!rec) return { ok: false, error: 'No live case for that conversation.' };
  var who = (data._user && data._user.name) || '';
  var bj = caseBriefJson_(rec);
  if (rec.issue_id && !bj.note_posted) {
    var found = findRow_(rec.issue_id);
    if (found) {
      try { chatwootNote_(id, found.record, data.app_url || getAppUrl_()); bj.note_posted = true; } catch (e) {}
    }
  }
  rec.status = 'closed';
  rec.last_activity = new Date().toISOString();
  rec.last_touched_by = who;
  rec.brief_json = JSON.stringify(bj);
  liveCaseSave_(rec);
  return { ok: true, closed: true, issue_id: rec.issue_id || '' };
}

// Opening a case clears the sheet-level unread flag (the per-person dot is
// handled in the browser). Deliberately does NOT count as touching the case.
function caseTouch_(data) {
  var rec = liveCaseFind_(chatwootConvId_(data.conversation_id || data.conversation));
  if (!rec) return { ok: true };
  if (isTrueLike_(rec.unread)) { rec.unread = false; liveCaseSave_(rec); }
  return { ok: true };
}

// ---- batch "tell the students" (Round 45) ----------------------------------
// One click on a resolved issue drafts the "it's fixed" message for EVERY
// affected student across the merged reports, each in the resolving
// instructor's voice. Drafts only - nothing sends itself, ever.
function batchStudentDrafts_(data) {
  var found = findRow_(data.issue_id);
  if (!found) return { ok: false, error: 'No issue found with that id.' };
  var i = found.record;
  var who = (data._user && data._user.name) || '';
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'no api key' };

  var seen = {};
  var students = [];
  function addSt(name, contact) {
    name = String(name || '').trim(); contact = String(contact || '').trim();
    if (!name && !contact) return;
    var key = (name + '|' + contact).toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    students.push({ name: name || 'there', contact: contact });
  }
  addSt(i.student_name, i.student_contact);
  var reps = []; try { reps = i.reports_json ? JSON.parse(i.reports_json) : []; } catch (e) { reps = []; }
  reps.forEach(function (rp) {
    if (rp.kind === 'question' || rp.kind === 'answer') return;
    addSt(rp.student_name, rp.student_contact);
  });
  if (!students.length) return { ok: false, error: 'No student names or contacts on this issue.' };
  var capped = students.length > 10;
  students = students.slice(0, 10);

  var guide = voiceGuideFor_(who);
  var fixNotes = i.dev_notes || i.resolution_note || '';
  var drafts = [];
  for (var s = 0; s < students.length; s++) {
    var st = students[s];
    var prompt = 'Draft a short, warm email from an instructor at Ardent Training (an online RYA sailing school) to a student.\n\n' +
      'GOAL: Tell the student the problem they reported has been fixed, and invite them to try again and shout if anything still looks off.\n\n' +
      'THE ISSUE:\n' + JSON.stringify({ summary: i.summary, lesson: i.lesson || i.lesson_code || '', fix_notes: fixNotes,
        student_first_name: String(st.name).split(' ')[0] }) + '\n' +
      (guide ? '\nWrite it in this instructor\'s own voice. Their style guide:\n"""\n' + guide.slice(0, 30000) + '\n"""\n' : '') +
      AI_TELLS_RULES_ +
    '\nRules: plain text only, no subject line, 60-140 words, greet the student by first name. ' +
      (guide
        ? 'Sign off exactly the way this instructor signs off in the style guide (their name: "' + (who || 'The Ardent team') + '"). '
        : 'Sign off as "' + (who || 'The Ardent team') + '". ') +
      'Promise no dates unless the fix notes give one. Return ONLY the email text, nothing else.';
    var text = '';
    try {
      var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({ model: DRAFT_MODEL, max_tokens: 1500, messages: [{ role: 'user', content: prompt }] })
      });
      if (res.getResponseCode() >= 200 && res.getResponseCode() < 300) {
        var parsed = JSON.parse(res.getContentText() || '{}');
        (parsed.content || []).forEach(function (c) { if (c.type === 'text') text += c.text; });
        text = text.trim();
      }
    } catch (e) { text = ''; }
    drafts.push({ name: st.name, contact: st.contact, text: text, failed: !text });
  }
  return { ok: true, voiced: !!guide, capped: capped, total_students: Object.keys(seen).length, drafts: drafts };
}

// ---- confusion -> content tweak (Round 45) ---------------------------------
// Three or more open student-confusion reports on one lesson is not three
// confused students, it is a sentence in the lesson doing the confusing.
// Draft what would stop it - as a SUGGESTION the course team can accept or
// dismiss, never applied by itself (same rule as the playbook suggestions).
function contentSuggestions_() {
  var raw = PropertiesService.getScriptProperties().getProperty('CONTENT_SUGGESTIONS');
  if (!raw) return [];
  try { var a = JSON.parse(raw); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}
function saveContentSuggestions_(arr) {
  PropertiesService.getScriptProperties().setProperty('CONTENT_SUGGESTIONS', JSON.stringify((arr || []).slice(-30)));
}
function contentSeen_() {
  var raw = PropertiesService.getScriptProperties().getProperty('CONTENT_SUGG_SEEN');
  if (!raw) return {};
  try { return JSON.parse(raw) || {}; } catch (e) { return {}; }
}
function confusionReview_() {
  var byLesson = {};
  (getIssues_().issues || []).forEach(function (i) {
    if (String(i.issue_type || '') !== 'student_confusion') return;
    var st = String(i.status || 'open').toLowerCase();
    if (st !== 'open' && st !== 'in_progress' && st !== 'with_dev') return;
    var code = String(i.lesson_code || '').trim();
    if (!code) return;
    (byLesson[code] = byLesson[code] || []).push(i);
  });
  var pending = contentSuggestions_();
  var seen = contentSeen_();
  var added = 0, hotspots = 0;
  Object.keys(byLesson).forEach(function (code) {
    var group = byLesson[code];
    if (group.length < 3) return;
    hotspots++;
    var already = pending.some(function (s) { return s.lesson_code === code; });
    if (already) return;
    // Only re-suggest once the pile has grown past the last accept/dismiss.
    if (seen[code] && group.length <= Number(seen[code])) return;
    var items = group.slice(0, 6).map(function (i) {
      return { summary: String(i.summary || '').slice(0, 200), detail: String(i.raw_text || '').slice(0, 600) };
    });
    var prompt = 'You are reviewing a lesson (code ' + code + ') in an online RYA sailing theory course for Ardent Training. ' +
      group.length + ' separate students have reported being confused by this lesson. Their reports:\n' +
      JSON.stringify(items) + '\n\n' +
      'Work out what the students keep misreading, and draft ONE practical content tweak the course team could make. ' +
      'The shape is: "students keep reading X as Y - a sentence on Z would stop it". Be concrete about the fix ' +
      '(what to add, reword, or illustrate) and keep it to two or three sentences. ' +
      'If the reports are genuinely about different things and no single tweak would help, return found false.\n' +
      'Return ONLY JSON, no prose, no fences: {"found": true or false, "suggestion": "<the drafted tweak>", "why": "<one sentence on the pattern in the reports>"}';
    var got = anthropicRaw_(ANTHROPIC_MODEL, prompt, 16000);
    var out = got.json;
    if (!out || !out.found || !out.suggestion) return;
    pending.push({ id: Utilities.getUuid(), lesson_code: code, count: group.length,
      suggestion: String(out.suggestion), why: String(out.why || ''),
      issue_ids: group.map(function (i) { return i.issue_id; }).slice(0, 12),
      created_at: new Date().toISOString() });
    added++;
  });
  if (added) saveContentSuggestions_(pending);
  return { ok: true, hotspots: hotspots, added: added, pending: pending.length };
}
function runConfusionReview_(data) {
  var r = confusionReview_();
  r.suggestions = contentSuggestions_();
  return r;
}
function listContentSuggestions_() { return { ok: true, suggestions: contentSuggestions_() }; }
function resolveContentSuggestion_(body) {
  var arr = contentSuggestions_();
  var kept = [], matched = null;
  for (var i = 0; i < arr.length; i++) {
    if (arr[i].id === body.id) matched = arr[i]; else kept.push(arr[i]);
  }
  if (!matched) return { ok: false, error: 'Suggestion not found (it may have already been actioned).' };
  var issueId = '';
  if (body.accept) {
    // Accepting queues the tweak as a course improvement on that lesson, so it
    // lands in the same course-fixes queue as everything else. Making the
    // change is still a human job - nothing touches lesson content from here.
    var r = addIssue_({
      _user: body._user,
      category: 'course_error',
      request_kind: 'improvement',
      lesson_code: matched.lesson_code,
      issue_type: 'student_confusion',
      summary: 'Content tweak: ' + String(matched.suggestion).slice(0, 240),
      raw_text: 'Drafted from ' + matched.count + ' student-confusion reports on ' + matched.lesson_code + '.\n\n' +
        'Suggested tweak: ' + matched.suggestion + '\n\nWhy: ' + (matched.why || '') +
        '\n\nSource issues: ' + (matched.issue_ids || []).join(', '),
      priority: 'medium',
      priority_reason: matched.count + ' students confused by the same lesson',
      no_merge: true
    });
    issueId = (r && r.ok && r.issue && r.issue.issue_id) || '';
  }
  var seen = contentSeen_();
  seen[matched.lesson_code] = matched.count;
  PropertiesService.getScriptProperties().setProperty('CONTENT_SUGG_SEEN', JSON.stringify(seen));
  saveContentSuggestions_(kept);
  return { ok: true, accepted: !!body.accept, issue_id: issueId };
}

function backendInfo_() {
  var p = PropertiesService.getScriptProperties();
  return {
    stamp: CODE_STAMP,
    version: p.getProperty('BACKEND_VERSION') || '',
    deployed_at: p.getProperty('BACKEND_DEPLOYED_AT') || '',
    note: p.getProperty('BACKEND_NOTE') || '',
    // What is allowed to reach Slack and where it goes, so it can be checked
    // without opening the code. Muting something and being unable to confirm it
    // is how a channel quietly starts up again. Reports whether each
    // destination property is SET, never the webhook itself.
    slack: (function () {
      var out = {};
      Object.keys(SLACK_NOTICES).forEach(function (k) {
        var n = SLACK_NOTICES[k];
        out[k] = { on: n.on, to: n.to || '(current channel)',
                   ready: !n.to || !!p.getProperty(n.to) };
      });
      return out;
    })()
  };
}

// ---- Auto-resolve TBC issues ----------------------------------------------

// Daily job: any issue sitting at "Resolved - TBC" with no activity for
// TBC_AUTO_RESOLVE_DAYS days is taken as fixed and moved to Resolved. The timer
// is updated_at, so a fresh report or a reopen resets the clock.
function autoResolveTbc() {
  var cutoff = Date.now() - TBC_AUTO_RESOLVE_DAYS * 24 * 3600 * 1000;
  var resolved = 0;
  ISSUE_SHEETS.forEach(function (name) {
    var sheet = sheetByName_(name);
    if (!sheet) return;
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return;
    var head = values[0];
    var idx = {}; head.forEach(function (h, i) { idx[h] = i; });
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if (!row[idx['issue_id']]) continue;
      if (String(row[idx['status']]).toLowerCase() !== 'resolved_tbc') continue;
      var updated = new Date(row[idx['updated_at']] || row[idx['submitted_at']]);
      if (isNaN(updated.getTime()) || updated.getTime() > cutoff) continue;
      var now = new Date().toISOString();
      sheet.getRange(r + 1, idx['status'] + 1).setValue('resolved');
      sheet.getRange(r + 1, idx['resolved_at'] + 1).setValue(now);
      sheet.getRange(r + 1, idx['updated_at'] + 1).setValue(now);
      resolved++;
    }
  });
  Logger.log('autoResolveTbc: resolved ' + resolved + ' TBC issue(s).');
  return resolved;
}

// Run this once from the editor to schedule autoResolveTbc to run every day.
// Safe to run again; it removes any existing copy first so you never end up
// with duplicate triggers.
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'autoResolveTbc') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoResolveTbc').timeBased().everyDays(1).atHour(2).create();
  Logger.log('Daily auto-resolve trigger installed (runs around 2am).');
}

// ---- One-time setup -------------------------------------------------------

/**
 * Run this once from the Apps Script editor (select setup, press Run).
 * It creates the Course Errors, Tech Issues, and Instructors tabs if missing,
 * writes the header rows, and pre-populates the instructor names. Safe to run
 * again; it will not wipe existing rows.
 */
// Round 61. Invite tokens made before this round have no expiry baked in, so
// tokenExpired_ reads them as good forever - a year-old invite email is still a
// way into an account. New ones carry an expiry (newInviteToken_), but the old
// ones already sitting in the sheet have to be cleared by hand, and clearing one
// is the only safe move: we cannot add an expiry to a token someone already has
// in their inbox without changing it, which breaks their link anyway.
//
// So: any outstanding invite token with no expiry is wiped. Anyone caught by it
// gets a fresh link from Resend on the Users page. Logs who, so there is a list
// to work from rather than a silent change. Runs from setup(), and is safe to
// run again - once they are cleared there is nothing left to find.
function expireLegacyInviteTokens_() {
  var sheet = usersSheet_();
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0]; var idx = {};
  head.forEach(function (h, i) { idx[h] = i; });
  if (idx['invite_token'] == null) return [];
  var cleared = [];
  for (var r = 1; r < values.length; r++) {
    var tok = String(values[r][idx['invite_token']] || '');
    // A dot means it already carries its own expiry, so leave it be.
    if (!tok || tok.indexOf('.') > -1) continue;
    sheet.getRange(r + 1, idx['invite_token'] + 1).setValue('');
    cleared.push(String(values[r][idx['email']] || '(no email)'));
  }
  if (cleared.length) {
    Logger.log('Round 61: cleared ' + cleared.length + ' invite link(s) that had no expiry. ' +
      'These people need a fresh one via Resend on the Users page: ' + cleared.join(', '));
  } else {
    Logger.log('Round 61: no invite links without an expiry were outstanding.');
  }
  return cleared;
}

function setup() {
  var ss = ss_();

  ISSUE_SHEETS.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    // A sheet only has the columns it was given. Appending a header past the
    // last one throws rather than growing the sheet, and every row write goes
    // out HEADERS.length wide, so make the room first (Round 55).
    if (sheet.getMaxColumns() < HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
    }
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    // Tracking references and chase dates are text, not maths. Left to itself
    // Sheets turns a long consignment number into scientific notation and a
    // chase date into a Date object an hour behind, so we pin both to plain
    // text before any row lands.
    // chatwoot_contact_id joins the list in Round 61 for the same reason: it is
    // digits, so Sheets treats it as a number, and a long enough one comes back
    // in scientific notation with the end of it rounded off. An identifier that
    // quietly loses its last few digits is worse than not having one. Pinned to
    // text while the column is still empty, so nothing needs repairing later.
    ['tracking_number', 'chase_at', 'chatwoot_contact_id'].forEach(function (col) {
      var c = HEADERS.indexOf(col);
      if (c < 0) return;
      sheet.getRange(1, c + 1, sheet.getMaxRows(), 1).setNumberFormat('@');
    });
  });

  var instructors = ss.getSheetByName(INSTRUCTORS_SHEET);
  if (!instructors) instructors = ss.insertSheet(INSTRUCTORS_SHEET);
  instructors.getRange(1, 1, 1, 2).setValues([['name', 'email']]);
  instructors.setFrozenRows(1);
  instructors.getRange(1, 1, 1, 2).setFontWeight('bold');

  var existingRows = instructors.getDataRange().getValues();
  var nameToRow = {};
  for (var i = 1; i < existingRows.length; i++) {
    nameToRow[String(existingRows[i][0]).trim()] = i + 1;
  }
  INSTRUCTORS.forEach(function (ins) {
    if (nameToRow[ins.name]) {
      instructors.getRange(nameToRow[ins.name], 2).setValue(ins.email); // refresh email
    } else {
      instructors.appendRow([ins.name, ins.email]);
    }
  });

  // Feedback tab for suggestions/bugs about the tracker itself.
  var fb = ss.getSheetByName(FEEDBACK_SHEET);
  if (!fb) fb = ss.insertSheet(FEEDBACK_SHEET);
  fb.getRange(1, 1, 1, FEEDBACK_HEADERS.length).setValues([FEEDBACK_HEADERS]);
  fb.setFrozenRows(1);
  fb.getRange(1, 1, 1, FEEDBACK_HEADERS.length).setFontWeight('bold');

  // Users tab for accounts/logins.
  var usersS = ss.getSheetByName(USERS_SHEET);
  if (!usersS) usersS = ss.insertSheet(USERS_SHEET);
  usersS.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]);
  usersS.setFrozenRows(1);
  usersS.getRange(1, 1, 1, USER_HEADERS.length).setFontWeight('bold');

  // Seed the first admin (Edd) as an invited account with every permission, so
  // there is someone who can log in and invite the rest.
  if (!findUserByEmail_('ehewett@ardent-training.com')) {
    var allPerms = {}; PERM_KEYS.forEach(function (k) { allPerms[k] = true; });
    var token = newInviteToken_();
    var seed = {
      email: 'ehewett@ardent-training.com', name: 'Edd', status: 'invited',
      perms_json: JSON.stringify(allPerms), pass_hash: '', pass_salt: '',
      invite_token: token, session_token: '', session_expires: '',
      created_at: new Date().toISOString()
    };
    usersS.appendRow(USER_HEADERS.map(function (k) { return seed[k]; }));
    var url = getAppUrl_();
    Logger.log('Seeded admin ehewett@ardent-training.com. Set your password at: ' +
      (url || '(set APP_URL first, then run adminInviteLink)') +
      ((url && url.indexOf('?') > -1) ? '&' : '?') + 'invite=' + token);
  }

  // Time-based triggers for the recheck reminders and the monthly checklist
  // review. Created once; re-running setup() won't duplicate them.
  try { ensureTriggers_(); } catch (e) { Logger.log('Trigger creation failed (create them by hand if needed): ' + e); }

  // Round 61: clear any invite link still outstanding from before invites had
  // an expiry. Logs who needs a fresh one.
  var clearedInvites = [];
  try { clearedInvites = expireLegacyInviteTokens_(); } catch (e) { Logger.log('Invite sweep failed: ' + e); }

  Logger.log('Setup complete. Course Errors, Tech Issues, Instructors, Users, and Feedback sheets are ready. ' +
    'If APP_URL is not set yet, set it after deploying then run adminInviteLink() to get your setup link.' +
    (clearedInvites.length ? ' NOTE: ' + clearedInvites.length + ' invite link(s) with no expiry were cleared - resend to: ' + clearedInvites.join(', ') : ''));
}

/**
 * Quick check you can run from the editor to confirm read and write work.
 * Appends a throwaway test issue to the Course Errors tab, then reads it back.
 * Delete that test row by hand afterwards.
 */
function selfTest() {
  var before = getIssues_().issues.length;
  var added = addIssue_({
    instructor_name: 'Edd',
    category: 'course_error',
    raw_text: 'SELF TEST - delete me',
    summary: 'Self test row',
    priority: 'low',
    issue_type: 'other',
    status: 'open'
  });
  var after = getIssues_().issues.length;
  Logger.log('Rows before: ' + before + ', after: ' + after +
    '. Added id: ' + added.issue.issue_id +
    '. Delete that test row from the Course Errors sheet once you are happy.');
}

// ---- One-off historical import ---------------------------------------------
//
// Pulls every report out of the two OLD spreadsheets (the Google Form responses
// and the manual trackers) and writes them into this tracker's sheets. Run
// importHistoricalIssues() once from the editor. It is guarded: if it finds
// previously imported rows it stops, so running it twice cannot duplicate data.
// It never posts to Slack and never calls the AI dedupe.
//
// Status mapping (agreed with Edd, 11 July 2026):
//   Course errors form:  Done "x" -> resolved, otherwise -> open
//   Tech issues form:    everything -> past (no resolution record was kept)
//   Product Updates/SB:  Completed -> resolved, everything else -> open
//
// Expected result (verified against an offline dry run of the same mapping):
//   864 rows total: course_error 539 (253 resolved / 286 open),
//   tech_issue 171 (all past), internal 154.

var IMPORT_COURSE_ERRORS_ID = '1GZxSuryrlIMdI_BiFMXOSZB3XYWncgrUpWX4Le6mMcw';
var IMPORT_TECH_ID = '10JfcQ159gwiJX40eHOf4LJXLhMe_pwQou8SqfRVO8lA';

var IMPORT_MODULES = {
  EN: { 1:'Welcome Aboard',2:"Cynthia's Boat",3:'Safety at Sea',4:'Understanding a Passage Plan',5:'Checking the Weather',6:'Leaving the Harbour',7:'Setting Course',8:'What Can We See?',9:'The Groats',10:'Going with the Flow',11:'Avoiding Collision',12:'Communicating with Other Vessels',13:"Approaching St Anthony's",14:'A Picnic at Anchor',15:'Returning Home',16:'Debrief' },
  DS: { 1:'Welcome To Ardent Training',2:'Nautical Terms',3:'Introduction to Navigation',4:'Navigation Instruments',5:'Position Fixing',6:'Ropework',7:'Tidal Heights',8:'Anchorwork',9:'Tidal Streams',10:'IRPCS',11:'Meteorology',12:'Safety',13:'Pilotage',14:'Passage Planning',15:'Marine Environment',16:'Final Exam',17:'Final Exam (assessment)' },
  YM: { 1:'Welcome',2:'Navigation Instruments',3:'Tidal Heights',4:'Tidal Streams',5:'Position Fixing and Chartwork',6:'Meteorology',7:'IRPCS',8:'Pilotage',9:'Safety',10:'Passage Planning',11:'Marine Environment',12:'Exams',13:'Chartwork and IRPCS (exam)',14:'Appraisal',15:'Passage Making' },
  FT: { 1:'Welcome',2:'Foundations of Navigation',3:'Navigation Instruments',4:'Position Fixing',5:'Tidal Heights',6:'Anchorwork',7:'Ropework',8:'Tidal Streams',9:'Chartwork',10:'Meteorology',11:'IRPCS',12:'Pilotage',13:'Safety',14:'Passage Planning',15:'Marine Environment',16:'Exams',17:'Final Part 1',18:'Appraisal',19:'Final Part 2' }
};
var IMPORT_COURSE_NAME = { EN:'Essential Navigation', DS:'Day Skipper', YM:'Yachtmaster', FT:'Fast Track' };

var IMPORT_NAME_MAP = {
  'edd':'Edd','edd hewett':'Edd','charly':'Charly','charly hewett':'Charly','charlie':'Charlie',
  'charlie triggs':'Charlie','ct':'Charlie','stu':'Stuart','stuart':'Stuart','tom':'Tom',
  'michelle':'Michelle','michelle skelton':'Michelle','laura':'Laura','luke':'Luke','holly':'Holly',
  'peter':'Peter','peter king':'Peter','charly / laura':'Charly','lauren':'Lauren','lauren hewett':'Lauren'
};

// Slide references the parser cannot safely read, fixed by hand after eyeballing
// the source rows. Keyed by spreadsheet row number in Form responses 2.
var IMPORT_CE_OVERRIDES = {
  368: ['Yachtmaster','Chartwork and IRPCS (exam)','YM.13'],   // 'YN.13.10.18 irpcsexam' (typo)
  228: ['Essential Navigation','Returning Home','EN.15'],      // 'Returning Home quiz - EN'
  341: ['Day Skipper','Final Exam','DS.16'],                   // 'Day Skipper Theory Exam'
  241: ['Yachtmaster','',''],                                  // 'Yachtmaster course'
  503: ['Essential Navigation','','']                          // 'Essential Nav'
};

function impClean_(v) { return v == null ? '' : String(v).replace(/\s+/g, ' ').trim(); }

function impIso_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v.toISOString();
  var s = impClean_(v);
  var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2}):(\d{2}))?/);
  if (m) return m[3] + '-' + m[2] + '-' + m[1] + 'T' + (m[4]||'00') + ':' + (m[5]||'00') + ':' + (m[6]||'00') + 'Z';
  return '';
}

function impSummarise_(text, fallback) {
  var t = impClean_(text);
  if (!t) return fallback || '';
  if (t.length <= 140) return t;
  var cut = t.slice(0, 140);
  var sp = cut.lastIndexOf(' ');
  return (sp > 0 ? cut.slice(0, sp) : cut) + '…';
}

function impName_(raw, fallback) {
  var d = fallback === undefined ? 'Imported' : fallback;
  var k = impClean_(raw).toLowerCase();
  if (!k) return d;
  if (IMPORT_NAME_MAP[k]) return IMPORT_NAME_MAP[k];
  var first = k.split(/[\s\/]+/)[0];
  if (IMPORT_NAME_MAP[first]) return IMPORT_NAME_MAP[first];
  return d;
}

// Parse a free-text slide/lesson reference into [course, module title, lesson code].
function impParseLesson_(ref) {
  var s = impClean_(ref);
  if (!s) return ['','',''];
  var m = s.match(/\b(EN|DS|YM|FT)[\s.]*?(\d{1,2})(?:[\s.]+(\d{1,3}))?(?:[\s.]+\d{1,3})?/i);
  if (!m) {
    var w = ' ' + s.toUpperCase() + ' ';
    for (var pfx in IMPORT_COURSE_NAME) {
      if (w.indexOf(' ' + pfx + ' ') === 0 || w.match(new RegExp('^\\s*' + pfx + '\\b'))) {
        return [IMPORT_COURSE_NAME[pfx], '', ''];
      }
    }
    return ['','',''];
  }
  var pfx2 = m[1].toUpperCase();
  var mod = parseInt(m[2], 10);
  var les = m[3] ? parseInt(m[3], 10) : null;
  var pad = function (n) { return (n < 10 ? '0' : '') + n; };
  return [
    IMPORT_COURSE_NAME[pfx2] || '',
    (IMPORT_MODULES[pfx2] || {})[mod] || '',
    pfx2 + '.' + pad(mod) + (les != null ? '.' + pad(les) : '')
  ];
}

function impDump_(pairs) {
  var out = [];
  pairs.forEach(function (p) {
    var v = p[1] instanceof Date ? p[1].toISOString() : impClean_(p[1]);
    if (v) out.push(p[0] + ': ' + v);
  });
  return out.join('\n');
}

function impRecord_(fields) {
  var rec = {};
  HEADERS.forEach(function (k) { rec[k] = ''; });
  rec.notified_students = false;
  rec.report_count = 1;
  for (var k in fields) rec[k] = fields[k];
  rec.issue_id = Utilities.getUuid();
  rec.reports_json = JSON.stringify([{
    kind: 'report', student_name: rec.student_name, student_contact: rec.student_contact,
    device_info: rec.device_info, instructor_name: rec.instructor_name, summary: rec.summary,
    priority: rec.priority, raw_text: rec.raw_text, recommended_steps: null, date: rec.submitted_at
  }]);
  return rec;
}

function importHistoricalIssues() {
  // Guard: refuse to run if imported rows already exist in either sheet.
  var already = 0;
  ISSUE_SHEETS.forEach(function (name) {
    var sheet = sheetByName_(name);
    if (!sheet) return;
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return;
    var head = values[0]; var idx = {}; head.forEach(function (h, i) { idx[h] = i; });
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][idx['raw_text']] || '').indexOf('Imported from') === 0) already++;
    }
  });
  if (already > 0) {
    throw new Error('Import aborted: found ' + already + ' previously imported row(s). ' +
      'Delete them first if you want to re-run the import.');
  }

  var records = [];
  var urgMap = { 1:'low', 2:'low', 3:'medium', 4:'high', 5:'high' };

  // ---- 1. Course errors form (Form responses 2) ----
  var ceRows = SpreadsheetApp.openById(IMPORT_COURSE_ERRORS_ID)
    .getSheetByName('Form responses 2').getDataRange().getValues();
  var lastCeDate = '';
  for (var r = 1; r < ceRows.length; r++) {
    var row = ceRows[r];
    var hasData = row.some(function (c) { return c !== '' && c != null; });
    if (!hasData) continue;
    var kind = impClean_(row[3]) || 'Error';
    var isImp = kind.toLowerCase().indexOf('improv') === 0;
    var slide = impClean_(isImp ? row[9] : row[4]) || impClean_(isImp ? row[4] : row[9]);
    var media = (impClean_(isImp ? row[10] : row[5]) || impClean_(isImp ? row[5] : row[10])).toLowerCase();
    var sev = isImp ? row[11] : row[6];
    if (sev === '' || sev == null) sev = isImp ? row[6] : row[11];
    var det = impClean_(isImp ? row[12] : row[7]) || impClean_(isImp ? row[7] : row[12]);
    var shots = [impClean_(row[8]), impClean_(row[13])].filter(String).join(', ');
    var parsed = IMPORT_CE_OVERRIDES[r + 1] || impParseLesson_(slide);
    var sevn = parseInt(sev, 10); if (isNaN(sevn)) sevn = null;
    var done = impClean_(row[15]).toLowerCase() === 'x';
    var when = impIso_(row[0]) || lastCeDate; // 37 old rows have no timestamp; use the previous row's
    if (impIso_(row[0])) lastCeDate = impIso_(row[0]);
    records.push(impRecord_({
      submitted_at: when, updated_at: when,
      instructor_name: impName_(row[1]),
      category: 'course_error',
      course: parsed[0], module: parsed[1], lesson_code: parsed[2],
      // The full slide/question ref as one string (e.g. EN.06.03.09, DS.10.19.09.2.M).
      // Only uppercased when it is actually a code, so "Glossary" stays "Glossary".
      lesson: /^(EN|DS|YM|FT)[\s.]*\d/i.test(slide) ? slide.toUpperCase() : slide,
      issue_type: isImp ? 'other' : 'content_error',
      summary: impSummarise_(det, slide ? (isImp ? 'Improvement on ' : 'Error on ') + slide : kind),
      priority: urgMap[sevn] || 'medium',
      priority_reason: sevn != null ? 'Imported: original ' + (isImp ? 'impact' : 'urgency') + ' ' + sevn + '/5' : 'Imported: no urgency given',
      image_urls: shots,
      status: done ? 'resolved' : 'open',
      resolution_note: done ? 'Imported from the old Course Errors/Improvements form; marked done in the old spreadsheet.' : '',
      request_kind: isImp ? 'improvement' : 'fix',
      assignee: impName_(row[14], impClean_(row[14])),
      media_kind: media === 'video' ? 'video' : media === 'text' ? 'text' : media === 'both' ? 'other' : '',
      double_checked: impClean_(row[2]).toLowerCase() === 'yes' ? true : '',
      impact: isImp ? (urgMap[sevn] || '') : '',
      raw_text: 'Imported from Course Errors/Improvements (Responses) — Form responses 2\n' + impDump_([
        ['Submitted', row[0]], ['Reported by', row[1]], ['Double checked', row[2]], ['Type', kind],
        ['Slide/question', slide], ['Video or text', media], ['Urgency/impact', sev], ['Details', det],
        ['Screenshots', shots], ['Who to fix', row[14]], ['Done', row[15]]])
    }));
  }

  // ---- 2. Tech issues form (Issue log - Form Responses) ----
  var techSs = SpreadsheetApp.openById(IMPORT_TECH_ID);
  var tfRows = techSs.getSheetByName('Issue log - Form Responses').getDataRange().getValues();
  var secMap = { website:'website', app:'app', both:'other' };
  var accessRe = /log\s?in|login|password|sign\s?in|access|locked|reset/i;
  for (var t = 1; t < tfRows.length; t++) {
    var tr = tfRows[t];
    if (!tr.some(function (c) { return c !== '' && c != null; })) continue;
    var parts = [impClean_(tr[1])];
    var section = '', priRaw = '';
    [impClean_(tr[2]), impClean_(tr[3])].forEach(function (x) {
      var xl = x.toLowerCase();
      if (secMap[xl]) { if (!section) section = secMap[xl]; }
      else if (xl === 'high' || xl === 'medium' || xl === 'low') { priRaw = xl; }
      else if (x) { parts.push(x); }
    });
    var dev = [impClean_(tr[6]), (tr[7] instanceof Date) ? '' : impClean_(tr[7])].filter(String).join(' ');
    var full = parts.filter(String).join('. ');
    var whenT = impIso_(tr[0]);
    records.push(impRecord_({
      submitted_at: whenT, updated_at: whenT,
      instructor_name: impName_(tr[4]),
      category: 'tech_issue',
      student_name: impClean_(tr[8]), student_contact: impClean_(tr[9]),
      device_info: dev,
      issue_type: accessRe.test(full) ? 'access_problem' : 'bug',
      summary: impSummarise_(full, 'Tech issue (no description)'),
      priority: priRaw || 'low',
      priority_reason: 'Imported historical record; priority from the old form where given.',
      image_urls: impClean_(tr[5]),
      status: 'past',
      resolution_note: 'Imported from the old tech issues form. No resolution record kept.',
      request_kind: 'fix',
      section: section || 'other',
      raw_text: 'Imported from Tech Fix Requests — Issue log - Form Responses\n' + impDump_([
        ['Submitted', tr[0]], ['Describe issue', tr[1]], ['Issue with', tr[2]], ['Priority', tr[3]],
        ['Owner', tr[4]], ['Images/videos', tr[5]], ['Device/OS/browser', tr[6]], ['Date', tr[7]],
        ['Student name', tr[8]], ['Student email', tr[9]], ['Other', tr[10]], ['Forwarded to tech team', tr[11]]])
    }));
  }

  // ---- 3 & 4. Product Updates (TR) and Bugs/Support fixes (SB) ----
  var areaMap = { 'website':'website', 'app':'app', 'instructor portal':'instructor_portal',
    'partner portal':'partner_portal', 'other':'other', 'course player':'course_player' };
  var bugish = /typo|broken|error|not work|doesn.t work|can.t |cant |cannot|fails|failing|404|covers|missing|wrong|fix\b|issue|slow|insecure|skips|crash/i;
  var copyish = /typo|change.*(text|wording|title|word)|spelling|rename|update.*(info|text|listing|profile)|add.*profile/i;
  var infra = /stripe|onesignal|sendgrid|shipstation|email.*junk|tracking|dpd|tax|vat|chatwoot|hotjar/i;
  // Match tabs by name prefix: the exact tail of the "Bugs, Issues, Support
  // Fix Requests..." tab name is unknown (it was truncated in the export this
  // mapping was built from), so a prefix match is the safe lookup.
  var impSheetByPrefix_ = function (ss, prefix) {
    var sheets = ss.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      if (sheets[s].getName().indexOf(prefix) === 0) return sheets[s];
    }
    throw new Error('Import: no tab found starting with "' + prefix + '"');
  };
  [['Product Updates', 'Product Updates (TR)'], ['Bugs, Issues, Support Fix', 'Bugs/Support fixes (SB)']].forEach(function (tabInfo) {
    var rows = impSheetByPrefix_(techSs, tabInfo[0]).getDataRange().getValues();
    var hi = -1;
    for (var i = 0; i < rows.length; i++) { if (rows[i][0] === 'Issue Code') { hi = i; break; } }
    var lastDate = null;
    for (var j = hi + 1; j < rows.length; j++) {
      var b = rows[j];
      var task = impClean_(b[1]);
      if (!task) continue;
      if (b[7]) lastDate = b[7];
      var date = b[7] || lastDate;
      var resolved = impClean_(b[10]).toLowerCase() === 'completed';
      var text = task + ' ' + impClean_(b[6]);
      var itype, kindB;
      if (tabInfo[0].indexOf('Bugs') === 0) { itype = 'bug'; kindB = 'fix'; }
      else if (copyish.test(task) && !/portal|function|feature/i.test(task)) { itype = 'content_copy'; kindB = 'fix'; }
      else if (infra.test(text)) { itype = 'infrastructure'; kindB = 'improvement'; }
      else if (bugish.test(task)) { itype = 'bug'; kindB = 'fix'; }
      else { itype = 'feature_request'; kindB = 'improvement'; }
      var priB = impClean_(b[3]).toLowerCase().replace(/!+$/, '');
      if (priB !== 'high' && priB !== 'medium' && priB !== 'low') {
        priB = priB === 'urgent' ? 'high' : (resolved ? 'low' : 'medium');
      }
      var whenB = impIso_(date);
      records.push(impRecord_({
        submitted_at: whenB, updated_at: whenB,
        instructor_name: impName_(b[4]),
        category: 'internal',
        issue_type: itype,
        summary: impSummarise_(task),
        priority: priB,
        priority_reason: 'Imported: original priority ' + (impClean_(b[3]) || 'not set') + '.',
        image_urls: String(b[5] || '').indexOf('http') === 0 ? impClean_(b[5]) : '',
        status: resolved ? 'resolved' : 'open',
        resolution_note: resolved ? 'Imported from the old ' + tabInfo[0] + ' tracker; original status "' + impClean_(b[10]) + '".' : '',
        request_kind: kindB,
        section: areaMap[impClean_(b[2]).toLowerCase()] || 'other',
        raw_text: 'Imported from Tech Fix Requests — ' + tabInfo[1] + '\n' + impDump_([
          ['Issue code', b[0]], ['Task', b[1]], ['Area', b[2]], ['Priority', b[3]], ['Owner', b[4]],
          ['Info/screenshot', b[5]], ['Notes', b[6]], ['Request date', b[7]], ['Work size', b[8]],
          ['End date', b[9]], ['Status', b[10]], ['Notes 2', b[11]]])
      }));
    }
  });

  // ---- Batch write, one setValues per target sheet ----
  var byTarget = {};
  records.forEach(function (rec) {
    var name = targetSheetName_(rec.category);
    (byTarget[name] = byTarget[name] || []).push(recordToRow_(rec));
  });
  for (var sheetName in byTarget) {
    var target = sheetByName_(sheetName);
    var rowsOut = byTarget[sheetName];
    target.getRange(target.getLastRow() + 1, 1, rowsOut.length, HEADERS.length).setValues(rowsOut);
  }

  // ---- Report ----
  var counts = { total: records.length, by_category: {}, by_status: {} };
  records.forEach(function (rec) {
    counts.by_category[rec.category] = (counts.by_category[rec.category] || 0) + 1;
    counts.by_status[rec.status] = (counts.by_status[rec.status] || 0) + 1;
  });
  Logger.log('Imported ' + counts.total + ' issue(s). By category: ' + JSON.stringify(counts.by_category) +
    '. By status: ' + JSON.stringify(counts.by_status) +
    '. Expected: 864 total {course_error:539, tech_issue:171, internal:154}, {open:361, resolved:332, past:171}.');
  return counts;
}

// One-off follow-up to the 11 July import (safe to run again; it only fills
// blanks). The import left the "lesson" column empty; this walks the imported
// course errors, pulls the original "Slide/question:" line back out of
// raw_text, and stores the full reference (e.g. EN.06.03.09, DS.10.19.09.2.M)
// in "lesson" — one long string, exactly as reported. The front end shows it on
// cards and uses its .K / .M suffix for the knowledge-check / module-assessment
// filter in Course fixes.
function patchImportedLessonRefs() {
  var sheet = sheetByName_(COURSE_SHEET);
  var values = sheet.getDataRange().getValues();
  var head = values[0]; var idx = {}; head.forEach(function (h, i) { idx[h] = i; });
  var patched = 0;
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!row[idx['issue_id']]) continue;
    if (String(row[idx['raw_text']] || '').indexOf('Imported from Course Errors') !== 0) continue;
    if (String(row[idx['lesson']] || '').trim() !== '') continue;
    var m = String(row[idx['raw_text']]).match(/^Slide\/question: (.+)$/m);
    if (!m) continue;
    var ref = m[1].replace(/\s+/g, ' ').trim();
    if (/^(EN|DS|YM|FT)[\s.]*\d/i.test(ref)) ref = ref.toUpperCase();
    sheet.getRange(r + 1, idx['lesson'] + 1).setValue(ref);
    patched++;
  }
  Logger.log('patchImportedLessonRefs: filled the lesson column on ' + patched + ' imported row(s).');
  return patched;
}

// Free-text questions over the whole issue log ("how many times has the
// resources page gone down?"). Builds a compact one-line-per-issue digest of
// every row in both sheets and lets the model count and summarise. Read-only:
// it never writes anything, so a wrong answer costs nothing but a shrug.
function askIssues_(body) {
  var q = String(body.question || '').trim();
  if (!q) return { ok: false, error: 'No question given.' };
  if (q.length > 500) return { ok: false, error: 'Question too long.' };
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'No API key configured.' };

  var issues = getIssues_().issues;
  var lines = ['date\tcategory\tstatus\tpriority\tkind\ttype\tcourse\tref\tsection\treports\tsummary'];
  issues.forEach(function (i) {
    var ref = i.lesson_code || '';
    var l = String(i.lesson || '');
    if (/^(EN|DS|YM|FT)[\s.]*\d/i.test(l)) ref = l; // full slide ref when we have it
    lines.push([
      String(i.submitted_at || '').slice(0, 10),
      i.category || '', i.status || '', i.priority || '',
      i.request_kind || '', i.issue_type || '',
      i.course || '', ref, i.section || '',
      i.report_count || 1,
      String(i.summary || '').replace(/\s+/g, ' ').slice(0, 110)
    ].join('\t'));
  });

  var prompt = 'You are answering a question from the team at Ardent Training (a sailing school) about their Bugs issue log.\n\n' +
    'Below is the full log, one issue per line, tab-separated. "past" status means an imported historical record with no resolution kept. ' +
    'Ref codes read COURSE.module.lesson.slide (.K = knowledge check, .M = module assessment).\n\n' +
    'Answer the question using ONLY this data. Give real counts and name the specific issues (date + summary fragment) when there are few enough to list (up to ten). ' +
    'If the data cannot answer the question, say so plainly. Note that the log only holds what was reported, so counts are lower bounds. ' +
    'Answer in plain text, no markdown.\n\n' +
    'Today: ' + new Date().toISOString().slice(0, 10) + '\n\n' +
    'THE LOG (' + issues.length + ' issues):\n' + lines.join('\n') + '\n\n' +
    'QUESTION: ' + q;

  var res;
  try {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      // 8000, not 1000. The model thinks before it answers and the thinking
      // spends from the same budget, so a thousand tokens went entirely on
      // thinking and the reply came back with no text block in it at all -
      // which is why Ask has been answering "Empty AI response" (found while
      // wiring Ask into the queues, FB-0228). Same trap as Round 44.
      payload: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 8000, messages: [{ role: 'user', content: prompt }] })
    });
  } catch (e) { return { ok: false, error: 'The request failed: ' + e }; }
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
    return { ok: false, error: 'The request failed (' + res.getResponseCode() + ').' };
  }
  var parsed; try { parsed = JSON.parse(res.getContentText()); } catch (e) { return { ok: false, error: 'The reply came back unreadable.' }; }
  tallyAi_(parsed);
  var text = '';
  if (parsed.content) for (var i = 0; i < parsed.content.length; i++) {
    if (parsed.content[i].type === 'text') text += parsed.content[i].text;
  }
  // A failure has to say what actually came back. "Empty AI response" sent us
  // looking at the front end for a fault that was a token budget all along.
  if (!text.trim()) {
    var blocks = (parsed.content || []).map(function (c) { return c.type + ':' + String(c.text || c.thinking || '').length; }).join(',');
    return { ok: false, error: 'The reply came back with no answer in it (stopped: ' + (parsed.stop_reason || '?') +
      '; blocks ' + (blocks || 'none') + '; ' + ((parsed.usage && parsed.usage.output_tokens) || '?') + ' output tokens).' };
  }
  return { ok: true, answer: text.trim() };
}

// Creates (or refreshes) a dedicated service login for Claude, so the assistant
// can drive the same API as everyone else — marking feedback needs_testing
// after a build round, closing test loops, logging issues — without borrowing
// a person's session. Run once from the editor; safe to run again (it rotates
// the token). The token lives in the Users sheet like any other session and
// goes through the normal permission checks; it is long-lived (5 years) so
// each work session doesn't need a fresh login. Actions taken with it show up
// under the name "Claude", so the audit trail says who did what.
function createClaudeServiceUser() {
  var email = 'claude-agent@ardent-training.com';
  var token = newToken_();
  var expires = new Date(Date.now() + 5 * 365 * 24 * 3600 * 1000).toISOString();
  var perms = {}; PERM_KEYS.forEach(function (k) { perms[k] = true; });
  var f = findUserByEmail_(email);
  if (f) {
    setCell_(f, 'status', 'active');
    setCell_(f, 'perms_json', JSON.stringify(perms));
    setCell_(f, 'session_token', token);
    setCell_(f, 'session_expires', expires);
    Logger.log('Claude service user refreshed; new token written to the Users sheet.');
    return 'refreshed';
  }
  var u = {
    email: email, name: 'Claude', status: 'active',
    perms_json: JSON.stringify(perms),
    pass_hash: '', pass_salt: '',            // no password: this account can never log in with one
    invite_token: '', session_token: token,
    session_expires: expires, created_at: new Date().toISOString()
  };
  usersSheet_().appendRow(USER_HEADERS.map(function (k) { return u[k]; }));
  Logger.log('Claude service user created; token written to the Users sheet.');
  return 'created';
}

// One-off for the 11 July feedback round: flips every feedback item still at
// "new" to "needs_testing" once the fixes for that round are deployed, so the
// Admin -> Feedback sub-tab shows them as ready for Edd to confirm.
function markNewFeedbackNeedsTesting() {
  var sheet = sheetByName_(FEEDBACK_SHEET);
  if (!sheet) return 0;
  var values = sheet.getDataRange().getValues();
  var head = values[0]; var idx = {}; head.forEach(function (h, i) { idx[h] = i; });
  var flipped = 0;
  for (var r = 1; r < values.length; r++) {
    if (!values[r][idx['id']]) continue;
    if (String(values[r][idx['status']] || '').toLowerCase() !== 'new') continue;
    sheet.getRange(r + 1, idx['status'] + 1).setValue('needs_testing');
    flipped++;
  }
  Logger.log('markNewFeedbackNeedsTesting: moved ' + flipped + ' feedback item(s) to needs_testing.');
  return flipped;
}

// ---- The backtest harness (r46) -------------------------------------------
//
// Two jobs, both scoped in "Scoping - Backtest harness - 9 August 2026":
// Job 1 backfills a KnownFixes corpus from historical resolved Chatwoot
// conversations; Job 2 replays past conversations against the live-case
// machinery and scores it. Everything here is key-gated (DEPLOY_KEY), reads
// Chatwoot but never writes to it, and writes only to its own two sheets.
// The two honesty rules live here: replays pass a cutoff so the lookup never
// sees a fix from the future, and draft judging is blind pairwise with the
// order randomised.

var KNOWNFIXES_SHEET = 'KnownFixes';
// Round 65 (Edd, FB-0231). The last three are APPENDED, never inserted - the
// array order IS the sheet column order, so an insert misaligns every existing
// row (the taxonomy rule, learned the hard way in Round 23).
//
// A fix is only a fix somewhere. Edd hit this on a Yachtmaster case that was
// offered a fix which "only applied to really old day skipper accounts, not
// Yachtmaster ever" - the corpus had no idea a fix HAS a scope, so every entry
// read as universal. course_scope is a canonical course name (or several,
// comma-separated) and applies_when is the plain-English caveat the
// conversation gave us. Blank means unscoped, which is how every row
// backfilled before today reads, so nothing changes for them.
var KNOWNFIX_HEADERS = ['conversation_id', 'resolved_date', 'problem', 'fix', 'category', 'lesson_code',
  'message_count', 'source', 'extracted_at', 'dup_of', 'course_scope', 'applies_when', 'flags_json'];
var BACKTESTLOG_SHEET = 'BacktestLog';
var BACKTESTLOG_HEADERS = ['type', 'conversation_id', 'data_json', 'created_at'];
var BACKFILL_OLDEST = '2026-02-10T00:00:00Z';  // ~6 months back, per the cap
// r47: raised from 500 on Edd's instruction - the first sweep only reached ~4
// days of history by Chatwoot's last-activity ordering, so the corpus had no
// real depth behind it. 2000 covers the original 500 plus the next ~1500, and
// the 6-month BACKFILL_OLDEST boundary still stops it either way.
var BACKFILL_MAX_ROWS = 2000;                  // conversations processed, per the cap

function knownFixesSheet_(create) {
  var sh = sheetByName_(KNOWNFIXES_SHEET);
  if (!sh && create) {
    sh = ss_().insertSheet(KNOWNFIXES_SHEET);
    sh.getRange(1, 1, 1, KNOWNFIX_HEADERS.length).setValues([KNOWNFIX_HEADERS]);
    sh.setFrozenRows(1);
  }
  // Round 65: the three appended columns land on a sheet that already has 930
  // rows under a ten-column header. knownFixRows_ maps values by the header
  // row, so without this the new cells would be written and never read back.
  // Appending only, and only what is missing, so it is safe to run on any
  // access and does nothing at all once the header is up to date.
  if (sh) {
    try {
      var last = sh.getLastColumn();
      if (last < KNOWNFIX_HEADERS.length) {
        var head = last ? sh.getRange(1, 1, 1, last).getValues()[0] : [];
        var add = [];
        for (var c = last; c < KNOWNFIX_HEADERS.length; c++) add.push(KNOWNFIX_HEADERS[c]);
        // Only append when what is already there is our header, in our order.
        var sane = true;
        for (var h = 0; h < head.length; h++) if (String(head[h]) !== KNOWNFIX_HEADERS[h]) sane = false;
        if (sane && add.length) sh.getRange(1, last + 1, 1, add.length).setValues([add]);
      }
    } catch (e) {}
  }
  return sh;
}
// ---- corpus corrections (Round 65, FB-0231) --------------------------------
// A shown known fix that was wrong needs somewhere to go. Edd: "this only
// applied to really old day skipper accounts. Not Yachtmaster ever." Until now
// there was nothing to press, so the corpus kept offering it and the next
// instructor made the same wasted trip.
//
// Two deliberate limits. The instructor's report NEVER rewrites the row: it is
// recorded against the entry as a flag and queued for Edd, the same
// approve-or-reject shape the playbook suggestions use. And a flag on its own
// only ever adds a caveat to future suggestions - narrowing or rewording the
// entry is Edd's press, not the machine's.
function kfFlagCount_(kf) {
  try { var a = JSON.parse(kf.flags_json || '[]'); return Array.isArray(a) ? a.length : 0; } catch (e) { return 0; }
}
function getKfCorrections_() {
  var raw = PropertiesService.getScriptProperties().getProperty('KNOWNFIX_CORRECTIONS');
  if (!raw) return [];
  try { var a = JSON.parse(raw); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}
function saveKfCorrections_(arr) {
  PropertiesService.getScriptProperties().setProperty('KNOWNFIX_CORRECTIONS', JSON.stringify(arr || []));
}
function listKnownFixFlags_() { return { ok: true, corrections: getKfCorrections_() }; }

// Filed by whoever was shown the suggestion. Records the objection on the row
// (a count and a note, nothing overwritten) and queues it for Edd.
function flagKnownFix_(body) {
  var why = String((body && body.why) || '').trim();
  if (!why) return { ok: false, error: 'Say what was wrong about it - "wrong" on its own tells the next person nothing.' };
  var corpusId = String((body && body.corpus_id) || '').trim();
  var who = (body && body._user && body._user.name) || '';
  var now = new Date().toISOString();
  var scope = kfNormaliseScope_(body && body.course_scope);
  var appliesWhen = String((body && body.applies_when) || '').trim().slice(0, 400);

  // Stamp the flag on the entry itself when we know which one it was. The
  // suggestion is AI-written from a shortlist, so it does not always come back
  // with an id; a correction with no id still reaches Edd, it just cannot
  // caveat future suggestions on its own.
  var stamped = false;
  if (corpusId) {
    var sh = knownFixesSheet_(false);
    var rows = sh ? knownFixRows_() : [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].conversation_id) !== corpusId) continue;
      var flags = [];
      try { flags = JSON.parse(rows[i].flags_json || '[]'); } catch (e) { flags = []; }
      if (!Array.isArray(flags)) flags = [];
      flags.push({ by: who, at: now, why: why.slice(0, 400) });
      if (flags.length > 20) flags = flags.slice(flags.length - 20);
      var col = KNOWNFIX_HEADERS.indexOf('flags_json') + 1;
      sh.getRange(rows[i]._rowNum, col).setValue(JSON.stringify(flags));
      stamped = true;
      break;
    }
  }

  var arr = getKfCorrections_();
  arr.push({
    id: Utilities.getUuid(),
    corpus_id: corpusId,
    by: who,
    created_at: now,
    why: why.slice(0, 1000),
    course_scope: scope,
    applies_when: appliesWhen,
    shown_fix: String((body && body.shown_fix) || '').slice(0, 800),
    case_summary: String((body && body.case_summary) || '').slice(0, 400),
    conversation_id: String((body && body.conversation_id) || ''),
    issue_id: String((body && body.issue_id) || ''),
    stamped: stamped
  });
  if (arr.length > 60) arr = arr.slice(arr.length - 60);
  saveKfCorrections_(arr);
  return { ok: true, stamped: stamped };
}

// Edd's press, and the ONLY thing that changes a corpus row's wording or scope.
// Approving writes the scope he confirmed onto the entry; rejecting drops the
// correction and leaves the entry exactly as it was.
function resolveKnownFixFlag_(body) {
  var arr = getKfCorrections_();
  var kept = [], matched = null;
  for (var i = 0; i < arr.length; i++) {
    if (arr[i].id === body.id) matched = arr[i]; else kept.push(arr[i]);
  }
  if (!matched) return { ok: false, error: 'Correction not found (it may have already been actioned).' };
  var applied = false;
  if (body.approve && matched.corpus_id) {
    var scope = kfNormaliseScope_(body.course_scope !== undefined ? body.course_scope : matched.course_scope);
    var appliesWhen = String((body.applies_when !== undefined ? body.applies_when : matched.applies_when) || '').slice(0, 400);
    var drop = !!body.drop_entry;
    var sh = knownFixesSheet_(false);
    var rows = sh ? knownFixRows_() : [];
    for (var j = 0; j < rows.length; j++) {
      if (String(rows[j].conversation_id) !== String(matched.corpus_id)) continue;
      if (drop) {
        // Retired rather than deleted: dup_of is already how this corpus takes
        // a row out of the lookup, and the history is still worth keeping.
        sh.getRange(rows[j]._rowNum, KNOWNFIX_HEADERS.indexOf('dup_of') + 1).setValue('retired_' + new Date().toISOString().slice(0, 10));
      } else {
        if (scope) sh.getRange(rows[j]._rowNum, KNOWNFIX_HEADERS.indexOf('course_scope') + 1).setValue(scope);
        if (appliesWhen) sh.getRange(rows[j]._rowNum, KNOWNFIX_HEADERS.indexOf('applies_when') + 1).setValue(appliesWhen);
      }
      applied = true;
      break;
    }
  }
  saveKfCorrections_(kept);
  return { ok: true, applied: applied };
}

function knownFixRows_() {
  var sh = knownFixesSheet_(false);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    if (!values[r][0]) continue;
    var rec = {};
    for (var c = 0; c < head.length; c++) rec[head[c]] = values[r][c];
    rec.conversation_id = String(rec.conversation_id);
    // Sheets coerces ISO strings to Dates on the way in; hand back ISO strings
    // so date maths and comparisons stay predictable (the 30 Jul coercion trap).
    if (rec.resolved_date instanceof Date) rec.resolved_date = rec.resolved_date.toISOString();
    rec._rowNum = r + 1;
    out.push(rec);
  }
  return out;
}
function backtestLogSheet_() {
  var sh = sheetByName_(BACKTESTLOG_SHEET);
  if (!sh) {
    sh = ss_().insertSheet(BACKTESTLOG_SHEET);
    sh.getRange(1, 1, 1, BACKTESTLOG_HEADERS.length).setValues([BACKTESTLOG_HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function btLog_(type, convId, obj) {
  backtestLogSheet_().appendRow([type, String(convId || ''), JSON.stringify(obj || {}), new Date().toISOString()]);
}
function btLogRows_(type) {
  var sh = sheetByName_(BACKTESTLOG_SHEET);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  var out = [];
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][0]) !== type) continue;
    var o = { conversation_id: String(values[r][1]), created_at: values[r][3] };
    try { o.data = JSON.parse(values[r][2]); } catch (e) { o.data = {}; }
    out.push(o);
  }
  return out;
}

// One conversation as structured turns, read-only, no Drive copies. The same
// filtering rules as chatwootImport_ (skip activity lines and private notes)
// so the replay sees exactly what the live brief would see.
function chatwootTurns_(convId) {
  var conv = chatwootCall_('/conversations/' + convId);
  var msgs = chatwootCall_('/conversations/' + convId + '/messages');
  var list = (msgs && (msgs.payload || (msgs.data && msgs.data.payload))) || [];
  var sender = (conv && conv.meta && conv.meta.sender) || {};
  var turns = [];
  list.forEach(function (m) {
    var t = Number(m.message_type);
    if (t !== 0 && t !== 1) return;
    if (m.private) return;
    var body = cleanChatwootBody_(m.content);
    var pics = (m.attachments || []).filter(function (a) { return String(a.file_type) === 'image'; }).length;
    if (!body && !pics) return;
    turns.push({
      who: t === 0 ? 'student' : 'agent',
      name: t === 0 ? (sender.name || 'Student') : ((m.sender && (m.sender.name || m.sender.available_name)) || 'Ardent'),
      at: m.created_at ? new Date(Number(m.created_at) * 1000).toISOString() : '',
      body: body || ('[shared ' + pics + ' screenshot' + (pics > 1 ? 's' : '') + ']')
    });
  });
  return {
    conversation_id: String(convId),
    student_name: sender.name || '',
    turns: turns,
    started_at: turns.length ? turns[0].at : '',
    ended_at: turns.length ? turns[turns.length - 1].at : '',
    status: String((conv && conv.status) || (conv && conv.payload && conv.payload.status) || '')
  };
}
function btTranscript_(turns, upTo) {
  var lines = [];
  for (var i = 0; i < turns.length && i < upTo; i++) {
    var t = turns[i];
    lines.push(t.name + (t.at ? ' (' + t.at.slice(0, 16).replace('T', ' ') + ')' : '') + ': ' + t.body);
  }
  return lines.join('\n\n');
}

// Job 1: sweep resolved Chatwoot history into KnownFixes. Resumable: every
// conversation looked at gets a 'seen' row in BacktestLog, so a re-run skips
// straight past it, and the caller steers with page/max_pages. Stops itself at
// the ~6-month / 500-conversation cap from the scoping doc.
function btBackfill_(body) {
  var budget = Math.min(Number(body.budget_ms) || 270000, 300000);
  var started = Date.now();
  var page = Math.max(1, Number(body.page) || 1);
  var maxPages = Math.max(1, Number(body.max_pages) || 3);
  var endPage = page + maxPages - 1;
  var seen = {};
  btLogRows_('seen').forEach(function (r) { seen[r.conversation_id] = true; });
  knownFixRows_().forEach(function (r) { seen[r.conversation_id] = true; });
  var seenCount = Object.keys(seen).length;
  var stats = { processed: 0, added: 0, skipped_seen: 0, nofix: 0, thin: 0, errors: 0, hit_oldest: false, hit_cap: false, pages_done: [] };
  var kfSheet = knownFixesSheet_(true);

  while (page <= endPage && Date.now() - started < budget) {
    var out;
    try { out = chatwootCall_('/conversations?status=resolved&page=' + page); }
    catch (e) { stats.errors++; break; }
    var payload = (out && out.data && out.data.payload) || (out && out.payload) || [];
    if (!payload.length) { page = 0; break; }   // ran out of history
    for (var i = 0; i < payload.length; i++) {
      if (Date.now() - started >= budget) break;
      var c = payload[i];
      var id = String(c.id);
      var lastAt = c.last_activity_at ? new Date(Number(c.last_activity_at) * 1000).toISOString() : '';
      if (lastAt && lastAt < BACKFILL_OLDEST) { stats.hit_oldest = true; break; }
      if (seenCount + stats.processed >= BACKFILL_MAX_ROWS) { stats.hit_cap = true; break; }
      if (seen[id]) { stats.skipped_seen++; continue; }
      stats.processed++;
      var t;
      try { t = chatwootTurns_(id); }
      catch (e2) { stats.errors++; btLog_('seen', id, { verdict: 'error', why: String(e2).slice(0, 120) }); continue; }
      var hasStudent = t.turns.some(function (x) { return x.who === 'student'; });
      var hasAgent = t.turns.some(function (x) { return x.who === 'agent'; });
      if (t.turns.length < 2 || !hasStudent || !hasAgent) {
        stats.thin++; btLog_('seen', id, { verdict: 'thin' }); continue;
      }
      var transcript = btTranscript_(t.turns, t.turns.length).slice(0, 12000);
      var got = anthropicRaw_(ANTHROPIC_MODEL, btExtractPrompt_(transcript), 1000);
      if (!got.json) { stats.errors++; btLog_('seen', id, { verdict: 'ai_failed', why: got.why }); continue; }
      var f = got.json;
      if (!f.found || !String(f.fix || '').trim() || !String(f.problem || '').trim()) {
        stats.nofix++; btLog_('seen', id, { verdict: 'nofix' }); continue;
      }
      kfSheet.appendRow([id, lastAt || t.ended_at, String(f.problem).slice(0, 800), String(f.fix).slice(0, 800),
        String(f.category || 'other'), String(f.lesson_code || ''), t.turns.length, 'chatwoot_backfill',
        new Date().toISOString(), '', kfNormaliseScope_(f.course_scope), String(f.applies_when || '').slice(0, 400), '']);
      stats.added++;
    }
    if (stats.hit_oldest || stats.hit_cap) break;
    stats.pages_done.push(page);
    page++;
  }
  var finished = stats.hit_oldest || stats.hit_cap || page === 0;
  return { ok: true, next_page: finished ? 0 : page, finished: finished, stats: stats, tally: AI_TALLY };
}
function btExtractPrompt_(transcript) {
  return 'You are building a "known fixes" memory for Ardent Training, an online RYA sailing school. ' +
    'Below is one RESOLVED support conversation from its history.\n\n' +
    'Decide whether this thread shows BOTH: (a) a specific problem the student described, AND (b) a fix that was ' +
    'confirmed to work or clearly given as the answer. A password reset that worked, a refund processed, a missing ' +
    'ebook re-sent, a clear how-to question answered - all count, as long as the resolution is visible in the thread. ' +
    'Chit-chat, unanswered threads, threads that fizzle out, and resolutions that happened somewhere else ("I\'ll call you") ' +
    'do NOT count. When in doubt, found is false.\n\n' +
    'THE CONVERSATION:\n"""\n' + transcript + '\n"""\n\n' +
    'Return ONLY JSON, no prose, no fences:\n' +
    '{"found": true or false,\n' +
    ' "problem": "<one or two plain sentences: what was going wrong, with the device or platform if relevant>",\n' +
    ' "fix": "<one or two plain sentences: the specific thing that fixed it or the answer given>",\n' +
    ' "category": "<one of: tech_issue, course_error, shipping, admin, other>",\n' +
    ' "lesson_code": "<e.g. DS.09.04 if one specific lesson is identifiable, else empty string>",\n' +
    // Round 65 (FB-0231). The scope has to come from the conversation, not from
    // a guess: an entry wrongly narrowed stops firing where it would have
    // helped, and one wrongly widened is the fault Edd reported. Say nothing
    // rather than say something shaky.
    ' "course_scope": "<ONLY if the thread makes clear the fix applies to one course and not others, name it from: Essential Navigation, Day Skipper, Yachtmaster, Fast Track, SRC, PPR. Several allowed, comma-separated. Empty string when the thread does not say, or when it plainly applies to any course - do NOT infer a scope from which course the student happened to be on>",\n' +
    ' "applies_when": "<ONLY if the thread states a condition on when this fix applies (an account age, an old signup route, a particular device or app version, a specific enrolment type), one short plain-English line saying it. Empty string otherwise>"}';
}

// Canonical course names only, so the comparison later is a straight match
// rather than a fuzzy one. Anything the model volunteers that is not one of
// our six is dropped: a scope we cannot compare against is worse than none,
// because it would silently exclude every case.
var KF_COURSES = ['Essential Navigation', 'Day Skipper', 'Yachtmaster', 'Fast Track', 'SRC', 'PPR'];
function kfNormaliseScope_(raw) {
  var s = String(raw || '').toLowerCase();
  if (!s.trim()) return '';
  var out = [];
  KF_COURSES.forEach(function (c) { if (s.indexOf(c.toLowerCase()) > -1) out.push(c); });
  return out.join(', ');
}
// Does a corpus entry apply to the case in front of us?
//   'yes'   - no scope recorded, or the course matches. Offer it as normal.
//   'no'    - scoped to other courses and we KNOW this case is not one of them.
//             Drop it: this is the Yachtmaster-offered-a-Day-Skipper-fix case.
//   'maybe' - scoped, but we do not know the case's course. Offer it WITH the
//             caveat attached rather than pretending the scope isn't there.
function kfScopeVerdict_(scope, course) {
  var sc = String(scope || '').trim();
  if (!sc) return 'yes';
  var c = String(course || '').trim().toLowerCase();
  if (!c) return 'maybe';
  var list = sc.toLowerCase().split(',').map(function (x) { return x.trim(); }).filter(Boolean);
  for (var i = 0; i < list.length; i++) if (list[i] && c.indexOf(list[i]) > -1) return 'yes';
  return 'no';
}

// Near-identical pairs get marked, not deleted: dup_of points at the row that
// stays live, and the lookup skips marked rows. The repeats are still real
// history (how often a fix recurs is a signal), so nothing is thrown away.
function btDedupe_() {
  var rows = knownFixRows_();
  var marked = 0;
  function tokens(r) {
    var m = {};
    ((r.problem || '') + ' ' + (r.fix || '')).toLowerCase().replace(/[^a-z0-9\s]/g, ' ').split(/\s+/)
      .forEach(function (w) { if (w.length > 3 && !FIX_STOPWORDS[w]) m[w] = true; });
    return m;
  }
  var toks = rows.map(tokens);
  var sh = knownFixesSheet_(false);
  var dupCol = KNOWNFIX_HEADERS.indexOf('dup_of') + 1;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].dup_of) continue;
    for (var j = i + 1; j < rows.length; j++) {
      if (rows[j].dup_of) continue;
      var a = toks[i], b = toks[j], inter = 0, uni = 0, k;
      for (k in a) { uni++; if (b[k]) inter++; }
      for (k in b) { if (!a[k]) uni++; }
      if (uni && inter / uni >= 0.75) {
        sh.getRange(rows[j]._rowNum, dupCol).setValue(rows[i].conversation_id);
        rows[j].dup_of = rows[i].conversation_id;
        marked++;
      }
    }
  }
  return { ok: true, rows: rows.length, marked_duplicates: marked };
}

// Job 2: replay one conversation turn by turn. After each student message the
// live-case brief runs on the transcript-so-far, with the known-fix lookup
// cut off at the conversation's own start date - the no-time-travel rule.
// Ground truth, the accuracy verdict, and the blind draft comparison all use
// VERIFIER_MODEL so the judge is not the model that wrote the briefs.
function btReplay_(body) {
  var id = chatwootConvId_(body.conversation_id || '');
  if (!id) return { ok: false, error: 'conversation_id required' };
  var maxTurns = Math.max(1, Math.min(Number(body.max_turns) || 6, 10));
  AI_TALLY = { calls: 0, in_tokens: 0, out_tokens: 0 };

  var t;
  try { t = chatwootTurns_(id); }
  catch (e) { return { ok: false, error: 'Chatwoot read failed: ' + String(e).slice(0, 160) }; }
  if (t.turns.length < 2) return { ok: false, error: 'too thin to replay' };
  var cutoff = t.started_at;
  var fullTranscript = btTranscript_(t.turns, t.turns.length).slice(0, 14000);

  // Ground truth: what actually fixed it, and who gave the fix.
  var truthGot = anthropicRaw_(VERIFIER_MODEL,
    'Read this RESOLVED support conversation from Ardent Training (online RYA sailing school).\n\n"""\n' + fullTranscript + '\n"""\n\n' +
    'Return ONLY JSON, no prose, no fences:\n' +
    '{"clear_fix": true or false (was a specific fix or answer visibly confirmed or clearly given?),\n' +
    ' "fix": "<one or two sentences: the specific thing that actually resolved it>",\n' +
    ' "resolved_by": "<first name of the team member who gave the fix, or empty>",\n' +
    ' "category": "<one of: tech_issue, course_error, shipping, admin, other>"}', 8000);
  if (!truthGot.json) return { ok: false, error: 'truth read failed: ' + truthGot.why };
  var truth = truthGot.json;

  // The replay: after each student message, what would the brief have said?
  var suggestions = [];
  var studentTurns = 0;
  for (var i = 0; i < t.turns.length; i++) {
    if (t.turns[i].who !== 'student') continue;
    studentTurns++;
    if (studentTurns > maxTurns) continue;   // keep counting M, stop briefing
    var imp = { transcript: btTranscript_(t.turns, i + 1), message_count: i + 1, images: [], link: '' };
    var core = caseBriefCore_(imp, cutoff);
    if (core.error) { suggestions.push({ student_turn: studentTurns, error: core.error }); continue; }
    suggestions.push({
      student_turn: studentTurns,
      at_turn_index: i + 1,
      next: core.bj.next || '',
      instructor_action: !!core.bj.instructor_action,
      known_fix: core.bj.fix || '',
      fix_based_on: core.bj.fix_based_on || ''
    });
  }

  // Accuracy: did any suggestion match the real fix, and how early?
  var accGot = anthropicRaw_(VERIFIER_MODEL,
    'A support system replayed a historical conversation. After each student message it suggested a next step, and ' +
    'sometimes a known fix from past cases. Judge it against what actually resolved the conversation.\n\n' +
    'THE FIX THAT ACTUALLY RESOLVED IT:\n' + JSON.stringify(truth.fix) + '\n\n' +
    'THE SYSTEM\'S SUGGESTIONS, in order:\n' + JSON.stringify(suggestions) + '\n\n' +
    'A suggestion MATCHES when acting on it would have produced the actual fix - the same specific action or answer. ' +
    'The same topic is not a match. A generic step (restart, clear cache, try another browser) only matches if the ' +
    'actual fix genuinely was that step. Judge the known_fix and the next step together - either can carry the match.\n\n' +
    'Return ONLY JSON, no prose, no fences:\n' +
    '{"first_match_student_turn": <number or null>,\n' +
    ' "why": "<one plain line: why it matched where it did, or why it never matched>"}', 8000);
  var acc = accGot.json || { first_match_student_turn: null, why: 'judge failed: ' + accGot.why };

  // Draft comparison: at the first substantive agent reply, write ours from
  // the same position and judge the pair blind, order randomised.
  var draft = null;
  var replyIdx = -1;
  for (var j = 0; j < t.turns.length; j++) {
    if (t.turns[j].who === 'agent' && t.turns[j].body.length > 60 && j > 0) { replyIdx = j; break; }
  }
  if (replyIdx > 0) {
    var realReply = t.turns[replyIdx].body;
    var agentName = t.turns[replyIdx].name || truth.resolved_by || '';
    var prefix = btTranscript_(t.turns, replyIdx);
    var ourDraft = btDraftAt_(prefix, suggestions, replyIdx, agentName, t.student_name);
    if (ourDraft) {
      var oursFirst = Math.random() < 0.5;
      var A = oursFirst ? ourDraft : realReply;
      var B = oursFirst ? realReply : ourDraft;
      var judgeGot = anthropicRaw_(VERIFIER_MODEL,
        'At the same point in a real support conversation (Ardent Training, online RYA sailing school), two different ' +
        'replies to the student were written. You do not know who wrote which. Judge the pair.\n\n' +
        'THE CONVERSATION SO FAR:\n"""\n' + prefix.slice(0, 8000) + '\n"""\n\n' +
        'WHAT EVENTUALLY RESOLVED THE CONVERSATION (context for judging correctness only):\n' + JSON.stringify(truth.fix) + '\n\n' +
        'REPLY A:\n"""\n' + A + '\n"""\n\n' +
        'REPLY B:\n"""\n' + B + '\n"""\n\n' +
        'Three verdicts, each "A", "B", or "tie":\n' +
        '- correctness: which points the student closer to what actually solved it (an invented or wrong step loses)\n' +
        '- helpfulness: clear, one actionable step not a wall of them, the right length for support\n' +
        '- voice: which reads more like a warm, plain-English human instructor\n\n' +
        'Return ONLY JSON, no prose, no fences:\n' +
        '{"correctness": "A|B|tie", "helpfulness": "A|B|tie", "voice": "A|B|tie", "note": "<one line>"}', 8000);
      var v = judgeGot.json;
      if (v) {
        function unblind(x) { return x === 'tie' ? 'tie' : ((x === 'A') === oursFirst ? 'ours' : 'theirs'); }
        draft = {
          agent: agentName, at_turn_index: replyIdx + 1,
          ours: ourDraft, theirs: realReply,
          correctness: unblind(v.correctness), helpfulness: unblind(v.helpfulness), voice: unblind(v.voice),
          note: v.note || ''
        };
      } else {
        draft = { agent: agentName, error: 'judge failed: ' + judgeGot.why };
      }
    }
  }

  var elapsedH = (t.started_at && t.ended_at) ?
    Math.round((new Date(t.ended_at) - new Date(t.started_at)) / 36000) / 100 : null;
  var result = {
    conversation_id: id,
    category: truth.category || '',
    clear_fix: !!truth.clear_fix,
    real_fix: truth.fix || '',
    turns_total: t.turns.length,
    student_turns_total: studentTurns,
    student_turns_briefed: Math.min(studentTurns, maxTurns),
    elapsed_hours: elapsedH,
    first_match_student_turn: acc.first_match_student_turn,
    match_why: acc.why || '',
    suggestions: suggestions,
    draft: draft,
    cutoff: cutoff,
    tally: AI_TALLY
  };
  btLog_('replay', id, result);
  return { ok: true, result: result };
}

// Our reply from the replay position: the same prompt shape the live
// caseDraftReply_ uses, voiced with the real agent's guide where one exists,
// but standing on the replay's own brief - no live case row is touched.
function btDraftAt_(prefix, suggestions, replyIdx, agentName, studentName) {
  var brief = null;
  for (var i = suggestions.length - 1; i >= 0; i--) {
    if (!suggestions[i].error && suggestions[i].at_turn_index <= replyIdx) { brief = suggestions[i]; break; }
  }
  if (!brief && suggestions.length && !suggestions[0].error) brief = suggestions[0];
  var guide = agentName ? voiceGuideFor_(agentName) : '';
  // r47: the SAME prompt the live caseDraftReply_ uses (draftReplyPrompt_), so
  // the benchmark measures what ships. The judged text is the message alone -
  // the real reply was also just the sent message, so that is the fair pair.
  var prompt = draftReplyPrompt_({
    recommended_next_step: (brief && brief.next) || '',
    next_step_is_an_instructor_action: !!(brief && brief.instructor_action),
    known_fix_from_a_past_issue: (brief && brief.known_fix) || '',
    student_first_name: String(studentName || '').split(' ')[0]
  }, prefix.slice(-2500), guide, agentName);
  var got = anthropicRaw_(DRAFT_MODEL, prompt, 8000);
  if (!got.json || !String(got.json.message || '').trim()) return null;
  return String(got.json.message).trim();
}

// The dispatcher: one key-gated action, several ops. 'sample' hands back the
// KnownFixes pool so the caller can stratify; 'results' hands back every
// replay row so the write-up happens outside the 6-minute window.
function backtest_(body) {
  var key = PropertiesService.getScriptProperties().getProperty('DEPLOY_KEY');
  if (!key || String(body.key || '') !== key) return { ok: false, error: 'bad deploy key' };
  var op = String(body.op || '');
  if (op === 'backfill') return btBackfill_(body);
  if (op === 'dedupe') return btDedupe_();
  if (op === 'replay') return btReplay_(body);
  if (op === 'sample') {
    return { ok: true, pool: knownFixRows_().map(function (r) {
      return { conversation_id: r.conversation_id, resolved_date: r.resolved_date, category: r.category,
        message_count: Number(r.message_count) || 0, lesson_code: r.lesson_code || '', dup_of: r.dup_of || '' };
    }) };
  }
  if (op === 'results') {
    // _logged_at lets the caller split runs apart: the r46 rows and any r47
    // re-run of the same conversations live in the same log.
    var replays = btLogRows_('replay');
    return { ok: true, count: replays.length, replays: replays.map(function (r) {
      var d = r.data || {}; d._logged_at = r.created_at; return d;
    }) };
  }
  if (op === 'peek') {
    // Read-only look at one conversation's turns, for verifying the wording of
    // a seed against what was actually said - never seed a guess.
    var pid = chatwootConvId_(body.conversation_id || '');
    if (!pid) return { ok: false, error: 'conversation_id required' };
    try {
      var pt = chatwootTurns_(pid);
      return { ok: true, conversation_id: pid, student_name: pt.student_name,
        started_at: pt.started_at, ended_at: pt.ended_at, turns: pt.turns };
    } catch (e) { return { ok: false, error: String(e).slice(0, 200) }; }
  }
  if (op === 'seed') {
    // r47: hand-curated KnownFixes entries from the backtest's miss catalogue.
    // Each entry names its source conversation; the resolved_date is read from
    // the REAL conversation (its last message) so the time-travel rule stays
    // honest - a seed can never fire on the conversation it was learned from.
    var entries = body.entries || [];
    var sh47 = knownFixesSheet_(true);
    var have = {};
    knownFixRows_().forEach(function (r) { have[r.conversation_id] = true; });
    var seeded = [], skipped = [], failed = [];
    for (var si = 0; si < entries.length; si++) {
      var en = entries[si] || {};
      var sid = String(en.conversation_id || '');
      if (!sid || !en.problem || !en.fix) { failed.push({ id: sid, why: 'missing fields' }); continue; }
      if (have[sid]) { skipped.push(sid); continue; }
      var when = String(en.resolved_date || '');
      if (!when) {
        try { when = chatwootTurns_(sid).ended_at || ''; }
        catch (e) { failed.push({ id: sid, why: 'date read failed: ' + String(e).slice(0, 80) }); continue; }
      }
      if (!when) { failed.push({ id: sid, why: 'no resolved date' }); continue; }
      sh47.appendRow([sid, when, String(en.problem).slice(0, 800), String(en.fix).slice(0, 800),
        String(en.category || 'other'), String(en.lesson_code || ''), Number(en.message_count) || 0,
        'seeded_r47', new Date().toISOString(), '']);
      have[sid] = true;
      seeded.push(sid);
    }
    return { ok: true, seeded: seeded, skipped_existing: skipped, failed: failed };
  }
  if (op === 'kf') {
    // Read back full KnownFixes rows for given conversation ids (or all when
    // none given) - lets a curation pass check what the extractor actually
    // wrote before deciding whether a hand-verified version should replace it.
    var want = {};
    (body.conversation_ids || []).forEach(function (x) { want[String(x)] = true; });
    var all = knownFixRows_();
    var hits = all.filter(function (r) { return !Object.keys(want).length || want[r.conversation_id]; });
    return { ok: true, rows: hits.map(function (r) {
      return { conversation_id: r.conversation_id, resolved_date: r.resolved_date, problem: r.problem,
        fix: r.fix, category: r.category, lesson_code: r.lesson_code, source: r.source, dup_of: r.dup_of || '', row: r._rowNum,
        // Round 65: a curation pass has to be able to see the scope, and this
        // is also how the appended columns get checked from outside the sheet.
        course_scope: r.course_scope || '', applies_when: r.applies_when || '', flags_json: r.flags_json || '' };
    }) };
  }
  if (op === 'kffix') {
    // Replace the problem/fix wording on an existing KnownFixes row (matched
    // by conversation_id) with a hand-verified version. Curation, not deletion:
    // the row keeps its id, date and place; source records the override.
    var fixes = body.entries || [];
    var sh2 = knownFixesSheet_(false);
    if (!sh2) return { ok: false, error: 'no KnownFixes sheet' };
    var byId = {};
    knownFixRows_().forEach(function (r) { byId[r.conversation_id] = r; });
    var done = [], missing = [];
    for (var fi = 0; fi < fixes.length; fi++) {
      var fe = fixes[fi] || {};
      var frow = byId[String(fe.conversation_id || '')];
      if (!frow) { missing.push(String(fe.conversation_id || '')); continue; }
      if (fe.problem) sh2.getRange(frow._rowNum, KNOWNFIX_HEADERS.indexOf('problem') + 1).setValue(String(fe.problem).slice(0, 800));
      if (fe.fix) sh2.getRange(frow._rowNum, KNOWNFIX_HEADERS.indexOf('fix') + 1).setValue(String(fe.fix).slice(0, 800));
      if (fe.category) sh2.getRange(frow._rowNum, KNOWNFIX_HEADERS.indexOf('category') + 1).setValue(String(fe.category));
      sh2.getRange(frow._rowNum, KNOWNFIX_HEADERS.indexOf('source') + 1).setValue('curated_r47');
      done.push(frow.conversation_id);
    }
    return { ok: true, updated: done, missing: missing };
  }
  if (op === 'suggest') {
    // Process patterns from the miss catalogue go through the SAME
    // suggest-and-approve queue as every other playbook idea - the playbook
    // itself never changes without Edd pressing approve.
    var sug = body.suggestions || [];
    var queued = 0;
    for (var qi = 0; qi < sug.length; qi++) {
      var s = sug[qi] || {};
      if (!s.suggestion) continue;
      addSuggestion_({
        id: Utilities.getUuid(), issue_id: '', summary: String(s.summary || 'From the 10 Aug backtest miss catalogue'),
        suggestion: String(s.suggestion), section: String(s.section || 'known issues'),
        created_at: new Date().toISOString()
      });
      queued++;
    }
    return { ok: true, queued: queued };
  }
  if (op === 'state') {
    var kf = knownFixRows_();
    return { ok: true, known_fixes: kf.length, dups: kf.filter(function (r) { return r.dup_of; }).length,
      seen: btLogRows_('seen').length, replays: btLogRows_('replay').length };
  }
  return { ok: false, error: 'unknown backtest op: ' + op };
}


/**
 * Ardent Reports ticket minting, for pasting into the Issue Tracker's Code.gs.
 *
 * WHAT THIS IS FOR
 * The Reports dashboard reads course telemetry out of Supabase. It used to do
 * that with one shared password baked into the page. That was survivable while
 * the page lived on Edd's machine and is not survivable once it is served from
 * a public GitHub Pages repo, where anyone could read the password out of the
 * source.
 *
 * So instead the tracker, which already knows who people are, vouches for them.
 * This mints a short-lived ticket saying "this is Charly, she holds analytics,
 * valid for the next half hour", signs it, and hands it over. Supabase checks
 * the signature and serves the data. The signing secret never leaves the two
 * servers, so nothing sensitive is in the published page at all.
 *
 * Supabase does NOT call back here to check anything. That is deliberate: a
 * round trip to Apps Script costs 1.8 to 2.8 seconds, and doing it per request
 * would put that delay on every chart. Verifying a signature takes microseconds.
 * This runs once when someone opens Reports, not once per query.
 *
 * ---------------------------------------------------------------------------
 * INSTALLATION, four steps
 *
 * 1. Generate the shared secret ONCE, and paste the same value into both
 *    places. Anything long and random works; from a terminal:
 *
 *        openssl rand -base64 48
 *
 *    Put it in Apps Script under Project Settings -> Script Properties, named
 *        REPORTS_TICKET_SECRET
 *    and in Supabase under Edge Functions -> Secrets, with the same name.
 *
 *    Do not paste it into a chat, a file in the repo, or an email. It only
 *    needs to exist in those two boxes.
 *
 * 2. Paste this whole file at the end of Code.gs.
 *
 * 3. In doPost's action switch, add:
 *        case 'reportsTicket': return reportsTicket_(user);
 *    matching the style of the cases already there.
 *
 * 4. In reqPerm_(action), add:
 *        case 'reportsTicket': return 'analytics';
 *    Nothing else changes. 'analytics' is reused rather than a new permission
 *    key added, because it already exists and is already held by exactly the
 *    people who should see Reports. If that ever stops being true, change this
 *    line and the matching REQUIRED_PERM constant in the Supabase function.
 *
 * There is no fifth step. No new sheet, no new column, no migration.
 * ---------------------------------------------------------------------------
 */

/* Half an hour. Short enough that a ticket copied out of a browser is close to
 * worthless, long enough that nobody re-mints mid-session. The dashboard
 * refreshes silently when one expires, so this can be shortened without
 * anybody noticing. */
var REPORTS_TICKET_MINUTES = 30;

/**
 * Mint a Reports ticket for an authenticated user.
 *
 * Called only via the action switch, so by the time we are here reqPerm_ has
 * already established that this person holds 'analytics'. The permission check
 * below is therefore a second one, on purpose: this function hands out a
 * credential to another system, and it should not depend on a caller elsewhere
 * in the file having got its wiring right.
 */
function reportsTicket_(user) {
  var secret = PropertiesService.getScriptProperties().getProperty('REPORTS_TICKET_SECRET');
  if (!secret) {
    // Configuration fault, not a permissions one, and worth saying so plainly:
    // silently returning "denied" here would send someone hunting through the
    // user list for a permission problem that does not exist.
    return { ok: false, error: 'REPORTS_TICKET_SECRET is not set in Script Properties' };
  }

  var pub = publicUser_(user);
  var perms = pub.perms || {};
  if (perms.analytics !== true) {
    return { ok: false, error: 'forbidden' };
  }

  var now = Math.floor(Date.now() / 1000);
  var payload = {
    email: pub.email,
    name: pub.name || null,
    // The whole permission set travels, not just the one being checked. It
    // costs nothing and it means adding a second Reports view with a different
    // permission later needs no change on this side.
    perms: perms,
    iat: now,
    exp: now + REPORTS_TICKET_MINUTES * 60
  };

  var payloadB64 = b64UrlEncode_(Utilities.newBlob(JSON.stringify(payload)).getBytes());
  var sig = Utilities.computeHmacSha256Signature(payloadB64, secret);
  var ticket = payloadB64 + '.' + b64UrlEncode_(sig);

  return {
    ok: true,
    ticket: ticket,
    expires_at: payload.exp,
    // Returned so the dashboard can name the person in its header without a
    // second call. It is the same data that is inside the ticket.
    email: pub.email,
    name: pub.name || null
  };
}

/**
 * Base64url, unpadded. The URL-safe alphabet and the stripped '=' are what the
 * Supabase side expects, and both are trivial to get subtly wrong, so this is
 * the only place either is done.
 */
function b64UrlEncode_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

/**
 * Optional, but worth running once after installing.
 *
 * Mints a ticket for whoever owns the script and prints the decoded payload to
 * the log, so a broken secret or a missing permission shows up here rather
 * than as a silent 401 in the browser three steps later.
 */
function testReportsTicket() {
  var secret = PropertiesService.getScriptProperties().getProperty('REPORTS_TICKET_SECRET');
  Logger.log('secret configured: ' + (secret ? 'yes, ' + secret.length + ' chars' : 'NO'));

  var fake = { email: 'test@ardent-training.com', name: 'Test', perms: { analytics: true } };
  var payload = {
    email: fake.email, name: fake.name, perms: fake.perms,
    iat: Math.floor(Date.now() / 1000),
    exp: Math.floor(Date.now() / 1000) + REPORTS_TICKET_MINUTES * 60
  };
  var b64 = b64UrlEncode_(Utilities.newBlob(JSON.stringify(payload)).getBytes());
  var ticket = b64 + '.' + b64UrlEncode_(Utilities.computeHmacSha256Signature(b64, secret || 'unset'));

  Logger.log('sample ticket length: ' + ticket.length);
  Logger.log('payload decodes to: ' + Utilities.newBlob(
    Utilities.base64DecodeWebSafe(b64)).getDataAsString());
}
