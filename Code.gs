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
function slackWebhook_() {
  return PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL') || '';
}

var COURSE_SHEET = 'Course Errors';
var TECH_SHEET = 'Tech Issues';
var INSTRUCTORS_SHEET = 'Instructors';
var ISSUE_SHEETS = [COURSE_SHEET, TECH_SHEET];

// Feedback on the tracker itself (bugs/improvements suggested by users).
var FEEDBACK_SHEET = 'Feedback';
var FEEDBACK_HEADERS = ['id', 'created_at', 'user_email', 'user_name', 'message', 'image_urls', 'status', 'context'];
// context = JSON snapshot of where the Feedback button was pressed (view, open
// issue, filters, viewport, browser), so a report carries its own crime scene.

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

// A "Resolved - TBC" issue auto-resolves after this many days of silence (no
// further reports or objections). The timer is the issue's updated_at, so any
// new activity resets it. 14 days per Edd (21 Jul); the sweep runs on a daily
// trigger created by ensureTriggers_.
var TBC_AUTO_RESOLVE_DAYS = 14;

// Column order for both issue sheets (A..V). Both tabs use the same columns
// so the code can treat them the same.
var HEADERS = [
  'issue_id',          // A
  'submitted_at',      // B
  'updated_at',        // C
  'instructor_name',   // D
  'category',          // E  course_error | tech_issue | internal (internal task/bug, admin-only logging)
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
  'recheck_at'         // AN when to remind the instructor to double-check a Resolved-TBC (blank = no reminder)
];

// The fixed pre-developer troubleshooting checklist for tech issues. Each item
// can end up done (tried), na (not relevant, e.g. an app-only step on a browser
// issue), or todo (relevant but not yet done). The AI pre-fills these from the
// conversation; a person can adjust any of them. Order here is the order shown.
var CHECKLIST_ITEMS = [
  { id: 'confirm_error',          group: 'Identify and record', scope: 'both',    label: "Confirmed exactly what's failing / what the student sees (screenshot if useful)" },
  { id: 'noted_device',           group: 'Identify and record', scope: 'both',    label: 'Noted device make, model, OS version, and browser or app' },
  { id: 'right_place',            group: 'Account and login',   scope: 'both',    label: 'Logging in via the right place (correct partner portal vs ardent-training.com)' },
  { id: 'email_correct',          group: 'Account and login',   scope: 'both',    label: "Email spelled correctly, and it's the one they registered with" },
  { id: 'password_reset',         group: 'Account and login',   scope: 'both',    label: 'Tried "forgot password", then typed email and password manually (no copy-paste)' },
  { id: 'social_signin_password', group: 'Account and login',   scope: 'app',     label: 'Social sign-in: created a password via "organisation -> forgot password"' },
  { id: 'logout_login',           group: 'Standard fixes',      scope: 'both',    label: 'Logged out and back in' },
  { id: 'restart_device',         group: 'Standard fixes',      scope: 'both',    label: 'Restarted the device (or closed and reopened the app/browser)' },
  { id: 'clear_cache',            group: 'Standard fixes',      scope: 'browser', label: 'Cleared cache / tried an incognito or private window' },
  { id: 'app_updated',            group: 'Standard fixes',      scope: 'app',     label: 'Checked the app is up to date' },
  { id: 'different_browser',      group: 'Standard fixes',      scope: 'browser', label: 'Tried a different browser' },
  { id: 'different_device',       group: 'Standard fixes',      scope: 'both',    label: 'Tried a different device' },
  { id: 'different_network',      group: 'Standard fixes',      scope: 'both',    label: 'Tried a different network (mobile data, hotspot, or another wifi)' },
  { id: 'vpn_adblock',            group: 'Standard fixes',      scope: 'both',    label: 'Turned off any VPN, ad blocker, or content/parental filter' },
  { id: 'storage_space',          group: 'Standard fixes',      scope: 'app',     label: "Checked there's free storage on the device (download/save problems)" }
];

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
    if (action === 'ping') return jsonOut({ ok: true, time: new Date().toISOString() });
    if (action === 'getInvite') return jsonOut(getInvite_(p.token));   // public: validate an invite link
    if (action === 'mirror') return jsonOut(mirror_(p));               // read-only, key-gated mirror for the local Cowork sync

    var user = userForToken_(p.token);
    if (!user) return jsonOut({ ok: false, error: 'unauthorized' });
    if (action === 'me') return jsonOut({ ok: true, user: publicUser_(user) });
    if (!hasPerm_(user, reqPerm_(action))) return jsonOut({ ok: false, error: 'forbidden' });

    if (action === 'getIssues') return jsonOut(getIssues_());
    if (action === 'getInstructors') return jsonOut(getInstructors_());
    if (action === 'listUsers') return jsonOut(listUsers_());
    if (action === 'getPlaybook') return jsonOut(getPlaybookEndpoint_());
    if (action === 'listPlaybookSuggestions') return jsonOut(listPlaybookSuggestions_());
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

    // Public auth actions (no session yet).
    // Self-deploy: gated by its own DEPLOY_KEY (script property), not a user
    // session — same pattern as the mirror. See deployBackend_ below.
    if (action === 'deployBackend') return jsonOut(deployBackend_(body));
    if (action === 'setSlackWebhook') return jsonOut(setSlackWebhook_(body));
    if (action === 'runSetup') return jsonOut(runSetup_(body));

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
    if (action === 'me') return jsonOut({ ok: true, user: publicUser_(user) });
    if (action === 'getIssues') return jsonOut(getIssues_());
    if (action === 'getInstructors') return jsonOut(getInstructors_());
    if (action === 'listUsers') return jsonOut(listUsers_());
    if (action === 'getPlaybook') return jsonOut(getPlaybookEndpoint_());
    if (action === 'listPlaybookSuggestions') return jsonOut(listPlaybookSuggestions_());
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
    if (action === 'saveChecklist') return jsonOut(saveChecklist_(body));
    if (action === 'assignIssue') return jsonOut(assignIssue_(body));
    if (action === 'uploadImage') return jsonOut(uploadImage_(body));
    if (action === 'attachImages') return jsonOut(attachImages_(body));
    if (action === 'extract') return jsonOut(extract_(body));
    if (action === 'askIssues') return jsonOut(askIssues_(body));
    if (action === 'suggestFix') return jsonOut(suggestFix_(body));
    if (action === 'troubleshoot') return jsonOut(troubleshoot_(body));
    if (action === 'matchUpdate') return jsonOut(matchUpdate_(body));
    if (action === 'inviteUser') return jsonOut(inviteUser_(body));
    if (action === 'updateUser') return jsonOut(updateUser_(body));
    if (action === 'changePassword') return jsonOut(changePassword_(body));
    if (action === 'adminResetLink') return jsonOut(adminResetLink_(body));
    if (action === 'savePlaybook') return jsonOut(savePlaybook_(body));
    if (action === 'resolvePlaybookSuggestion') return jsonOut(resolvePlaybookSuggestion_(body));
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
  var p = permsOf_(user);
  if (req === 'devcourse') return !!(p.dev || p.course);
  // 'work' = anyone who actually works issues, so they can tick the checklist
  // wherever it shows (Track or the Developers / Course queues).
  if (req === 'work') return !!(p.log || p.manage || p.dev || p.course);
  return !!p[req];
}
function reqPerm_(action) {
  switch (action) {
    case 'addIssue': case 'addUpdate': case 'extract': case 'suggestFix': case 'troubleshoot': case 'matchUpdate': case 'attachImages': return 'log';
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
    case 'saveChecklist': case 'assignIssue': case 'getAssignees': return 'work';
    case 'inviteUser': case 'updateUser': case 'adminResetLink': case 'listUsers':
    case 'getPlaybook': case 'savePlaybook': case 'listPlaybookSuggestions': case 'resolvePlaybookSuggestion':
    case 'getFeedback': case 'updateFeedback': case 'deleteFeedback': return 'users';
    // uploadImage and addFeedback are available to any logged-in user (feedback
    // screenshots, etc.), handled by the default below.
    case 'getIssues': case 'getInstructors': case 'me': return 'any';
    default: return 'any';
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
  var token = newToken_();
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
    var token = newToken_();
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
  found.record.dev_notes = body.dev_notes || '';
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
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- Sheet helpers --------------------------------------------------------

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function sheetByName_(name) { return ss_().getSheetByName(name); }

// course_error -> Course Errors; anything else (incl tech_issue) decided here.
function targetSheetName_(category) {
  var c = String(category).toLowerCase();
  // Internal tasks/bugs live in the Tech Issues sheet, kept apart by their
  // category value, so we don't need a third tab or a sheet migration.
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
      all.push(obj);
    }
  });
  return { ok: true, issues: all };
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
  var now = new Date().toISOString();
  var category = (data.category || 'course_error').toLowerCase();

  // One entry describing this particular report (who hit it, who logged it).
  // We keep its own priority and raw text too, so a wrongly merged report can
  // be split back out cleanly later.
  var report = {
    kind: 'report',
    student_name: data.student_name || '',
    student_contact: data.student_contact || '',
    device_info: data.device_info || '',
    instructor_name: data.instructor_name || '',
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
    category: category,
    raw_text: data.raw_text || '',
    student_name: data.student_name || '',
    student_contact: data.student_contact || '',
    device_info: data.device_info || '',
    course: data.course || '',
    module: data.module || '',
    lesson: data.lesson || '',
    lesson_code: data.lesson_code || '',
    issue_type: data.issue_type || '',
    summary: data.summary || '',
    priority: (data.priority || '').toLowerCase(),
    priority_reason: data.priority_reason || '',
    image_urls: normaliseImageUrls_(data.image_urls),
    // If the instructor has already given the student the suggested fix, this
    // lands as "Resolved - TBC" with that fix saved, and will auto-resolve
    // after a quiet spell unless it comes back.
    status: data.resolved ? 'resolved' : (data.tbc ? 'resolved_tbc' : (data.status || 'open')),
    resolved_at: data.resolved ? (data.resolved_at || now) : '',
    resolution_note: (data.resolved || data.tbc) ? (data.resolution_note || '') : '',
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
    recheck_at: ''
  };

  // Auto-route to the right fix queue (unless the instructor already settled it
  // with a suggested fix). Every course error needs a content fix, so it always
  // goes to the course team. A tech issue goes to the developers when it is high
  // priority (something is badly enough broken to warrant it) OR when the AI
  // judges it a genuine code bug rather than a user-side step.
  // Improvements (feature/enhancement requests) are a calmer backlog: they do
  // NOT auto-route to a fix team, they just sit open for the team to review.
  if (!data.tbc && !data.resolved && issue.request_kind !== 'improvement') {
    if (category === 'course_error') {
      issue.dev_passed_at = new Date().toISOString();
      issue.status = 'with_dev';
    } else if (category === 'internal') {
      // Internal tasks are logged deliberately by an admin, so no AI judgement
      // needed: genuine defects and infrastructure work go straight to the dev
      // queue; content, admin, and feature items sit open for triage.
      if (issue.issue_type === 'bug' || issue.issue_type === 'infrastructure') {
        issue.dev_passed_at = new Date().toISOString();
        issue.status = 'with_dev';
      }
    } else if (category === 'tech_issue' && String(issue.priority).toLowerCase() === 'high') {
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

  return { ok: true, issue: issue, merged: false };
}

// Roll a new report into an existing issue row: add its student to the list,
// bump the priority a level (capped at high), bump the count, and keep an
// audit note in raw_text. Returns the same shape as addIssue_.
function addReportToIssue_(id, data, report) {
  var found = findRow_(id);
  if (!found) return { ok: false, error: 'matched issue not found: ' + id };
  var rec = found.record;

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
  // something that has landed resolved or Resolved - TBC.
  if (String(rec.priority).toLowerCase() === 'high' &&
      rec.status !== 'resolved' && rec.status !== 'resolved_tbc') {
    try { sendSlack_(rec, data.app_url || getAppUrl_()); } catch (e) {}
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

  // Without the manage permission (dev / course team), only the priority may
  // be tweaked - that's the one edit their drawer offers (Edd, 21 Jul).
  if (data._user && !hasPerm_(data._user, 'manage')) {
    var allowedKeys = { priority: 1, priority_reason: 1 };
    var blocked = HEADERS.filter(function (k) {
      return data.hasOwnProperty(k) && !allowedKeys[k] && k !== 'issue_id';
    });
    if (blocked.length) return { ok: false, error: 'Only priority can be edited from this queue.' };
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
    raw_text: note, recommended_steps: data.recommended_steps || null, date: new Date().toISOString()
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
  // in, reopen it (the suggested fix evidently did not stick).
  if (String(rec.status).toLowerCase() === 'resolved_tbc') rec.status = 'open';

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
  }

  rec.updated_at = new Date().toISOString();
  found.sheet.getRange(found.rowNum, 1, 1, HEADERS.length).setValues([recordToRow_(rec)]);

  if (String(rec.priority).toLowerCase() === 'high' &&
      rec.status !== 'resolved' && rec.status !== 'resolved_tbc') {
    try { sendSlack_(rec, data.app_url || getAppUrl_()); } catch (e) {}
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
  rec.status = 'dev_fixed';
  if (data.dev_notes != null) rec.dev_notes = data.dev_notes;
  rec.updated_at = new Date().toISOString();
  found.sheet.getRange(found.rowNum, 1, 1, HEADERS.length).setValues([recordToRow_(rec)]);
  if (data.notify_student) {
    try { sendNotifyStudentSlack_(rec, data.app_url || getAppUrl_()); } catch (e) {}
  }
  return { ok: true };
}

function sendNotifyStudentSlack_(issue, appUrl) {
  var isCourse = String(issue.category).toLowerCase() === 'course_error';
  var student = (issue.student_name || '') + (issue.student_contact ? ' (' + issue.student_contact + ')' : '');
  var text = [
    ':white_check_mark: *' + (isCourse ? 'Course fix done' : 'Fix done') + ' - student to notify*',
    '*Lesson:* ' + (issue.lesson || '-') + ' (' + (issue.lesson_code || '-') + ')',
    '*Summary:* ' + (issue.summary || '-'),
    '*Fix notes:* ' + (issue.dev_notes || '-'),
    '*Student:* ' + (student || '-'),
    '*Logged by:* ' + (issue.instructor_name || '-'),
    '',
    'Instructors: please let the student know this is sorted (it\'s also in your Actions tab).',
    'Open this issue: ' + issueLink_(issue, appUrl)
  ].join('\n');
  UrlFetchApp.fetch(slackWebhook_(), {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ text: text })
  });
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
  var isCourse = String(issue.category).toLowerCase() === 'course_error';
  var toInstructor = issue.dev_query_target === 'instructor';
  var text = [
    ':grey_question: *' + (issue.dev_query_by || 'Someone') + ' has a question for ' +
      (toInstructor ? (issue.instructor_name || 'the instructor') : 'the admins') + '*',
    '*Lesson:* ' + (issue.lesson || '-') + ' (' + (issue.lesson_code || '-') + ')',
    '*Summary:* ' + (issue.summary || '-'),
    '*Question:* ' + (issue.dev_query || '-'),
    '',
    toInstructor
      ? ((issue.instructor_name || 'The instructor') + ', reply from your Actions tab so the fix can carry on.')
      : ('Reply in the tracker so ' + (isCourse ? 'the course team' : 'the developer') + ' can carry on.'),
    'Open this issue: ' + issueLink_(issue, appUrl)
  ].join('\n');
  UrlFetchApp.fetch(slackWebhook_(), {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ text: text })
  });
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
  var text = [
    ':speech_balloon: *Question answered*' + (asker ? ' - ' + asker + ', this one\'s for you' : ''),
    '*Lesson:* ' + (issue.lesson || '-') + ' (' + (issue.lesson_code || '-') + ')',
    '*Summary:* ' + (issue.summary || '-'),
    '*Question:* ' + (question || '-'),
    '*Reply:* ' + (reply || '-'),
    '',
    'Open this issue: ' + issueLink_(issue, appUrl)
  ].join('\n');
  UrlFetchApp.fetch(slackWebhook_(), {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ text: text })
  });
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
  var text = [
    ':alarm_clock: *Workaround needs a double-check*',
    '*Lesson:* ' + (rec.lesson || '-') + ' (' + (rec.lesson_code || '-') + ')',
    '*Summary:* ' + (rec.summary || '-'),
    '*Workaround given:* ' + (rec.resolution_note || '-'),
    '',
    (rec.instructor_name || 'Whoever logged this') + " couldn't check at the time: is this a one-off for that student, or a fault for everyone? " +
      "One-off - leave it as Resolved - TBC. Everyone - reopen it so it gets a proper fix.",
    'Open this issue: ' + issueLink_(rec, appUrl)
  ].join('\n');
  UrlFetchApp.fetch(slackWebhook_(), {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ text: text })
  });
}

// Kept only so any leftover daily trigger from the first Round 13 deploy
// doesn't error: recheck reminders now go to Slack immediately (see
// requestRecheck_), so there is nothing for a scheduled pass to do. The
// trigger itself is removed by ensureTriggers_ next time setup() runs.
function sendRecheckReminders() {}

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
  try {
    UrlFetchApp.fetch(slackWebhook_(), {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ text: text })
    });
  } catch (e) {}
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
    return { ok: true, version: v.versionNumber, deployment: target.deploymentId };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

// Keep the checklist-review (monthly) trigger in place, and clear out the old
// daily recheck trigger if one exists (rechecks ping Slack immediately now).
// Called from setup(), safe to run repeatedly.
function ensureTriggers_() {
  var haveMonthly = false, haveTbc = false, haveBackup = false, haveDigest = false;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendRecheckReminders') ScriptApp.deleteTrigger(t);
    if (t.getHandlerFunction() === 'monthlyChecklistReview') haveMonthly = true;
    if (t.getHandlerFunction() === 'autoResolveTbc') haveTbc = true;
    if (t.getHandlerFunction() === 'weeklyBackup') haveBackup = true;
    if (t.getHandlerFunction() === 'weeklyDigest') haveDigest = true;
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
}

// Monday-morning state of the tracker, posted to Slack so the team sees where
// things stand without opening the app. Runs on a weekly trigger (8am, after
// the 7am backup).
function weeklyDigest() {
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
  try {
    UrlFetchApp.fetch(slackWebhook_(), {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ text: lines.join('\n') })
    });
  } catch (e) {}
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

// Run setup() remotely (DEPLOY_KEY gated), so schema/trigger changes shipped
// via deployBackend don't need anyone in the editor either.
function runSetup_(data) {
  var key = PropertiesService.getScriptProperties().getProperty('DEPLOY_KEY');
  if (!key || String(data.key || '') !== key) return { ok: false, error: 'bad deploy key' };
  setup();
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

  // Drive access needs its own OAuth grant. If the script was redeployed
  // without re-authorising, this is where it fails, so give a clear error the
  // front-end can show instead of a bare exception.
  var folder, file;
  try {
    folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    file = folder.createFile(blob);
  } catch (e) {
    return { ok: false, error: 'Drive access is not authorised for this deployment. An admin needs to open the Apps Script editor, run authorizeDrive() once, approve the permission, and redeploy. (' + String(e) + ')' };
  }
  // Sharing can be blocked by a Workspace policy even when the upload worked.
  // Keep the file either way; a private link is better than losing the image.
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

  var id = file.getId();
  return {
    ok: true,
    file_id: id,
    url: 'https://drive.google.com/uc?export=view&id=' + id,
    open_url: 'https://drive.google.com/file/d/' + id + '/view'
  };
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

  var prompt = buildExtractionPrompt_(rawText);
  var payload = {
    // A single thread can now split into several full issue objects, so give the
    // model enough room. 1024 truncated the JSON mid-string on long multi-topic
    // threads (exactly the ones this splitting is for) and the parse then failed.
    model: EXTRACTION_MODEL,
    max_tokens: 4096,
    messages: [{ role: 'user', content: prompt }]
  };

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload)
  });

  var code = res.getResponseCode();
  var bodyText = res.getContentText();
  if (code < 200 || code >= 300) {
    return { ok: false, error: 'Anthropic API error ' + code + ': ' + bodyText };
  }

  var parsed = JSON.parse(bodyText);
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
  catch (parseErr) { return { ok: false, error: 'Could not parse model output as JSON', raw: text }; }

  return { ok: true, fields: fields };
}

// Given a new issue and a set of past RESOLVED issues (with the fix that was
// used), ask the AI whether this is a known, already-solved problem and, if so,
// recommend the fix to the instructor. Returns { found, fix, based_on }.
function suggestFix_(data) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: true, found: false };

  var newIssue = data.new_issue || {};
  var candidates = data.candidates || [];
  if (!candidates.length) return { ok: true, found: false };

  var prompt = 'An instructor has just logged an issue with an online sailing course platform (Ardent Training). ' +
    'Below are PAST issues that were already resolved, each with the fix that was applied. ' +
    'If the new issue clearly matches one of these known, already-solved problems, write a short recommended fix for the instructor: ' +
    '1 to 3 plain, practical sentences they can act on or pass to the student. ' +
    'Base it on how the matching past issue was actually resolved.\n\n' +
    'Hold a HIGH bar for "found". A genuine match means the SAME failure mode in the same part of the platform - ' +
    'the same thing failing in the same way, not merely the same lesson, the same device, or a similar-sounding symptom. ' +
    'Never pad a weak match into advice, and never suggest generic steps (restart, reinstall, clear cache, log out and in) ' +
    'unless the matching past issue was genuinely resolved by exactly that step - the logging form already walks instructors ' +
    'through the generic checklist, so repeating it here is noise. A wrong suggestion wastes the student\'s time and the ' +
    'instructor\'s trust; when in doubt, return found false. Most new issues do NOT have a matching past fix.\n\n' +
    'NEW issue:\n' + JSON.stringify(newIssue) + '\n\n' +
    'PAST resolved issues (summary + how it was fixed, most relevant first):\n' + JSON.stringify(candidates) + '\n\n' +
    'Return ONLY JSON: {"found": true or false, "fix": "<the recommendation, or empty string>", "based_on": "<short reference to the matching past issue, or empty string>"}. No prose, no markdown fences.';

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
  var text = '';
  if (parsed.content) for (var i = 0; i < parsed.content.length; i++) {
    if (parsed.content[i].type === 'text') text += parsed.content[i].text;
  }
  text = text.trim().replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
  var out; try { out = JSON.parse(text); } catch (e) { return { ok: true, found: false }; }

  if (!out || !out.found || !out.fix) return { ok: true, found: false };
  return { ok: true, found: true, fix: String(out.fix), based_on: String(out.based_on || '') };
}

// The default tech troubleshooting playbook. The live one is stored as a script
// property (PLAYBOOK_TEXT) so it can be edited from the Admin page and grow as
// we learn from resolved issues; this constant is the fallback / starting point.
var DEFAULT_PLAYBOOK = [
  'ARDENT TECH TROUBLESHOOTING PLAYBOOK (for the instructor helping a student):',
  '',
  'KNOWN ACCOUNT ISSUES (check these first, each has a specific fix):',
  '- Cannot purchase a course / error at purchase: they probably already have an account from a free trial. Get them to log in first (search their email), then purchase.',
  '- Login not recognising email / login problems: they likely bought through a partner school but are trying to log in at ardent-training.com. Check the instructor portal (services.ardent-training.com) and send them to the correct partner portal.',
  '- Double-check they are not misspelling their email (a screenshot confirms this).',
  '- They may be using a different email than the one they registered with: check the instructor portal.',
  '- If still stuck: get them to use "forgot password" to reset, then type email and password manually (no copy-paste).',
  '- Still cannot log in: reset their password yourself in the students tab of the instructor portal and try logging in as them. If you can log in, it is a device/browser issue on their end, move to the step-by-step list.',
  '- APP login for practical partner students: if they signed up with a social sign-in (Google/Facebook) they may have no password, and social sign-in does not work on the app for partner-school bookings. They must tap "organisation", find their school, then "forgot password" to create a password (or do it for them).',
  '',
  'BROWSER (web) ISSUES, in order, confirming each step with the student:',
  '1. Log out of Ardent Training and log back in.',
  '2. Try an incognito/private window. If that works, clear cache and/or disable extensions for a permanent fix.',
  '3. Try a different browser. If that works, clear cache and/or disable extensions.',
  '4. Try a different device. If that works, note the make, model and OS version of the one with the problem.',
  '5. Try a different internet connection/wifi. If that works, note their ISP.',
  '6. Still stuck: get screenshots and note device make/model/browser/OS, then log it here so it reaches the team. Tell the student it is with the developers and apologise.',
  '',
  'APP (mobile app) ISSUES, in order:',
  '1. Check on your own device. If it fails for you too, try the web version; if it only fails on the app, tell them to use the web version and flag it.',
  '2. Log out and back in.',
  '3. Restart the device.',
  '4. Check the app is up to date.',
  '5. Try a different device.',
  '6. Try a different network (mobile data, a phone hotspot, or another wifi are all the same step). A "check network connection" message often means something on that network or device is blocking the app, not that the internet is down.',
  '7. Turn off any VPN, ad blocker, or content/parental filter on the device.',
  '8. If it is about downloading lessons, check there is free space on the device.',
  '9. Read the pattern: if it works on some of their devices or networks but not others (e.g. fine on a phone, failing on tablets), it is more likely a setting on the failing device or network than a bug in our app, so rule the above out before escalating.',
  '10. Still stuck after genuinely trying all of the above: get screenshots and note device make/model/OS, then log it here so it reaches the team. Tell the student it is with the developers and apologise.'
].join('\n');

// Read the conversation the instructor pasted, work out what has been tried,
// and suggest the next steps from the playbook (pointing out anything missed).
function troubleshoot_(data) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: true, found: false };
  var raw = data.raw_text || '';
  if (!raw) return { ok: true, found: false };

  var prompt = 'You are helping an Ardent Training instructor troubleshoot a student tech issue. ' +
    'Below is the troubleshooting playbook, then the conversation or notes the instructor pasted. ' +
    'Work out what has ALREADY been tried in the conversation, then list the NEXT things the instructor should get the student to try, in the playbook order, skipping anything already done. ' +
    'Only count a step as already tried if it is EXPLICITLY described in the conversation or history. Do not assume or infer that a step was tried. In particular, the app working on another device (such as their phone) does NOT mean a different network was tried, so never list "a different network" as already tried unless the student actually says they tried one. ' +
    'Be helpful and thorough: even if a lot has been tried, there are almost always remaining steps, so list every relevant playbook step that has not been explicitly done rather than concluding nothing is left. ' +
    'Pay attention to the pattern: if it works on some of the student\'s devices or networks but not others (for example fine on a phone but failing on tablets, or a "check network connection" message while other things work), that points to a setting on the failing device or network rather than a bug in our app, so suggest ruling those out first (a different network, and turning off any VPN, ad blocker or content filter). ' +
    'Keep the list short, distinct and non-overlapping: aim for 2 to 4 genuinely different next steps and never repeat the same step in different words. Treat mobile data, a phone hotspot and a different wifi as ONE step (trying a different network); if any different network has already been tried, do NOT suggest another network step. ' +
    'Do NOT tell them to escalate to Edd, Charlie, Stu or anyone, or to message Slack; that happens automatically when an issue is high priority. Once the relevant steps have genuinely all been tried, just say to submit it so it reaches the team. ' +
    'Point out any step that seems to have been missed or done out of order. If it matches one of the known account issues, name it and give that specific fix first. ' +
    'Decide from the conversation whether it is a browser/web issue or a mobile app issue and use the matching list. ' +
    'Focus every step on the thing that is actually FAILING. If another route already works for them (for example the website works, or it works on another device), that is only a temporary workaround, so do NOT suggest troubleshooting the part that already works. A "different network" step means getting them to try the FAILING thing (e.g. the app) on a different network (mobile data, a phone hotspot, or another wifi, which are all one and the same step), never trying the website on a different network. ' +
    'Only suggest checking free storage when the problem is about downloading or saving content, not for a login or "check network connection" problem. ' +
    'Keep each suggestion short and practical, addressed to the instructor. Do not mention filling out any external form.\n\n' +
    'PLAYBOOK:\n' + getPlaybook_() + '\n\n' +
    'CHECKLIST ITEMS (the team ticks these off before an issue reaches the developers). For each one, decide its state from the conversation:\n' +
    checklistItemsForPrompt_() + '\n' +
    'State rules: "done" only if the conversation EXPLICITLY says that step was tried (same strict rule as the steps above; the same network caveat applies). ' +
    '"na" if the step is not relevant to THIS issue, for example app-only steps (app up to date, social sign-in password, free storage) on a browser/web issue, browser-only steps (cleared cache/incognito, a different browser) on a mobile-app issue, the login/account steps when it is not a login problem, or free storage when it is not a download/save problem. ' +
    'Universal steps are NEVER "na": restarting the device, logging out and back in, trying a different device, and trying a different network apply to every platform and every kind of issue, so they can only be "done" or "todo". ' +
    '"todo" if the step is relevant but has not been done yet. When unsure between done and todo, choose todo. When unsure between na and todo, choose todo.\n\n' +
    (data.existing_history ? 'EARLIER HISTORY ON THIS SAME ISSUE (already logged; treat anything here as already tried):\n"""\n' + data.existing_history + '\n"""\n\n' : '') +
    'NEW CONVERSATION / NOTES:\n"""\n' + raw + '\n"""\n\n' +
    'Return ONLY JSON: {"found": true or false, "steps": ["short next step", ...], "escalate": true or false, "note": "<one short line such as an escalation note, or empty string>", "checklist": {"<item id>": "done | na | todo", ...}}. ' +
    'Include every checklist item id in the checklist object. ' +
    'Set found false only if there is genuinely nothing useful to suggest. No prose, no markdown fences.';

  var res;
  try {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 600, messages: [{ role: 'user', content: prompt }] })
    });
  } catch (e) { return { ok: true, found: false }; }
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) return { ok: true, found: false };

  var parsed; try { parsed = JSON.parse(res.getContentText()); } catch (e) { return { ok: true, found: false }; }
  var text = '';
  if (parsed.content) for (var i = 0; i < parsed.content.length; i++) {
    if (parsed.content[i].type === 'text') text += parsed.content[i].text;
  }
  text = text.trim().replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
  var out; try { out = JSON.parse(text); } catch (e) { return { ok: true, found: false }; }
  var checklist = normaliseChecklist_(out.checklist);
  if (!out || !out.found || !out.steps || !out.steps.length) return { ok: true, found: false, checklist: checklist };
  return { ok: true, found: true, steps: out.steps, escalate: !!out.escalate, note: out.note || '', checklist: checklist };
}

// A compact, numbered list of the checklist items for the AI prompt, with a
// scope hint so it can mark app/browser-only steps "na" on the wrong platform.
function checklistItemsForPrompt_() {
  return CHECKLIST_ITEMS.map(function (it) {
    var scope = it.scope === 'app' ? ' [app only]' : it.scope === 'browser' ? ' [browser/web only]' : '';
    return '- ' + it.id + ': ' + it.label + scope;
  }).join('\n');
}

// Keep only known item ids and valid states, so a wobbly AI reply cannot put
// junk in the checklist. Anything missing or odd is left out (the front-end
// treats a missing item as "todo").
function normaliseChecklist_(obj) {
  var clean = {};
  if (!obj || typeof obj !== 'object') return clean;
  CHECKLIST_ITEMS.forEach(function (it) {
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

function addFeedback_(data) {
  var sheet = sheetByName_(FEEDBACK_SHEET);
  if (!sheet) return { ok: false, error: 'Feedback sheet missing. Run setup() once.' };
  if (!data.message && !data.image_urls) return { ok: false, error: 'Add a message first.' };
  var u = data._user || {};
  var row = {
    id: Utilities.getUuid(),
    created_at: new Date().toISOString(),
    user_email: u.email || '',
    user_name: u.name || '',
    message: data.message || '',
    image_urls: normaliseImageUrls_(data.image_urls),
    status: 'new',
    context: typeof data.context === 'string' ? data.context : (data.context ? JSON.stringify(data.context) : '')
  };
  sheet.appendRow(FEEDBACK_HEADERS.map(function (k) { return row[k]; }));
  return { ok: true };
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
    var o = {}; FEEDBACK_HEADERS.forEach(function (k) { o[k] = values[r][idx[k]]; });
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

function buildExtractionPrompt_(rawText) {
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
    '- category: "course_error" if the problem is with the lesson content or teaching material itself (wrong information, a confusing or incorrect explanation, a typo in a lesson, a mislabelled diagram, a quiz answer being wrong, OR a specific lesson that will not open / shows a 404 / page not found, which usually means that lesson was not uploaded properly and the course team needs to re-upload it). "tech_issue" if the problem is with the platform, website or app generally: video or audio not playing, login or access problems, a button that does not work, progress not saving, or anything device or browser specific.',
    '- likely_internal: true if this is NOT a student-facing problem at all but an internal one: instructors talking to each other about the instructor portal, the partner portal, admin tools, or company systems, with no student blocked from learning. A conversation between staff about wrong data shown in the instructor portal is internal. A student unable to watch a video is not. Return false when in doubt.',
    '- section: which part of the platform the problem lives in: one of ["website", "instructor_portal", "partner_portal", "course_player", "app", "other"], or null if unclear. "website" is the public ardent-training.com site, "course_player" is where students take lessons, "app" is the mobile app.',
    '- student_name: string or null',
    '- student_contact: the student email address if it appears anywhere in the text (prefer an email over a phone number). Return null if no email or contact is found, so the instructor can be asked for it.',
    '- device_info: the students device, operating system and browser if mentioned (e.g. "iPad, iOS 17, Safari" or "Windows 11, Chrome"), or null',
    '- course: one of the canonical course names above, but ONLY when the issue is actually about that course\'s content or a student\'s access to it. If the problem is with the instructor portal, partner portal, website, or another internal system, return null even if a course happens to be mentioned in passing. Do not guess a course from weak clues.',
    '- module: module title string or null',
    '- lesson: the FULL slide/question code exactly as written when one appears in the text (e.g. "EN.06.03.09" or "DS.10.19.09.2.M", one long string, not broken down), otherwise the lesson title if known, or null',
    '- lesson_code: lesson code string (e.g. DS.09.04) or null',
    '- issue_type: one of ["bug", "content_error", "student_confusion", "access_problem", "other"]',
    '- request_kind: "improvement" if the report is asking for a NEW feature, an enhancement, or an "it would be nice if" change rather than reporting something broken or wrong (this applies to both course content and the platform, for example "could we add a glossary" or "the player should remember playback speed"); otherwise "fix" for a bug, an error, or something not working or incorrect as it stands. When in doubt, choose "fix". Most reports are "fix".',
    '- media_kind: for a course_error only, which part of the lesson it concerns: "video" if it is about a video or animation, "text" if it is about written text, a diagram, or quiz wording, otherwise "other". Return null for tech_issue.',
    '- impact: for an improvement only, a rough impact rating of "low", "medium", or "high" based on how much it would benefit students. Return null for a fix.',
    '- summary: one or two plain-English sentences summarising the issue. Keep the specific detail someone would need to reproduce it: which page or view, and HOW it is reached when that matters (e.g. "opened via the three-dots menu on the Students page" rather than just "the student profile page"). If the report describes two different symptoms, name both rather than blending them into one vague sentence.',
    '- priority: one of ["high", "medium", "low"]',
    '- priority_reason: one sentence explaining the priority',
    '- resolution_status: "resolved" if the pasted conversation shows this problem was ALREADY sorted out in the chat itself (an instructor gave a definitive answer or fix, the thread says "Conversation was marked resolved by ...", or the student confirms it works now, e.g. "that solved it", "thanks, working now", "all good"). "tbc" if a fix or answer was given but the student has not yet confirmed it worked. "open" if it is still unresolved, or was only logged to hand to the developers. When in doubt, "open".',
    '- resolution_note: when resolution_status is "resolved" or "tbc", one or two sentences stating the actual answer or fix that was given (what resolved it), otherwise null.',
    '- resolved_by: the staff member who resolved it or gave the answer (from "marked resolved by X", or whoever replied with the fix), or null.',
    '- resolved_at: the date the resolution happened if it can be read from the text (ISO 8601 if possible, otherwise the date as written), or null.',
    '- sub_issues: a single pasted thread can hold SEVERAL separate problems raised over time (different pages, features, slides, or topics, each fixed independently, and each possibly at a different date or already resolved). If it holds more than one, return an array with one FULL entry per distinct problem, each carrying the SAME fields as above (category, likely_internal, section, student_name, student_contact, device_info, course, module, lesson, lesson_code, issue_type, request_kind, media_kind, impact, summary, priority, priority_reason, resolution_status, resolution_note, resolved_by, resolved_at). Put the primary or most urgent problem in the top-level fields AND as the FIRST array entry, so the array is the complete set. If it is really one problem (or one problem with knock-on effects), return null. Never split a single problem, and never blend unrelated topics into one entry.',
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
    '- high: a factual or safety-critical content error; OR a tech issue where the troubleshooting steps relevant to this situation have genuinely been tried and the student is still completely blocked with no workaround. Judge relevance by the case: for a browser/website problem the relevant fixes are things like clearing cache, an incognito window, a different browser, a different network; for a mobile-app problem they are reinstalling or updating the app, restarting, a different device, a different network. Do NOT expect app-only fixes for a browser problem or browser-only fixes for an app problem, and do not hold back from high just because an irrelevant fix was not tried.',
    '- high ALSO covers an outage that hits everyone rather than one student: a page, resource, video host, or the site itself down or erroring for all users (e.g. a resource page returning an error). User-side troubleshooting cannot fix a down server, so never hold one of these at medium because steps were not tried.',
    '- medium: a real problem but the student is not fully blocked, has a workaround, or the relevant fixes have not all been tried yet.',
    '- low: minor or cosmetic, a one-off, or something very likely solved by a simple relevant step the student has not tried yet.',
    '',
    'Raw text:',
    '"""',
    rawText,
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

function sendSlack_(issue, appUrl) {
  var c = String(issue.category).toLowerCase();
  var area = c === 'tech_issue' ? 'Tech issue' : (c === 'internal' ? 'Internal' : 'Course error');
  if (issue.section) area += ' · ' + String(issue.section).replace(/_/g, ' ');
  var text = [
    ':red_circle: *High priority issue logged* (' + area + ')',
    '*Lesson:* ' + (issue.lesson || '-') + ' (' + (issue.lesson_code || '-') + ')',
    '*Type:* ' + (issue.issue_type || '-'),
    '*Summary:* ' + (issue.summary || '-'),
    '*Student:* ' + (issue.student_name || '-') + ' (' + (issue.student_contact || '-') + ')',
    '*Device:* ' + (issue.device_info || '-'),
    '*Logged by:* ' + (issue.instructor_name || '-'),
    '*Submitted:* ' + (issue.submitted_at || new Date().toISOString()),
    '',
    'View in Bugs: ' + issueLink_(issue, appUrl)
  ].join('\n');

  UrlFetchApp.fetch(slackWebhook_(), {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({ text: text })
  });

  return { ok: true };
}

function getAppUrl_() {
  return PropertiesService.getScriptProperties().getProperty('APP_URL') || '';
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
function setup() {
  var ss = ss_();

  ISSUE_SHEETS.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
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
    var token = newToken_();
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

  Logger.log('Setup complete. Course Errors, Tech Issues, Instructors, Users, and Feedback sheets are ready. ' +
    'If APP_URL is not set yet, set it after deploying then run adminInviteLink() to get your setup link.');
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
      payload: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 1000, messages: [{ role: 'user', content: prompt }] })
    });
  } catch (e) { return { ok: false, error: 'AI request failed: ' + e }; }
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
    return { ok: false, error: 'AI request failed (' + res.getResponseCode() + ').' };
  }
  var parsed; try { parsed = JSON.parse(res.getContentText()); } catch (e) { return { ok: false, error: 'Bad AI response.' }; }
  var text = '';
  if (parsed.content) for (var i = 0; i < parsed.content.length; i++) {
    if (parsed.content[i].type === 'text') text += parsed.content[i].text;
  }
  if (!text.trim()) return { ok: false, error: 'Empty AI response.' };
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
