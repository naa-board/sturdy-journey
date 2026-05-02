/**
 * North Aldercroft Association — Meet Notes → Website (Automatic)
 *
 * Two modes:
 *
 *   AUTOMATIC  A monthly Apps Script trigger (1st of each month) watches a
 *              Google Drive folder. When Google Meet saves new Gemini AI notes
 *              there, the script picks them up, calls Claude to generate a
 *              summary, and opens a GitHub pull request — no human action required.
 *
 *   MANUAL     A board member can also trigger it from inside any open
 *              Google Doc via Extensions → NAA Website → Post summary…
 *
 * In both cases, a board member reviews and merges the GitHub PR to publish.
 *
 * ─── ONE-TIME SETUP ──────────────────────────────────────────────────────────
 *
 * 1. Open script.google.com → New project. Paste this file in.
 *
 * 2. Project Settings → Script Properties — add these four values:
 *
 *      ANTHROPIC_API_KEY    Your Anthropic API key  (console.anthropic.com)
 *      GITHUB_TOKEN         GitHub fine-grained personal access token
 *                           (Settings → Developer settings → Fine-grained tokens)
 *                           Needs Contents and Pull requests: Read & Write
 *                           on the scott4382/sturdy-journey repo.
 *      NOTIFY_EMAIL         Email address to notify when a PR is ready (optional)
 *
 * 3. Fill in DRIVE_FOLDER_ID below.
 *    – In Google Drive, tell Meet where to save notes:
 *      Meet settings → Recording → Choose a folder → pick one
 *    – Copy the folder ID from its URL:
 *      drive.google.com/drive/folders/THIS_PART_HERE
 *    – Paste it as the value of DRIVE_FOLDER_ID below.
 *
 * 4. Change GITHUB_BRANCH to 'main' once the site is live.
 *
 * 5. Run  setupAutomaticTrigger()  once from the Apps Script editor.
 *    That's it — the script will run on the 1st of each month from then on.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 */

// ── Config ────────────────────────────────────────────────────────────────────

const DRIVE_FOLDER_ID = 'YOUR_FOLDER_ID_HERE'; // ← paste your Drive folder ID
const GITHUB_OWNER    = 'naa-board';
const GITHUB_REPO     = 'legendary-umbrella';
const GITHUB_BRANCH   = 'main';                // ← change if your branch differs
const CLAUDE_MODEL    = 'claude-sonnet-4-6';

// ── Automatic trigger setup ───────────────────────────────────────────────────

/**
 * Run this ONCE from the Apps Script editor.
 * Installs a monthly trigger (1st of each month) — after that everything is automatic.
 */
function setupAutomaticTrigger() {
  // Remove any previous automatic triggers to avoid duplicates
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkForNewNotes')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('checkForNewNotes')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  Logger.log('Monthly trigger installed. The script will check for new meeting notes on the 1st of each month.');

  // Also install the manual Doc menu trigger
  installDocMenuTrigger();
}

// ── Automatic: poll Drive folder ──────────────────────────────────────────────

/**
 * Runs on the 1st of each month. Finds any new Google Docs in DRIVE_FOLDER_ID
 * that haven't been processed yet, and creates a GitHub PR for each one.
 */
function checkForNewNotes() {
  const props    = PropertiesService.getScriptProperties();
  const folder   = DriveApp.getFolderById(DRIVE_FOLDER_ID);

  // Look for Google Docs created in the last 35 days to catch any meetings from the past month
  const since    = new Date(Date.now() - 35 * 24 * 60 * 60 * 1000);
  const files    = folder.getFilesByType(MimeType.GOOGLE_DOCS);

  const toProcess = [];
  while (files.hasNext()) {
    const file = files.next();
    if (file.getDateCreated() >= since && !props.getProperty('processed_' + file.getId())) {
      toProcess.push(file);
    }
  }

  if (toProcess.length === 0) return;

  Logger.log(`Found ${toProcess.length} new file(s) to process.`);

  toProcess.forEach(file => {
    try {
      Logger.log(`Processing: ${file.getName()} (${file.getId()})`);
      runPipelineForDoc_(DocumentApp.openById(file.getId()), file.getName());
      props.setProperty('processed_' + file.getId(), new Date().toISOString());
    } catch (e) {
      Logger.log(`Error processing ${file.getName()}: ${e.message}`);
      notifyError_(file.getName(), e);
    }
  });
}

// ── Manual: Google Doc menu ───────────────────────────────────────────────────

function onOpen() {
  DocumentApp.getUi()
    .createMenu('NAA Website')
    .addItem('Post summary to website…', 'showConfirmDialog')
    .addToUi();
}

function installDocMenuTrigger() {
  try {
    // For standalone scripts the onOpen trigger is automatic;
    // this is only needed if you want it in a specific Doc.
    Logger.log('Doc menu trigger: use Extensions > Apps Script from inside a Doc to bind it there.');
  } catch (e) { /* not bound to a Doc */ }
}

function showConfirmDialog() {
  const ui  = DocumentApp.getUi();
  const doc = DocumentApp.getActiveDocument();

  const result = ui.alert(
    'Post summary to website',
    `Send "${doc.getName()}" to Claude and open a GitHub pull request?\n\n` +
    `A board member will review and merge the PR to publish.`,
    ui.ButtonSet.YES_NO
  );

  if (result === ui.Button.YES) {
    try {
      const prUrl = runPipelineForDoc_(doc, doc.getName());
      ui.alert('Done!', `Pull request created:\n${prUrl}\n\nReview and merge it to publish.`, ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('Error', e.message, ui.ButtonSet.OK);
    }
  }
}

// ── Core pipeline ─────────────────────────────────────────────────────────────

function runPipelineForDoc_(doc, docTitle) {
  const notes = extractDocumentText_(doc);

  if (notes.trim().length < 100) {
    throw new Error('Document is too short to summarise. Is this the right file?');
  }

  const { html, title, dateLabel, slug } = generateSummary_(notes, docTitle);
  const prUrl = createPullRequest_(html, title, dateLabel, slug);

  notifyPrReady_(title, prUrl);
  return prUrl;
}

// ── Document text extraction ──────────────────────────────────────────────────

function extractDocumentText_(doc) {
  return doc.getBody().getParagraphs()
    .map(p => p.getText())
    .filter(t => t.trim())
    .join('\n');
}

// ── Claude summarisation ──────────────────────────────────────────────────────

function generateSummary_(notes, docTitle) {
  const apiKey = getProperty_('ANTHROPIC_API_KEY');
  const today  = new Date().toISOString().slice(0, 10);

  const prompt =
`You are helping maintain the website for the North Aldercroft Association, a private road HOA in the Santa Cruz Mountains. Below are the AI-generated notes from a board meeting (produced by Google Meet's Gemini feature). Your job is to turn them into a clear, friendly summary that residents can quickly read.

MEETING NOTES:
${notes}

Write a meeting summary following these rules exactly:
1. Tone: Plain, friendly language. Write for neighbors, not lawyers. No jargon.
2. Start with a "The Short Version" section (2–3 sentences max) covering the most important decisions or news.
3. Add 3–6 more sections, each covering one topic. Use a short descriptive heading — be specific (e.g. "Spring Road Repair", "Dues Update"), not generic (e.g. "Discussion").
4. End with a "What's Next" section covering the next meeting date (if mentioned) or key upcoming actions.
5. Keep each section to 2–4 sentences. Do not pad.
6. If road work, dues, costs, or deadlines were discussed, include them clearly.

Output ONLY the following structured response — no extra explanation:

TITLE: [short meeting title, e.g. "March 2025 Board Meeting Summary"]
DATE_ISO: [meeting date as YYYY-MM-DD, inferred from the notes; use ${today} if unclear]
DATE_LABEL: [human-readable, e.g. "March 15, 2025"]
MINUTES_FILE: [filename of the formal minutes page if one exists, e.g. 2025-03.html; leave blank if none]
---HTML_BODY---
[section headings and paragraphs as HTML — <h2> and <p> tags only, no wrapper divs]
---END---`;

  const res  = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    headers: {
      'x-api-key':          apiKey,
      'anthropic-version':  '2023-06-01',
      'content-type':       'application/json',
    },
    payload: JSON.stringify({
      model:      CLAUDE_MODEL,
      max_tokens: 2048,
      messages:   [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    throw new Error(`Claude API error ${res.getResponseCode()}: ${res.getContentText()}`);
  }

  const text        = JSON.parse(res.getContentText()).content[0].text;
  const title       = (text.match(/^TITLE:\s*(.+)$/m)      || [])[1]?.trim() || 'Board Meeting Summary';
  const dateIso     = (text.match(/^DATE_ISO:\s*(.+)$/m)    || [])[1]?.trim() || today;
  const dateLabel   = (text.match(/^DATE_LABEL:\s*(.+)$/m)  || [])[1]?.trim() || today;
  const minutesFile = (text.match(/^MINUTES_FILE:\s*(.*)$/m)|| [])[1]?.trim() || '';
  const htmlBody    = (text.match(/---HTML_BODY---([\s\S]*?)---END---/) || [])[1]?.trim() || '';
  const slug        = `${dateIso}-summary`;

  return { html: buildHtmlPage_(title, dateLabel, slug, htmlBody, minutesFile), title, dateLabel, slug };
}

// ── HTML page builder ─────────────────────────────────────────────────────────

function buildHtmlPage_(title, dateLabel, slug, htmlBody, minutesFile) {
  const minutesNote = minutesFile
    ? `\n      <p class="update-note">AI-generated summary. For the complete official record, see the <a href="../minutes/${minutesFile}">full meeting minutes</a>.</p>`
    : '\n      <p class="update-note">AI-generated summary from Google Meet Gemini notes.</p>';

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml_(title)} — North Aldercroft Association</title>
  <link rel="stylesheet" href="../css/style.css">
</head>
<body>

<header class="site-header">
  <div class="header-inner">
    <a class="logo" href="../index.html">
      <span class="logo-primary">North Aldercroft</span>
      <span class="logo-secondary">Association</span>
    </a>
    <nav aria-label="Main navigation">
      <ul>
        <li><a href="../index.html">Home</a></li>
        <li><a href="../news.html">News &amp; Updates</a></li>
        <li><a href="../residents.html">For Residents</a></li>
        <li><a href="../buyers.html">For Buyers</a></li>
        <li><a href="../documents.html">Documents</a></li>
        <li><a href="../contact.html">Contact</a></li>
      </ul>
    </nav>
    <button class="nav-toggle" aria-label="Open navigation" aria-expanded="false"><span></span><span></span><span></span></button>
  </div>
</header>

<main>
  <div class="update-doc">

    <a class="back-link" href="../residents.html">&larr; Back to Resident Resources</a>

    <div class="update-doc-header">
      <div class="update-meta-row">
        <span class="update-tag update-tag-summary">Meeting Summary</span>
        <span class="update-date">${escapeHtml_(dateLabel)}</span>
      </div>
      <h1>${escapeHtml_(title)}</h1>${minutesNote}
    </div>

${htmlBody}

  </div>
</main>

<footer>
  &copy; <script>document.write(new Date().getFullYear())</script> North Aldercroft Association
</footer>

<script src="../js/nav.js"></script>
</body>
</html>
`;
}

// ── GitHub pull request ───────────────────────────────────────────────────────

function createPullRequest_(html, title, dateLabel, slug) {
  const token    = getProperty_('GITHUB_TOKEN');
  const filePath = `updates/${slug}.html`;
  const branch   = `meeting-summary-${slug}`;

  const baseSha = getRefSha_(token, GITHUB_BRANCH);

  githubPost_(token, `/repos/${GITHUB_OWNER}/${GITHUB_REPO}/git/refs`, {
    ref: `refs/heads/${branch}`,
    sha: baseSha,
  });

  githubPut_(token, `/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${filePath}`, {
    message: `Add meeting summary: ${title}`,
    content: Utilities.base64Encode(html),
    branch:  branch,
  });

  const { content: currentJs, sha: indexSha } = getFileContent_(token, 'updates/index.js', GITHUB_BRANCH);
  githubPut_(token, `/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/updates/index.js`, {
    message: `Update updates index for: ${title}`,
    content: Utilities.base64Encode(prependToIndex_(currentJs, slug, dateLabel, title)),
    sha:     indexSha,
    branch:  branch,
  });

  const pr = githubPost_(token, `/repos/${GITHUB_OWNER}/${GITHUB_REPO}/pulls`, {
    title: `Meeting summary: ${title}`,
    body:  `Auto-generated from Google Meet Gemini notes.\n\n**Date:** ${dateLabel}\n**File:** \`${filePath}\`\n\nReview the summary below, then merge to publish.`,
    head:  branch,
    base:  GITHUB_BRANCH,
  });

  return pr.html_url;
}

// ── updates/index.js manipulation ────────────────────────────────────────────

function prependToIndex_(currentJs, slug, dateLabel, title) {
  const dateIso  = slug.replace(/-summary$/, '');
  const newEntry = `  { date: "${dateIso}", title: "${escapeJs_(title)}", type: "summary", file: "${slug}.html" },`;
  return currentJs.replace(/(window\.UPDATES_INDEX\s*=\s*\[)/, `$1\n${newEntry}`);
}

// ── Notifications ─────────────────────────────────────────────────────────────

function notifyPrReady_(title, prUrl) {
  try {
    const email = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');
    if (!email) return;
    MailApp.sendEmail(
      email,
      `[NAA Website] Meeting summary ready for review: ${title}`,
      `A meeting summary has been drafted automatically and is waiting for review.\n\n` +
      `Title: ${title}\n` +
      `Review and merge here: ${prUrl}\n\n` +
      `Once merged, the summary will appear on the website.`
    );
  } catch (e) { Logger.log('Email notification failed: ' + e.message); }
}

function notifyError_(fileName, error) {
  try {
    const email = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');
    if (!email) return;
    MailApp.sendEmail(
      email,
      `[NAA Website] Error processing meeting notes: ${fileName}`,
      `The automatic summary pipeline encountered an error.\n\nFile: ${fileName}\nError: ${error.message}\n\nYou can process this doc manually via Extensions → NAA Website → Post summary to website.`
    );
  } catch (e) { Logger.log('Error notification failed: ' + e.message); }
}

// ── GitHub API helpers ────────────────────────────────────────────────────────

function getRefSha_(token, branch) {
  return githubFetch_(token, 'get', `/repos/${GITHUB_OWNER}/${GITHUB_REPO}/git/refs/heads/${branch}`, null).object.sha;
}

function getFileContent_(token, path, branch) {
  const res = githubFetch_(token, 'get', `/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${path}?ref=${branch}`, null);
  return {
    content: Utilities.newBlob(Utilities.base64Decode(res.content)).getDataAsString(),
    sha:     res.sha,
  };
}

function githubPost_(token, path, payload) { return githubFetch_(token, 'post', path, payload); }
function githubPut_(token, path, payload)  { return githubFetch_(token, 'put',  path, payload); }

function githubFetch_(token, method, path, payload) {
  const res = UrlFetchApp.fetch(`https://api.github.com${path}`, {
    method,
    headers: {
      Authorization:          `Bearer ${token}`,
      Accept:                 'application/vnd.github+json',
      'Content-Type':         'application/json',
      'X-GitHub-Api-Version': '2022-11-28',
    },
    payload:            payload ? JSON.stringify(payload) : undefined,
    muteHttpExceptions: true,
  });

  const status = res.getResponseCode();
  const body   = res.getContentText();
  if (status < 200 || status >= 300) {
    throw new Error(`GitHub ${method.toUpperCase()} ${path} → ${status}: ${body}`);
  }
  return JSON.parse(body);
}

// ── Utilities ─────────────────────────────────────────────────────────────────

function getProperty_(key) {
  const val = PropertiesService.getScriptProperties().getProperty(key);
  if (!val) throw new Error(`Script property "${key}" not set — see setup instructions at the top of this file.`);
  return val;
}

function escapeHtml_(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function escapeJs_(s) {
  return String(s).replace(/\\/g,'\\\\').replace(/"/g,'\\"');
}
