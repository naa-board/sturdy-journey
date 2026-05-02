/**
 * North Aldercroft Association — Meet Notes → Website
 *
 * Reads the open Google Doc (Google Meet Gemini AI notes), sends it to Claude
 * to generate a plain-language summary, then opens a GitHub pull request that
 * adds the summary page to the site. A board member merges the PR to publish.
 *
 * ─── ONE-TIME SETUP ──────────────────────────────────────────────────────────
 *
 * 1. Open the Google Apps Script editor (script.google.com → New project).
 *    Paste this entire file in.
 *
 * 2. Add your secrets via Project Settings → Script Properties:
 *
 *      ANTHROPIC_API_KEY   Your Anthropic API key (console.anthropic.com)
 *      GITHUB_TOKEN        A GitHub personal access token (github.com → Settings
 *                          → Developer settings → Personal access tokens → Fine-
 *                          grained token). Needs Read+Write access to Contents
 *                          and Pull Requests on this repo.
 *      NOTIFY_EMAIL        (Optional) Email address to notify when a PR is ready.
 *
 * 3. Change GITHUB_BRANCH below to whatever branch you want PRs targeting
 *    (usually 'main' once the site is live).
 *
 * 4. Run installDocMenuTrigger() once (click the function name in the editor
 *    and hit Run). This makes the "NAA Website" menu appear in every Google Doc
 *    opened by the same account.
 *
 * ─── USAGE (after setup) ─────────────────────────────────────────────────────
 *
 * After each board meeting:
 *   1. Open the Gemini notes doc that Google Meet saved to Drive.
 *   2. Click  Extensions → NAA Website → Post summary to website…
 *   3. Confirm in the dialog.
 *   4. Click the PR link that appears and review the draft summary.
 *   5. Merge the PR — the site updates automatically.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 */

// ── Config ────────────────────────────────────────────────────────────────────

const GITHUB_OWNER  = 'scott4382';
const GITHUB_REPO   = 'sturdy-journey';
const GITHUB_BRANCH = 'main';   // ← change if your production branch differs

// The Claude model to use for summarisation
const CLAUDE_MODEL  = 'claude-sonnet-4-6';

// ── Menu setup ────────────────────────────────────────────────────────────────

function onOpen() {
  DocumentApp.getUi()
    .createMenu('NAA Website')
    .addItem('Post summary to website…', 'showConfirmDialog')
    .addToUi();
}

/** Run once to install the onOpen trigger across all Docs. */
function installDocMenuTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onOpen')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onOpen')
    .forDocument(DocumentApp.getActiveDocument())
    .onOpen()
    .create();

  DocumentApp.getUi().alert('Setup complete! Reload the doc to see the NAA Website menu.');
}

// ── Dialog ────────────────────────────────────────────────────────────────────

function showConfirmDialog() {
  const ui   = DocumentApp.getUi();
  const doc  = DocumentApp.getActiveDocument();
  const title = doc.getName();

  const result = ui.alert(
    'Post summary to website',
    `This will:\n` +
    `1. Send the contents of "${title}" to Claude for summarisation\n` +
    `2. Create a GitHub pull request with the generated summary page\n\n` +
    `A board member will need to review and merge the PR to publish it.\n\n` +
    `Continue?`,
    ui.ButtonSet.YES_NO
  );

  if (result === ui.Button.YES) {
    runPipeline();
  }
}

// ── Main pipeline ─────────────────────────────────────────────────────────────

function runPipeline() {
  const ui = DocumentApp.getUi();

  try {
    // 1. Read the document
    ui.alert('Step 1/3: Reading document…');
    const doc   = DocumentApp.getActiveDocument();
    const notes = extractDocumentText_(doc);
    const docTitle = doc.getName();

    if (notes.trim().length < 100) {
      ui.alert('Error', 'The document looks too short to summarise. Is this the right doc?', ui.ButtonSet.OK);
      return;
    }

    // 2. Generate summary via Claude
    ui.alert('Step 2/3: Generating summary with Claude… (may take 20–30 seconds)');
    const { html, title, dateLabel, slug } = generateSummary_(notes, docTitle);

    // 3. Create GitHub PR
    ui.alert('Step 3/3: Creating GitHub pull request…');
    const prUrl = createPullRequest_(html, title, dateLabel, slug);

    // 4. Notify
    const email = getProperty_('NOTIFY_EMAIL');
    if (email) {
      MailApp.sendEmail(email, `[NAA Website] New summary ready for review: ${title}`,
        `A board meeting summary has been drafted and is ready for review.\n\n` +
        `Title: ${title}\n` +
        `Review PR: ${prUrl}\n\n` +
        `Merge the PR to publish it on the website.`
      );
    }

    ui.alert(
      'Done!',
      `Pull request created:\n${prUrl}\n\n` +
      `Review the summary there. Merge it to publish on the site.`,
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('Error', `Something went wrong:\n\n${e.message}`, ui.ButtonSet.OK);
    Logger.log(e.stack);
  }
}

// ── Document text extraction ──────────────────────────────────────────────────

function extractDocumentText_(doc) {
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();
  return paragraphs
    .map(p => p.getText())
    .filter(t => t.trim())
    .join('\n');
}

// ── Claude summarisation ──────────────────────────────────────────────────────

function generateSummary_(notes, docTitle) {
  const apiKey = getProperty_('ANTHROPIC_API_KEY');
  const today  = new Date().toISOString().slice(0, 10); // YYYY-MM-DD

  const prompt = `You are helping maintain the website for the North Aldercroft Association, a private road HOA in the Santa Cruz Mountains. Below are the AI-generated notes from a board meeting (produced by Google Meet's Gemini feature). Your job is to turn them into a clear, friendly summary that residents can quickly read.

MEETING NOTES:
${notes}

Write a meeting summary following these rules exactly:
1. Tone: Plain, friendly language. Write for neighbors, not lawyers. No jargon.
2. Start with a "The Short Version" section (2–3 sentences max) covering the most important decisions or news.
3. Add 3–6 more sections, each covering one topic. Use a short descriptive heading (not generic like "Discussion" — be specific, e.g. "Spring Road Repair", "Dues Update").
4. End with a "What's Next" section covering the next meeting date (if mentioned) or key upcoming actions.
5. Keep each section to 2–4 sentences. Do not pad.
6. If road work, dues, costs, or deadlines were discussed, make sure they appear clearly.

Output ONLY the following, with no extra explanation:

TITLE: [short meeting title, e.g. "March 2025 Board Meeting Summary"]
DATE_ISO: [meeting date as YYYY-MM-DD, infer from the notes or use ${today}]
DATE_LABEL: [human-readable date, e.g. "March 15, 2025"]
MINUTES_FILE: [leave blank if no formal minutes file yet]
---HTML_BODY---
[the section headings and paragraphs as HTML — only <h2> and <p> tags, no wrapper divs or other elements]
---END---`;

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json',
    },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: 2048,
      messages: [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  const body   = JSON.parse(response.getContentText());

  if (status !== 200) {
    throw new Error(`Claude API error ${status}: ${JSON.stringify(body)}`);
  }

  const text = body.content[0].text;

  // Parse the structured output
  const title      = (text.match(/^TITLE:\s*(.+)$/m)     || [])[1]?.trim() || 'Board Meeting Summary';
  const dateIso    = (text.match(/^DATE_ISO:\s*(.+)$/m)   || [])[1]?.trim() || today;
  const dateLabel  = (text.match(/^DATE_LABEL:\s*(.+)$/m) || [])[1]?.trim() || today;
  const minutesFile = (text.match(/^MINUTES_FILE:\s*(.+)$/m) || [])[1]?.trim() || '';
  const htmlBody   = (text.match(/---HTML_BODY---([\s\S]*?)---END---/) || [])[1]?.trim() || '';

  // Generate a URL-safe slug from the date + "summary"
  const slug = `${dateIso}-summary`;

  // Build the full HTML page
  const fullHtml = buildHtmlPage_(title, dateLabel, slug, htmlBody, minutesFile);

  return { html: fullHtml, title, dateLabel, slug };
}

// ── HTML page builder ─────────────────────────────────────────────────────────

function buildHtmlPage_(title, dateLabel, slug, htmlBody, minutesFile) {
  const minutesNote = minutesFile
    ? `\n      <p class="update-note">This is an AI-generated summary of the board meeting. For the complete official record, see the <a href="../minutes/${minutesFile}">full meeting minutes</a>.</p>`
    : '';

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

  // 1. Get the SHA of the base branch tip
  const baseSha = getRefSha_(token, GITHUB_BRANCH);

  // 2. Create the new branch
  githubPost_(token, `/repos/${GITHUB_OWNER}/${GITHUB_REPO}/git/refs`, {
    ref: `refs/heads/${branch}`,
    sha: baseSha,
  });

  // 3. Create the new HTML file
  githubPut_(token, `/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${filePath}`, {
    message: `Add meeting summary: ${title}`,
    content: Utilities.base64Encode(html),
    branch: branch,
  });

  // 4. Read and update updates/index.js
  const indexPath     = 'updates/index.js';
  const { content: currentJs, sha: indexSha } = getFileContent_(token, indexPath, GITHUB_BRANCH);
  const updatedJs     = prependToIndex_(currentJs, slug, dateLabel, title);

  githubPut_(token, `/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${indexPath}`, {
    message: `Update updates index for: ${title}`,
    content: Utilities.base64Encode(updatedJs),
    sha:     indexSha,
    branch:  branch,
  });

  // 5. Open the pull request
  const pr = githubPost_(token, `/repos/${GITHUB_OWNER}/${GITHUB_REPO}/pulls`, {
    title:  `Meeting summary: ${title}`,
    body:   `Auto-generated from Google Meet Gemini notes via Apps Script.\n\n` +
            `**Date:** ${dateLabel}\n` +
            `**File:** \`${filePath}\`\n\n` +
            `Review the summary below. Merge to publish on the site.`,
    head:   branch,
    base:   GITHUB_BRANCH,
  });

  return pr.html_url;
}

// ── updates/index.js manipulation ────────────────────────────────────────────

function prependToIndex_(currentJs, slug, dateLabel, title) {
  // Parse the date from the slug (YYYY-MM-DD-summary → YYYY-MM-DD)
  const dateIso = slug.replace(/-summary$/, '');
  const file    = `${slug}.html`;

  const newEntry =
    `  { date: "${dateIso}", title: "${escapeJs_(title)}", type: "summary", file: "${file}" },`;

  // Insert after "window.UPDATES_INDEX = ["
  return currentJs.replace(
    /(window\.UPDATES_INDEX\s*=\s*\[)/,
    `$1\n${newEntry}`
  );
}

// ── GitHub API helpers ────────────────────────────────────────────────────────

function getRefSha_(token, branch) {
  const res = githubFetch_(token, 'get',
    `/repos/${GITHUB_OWNER}/${GITHUB_REPO}/git/refs/heads/${branch}`, null);
  return res.object.sha;
}

function getFileContent_(token, path, branch) {
  const res = githubFetch_(token, 'get',
    `/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${path}?ref=${branch}`, null);
  const content = Utilities.newBlob(Utilities.base64Decode(res.content)).getDataAsString();
  return { content, sha: res.sha };
}

function githubPost_(token, path, payload) {
  return githubFetch_(token, 'post', path, payload);
}

function githubPut_(token, path, payload) {
  return githubFetch_(token, 'put', path, payload);
}

function githubFetch_(token, method, path, payload) {
  const url = `https://api.github.com${path}`;
  const options = {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept:        'application/vnd.github+json',
      'Content-Type':'application/json',
      'X-GitHub-Api-Version': '2022-11-28',
    },
    muteHttpExceptions: true,
  };
  if (payload) options.payload = JSON.stringify(payload);

  const response = UrlFetchApp.fetch(url, options);
  const status   = response.getResponseCode();
  const body     = response.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error(`GitHub API ${method.toUpperCase()} ${path} → ${status}: ${body}`);
  }

  return JSON.parse(body);
}

// ── Utilities ─────────────────────────────────────────────────────────────────

function getProperty_(key) {
  const val = PropertiesService.getScriptProperties().getProperty(key);
  if (!val) throw new Error(`Script property "${key}" is not set. See setup instructions at the top of this file.`);
  return val;
}

function escapeHtml_(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escapeJs_(str) {
  return String(str).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}
