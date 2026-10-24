const CONFIG = {
  numHighlights: 4,
  subject: 'Daily Highlights',
  senderName: 'Daily Highlights',
  triggerHour: 8,
  skipHiddenSheets: true,
  excludedSheets: [],
  columns: { highlight: 0, author: 1, note: 2, page: 3, sourceUrl: 4 },
};

const COLUMN_COUNT = Math.max(...Object.values(CONFIG.columns)) + 1;
const EMAIL_PATTERN = /^[^\s@,<>]+@[^\s@,<>]+\.[^\s@,<>]+$/;
const ERROR_VALUE = /^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|ERROR!)$/;

const FONT_SANS = "'Inter', -apple-system, 'Segoe UI', Helvetica, Arial, sans-serif";
const FONT_SERIF = "Georgia, 'Times New Roman', serif";

const EMAIL_STYLES = `
  :root { color-scheme: light dark; supported-color-schemes: light dark; }
  a { color: inherit; }
  @media (prefers-color-scheme: dark) {
    .bg { background-color: #1e1e1e !important; }
    .card { background-color: #272727 !important; border-color: #363633 !important; }
    .tab { background-color: #363633 !important; color: #d4d0c8 !important; }
    .ink { color: #d4d0c8 !important; }
    .muted { color: #a8a49c !important; }
    .faint { color: #787470 !important; }
    .rule { border-color: #4a4a46 !important; }
    .dots { color: #4a4a46 !important; }
  }
`;

function sendRandomHighlight() {
  const settings = getSettings();
  const candidates = listCandidates(SpreadsheetApp.openById(settings.spreadsheetId));
  sendDigest(settings, candidates);
}

function sendDigest(settings, candidates) {
  validateConfig();

  if (candidates.length === 0) {
    throw new Error('No valid highlights found. Check that column A is filled and row 1 is a header.');
  }

  const highlights = pickRandom(candidates, CONFIG.numHighlights).map(loadHighlight);
  deliver(settings.recipients, renderEmail(highlights), renderPlainText(highlights));
  console.log(`Sent ${highlights.length} highlight(s) to ${settings.recipients.length} recipient(s).`);
}

function validateConfig() {
  if (!Number.isInteger(CONFIG.numHighlights) || CONFIG.numHighlights < 1) {
    throw new Error('CONFIG.numHighlights must be a positive integer.');
  }
  if (!Number.isInteger(CONFIG.triggerHour) || CONFIG.triggerHour < 0 || CONFIG.triggerHour > 23) {
    throw new Error('CONFIG.triggerHour must be an integer from 0 to 23.');
  }
}

function getSettings() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  const recipients = parseRecipients(props.getProperty('RECIPIENT_EMAIL'));

  if (!spreadsheetId) {
    throw new Error('Missing script property SPREADSHEET_ID (Project Settings > Script properties).');
  }
  if (recipients.length === 0) {
    throw new Error('Missing script property RECIPIENT_EMAIL. Use a comma-separated list for several recipients.');
  }

  const invalid = recipients.filter(email => !EMAIL_PATTERN.test(email));
  if (invalid.length > 0) {
    throw new Error(`Invalid address in RECIPIENT_EMAIL: ${invalid.join(', ')}`);
  }

  return { spreadsheetId, recipients };
}

function parseRecipients(raw) {
  const emails = (raw || '').split(',').map(email => email.trim()).filter(Boolean);
  return [...new Set(emails)];
}

function listCandidates(spreadsheet) {
  return spreadsheet.getSheets().filter(isIncluded).flatMap(listSheetCandidates);
}

function isIncluded(sheet) {
  if (CONFIG.skipHiddenSheets && sheet.isSheetHidden()) return false;
  return !CONFIG.excludedSheets.includes(sheet.getName());
}

function listSheetCandidates(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const column = CONFIG.columns.highlight + 1;
  const candidates = [];

  sheet.getRange(2, column, lastRow - 1, 1).getValues().forEach(([value], i) => {
    if (isUsableText(value)) {
      candidates.push({ sheet, row: i + 2 });
    }
  });

  return candidates;
}

function isUsableText(value) {
  const text = String(value).trim();
  return text !== '' && !ERROR_VALUE.test(text);
}

function pickRandom(items, count) {
  const pool = items.slice();
  const size = Math.min(count, pool.length);

  for (let i = 0; i < size; i++) {
    const j = i + Math.floor(Math.random() * (pool.length - i));
    [pool[i], pool[j]] = [pool[j], pool[i]];
  }

  return pool.slice(0, size);
}

function loadHighlight({ sheet, row }) {
  const columnCount = Math.min(COLUMN_COUNT, sheet.getMaxColumns());
  const range = sheet.getRange(row, 1, 1, columnCount);
  const [values] = range.getValues();
  const [richText] = range.getRichTextValues();
  const { highlight, author, note, page, sourceUrl } = CONFIG.columns;
  const sourceCell = richText[sourceUrl];

  return {
    title: sheet.getName(),
    highlight: richText[highlight],
    author: cellText(values[author]),
    note: hasText(richText[note]) ? richText[note] : null,
    page: normalizePage(values[page]),
    sourceUrl: (sourceCell && sourceCell.getLinkUrl()) || cellText(values[sourceUrl]),
  };
}

function hasText(richText) {
  return Boolean(richText && richText.getText().trim());
}

function cellText(value) {
  return value === undefined || value === null || value instanceof Date ? '' : String(value).trim();
}

function normalizePage(value) {
  const text = cellText(value);
  return text === '0' ? '' : text;
}

function renderEmail(highlights) {
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="color-scheme" content="light dark">
  <meta name="supported-color-schemes" content="light dark">
  <title>${escapeHtml(CONFIG.subject)}</title>
  <style>${EMAIL_STYLES}</style>
</head>
<body style="margin:0; padding:0;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="bg" bgcolor="#f0ede6" style="background-color:#f0ede6;">
  <tr>
    <td align="center" style="padding:32px 16px 40px; font-family:${FONT_SANS};">
      <table role="presentation" align="center" width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:560px; margin:0 auto;">
        <tr>
          <td style="padding:0 4px 28px;">
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
              <tr>
                <td class="ink" style="font-family:${FONT_SERIF}; font-size:22px; color:#1c1e1a;">${escapeHtml(CONFIG.subject.toLowerCase())}.</td>
                <td class="muted" align="right" style="font-family:${FONT_SANS}; font-size:13px; color:#5a5d54;">${escapeHtml(date)}</td>
              </tr>
            </table>
          </td>
        </tr>
        ${highlights.map(renderHighlight).join('')}
      </table>
    </td>
  </tr>
</table>
</body>
</html>`;
}

function renderHighlight(item) {
  const hasNote = Boolean(item.note);
  const footerParts = [];

  if (item.author) {
    footerParts.push(`<span class="faint" style="color:#80837a;">by <span class="muted" style="color:#5a5d54; font-weight:500;">${escapeHtml(item.author)}</span></span>`);
  }
  if (item.page) {
    footerParts.push(`<span class="faint" style="color:#80837a;">Page ${escapeHtml(item.page)}</span>`);
  }
  if (isSafeUrl(item.sourceUrl)) {
    footerParts.push(`<a href="${escapeHtml(item.sourceUrl)}" class="faint" style="color:#80837a; text-decoration:underline;">Source</a>`);
  }

  const footer = footerParts.length
    ? `<div class="dots" style="font-size:14px; line-height:1; letter-spacing:4px; color:#c9c4b6; margin:${hasNote ? 10 : 14}px 0 8px;">&middot;&middot;&middot;</div>
          <div style="font-size:12px; line-height:1.5;">${footerParts.join('<span style="display:inline-block; width:12px;"></span>')}</div>`
    : '';

  const note = hasNote
    ? `<div class="rule" style="width:32px; border-top:1px solid #c9c4b6; margin:16px 0 12px; font-size:0; line-height:0;">&nbsp;</div>
          <div class="muted" style="font-size:13px; line-height:1.55; color:#5a5d54;">${convertRichTextToHtml(item.note)}</div>`
    : '';

  return `<tr>
  <td style="padding:0 0 12px; font-family:${FONT_SANS};">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="14" style="font-size:0; line-height:0;">&nbsp;</td>
        <td class="tab" bgcolor="#e2ddd0" style="background-color:#e2ddd0; color:#3c3f37; font-family:${FONT_SANS}; font-size:12px; line-height:1.4; padding:4px 12px; border-radius:6px 6px 0 0;">${escapeHtml(item.title)}</td>
      </tr>
    </table>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td class="card" bgcolor="#ffffff" style="background-color:#ffffff; border:1px solid #e3dfd5; border-radius:8px; padding:16px 18px;">
          <div class="ink" style="font-size:15px; line-height:1.65; color:#1c1e1a;">${convertRichTextToHtml(item.highlight)}</div>
          ${note}${footer}
        </td>
      </tr>
    </table>
  </td>
</tr>`;
}

function renderPlainText(highlights) {
  return highlights.map(item => {
    const lines = [`[${item.title}]`, item.highlight.getText().trim()];

    if (item.note) {
      lines.push('', `Note: ${item.note.getText().trim()}`);
    }

    const meta = [
      item.author && `by ${item.author}`,
      item.page && `Page ${item.page}`,
      isSafeUrl(item.sourceUrl) && item.sourceUrl,
    ].filter(Boolean);

    if (meta.length > 0) {
      lines.push('', meta.join(' | '));
    }

    return lines.join('\n');
  }).join('\n\n---\n\n');
}

function convertRichTextToHtml(richText) {
  return richText.getRuns().map(run => {
    const text = escapeHtml(run.getText()).replace(/\n/g, '<br>');
    const textStyle = run.getTextStyle();
    let style = '';

    if (textStyle.isBold()) style += 'font-weight:600;';
    if (textStyle.isItalic()) style += 'font-style:italic;';
    if (textStyle.isUnderline()) style += 'text-decoration:underline;';

    let html = style ? `<span style="${style}">${text}</span>` : text;

    const linkUrl = run.getLinkUrl();
    if (isSafeUrl(linkUrl, true)) {
      html = `<a href="${escapeHtml(linkUrl)}">${html}</a>${renderLinkTarget(linkUrl, run.getText())}`;
    }

    return html;
  }).join('');
}

function renderLinkTarget(url, visibleText) {
  const target = linkTarget(url);
  if (!target || visibleText.toLowerCase().includes(target)) return '';
  return `<span class="faint" style="color:#80837a; font-size:0.85em;"> (${escapeHtml(target)})</span>`;
}

function linkTarget(url) {
  const web = /^https?:\/\/([^\/?#]+)/i.exec(url);
  if (web) {
    return web[1].replace(/^.*@/, '').replace(/:\d+$/, '').replace(/^www\./i, '').toLowerCase();
  }
  const mail = /^mailto:([^?]+)/i.exec(url);
  return mail ? mail[1].toLowerCase() : '';
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function isSafeUrl(url, allowMailto = false) {
  if (!url) return false;
  const pattern = allowMailto ? /^(https?:\/\/|mailto:)/i : /^https?:\/\//i;
  return pattern.test(url);
}

function deliver(recipients, htmlBody, body) {
  if (MailApp.getRemainingDailyQuota() < recipients.length) {
    throw new Error('Not enough daily email quota left to send to all recipients.');
  }

  const failures = [];

  recipients.forEach(to => {
    try {
      MailApp.sendEmail({ to, subject: CONFIG.subject, body, htmlBody, name: CONFIG.senderName });
    } catch (error) {
      failures.push(`${to}: ${error.message}`);
    }
  });

  if (failures.length > 0) {
    throw new Error(`Failed to send to ${failures.length} recipient(s): ${failures.join('; ')}`);
  }
}

function setup() {
  const settings = getSettings();
  const candidates = listCandidates(SpreadsheetApp.openById(settings.spreadsheetId));

  sendDigest(settings, candidates);
  console.log(`Setup OK: ${candidates.length} highlight(s) available. A test email was sent.`);
}

function createDailyTrigger() {
  validateConfig();
  deleteTriggers();
  ScriptApp.newTrigger('sendRandomHighlight')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.triggerHour)
    .create();
  console.log(`Daily trigger created for around ${CONFIG.triggerHour}:00.`);
}

function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'sendRandomHighlight');

  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  console.log(`Deleted ${triggers.length} trigger(s).`);
}