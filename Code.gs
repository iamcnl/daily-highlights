const CONFIG = {
  numHighlights: 4,
  subject: 'Daily Highlights',
  senderName: 'Daily Highlights',
  triggerHour: 8,
  columnCount: 5,
  columns: { highlight: 0, author: 1, note: 2, page: 3, sourceUrl: 4 },
};

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
  try {
    const settings = getSettings();
    const spreadsheet = SpreadsheetApp.openById(settings.spreadsheetId);
    const highlights = loadHighlights(spreadsheet);

    if (highlights.length === 0) {
      Logger.log('No valid highlights found. Check that column A is filled and row 1 is a header.');
      return;
    }

    const selected = selectHighlights(highlights, CONFIG.numHighlights);
    sendDigest(settings.recipients, renderEmail(selected));
    Logger.log(`Sent ${selected.length} highlight(s) to ${settings.recipients.length} recipient(s).`);
  } catch (error) {
    Logger.log(`Daily Highlights failed: ${error.message}`);
    throw error;
  }
}

function getSettings() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  const recipientProperty = props.getProperty('RECIPIENT_EMAIL');

  if (!spreadsheetId) {
    throw new Error('Missing script property SPREADSHEET_ID (Project Settings > Script properties).');
  }
  if (!recipientProperty) {
    throw new Error('Missing script property RECIPIENT_EMAIL. Use a comma-separated list for several recipients.');
  }

  const recipients = recipientProperty.split(',').map(email => email.trim()).filter(Boolean);

  return { spreadsheetId, recipients };
}

function loadHighlights(spreadsheet) {
  return spreadsheet.getSheets().flatMap(readSheet);
}

function readSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const columnCount = Math.min(CONFIG.columnCount, sheet.getMaxColumns());
  const range = sheet.getRange(2, 1, lastRow - 1, columnCount);
  const values = range.getValues();
  const richTextValues = range.getRichTextValues();
  const { highlight, author, note, page, sourceUrl } = CONFIG.columns;
  const title = sheet.getName();
  const highlights = [];

  values.forEach((row, i) => {
    const highlightText = richTextValues[i][highlight];
    if (!highlightText || !highlightText.getText().trim()) {
      return;
    }

    highlights.push({
      title,
      highlight: highlightText,
      author: cellText(row, author),
      note: richTextValues[i][note] || null,
      page: normalizePage(row[page]),
      sourceUrl: cellText(row, sourceUrl),
    });
  });

  return highlights;
}

function cellText(row, index) {
  const value = row[index];
  return value === undefined || value === null || value instanceof Date ? '' : String(value).trim();
}

function normalizePage(value) {
  if (value === undefined || value === null || value === '' || value === 0 || value instanceof Date) {
    return '';
  }
  return String(value).trim();
}

function selectHighlights(highlights, count) {
  return shuffleArray(highlights.slice()).slice(0, count);
}

function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

function renderEmail(highlights) {
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="color-scheme" content="light dark">
  <meta name="supported-color-schemes" content="light dark">
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
  const hasNote = Boolean(item.note && item.note.getText().trim());
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
      html = `<a href="${escapeHtml(linkUrl)}">${html}</a>`;
    }

    return html;
  }).join('');
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function isSafeUrl(url, allowMailto) {
  if (!url) return false;
  const pattern = allowMailto ? /^(https?:\/\/|mailto:)/i : /^https?:\/\//i;
  return pattern.test(url);
}

function sendDigest(recipients, htmlBody) {
  MailApp.sendEmail({
    to: recipients.join(','),
    subject: CONFIG.subject,
    htmlBody,
    name: CONFIG.senderName,
  });
}

function setup() {
  const settings = getSettings();
  const spreadsheet = SpreadsheetApp.openById(settings.spreadsheetId);
  const count = loadHighlights(spreadsheet).length;

  if (count === 0) {
    throw new Error('Spreadsheet opened, but no valid highlights were found.');
  }

  sendRandomHighlight();
  Logger.log(`Setup OK: ${count} highlight(s) available. A test email was sent.`);
}

function createDailyTrigger() {
  deleteTriggers();
  ScriptApp.newTrigger('sendRandomHighlight')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.triggerHour)
    .create();
  Logger.log(`Daily trigger created for around ${CONFIG.triggerHour}:00.`);
}

function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'sendRandomHighlight');

  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  Logger.log(`Deleted ${triggers.length} trigger(s).`);
}