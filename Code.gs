const CONFIG = {
  numHighlights: 4,
  subject: 'Daily Highlights',
  senderName: 'Daily Highlights',
  triggerHour: 8,
  columnCount: 5,
  columns: { highlight: 0, author: 1, note: 2, page: 3, sourceUrl: 4 },
};

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
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy');

  return `<div style="text-align:center; max-width:600px; margin:auto;">
  <h2 style="margin-bottom:0;">${escapeHtml(CONFIG.subject)}</h2>
  <p style="margin-top:4px; font-size:16px; color:#555;">${date}</p><br>
  ${highlights.map(renderHighlight).join('')}
</div>`;
}

function renderHighlight(item) {
  const author = item.author
    ? ` <span style="font-size:11px; color:#525453;">by ${escapeHtml(item.author)}</span>`
    : '';

  const page = item.page
    ? `<p style="font-size:11.5px; color:#a6a6a4; margin:2px 0 2px 10px;">(Page ${escapeHtml(item.page)})</p>`
    : '';

  const source = isSafeUrl(item.sourceUrl)
    ? `<p style="font-size:11.5px; margin:2px 0 2px 10px;"><a href="${escapeHtml(item.sourceUrl)}" style="color:#a6a6a4;">Source</a></p>`
    : '';

  const note = item.note && item.note.getText().trim()
    ? `<div style="background-color:#f9f9f9; padding:2px 15px; border-left:4px solid #ccc;">
        <p><strong>Note:</strong> ${convertRichTextToHtml(item.note)}</p>
      </div>`
    : '';

  return `<div style="border:1px solid #ccc; padding:10px; margin:10px 0; text-align:left;">
    <p style="color:#1f1e1b; font-size:17px; font-family:Verdana; padding:0 5px;"><strong>${escapeHtml(item.title)}</strong>${author}</p>
    <hr style="border:none; border-top:1px solid #ccc; margin:5px 0;">
    <blockquote style="font-size:13px; text-align:left;">${convertRichTextToHtml(item.highlight)}</blockquote>
    ${page}${source}${note}
  </div><br>`;
}

function convertRichTextToHtml(richText) {
  return richText.getRuns().map(run => {
    const text = escapeHtml(run.getText()).replace(/\n/g, '<br>');
    const textStyle = run.getTextStyle();
    let style = '';

    if (textStyle.isBold()) style += 'font-weight:bold;';
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