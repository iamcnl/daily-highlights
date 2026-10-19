function sendRandomHighlight() {
  const scriptProperties = PropertiesService.getScriptProperties();

  const spreadsheetId = scriptProperties.getProperty('SPREADSHEET_ID') || 'YOUR_SPREADSHEET_ID';
  const recipientEmail = scriptProperties.getProperty('RECIPIENT_EMAIL') || 'your_email@example.com';
  var numHighlights = 4;
  
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheets = spreadsheet.getSheets();
  
  var allRows = [];
  
sheets.forEach(function(sheet) {
    var lastRow = sheet.getLastRow();
    
    // Skip empty sheets immediately to save processing time
    if (lastRow === 0) {
      return;
    }
    
    // Fetch only the rows and columns that actually have data (Columns A to D)
    var range = sheet.getRange(1, 1, lastRow, 4);
    var data = range.getValues();
    var richTextValues = range.getRichTextValues();
    
    var cellC1Value = sheet.getRange('C1').getValue();
    var sheetName = sheet.getName();
    var formattedSheetName = '<strong>' + sheetName + '</strong>';
    
    if (cellC1Value) {
      formattedSheetName += ' <span style="font-size: 11px; color:#525453;"> ' + ' by ' + cellC1Value + ' </span>';
    }

    let rows = [];

    if (data.length > 1) {
      rows = data.slice(1).filter((row, i) => row[0] || row[1]).map((row, i) => {
        return [formattedSheetName, richTextValues[i + 1][0], richTextValues[i + 1][1], row[3]]; // Column D (plain text)
      });
    } else if (data.length === 1) {
      rows = data.filter((row, i) => row[0] || row[1]).map((row, i) => {
        return [formattedSheetName, richTextValues[i][0], richTextValues[i][1], row[3]]; // Column D (plain text)
      });
    }
    
    allRows = allRows.concat(rows);
  });
  
  if (allRows.length === 0) {
    Logger.log('No valid rows available.');
    return;
  }
  
  allRows = shuffleArray(allRows);
  var selectedRows = allRows.slice(0, Math.min(numHighlights, allRows.length));
  
  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'MMMM dd, yyyy');
  
  var body = '<div style="text-align:center; max-width:600px; margin:auto;">';
  body += '<h2 style="margin-bottom: 0;">Daily Highlights</h2>';
  body += '<p style="margin-top: 4px; font-size: 16px; color: #555;">' + formattedDate + '</p><br>';
  
  selectedRows.forEach(function(row) {
    var sheetName = row[0];
    var highlight = row[1];
    var note = row[2];
    var page = row[3]; // Column D (plain text)

    //sheetName
    body += '<div style="border: 1px solid #ccc; padding: 10px; margin: 10px 0; text-align:left;">';
    body += '<p style="color: #1f1e1b; font-size: 17px; font-family: Verdana; padding: 0px 5px;">' + sheetName + '</p>';

    //highlight
    body += '<hr style="border:none; border-top: 1px solid #ccc; margin: 5px 0;">';
    body += '<blockquote style="font-size: 13px; text-align: left;">' + convertRichTextToHtml(highlight) + '</blockquote>';

    //page
    if (page !== undefined && page !== null && page !== '') { // Ensure page is not undefined, null, or empty
      body += '<div style="display: inline-block;">';
      body += '<p style="font-size: 11.5px; color: #a6a6a4; margin: 2px 0 2px 10px">(Page ' + page + ')</p>';
      body += '</div>';
    }

    //note
    if (note.getText()) {
      body += '<div style="background-color: #f9f9f9; padding: 2px 15px 2px; border-left: 4px solid #ccc;">';
      body += '<p><strong>Note:</strong> ' + convertRichTextToHtml(note) + '</p>';
      body += '</div>';
    }
    
    body += '</div>';
    body += '<br>'; // Add an empty row between highlights
  });
  body += '</div>';
  
  var subject = 'Daily Highlights';
  Logger.log('Email body: ' + body);

  MailApp.sendEmail({
    to: recipientEmail,
    subject: subject,
    htmlBody: body,
    name: "Daily Highlights"
  });
}

// Helper function to shuffle an array
function shuffleArray(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
}

// Helper function to convert RichTextValue to HTML
function convertRichTextToHtml(richText) {
  var text = richText.getText();
  var html = '';
  
  for (var i = 0; i < richText.getRuns().length; i++) {
    var run = richText.getRuns()[i];
    var runText = run.getText().replace(/\n/g, '<br>');
    var style = '';
    var textStyle = run.getTextStyle();
    
    if (textStyle.isBold()) {
      style += 'font-weight:bold;';
    }
    if (textStyle.isItalic()) {
      style += 'font-style:italic;';
    }
    if (textStyle.isUnderline()) {
      style += 'text-decoration:underline;';
    }
    
    if (style) {
      html += '<span style="' + style + '">' + runText + '</span>';
    } else {
      html += runText;
    }
  }
  
  return html;
}
