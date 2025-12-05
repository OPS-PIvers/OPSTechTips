/**
 * Professional Newsletter Generator for Google Apps Script
 * Refactored for Brand Consistency, Maintainability, and Gmail Reliability
 */

// --- BRAND CONFIGURATION ---
const BRAND = {
  colors: {
    primaryBlue: '#2d3f89',
    primaryRed: '#ad2122',
    primaryGray: '#666666',
    secondaryBlue: '#4356a0',
    secondaryRed: '#c13435',
    secondaryGray: '#999999',
    lightBg: '#f3f3f3',
    contentBg: '#ffffff',
    accentBg: '#eaecf5'
  },
  fonts: {
    headings: 'Lexend, Arial, sans-serif',
    body: 'Roboto, Arial, sans-serif'
  },
  gradients: {
    blue: 'linear-gradient(135deg, #2d3f89 0%, #4356a0 100%)',
    red: 'linear-gradient(135deg, #ad2122 0%, #c13435 100%)',
    light: 'linear-gradient(135deg, #eaecf5 0%, #f3f3f3 100%)'
  }
};

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Newsletter Tools')
    .addItem('Send Newsletter', 'showColumnPicker')
    .addItem('Create Draft Newsletter', 'showDraftPicker')
    .addItem('Preview Newsletter', 'showPreviewPicker')
    .addSeparator()
    .addItem('Generate HTML Only', 'showGeneratePicker')
    .addToUi();
}

// --- UI FUNCTIONS ---

function showColumnPicker() { showDialog('send', 'Select Newsletter to Send'); }
function showDraftPicker() { showDialog('draft', 'Select Newsletter to Create Draft'); }
function showPreviewPicker() { showDialog('preview', 'Select Newsletter to Preview'); }
function showGeneratePicker() { showDialog('generate', 'Select Newsletter to Generate'); }

function showDialog(action, title) {
  const html = createColumnPickerDialog(action);
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(350)
    .setTitle(title);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}

function createColumnPickerDialog(action) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const columns = ['B', 'C', 'D', 'E', 'F'];
  const options = [];
  
  // Batch read rows 1 (Date) and 2 (Title) for columns B-F
  const range = sheet.getRange('B1:F2');
  const values = range.getValues(); // Returns 2D array: [[B1, C1, D1, E1, F1], [B2, C2, D2, E2, F2]] (row 0 = dates, row 1 = titles)

  columns.forEach((col, index) => {
    const dateCell = values[0][index];
    const dateStr = dateCell ? Utilities.formatDate(new Date(dateCell), Session.getScriptTimeZone(), 'MM/dd/yyyy') : 'No Date';

    const titleCell = values[1][index];
    const titleStr = titleCell ? titleCell.toString().substring(0, 30) + (titleCell.toString().length > 30 ? '...' : '') : 'No Title';
    
    options.push({
      column: col,
      date: dateStr,
      title: titleStr
    });
  });

  try {
    const template = HtmlService.createTemplateFromFile('Picker');
    template.action = action;
    template.options = options;
    return template.evaluate().getContent();
  } catch (e) {
    return '<h3>Error: Picker.html file missing. Please create it.</h3>';
  }
}

// --- CORE LOGIC ---

function generateNewsletterHTMLFromColumn(column) {
  try {
    SpreadsheetApp.getActive().toast('Generating HTML...', 'Status', 3);
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = getNewsletterDataFromColumn(sheet, column);
    const html = createNewsletterHTML(data);
    return html;
  } catch (error) {
    console.error(`Error generating HTML: ${error.message}`);
    throw error;
  }
}

function sendNewsletterFromColumn(column) { return processEmailAction(column, 'send'); }
function createDraftNewsletterFromColumn(column) { return processEmailAction(column, 'draft'); }

function processEmailAction(column, mode) {
  try {
    SpreadsheetApp.getActive().toast(`Processing ${mode}...`, 'Status', -1);
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = getNewsletterDataFromColumn(sheet, column);
    
    if (!data.to) throw new Error(`No "To" recipients in column ${column}`);
    if (!data.title) throw new Error(`No Title in column ${column}`);
    
    const html = createNewsletterHTML(data);
    const subject = stripHtmlTags(data.title) + (data.date ? ' - ' + Utilities.formatDate(new Date(data.date), Session.getScriptTimeZone(), 'MM/dd/yyyy') : '');
    
    const options = {
      htmlBody: html,
      cc: data.cc || '',
      bcc: data.bcc || '',
      attachments: []
    };

    if (mode === 'send') {
      GmailApp.sendEmail(data.to, subject, '', options);
      SpreadsheetApp.getActive().toast('Email sent successfully!', 'Success', 5);
    } else {
      GmailApp.createDraft(data.to, subject, '', options);
      SpreadsheetApp.getActive().toast('Draft created successfully!', 'Success', 5);
    }
    return true;
  } catch (error) {
    SpreadsheetApp.getActive().toast('Error: ' + error.message, 'Error', 10);
    console.error(error);
    throw error;
  }
}

// --- DATA EXTRACTION ---

function getNewsletterDataFromColumn(sheet, column) {
  // Batch read rows 1-23
  // 1-based indices in sheet:
  // 1: Date
  // 2: Title
  // 3: Subtitle
  // 4-8: Topic 1
  // 9-13: Topic 2
  // 14-18: Topic 3
  // 19: Final Button URL
  // 20: To
  // 21: Cc
  // 22: Bcc
  // 23: Layout Style

  const range = sheet.getRange(`${column}1:${column}23`);
  const values = range.getValues(); // 2D array [row][col]
  const richTextValues = range.getRichTextValues(); // 2D array [row][col]

  const getVal = (rowIdx) => {
    const v = values[rowIdx - 1][0];
    return v === undefined ? '' : v;
  };

  const getRichVal = (rowIdx) => {
    const v = richTextValues[rowIdx - 1][0];
    return v;
  };

  const getFmt = (rowIdx, mode) => {
    const rtv = getRichVal(rowIdx);
    const v = getVal(rowIdx);
    return getFormattedCellValueFromData(rtv, v, mode);
  };

  const extractTopicBatch = (startRow) => {
    return {
      title: getFmt(startRow, 'header'),
      url: getVal(startRow + 1),
      description: getFmt(startRow + 2, 'body'),
      buttonText: getVal(startRow + 3),
      buttonUrl: getVal(startRow + 4)
    };
  };
  
  const data = {
    date: getVal(1),
    title: getFmt(2, 'header'),
    subtitle: getFmt(3, 'header'),
    topic1: extractTopicBatch(4),
    topic2: extractTopicBatch(9),
    topic3: extractTopicBatch(14),
    finalButtonUrl: getVal(19),
    to: getVal(20),
    cc: getVal(21),
    bcc: getVal(22),
    layoutStyle: getVal(23)
  };
  
  sanitizeNewsletterData(data);
  return data;
}

function sanitizeNewsletterData(data) {
  if (data.title) data.title = sanitizeHtml(data.title);
  if (data.subtitle) data.subtitle = sanitizeHtml(data.subtitle);
  ['topic1', 'topic2', 'topic3'].forEach(t => {
    if (data[t].title) data[t].title = sanitizeHtml(data[t].title);
    if (data[t].description) data[t].description = sanitizeHtml(data[t].description);
  });
}

// --- HTML GENERATION (Mobile Optimized) ---

function createNewsletterHTML(data) {
  const topics = [data.topic1, data.topic2, data.topic3]
    .filter(t => t.title && (t.url || t.description))
    .map(t => ({
      ...t,
      url: convertDriveImageUrl(t.url),
      description: t.description || ''
    }));

  const layoutStyle = (data.layoutStyle || 'offset').trim().toLowerCase();
  
  let topicHTML = '';
  if (layoutStyle === 'stacked') topicHTML = generateStackedLayout(topics);
  else if (layoutStyle === 'hero') topicHTML = generateHeroLayout(topics);
  else topicHTML = generateOffsetLayout(topics);

  const logos = getLogosFromConfig();
  const preheaderText = data.subtitle ? stripHtmlTags(data.subtitle).substring(0, 100) : 'Orono Technology Newsletter';

  return `
<!DOCTYPE html>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="format-detection" content="telephone=no, date=no, address=no, email=no">
    <meta name="color-scheme" content="light dark">
    <meta name="supported-color-schemes" content="light dark">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>${stripHtmlTags(data.title) || 'Newsletter'}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Lexend:wght@400;600;700&family=Roboto:wght@400;500;700&display=swap');
        
        /* RESET */
        body { margin: 0; padding: 0; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; background-color: ${BRAND.colors.lightBg}; min-width: 320px; }
        table, td { border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; }
        img { border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic; max-width: 100%; }
        
        /* TYPOGRAPHY */
        body, td { font-family: ${BRAND.fonts.body}; font-size: 11pt; color: ${BRAND.colors.primaryGray}; }
        h1, h2, h3 { font-family: ${BRAND.fonts.headings}; margin: 0; word-break: break-word; }
        p { word-break: break-word; margin: 0 0 10px 0; }
        a { text-decoration: none; color: inherit; }
        
        /* DARK MODE */
        @media (prefers-color-scheme: dark) {
            body, .body { background-color: #1a1a1a !important; color: #e0e0e0 !important; }
            .container { background-color: #2d2d2d !important; border-radius: 8px !important; overflow: hidden !important; }
            .header-title { color: #ffffff !important; }
            h1, h2, h3, h4, h5, h6 { color: #ffffff !important; }
            td[style*="background-color: #ffffff"] { background-color: #2d2d2d !important; color: #ffffff !important; }
            .description-box { background-color: #333333 !important; border-left-color: ${BRAND.colors.secondaryBlue} !important; }
            .description-text { color: #cccccc !important; }
        }

        /* RESPONSIVE MEDIA QUERIES */
        @media screen and (max-width: 600px) {
            .main-container { width: 100% !important; max-width: 100% !important; }
            .content-padding { padding: 20px !important; }
            .responsive-cell { display: block !important; width: 100% !important; padding: 0 0 20px 0 !important; }
            .responsive-image img { width: 100% !important; height: auto !important; }
            .header-title { font-size: 24pt !important; }
        }
        
        /* GMAIL BLUE LINK FIX */
        a[x-apple-data-detectors] { color: inherit !important; text-decoration: none !important; font-size: inherit !important; font-family: inherit !important; font-weight: inherit !important; line-height: inherit !important; }
    </style>
</head>
<body id="body" class="body" style="margin: 0; padding: 0; background-color: ${BRAND.colors.lightBg}; font-family: ${BRAND.fonts.body}; color: ${BRAND.colors.primaryGray};">
    
    <!-- PREHEADER HACK -->
    <div style="display: none; max-height: 0px; overflow: hidden; mso-hide: all;">
      ${preheaderText}
    </div>
    <div style="display: none; max-height: 0px; overflow: hidden;">
      &nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;
    </div>

    <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" style="background-color: ${BRAND.colors.lightBg}; padding: 20px 0;">
        <tr>
            <td align="center">
                <!-- Outlook Wrapper Force 600px -->
                <!--[if (gte mso 9)|(IE)]>
                <table width="600" align="center" cellpadding="0" cellspacing="0" border="0"><tr><td>
                <![endif]-->
                
                <!-- Main Container: 600px on Desktop (fixed attribute), 100% on Mobile (via Class) -->
                <table class="main-container" width="600" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="width: 600px; max-width: 100%;">
                    <tr>
                        <td>
                            <table width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" class="container" style="background-color: ${BRAND.colors.contentBg}; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 12px rgba(45, 63, 137, 0.1);">
                                
                                <!-- WHITE LOGO HEADER -->
                                <tr>
                                    <td style="background-color: #ffffff; padding: 25px 30px; text-align: center; border-bottom: 1px solid ${BRAND.colors.accentBg};">
                                        ${logos.main ? `<img src="${logos.main}" alt="Orono Technology" width="200" style="width: 100%; max-width: 200px; height: auto; display: inline-block; border: 0;">` : ''}
                                    </td>
                                </tr>

                                <!-- HERO TITLE SECTION -->
                                <tr>
                                    <td class="header-padding" style="background-color: ${BRAND.colors.primaryBlue}; background: ${BRAND.gradients.blue}; padding: 40px 30px; text-align: center;">
                                        ${data.date ? `<div style="color: ${BRAND.colors.accentBg}; font-size: 11pt; letter-spacing: 1px; margin-bottom: 10px; text-transform: uppercase;">${Utilities.formatDate(new Date(data.date), Session.getScriptTimeZone(), 'MMMM yyyy')}</div>` : ''}
                                        ${data.title ? `<h1 class="header-title" style="font-family: ${BRAND.fonts.headings}; color: ${BRAND.colors.contentBg}; font-size: 28pt; margin: 0 0 10px 0; line-height: 1.2;">${data.title}</h1>` : ''}
                                        ${data.subtitle ? `<p style="color: ${BRAND.colors.accentBg}; font-size: 13pt; margin: 0; line-height: 1.4;">${data.subtitle}</p>` : ''}
                                    </td>
                                </tr>

                                <!-- Content -->
                                <tr>
                                    <td class="content-padding" style="padding: 40px 30px;">
                                        ${topicHTML}
                                        ${data.finalButtonUrl ? generateCallToAction(data.finalButtonUrl) : ''}
                                    </td>
                                </tr>
                                <!-- Footer -->
                                <tr>
                                    <td style="background-color: #1d2a5d; padding: 25px 30px; text-align: right;">
                                        ${logos.secondary ? `<img src="${logos.secondary}" alt="Icon" width="60" style="width: 60px; max-width: 60px; height: auto; margin-bottom: 15px; display: inline-block; border: 0;">` : ''}
                                        <p style="color: ${BRAND.colors.accentBg}; font-size: 10pt; margin: 0; line-height: 1.5;">
                                            ${new Date().getFullYear()} Orono Technology Digital Learning Hub<br>
                                            <span style="color: ${BRAND.colors.secondaryBlue};">Empowering Digital Learning and Innovation</span>
                                        </p>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
                <!--[if (gte mso 9)|(IE)]>
                </td></tr></table>
                <![endif]-->
            </td>
        </tr>
    </table>
</body>
</html>`;
}

// --- LAYOUT GENERATORS ---

function createButtonHTML(text, url, style = 'blue', padding = '10px 20px', fontSize = '11pt') {
  const bgColor = style === 'red' ? BRAND.colors.primaryRed : BRAND.colors.primaryBlue;
  const gradient = style === 'red' ? BRAND.gradients.red : BRAND.gradients.blue;
  
  // Parse padding string (e.g., '10px 20px') for border-based technique
  const parts = padding.trim().split(' ');
  const py = parts[0];
  const px = parts.length > 1 ? parts[1] : parts[0];

  return `<a href="${url}" style="background-color: ${bgColor}; background: ${gradient}; color: #ffffff; text-decoration: none; border-top: ${py} solid ${bgColor}; border-bottom: ${py} solid ${bgColor}; border-left: ${px} solid ${bgColor}; border-right: ${px} solid ${bgColor}; border-radius: 6px; font-size: ${fontSize}; font-weight: 600; font-family: ${BRAND.fonts.headings}; display: inline-block; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.25);">${text}</a>`;
}

function generateCallToAction(url) {
  return `
    <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" style="margin-top: 40px;">
        <tr>
            <td align="center" style="background-color: ${BRAND.colors.accentBg}; background: ${BRAND.gradients.light}; padding: 30px; border-radius: 8px;">
                <h3 style="font-family: ${BRAND.fonts.headings}; color: ${BRAND.colors.primaryBlue}; font-size: 18pt; margin: 0 0 20px 0;">Ready to Learn More?</h3>
                ${createButtonHTML('Visit the Orono Technology Digital Learning Hub to learn more', url, 'red', '14px 32px', '14pt')}
            </td>
        </tr>
    </table>`;
}

function generateOffsetLayout(topics) {
  return topics.map((topic, index) => {
    const divider = index > 0 ? getDividerHTML() : '';
    const isEven = index % 2 === 0;
    
    const imageCell = topic.url ? `
      <td width="33%" class="responsive-cell" style="padding: ${isEven ? '0 20px 0 0' : '0 0 0 20px'}; vertical-align: top;">
          <div class="responsive-image" style="border-radius: 8px; overflow: hidden; border: 1px solid ${BRAND.colors.accentBg};">
              <img src="${topic.url}" alt="${topic.title}" style="width: 100%; height: auto; display: block; border: 0; max-width: 100%;">
          </div>
      </td>` : '';

    const contentCell = `
      <td class="responsive-cell" style="vertical-align: top; padding: 10px 0;">
          <h2 style="font-family: ${BRAND.fonts.headings}; color: ${BRAND.colors.primaryBlue}; font-size: 24pt; font-weight: 600; margin: 0 0 15px 0;">${topic.title}</h2>
          ${topic.description ? `<div class="description-box" style="background-color: ${BRAND.colors.accentBg}; padding: 18px; border-radius: 6px; border-left: 4px solid ${BRAND.colors.primaryBlue};"><div class="description-text" style="color: ${BRAND.colors.primaryGray}; font-size: 11pt; line-height: 1.6;">${topic.description}</div></div>` : ''}
          ${topic.buttonText && topic.buttonUrl ? `<div style="text-align: center; margin-top: 15px;">${createButtonHTML(topic.buttonText, topic.buttonUrl)}</div>` : ''}
      </td>`;

    return divider + `
      <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation">
          <tr>${isEven ? imageCell + contentCell : contentCell + imageCell}</tr>
      </table>`;
  }).join('');
}

function generateStackedLayout(topics) {
  return topics.map((topic, index) => {
    return (index > 0 ? getDividerHTML() : '') + `
      <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation">
          <tr>
              <td>
                  <h2 style="font-family: ${BRAND.fonts.headings}; color: ${BRAND.colors.primaryBlue}; font-size: 24pt; margin: 0 0 15px 0;">${topic.title}</h2>
                  ${topic.url ? `<div style="margin-bottom: 20px; border-radius: 8px; overflow: hidden; border: 1px solid ${BRAND.colors.accentBg};"><img src="${topic.url}" alt="${topic.title}" style="width: 100%; height: auto; display: block; border: 0; max-width: 100%;"></div>` : ''}
                  ${topic.description ? `<div class="description-box" style="background-color: ${BRAND.colors.accentBg}; padding: 20px; border-radius: 6px; border-left: 4px solid ${BRAND.colors.primaryBlue};"><div class="description-text" style="color: ${BRAND.colors.primaryGray}; font-size: 11pt; line-height: 1.6;">${topic.description}</div></div>` : ''}
                  ${topic.buttonText && topic.buttonUrl ? `<div style="text-align: center; margin-top: 15px;">${createButtonHTML(topic.buttonText, topic.buttonUrl)}</div>` : ''}
              </td>
          </tr>
      </table>`;
  }).join('');
}

function generateHeroLayout(topics) {
  if (topics.length === 0) return '';
  let html = '';
  const hero = topics[0];
  
  html += `
    <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation">
      <tr>
        <td>
          <h2 style="font-family: ${BRAND.fonts.headings}; color: ${BRAND.colors.primaryBlue}; font-size: 24pt; margin: 0 0 20px 0; text-align: center;">${hero.title}</h2>
          ${hero.url ? `<div style="margin-bottom: 25px; border-radius: 12px; overflow: hidden; border: 1px solid ${BRAND.colors.accentBg};"><img src="${hero.url}" alt="${hero.title}" style="width: 100%; height: auto; display: block; border: 0; max-width: 100%;"></div>` : ''}
          ${hero.description ? `<div class="description-box" style="background-color: ${BRAND.colors.accentBg}; background: ${BRAND.gradients.light}; padding: 25px; border-radius: 8px; border-left: 4px solid ${BRAND.colors.primaryBlue};"><div class="description-text" style="color: ${BRAND.colors.primaryGray}; font-size: 11pt; line-height: 1.6; text-align: center;">${hero.description}</div></div>` : ''}
          ${hero.buttonText && hero.buttonUrl ? `<div style="text-align: center; margin-top: 20px;">${createButtonHTML(hero.buttonText, hero.buttonUrl, 'blue', '12px 24px', '12pt')}</div>` : ''}
        </td>
      </tr>
    </table>`;

  if (topics.length > 1) {
    html += getDividerHTML();
    html += generateOffsetLayout(topics.slice(1));
  }
  return html;
}

function getDividerHTML() {
  return `<table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" style="margin-bottom: 35px;"><tr><td style="border-bottom: 1px solid ${BRAND.colors.accentBg};"></td></tr></table>`;
}

// --- UTILITIES & RICH TEXT HANDLING ---

function getLogosFromConfig() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) return { main: '', secondary: '' };
    const main = sheet.getRange('A2').getValue();
    const sec = sheet.getRange('B2').getValue();
    return {
      main: main ? convertDriveImageUrl(main.toString()) : '',
      secondary: sec ? convertDriveImageUrl(sec.toString()) : ''
    };
  } catch (e) {
    console.error('Config sheet error:', e);
    return { main: '', secondary: '' };
  }
}

function convertDriveImageUrl(url) {
  if (!url || typeof url !== 'string') return '';
  const match = url.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
  return match ? `https://drive.google.com/uc?export=view&id=${match[1]}` : url;
}

// --- RESTORED RICH TEXT LOGIC (With Headers Fix) ---

/**
 * Kept for backward compatibility if used elsewhere,
 * but getNewsletterDataFromColumn now uses getFormattedCellValueFromData
 */
function getFormattedCellValue(sheet, cellAddress, mode = 'body') {
  if (!sheet || !cellAddress) return '';
  try {
    const range = sheet.getRange(cellAddress);
    const richTextValue = range.getRichTextValue();
    const plainValue = range.getValue();
    return getFormattedCellValueFromData(richTextValue, plainValue, mode);
  } catch (error) {
    return sheet.getRange(cellAddress).getValue() || '';
  }
}

/**
 * Processes pre-fetched cell data (for batch operations)
 * @param {RichTextValue} richTextValue - The rich text value from the cell
 * @param {*} plainValue - The plain value from the cell
 * @param {string} mode - Processing mode ('header' or 'body')
 * @return {string} Formatted HTML string
 */
function getFormattedCellValueFromData(richTextValue, plainValue, mode = 'body') {
  // If rich text exists, convert it.
  if (richTextValue && richTextValue.getRuns().length > 0) {
    return convertRichTextToHtml(richTextValue, mode);
  }

  // Fallback to plain value, but STILL process line breaks
  if (plainValue && typeof plainValue === 'string') {
    return mode === 'header' ? processTextForHeaders(plainValue) : processTextForBody(plainValue);
  }
  return plainValue ? plainValue.toString() : '';
}

function convertRichTextToHtml(richTextValue, mode) {
  if (!richTextValue) return '';
  const textRuns = richTextValue.getRuns();
  let contentWithTags = '';
  
  for (const run of textRuns) {
    let runText = run.getText();
    const textStyle = run.getTextStyle();
    const isBold = textStyle.isBold();
    const isItalic = textStyle.isItalic();

    // Split run text by paragraph breaks (double newlines)
    // This prevents styling tags from spanning across paragraphs, which causes invalid HTML
    const parts = runText.split(/(\n{2,})/);

    for (const part of parts) {
      if (part.match(/^\n{2,}$/)) {
        // Append delimiters (newlines) without styling
        contentWithTags += part;
      } else if (part.length > 0) {
        // Apply styling to text content
        let partText = part;
        if (isBold) partText = `<strong>${partText}</strong>`;
        if (isItalic) partText = `<em>${partText}</em>`;
        contentWithTags += partText;
      }
    }
  }
  return mode === 'header' ? processTextForHeaders(contentWithTags) : processTextForBody(contentWithTags);
}

// NEW: For Headers (Titles) - Uses <br> instead of <p> to avoid nested tags
function processTextForHeaders(text) {
  if (!text || typeof text !== 'string') return '';
  // Normalize and just replace newlines with <br>
  let processedText = text.replace(/\r\n|\r/g, '\n').trim();
  return processedText.replace(/\n/g, '<br>');
}

// EXISTING: For Body (Descriptions) - Uses <p> for proper block formatting
function processTextForBody(text) {
  if (!text || typeof text !== 'string') return '';

  let processedText = text.replace(/\r\n|\r/g, '\n').trim();
  const paragraphs = processedText.split(/\n{2,}/);

  return paragraphs.map(p => {
    if (p.trim() === '') return '';
    const content = p.replace(/\n/g, '<br>');
    return `<p style="margin: 0 0 10px 0;">${content}</p>`;
  }).join('');
}

function stripHtmlTags(html) { return html ? html.replace(/<[^>]+>/g, '') : ''; }
function sanitizeHtml(html) { return html; }