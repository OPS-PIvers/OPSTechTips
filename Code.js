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
    .addItem('Generate Landing Page', 'generateLandingPage')
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
  const lastCol = sheet.getLastColumn();
  if (lastCol < 2) return '<h3>No newsletter columns found.</h3>';
  
  // Read row 1 (Date), row 2 (Title), and row 4 (Topic 1 Title) 
  // from Column B (2) to the end
  const dataRange = sheet.getRange(1, 2, 4, lastCol - 1);
  const dataValues = dataRange.getValues(); 
  
  const options = [];
  for (let i = 0; i < (lastCol - 1); i++) {
    const topic1Title = dataValues[3][i]; // Row 4
    
    // Only include columns that have a Topic 1 Title in row 4
    if (!topic1Title || topic1Title.toString().trim() === '') continue;

    const dateCell = dataValues[0][i]; // Row 1
    const titleCell = dataValues[1][i]; // Row 2
    
    const dateStr = dateCell ? Utilities.formatDate(new Date(dateCell), Session.getScriptTimeZone(), 'MM/dd/yyyy') : 'No Date';
    const displayTitle = titleCell ? titleCell.toString() : 'Untitled Newsletter';
    const titleStr = displayTitle.substring(0, 30) + (displayTitle.length > 30 ? '...' : '');
    
    const actualColLetter = getColumnLetter(i + 2);

    options.push({
      column: actualColLetter,
      date: dateStr,
      title: titleStr
    });
  }

  if (options.length === 0) return '<h3>No newsletters found with titles.</h3>';

  try {
    const template = HtmlService.createTemplateFromFile('Picker');
    template.action = action;
    template.options = options;
    return template.evaluate().getContent();
  } catch (e) {
    return '<h3>Error: Picker.html file missing. Please create it.</h3>';
  }
}

/**
 * Helper to convert column index to Letter (e.g., 2 -> B, 28 -> AB)
 */
function getColumnLetter(col) {
  let letter = '';
  while (col > 0) {
    let temp = (col - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

// --- CORE LOGIC ---

function generateNewsletterHTMLFromColumn(column) {
  try {
    SpreadsheetApp.getActive().toast('Generating HTML...', 'Status', 3);
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = getNewsletterDataFromColumn(sheet, column);
    const html = createNewsletterHTML(data);
    // Encode to HTML entities to match Draft output and ensure stability
    return encodeHtmlEntities(html);
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
    
    let html = createNewsletterHTML(data);
    const subject = stripHtmlTags(data.title) + (data.date ? ' - ' + Utilities.formatDate(new Date(data.date), Session.getScriptTimeZone(), 'MM/dd/yyyy') : '');
    
    // Fix for emojis: Convert non-ASCII characters to HTML entities
    // This bypasses transport encoding issues by sending safe ASCII text that renders as Unicode
    const htmlBody = encodeHtmlEntities(html);

    const options = {
      htmlBody: htmlBody,
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

function generateLandingPage() {
  try {
    SpreadsheetApp.getActive().toast('Building Landing Page...', 'Status', -1);
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastCol = sheet.getLastColumn();
    
    if (lastCol < 2) throw new Error('No newsletter columns found.');
    
    // Batch read row 4 for all columns to find valid newsletters quickly
    const topic1Range = sheet.getRange(4, 2, 1, lastCol - 1);
    const topic1Values = topic1Range.getValues()[0];
    
    const newsletters = [];
    for (let i = 0; i < topic1Values.length; i++) {
      const topic1Title = topic1Values[i];
      
      // Only include if Row 4 (Topic 1 Title) is not empty
      if (topic1Title && topic1Title.toString().trim() !== '') {
        const colLetter = getColumnLetter(i + 2);
        const data = getNewsletterDataFromColumn(sheet, colLetter);
        newsletters.push({
          ...data,
          column: colLetter
        });
      }
    }

    if (newsletters.length === 0) {
      throw new Error('No newsletters found with a Topic 1 Title in Row 4.');
    }

    const html = createLandingPageHTML(newsletters);
    const encodedHtml = encodeHtmlEntities(html);
    
    // Reuse the Picker dialog logic to show the generated HTML
    const template = HtmlService.createTemplateFromFile('Picker');
    template.action = 'generate';
    template.options = []; // Not needed for the result view
    
    const htmlOutput = HtmlService.createHtmlOutput(template.evaluate().getContent())
      .setWidth(450)
      .setHeight(500)
      .setTitle('Landing Page HTML');
    
    // We need to pass the result to the client side
    // Since we're reusing the template, we'll use a trick or just show a separate simple dialog
    showGeneratedHtmlDialog(encodedHtml, 'Landing Page HTML');
    
    SpreadsheetApp.getActive().toast('Landing Page Generated!', 'Success', 3);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
  }
}

function showGeneratedHtmlDialog(html, title) {
  const css = `
    <style>
      body { font-family: 'Roboto', sans-serif; padding: 20px; background: #f9fafb; }
      h3 { font-family: 'Lexend', sans-serif; color: #2d3f89; margin-top: 0; }
      textarea { width: 100%; height: 300px; margin: 15px 0; border: 1px solid #d1d5db; border-radius: 8px; padding: 12px; font-family: monospace; resize: none; }
      .btn { background: #2d3f89; color: white; padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; font-weight: 600; }
      .btn:hover { background: #1d2a5d; }
    </style>
  `;
  
  const content = `
    ${css}
    <h3>${title}</h3>
    <p style="font-size: 14px; color: #666;">Copy this HTML to create a summary page of all your tech tips.</p>
    <textarea id="output" readonly>${html}</textarea>
    <button class="btn" onclick="copyText()">Copy to Clipboard</button>
    <script>
      function copyText() {
        const copyText = document.getElementById("output");
        copyText.select();
        document.execCommand("copy");
        const btn = event.target;
        btn.innerText = "Copied!";
        btn.style.background = "#10b981";
        setTimeout(() => { btn.innerText = "Copy to Clipboard"; btn.style.background = "#2d3f89"; }, 2000);
      }
    </script>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(content)
    .setWidth(450)
    .setHeight(480)
    .setTitle(title);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}

function createLandingPageHTML(newsletters) {
  const logos = getLogosFromConfig();
  
  const cardsHTML = newsletters.map(n => {
    // Get the first available image and convert it to preview format
    const rawImage = n.topic1.url || n.topic2.url || n.topic3.url || '';
    const heroImage = convertDriveImageUrl(rawImage);
    const dateStr = n.date ? Utilities.formatDate(new Date(n.date), Session.getScriptTimeZone(), 'MMMM d, yyyy') : '';
    
    return `
      <!-- Card -->
      <div class="card" style="background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.08); transition: transform 0.2s; margin-bottom: 30px;">
        ${heroImage ? `<div style="height: 200px; overflow: hidden;"><img src="${heroImage}" style="width: 100%; height: 100%; object-fit: cover;"></div>` : ''}
        <div style="padding: 25px;">
          <div style="color: ${BRAND.colors.primaryBlue}; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 8px;">${dateStr}</div>
          <h2 style="font-family: ${BRAND.fonts.headings}; color: #333; margin: 0 0 12px 0; font-size: 20px; line-height: 1.3;">${n.title}</h2>
          <div style="color: #666; font-size: 15px; line-height: 1.5; margin-bottom: 20px;">${stripHtmlTags(n.subtitle || '').substring(0, 120)}...</div>
          <a href="#" style="display: inline-block; background: ${BRAND.colors.primaryBlue}; color: white; padding: 10px 20px; border-radius: 6px; text-decoration: none; font-weight: 600; font-size: 14px;">View Tech Tip</a>
        </div>
      </div>`;
  }).join('');

  return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tech Tips Hub</title>
    <link href="https://fonts.googleapis.com/css2?family=Lexend:wght@600&family=Roboto:wght@400;500&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Roboto', sans-serif; margin: 0; padding: 0; background-color: #f3f4f6; color: #333; }
        .header { background: ${BRAND.colors.primaryBlue}; background: ${BRAND.gradients.blue}; color: white; padding: 60px 20px; text-align: center; }
        .container { max-width: 1000px; margin: -40px auto 60px auto; padding: 0 20px; }
        .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 30px; }
        h1 { font-family: 'Lexend', sans-serif; font-size: 36px; margin: 0 0 10px 0; }
        .subtitle { font-size: 18px; opacity: 0.9; }
        .card { background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.08); height: 100%; display: flex; flex-direction: column; }
        .card:hover { transform: translateY(-5px); transition: transform 0.3s ease; }
        @media (max-width: 600px) { .grid { grid-template-columns: 1fr; } }
    </style>
</head>
<body>
    <div class="header">
        ${logos.main ? `<img src="${logos.main}" alt="Logo" style="max-width: 180px; margin-bottom: 20px; filter: brightness(0) invert(1);">` : ''}
        <h1>Digital Learning Tech Tips</h1>
        <p class="subtitle">Explore our latest guides and resources for modern education.</p>
    </div>
    
    <div class="container">
        <div class="grid">
            ${cardsHTML}
        </div>
    </div>
    
    <footer style="text-align: center; padding: 40px; color: #999; font-size: 14px;">
        &copy; ${new Date().getFullYear()} Orono Technology Digital Learning Hub
    </footer>
</body>
</html>`;
}
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
    .filter((t, i) => (t.title || i === 2) && (t.url || t.description))
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
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="color-scheme" content="light only">
    <meta name="supported-color-schemes" content="light only">
    <!--[if mso]>
    <noscript>
    <xml>
    <o:OfficeDocumentSettings>
    <o:PixelsPerInch>96</o:PixelsPerInch>
    </o:OfficeDocumentSettings>
    </xml>
    </noscript>
    <![endif]-->
    <title>${stripHtmlTags(data.title) || 'Newsletter'}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Lexend:wght@400;600;700&family=Roboto:wght@400;500;700&display=swap');
        
        /* RESET */
        body { margin: 0; padding: 0; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; background-color: ${BRAND.colors.lightBg}; min-width: 320px; }
        table, td { border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; }
        img { border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic; max-width: 100%; }
        
        /* TYPOGRAPHY */
        body, td { font-family: ${BRAND.fonts.body}; font-size: 12pt; color: ${BRAND.colors.primaryGray}; line-height: 1.6; }
        h1, h2, h3 { font-family: ${BRAND.fonts.headings}; margin: 0; word-break: break-word; line-height: 1.3; }
        p { word-break: break-word; margin: 0 0 16px 0; }
        a { text-decoration: none; color: inherit; }
        p a { color: ${BRAND.colors.primaryBlue} !important; text-decoration: underline !important; }
        
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
<body id="body" style="margin: 0; padding: 0; background-color: ${BRAND.colors.lightBg}; font-family: ${BRAND.fonts.body}; color: ${BRAND.colors.primaryGray};">
    
    <!-- PREHEADER HACK -->
    <div style="display: none; max-height: 0px; overflow: hidden;">
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
                                        <p style="color: ${BRAND.colors.accentBg}; font-size: 12pt; margin: 0; line-height: 1.5;">
                                            ${new Date().getFullYear()} Orono Technology Digital Learning Hub<br>
                                            <span style="color: ${BRAND.colors.secondaryGray}; font-size: 10pt;">Empowering Responsible Digital Learning and Innovation</span>
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

function createButtonHTML(text, url, style = 'blue', padding = '14px 28px', fontSize = '12pt') {
  const bgColor = style === 'red' ? BRAND.colors.primaryRed : BRAND.colors.primaryBlue;
  const gradient = style === 'red' ? BRAND.gradients.red : BRAND.gradients.blue;
  
  return `<a href="${url}" style="background-color: ${bgColor}; background: ${gradient}; color: #ffffff; text-decoration: none; padding: ${padding}; border-radius: 6px; font-size: ${fontSize}; font-weight: 600; font-family: ${BRAND.fonts.headings}; display: inline-block; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.25); border: 0; line-height: 1.2;">${text}</a>`;
}

function generateCallToAction(url) {
  return `
    <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" style="margin-top: 40px;">
        <tr>
            <td align="center" style="background-color: ${BRAND.colors.accentBg}; background: ${BRAND.gradients.light}; padding: 30px; border-radius: 8px;">
                <h3 style="font-family: ${BRAND.fonts.headings}; color: ${BRAND.colors.primaryBlue}; font-size: 18pt; margin: 0 0 20px 0;">Ready to Learn More?</h3>
                ${createButtonHTML('Visit the Digital Learning Hub', url, 'red', '14px 32px', '14pt')}
            </td>
        </tr>
    </table>`;
}

function generateOffsetLayout(topics) {
  return topics.map((topic, index) => {
    const divider = index > 0 ? getDividerHTML() : '';
    const isEven = index % 2 === 0;
    
    const imageCell = topic.url ? `
      <td width="240" class="responsive-cell" valign="top" style="width: 240px; padding: ${isEven ? '0 20px 0 0' : '0 0 0 20px'};">
          <div style="border-radius: 8px; overflow: hidden; border: 1px solid ${BRAND.colors.accentBg}; font-size: 0; line-height: 0;">
              <img src="${topic.url}" alt="${escapeAttribute(topic.title)}" width="240" style="width: 100%; height: auto; display: block; border: 0;">
          </div>
      </td>` : '';

    const contentCell = `
      <td class="responsive-cell" valign="top" style="vertical-align: top;">
          ${topic.title ? `<h2 style="font-family: ${BRAND.fonts.headings}; color: ${BRAND.colors.primaryBlue}; font-size: 20pt; font-weight: 600; margin: 0 0 12px 0; line-height: 1.3;">${topic.title}</h2>` : ''}
          ${topic.description ? `<div style="color: ${BRAND.colors.primaryGray}; font-size: 11pt; line-height: 1.6; margin-bottom: 15px;">${topic.description}</div>` : ''}
          ${topic.buttonText && topic.buttonUrl ? `<div style="text-align: left;">${createButtonHTML(topic.buttonText, topic.buttonUrl)}</div>` : ''}
      </td>`;

    return divider + `
      <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" style="margin-bottom: 20px;">
          <tr>${isEven ? imageCell + contentCell : contentCell + imageCell}</tr>
      </table>`;
  }).join('');
}

function generateStackedLayout(topics) {
  return topics.map((topic, index) => {
    const divider = index > 0 ? getDividerHTML() : '';
    
    return divider + `
      <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" style="margin-bottom: 30px;">
          <tr>
              <td>
                  ${topic.title ? `<h2 style="font-family: ${BRAND.fonts.headings}; color: ${BRAND.colors.primaryBlue}; font-size: 24pt; margin: 0 0 20px 0; text-align: center; line-height: 1.3;">${topic.title}</h2>` : ''}
                  ${topic.url ? `<div style="margin-bottom: 25px; border-radius: 12px; overflow: hidden; border: 1px solid ${BRAND.colors.accentBg}; font-size: 0; line-height: 0;"><img src="${topic.url}" alt="${escapeAttribute(topic.title)}" width="540" style="width: 100%; height: auto; display: block; border: 0;"></div>` : ''}
                  ${topic.description ? `<div style="background-color: ${BRAND.colors.accentBg}; padding: 20px; border-radius: 10px; border-left: 5px solid ${BRAND.colors.primaryBlue};"><div style="color: ${BRAND.colors.primaryGray}; font-size: 11pt; line-height: 1.6; text-align: left;">${topic.description}</div></div>` : ''}
                  ${topic.buttonText && topic.buttonUrl ? `<div style="text-align: center; margin-top: 25px;">${createButtonHTML(topic.buttonText, topic.buttonUrl, 'blue', '12px 24px', '12pt')}</div>` : ''}
              </td>
          </tr>
      </table>`;
  }).join('');
}

function generateHeroLayout(topics) {
  if (topics.length === 0) return '';
  let html = '';
  const hero = topics[0];
  
  // Hero Section
  html += `
    <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" style="margin-bottom: 30px;">
      <tr>
        <td>
          ${hero.title ? `<h2 style="font-family: ${BRAND.fonts.headings}; color: ${BRAND.colors.primaryBlue}; font-size: 26pt; margin: 0 0 20px 0; text-align: center; line-height: 1.2;">${hero.title}</h2>` : ''}
          ${hero.url ? `<div style="margin-bottom: 25px; border-radius: 12px; overflow: hidden; border: 1px solid ${BRAND.colors.accentBg}; font-size: 0; line-height: 0;"><img src="${hero.url}" alt="${escapeAttribute(hero.title)}" width="540" style="width: 100%; height: auto; display: block; border: 0;"></div>` : ''}
          ${hero.description ? `<div style="background-color: ${BRAND.colors.accentBg}; padding: 25px; border-radius: 10px; border-left: 5px solid ${BRAND.colors.primaryBlue};"><div style="color: ${BRAND.colors.primaryGray}; font-size: 12pt; line-height: 1.6; text-align: left;">${hero.description}</div></div>` : ''}
          ${hero.buttonText && hero.buttonUrl ? `<div style="text-align: center; margin-top: 25px;">${createButtonHTML(hero.buttonText, hero.buttonUrl, 'blue', '14px 30px', '13pt')}</div>` : ''}
        </td>
      </tr>
    </table>`;

  // Sub-sections
  const subTopics = topics.slice(1);
  if (subTopics.length > 0) {
    html += getDividerHTML();
    
    if (subTopics.length === 2) {
      const t1 = subTopics[0];
      const t2 = subTopics[1];
      
      const renderGridItem = (topic) => `
        <td class="responsive-cell" width="260" valign="top" style="width: 260px; background-color: ${BRAND.colors.accentBg}; border-radius: 10px;">
           <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation">
             <tr>
               <td style="padding: 20px;">
                 ${topic.title ? `<h3 style="font-family: ${BRAND.fonts.headings}; color: ${BRAND.colors.primaryBlue}; font-size: 16pt; margin: 0 0 15px 0; line-height: 1.3;">${topic.title}</h3>` : ''}
                 ${topic.url ? `<div style="margin-bottom: 15px; border-radius: 8px; overflow: hidden; border: 1px solid #d1d5db; font-size: 0; line-height: 0;"><img src="${topic.url}" alt="${escapeAttribute(topic.title)}" width="220" style="width: 100%; height: auto; display: block; border: 0;"></div>` : ''}
                 ${topic.description ? `<div style="color: ${BRAND.colors.primaryGray}; font-size: 10pt; line-height: 1.5; margin-bottom: 15px;">${topic.description}</div>` : ''}
                 ${topic.buttonText && topic.buttonUrl ? `<div style="text-align: left;">${createButtonHTML(topic.buttonText, topic.buttonUrl, 'blue', '10px 20px', '10pt')}</div>` : ''}
               </td>
             </tr>
           </table>
        </td>`;

      html += `
        <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation">
          <tr>
            ${renderGridItem(t1)}
            <td width="20" class="responsive-cell" style="width: 20px; font-size: 0; line-height: 0;">&nbsp;</td>
            ${renderGridItem(t2)}
          </tr>
        </table>`;
    } else {
      html += generateStackedLayout(subTopics);
    }
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
  
  const cleanUrl = url.trim();
  
  // Extract ID from various Google Drive URL formats
  let id = '';
  const patterns = [
    /drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/,
    /drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/,
    /drive\.google\.com\/uc\?id=([a-zA-Z0-9_-]+)/,
    /docs\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/
  ];
  
  for (const pattern of patterns) {
    const match = cleanUrl.match(pattern);
    if (match) {
      id = match[1];
      break;
    }
  }
  
  // If no match but it looks like just an ID
  if (!id && cleanUrl.length > 20 && !cleanUrl.includes('/') && !cleanUrl.includes('.')) {
    id = cleanUrl;
  }
  
  // Use the Thumbnail/Preview endpoint which is significantly more reliable for embedding
  // sz=w1200 ensures we get a high-quality version suitable for retina displays
  return id ? `https://drive.google.com/thumbnail?id=${id}&sz=w1200` : cleanUrl;
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
function escapeAttribute(text) {
  return stripHtmlTags(text).replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function sanitizeHtml(html) { return html; }

/**
 * Encodes non-ASCII characters (including emojis) to HTML entities.
 * This ensures they survive the email transport layers without turning into '???'.
 */
function encodeHtmlEntities(text) {
  const chars = [];
  for (const char of text) {
    const code = char.codePointAt(0);
    if (code > 127) {
      chars.push('&#' + code + ';');
    } else {
      chars.push(char);
    }
  }
  return chars.join('');
}