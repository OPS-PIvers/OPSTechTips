/**
 * Professional Newsletter Generator for Google Apps Script
 * Multi-user system with date/column selection
 * 
 * Sheet Structure:
 * Row 1: Dates across columns B, C, D, E, F
 * Column A: Labels (A2=Title, A3=Subtitle, A4=Topic1Title, etc.)
 * Each column B-F: Newsletter data for that date/person
 * 
 * Data Layout (same for each column):
 * Row 2: Title, Row 3: Subtitle
 * Row 4: Topic 1 Title, Row 5: Topic 1 URL, Row 6: Topic 1 Text
 * Row 7: Topic 2 Title, Row 8: Topic 2 URL, Row 9: Topic 2 Description
 * Row 10: Topic 3 Title, Row 11: Topic 3 URL, Row 12: Topic 3 Description
 * Row 13: Final Button URL
 * Row 14: To, Row 15: CC, Row 16: BCC
 * Row 17: Layout Style ("Stacked", "Offset", or "Hero" - defaults to "Offset")
 */

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Newsletter Tools')
    .addItem('Send Newsletter', 'showColumnPicker')
    .addItem('Preview Newsletter', 'showPreviewPicker')
    .addSeparator()
    .addItem('Generate HTML Only', 'showGeneratePicker')
    .addToUi();
}

/**
 * Shows dialog to select column/date for sending
 */
function showColumnPicker() {
  const html = createColumnPickerDialog('send');
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(300)
    .setTitle('Select Newsletter to Send');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Send Newsletter');
}

/**
 * Shows dialog to select column/date for preview
 */
function showPreviewPicker() {
  const html = createColumnPickerDialog('preview');
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(300)
    .setTitle('Select Newsletter to Preview');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Preview Newsletter');
}

/**
 * Shows dialog to select column/date for HTML generation
 */
function showGeneratePicker() {
  const html = createColumnPickerDialog('generate');
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(300)
    .setTitle('Select Newsletter to Generate');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Generate HTML');
}

/**
 * Creates HTML dialog for column selection
 * @param {string} action - The action to perform (send, preview, generate)
 * @returns {string} HTML for dialog
 */
function createColumnPickerDialog(action) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const columns = ['B', 'C', 'D', 'E', 'F'];
  const options = [];
  
  // Get dates from row 1 for each column
  columns.forEach(col => {
    const dateCell = sheet.getRange(col + '1').getValue();
    const dateStr = dateCell ? Utilities.formatDate(new Date(dateCell), Session.getScriptTimeZone(), 'MM/dd/yyyy') : 'No Date';
    const titleCell = sheet.getRange(col + '2').getValue();
    const titleStr = titleCell ? titleCell.toString().substring(0, 30) + '...' : 'No Title';
    
    options.push({
      column: col,
      date: dateStr,
      title: titleStr,
      label: `${dateStr} - ${titleStr}`
    });
  });
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        .option { margin: 10px 0; padding: 10px; border: 1px solid #ddd; border-radius: 5px; }
        .option:hover { background-color: #f5f5f5; }
        .btn { background-color: #2d3f89; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; margin: 5px; }
        .btn:hover { background-color: #1d2a5d; }
        .btn-cancel { background-color: #666666; }
      </style>
    </head>
    <body>
      <h3>Select Newsletter to ${action.charAt(0).toUpperCase() + action.slice(1)}:</h3>
      <form>
        ${options.map(opt => `
          <div class="option">
            <label>
              <input type="radio" name="column" value="${opt.column}"> 
              <strong>${opt.date}</strong><br>
              <small>${opt.title}</small>
            </label>
          </div>
        `).join('')}
      </form>
      <br>
      <button class="btn" onclick="executeAction()">${action.charAt(0).toUpperCase() + action.slice(1)}</button>
      <button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>
      
      <script>
        function executeAction() {
          const selected = document.querySelector('input[name="column"]:checked');
          if (!selected) {
            alert('Please select a newsletter to ${action}.');
            return;
          }
          
          const column = selected.value;
          const action = '${action}';
          
          if (action === 'send') {
            google.script.run
              .withSuccessHandler(() => {
                alert('Newsletter sent successfully!');
                google.script.host.close();
              })
              .withFailureHandler((error) => {
                alert('Error sending newsletter: ' + error.message);
              })
              .sendNewsletterFromColumn(column);
          } else if (action === 'preview') {
            google.script.run
              .withSuccessHandler((html) => {
                const newWindow = window.open();
                newWindow.document.write(html);
                google.script.host.close();
              })
              .withFailureHandler((error) => {
                alert('Error generating preview: ' + error.message);
              })
              .generateNewsletterHTMLFromColumn(column);
          } else if (action === 'generate') {
            google.script.run
              .withSuccessHandler((html) => {
                alert('HTML generated and logged. Check execution transcript for details.');
                google.script.host.close();
              })
              .withFailureHandler((error) => {
                alert('Error generating HTML: ' + error.message);
              })
              .generateNewsletterHTMLFromColumn(column);
          }
        }
      </script>
    </body>
    </html>
  `;
}

/**
 * Generates HTML newsletter from specified column
 * @param {string} column - Column letter (B, C, D, E, F)
 * @returns {string} Complete HTML newsletter
 */
function generateNewsletterHTMLFromColumn(column) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = getNewsletterDataFromColumn(sheet, column);
    
    const html = createNewsletterHTML(data);
    
    console.log(`Generated HTML Newsletter from Column ${column}`);
    console.log('HTML length:', html.length, 'characters');
    return html;
    
  } catch (error) {
    console.error(`Error generating newsletter HTML from column ${column}:`, error);
    throw new Error('Failed to generate newsletter HTML: ' + error.message);
  }
}

/**
 * Sends newsletter email from specified column
 * @param {string} column - Column letter (B, C, D, E, F)
 * @returns {boolean} Success status
 */
function sendNewsletterFromColumn(column) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = getNewsletterDataFromColumn(sheet, column);
    
    if (!data.to) {
      throw new Error(`No recipients specified in "To" field for column ${column}`);
    }
    
    if (!data.title) {
      throw new Error(`Newsletter title is required for column ${column}`);
    }
    
    const html = createNewsletterHTML(data);
    const subject = data.title + (data.date ? ' - ' + Utilities.formatDate(new Date(data.date), Session.getScriptTimeZone(), 'MM/dd/yyyy') : '');
    
    GmailApp.sendEmail(
      data.to,
      subject,
      '',
      {
        htmlBody: html,
        cc: data.cc || '',
        bcc: data.bcc || '',
        attachments: []
      }
    );
    
    console.log(`Newsletter sent successfully from column ${column} to:`, data.to);
    return true;
    
  } catch (error) {
    console.error(`Error sending newsletter from column ${column}:`, error);
    throw new Error('Failed to send newsletter: ' + error.message);
  }
}

/**
 * Legacy functions for backward compatibility (use Column B)
 */
function generateNewsletterHTML() {
  return generateNewsletterHTMLFromColumn('B');
}

function sendNewsletterEmail() {
  return sendNewsletterFromColumn('B');
}

/**
 * Extracts newsletter data from specified column
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The active sheet
 * @param {string} column - Column letter (B, C, D, E, F)
 * @returns {Object} Newsletter data object
 */
function getNewsletterDataFromColumn(sheet, column) {
  const data = {
    date: sheet.getRange(column + '1').getValue(),
    title: sheet.getRange(column + '2').getValue(),
    subtitle: sheet.getRange(column + '3').getValue(),
    topic1: {
      title: sheet.getRange(column + '4').getValue(),
      url: sheet.getRange(column + '5').getValue(),
      text: sheet.getRange(column + '6').getValue()
    },
    topic2: {
      title: sheet.getRange(column + '7').getValue(),
      url: sheet.getRange(column + '8').getValue(),
      description: sheet.getRange(column + '9').getValue()
    },
    topic3: {
      title: sheet.getRange(column + '10').getValue(),
      url: sheet.getRange(column + '11').getValue(),
      description: sheet.getRange(column + '12').getValue()
    },
    finalButtonUrl: sheet.getRange(column + '13').getValue(),
    to: sheet.getRange(column + '14').getValue(),
    cc: sheet.getRange(column + '15').getValue(),
    bcc: sheet.getRange(column + '16').getValue(),
    layoutStyle: sheet.getRange(column + '17').getValue()
  };
  
  return data;
}

/**
 * Legacy function for backward compatibility
 */
function getNewsletterData(sheet) {
  return getNewsletterDataFromColumn(sheet, 'B');
}

/**
 * Converts Google Drive sharing URL to direct image URL, or any URL to base64 for email embedding
 * @param {string} url - Image URL
 * @returns {string} Base64 data URL or original URL if conversion fails
 */
function convertDriveImageUrl(url) {
  if (!url || typeof url !== 'string') return '';
  
  // First handle Google Drive URLs
  const drivePattern = /drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/;
  const match = url.match(drivePattern);
  
  let imageUrl = url;
  if (match && match[1]) {
    imageUrl = `https://drive.google.com/uc?id=${match[1]}`;
  }
  
  // Convert any image URL to base64 for reliable email delivery
  return convertImageToBase64(imageUrl);
}

/**
 * Converts image URL to base64 data URL for email embedding
 * @param {string} url - Image URL
 * @returns {string} Base64 data URL or original URL if conversion fails
 */
function convertImageToBase64(url) {
  try {
    console.log('Converting image to base64:', url);
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; GoogleAppsScript)'
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      console.error('Failed to fetch image:', response.getResponseCode(), response.getContentText());
      return url; // fallback to original URL
    }
    
    const blob = response.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    
    // Validate it's an image
    if (!mimeType.startsWith('image/')) {
      console.error('URL does not return an image:', mimeType);
      return url;
    }
    
    // Check size (warn if over 100KB)
    const sizeKB = blob.getBytes().length / 1024;
    if (sizeKB > 100) {
      console.warn(`Large image detected: ${sizeKB.toFixed(1)}KB. Consider optimizing for faster email loading.`);
    }
    
    const base64Url = `data:${mimeType};base64,${base64}`;
    console.log(`Image converted successfully: ${sizeKB.toFixed(1)}KB`);
    return base64Url;
    
  } catch (error) {
    console.error('Failed to convert image to base64:', error);
    return url; // fallback to original URL
  }
}

/**
 * Creates the complete HTML newsletter
 * @param {Object} data - Newsletter data
 * @returns {string} Complete HTML newsletter
 */
function createNewsletterHTML(data) {
  const topics = [];
  
  if (data.topic1.title && data.topic1.url) {
    topics.push({
      title: data.topic1.title,
      url: convertDriveImageUrl(data.topic1.url),
      description: data.topic1.text || ''
    });
  }
  
  if (data.topic2.title && data.topic2.url) {
    topics.push({
      title: data.topic2.title,
      url: convertDriveImageUrl(data.topic2.url),
      description: data.topic2.description || ''
    });
  }
  
  if (data.topic3.title && data.topic3.url) {
    topics.push({
      title: data.topic3.title,
      url: convertDriveImageUrl(data.topic3.url),
      description: data.topic3.description || ''
    });
  }
  
  const layoutStyle = data.layoutStyle ? data.layoutStyle.toLowerCase() : 'offset';
  
  let topicHTML;
  if (layoutStyle === 'stacked') {
    topicHTML = generateStackedLayout(topics);
  } else if (layoutStyle === 'hero') {
    topicHTML = generateHeroLayout(topics);
  } else {
    topicHTML = generateOffsetLayout(topics);
  }
  
  const html = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${data.title || 'Newsletter'}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Lato:wght@300;400;600;700&display=swap');
    </style>
</head>
<body style="margin: 0; padding: 0; background-color: #f3f3f3; font-family: 'Lato', Arial, sans-serif;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #f3f3f3; padding: 20px 0;">
        <tr>
            <td align="center">
                <table width="600" cellpadding="0" cellspacing="0" border="0" style="background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 12px rgba(45, 63, 137, 0.1);">
                    
                    <!-- Header -->
                    <tr>
                        <td style="background: linear-gradient(135deg, #2d3f89 0%, #4356a0 100%); padding: 40px 30px; text-align: center;">
                            ${data.date ? `<div style="color: #eaecf5; font-size: 14px; font-weight: 400; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 1px;">${Utilities.formatDate(new Date(data.date), Session.getScriptTimeZone(), 'MMMM yyyy')}</div>` : ''}
                            ${data.title ? `<h1 style="color: #ffffff; font-size: 26px; font-weight: 700; margin: 0 0 10px 0; line-height: 1.2;">${data.title}</h1>` : ''}
                            ${data.subtitle ? `<p style="color: #eaecf5; font-size: 16px; font-weight: 400; margin: 0; line-height: 1.4;">${data.subtitle}</p>` : ''}
                        </td>
                    </tr>
                    
                    <!-- Content -->
                    <tr>
                        <td style="padding: 40px 30px;">
                            
                            ${topicHTML}
                            
                            ${data.finalButtonUrl ? `
                            <!-- Call to Action -->
                            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top: 40px;">
                                <tr>
                                    <td align="center" style="background: linear-gradient(135deg, #eaecf5 0%, #f3f3f3 100%); padding: 30px; border-radius: 8px;">
                                        <h3 style="color: #2d3f89; font-size: 18px; font-weight: 600; margin: 0 0 20px 0;">Ready to Learn More?</h3>
                                        <a href="${data.finalButtonUrl}" style="display: inline-block; background: linear-gradient(135deg, #ad2122 0%, #c13435 100%); color: #ffffff; text-decoration: none; padding: 14px 32px; border-radius: 6px; font-size: 16px; font-weight: 600; transition: all 0.3s ease; box-shadow: 0 3px 8px rgba(173, 33, 34, 0.3);">
                                            Visit the Orono Technology Digital Learning Hub to learn more
                                        </a>
                                    </td>
                                </tr>
                            </table>
                            ` : ''}
                            
                        </td>
                    </tr>
                    
                    <!-- Footer -->
                    <tr>
                        <td style="background-color: #1d2a5d; padding: 25px 30px; text-align: center;">
                            <p style="color: #eaecf5; font-size: 12px; font-weight: 400; margin: 0; line-height: 1.5;">
                                ${new Date().getFullYear()} Orono Technology Digital Learning Hub<br>
                                <span style="color: #4356a0;">Empowering Digital Learning and Innovation</span>
                            </p>
                        </td>
                    </tr>
                    
                </table>
            </td>
        </tr>
    </table>
</body>
</html>`;
  
  return html;
}

/**
 * Generates stacked (full-width) layout for topics
 * @param {Array} topics - Array of topic objects
 * @returns {string} HTML for stacked layout
 */
function generateStackedLayout(topics) {
  return topics.map((topic, index) => `
                            <!-- Topic ${index + 1} - Stacked Layout -->
                            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 35px;">
                                <tr>
                                    <td>
                                        <h2 style="color: #1d2a5d; font-size: 22px; font-weight: 600; margin: 0 0 15px 0; line-height: 1.3;">${topic.title}</h2>
                                        
                                        ${topic.url ? `
                                        <div style="margin-bottom: 20px; border-radius: 8px; overflow: hidden; border: 1px solid #eaecf5;">
                                            <img src="${topic.url}" alt="${topic.title}" style="width: 100%; height: auto; display: block; max-height: 300px; object-fit: cover;">
                                        </div>
                                        ` : ''}
                                        
                                        ${topic.description ? `
                                        <div style="background-color: #eaecf5; padding: 20px; border-radius: 6px; border-left: 4px solid #2d3f89;">
                                            <p style="color: #333333; font-size: 14px; font-weight: 400; margin: 0; line-height: 1.6;">${topic.description}</p>
                                        </div>
                                        ` : ''}
                                    </td>
                                </tr>
                            </table>
                            `).join('');
}

/**
 * Generates hero layout for topics (main feature + two columns)
 * @param {Array} topics - Array of topic objects
 * @returns {string} HTML for hero layout
 */
function generateHeroLayout(topics) {
  let html = '';
  
  if (topics.length > 0) {
    const heroTopic = topics[0];
    html += `
                            <!-- Hero Section -->
                            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 40px;">
                                <tr>
                                    <td>
                                        <h2 style="color: #1d2a5d; font-size: 24px; font-weight: 700; margin: 0 0 20px 0; line-height: 1.2; text-align: center;">${heroTopic.title}</h2>
                                        
                                        ${heroTopic.url ? `
                                        <div style="margin-bottom: 25px; border-radius: 12px; overflow: hidden; border: 1px solid #eaecf5;">
                                            <img src="${heroTopic.url}" alt="${heroTopic.title}" style="width: 100%; height: auto; display: block; max-height: 350px; object-fit: cover;">
                                        </div>
                                        ` : ''}
                                        
                                        ${heroTopic.description ? `
                                        <div style="background: linear-gradient(135deg, #eaecf5 0%, #f3f3f3 100%); padding: 25px; border-radius: 8px; border-left: 4px solid #2d3f89;">
                                            <p style="color: #333333; font-size: 16px; font-weight: 400; margin: 0; line-height: 1.6; text-align: center;">${heroTopic.description}</p>
                                        </div>
                                        ` : ''}
                                    </td>
                                </tr>
                            </table>
                            `;
  }
  
  if (topics.length > 1) {
    const leftTopic = topics[1];
    const rightTopic = topics.length > 2 ? topics[2] : null;
    
    html += `
                            <!-- Two Column Section -->
                            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 35px;">
                                <tr>
                                    <!-- Left Column -->
                                    <td width="48%" style="vertical-align: top; padding-right: ${rightTopic ? '15px' : '0'};">
                                        <h3 style="color: #1d2a5d; font-size: 18px; font-weight: 600; margin: 0 0 15px 0; line-height: 1.3;">${leftTopic.title}</h3>
                                        
                                        ${leftTopic.url ? `
                                        <div style="margin-bottom: 15px; border-radius: 6px; overflow: hidden; border: 1px solid #eaecf5;">
                                            <img src="${leftTopic.url}" alt="${leftTopic.title}" style="width: 100%; height: auto; display: block; max-height: 200px; object-fit: cover;">
                                        </div>
                                        ` : ''}
                                        
                                        ${leftTopic.description ? `
                                        <div style="background-color: #eaecf5; padding: 15px; border-radius: 6px; border-left: 3px solid #2d3f89;">
                                            <p style="color: #333333; font-size: 13px; font-weight: 400; margin: 0; line-height: 1.5;">${leftTopic.description}</p>
                                        </div>
                                        ` : ''}
                                    </td>
                                    
                                    ${rightTopic ? `
                                    <!-- Right Column -->
                                    <td width="4%" style="padding: 0;"></td>
                                    <td width="48%" style="vertical-align: top; padding-left: 15px;">
                                        <h3 style="color: #1d2a5d; font-size: 18px; font-weight: 600; margin: 0 0 15px 0; line-height: 1.3;">${rightTopic.title}</h3>
                                        
                                        ${rightTopic.url ? `
                                        <div style="margin-bottom: 15px; border-radius: 6px; overflow: hidden; border: 1px solid #eaecf5;">
                                            <img src="${rightTopic.url}" alt="${rightTopic.title}" style="width: 100%; height: auto; display: block; max-height: 200px; object-fit: cover;">
                                        </div>
                                        ` : ''}
                                        
                                        ${rightTopic.description ? `
                                        <div style="background-color: #eaecf5; padding: 15px; border-radius: 6px; border-left: 3px solid #2d3f89;">
                                            <p style="color: #333333; font-size: 13px; font-weight: 400; margin: 0; line-height: 1.5;">${rightTopic.description}</p>
                                        </div>
                                        ` : ''}
                                    </td>
                                    ` : `<td width="52%"></td>`}
                                </tr>
                            </table>
                            `;
  }
  
  return html;
}

/**
 * Generates offset (alternating) layout for topics
 * @param {Array} topics - Array of topic objects
 * @returns {string} HTML for offset layout
 */
function generateOffsetLayout(topics) {
  return topics.map((topic, index) => {
    const isEven = index % 2 === 0;
    const imageCell = topic.url ? `
        <td width="250" style="padding: ${isEven ? '0 20px 0 0' : '0 0 0 20px'}; vertical-align: top;">
            <div style="border-radius: 8px; overflow: hidden; border: 1px solid #eaecf5;">
                <img src="${topic.url}" alt="${topic.title}" style="width: 250px; height: 180px; display: block; object-fit: cover;">
            </div>
        </td>
    ` : '';
    
    const contentCell = `
        <td style="vertical-align: top; padding: 10px 0;">
            <h2 style="color: #1d2a5d; font-size: 22px; font-weight: 600; margin: 0 0 15px 0; line-height: 1.3;">${topic.title}</h2>
            ${topic.description ? `
            <div style="background-color: #eaecf5; padding: 18px; border-radius: 6px; border-left: 4px solid #2d3f89;">
                <p style="color: #333333; font-size: 14px; font-weight: 400; margin: 0; line-height: 1.6;">${topic.description}</p>
            </div>
            ` : ''}
        </td>
    `;
    
    return `
                            <!-- Topic ${index + 1} - Offset Layout -->
                            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 35px;">
                                <tr>
                                    ${isEven ? imageCell + contentCell : contentCell + imageCell}
                                </tr>
                            </table>
                            `;
  }).join('');
}

/**
 * Test function to preview newsletter HTML from any column
 * @param {string} column - Column letter (B, C, D, E, F) - defaults to B
 */
function testNewsletterGeneration(column = 'B') {
  try {
    const html = generateNewsletterHTMLFromColumn(column);
    const sheet = SpreadsheetApp.getActiveSheet();
    const layoutStyle = sheet.getRange(column + '17').getValue() || 'Offset';
    const date = sheet.getRange(column + '1').getValue();
    const title = sheet.getRange(column + '2').getValue();
    
    console.log(`Test completed successfully for Column ${column}`);
    console.log('Date:', date ? Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'MM/dd/yyyy') : 'No date');
    console.log('Title:', title || 'No title');
    console.log('Layout Style:', layoutStyle);
    console.log('HTML length:', html.length, 'characters');
    return html;
  } catch (error) {
    console.error(`Test failed for column ${column}:`, error);
    return null;
  }
}

/**
 * Utility function to validate email addresses
 * @param {string} email - Email address to validate
 * @returns {boolean} True if valid email format
 */
function isValidEmail(email) {
  if (!email) return false;
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}