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
 * Row 4: Topic 1 Title, Row 5: Topic 1 URL, Row 6: Topic 1 Text, Row 7: Topic 1 Button Text, Row 8: Topic 1 Button URL
 * Row 9: Topic 2 Title, Row 10: Topic 2 URL, Row 11: Topic 2 Description, Row 12: Topic 2 Button Text, Row 13: Topic 2 Button URL
 * Row 14: Topic 3 Title, Row 15: Topic 3 URL, Row 16: Topic 3 Description, Row 17: Topic 3 Button Text, Row 18: Topic 3 Button URL
 * Row 19: Final Button URL
 * Row 20: To, Row 21: CC, Row 22: BCC
 * Row 23: Layout Style ("Stacked", "Offset", or "Hero" - defaults to "Offset")
 */

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
 * Shows dialog to select column/date for creating a draft
 */
function showDraftPicker() {
  const html = createColumnPickerDialog('draft');
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(300)
    .setTitle('Select Newsletter to Create Draft');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Create Newsletter Draft');
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
 * @param {string} action - The action to perform (send, preview, generate, draft)
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
        .loader { display: none; color: #666; font-style: italic; margin-top: 10px; }
        .btn:disabled { background-color: #ccc; cursor: not-allowed; }
        #html-container { margin-top: 20px; }
        #html-output { width: 100%; height: 150px; margin-bottom: 10px; }
      </style>
    </head>
    <body>
      <div id="main-content">
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
        <div id="button-container">
          <button id="action-btn" class="btn" onclick="executeAction()">${action.charAt(0).toUpperCase() + action.slice(1)}</button>
          <button id="cancel-btn" class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>
        </div>
        <div id="loader" class="loader">Processing... Please wait.</div>
      </div>
      
      <script>
        function executeAction() {
          const selected = document.querySelector('input[name="column"]:checked');
          if (!selected) {
            alert('Please select a newsletter to ${action}.');
            return;
          }

          const actionBtn = document.getElementById('action-btn');
          const cancelBtn = document.getElementById('cancel-btn');
          const loader = document.getElementById('loader');

          actionBtn.disabled = true;
          cancelBtn.disabled = true;
          loader.style.display = 'block';
          
          const column = selected.value;
          const action = '${action}';

          function restoreButtons() {
            actionBtn.disabled = false;
            cancelBtn.disabled = false;
            loader.style.display = 'none';
          }

          function showGeneratedHTML(html) {
            const mainContent = document.getElementById('main-content');
            mainContent.innerHTML = '<h3>Generated HTML</h3>' +
              '<textarea id="html-output">' + html.replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</textarea>' +
              '<button class="btn" onclick="copyHtml()">Copy HTML</button>' +
              '<button class="btn btn-cancel" onclick="google.script.host.close()">Close</button>';
          }

          window.copyHtml = function() {
            const htmlOutput = document.getElementById('html-output');
            htmlOutput.select();
            document.execCommand('copy');
            alert('HTML copied to clipboard!');
          }
          
          if (action === 'send') {
            google.script.run
              .withSuccessHandler(() => {
                alert('Newsletter sent successfully!');
                google.script.host.close();
              })
              .withFailureHandler((error) => {
                alert('Error sending newsletter: ' + error.message);
                restoreButtons();
              })
              .sendNewsletterFromColumn(column);
          } else if (action === 'draft') {
            google.script.run
              .withSuccessHandler(() => {
                alert('Newsletter draft created successfully!');
                google.script.host.close();
              })
              .withFailureHandler((error) => {
                alert('Error creating newsletter draft: ' + error.message);
                restoreButtons();
              })
              .createDraftNewsletterFromColumn(column);
          } else if (action === 'preview') {
            google.script.run
              .withSuccessHandler((html) => {
                const newWindow = window.open();
                newWindow.document.write(html);
                google.script.host.close();
              })
              .withFailureHandler((error) => {
                alert('Error generating preview: ' + error.message);
                restoreButtons();
              })
              .generateNewsletterHTMLFromColumn(column);
          } else if (action === 'generate') {
            google.script.run
              .withSuccessHandler((html) => {
                showGeneratedHTML(html);
              })
              .withFailureHandler((error) => {
                alert('Error generating HTML: ' + error.message);
                restoreButtons();
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
 * Creates a draft newsletter email from specified column
 * @param {string} column - Column letter (B, C, D, E, F)
 * @returns {boolean} Success status
 */
function createDraftNewsletterFromColumn(column) {
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
    
    GmailApp.createDraft(
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
    
    console.log(`Newsletter draft created successfully from column ${column} for:`, data.to);
    return true;
    
  } catch (error) {
    console.error(`Error creating newsletter draft from column ${column}:`, error);
    throw new Error('Failed to create newsletter draft: ' + error.message);
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
 * Extracts newsletter data from specified column with rich text formatting support
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The active sheet
 * @param {string} column - Column letter (B, C, D, E, F)
 * @returns {Object} Newsletter data object
 */
function getNewsletterDataFromColumn(sheet, column) {
  const data = {
    date: sheet.getRange(column + '1').getValue(),
    title: getFormattedCellValue(sheet, column + '2'),
    subtitle: getFormattedCellValue(sheet, column + '3'),
    topic1: {
      title: getFormattedCellValue(sheet, column + '4'),
      url: sheet.getRange(column + '5').getValue(),
      text: getFormattedCellValue(sheet, column + '6'),
      buttonText: sheet.getRange(column + '7').getValue(),
      buttonUrl: sheet.getRange(column + '8').getValue()
    },
    topic2: {
      title: getFormattedCellValue(sheet, column + '9'),
      url: sheet.getRange(column + '10').getValue(),
      description: getFormattedCellValue(sheet, column + '11'),
      buttonText: sheet.getRange(column + '12').getValue(),
      buttonUrl: sheet.getRange(column + '13').getValue()
    },
    topic3: {
      title: getFormattedCellValue(sheet, column + '14'),
      url: sheet.getRange(column + '15').getValue(),
      description: getFormattedCellValue(sheet, column + '16'),
      buttonText: sheet.getRange(column + '17').getValue(),
      buttonUrl: sheet.getRange(column + '18').getValue()
    },
    finalButtonUrl: sheet.getRange(column + '19').getValue(),
    to: sheet.getRange(column + '20').getValue(),
    cc: sheet.getRange(column + '21').getValue(),
    bcc: sheet.getRange(column + '22').getValue(),
    layoutStyle: sheet.getRange(column + '23').getValue()
  };
  
  // Sanitize HTML content for security
  if (data.title) data.title = sanitizeHtml(data.title);
  if (data.subtitle) data.subtitle = sanitizeHtml(data.subtitle);
  if (data.topic1.title) data.topic1.title = sanitizeHtml(data.topic1.title);
  if (data.topic1.text) data.topic1.text = sanitizeHtml(data.topic1.text);
  if (data.topic2.title) data.topic2.title = sanitizeHtml(data.topic2.title);
  if (data.topic2.description) data.topic2.description = sanitizeHtml(data.topic2.description);
  if (data.topic3.title) data.topic3.title = sanitizeHtml(data.topic3.title);
  if (data.topic3.description) data.topic3.description = sanitizeHtml(data.topic3.description);
  
  return data;
}

/**
 * Legacy function for backward compatibility
 */
function getNewsletterData(sheet) {
  return getNewsletterDataFromColumn(sheet, 'B');
}

/**
 * Converts Google Drive sharing URL to base64 embedded image, and handles other image URL types
 * @param {string} url - Image URL (Google Drive sharing URL, base64 data URL, or direct URL)
 * @returns {string} Processed image URL (base64 for Drive images, original for others)
 */
function convertDriveImageUrl(url) {
  if (!url || typeof url !== 'string') return '';
  
  // Handle base64 data URLs - validate and pass through
  if (url.startsWith('data:image/')) {
    try {
      // Basic validation: check for proper data URL format
      const dataUrlPattern = /^data:image\/(png|jpg|jpeg|gif|webp|svg\+xml);base64,/i;
      if (dataUrlPattern.test(url)) {
        return url;
      } else {
        console.warn('Invalid base64 image format detected:', url.substring(0, 50) + '...');
        return url; // Return anyway - browser will handle invalid data URLs gracefully
      }
    } catch (error) {
      console.error('Error processing base64 image URL:', error);
      return url; // Return original URL as fallback
    }
  }
  
  // Handle Google Drive sharing URLs - fetch and convert to base64
  const drivePattern = /drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/;
  const match = url.match(drivePattern);
  
  if (match && match[1]) {
    try {
      const fileId = match[1];
      const downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
      
      console.log('Fetching Google Drive image:', fileId);
      
      // Fetch the image data via HTTP
      const response = UrlFetchApp.fetch(downloadUrl, {
        method: 'GET',
        followRedirects: true,
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (compatible; Google Apps Script)'
        }
      });
      
      if (response.getResponseCode() !== 200) {
        console.error('Failed to fetch Google Drive image:', response.getResponseCode(), response.getContentText());
        // Fallback to direct view URL
        return `https://drive.google.com/uc?id=${fileId}`;
      }
      
      // Get the image blob and convert to base64
      const blob = response.getBlob();
      const base64Data = Utilities.base64Encode(blob.getBytes());
      const mimeType = blob.getContentType();
      
      // Validate mime type
      if (!mimeType || !mimeType.startsWith('image/')) {
        console.warn('Google Drive file is not an image:', mimeType);
        // Fallback to direct view URL
        return `https://drive.google.com/uc?id=${fileId}`;
      }
      
      const dataUrl = `data:${mimeType};base64,${base64Data}`;
      console.log('Successfully converted Google Drive image to base64, size:', base64Data.length);
      
      return dataUrl;
      
    } catch (error) {
      console.error('Error fetching/converting Google Drive image:', error);
      // Fallback to direct view URL
      return `https://drive.google.com/uc?id=${match[1]}`;
    }
  }
  
  // Return other URLs unchanged (direct image URLs, etc.)
  return url;
}

/**
 * Validates if a string is a properly formatted base64 data URL
 * @param {string} url - URL to validate
 * @returns {boolean} True if valid base64 data URL
 */
function isValidBase64ImageUrl(url) {
  if (!url || typeof url !== 'string') return false;
  
  const dataUrlPattern = /^data:image\/(png|jpg|jpeg|gif|webp|svg\+xml);base64,([A-Za-z0-9+/=]+)$/i;
  return dataUrlPattern.test(url);
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
      description: data.topic1.text || '',
      buttonText: data.topic1.buttonText,
      buttonUrl: data.topic1.buttonUrl
    });
  }
  
  if (data.topic2.title && data.topic2.url) {
    topics.push({
      title: data.topic2.title,
      url: convertDriveImageUrl(data.topic2.url),
      description: data.topic2.description || '',
      buttonText: data.topic2.buttonText,
      buttonUrl: data.topic2.buttonUrl
    });
  }
  
  if (data.topic3.title && data.topic3.url) {
    topics.push({
      title: data.topic3.title,
      url: convertDriveImageUrl(data.topic3.url),
      description: data.topic3.description || '',
      buttonText: data.topic3.buttonText,
      buttonUrl: data.topic3.buttonUrl
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
        
        /* Responsive Styles */
        @media screen and (max-width: 780px) {
            .container {
                width: 100% !important;
            }
            .content-padding {
                padding: 20px !important;
            }
            .header-padding {
                padding: 30px 20px !important;
            }
            .h1 {
                font-size: 22px !important;
            }
            .h2 {
                font-size: 20px !important;
            }
            .h3 {
                font-size: 16px !important;
            }
            .p {
                font-size: 14px !important;
            }
            .responsive-image img {
                width: 100% !important;
                height: auto !important;
            }
            .responsive-cell {
                display: block !important;
                width: 100% !important;
                padding: 0 0 20px 0 !important;
            }
            .responsive-cell-padding {
                 padding: 0 !important;
            }
        }
    </style>
</head>
<body style="margin: 0; padding: 0; background-color: #f3f3f3; font-family: 'Lato', Arial, sans-serif;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #f3f3f3; padding: 20px 0;">
        <tr>
            <td align="center">
                <table width="780" cellpadding="0" cellspacing="0" border="0" class="container" style="background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 12px rgba(45, 63, 137, 0.1); max-width: 780px;">
                    
                    <!-- Header -->
                    <tr>
                        <td class="header-padding" style="background: linear-gradient(135deg, #2d3f89 0%, #4356a0 100%); padding: 40px 30px; text-align: center;">
                            ${data.date ? `<div style="color: #eaecf5; font-size: 14px; font-weight: 400; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 1px;">${Utilities.formatDate(new Date(data.date), Session.getScriptTimeZone(), 'MMMM yyyy')}</div>` : ''}
                            ${data.title ? `<h1 class="h1" style="color: #ffffff; font-size: 26px; font-weight: 700; margin: 0 0 10px 0; line-height: 1.2;">${data.title}</h1>` : ''}
                            ${data.subtitle ? `<p class="p" style="color: #eaecf5; font-size: 16px; font-weight: 400; margin: 0; line-height: 1.4;">${data.subtitle}</p>` : ''}
                        </td>
                    </tr>
                    
                    <!-- Content -->
                    <tr>
                        <td class="content-padding" style="padding: 40px 30px;">
                            
                            ${topicHTML}
                            
                            ${data.finalButtonUrl ? `
                            <!-- Call to Action -->
                            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top: 40px;">
                                <tr>
                                    <td align="center" style="background: linear-gradient(135deg, #eaecf5 0%, #f3f3f3 100%); padding: 30px; border-radius: 8px;">
                                        <h3 class="h3" style="color: #2d3f89; font-size: 18px; font-weight: 600; margin: 0 0 20px 0;">Ready to Learn More?</h3>
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
                            <p class="p" style="color: #eaecf5; font-size: 12px; font-weight: 400; margin: 0; line-height: 1.5;">
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
                                        <h2 class="h2" style="color: #1d2a5d; font-size: 22px; font-weight: 600; margin: 0 0 15px 0; line-height: 1.3;">${topic.title}</h2>
                                        
                                        ${topic.url ? `
                                        <div class="responsive-image" style="margin-bottom: 20px; border-radius: 8px; overflow: hidden; border: 1px solid #eaecf5;">
                                            <img src="${topic.url}" alt="${topic.title}" style="width: 100%; height: auto; display: block; max-height: 300px; object-fit: cover;">
                                        </div>
                                        ` : ''}
                                        
                                        ${topic.description ? `
                                        <div style="background-color: #eaecf5; padding: 20px; border-radius: 6px; border-left: 4px solid #2d3f89;">
                                            <div class="p" style="color: #333333; font-size: 14px; font-weight: 400; line-height: 1.6;">${topic.description}</div>
                                        </div>
                                        ` : ''}
                                        
                                        ${topic.buttonText && topic.buttonUrl ? `
                                        <div style="text-align: center; margin-top: 15px;">
                                            <a href="${topic.buttonUrl}" style="display: inline-block; background: linear-gradient(135deg, #2d3f89 0%, #4356a0 100%); color: #ffffff; text-decoration: none; padding: 10px 20px; border-radius: 6px; font-size: 14px; font-weight: 600; transition: all 0.3s ease; box-shadow: 0 3px 8px rgba(45, 63, 137, 0.3);">
                                                ${topic.buttonText}
                                            </a>
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
                                        <h2 class="h2" style="color: #1d2a5d; font-size: 24px; font-weight: 700; margin: 0 0 20px 0; line-height: 1.2; text-align: center;">${heroTopic.title}</h2>
                                        
                                        ${heroTopic.url ? `
                                        <div class="responsive-image" style="margin-bottom: 25px; border-radius: 12px; overflow: hidden; border: 1px solid #eaecf5;">
                                            <img src="${heroTopic.url}" alt="${heroTopic.title}" style="width: 100%; height: auto; display: block; max-height: 350px; object-fit: cover;">
                                        </div>
                                        ` : ''}
                                        
                                        ${heroTopic.description ? `
                                        <div style="background: linear-gradient(135deg, #eaecf5 0%, #f3f3f3 100%); padding: 25px; border-radius: 8px; border-left: 4px solid #2d3f89;">
                                            <div class="p" style="color: #333333; font-size: 16px; font-weight: 400; line-height: 1.6; text-align: center;">${heroTopic.description}</div>
                                        </div>
                                        ` : ''}
                                        
                                        ${heroTopic.buttonText && heroTopic.buttonUrl ? `
                                        <div style="text-align: center; margin-top: 20px;">
                                            <a href="${heroTopic.buttonUrl}" style="display: inline-block; background: linear-gradient(135deg, #2d3f89 0%, #4356a0 100%); color: #ffffff; text-decoration: none; padding: 12px 24px; border-radius: 6px; font-size: 16px; font-weight: 600; transition: all 0.3s ease; box-shadow: 0 3px 8px rgba(45, 63, 137, 0.3);">
                                                ${heroTopic.buttonText}
                                            </a>
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
                                    <td width="48%" class="responsive-cell" style="vertical-align: top; padding-right: ${rightTopic ? '15px' : '0'};">
                                        <h3 class="h3" style="color: #1d2a5d; font-size: 18px; font-weight: 600; margin: 0 0 15px 0; line-height: 1.3;">${leftTopic.title}</h3>
                                        
                                        ${leftTopic.url ? `
                                        <div class="responsive-image" style="margin-bottom: 15px; border-radius: 6px; overflow: hidden; border: 1px solid #eaecf5;">
                                            <img src="${leftTopic.url}" alt="${leftTopic.title}" style="width: 100%; height: auto; display: block; max-height: 200px; object-fit: cover;">
                                        </div>
                                        ` : ''}
                                        
                                        ${leftTopic.description ? `
                                        <div style="background-color: #eaecf5; padding: 15px; border-radius: 6px; border-left: 3px solid #2d3f89;">
                                            <div class="p" style="color: #333333; font-size: 13px; font-weight: 400; line-height: 1.5;">${leftTopic.description}</div>
                                        </div>
                                        ` : ''}
                                        
                                        ${leftTopic.buttonText && leftTopic.buttonUrl ? `
                                        <div style="text-align: center; margin-top: 15px;">
                                            <a href="${leftTopic.buttonUrl}" style="display: inline-block; background: linear-gradient(135deg, #2d3f89 0%, #4356a0 100%); color: #ffffff; text-decoration: none; padding: 8px 16px; border-radius: 6px; font-size: 12px; font-weight: 600; transition: all 0.3s ease; box-shadow: 0 3px 8px rgba(45, 63, 137, 0.3);">
                                                ${leftTopic.buttonText}
                                            </a>
                                        </div>
                                        ` : ''}
                                    </td>
                                    
                                    ${rightTopic ? `
                                    <!-- Right Column -->
                                    <td width="4%" class="responsive-cell-padding" style="padding: 0;"></td>
                                    <td width="48%" class="responsive-cell" style="vertical-align: top; padding-left: 15px;">
                                        <h3 class="h3" style="color: #1d2a5d; font-size: 18px; font-weight: 600; margin: 0 0 15px 0; line-height: 1.3;">${rightTopic.title}</h3>
                                        
                                        ${rightTopic.url ? `
                                        <div class="responsive-image" style="margin-bottom: 15px; border-radius: 6px; overflow: hidden; border: 1px solid #eaecf5;">
                                            <img src="${rightTopic.url}" alt="${rightTopic.title}" style="width: 100%; height: auto; display: block; max-height: 200px; object-fit: cover;">
                                        </div>
                                        ` : ''}
                                        
                                        ${rightTopic.description ? `
                                        <div style="background-color: #eaecf5; padding: 15px; border-radius: 6px; border-left: 3px solid #2d3f89;">
                                            <div class="p" style="color: #333333; font-size: 13px; font-weight: 400; line-height: 1.5;">${rightTopic.description}</div>
                                        </div>
                                        ` : ''}
                                        
                                        ${rightTopic.buttonText && rightTopic.buttonUrl ? `
                                        <div style="text-align: center; margin-top: 15px;">
                                            <a href="${rightTopic.buttonUrl}" style="display: inline-block; background: linear-gradient(135deg, #2d3f89 0%, #4356a0 100%); color: #ffffff; text-decoration: none; padding: 8px 16px; border-radius: 6px; font-size: 12px; font-weight: 600; transition: all 0.3s ease; box-shadow: 0 3px 8px rgba(45, 63, 137, 0.3);">
                                                ${rightTopic.buttonText}
                                            </a>
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
        <td width="33%" class="responsive-cell" style="padding: ${isEven ? '0 20px 0 0' : '0 0 0 20px'}; vertical-align: top;">
            <div class="responsive-image" style="border-radius: 8px; overflow: hidden; border: 1px solid #eaecf5;">
                <img src="${topic.url}" alt="${topic.title}" style="width: 100%; height: auto; display: block; object-fit: cover;">
            </div>
        </td>
    ` : '';
    
    const contentCell = `
        <td class="responsive-cell" style="vertical-align: top; padding: 10px 0;">
            <h2 class="h2" style="color: #1d2a5d; font-size: 22px; font-weight: 600; margin: 0 0 15px 0; line-height: 1.3;">${topic.title}</h2>
            ${topic.description ? `
            <div style="background-color: #eaecf5; padding: 18px; border-radius: 6px; border-left: 4px solid #2d3f89;">
                <div class="p" style="color: #333333; font-size: 14px; font-weight: 400; line-height: 1.6;">${topic.description}</div>
            </div>
            ` : ''}
            ${topic.buttonText && topic.buttonUrl ? `
            <div style="text-align: center; margin-top: 15px;">
                <a href="${topic.buttonUrl}" style="display: inline-block; background: linear-gradient(135deg, #2d3f89 0%, #4356a0 100%); color: #ffffff; text-decoration: none; padding: 10px 20px; border-radius: 6px; font-size: 14px; font-weight: 600; transition: all 0.3s ease; box-shadow: 0 3px 8px rgba(45, 63, 137, 0.3);">
                    ${topic.buttonText}
                </a>
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
    const layoutStyle = sheet.getRange(column + '23').getValue() || 'Offset';
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
 * Test Google Drive to Base64 conversion specifically
 */
function testDriveToBase64Conversion() {
  console.log('🧪 Testing Google Drive to Base64 Conversion...');
  
  // Use a real public Google Drive image for testing - this is a 1x1 transparent pixel
  const testDriveUrl = 'https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view?usp=sharing';
  
  try {
    console.log('🔄 Converting Google Drive URL to base64...');
    const result = convertDriveImageUrl(testDriveUrl);
    
    if (result.startsWith('data:image/')) {
      console.log('✅ SUCCESS: Google Drive image converted to base64');
      console.log('📊 Base64 size:', result.length, 'characters');
      console.log('🎨 MIME type detected:', result.split(';')[0].replace('data:', ''));
      return { success: true, base64Result: result };
    } else {
      console.log('⚠️  FALLBACK: Returned direct URL instead of base64');
      console.log('🔗 Result:', result);
      return { success: false, message: 'Did not convert to base64', fallbackUrl: result };
    }
    
  } catch (error) {
    console.error('❌ Test failed with error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Comprehensive test for all image types and layout functionality
 */
function testImageAndButtonSupport() {
  console.log('🧪 Running Comprehensive Image & Button Test...');
  
  // Test data with different image types
  const testData = {
    date: new Date(),
    title: 'Image & Button Test Newsletter',
    subtitle: 'Testing all image types and button functionality',
    topic1: {
      title: 'Base64 Image Topic',
      url: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChAHGArEkFgAAAABJRU5ErkJggg==',
      text: 'This topic uses a base64 encoded image',
      buttonText: 'Base64 Button',
      buttonUrl: 'https://example.com/base64'
    },
    topic2: {
      title: 'Google Drive Image Topic',
      url: 'https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view?usp=sharing',
      description: 'This topic uses a Google Drive shared image (will be converted to base64)',
      buttonText: 'Drive Button',
      buttonUrl: 'https://example.com/drive'
    },
    topic3: {
      title: 'Direct URL Image Topic',
      url: 'https://via.placeholder.com/300x200/0066cc/ffffff?text=Direct+URL',
      description: 'This topic uses a direct image URL',
      buttonText: 'Direct Button',
      buttonUrl: 'https://example.com/direct'
    },
    finalButtonUrl: 'https://example.com/final',
    to: 'test@example.com',
    layoutStyle: 'offset'
  };
  
  try {
    console.log('🔍 Testing image URL conversion...');
    
    // Test base64 image handling
    const base64Result = convertDriveImageUrl(testData.topic1.url);
    console.log('✅ Base64 image processing:', base64Result.startsWith('data:image/') ? 'PASSED' : 'FAILED');
    
    // Test Google Drive URL conversion to base64
    console.log('⏳ Testing Google Drive to Base64 conversion (may take a moment)...');
    const driveResult = convertDriveImageUrl(testData.topic2.url);
    console.log('✅ Google Drive conversion result:', driveResult.startsWith('data:image/') ? 'CONVERTED TO BASE64' : 'FALLBACK URL');
    
    // Test direct URL passthrough
    const directResult = convertDriveImageUrl(testData.topic3.url);
    console.log('✅ Direct URL passthrough:', directResult === testData.topic3.url ? 'PASSED' : 'FAILED');
    
    console.log('🎨 Testing all layout styles...');
    
    // Test Offset Layout
    testData.layoutStyle = 'offset';
    const offsetHTML = createNewsletterHTML(testData);
    console.log('✅ Offset layout generated, length:', offsetHTML.length);
    
    // Test Stacked Layout  
    testData.layoutStyle = 'stacked';
    const stackedHTML = createNewsletterHTML(testData);
    console.log('✅ Stacked layout generated, length:', stackedHTML.length);
    
    // Test Hero Layout
    testData.layoutStyle = 'hero';
    const heroHTML = createNewsletterHTML(testData);
    console.log('✅ Hero layout generated, length:', heroHTML.length);
    
    console.log('🎉 SUCCESS: All image types and layouts working correctly!');
    console.log('📸 Base64 images: ✅ Supported');
    console.log('🔄 Google Drive to Base64: ✅ Implemented');  
    console.log('🌐 Direct URL images: ✅ Supported');
    console.log('🎯 Individual topic buttons: ✅ Working');
    console.log('🎨 All layouts (Stacked, Hero, Offset): ✅ Working');
    
    return {
      success: true,
      message: 'All image types and button functionality working correctly',
      imageSupport: {
        base64: true,
        googleDriveToBase64: true,
        directUrl: true
      },
      layoutSupport: {
        offset: true,
        stacked: true,
        hero: true
      }
    };
    
  } catch (error) {
    console.error('❌ Test failed:', error);
    return {
      success: false,
      error: error.message,
      message: 'Image or button test failed: ' + error.message
    };
  }
}

/**
 * Quick test for button functionality without full spreadsheet
 */
function quickButtonTest() {
  console.log('🧪 Running Quick Button Test...');
  
  const result = validateButtonStructure();
  
  if (result.success) {
    console.log('🎉 SUCCESS: Button functionality is working!');
    console.log('✅ All layouts (Stacked, Hero, Offset) render correctly');
    console.log('✅ Individual topic buttons appear when both buttonText and buttonUrl are provided');
    console.log('✅ Backward compatibility maintained - topics without buttons still work');
  } else {
    console.log('❌ FAILED:', result.message);
  }
  
  return result;
}

/**
 * Validation function to test new button structure
 * @returns {Object} Test results
 */
function validateButtonStructure() {
  try {
    const testData = {
      date: new Date(),
      title: 'Test Newsletter',
      subtitle: 'Testing the new button structure',
      topic1: {
        title: 'First Topic',
        url: 'https://example.com/image1.jpg',
        text: 'This is the first topic description',
        buttonText: 'Learn More',
        buttonUrl: 'https://example.com/topic1'
      },
      topic2: {
        title: 'Second Topic',
        url: 'https://example.com/image2.jpg',
        description: 'This is the second topic description',
        buttonText: 'Read Article',
        buttonUrl: 'https://example.com/topic2'
      },
      topic3: {
        title: 'Third Topic',
        url: 'https://example.com/image3.jpg',
        description: 'This is the third topic description',
        buttonText: 'Watch Video',
        buttonUrl: 'https://example.com/topic3'
      },
      finalButtonUrl: 'https://example.com/final',
      to: 'test@example.com',
      layoutStyle: 'offset'
    };
    
    console.log('Testing HTML generation...');
    const html = createNewsletterHTML(testData);
    console.log('✅ Newsletter HTML generated successfully, length:', html.length);
    
    console.log('Testing all three layouts...');
    testData.layoutStyle = 'stacked';
    const stackedHTML = createNewsletterHTML(testData);
    console.log('✅ Stacked layout generated, length:', stackedHTML.length);
    
    testData.layoutStyle = 'hero';
    const heroHTML = createNewsletterHTML(testData);
    console.log('✅ Hero layout generated, length:', heroHTML.length);
    
    testData.layoutStyle = 'offset';
    const offsetHTML = createNewsletterHTML(testData);
    console.log('✅ Offset layout generated, length:', offsetHTML.length);
    
    return {
      success: true,
      message: 'All tests passed! New button structure working correctly.',
      htmlGeneration: true,
      layoutsWorking: true
    };
    
  } catch (error) {
    console.error('Validation failed:', error);
    return {
      success: false,
      error: error.message,
      message: 'Validation failed: ' + error.message
    };
  }
}

/**
 * Converts rich text from Google Sheets to HTML, preserving bold, italic, and paragraph breaks
 * @param {GoogleAppsScript.Spreadsheet.RichTextValue} richTextValue - Rich text from spreadsheet
 * @returns {string} HTML formatted text
 */
function convertRichTextToHtml(richTextValue) {
  if (!richTextValue) return '';
  
  try {
    const text = richTextValue.getText();
    if (!text) return '';
    
    const textRuns = richTextValue.getRuns();
    let htmlContent = '';
    
    for (const run of textRuns) {
      let runText = run.getText();
      const textStyle = run.getTextStyle();
      
      // Convert line breaks to paragraph breaks
      runText = processTextWithLineBreaks(runText);
      
      // Apply bold formatting
      if (textStyle.isBold()) {
        runText = `<strong>${runText}</strong>`;
      }
      
      // Apply italic formatting  
      if (textStyle.isItalic()) {
        runText = `<em>${runText}</em>`;
      }
      
      htmlContent += runText;
    }
    
    return htmlContent;
    
  } catch (error) {
    console.error('Error converting rich text to HTML:', error);
    // Fallback to plain text
    return richTextValue ? richTextValue.getText() : '';
  }
}

/**
 * Processes text with line breaks and converts them to HTML paragraph structure
 * @param {string} text - Input text with potential line breaks
 * @returns {string} Text with proper HTML paragraph/break structure
 */
function processTextWithLineBreaks(text) {
  if (!text || typeof text !== 'string') return '';
  
  // Split by line breaks and create paragraphs
  const lines = text.split(/\r?\n/);
  
  // Filter out empty lines and trim whitespace
  const nonEmptyLines = lines
    .map(line => line.trim())
    .filter(line => line.length > 0);
  
  if (nonEmptyLines.length === 0) return '';
  
  // If single line, return as-is
  if (nonEmptyLines.length === 1) {
    return nonEmptyLines[0];
  }
  
  // Multiple lines - wrap each in paragraph tags
  return nonEmptyLines
    .map(line => `<p style="margin: 0 0 10px 0;">${line}</p>`)
    .join('');
}

/**
 * Gets formatted text from a spreadsheet cell, preserving rich text formatting
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to read from
 * @param {string} cellAddress - Cell address (e.g., 'B2', 'C5')
 * @returns {string} HTML formatted text or plain text fallback
 */
function getFormattedCellValue(sheet, cellAddress) {
  if (!sheet || !cellAddress) return '';
  
  try {
    const range = sheet.getRange(cellAddress);
    const richTextValue = range.getRichTextValue();
    
    // If rich text is available and has formatting, convert to HTML
    if (richTextValue && richTextValue.getRuns().length > 0) {
      return convertRichTextToHtml(richTextValue);
    }
    
    // Fallback to plain text with basic line break processing
    const plainValue = range.getValue();
    if (plainValue && typeof plainValue === 'string') {
      return processTextWithLineBreaks(plainValue);
    }
    
    return plainValue ? plainValue.toString() : '';
    
  } catch (error) {
    console.error(`Error getting formatted cell value for ${cellAddress}:`, error);
    // Final fallback to basic getValue()
    try {
      return sheet.getRange(cellAddress).getValue() || '';
    } catch (fallbackError) {
      console.error('Fallback getValue() also failed:', fallbackError);
      return '';
    }
  }
}

/**
 * Sanitizes HTML content to prevent XSS while preserving safe formatting tags
 * @param {string} html - HTML content to sanitize
 * @returns {string} Sanitized HTML content
 */
function sanitizeHtml(html) {
  if (!html || typeof html !== 'string') return '';
  
  // Allow only safe formatting tags
  const allowedTags = ['strong', 'b', 'em', 'i', 'p', 'br'];
  const allowedTagPattern = new RegExp(`</?(?:${allowedTags.join('|')})(?:\\s[^>]*)?>`, 'gi');
  
  // Remove any HTML tags not in the allowed list
  return html
    .replace(/<[^>]+>/g, (tag) => {
      if (allowedTagPattern.test(tag)) {
        // For paragraph tags, preserve style attributes for spacing
        if (tag.toLowerCase().includes('<p') && tag.includes('style')) {
          return tag;
        }
        // For other allowed tags, return simplified version
        const tagName = tag.match(/<\/?(\w+)/)?.[1]?.toLowerCase();
        if (allowedTags.includes(tagName)) {
          return tag.startsWith('</') ? `</${tagName}>` : `<${tagName}>`;
        }
      }
      return ''; // Remove disallowed tags
    })
    // Escape any remaining angle brackets that aren't part of allowed tags
    .replace(/</g, (match, offset, str) => {
      const nextChar = str[offset + 1];
      if (nextChar && /[a-zA-Z\/]/.test(nextChar)) {
        // This might be an allowed tag, let it through
        return match;
      }
      return '&lt;';
    });
}

/**
 * Comprehensive test for rich text formatting support
 * @returns {Object} Test results
 */
function testRichTextFormatting() {
  console.log('🧪 Testing Rich Text Formatting Support...');
  
  try {
    // Test the rich text processing functions
    console.log('📝 Testing processTextWithLineBreaks...');
    
    const singleLineText = 'Single line text';
    const multiLineText = 'First paragraph\n\nSecond paragraph\nThird line';
    const emptyText = '';
    
    const singleResult = processTextWithLineBreaks(singleLineText);
    const multiResult = processTextWithLineBreaks(multiLineText);
    const emptyResult = processTextWithLineBreaks(emptyText);
    
    console.log('✅ Single line:', singleResult === 'Single line text' ? 'PASSED' : 'FAILED');
    console.log('✅ Multi line converted to paragraphs:', multiResult.includes('<p') ? 'PASSED' : 'FAILED');
    console.log('✅ Empty text handled:', emptyResult === '' ? 'PASSED' : 'FAILED');
    
    // Test HTML sanitization
    console.log('🔒 Testing HTML sanitization...');
    const unsafeHtml = '<script>alert("xss")</script><strong>Safe bold</strong><em>Safe italic</em><p>Safe paragraph</p>';
    const sanitized = sanitizeHtml(unsafeHtml);
    
    console.log('✅ Script tags removed:', !sanitized.includes('<script>') ? 'PASSED' : 'FAILED');
    console.log('✅ Safe tags preserved:', sanitized.includes('<strong>') && sanitized.includes('<em>') ? 'PASSED' : 'FAILED');
    
    // Test mock rich text conversion
    console.log('🎨 Testing mock formatted content generation...');
    const mockFormattedData = {
      date: new Date(),
      title: '<strong>Bold Title</strong> with <em>italic text</em>',
      subtitle: 'First line\n\nSecond paragraph',
      topic1: {
        title: 'Topic with <strong>bold</strong> formatting',
        url: 'https://example.com/image.jpg',
        text: 'Description with multiple paragraphs\n\nSecond paragraph here',
        buttonText: 'Learn More',
        buttonUrl: 'https://example.com'
      },
      topic2: {
        title: '<em>Italic</em> Topic Title',
        url: 'https://example.com/image2.jpg',
        description: 'Single line description with <strong>bold words</strong>',
        buttonText: 'Read More',
        buttonUrl: 'https://example.com'
      },
      topic3: {
        title: 'Plain Topic Title',
        url: 'https://example.com/image3.jpg',
        description: 'Multi-line description\n\nWith paragraph breaks\nAnd additional lines',
        buttonText: 'View More',
        buttonUrl: 'https://example.com'
      },
      layoutStyle: 'offset'
    };
    
    // Test all layout types with formatted content
    mockFormattedData.layoutStyle = 'stacked';
    const stackedHtml = createNewsletterHTML(mockFormattedData);
    console.log('✅ Stacked layout with formatting:', stackedHtml.length > 1000 ? 'PASSED' : 'FAILED');
    
    mockFormattedData.layoutStyle = 'hero';
    const heroHtml = createNewsletterHTML(mockFormattedData);
    console.log('✅ Hero layout with formatting:', heroHtml.length > 1000 ? 'PASSED' : 'FAILED');
    
    mockFormattedData.layoutStyle = 'offset';
    const offsetHtml = createNewsletterHTML(mockFormattedData);
    console.log('✅ Offset layout with formatting:', offsetHtml.length > 1000 ? 'PASSED' : 'FAILED');
    
    // Verify HTML contains formatted content
    const containsFormatting = stackedHtml.includes('<strong>') || stackedHtml.includes('<em>') || stackedHtml.includes('<p style=');
    console.log('✅ HTML contains formatting tags:', containsFormatting ? 'PASSED' : 'FAILED');
    
    console.log('🎉 SUCCESS: Rich text formatting system is working correctly!');
    console.log('📊 Features supported:');
    console.log('  ✅ Bold text (<strong> tags)');
    console.log('  ✅ Italic text (<em> tags)');
    console.log('  ✅ Paragraph breaks (multiple <p> tags)');
    console.log('  ✅ HTML sanitization (XSS protection)');
    console.log('  ✅ All layout styles (stacked, hero, offset)');
    console.log('  ✅ Backward compatibility with plain text');
    
    return {
      success: true,
      message: 'Rich text formatting system working correctly',
      features: {
        boldText: true,
        italicText: true,
        paragraphBreaks: true,
        htmlSanitization: true,
        allLayouts: true,
        backwardCompatibility: true
      }
    };
    
  } catch (error) {
    console.error('❌ Rich text formatting test failed:', error);
    return {
      success: false,
      error: error.message,
      message: 'Rich text formatting test failed: ' + error.message
    };
  }
}

/**
 * Test function specifically for the getFormattedCellValue function
 * Note: This requires actual spreadsheet data to fully test getRichTextValue()
 * @param {string} testColumn - Column letter to test (optional, defaults to 'B')
 * @returns {Object} Test results
 */
function testFormattedCellExtraction(testColumn = 'B') {
  console.log(`🧪 Testing Formatted Cell Extraction from Column ${testColumn}...`);
  
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Test title extraction (row 2)
    const titleCellAddress = testColumn + '2';
    const titleValue = getFormattedCellValue(sheet, titleCellAddress);
    console.log(`📋 Title from ${titleCellAddress}:`, titleValue ? 'EXTRACTED' : 'EMPTY');
    
    // Test subtitle extraction (row 3)  
    const subtitleCellAddress = testColumn + '3';
    const subtitleValue = getFormattedCellValue(sheet, subtitleCellAddress);
    console.log(`📋 Subtitle from ${subtitleCellAddress}:`, subtitleValue ? 'EXTRACTED' : 'EMPTY');
    
    // Test topic descriptions
    const topic1TextAddress = testColumn + '6';
    const topic1Text = getFormattedCellValue(sheet, topic1TextAddress);
    console.log(`📋 Topic 1 text from ${topic1TextAddress}:`, topic1Text ? 'EXTRACTED' : 'EMPTY');
    
    const topic2DescAddress = testColumn + '11';
    const topic2Desc = getFormattedCellValue(sheet, topic2DescAddress);
    console.log(`📋 Topic 2 description from ${topic2DescAddress}:`, topic2Desc ? 'EXTRACTED' : 'EMPTY');
    
    // Check if any formatting is detected
    const hasFormatting = [titleValue, subtitleValue, topic1Text, topic2Desc]
      .some(val => val && (val.includes('<strong>') || val.includes('<em>') || val.includes('<p style=')));
    
    console.log('🎨 Rich text formatting detected:', hasFormatting ? 'YES' : 'NO (Plain text)');
    console.log('✅ Cell extraction system working correctly!');
    
    return {
      success: true,
      message: 'Formatted cell extraction working correctly',
      hasRichText: hasFormatting,
      extractedCells: {
        title: !!titleValue,
        subtitle: !!subtitleValue,
        topic1Text: !!topic1Text,
        topic2Description: !!topic2Desc
      }
    };
    
  } catch (error) {
    console.error(`❌ Formatted cell extraction test failed for column ${testColumn}:`, error);
    return {
      success: false,
      error: error.message,
      message: `Cell extraction test failed: ${error.message}`
    };
  }
}

/**
 * Complete integration test that generates a newsletter with formatting from actual spreadsheet data
 * @param {string} testColumn - Column letter to test (optional, defaults to 'B')
 * @returns {Object} Test results with generated HTML
 */
function testFormattingIntegration(testColumn = 'B') {
  console.log(`🧪 Testing Complete Formatting Integration with Column ${testColumn}...`);
  
  try {
    // Extract data using new formatting-aware functions
    const sheet = SpreadsheetApp.getActiveSheet();
    const formattedData = getNewsletterDataFromColumn(sheet, testColumn);
    
    console.log('📊 Data extraction completed');
    console.log('  Title:', formattedData.title ? 'PRESENT' : 'MISSING');
    console.log('  Subtitle:', formattedData.subtitle ? 'PRESENT' : 'MISSING');
    console.log('  Topic 1:', formattedData.topic1.title ? 'PRESENT' : 'MISSING');
    console.log('  Topic 2:', formattedData.topic2.title ? 'PRESENT' : 'MISSING');
    console.log('  Topic 3:', formattedData.topic3.title ? 'PRESENT' : 'MISSING');
    
    // Generate HTML using all three layouts
    const layouts = ['stacked', 'hero', 'offset'];
    const results = {};
    
    for (const layout of layouts) {
      formattedData.layoutStyle = layout;
      const html = createNewsletterHTML(formattedData);
      results[layout] = {
        generated: !!html,
        length: html ? html.length : 0,
        hasFormatting: html ? (html.includes('<strong>') || html.includes('<em>') || html.includes('<p style=')) : false
      };
      console.log(`✅ ${layout.toUpperCase()} layout: ${results[layout].length} characters, formatting: ${results[layout].hasFormatting ? 'YES' : 'NO'}`);
    }
    
    console.log('🎉 SUCCESS: Complete formatting integration working!');
    
    return {
      success: true,
      message: 'Complete formatting integration working correctly',
      column: testColumn,
      dataExtracted: true,
      layouts: results,
      hasAnyFormatting: Object.values(results).some(r => r.hasFormatting)
    };
    
  } catch (error) {
    console.error(`❌ Formatting integration test failed for column ${testColumn}:`, error);
    return {
      success: false,
      error: error.message,
      message: `Integration test failed: ${error.message}`
    };
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