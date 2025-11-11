# Project Overview

This project is a Google Apps Script for generating and sending professional HTML newsletters. The script is designed to be used with a Google Sheet, where each column represents a different newsletter. The script provides a user-friendly interface within the Google Sheet to send, draft, preview, or generate the HTML for the newsletters.

## Main Technologies

*   **Google Apps Script:** The core of the project is written in JavaScript using the Google Apps Script platform.
*   **Google Sheets:** The data for the newsletters is stored and managed in a Google Sheet.
*   **HTML/CSS:** The newsletters are generated as HTML with inline CSS for maximum compatibility with email clients.
*   **Gmail:** The script uses the Gmail API to send the newsletters.

## Architecture

The script is organized into several functions:

*   `onOpen()`: Creates the custom menu in the Google Sheet when the spreadsheet is opened.
*   `show...Picker()` functions: These functions create the HTML dialogs for the user to select the newsletter to work with.
*   `createColumnPickerDialog()`: This function generates the HTML for the column selection dialog.
*   `generateNewsletterHTMLFromColumn()`: This is the core function that generates the HTML for the newsletter based on the data in the specified column.
*   `sendNewsletterFromColumn()` and `createDraftNewsletterFromColumn()`: These functions handle the sending and drafting of the newsletters.
*   `getNewsletterDataFromColumn()`: This function extracts the data for the newsletter from the Google Sheet.
*   `convertDriveImageUrl()`: This function handles the conversion of Google Drive image URLs to base64 encoded images.
*   `createNewsletterHTML()`: This function constructs the final HTML for the newsletter, including the header, content, and footer.
*   `generate...Layout()` functions: These functions generate the HTML for the different topic layouts (Stacked, Hero, Offset).
*   `getFormattedCellValue()` and `convertRichTextToHtml()`: These functions handle the extraction and conversion of rich text formatting from the Google Sheet.
*   Test functions: The script includes a comprehensive set of test functions to validate the functionality.

# Building and Running

This is a Google Apps Script project, so there is no traditional build process. To run the script, you need to:

1.  Open the Google Sheet that is associated with the script.
2.  A "Newsletter Tools" menu should appear in the menu bar.
3.  Select one of the options from the menu to send, draft, preview, or generate a newsletter.

To run the test functions, you need to open the script editor and run the test functions manually.

# Development Conventions

*   The code is well-documented with JSDoc comments.
*   The script uses a consistent naming convention for functions and variables.
*   The script includes a comprehensive set of test functions to ensure the quality of the code.
*   The script is designed to be easily extensible with new layouts or features.
