AI Bulk Translator for Excel
A modern Office Add-in for Microsoft Excel that bulk-translates selected cell ranges or entire workbooks using the Google Gemini API.

This add-in is designed for users who work with large datasets and need to quickly localize multilingual documents. Thanks to its smart features, it can safely and efficiently translate thousands of rows without hitting the free API limits of Google.

✨ Features
Comprehensive Translation: Translate selected cells or all sheets in a workbook with a single click.

Flexible Output Options: Choose between overwriting the existing data ("Replace in Place") or exporting the translations to a brand new sheet while preserving all formatting ("Translate to New Sheet").

Google AI Support: Allows selection between Gemini Flash (optimized for free-tier usage) and Gemini Pro models.

Smart Batching: An intelligent system that processes large files by considering both cell count and character limits to work efficiently within API constraints.

Resilient Error Handling: Features an "exponential backoff" mechanism that automatically waits and retries when API rate limits (429 errors) are exceeded, complete with a visual countdown timer.

Caching System: Caches translations within the current session to avoid re-querying the API for the same text, increasing speed and preserving API quota.

Modern UI: A clean and user-friendly interface designed with Google Material Design, featuring both light and dark theme support.

Data Safety: Protects your original data by displaying any potential API errors in the status bar instead of writing them into the cells.

Additional Functionality: Includes a feature to translate the name of the active sheet.

🚀 Installation
Method 1: Via Microsoft AppSource (Recommended)
Once our add-in is published on Microsoft AppSource, you can install it with one click by navigating to Insert > Get Add-ins in Excel and searching for "AI Bulk Translator".

Method 2: Sideloading for Development
To test or sideload this add-in for personal use:

Download the manifest.xml file from this repository.

Follow the official  to load the manifest file in Excel.

📋 How to Use
In Excel, navigate to the Home tab and click the "Start Translator" button on the ribbon to open the add-in pane.

Get Your API Key: The add-in requires an API key to use the Google AI API (see section below).

Save the Key: Paste your API key into the designated field in the add-in pane and click "Save Key".

Configure Settings:

AI Model: Choose between the Flash (free tier) or Pro (paid) model.

Target Language: Select the language you want to translate your text into.

Translation Mode: Choose how the translation should be applied ("Replace in Place" or "Translate to New Sheet").

Start Translating:

To translate a specific area, select the cells and click "Translate Selection".

To translate the entire workbook, click "Translate All Sheets".

To translate only the name of the active sheet, click "Translate Active Sheet Name".

You can monitor the progress in the Status section at the bottom of the pane.

🔑 Getting a Google AI API Key
Go to .

Sign in with your Google account.

Click "Create API key in new project" to generate a new key.

Copy the generated key and paste it into the add-in.

🛠️ Technology Stack
Platform: Office Add-ins

Primary Language: JavaScript (ES6+ async/await)

API: Office.js, Google Gemini API

Interface: HTML5, CSS3

📜 MIT License
This project is licensed under the MIT
