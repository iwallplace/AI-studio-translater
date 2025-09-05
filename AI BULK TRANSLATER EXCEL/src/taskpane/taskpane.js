/*
 * AI Translator for Excel - Main Logic
 * Version: 1.0.0
 * Features:
 * - Translates selected range or entire workbook.
 * - Two modes: "Replace in Place" or "Translate to New Sheet".
 * - Supports Gemini Flash (free tier) and Pro models.
 * - Smart batching system based on cell and character count to handle large data.
 * - Session-based caching to avoid re-translating text.
 * - Automatic retry (exponential backoff) for API rate limit errors (429).
 * - Visual countdown timer for rate limit delays.
 * - Robust error handling to protect cell data.
 * - Google Material Design theme with light/dark mode support.
 */

// --- CONSTANTS ---
const STANDARD_MODEL = "gemini-1.5-flash";
const PRO_MODEL = "gemini-1.5-pro";
const BATCH_CELL_LIMIT = 100; // Max number of cells per batch
const BATCH_CHAR_LIMIT = 15000; // Max total characters per batch to keep API request size safe
const SAFE_MODE_DELAY = 1200; // Delay in ms for the Flash model to respect the 60 RPM limit
const EXCEL_CELL_CHAR_LIMIT = 32767; // Max characters allowed in a single Excel cell

// --- CACHE ---
/** @type {Map<string, string>} Session-only cache for translations. Cleared on language change or reload. */
const translationCache = new Map();

/**
 * Office.onReady is called when the Office platform is ready to host the add-in.
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    try {
      // Apply the host theme for a native look and feel.
      const theme = Office.context.officeTheme;
      if (theme && ['#000000', '#2B2B2B', '#3c3c3c'].includes(theme.bodyBackgroundColor)) {
        document.body.classList.add("dark-mode");
      }
      
      // Assign event listeners to all interactive elements.
      document.getElementById("save-key-button").onclick = saveApiKey;
      document.getElementById("edit-key-button").onclick = editApiKey;
      document.getElementById("translate-workbook-button").onclick = runWorkbookTranslation;
      document.getElementById("translate-selection-button").onclick = runSelectionTranslation;
      document.getElementById("translate-sheet-name-button").onclick = runTranslateSheetName;
      document.getElementById("model-selection-select").onchange = updateModelDescription;
      document.getElementById("target-language-select").onchange = clearCache;
      
      // Populate the model selection dropdown.
      const modelSelect = document.getElementById("model-selection-select");
      modelSelect.innerHTML = `
          <option value="${STANDARD_MODEL}">Flash (Standard)</option>
          <option value="${PRO_MODEL}">Pro (Fast)</option>
      `;
      
      // Load initial state.
      loadApiKey();
      updateModelDescription();
    } catch (error) {
      console.error("Initialization error:", error);
      updateStatus("Error during startup.", error.message, null, true);
    }
  }
});

// --- CACHE MANAGEMENT ---

/**
 * Clears the session translation cache. Typically called when the target language changes.
 */
function clearCache() {
    translationCache.clear();
    console.log("Translation cache cleared due to language change.");
}

// --- UI & DATA MANAGEMENT ---

/**
 * Loads the API key from document settings and updates the UI.
 */
function loadApiKey() {
    const apiKey = Office.context.document.settings.get("apiKey");
    if (apiKey && apiKey.trim() !== "") {
        document.getElementById("masked-key").textContent = maskApiKey(apiKey);
        showDisplayMode();
    } else {
        showInputMode();
    }
}

/**
 * Saves the API key from the input field to the current document's settings.
 */
function saveApiKey() {
    const apiKeyInput = document.getElementById("api-key-input");
    const apiKey = apiKeyInput.value.trim();
    if (apiKey) {
        Office.context.document.settings.set("apiKey", apiKey);
        Office.context.document.settings.saveAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                document.getElementById("masked-key").textContent = maskApiKey(apiKey);
                showDisplayMode();
                updateStatus("API Key saved successfully.", null, null, false);
            } else {
                console.error("Failed to save API key:", result.error.message);
                updateStatus("Error: Could not save API key.", result.error.message, null, true);
            }
        });
    } else {
        updateStatus("Please enter a valid API key.", null, null, true);
    }
}

/**
 * Switches the UI to allow editing of the API key.
 */
function editApiKey() {
    document.getElementById("api-key-input").value = Office.context.document.settings.get("apiKey") || "";
    showInputMode();
}

/**
 * Shows the UI section for entering an API key.
 */
function showInputMode() {
    document.getElementById("key-input-section").style.display = "block";
    document.getElementById("key-display-section").style.display = "none";
    document.getElementById("translation-section").style.display = "none";
}

/**
 * Shows the main UI section for translation functions.
 */
function showDisplayMode() {
    document.getElementById("key-input-section").style.display = "none";
    document.getElementById("key-display-section").style.display = "block";
    document.getElementById("translation-section").style.display = "block";
}

/**
 * Masks an API key for display purposes (e.g., "AIza...J8ZU").
 * @param {string} apiKey The API key to mask.
 * @returns {string} The masked key.
 */
function maskApiKey(apiKey) {
    if (apiKey.length > 8) {
        return `${apiKey.substring(0, 4)}...${apiKey.substring(apiKey.length - 4)}`;
    }
    return apiKey;
}

/**
 * Updates the descriptive text below the AI model dropdown based on the current selection.
 */
function updateModelDescription() {
    const model = document.getElementById("model-selection-select").value;
    const descriptionEl = document.getElementById("model-recommendation");
    if (model === STANDARD_MODEL) {
        descriptionEl.innerHTML = "<b>Ideal for free API keys.</b> Due to a 60 requests/minute limit, this add-in uses a smart delay and retry system to translate large files safely.";
    } else {
        descriptionEl.innerHTML = "<b>For users paying for API usage.</b> Offers the highest speed and most powerful translation performance without limits.";
    }
}

/**
 * Updates the status bar with messages, progress, and error states.
 * @param {string} message The main status message.
 * @param {string|null} detail A secondary, smaller message.
 * @param {number|null} progress A percentage from 0 to 100 to display the progress bar.
 * @param {boolean} isError If true, the message is styled as an error.
 */
function updateStatus(message, detail, progress, isError = false) {
    document.getElementById("status-text").innerText = message;
    document.getElementById("status-text").style.color = isError ? "#ff4d4d" : "inherit";
    document.getElementById("status-detail-text").innerText = detail || "";

    const progressIndicator = document.getElementById("progress-indicator");
    const progressBar = document.getElementById("progress-bar");
    const progressPercentage = document.getElementById("progress-percentage");

    if (progress != null) {
        progressIndicator.style.display = "block";
        progressBar.style.width = `${progress}%`;
        progressPercentage.innerText = `${Math.round(progress)}%`;
    } else {
        progressIndicator.style.display = "none";
        progressPercentage.innerText = "";
    }
}

/**
 * A simple promise-based delay function.
 * @param {number} ms Milliseconds to wait.
 */
function sleep(ms) { 
    return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Displays a visual countdown timer in the status bar for a given duration.
 * @param {number} duration The duration of the countdown in milliseconds.
 */
function startCountdownTimer(duration) {
    return new Promise(resolve => {
        let timeLeft = duration;
        const statusDetailEl = document.getElementById("status-detail-text");
        const intervalId = setInterval(() => {
            timeLeft -= 100;
            if (timeLeft <= 0) {
                clearInterval(intervalId);
                statusDetailEl.innerText = "";
                resolve();
            } else {
                statusDetailEl.innerText = `Next request in ${(timeLeft / 1000).toFixed(1)}s...`;
            }
        }, 100);
    });
}

// --- CORE TRANSLATION LOGIC ---

/**
 * The main translation engine. It reads a range, checks the cache, calls the API for new text,
 * processes results, and writes them back to the sheet.
 * @param {Excel.RequestContext} context The request context.
 * @param {Excel.Range} range The Excel range to process.
 * @param {object} options An object containing apiKey, targetLanguage, model, and mode.
 * @returns {Promise<{totalErrors: number, firstErrorMessage: string|null}>} An object with the error count and the first error message.
 */
async function translateRange(context, range, options) {
    updateStatus("Reading data from sheet...", null, 0);
    range.load(["values", "address", "worksheet"]);
    await context.sync();

    const originalValues = range.values;
    const cellsToTranslate = [];
    const uniqueTexts = new Map();

    for (let i = 0; i < originalValues.length; i++) {
        for (let j = 0; j < originalValues[i].length; j++) {
            const cellValue = originalValues[i][j];
            if (typeof cellValue === 'string' && cellValue.trim() !== "") {
                if (!uniqueTexts.has(cellValue)) { uniqueTexts.set(cellValue, null); }
                cellsToTranslate.push({ row: i, col: j, text: cellValue });
            }
        }
    }

    if (cellsToTranslate.length === 0) {
        return { totalErrors: 0, firstErrorMessage: null };
    }

    // Check cache for existing translations
    const uniqueTextArray = Array.from(uniqueTexts.keys());
    const textsToFetchFromApi = [];
    for (const text of uniqueTextArray) {
        if (translationCache.has(text)) {
            uniqueTexts.set(text, translationCache.get(text));
        } else {
            textsToFetchFromApi.push(text);
        }
    }

    const cachedCount = uniqueTextArray.length - textsToFetchFromApi.length;
    if (cachedCount > 0) {
        updateStatus(`Found ${cachedCount} translations in cache.`, "Checking for new text...", 10);
        await sleep(500);
    }
    
    let totalErrors = 0;
    let firstErrorMessage = null;

    if (textsToFetchFromApi.length > 0) {
        // Smart batching based on both cell count and total characters
        const allBatches = [];
        let currentBatch = [];
        let currentCharCount = 0;
        for (const text of textsToFetchFromApi) {
            const textLength = text.length;
            if (textLength > BATCH_CHAR_LIMIT) {
                if(currentBatch.length > 0) { allBatches.push(currentBatch); }
                allBatches.push([text]);
                currentBatch = [];
                currentCharCount = 0;
                continue;
            }
            if (currentBatch.length > 0 && (currentCharCount + textLength > BATCH_CHAR_LIMIT || currentBatch.length >= BATCH_CELL_LIMIT)) {
                allBatches.push(currentBatch);
                currentBatch = [];
                currentCharCount = 0;
            }
            currentBatch.push(text);
            currentCharCount += textLength;
        }
        if (currentBatch.length > 0) { allBatches.push(currentBatch); }

        // Process all batches
        const totalBatches = allBatches.length;
        for (let i = 0; i < totalBatches; i++) {
            const batch = allBatches[i];
            const overallProgress = 10 + ((i + 1) / totalBatches) * 80;
            updateStatus(`Translating...`, `Processing batch ${i + 1} of ${totalBatches} from API`, overallProgress);
            
            const translatedBatch = await callGeminiBatch(batch, options.apiKey, options.targetLanguage, options.model, overallProgress);
            
            for (let j = 0; j < batch.length; j++) {
                const originalText = batch[j];
                const translatedResult = translatedBatch[j];

                // CRITICAL: Check for errors. If an error is returned, do not write it to the cell.
                if (typeof translatedResult === 'string' && (translatedResult.startsWith("API Error") || translatedResult.startsWith("Blocked"))) {
                    totalErrors++;
                    if (!firstErrorMessage) {
                        firstErrorMessage = translatedResult; // Save the first error message to display
                    }
                } else {
                    // This is a valid translation.
                    const finalText = translatedResult || originalText; // Fallback to original
                    uniqueTexts.set(originalText, finalText);
                    translationCache.set(originalText, finalText);
                }
            }

            // If using the free model, wait with a visual countdown timer.
            if (options.model === STANDARD_MODEL && totalBatches > 1 && i < totalBatches - 1) {
                updateStatus("Waiting for API rate limit...", null, overallProgress);
                await startCountdownTimer(SAFE_MODE_DELAY);
            }
        }
    }
    
    updateStatus("Writing translations...", `Applying changes...`, 95);

    const newValues = JSON.parse(JSON.stringify(originalValues));
    for (const cell of cellsToTranslate) {
        // Only get a value here if the translation was successful. Otherwise, it's null.
        const translatedText = uniqueTexts.get(cell.text);
        if (translatedText) {
            let textToWrite = translatedText;
            if (typeof textToWrite === 'string' && textToWrite.length > EXCEL_CELL_CHAR_LIMIT) {
                textToWrite = textToWrite.substring(0, EXCEL_CELL_CHAR_LIMIT);
            }
            newValues[cell.row][cell.col] = textToWrite;
        }
        // If translatedText is null (because an error occurred), the original value is kept.
    }

    if (options.mode === 'replace') {
        range.values = newValues;
        await context.sync();
    } else { // 'newSheet'
        // This robust method ensures new sheet creation is reliable across Excel versions.
        const sourceSheet = range.worksheet;
        const worksheets = context.workbook.worksheets;

        // 1. Get a snapshot of sheet names BEFORE the copy.
        worksheets.load("items/name");
        await context.sync();
        const existingSheetNames = new Set(worksheets.items.map(s => s.name));

        // 2. Perform the copy operation.
        sourceSheet.copy("After");
        await context.sync();

        // 3. Reload the worksheets collection and FIND the new sheet by comparing names.
        worksheets.load("items/name");
        await context.sync();
        
        let newSheet = null;
        for (const sheet of worksheets.items) {
            if (!existingSheetNames.has(sheet.name)) {
                newSheet = sheet;
                break;
            }
        }

        // 4. If found, write the data to the new sheet.
        if (newSheet) {
            const address = range.address;
            const localAddress = address.includes('!') ? address.substring(address.indexOf('!') + 1) : address;
            newSheet.getRange(localAddress).values = newValues;
            newSheet.activate();
            await context.sync();
        } else {
            throw new Error("Fatal: Could not find the newly created worksheet after copy operation.");
        }
    }
    
    return { totalErrors, firstErrorMessage };
}

/**
 * Main function to run translation on the user's selected range.
 */
async function runSelectionTranslation() {
    const options = { apiKey: Office.context.document.settings.get("apiKey"), targetLanguage: document.getElementById("target-language-select").value, model: document.getElementById("model-selection-select").value, mode: document.querySelector('input[name="translation-mode"]:checked').value };
    document.querySelectorAll("button").forEach(b => b.disabled = true);
    try {
        await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();
            const result = await translateRange(context, selectedRange, options);
            if (result.totalErrors > 0) {
                updateStatus(`${result.totalErrors} cells could not be translated.`, result.firstErrorMessage, 100, true);
            } else {
                updateStatus("Selection translated successfully!", null, 100, false);
            }
        });
    } catch (error) {
        updateStatus("An unexpected error occurred.", `${error.name}: ${error.message}`, 100, true);
        console.error(JSON.stringify(error, null, 2));
    } finally {
        document.querySelectorAll("button").forEach(b => b.disabled = false);
    }
}

/**
 * Main function to run translation on all sheets in the workbook.
 */
async function runWorkbookTranslation() {
    const options = { apiKey: Office.context.document.settings.get("apiKey"), targetLanguage: document.getElementById("target-language-select").value, model: document.getElementById("model-selection-select").value, mode: document.querySelector('input[name="translation-mode"]:checked').value };
    document.querySelectorAll("button").forEach(b => b.disabled = true);
    let totalWorkbookErrors = 0;
    let firstOverallErrorMessage = null;

    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name");
            await context.sync();
            const originalSheets = [...worksheets.items];
            for (let i = 0; i < originalSheets.length; i++) {
                const sheet = originalSheets[i];
                if (options.mode === 'replace') { sheet.activate(); }
                const progress = ((i + 1) / originalSheets.length) * 100;
                updateStatus(`Processing sheet ${i + 1}/${originalSheets.length}: '${sheet.name}'`, null, progress);
                try {
                    const usedRange = sheet.getUsedRange(true);
                    const result = await translateRange(context, usedRange, options);
                    totalWorkbookErrors += result.totalErrors;
                    if (result.firstErrorMessage && !firstOverallErrorMessage) {
                        firstOverallErrorMessage = result.firstErrorMessage;
                    }
                } catch (error) {
                    if (error.code === "ItemNotFound") { console.log(`Sheet '${sheet.name}' is empty. Skipping.`); } 
                    else { throw error; }
                }
            }
        });
        if (totalWorkbookErrors > 0) {
            updateStatus(`Workbook translation completed with ${totalWorkbookErrors} errors.`, firstOverallErrorMessage, 100, true);
        } else {
            updateStatus("Entire workbook translated successfully!", null, 100, false);
        }
    } catch (error) {
        updateStatus("An unexpected error occurred.", `${error.name}: ${error.message}`, 100, true);
        console.error(JSON.stringify(error, null, 2));
    } finally {
        document.querySelectorAll("button").forEach(b => b.disabled = false);
    }
}

/**
 * Translates the name of the currently active sheet.
 */
async function runTranslateSheetName() {
    const options = { apiKey: Office.context.document.settings.get("apiKey"), targetLanguage: document.getElementById("target-language-select").value, model: document.getElementById("model-selection-select").value };
    document.querySelectorAll("button").forEach(b => b.disabled = true);
    try {
        await Excel.run(async (context) => {
            updateStatus("Translating sheet name...", null, 20);
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load('name');
            const worksheets = context.workbook.worksheets;
            worksheets.load('items/name');
            await context.sync();
            const existingNames = worksheets.items.map(s => s.name.toLowerCase());
            const sheetName = sheet.name;
            let translatedSheetName;

            if (translationCache.has(sheetName)) {
                translatedSheetName = translationCache.get(sheetName);
            } else {
                const translatedNameArray = await callGeminiBatch([sheetName], options.apiKey, options.targetLanguage, options.model);
                const firstResult = translatedNameArray[0];
                if (typeof firstResult === 'string' && firstResult.startsWith("API Error")) {
                    throw new Error(firstResult);
                }
                translatedSheetName = firstResult;
                translationCache.set(sheetName, translatedSheetName);
            }
            
            let finalName = translatedSheetName.replace(/[:\\/?*[\]]/g, '').substring(0, 31);
            if (sheet.name.toLowerCase() !== finalName.toLowerCase()) {
                const uniqueName = await getUniqueSheetName(finalName, existingNames);
                sheet.name = uniqueName;
            }
            await context.sync();
            updateStatus("Active sheet name translated!", null, 100);
        });
    } catch (error) {
        updateStatus("Error translating sheet name.", `${error.name}: ${error.message}`, 100, true);
        console.error(JSON.stringify(error, null, 2));
    } finally {
        document.querySelectorAll("button").forEach(b => b.disabled = false);
    }
}

/**
 * Generates a unique sheet name by appending a number if the name already exists.
 * @param {string} baseName The desired new name for the sheet.
 * @param {string[]} existingNames An array of existing sheet names.
 * @returns {string} A unique sheet name.
 */
async function getUniqueSheetName(baseName, existingNames) {
    let finalName = baseName;
    let counter = 1;
    while (existingNames.includes(finalName.toLowerCase())) {
        const suffix = `_${counter}`;
        finalName = (baseName.length > 31 - suffix.length) ? baseName.substring(0, 31 - suffix.length) + suffix : baseName + suffix;
        counter++;
    }
    return finalName;
}

// --- API CALL ---

/**
 * Calls the Google AI API with a batch of texts and handles retries for rate limiting.
 * @param {string[]} texts An array of unique strings to be translated.
 * @param {string} apiKey The user's Google AI API key.
 * @param {string} targetLanguage The language to translate the texts into.
 * @param {string} modelName The AI model to use.
 * @param {number} progress The current progress percentage for status updates.
 * @returns {Promise<string[]>} A promise that resolves to an array of translated strings or error messages.
 */
async function callGeminiBatch(texts, apiKey, targetLanguage, modelName, progress) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
    const prompt = `You are a translation API. Your only function is to translate text. Translate each string in the following JSON array to ${targetLanguage}. Detect the source language. Your response MUST BE ONLY a valid JSON array of strings containing the translations in the exact same order. Do not include any other text, markdown, or explanations. Input: ${JSON.stringify(texts)}`;
    const requestBody = { "contents": [{ "parts": [{ "text": prompt }] }] };
    
    const MAX_RETRIES = 3;
    let attempt = 0;
    
    while (attempt < MAX_RETRIES) {
        try {
            const response = await fetch(endpoint, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(requestBody) });

            if (response.status === 429) {
                attempt++;
                if (attempt >= MAX_RETRIES) { return texts.map(() => `API Error: Rate limit exceeded after ${MAX_RETRIES} retries.`); }
                const delay = Math.pow(2, attempt) * 1000;
                const waitMessage = `Rate limit hit. Waiting ${delay / 1000}s before retrying... (Attempt ${attempt}/${MAX_RETRIES})`;
                updateStatus("Translating...", waitMessage, progress, false);
                await sleep(delay);
                continue; // Retry the request
            }

            const data = await response.json();
            if (!response.ok) {
                const errorMessage = data?.error?.message || JSON.stringify(data);
                if (errorMessage.toLowerCase().includes("request payload size")) { return texts.map(() => `API Error: Request size is too large.`); }
                return texts.map(() => `API Error: ${response.status} - ${errorMessage}`);
            }

            if (!data.candidates || data.candidates.length === 0) {
                const finishReason = data.promptFeedback?.blockReason || "Safety Filter";
                return texts.map(() => `Blocked: ${finishReason}`);
            }

            let responseText = data.candidates[0].content.parts[0].text;
            responseText = responseText.trim().replace(/^```json\s*/, '').replace(/```$/, '');

            try {
                const parsed = JSON.parse(responseText);
                if (Array.isArray(parsed) && parsed.length === texts.length) { return parsed; }
                if (texts.length === 1 && typeof parsed === 'string') { return [parsed]; }
                throw new Error("Mismatched length or invalid type");
            } catch (e) {
                const jsonMatch = responseText.match(/\[.*\]/s);
                if (jsonMatch) {
                    try { return JSON.parse(jsonMatch[0]); } catch (e2) { console.error("Failed to parse extracted JSON:", jsonMatch[0]); }
                }
                return texts.map(() => "API Error: Invalid format from AI.");
            }
        } catch (error) {
            return texts.map(() => `Network Error: ${error.message}`);
        }
    }
    return texts.map(() => `API Error: Failed after all retries.`);
}