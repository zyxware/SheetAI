/*******************************
 * Google Apps Script for OpenAI API Integration
 * Made for you by https://www.zyxware.com
 * Features:
 * - Batch Processing Capability
 * - Executes OpenAI prompts on Google Sheets data
 * - Saves results back to the Data sheet
 * - Logs execution (tokens, costs)
 *********************************/

/**
 * Configuration keys used in the Config sheet
 */
const CONFIG_KEYS = {
  API_KEY: 'API_KEY',
  DEFAULT_MODEL: 'DEFAULT_MODEL',
  DEBUG: 'DEBUG',
  BATCH_SIZE: 'BATCH_SIZE'
};

/**
 * Default values for configuration
 */
const CONFIG_DEFAULTS = {
  DEFAULT_MODEL: 'gpt-4o-mini',
  DEBUG: false,
  BATCH_SIZE: 2000
};

/**
 * Gets a configuration value from the Config sheet
 * @param {string} key - The configuration key
 * @returns {any} The configuration value or undefined if not found
 */
function getConfigValue(key) {
  var configSheet = getSheet('Config');
  
  // If Config sheet doesn't exist, return undefined
  if (!configSheet) {
    return undefined;
  }
  
  var configData = configSheet.getDataRange().getValues();
  
  // Look for the key in the Config sheet
  for (var i = 0; i < configData.length; i++) {
    if (configData[i][0] === key) {
      return configData[i][1];
    }
  }
  
  // Key not found
  return undefined;
}

/**
 * Gets the API key from the Config sheet
 * @returns {string} The API key
 */
function getApiKey() {
  return getConfigValue(CONFIG_KEYS.API_KEY);
}

/**
 * Gets the default model from the Config sheet or uses the default
 * @returns {string} The default model
 */
function getDefaultModel() {
  var model = getConfigValue(CONFIG_KEYS.DEFAULT_MODEL);
  return model !== undefined ? model : CONFIG_DEFAULTS.DEFAULT_MODEL;
}

/**
 * Gets the batch size from the Config sheet or uses the default
 * @returns {number} The batch size
 */
function getBatchSize() {
  var size = getConfigValue(CONFIG_KEYS.BATCH_SIZE);
  return size !== undefined ? size : CONFIG_DEFAULTS.BATCH_SIZE;
}

/**
 * Checks if debug mode is enabled
 * @returns {boolean} True if debug mode is enabled
 */
function isDebugModeEnabled() {
  var debug = getConfigValue(CONFIG_KEYS.DEBUG);
  if (debug === undefined) {
    return CONFIG_DEFAULTS.DEBUG;
  }
  return debug === true || debug === 'TRUE' || debug === 'Yes' || debug === 'true' || debug === 1 || debug === '1';
}

/**
 * Validates that required configuration is set
 * @returns {boolean} True if configuration is valid
 */
function validateConfig() {
  var apiKey = getApiKey();
  
  if (!apiKey) {
    SpreadsheetApp.getUi().alert(
      'Configuration Error',
      'API_KEY is not set. Please create a sheet named "Config" with columns "Key" and "Value", ' +
      'and add a row with Key="API_KEY" and Value=your_openai_api_key.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
  
  return true;
}

// OpenAI Pricing Configuration
const PRICING_CONFIG = {
  "gpt-4.5-preview": { "input_per_1m": 75.00, "cached_input_per_1m": 37.50, "output_per_1m": 150.00 },
  "gpt-4o": { "input_per_1m": 2.50, "cached_input_per_1m": 1.25, "output_per_1m": 10.00 },
  "gpt-4o-mini": { "input_per_1m": 0.15, "cached_input_per_1m": 0.075, "output_per_1m": 0.60 },
  "gpt-4o-mini-audio-preview": { "input_per_1m": 0.15, "cached_input_per_1m": 0.075, "output_per_1m": 0.60 },
  "gpt-4o-audio-preview": { "input_per_1m": 2.50, "cached_input_per_1m": 1.25, "output_per_1m": 10.00 },
  "gpt-4o-mini-realtime-preview": { "input_per_1m": 0.60, "cached_input_per_1m": 0.30, "output_per_1m": 2.40 },
  "gpt-4o-realtime-preview": { "input_per_1m": 5.00, "cached_input_per_1m": 2.50, "output_per_1m": 20.00 },
  "o3-mini": { "input_per_1m": 1.10, "cached_input_per_1m": 0.55, "output_per_1m": 4.40 },
  "o1-mini": { "input_per_1m": 1.10, "cached_input_per_1m": 0.55, "output_per_1m": 4.40 },
  "o1": { "input_per_1m": 15.00, "cached_input_per_1m": 7.50, "output_per_1m": 60.00 }
};

// OpenAI Batch API Pricing Configuration
const PRICING_CONFIG_BATCH = {
"gpt-4o-mini": { "input_per_1m": 0.075, "output_per_1m": 0.30 },
"o3-mini": { "input_per_1m": 0.55, "output_per_1m": 2.20 },
"o1-mini": { "input_per_1m": 0.55, "output_per_1m": 2.20 },
"o1": { "input_per_1m": 7.50, "output_per_1m": 30.00 },
"gpt-4o": { "input_per_1m": 1.25, "output_per_1m": 5.00 },
"gpt-4.5-preview": { "input_per_1m": 37.50, "output_per_1m": 75.00 }
};
  
/* ======== UI Functions ======== */
function onOpen() {
  // Create the menu
  SpreadsheetApp.getUi()
    .createMenu('OpenAI Tools')
    .addItem('Run for First 10 Rows', 'runPromptsForFirst10Rows')
    .addItem('Run for All Rows', 'runPromptsForAllRows')
    .addSeparator()
    .addItem('Create Batch', 'createBatchWithConfigLimit')
    .addItem('Check and Process Batch', 'checkAndProcessNextCompletedBatch')
    .addItem('Check Batch Status', 'checkBatchStatus')
    .addToUi();
}
  
function runPromptsForFirst10Rows() {
  if (!validateConfig()) return;
  runPrompts(10);
}
  
function runPromptsForAllRows() {
  if (!validateConfig()) return;
  runPrompts(Infinity);
}
  
function createBatchWithConfigLimit() {
  if (!validateConfig()) return;
  
  // Check if a batch is already being created
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    showDebugAlert('Batch Creation in Progress', 
                  'A batch is already being created. Please wait until it completes.', 
                  SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  try {
    // Get the batch size from config
    var batchSize = getBatchSize();
    createBatch(Infinity, batchSize);
  } finally {
    lock.releaseLock();
  }
}
  
function checkBatchStatus() {
  if (!validateConfig()) return;
  
  var ui = SpreadsheetApp.getUi();
  
  try {
    // First check our local Batch Status sheet
    var batchStatusSheet = getSheet('Batch Status');
    if (batchStatusSheet.getLastRow() <= 1) {
      showDebugAlert('No Batches', 'No batch jobs were found in the Batch Status sheet.', ui.ButtonSet.OK);
      return;
    }
    
    var batchData = batchStatusSheet.getDataRange().getValues();
    var headers = batchData[0];
    
    // Define column indices based on expected structure
    var batchIdColIndex = headers.indexOf("Batch ID");
    var openAIBatchIdColIndex = headers.indexOf("OpenAI Batch ID");
    var statusColIndex = headers.indexOf("Status");
    var createdAtColIndex = headers.indexOf("Created At");
    var lastCheckedColIndex = headers.indexOf("Last Checked At");
    var inputFileIdColIndex = headers.indexOf("Input File ID");
    var outputFileIdColIndex = headers.indexOf("Output File ID");
    var errorFileIdColIndex = headers.indexOf("Error File ID");
    var totalRequestsColIndex = headers.indexOf("Total Requests");
    var completedColIndex = headers.indexOf("Completed");
    var failedColIndex = headers.indexOf("Failed");
    var processedColIndex = headers.indexOf("Processed");
    
    // Add Processed column if it doesn't exist
    if (processedColIndex < 0) {
      processedColIndex = headers.length;
      batchStatusSheet.getRange(1, processedColIndex + 1).setValue("Processed");
      headers.push("Processed");
      
      // Initialize all existing rows with "No" for Processed
      for (var i = 1; i < batchData.length; i++) {
        batchStatusSheet.getRange(i + 1, processedColIndex + 1).setValue("No");
      }
    }
    
    if (openAIBatchIdColIndex < 0 || statusColIndex < 0) {
      showDebugAlert('Error', 'Batch Status sheet is missing required columns.', ui.ButtonSet.OK, true);
      return;
    }
    
    // Get all batches from OpenAI
    var openAIBatches = fetchAllBatches();
    var openAIBatchesMap = {};
    
    // Create a map for quick lookup
    for (var i = 0; i < openAIBatches.length; i++) {
      openAIBatchesMap[openAIBatches[i].id] = openAIBatches[i];
    }
    
    var updatedCount = 0;
    
    // Update status for each batch in our sheet
    for (var i = 1; i < batchData.length; i++) {
      var openAIBatchId = batchData[i][openAIBatchIdColIndex];
      var currentStatus = batchData[i][statusColIndex];
      var currentProcessed = batchData[i][processedColIndex] || "No";
      
      // Skip batches that are already processed
      if (currentProcessed === "Yes") continue;
      
      // Check if this batch exists in OpenAI
      if (openAIBatchId && openAIBatchesMap[openAIBatchId]) {
        var openAIBatch = openAIBatchesMap[openAIBatchId];
        
        // If the status has changed or we need to update counts
        if (openAIBatch.status !== currentStatus || 
            (completedColIndex >= 0 && openAIBatch.request_counts.completed !== batchData[i][completedColIndex])) {
          
          // Get the full batch details
          var fullBatchDetails = retrieveBatch(openAIBatchId);
          
          // Update status
          batchStatusSheet.getRange(i + 1, statusColIndex + 1).setValue(fullBatchDetails.status);
          batchStatusSheet.getRange(i + 1, lastCheckedColIndex + 1).setValue(new Date().toISOString());
          
          // Update request counts
          if (totalRequestsColIndex >= 0) {
            batchStatusSheet.getRange(i + 1, totalRequestsColIndex + 1).setValue(fullBatchDetails.request_counts.total);
          }
          
          if (completedColIndex >= 0) {
            batchStatusSheet.getRange(i + 1, completedColIndex + 1).setValue(fullBatchDetails.request_counts.completed);
          }
          
          if (failedColIndex >= 0) {
            batchStatusSheet.getRange(i + 1, failedColIndex + 1).setValue(fullBatchDetails.request_counts.failed);
          }
          
          // Update the output file ID if available
          if (fullBatchDetails.output_file_id && outputFileIdColIndex >= 0) {
            batchStatusSheet.getRange(i + 1, outputFileIdColIndex + 1).setValue(fullBatchDetails.output_file_id);
          }
          
          updatedCount++;
          debugLog(`Updated batch ${openAIBatchId} status from ${currentStatus} to ${fullBatchDetails.status}`);
        }
      }
    }
    
    if (updatedCount > 0) {
      showDebugAlert('Batch Status Updated', `Updated status for ${updatedCount} batches.`, ui.ButtonSet.OK);
    } else {
      showDebugAlert('No Updates', 'No batch status updates were needed.', ui.ButtonSet.OK);
    }
    
    // Activate the Batch Status sheet
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(batchStatusSheet);
    
  } catch (e) {
    debugLog('Error checking batch status: ' + e.toString());
    showDebugAlert('Error', 'Failed to check batch status: ' + e.toString(), ui.ButtonSet.OK, true);
  }
}
  
function checkAndProcessNextCompletedBatch() {
  if (!validateConfig()) return;
  
  // Check if a batch is already being processed
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    showDebugAlert('Batch Processing in Progress', 
                  'A batch is already being checked or processed. Please wait until it completes.', 
                  SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  try {
    var ui = SpreadsheetApp.getUi();
    
    // Force debug logging for this function
    Logger.log("Starting checkAndProcessNextCompletedBatch");
    
    // First check our local Batch Status sheet to find batches that need processing
    var batchStatusSheet = getSheet('Batch Status');
    if (batchStatusSheet.getLastRow() <= 1) {
      Logger.log("No batches found in Batch Status sheet");
      showDebugAlert('No Batches', 'No batch jobs were found in the Batch Status sheet.', ui.ButtonSet.OK);
      return;
    }
    
    var batchData = batchStatusSheet.getDataRange().getValues();
    var headers = batchData[0];
    Logger.log("Found " + (batchData.length - 1) + " batches in Batch Status sheet");
    
    // Define column indices based on expected structure
    var batchIdColIndex = headers.indexOf("Batch ID");
    var openAIBatchIdColIndex = headers.indexOf("OpenAI Batch ID");
    var statusColIndex = headers.indexOf("Status");
    var outputFileIdColIndex = headers.indexOf("Output File ID");
    var processedColIndex = headers.indexOf("Processed");
    var lastCheckedColIndex = headers.indexOf("Last Checked At");
    
    Logger.log("Column indices - Batch ID: " + batchIdColIndex + 
               ", OpenAI Batch ID: " + openAIBatchIdColIndex + 
               ", Status: " + statusColIndex + 
               ", Output File ID: " + outputFileIdColIndex + 
               ", Processed: " + processedColIndex);
    
    // Add Processed column if it doesn't exist
    if (processedColIndex < 0) {
      processedColIndex = headers.length;
      batchStatusSheet.getRange(1, processedColIndex + 1).setValue("Processed");
      headers.push("Processed");
      
      // Initialize all existing rows with "No" for Processed
      for (var i = 1; i < batchData.length; i++) {
        batchStatusSheet.getRange(i + 1, processedColIndex + 1).setValue("No");
      }
      Logger.log("Added Processed column");
    }
    
    if (openAIBatchIdColIndex < 0 || statusColIndex < 0) {
      Logger.log("Missing required columns in Batch Status sheet");
      showDebugAlert('Error', 'Batch Status sheet is missing required columns.', ui.ButtonSet.OK, true);
      return;
    }
    
    // Get all batches from OpenAI to update status
    Logger.log("Fetching batches from OpenAI");
    var openAIBatches = fetchAllBatches();
    var openAIBatchesMap = {};
    
    // Create a map for quick lookup
    for (var i = 0; i < openAIBatches.length; i++) {
      openAIBatchesMap[openAIBatches[i].id] = openAIBatches[i];
    }
    Logger.log("Found " + openAIBatches.length + " batches in OpenAI");
    
    // First, update the status of all batches
    var updatedCount = 0;
    for (var i = 1; i < batchData.length; i++) {
      var batchId = batchData[i][batchIdColIndex];
      var openAIBatchId = batchData[i][openAIBatchIdColIndex];
      var currentStatus = batchData[i][statusColIndex];
      var currentProcessed = batchData[i][processedColIndex] || "No";
      
      Logger.log("Checking batch " + batchId + " (OpenAI ID: " + openAIBatchId + ") - Status: " + currentStatus + ", Processed: " + currentProcessed);
      
      // Skip batches that are already processed
      if (currentProcessed === "Yes") {
        Logger.log("Skipping already processed batch: " + batchId);
        continue;
      }
      
      // Check if this batch exists in OpenAI
      if (openAIBatchId && openAIBatchesMap[openAIBatchId]) {
        var openAIBatch = openAIBatchesMap[openAIBatchId];
        
        // If the status has changed or we need to update
        if (openAIBatch.status !== currentStatus || !batchData[i][outputFileIdColIndex]) {
          Logger.log("Status changed for batch " + batchId + " from " + currentStatus + " to " + openAIBatch.status);
          
          // Get the full batch details
          var fullBatchDetails = retrieveBatch(openAIBatchId);
          
          // Update status
          batchStatusSheet.getRange(i + 1, statusColIndex + 1).setValue(fullBatchDetails.status);
          batchStatusSheet.getRange(i + 1, lastCheckedColIndex + 1).setValue(new Date().toISOString());
          
          // Update the output file ID if available
          if (fullBatchDetails.output_file_id && outputFileIdColIndex >= 0) {
            batchStatusSheet.getRange(i + 1, outputFileIdColIndex + 1).setValue(fullBatchDetails.output_file_id);
            Logger.log("Updated output file ID for batch " + batchId + ": " + fullBatchDetails.output_file_id);
          }
          
          updatedCount++;
        }
      }
    }
    
    Logger.log("Updated " + updatedCount + " batches");
    
    // Now find a completed batch to process
    var batchToProcess = null;
    var batchRowIndex = -1;
    
    // Refresh the data after updates
    batchData = batchStatusSheet.getDataRange().getValues();
    
    for (var i = 1; i < batchData.length; i++) {
      var batchId = batchData[i][batchIdColIndex];
      var openAIBatchId = batchData[i][openAIBatchIdColIndex];
      var currentStatus = batchData[i][statusColIndex];
      var currentProcessed = batchData[i][processedColIndex] || "No";
      var outputFileId = batchData[i][outputFileIdColIndex];
      
      // Only process completed batches that have an output file and aren't already processed
      if (currentStatus === "completed" && outputFileId && currentProcessed === "No") {
        batchToProcess = {
          id: openAIBatchId,
          output_file_id: outputFileId
        };
        batchRowIndex = i;
        Logger.log("Found completed batch to process: " + batchId);
        break;
      }
    }
    
    // If we found a batch to process
    if (batchToProcess) {
      var batchId = batchData[batchRowIndex][batchIdColIndex];
      var openAIBatchId = batchToProcess.id;
      
      Logger.log("Processing batch " + batchId + " (OpenAI ID: " + openAIBatchId + ")");
      
      // Process the batch
      var processed = processBatchById(batchId, openAIBatchId);
      
      if (processed) {
        showDebugAlert('Batch Processed', 
                     `Successfully processed batch ${batchId}.\n\nOpenAI Batch ID: ${openAIBatchId}`, 
                     ui.ButtonSet.OK);
      } else {
        showDebugAlert('Processing Failed', 
                     `Failed to process batch ${batchId}.\n\nOpenAI Batch ID: ${openAIBatchId}\n\nCheck the Error Log for details.`, 
                     ui.ButtonSet.OK, true);
      }
    } else {
      Logger.log("No completed batches found to process");
      showDebugAlert('No Batches to Process', 
                   'No completed batches were found that need processing. Batches may still be in progress at OpenAI.', 
                   ui.ButtonSet.OK);
    }
  } catch (e) {
    Logger.log("Error in checkAndProcessNextCompletedBatch: " + e.toString());
    showDebugAlert('Error', 'Failed to check or process batches: ' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK, true);
  } finally {
    lock.releaseLock();
  }
}
  
/* ======== Utility Functions ======== */
  
function getApiKey() {
  return getSheet('Config').getRange('B1').getValue();
}
  
function getDefaultModel() {
  var configSheet = getSheet('Config');
  var model = configSheet.getRange('B2').getValue();
  return model || "gpt-4o-mini";
}
  
/**
 * Gets only the active prompts from the Prompts sheet
 * @return {Array} Array of active prompts
 */
function getActivePrompts() {
  var promptsSheet = getSheet('Prompts');
  var promptsData = promptsSheet.getDataRange().getValues();
  
  // Check if we have headers
  if (promptsData.length <= 1) {
    return [];
  }
  
  var headers = promptsData[0];
  var activeColIndex = headers.indexOf("Active");
  
  // If Active column doesn't exist, add it
  if (activeColIndex < 0) {
    activeColIndex = headers.length;
    promptsSheet.getRange(1, activeColIndex + 1).setValue("Active");
    
    // Set all existing prompts as active by default
    if (promptsData.length > 1) {
      var activeRange = promptsSheet.getRange(2, activeColIndex + 1, promptsData.length - 1, 1);
      activeRange.setValue(1);
    }
    
    // Refresh the data after adding the column
    promptsData = promptsSheet.getDataRange().getValues();
    headers = promptsData[0];
  }
  
  // Filter to only include active prompts (where Active = 1)
  var activePrompts = [];
  for (var i = 1; i < promptsData.length; i++) {
    if (promptsData[i][activeColIndex] === 1) {
      activePrompts.push(promptsData[i]);
    }
  }
  
  return activePrompts;
}
  
function getBatchSize() {
  var size = getConfigValue(CONFIG_KEYS.BATCH_SIZE);
  return size !== undefined ? size : CONFIG_DEFAULTS.BATCH_SIZE;
}
  
function getPricing(model) {
  return PRICING_CONFIG[model.toLowerCase()] || PRICING_CONFIG["gpt-4o-mini"];
}
  
function getSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  return sheet || ss.insertSheet(sheetName);
}
  
/* ======== Logging Functions ======== */
  
function logExecution(rowIndex, promptSent, responseReceived, inputTokens, outputTokens, totalTokens, model) {
  var cost = calculateCost(model, inputTokens, outputTokens);
  var logSheet = getSheet('Execution Log');

  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(["Timestamp", "Row", "Model", "Prompt Sent", "Response Received", "Input Tokens", "Output Tokens", "Total Tokens", "Cost (USD)"]);
  }

  logSheet.appendRow([
    new Date().toISOString(),
    rowIndex,
    model,
    promptSent,
    responseReceived,
    inputTokens,
    outputTokens,
    totalTokens,
    cost.toFixed(6)
  ]);
}
  
function logError(rowIndex, errorMessage) {
  getSheet('Error Log').appendRow([new Date().toISOString(), rowIndex, errorMessage]);
}
  
/* ======== Main Function to Run Prompts ======== */
function runPrompts(maxRows) {
  var ui = SpreadsheetApp.getUi();
  var apiKey = getApiKey();
  
  if (!apiKey) {
    showDebugAlert('Error', 'API key is missing. Please add it to the Config sheet.', ui.ButtonSet.OK, true);
    return;
  }
  
  try {
    // Record start time for this execution
    var startTime = new Date();
    var debugMode = isDebugModeEnabled();
    
    // Get active prompts from the Prompts sheet
    var promptsSheet = getSheet('Prompts');
    if (!promptsSheet) {
      showDebugAlert('Error', 'Prompts sheet not found.', ui.ButtonSet.OK, true);
      return;
    }
    
    var promptsData = promptsSheet.getDataRange().getValues();
    var promptsHeaders = promptsData[0];
    
    // Find column indices in the Prompts sheet
    var promptNameIndex = -1;
    var promptTextIndex = -1;
    var modelIndex = -1;
    var activeIndex = -1;
    
    for (var i = 0; i < promptsHeaders.length; i++) {
      var header = promptsHeaders[i];
      if (header === null || header === undefined) continue;
      
      var headerStr = String(header); // Safely convert to string
      
      if (headerStr === 'Prompt Name') promptNameIndex = i;
      if (headerStr === 'Prompt Text') promptTextIndex = i;
      if (headerStr === 'Model') modelIndex = i;
      if (headerStr === 'Active') activeIndex = i;
    }
    
    if (promptNameIndex < 0 || promptTextIndex < 0) {
      showDebugAlert('Error', 'Prompts sheet is missing required columns.', ui.ButtonSet.OK, true);
      return;
    }
    
    // Get active prompts
    var activePrompts = [];
    for (var i = 1; i < promptsData.length; i++) {
      var isActive = activeIndex >= 0 ? 
                    (promptsData[i][activeIndex] === 1 || 
                     promptsData[i][activeIndex] === '1' || 
                     promptsData[i][activeIndex] === true || 
                     promptsData[i][activeIndex] === 'TRUE' || 
                     promptsData[i][activeIndex] === 'Yes' || 
                     promptsData[i][activeIndex] === 'true') : true;
      
      if (isActive) {
        activePrompts.push({
          name: promptsData[i][promptNameIndex],
          text: promptsData[i][promptTextIndex],
          model: modelIndex >= 0 ? promptsData[i][modelIndex] : getDefaultModel()
        });
      }
    }
    
    if (activePrompts.length === 0) {
      showDebugAlert('No Active Prompts', 'No active prompts found in the Prompts sheet.', ui.ButtonSet.OK);
      return;
    }
    
    // Get data from the Data sheet
    var dataSheet = getSheet('Data');
    if (!dataSheet) {
      showDebugAlert('Error', 'Data sheet not found.', ui.ButtonSet.OK, true);
      return;
    }
    
    var dataRange = dataSheet.getDataRange().getValues();
    var headers = dataRange[0];
    
    // Find Status column or add it if it doesn't exist
    var statusColIndex = -1;
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i];
      if (header === null || header === undefined) continue;
      
      var headerStr = String(header); // Safely convert to string
      if (headerStr === 'Status') {
        statusColIndex = i;
        break;
      }
    }
    
    if (statusColIndex < 0) {
      statusColIndex = headers.length;
      dataSheet.getRange(1, statusColIndex + 1).setValue('Status');
      headers.push('Status');
    }
    
    // Find rows that need processing (status is 0 or empty)
    var rowsToProcess = [];
    for (var i = 1; i < dataRange.length && rowsToProcess.length < maxRows; i++) {
      var status = dataRange[i][statusColIndex];
      if (status === 0 || status === '' || status === null || status === undefined) {
        rowsToProcess.push(i);
      }
    }
    
    if (rowsToProcess.length === 0) {
      showDebugAlert('No Data', 'No rows found that need processing.', ui.ButtonSet.OK);
      return;
    }
    
    // Track metrics
    var promptMetrics = {};
    var totalProcessed = 0;
    var totalErrors = 0;
    
    // Process each row
    for (var i = 0; i < rowsToProcess.length; i++) {
      var rowIndex = rowsToProcess[i];
      var rowNumber = rowIndex + 1; // +1 for the actual row number in the sheet
      
      try {
        // Mark the row as in progress (status = 1)
        dataSheet.getRange(rowNumber, statusColIndex + 1).setValue(1);
        
        // Process each active prompt for this row
        for (var j = 0; j < activePrompts.length; j++) {
          var prompt = activePrompts[j];
          var promptName = prompt.name;
          var promptTemplate = prompt.text;
          var model = prompt.model || getDefaultModel();
          
          try {
            // Replace placeholders in the prompt template
            var promptText = promptTemplate;
            
            // Find all placeholders in the format {{Column Name}}
            var placeholders = promptTemplate.match(/\{\{([^}]+)\}\}/g) || [];
            
            for (var k = 0; k < placeholders.length; k++) {
              var placeholder = placeholders[k];
              var columnName = placeholder.substring(2, placeholder.length - 2).trim();
              
              // Find the column index
              var columnIndex = -1;
              for (var l = 0; l < headers.length; l++) {
                var header = headers[l];
                if (header === null || header === undefined) continue;
                
                var headerStr = String(header); // Safely convert to string
                if (headerStr === columnName) {
                  columnIndex = l;
                  break;
                }
              }
              
              if (columnIndex >= 0) {
                var cellValue = dataRange[rowIndex][columnIndex] || '';
                promptText = promptText.replace(placeholder, cellValue);
              }
            }
            
            // Call the OpenAI API
            var apiCallStartTime = new Date();
            var response = callOpenAI(apiKey, model, promptText);
            var apiCallEndTime = new Date();
            var apiCallDuration = (apiCallEndTime - apiCallStartTime) / 1000; // Duration in seconds
            
            var responseText = response.text;
            var parsedResponse = response.parsedJson;
            var inputTokens = response.inputTokens;
            var outputTokens = response.outputTokens;
            var totalTokens = response.totalTokens;
            var cost = calculateCost(model, inputTokens, outputTokens, true);
            
            // Update metrics
            if (!promptMetrics[promptName]) {
              promptMetrics[promptName] = {
                count: 0,
                inputTokens: 0,
                outputTokens: 0,
                totalTokens: 0,
                cost: 0,
                model: model,
                duration: 0
              };
            }
            
            promptMetrics[promptName].count++;
            promptMetrics[promptName].inputTokens += inputTokens;
            promptMetrics[promptName].outputTokens += outputTokens;
            promptMetrics[promptName].totalTokens += totalTokens;
            promptMetrics[promptName].cost += cost;
            promptMetrics[promptName].duration += apiCallDuration;
            
            // Save response to the Data sheet
            saveResponseToDataSheet(dataSheet, headers, rowIndex, parsedResponse, promptName);
            
            // Log execution
            if (debugMode) {
              addExecutionLog(
                new Date(),
                rowNumber,
                model,
                promptName,
                responseText,
                inputTokens,
                outputTokens,
                totalTokens,
                cost
              );
            }
            
            totalProcessed++;
          } catch (e) {
            logError(rowNumber + 1, `Error processing ${promptName}: ${e.toString()}`);
            totalErrors++;
          }
        }
        
        // Mark the row as completed (status = 1 for non batch mode)
        dataSheet.getRange(rowNumber, statusColIndex + 1).setValue(1);
      } catch (e) {
        logError(rowNumber + 1, e.toString());
        totalErrors++;
      }
    }
    
    // Record end time and calculate duration
    var endTime = new Date();
    var executionDuration = (endTime - startTime) / 1000; // Duration in seconds
    
    // Add summary entries for each prompt
    if (Object.keys(promptMetrics).length > 0) {
      for (var promptName in promptMetrics) {
        var metrics = promptMetrics[promptName];
        addPromptSummary(
          startTime,
          endTime,
          metrics.duration,
          promptName,
          metrics.count,
          metrics.inputTokens,
          metrics.outputTokens,
          metrics.cost
        );
      }
    }
    
    showDebugAlert('Processing Complete', 
                 `Processed ${totalProcessed} prompts with ${totalErrors} errors.`, 
                 ui.ButtonSet.OK);
  } catch (e) {
    debugLog('Error running prompts: ' + e.toString());
    showDebugAlert('Error', 'Failed to run prompts: ' + e.toString(), ui.ButtonSet.OK, true);
  }
}
  
/* ======== Cost Summary Functions ======== */
function addPromptSummary(startTime, endTime, durationSeconds, promptName, rowsExecuted, inputTokens, outputTokens, cost) {
  var costSummarySheet = getSheet('Cost Summary');
  
  // Initialize headers if sheet is empty
  if (costSummarySheet.getLastRow() === 0) {
    costSummarySheet.appendRow([
      "Date", 
      "Start Time", 
      "End Time", 
      "Duration (sec)", 
      "Prompt Title", 
      "No. of Rows Executed", 
      "Total Input Tokens", 
      "Total Output Tokens", 
      "Total Tokens", 
      "Total Cost (USD)"
    ]);
  }
  
  // Format date and times
  var dateStr = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var startTimeStr = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "HH:mm:ss");
  var endTimeStr = Utilities.formatDate(endTime, Session.getScriptTimeZone(), "HH:mm:ss");
  
  // Add new row for this prompt execution
  costSummarySheet.appendRow([
    dateStr,
    startTimeStr,
    endTimeStr,
    durationSeconds.toFixed(1),
    promptName,
    rowsExecuted,
    inputTokens,
    outputTokens,
    inputTokens + outputTokens,
    cost.toFixed(6)
  ]);
  
  // Format the cost column as currency
  var lastRow = costSummarySheet.getLastRow();
  costSummarySheet.getRange(lastRow, 10).setNumberFormat("$0.000000");
  
  // Format the duration as number with 1 decimal place
  costSummarySheet.getRange(lastRow, 4).setNumberFormat("0.0");
}
  
/* ======== Save Cleaned Response to Data Sheet ======== */
function saveResponseToDataSheet(sheet, headers, rowIndex, response, promptName) {
  try {
    // No need to parse the response again as it's already a JSON object
    var parsedResponse = response;

    for (var key in parsedResponse) {
      if (Object.prototype.hasOwnProperty.call(parsedResponse, key)) {
        var colName = promptName + ' - ' + key; // Format: Prompt Name - Key
        var colIndex = headers.indexOf(colName);

        // If the column does not exist, create it
        if (colIndex < 0) {
          colIndex = headers.length;
          sheet.getRange(1, colIndex + 1).setValue(colName);
          headers.push(colName);
        }

        // Write response data to the correct cell in the row
        sheet.getRange(rowIndex + 1, colIndex + 1).setValue(parsedResponse[key]);
      }
    }
  } catch (e) {
    debugLog("Error processing response: " + e);
    logError(rowIndex + 1, "Error processing response: " + e.toString());
  }
}
  
/* ======== OpenAI API Function ======== */
function callOpenAI(apiKey, model, prompt) {
  var url = 'https://api.openai.com/v1/chat/completions';

  var payload = {
    model: model,
    messages: [
      { role: 'system', content: 'You are a helpful assistant. Return valid JSON only.' },
      { role: 'user', content: prompt }
    ],
    temperature: 0.0,
    max_tokens: 256,
    seed: 42,
    response_format: { type: "json_object" }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var jsonText = response.getContentText().trim();
  var json = JSON.parse(jsonText);

  if (json.error) {
    throw new Error('OpenAI API error: ' + json.error.message);
  }

  // Parse the content string into a JSON object
  var contentStr = json.choices[0]?.message?.content || '{}';
  var parsedContent;
  
  try {
    parsedContent = JSON.parse(contentStr);
  } catch (e) {
    throw new Error('Failed to parse OpenAI response as JSON: ' + e.toString());
  }
  
  return {
    text: JSON.stringify(parsedContent),
    parsedJson: parsedContent,
    inputTokens: json.usage?.prompt_tokens || 0,
    outputTokens: json.usage?.completion_tokens || 0,
    totalTokens: json.usage?.total_tokens || 0
  };
}
  
/* ======== Placeholder Replacement Function ======== */
function replaceVariables(prompt, headers, rowData) {
  var finalPrompt = prompt;
  for (var colIndex = 0; colIndex < headers.length; colIndex++) {
    var placeholder = '{{' + headers[colIndex].trim() + '}}';
    var value = rowData[colIndex] !== undefined ? rowData[colIndex] : "N/A";
    finalPrompt = finalPrompt.replace(new RegExp(placeholder, 'g'), value);
  }
  return finalPrompt;
}
  
/* ======== Calculate OpenAI API Cost ======== */
function calculateCost(model, inputTokens, outputTokens, isBatch = false) {
  // Normalize model name to lowercase
  var modelLower = model.toLowerCase();
  
  // Get pricing configuration based on whether this is a batch request
  var pricingConfig = isBatch ? PRICING_CONFIG_BATCH : PRICING_CONFIG;
  
  // Get pricing for the model, or use default if not found
  var pricing = pricingConfig[modelLower];
  if (!pricing) {
    // If model not found in batch pricing but this is a batch, try standard pricing
    if (isBatch && PRICING_CONFIG[modelLower]) {
      pricing = {
        input_per_1m: PRICING_CONFIG[modelLower].input_per_1m / 2,
        output_per_1m: PRICING_CONFIG[modelLower].output_per_1m / 2
      };
    } else {
      // Default to gpt-4o-mini pricing if model not found
      pricing = pricingConfig["gpt-4o-mini"];
    }
  }
  
  // Calculate cost
  var inputCost = (inputTokens / 1000000) * pricing.input_per_1m;
  var outputCost = (outputTokens / 1000000) * pricing.output_per_1m;
  
  return inputCost + outputCost;
}
  
/* ======== Batch Processing Functions ======== */
  
/**
 * Creates a batch job for the specified number of rows
 */
function createBatch(maxRows, batchSize) {
  var ui = SpreadsheetApp.getUi();
  var apiKey = getApiKey();
  
  if (!apiKey) {
    showDebugAlert('Error', 'API key is missing. Please add it to the Config sheet.', ui.ButtonSet.OK, true);
    return;
  }
  
  try {
    // Find the next set of rows to process
    var nextBatchInfo = findNextBatchRows(maxRows, batchSize);
    
    if (!nextBatchInfo || nextBatchInfo.startRow > nextBatchInfo.endRow) {
      showDebugAlert('No Data', 'No more rows to process or all rows are already processed.', ui.ButtonSet.OK);
      return;
    }
    
    // Prepare the batch data
    var batchData = prepareBatchDataRange(nextBatchInfo.startRow, nextBatchInfo.endRow);
    
    // Check if there are any requests to process
    if (!batchData || !batchData.requests || batchData.requests.length === 0) {
      showDebugAlert('No Data', 'No requests to process in the selected rows.', ui.ButtonSet.OK);
      return;
    }
    
    // Check batch size limits
    if (batchData.requests.length > 50000) {
      showDebugAlert('Batch Too Large', 
               'This batch contains ' + batchData.requests.length + ' requests, which exceeds the OpenAI limit of 50,000 requests per batch. Please reduce the ' + CONFIG_KEYS.BATCH_SIZE + ' in Config.',
               ui.ButtonSet.OK);
      return;
    }
    
    // Create the batch job
    var batch = createBatchJob(batchData.requests);
    
    // Store batch information in the Batch Status sheet
    var batchId = storeBatchInfo(batch, batchData.rowIndices);
    
    // Update the Data sheet with batch IDs
    updateDataSheetWithBatchId(batchData.rowIndices, batchId);
    
    showDebugAlert('Success', 
             `Batch job created successfully!\n\nProcessed rows ${nextBatchInfo.startRow} to ${nextBatchInfo.endRow}\nBatch ID: ${batch.id}\nStatus: ${batch.status}\nTotal Requests: ${batchData.requests.length}\n\n${nextBatchInfo.remainingRows > 0 ? 'There are ' + nextBatchInfo.remainingRows + ' more rows to process. Run "Create Batch" again to process the next set.' : 'All rows have been processed.'}`, 
             ui.ButtonSet.OK);
             
  } catch (e) {
    debugLog('Error creating batch: ' + e.toString());
    showDebugAlert('Error', 'Failed to create batch: ' + e.toString(), ui.ButtonSet.OK, true);
  }
}
  
/**
 * Finds the next set of rows to process
 */
function findNextBatchRows(maxRows, batchSize) {
  var dataSheet = getSheet('Data');
  
  if (!dataSheet) {
    return null;
  }
  
  var lastRow = dataSheet.getLastRow();
  
  // If there's only a header row or no data at all
  if (lastRow <= 1) {
    return null;
  }
  
  var dataRange = dataSheet.getRange(1, 1, lastRow, dataSheet.getLastColumn()).getValues();
  var headers = dataRange[0];
  
  // Find the Status column
  var statusColIndex = headers.indexOf("Status");
  if (statusColIndex < 0) {
    // If no Status column exists, add one
    statusColIndex = headers.length;
    dataSheet.getRange(1, statusColIndex + 1).setValue("Status");
    
    // Update our local copy of the data
    headers.push("Status");
    for (var i = 1; i < dataRange.length; i++) {
      dataRange[i][statusColIndex] = 0;
    }
  }
  
  // Find the first row that hasn't been processed yet
  var startRow = -1;
  for (var i = 1; i < dataRange.length; i++) {
    if (!dataRange[i][statusColIndex] || dataRange[i][statusColIndex] === 0) {
      startRow = i + 1; // +1 because row indices are 1-based
      break;
    }
  }
  
  // If no unprocessed rows were found
  if (startRow === -1) {
    return {
      startRow: 0,
      endRow: 0,
      remainingRows: 0
    };
  }
  
  // Calculate the end row based on batch size
  var endRow = Math.min(startRow + batchSize - 1, lastRow);
  
  // Count how many rows are left after this batch
  var remainingRows = 0;
  for (var i = endRow; i < dataRange.length; i++) {
    if (!dataRange[i][statusColIndex] || dataRange[i][statusColIndex] === 0) {
      remainingRows++;
    }
  }
  
  return {
    startRow: startRow,
    endRow: endRow,
    remainingRows: remainingRows
  };
}
  
/**
 * Prepares batch data for a specific range of rows
 */
function prepareBatchDataRange(startRow, endRow) {
  // If invalid range, return empty result
  if (startRow <= 0 || endRow <= 0 || startRow > endRow) {
    return {
      requests: [],
      rowIndices: []
    };
  }
  
  var dataSheet = getSheet('Data');
  
  if (!dataSheet) {
    return {
      requests: [],
      rowIndices: []
    };
  }
  
  var lastRow = dataSheet.getLastRow();
  
  // Adjust endRow if it exceeds the actual data
  endRow = Math.min(endRow, lastRow);
  
  // If we're only looking at the header row
  if (startRow === 1 && endRow === 1) {
    return {
      requests: [],
      rowIndices: []
    };
  }
  
  var dataRange = dataSheet.getRange(1, 1, lastRow, dataSheet.getLastColumn()).getValues();
  var headers = dataRange[0];
  
  // Ensure Status column exists
  var statusColIndex = headers.indexOf("Status");
  if (statusColIndex < 0) {
    statusColIndex = headers.length;
    dataSheet.getRange(1, statusColIndex + 1).setValue("Status");
    headers.push("Status");
  }
  
  // Ensure Batch ID column exists
  var batchIdColIndex = headers.indexOf("Batch ID");
  if (batchIdColIndex < 0) {
    batchIdColIndex = headers.length;
    dataSheet.getRange(1, batchIdColIndex + 1).setValue("Batch ID");
    headers.push("Batch ID");
  }
  
  // Get only active prompts
  var prompts = getActivePrompts();
  
  var requests = [];
  var rowIndices = []; // Stores row indices that are part of this batch
  
  // Process only rows in the specified range
  for (var rowIndex = startRow - 1; rowIndex < endRow; rowIndex++) {
    // Skip rows that are already processed or already part of a batch
    if (rowIndex >= dataRange.length || 
        (dataRange[rowIndex][statusColIndex] > 0 || 
         dataRange[rowIndex][batchIdColIndex])) {
      continue;
    }
    
    // Add this row to the batch
    rowIndices.push(rowIndex + 1);
    
    for (var p = 0; p < prompts.length; p++) {
      var promptName = prompts[p][0];
      var promptText = prompts[p][1];
      var model = (prompts[p][2] || getDefaultModel()).toLowerCase();
      
      var finalPrompt = replaceVariables(promptText, headers, dataRange[rowIndex]);
      
      // Create a unique custom_id that encodes all necessary information
      // Format: row-{rowIndex}-prompt-{promptIndex}-{promptName}
      // This way we can extract all needed info from the custom_id itself
      var customId = `row-${rowIndex+1}-prompt-${p+1}-${encodeURIComponent(promptName)}`;
      
      // Create the request object
      var request = {
        custom_id: customId,
        method: "POST",
        url: "/v1/chat/completions",
        body: {
          model: model,
          messages: [
            { role: 'system', content: 'You are a helpful assistant. Return valid JSON only.' },
            { role: 'user', content: finalPrompt }
          ],
          temperature: 0.0,
          max_tokens: 256,
          seed: 42,
          response_format: { type: "json_object" }
        }
      };
      
      requests.push(request);
    }
  }
  
  return {
    requests: requests,
    rowIndices: rowIndices
  };
}
  
/**
 * Updates the Data sheet with batch IDs for the rows in this batch
 */
function updateDataSheetWithBatchId(rowIndices, batchId) {
  if (!rowIndices || rowIndices.length === 0) return;
  
  var dataSheet = getSheet('Data');
  var headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  var batchIdColIndex = headers.indexOf("Batch ID");
  var statusColIndex = headers.indexOf("Status");
  
  if (batchIdColIndex < 0 || statusColIndex < 0) return;
  
  // Update each row with the batch ID and set status to 1 (batch uploaded)
  for (var i = 0; i < rowIndices.length; i++) {
    var rowIndex = rowIndices[i];
    dataSheet.getRange(rowIndex, batchIdColIndex + 1).setValue(batchId);
    dataSheet.getRange(rowIndex, statusColIndex + 1).setValue(1); // Status 1 = batch uploaded
  }
}
  
/**
 * Stores batch information in the Batch Status sheet
 */
function storeBatchInfo(batch, rowIndices) {
  var batchStatusSheet = getSheet('Batch Status');
  
  // Create the sheet if it doesn't exist
  if (!batchStatusSheet) {
    batchStatusSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Batch Status');
    batchStatusSheet.appendRow([
      "Batch ID",
      "OpenAI Batch ID",
      "Status",
      "Created At",
      "Last Checked At",
      "Input File ID",
      "Output File ID",
      "Error File ID",
      "Total Requests",
      "Completed",
      "Failed",
      "Processed"
    ]);
    
    // Format the header row
    batchStatusSheet.getRange(1, 1, 1, 12).setFontWeight('bold');
    batchStatusSheet.setFrozenRows(1);
  }
  
  // Generate a unique batch ID
  var batchId = Utilities.getUuid();
  
  // Add the batch information
  batchStatusSheet.appendRow([
    batchId,
    batch.id,
    batch.status,
    new Date().toISOString(),
    new Date().toISOString(),
    batch.input_file_id || "",
    batch.output_file_id || "",
    batch.error_file_id || "",
    batch.request_counts.total,
    batch.request_counts.completed,
    batch.request_counts.failed,
    "No"
  ]);
  
  return batchId;
}
  
/**
 * Processes a specific batch by its OpenAI batch ID
 */
function processBatchById(batchId, openAIBatchId) {
  Logger.log("Starting processBatchById for batch " + batchId + " (OpenAI ID: " + openAIBatchId + ")");
  
  var batchStatusSheet = getSheet('Batch Status');
  var batchData = batchStatusSheet.getDataRange().getValues();
  var headers = batchData[0];
  
  // Define column indices based on expected structure
  var batchIdColIndex = headers.indexOf("Batch ID");
  var openAIBatchIdColIndex = headers.indexOf("OpenAI Batch ID");
  var statusColIndex = headers.indexOf("Status");
  var outputFileIdColIndex = headers.indexOf("Output File ID");
  var processedColIndex = headers.indexOf("Processed");
  
  Logger.log("Column indices - Batch ID: " + batchIdColIndex + 
             ", OpenAI Batch ID: " + openAIBatchIdColIndex + 
             ", Status: " + statusColIndex + 
             ", Output File ID: " + outputFileIdColIndex + 
             ", Processed: " + processedColIndex);
  
  if (openAIBatchIdColIndex < 0 || batchIdColIndex < 0 || outputFileIdColIndex < 0) {
    Logger.log("Missing required columns in Batch Status sheet");
    return false;
  }
  
  var processed = false;
  
  // Find the batch in our sheet
  for (var i = 1; i < batchData.length; i++) {
    var currentBatchId = batchData[i][batchIdColIndex];
    var currentOpenAIBatchId = batchData[i][openAIBatchIdColIndex];
    
    if (currentBatchId === batchId || currentOpenAIBatchId === openAIBatchId) {
      Logger.log("Found batch in row " + (i+1));
      
      var outputFileId = batchData[i][outputFileIdColIndex];
      var currentProcessed = processedColIndex >= 0 ? batchData[i][processedColIndex] : "No";
      
      // Skip if already processed
      if (currentProcessed === "Yes") {
        Logger.log("Batch " + batchId + " has already been processed");
        return true;
      }
      
      if (!outputFileId) {
        Logger.log("Warning: No output file ID for batch " + batchId);
        return false;
      }
      
      try {
        Logger.log("Downloading output file " + outputFileId);
        var outputContent = downloadFileFromOpenAI(outputFileId);
        if (!outputContent) {
          Logger.log("Warning: Could not download output file for batch " + batchId);
          return false;
        }
        
        Logger.log("Processing output file content (length: " + outputContent.length + ")");
        // Log a sample of the content
        Logger.log("Content sample: " + outputContent.substring(0, 200) + "...");
        
        // Process the results
        var result = processOutputFile(outputContent, null, batchId);
        Logger.log("Processed output file: " + JSON.stringify(result));
        
        // Update the status to "processed" in our sheet
        if (statusColIndex >= 0) {
          batchStatusSheet.getRange(i + 1, statusColIndex + 1).setValue("processed");
        }
        
        // Mark as processed
        if (processedColIndex >= 0) {
          batchStatusSheet.getRange(i + 1, processedColIndex + 1).setValue("Yes");
        }
        
        Logger.log("Successfully processed batch " + batchId + " with " + result.success + " successful requests and " + result.failed + " failed requests");
        processed = true;
      } catch (e) {
        Logger.log("Error processing batch " + batchId + ": " + e.toString());
        // Log the error but continue
        addErrorLog(new Date(), 0, "Processing Error", "Error processing batch " + batchId + ": " + e.toString(), batchId);
      }
      break;
    }
  }
  
  if (!processed) {
    Logger.log("Warning: Could not find batch " + batchId + " in Batch Status sheet for processing");
  }
  
  return processed;
}
  
/**
 * Fetches all batches from OpenAI
 */
function fetchAllBatches() {
  var apiKey = getApiKey();
  
  var options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + apiKey },
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch('https://api.openai.com/v1/batches', options);
  var responseJson = JSON.parse(response.getContentText());
  
  if (responseJson.error) {
    throw new Error('OpenAI API error: ' + responseJson.error.message);
  }
  
  return responseJson.data || [];
}
  
/**
 * Creates a batch job with the given requests
 * @param {Array} requests - Array of request objects
 * @returns {Object} The created batch object
 */
function createBatchJob(requests) {
  var apiKey = getApiKey();
  
  // Create the JSONL content
  var jsonlContent = createJsonlContent(requests);
  
  // Estimate JSONL file size
  var estimatedSizeMB = jsonlContent.length / (1024 * 1024);
  if (estimatedSizeMB > 200) {
    throw new Error('The estimated batch file size is ' + estimatedSizeMB.toFixed(2) + 
                   ' MB, which exceeds the OpenAI limit of 200 MB. Please reduce the ' + 
                   CONFIG_KEYS.BATCH_SIZE + ' in Config.');
  }
  
  // Upload the file to OpenAI
  var fileId = uploadFileToOpenAI(jsonlContent);
  if (!fileId) {
    throw new Error('Failed to upload batch file to OpenAI.');
  }
  
  // Create the batch
  var batch = createOpenAIBatch(fileId);
  if (!batch) {
    throw new Error('Failed to create batch job.');
  }
  
  return batch;
}
  
/**
 * Creates a JSONL string from an array of request objects
 * @param {Array} requests - Array of request objects
 * @returns {string} JSONL content
 */
function createJsonlContent(requests) {
  return requests.map(function(request) {
    return JSON.stringify(request);
  }).join('\n');
}
  
/**
 * Uploads a file to OpenAI
 */
function uploadFileToOpenAI(jsonlContent) {
  var apiKey = getApiKey();
  var boundary = Utilities.getUuid();
  
  var metadata = {
    purpose: "batch"
  };
  
  var payload = "--" + boundary + "\r\n" +
                "Content-Disposition: form-data; name=\"purpose\"\r\n\r\n" +
                metadata.purpose + "\r\n" +
                "--" + boundary + "\r\n" +
                "Content-Disposition: form-data; name=\"file\"; filename=\"batch_requests.jsonl\"\r\n" +
                "Content-Type: application/jsonl\r\n\r\n" +
                jsonlContent + "\r\n" +
                "--" + boundary + "--";
  
  var options = {
    method: 'post',
    contentType: 'multipart/form-data; boundary=' + boundary,
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: payload,
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch('https://api.openai.com/v1/files', options);
  var responseJson = JSON.parse(response.getContentText());
  
  if (responseJson.error) {
    throw new Error('OpenAI API error: ' + responseJson.error.message);
  }
  
  return responseJson.id;
}
  
/**
 * Creates a batch job in OpenAI
 */
function createOpenAIBatch(fileId) {
  var apiKey = getApiKey();
  
  var payload = {
    input_file_id: fileId,
    endpoint: "/v1/chat/completions",
    completion_window: "24h"
  };
  
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch('https://api.openai.com/v1/batches', options);
  var responseJson = JSON.parse(response.getContentText());
  
  if (responseJson.error) {
    throw new Error('OpenAI API error: ' + responseJson.error.message);
  }
  
  return responseJson;
}
  
/**
 * Retrieves batch information from OpenAI
 */
function retrieveBatch(batchId) {
  var apiKey = getApiKey();
  
  var options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + apiKey },
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(`https://api.openai.com/v1/batches/${batchId}`, options);
  var responseJson = JSON.parse(response.getContentText());
  
  if (responseJson.error) {
    throw new Error('OpenAI API error: ' + responseJson.error.message);
  }
  
  return responseJson;
}
  
/**
 * Downloads a file from OpenAI
 */
function downloadFileFromOpenAI(fileId) {
  var apiKey = getApiKey();
  
  var options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + apiKey },
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(`https://api.openai.com/v1/files/${fileId}/content`, options);
  return response.getContentText();
}
  
/**
 * Processes the output file and updates the Data sheet
 */
function processOutputFile(outputContent, rowMap, batchId) {
  Logger.log("Starting processOutputFile for batch " + batchId);
  
  var dataSheet = getSheet('Data');
  var headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  
  // Find Status and Batch ID columns
  var statusColIndex = headers.indexOf("Status");
  var batchIdColIndex = headers.indexOf("Batch ID");
  
  Logger.log("Column indices - Status: " + statusColIndex + ", Batch ID: " + batchIdColIndex);
  
  var lines = outputContent.split('\n').filter(line => line.trim()); // Filter out empty lines
  Logger.log("Found " + lines.length + " lines in output file");
  
  // Calculate the total number of requests from the number of lines
  var totalRequests = lines.length;
  var successfulRequests = 0;
  var failedRequests = 0;
  
  // Track metrics for cost summary
  var promptMetrics = {};
  
  for (var i = 0; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    
    try {
      var result = JSON.parse(lines[i]);
      var customId = result.custom_id;
      
      Logger.log("Processing line " + (i+1) + " with custom_id: " + customId);
      
      // Parse the custom_id to extract row and prompt information
      // Format: row-{rowIndex}-prompt-{promptIndex}-{promptName}
      var customIdParts = customId.split('-');
      if (customIdParts.length < 5 || customIdParts[0] !== 'row' || customIdParts[2] !== 'prompt') {
        Logger.log("Warning: Invalid custom_id format: " + customId);
        addErrorLog(new Date(), 0, "Invalid Format", "Invalid custom_id format: " + customId, batchId);
        continue;
      }
      
      var rowNumber = parseInt(customIdParts[1]);
      var promptIndex = parseInt(customIdParts[3]);
      
      // Extract the prompt name (which might contain hyphens)
      var promptNameEncoded = customIdParts.slice(4).join('-');
      var promptName = decodeURIComponent(promptNameEncoded);
      
      Logger.log("Parsed custom_id - Row: " + rowNumber + ", Prompt Index: " + promptIndex + ", Prompt Name: " + promptName);
      
      if (isNaN(rowNumber) || isNaN(promptIndex)) {
        Logger.log("Warning: Invalid row or prompt index in custom_id: " + customId);
        addErrorLog(new Date(), 0, "Invalid Format", "Invalid row or prompt index in custom_id: " + customId, batchId);
        continue;
      }
      
      if (result.error) {
        // Log the error
        Logger.log("API Error for row " + rowNumber + ": " + result.error.message);
        addErrorLog(new Date(), rowNumber, "API Error", `Batch error for ${promptName}: ${result.error.message}`, batchId);
        failedRequests++;
        continue;
      }
      
      var response = result.response;
      if (response && response.status_code === 200 && response.body) {
        var responseBody = response.body;
        var content = responseBody.choices[0].message.content;
        
        try {
          Logger.log("Parsing content for row " + rowNumber);
          var parsedContent = JSON.parse(content);
          
          // Save the response to the Data sheet
          saveResponseToDataSheet(dataSheet, headers, rowNumber - 1, parsedContent, promptName);
          
          // Mark the row as processed (status = 2 for batch completed)
          if (statusColIndex >= 0) {
            dataSheet.getRange(rowNumber, statusColIndex + 1).setValue(2);
          }
          
          // Set the batch ID if it's not already set
          if (batchIdColIndex >= 0) {
            var currentBatchId = dataSheet.getRange(rowNumber, batchIdColIndex + 1).getValue();
            if (!currentBatchId) {
              dataSheet.getRange(rowNumber, batchIdColIndex + 1).setValue(batchId);
            }
          }
          
          // Track usage for cost summary
          var model = responseBody.model;
          var inputTokens = responseBody.usage.prompt_tokens;
          var outputTokens = responseBody.usage.completion_tokens;
          var totalTokens = responseBody.usage.total_tokens;
          var cost = calculateCost(model, inputTokens, outputTokens, true);
          
          Logger.log("Successfully processed row " + rowNumber + " with model " + model);
          
          // Add to execution log with the response content
          addExecutionLog(
            new Date(),
            rowNumber,
            model,
            promptName,
            content, // Include the actual response content
            inputTokens,
            outputTokens,
            totalTokens,
            cost
          );
          
          if (!promptMetrics[promptName]) {
            promptMetrics[promptName] = {
              count: 0,
              inputTokens: 0,
              outputTokens: 0,
              totalTokens: 0,
              cost: 0,
              model: model,
              duration: 0 // We don't have individual durations for batch requests
            };
          }
          
          promptMetrics[promptName].count++;
          promptMetrics[promptName].inputTokens += inputTokens;
          promptMetrics[promptName].outputTokens += outputTokens;
          promptMetrics[promptName].totalTokens += totalTokens;
          promptMetrics[promptName].cost += cost;
          
          successfulRequests++;
        } catch (e) {
          Logger.log("Error parsing content for row " + rowNumber + ": " + e.toString());
          addErrorLog(new Date(), rowNumber, "Parse Error", `Error parsing content for ${promptName}: ${e.toString()}`, batchId);
          failedRequests++;
        }
      } else {
        Logger.log("Invalid response for row " + rowNumber + ": " + JSON.stringify(response));
        addErrorLog(new Date(), rowNumber, "Invalid Response", `Invalid response for ${promptName}`, batchId);
        failedRequests++;
      }
    } catch (e) {
      Logger.log("Error processing result in line " + (i+1) + ": " + e.toString());
      addErrorLog(new Date(), 0, "Processing Error", `Error processing batch result: ${e.toString()}`, batchId);
      failedRequests++;
    }
  }
  
  // Add cost summary entries
  var startTime = new Date();
  var endTime = new Date();
  
  for (var promptName in promptMetrics) {
    var metrics = promptMetrics[promptName];
    addPromptSummary(
      startTime,
      endTime,
      0, // We don't have duration for batch requests
      promptName + " (Batch)",
      metrics.count,
      metrics.inputTokens,
      metrics.outputTokens,
      metrics.cost
    );
  }
  
  Logger.log("Finished processing output file - Total: " + totalRequests + ", Success: " + successfulRequests + ", Failed: " + failedRequests);
  
  return {
    total: totalRequests,
    success: successfulRequests,
    failed: failedRequests
  };
}
  
/**
 * Shows an alert only if debug mode is enabled, or always for errors
 * @param {string} title - The alert title
 * @param {string} message - The alert message
 * @param {ButtonSet} buttons - The buttons to display
 * @param {boolean} isError - Whether this is an error message
 */
function showDebugAlert(title, message, buttons, isError = false) {
  // Always show error messages as popups
  if (isError) {
    SpreadsheetApp.getUi().alert(title, message, buttons);
    Logger.log(`ERROR - ${title}: ${message}`);
    return;
  }
  
  // For non-error messages, only show/log if debug mode is enabled
  if (isDebugModeEnabled()) {
    SpreadsheetApp.getUi().alert(title, message, buttons);
    Logger.log(`${title}: ${message}`);
  }
  // No logging if debug mode is disabled and it's not an error
}
  
/**
 * Logs a message only if debug mode is enabled
 * @param {string} message - The message to log
 */
function debugLog(message) {
  if (isDebugModeEnabled()) {
    Logger.log(message);
  }
}
  
/**
 * Adds an entry to the Execution Log sheet
 * @param {Date} timestamp - When the execution occurred
 * @param {number} row - The row number in the Data sheet
 * @param {string} model - The model used
 * @param {string} promptName - The name of the prompt
 * @param {string} responseContent - The response content received
 * @param {number} inputTokens - Number of input tokens
 * @param {number} outputTokens - Number of output tokens
 * @param {number} totalTokens - Total tokens used
 * @param {number} cost - The cost in USD
 */
function addExecutionLog(timestamp, row, model, promptName, responseContent, inputTokens, outputTokens, totalTokens, cost) {
  var logSheet = getSheet('Execution Log');
  
  // Add headers if the sheet is empty
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow([
      "Timestamp",
      "Row",
      "Model",
      "Prompt Sent",
      "Response Received",
      "Input Tokens",
      "Output Tokens",
      "Total Tokens",
      "Cost (USD)"
    ]);
    
    // Format the header row
    logSheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    logSheet.setFrozenRows(1);
  }
  
  // Add the log entry
  logSheet.appendRow([
    timestamp,
    row,
    model,
    promptName,
    responseContent,
    inputTokens,
    outputTokens,
    totalTokens,
    cost
  ]);
}
  
/**
 * Adds an entry to the Error Log sheet
 * @param {Date} timestamp - When the error occurred
 * @param {number} row - The row number in the Data sheet
 * @param {string} errorType - Type of error
 * @param {string} errorMessage - The error message
 * @param {string} batchId - The batch ID (if applicable)
 */
function addErrorLog(timestamp, row, errorType, errorMessage, batchId = '') {
  var errorSheet = getSheet('Error Log');
  
  // Add headers if the sheet is empty
  if (errorSheet.getLastRow() === 0) {
    errorSheet.appendRow([
      "Timestamp",
      "Row",
      "Error Type",
      "Error Message",
      "Batch ID"
    ]);
    
    // Format the header row
    errorSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    errorSheet.setFrozenRows(1);
  }
  
  // Add the error entry
  errorSheet.appendRow([
    timestamp,
    row,
    errorType,
    errorMessage,
    batchId
  ]);
}
  
/**
 * Monitors changes to the Config sheet
 * @param {Event} e - The event object
 */
function onEditConfig(e) {
  if (!e || !e.source) return;
  
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'Config') return;
  
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  var value = range.getValue();
  
  // Log the change
  Logger.log(`Config sheet edited: Row ${row}, Column ${col}, Value: ${value}`);
  
  // If this is a key cell (column 1)
  if (col === 1) {
    Logger.log(`Config key changed to: ${value}`);
  }
  
  // If this is a value cell (column 2)
  if (col === 2) {
    // Get the key
    var key = sheet.getRange(row, 1).getValue();
    Logger.log(`Config value changed for key "${key}": ${value}`);
  }
}