/*******************************
 * Google Apps Script for OpenAI API Integration
 * Made for you by https://www.zyxware.com
 * Features:
 * - Batch Processing Capability
 * - Executes OpenAI prompts on Google Sheets data
 * - Saves results back to the Data sheet
 * - Logs execution (tokens, costs)
 *********************************/

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
    SpreadsheetApp.getUi()
      .createMenu('OpenAI Tools')
      .addItem('Run for First 10 Rows', 'runPromptsForFirst10Rows')
      .addItem('Run for All Rows', 'runPromptsForAllRows')
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Batch Processing')
        .addItem('Create Batch for First 10 Rows', 'createBatchForFirst10Rows')
        .addItem('Create Batch for All Rows', 'createBatchForAllRows')
        .addItem('Check Batch Status', 'checkBatchStatus')
        .addItem('Process Completed Batch', 'processCompletedBatch')
        .addItem('Cancel Current Batch', 'cancelCurrentBatch')
        .addItem('List All Batches', 'listAllBatches'))
      .addToUi();
  }
  
  function runPromptsForFirst10Rows() {
    runPrompts(10);
  }
  
  function runPromptsForAllRows() {
    runPrompts(Infinity);
  }
  
  function createBatchForFirst10Rows() {
    createBatch(10);
  }
  
  function createBatchForAllRows() {
    createBatch(Infinity);
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
  
  function isDebugMode() {
    return getSheet('Config').getRange('B3').getValue().toLowerCase() === "on";
  }
  
  function getBatchSize() {
    var configSheet = getSheet('Config');
    // Check if the batch size configuration exists
    if (configSheet.getRange('A4').getValue() !== "Batch Size") {
      // Add the configuration if it doesn't exist
      configSheet.getRange('A4').setValue("Batch Size");
      configSheet.getRange('B4').setValue(5000);
    }
    
    var batchSize = parseInt(configSheet.getRange('B4').getValue());
    return isNaN(batchSize) || batchSize <= 0 ? 5000 : batchSize;
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
    // Record start time for this execution
    var startTime = new Date();
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = getSheet('Data');
    var apiKey = getApiKey();
    var debugMode = isDebugMode();
    var lastRow = dataSheet.getLastRow();
    var dataRange = dataSheet.getRange(1, 1, lastRow, dataSheet.getLastColumn()).getValues();
    var headers = dataRange[0];
  
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
  
    // Get only active prompts instead of all prompts
    var prompts = getActivePrompts();
    var rowsToProcess = Math.min(lastRow, 1 + maxRows);
  
    // Track metrics per prompt type
    var promptMetrics = {};
    var rowsExecuted = 0;
  
    for (var rowIndex = 1; rowIndex < rowsToProcess; rowIndex++) {
      if (dataRange[rowIndex][statusColIndex] > 0) continue;
      
      rowsExecuted++;
      
      // Set Batch ID to 0 for non-batch processing
      dataSheet.getRange(rowIndex + 1, batchIdColIndex + 1).setValue("0");
      
      for (var p of prompts) {
        var promptName = p[0];
        var promptText = p[1];
        var model = (p[2] || getDefaultModel()).toLowerCase();
  
        var finalPrompt = replacePlaceholders(promptText, headers, dataRange[rowIndex]);
  
        try {
          // Record start time for this API call
          var apiCallStartTime = new Date();
          
          var apiResponse = callOpenAI(apiKey, model, finalPrompt);
          
          // Record end time and calculate duration for this API call
          var apiCallEndTime = new Date();
          var apiCallDuration = (apiCallEndTime - apiCallStartTime) / 1000; // Duration in seconds
          
          var responseText = apiResponse.text;
          var parsedResponse = apiResponse.parsedJson;
          var inputTokens = apiResponse.inputTokens;
          var outputTokens = apiResponse.outputTokens;
          var totalTokens = apiResponse.totalTokens;
          var cost = calculateCost(model, inputTokens, outputTokens);
          
          // Update per-prompt metrics
          if (!promptMetrics[promptName]) {
            promptMetrics[promptName] = {
              count: 0,
              inputTokens: 0,
              outputTokens: 0,
              totalTokens: 0,
              cost: 0,
              model: model,
              duration: 0 // Add duration tracking
            };
          }
          promptMetrics[promptName].count++;
          promptMetrics[promptName].inputTokens += inputTokens;
          promptMetrics[promptName].outputTokens += outputTokens;
          promptMetrics[promptName].totalTokens += totalTokens;
          promptMetrics[promptName].cost += cost;
          promptMetrics[promptName].duration += apiCallDuration; // Add this API call's duration to the total
  
          // Save response back to the Data sheet
          saveResponseToDataSheet(dataSheet, headers, rowIndex, parsedResponse, promptName);
  
          if (debugMode) {
            logExecution(rowIndex + 1, finalPrompt, responseText, inputTokens, outputTokens, totalTokens, model);
          }
  
        } catch (e) {
          logError(rowIndex + 1, e.toString());
        }
      }
  
      // When processing is complete, set status to 1 for non-batch processing
      dataSheet.getRange(rowIndex + 1, statusColIndex + 1).setValue(1);
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
          metrics.duration, // Use the accumulated duration for this prompt
          promptName, 
          rowsExecuted, 
          metrics.inputTokens, 
          metrics.outputTokens, 
          metrics.cost
        );
      }
      
      // Calculate totals for toast notification
      var totalPromptsExecuted = 0;
      var totalInputTokens = 0;
      var totalOutputTokens = 0;
      var totalCost = 0;
      var totalDuration = 0;
      
      for (var promptName in promptMetrics) {
        totalPromptsExecuted += promptMetrics[promptName].count;
        totalInputTokens += promptMetrics[promptName].inputTokens;
        totalOutputTokens += promptMetrics[promptName].outputTokens;
        totalCost += promptMetrics[promptName].cost;
        totalDuration += promptMetrics[promptName].duration;
      }
      
      // Show a toast notification with the summary
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Executed ${totalPromptsExecuted} prompts\nTotal tokens: ${totalInputTokens + totalOutputTokens}\nTotal cost: $${totalCost.toFixed(4)}\nTotal API time: ${totalDuration.toFixed(1)} seconds\nTotal execution time: ${executionDuration.toFixed(1)} seconds`,
        'Execution Complete',
        30
      );
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
      Logger.log("Error processing response: " + e);
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
  function replacePlaceholders(prompt, headers, rowData) {
    var finalPrompt = prompt;
    for (var colIndex = 0; colIndex < headers.length; colIndex++) {
      var placeholder = '{{' + headers[colIndex].trim() + '}}';
      var value = rowData[colIndex] !== undefined ? rowData[colIndex] : "N/A";
      finalPrompt = finalPrompt.replace(new RegExp(placeholder, 'g'), value);
    }
    return finalPrompt;
  }
  
  /* ======== Calculate OpenAI API Cost ======== */
  function calculateCost(model, inputTokens, outputTokens) {
    var pricing = getPricing(model);
    var inputRate = (pricing.input_per_1m || 0) / 1000000; // Convert per 1M rate to per token
    var outputRate = (pricing.output_per_1m || 0) / 1000000;
  
    var inputCost = inputTokens * inputRate;
    var outputCost = outputTokens * outputRate;
  
    return inputCost + outputCost;
  }
  
  /* ======== Batch Processing Functions ======== */
  
  /**
   * Creates a batch processing job for the specified number of rows
   */
  function createBatch(maxRows) {
    var ui = SpreadsheetApp.getUi();
    var apiKey = getApiKey();
    
    if (!apiKey) {
      ui.alert('Error', 'API key is missing. Please add it to the Config sheet.', ui.ButtonSet.OK);
      return;
    }
    
    try {
      // Get the configured batch size
      var batchSize = getBatchSize();
      
      // Find the next set of rows to process
      var nextBatchInfo = findNextBatchRows(maxRows, batchSize);
      
      if (!nextBatchInfo || nextBatchInfo.startRow > nextBatchInfo.endRow) {
        ui.alert('No Data', 'No more rows to process or all rows are already processed.', ui.ButtonSet.OK);
        return;
      }
      
      // Prepare the batch file for the selected range of rows
      var batchData = prepareBatchDataRange(nextBatchInfo.startRow, nextBatchInfo.endRow);
      if (!batchData || batchData.requests.length === 0) {
        ui.alert('No Data', 'No rows to process in the selected range.', ui.ButtonSet.OK);
        return;
      }
      
      // Check batch size limits
      if (batchData.requests.length > 50000) {
        ui.alert('Batch Too Large', 
                 'This batch contains ' + batchData.requests.length + ' requests, which exceeds the OpenAI limit of 50,000 requests per batch. Please reduce the batch size in Config.',
                 ui.ButtonSet.OK);
        return;
      }
      
      // Create the JSONL content
      var jsonlContent = createJsonlContent(batchData.requests);
      
      // Estimate JSONL file size
      var estimatedSizeMB = jsonlContent.length / (1024 * 1024);
      if (estimatedSizeMB > 200) {
        ui.alert('Batch File Too Large', 
                 'The estimated batch file size is ' + estimatedSizeMB.toFixed(2) + ' MB, which exceeds the OpenAI limit of 200 MB. Please reduce the batch size in Config.',
                 ui.ButtonSet.OK);
        return;
      }
      
      // Upload the file to OpenAI
      var fileId = uploadFileToOpenAI(jsonlContent);
      if (!fileId) {
        ui.alert('Error', 'Failed to upload batch file to OpenAI.', ui.ButtonSet.OK);
        return;
      }
      
      // Create the batch
      var batch = createOpenAIBatch(fileId);
      if (!batch) {
        ui.alert('Error', 'Failed to create batch job.', ui.ButtonSet.OK);
        return;
      }
      
      // Store batch information in the Batch Status sheet
      var batchId = storeBatchInfo(batch, batchData.rowMap);
      
      // Update the Data sheet with batch IDs
      updateDataSheetWithBatchId(batchData.rowIndices, batchId);
      
      ui.alert('Success', 
               `Batch job created successfully!\n\nProcessed rows ${nextBatchInfo.startRow} to ${nextBatchInfo.endRow}\nBatch ID: ${batch.id}\nStatus: ${batch.status}\nTotal Requests: ${batchData.requests.length}\n\n${nextBatchInfo.remainingRows > 0 ? 'There are ' + nextBatchInfo.remainingRows + ' more rows to process. Run "Create Batch" again to process the next set.' : 'All rows have been processed.'}`, 
               ui.ButtonSet.OK);
               
    } catch (e) {
      Logger.log('Error creating batch: ' + e.toString());
      ui.alert('Error', 'Failed to create batch: ' + e.toString(), ui.ButtonSet.OK);
    }
  }
  
  /**
   * Finds the next set of rows to process based on batch size
   */
  function findNextBatchRows(maxRows, batchSize) {
    var dataSheet = getSheet('Data');
    var lastRow = dataSheet.getLastRow();
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
    
    var rowsToProcess = Math.min(lastRow, 1 + maxRows);
    var startRow = -1;
    var endRow = -1;
    var remainingRows = 0;
    
    // Find the first unprocessed row
    for (var rowIndex = 1; rowIndex < rowsToProcess; rowIndex++) {
      // Skip rows that are already processed (status = 1 or 2) or already part of a batch
      if (dataRange[rowIndex][statusColIndex] > 0 || 
          (batchIdColIndex >= 0 && dataRange[rowIndex][batchIdColIndex] && dataRange[rowIndex][batchIdColIndex] !== "0")) {
        continue;
      }
      
      if (startRow === -1) {
        startRow = rowIndex + 1; // +1 because rowIndex is 0-based but sheet rows are 1-based
      }
      
      // If we've found enough rows for this batch, set the end row
      if (rowIndex - (startRow - 1) + 1 >= batchSize) { // +1 to include the current row
        endRow = rowIndex + 1; // +1 to include the current row in sheet coordinates
        break;
      }
    }
    
    // If we didn't find enough rows to fill a batch, use all remaining rows
    if (startRow !== -1 && endRow === -1) {
      endRow = rowsToProcess; // Use the actual last row to process (not -1)
    }
    
    // Count remaining rows after this batch
    if (endRow !== -1 && endRow < rowsToProcess) {
      for (var rowIndex = endRow; rowIndex < rowsToProcess; rowIndex++) { // Start from endRow (not +1)
        if (dataRange[rowIndex][statusColIndex] === 0 && 
            (!dataRange[rowIndex][batchIdColIndex] || dataRange[rowIndex][batchIdColIndex] === "0")) {
          remainingRows++;
        }
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
    var dataSheet = getSheet('Data');
    var lastRow = dataSheet.getLastRow();
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
    
    // Get only active prompts instead of all prompts
    var prompts = getActivePrompts();
    
    var requests = [];
    var rowMap = {}; // Simplified row map
    var rowIndices = []; // Stores row indices that are part of this batch
    
    // Process only rows in the specified range
    for (var rowIndex = startRow - 1; rowIndex < endRow; rowIndex++) {
      // Skip rows that are already processed or already part of a batch
      if (dataRange[rowIndex][statusColIndex] > 0 || 
          (batchIdColIndex >= 0 && dataRange[rowIndex][batchIdColIndex] && dataRange[rowIndex][batchIdColIndex] !== "0")) {
        continue;
      }
      
      // Add this row to the list of rows in this batch
      rowIndices.push(rowIndex + 1);
      
      for (var p = 0; p < prompts.length; p++) {
        var promptName = prompts[p][0];
        var promptText = prompts[p][1];
        var model = (prompts[p][2] || getDefaultModel()).toLowerCase();
        
        var finalPrompt = replacePlaceholders(promptText, headers, dataRange[rowIndex]);
        
        // Create a unique custom_id for this request
        var customId = `row-${rowIndex+1}-prompt-${p+1}`;
        
        // Simplified row map - just store the essential information
        rowMap[customId] = {
          row: rowIndex + 1,
          promptName: promptName
        };
        
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
      rowMap: rowMap,
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
  function storeBatchInfo(batch, rowMap) {
    var batchStatusSheet = getSheet('Batch Status');
    
    // Initialize headers if sheet is empty
    if (batchStatusSheet.getLastRow() === 0) {
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
        "Row Mapping"
      ]);
    }
    
    // Generate a unique batch ID for our system
    var batchId = Utilities.getUuid();
    
    // Format the creation date
    var createdDate = new Date(batch.created_at * 1000).toISOString();
    var currentDate = new Date().toISOString();
    
    // Store the row mapping as JSON
    var rowMapJson = JSON.stringify(rowMap);
    
    batchStatusSheet.appendRow([
      batchId,
      batch.id,
      batch.status,
      createdDate,
      currentDate,
      batch.input_file_id,
      batch.output_file_id || "",
      batch.error_file_id || "",
      batch.request_counts.total,
      batch.request_counts.completed,
      batch.request_counts.failed,
      rowMapJson
    ]);
    
    return batchId;
  }
  
  /**
   * Checks the status of the most recent batch
   */
  function checkBatchStatus() {
    var ui = SpreadsheetApp.getUi();
    var batchStatusSheet = getSheet('Batch Status');
    
    if (batchStatusSheet.getLastRow() <= 1) {
      ui.alert('No Batches', 'No batch jobs have been created yet.', ui.ButtonSet.OK);
      return;
    }
    
    // Get all batches
    var batchData = batchStatusSheet.getDataRange().getValues();
    var headers = batchData[0];
    
    // Find column indices
    var batchIdColIndex = headers.indexOf("Batch ID");
    var openAIBatchIdColIndex = headers.indexOf("OpenAI Batch ID");
    var statusColIndex = headers.indexOf("Status");
    var lastCheckedColIndex = headers.indexOf("Last Checked At");
    var outputFileIdColIndex = headers.indexOf("Output File ID");
    
    if (batchIdColIndex < 0 || openAIBatchIdColIndex < 0 || statusColIndex < 0) {
      ui.alert('Error', 'Batch Status sheet has invalid format.', ui.ButtonSet.OK);
      return;
    }
    
    // Find batches that are not in a final state
    var activeBatches = [];
    for (var i = 1; i < batchData.length; i++) {
      var status = batchData[i][statusColIndex];
      if (status !== 'completed' && status !== 'failed' && status !== 'expired' && status !== 'cancelled') {
        activeBatches.push({
          rowIndex: i + 1,
          batchId: batchData[i][batchIdColIndex],
          openAIBatchId: batchData[i][openAIBatchIdColIndex]
        });
      }
    }
    
    if (activeBatches.length === 0) {
      ui.alert('No Active Batches', 'There are no active batches to check. All batches are in a final state.', ui.ButtonSet.OK);
      return;
    }
    
    // Check each active batch
    var updatedBatches = 0;
    var completedBatches = 0;
    
    for (var i = 0; i < activeBatches.length; i++) {
      var batchInfo = activeBatches[i];
      
      try {
        var batch = retrieveBatch(batchInfo.openAIBatchId);
        if (!batch) continue;
        
        // Update the batch information in the sheet
        batchStatusSheet.getRange(batchInfo.rowIndex, statusColIndex + 1).setValue(batch.status);
        batchStatusSheet.getRange(batchInfo.rowIndex, lastCheckedColIndex + 1).setValue(new Date().toISOString());
        batchStatusSheet.getRange(batchInfo.rowIndex, outputFileIdColIndex + 1).setValue(batch.output_file_id || "");
        
        updatedBatches++;
        
        if (batch.status === 'completed') {
          completedBatches++;
        }
      } catch (e) {
        Logger.log('Error checking batch status: ' + e.toString());
      }
    }
    
    if (completedBatches > 0) {
      var response = ui.alert('Batches Ready', 
                             `${completedBatches} batch(es) are complete and ready to process. Would you like to process them now?`, 
                             ui.ButtonSet.YES_NO);
                             
      if (response === ui.Button.YES) {
        processCompletedBatches();
      }
    } else {
      ui.alert('Batch Status Updated', 
               `Updated status for ${updatedBatches} batch(es). No batches are ready for processing yet.`, 
               ui.ButtonSet.OK);
    }
  }
  
  /**
   * Processes all completed batches
   */
  function processCompletedBatches() {
    var ui = SpreadsheetApp.getUi();
    var batchStatusSheet = getSheet('Batch Status');
    
    if (batchStatusSheet.getLastRow() <= 1) {
      ui.alert('No Batches', 'No batch jobs have been created yet.', ui.ButtonSet.OK);
      return;
    }
    
    // Get all batches
    var batchData = batchStatusSheet.getDataRange().getValues();
    var headers = batchData[0];
    
    // Find column indices
    var batchIdColIndex = headers.indexOf("Batch ID");
    var statusColIndex = headers.indexOf("Status");
    var outputFileIdColIndex = headers.indexOf("Output File ID");
    var rowMappingColIndex = headers.indexOf("Row Mapping");
    
    if (batchIdColIndex < 0 || statusColIndex < 0 || outputFileIdColIndex < 0 || rowMappingColIndex < 0) {
      ui.alert('Error', 'Batch Status sheet has invalid format.', ui.ButtonSet.OK);
      return;
    }
    
    // Find completed batches that haven't been processed
    var completedBatches = [];
    for (var i = 1; i < batchData.length; i++) {
      if (batchData[i][statusColIndex] === 'completed' && batchData[i][outputFileIdColIndex]) {
        completedBatches.push({
          rowIndex: i + 1,
          batchId: batchData[i][batchIdColIndex],
          outputFileId: batchData[i][outputFileIdColIndex],
          rowMapping: batchData[i][rowMappingColIndex]
        });
      }
    }
    
    if (completedBatches.length === 0) {
      ui.alert('No Completed Batches', 'There are no completed batches to process.', ui.ButtonSet.OK);
      return;
    }
    
    // Process each completed batch
    var totalProcessed = 0;
    var totalSuccess = 0;
    var totalFailed = 0;
    
    for (var i = 0; i < completedBatches.length; i++) {
      var batchInfo = completedBatches[i];
      
      try {
        // Download the output file
        var outputContent = downloadFileFromOpenAI(batchInfo.outputFileId);
        if (!outputContent) continue;
        
        // Parse the row mapping
        var rowMap = JSON.parse(batchInfo.rowMapping);
        
        // Process the results
        var results = processOutputFile(outputContent, rowMap, batchInfo.batchId);
        
        totalProcessed += results.total;
        totalSuccess += results.success;
        totalFailed += results.failed;
        
      } catch (e) {
        Logger.log('Error processing batch: ' + e.toString());
      }
    }
    
    // Show summary
    ui.alert('Batch Processing Complete', 
             `Successfully processed ${totalSuccess} out of ${totalProcessed} requests.\n${totalFailed} requests failed.`, 
             ui.ButtonSet.OK);
  }
  
  /**
   * Processes the output file and updates the Data sheet
   */
  function processOutputFile(outputContent, rowMap, batchId) {
    var dataSheet = getSheet('Data');
    var headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
    
    // Find Status and Batch ID columns
    var statusColIndex = headers.indexOf("Status");
    var batchIdColIndex = headers.indexOf("Batch ID");
    
    var lines = outputContent.split('\n').filter(line => line.trim()); // Filter out empty lines
    
    // Calculate the actual number of requests from the row mapping
    var totalRequests = Object.keys(rowMap).length;
    var successfulRequests = 0;
    var failedRequests = 0;
    
    // Track metrics for cost summary
    var promptMetrics = {};
    
    for (var i = 0; i < lines.length; i++) {
      if (!lines[i].trim()) continue;
      
      try {
        var result = JSON.parse(lines[i]);
        var customId = result.custom_id;
        var rowInfo = rowMap[customId];
        
        if (!rowInfo) {
          Logger.log('Warning: No row mapping found for custom_id: ' + customId);
          continue;
        }
        
        if (result.error) {
          // Log the error
          logError(rowInfo.row, `Batch error for ${rowInfo.promptName}: ${result.error.message}`);
          failedRequests++;
          continue;
        }
        
        var response = result.response;
        if (response && response.status_code === 200 && response.body) {
          var responseBody = response.body;
          var content = responseBody.choices[0].message.content;
          
          try {
            var parsedContent = JSON.parse(content);
            
            // Save the response to the Data sheet
            saveResponseToDataSheet(dataSheet, headers, rowInfo.row - 1, parsedContent, rowInfo.promptName);
            
            // Mark the row as processed (status = 2 for batch completed)
            if (statusColIndex >= 0) {
              dataSheet.getRange(rowInfo.row, statusColIndex + 1).setValue(2);
            }
            
            // Verify the batch ID matches
            if (batchIdColIndex >= 0) {
              var currentBatchId = dataSheet.getRange(rowInfo.row, batchIdColIndex + 1).getValue();
              if (currentBatchId !== batchId) {
                Logger.log(`Warning: Batch ID mismatch for row ${rowInfo.row}. Expected: ${batchId}, Found: ${currentBatchId}`);
              }
            }
            
            // Track usage for cost summary
            var model = responseBody.model;
            var inputTokens = responseBody.usage.prompt_tokens;
            var outputTokens = responseBody.usage.completion_tokens;
            var totalTokens = responseBody.usage.total_tokens;
            var cost = calculateCost(model, inputTokens, outputTokens);
            
            if (!promptMetrics[rowInfo.promptName]) {
              promptMetrics[rowInfo.promptName] = {
                count: 0,
                inputTokens: 0,
                outputTokens: 0,
                totalTokens: 0,
                cost: 0,
                model: model,
                duration: 0 // We don't have individual durations for batch requests
              };
            }
            
            promptMetrics[rowInfo.promptName].count++;
            promptMetrics[rowInfo.promptName].inputTokens += inputTokens;
            promptMetrics[rowInfo.promptName].outputTokens += outputTokens;
            promptMetrics[rowInfo.promptName].totalTokens += totalTokens;
            promptMetrics[rowInfo.promptName].cost += cost;
            
            successfulRequests++;
          } catch (e) {
            Logger.log('Error parsing content: ' + e.toString());
            logError(rowInfo.row, `Error parsing content for ${rowInfo.promptName}: ${e.toString()}`);
            failedRequests++;
          }
        } else {
          Logger.log('Invalid response: ' + JSON.stringify(response));
          logError(rowInfo.row, `Invalid response for ${rowInfo.promptName}`);
          failedRequests++;
        }
      } catch (e) {
        Logger.log('Error processing result: ' + e.toString());
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
        metrics.cost * 0.5 // 50% discount for batch processing
      );
    }
    
    return {
      total: totalRequests,
      success: successfulRequests,
      failed: failedRequests
    };
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
   * Lists all batches
   */
  function listAllBatches() {
    var ui = SpreadsheetApp.getUi();
    
    try {
      var batches = fetchAllBatches();
      
      if (batches.length === 0) {
        ui.alert('No Batches', 'No batch jobs were found.', ui.ButtonSet.OK);
        return;
      }
      
      // Create a sheet to display the batches
      var batchListSheet = getSheet('Batch List');
      batchListSheet.clear();
      
      // Add headers
      batchListSheet.appendRow([
        "Batch ID", 
        "Status", 
        "Created At", 
        "Total Requests", 
        "Completed", 
        "Failed"
      ]);
      
      // Add batch data
      for (var i = 0; i < batches.length; i++) {
        var batch = batches[i];
        var createdDate = new Date(batch.created_at * 1000).toISOString();
        
        batchListSheet.appendRow([
          batch.id,
          batch.status,
          createdDate,
          batch.request_counts.total,
          batch.request_counts.completed,
          batch.request_counts.failed
        ]);
      }
      
      // Format the sheet
      batchListSheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      batchListSheet.autoResizeColumns(1, 6);
      
      // Activate the sheet
      SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(batchListSheet);
      
      ui.alert('Batch List', `Found ${batches.length} batches.`, ui.ButtonSet.OK);
      
    } catch (e) {
      Logger.log('Error listing batches: ' + e.toString());
      ui.alert('Error', 'Failed to list batches: ' + e.toString(), ui.ButtonSet.OK);
    }
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
   * Creates JSONL content from request objects
   */
  function createJsonlContent(requests) {
    return requests.map(function(req) {
      return JSON.stringify(req);
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
   * Cancels a batch in OpenAI
   */
  function cancelBatch(batchId) {
    var apiKey = getApiKey();
    
    var options = {
      method: 'post',
      headers: { Authorization: 'Bearer ' + apiKey },
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(`https://api.openai.com/v1/batches/${batchId}/cancel`, options);
    var responseJson = JSON.parse(response.getContentText());
    
    if (responseJson.error) {
      throw new Error('OpenAI API error: ' + responseJson.error.message);
    }
    
    return responseJson;
  }
  
  /**
   * Processes a completed batch
   */
  function processCompletedBatch() {
    processCompletedBatches();
  }
  
  /**
   * Cancels the current batch
   */
  function cancelCurrentBatch() {
    var ui = SpreadsheetApp.getUi();
    var batchStatusSheet = getSheet('Batch Status');
    
    if (batchStatusSheet.getLastRow() <= 1) {
      ui.alert('No Batches', 'No batch jobs have been created yet.', ui.ButtonSet.OK);
      return;
    }
    
    var lastRow = batchStatusSheet.getLastRow();
    var batchData = batchStatusSheet.getRange(lastRow, 1, 1, 12).getValues()[0];
    var batchId = batchData[1]; // OpenAI Batch ID is in column 2
    var status = batchData[2]; // Status is in column 3
    
    if (status === 'completed' || status === 'failed' || status === 'expired' || status === 'cancelled') {
      ui.alert('Batch Already Finished', 
               `This batch is already in a final state: ${status}. It cannot be cancelled.`, 
               ui.ButtonSet.OK);
      return;
    }
    
    var response = ui.alert('Confirm Cancellation', 
                           'Are you sure you want to cancel this batch? This action cannot be undone.', 
                           ui.ButtonSet.YES_NO);
                           
    if (response !== ui.Button.YES) {
      return;
    }
    
    try {
      var result = cancelBatch(batchId);
      
      // Update the status in the sheet
      batchStatusSheet.getRange(lastRow, 3).setValue(result.status);
      
      ui.alert('Batch Cancelled', 
               `The batch has been cancelled. Status: ${result.status}`, 
               ui.ButtonSet.OK);
               
    } catch (e) {
      Logger.log('Error cancelling batch: ' + e.toString());
      ui.alert('Error', 'Failed to cancel batch: ' + e.toString(), ui.ButtonSet.OK);
    }
  }
  
  
  
