/*******************************
 * Google Apps Script for OpenAI API Integration
 * Made for you by https://www.zyxware.com
 * Features:
 * - Executes OpenAI prompts on Google Sheets data
 * - Saves results back to the Data sheet
 * - Logs execution (tokens, costs)
 *********************************/

// OpenAI Pricing Configuration
const PRICING_CONFIG = {
    "gpt-4.5-preview": { "input_per_1m": 75.00, "cached_input_per_1m": 37.50, "output_per_1m": 150.00 },
    "gpt-4o": { "input_per_1m": 2.50, "cached_input_per_1m": 1.25, "output_per_1m": 10.00 },
    "gpt-4o-mini": { "input_per_1m": 0.15, "cached_input_per_1m": 0.075, "output_per_1m": 0.60 },
    "o3-mini": { "input_per_1m": 1.10, "cached_input_per_1m": 0.55, "output_per_1m": 4.40 },
    "o1-mini": { "input_per_1m": 1.10, "cached_input_per_1m": 0.55, "output_per_1m": 4.40 }
  };
  
  /* ======== UI Functions ======== */
  function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu('OpenAI Tools')
      .addItem('Run for First 10 Rows', 'runPromptsForFirst10Rows')
      .addItem('Run for All Rows', 'runPromptsForAllRows')
      .addToUi();
  }
  
  function runPromptsForFirst10Rows() {
    runPrompts(10);
  }
  
  function runPromptsForAllRows() {
    runPrompts(Infinity);
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
  
  function isDebugMode() {
    return getSheet('Config').getRange('B3').getValue().toLowerCase() === "on";
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
  
    var promptsSheet = getSheet('Prompts');
    var prompts = promptsSheet.getDataRange().getValues().slice(1);
    var rowsToProcess = Math.min(lastRow, 1 + maxRows);
  
    // Track metrics per prompt type
    var promptMetrics = {};
    var rowsExecuted = 0;
  
    for (var rowIndex = 1; rowIndex < rowsToProcess; rowIndex++) {
      if (dataRange[rowIndex][statusColIndex] == 1) continue;
      
      rowsExecuted++;
      
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
  
  
  
