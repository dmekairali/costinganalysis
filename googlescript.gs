
/**
 * Google Apps Script Backend for Advanced Costing Review Dashboard
 * Handles OpenAI integration, Google Sheets operations, and data processing
 */

const CONFIG = {
  // API key is fetched from Script Properties for security
  OPENAI_API_KEY: PropertiesService.getScriptProperties().getProperty('openai_key'),
  
  // Google Sheets configuration
  SHEET_ID: '1J17ADExIvGd8WNFODQtouhv9kWnU9aPzPnaMbazmwtg',
  DATA_SHEET_NAME: 'SFG REVIEW SYSTEM',
  
  // Row configuration
  HEADER_ROW: 7,        // Row number where headers are located
  DATA_START_ROW: 8,    // Row number where data starts
  
  // Column mapping - specify which columns contain your data
  COLUMN_MAPPING: {
    productName: 'G',      // or use column index: 1
    productionId: 'B',     // or use column index: 2
    package: 'D',          // or use column index: 3
    qty: 'N',              // or use column index: 4
    notes: 'BJ',           // Notes column
    
    // CA Cost columns
    caCorrect: 'CX',        // Column for CA Correct value
    caBenchmark: 'CY',      // Column for CA Benchmark cost
    caLast3Prod: 'CZ',      // Column for CA Last 3 Production avg
    caLast3Month: 'DA',     // Column for CA Last 3 Month avg
    caLast12Month: 'DB',    // Column for CA Last 12 Month avg
    caNewBenchmark: 'CX',   // Column for CA New Benchmark
    caStatus: 'DD',         // Column for CA Status
    
    // TOC Cost columns
    tocCorrect: 'DE',
    tocBenchmark: 'DF',
    tocLast3Prod: 'DG',
    tocLast3Month: 'DH',
    tocLast12Month: 'DI',
    tocNewBenchmark: 'DE',
    tocStatus: 'DK',

    // Time columns
    timeCorrect: 'DL',
    timeBenchmark: 'DM',
    timeLast3Prod: 'DN',
    timeLast3Month: 'DO',
    timeLast12Month: 'DP',
    timeNewBenchmark: 'DL',
    timeStatus: 'DR'
  },
  
  // OpenAI configuration
  OPENAI_MODEL: 'gpt-4o',
  OPENAI_API_URL: 'https://api.openai.com/v1/chat/completions'
};

function columnLetterToIndex(letter) {
  let column = 0, length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}



/**
 * Main function to serve the HTML page
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Include HTML files (CSS and JS)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * Simplified import function with direct filtering
 */
function importCostingData() {
  try {
    const sheet = getDataSheet();
    
    // Get the data range
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    console.log(`Sheet dimensions: ${lastRow} rows, ${lastCol} columns`);
    
    if (lastRow < CONFIG.DATA_START_ROW) {
      console.log('No data rows found');
      return [];
    }
    
    // Get headers if they exist
    let headers = [];
    if (CONFIG.HEADER_ROW && CONFIG.HEADER_ROW <= lastRow && CONFIG.HEADER_ROW > 0) {
      headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
      console.log('Headers found:', headers.slice(0, 10));
    }
    
    // Calculate data range dimensions
    const dataRows = lastRow - CONFIG.DATA_START_ROW + 1;
    
    if (dataRows <= 0) {
      console.log('No data rows to process');
      return [];
    }
    
    // Get data starting from DATA_START_ROW
    console.log(`Getting data range: Row ${CONFIG.DATA_START_ROW}, Col 1, ${dataRows} rows, ${lastCol} cols`);
    const dataRange = sheet.getRange(CONFIG.DATA_START_ROW, 1, dataRows, lastCol);
    const rawData = dataRange.getValues();
    
    console.log(`Processing ${rawData.length} raw data rows`);
    
    const costingData = [];
    let skippedRows = 0;
    
    for (let i = 0; i < rawData.length; i++) {
      const row = rawData[i];
      const rowIndex = CONFIG.DATA_START_ROW + i;
      
      try {
        // Extract data using column mapping
        const extractedData = extractRowData(row, CONFIG.COLUMN_MAPPING, headers);
        
        // Apply direct filter condition: B is not null AND CU is not null AND CV is null
        const columnB = row[1]; // Column B (0-based index = 1)
        const columnCU = row[98]; // Column CU (0-based index = 98)
        const columnCV = row[99]; // Column CV (0-based index = 99)
        
        const conditionB = columnB !== null && columnB !== undefined && columnB !== '';
        const conditionCU = columnCU !== null && columnCU !== undefined && columnCU !== '';
        const conditionCV = columnCV === null || columnCV === undefined || columnCV === '';
        
        if (!(conditionB && conditionCU && conditionCV)) {
          skippedRows++;
          continue;
        }
        
        // Create the item object
        const item = createItemFromData(extractedData);
        costingData.push(item);
        
      } catch (error) {
        console.error(`Error processing row ${rowIndex}:`, error.message);
        skippedRows++;
      }
    }
    
    console.log(`Import completed: ${costingData.length} items imported, ${skippedRows} rows skipped`);
    
    return costingData;
    
  } catch (error) {
    console.error('Error importing data:', error);
    throw new Error('Failed to import data from Google Sheets: ' + error.message);
  }
}

/**
 * Extract data from row using column mapping
 */
function extractRowData(row, columnMapping, headers) {
  const data = {};
  
  for (const [key, columnRef] of Object.entries(columnMapping)) {
    let columnIndex;
    
    if (typeof columnRef === 'string') {
      // Handle column letter (A, B, C, etc.)
      if (columnRef.match(/^[A-Z]+$/)) {
        columnIndex = columnLetterToIndex(columnRef) - 1; // Convert to 0-based index
      } else {
        // Handle column name lookup in headers
        columnIndex = headers.findIndex(header => 
          header.toString().toLowerCase().includes(columnRef.toLowerCase())
        );
      }
    } else if (typeof columnRef === 'number') {
      // Handle direct column index (1-based to 0-based)
      columnIndex = columnRef - 1;
    }
    
    if (columnIndex >= 0 && columnIndex < row.length) {
      data[key] = row[columnIndex];
    } else {
      console.warn(`Column mapping failed for ${key}: ${columnRef}`);
      data[key] = null;
    }
  }
  
  return data;
}

/**
 * Convert column letter to index (A=1, B=2, etc.)
 */
function columnLetterToIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return index;
}


/**
 * Create item object from extracted data
 */
function createItemFromData(data) {
  return {
    productName: (data.productName || '').toString().trim(),
    productionId: (data.productionId || '').toString().trim(),
    package: (data.package || '').toString().trim(),
    qty: parseInt(data.qty) || 0,
    notes: (data.notes || '').toString().trim(),
    ca: {
      correct: parseFloat(data.caCorrect) || 0,
      benchmark: parseFloat(data.caBenchmark) || 0,
      last3Prod: parseFloat(data.caLast3Prod) || 0,
      last3Month: parseFloat(data.caLast3Month) || 0,
      last12Month: parseFloat(data.caLast12Month) || 0,
      newBenchmark: parseFloat(data.caNewBenchmark) || 0,
      status: normalizeStatus(data.caStatus)
    },
    toc: {
      correct: parseFloat(data.tocCorrect) || 0,
      benchmark: parseFloat(data.tocBenchmark) || 0,
      last3Prod: parseFloat(data.tocLast3Prod) || 0,
      last3Month: parseFloat(data.tocLast3Month) || 0,
      last12Month: parseFloat(data.tocLast12Month) || 0,
      newBenchmark: parseFloat(data.tocNewBenchmark) || 0,
      status: normalizeStatus(data.tocStatus)
    },
    time: {
      correct: normalizeTime(data.timeCorrect),
      benchmark: normalizeTime(data.timeBenchmark),
      last3Prod: normalizeTime(data.timeLast3Prod),
      last3Month: normalizeTime(data.timeLast3Month),
      last12Month: normalizeTime(data.timeLast12Month),
      newBenchmark: normalizeTime(data.timeNewBenchmark),
      status: normalizeStatus(data.timeStatus)
    }
  };
}

/**
 * Normalize status values
 */
function normalizeStatus(status) {
  const statusStr = (status || '').toString().toLowerCase().trim();
  
  if (statusStr.includes('no need')) return 'no_need';
  if (statusStr.includes('update benchmark cost')) return 'update';
  
  return 'no_need'; // Default to 'update'
}

/**
 * Normalize time values
 */
function normalizeTime(timeValue) {
  if (!timeValue) return '0:00:00';

  // If the value is a Date object, format it directly
  if (timeValue instanceof Date) {
    // Format the date object to HH:mm:ss. Using "UTC" is important to avoid timezone shifts.
    return Utilities.formatDate(timeValue, "UTC", "HH:mm:ss");
  }
  
  const timeStr = timeValue.toString().trim();
  
  // If it's already in HH:MM:SS format
  if (timeStr.match(/^\d{1,2}:\d{2}:\d{2}$/)) {
    return timeStr;
  }
  
  // If it's in HH:MM format
  if (timeStr.match(/^\d{1,2}:\d{2}$/)) {
    return timeStr + ':00';
  }
  
  // If it's a decimal number (hours)
  const hours = parseFloat(timeStr);
  if (!isNaN(hours)) {
    const h = Math.floor(hours);
    const m = Math.floor((hours - h) * 60);
    const s = Math.floor(((hours - h) * 60 - m) * 60);
    return `${h}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
  }
  
  return '0:00:00';
}



/**
 * Save costing data to Google Sheets by updating individual cells.
 * Note: This method can be slow if updating a large number of rows simultaneously,
 * as it makes multiple write calls to the Google Sheets API.
 */
function saveCostingData(updatedCostingData) {
  try {
    const sheet = getDataSheet();
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    // Create a map of production IDs to row index for faster lookups
    const productionIdToRowIndex = {};
    for (let i = CONFIG.HEADER_ROW; i < values.length; i++) {
      const productionId = values[i][columnLetterToIndex(CONFIG.COLUMN_MAPPING.productionId) - 1];
      if (productionId) {
        productionIdToRowIndex[productionId] = i + 1; // Use 1-based index for getRange
      }
    }

    updatedCostingData.forEach(item => {
      const rowIndex = productionIdToRowIndex[item.productionId];
      if (rowIndex !== undefined) {
        // Update the 6 specific columns for the given row
        sheet.getRange(rowIndex, columnLetterToIndex(CONFIG.COLUMN_MAPPING.caNewBenchmark)).setValue(item.ca.newBenchmark);
        sheet.getRange(rowIndex, columnLetterToIndex(CONFIG.COLUMN_MAPPING.caStatus)).setValue(item.ca.status);
        sheet.getRange(rowIndex, columnLetterToIndex(CONFIG.COLUMN_MAPPING.tocNewBenchmark)).setValue(item.toc.newBenchmark);
        sheet.getRange(rowIndex, columnLetterToIndex(CONFIG.COLUMN_MAPPING.tocStatus)).setValue(item.toc.status);
        sheet.getRange(rowIndex, columnLetterToIndex(CONFIG.COLUMN_MAPPING.timeNewBenchmark)).setValue(item.time.newBenchmark);
        sheet.getRange(rowIndex, columnLetterToIndex(CONFIG.COLUMN_MAPPING.timeStatus)).setValue(item.time.status);
      }
    });
    
    return { success: true, message: 'Data saved successfully' };
    
  } catch (error) {
    console.error('Error saving data:', error);
    throw new Error('Failed to save data to Google Sheets: ' + error.message);
  }
}

/**
 * Run OpenAI analysis on cost deviations
 */
function runOpenAIAnalysis(deviations) {
  try {
    if (!CONFIG.OPENAI_API_KEY || CONFIG.OPENAI_API_KEY === 'your-openai-api-key-here') {
      throw new Error('OpenAI API key not configured');
    }
    
    const prompt = buildAnalysisPrompt(deviations);
    const response = callOpenAI(prompt);
    
    return parseAIResponse(response, deviations);
    
  } catch (error) {
    console.error('OpenAI Analysis Error:', error);
    
    // Fallback to local analysis
    return generateLocalAnalysis(deviations);
  }
}

/**
 * Build prompt for OpenAI analysis
 */
function buildAnalysisPrompt(deviations) {
  const formatPromptValue = (value, prefix = '₹') => {
    if (typeof value === 'string') return value; // 'No data'
    if (typeof value !== 'number') return value;
    return `${prefix}${value.toFixed(2)}`;
  };

  let prompt = `You are a cost analysis expert. For each item below, decide if the benchmark cost needs to be updated based on recent data.

COST DEVIATIONS:
`;

  deviations.forEach((dev, index) => {
    prompt += `
${index + 1}. Product: ${dev.product} (${dev.productionId})
   Type: ${dev.type}
   Current Benchmark: ${formatPromptValue(dev.currentBenchmark)}
   Proposed Benchmark: ${formatPromptValue(dev.proposedBenchmark)}
   Deviation: ${dev.deviation.toFixed(1)}%
   Last 3 Production Avg: ${formatPromptValue(dev.last3Avg)}
   Last 3 Month Avg: ${formatPromptValue(dev.last3MonthAvg)}
   Last 12 Month Avg: ${formatPromptValue(dev.last12MonthAvg)}
`;
    if (dev.notes) {
        prompt += `   Human Notes: ${dev.notes}\n`;
    }
  });

  prompt += `

For each deviation, provide your recommendation in the following JSON format as an array of objects. Do not include any other text or explanations outside of the JSON.

[
  {
    "status": "update" or "no_need",
    "decisionNotes": "Your brief analysis and reasoning here (max 50 words)."
  }
]

Guidelines for your decision:
- Base your decision on whether the 'Proposed Benchmark' is a realistic reflection of recent costs (Last 3 Production Avg, Last 3 Month Avg).
- Use any provided 'Human Notes' as additional context, but do not base your decision solely on them.
- If the 'Proposed Benchmark' aligns with recent trends and data, the status should be 'update'.
- If the 'Proposed Benchmark' seems anomalous, or if the deviation is insignificant, the status should be 'no_need'.
- If crucial data points are missing (indicated by "No data"), state this in your notes and be cautious. You might recommend 'no_need' pending more data.
- Your 'decisionNotes' should be a concise summary of your reasoning.

Data Security:
- The data provided is confidential. Do not repeat or disclose any part of the input data in your response, other than what is explicitly required by the JSON format.
- Your 'decisionNotes' should be a summary of your analysis, not a copy of the input data.`;

  return prompt;
}

/**
 * Call OpenAI API
 */
function callOpenAI(prompt) {
  const payload = {
    model: CONFIG.OPENAI_MODEL,
    messages: [
      {
        role: "system",
        content: "You are a manufacturing cost analysis expert with expertise in benchmark optimization and cost variance analysis."
      },
      {
        role: "user", 
        content: prompt
      }
    ],
    max_tokens: 2000,
    temperature: 0.3
  };

  const options = {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.OPENAI_API_KEY,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(CONFIG.OPENAI_API_URL, options);
  const responseData = JSON.parse(response.getContentText());
  
  if (responseData.error) {
    throw new Error('OpenAI API Error: ' + responseData.error.message);
  }
  
  return responseData.choices[0].message.content;
}

/**
 * Parse AI response and combine with deviation data
 */
function parseAIResponse(aiResponse, deviations) {
  try {
    // The AI response might be enclosed in ```json ... ```, so let's strip that.
    const cleanedResponse = aiResponse.replace(/```json\n?/g, '').replace(/\n?```/g, '');
    const recommendations = JSON.parse(cleanedResponse);
    
    return deviations.map((dev, index) => {
      const aiRec = recommendations[index] || {};
      
      return {
        ...dev,
        status: aiRec.status || 'no_need', // Default to 'no_need' if missing
        decisionNotes: aiRec.decisionNotes || 'AI response format error.'
      };
    });
    
  } catch (error) {
    console.error('Error parsing AI response:', error, 'Raw response:', aiResponse);
    // Fallback to local analysis if parsing fails
    return generateLocalAnalysis(deviations);
  }
}

/**
 * Generate local analysis as fallback
 */
function generateLocalAnalysis(deviations) {
  return deviations.map(dev => {
    let status = 'no_need';
    let decisionNotes = '';
    const deviation = parseFloat(dev.deviation);

    if (dev.proposedBenchmark === 'No data') {
        status = 'no_need';
        decisionNotes = 'Cannot recommend update without a proposed benchmark.';
    } else if (deviation > 10) {
        status = 'update';
        decisionNotes = `Significant deviation of ${deviation.toFixed(1)}% suggests an update is needed.`;
    } else if (deviation > 5) {
        status = 'update';
        decisionNotes = `Moderate deviation of ${deviation.toFixed(1)}% suggests an update might be needed. Review recommended.`;
    } else {
        status = 'no_need';
        decisionNotes = `Deviation of ${deviation.toFixed(1)}% is within acceptable limits. No update needed.`;
    }

    return {
      ...dev,
      status: status,
      decisionNotes: decisionNotes + ' (Local Fallback Analysis)'
    };
  });
}

/**
 * Export data to different formats
 */
function exportData(costingData, format) {
  try {
    switch(format.toLowerCase()) {
      case 'excel':
        return exportToExcel(costingData);
      case 'pdf':
        return exportToPDF(costingData);
      default:
        throw new Error('Unsupported export format');
    }
  } catch (error) {
    console.error('Export Error:', error);
    throw new Error('Failed to export data: ' + error.message);
  }
}

/**
 * Export to Excel format
 */
function exportToExcel(costingData) {
  // Create a new spreadsheet for export
  const exportSheet = SpreadsheetApp.create('Costing_Export_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  const sheet = exportSheet.getActiveSheet();
  
  // Add headers
  const headers = [
    'Product Name', 'Production ID', 'Package', 'Qty',
    'CA Correct', 'CA Benchmark', 'CA Last 3 Prod', 'CA Last 3 Month', 'CA Last 12 Month', 'CA New Benchmark', 'CA Status',
    'TOC Correct', 'TOC Benchmark', 'TOC Last 3 Prod', 'TOC Last 3 Month', 'TOC Last 12 Month', 'TOC New Benchmark', 'TOC Status',
    'Time Correct', 'Time Benchmark', 'Time Last 3 Prod', 'Time Last 3 Month', 'Time Last 12 Month', 'Time New Benchmark', 'Time Status'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Add data
  const rows = costingData.map(item => [
    item.productName, item.productionId, item.package, item.qty,
    item.ca.correct, item.ca.benchmark, item.ca.last3Prod, item.ca.last3Month, item.ca.last12Month, item.ca.newBenchmark, item.ca.status,
    item.toc.correct, item.toc.benchmark, item.toc.last3Prod, item.toc.last3Month, item.toc.last12Month, item.toc.newBenchmark, item.toc.status,
    item.time.correct, item.time.benchmark, item.time.last3Prod, item.time.last3Month, item.time.last12Month, item.time.newBenchmark, item.time.status
  ]);
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  // Format the sheet
  formatExportSheet(sheet);
  
  return {
    success: true,
    url: exportSheet.getUrl(),
    message: 'Excel export created successfully'
  };
}

/**
 * Export to PDF format
 */
function exportToPDF(costingData) {
  // Create a temporary document for PDF export
  const doc = DocumentApp.create('Costing_Report_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  const body = doc.getBody();
  
  // Add title
  const title = body.appendParagraph('Advanced Costing Review Report');
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);
  
  // Add summary
  body.appendParagraph('Generated on: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
  body.appendParagraph('Total Products: ' + costingData.length);
  
  // Add data table
  const table = body.appendTable();
  
  // Add headers
  const headerRow = table.appendTableRow();
  ['Product', 'Production ID', 'CA Current', 'CA Proposed', 'CA Status', 'TOC Current', 'TOC Proposed', 'TOC Status'].forEach(header => {
    headerRow.appendTableCell(header);
  });
  
  // Add data rows
  costingData.forEach(item => {
    const dataRow = table.appendTableRow();
    dataRow.appendTableCell(item.productName);
    dataRow.appendTableCell(item.productionId);
    dataRow.appendTableCell('₹' + item.ca.benchmark.toFixed(2));
    dataRow.appendTableCell('₹' + item.ca.newBenchmark.toFixed(2));
    dataRow.appendTableCell(item.ca.status);
    dataRow.appendTableCell('₹' + item.toc.benchmark.toFixed(2));
    dataRow.appendTableCell('₹' + item.toc.newBenchmark.toFixed(2));
    dataRow.appendTableCell(item.toc.status);
  });
  
  // Save and convert to PDF
  doc.saveAndClose();
  
  return {
    success: true,
    url: doc.getUrl(),
    message: 'PDF report created successfully'
  };
}

/**
 * Get or create the data sheet
 */
function getDataSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.DATA_SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`Sheet "${CONFIG.DATA_SHEET_NAME}" not found in spreadsheet`);
    }
    
    return sheet;
  } catch (error) {
    console.error('Error getting data sheet:', error);
    throw new Error('Failed to access Google Sheets: ' + error.message);
  }
}
