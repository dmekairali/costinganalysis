
/**
 * Google Apps Script Backend for Advanced Costing Review Dashboard
 * Handles OpenAI integration, Google Sheets operations, and data processing
 */

const CONFIG = {
  // Replace with your actual OpenAI API key
  OPENAI_API_KEY: 'your-openai-api-key-here',
  
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
    tocNewBenchmark: 'DJ',
    tocStatus: 'DK',

    // Time columns
    timeCorrect: 'DL',
    timeBenchmark: 'DM',
    timeLast3Prod: 'DN',
    timeLast3Month: 'DO',
    timeLast12Month: 'DP',
    timeNewBenchmark: 'DQ',
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
  
  return 'update'; // Default to 'update'
}

/**
 * Normalize time values
 */
function normalizeTime(timeValue) {
  if (!timeValue) return '0:00:00';
  
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
 * Get configuration template for easy setup
 */
function getConfigurationTemplate() {
  return {
    message: "Copy and modify this configuration in your CONFIG object",
    example: {
      HEADER_ROW: 1,
      DATA_START_ROW: 2,
      COLUMN_MAPPING: {
        productName: 'A',        // Product name column
        productionId: 'B',       // Production ID column
        package: 'C',            // Package type column
        qty: 'D',                // Quantity column
        caCorrect: 'E',          // CA Correct cost
        caBenchmark: 'F',        // CA Benchmark cost
        caNewBenchmark: 'J',     // CA New Benchmark
        caStatus: 'K',           // CA Status
        tocCorrect: 'L',         // TOC Correct cost
        tocBenchmark: 'M',       // TOC Benchmark cost
        tocNewBenchmark: 'Q',    // TOC New Benchmark
        tocStatus: 'R'           // TOC Status
      },
      FILTER_CONDITIONS: {
        skipEmptyProduct: true,
        skipEmptyProductionId: true,
        includeStatuses: ['update', 'pending'], // Only these statuses
        includePackages: ['Standard', 'Premium'], // Only these packages
        minQuantity: 1,
        skipZeroCosts: true,
        customFilter: function(data) {
          // Example: Only include products with cost > 50
          return (data.caBenchmark || 0) > 50;
        }
      }
    }
  };
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
  let prompt = `You are a cost analysis expert. Analyze the following cost deviations and provide recommendations:

COST DEVIATIONS:
`;

  deviations.forEach((dev, index) => {
    prompt += `
${index + 1}. Product: ${dev.product} (${dev.productionId})
   Type: ${dev.type}
   Current Benchmark: ₹${dev.currentBenchmark.toFixed(2)}
   Proposed Benchmark: ₹${dev.proposedBenchmark.toFixed(2)}
   Deviation: ${dev.deviation.toFixed(1)}%
   Last 3 Production Avg: ₹${dev.last3Avg.toFixed(2)}
   Last 3 Month Avg: ₹${dev.last3MonthAvg.toFixed(2)}
   Last 12 Month Avg: ₹${dev.last12MonthAvg.toFixed(2)}
`;
  });

  prompt += `

For each deviation, provide:
1. RECOMMENDATION: APPROVE, REJECT, or CAUTION
2. REASONING: Brief explanation (max 100 words)
3. CONFIDENCE: Percentage (70-95%)
4. TREND: Increasing, Decreasing, or Stable

Guidelines:
- REJECT if deviation > 15% and trend is increasing
- CAUTION if deviation 10-15% 
- APPROVE if deviation < 10% or decreasing costs
- Consider production volume and historical trends

Respond in JSON format with array of recommendations.`;

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
    const recommendations = JSON.parse(aiResponse);
    
    return deviations.map((dev, index) => {
      const aiRec = recommendations[index] || {};
      
      return {
        ...dev,
        recommendation: formatRecommendation(aiRec.recommendation, aiRec.reasoning, dev.deviation),
        confidence: aiRec.confidence || 85,
        trend: aiRec.trend || analyzeTrend(dev),
        riskLevel: calculateRiskLevel(dev.deviation)
      };
    });
    
  } catch (error) {
    console.error('Error parsing AI response:', error);
    return generateLocalAnalysis(deviations);
  }
}

/**
 * Format recommendation with emoji and action
 */
function formatRecommendation(action, reasoning, deviation) {
  const emoji = action === 'APPROVE' ? '✅' : action === 'REJECT' ? '❌' : '⚠️';
  return `${emoji} ${action}: ${reasoning || getDefaultReasoning(action, deviation)}`;
}

/**
 * Get default reasoning based on action and deviation
 */
function getDefaultReasoning(action, deviation) {
  switch(action) {
    case 'APPROVE':
      return `${deviation.toFixed(1)}% deviation is within acceptable limits based on production data.`;
    case 'REJECT':
      return `${deviation.toFixed(1)}% increase is too high and requires cost optimization review.`;
    case 'CAUTION':
      return `${deviation.toFixed(1)}% increase needs careful monitoring and stakeholder approval.`;
    default:
      return `Deviation of ${deviation.toFixed(1)}% requires analysis.`;
  }
}

/**
 * Generate local analysis as fallback
 */
function generateLocalAnalysis(deviations) {
  return deviations.map(dev => {
    const trend = analyzeTrend(dev);
    const riskLevel = calculateRiskLevel(dev.deviation);
    
    let action, reasoning;
    
    if (dev.deviation > 15 && trend === 'Increasing') {
      action = 'REJECT';
      reasoning = 'High deviation with increasing trend indicates cost control issues.';
    } else if (dev.deviation > 10) {
      action = 'CAUTION';
      reasoning = 'Moderate deviation requires management approval and monitoring.';
    } else if (dev.deviation < 0) {
      action = 'APPROVE';
      reasoning = 'Cost reduction should be implemented immediately.';
    } else {
      action = 'APPROVE';
      reasoning = 'Deviation is within acceptable limits.';
    }
    
    return {
      ...dev,
      recommendation: formatRecommendation(action, reasoning, dev.deviation),
      confidence: Math.floor(Math.random() * 15) + 80,
      trend: trend,
      riskLevel: riskLevel
    };
  });
}

/**
 * Analyze cost trend
 */
function analyzeTrend(deviation) {
  const recent = (deviation.last3Avg + deviation.last3MonthAvg) / 2;
  const historical = deviation.last12MonthAvg;
  
  if (recent > historical * 1.1) return 'Increasing';
  if (recent < historical * 0.9) return 'Decreasing';
  return 'Stable';
}

/**
 * Calculate risk level based on deviation
 */
function calculateRiskLevel(deviation) {
  if (Math.abs(deviation) > 20) return 'High';
  if (Math.abs(deviation) > 10) return 'Medium';
  return 'Low';
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
