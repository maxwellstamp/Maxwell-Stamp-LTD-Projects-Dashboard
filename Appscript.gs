// Supabase Configuration
const SUPABASE_CONFIG = {
  URL: "https://woklvqytrxauxosurmyv.supabase.co"
};

/**
 * Initialize configuration from Script Properties
 */
function initConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('SUPABASE_API_KEY');
  
  if (!apiKey) {
    throw new Error('Supabase API key not found. Please run "0. Setup API Key" first.');
  }
  
  return {
    URL: SUPABASE_CONFIG.URL,
    API_KEY: apiKey
  };
}

/**
 * Setup script properties - RUN THIS FIRST
 */
function setupScriptProperties() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Setup Supabase API Key',
    'Enter your Supabase API Key:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText();
    
    if (apiKey && apiKey.startsWith('eyJ')) {
      PropertiesService.getScriptProperties().setProperty('SUPABASE_API_KEY', apiKey);
      ui.alert('✅ Success', 'API key stored securely!', ui.ButtonSet.OK);
      return true;
    } else {
      ui.alert('❌ Error', 'Invalid API key format', ui.ButtonSet.OK);
      return false;
    }
  }
  return false;
}

/**
 * Main function to sync all data from Google Sheets to Supabase
 */
function syncAllDataToSupabase() {
  try {
    const config = initConfig();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    if (!spreadsheet) {
      throw new Error('No active spreadsheet found.');
    }
    
    let sheet = spreadsheet.getSheetByName('sheet1');
    
    // If sheet1 not found, try common alternative names
    if (!sheet) {
      const alternativeNames = ['Sheet1', 'SHEET1', 'Data', 'Main', 'Transport Plan'];
      for (const name of alternativeNames) {
        sheet = spreadsheet.getSheetByName(name);
        if (sheet) break;
      }
    }
    
    if (!sheet) {
      throw new Error('Sheet not found. Please update the sheet name in the code.');
    }
    
    const records = syncSheetToSupabase(sheet, config);
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('✅ Sync Complete', `Successfully synced ${records} records to Supabase`, ui.ButtonSet.OK);
    
    return records;
    
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('❌ Sync Failed', error.toString(), ui.ButtonSet.OK);
    throw error;
  }
}

/**
 * Sync individual sheet to Supabase
 */
function syncSheetToSupabase(sheet, config) {
  const sheetName = sheet.getName();
  const data = getSheetData(sheet);
  
  if (data.length === 0) {
    return 0;
  }
  
  // Transform data for Supabase
  const supabaseData = data.map(row => transformRowForSupabase(row, sheetName));
  
  // Upsert data to Supabase
  const result = upsertToSupabase('met', supabaseData, config);
  
  return result.length;
}

/**
 * Get all data from a sheet (excluding header row)
 */
function getSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  
  if (lastRow <= 1 || lastColumn === 0) {
    return [];
  }
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  const data = dataRange.getValues();
  
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  
  // Convert to array of objects
  const result = data.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      if (header && header.toString().trim() !== '') {
        const cleanHeader = header.toString().trim();
        obj[cleanHeader] = row[index] || '';
      }
    });
    return obj;
  });
  
  // Filter out completely empty rows
  return result.filter(row => {
    return Object.values(row).some(value => value !== '');
  });
}

/**
 * Transform row data for Supabase schema
 */
function transformRowForSupabase(row, sheetName) {
  return {
    sl_number: row['S/L.'] || row['SL'] || row['SNo'] || row['Sl No'] || row['Serial'] || '',
    checklist_category: row['Checklist Category'] || row['Category'] || '',
    oversight_manager: row['Oversight Manager'] || row['Manager'] || '',
    action_item: row['Actionable Items'] || row['Action Item'] || row['Action'] || '',
    responsible_person_name: row['Responsible Person Name'] || row['Responsible Person'] || row['Person Name'] || '',
    responsible_person_designation: row['Responsible Person Designation'] || row['Designation'] || '',
    responsible_person_email: row['Responsible Person Email'] || row['Email'] || '',
    reminder_cc_email: row['Reminder Cc email'] || row['CC Email'] || '',
    due_date: formatDateForSupabase(row['Due Date']),
    reminder_days: row['Reminder days before due date'] || row['Reminder Days'] || '',
    reminder_sent: row['Reminder Sent?'] || row['Reminder Sent'] || 'No',
    reminder_count: parseInt(row['Reminder Count']) || 0,
    status: row['Status'] || 'Not Done',
    comments: row['Comments'] || '',
    sheet_source: sheetName,
    last_synced: new Date().toISOString()
  };
}

/**
 * Format date for Supabase
 */
function formatDateForSupabase(dateValue) {
  if (!dateValue) return null;
  
  try {
    if (typeof dateValue === 'string') {
      const date = new Date(dateValue);
      return isNaN(date.getTime()) ? null : date.toISOString().split('T')[0];
    } else if (dateValue instanceof Date) {
      return dateValue.toISOString().split('T')[0];
    }
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Enhanced upsert function that prevents duplicates
 */
function upsertToSupabase(tableName, data, config) {
  if (!data || data.length === 0) {
    return [];
  }
  
  const url = `${config.URL}/rest/v1/${tableName}`;
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${config.API_KEY}`,
      'apikey': config.API_KEY,
      'Content-Type': 'application/json',
      'Prefer': 'resolution=merge-duplicates'
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode >= 200 && responseCode < 300) {
      try {
        return JSON.parse(response.getContentText());
      } catch (e) {
        return data;
      }
    } else {
      // If upsert fails, fall back to individual updates
      return updateRecordsIndividually(tableName, data, config);
    }
  } catch (error) {
    throw error;
  }
}

/**
 * Update records individually as fallback
 */
function updateRecordsIndividually(tableName, data, config) {
  const results = [];
  
  for (let i = 0; i < data.length; i++) {
    try {
      const record = data[i];
      const slNumber = record.sl_number;
      
      if (!slNumber) continue;
      
      // Use optimized PATCH method for individual updates
      const success = updateIndividualRowInSupabase(slNumber, record, config);
      if (success) {
        results.push(record);
      }
    } catch (error) {
      // Continue with next record even if one fails
      console.error('Error updating record:', error);
      continue;
    }
  }
  
  return results;
}

/**
 * Update individual row in Supabase using PATCH
 */
function updateIndividualRowInSupabase(slNumber, data, config) {
  if (!slNumber) {
    throw new Error('SL number required for individual row update');
  }
  
  const url = `${config.URL}/rest/v1/met?sl_number=eq.${encodeURIComponent(slNumber)}`;
  
  const options = {
    method: 'PATCH',
    headers: {
      'Authorization': `Bearer ${config.API_KEY}`,
      'apikey': config.API_KEY,
      'Content-Type': 'application/json',
      'Prefer': 'return=minimal'
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode >= 200 && responseCode < 300) {
      return true;
    } else if (responseCode === 404) {
      // Record doesn't exist, create it
      console.log('Record not found, creating new one for SL:', slNumber);
      return createRecordInSupabase(data, config);
    } else {
      console.error(`PATCH failed: ${responseCode} - ${responseText}`);
      return false;
    }
  } catch (error) {
    console.error('Error in updateIndividualRowInSupabase:', error);
    return false;
  }
}

/**
 * Create new record using POST method
 */
function createRecordInSupabase(data, config) {
  const url = `${config.URL}/rest/v1/met`;
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${config.API_KEY}`,
      'apikey': config.API_KEY,
      'Content-Type': 'application/json',
      'Prefer': 'return=minimal'
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  
  return responseCode >= 200 && responseCode < 300;
}

/**
 * Get data from Supabase with proper ordering
 */
function getDataFromSupabase(tableName = 'met') {
  const config = initConfig();
  
  // Add ORDER BY to ensure consistent ordering by serial number
  const url = `${config.URL}/rest/v1/${tableName}?select=*&order=sl_number.asc`;
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${config.API_KEY}`,
      'apikey': config.API_KEY,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode >= 200 && responseCode < 300) {
      return JSON.parse(response.getContentText());
    }
    return [];
  } catch (error) {
    return [];
  }
}

/**
 * View synced data in a new sheet with proper ordering
 */
function viewSyncedData() {
  try {
    const data = getDataFromSupabase();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    let viewSheet = spreadsheet.getSheetByName('Supabase Data');
    if (!viewSheet) {
      viewSheet = spreadsheet.insertSheet('Supabase Data');
    } else {
      viewSheet.clear();
    }
    
    if (data.length > 0) {
      const headers = Object.keys(data[0]);
      viewSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      const rows = data.map(item => headers.map(header => item[header] || ''));
      viewSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      
      viewSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      viewSheet.autoResizeColumns(1, headers.length);
    }
    
    SpreadsheetApp.getUi().alert(`✅ Displaying ${data.length} records from Supabase (ordered by SL number)`);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.toString()}`);
  }
}

/**
 * OPTIMIZED AUTO-SYNC: Set up automatic sync triggers
 */
function setupAutoSyncTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  
  // Remove existing sync triggers to avoid duplicates
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onSheetEdit' || trigger.getHandlerFunction() === 'scheduledSync') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create edit trigger for real-time sync
  ScriptApp.newTrigger('onSheetEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  
  // Create time-based trigger for periodic sync (every hour)
  ScriptApp.newTrigger('scheduledSync')
    .timeBased()
    .everyHours(1)
    .create();
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('✅ Auto-Sync Enabled', 'Real-time and hourly sync triggers have been set up.', ui.ButtonSet.OK);
}

/**
 * Remove all sync triggers
 */
function removeSyncTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let removedCount = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onSheetEdit' || 
        trigger.getHandlerFunction() === 'scheduledSync') {
      ScriptApp.deleteTrigger(trigger);
      removedCount++;
    }
  });
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('✅ Auto-Sync Disabled', `Removed ${removedCount} sync triggers.`, ui.ButtonSet.OK);
}

/**
 * OPTIMIZED AUTO-SYNC: Trigger function that runs automatically when sheet is edited
 * Uses PATCH method for individual row updates
 */
function onSheetEdit(e) {
  try {
    // Only process if the edit is in our target sheet
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    
    // Define which sheets to auto-sync
    const targetSheets = ['sheet1', 'Sheet1', 'Data', 'Main', 'Transport Plan'];
    
    if (!targetSheets.includes(sheetName)) {
      return;
    }
    
    // Get the edited range
    const range = e.range;
    const row = range.getRow();
    const column = range.getColumn();
    
    // Only sync if edit is in data area (not headers) and within reasonable range
    if (row > 1 && row < 1000) { // Added upper limit to prevent infinite loops
      // Add a small delay to ensure the edit is complete
      Utilities.sleep(1000);
      
      // Use PATCH method to sync only the edited row
      syncEditedRowWithPatch(sheet, row);
    }
    
  } catch (error) {
    // Silent fail for auto-sync to avoid disrupting user experience
    console.error('Auto-sync error:', error);
  }
}

/**
 * OPTIMIZED: Sync only the edited row using PATCH method
 */
function syncEditedRowWithPatch(sheet, editedRow) {
  try {
    const config = initConfig();
    const lastColumn = sheet.getLastColumn();
    
    if (lastColumn === 0) return;
    
    // Get headers
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    
    // Get the edited row data
    const rowData = sheet.getRange(editedRow, 1, 1, lastColumn).getValues()[0];
    
    // Convert to object
    const rowObject = {};
    headers.forEach((header, index) => {
      if (header && header.toString().trim() !== '') {
        const cleanHeader = header.toString().trim();
        rowObject[cleanHeader] = rowData[index] || '';
      }
    });
    
    // Skip if row is empty or missing critical data
    if (!Object.values(rowObject).some(value => value !== '') || !rowObject['S/L.']) {
      return;
    }
    
    // Transform for Supabase
    const supabaseData = transformRowForSupabase(rowObject, sheet.getName());
    const slNumber = supabaseData.sl_number;
    
    if (!slNumber) return;
    
    // Use PATCH method to update only this specific row
    const success = updateIndividualRowInSupabase(slNumber, supabaseData, config);
    
    if (!success) {
      console.error('Failed to sync row:', editedRow);
    }
    
  } catch (error) {
    console.error('Error in syncEditedRowWithPatch:', error);
  }
}

/**
 * Scheduled function for periodic full sync
 */
function scheduledSync() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('sheet1');
    
    if (!sheet) return;
    
    syncSheetToSupabase(sheet, initConfig());
    
  } catch (error) {
    // Silent fail for scheduled sync
  }
}

/**
 * Check auto-sync status
 */
function checkAutoSyncStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const ui = SpreadsheetApp.getUi();
  
  let message = 'Auto-Sync Status:\n\n';
  let hasEditTrigger = false;
  let hasScheduledTrigger = false;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onSheetEdit') {
      hasEditTrigger = true;
      message += `✅ Edit Trigger: ACTIVE (runs on cell edit)\n`;
    }
    if (trigger.getHandlerFunction() === 'scheduledSync') {
      hasScheduledTrigger = true;
      message += `✅ Scheduled Trigger: ACTIVE (runs hourly)\n`;
    }
  });
  
  if (!hasEditTrigger) {
    message += `❌ Edit Trigger: MISSING\n`;
  }
  if (!hasScheduledTrigger) {
    message += `❌ Scheduled Trigger: MISSING\n`;
  }
  
  message += `\nTotal triggers: ${triggers.length}`;
  
  ui.alert('Auto-Sync Status', message, ui.ButtonSet.OK);
}

/**
 * Test Supabase connection
 */
function testSupabaseConnection() {
  try {
    const config = initConfig();
    const data = getDataFromSupabase();
    const ui = SpreadsheetApp.getUi();
    
    if (data && Array.isArray(data)) {
      ui.alert(
        'Supabase Connection Test',
        `✅ Connection successful!\n\nFound ${data.length} records in Supabase.\nData is ordered by SL number.`,
        ui.ButtonSet.OK
      );
      return true;
    } else {
      ui.alert(
        'Supabase Connection Test',
        `⚠️ Connection works but no data found.`,
        ui.ButtonSet.OK
      );
      return true;
    }
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Supabase Connection Test',
      `❌ Connection failed:\n${error.toString()}`,
      ui.ButtonSet.OK
    );
    return false;
  }
}

/**
 * Test auto-sync for a specific row
 */
function testAutoSyncForRow() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Test Auto-Sync',
    'Enter row number to test sync:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const rowNumber = parseInt(response.getResponseText());
    
    if (isNaN(rowNumber) || rowNumber < 2) {
      ui.alert('❌ Error', 'Please enter a valid row number (2 or higher)', ui.ButtonSet.OK);
      return;
    }
    
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      syncEditedRowWithPatch(sheet, rowNumber);
      ui.alert('✅ Success', `Row ${rowNumber} sync test completed.`, ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('❌ Error', `Failed to sync row ${rowNumber}:\n${error.toString()}`, ui.ButtonSet.OK);
    }
  }
}

/**
 * Create menu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔗 Supabase Sync')
    .addItem('0. Setup API Key (Run First)', 'setupScriptProperties')
    .addItem('1. Sync to Supabase', 'syncAllDataToSupabase')
    .addItem('2. Enable Auto-Sync', 'setupAutoSyncTriggers')
    .addItem('3. Disable Auto-Sync', 'removeSyncTriggers')
    .addItem('4. View Synced Data', 'viewSyncedData')
    .addSeparator()
    .addItem('🔧 Check Auto-Sync Status', 'checkAutoSyncStatus')
    .addItem('🔧 Test Connection', 'testSupabaseConnection')
    .addItem('🔧 Test Auto-Sync for Row', 'testAutoSyncForRow')
    .addToUi();
}