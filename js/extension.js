let selectedSheets = []; // Array of selected worksheet objects
let sheetSelectorCount = 0; // Counter for sheet selectors
let combinedData = null;
let isInitialized = false;
let apiWaitAttempts = 0;
const MAX_API_WAIT_ATTEMPTS = 100; // 10 seconds max wait

// Wait for tableau API to be available, then initialize
function waitForTableauAPI() {
  apiWaitAttempts++;
  
  // Check if Tableau Desktop injected the API automatically (this is the normal way)
  // OR if it was loaded from local file
  if (typeof tableau !== 'undefined' && tableau.extensions) {
    // Tableau API is available, initialize
    console.log('Tableau API found, initializing...');
      tableau.extensions.initializeAsync().then(() => {
        console.log('Extension initialized');
        isInitialized = true;
        setupEventListeners(); // Create sheet selectors first
        loadSheets(); // Then load and populate worksheets
      }).catch((error) => {
      console.error('Error initializing extension:', error);
      showStatus('Error initializing extension: ' + error.message, 'error');
    });
  } else if (apiWaitAttempts < MAX_API_WAIT_ATTEMPTS) {
    // Tableau API not ready yet, wait a bit and try again
    if (apiWaitAttempts === 1) {
      console.log('Waiting for Tableau API to load... (attempt ' + apiWaitAttempts + '/' + MAX_API_WAIT_ATTEMPTS + ')');
    }
    // Log progress every 10 attempts
    if (apiWaitAttempts % 10 === 0) {
      console.log('Still waiting... attempt ' + apiWaitAttempts + '/' + MAX_API_WAIT_ATTEMPTS);
    }
    setTimeout(waitForTableauAPI, 100);
  } else {
    // Timeout - API didn't load
    console.error('Tableau API failed to load after waiting');
    console.error('Current window.location:', window.location.href);
    console.error('typeof tableau:', typeof tableau);
    console.error('Script load status:', {
      loaded: window.tableauScriptLoaded,
      error: window.tableauScriptLoadError
    });
    
    let errorMsg = 'Tableau Extensions API failed to load. ';
    if (window.tableauScriptLoadError) {
      errorMsg += 'The script failed to load from CDN. This might be due to network/firewall restrictions. ';
    } else if (!window.tableauScriptLoaded) {
      errorMsg += 'The script may not have loaded. ';
    }
    errorMsg += 'Please check: 1) You are running in Tableau Desktop/Server, 2) Your internet connection, 3) Browser console (F12) for detailed errors.';
    
    showStatus(errorMsg, 'error');
  }
}

// Start waiting for Tableau API when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', waitForTableauAPI);
} else {
  // DOM is already ready
  waitForTableauAPI();
}

// Load available sheets from the dashboard
async function loadSheets() {
  try {
    // Check if tableau API is available
    if (typeof tableau === 'undefined' || !tableau.extensions) {
      showStatus('Tableau Extensions API not loaded. Please wait...', 'error');
      console.error('Tableau API not available');
      return;
    }
    
    // Check if initialized
    if (!isInitialized) {
      showStatus('Extension not initialized yet. Please wait...', 'error');
      console.error('Extension not initialized');
      return;
    }
    
    showStatus('Loading worksheets...', 'info');
    
    const dashboardContent = tableau.extensions.dashboardContent;
    const workbook = dashboardContent.workbook;
    
    console.log('DashboardContent:', dashboardContent);
    console.log('Workbook:', workbook);
    
    let worksheets = [];
    
    // Method 1: Try workbook.worksheets (gets ALL worksheets in workbook)
    if (workbook) {
      try {
        // Access worksheets - this should be a collection/array
        worksheets = Array.from(workbook.worksheets || []);
        console.log('Method 1 - Got worksheets from workbook.worksheets:', worksheets.length);
        
        // If empty, try accessing as property directly
        if (worksheets.length === 0 && workbook.worksheets) {
          // Try iterating if it's a collection
          worksheets = [];
          for (let i = 0; i < workbook.worksheets.length; i++) {
            worksheets.push(workbook.worksheets[i]);
          }
          console.log('Method 1b - Iterated worksheets:', worksheets.length);
        }
      } catch (e) {
        console.log('Error accessing workbook.worksheets:', e);
      }
    }
    
    // Method 2: If still empty, try dashboard worksheets (only sheets on dashboard)
    if (worksheets.length === 0) {
      try {
        const dashboard = dashboardContent.dashboard;
        if (dashboard && dashboard.worksheets) {
          worksheets = Array.from(dashboard.worksheets || []);
          console.log('Method 2 - Got worksheets from dashboard.worksheets:', worksheets.length);
        }
      } catch (e) {
        console.log('Error accessing dashboard.worksheets:', e);
      }
    }
    
    console.log('Total worksheets found:', worksheets.length);
    
    if (worksheets.length > 0) {
      console.log('Worksheet names:', worksheets.map((w, i) => `${i}: ${w.name}`));
    } else {
      console.warn('No worksheets found!');
      console.log('Available properties on workbook:', Object.keys(workbook || {}));
      console.log('Available properties on dashboardContent:', Object.keys(dashboardContent || {}));
    }
    
    if (worksheets.length === 0) {
      showStatus('No worksheets found. Make sure your workbook has worksheets. Check console for details.', 'error');
      return;
    }
    
    // Store worksheets array for later access
    window.worksheetsArray = worksheets;
    
    // Ensure at least one sheet selector exists
    const existingSelectors = document.querySelectorAll('.sheet-selector select');
    if (existingSelectors.length === 0) {
      console.log('No sheet selectors found, creating default ones...');
      // Create at least one selector if none exist
      addSheetSelector();
    }
    
    // Populate all existing sheet selectors
    document.querySelectorAll('.sheet-selector select').forEach(select => {
      // Clear existing options except the first one
      while (select.options.length > 1) {
        select.remove(1);
      }
      
      // Add worksheets
      worksheets.forEach((worksheet, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = worksheet.name;
        select.appendChild(option);
      });
    });
    
    showStatus(`Loaded ${worksheets.length} worksheet(s)`, 'success');
  } catch (error) {
    console.error('Error loading sheets:', error);
    console.error('Error stack:', error.stack);
    showStatus('Error loading worksheets: ' + error.message, 'error');
  }
}

// Add a new sheet selector
function addSheetSelector() {
  sheetSelectorCount++;
  const selectorId = `sheetSelector_${sheetSelectorCount}`;
  const selectId = `sheetSelect_${sheetSelectorCount}`;
  
  const sheetSelectorsDiv = document.getElementById('sheetSelectors');
  const sheetSelector = document.createElement('div');
  sheetSelector.className = 'sheet-selector';
  sheetSelector.id = selectorId;
  
  const label = document.createElement('label');
  label.textContent = `Sheet ${sheetSelectorCount}:`;
  label.setAttribute('for', selectId);
  
  const select = document.createElement('select');
  select.id = selectId;
  select.innerHTML = '<option value="">Select a sheet...</option>';
  
  // Populate with worksheets if available
  if (window.worksheetsArray) {
    window.worksheetsArray.forEach((worksheet, index) => {
      const option = document.createElement('option');
      option.value = index;
      option.textContent = worksheet.name;
      select.appendChild(option);
    });
  }
  
  // Add change event listener
  select.addEventListener('change', function() {
    updateSelectedSheets();
    checkReadyState();
  });
  
  const removeBtn = document.createElement('button');
  removeBtn.className = 'remove-btn';
  removeBtn.textContent = 'Remove';
  removeBtn.onclick = function() {
    sheetSelector.remove();
    updateSelectedSheets();
    checkReadyState();
  };
  
  sheetSelector.appendChild(label);
  sheetSelector.appendChild(select);
  sheetSelector.appendChild(removeBtn);
  sheetSelectorsDiv.appendChild(sheetSelector);
  
  checkReadyState();
}

// Remove a sheet selector
function removeSheetSelector(selectorId) {
  const selector = document.getElementById(selectorId);
  if (selector) {
    selector.remove();
    updateSelectedSheets();
    checkReadyState();
  }
}

// Update the selectedSheets array based on current selections
function updateSelectedSheets() {
  selectedSheets = [];
  document.querySelectorAll('.sheet-selector select').forEach((select, index) => {
    if (select.value !== '' && window.worksheetsArray) {
      const worksheet = window.worksheetsArray[parseInt(select.value)];
      if (worksheet) {
        selectedSheets.push({
          index: index + 1,
          worksheet: worksheet,
          name: worksheet.name
        });
      }
    }
  });
  console.log('Selected sheets:', selectedSheets.map(s => s.name));
}

// Setup event listeners
function setupEventListeners() {
  // Add initial sheet selector
  addSheetSelector();
  addSheetSelector(); // Add second one by default
}

// Check if at least one sheet is selected
function checkReadyState() {
  const downloadBtn = document.getElementById('downloadBtn');
  updateSelectedSheets();
  if (selectedSheets.length >= 1) {
    downloadBtn.disabled = false;
  } else {
    downloadBtn.disabled = true;
  }
}

// Get data from a worksheet
async function getWorksheetData(worksheet) {
  return new Promise((resolve, reject) => {
    worksheet.getSummaryDataAsync().then((dataTable) => {
      const columns = dataTable.columns;
      const data = dataTable.data;
      
      // Extract column names
      const columnNames = columns.map(col => col.fieldName || col.name);
      
      // Convert to array of objects
      const rows = data.map(row => {
        const rowObj = {};
        row.forEach((value, index) => {
          rowObj[columnNames[index]] = value.formattedValue || value.value;
        });
        return rowObj;
      });
      
      resolve({
        columns: columnNames,
        rows: rows
      });
    }).catch(reject);
  });
}

// Combine data from multiple sheets
async function combineData() {
  try {
    if (selectedSheets.length === 0) {
      showStatus('Please select at least one sheet', 'error');
      return;
    }
    
    showStatus('Loading data from sheets...', 'info');
    
    // Get data from all selected sheets
    const sheetsData = [];
    for (const sheetInfo of selectedSheets) {
      const data = await getWorksheetData(sheetInfo.worksheet);
      sheetsData.push({
        sheetIndex: sheetInfo.index,
        sheetName: sheetInfo.name,
        data: data
      });
    }
    
    // Use first sheet as primary order source
    const primarySheet = sheetsData[0];
    const dimensionCol = primarySheet.data.columns[0];
    
    console.log('Primary sheet:', primarySheet.sheetName);
    console.log('Total sheets to combine:', sheetsData.length);
    console.log('Dimension column:', dimensionCol);
    
    // Reverse rows to match dashboard order (Tableau often returns data in reverse)
    const reversedPrimaryRows = [...primarySheet.data.rows].reverse();
    
    // Reverse all other sheets' rows too
    sheetsData.forEach(sheetInfo => {
      sheetInfo.reversedRows = [...sheetInfo.data.rows].reverse();
    });
    
    // Build headers - start with dimension column
    const headers = [dimensionCol];
    
    // Add columns from each sheet (excluding dimension column)
    sheetsData.forEach(sheetInfo => {
      sheetInfo.data.columns.slice(1).forEach(col => {
        headers.push(`${col} (Sheet ${sheetInfo.sheetIndex})`);
      });
    });
    
    // Combine data by aligning all sheets by index position
    const combined = [];
    const maxRows = Math.max(...sheetsData.map(s => s.reversedRows.length));
    
    for (let i = 0; i < maxRows; i++) {
      const combinedRow = {};
      
      // Get dimension value from primary sheet
      if (reversedPrimaryRows[i]) {
        combinedRow[dimensionCol] = String(reversedPrimaryRows[i][dimensionCol] || '').trim();
      } else {
        // If primary sheet doesn't have this row, try to get from other sheets
        for (const sheetInfo of sheetsData) {
          if (sheetInfo.reversedRows[i]) {
            const dimCol = sheetInfo.data.columns[0];
            combinedRow[dimensionCol] = String(sheetInfo.reversedRows[i][dimCol] || '').trim();
            break;
          }
        }
      }
      
      // Add data from each sheet
      sheetsData.forEach(sheetInfo => {
        const sheetRow = sheetInfo.reversedRows[i] || null;
        const dimCol = sheetInfo.data.columns[0];
        
        sheetInfo.data.columns.slice(1).forEach(col => {
          const headerName = `${col} (Sheet ${sheetInfo.sheetIndex})`;
          if (sheetRow) {
            combinedRow[headerName] = sheetRow[col] || '';
          } else {
            combinedRow[headerName] = '';
          }
        });
      });
      
      combined.push(combinedRow);
    }
    
    console.log('Combined data:', {
      headers: headers,
      rowCount: combined.length,
      firstRow: combined[0]
    });
    
    combinedData = {
      headers: headers,
      rows: combined
    };
    
    return combinedData;
  } catch (error) {
    console.error('Error combining data:', error);
    showStatus('Error: ' + error.message, 'error');
    throw error;
  }
}

// Preview combined data
async function previewData() {
  try {
    const data = await combineData();
    
    const previewDiv = document.getElementById('preview');
    previewDiv.style.display = 'block';
    previewDiv.innerHTML = '<h3>Preview:</h3>';
    
    const table = document.createElement('table');
    
    // Create header row
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    data.headers.forEach(header => {
      const th = document.createElement('th');
      th.textContent = header;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // Create body rows
    const tbody = document.createElement('tbody');
    data.rows.forEach(row => {
      const tr = document.createElement('tr');
      data.headers.forEach(header => {
        const td = document.createElement('td');
        td.textContent = row[header] || '';
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    
    previewDiv.appendChild(table);
    showStatus('Data preview loaded successfully!', 'success');
  } catch (error) {
    showStatus('Error previewing data: ' + error.message, 'error');
  }
}

// Download data as Excel
async function downloadData() {
  try {
    if (!combinedData) {
      await combineData();
    }
    
    // Convert to worksheet format for XLSX
    const wsData = [];
    
    // Add headers
    wsData.push(combinedData.headers);
    
    // Add rows
    combinedData.rows.forEach(row => {
      const rowData = combinedData.headers.map(header => row[header] || '');
      wsData.push(rowData);
    });
    
    // Create workbook
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    
    // Set column widths
    const colWidths = combinedData.headers.map(() => ({ wch: 20 }));
    ws['!cols'] = colWidths;
    
    XLSX.utils.book_append_sheet(wb, ws, 'Combined Data');
    
    // Generate filename with timestamp
    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    const filename = `Combined_Loan_Application_Data_${timestamp}.xlsx`;
    
    // Download
    XLSX.writeFile(wb, filename);
    
    showStatus('File downloaded successfully: ' + filename, 'success');
  } catch (error) {
    console.error('Error downloading data:', error);
    showStatus('Error downloading data: ' + error.message, 'error');
  }
}

// Refresh worksheets (called from button)
function refreshWorksheets() {
  if (typeof tableau === 'undefined' || !tableau.extensions) {
    showStatus('Tableau Extensions API not loaded. Please wait and try again.', 'error');
    // Try to reinitialize
    waitForTableauAPI();
    return;
  }
  if (!isInitialized) {
    showStatus('Extension not initialized yet. Please wait...', 'error');
    return;
  }
  loadSheets();
}

// Show status message
function showStatus(message, type) {
  const statusDiv = document.getElementById('status');
  statusDiv.textContent = message;
  statusDiv.className = 'status ' + type;
  
  if (type === 'success' || type === 'error') {
    setTimeout(() => {
      statusDiv.className = 'status';
      statusDiv.textContent = '';
    }, 5000);
  }
}

