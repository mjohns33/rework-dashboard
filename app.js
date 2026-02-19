// Rework Tracking Dashboard
const STORAGE_KEY = 'rework-dashboard-data-v13';
let reworkData = [];
let trendChart = null;
let causeChart = null;
const AI_BASE_URL = 'http://localhost:3001';
const AI_ENDPOINT = `${AI_BASE_URL}/api/ai-insights`;
const AI_HEALTH_ENDPOINT = `${AI_BASE_URL}/health`;
const DEFAULT_PERFORMANCE_GOALS = {
  reworkCost: 7000000,
  releaseRate: 40,
  rootCauseAssignment: 90
};
let PERFORMANCE_GOALS = { ...DEFAULT_PERFORMANCE_GOALS };

// Track sort state for defect drivers table
let defectDriversSortMode = 'cases'; // Options: 'cases', 'cost', 'lag'

// Initialize dashboard on page load
document.addEventListener('DOMContentLoaded', () => {
  loadStoredData();
  hookUpDateEvents();
  setupDefectDriversSorting();
  setupAIInsights();
  renderDashboard();
});

// ---------- File Load ----------
function loadCSV() {
    if (window && window.console) {
      console.log('DEBUG: loadCSV called, file input:', document.getElementById('csvFile').files);
    }
  const fileInput = document.getElementById('csvFile');
  const loadBtn = document.getElementById('loadBtn');
  const file = fileInput.files[0];
  const isExcelFile = /\.(xlsx|xls)$/i.test(file?.name || '');

  if (!file) {
    showMessage('Please select a CSV file', 'error');
    return;
  }

  const maxSize = 50 * 1024 * 1024; // 50 MB limit
  if (file.size > maxSize) {
    showMessage(`File too large (${(file.size / 1024 / 1024).toFixed(1)} MB). Max 50 MB.`, 'error');
    return;
  }

  loadBtn.disabled = true;
  const spinner = document.getElementById('loadingSpinner');
  spinner.classList.add('active');
  showMessage(`Loading ${file.name}... please wait`, 'info');

  const reader = new FileReader();
  reader.onload = (e) => {
    if (window && window.console) {
      console.log('DEBUG: FileReader onload triggered');
    }
    try {
      let csv = '';
      let goalsApplied = 0;

      if (isExcelFile) {
        if (typeof XLSX === 'undefined') {
          showMessage('Excel parsing library failed to load. Please refresh and try again.', 'error');
          loadBtn.disabled = false;
          spinner.classList.remove('active');
          return;
        }

        const workbook = XLSX.read(e.target.result, { type: 'array' });
        csv = extractDataCSVFromWorkbook(workbook);
        goalsApplied = applyGoalsFromWorkbook(workbook);
      } else {
        csv = e.target.result;
        resetPerformanceGoals();
      }

      if (window && window.console) {
        console.log('DEBUG: Data file contents (first 500 chars):', String(csv).slice(0, 500));
      }

      showMessage('Parsing data...', 'info');
      setTimeout(() => {
        let data = [];
        try {
          data = parseCSV(csv);
          if (window && window.console) {
            console.log('DEBUG: parseCSV returned', data.length, 'rows');
          }
        } catch (parseErr) {
          let goalsFromCSV = 0;
          const isDateColumnError = /date column/i.test(String(parseErr?.message || ''));

          if (!isExcelFile && isDateColumnError) {
            goalsFromCSV = applyGoalsFromCSVText(csv);
          }

          if (goalsFromCSV > 0) {
            if (reworkData.length > 0) {
              renderDashboard();
            } else {
              renderPerformanceGoals([]);
              updateDataInfo();
            }
            const dataNote = reworkData.length > 0
              ? ` Existing data remains loaded (${reworkData.length} records).`
              : ' Upload a date-based data export to populate dashboard records.';
            showMessage(`✓ Loaded ${goalsFromCSV} goal(s) from ${file.name}.${dataNote}`, 'success');
            loadBtn.disabled = false;
            spinner.classList.remove('active');
            return;
          }

          showMessage(`CSV parsing failed: ${parseErr.message}`, 'error');
          if (window && window.console) {
            console.error('DEBUG: CSV parsing failed:', parseErr);
          }
          loadBtn.disabled = false;
          spinner.classList.remove('active');
          return;
        }
        if (data.length === 0) {
          showMessage('CSV file is empty or invalid format', 'error');
          if (window && window.console) {
            console.error('DEBUG: CSV file is empty or invalid format');
          }
          loadBtn.disabled = false;
          spinner.classList.remove('active');
          return;
        }
        reworkData = data;
        saveData();
        populateLocationFilter();
        renderDashboard();
        const goalsNote = goalsApplied > 0 ? ` • ${goalsApplied} goal(s) loaded from Goals sheet` : '';
        showMessage(`✓ Successfully loaded ${data.length} records from ${file.name}${goalsNote}`, 'success');
        loadBtn.disabled = false;
        spinner.classList.remove('active');
      }, 100);
    } catch (error) {
      showMessage(`Error parsing CSV: ${error.message}`, 'error');
      console.error(error);
      loadBtn.disabled = false;
      spinner.classList.remove('active');
    }
  };

  reader.onerror = (e) => {
    showMessage('Error reading file. Please try again.', 'error');
    loadBtn.disabled = false;
    spinner.classList.remove('active');
  };

  if (isExcelFile) {
    reader.readAsArrayBuffer(file);
  } else {
    reader.readAsText(file);
  }
}

function resetPerformanceGoals() {
  PERFORMANCE_GOALS = { ...DEFAULT_PERFORMANCE_GOALS };
}

function normalizeHeaderCell(value) {
  return String(value || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function rowHasAnyValue(row) {
  return Array.isArray(row) && row.some((cell) => String(cell || '').trim() !== '');
}

function scoreDataHeaderRow(row) {
  if (!Array.isArray(row)) return 0;

  const headers = row.map((cell) => normalizeHeaderCell(cell));
  const hasMatch = (tokens) => headers.some((header) => tokens.some((token) => header.includes(token)));

  let score = 0;
  if (hasMatch(['hold date', 'production date', 'date produced', 'prod date', 'date', 'day'])) score += 5;
  if (hasMatch(['root cause', 'cause', 'reason', 'root'])) score += 2;
  if (hasMatch(['cases produced', 'casesproduced'])) score += 2;
  if (hasMatch(['cases reworked', 'casesreworked'])) score += 2;
  if (hasMatch(['disposition', 'status'])) score += 1;
  if (hasMatch(['plant', 'work center', 'site', 'location'])) score += 1;

  return score;
}

function buildCSVFromRows(rows, startRowIndex) {
  const sliced = rows.slice(startRowIndex).filter((row) => rowHasAnyValue(row));
  if (sliced.length === 0) return '';
  return Papa.unparse(sliced);
}

function extractDataCSVFromWorkbook(workbook) {
  if (!workbook || !Array.isArray(workbook.SheetNames) || workbook.SheetNames.length === 0) {
    throw new Error('Workbook has no sheets');
  }

  const nonGoalSheets = workbook.SheetNames.filter((name) => !String(name).toLowerCase().includes('goal'));
  const candidateSheets = nonGoalSheets.length > 0 ? nonGoalSheets : workbook.SheetNames;

  let bestMatch = null;

  candidateSheets.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!Array.isArray(rows) || rows.length === 0) return;

    const maxScanRows = Math.min(rows.length, 35);
    for (let rowIndex = 0; rowIndex < maxScanRows; rowIndex += 1) {
      const row = rows[rowIndex];
      if (!rowHasAnyValue(row)) continue;

      const score = scoreDataHeaderRow(row);
      if (!bestMatch || score > bestMatch.score) {
        bestMatch = {
          sheetName,
          rows,
          rowIndex,
          score
        };
      }
    }
  });

  if (!bestMatch || bestMatch.score < 5) {
    throw new Error('Could not find a data table with a Date header in the workbook.');
  }

  if (window && window.console) {
    console.log('DEBUG: Selected workbook data sheet:', bestMatch.sheetName, 'header row index:', bestMatch.rowIndex, 'score:', bestMatch.score);
  }

  const csv = buildCSVFromRows(bestMatch.rows, bestMatch.rowIndex);
  if (!csv) {
    throw new Error('Unable to extract data rows from workbook.');
  }

  return csv;
}

function parseGoalNumber(value) {
  if (value === null || value === undefined) return NaN;
  const raw = String(value).trim();
  if (!raw) return NaN;
  const normalized = raw.replace(/[%,$\s]/g, '').replace(/,/g, '');
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : NaN;
}

function mapGoalMetric(metricLabel) {
  const metric = String(metricLabel || '').toLowerCase().replace(/\s+/g, ' ').trim();
  if (!metric) return null;
  if (metric.includes('rework cost')) return 'reworkCost';
  if (metric.includes('release rate') || metric.includes('released')) return 'releaseRate';
  if (metric.includes('root cause assignment') || metric.includes('assignment')) return 'rootCauseAssignment';
  if (metric.includes('rework')) return 'reworkCost';
  if (metric.includes('release')) return 'releaseRate';
  return null;
}

function applyGoalsFromWorkbook(workbook) {
  resetPerformanceGoals();
  if (!workbook || !Array.isArray(workbook.SheetNames)) return 0;

  const goalsSheetName = workbook.SheetNames.find((name) => String(name).toLowerCase().includes('goal'));
  if (!goalsSheetName) return 0;

  const goalsSheet = workbook.Sheets[goalsSheetName];
  if (!goalsSheet) return 0;

  const rows = XLSX.utils.sheet_to_json(goalsSheet, { header: 1, defval: '' });
  let appliedCount = 0;

  rows.forEach((row) => {
    if (!Array.isArray(row) || row.length === 0) return;

    const textCells = row
      .map((cell) => String(cell).trim())
      .filter((cell) => cell.length > 0);
    if (textCells.length === 0) return;

    const metricKey = mapGoalMetric(textCells[0]);
    if (!metricKey) return;

    const valueCandidate = row
      .map((cell) => parseGoalNumber(cell))
      .find((num) => Number.isFinite(num));

    if (!Number.isFinite(valueCandidate)) return;

    PERFORMANCE_GOALS[metricKey] = valueCandidate;
    appliedCount += 1;
  });

  return appliedCount;
}

function applyGoalsFromCSVText(csvText) {
  if (!csvText) return 0;

  const parsed = Papa.parse(csvText, {
    header: false,
    skipEmptyLines: true
  });

  const rows = Array.isArray(parsed.data) ? parsed.data : [];
  if (rows.length === 0) return 0;

  let appliedCount = 0;
  const appliedKeys = new Set();

  const applyGoal = (metricLabel, value) => {
    const metricKey = mapGoalMetric(metricLabel);
    if (!metricKey || !Number.isFinite(value)) return;
    PERFORMANCE_GOALS[metricKey] = value;
    if (!appliedKeys.has(metricKey)) {
      appliedKeys.add(metricKey);
      appliedCount += 1;
    }
  };

  // Horizontal pattern: header row + value row
  const maxPairScan = Math.min(rows.length - 1, 12);
  for (let i = 0; i < maxPairScan; i += 1) {
    const headerRow = Array.isArray(rows[i]) ? rows[i] : [];
    const valueRow = Array.isArray(rows[i + 1]) ? rows[i + 1] : [];
    if (headerRow.length === 0 || valueRow.length === 0) continue;

    for (let col = 0; col < headerRow.length; col += 1) {
      const goalValue = parseGoalNumber(valueRow[col]);
      applyGoal(headerRow[col], goalValue);
    }
  }

  // Vertical pattern: Metric in first column, goal value in same row
  const maxVerticalScan = Math.min(rows.length, 30);
  for (let i = 0; i < maxVerticalScan; i += 1) {
    const row = Array.isArray(rows[i]) ? rows[i] : [];
    if (row.length === 0) continue;

    const label = String(row[0] || '').trim();
    if (!label) continue;

    const valueCandidate = row
      .map((cell) => parseGoalNumber(cell))
      .find((num) => Number.isFinite(num));

    applyGoal(label, valueCandidate);
  }

  return appliedCount;
}

// ---------- CSV Parsing ----------
function parseCSV(csv) {
  // Use Papa Parse to handle complex CSV with quoted fields
  const result = Papa.parse(csv, {
    header: false,
    skipEmptyLines: true
  });

  if (!result.data || result.data.length < 2) return [];

  const normalizeHeader = (value) => String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');

  const scoreHeaderRow = (row) => {
    if (!Array.isArray(row)) return 0;
    const normalized = row.map((cell) => normalizeHeader(cell));
    const hasAny = (tokens) => normalized.some((item) => tokens.some((token) => item.includes(token)));
    let score = 0;
    if (hasAny(['holddate', 'productiondate', 'dateproduced', 'proddate', 'date', 'day'])) score += 5;
    if (hasAny(['casesproduced'])) score += 2;
    if (hasAny(['casesreworked'])) score += 2;
    if (hasAny(['rootcause', 'cause', 'reason'])) score += 1;
    if (hasAny(['disposition', 'status'])) score += 1;
    return score;
  };

  const maxScan = Math.min(result.data.length, 40);
  let headerRowIndex = 0;
  let bestScore = -1;
  for (let i = 0; i < maxScan; i += 1) {
    const score = scoreHeaderRow(result.data[i]);
    if (score > bestScore) {
      bestScore = score;
      headerRowIndex = i;
    }
  }

  if (bestScore < 5) {
    throw new Error('CSV must contain a Date column');
  }

  const header = result.data[headerRowIndex].map(h => String(h).toLowerCase().trim());
  const normalizedHeader = header.map(h => normalizeHeader(h));

  // Print normalized header array for debugging
  if (window && window.console) {
    console.log('DEBUG: Header row index selected:', headerRowIndex);
    console.log('DEBUG: Normalized header array:', normalizedHeader);
  }

  if (window && window.console) {
    console.log('Parsed CSV header:', header);
  }
  const data = [];

  // Resolve column indexes once
  const getColIndex = (keywords) => {
    const normKeywords = keywords.map((keyword) => normalizeHeader(keyword));
    return normalizedHeader.findIndex(h => normKeywords.some(k => h.includes(k)));
  };

  const prodDateCol = getColIndex(['production date', 'product date', 'date produced', 'prod date']);
  const holdDateCol = getColIndex(['hold date', 'date placed', 'date held', 'hold']);
  const dateCol = holdDateCol !== -1 ? holdDateCol : (prodDateCol !== -1 ? prodDateCol : getColIndex(['date', 'day']));

  const descCol = getColIndex(['description', 'item description', 'desc', 'product', 'item']);
  const typeCol = getColIndex(['type', 'item type', 'category']);
  const plantNameCol = getColIndex(['plantname']);
  const workCenterCol = getColIndex(['workcentertext', 'work center text']);
  const siteCol = getColIndex(['site']);
  // Robustly find the disposition column (case-insensitive, trimmed, supports 'disposition', 'status', etc.)
  const dispositionCol = normalizedHeader.findIndex(h => {
    const norm = h.replace(/\s+/g, '').toLowerCase();
    return norm === 'disposition' || norm === 'status';
  });
  const causeCol = getColIndex(['cause', 'root', 'root cause', 'reason']);
  // Robust mapping for 'Cases Reworked', 'casesProduced', and 'Cost' columns
  const normalize = h => h.replace(/\s+/g, '').toLowerCase();
  const workOrderCol = header.findIndex(h => {
    const norm = String(h).replace(/[^a-z0-9]/gi, '').toLowerCase();
    return norm === 'mfgord' || norm === 'mfgorder' || norm === 'wo' || norm === 'workorder' || norm === 'workordernumber' || norm === 'workorderid' || norm.includes('workorder') || norm === 'batchnumber' || norm === 'batch';
  });
  const mfgordCol = header.findIndex(h => {
    const norm = String(h).replace(/[^a-z0-9]/gi, '').toLowerCase();
    return norm === 'mfgord' || norm === 'mfgorder';
  });
  const casesProducedCol = header.findIndex(h => normalize(h) === 'casesproduced');
  const casesReworkedCol = header.findIndex(h => normalize(h) === 'casesreworked');
  // Find cost column: prefer 'cost', but fallback to '$ scrap', '$ rework', etc.
  let costCol = header.findIndex(h => normalize(h) === 'cost');
  if (costCol === -1) {
    costCol = header.findIndex(h => normalize(h).includes('scrap') && h.includes('$'));
    if (costCol === -1) {
      costCol = header.findIndex(h => normalize(h).includes('rework') && h.includes('$'));
    }
  }
  const reworkPercentCol = getColIndex(['rework %', 'rework percent', 'percentage', 'percent']);
  const costImpactCol = header.findIndex(h => normalize(h).includes('costimpact'));
  const reworkLagCol = header.findIndex(h => normalize(h).includes('reworklag'));
  const goalReworkCostCol = getColIndex(['rework cost', 'goal rework cost', 'target rework cost']);
  const goalReleaseRateCol = getColIndex(['release rate', 'goal release rate', 'target release rate']);
  const goalRootCauseAssignmentCol = getColIndex(['root cause assignment', 'goal root cause assignment', 'target root cause assignment']);

  // Place debug log after all column index initialization
  if (window && window.console) {
    console.log('Disposition column index:', dispositionCol);
    console.log('DEBUG: mfgordCol index:', mfgordCol, 'header value:', mfgordCol !== -1 ? header[mfgordCol] : 'NOT FOUND');
    console.log('DEBUG: workOrderCol index:', workOrderCol, 'header value:', workOrderCol !== -1 ? header[workOrderCol] : 'NOT FOUND');
    console.log('DEBUG: reworkLagCol index:', reworkLagCol, 'header value:', reworkLagCol !== -1 ? header[reworkLagCol] : 'NOT FOUND');
    if (casesProducedCol !== -1 && header[casesProducedCol] !== undefined) {
      console.log('DEBUG: Detected casesProducedCol index:', casesProducedCol, 'header:', header[casesProducedCol]);
    } else {
      console.warn('WARNING: casesProducedCol not found in header!');
    }
    if (casesReworkedCol !== -1 && header[casesReworkedCol] !== undefined) {
      console.log('DEBUG: Detected casesReworkedCol index:', casesReworkedCol, 'header:', header[casesReworkedCol]);
    } else {
      console.warn('WARNING: casesReworkedCol not found in header!');
    }
  }

  if (dateCol === -1) {
    throw new Error('CSV must contain a Date column');
  }

  for (let i = headerRowIndex + 1; i < result.data.length; i++) {
    const row = result.data[i].map(v => String(v).trim());
    if (row.every(v => !v)) continue; // Skip completely empty rows

    const dateValue = row[dateCol];
    if (!dateValue) continue;

    // Parse produced, reworked, and cost from correct columns, robustly removing commas, $, spaces, and quotes
    const casesProduced = casesProducedCol !== -1 ? parseInt(String(row[casesProducedCol]).replace(/,/g, '')) || 0 : 0;
    const casesReworked = casesReworkedCol !== -1 ? parseInt(String(row[casesReworkedCol]).replace(/,/g, '')) || 0 : 0;
    let cost = 0;
    if (costCol !== -1) {
      let rawCost = String(row[costCol]).replace(/[$,\s\"]/g, '');
      // Remove any remaining non-numeric except dot and minus
      rawCost = rawCost.replace(/[^0-9.\-]/g, '');
      cost = parseFloat(rawCost) || 0;
    }

    let costImpact = cost;
    if (costImpactCol !== -1) {
      let rawImpact = String(row[costImpactCol]).replace(/[$,\s\"]/g, '');
      rawImpact = rawImpact.replace(/[^0-9.\-]/g, '');
      costImpact = parseFloat(rawImpact) || cost;
    }

    const goalReworkCost = goalReworkCostCol !== -1 ? parseGoalNumber(row[goalReworkCostCol]) : NaN;
    const goalReleaseRate = goalReleaseRateCol !== -1 ? parseGoalNumber(row[goalReleaseRateCol]) : NaN;
    const goalRootCauseAssignment = goalRootCauseAssignmentCol !== -1 ? parseGoalNumber(row[goalRootCauseAssignmentCol]) : NaN;

    // Assign costRework and costScrap per row for correct aggregation
    let costRework = 0, costScrap = 0;
    const disp = dispositionCol !== -1 ? String(row[dispositionCol]).trim().toLowerCase() : '';
    const unitCost = casesProduced > 0 ? cost / casesProduced : 0;
    
    // For rework: calculate unit cost, then multiply by reworked cases
    if (disp === 'rework' || disp.startsWith('rework ')) {
      costRework = unitCost * casesReworked;
    }
    // For scrap: calculate unit cost, then multiply by scrapped cases
    if (disp === 'scrap' || disp === 'scrapped' || disp.startsWith('scrap ')) {
      const scrappedCases = Math.max(casesProduced - casesReworked, 0);
      costScrap = unitCost * scrappedCases;
    }
    const mfgordValue = mfgordCol !== -1 ? row[mfgordCol] : (row[1] || undefined);
    const woValue = mfgordValue || (workOrderCol !== -1 ? row[workOrderCol] : undefined);

    data.push({
      wo: woValue,
      workOrder: woValue,
      mfgOrder: woValue,
      mfgord: mfgordValue,
      productionDate: prodDateCol !== -1 ? row[prodDateCol] : '',
      holdDate: holdDateCol !== -1 ? row[holdDateCol] : dateValue,
      description: descCol !== -1 ? row[descCol] : 'Unknown',
      itemType: typeCol !== -1 ? row[typeCol] : 'Unknown',
      disposition: dispositionCol !== -1 ? row[dispositionCol] : undefined,
      location:
        plantNameCol !== -1 && row[plantNameCol]
          ? row[plantNameCol]
          : (workCenterCol !== -1 && row[workCenterCol]
              ? row[workCenterCol]
              : (siteCol !== -1 && row[siteCol]
                  ? row[siteCol]
                  : 'Unknown')),
      rootCause: causeCol !== -1 ? row[causeCol] : 'Unknown',
      casesProduced: casesProduced,
      casesReworked: casesReworked,
      cost: cost,
      costImpact: costImpact,
      reworkLag: reworkLagCol !== -1 ? row[reworkLagCol] : undefined,
      costRework: costRework,
      costScrap: costScrap,
      goalReworkCost,
      goalReleaseRate,
      goalRootCauseAssignment
      // Removed per-row reworkPercent for clarity in overall metric
    });
  }

  // Debug: log first 3 parsed records to verify mfgord values
  if (window && window.console && data.length > 0) {
    console.log('DEBUG: First 3 parsed records:');
    data.slice(0, 3).forEach((d, i) => {
      console.log(`  Record ${i}:`, {
        mfgord: d.mfgord,
        wo: d.wo,
        workOrder: d.workOrder,
        rootCause: d.rootCause,
        disposition: d.disposition,
        reworkLag: d.reworkLag
      });
    });
  }

  return data;
}

// ---------- Storage ----------
function saveData() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(reworkData));
    updateDataInfo();
  } catch (e) {
    // Storage quota exceeded - skip saving, data will need to be re-uploaded
    if (window && window.console) {
      console.warn('Data too large for localStorage. You will need to re-upload the CSV file on page reload.');
    }
  }
}

function loadStoredData() {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (stored) {
    try {
      reworkData = JSON.parse(stored);
    } catch {
      reworkData = [];
    }
  }
}

function clearData() {
  if (!confirm('Clear all rework data? This cannot be undone.')) return;
  reworkData = [];
  localStorage.removeItem(STORAGE_KEY);
  renderDashboard();
  showMessage('All data cleared', 'info');
}

function populateLocationFilter() {
  const locations = new Set();
  reworkData.forEach((d) => {
    if (d.location) locations.add(d.location);
  });
  const dropdownList = document.getElementById('dropdownList');
  const dropdownSelected = document.getElementById('dropdownSelected');
  if (!dropdownList || !dropdownSelected) return;
  // Save previous selection
  const prevSelected = Array.from(dropdownList.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value);
  // Clear list
  dropdownList.innerHTML = '';
  // Add "All Locations" checkbox
  const allLabel = document.createElement('label');
  const allCheckbox = document.createElement('input');
  allCheckbox.type = 'checkbox';
  allCheckbox.value = '';
  allCheckbox.checked = prevSelected.length === 0;
  allLabel.appendChild(allCheckbox);
  allLabel.appendChild(document.createTextNode('All Locations'));
  dropdownList.appendChild(allLabel);
  // Add sorted locations
  const sortedLocations = Array.from(locations).sort();
  sortedLocations.forEach((loc) => {
    const label = document.createElement('label');
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.value = loc;
    checkbox.checked = prevSelected.includes(loc);
    label.appendChild(checkbox);
    label.appendChild(document.createTextNode(loc));
    dropdownList.appendChild(label);
  });
  // Update selected text
  const updateSelectedText = () => {
    const checked = Array.from(dropdownList.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value);
    if (checked.length === 0 || checked.includes('')) {
      dropdownSelected.textContent = 'All Locations';
    } else {
      dropdownSelected.textContent = checked.length === 1 ? checked[0] : `${checked.length} selected`;
    }
  };
  dropdownList.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    cb.addEventListener('change', () => {
      // If "All Locations" is checked, uncheck others
      if (cb.value === '') {
        if (cb.checked) {
          dropdownList.querySelectorAll('input[type="checkbox"]').forEach(other => {
            if (other.value !== '') other.checked = false;
          });
        }
      } else {
        // If any location is checked, uncheck "All Locations"
        if (cb.checked) {
          dropdownList.querySelector('input[type="checkbox"][value=""]').checked = false;
        }
      }
      updateSelectedText();
      renderDashboard();
    });
  });
  updateSelectedText();
  // Dropdown open/close logic
  dropdownSelected.onclick = () => {
    dropdownList.style.display = dropdownList.style.display === 'none' ? 'block' : 'none';
  };
  document.addEventListener('click', (e) => {
    if (!dropdownList.contains(e.target) && !dropdownSelected.contains(e.target)) {
      dropdownList.style.display = 'none';
    }
  });
}

// ---------- UI Helpers ----------
function showMessage(msg, type) {
  const msgDiv = document.getElementById('message');
  msgDiv.className = type;
  msgDiv.textContent = msg;
  msgDiv.style.display = 'block';
  let timeout = 0;
  if (type === 'error') {
    timeout = 10000;
  } else {
    timeout = 4000;
  }
  if (timeout > 0) {
    clearTimeout(msgDiv._hideTimeout);
    msgDiv._hideTimeout = setTimeout(() => {
      msgDiv.style.display = 'none';
    }, timeout);
  }
}

function updateDataInfo() {
  const info = reworkData.length > 0 ? `${reworkData.length} records loaded` : 'No data loaded';
  document.getElementById('dataInfo').textContent = info;
}

function hookUpDateEvents() {
  const start = document.getElementById('startDate');
  const end = document.getElementById('endDate');
  const location = document.getElementById('locationFilter');
  const costTimeView = document.getElementById('costTimeView');
  if (start) start.addEventListener('change', renderDashboard);
  if (end) end.addEventListener('change', renderDashboard);
  if (location) location.addEventListener('change', renderDashboard);
  if (location) location.addEventListener('input', renderDashboard);
  if (costTimeView) costTimeView.addEventListener('change', renderDashboard);
}

// ---------- Dates & Filtering ----------
function parseFlexibleDate(s) {
  if (!s) return null;
  const only = String(s).trim().split(' ')[0];

  // YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(only)) {
    const [y, m, d] = only.split('-').map(Number);
    return new Date(y, m - 1, d);
  }
  // MM/DD/YYYY or M/D/YY
  if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(only)) {
    const [m, d, yRaw] = only.split('/').map(Number);
    const y = yRaw < 100 ? 2000 + yRaw : yRaw;
    return new Date(y, m - 1, d);
  }
  // Fallback
  const dt = new Date(only);
  return isNaN(dt) ? null : dt;
}

function pad2(value) {
  return String(value).padStart(2, '0');
}

function formatYMD(date) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

function getWeekStart(date) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const dayIndex = (d.getDay() + 6) % 7; // Monday = 0
  d.setDate(d.getDate() - dayIndex);
  d.setHours(0, 0, 0, 0);
  return d;
}

function getFilteredData() {
  const startVal = document.getElementById('startDate')?.value || '';
  const endVal = document.getElementById('endDate')?.value || '';
  const dropdownList = document.getElementById('dropdownList');
  let locationVals = [];
  if (dropdownList) {
    locationVals = Array.from(dropdownList.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value).filter(v => v);
  }
  const start = startVal ? new Date(startVal) : null;
  const end = endVal ? new Date(endVal) : null;
  // end-of-day for inclusive filtering
  const endEOD = end ? new Date(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59, 999) : null;

  const filtered = reworkData.filter((d) => {
    // Date filtering
    const source = d.holdDate || d.productionDate;
    const dt = parseFlexibleDate(source);
    if (!dt) return false;
    if (start && dt < start) return false;
    if (endEOD && dt > endEOD) return false;
    // Location filtering
    if (locationVals.length > 0 && !locationVals.includes(d.location)) return false;
    return true;
  });

  updateDateRangeDisplay(start, end, filtered.length);
  return filtered;
}

function updateDateRangeDisplay(start, end, count) {
  const el = document.getElementById('dateRangeDisplay');
  const fmt = (dt) => dt?.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
  if (!start && !end) {
    el.textContent = `Showing all records (${count})`;
  } else if (start && end) {
    el.textContent = `Showing ${fmt(start)} → ${fmt(end)} (${count} records)`;
  } else if (start) {
    el.textContent = `From ${fmt(start)} onward (${count} records)`;
  } else {
    el.textContent = `Up to ${fmt(end)} (${count} records)`;
  }
}

// ---------- Render ----------
function renderDashboard() {
  const active = getFilteredData();

  if (reworkData.length === 0) {
    showMessage('No data loaded. Please upload a CSV file.', 'error');
  }

  if (active.length === 0) {
    renderEmptyState();
    updateDataInfo();
    showMessage('No records match your filters or the file failed to load.', 'error');
    return;
  }

  showMessage(`✓ Successfully loaded ${active.length} records.`, 'success');
  renderPerformanceGoals(active);
  calculateMetrics(active);
  renderCharts(active);
  updateDataInfo();
}

function renderEmptyState() {
  renderPerformanceGoals([]);
  document.getElementById('totalItems').textContent = '0';
  document.getElementById('reworkPercent').textContent = '0%';
  document.getElementById('topCause').textContent = '—';
  const daysTrackedField = document.getElementById('daysTracked');
  if (daysTrackedField) {
    daysTrackedField.textContent = '0';
  }
  const defectsBody = document.getElementById('defectsBody') || document.querySelector('#defectDriversTable tbody');
  if (defectsBody) {
    defectsBody.innerHTML =
      '<tr><td colspan="5" style="text-align: center; color: #999;">No data loaded</td></tr>';
  }

  if (trendChart) trendChart.destroy();
  if (causeChart) causeChart.destroy();
}

function getPerformanceStatus(current, goal, direction) {
  if (!Number.isFinite(current) || !Number.isFinite(goal)) {
    return { text: 'N/A', className: '' };
  }

  const delta = current - goal;
  const absDelta = Math.abs(delta);
  const atGoalTolerance = Math.max(0.2, Math.abs(goal) * 0.01);
  const nearTolerance = Math.max(1.0, Math.abs(goal) * 0.05);

  if (absDelta <= atGoalTolerance) {
    return { text: 'At Goal', className: 'status-at-goal' };
  }

  if (direction === 'higher') {
    if (delta > atGoalTolerance) {
      return { text: 'Exceeds Goal', className: 'status-on-track' };
    }
    if (Math.abs(delta) <= nearTolerance || current >= goal - nearTolerance) {
      return { text: 'Approaching', className: 'status-approaching' };
    }
    return { text: 'Off Track', className: 'status-off-track' };
  }

  if (delta < -atGoalTolerance) {
    return { text: 'Exceeds Goal', className: 'status-on-track' };
  }
  if (Math.abs(delta) <= nearTolerance || current <= goal + nearTolerance) {
    return { text: 'Approaching', className: 'status-approaching' };
  }
  return { text: 'Off Track', className: 'status-off-track' };
}

function formatPerformanceValue(value, formatType) {
  if (!Number.isFinite(value)) return 'N/A';
  if (formatType === 'currency') {
    return `$${value.toLocaleString(undefined, { maximumFractionDigits: 2 })}`;
  }
  return `${value.toFixed(1)}%`;
}

function getGoalValueFromData(data, propertyName) {
  const candidate = data
    .map((row) => Number(row[propertyName]))
    .find((value) => Number.isFinite(value));
  return Number.isFinite(candidate) ? candidate : NaN;
}

function calculatePopulationFactor(filteredData, fullData) {
  const safeFiltered = Array.isArray(filteredData) ? filteredData : [];
  const safeFull = Array.isArray(fullData) ? fullData : [];

  const filteredHoldUnits = safeFiltered.reduce((sum, row) => sum + (parseInt(row.casesProduced) || 0), 0);
  const fullHoldUnits = safeFull.reduce((sum, row) => sum + (parseInt(row.casesProduced) || 0), 0);

  if (fullHoldUnits > 0) {
    return Math.max(filteredHoldUnits / fullHoldUnits, 0);
  }

  if (safeFull.length > 0) {
    return Math.max(safeFiltered.length / safeFull.length, 0);
  }

  return 1;
}

function updatePerformanceGoalsNote(populationFactor, filteredCount, fullCount) {
  const noteEl = document.getElementById('performanceGoalsNote');
  if (!noteEl) return;

  if (!Number.isFinite(populationFactor) || fullCount === 0) {
    noteEl.textContent = 'Goals scaled to 100.0% of total population.';
    return;
  }

  const pct = (populationFactor * 100).toFixed(1);
  noteEl.textContent = `Goals scaled to ${pct}% of total population (${filteredCount.toLocaleString()} of ${fullCount.toLocaleString()} records).`;
}

function scaleGoalByPopulation(goalValue, populationFactor) {
  if (!Number.isFinite(goalValue)) return NaN;
  if (!Number.isFinite(populationFactor) || populationFactor <= 0) return goalValue;
  return goalValue * populationFactor;
}

function renderPerformanceGoals(data) {
  const safeData = Array.isArray(data) ? data : [];

  const totalHoldUnits = safeData.reduce((sum, d) => sum + (parseInt(d.casesProduced) || 0), 0);
  const totalReworkCost = safeData.reduce((sum, d) => sum + (Number(d.costRework) || 0), 0);

  let releasedUnits = 0;
  let assignedRootCauseRows = 0;

  safeData.forEach((d) => {
    const disp = (typeof d.disposition === 'string' ? d.disposition : '').toLowerCase().replace(/\s+/g, '');
    const produced = parseInt(d.casesProduced) || 0;
    const reworked = parseInt(d.casesReworked) || 0;
    const rootCause = String(d.rootCause || '').trim().toLowerCase();

    if (disp === 'rework' || disp.includes('release')) {
      releasedUnits += Math.max(produced - reworked, 0);
    }
    if (rootCause && rootCause !== 'unknown' && rootCause !== 'other' && rootCause !== 'n/a') {
      assignedRootCauseRows += 1;
    }
  });

  const releaseRate = totalHoldUnits > 0 ? (releasedUnits / totalHoldUnits) * 100 : NaN;
  const rootCauseAssignment = safeData.length > 0 ? (assignedRootCauseRows / safeData.length) * 100 : NaN;

  const goalReworkCostFromData = getGoalValueFromData(safeData, 'goalReworkCost');
  const goalReleaseRateFromData = getGoalValueFromData(safeData, 'goalReleaseRate');
  const goalRootCauseAssignmentFromData = getGoalValueFromData(safeData, 'goalRootCauseAssignment');

  const effectiveGoalReworkCost = Number.isFinite(goalReworkCostFromData) ? goalReworkCostFromData : DEFAULT_PERFORMANCE_GOALS.reworkCost;
  const effectiveGoalReleaseRate = Number.isFinite(goalReleaseRateFromData) ? goalReleaseRateFromData : DEFAULT_PERFORMANCE_GOALS.releaseRate;
  const effectiveGoalRootCauseAssignment = Number.isFinite(goalRootCauseAssignmentFromData) ? goalRootCauseAssignmentFromData : DEFAULT_PERFORMANCE_GOALS.rootCauseAssignment;

  const populationFactor = calculatePopulationFactor(safeData, reworkData);
  updatePerformanceGoalsNote(populationFactor, safeData.length, reworkData.length);

  const adjustedGoalReworkCost = scaleGoalByPopulation(effectiveGoalReworkCost, populationFactor);
  const adjustedGoalReleaseRate = scaleGoalByPopulation(effectiveGoalReleaseRate, populationFactor);
  const adjustedGoalRootCauseAssignment = scaleGoalByPopulation(effectiveGoalRootCauseAssignment, populationFactor);

  updatePerformanceTile('perfReworkCost', totalReworkCost, adjustedGoalReworkCost, 'lower', 'currency');
  updatePerformanceTile('perfReleaseRate', releaseRate, adjustedGoalReleaseRate, 'lower', 'percent');
  updatePerformanceTile('perfRootCauseAssignment', rootCauseAssignment, adjustedGoalRootCauseAssignment, 'higher', 'percent');
}

function updatePerformanceTile(tileIdPrefix, current, goal, direction, formatType = 'percent') {
  const currentEl = document.getElementById(`${tileIdPrefix}Current`);
  const goalEl = document.getElementById(`${tileIdPrefix}Goal`);
  const statusEl = document.getElementById(`${tileIdPrefix}Status`);
  if (!currentEl || !goalEl || !statusEl) return;

  const normalizedGoal = formatType === 'percent' && Number.isFinite(goal)
    ? Math.min(Math.max(goal, 0), 100)
    : goal;

  currentEl.textContent = formatPerformanceValue(current, formatType);
  goalEl.textContent = formatPerformanceValue(normalizedGoal, formatType);

  const status = getPerformanceStatus(current, normalizedGoal, direction);
  statusEl.className = 'performance-status';
  if (status.className) {
    statusEl.classList.add(status.className);
  }
  statusEl.textContent = status.text;
}

// ---------- Metrics ----------
function calculateMetrics(data) {
      // ...existing code...
    // Debug: print first 10 rows of casesProduced and casesReworked
    if (window && window.console) {
      console.log('Sample casesProduced:', data.slice(0, 10).map(d => d.casesProduced));
      console.log('Sample casesReworked:', data.slice(0, 10).map(d => d.casesReworked));
    }
  // ...existing code...
  // ...existing code...
  // Debug: log first 10 disposition values to check parsing
  if (window && window.console) {
    console.log('First 10 disposition values:', data.slice(0, 10).map(d => d.disposition));
  }

  // Total Hold Units = sum of casesProduced
  const totalHoldUnits = data.reduce((sum, d) => sum + (parseInt(d.casesProduced) || 0), 0);
  // Total Items Reworked = sum of casesReworked
  const totalItems = data.reduce((sum, d) => sum + (parseInt(d.casesReworked) || 0), 0);
  // % Released: For 'rework' or 'release' disposition, released = casesProduced - casesReworked
  let releasedUnits = 0;
  let reworkCost = 0;
  let scrapUnits = 0;
  let scrapCost = 0;
  let debugReworkRows = [];
  let foundRework = false, foundScrap = false;
  let debugReworkSumRows = [];
  data.forEach((d, idx) => {
    let disp = (typeof d.disposition === 'string' ? d.disposition : '').toLowerCase().replace(/\s+/g, '');
    const produced = parseInt(d.casesProduced) || 0;
    const reworked = parseInt(d.casesReworked) || 0;
    // $ Rework: disposition exactly 'rework' (robust to whitespace/case)
    if (disp === 'rework') {
      reworkCost += d.costRework || 0;
      foundRework = true;
      if (debugReworkSumRows.length < 10) debugReworkSumRows.push({row: idx+1, disp: d.disposition, cost: d.costRework});
      releasedUnits += Math.max(produced - reworked, 0);
    }
    // Released: disposition contains 'release'
    else if (disp.includes('release')) {
      releasedUnits += Math.max(produced - reworked, 0);
    }
    // Scrap: disposition contains 'scrap'
    if (disp.includes('scrap')) {
      scrapUnits += Math.max(produced - reworked, 0);
      scrapCost += d.costScrap || 0;
      foundScrap = true;
    }
  });
  if (window && window.console) {
    console.log('DEBUG: First 10 rows summed for $ Rework:', debugReworkSumRows);
  }
  if (window && window.console) {
    console.log('DEBUG: First 10 Rework rows:', debugReworkRows.slice(0, 10));
    console.log('DEBUG: Total reworkCost after loop:', reworkCost);
  }
  const percentReleased = totalHoldUnits > 0 ? ((releasedUnits / totalHoldUnits) * 100).toFixed(1) : 'N/A';
  const percentScrapped = totalHoldUnits > 0 ? ((scrapUnits / totalHoldUnits) * 100).toFixed(1) : 'N/A';
  let reworkPercent = 0;
  if (totalHoldUnits > 0) reworkPercent = ((totalItems / totalHoldUnits) * 100).toFixed(1);
  const uniqueDates = new Set(data.map((d) => d.holdDate || d.productionDate)).size;

  // Debug: print totals after assignment
  if (window && window.console) {
    console.log('DEBUG: totalHoldUnits (should be casesProduced):', totalHoldUnits);
    console.log('DEBUG: totalItems (should be casesReworked):', totalItems);
    console.log('Total releasedUnits:', releasedUnits);
    console.log('Total reworkCost:', reworkCost);
    console.log('Total scrapUnits:', scrapUnits);
    console.log('Total scrapCost:', scrapCost);
    console.log('Calculated Rework %:', reworkPercent);
    console.log('Calculated % Scrapped:', percentScrapped);
  }

  const causeCounts = {};
  data.forEach((d) => {
    const cause = d.rootCause || 'Unknown';
    causeCounts[cause] = (causeCounts[cause] || 0) + (parseInt(d.casesReworked) || 0);
  });
  const topCause =
    Object.keys(causeCounts).length > 0
      ? Object.entries(causeCounts).sort((a, b) => b[1] - a[1])[0][0]
      : '—';

  // Assign metrics to dashboard fields (executive summary)
  document.getElementById('totalHoldUnits').textContent = totalHoldUnits.toLocaleString(); // casesProduced
  document.getElementById('totalItems').textContent = totalItems.toLocaleString(); // casesReworked
  document.getElementById('percentReleased').textContent = percentReleased !== 'N/A' ? `${percentReleased}%` : 'N/A';
  document.getElementById('reworkPercent').textContent = totalHoldUnits > 0 ? `${reworkPercent}%` : 'N/A';
  document.getElementById('topCause').textContent = (topCause || '—').substring(0, 25);
  const daysTrackedField = document.getElementById('daysTracked');
  if (daysTrackedField) {
    daysTrackedField.textContent = uniqueDates.toLocaleString();
  }
  // Cost of Rework and Scrap (executive summary fields)
  // Robustly check for all metric elements and show visible error if missing
  const reworkCostField = document.getElementById('costRework');
  if (reworkCostField) {
    if (foundRework && reworkCost > 0) {
      reworkCostField.textContent = `$${reworkCost.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}`;
    } else {
      reworkCostField.textContent = 'N/A';
    }
  } else {
    showMessage('Dashboard error: $ Rework metric element (id="costRework") is missing from the HTML. Please add <div class="metric-value" id="costRework"></div> to your metrics section.', 'error');
    console.error('Dashboard error: #costRework element not found in DOM.');
  }

  const scrapCostField = document.getElementById('costScrap');
  if (scrapCostField) {
    if (foundScrap && scrapCost > 0) {
      scrapCostField.textContent = `$${scrapCost.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}`;
    } else {
      scrapCostField.textContent = 'N/A';
    }
  } else {
    showMessage('Dashboard error: $ Scrap metric element (id="costScrap") is missing from the HTML. Please add <div class="metric-value" id="costScrap"></div> to your metrics section.', 'error');
    console.error('Dashboard error: #costScrap element not found in DOM.');
  }

  const percentScrappedField = document.getElementById('percentScrapped');
  if (percentScrappedField) {
    percentScrappedField.textContent = percentScrapped !== 'N/A' ? `${percentScrapped}%` : 'N/A';
  } else {
    showMessage('Dashboard error: % Scrapped metric element (id="percentScrapped") is missing from the HTML. Please add <div class="metric-value" id="percentScrapped"></div> to your metrics section.', 'error');
    console.error('Dashboard error: #percentScrapped element not found in DOM.');
  }
}

// ---------- Charts ----------
function renderCharts(data) {
  renderDispositionDonut(data);
  renderTimeCostBar(data);
  renderRootCauseDonut(data);
  renderDefectParetoBar(data);
  renderDefectDriversTable(data);
}

// Disposition Mix Donut Chart
function renderDispositionDonut(data) {
  const ctx = document.getElementById('dispositionDonut');
  if (!ctx) return;

  const parseNumberText = (text) => {
    if (!text) return NaN;
    const cleaned = String(text).replace(/[^0-9.\-]/g, '');
    return parseFloat(cleaned);
  };

  const totalHoldUnits = parseNumberText(document.getElementById('totalHoldUnits')?.textContent);
  const percentReleased = parseNumberText(document.getElementById('percentReleased')?.textContent);
  const percentReworked = parseNumberText(document.getElementById('reworkPercent')?.textContent);
  const percentScrapped = parseNumberText(document.getElementById('percentScrapped')?.textContent);

  let mix = { Released: 0, Reworked: 0, Scrapped: 0 };

  // Primary method: use already-calculated KPI percentages × total hold units
  if (
    Number.isFinite(totalHoldUnits) && totalHoldUnits > 0 &&
    Number.isFinite(percentReleased) && Number.isFinite(percentReworked) && Number.isFinite(percentScrapped)
  ) {
    mix = {
      Released: (totalHoldUnits * percentReleased) / 100,
      Reworked: (totalHoldUnits * percentReworked) / 100,
      Scrapped: (totalHoldUnits * percentScrapped) / 100
    };
  } else {
    // Fallback: robust disposition parsing from raw data
    data.forEach((d) => {
      const disp = (typeof d.disposition === 'string' ? d.disposition : '').trim().toLowerCase();
      const produced = parseInt(d.casesProduced) || 0;
      const reworked = parseInt(d.casesReworked) || 0;
      if (disp.includes('release')) mix.Released += Math.max(produced - reworked, 0);
      if (disp.includes('rework')) mix.Reworked += reworked;
      if (disp.includes('scrap')) mix.Scrapped += Math.max(produced - reworked, 0);
    });
  }

  if (window.dispositionDonutChart) window.dispositionDonutChart.destroy();
  window.dispositionDonutChart = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: Object.keys(mix),
      datasets: [{
        data: Object.values(mix),
        backgroundColor: ['#b39b73', '#d94b59', '#7f1320'],
        radius: '84%',
        cutout: '58%'
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      aspectRatio: 1.2,
      plugins: {
        legend: { position: 'bottom' },
        tooltip: {
          callbacks: {
            label: (context) => {
              const label = context.label || 'Disposition';
              const value = Number(context.raw) || 0;
              return `${label}: ${value.toLocaleString()} cases`;
            }
          }
        }
      }
    },
  });
}

// Time vs Cost Bar Chart
function renderTimeCostBar(data) {
  const ctx = document.getElementById('timeCostBar');
  if (!ctx) return;
  const granularity = (document.getElementById('costTimeView')?.value || 'month').toLowerCase();
  const bucketMap = new Map();

  data.forEach((d) => {
    const sourceDate = d.holdDate || d.productionDate;
    const dt = parseFlexibleDate(sourceDate);
    if (!dt) return;

    let key = '';
    let label = '';
    let sortKey = 0;

    if (granularity === 'day') {
      key = formatYMD(dt);
      label = dt.toLocaleDateString(undefined, { month: '2-digit', day: '2-digit', year: 'numeric' });
      sortKey = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate()).getTime();
    } else if (granularity === 'week') {
      const weekStart = getWeekStart(dt);
      key = `week-${formatYMD(weekStart)}`;
      label = `Week of ${weekStart.toLocaleDateString(undefined, { month: '2-digit', day: '2-digit', year: 'numeric' })}`;
      sortKey = weekStart.getTime();
    } else {
      const monthStart = new Date(dt.getFullYear(), dt.getMonth(), 1);
      key = `month-${dt.getFullYear()}-${pad2(dt.getMonth() + 1)}`;
      label = monthStart.toLocaleDateString(undefined, { month: 'short', year: 'numeric' });
      sortKey = monthStart.getTime();
    }

    if (!bucketMap.has(key)) {
      bucketMap.set(key, { label, sortKey, rework: 0, scrap: 0 });
    }

    const bucket = bucketMap.get(key);
    const disp = (typeof d.disposition === 'string' ? d.disposition : '').toLowerCase();
    if (disp.includes('rework')) bucket.rework += d.costRework || 0;
    if (disp.includes('scrap')) bucket.scrap += d.costScrap || 0;
  });

  const sortedBuckets = Array.from(bucketMap.values()).sort((a, b) => a.sortKey - b.sortKey);
  const labels = sortedBuckets.map((b) => b.label);
  const reworkData = sortedBuckets.map((b) => b.rework);
  const scrapData = sortedBuckets.map((b) => b.scrap);

  const xAxisTitle = granularity === 'day'
    ? 'Date (Day)'
    : (granularity === 'week' ? 'Week' : 'Month');

  if (window.timeCostBarChart) window.timeCostBarChart.destroy();
  window.timeCostBarChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Rework', data: reworkData, backgroundColor: '#b39b73' },
        { label: 'Scrap', data: scrapData, backgroundColor: '#b21f2d' },
      ],
    },
    options: {
      responsive: true,
      plugins: { legend: { position: 'top' } },
      scales: {
        x: {
          stacked: true,
          title: { display: true, text: xAxisTitle }
        },
        y: {
          stacked: true,
          beginAtZero: true,
          title: { display: true, text: 'Cost Impact ($)' }
        }
      }
    },
  });
}

// Root Cause Breakdown Donut Chart
function renderRootCauseDonut(data) {
  const ctx = document.getElementById('rootCauseDonut');
  if (!ctx) return;
  const counts = {};
  data.forEach((d) => {
    const cause = d.rootCause || 'Unknown';
    counts[cause] = (counts[cause] || 0) + (d.casesReworked || 0);
  });

  // Keep donut in line with Pareto: same ranking basis (top root causes by cases reworked)
  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  const top = sorted.slice(0, 10);
  const otherTotal = sorted.slice(10).reduce((sum, [, value]) => sum + value, 0);

  const labels = top.map(([cause]) => cause);
  const values = top.map(([, value]) => value);
  if (otherTotal > 0) {
    labels.push('Other');
    values.push(otherTotal);
  }

  if (window.rootCauseDonutChart) window.rootCauseDonutChart.destroy();
  window.rootCauseDonutChart = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: ['#b21f2d', '#d94b59', '#7f1320', '#b39b73', '#8b6b3f', '#f0dcb5', '#d4c2a1', '#e7d9c1', '#c58e72', '#a85656', '#9ca3af'],
        radius: '82%',
        cutout: '60%'
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      aspectRatio: 1.45,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { boxWidth: 12, font: { size: 11 } }
        },
        tooltip: {
          callbacks: {
            label: (context) => `${context.label}: ${Number(context.raw || 0).toLocaleString()} cases`
          }
        }
      }
    },
  });
}

// Pareto Bar Chart of Top Defect Drivers
function renderDefectParetoBar(data) {
  const ctx = document.getElementById('defectParetoBar');
  if (!ctx) return;
  const counts = {};
  data.forEach(d => {
    const driver = d.rootCause || 'Unknown';
    counts[driver] = (counts[driver] || 0) + (d.casesReworked || 0);
  });
  const sorted = Object.entries(counts).sort((a,b) => b[1]-a[1]).slice(0,10);
  const labels = sorted.map(([k]) => k);
  const values = sorted.map(([_,v]) => v);
  if (window.defectParetoBarChart) window.defectParetoBarChart.destroy();
  window.defectParetoBarChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Cases Reworked',
        data: values,
        backgroundColor: '#b39b73',
      }],
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (context) => `Cases Reworked: ${Number(context.raw || 0).toLocaleString()} cases`
          }
        }
      },
      indexAxis: 'y',
      scales: {
        x: {
          beginAtZero: true,
          title: { display: true, text: 'Cases Reworked (Cases)' },
          ticks: {
            callback: (value) => Number(value).toLocaleString()
          }
        },
        y: {
          title: { display: true, text: 'Root Cause' }
        }
      }
    },
  });
}

// Table of Top Defect Drivers
function renderDefectDriversTable(data) {
  const tbody = document.querySelector('#defectDriversTable tbody');
  if (!tbody) return;
  
  // Sort based on selected mode
  const sorted = data.slice().sort((a,b) => {
    if (defectDriversSortMode === 'cost') {
      return (b.cost||0)-(a.cost||0);
    } else if (defectDriversSortMode === 'lag') {
      const lagA = Number(String(a.reworkLag || '0').replace(/[^0-9.\-]/g, '')) || 0;
      const lagB = Number(String(b.reworkLag || '0').replace(/[^0-9.\-]/g, '')) || 0;
      return lagB - lagA;
    }
    // Default: sort by cases held (casesProduced)
    return (b.casesProduced||0)-(a.casesProduced||0);
  }).slice(0,10);
  
  // Debug: log what values the table is receiving
  if (window && window.console && sorted.length > 0) {
    console.log('DEBUG: renderDefectDriversTable - First 3 sorted records:');
    sorted.slice(0, 3).forEach((d, i) => {
      console.log(`  Record ${i}:`, {
        mfgord: d.mfgord,
        mfgOrder: d.mfgOrder,
        wo: d.wo,
        workOrder: d.workOrder,
        rootCause: d.rootCause,
        reworkLag: d.reworkLag
      });
    });
  }
  
  tbody.innerHTML = sorted.map(d => {
    const wo = String(d.mfgord || d.mfgOrder || d.wo || d.workOrder || '').trim() || '—';
    const date = d.holdDate || d.productionDate || '—';
    const cases = Number(d.casesProduced || 0).toLocaleString();
    const cause = d.rootCause || '—';
    const disp = d.disposition || '—';
    const cost = d.cost || 0;
    const lagRaw = String(d.reworkLag || '').trim();
    const lagNum = lagRaw ? Number(lagRaw.replace(/[^0-9.\-]/g, '')) : NaN;
    const age = Number.isFinite(lagNum) ? `${lagNum} days` : (lagRaw || '—');
    return `<tr><td>${wo}</td><td>${date}</td><td>${cases}</td><td>${cause}</td><td>${disp}</td><td>$${Number(cost).toLocaleString()}</td><td>${age}</td></tr>`;
  }).join('');
  
  // Update sort icons
  const casesIcon = document.getElementById('casesSortIcon');
  const costIcon = document.getElementById('costSortIcon');
  const lagIcon = document.getElementById('lagSortIcon');
  
  if (casesIcon) casesIcon.textContent = defectDriversSortMode === 'cases' ? '↓' : '';
  if (costIcon) costIcon.textContent = defectDriversSortMode === 'cost' ? '↓' : '';
  if (lagIcon) lagIcon.textContent = defectDriversSortMode === 'lag' ? '↓' : '';
}

// Add click handlers for sorting
function setupDefectDriversSorting() {
  const casesHeader = document.getElementById('casesHeader');
  const costHeader = document.getElementById('costHeader');
  const lagHeader = document.getElementById('lagHeader');
  
  if (casesHeader) {
    casesHeader.addEventListener('click', () => {
      defectDriversSortMode = 'cases';
      renderDashboard();
    });
  }
  
  if (costHeader) {
    costHeader.addEventListener('click', () => {
      defectDriversSortMode = 'cost';
      renderDashboard();
    });
  }
  
  if (lagHeader) {
    lagHeader.addEventListener('click', () => {
      defectDriversSortMode = 'lag';
      renderDashboard();
    });
  }
}

function setupAIInsights() {
  const button = document.getElementById('generate-ai-insights-btn');
  const refreshButton = document.getElementById('refresh-ai-health-btn');

  if (refreshButton) {
    refreshButton.addEventListener('click', () => {
      checkAIBackendHealth(true);
    });
  }

  checkAIBackendHealth(false);

  if (!button) return;
  button.addEventListener('click', generateAIInsights);
}

function setAIHealthStatus(state) {
  const statusEl = document.getElementById('ai-health-status');
  if (!statusEl) return;

  statusEl.classList.remove('online', 'offline', 'checking');

  if (state === 'online') {
    statusEl.classList.add('online');
    statusEl.textContent = 'Backend: Connected';
    return;
  }

  if (state === 'checking') {
    statusEl.classList.add('checking');
    statusEl.textContent = 'Backend: Checking...';
    return;
  }

  statusEl.classList.add('offline');
  statusEl.textContent = 'Backend: Offline';
}

async function checkAIBackendHealth(showErrorMessage) {
  setAIHealthStatus('checking');
  try {
    const response = await fetch(AI_HEALTH_ENDPOINT, {
      method: 'GET',
      headers: { 'Content-Type': 'application/json' }
    });
    if (!response.ok) {
      throw new Error(`Health check failed with status ${response.status}`);
    }
    setAIHealthStatus('online');
  } catch (error) {
    setAIHealthStatus('offline');
    if (showErrorMessage) {
      const result = document.getElementById('ai-insights-result');
      if (result) {
        result.innerHTML = 'Backend is not reachable. Start the API at <strong>http://localhost:3001</strong> and refresh status.';
      }
    }
  }
}

function buildAIContext(data) {
  const safeData = Array.isArray(data) ? data : [];
  const totalHoldUnits = safeData.reduce((sum, d) => sum + (parseInt(d.casesProduced) || 0), 0);
  const totalItemsReworked = safeData.reduce((sum, d) => sum + (parseInt(d.casesReworked) || 0), 0);
  const reworkPercent = totalHoldUnits > 0 ? (totalItemsReworked / totalHoldUnits) * 100 : 0;

  let scrapUnits = 0;
  const locationCounts = {};
  const causeCounts = {};
  const dispositionCounts = {};

  safeData.forEach((d) => {
    const disposition = String(d.disposition || 'Unknown').trim();
    const dispositionKey = disposition || 'Unknown';
    dispositionCounts[dispositionKey] = (dispositionCounts[dispositionKey] || 0) + 1;

    const location = String(d.location || 'Unknown').trim() || 'Unknown';
    locationCounts[location] = (locationCounts[location] || 0) + (parseInt(d.casesReworked) || 0);

    const cause = String(d.rootCause || 'Unknown').trim() || 'Unknown';
    causeCounts[cause] = (causeCounts[cause] || 0) + (parseInt(d.casesReworked) || 0);

    const dispNormalized = disposition.toLowerCase();
    if (dispNormalized.includes('scrap')) {
      const produced = parseInt(d.casesProduced) || 0;
      const reworked = parseInt(d.casesReworked) || 0;
      scrapUnits += Math.max(produced - reworked, 0);
    }
  });

  const topRootCause = Object.entries(causeCounts).sort((a, b) => b[1] - a[1])[0]?.[0] || 'Unknown';
  const topLocation = Object.entries(locationCounts).sort((a, b) => b[1] - a[1])[0]?.[0] || 'Unknown';
  const topDisposition = Object.entries(dispositionCounts).sort((a, b) => b[1] - a[1])[0]?.[0] || 'Unknown';
  const percentScrapped = totalHoldUnits > 0 ? (scrapUnits / totalHoldUnits) * 100 : 0;

  return {
    totalHoldUnits,
    totalItemsReworked,
    reworkPercent,
    percentScrapped,
    topRootCause,
    topLocation,
    topDisposition
  };
}

function buildAISummaryText(data, metrics) {
  const causeCounts = {};
  data.forEach((d) => {
    const cause = String(d.rootCause || 'Unknown').trim() || 'Unknown';
    causeCounts[cause] = (causeCounts[cause] || 0) + (parseInt(d.casesReworked) || 0);
  });

  const topCauses = Object.entries(causeCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([cause, count]) => `${cause}: ${Number(count).toLocaleString()} cases`)
    .join('; ');

  return [
    `Records analyzed: ${data.length}.`,
    `Total hold units: ${Number(metrics.totalHoldUnits).toLocaleString()}.`,
    `Total items reworked: ${Number(metrics.totalItemsReworked).toLocaleString()}.`,
    `Rework percent: ${Number(metrics.reworkPercent).toFixed(1)}%.`,
    `Scrap percent: ${Number(metrics.percentScrapped).toFixed(1)}%.`,
    `Top root cause: ${metrics.topRootCause}.`,
    `Top location: ${metrics.topLocation}.`,
    `Most common disposition: ${metrics.topDisposition}.`,
    `Top cause breakdown: ${topCauses || 'N/A'}.`
  ].join(' ');
}

function renderAIInsights(resultElement, insights, keyPhrases) {
  const safeInsights = (Array.isArray(insights) ? insights : []).slice(0, 3);
  if (safeInsights.length === 0) {
    resultElement.textContent = 'No actionable insights were returned.';
    return;
  }

  const insightsList = safeInsights
    .map((line) => `<li>${String(line)}</li>`)
    .join('');

  const phraseText = Array.isArray(keyPhrases) && keyPhrases.length > 0
    ? `<div class="ai-insights-meta"><strong>Key Themes:</strong> ${keyPhrases.map((k) => String(k)).join(', ')}</div>`
    : '';

  resultElement.innerHTML = `${phraseText}<ol class="ai-insights-list">${insightsList}</ol>`;
}

function buildLocalInsights(data, metrics) {
  const insights = [];
  const safeData = Array.isArray(data) ? data : [];
  const reworkPercent = Number(metrics.reworkPercent || 0);
  const scrapPercent = Number(metrics.percentScrapped || 0);
  const topCause = String(metrics.topRootCause || 'Unknown');
  const topLocation = String(metrics.topLocation || 'Unknown');
  const totalReworked = Number(metrics.totalItemsReworked || 0);
  const totalHold = Number(metrics.totalHoldUnits || 0);

  if (reworkPercent >= 5) {
    insights.push(`Rework is high at ${reworkPercent.toFixed(1)}%. Launch immediate containment at ${topLocation} focused on ${topCause}.`);
  } else if (reworkPercent >= 3) {
    insights.push(`Rework is elevated at ${reworkPercent.toFixed(1)}%. Increase first-pass checks on lines with recurring ${topCause}.`);
  } else {
    insights.push(`Rework is stable at ${reworkPercent.toFixed(1)}%. Maintain controls and watch for drift in ${topCause}.`);
  }

  if (scrapPercent >= 2) {
    insights.push(`Scrap is ${scrapPercent.toFixed(1)}%. Add pre-release quality verification to reduce irreversible losses.`);
  }

  insights.push(`Top driver is ${topCause}. Assign an owner to verify root cause and close corrective actions this week.`);

  if (totalHold > 0) {
    insights.push(`Volume context: ${totalReworked.toLocaleString()} reworked of ${totalHold.toLocaleString()} hold units.`);
  }

  const causeCounts = {};
  safeData.forEach((d) => {
    const cause = String(d.rootCause || 'Unknown').trim() || 'Unknown';
    causeCounts[cause] = (causeCounts[cause] || 0) + (parseInt(d.casesReworked) || 0);
  });

  const topCauses = Object.entries(causeCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([cause, count]) => `${cause} (${Number(count).toLocaleString()} cases)`);

  if (topCauses.length > 0) {
    insights.push(`Top contributors: ${topCauses.join(', ')}.`);
  }

  return insights.slice(0, 3);
}

async function generateAIInsights() {
  const button = document.getElementById('generate-ai-insights-btn');
  const result = document.getElementById('ai-insights-result');
  if (!button || !result) return;

  if (!Array.isArray(reworkData) || reworkData.length === 0) {
    result.textContent = 'Load CSV data first, then generate insights.';
    return;
  }

  const active = getFilteredData();
  if (active.length === 0) {
    result.textContent = 'No records match your current filters.';
    return;
  }

  const metrics = buildAIContext(active);
  const summaryText = buildAISummaryText(active, metrics);

  button.disabled = true;
  const originalLabel = button.textContent;
  button.textContent = 'Generating...';
  result.textContent = 'Generating AI insights...';

  try {
    const response = await fetch(AI_ENDPOINT, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        summaryText,
        metrics
      })
    });

    if (!response.ok) {
      const errorPayload = await response.json().catch(() => ({}));
      throw new Error(errorPayload.error || `Request failed with status ${response.status}`);
    }

    const payload = await response.json();
    renderAIInsights(result, payload.insights, payload.keyPhrases);
    setAIHealthStatus('online');
  } catch (error) {
    console.error('AI insights request failed:', error);
    setAIHealthStatus('offline');
    const localInsights = buildLocalInsights(active, metrics);
    const localInsightsList = localInsights.map((line) => `<li>${String(line)}</li>`).join('');
    result.innerHTML = `<div class="ai-insights-meta"><strong>Running in Local Mode:</strong> Backend is unavailable, so these recommendations are generated directly in the dashboard.</div><ol class="ai-insights-list">${localInsightsList}</ol>`;
  } finally {
    button.disabled = false;
    button.textContent = originalLabel;
  }
}

function renderTrendChart(data) {
  const dailyData = {};
  data.forEach((d) => {
    const dateKey = d.holdDate || d.productionDate;
    dailyData[dateKey] = (dailyData[dateKey] || 0) + (parseInt(d.casesReworked) || 0);
  });

  const dates = Object.keys(dailyData).sort(
    (a, b) => (parseFlexibleDate(a) || 0) - (parseFlexibleDate(b) || 0)
  );
  const counts = dates.map((d) => dailyData[d]);

  if (trendChart) trendChart.destroy();

  const ctx = document.getElementById('trendChart').getContext('2d');
  trendChart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: dates,
      datasets: [
        {
          label: 'Cases Reworked',
          data: counts,
          borderColor: '#1e40af',
          backgroundColor: 'rgba(30, 64, 175, 0.05)',
          borderWidth: 2,
          fill: true,
          tension: 0.4,
          pointBackgroundColor: '#1e40af',
          pointRadius: 4,
          pointHoverRadius: 6
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: { legend: { display: false } },
      scales: {
        y: { beginAtZero: true, title: { display: true, text: 'Cases' } }
      }
    }
  });
}

function renderCauseChart(data) {
  const causeCounts = {};
  data.forEach((d) => {
    const cause = d.rootCause || 'Unknown';
    causeCounts[cause] = (causeCounts[cause] || 0) + (parseInt(d.casesReworked) || 0);
  });

  const causes = Object.keys(causeCounts);
  const counts = Object.values(causeCounts);
  const colors = ['#1e40af','#ea580c','#059669','#7c3aed','#0369a1','#b91c1c','#d97706','#16a34a','#065f46','#2563eb'];

  if (causeChart) causeChart.destroy();

  const ctx = document.getElementById('causeChart').getContext('2d');
  causeChart = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: causes,
      datasets: [
        {
          data: counts,
          backgroundColor: colors.slice(0, causes.length),
          borderColor: '#fff',
          borderWidth: 2
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: { position: 'bottom', labels: { padding: 15, font: { size: 12 } } }
      }
    }
  });
}

// ---------- Table ----------
function renderDefectsTable(data) {
  const itemCounts = {};
  let totalReworked = 0;

  data.forEach((d) => {
    const key = d.description || d.itemType || 'Unknown';
    const count = parseInt(d.casesReworked) || 0;
    itemCounts[key] = (itemCounts[key] || 0) + count;
    totalReworked += count;
  });

  const items = Object.entries(itemCounts)
    .map(([name, count]) => ({
      name,
      count,
      percent: totalReworked > 0 ? ((count / totalReworked) * 100).toFixed(1) : 0
    }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 10);

  const tbody = document.getElementById('defectsBody');
  if (!tbody) {
    // No separate defects table exists, skip rendering
    return;
  }
  tbody.innerHTML =
    items.length > 0
      ? items
          .map(
            (item) => `
      <tr>
        <td>${item.name}</td>
        <td>${Number(item.count).toLocaleString()}</td>
        <td>${item.percent}%</td>
      </tr>`
          )
          .join('')
      : '<tr><td colspan="3" style="text-align: center; color: #999;">No data loaded</td></tr>';
}
