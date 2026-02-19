// Rework Tracking Dashboard
const STORAGE_KEY = 'rework-dashboard-data-v13';
let reworkData = [];
let trendChart = null;
let causeChart = null;

// Track sort state for defect drivers table
let defectDriversSortMode = 'cases'; // Options: 'cases', 'cost', 'lag'

// Initialize dashboard on page load
document.addEventListener('DOMContentLoaded', () => {
  loadStoredData();
  hookUpDateEvents();
  setupDefectDriversSorting();
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
      const csv = e.target.result;
      if (window && window.console) {
        console.log('DEBUG: CSV file contents (first 500 chars):', csv.slice(0, 500));
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
        showMessage(`✓ Successfully loaded ${data.length} records from ${file.name}`, 'success');
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

  reader.readAsText(file);
}

// ---------- CSV Parsing ----------
function parseCSV(csv) {
  // Use Papa Parse to handle complex CSV with quoted fields
  const result = Papa.parse(csv, {
    header: false,
    skipEmptyLines: true
  });

  // Print normalized header array for debugging
  if (window && window.console && result && result.data && result.data[0]) {
    const normalizedHeader = result.data[0].map(h => h.replace(/\s+/g, '').toLowerCase());
    console.log('DEBUG: Normalized header array:', normalizedHeader);
  }

  if (!result.data || result.data.length < 2) return [];

  const header = result.data[0].map(h => String(h).toLowerCase().trim());
  if (window && window.console) {
    console.log('Parsed CSV header:', header);
  }
  const data = [];

  // Resolve column indexes once
  const getColIndex = (keywords) => {
    return header.findIndex(h => keywords.some(k => h.includes(k)));
  };

  const prodDateCol = getColIndex(['production date', 'product date', 'date produced', 'prod date']);
  const holdDateCol = getColIndex(['hold date', 'date placed', 'date held', 'hold']);
  const dateCol = prodDateCol !== -1 ? prodDateCol : (holdDateCol !== -1 ? holdDateCol : getColIndex(['date', 'day']));

  const descCol = getColIndex(['description', 'item description', 'desc', 'product', 'item']);
  const typeCol = getColIndex(['type', 'item type', 'category']);
  const plantNameCol = getColIndex(['plantname']);
  const workCenterCol = getColIndex(['workcentertext', 'work center text']);
  const siteCol = getColIndex(['site']);
  // Robustly find the disposition column (case-insensitive, trimmed, supports 'disposition', 'status', etc.)
  const dispositionCol = header.findIndex(h => {
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

  for (let i = 1; i < result.data.length; i++) {
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
      costScrap: costScrap
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
  calculateMetrics(active);
  renderCharts(active);
  updateDataInfo();
}

function renderEmptyState() {
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
