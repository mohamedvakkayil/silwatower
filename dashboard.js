// Dashboard JavaScript

// Format currency - Always AED with 2 decimal places
function formatCurrency(amount) {
    if (amount === null || amount === undefined || isNaN(amount)) {
        return 'AED 0.00';
    }
    const numAmount = parseFloat(amount);
    if (isNaN(numAmount)) {
        return 'AED 0.00';
    }
    // Format as AED with exactly 2 decimal places
    return 'AED ' + numAmount.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// Format number
function formatNumber(num) {
    return new Intl.NumberFormat('en-AE').format(num);
}

// Load and display dashboard data
async function loadDashboardData() {
    try {
        // Try to fetch from API endpoint first (when running on server)
        let response;
        try {
            response = await fetch('/api/dashboard-data');
            if (!response.ok) {
                throw new Error('API not available');
            }
            
            const apiData = await response.json();
            
            if (apiData.error) {
                throw new Error(apiData.error);
            }
            
            // Transform API data to match expected format
            const data = {
                all_boards: apiData.allBoards || [],
                total_estimate: apiData.allBoardsTotal || 0,
                total_load: apiData.allBoardsTotalLoad || 0,
                total_items: apiData.allBoardsTotalItems || 0,
                count: apiData.allBoardsCount || 0,
                mdb_boards: apiData.mdbBoards || [],
                mdb_total_estimate: apiData.totalEstimate || 0,
                mdb_total_load: apiData.totalLoad || 0,
                mdb_total_items: apiData.totalItems || 0,
                last_updated: apiData.lastUpdated
            };
            
            processDashboardData(data);
            updateMDBSummaryCards(data);
            
            // Show last updated time if available
            if (data.last_updated) {
                showLastUpdated(data.last_updated);
            }
            
            return;
        } catch (apiError) {
            console.warn('API fetch failed, trying JSON file:', apiError);
            
            // Fallback: Try to fetch from JSON file
            response = await fetch('all_boards_data.json');
            if (!response.ok) {
                throw new Error('Failed to load data');
            }
            
            const data = await response.json();
            processDashboardData(data);
            return;
        }
        
    } catch (error) {
        console.warn('Fetch failed, trying embedded data:', error);
        
        // Fallback to embedded data (works with file:// protocol)
        if (window.dashboardData && window.dashboardData.allBoards) {
            const data = {
                all_boards: window.dashboardData.allBoards,
                total_estimate: window.dashboardData.allBoardsTotal || 0,
                total_load: window.dashboardData.allBoardsTotalLoad || 0,
                total_items: window.dashboardData.allBoardsTotalItems || 0,
                count: window.dashboardData.allBoardsCount || 0,
                mdb_boards: window.dashboardData.mdbBoards || []
            };
            processDashboardData(data);
            updateMDBSummaryCards(data);
        } else {
            console.error('No data available');
        }
    }
}

// Update MDB summary cards with main MDB data
function updateMDBSummaryCards(data) {
    const mdbBoards = data.mdb_boards || [];
    
    mdbBoards.forEach(mdb => {
        const mdbName = mdb.name.toUpperCase();
        let mdbId = '';
        
        if (mdbName.includes('MDB1')) {
            mdbId = 'mdb1';
        } else if (mdbName.includes('MDB2')) {
            mdbId = 'mdb2';
        } else if (mdbName.includes('MDB3')) {
            mdbId = 'mdb3';
        } else if (mdbName.includes('MDB.GF.04') || mdbName.includes('MDB4')) {
            mdbId = 'mdb4';
        }
        
        if (mdbId) {
            const estimateEl = document.getElementById(`${mdbId}-estimate`);
            const loadEl = document.getElementById(`${mdbId}-load`);
            const itemsEl = document.getElementById(`${mdbId}-items`);
            
            if (estimateEl) {
                estimateEl.textContent = formatCurrency(mdb.estimate || 0);
            }
            if (loadEl) {
                loadEl.textContent = `${(mdb.load || 0).toFixed(2)} kW`;
            }
            if (itemsEl) {
                itemsEl.textContent = formatNumber(mdb.items || 0);
            }
        }
    });
}

// Show last updated time
function showLastUpdated(timestamp) {
    let lastUpdatedEl = document.getElementById('last-updated');
    if (!lastUpdatedEl) {
        // Create element if it doesn't exist
        const dashboardHeader = document.querySelector('.dashboard-header .container');
        if (dashboardHeader) {
            lastUpdatedEl = document.createElement('div');
            lastUpdatedEl.id = 'last-updated';
            lastUpdatedEl.style.cssText = 'text-align: center; margin-top: 1rem; font-size: 0.875rem; color: var(--text-secondary);';
            dashboardHeader.appendChild(lastUpdatedEl);
        }
    }
    
    if (lastUpdatedEl) {
        const date = new Date(timestamp);
        const formattedDate = date.toLocaleString('en-AE', {
            year: 'numeric',
            month: 'short',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
        lastUpdatedEl.textContent = `Last updated: ${formattedDate}`;
    }
}

// Store boards data globally
let allBoardsData = [];
let boardsByMDB = {
    'MDB1': [],
    'MDB2': [],
    'MDB3': [],
    'MDB4': []
};

// Process and display dashboard data
function processDashboardData(data) {
    allBoardsData = data.all_boards || [];
    
    // Group boards by MDB
    boardsByMDB = {
        'MDB1': [],
        'MDB2': [],
        'MDB3': [],
        'MDB4': []
    };
    
    allBoardsData.forEach(board => {
        const mdb = board.mdb;
        if (mdb && boardsByMDB[mdb]) {
            boardsByMDB[mdb].push(board);
        }
    });
    
    // Display each MDB section
    displayMDBSection('MDB1', boardsByMDB['MDB1']);
    displayMDBSection('MDB2', boardsByMDB['MDB2']);
    displayMDBSection('MDB3', boardsByMDB['MDB3']);
    displayMDBSection('MDB4', boardsByMDB['MDB4']);
    
    // Set up filter handler
    setupFilter();
}

// Display MDB section with its boards
function displayMDBSection(mdbName, boards) {
    // Calculate MDB statistics
    const mdbEstimate = boards.reduce((sum, board) => sum + (board.estimate || 0), 0);
    const mdbLoad = boards.reduce((sum, board) => sum + (board.load || 0), 0);
    const mdbItems = boards.reduce((sum, board) => sum + (board.items || 0), 0);
    const mdbBoardsCount = boards.length;
    
    // Update MDB section statistics
    const mdbId = mdbName.toLowerCase();
    document.getElementById(`${mdbId}-estimate`).textContent = formatCurrency(mdbEstimate);
    document.getElementById(`${mdbId}-load`).textContent = `${mdbLoad.toFixed(2)} kW`;
    document.getElementById(`${mdbId}-items`).textContent = formatNumber(mdbItems);
    document.getElementById(`${mdbId}-boards-count`).textContent = formatNumber(mdbBoardsCount);
    
    // Display boards
    const container = document.getElementById(`${mdbId}-boards`);
    container.innerHTML = '';
    
    if (boards.length === 0) {
        container.innerHTML = '<div class="no-boards-message">No boards found for this MDB</div>';
        return;
    }
    
    // Sort boards by estimate (descending)
    const sortedBoards = [...boards].sort((a, b) => (b.estimate || 0) - (a.estimate || 0));
    
    // Create table
    const table = document.createElement('table');
    table.className = 'mdb-boards-table';
    
    // Table header
    const thead = document.createElement('thead');
    thead.innerHTML = `
        <tr>
            <th>Board Name</th>
            <th>KIND</th>
            <th>Load (kW)</th>
            <th>Items</th>
            <th>Estimate (AED)</th>
        </tr>
    `;
    table.appendChild(thead);
    
    // Table body
    const tbody = document.createElement('tbody');
    sortedBoards.forEach(board => {
        const row = document.createElement('tr');
        const load = board.load ? board.load.toFixed(2) : 'N/A';
        const items = board.items || 0;
        const kind = board.kind || 'N/A';
        const isMainMDB = board.name === mdbName;
        
        row.innerHTML = `
            <td class="board-name-cell">
                ${board.name}
                ${isMainMDB ? '<span class="main-mdb-badge">Main</span>' : ''}
            </td>
            <td><span class="board-kind-badge">${kind}</span></td>
            <td class="board-load-cell">${load}</td>
            <td class="board-items-cell">${items}</td>
            <td class="board-estimate-cell">${formatCurrency(board.estimate)}</td>
        `;
        
        // Make row clickable
        row.style.cursor = 'pointer';
        row.addEventListener('click', () => {
            showBoardDetails(board.name);
        });
        
        tbody.appendChild(row);
    });
    table.appendChild(tbody);
    
    container.appendChild(table);
}

// Setup filter functionality
function setupFilter() {
    const filterSelect = document.getElementById('mdb-filter');
    if (!filterSelect) return;
    
    filterSelect.addEventListener('change', (e) => {
        const selectedValue = e.target.value;
        filterMDBSections(selectedValue);
    });
    
    // Initial filter (show all)
    filterMDBSections('all');
}

// Filter MDB sections based on selection
function filterMDBSections(selectedMDB) {
    const allSections = document.querySelectorAll('.mdb-section');
    
    allSections.forEach(section => {
        if (selectedMDB === 'all') {
            section.style.display = 'block';
        } else {
            const sectionMDB = section.getAttribute('data-mdb');
            if (sectionMDB === selectedMDB) {
                section.style.display = 'block';
            } else {
                section.style.display = 'none';
            }
        }
    });
}

// Show board details modal
async function showBoardDetails(boardName) {
    const modal = document.getElementById('board-detail-modal');
    const loading = document.getElementById('detail-loading');
    const content = document.getElementById('detail-content');
    const error = document.getElementById('detail-error');
    const boardNameEl = document.getElementById('detail-board-name');
    
    // Show modal
    modal.style.display = 'block';
    document.body.style.overflow = 'hidden';
    
    // Add history state for back button support
    history.pushState({ modal: true, boardName: boardName }, '', `#board-${boardName}`);
    
    // Reset content
    boardNameEl.textContent = boardName;
    loading.style.display = 'block';
    content.style.display = 'none';
    error.style.display = 'none';
    
    try {
        // Try to fetch from API endpoint (if running server)
        let response;
        try {
            response = await fetch(`/api/board-details?name=${encodeURIComponent(boardName)}`);
            if (!response.ok) throw new Error('API not available');
        } catch (e) {
            // Fallback: use Python script via fetch (requires server)
            // For file:// protocol, we'll show a message
            throw new Error('Server not available. Please run a local server to view details.');
        }
        
        const data = await response.json();
        
        if (data.error) {
            throw new Error(data.error);
        }
        
        // Display details
        displayBoardDetails(data);
        loading.style.display = 'none';
        content.style.display = 'block';
        
    } catch (err) {
        loading.style.display = 'none';
        error.style.display = 'block';
        error.innerHTML = `
            <p><strong>Error loading board details:</strong></p>
            <p>${err.message}</p>
            <p style="margin-top: 1rem; font-size: 0.875rem; color: var(--text-secondary);">
                To view board details, run a local server:<br>
                <code>python3 -m http.server 8000</code><br>
                Then access: <code>http://localhost:8000/dashboard.html</code>
            </p>
        `;
    }
}

// Display board details
function displayBoardDetails(data) {
    const summaryEl = document.getElementById('detail-summary');
    const itemsEl = document.getElementById('detail-items');
    
    // Display metadata and summary
    let summaryHTML = '<div class="detail-summary-grid">';
    
    // Display metadata (KIND, MDB, SMDB)
    if (data.metadata) {
        if (data.metadata.kind) {
            summaryHTML += `
                <div class="detail-summary-item">
                    <span class="detail-summary-label">KIND:</span>
                    <span class="detail-summary-value">${data.metadata.kind}</span>
                </div>
            `;
        }
        if (data.metadata.mdb) {
            summaryHTML += `
                <div class="detail-summary-item">
                    <span class="detail-summary-label">MDB:</span>
                    <span class="detail-summary-value">${data.metadata.mdb}</span>
                </div>
            `;
        }
        if (data.metadata.smdb) {
            summaryHTML += `
                <div class="detail-summary-item">
                    <span class="detail-summary-label">SMDB:</span>
                    <span class="detail-summary-value">${data.metadata.smdb}</span>
                </div>
            `;
        }
        if (data.metadata.load) {
            summaryHTML += `
                <div class="detail-summary-item">
                    <span class="detail-summary-label">Load:</span>
                    <span class="detail-summary-value">${data.metadata.load.toFixed(2)} kW</span>
                </div>
            `;
        }
    }
    
    // Display summary values
    if (data.summary) {
        if (data.summary.net_total !== undefined && data.summary.net_total !== null && data.summary.net_total !== 0) {
            summaryHTML += `
                <div class="detail-summary-item">
                    <span class="detail-summary-label">Net Total:</span>
                    <span class="detail-summary-value">${formatCurrency(data.summary.net_total)}</span>
                </div>
            `;
        }
        if (data.summary.no_of_units !== undefined && data.summary.no_of_units !== null && data.summary.no_of_units !== 0) {
            const unitsValue = parseFloat(data.summary.no_of_units);
            const formattedUnits = !isNaN(unitsValue) ? unitsValue.toFixed(2) : data.summary.no_of_units;
            summaryHTML += `
                <div class="detail-summary-item">
                    <span class="detail-summary-label">No. of Units:</span>
                    <span class="detail-summary-value">${formattedUnits}</span>
                </div>
            `;
        }
    }
    
    summaryHTML += '</div>';
    summaryEl.innerHTML = summaryHTML;
    
    // Display items table
    if (data.items && data.items.length > 0) {
        // Find AMOUNT column (case-insensitive)
        const amountColumnNames = ['AMOUNT', 'Amount', 'amount', 'AMT', 'Amt', 'amt'];
        let amountColumnName = null;
        
        // Get headers from first item
        const headers = Object.keys(data.items[0]);
        
        // Find the amount column
        for (const header of headers) {
            if (amountColumnNames.includes(header) || header.toUpperCase().includes('AMOUNT')) {
                amountColumnName = header;
                break;
            }
        }
        
        // Filter out rows where AMOUNT is 0 or null
        const filteredItems = data.items.filter(item => {
            if (!amountColumnName) return true; // If no AMOUNT column found, keep all rows
            
            const amountValue = item[amountColumnName];
            // Remove rows where AMOUNT is 0, null, undefined, or empty string
            if (amountValue === null || amountValue === undefined || amountValue === '' || amountValue === 0) {
                return false;
            }
            // Also check if it's a string "0"
            if (typeof amountValue === 'string' && amountValue.trim() === '0') {
                return false;
            }
            return true;
        });
        
        if (filteredItems.length === 0) {
            itemsEl.innerHTML = '<p class="no-items-message">No items found with valid amounts.</p>';
            return;
        }
        
        // Define preferred column order for display
        // Note: 'back' is excluded as it's just a hyperlink, not a data column
        const preferredOrder = ['BRAND', 'ITEM', 'DESCRIPTION', 'PRICE', 'QTY', 'QUANTITY', 'AMOUNT'];
        
        // Filter out excluded columns (hyperlinks, not data)
        const excludedColumns = ['back', 'BACK', 'list', 'LIST'];
        const filteredHeaders = headers.filter(h => !excludedColumns.includes(h.toUpperCase()));
        
        // Sort headers: preferred first, then others
        const orderedHeaders = [];
        const remainingHeaders = [...filteredHeaders];
        
        // Add preferred headers in order
        for (const prefHeader of preferredOrder) {
            const foundIndex = remainingHeaders.findIndex(h => h.toUpperCase() === prefHeader.toUpperCase());
            if (foundIndex !== -1) {
                orderedHeaders.push(remainingHeaders[foundIndex]);
                remainingHeaders.splice(foundIndex, 1);
            }
        }
        
        // Add remaining headers
        orderedHeaders.push(...remainingHeaders);
        
        let tableHTML = '<table class="detail-items-table"><thead><tr>';
        orderedHeaders.forEach(header => {
            tableHTML += `<th>${header}</th>`;
        });
        tableHTML += '</tr></thead><tbody>';
        
        // Find summary section: from first "TOTAL" row to "NET TOTAL" row
        let summaryStartIndex = -1;
        let summaryEndIndex = -1;
        
        // Find where summary section starts (first row with "TOTAL" but not "NET TOTAL")
        filteredItems.forEach((item, index) => {
            const rowValues = Object.values(item);
            const hasTotal = rowValues.some(val => {
                const strVal = String(val || '').toUpperCase();
                return strVal.includes('TOTAL') && !strVal.includes('NET TOTAL');
            });
            if (hasTotal && summaryStartIndex === -1) {
                summaryStartIndex = index;
            }
            
            // Find NET TOTAL row
            const hasNetTotal = rowValues.some(val => {
                const strVal = String(val || '').toUpperCase();
                return strVal.includes('NET TOTAL');
            });
            if (hasNetTotal) {
                summaryEndIndex = index;
            }
        });
        
        filteredItems.forEach((item, index) => {
            // Check if this row is in the summary section (from TOTAL to NET TOTAL)
            const rowValues = Object.values(item);
            const hasTotal = rowValues.some(val => {
                const strVal = String(val || '').toUpperCase();
                return strVal.includes('TOTAL');
            });
            
            // Determine if row is in summary section
            const isInSummaryRange = summaryStartIndex !== -1 && 
                                    summaryEndIndex !== -1 && 
                                    index >= summaryStartIndex && 
                                    index <= summaryEndIndex;
            
            const isSummaryRow = isInSummaryRange || hasTotal;
            const isNetTotalRow = rowValues.some(val => {
                const strVal = String(val || '').toUpperCase();
                return strVal.includes('NET TOTAL');
            });
            
            // Apply appropriate class
            let rowClass = '';
            if (isNetTotalRow) {
                rowClass = 'summary-row net-total-row';
            } else if (isSummaryRow) {
                rowClass = 'summary-row';
            }
            
            tableHTML += `<tr class="${rowClass}">`;
            
            orderedHeaders.forEach(header => {
                const value = item[header];
                const headerUpper = header.toUpperCase();
                
                // Check if this is an AMOUNT column and format as currency
                const isAmountColumn = amountColumnName && 
                                      (header === amountColumnName || 
                                       headerUpper.includes('AMOUNT') ||
                                       headerUpper.includes('AMT'));
                
                // Check if this is a PRICE/RATE column (format to 2 decimals, no AED)
                const isPriceColumn = headerUpper.includes('PRICE') || 
                                     headerUpper.includes('RATE') ||
                                     headerUpper.includes('UNIT PRICE');
                
                // Check if value is 0, null, undefined, or empty
                const isEmpty = value === null || 
                               value === undefined || 
                               value === '' || 
                               value === 0 ||
                               (typeof value === 'string' && value.trim() === '0');
                
                // Format display value
                let displayValue = '';
                if (!isEmpty) {
                    if (isAmountColumn) {
                        // Format as AED currency with 2 decimals
                        const numValue = parseFloat(value);
                        if (!isNaN(numValue)) {
                            displayValue = formatCurrency(numValue);
                        } else {
                            displayValue = value;
                        }
                    } else if (isPriceColumn) {
                        // Format as number with 2 decimals (no AED)
                        const numValue = parseFloat(value);
                        if (!isNaN(numValue)) {
                            displayValue = numValue.toFixed(2);
                        } else {
                            displayValue = value;
                        }
                    } else {
                        displayValue = value;
                    }
                }
                
                let cellClass = '';
                if (isEmpty) {
                    cellClass = 'empty-cell';
                } else if (isAmountColumn) {
                    cellClass = 'amount-cell';
                } else if (isPriceColumn) {
                    cellClass = 'price-cell';
                }
                
                tableHTML += `<td class="${cellClass}">${displayValue}</td>`;
            });
            tableHTML += '</tr>';
        });
        
        tableHTML += '</tbody></table>';
        itemsEl.innerHTML = tableHTML;
    } else {
        itemsEl.innerHTML = '<p class="no-items-message">No items found in this board sheet.</p>';
    }
}

// Close detail modal
function closeDetailModal() {
    const modal = document.getElementById('board-detail-modal');
    if (modal) {
        modal.style.display = 'none';
        document.body.style.overflow = '';
        
        // Clear content
        document.getElementById('detail-content').style.display = 'none';
        document.getElementById('detail-error').style.display = 'none';
        document.getElementById('detail-loading').style.display = 'none';
        
        // Update history to remove modal state
        if (history.state && history.state.modal) {
            history.back();
        } else {
            history.replaceState(null, '', window.location.pathname);
        }
    }
}

// Close modal on Escape key
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
        closeDetailModal();
    }
});

// Close modal when clicking overlay
document.addEventListener('DOMContentLoaded', () => {
    const modal = document.getElementById('board-detail-modal');
    if (modal) {
        const overlay = modal.querySelector('.detail-modal-overlay');
        if (overlay) {
            overlay.addEventListener('click', closeDetailModal);
        }
    }
    
    // Close modal when navigating back (browser back button)
    window.addEventListener('popstate', () => {
        closeDetailModal();
    });
});

// Check if URL has board hash and open modal
function checkHashAndOpenModal() {
    const hash = window.location.hash;
    if (hash && hash.startsWith('#board-')) {
        const boardName = hash.replace('#board-', '');
        if (boardName) {
            // Wait a bit for data to load, then open modal
            setTimeout(() => {
                showBoardDetails(decodeURIComponent(boardName));
            }, 500);
        }
    }
}

// Auto-refresh interval (in milliseconds) - refresh every 30 seconds
let autoRefreshInterval = null;
const AUTO_REFRESH_INTERVAL = 30000; // 30 seconds

// Start auto-refresh
function startAutoRefresh() {
    // Clear any existing interval
    if (autoRefreshInterval) {
        clearInterval(autoRefreshInterval);
    }
    
    // Only start auto-refresh if we're on a server (not file:// protocol)
    if (window.location.protocol === 'http:' || window.location.protocol === 'https:') {
        autoRefreshInterval = setInterval(() => {
            console.log('Auto-refreshing dashboard data...');
            loadDashboardData();
        }, AUTO_REFRESH_INTERVAL);
        
        console.log(`Auto-refresh enabled: updating every ${AUTO_REFRESH_INTERVAL / 1000} seconds`);
    }
}

// Stop auto-refresh
function stopAutoRefresh() {
    if (autoRefreshInterval) {
        clearInterval(autoRefreshInterval);
        autoRefreshInterval = null;
    }
}

// Initialize dashboard on page load
document.addEventListener('DOMContentLoaded', () => {
    loadDashboardData();
    
    // Start auto-refresh
    startAutoRefresh();
    
    // Check for hash on load
    checkHashAndOpenModal();
    
    // Also check when hash changes
    window.addEventListener('hashchange', () => {
        const hash = window.location.hash;
        if (hash && hash.startsWith('#board-')) {
            const boardName = hash.replace('#board-', '');
            if (boardName) {
                showBoardDetails(decodeURIComponent(boardName));
            }
        } else {
            closeDetailModal();
        }
    });
    
    // Stop auto-refresh when page is hidden (to save resources)
    document.addEventListener('visibilitychange', () => {
        if (document.hidden) {
            stopAutoRefresh();
        } else {
            startAutoRefresh();
            loadDashboardData(); // Refresh immediately when page becomes visible
        }
    });
});
