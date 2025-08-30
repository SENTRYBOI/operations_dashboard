/**
 * ADVANCED OPERATIONS DASHBOARD WITH KPI ANALYTICS - MAIN JAVASCRIPT
 * Professional operations tracking with dynamic calculations and comprehensive KPI monitoring
 * Features: Daily/Monthly views, downtime-based percentages, real-time sync, KPI analytics
 */

// ========================================
// GLOBAL VARIABLES & CONFIGURATION
// ========================================

// Core data storage
let operationalData = JSON.parse(localStorage.getItem('operationalData')) || {};
let skillsList = JSON.parse(localStorage.getItem('skillsList')) || [
    'Business Clarity',
    'Revenue',
    'BPT Processing',
    'Stock Aging',
    'Inventory',
    'AMB for Manufacturing',
    'PMI for Manufacturing',
    'BPM Stock Reconciliation'
];

// Enhanced Landscapes with dynamic addition
let landscapesList = JSON.parse(localStorage.getItem('landscapesList')) || [
    'Europe', 'NAM', 'IA', 'SEA', 'Global'
];

// Application state
let currentView = 'monthly';
let currentLandscape = 'Europe';
let currentMonth = new Date().toISOString().substring(0, 7);
let currentDate = new Date().toISOString().split('T')[0];
let editingKey = null;

// KPI tracking variables
let monthlyKPICache = {};
let yearlyKPICache = {};
let weeklyKPICache = {};

// ========================================
// DYNAMIC PERCENTAGE CALCULATION ENGINE
// ========================================

/**
 * Calculate performance percentage based on downtime hours
 * @param {number} downtimeHours - Hours of downtime (0-24)
 * @returns {number} Performance percentage (0-100)
 */
function calculatePercentageFromDowntime(downtimeHours) {
    const downtime = parseFloat(downtimeHours) || 0;
    
    if (downtime <= 1.2) {
        // Green: 95-100% - Linear scale from 100% to 95%
        return Math.max(95, Math.round(100 - (downtime * 4.17)));
    } else if (downtime <= 13.2) {
        // Amber: 45-95% - Scale from 95% to 45%
        const range = 13.2 - 1.2; // 12 hours range
        const percentage = 95 - ((downtime - 1.2) / range) * 50;
        return Math.round(percentage);
    } else {
        // Red: 0-45% - Scale based on remaining hours
        const maxDowntime = 24;
        const remaining = Math.max(0, maxDowntime - downtime);
        return Math.round((remaining / (maxDowntime - 13.2)) * 45);
    }
}

/**
 * Determine status color from percentage
 * @param {number} percentage - Performance percentage
 * @returns {string} Status color ('green', 'amber', 'red')
 */
function getStatusFromPercentage(percentage) {
    if (percentage >= 95) return 'green';
    if (percentage >= 45) return 'amber';
    return 'red';
}

/**
 * Update modal percentage field when downtime changes
 */
function updateStatusFromDowntime() {
    const downtime = parseFloat(document.getElementById('modalDowntime').value) || 0;
    const percentage = calculatePercentageFromDowntime(downtime);
    const status = getStatusFromPercentage(percentage);
    
    document.getElementById('modalPercentage').value = `${percentage}% (${status.toUpperCase()})`;
}

// ========================================
// KPI CALCULATION FUNCTIONS
// ========================================

/**
 * Calculate monthly KPIs for current landscape and month
 * @returns {Object} Monthly KPI data
 */
function calculateMonthlyKPIs() {
    const [year, month] = currentMonth.split('-').map(Number);
    const firstDay = new Date(year, month - 1, 1);
    const lastDay = new Date(year, month, 0);
    
    let totalEntries = 0;
    let totalPercentage = 0;
    let greenCount = 0;
    let amberCount = 0;
    let redCount = 0;
    let previousMonthAvg = 0;
    
    // Calculate current month stats
    for (let date = new Date(firstDay); date <= lastDay; date.setDate(date.getDate() + 1)) {
        const dateStr = date.toISOString().split('T')[0];
        
        skillsList.forEach(skill => {
            const entryKey = `${currentLandscape}_${skill}_${dateStr}`;
            const entry = operationalData[entryKey];
            
            if (entry && entry.percentage !== undefined) {
                totalEntries++;
                totalPercentage += entry.percentage;
                
                if (entry.status === 'green') greenCount++;
                else if (entry.status === 'amber') amberCount++;
                else if (entry.status === 'red') redCount++;
            }
        });
    }
    
    // Calculate previous month average for trend
    const prevMonth = month - 1 === 0 ? 12 : month - 1;
    const prevYear = month - 1 === 0 ? year - 1 : year;
    const prevFirstDay = new Date(prevYear, prevMonth - 1, 1);
    const prevLastDay = new Date(prevYear, prevMonth, 0);
    
    let prevTotalEntries = 0;
    let prevTotalPercentage = 0;
    
    for (let date = new Date(prevFirstDay); date <= prevLastDay; date.setDate(date.getDate() + 1)) {
        const dateStr = date.toISOString().split('T')[0];
        
        skillsList.forEach(skill => {
            const entryKey = `${currentLandscape}_${skill}_${dateStr}`;
            const entry = operationalData[entryKey];
            
            if (entry && entry.percentage !== undefined) {
                prevTotalEntries++;
                prevTotalPercentage += entry.percentage;
            }
        });
    }
    
    previousMonthAvg = prevTotalEntries > 0 ? Math.round(prevTotalPercentage / prevTotalEntries) : 0;
    
    const overallAvg = totalEntries > 0 ? Math.round(totalPercentage / totalEntries) : 0;
    const greenPercentage = totalEntries > 0 ? Math.round((greenCount / totalEntries) * 100) : 0;
    const amberPercentage = totalEntries > 0 ? Math.round((amberCount / totalEntries) * 100) : 0;
    const redPercentage = totalEntries > 0 ? Math.round((redCount / totalEntries) * 100) : 0;
    
    let trendIndicator = '📊';
    if (overallAvg > previousMonthAvg) trendIndicator = '📈';
    else if (overallAvg < previousMonthAvg) trendIndicator = '📉';
    
    return {
        overallAvg,
        greenPercentage,
        amberPercentage,
        redPercentage,
        trendIndicator,
        totalEntries,
        previousMonthAvg
    };
}

/**
 * Calculate yearly KPIs for current landscape and year
 * @returns {Object} Yearly KPI data
 */
function calculateYearlyKPIs() {
    const currentYear = new Date(currentMonth).getFullYear();
    const startOfYear = new Date(currentYear, 0, 1);
    const endOfYear = new Date(currentYear, 11, 31);
    
    let totalEntries = 0;
    let totalPercentage = 0;
    let greenCount = 0;
    let amberCount = 0;
    let redCount = 0;
    let previousYearAvg = 0;
    
    // Calculate current year stats
    for (let date = new Date(startOfYear); date <= endOfYear; date.setDate(date.getDate() + 1)) {
        const dateStr = date.toISOString().split('T')[0];
        
        skillsList.forEach(skill => {
            const entryKey = `${currentLandscape}_${skill}_${dateStr}`;
            const entry = operationalData[entryKey];
            
            if (entry && entry.percentage !== undefined) {
                totalEntries++;
                totalPercentage += entry.percentage;
                
                if (entry.status === 'green') greenCount++;
                else if (entry.status === 'amber') amberCount++;
                else if (entry.status === 'red') redCount++;
            }
        });
    }
    
    // Calculate previous year average for trend
    const prevStartOfYear = new Date(currentYear - 1, 0, 1);
    const prevEndOfYear = new Date(currentYear - 1, 11, 31);
    
    let prevTotalEntries = 0;
    let prevTotalPercentage = 0;
    
    for (let date = new Date(prevStartOfYear); date <= prevEndOfYear; date.setDate(date.getDate() + 1)) {
        const dateStr = date.toISOString().split('T')[0];
        
        skillsList.forEach(skill => {
            const entryKey = `${currentLandscape}_${skill}_${dateStr}`;
            const entry = operationalData[entryKey];
            
            if (entry && entry.percentage !== undefined) {
                prevTotalEntries++;
                prevTotalPercentage += entry.percentage;
            }
        });
    }
    
    previousYearAvg = prevTotalEntries > 0 ? Math.round(prevTotalPercentage / prevTotalEntries) : 0;
    
    const overallAvg = totalEntries > 0 ? Math.round(totalPercentage / totalEntries) : 0;
    const greenPercentage = totalEntries > 0 ? Math.round((greenCount / totalEntries) * 100) : 0;
    const amberPercentage = totalEntries > 0 ? Math.round((amberCount / totalEntries) * 100) : 0;
    const redPercentage = totalEntries > 0 ? Math.round((redCount / totalEntries) * 100) : 0;
    
    let trendIndicator = '📊';
    if (overallAvg > previousYearAvg) trendIndicator = '📈';
    else if (overallAvg < previousYearAvg) trendIndicator = '📉';
    
    return {
        overallAvg,
        greenPercentage,
        amberPercentage,
        redPercentage,
        trendIndicator,
        totalEntries,
        previousYearAvg
    };
}

/**
 * Calculate weekly KPIs for each week in current month
 * @param {Array} weeks - Week data array
 * @returns {Array} Weekly KPI data
 */
function calculateWeeklyKPIs(weeks) {
    const weeklyKPIs = [];
    
    weeks.forEach((week, index) => {
        let totalEntries = 0;
        let totalPercentage = 0;
        
        week.dates.forEach(dateStr => {
            skillsList.forEach(skill => {
                const entryKey = `${currentLandscape}_${skill}_${dateStr}`;
                const entry = operationalData[entryKey];
                
                if (entry && entry.percentage !== undefined) {
                    totalEntries++;
                    totalPercentage += entry.percentage;
                }
            });
        });
        
        const avgPercentage = totalEntries > 0 ? Math.round(totalPercentage / totalEntries) : 0;
        const status = getStatusFromPercentage(avgPercentage);
        
        weeklyKPIs.push({
            week: week.label,
            percentage: avgPercentage,
            status: status,
            totalEntries: totalEntries
        });
    });
    
    return weeklyKPIs;
}

/**
 * Update KPI displays in the dashboard
 */
function updateKPIDisplays() {
    const monthlyKPIs = calculateMonthlyKPIs();
    const yearlyKPIs = calculateYearlyKPIs();
    
    // Update monthly KPIs
    document.getElementById('currentMonthDisplay').textContent = new Date(currentMonth).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    document.getElementById('monthlyOverallAvailability').textContent = `${monthlyKPIs.overallAvg}%`;
    document.getElementById('monthlyGreenPercentage').textContent = `${monthlyKPIs.greenPercentage}%`;
    document.getElementById('monthlyAmberPercentage').textContent = `${monthlyKPIs.amberPercentage}%`;
    document.getElementById('monthlyRedPercentage').textContent = `${monthlyKPIs.redPercentage}%`;
    document.getElementById('monthlyTrendIndicator').textContent = monthlyKPIs.trendIndicator;
    
    // Update yearly KPIs
    document.getElementById('currentYearDisplay').textContent = new Date(currentMonth).getFullYear();
    document.getElementById('yearlyOverallAvailability').textContent = `${yearlyKPIs.overallAvg}%`;
    document.getElementById('yearlyGreenPercentage').textContent = `${yearlyKPIs.greenPercentage}%`;
    document.getElementById('yearlyAmberPercentage').textContent = `${yearlyKPIs.amberPercentage}%`;
    document.getElementById('yearlyRedPercentage').textContent = `${yearlyKPIs.redPercentage}%`;
    document.getElementById('yearlyTrendIndicator').textContent = yearlyKPIs.trendIndicator;
    
    // Cache the results
    monthlyKPICache = monthlyKPIs;
    yearlyKPICache = yearlyKPIs;
}

/**
 * Generate and display weekly KPIs above the monthly table
 * @param {Array} weeks - Week data array
 */
function generateWeeklyKPIs(weeks) {
    const weeklyKPIs = calculateWeeklyKPIs(weeks);
    let kpiHTML = '';
    
    weeklyKPIs.forEach(weekKPI => {
        const statusClass = weekKPI.status;
        
        kpiHTML += `
            <div class="weekly-kpi-item">
                <div class="weekly-kpi-value" style="color: ${getStatusColor(statusClass)}">
                    ${weekKPI.percentage}%
                </div>
                <div class="weekly-kpi-label">
                    ${weekKPI.week} Availability
                </div>
            </div>
        `;
    });
    
    document.getElementById('weeklyKPIGrid').innerHTML = kpiHTML;
    weeklyKPICache = weeklyKPIs;
}

/**
 * Get status color for display
 * @param {string} status - Status string
 * @returns {string} Color code
 */
function getStatusColor(status) {
    switch(status) {
        case 'green': return '#28a745';
        case 'amber': return '#ffc107';
        case 'red': return '#dc3545';
        default: return '#666';
    }
}

// ========================================
// INITIALIZATION & SETUP
// ========================================

/**
 * Initialize the application when DOM is loaded
 */
document.addEventListener('DOMContentLoaded', function() {
    console.log('🚀 Advanced Operations Dashboard with KPI Analytics - Initializing...');
    
    initializeLandscapes();
    document.getElementById('monthSelect').value = currentMonth;
    document.getElementById('dateSelect').value = currentDate;
    
    updateSkillList();
    updateKPIDisplays();
    loadCurrentView();
    
    // Set up auto-save interval
    setInterval(autoSave, 30000);
    
    console.log('✅ Dashboard with KPI analytics initialized successfully');
});

// ========================================
// LANDSCAPE MANAGEMENT
// ========================================

/**
 * Initialize landscape dropdowns
 */
function initializeLandscapes() {
    const landscapeSelect = document.getElementById('landscapeSelect');
    const historyFilter = document.getElementById('historyFilter');
    
    // Clear existing options
    landscapeSelect.innerHTML = '';
    if (historyFilter) {
        historyFilter.innerHTML = '<option value="">All Landscapes</option>';
    }
    
    // Populate landscapes
    landscapesList.forEach(landscape => {
        const option = document.createElement('option');
        option.value = landscape;
        option.textContent = landscape;
        landscapeSelect.appendChild(option);
        
        if (historyFilter) {
            const filterOption = document.createElement('option');
            filterOption.value = landscape;
            filterOption.textContent = landscape;
            historyFilter.appendChild(filterOption);
        }
    });
    
    // Set default
    landscapeSelect.value = currentLandscape;
}

/**
 * Show modal to add new landscape
 */
function addNewLandscape() {
    document.getElementById('addLandscapeModal').style.display = 'block';
    document.getElementById('newLandscapeName').focus();
}

/**
 * Close add landscape modal
 */
function closeAddLandscapeModal() {
    document.getElementById('addLandscapeModal').style.display = 'none';
    document.getElementById('newLandscapeName').value = '';
}

/**
 * Save new landscape to system
 */
function saveNewLandscape() {
    const newLandscape = document.getElementById('newLandscapeName').value.trim();
    
    if (!newLandscape) {
        alert('❌ Please enter a landscape name!');
        return;
    }
    
    if (landscapesList.includes(newLandscape)) {
        alert('❌ Landscape already exists!');
        return;
    }
    
    // Add to list and save
    landscapesList.push(newLandscape);
    localStorage.setItem('landscapesList', JSON.stringify(landscapesList));
    
    // Update dropdowns
    initializeLandscapes();
    
    // Select new landscape
    document.getElementById('landscapeSelect').value = newLandscape;
    currentLandscape = newLandscape;
    
    closeAddLandscapeModal();
    loadCurrentView();
    alert(`✅ "${newLandscape}" landscape added successfully!`);
}

// ========================================
// VIEW MANAGEMENT & NAVIGATION
// ========================================

/**
 * Switch between monthly and daily views
 * @param {string} view - 'monthly' or 'daily'
 */
function switchView(view) {
    currentView = view;
    
    // Update toggle buttons
    document.querySelectorAll('.toggle-btn').forEach(btn => {
        btn.classList.remove('active');
        btn.setAttribute('aria-selected', 'false');
    });
    event.target.classList.add('active');
    event.target.setAttribute('aria-selected', 'true');
    
    // Show/hide controls and views
    if (view === 'daily') {
        document.getElementById('monthControl').style.display = 'none';
        document.getElementById('dateControl').style.display = 'flex';
        document.getElementById('monthlyView').style.display = 'none';
        document.getElementById('dailyView').style.display = 'block';
    } else {
        document.getElementById('monthControl').style.display = 'flex';
        document.getElementById('dateControl').style.display = 'none';
        document.getElementById('monthlyView').style.display = 'block';
        document.getElementById('dailyView').style.display = 'none';
    }
    
    loadCurrentView();
}

/**
 * Load and render current view
 */
function loadCurrentView() {
    currentLandscape = document.getElementById('landscapeSelect').value;
    
    if (currentView === 'monthly') {
        currentMonth = document.getElementById('monthSelect').value;
        generateMonthlyView();
        updateKPIDisplays();
    } else {
        currentDate = document.getElementById('dateSelect').value;
        generateDailyView();
    }
    
    updateSummaryCards();
}

// ========================================
// WEEKLY AGGREGATION & CALCULATIONS
// ========================================

/**
 * Calculate weekly status from daily entries
 * @param {string} skill - Skill name
 * @param {Array} weekDates - Array of date strings for the week
 * @returns {Object} Weekly status and percentage
 */
function calculateWeeklyStatus(skill, weekDates) {
    let totalPercentage = 0;
    let validEntries = 0;
    
    weekDates.forEach(date => {
        const entryKey = `${currentLandscape}_${skill}_${date}`;
        const entry = operationalData[entryKey];
        if (entry) {
            totalPercentage += entry.percentage || 0;
            validEntries++;
        }
    });
    
    if (validEntries === 0) return { percentage: 0, status: 'none' };
    
    const avgPercentage = Math.round(totalPercentage / validEntries);
    const status = getStatusFromPercentage(avgPercentage);
    
    return { percentage: avgPercentage, status: status };
}

// ========================================
// MONTHLY VIEW GENERATION
// ========================================

/**
 * Generate and render monthly view
 */
function generateMonthlyView() {
    const weeks = generateWeeksForMonth();
    generateWeeklyKPIs(weeks);
    createMonthlyHeaders(weeks);
    createMonthlyBody(weeks);
}

/**
 * Generate week data for current month
 * @returns {Array} Array of week objects
 */
function generateWeeksForMonth() {
    const [year, month] = currentMonth.split('-').map(Number);
    const firstDay = new Date(year, month - 1, 1);
    const lastDay = new Date(year, month, 0);
    
    let weeks = [];
    let currentDate = new Date(firstDay);
    let weekNum = 1;

    while (currentDate <= lastDay) {
        const weekStart = new Date(currentDate);
        const weekEnd = new Date(currentDate);
        weekEnd.setDate(currentDate.getDate() + 6);
        
        // Get all dates in this week
        const weekDates = [];
        for (let d = new Date(weekStart); d <= weekEnd && d <= lastDay; d.setDate(d.getDate() + 1)) {
            weekDates.push(d.toISOString().split('T')[0]);
        }
        
        weeks.push({
            number: weekNum,
            label: `WK ${weekNum}`,
            start: weekStart.getDate(),
            end: Math.min(weekEnd.getDate(), lastDay.getDate()),
            dates: weekDates
        });
        
        currentDate.setDate(currentDate.getDate() + 7);
        weekNum++;
    }
    
    return weeks;
}

/**
 * Create monthly view headers
 * @param {Array} weeks - Week data array
 */
function createMonthlyHeaders(weeks) {
    let headerHTML = '<tr><th class="skill-header">Skills / Operational Areas</th>';
    
    weeks.forEach(week => {
        headerHTML += `<th>${week.label}<br><small>${week.start}-${week.end}</small></th>`;
    });
    
    headerHTML += '</tr>';
    document.getElementById('monthlyHeaders').innerHTML = headerHTML;
}

/**
 * Create monthly view body with weekly aggregation
 * @param {Array} weeks - Week data array
 */
function createMonthlyBody(weeks) {
    let bodyHTML = '';
    
    skillsList.forEach(skill => {
        bodyHTML += `<tr><td class="skill-cell">${skill}</td>`;
        
        weeks.forEach(week => {
            const weeklyData = calculateWeeklyStatus(skill, week.dates);
            const statusClass = weeklyData.status;
            const percentage = weeklyData.percentage;
            
            // Create daily indicators
            let dailyIndicators = '<div class="daily-indicators">';
            week.dates.forEach((date, index) => {
                const entryKey = `${currentLandscape}_${skill}_${date}`;
                const entry = operationalData[entryKey];
                const dayStatus = entry ? getStatusFromPercentage(entry.percentage) : 'none';
                const dayName = ['S','M','T','W','T','F','S'][new Date(date).getDay()];
                
                dailyIndicators += `
                    <div class="daily-indicator ${dayStatus}" 
                         title="${dayName}: ${entry ? entry.percentage + '%' : 'No data'}" 
                         onclick="goToDailyView('${date}')"></div>
                `;
            });
            dailyIndicators += '</div>';
            
            bodyHTML += `
                <td class="status-cell" onclick="editWeeklyEntry('${skill}', ${week.number}, '${JSON.stringify(week.dates).replace(/"/g, '&quot;')}')">
                    <div class="status-dot ${statusClass}" title="Click to view details - Week ${week.number}">
                        ${percentage}%
                    </div>
                    ${dailyIndicators}
                </td>
            `;
        });
        
        bodyHTML += '</tr>';
    });
    
    document.getElementById('monthlyBody').innerHTML = bodyHTML;
}

/**
 * Navigate to daily view for specific date
 * @param {string} date - Date string (YYYY-MM-DD)
 */
function goToDailyView(date) {
    // Switch to daily view and set the date
    document.getElementById('dateSelect').value = date;
    currentDate = date;
    
    // Update view
    currentView = 'daily';
    document.getElementById('monthControl').style.display = 'none';
    document.getElementById('dateControl').style.display = 'flex';
    document.getElementById('monthlyView').style.display = 'none';
    document.getElementById('dailyView').style.display = 'block';
    
    // Update toggle buttons
    document.querySelectorAll('.toggle-btn').forEach(btn => {
        btn.classList.remove('active');
        btn.setAttribute('aria-selected', 'false');
    });
    document.querySelector('.toggle-btn:last-child').classList.add('active');
    document.querySelector('.toggle-btn:last-child').setAttribute('aria-selected', 'true');
    
    generateDailyView();
}

/**
 * Handle weekly entry click in monthly view
 * @param {string} skill - Skill name
 * @param {number} weekNumber - Week number
 * @param {string} weekDatesStr - JSON string of week dates
 */
function editWeeklyEntry(skill, weekNumber, weekDatesStr) {
    const weekDates = JSON.parse(weekDatesStr.replace(/&quot;/g, '"'));
    
    // Show options: View details or go to daily view
    const choice = confirm(`📊 Weekly Summary for ${skill} - Week ${weekNumber}\n\nClick OK to go to Daily View for editing\nClick Cancel to stay in Monthly View`);
    
    if (choice) {
        // Go to daily view with the first date of the week
        goToDailyView(weekDates[0]);
    }
}

// ========================================
// DAILY VIEW GENERATION
// ========================================

/**
 * Generate and render daily view
 */
function generateDailyView() {
    const date = new Date(currentDate);
    const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    
    let gridHTML = '<div class="daily-header">Skills</div>';
    
    // Create day headers
    for (let i = 0; i < 7; i++) {
        const dayDate = new Date(date);
        dayDate.setDate(date.getDate() - date.getDay() + i);
        gridHTML += `<div class="daily-header">${dayNames[i]}<br><small>${dayDate.getDate()}/${dayDate.getMonth() + 1}</small></div>`;
    }
    
    // Create skill rows
    skillsList.forEach(skill => {
        gridHTML += `<div class="daily-skill">${skill}</div>`;
        
        for (let i = 0; i < 7; i++) {
            const dayDate = new Date(date);
            dayDate.setDate(date.getDate() - date.getDay() + i);
            const dateStr = dayDate.toISOString().split('T')[0];
            const entryKey = `${currentLandscape}_${skill}_${dateStr}`;
            const entry = operationalData[entryKey];
            const statusClass = entry ? getStatusFromPercentage(entry.percentage) : 'none';
            const percentage = entry ? entry.percentage : 0;
            const downtime = entry ? entry.downtime || 0 : 0;
            
            gridHTML += `
                <div class="daily-cell" onclick="editEntry('${entryKey}', '${skill}', '${dateStr}')">
                    <div class="status-dot ${statusClass}" title="Click to edit ${skill} - ${dateStr}">
                        ${entry ? '✓' : ''}
                    </div>
                    <div class="percentage-value">${percentage}%</div>
                    <div class="downtime-display">${downtime}h down</div>
                </div>
            `;
        }
    });
    
    document.getElementById('dailyGrid').innerHTML = gridHTML;
    updateDailyStats();
}

/**
 * Update daily statistics display
 */
function updateDailyStats() {
    const date = new Date(currentDate);
    let todayTotal = 0, todaySum = 0;
    let weekTotal = 0, weekSum = 0;

    // Today's stats
    const todayKey = currentDate;
    skillsList.forEach(skill => {
        const entryKey = `${currentLandscape}_${skill}_${todayKey}`;
        const entry = operationalData[entryKey];
        todayTotal++;
        if (entry) {
            todaySum += entry.percentage || 0;
        }
    });

    // Week's stats
    for (let i = 0; i < 7; i++) {
        const dayDate = new Date(date);
        dayDate.setDate(date.getDate() - date.getDay() + i);
        const dateStr = dayDate.toISOString().split('T')[0];
        
        skillsList.forEach(skill => {
            const entryKey = `${currentLandscape}_${skill}_${dateStr}`;
            const entry = operationalData[entryKey];
            weekTotal++;
            if (entry) {
                weekSum += entry.percentage || 0;
            }
        });
    }

    const dailyCompletion = todayTotal > 0 ? Math.round(todaySum / todayTotal) : 0;
    const weeklyAverage = weekTotal > 0 ? Math.round(weekSum / weekTotal) : 0;
    const trend = dailyCompletion >= weeklyAverage ? '📈' : '📉';

    document.getElementById('dailyCompletion').textContent = `${dailyCompletion}%`;
    document.getElementById('weeklyAverage').textContent = `${weeklyAverage}%`;
    document.getElementById('trendIndicator').textContent = trend;
}

// ========================================
// ENTRY MANAGEMENT
// ========================================

/**
 * Open edit modal for an entry
 * @param {string} key - Entry key
 * @param {string} skill - Skill name
 * @param {string} period - Period/date string
 */
function editEntry(key, skill, period) {
    editingKey = key;
    const entry = operationalData[key] || {};
    
    document.getElementById('modalSkill').value = skill;
    document.getElementById('modalPeriod').value = period;
    document.getElementById('modalDate').value = entry.date || period;
    document.getElementById('modalNotes').value = entry.notes || '';
    document.getElementById('modalDowntime').value = entry.downtime || 0;
    document.getElementById('modalIncident').value = entry.incident || '';
    
    // Update percentage display
    updateStatusFromDowntime();
    
    document.getElementById('editModal').style.display = 'block';
    document.getElementById('editModal').setAttribute('aria-hidden', 'false');
}

/**
 * Save entry with calculated percentage
 */
function saveEntry() {
    const downtime = parseFloat(document.getElementById('modalDowntime').value) || 0;
    const percentage = calculatePercentageFromDowntime(downtime);
    const status = getStatusFromPercentage(percentage);
    
    const entry = {
        status: status,
        date: document.getElementById('modalDate').value,
        notes: document.getElementById('modalNotes').value || '',
        percentage: percentage,
        downtime: downtime,
        incident: document.getElementById('modalIncident').value || '',
        landscape: currentLandscape,
        skill: document.getElementById('modalSkill').value,
        period: document.getElementById('modalPeriod').value,
        timestamp: new Date().toISOString()
    };
    
    operationalData[editingKey] = entry;
    localStorage.setItem('operationalData', JSON.stringify(operationalData));
    
    closeModal();
    loadCurrentView(); // Refresh both views and sync them
    updateKPIDisplays(); // Update KPIs after data change
    alert('✅ Entry saved successfully with calculated percentage and status!');
}

/**
 * Delete current entry
 */
function deleteEntry() {
    if (confirm('🗑️ Delete this entry?')) {
        delete operationalData[editingKey];
        localStorage.setItem('operationalData', JSON.stringify(operationalData));
        
        closeModal();
        loadCurrentView();
        updateKPIDisplays(); // Update KPIs after data change
        alert('✅ Entry deleted!');
    }
}

/**
 * Close edit modal
 */
function closeModal() {
    document.getElementById('editModal').style.display = 'none';
    document.getElementById('editModal').setAttribute('aria-hidden', 'true');
    editingKey = null;
}

// ========================================
// SKILL MANAGEMENT
// ========================================

/**
 * Update skill list display
 */
function updateSkillList() {
    let listHTML = '';
    skillsList.forEach((skill, index) => {
        listHTML += `
            <div class="skill-tag" role="listitem">
                ${skill}
                <button class="skill-remove" onclick="removeSkill(${index})" title="Remove skill" aria-label="Remove ${skill}">×</button>
            </div>
        `;
    });
    document.getElementById('skillList').innerHTML = listHTML;
}

/**
 * Show prompt to add new skill
 */
function showAddSkillModal() {
    const skillName = prompt('📝 Enter new skill name:');
    if (skillName && skillName.trim()) {
        addSkill(skillName.trim());
    }
}

/**
 * Add new skill to system
 * @param {string} skillName - Skill name (optional, gets from input if not provided)
 */
function addSkill(skillName = null) {
    if (!skillName) {
        skillName = document.getElementById('newSkillInput').value.trim();
    }
    
    if (skillName && !skillsList.includes(skillName)) {
        skillsList.push(skillName);
        localStorage.setItem('skillsList', JSON.stringify(skillsList));
        document.getElementById('newSkillInput').value = '';
        updateSkillList();
        loadCurrentView();
        updateKPIDisplays(); // Update KPIs after skill change
        alert(`✅ "${skillName}" added successfully!`);
    } else if (skillsList.includes(skillName)) {
        alert('❌ Skill already exists!');
    } else {
        alert('❌ Please enter a skill name!');
    }
}

/**
 * Remove skill from system
 * @param {number} index - Skill index in array
 */
function removeSkill(index) {
    if (confirm(`🗑️ Remove "${skillsList[index]}"?`)) {
        const skillName = skillsList[index];
        skillsList.splice(index, 1);
        
        // Remove related entries
        Object.keys(operationalData).forEach(key => {
            if (key.includes(skillName)) {
                delete operationalData[key];
            }
        });
        
        localStorage.setItem('skillsList', JSON.stringify(skillsList));
        localStorage.setItem('operationalData', JSON.stringify(operationalData));
        
        updateSkillList();
        loadCurrentView();
        updateKPIDisplays(); // Update KPIs after skill removal
        alert('✅ Skill removed!');
    }
}

// ========================================
// SUMMARY CARDS & STATISTICS
// ========================================

/**
 * Update summary cards with current statistics
 */
function updateSummaryCards() {
    let total = 0, green = 0, amber = 0, red = 0;
    
    Object.values(operationalData).forEach(entry => {
        if (entry.landscape === currentLandscape) {
            total++;
            switch(entry.status) {
                case 'green': green++; break;
                case 'amber': amber++; break;
                case 'red': red++; break;
            }
        }
    });
    
    document.getElementById('totalCount').textContent = total;
    document.getElementById('greenCount').textContent = green;
    document.getElementById('amberCount').textContent = amber;
    document.getElementById('redCount').textContent = red;
}

// ========================================
// EXCEL EXPORT FUNCTIONALITY
// ========================================

/**
 * Export monthly view to Excel with KPI data
 */
function exportToExcel() {
    const weeks = generateWeeksForMonth();
    const [year, month] = currentMonth.split('-').map(Number);
    const monthName = new Date(year, month - 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    
    let htmlTable = generateExcelTable(weeks, monthName);
    downloadExcelFile(htmlTable, `Advanced_KPI_Operations_Report_${currentLandscape}_${monthName.replace(' ', '_')}.xls`);
    
    alert('✅ Advanced Excel report with KPI analytics exported successfully!');
}

/**
 * Generate Excel-compatible HTML table with KPI data
 * @param {Array} weeks - Week data
 * @param {string} monthName - Month name string
 * @returns {string} HTML table string
 */
function generateExcelTable(weeks, monthName) {
    const monthlyKPIs = monthlyKPICache || calculateMonthlyKPIs();
    const yearlyKPIs = yearlyKPICache || calculateYearlyKPIs();
    const weeklyKPIs = weeklyKPICache || calculateWeeklyKPIs(weeks);
    
    let htmlTable = `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }
                .header { background-color: #4472C4; color: white; font-weight: bold; text-align: center; padding: 8px; font-size: 11px; }
                .subheader { background-color: #D9E1F2; font-weight: bold; text-align: center; padding: 6px; font-size: 10px; border: 1px solid #B4C6E7; }
                .kpi-header { background-color: #70AD47; color: white; font-weight: bold; text-align: center; padding: 8px; font-size: 11px; }
                .kpi-cell { background-color: #E2EFDA; text-align: center; padding: 6px; font-size: 10px; border: 1px solid #A9D08E; }
                .skill-cell { background-color: #E2EFDA; font-weight: bold; padding: 8px; font-size: 10px; border: 1px solid #A9D08E; text-align: left; }
                .data-cell { text-align: center; padding: 8px; font-size: 12px; border: 1px solid #D0D7DE; font-weight: bold; }
                .green-cell { background-color: #C6EFCE; color: #006100; }
                .amber-cell { background-color: #FFEB9C; color: #9C5700; }
                .red-cell { background-color: #FFC7CE; color: #9C0006; }
                .white-cell { background-color: #FFFFFF; color: #767676; }
                .status-dot { font-size: 16px; }
            </style>
        </head>
        <body>
            <table border="1">
                <tr>
                    <td colspan="${weeks.length + 1}" class="header">
                        Advanced Operations Dashboard with KPI Analytics - ${monthName}
                    </td>
                </tr>
                <tr>
                    <td colspan="${weeks.length + 1}" class="subheader">
                        Landscape: ${currentLandscape} | Report Generated: ${new Date().toLocaleDateString()} | Dynamic Downtime-Based Calculation with KPI Tracking
                    </td>
                </tr>
                <tr>
                    <td colspan="${weeks.length + 1}" class="kpi-header">
                        MONTHLY KPI SUMMARY
                    </td>
                </tr>
                <tr>
                    <td class="kpi-cell">Overall Availability: ${monthlyKPIs.overallAvg}%</td>
                    <td class="kpi-cell">Green Performance: ${monthlyKPIs.greenPercentage}%</td>
                    <td class="kpi-cell">Amber Performance: ${monthlyKPIs.amberPercentage}%</td>
                    <td class="kpi-cell">Red Performance: ${monthlyKPIs.redPercentage}%</td>
                    <td class="kpi-cell">Trend: ${monthlyKPIs.trendIndicator}</td>
    `;
    
    // Add remaining columns if weeks.length > 4
    for (let i = 4; i < weeks.length; i++) {
        htmlTable += '<td class="kpi-cell">-</td>';
    }
    htmlTable += '</tr>';
    
    // Weekly KPIs row
    htmlTable += '<tr><td class="kpi-header">Weekly Availability</td>';
    weeklyKPIs.forEach(weekKPI => {
        htmlTable += `<td class="kpi-cell">${weekKPI.week}: ${weekKPI.percentage}%</td>`;
    });
    htmlTable += '</tr>';
    
    // Main data table
    htmlTable += `
                <tr>
                    <td class="skill-cell" style="width: 250px;">Operational Description</td>
    `;
    
    weeks.forEach(week => {
        htmlTable += `<td class="header" style="width: 80px;">${week.label}<br><small>${week.start}-${week.end}</small></td>`;
    });
    htmlTable += '</tr>';
    
    skillsList.forEach(skill => {
        htmlTable += `<tr><td class="skill-cell">${skill}</td>`;
        
        weeks.forEach(week => {
            const weeklyData = calculateWeeklyStatus(skill, week.dates);
            let cellClass = 'white-cell';
            let statusDot = '';
            let percentage = '';
            
            if (weeklyData.status !== 'none') {
                percentage = `${weeklyData.percentage}.0%`;
                switch(weeklyData.status) {
                    case 'green':
                        cellClass = 'green-cell';
                        statusDot = '●';
                        break;
                    case 'amber':
                        cellClass = 'amber-cell';
                        statusDot = '●';
                        break;
                    case 'red':
                        cellClass = 'red-cell';
                        statusDot = '●';
                        break;
                }
            } else {
                statusDot = '○';
                percentage = '';
            }
            
            htmlTable += `
                <td class="data-cell ${cellClass}">
                    <span class="status-dot">${statusDot}</span><br>
                    <small>${percentage}</small>
                </td>
            `;
        });
        
        htmlTable += '</tr>';
    });
    
    htmlTable += `
                <tr>
                    <td colspan="${weeks.length + 1}" class="subheader">
                        Legend: ● Green (95-100%) | ● Amber (45-95%) | ● Red (0-45%) | ○ No Data - Advanced KPI Analytics with Dynamic Downtime Calculation
                    </td>
                </tr>
            </table>
        </body>
        </html>
    `;
    
    return htmlTable;
}

/**
 * Download Excel file
 * @param {string} content - HTML content
 * @param {string} filename - File name
 */
function downloadExcelFile(content, filename) {
    const blob = new Blob([content], { type: 'application/vnd.ms-excel;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

/**
 * Export history to Excel with KPI analytics
 */
function exportHistoryToExcel() {
    let htmlTable = generateHistoryExcelTable();
    downloadExcelFile(htmlTable, `Operations_History_KPI_Analytics_${new Date().toISOString().split('T')[0]}.xls`);
    alert('✅ Complete history with KPI analytics exported successfully!');
}

/**
 * Generate history Excel table with KPI data
 * @returns {string} HTML table string
 */
function generateHistoryExcelTable() {
    let htmlTable = `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }
                .header { background-color: #4472C4; color: white; font-weight: bold; text-align: center; padding: 8px; font-size: 12px; }
                .kpi-header { background-color: #70AD47; color: white; font-weight: bold; text-align: center; padding: 8px; font-size: 12px; }
                .data-cell { text-align: center; padding: 6px; font-size: 11px; border: 1px solid #D0D7DE; }
                .green-cell { background-color: #C6EFCE; color: #006100; }
                .amber-cell { background-color: #FFEB9C; color: #9C5700; }
                .red-cell { background-color: #FFC7CE; color: #9C0006; }
            </style>
        </head>
        <body>
            <table border="1">
                <tr>
                    <td colspan="9" class="header">Complete Operations History Report with KPI Analytics - Dynamic Downtime-Based Calculation</td>
                </tr>
                <tr>
                    <td colspan="9" class="kpi-header">Monthly Average: ${monthlyKPICache ? monthlyKPICache.overallAvg : 0}% | Yearly Average: ${yearlyKPICache ? yearlyKPICache.overallAvg : 0}%</td>
                </tr>
                <tr class="header">
                    <td>Date</td><td>Landscape</td><td>Skill</td><td>Period</td><td>Status</td><td>Performance %</td><td>Downtime (h)</td><td>Incident</td><td>Notes</td>
                </tr>
    `;
    
    Object.entries(operationalData).forEach(([key, entry]) => {
        let cellClass = '';
        switch(entry.status) {
            case 'green': cellClass = 'green-cell'; break;
            case 'amber': cellClass = 'amber-cell'; break;
            case 'red': cellClass = 'red-cell'; break;
        }
        
        htmlTable += `
            <tr>
                <td class="data-cell">${entry.date}</td>
                <td class="data-cell">${entry.landscape}</td>
                <td class="data-cell">${entry.skill}</td>
                <td class="data-cell">${entry.period}</td>
                <td class="data-cell ${cellClass}">● ${entry.status.toUpperCase()}</td>
                <td class="data-cell">${entry.percentage}%</td>
                <td class="data-cell">${entry.downtime || 0}h</td>
                <td class="data-cell">${entry.incident || ''}</td>
                <td class="data-cell">${entry.notes || ''}</td>
            </tr>
        `;
    });
    
    htmlTable += '</table></body></html>';
    return htmlTable;
}

// ========================================
// HISTORY MANAGEMENT
// ========================================

/**
 * Show history modal
 */
function showHistoryModal() {
    populateHistoryTable();
    document.getElementById('historyModal').style.display = 'block';
    document.getElementById('historyModal').setAttribute('aria-hidden', 'false');
}

/**
 * Close history modal
 */
function closeHistoryModal() {
    document.getElementById('historyModal').style.display = 'none';
    document.getElementById('historyModal').setAttribute('aria-hidden', 'true');
}

/**
 * Populate history table
 */
function populateHistoryTable() {
    const tbody = document.getElementById('historyTableBody');
    const filter = document.getElementById('historyFilter') ? document.getElementById('historyFilter').value : '';
    let entries = [];
    
    Object.entries(operationalData).forEach(([key, entry]) => {
        if (!filter || entry.landscape === filter) {
            entries.push({key, ...entry});
        }
    });
    
    entries.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    let html = '';
    entries.forEach(entry => {
        const statusColor = entry.status === 'green' ? '#28a745' : 
                           entry.status === 'amber' ? '#ffc107' : 
                           entry.status === 'red' ? '#dc3545' : '#999';
        
        html += `
            <tr onclick="editHistoryEntry('${entry.key}')" style="cursor: pointer;" 
                onmouseover="this.style.backgroundColor='#f8f9fa'" 
                onmouseout="this.style.backgroundColor='white'">
                <td>${new Date(entry.timestamp).toLocaleDateString()}</td>
                <td>${entry.landscape}</td>
                <td>${entry.skill}</td>
                <td>${entry.period}</td>
                <td><span style="color: ${statusColor}; font-weight: bold; font-size: 18px;">●</span> ${entry.status}</td>
                <td>${entry.percentage}%</td>
                <td>${entry.downtime || 0}h</td>
                <td>${entry.incident || '-'}</td>
                <td>
                    <button class="edit-btn" onclick="event.stopPropagation(); editHistoryEntry('${entry.key}')">Edit</button>
                    <button class="delete-btn" onclick="event.stopPropagation(); deleteHistoryEntry('${entry.key}')">Delete</button>
                </td>
            </tr>
        `;
    });
    
    tbody.innerHTML = html;
}

/**
 * Filter history table
 */
function filterHistory() {
    populateHistoryTable();
}

/**
 * Edit history entry
 * @param {string} key - Entry key
 */
function editHistoryEntry(key) {
    closeHistoryModal();
    const entry = operationalData[key];
    editEntry(key, entry.skill, entry.period);
}

/**
 * Delete history entry
 * @param {string} key - Entry key
 */
function deleteHistoryEntry(key) {
    if (confirm('🗑️ Delete this history entry?')) {
        delete operationalData[key];
        localStorage.setItem('operationalData', JSON.stringify(operationalData));
        populateHistoryTable();
        loadCurrentView();
        updateKPIDisplays(); // Update KPIs after deletion
        alert('✅ Entry deleted!');
    }
}

// ========================================
// DATA MANAGEMENT
// ========================================

/**
 * Clear current view data
 */
function clearCurrentData() {
    const confirmText = currentView === 'monthly' ? 
        `Clear all data for ${currentMonth} in ${currentLandscape}?` :
        `Clear all data for ${currentDate} in ${currentLandscape}?`;
    
    if (confirm(`🗑️ ${confirmText}`)) {
        Object.keys(operationalData).forEach(key => {
            if (key.includes(currentLandscape)) {
                if (currentView === 'monthly' && key.includes(currentMonth)) {
                    delete operationalData[key];
                } else if (currentView === 'daily' && key.includes(currentDate)) {
                    delete operationalData[key];
                }
            }
        });
        
        localStorage.setItem('operationalData', JSON.stringify(operationalData));
        loadCurrentView();
        updateKPIDisplays(); // Update KPIs after clearing data
        alert('✅ Data cleared!');
    }
}

/**
 * Import data from file
 */
function importData() {
    document.getElementById('fileInput').click();
}

/**
 * Handle file import
 * @param {Event} event - File input change event
 */
function handleFileImport(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const importData = JSON.parse(e.target.result);
            
            if (confirm('📤 Import data? This will merge with existing data.')) {
                operationalData = {...operationalData, ...importData.data || importData};
                if (importData.skills) {
                    skillsList = [...new Set([...skillsList, ...importData.skills])];
                }
                if (importData.landscapes) {
                    landscapesList = [...new Set([...landscapesList, ...importData.landscapes])];
                    localStorage.setItem('landscapesList', JSON.stringify(landscapesList));
                    initializeLandscapes();
                }
                
                localStorage.setItem('operationalData', JSON.stringify(operationalData));
                localStorage.setItem('skillsList', JSON.stringify(skillsList));
                
                updateSkillList();
                loadCurrentView();
                updateKPIDisplays(); // Update KPIs after import
                alert('✅ Data imported successfully with KPI analytics updated!');
            }
        } catch (error) {
            alert('❌ Invalid file format!');
            console.error('Import error:', error);
        }
    };
    reader.readAsText(file);
}

// ========================================
// AUTO-SAVE & UTILITY FUNCTIONS
// ========================================

/**
 * Auto-save data to localStorage
 */
function autoSave() {
    if (Object.keys(operationalData).length > 0) {
        localStorage.setItem('operationalData', JSON.stringify(operationalData));
        localStorage.setItem('skillsList', JSON.stringify(skillsList));
        localStorage.setItem('landscapesList', JSON.stringify(landscapesList));
        console.log('🔄 Auto-save completed with KPI data');
    }
}

// ========================================
// EVENT LISTENERS & MODAL MANAGEMENT
// ========================================

/**
 * Close modals when clicking outside
 */
window.onclick = function(event) {
    if (event.target.classList.contains('modal')) {
        event.target.style.display = 'none';
        event.target.setAttribute('aria-hidden', 'true');
    }
}

/**
 * Handle keyboard events for accessibility
 */
document.addEventListener('keydown', function(event) {
    // Close modals with Escape key
    if (event.key === 'Escape') {
        const visibleModals = document.querySelectorAll('.modal[style*="block"]');
        visibleModals.forEach(modal => {
            modal.style.display = 'none';
            modal.setAttribute('aria-hidden', 'true');
        });
    }
});

// ========================================
// DEVELOPMENT & DEBUGGING
// ========================================

/**
 * Log application state with KPI data (for debugging)
 */
function debugAppState() {
    console.log('📊 Dashboard with KPI Analytics Debug Information:');
    console.log('Current View:', currentView);
    console.log('Current Landscape:', currentLandscape);
    console.log('Current Month:', currentMonth);
    console.log('Current Date:', currentDate);
    console.log('Skills Count:', skillsList.length);
    console.log('Landscapes Count:', landscapesList.length);
    console.log('Data Entries Count:', Object.keys(operationalData).length);
    console.log('Monthly KPIs:', monthlyKPICache);
    console.log('Yearly KPIs:', yearlyKPICache);
    console.log('Weekly KPIs:', weeklyKPICache);
    console.log('Full Data:', operationalData);
}

// Make debug function available globally for development
window.debugDashboard = debugAppState;

console.log('✅ Advanced Operations Dashboard with KPI Analytics JavaScript loaded successfully');
console.log('🔧 Type "debugDashboard()" in console for debug information including KPI data');