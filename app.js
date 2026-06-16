// Election Result Analysis Dashboard - Application Logic
// Handles state, navigation, filtering, rendering, exports, and in-browser uploads

// 1. STATE MANAGEMENT
let state = {
  electionData: JSON.parse(JSON.stringify(DEFAULT_ELECTION_DATA)), // Deep clone default data
  activePage: 'district-overview',
  selectedAC: 209,
  theme: 'dark',
  
  // Booth Ledger Filters
  boothFilters: {
    ac: 'all',
    winner: 'all',
    margin: 'all',
    turnout: 'all',
    search: '',
    page: 1,
    pageSize: 15
  },
  
  // Agent Filters
  agentFilters: {
    ac: '209',
    party: 'all',
    rating: 'all',
    search: '',
    page: 1,
    pageSize: 15
  },
  
  // Cluster Settings
  clusterSettings: {
    ac: 209,
    start: 1,
    end: 50
  },
  
  // AI Insights settings
  aiSettings: {
    ac: 209
  },

  // Active files uploaded
  uploadedFiles: [
    { name: 'AC209_PARAMAKUDI(SC)_Election_Results_2026.xlsx', category: 'Results', ac: '209', records: 323, status: 'Default' },
    { name: 'AC209_PARAMAKUDI(SC)_BLA_Details.xlsx', category: 'Agents', ac: '209', records: 629, status: 'Default' },
    { name: 'AC210_Tiruvadanai_Election_Results_2026.xlsx', category: 'Results', ac: '210', records: 378, status: 'Default' },
    { name: 'AC210_Tiruvadanai_BLA2_Details.xlsx', category: 'Agents', ac: '210', records: 1583, status: 'Default' },
    { name: 'AC211_Ramanathapuram_Election_Results_2026.xlsx', category: 'Results', ac: '211', records: 399, status: 'Default' },
    { name: 'AC211_Ramanathapuram_BLA2_Details.xlsx', category: 'Agents', ac: '211', records: 1497, status: 'Default' },
    { name: 'AC212_Mudhukulathur_Election_Results_2026.xlsx', category: 'Results', ac: '212', records: 414, status: 'Default' },
    { name: 'AC212_Mudhukulathur_BLA2_Details.xlsx', category: 'Agents', ac: '212', records: 1765, status: 'Default' }
  ]
};

// Store active Chart.js instances to avoid canvas reuse errors
let activeCharts = {};

// Party hex color mapper
const partyColors = {
  'DMK': '#f59e0b',
  'TVK': '#8b5cf6',
  'AIADMK': '#10b981',
  'BJP': '#f97316',
  'NTK': '#ef4444',
  'INC': '#3b82f6',
  'AIPTMMK': '#ec4899',
  'PT': '#06b6d4',
  'Others': '#6b7280',
  'NOTA': '#9ca3af'
};

// 2. DOCUMENT READY & INIT
document.addEventListener('DOMContentLoaded', () => {
  initNavigation();
  initTheme();
  initFilters();
  renderActivePage();
  
  // PDF Export
  document.getElementById('export-pdf-btn').addEventListener('click', exportDashboardPDF);
});

// 3. THEME TOGGLE
function initTheme() {
  const toggle = document.getElementById('theme-toggle');
  const iconContainer = document.getElementById('theme-icon-container');
  
  toggle.addEventListener('click', () => {
    state.theme = state.theme === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', state.theme);
    
    if (state.theme === 'dark') {
      iconContainer.innerHTML = '<i class="fa-solid fa-moon"></i> Dark';
    } else {
      iconContainer.innerHTML = '<i class="fa-solid fa-sun"></i> Light';
    }
    
    // Re-render active page to update chart text colors if necessary
    renderActivePage();
  });
}

// 4. NAVIGATION ROUTING
function initNavigation() {
  const menuItems = document.querySelectorAll('.sidebar-menu .menu-item');
  const toggleBtn = document.getElementById('menu-toggle');
  const sidebar = document.getElementById('app-sidebar');
  
  menuItems.forEach(item => {
    item.addEventListener('click', (e) => {
      e.preventDefault();
      
      const page = item.getAttribute('data-page');
      
      // Update sidebar state
      menuItems.forEach(i => i.classList.remove('active'));
      item.classList.add('active');
      
      // Close sidebar on mobile
      sidebar.classList.remove('active');
      
      // If constituency link is clicked
      if (page.startsWith('ac-')) {
        state.selectedAC = parseInt(page.split('-')[1]);
        state.activePage = 'constituency-view';
      } else {
        state.activePage = page;
      }
      
      renderActivePage();
    });
  });
  
  toggleBtn.addEventListener('click', () => {
    sidebar.classList.toggle('active');
  });
  
  // Close modal click handlers
  document.getElementById('modal-close-btn').addEventListener('click', hideModal);
  document.getElementById('detail-modal').addEventListener('click', (e) => {
    if (e.target.id === 'detail-modal') hideModal();
  });
}

// 5. FILTER PANELS INITIALIZATION
function initFilters() {
  // Booth Table Filters
  document.getElementById('filter-booth-ac').addEventListener('change', (e) => {
    state.boothFilters.ac = e.target.value;
    state.boothFilters.page = 1;
    renderBoothLedgerTable();
  });
  document.getElementById('filter-booth-winner').addEventListener('change', (e) => {
    state.boothFilters.winner = e.target.value;
    state.boothFilters.page = 1;
    renderBoothLedgerTable();
  });
  document.getElementById('filter-booth-margin').addEventListener('change', (e) => {
    state.boothFilters.margin = e.target.value;
    state.boothFilters.page = 1;
    renderBoothLedgerTable();
  });
  document.getElementById('filter-booth-turnout').addEventListener('change', (e) => {
    state.boothFilters.turnout = e.target.value;
    state.boothFilters.page = 1;
    renderBoothLedgerTable();
  });
  document.getElementById('search-booth-input').addEventListener('input', (e) => {
    state.boothFilters.search = e.target.value;
    state.boothFilters.page = 1;
    renderBoothLedgerTable();
  });
  document.getElementById('reset-booth-filters-btn').addEventListener('click', () => {
    document.getElementById('filter-booth-ac').value = 'all';
    document.getElementById('filter-booth-winner').value = 'all';
    document.getElementById('filter-booth-margin').value = 'all';
    document.getElementById('filter-booth-turnout').value = 'all';
    document.getElementById('search-booth-input').value = '';
    
    state.boothFilters.ac = 'all';
    state.boothFilters.winner = 'all';
    state.boothFilters.margin = 'all';
    state.boothFilters.turnout = 'all';
    state.boothFilters.search = '';
    state.boothFilters.page = 1;
    renderBoothLedgerTable();
  });

  // Agent Filters
  document.getElementById('filter-agent-ac').addEventListener('change', (e) => {
    state.agentFilters.ac = e.target.value;
    state.agentFilters.page = 1;
    populateAgentPartyDropdown();
    renderAgentPerformanceSection();
  });
  document.getElementById('filter-agent-party').addEventListener('change', (e) => {
    state.agentFilters.party = e.target.value;
    state.agentFilters.page = 1;
    renderAgentPerformanceSection();
  });
  document.getElementById('filter-agent-rating').addEventListener('change', (e) => {
    state.agentFilters.rating = e.target.value;
    state.agentFilters.page = 1;
    renderAgentPerformanceSection();
  });
  document.getElementById('search-agent-input').addEventListener('input', (e) => {
    state.agentFilters.search = e.target.value;
    state.agentFilters.page = 1;
    renderAgentPerformanceSection();
  });

  // Cluster Analysis
  document.getElementById('cluster-constituency-select').addEventListener('change', (e) => {
    state.clusterSettings.ac = parseInt(e.target.value);
    populatePredefinedClusters();
  });
  document.getElementById('btn-analyze-cluster').addEventListener('click', () => {
    state.clusterSettings.start = parseInt(document.getElementById('cluster-range-start').value) || 1;
    state.clusterSettings.end = parseInt(document.getElementById('cluster-range-end').value) || 50;
    renderClusterAnalysis();
  });

  // AI Insights
  document.getElementById('ai-constituency-select').addEventListener('change', (e) => {
    state.aiSettings.ac = parseInt(e.target.value);
    renderAIInsightsSection();
  });
  
  // Ledger Export Handlers
  document.getElementById('export-booth-csv').addEventListener('click', () => exportBoothTable('csv'));
  document.getElementById('export-booth-xlsx').addEventListener('click', () => exportBoothTable('xlsx'));
  document.getElementById('export-agent-csv').addEventListener('click', exportAgentTable);
}

// 6. MAIN ROUTER RENDERER
function renderActivePage() {
  // Hide all sections
  const sections = document.querySelectorAll('.dashboard-section');
  sections.forEach(s => s.classList.remove('active'));
  
  // Reset Display Titles
  const titleEl = document.getElementById('page-display-title');
  const subtitleEl = document.getElementById('page-display-subtitle');
  
  // Destroy previous charts to free memory
  Object.keys(activeCharts).forEach(key => {
    if (activeCharts[key]) {
      activeCharts[key].destroy();
      activeCharts[key] = null;
    }
  });

  if (state.activePage === 'district-overview') {
    document.getElementById('sect-district-overview').classList.add('active');
    titleEl.innerText = "District Overview";
    subtitleEl.innerText = `Ramanathapuram District General Election Results | 4 Assembly Constituencies`;
    renderDistrictOverview();
  } 
  else if (state.activePage === 'constituency-view') {
    document.getElementById('sect-constituency-view').classList.add('active');
    const acData = state.electionData.constituencies[state.selectedAC];
    titleEl.innerText = `AC ${acData.ac_no} ${acData.ac_name}`;
    subtitleEl.innerText = `Constituency Level Analytical Intelligence & Booth Matrix`;
    renderConstituencyView();
  }
  else if (state.activePage === 'booth-analysis') {
    document.getElementById('sect-booth-analysis').classList.add('active');
    titleEl.innerText = "Booth-Wise Result Ledger";
    subtitleEl.innerText = "Deep filterable audit ledger of voting performance across all 1,514 polling stations";
    renderBoothLedgerTable();
  }
  else if (state.activePage === 'booth-cluster') {
    document.getElementById('sect-booth-cluster').classList.add('active');
    titleEl.innerText = "Booth Cluster Analytics";
    subtitleEl.innerText = "Define and compare range of booths to isolate voter concentration patterns";
    populatePredefinedClusters();
    renderClusterAnalysis();
  }
  else if (state.activePage === 'agent-analysis') {
    document.getElementById('sect-agent-analysis').classList.add('active');
    titleEl.innerText = "Booth Agent Intelligence Matrix";
    subtitleEl.innerText = "Electoral effectiveness ratings, conversion ratios, and turnout analysis for BLA-2 agents";
    populateAgentPartyDropdown();
    renderAgentPerformanceSection();
  }
  else if (state.activePage === 'ai-insights') {
    document.getElementById('sect-ai-insights').classList.add('active');
    titleEl.innerText = "AI Strategic Insights & Re-Election Planning";
    subtitleEl.innerText = "Algorithmic targeting model based on victory margins, turnout conversion, and stronghold audits";
    renderAIInsightsSection();
  }
}

// ==============================================
// PAGE RENDERERS
// ==============================================

// A. DISTRICT OVERVIEW PAGE
function renderDistrictOverview() {
  const sum = state.electionData.summary;
  
  // Populate summary cards
  const cardsContainer = document.getElementById('district-summary-cards');
  cardsContainer.innerHTML = `
    <div class="glass-card summary-card">
      <span class="summary-label">Total Votes Polled</span>
      <span class="summary-value">${sum.total_votes_polled.toLocaleString()}</span>
      <span class="summary-subtitle" style="font-size:0.75rem; color:var(--text-muted);">EVM + Postal Ballot</span>
    </div>
    <div class="glass-card summary-card">
      <span class="summary-label">Total Registered Electors</span>
      <span class="summary-value">${sum.total_electors.toLocaleString()}</span>
      <span class="summary-subtitle" style="font-size:0.75rem; color:var(--text-muted);">District Population baseline</span>
    </div>
    <div class="glass-card summary-card">
      <span class="summary-label">District Turnout</span>
      <span class="summary-value">${sum.polling_percentage}%</span>
      <span class="summary-trend trend-up"><i class="fa-solid fa-circle-check"></i> High Turnout</span>
    </div>
    <div class="glass-card summary-card">
      <span class="summary-label">Total Polling Booths</span>
      <span class="summary-value">${sum.total_booths}</span>
      <span class="summary-subtitle" style="font-size:0.75rem; color:var(--text-muted);">across 4 constituencies</span>
    </div>
  `;

  // Winners Table
  const tbody = document.querySelector('#district-winner-table tbody');
  tbody.innerHTML = '';
  Object.values(state.electionData.constituencies).forEach(ac => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${ac.ac_no}</strong></td>
      <td><strong>${ac.ac_name}</strong></td>
      <td>${ac.total_votes_polled.toLocaleString()}</td>
      <td><strong>${ac.winning_candidate.split('(')[0].trim()}</strong></td>
      <td><span class="badge badge-${ac.winning_party.toLowerCase()}">${ac.winning_party}</span></td>
      <td><span style="font-weight: 700; color: #10b981;">+${ac.victory_margin.toLocaleString()}</span></td>
      <td><strong>${ac.polling_percentage}%</strong></td>
      <td>
        <button class="btn btn-secondary" style="padding: 4px 8px; font-size: 0.75rem;" onclick="viewACDetails(${ac.ac_no})">
          <i class="fa-solid fa-chart-line"></i> View
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  // Render Charts
  renderDistrictCharts();
}

function viewACDetails(acNo) {
  state.selectedAC = acNo;
  state.activePage = 'constituency-view';
  
  // Update active class in sidebar menu
  const menuItems = document.querySelectorAll('.sidebar-menu .menu-item');
  menuItems.forEach(item => {
    item.classList.remove('active');
    if (item.getAttribute('data-page') === `ac-${acNo}`) {
      item.classList.add('active');
    }
  });
  
  renderActivePage();
}

function renderDistrictCharts() {
  const sum = state.electionData.summary;
  
  // 1. Party Vote Share (Pie Chart)
  const pieCtx = document.getElementById('district-vote-share-chart').getContext('2d');
  const labels = [];
  const data = [];
  const colors = [];
  
  // Sort party vote shares
  const sortedShares = Object.entries(sum.party_vote_share)
    .filter(([p, s]) => p !== 'NOTA' && s > 0.5)
    .sort((a,b) => b[1] - a[1]);
    
  // Add NOTA
  if (sum.party_vote_share['NOTA']) {
    sortedShares.push(['NOTA', sum.party_vote_share['NOTA']]);
  }
  
  sortedShares.forEach(([party, share]) => {
    labels.push(party);
    data.push(share);
    colors.push(partyColors[party] || partyColors['Others']);
  });

  activeCharts['district-pie'] = new Chart(pieCtx, {
    type: 'doughnut',
    data: {
      labels: labels,
      datasets: [{
        data: data,
        backgroundColor: colors,
        borderColor: state.theme === 'dark' ? '#0f172a' : '#ffffff',
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'right',
          labels: {
            color: state.theme === 'dark' ? '#94a3b8' : '#475569',
            font: { family: 'Plus Jakarta Sans', size: 11, weight: 600 }
          }
        },
        tooltip: {
          callbacks: {
            label: function(context) {
              return ` ${context.label}: ${context.raw}%`;
            }
          }
        }
      },
      cutout: '60%'
    }
  });

  // 2. Constituency Turnout Comparison (Bar Chart)
  const barCtx = document.getElementById('district-turnout-comparison-chart').getContext('2d');
  const acLabels = [];
  const acTurnouts = [];
  const barColors = [];
  
  Object.values(state.electionData.constituencies).forEach(ac => {
    acLabels.push(ac.ac_name);
    acTurnouts.push(ac.polling_percentage);
    barColors.push(partyColors[ac.winning_party] || '#6366f1');
  });

  activeCharts['district-bar'] = new Chart(barCtx, {
    type: 'bar',
    data: {
      labels: acLabels,
      datasets: [{
        label: 'Turnout %',
        data: acTurnouts,
        backgroundColor: barColors.map(c => c + 'cc'), // Add opacity
        borderColor: barColors,
        borderWidth: 1,
        borderRadius: 8
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false }
      },
      scales: {
        y: {
          min: 50,
          max: 100,
          grid: { color: state.theme === 'dark' ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.05)' },
          ticks: { color: state.theme === 'dark' ? '#94a3b8' : '#475569' }
        },
        x: {
          grid: { display: false },
          ticks: { color: state.theme === 'dark' ? '#94a3b8' : '#475569' }
        }
      }
    }
  });
}

// B. INDIVIDUAL CONSTITUENCY VIEW
function renderConstituencyView() {
  const ac = state.electionData.constituencies[state.selectedAC];
  
  // Populate summaries
  const cardsContainer = document.getElementById('ac-summary-cards');
  cardsContainer.innerHTML = `
    <div class="glass-card summary-card">
      <span class="summary-label">Total Votes Polled</span>
      <span class="summary-value">${ac.total_votes_polled.toLocaleString()}</span>
      <span class="summary-subtitle" style="font-size:0.75rem; color:var(--text-muted);">EVM + Postal Ballots</span>
    </div>
    <div class="glass-card summary-card">
      <span class="summary-label">Victory Winner</span>
      <span class="summary-value" style="font-size: 1.5rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">
        ${ac.winning_candidate.split('(')[0].trim()}
      </span>
      <span class="summary-subtitle" style="margin-top: 2px;">
        <span class="badge badge-${ac.winning_party.toLowerCase()}">${ac.winning_party}</span>
      </span>
    </div>
    <div class="glass-card summary-card">
      <span class="summary-label">Victory Margin</span>
      <span class="summary-value" style="color:#10b981;">+${ac.victory_margin.toLocaleString()}</span>
      <span class="summary-subtitle" style="font-size:0.75rem; color:var(--text-muted);">votes over runner-up</span>
    </div>
    <div class="glass-card summary-card">
      <span class="summary-label">AC Turnout %</span>
      <span class="summary-value">${ac.polling_percentage}%</span>
      <span class="summary-subtitle" style="font-size:0.75rem; color:var(--text-muted);">${ac.total_booths} Polling Stations</span>
    </div>
  `;

  // Candidate vote stack list
  const stackContainer = document.getElementById('ac-vote-stack');
  stackContainer.innerHTML = '';
  
  // Show top candidates
  ac.candidates.forEach(cand => {
    const cleanName = cand.name.split('(')[0].trim();
    const partyParts = cand.name.split('(');
    const party = partyParts.length > 1 ? partyParts[partyParts.length - 1].replace(')', '') : 'Others';
    const color = partyColors[party] || partyColors['Others'];
    
    const stackItem = document.createElement('div');
    stackItem.className = 'vote-stack-item';
    stackItem.innerHTML = `
      <div class="vote-stack-header" style="margin-bottom: 2px;">
        <span><strong>${cleanName}</strong> <span style="font-size:0.75rem; color:var(--text-muted);">(${party})</span></span>
        <span><strong>${cand.votes.toLocaleString()}</strong> (${cand.share_pct}%)</span>
      </div>
      <div style="font-size: 0.75rem; color: var(--text-muted); margin-bottom: 6px; display: flex; gap: 8px;">
        <span>EVM: <strong>${(cand.evm_votes || 0).toLocaleString()}</strong></span>
        <span>|</span>
        <span>Postal: <strong>${(cand.postal_votes || 0).toLocaleString()}</strong></span>
      </div>
      <div class="vote-stack-progress">
        <div class="vote-stack-bar" style="width: ${cand.share_pct}%; background-color: ${color};"></div>
      </div>
    `;
    stackContainer.appendChild(stackItem);
  });

  // Load strong / weak / close booths lists
  const sortedByMarginDesc = [...ac.booths].sort((a,b) => b.margin - a.margin);
  
  // 1. Strongest Booths (highest winning votes for the winner candidate)
  const strongBody = document.querySelector('#ac-strong-booths-table tbody');
  strongBody.innerHTML = '';
  // Strongest booths for the constituency winner
  const winnerParty = ac.winning_party;
  const winnerStrongBooths = [...ac.booths]
    .filter(b => b.winner_party === winnerParty)
    .sort((a,b) => b.candidate_votes[ac.winning_candidate] - a.candidate_votes[ac.winning_candidate])
    .slice(0, 5);
    
  winnerStrongBooths.forEach(b => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>#${b.booth_no}</strong></td>
      <td><span class="badge badge-${b.winner_party.toLowerCase()}">${b.winner_party}</span></td>
      <td>${b.candidate_votes[ac.winning_candidate]}</td>
      <td><span style="color:#10b981; font-weight:700;">+${b.margin}</span></td>
    `;
    strongBody.appendChild(tr);
  });

  // 2. Weakest Booths (lowest votes for the constituency winner or biggest losses)
  const weakBody = document.querySelector('#ac-weak-booths-table tbody');
  weakBody.innerHTML = '';
  const constituencyWinnerLosses = [...ac.booths]
    .filter(b => b.winner_party !== winnerParty)
    .sort((a,b) => b.margin - a.margin) // biggest defeats first
    .slice(0, 5);
    
  // If no losses, just take candidate's lowest vote booths
  const weakList = constituencyWinnerLosses.length >= 5 ? constituencyWinnerLosses : 
    [...ac.booths].sort((a,b) => a.candidate_votes[ac.winning_candidate] - b.candidate_votes[ac.winning_candidate]).slice(0,5);

  weakList.forEach(b => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>#${b.booth_no}</strong></td>
      <td><span class="badge badge-${b.winner_party.toLowerCase()}">${b.winner_party}</span></td>
      <td>${b.candidate_votes[ac.winning_candidate] || 0}</td>
      <td><span style="color:#ef4444; font-weight:700;">-${b.margin}</span></td>
    `;
    weakBody.appendChild(tr);
  });

  // 3. Close Contest Booths (margin < 30 votes)
  const closeBody = document.querySelector('#ac-close-booths-table tbody');
  closeBody.innerHTML = '';
  const closeContestBooths = [...ac.booths]
    .sort((a,b) => a.margin - b.margin)
    .slice(0, 5);

  closeContestBooths.forEach(b => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>#${b.booth_no}</strong></td>
      <td><span class="badge badge-${b.winner_party.toLowerCase()}">${b.winner_party}</span></td>
      <td><span style="color:#f59e0b; font-weight:700;">${b.margin}</span></td>
      <td>${b.turnout_pct}%</td>
    `;
    closeBody.appendChild(tr);
  });

  // 4. Heatmap Rendering
  const heatmapGrid = document.getElementById('ac-heatmap-grid');
  heatmapGrid.innerHTML = '';
  ac.booths.forEach(b => {
    const cell = document.createElement('div');
    cell.className = 'heatmap-cell';
    const color = partyColors[b.winner_party] || '#6b7280';
    cell.style.backgroundColor = color;
    cell.innerText = b.booth_no;
    cell.setAttribute('data-tooltip', `Booth ${b.booth_no}: ${b.booth_name} | Winner: ${b.winner_party} (+${b.margin} votes)`);
    
    cell.addEventListener('click', () => showBoothDetails(b, ac.ac_no));
    heatmapGrid.appendChild(cell);
  });

  // Constituency charts
  renderConstituencyCharts(ac);
}

function renderConstituencyCharts(ac) {
  // Aggregate party totals in constituency
  const partyVotes = {};
  ac.booths.forEach(b => {
    Object.entries(b.votes).forEach(([p, v]) => {
      if (p !== 'NOTA') {
        partyVotes[p] = (partyVotes[p] || 0) + v;
      }
    });
  });

  const labels = Object.keys(partyVotes).sort((a,b) => partyVotes[b] - partyVotes[a]).slice(0, 5);
  const data = labels.map(l => partyVotes[l]);
  const colors = labels.map(l => partyColors[l] || '#6366f1');

  const ctx = document.getElementById('ac-party-performance-chart').getContext('2d');
  
  activeCharts['ac-bar'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Votes Secured',
        data: data,
        backgroundColor: colors.map(c => c + 'dd'),
        borderColor: colors,
        borderWidth: 1,
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false }
      },
      scales: {
        y: {
          grid: { color: state.theme === 'dark' ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.05)' },
          ticks: { color: state.theme === 'dark' ? '#94a3b8' : '#475569' }
        },
        x: {
          grid: { display: false },
          ticks: { color: state.theme === 'dark' ? '#94a3b8' : '#475569' }
        }
      }
    }
  });
}

// C. BOOTH LEDGER ANALYSIS TABLE (SEARCHABLE & FILTERABLE)
function renderBoothLedgerTable() {
  const f = state.boothFilters;
  
  // Assemble booths across matching constituencies
  let boothsList = [];
  Object.keys(state.electionData.constituencies).forEach(acKey => {
    const acNo = parseInt(acKey);
    if (f.ac === 'all' || f.ac == acNo) {
      state.electionData.constituencies[acKey].booths.forEach(b => {
        boothsList.push({
          ...b,
          ac_no: acNo,
          ac_name: state.electionData.constituencies[acKey].ac_name
        });
      });
    }
  });

  // Apply filters
  // 1. Search filter
  if (f.search.trim()) {
    const q = f.search.toLowerCase().trim();
    boothsList = boothsList.filter(b => 
      b.booth_name.toLowerCase().includes(q) || 
      b.booth_no.toString() === q ||
      b.winner_party.toLowerCase().includes(q)
    );
  }

  // 2. Winner party filter
  if (f.winner !== 'all') {
    boothsList = boothsList.filter(b => b.winner_party === f.winner);
  }

  // 3. Margin range filter
  if (f.margin !== 'all') {
    if (f.margin === 'high') {
      boothsList = boothsList.filter(b => b.margin >= 200);
    } else if (f.margin === 'mid') {
      boothsList = boothsList.filter(b => b.margin >= 50 && b.margin < 200);
    } else if (f.margin === 'close') {
      boothsList = boothsList.filter(b => b.margin < 50);
    }
  }

  // 4. Turnout filter
  if (f.turnout !== 'all') {
    if (f.turnout === 'high') {
      boothsList = boothsList.filter(b => b.turnout_pct >= 80);
    } else if (f.turnout === 'low') {
      boothsList = boothsList.filter(b => b.turnout_pct < 72);
    }
  }

  // Calculate pages
  const totalCount = boothsList.length;
  const totalPages = Math.ceil(totalCount / f.pageSize) || 1;
  if (f.page > totalPages) f.page = totalPages;
  
  const startIndex = (f.page - 1) * f.pageSize;
  const paginatedBooths = boothsList.slice(startIndex, startIndex + f.pageSize);

  // Determine Table Headers
  const tableHeader = document.querySelector('#booth-ledger-table theological'); // Wait, let's look at index.html: thethead is empty!
  const thead = document.querySelector('#booth-ledger-table').createTHead();
  thead.innerHTML = '';
  const headerRow = thead.insertRow();
  
  const columns = ['Booth No', 'Booth Name', 'Constituency', 'Total Votes', 'Turnout %'];
  
  // If specific constituency is selected, we can show its candidate vote columns!
  let candidatesForAC = [];
  if (f.ac !== 'all') {
    const acNo = parseInt(f.ac);
    candidatesForAC = state.electionData.constituencies[acNo].candidates.slice(0, 4); // top 4 candidates
    candidatesForAC.forEach(c => {
      columns.push(c.name.split('(')[0].trim());
    });
  } else {
    // Show top district parties as columns
    columns.push('DMK', 'TVK', 'AIADMK');
  }
  
  columns.push('Winner in Booth', 'Winning Margin', 'Actions');

  columns.forEach(col => {
    const th = document.createElement('th');
    th.innerText = col;
    headerRow.appendChild(th);
  });

  // Populate Table Body
  const tbody = document.querySelector('#booth-ledger-table tbody') || document.querySelector('#booth-ledger-table').createTBody();
  tbody.innerHTML = '';
  
  if (paginatedBooths.length === 0) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="${columns.length}" style="text-align:center; padding: 30px; color: var(--text-secondary);">No polling stations match the specified filters.</td>`;
    tbody.appendChild(tr);
  } else {
    paginatedBooths.forEach(b => {
      const tr = document.createElement('tr');
      
      let html = `
        <td><strong>#${b.booth_no}</strong></td>
        <td style="max-width: 250px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${b.booth_name}">${b.booth_name}</td>
        <td><strong>${b.ac_name}</strong></td>
        <td>${b.total_votes.toLocaleString()}</td>
        <td><strong>${b.turnout_pct}%</strong></td>
      `;
      
      if (f.ac !== 'all') {
        candidatesForAC.forEach(c => {
          const votesVal = b.candidate_votes[c.name] || 0;
          html += `<td>${votesVal}</td>`;
        });
      } else {
        html += `
          <td>${b.votes['DMK'] || 0}</td>
          <td>${b.votes['TVK'] || 0}</td>
          <td>${b.votes['AIADMK'] || 0}</td>
        `;
      }
      
      html += `
        <td><span class="badge badge-${b.winner_party.toLowerCase()}">${b.winner_party}</span></td>
        <td><span style="font-weight: 700; color: #10b981;">+${b.margin}</span></td>
        <td>
          <button class="btn btn-secondary" style="padding: 4px 8px; font-size: 0.75rem;" onclick="viewBoothFromId(${b.ac_no}, ${b.booth_no})">
            <i class="fa-solid fa-circle-info"></i> Details
          </button>
        </td>
      `;
      
      tr.innerHTML = html;
      tbody.appendChild(tr);
    });
  }

  // Update Pagination Controls
  document.getElementById('booth-pagination-info').innerText = `Showing ${totalCount === 0 ? 0 : startIndex + 1} to ${Math.min(startIndex + f.pageSize, totalCount)} of ${totalCount} booths`;
  
  const pagControls = document.getElementById('booth-pagination-controls');
  pagControls.innerHTML = '';
  
  // Previous button
  const prevBtn = document.createElement('button');
  prevBtn.className = 'page-btn';
  prevBtn.innerHTML = '<i class="fa-solid fa-chevron-left"></i>';
  prevBtn.disabled = f.page === 1;
  prevBtn.addEventListener('click', () => {
    f.page--;
    renderBoothLedgerTable();
  });
  pagControls.appendChild(prevBtn);

  // Page numbers
  const maxPagesToShow = 5;
  let startPage = Math.max(1, f.page - 2);
  let endPage = Math.min(totalPages, startPage + maxPagesToShow - 1);
  if (endPage - startPage < maxPagesToShow - 1) {
    startPage = Math.max(1, endPage - maxPagesToShow + 1);
  }

  for (let p = startPage; p <= endPage; p++) {
    const pBtn = document.createElement('button');
    pBtn.className = `page-btn ${f.page === p ? 'active' : ''}`;
    pBtn.innerText = p;
    pBtn.addEventListener('click', () => {
      f.page = p;
      renderBoothLedgerTable();
    });
    pagControls.appendChild(pBtn);
  }

  // Next button
  const nextBtn = document.createElement('button');
  nextBtn.className = 'page-btn';
  nextBtn.innerHTML = '<i class="fa-solid fa-chevron-right"></i>';
  nextBtn.disabled = f.page === totalPages;
  nextBtn.addEventListener('click', () => {
    f.page++;
    renderBoothLedgerTable();
  });
  pagControls.appendChild(nextBtn);
}

function viewBoothFromId(acNo, boothNo) {
  const ac = state.electionData.constituencies[acNo];
  const booth = ac.booths.find(b => b.booth_no === boothNo);
  if (booth) {
    showBoothDetails(booth, acNo);
  }
}

// D. BOOTH CLUSTER ANALYSIS
function populatePredefinedClusters() {
  const acNo = state.clusterSettings.ac;
  const acData = state.electionData.constituencies[acNo];
  const totalBooths = acData.total_booths;
  
  const container = document.getElementById('predefined-cluster-buttons');
  container.innerHTML = '';
  
  const step = 25;
  for (let i = 1; i <= totalBooths; i += step) {
    const end = Math.min(i + step - 1, totalBooths);
    const btn = document.createElement('button');
    btn.className = 'cluster-btn';
    if (state.clusterSettings.start === i && state.clusterSettings.end === end) {
      btn.classList.add('active');
    }
    btn.innerText = `Booth ${i} to ${end}`;
    btn.addEventListener('click', () => {
      document.getElementById('cluster-range-start').value = i;
      document.getElementById('cluster-range-end').value = end;
      state.clusterSettings.start = i;
      state.clusterSettings.end = end;
      
      // Update active state
      document.querySelectorAll('.cluster-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      
      renderClusterAnalysis();
    });
    container.appendChild(btn);
  }
}

function renderClusterAnalysis() {
  const s = state.clusterSettings;
  const ac = state.electionData.constituencies[s.ac];
  
  // Filter booths in range
  const filteredBooths = ac.booths.filter(b => b.booth_no >= s.start && b.booth_no <= s.end);
  const boothCount = filteredBooths.length;
  
  // Accumulate
  let totalVotes = 0;
  let totalElectors = 0;
  const partyVotes = {};
  
  filteredBooths.forEach(b => {
    totalVotes += b.total_votes;
    totalElectors += b.electors;
    Object.entries(b.votes).forEach(([p, v]) => {
      partyVotes[p] = (partyVotes[p] || 0) + v;
    });
  });

  const turnoutPct = totalElectors > 0 ? ((totalVotes / totalElectors) * 100).toFixed(2) : '0.00';
  
  // Find cluster winner
  const sortedParties = Object.entries(partyVotes)
    .filter(([p, v]) => p !== 'NOTA')
    .sort((a,b) => b[1] - a[1]);
  const clusterWinner = sortedParties.length > 0 ? `${sortedParties[0][0]} (${((sortedParties[0][1]/totalVotes)*100).toFixed(1)}%)` : 'None';
  
  // Populate metrics card
  document.getElementById('cluster-val-booths').innerText = boothCount;
  document.getElementById('cluster-val-votes').innerText = totalVotes.toLocaleString();
  document.getElementById('cluster-val-turnout').innerText = `${turnoutPct}%`;
  document.getElementById('cluster-val-winner').innerText = clusterWinner;
  
  if (sortedParties.length > 0) {
    const winnerPartyName = sortedParties[0][0];
    document.getElementById('cluster-val-winner').className = `summary-value badge badge-${winnerPartyName.toLowerCase()}`;
    document.getElementById('cluster-val-winner').style.display = 'inline-flex';
    document.getElementById('cluster-val-winner').style.fontSize = '1.05rem';
    document.getElementById('cluster-val-winner').style.padding = '8px 12px';
  } else {
    document.getElementById('cluster-val-winner').className = 'summary-value';
  }

  // Populate Progress share bars
  const progressContainer = document.getElementById('cluster-party-progress-stack');
  progressContainer.innerHTML = '';
  
  sortedParties.forEach(([party, votes]) => {
    const share = totalVotes > 0 ? ((votes / totalVotes) * 100).toFixed(1) : 0;
    const color = partyColors[party] || '#6366f1';
    
    const div = document.createElement('div');
    div.className = 'vote-stack-item';
    div.innerHTML = `
      <div class="vote-stack-header">
        <span><strong>${party}</strong></span>
        <span><strong>${votes.toLocaleString()}</strong> (${share}%)</span>
      </div>
      <div class="vote-stack-progress">
        <div class="vote-stack-bar" style="width: ${share}%; background-color: ${color};"></div>
      </div>
    `;
    progressContainer.appendChild(div);
  });

  // Render Cluster Pie Chart
  renderClusterChart(partyVotes, totalVotes);
}

function renderClusterChart(partyVotes, totalVotes) {
  const labels = [];
  const data = [];
  const colors = [];
  
  // Sort
  const sorted = Object.entries(partyVotes)
    .sort((a,b) => b[1] - a[1])
    .slice(0, 5);
    
  sorted.forEach(([party, votes]) => {
    const share = totalVotes > 0 ? ((votes / totalVotes) * 100).toFixed(1) : 0;
    if (share > 0.5) {
      labels.push(party);
      data.push(share);
      colors.push(partyColors[party] || '#6b7280');
    }
  });

  const ctx = document.getElementById('cluster-vote-share-chart').getContext('2d');
  
  if (activeCharts['cluster-pie']) {
    activeCharts['cluster-pie'].destroy();
  }
  
  activeCharts['cluster-pie'] = new Chart(ctx, {
    type: 'pie',
    data: {
      labels: labels,
      datasets: [{
        data: data,
        backgroundColor: colors,
        borderColor: state.theme === 'dark' ? '#0f172a' : '#ffffff',
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'right',
          labels: {
            color: state.theme === 'dark' ? '#94a3b8' : '#475569',
            font: { family: 'Plus Jakarta Sans', size: 11, weight: 600 }
          }
        }
      }
    }
  });
}

// Helper to mask phone numbers for privacy
function maskPhoneNumber(phone) {
  if (!phone) return '';
  const clean = phone.trim();
  if (clean.includes('******')) return clean;
  
  const parts = clean.split(' ');
  if (parts.length === 2) {
    const country = parts[0];
    const local = parts[1];
    if (local.length >= 7) {
      return `${country} ******${local.slice(-4)}`;
    }
  }
  
  const noSpace = clean.replace(/\s+/g, '');
  if (noSpace.length >= 10) {
    return `${noSpace.slice(0, noSpace.length - 7)}******${noSpace.slice(-4)}`;
  }
  return clean;
}

// E. BOOTH AGENT PERFORMANCE ANALYSIS
function populateAgentPartyDropdown() {
  const acNo = state.agentFilters.ac;
  const acData = state.electionData.constituencies[acNo];
  
  // Find unique parties of agents in this constituency
  const parties = new Set();
  acData.agents.forEach(ag => {
    parties.add(ag.party);
  });
  
  const dropdown = document.getElementById('filter-agent-party');
  dropdown.innerHTML = '<option value="all">All Parties</option>';
  
  parties.forEach(p => {
    const opt = document.createElement('option');
    opt.value = p;
    opt.innerText = p;
    dropdown.appendChild(opt);
  });
  
  // Restore filter selection if valid
  if (parties.has(state.agentFilters.party)) {
    dropdown.value = state.agentFilters.party;
  } else {
    state.agentFilters.party = 'all';
    dropdown.value = 'all';
  }
}

function renderAgentPerformanceSection() {
  const f = state.agentFilters;
  const ac = state.electionData.constituencies[f.ac];
  
  // Filter Agents
  let list = [...ac.agents];
  
  if (f.party !== 'all') {
    list = list.filter(ag => ag.party === f.party);
  }
  
  if (f.search.trim()) {
    const q = f.search.toLowerCase().trim();
    list = list.filter(ag => 
      ag.agent_name.toLowerCase().includes(q) || 
      ag.mobile_no.includes(q) ||
      ag.area.toLowerCase().includes(q) ||
      ag.booth_no.toString() === q
    );
  }

  if (f.rating !== 'all') {
    if (f.rating === 'top') {
      list = list.filter(ag => ag.effectiveness_score >= 75);
    } else if (f.rating === 'low') {
      list = list.filter(ag => ag.effectiveness_score < 45);
    } else if (f.rating === 'high-turn-low-vote') {
      list = list.filter(ag => ag.booth_management_score >= 80 && ag.vote_conversion_score < 25);
    } else if (f.rating === 'low-turn-high-support') {
      list = list.filter(ag => ag.booth_management_score < 72 && ag.vote_conversion_score >= 45);
    }
  }

  // Populate side tables (Top, Underperforming, Low Conversion)
  const partyFilteredList = f.party !== 'all' ? [...ac.agents].filter(ag => ag.party === f.party) : [...ac.agents];
  
  // 1. Top Performers
  const topBody = document.querySelector('#agent-top-performers tbody');
  topBody.innerHTML = '';
  const topAgents = [...partyFilteredList].sort((a,b) => b.effectiveness_score - a.effectiveness_score).slice(0, 5);
  topAgents.forEach(ag => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${ag.agent_name}</strong><br><span style="font-size:0.75rem; color:var(--text-muted);">${ag.party.split('-')[0]}</span></td>
      <td style="max-width:120px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${ag.booth_name}">Booth ${ag.booth_no}</td>
      <td><span style="color:#10b981; font-weight:700;">${ag.effectiveness_score}</span></td>
    `;
    topBody.appendChild(tr);
  });

  // 2. Underperformers
  const lowBody = document.querySelector('#agent-under-performers tbody');
  lowBody.innerHTML = '';
  const lowAgents = [...partyFilteredList].sort((a,b) => a.effectiveness_score - b.effectiveness_score).slice(0, 5);
  lowAgents.forEach(ag => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${ag.agent_name}</strong><br><span style="font-size:0.75rem; color:var(--text-muted);">${ag.party.split('-')[0]}</span></td>
      <td style="max-width:120px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${ag.booth_name}">Booth ${ag.booth_no}</td>
      <td><span style="color:#ef4444; font-weight:700;">${ag.effectiveness_score}</span></td>
    `;
    lowBody.appendChild(tr);
  });

  // 3. Low Conversion Zones (Turnout > 80, support < 25)
  const lowConvBody = document.querySelector('#agent-low-conversion tbody');
  lowConvBody.innerHTML = '';
  const lowConvAgents = [...partyFilteredList]
    .filter(ag => ag.booth_management_score >= 80 && ag.vote_conversion_score < 25)
    .sort((a,b) => a.vote_conversion_score - b.vote_conversion_score)
    .slice(0, 5);
    
  if (lowConvAgents.length === 0) {
    lowConvBody.innerHTML = '<tr><td colspan="3" style="text-align:center; color:var(--text-muted); padding:15px;">No critical conversion alerts found.</td></tr>';
  } else {
    lowConvAgents.forEach(ag => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong>${ag.agent_name}</strong><br><span style="font-size:0.75rem; color:var(--text-muted);">${ag.party.split('-')[0]}</span></td>
        <td style="max-width:120px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${ag.booth_name}">Booth ${ag.booth_no}</td>
        <td><strong>${ag.booth_management_score}% / <span style="color:#ef4444;">${ag.vote_conversion_score}%</span></strong></td>
      `;
      lowConvBody.appendChild(tr);
    });
  }

  // Populate Master Table
  const totalCount = list.length;
  const totalPages = Math.ceil(totalCount / f.pageSize) || 1;
  if (f.page > totalPages) f.page = totalPages;
  
  const startIndex = (f.page - 1) * f.pageSize;
  const paginatedList = list.slice(startIndex, startIndex + f.pageSize);
  
  const tbody = document.querySelector('#agent-performance-table tbody');
  tbody.innerHTML = '';
  
  if (paginatedList.length === 0) {
    tbody.innerHTML = `<tr><td colspan="12" style="text-align:center; padding:30px; color:var(--text-muted);">No agents match the specified filters.</td></tr>`;
  } else {
    paginatedList.forEach(ag => {
      const tr = document.createElement('tr');
      const partyShort = ag.party.split('-')[0];
      const effectColor = ag.effectiveness_score >= 70 ? '#10b981' : ag.effectiveness_score < 45 ? '#ef4444' : '#f59e0b';
      
      tr.innerHTML = `
        <td><strong>${ag.rank}</strong></td>
        <td><strong>#${ag.booth_no}</strong></td>
        <td style="max-width: 180px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${ag.booth_name}">${ag.booth_name}</td>
        <td><strong>${ag.agent_name}</strong></td>
        <td><span style="font-family: monospace;">${maskPhoneNumber(ag.mobile_no)}</span></td>
        <td><span class="badge badge-${partyShort.toLowerCase()}">${partyShort}</span></td>
        <td>${ag.candidate_votes}</td>
        <td>${ag.opponent_votes}</td>
        <td style="font-weight:700; color:${ag.margin >= 0 ? '#10b981':'#ef4444'}">${ag.margin >= 0 ? '+' : ''}${ag.margin}</td>
        <td><strong>${ag.vote_conversion_score}%</strong></td>
        <td><strong>${ag.booth_management_score}%</strong></td>
        <td><strong style="color: ${effectColor}; font-size:1rem;">${ag.effectiveness_score}</strong></td>
      `;
      tbody.appendChild(tr);
    });
  }

  // Pagination for agents
  document.getElementById('agent-pagination-info').innerText = `Showing ${totalCount === 0 ? 0 : startIndex + 1} to ${Math.min(startIndex + f.pageSize, totalCount)} of ${totalCount} agents`;
  
  const pagControls = document.getElementById('agent-pagination-controls');
  pagControls.innerHTML = '';
  
  const prevBtn = document.createElement('button');
  prevBtn.className = 'page-btn';
  prevBtn.innerHTML = '<i class="fa-solid fa-chevron-left"></i>';
  prevBtn.disabled = f.page === 1;
  prevBtn.addEventListener('click', () => {
    f.page--;
    renderAgentPerformanceSection();
  });
  pagControls.appendChild(prevBtn);

  let startPage = Math.max(1, f.page - 2);
  let endPage = Math.min(totalPages, startPage + 4);
  if (endPage - startPage < 4) {
    startPage = Math.max(1, endPage - 4);
  }

  for (let p = startPage; p <= endPage; p++) {
    const pBtn = document.createElement('button');
    pBtn.className = `page-btn ${f.page === p ? 'active' : ''}`;
    pBtn.innerText = p;
    pBtn.addEventListener('click', () => {
      f.page = p;
      renderAgentPerformanceSection();
    });
    pagControls.appendChild(pBtn);
  }

  const nextBtn = document.createElement('button');
  nextBtn.className = 'page-btn';
  nextBtn.innerHTML = '<i class="fa-solid fa-chevron-right"></i>';
  nextBtn.disabled = f.page === totalPages;
  nextBtn.addEventListener('click', () => {
    f.page++;
    renderAgentPerformanceSection();
  });
  pagControls.appendChild(nextBtn);
}

// F. AI INSIGHTS & STRATEGIC RECOMMENDATIONS
function renderAIInsightsSection() {
  const acNo = state.aiSettings.ac;
  const ac = state.electionData.constituencies[acNo];
  
  // 1. Narrative strategic summary
  const summaryEl = document.getElementById('ai-narrative-summary');
  
  let narrative = `In **Assembly Constituency ${ac.ac_name} (${ac.ac_no})**, the election concluded with a victory for **${ac.winning_candidate.split('(')[0].trim()}** representing **${ac.winning_party}**, securing a final lead of **${ac.victory_margin.toLocaleString()}** votes. 
  Our algorithmic targeting model evaluated all **${ac.total_booths}** booths in this constituency and extracted critical operational intelligence:
  <br><br>
  - **Electoral Dominance**: ${ac.winning_party} captured a majority of the booths, winning in **${ac.booths.filter(b => b.winner_party === ac.winning_party).length}** polling stations. 
  - **Marginal Swing Zones**: We identified **${ac.booths.filter(b => b.margin < 50).length}** booths where the margin of victory or defeat was **under 50 votes**. In the next campaign, shifting just **25 votes** in these key locations will swing the booth outcome.
  - **Turnout Disconnects**: There are **${ac.booths.filter(b => b.turnout_pct >= 80 && (b.candidate_votes[ac.winning_candidate]/b.total_votes * 100) < 30).length}** booths with exceptionally high turnout (>80%) but below-average conversion for the leading alliance, indicating strong administrative turnout but weak ideological conversion. Priority focus should be placed on these agent areas.`;
  
  summaryEl.innerHTML = narrative;

  // 2. High Priority Swing Table (<100 votes loss)
  const swingBody = document.querySelector('#ai-swing-booths-table tbody');
  swingBody.innerHTML = '';
  
  const constituencyWinner = ac.winning_candidate;
  const winnerParty = ac.winning_party;
  
  // Find booths lost by constituency winner by less than 100 votes
  const swingBooths = ac.booths
    .filter(b => b.winner_party !== winnerParty && b.margin < 100)
    .sort((a,b) => a.margin - b.margin)
    .slice(0, 10);
    
  if (swingBooths.length === 0) {
    swingBody.innerHTML = '<tr><td colspan="5" style="text-align:center; color:var(--text-muted); padding:15px;">No narrow defeats (<100 votes) recorded. Excellent stronghold maintenance.</td></tr>';
  } else {
    swingBooths.forEach(b => {
      const tr = document.createElement('tr');
      const winnerPartyShort = b.winner_party;
      tr.innerHTML = `
        <td><strong>#${b.booth_no}</strong></td>
        <td style="max-width:150px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${b.booth_name}">${b.booth_name}</td>
        <td>${b.candidate_votes[constituencyWinner] || 0} <span style="font-size:0.75rem; color:var(--text-muted);">vs ${b.candidate_votes[b.winner]} (${winnerPartyShort})</span></td>
        <td><span style="color:#ef4444; font-weight:700;">-${b.margin} votes</span></td>
        <td><strong>${b.turnout_pct}%</strong></td>
      `;
      swingBody.appendChild(tr);
    });
  }

  // 3. High Turnout / Low Conversion Table (Turnout > 78, winning party votes share < 30%)
  const turnBody = document.querySelector('#ai-high-turn-low-conv-table tbody');
  turnBody.innerHTML = '';
  
  const highTurnLowConv = ac.booths
    .filter(b => b.turnout_pct >= 78 && (b.candidate_votes[constituencyWinner]/b.total_votes * 100) < 30)
    .sort((a,b) => (a.candidate_votes[constituencyWinner]/a.total_votes) - (b.candidate_votes[constituencyWinner]/b.total_votes))
    .slice(0, 10);
    
  if (highTurnLowConv.length === 0) {
    turnBody.innerHTML = '<tr><td colspan="5" style="text-align:center; color:var(--text-muted); padding:15px;">No turnout-conversion anomalies detected.</td></tr>';
  } else {
    highTurnLowConv.forEach(b => {
      const tr = document.createElement('tr');
      const voteShare = ((b.candidate_votes[constituencyWinner]/b.total_votes) * 100).toFixed(1);
      
      // Find BLA-2 agent name for constituency winner's party
      const agentObj = ac.agents.find(ag => ag.booth_no === b.booth_no && ag.party.includes(winnerParty));
      const agentName = agentObj ? agentObj.agent_name : 'No Agent Assigned';
      
      tr.innerHTML = `
        <td><strong>#${b.booth_no}</strong></td>
        <td style="max-width:150px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${b.booth_name}">${b.booth_name}</td>
        <td><strong>${b.turnout_pct}%</strong></td>
        <td><span style="color:#ef4444; font-weight:700;">${voteShare}%</span></td>
        <td>${agentName}</td>
      `;
      turnBody.appendChild(tr);
    });
  }

  // 4. Actionable political recommendations
  const recList = document.getElementById('ai-recommendations-list');
  recList.innerHTML = `
    <li class="insight-list-item">
      <strong>Agent Realignment in Low Conversion Zones</strong>: Initiate audit of BLA-2 agents in booths with >80% turnout but <25% vote conversion. Re-train or replace underperforming booth coordinators.
    </li>
    <li class="insight-list-item">
      <strong>Marginal Swing Focus Group</strong>: Target the ${ac.booths.filter(b => b.margin < 50).length} close-contest booths (<50 margin) with targeted localized beneficiary schemes and door-to-door micro-campaigning.
    </li>
    <li class="insight-list-item">
      <strong>Stronghold Defense Strategy</strong>: Maintain current levels in the top 20 stronghold areas, ensuring BLA agents receive recognition to keep voter retention high.
    </li>
    <li class="insight-list-item">
      <strong>Mobile Contact Campaign</strong>: Leverage BLA mobile numbers to build a localized WhatsApp broadcast matrix. Directly push candidate video testimonies to electors in swing zones.
    </li>
  `;
}

// ==============================================
// DRILL-DOWN MODAL WINDOW
// ==============================================
function showBoothDetails(booth, acNo) {
  const overlay = document.getElementById('detail-modal');
  const title = document.getElementById('modal-title');
  const body = document.getElementById('modal-body');
  
  title.innerText = `Detailed View: Polling Station No. ${booth.booth_no}`;
  overlay.classList.add('active');
  
  // Find all agents in this booth
  const ac = state.electionData.constituencies[acNo];
  const boothAgents = ac.agents.filter(ag => ag.booth_no === booth.booth_no);
  
  let agentsHtml = '';
  if (boothAgents.length === 0) {
    agentsHtml = '<p style="color:var(--text-muted); font-style:italic;">No BLA-2 agents registered for this booth in the system.</p>';
  } else {
    agentsHtml = `
      <div class="table-wrapper" style="margin-top:12px;">
        <table class="custom-table" style="font-size:0.85rem;">
          <thead>
            <tr>
              <th>Agent Name</th>
              <th>Party</th>
              <th>Mobile No</th>
              <th>Score</th>
            </tr>
          </thead>
          <tbody>
            ${boothAgents.map(ag => `
              <tr>
                <td><strong>${ag.agent_name}</strong></td>
                <td><span class="badge badge-${ag.party.split('-')[0].toLowerCase()}">${ag.party.split('-')[0]}</span></td>
                <td><span style="font-family:monospace;">${maskPhoneNumber(ag.mobile_no)}</span></td>
                <td><strong style="color: ${ag.effectiveness_score >= 70 ? '#10b981' : ag.effectiveness_score < 45 ? '#ef4444' : '#f59e0b'}">${ag.effectiveness_score}</strong></td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  // Generate candidate votes rows
  const sortedCandidates = Object.entries(booth.candidate_votes).sort((a,b) => b[1] - a[1]);
  const candidatesHtml = sortedCandidates.map(([candName, votes]) => {
    const pct = booth.total_votes > 0 ? ((votes / booth.total_votes) * 100).toFixed(1) : 0;
    const pParts = candName.split('(');
    const p = pParts.length > 1 ? pParts[pParts.length - 1].replace(')', '') : 'Others';
    const color = partyColors[p] || '#6b7280';
    return `
      <div style="margin-bottom: 12px;">
        <div style="display:flex; justify-content:space-between; font-size:0.85rem; margin-bottom:4px;">
          <span><strong>${candName.split('(')[0].trim()}</strong> <span style="color:var(--text-muted);">(${p})</span></span>
          <span><strong>${votes.toLocaleString()} votes</strong> (${pct}%)</span>
        </div>
        <div class="strength-bar-container">
          <div class="strength-bar" style="width:${pct}%; background-color:${color};"></div>
        </div>
      </div>
    `;
  }).join('');

  body.innerHTML = `
    <div style="display:grid; grid-template-columns:1fr 1fr; gap:20px; margin-bottom:24px;">
      <div>
        <p style="color:var(--text-muted); font-size:0.8rem; text-transform:uppercase; font-weight:600;">Building/Location</p>
        <p style="font-weight:600; font-size:1rem; margin-top:4px;">${booth.booth_name}</p>
        
        <p style="color:var(--text-muted); font-size:0.8rem; text-transform:uppercase; font-weight:600; margin-top:16px;">Assembly Constituency</p>
        <p style="font-weight:600; font-size:1rem; margin-top:4px;">AC ${acNo} ${ac.ac_name}</p>
      </div>
      <div>
        <p style="color:var(--text-muted); font-size:0.8rem; text-transform:uppercase; font-weight:600;">Booth Statistics</p>
        <p style="margin-top:6px;">Total Electors: <strong>${booth.electors.toLocaleString()}</strong></p>
        <p style="margin-top:4px;">Total Votes Polled: <strong>${booth.total_votes.toLocaleString()}</strong></p>
        <p style="margin-top:4px;">Turnout Percentage: <strong style="color:#6366f1;">${booth.turnout_pct}%</strong></p>
        <p style="margin-top:4px;">Winner: <span class="badge badge-${booth.winner_party.toLowerCase()}" style="margin-top:2px;">${booth.winner_party}</span> (margin +${booth.margin})</p>
      </div>
    </div>
    
    <div class="menu-divider" style="margin:20px 0;"></div>
    
    <div class="card-title" style="font-size:1.05rem; margin-bottom:12px;">Candidate Breakdown</div>
    <div style="margin-bottom:24px;">
      ${candidatesHtml}
    </div>
    
    <div class="menu-divider" style="margin:20px 0;"></div>
    
    <div class="card-title" style="font-size:1.05rem; margin-bottom:4px;">Registered Booth Agents (BLA-2)</div>
    ${agentsHtml}
  `;
}

function hideModal() {
  document.getElementById('detail-modal').classList.remove('active');
}

// ==============================================
// EXPORT INTEGRATIONS
// ==============================================

// Helper to sanitize filename strings
function getExportFilename(base) {
  return `${base}_${new Date().toISOString().slice(0, 10)}`;
}

// 1. Export Booth Ledger Table to CSV or Excel
function exportBoothTable(format) {
  const f = state.boothFilters;
  
  // Assemble booths across matching constituencies
  let boothsList = [];
  Object.keys(state.electionData.constituencies).forEach(acKey => {
    const acNo = parseInt(acKey);
    if (f.ac === 'all' || f.ac == acNo) {
      state.electionData.constituencies[acKey].booths.forEach(b => {
        boothsList.push({
          ...b,
          ac_no: acNo,
          ac_name: state.electionData.constituencies[acKey].ac_name
        });
      });
    }
  });

  // Apply filters
  if (f.search.trim()) {
    const q = f.search.toLowerCase().trim();
    boothsList = boothsList.filter(b => 
      b.booth_name.toLowerCase().includes(q) || 
      b.booth_no.toString() === q
    );
  }
  if (f.winner !== 'all') {
    boothsList = boothsList.filter(b => b.winner_party === f.winner);
  }
  if (f.margin !== 'all') {
    if (f.margin === 'high') boothsList = boothsList.filter(b => b.margin >= 200);
    else if (f.margin === 'mid') boothsList = boothsList.filter(b => b.margin >= 50 && b.margin < 200);
    else if (f.margin === 'close') boothsList = boothsList.filter(b => b.margin < 50);
  }
  if (f.turnout !== 'all') {
    if (f.turnout === 'high') boothsList = boothsList.filter(b => b.turnout_pct >= 80);
    else if (f.turnout === 'low') boothsList = boothsList.filter(b => b.turnout_pct < 72);
  }

  // Create columns based on selected AC
  const headers = ['Booth No', 'Booth Name', 'Constituency', 'Total Electors', 'Total Votes Polled', 'Turnout %'];
  
  // Dynamic columns for candidates
  let candNames = [];
  if (f.ac !== 'all') {
    const acNo = parseInt(f.ac);
    candNames = state.electionData.constituencies[acNo].candidates.map(c => c.name);
  } else {
    // Collect all unique candidates in the current list
    const candSet = new Set();
    boothsList.forEach(b => {
      Object.keys(b.candidate_votes).forEach(k => candSet.add(k));
    });
    candNames = Array.from(candSet);
  }
  
  headers.push(...candNames);
  headers.push('Winner in Booth', 'Winning Margin');

  // Map rows
  const rows = boothsList.map(b => {
    const row = [
      b.booth_no,
      b.booth_name,
      b.ac_name,
      b.electors,
      b.total_votes,
      b.turnout_pct
    ];
    
    candNames.forEach(c => {
      row.push(b.candidate_votes[c] || 0);
    });
    
    row.push(b.winner, b.margin);
    return row;
  });

  const fileTitle = getExportFilename(`Booth_Ledger_AC_${f.ac}`);
  
  if (format === 'csv') {
    // Generate CSV
    let csvContent = "data:text/csv;charset=utf-8,";
    csvContent += headers.map(h => `"${h.replace(/"/g, '""')}"`).join(",") + "\n";
    
    rows.forEach(r => {
      csvContent += r.map(val => {
        if (typeof val === 'string') return `"${val.replace(/"/g, '""')}"`;
        return val;
      }).join(",") + "\n";
    });
    
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `${fileTitle}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  } else {
    // Generate XLSX via SheetJS
    const wb = XLSX.utils.book_new();
    const dataMatrix = [headers, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(dataMatrix);
    XLSX.utils.book_append_sheet(wb, ws, "Booth Results");
    XLSX.writeFile(wb, `${fileTitle}.xlsx`);
  }
}

// 2. Export Agent Performance Table to CSV
function exportAgentTable() {
  const f = state.agentFilters;
  const ac = state.electionData.constituencies[f.ac];
  
  let list = [...ac.agents];
  if (f.party !== 'all') {
    list = list.filter(ag => ag.party === f.party);
  }
  if (f.search.trim()) {
    const q = f.search.toLowerCase().trim();
    list = list.filter(ag => 
      ag.agent_name.toLowerCase().includes(q) || 
      ag.mobile_no.includes(q) ||
      ag.area.toLowerCase().includes(q)
    );
  }
  if (f.rating !== 'all') {
    if (f.rating === 'top') list = list.filter(ag => ag.effectiveness_score >= 75);
    else if (f.rating === 'low') list = list.filter(ag => ag.effectiveness_score < 45);
    else if (f.rating === 'high-turn-low-vote') list = list.filter(ag => ag.booth_management_score >= 80 && ag.vote_conversion_score < 25);
    else if (f.rating === 'low-turn-high-support') list = list.filter(ag => ag.booth_management_score < 72 && ag.vote_conversion_score >= 45);
  }

  const headers = ['Rank', 'Booth No', 'Booth Name', 'Agent Name', 'Mobile No', 'Party', 'Candidate Votes', 'Opponent Votes', 'Margin', 'Vote Conversion %', 'Turnout %', 'Effectiveness Score'];
  const rows = list.map(ag => [
    ag.rank,
    ag.booth_no,
    ag.booth_name,
    ag.agent_name,
    maskPhoneNumber(ag.mobile_no),
    ag.party,
    ag.candidate_votes,
    ag.opponent_votes,
    ag.margin,
    ag.vote_conversion_score,
    ag.booth_management_score,
    ag.effectiveness_score
  ]);

  const fileTitle = getExportFilename(`Agent_Performance_AC_${f.ac}`);
  
  let csvContent = "data:text/csv;charset=utf-8,";
  csvContent += headers.map(h => `"${h.replace(/"/g, '""')}"`).join(",") + "\n";
  
  rows.forEach(r => {
    csvContent += r.map(val => {
      if (typeof val === 'string') return `"${val.replace(/"/g, '""')}"`;
      return val;
    }).join(",") + "\n";
  });
  
  const encodedUri = encodeURI(csvContent);
  const link = document.createElement("a");
  link.setAttribute("href", encodedUri);
  link.setAttribute("download", `${fileTitle}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// 3. Export PDF Report
function exportDashboardPDF() {
  const loader = document.getElementById('loading-overlay');
  const loaderText = document.getElementById('loading-text');
  
  loaderText.innerText = "Generating styled PDF report...";
  loader.classList.add('active');
  
  // Set temporary styling configuration optimized for printing
  const element = document.getElementById('pdf-root');
  
  const opt = {
    margin:       10,
    filename:     `${getExportFilename(state.activePage)}.pdf`,
    image:        { type: 'jpeg', quality: 0.98 },
    html2canvas:  { scale: 2, useCORS: true },
    jsPDF:        { unit: 'mm', format: 'a4', orientation: 'landscape' }
  };
  
  // Trigger html2pdf export
  html2pdf().set(opt).from(element).save().then(() => {
    loader.classList.remove('active');
  }).catch(err => {
    console.error("PDF generation failed:", err);
    loader.classList.remove('active');
    alert("Failed to export PDF: " + err.message);
  });
}


