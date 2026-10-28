/**
 * Dead Stock & Hardware Inventory Register - Application Logic
 * Supports Bilingual (Marathi/English) Inline Equipment Registration, Course Chooser,
 * Clickable KPIs, Live Filters, Cloud Drive & Excel Attachment Management (Add & Delete)
 */

// Application State
let assets = [];
let departments = [];
let activeFilters = {
  search: '',
  department: 'ALL',
  type: 'Computer', // default to Computer matching the primary tab
  status: 'ALL',
  hasExcel: false
};
let currentViewMode = 'table'; // 'table' or 'cards'
let chartInstance = null;
let chartType = 'doughnut';
let currentExcelTargetAsset = null;
let currentQuickAsset = null;

// Top / Popular College Courses for quick chooser bar
const TOP_COURSES = [
  'ALL',
  'B.Sc IT',
  'BCS',
  'BSc CS',
  'MSc CS',
  'Jr. IT Computer Lab',
  'Office',
  'Mathematics',
  'Language Lab',
  'Library',
  'Statistics',
  'Electronics',
  'PG Lab (Physics & Chemistry)',
  'Geography',
  'Botany',
  'Zoology',
  'Chemistry BSc',
  'Chemistry MSc',
  'ERP',
  'NAAC'
];

// DOM Element References
const kpiTotalCount = document.getElementById('kpiTotalCount');
const kpiTotalWorking = document.getElementById('kpiTotalWorking');
const kpiTotalFaulty = document.getElementById('kpiTotalFaulty');
const kpiProgressWorking = document.getElementById('kpiProgressWorking');
const kpiProgressFaulty = document.getElementById('kpiProgressFaulty');

const kpiPcCount = document.getElementById('kpiPcCount');
const kpiPcWorking = document.getElementById('kpiPcWorking');
const kpiPcFaulty = document.getElementById('kpiPcFaulty');
const kpiPcHealth = document.getElementById('kpiPcHealth');

const kpiPrinterCount = document.getElementById('kpiPrinterCount');
const kpiPrinterWorking = document.getElementById('kpiPrinterWorking');
const kpiPrinterFaulty = document.getElementById('kpiPrinterFaulty');
const kpiPrinterHealth = document.getElementById('kpiPrinterHealth');

const kpiProjectorCount = document.getElementById('kpiProjectorCount');
const kpiProjectorWorking = document.getElementById('kpiProjectorWorking');
const kpiProjectorFaulty = document.getElementById('kpiProjectorFaulty');
const kpiProjectorHealth = document.getElementById('kpiProjectorHealth');

const kpiExcelCount = document.getElementById('kpiExcelCount');
const kpiPcsWithExcel = document.getElementById('kpiPcsWithExcel');

// Navigation Tab Badges
const tabPcBadge = document.getElementById('tabPcBadge');
const tabPrinterBadge = document.getElementById('tabPrinterBadge');
const tabProjectorBadge = document.getElementById('tabProjectorBadge');

// Inline Add Equipment Panel
const inlineAddEquipmentPanel = document.getElementById('inlineAddEquipmentPanel');
const inlineAssetType = document.getElementById('inlineAssetType');
const inlinePcForm = document.getElementById('inlinePcForm');
const inlinePrinterForm = document.getElementById('inlinePrinterForm');
const inlineProjectorForm = document.getElementById('inlineProjectorForm');

const quickCourseChips = document.getElementById('quickCourseChips');
const courseFilterSearch = document.getElementById('courseFilterSearch');
const courseChooserCount = document.getElementById('courseChooserCount');
const departmentMatrixContainer = document.getElementById('departmentMatrixContainer');
const activeFilterBadge = document.getElementById('activeFilterBadge');
const deptSelectFilter = document.getElementById('deptSelectFilter');
const searchInput = document.getElementById('searchInput');
const btnResetFilters = document.getElementById('btnResetFilters');
const showingAssetsCount = document.getElementById('showingAssetsCount');

const activeFilterBanner = document.getElementById('activeFilterBanner');
const bannerText = document.getElementById('bannerText');
const btnClearActiveBanner = document.getElementById('btnClearActiveBanner');

const tableContainer = document.getElementById('tableContainer');
const assetTableBody = document.getElementById('assetTableBody');
const assetCardsContainer = document.getElementById('assetCardsContainer');

// Initialize Application
document.addEventListener('DOMContentLoaded', () => {
  loadData();
  setupEventListeners();
  renderAll();
});

// Load Data from LocalStorage (Starts at ZERO by default)
const STORAGE_KEY_ASSETS = 'willingdon_assets_clean_v1';
const STORAGE_KEY_DEPTS = 'willingdon_depts_clean_v1';
const STORAGE_KEY_CLEARED = 'willingdon_cleared_clean_v1';

function loadData() {
  // Clear any legacy test data
  localStorage.removeItem('assetpulse_assets_v4');
  localStorage.removeItem('assetpulse_depts_v4');
  localStorage.removeItem('assetpulse_assets_v3');
  localStorage.removeItem('willingdon_auth_user_v1');

  const isCleared = localStorage.getItem(STORAGE_KEY_CLEARED);
  if (isCleared === 'true') {
    assets = [];
    departments = [];
    return;
  }

  const storedAssets = localStorage.getItem(STORAGE_KEY_ASSETS);
  const storedDepts = localStorage.getItem(STORAGE_KEY_DEPTS);

  if (storedAssets !== null && storedDepts !== null) {
    try {
      assets = JSON.parse(storedAssets);
      departments = JSON.parse(storedDepts);
    } catch (e) {
      assets = [];
      departments = [];
    }
  } else {
    // When entering website, default to ZERO data
    assets = [];
    departments = [];
    localStorage.setItem(STORAGE_KEY_CLEARED, 'true');
    saveData();
  }
}

function useInitialData() {
  assets = [];
  departments = [];
  localStorage.setItem(STORAGE_KEY_CLEARED, 'true');
  saveData();
}

function saveData() {
  if (assets.length === 0) {
    localStorage.setItem(STORAGE_KEY_CLEARED, 'true');
  } else {
    localStorage.removeItem(STORAGE_KEY_CLEARED);
  }
  localStorage.setItem(STORAGE_KEY_ASSETS, JSON.stringify(assets));
  localStorage.setItem(STORAGE_KEY_DEPTS, JSON.stringify(departments));
}

// Complete Erase: Resets everything to ZERO
function eraseAllInventoryData() {
  const count = assets.length;
  const currentFile = localStorage.getItem('assetpulse_active_excel_filename') || 'Revised_Inventory_wdn_updated_05-05-2026.xlsx';
  const msg = `⚠️ पुष्टीकरण (CONFIRMATION):\n\nतुम्हाला जोडलेली एक्सेल फाईल "${currentFile}" आणि इन्व्हेंटरीमधील सर्व (${count}) उपकरणांचा डेटा कायमचा नष्ट करायचा आहे का?\n\n- संगणक (Computers): 0\n- प्रिंटर (Printers): 0\n- प्रोजेक्टर (Projectors): 0\n- एकूण संख्या (Total): 0\n\n(Are you sure you want to delete "${currentFile}" and completely erase ALL ${count} items? Everything will be reset to ZERO).`;

  if (confirm(msg)) {
    assets = [];
    departments = [];
    localStorage.removeItem('assetpulse_active_excel_filename');
    localStorage.setItem(STORAGE_KEY_CLEARED, 'true');
    saveData();
    renderAll();
    renderGlobalExcelTable();
    updateImportModalActiveFile();
    showToast('🗑️ एक्सेल फाईल व सर्व डेटा पूर्णपणे नष्ट केला - संख्या 0 झाली (All data erased - Count is 0)');
  }
}

// Restore default sample data (580 items)
function restoreDefaultData() {
  if (confirm('तुम्हाला ५८० उपकरणांचा मूळ नमुना डेटा पूर्ववत करायचा आहे का?\n(Restore original default 580 items?)')) {
    localStorage.removeItem(STORAGE_KEY_CLEARED);
    localStorage.setItem('assetpulse_active_excel_filename', 'Revised_Inventory_wdn_updated_05-05-2026.xlsx');
    useInitialData();
    renderAll();
    renderGlobalExcelTable();
    updateImportModalActiveFile();
    showToast('🔄 मूळ नमुना डेटा (580 उपकरणे) पूर्ववत करण्यात आला (Default data restored)');
  }
}

// Master Render
function renderAll() {
  renderKpis();
  renderCourseChips();
  renderDepartmentMatrix();
  renderDeptDropdowns();
  renderChart();
  renderInventoryList();
  updateBanner();
  updateImportModalActiveFile();
}

// Render Top KPIs & Tab Badges
function renderKpis() {
  const total = assets.length;
  const working = assets.filter(a => a.status === 'Working').length;
  const faulty = total - working;

  const pcs = assets.filter(a => a.type === 'Computer');
  const pcWorking = pcs.filter(a => a.status === 'Working').length;
  const pcFaulty = pcs.length - pcWorking;

  const printers = assets.filter(a => a.type === 'Printer');
  const prnWorking = printers.filter(a => a.status === 'Working').length;
  const prnFaulty = printers.length - prnWorking;

  const projectors = assets.filter(a => a.type === 'Projector');
  const prjWorking = projectors.filter(a => a.status === 'Working').length;
  const prjFaulty = projectors.length - prjWorking;

  let totalExcelFiles = 0;
  let pcsWithExcelCount = 0;
  assets.forEach(a => {
    if (a.excelAttachments && a.excelAttachments.length > 0) {
      totalExcelFiles += a.excelAttachments.length;
      pcsWithExcelCount++;
    }
  });

  if (kpiTotalCount) kpiTotalCount.textContent = total;
  if (kpiTotalWorking) kpiTotalWorking.textContent = working;
  if (kpiTotalFaulty) kpiTotalFaulty.textContent = faulty;
  
  const pctWorking = total > 0 ? (working / total) * 100 : 0;
  const pctFaulty = total > 0 ? (faulty / total) * 100 : 0;
  if (kpiProgressWorking) kpiProgressWorking.style.width = `${pctWorking}%`;
  if (kpiProgressFaulty) kpiProgressFaulty.style.width = `${pctFaulty}%`;

  if (kpiPcCount) kpiPcCount.textContent = pcs.length;
  if (kpiPcWorking) kpiPcWorking.textContent = pcWorking;
  if (kpiPcFaulty) kpiPcFaulty.textContent = pcFaulty;
  if (kpiPcHealth) kpiPcHealth.textContent = pcs.length > 0 ? `${Math.round((pcWorking / pcs.length) * 100)}% operational` : '0%';

  if (kpiPrinterCount) kpiPrinterCount.textContent = printers.length;
  if (kpiPrinterWorking) kpiPrinterWorking.textContent = prnWorking;
  if (kpiPrinterFaulty) kpiPrinterFaulty.textContent = prnFaulty;
  if (kpiPrinterHealth) kpiPrinterHealth.textContent = printers.length > 0 ? `${Math.round((prnWorking / printers.length) * 100)}% operational` : '0%';

  if (kpiProjectorCount) kpiProjectorCount.textContent = projectors.length;
  if (kpiProjectorWorking) kpiProjectorWorking.textContent = prjWorking;
  if (kpiProjectorFaulty) kpiProjectorFaulty.textContent = prjFaulty;
  if (kpiProjectorHealth) kpiProjectorHealth.textContent = projectors.length > 0 ? `${Math.round((prjWorking / projectors.length) * 100)}% operational` : '0%';

  if (kpiExcelCount) kpiExcelCount.textContent = totalExcelFiles;
  if (kpiPcsWithExcel) kpiPcsWithExcel.textContent = pcsWithExcelCount;

  // Update primary nav tab badges
  if (tabPcBadge) tabPcBadge.textContent = pcs.length;
  if (tabPrinterBadge) tabPrinterBadge.textContent = printers.length;
  if (tabProjectorBadge) tabProjectorBadge.textContent = projectors.length;

  // Update category filter pill buttons
  const pillAll = document.querySelector('.pill-btn[data-filter-type="ALL"]');
  const pillComp = document.querySelector('.pill-btn[data-filter-type="Computer"]');
  const pillPrn = document.querySelector('.pill-btn[data-filter-type="Printer"]');
  const pillPrj = document.querySelector('.pill-btn[data-filter-type="Projector"]');

  if (pillAll) pillAll.textContent = `All Equipment (${total})`;
  if (pillComp) pillComp.textContent = `🖥️ Computers (${pcs.length})`;
  if (pillPrn) pillPrn.textContent = `🖨️ Printers (${printers.length})`;
  if (pillPrj) pillPrj.textContent = `📽️ Projectors (${projectors.length})`;

  // Highlight active primary nav tab
  document.querySelectorAll('.college-nav-tabs .nav-tab-btn[data-nav-tab]').forEach(tab => {
    const tabName = tab.dataset.navTab;
    if (
      (tabName === 'computers' && activeFilters.type === 'Computer' && inlineAddEquipmentPanel.style.display === 'none') ||
      (tabName === 'printers' && activeFilters.type === 'Printer' && inlineAddEquipmentPanel.style.display === 'none') ||
      (tabName === 'projectors' && activeFilters.type === 'Projector' && inlineAddEquipmentPanel.style.display === 'none')
    ) {
      tab.classList.add('active');
    } else if (tabName !== 'add-item' && tabName !== 'excel-import') {
      tab.classList.remove('active');
    }
  });
}

// Render Interactive Course / Department Chooser Bar
function renderCourseChips() {
  if (!quickCourseChips) return;
  quickCourseChips.innerHTML = '';
  const search = courseFilterSearch ? courseFilterSearch.value.trim().toLowerCase() : '';

  let displayCourses = [...TOP_COURSES];
  if (search) {
    displayCourses = departments.filter(d => d.toLowerCase().includes(search));
    if (!displayCourses.includes('ALL')) displayCourses.unshift('ALL');
  }

  displayCourses.forEach(course => {
    let count = 0;
    let label = course;
    let icon = '🏫';

    if (course === 'ALL') {
      count = activeFilters.type === 'ALL'
        ? assets.length
        : assets.filter(a => a.type === activeFilters.type).length;
      label = `🎓 All Courses (${count})`;
    } else {
      count = activeFilters.type === 'ALL'
        ? assets.filter(a => a.department === course).length
        : assets.filter(a => a.department === course && a.type === activeFilters.type).length;

      if (course.includes('IT') || course.includes('CS')) icon = '💻';
      else if (course.includes('BCS')) icon = '🖥️';
      else if (course.includes('Office')) icon = '🏢';
      else if (course.includes('Lab')) icon = '🔬';
      else if (course.includes('Math') || course.includes('Stat')) icon = '📐';
      else if (course.includes('Library')) icon = '📚';
      else if (course.includes('Hostel')) icon = '🏠';
    }

    if (course !== 'ALL' && count === 0 && activeFilters.department !== course) return;

    const chip = document.createElement('div');
    const isSelected = activeFilters.department === course;
    chip.className = `course-chip ${isSelected ? 'active' : ''}`;
    chip.setAttribute('role', 'button');
    chip.setAttribute('tabindex', '0');
    chip.title = `Click to display all computers and items in ${course}`;

    chip.innerHTML = `
      <span>${course === 'ALL' ? label : `${icon} ${label}`}</span>
      <span class="course-chip-count">${count}</span>
    `;

    chip.addEventListener('click', () => {
      selectCourse(course);
    });

    quickCourseChips.appendChild(chip);
  });
}

// Handle Course Selection
function selectCourse(course) {
  activeFilters.department = course;
  if (deptSelectFilter) deptSelectFilter.value = course;

  updateActiveBadge();
  renderCourseChips();
  renderDepartmentMatrix();
  renderInventoryList();
  updateBanner();
  scrollToInventory();
  showToast(`Displaying equipment for ${course === 'ALL' ? 'All Departments' : course}`);
}

// Update Active Banner & Counters
function updateBanner() {
  const filteredCount = getFilteredAssets().length;
  if (showingAssetsCount) showingAssetsCount.textContent = filteredCount;

  if (activeFilters.department === 'ALL' && activeFilters.type === 'ALL' && activeFilters.status === 'ALL' && !activeFilters.hasExcel && !activeFilters.search) {
    if (activeFilterBanner) activeFilterBanner.style.display = 'none';
    if (courseChooserCount) courseChooserCount.textContent = `All (${assets.length} Assets)`;
  } else {
    if (activeFilterBanner) activeFilterBanner.style.display = 'flex';
    let desc = '';
    const typeLabel = activeFilters.type === 'Computer' ? 'Computers' : activeFilters.type === 'Printer' ? 'Printers' : activeFilters.type === 'Projector' ? 'Projectors' : 'Equipment';

    if (activeFilters.department !== 'ALL') {
      const deptPcs = assets.filter(a => a.department === activeFilters.department && a.type === activeFilters.type).length;
      const deptWork = assets.filter(a => a.department === activeFilters.department && a.type === activeFilters.type && a.status === 'Working').length;
      const deptFault = deptPcs - deptWork;
      desc = `Course: <strong>${escapeHtml(activeFilters.department)}</strong> &bull; Showing <strong>${filteredCount}</strong> ${typeLabel} (🟢 ${deptWork} Working, 🔴 ${deptFault} Faulty)`;
      if (courseChooserCount) courseChooserCount.textContent = `${activeFilters.department} (${filteredCount} items)`;
    } else if (activeFilters.type !== 'ALL') {
      desc = `Viewing Category: <strong>${escapeHtml(activeFilters.type)}</strong> &bull; Found <strong>${filteredCount}</strong> items`;
      if (courseChooserCount) courseChooserCount.textContent = `${activeFilters.type} (${filteredCount})`;
    } else if (activeFilters.status !== 'ALL') {
      desc = `Status: <strong>${escapeHtml(activeFilters.status)}</strong> &bull; Found <strong>${filteredCount}</strong> items`;
    } else if (activeFilters.hasExcel) {
      desc = `Filtered: <strong>Assets with Excel Audits</strong> &bull; Found <strong>${filteredCount}</strong> items`;
    } else if (activeFilters.search) {
      desc = `Search: <strong>"${escapeHtml(activeFilters.search)}"</strong> &bull; Found <strong>${filteredCount}</strong> items`;
    }

    if (bannerText) bannerText.innerHTML = desc;
  }
}

// Render All 60 Departments Grid
function renderDepartmentMatrix() {
  if (!departmentMatrixContainer) return;
  departmentMatrixContainer.innerHTML = '';

  departments.forEach(dept => {
    const deptAssets = assets.filter(a => a.department === dept);
    if (deptAssets.length === 0) return;

    const workingCount = deptAssets.filter(a => a.status === 'Working').length;
    const faultyCount = deptAssets.length - workingCount;
    const isSelected = activeFilters.department === dept;

    const card = document.createElement('div');
    card.className = `dept-pill-card ${isSelected ? 'active' : ''}`;
    card.setAttribute('role', 'button');
    card.setAttribute('tabindex', '0');
    card.title = `Click cursor to display all equipment in ${dept}`;

    card.innerHTML = `
      <div class="dept-card-top">
        <h4 class="dept-name" title="${escapeHtml(dept)}">${escapeHtml(dept)}</h4>
        <span class="badge-counter">${deptAssets.length}</span>
      </div>
      <div class="dept-device-counts">
        <span class="count-working">🟢 ${workingCount} Working</span>
        <span class="count-faulty">🔴 ${faultyCount} Faulty</span>
      </div>
      <div class="dept-types-badge">
        <span>🖥️ ${deptAssets.filter(a => a.type === 'Computer').length} PCs</span>
        <span>🖨️ ${deptAssets.filter(a => a.type === 'Printer').length} Prn</span>
        <span>📽️ ${deptAssets.filter(a => a.type === 'Projector').length} Prj</span>
      </div>
    `;

    card.addEventListener('click', () => {
      if (activeFilters.department === dept) {
        selectCourse('ALL');
      } else {
        selectCourse(dept);
      }
    });

    departmentMatrixContainer.appendChild(card);
  });
}

function updateActiveBadge() {
  if (!activeFilterBadge) return;
  if (activeFilters.department === 'ALL') {
    activeFilterBadge.textContent = 'Showing All';
    activeFilterBadge.style.background = 'rgba(59, 130, 246, 0.15)';
    activeFilterBadge.style.color = '#60a5fa';
  } else {
    activeFilterBadge.textContent = `Course: ${activeFilters.department}`;
    activeFilterBadge.style.background = 'rgba(16, 185, 129, 0.2)';
    activeFilterBadge.style.color = '#34d399';
  }
}

// Populate all Department Dropdowns
function renderDeptDropdowns() {
  const currentVal = deptSelectFilter ? deptSelectFilter.value || 'ALL' : 'ALL';
  if (deptSelectFilter) {
    deptSelectFilter.innerHTML = `<option value="ALL">All Departments (${assets.length})</option>`;
  }

  const formDeptSelect = document.getElementById('formDepartment');
  const inlineDeptSelect = document.getElementById('inlineDepartment');
  const inlinePrinterDept = document.getElementById('inlinePrinterDept');
  const inlineProjectorDept = document.getElementById('inlineProjectorDept');

  if (formDeptSelect) formDeptSelect.innerHTML = '';
  if (inlineDeptSelect) inlineDeptSelect.innerHTML = '';
  if (inlinePrinterDept) inlinePrinterDept.innerHTML = '';
  if (inlineProjectorDept) inlineProjectorDept.innerHTML = '';

  departments.forEach(dept => {
    const count = assets.filter(a => a.department === dept).length;
    
    if (deptSelectFilter) {
      const opt = document.createElement('option');
      opt.value = dept;
      opt.textContent = `${dept} (${count})`;
      deptSelectFilter.appendChild(opt);
    }

    const createOpt = () => {
      const o = document.createElement('option');
      o.value = dept;
      o.textContent = dept;
      return o;
    };

    if (formDeptSelect) formDeptSelect.appendChild(createOpt());
    if (inlineDeptSelect) inlineDeptSelect.appendChild(createOpt());
    if (inlinePrinterDept) inlinePrinterDept.appendChild(createOpt());
    if (inlineProjectorDept) inlineProjectorDept.appendChild(createOpt());
  });

  if (deptSelectFilter) deptSelectFilter.value = currentVal;
}

// Render Chart
function renderChart() {
  const canvas = document.getElementById('statusChart');
  if (!canvas || typeof Chart === 'undefined') return;

  const targetAssets = activeFilters.department === 'ALL'
    ? (activeFilters.type === 'ALL' ? assets : assets.filter(a => a.type === activeFilters.type))
    : (activeFilters.type === 'ALL' ? assets.filter(a => a.department === activeFilters.department) : assets.filter(a => a.department === activeFilters.department && a.type === activeFilters.type));

  const working = targetAssets.filter(a => a.status === 'Working').length;
  const faulty = targetAssets.length - working;

  if (chartInstance) {
    chartInstance.destroy();
  }

  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  const textColor = isDark ? '#94a3b8' : '#64748b';

  if (chartType === 'doughnut') {
    chartInstance = new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels: ['Operational / Working', 'Not Working / Faulty'],
        datasets: [{
          data: [working, faulty],
          backgroundColor: ['#10b981', '#ef4444'],
          borderColor: isDark ? '#1e293b' : '#ffffff',
          borderWidth: 2
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: { color: textColor, font: { size: 12 } }
          }
        },
        cutout: '65%'
      }
    });
  } else {
    const topDepts = [...departments]
      .map(d => ({
        name: d,
        working: assets.filter(a => a.department === d && a.status === 'Working').length,
        faulty: assets.filter(a => a.department === d && a.status === 'Not Working').length
      }))
      .filter(d => (d.working + d.faulty) > 0)
      .sort((a, b) => (b.working + b.faulty) - (a.working + a.faulty))
      .slice(0, 8);

    chartInstance = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: topDepts.map(d => d.name.length > 12 ? d.name.substring(0, 11) + '..' : d.name),
        datasets: [
          { label: 'Working', data: topDepts.map(d => d.working), backgroundColor: '#10b981' },
          { label: 'Faulty', data: topDepts.map(d => d.faulty), backgroundColor: '#ef4444' }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { stacked: true, ticks: { color: textColor, font: { size: 10 } } },
          y: { stacked: true, ticks: { color: textColor } }
        },
        plugins: {
          legend: { position: 'bottom', labels: { color: textColor, font: { size: 11 } } }
        }
      }
    });
  }
}

// Filter assets helper
function getFilteredAssets() {
  return assets.filter(asset => {
    // Search query
    if (activeFilters.search) {
      const q = activeFilters.search.toLowerCase();
      const match =
        (asset.id && asset.id.toLowerCase().includes(q)) ||
        (asset.name && asset.name.toLowerCase().includes(q)) ||
        (asset.department && asset.department.toLowerCase().includes(q)) ||
        (asset.serialNumber && asset.serialNumber.toLowerCase().includes(q)) ||
        (asset.specs && asset.specs.toLowerCase().includes(q)) ||
        (asset.status && asset.status.toLowerCase().includes(q)) ||
        (asset.assignedUser && asset.assignedUser.toLowerCase().includes(q));
      if (!match) return false;
    }

    // Type filter
    if (activeFilters.type !== 'ALL' && asset.type !== activeFilters.type) {
      return false;
    }

    // Department / Course filter
    if (activeFilters.department !== 'ALL' && asset.department !== activeFilters.department) {
      return false;
    }

    // Status filter
    if (activeFilters.status !== 'ALL' && asset.status !== activeFilters.status) {
      return false;
    }

    // Has Excel filter
    if (activeFilters.hasExcel && (!asset.excelAttachments || asset.excelAttachments.length === 0)) {
      return false;
    }

    return true;
  });
}

// Render Inventory List (Table or Cards view)
function renderInventoryList() {
  const filtered = getFilteredAssets();

  if (currentViewMode === 'table') {
    if (tableContainer) tableContainer.style.display = 'block';
    if (assetCardsContainer) assetCardsContainer.style.display = 'none';
    renderTableView(filtered);
  } else {
    if (tableContainer) tableContainer.style.display = 'none';
    if (assetCardsContainer) assetCardsContainer.style.display = 'grid';
    renderCardsView(filtered);
  }
}

// Render Table View
function renderTableView(filtered) {
  if (!assetTableBody) return;
  assetTableBody.innerHTML = '';

  if (filtered.length === 0) {
    if (assets.length === 0) {
      assetTableBody.innerHTML = `
        <tr>
          <td colspan="7" style="text-align: center; padding: 48px 20px; color: var(--text-muted);">
            <div style="font-size: 2.5rem; margin-bottom: 10px;">🗑️</div>
            <h3 style="font-size: 1.1rem; color: var(--text-main); margin-bottom: 6px;">सर्व डेटा यशस्वीरित्या नष्ट केला आहे (All Data Erased - Count is 0)</h3>
            <p style="font-size: 0.85rem; margin-bottom: 16px;">नवीन डेटा जोडण्यासाठी वरील <strong>"📥 एक्सेल इम्पोर्ट"</strong> वापरा किंवा मूळ नमुना डेटा परत आणा.</p>
            <button type="button" class="btn btn-primary btn-sm btn-restore-sample-trigger" style="font-weight: 700;">
              🔄 मूळ नमुना डेटा परत आणा (Restore Sample Data)
            </button>
          </td>
        </tr>
      `;
      const btnR = assetTableBody.querySelector('.btn-restore-sample-trigger');
      if (btnR) btnR.addEventListener('click', restoreDefaultData);
      return;
    }

    const itemType = activeFilters.type === 'Projector' ? 'projectors' : activeFilters.type === 'Printer' ? 'printers' : 'computers';
    assetTableBody.innerHTML = `
      <tr>
        <td colspan="7" style="text-align: center; padding: 40px; color: var(--text-muted);">
          <div style="font-size: 2.5rem; margin-bottom: 8px;">🔍</div>
          <h4>No ${itemType} found under current filter</h4>
          <p style="font-size: 0.85rem; margin-top: 6px;">
            <button class="btn btn-secondary btn-sm" onclick="resetAllFilters()">Reset All Filters</button>
          </p>
        </td>
      </tr>
    `;
    return;
  }

  filtered.forEach(asset => {
    const tr = document.createElement('tr');
    tr.setAttribute('data-asset-id', asset.id);

    const isWorking = asset.status === 'Working';
    const typeIcon = asset.type === 'Computer' ? '🖥️' : asset.type === 'Printer' ? '🖨️' : '📽️';
    const excelCount = asset.excelAttachments ? asset.excelAttachments.length : 0;

    let specBadgesHtml = '';
    if (asset.specs) {
      const parts = asset.specs.split(',').map(s => s.trim());
      parts.forEach(p => {
        let cls = 'spec-tag';
        const lower = p.toLowerCase();
        if (lower.includes('i5') || lower.includes('i3') || lower.includes('ryzen') || lower.includes('duo') || lower.includes('pentium')) {
          cls += ' cpu';
        } else if (lower.includes('gb ram') || lower.includes('mb ram') || lower.includes('16gb') || lower.includes('8gb')) {
          cls += ' ram';
        } else if (lower.includes('ssd') || lower.includes('hdd') || lower.includes('tb') || lower.includes('wd') || lower.includes('seagate')) {
          cls += ' storage';
        } else if (lower.includes('win') || lower.includes('ubuntu')) {
          cls += ' os';
        }
        specBadgesHtml += `<span class="${cls}">${escapeHtml(p)}</span>`;
      });
    }

    tr.innerHTML = `
      <td>
        <div style="display: flex; align-items: center; gap: 10px;">
          <span style="font-size: 1.4rem;" title="${asset.type}">${typeIcon}</span>
          <div>
            <div style="font-weight: 700; color: var(--text-main); font-size: 0.92rem;">${escapeHtml(asset.name)}</div>
            <div style="display: flex; gap: 6px; align-items: center; margin-top: 2px;">
              <span class="asset-tag-code">${escapeHtml(asset.id)}</span>
              <span style="font-size: 0.75rem; color: var(--text-muted); font-family: monospace;">${escapeHtml(asset.serialNumber || '')}</span>
            </div>
          </div>
        </div>
      </td>
      <td>
        <div style="font-weight: 600; color: #60a5fa; cursor: pointer;" data-action="filter-course" data-dept="${escapeHtml(asset.department)}" title="Click to view all ${escapeHtml(asset.department)} items">
          ${escapeHtml(asset.department)}
        </div>
        <div style="font-size: 0.78rem; color: var(--text-muted);">${escapeHtml(asset.location || '')}</div>
      </td>
      <td>
        <div>${escapeHtml(asset.assignedUser || 'Shared Lab')}</div>
        <div style="font-size: 0.75rem; color: var(--text-muted);">${escapeHtml(asset.ipAddress || '')}</div>
      </td>
      <td style="max-width: 320px;">
        <div class="spec-badges-wrapper">
          ${specBadgesHtml || `<span class="spec-tag">${escapeHtml(asset.specs || 'N/A')}</span>`}
        </div>
        ${asset.faultDetails ? `
          <div style="font-size: 0.76rem; color: #ef4444; margin-top: 4px; font-weight: 600;">
            ⚠️ Issue: ${escapeHtml(asset.faultDetails.issue || 'Fault reported')}
          </div>
        ` : ''}
      </td>
      <td>
        <button class="status-badge ${isWorking ? 'working' : 'faulty'}" data-action="toggle-status" data-id="${asset.id}" title="Click to toggle status">
          <span class="status-dot"></span>
          <span>${isWorking ? 'Working' : 'Not Working'}</span>
        </button>
      </td>
      <td>
        <div style="display: flex; align-items: center; gap: 4px;">
          <button class="btn btn-sm ${excelCount > 0 ? 'btn-excel' : 'btn-secondary'}" data-action="manage-excel" data-id="${asset.id}" style="padding: 4px 8px; font-size: 0.78rem;">
            📊 ${excelCount > 0 ? `${excelCount} File${excelCount > 1 ? 's' : ''}` : 'Attach'}
          </button>
          ${excelCount > 0 ? `
            <button class="btn btn-sm btn-danger-outline" data-action="delete-asset-excel" data-id="${asset.id}" title="Delete Excel files for ${asset.id}" style="padding: 4px 6px; font-size: 0.75rem;">
              🗑️
            </button>
          ` : ''}
        </div>
      </td>
      <td style="text-align: right; white-space: nowrap;">
        <button class="btn btn-secondary btn-sm" data-action="quick-view" data-id="${asset.id}" title="View full hardware specs">
          🔍 Specs
        </button>
        <button class="btn btn-icon-only" data-action="edit-asset" data-id="${asset.id}" title="Edit asset">
          ✏️
        </button>
        <button class="btn btn-icon-only" data-action="delete-asset" data-id="${asset.id}" title="Delete asset" style="color: #ef4444;">
          🗑️
        </button>
      </td>
    `;

    assetTableBody.appendChild(tr);
  });
}

// Render Cards View
function renderCardsView(filtered) {
  if (!assetCardsContainer) return;
  assetCardsContainer.innerHTML = '';

  if (filtered.length === 0) {
    if (assets.length === 0) {
      assetCardsContainer.innerHTML = `
        <div style="grid-column: 1 / -1; text-align: center; padding: 48px 20px; background: var(--bg-secondary); border-radius: var(--radius-lg); border: 1px dashed var(--border-color);">
          <div style="font-size: 2.5rem; margin-bottom: 10px;">🗑️</div>
          <h3 style="font-size: 1.1rem; color: var(--text-main); margin-bottom: 6px;">सर्व डेटा नष्ट केला आहे (All Data Has Been Erased - Total: 0)</h3>
          <p style="font-size: 0.85rem; color: var(--text-muted); margin-bottom: 16px;">नवीन डेटा इंपोर्ट करण्यासाठी <strong>"📥 एक्सेल इम्पोर्ट"</strong> वापरा किंवा मूळ नमुना डेटा परत आणा.</p>
          <button type="button" class="btn btn-primary btn-sm btn-cards-restore-sample" style="font-weight: 700;">
            🔄 मूळ नमुना डेटा परत आणा (Restore Sample Data)
          </button>
        </div>
      `;
      const btnR = assetCardsContainer.querySelector('.btn-cards-restore-sample');
      if (btnR) btnR.addEventListener('click', restoreDefaultData);
      return;
    }

    assetCardsContainer.innerHTML = `
      <div style="grid-column: 1 / -1; text-align: center; padding: 40px; color: var(--text-muted);">
        <h4>No equipment matched the filters</h4>
      </div>
    `;
    return;
  }

  filtered.forEach(asset => {
    const isWorking = asset.status === 'Working';
    const typeIcon = asset.type === 'Computer' ? '🖥️' : asset.type === 'Printer' ? '🖨️' : '📽️';
    const excelCount = asset.excelAttachments ? asset.excelAttachments.length : 0;

    const card = document.createElement('div');
    card.className = `asset-spec-card ${isWorking ? 'card-working' : 'card-faulty'}`;

    card.innerHTML = `
      <div>
        <div class="asset-card-title-row">
          <div style="display: flex; align-items: center; gap: 8px;">
            <span style="font-size: 1.5rem;">${typeIcon}</span>
            <div>
              <div style="font-weight: 700; font-size: 0.95rem;">${escapeHtml(asset.name)}</div>
              <div style="display: flex; gap: 6px; align-items: center; margin-top: 2px;">
                <span class="asset-tag-code">${escapeHtml(asset.id)}</span>
                <span style="font-size: 0.74rem; color: #60a5fa; cursor: pointer;" data-action="filter-course" data-dept="${escapeHtml(asset.department)}">
                  🏫 ${escapeHtml(asset.department)}
                </span>
              </div>
            </div>
          </div>
          <button class="status-badge ${isWorking ? 'working' : 'faulty'}" data-action="toggle-status" data-id="${asset.id}">
            <span class="status-dot"></span>
            <span>${isWorking ? 'Working' : 'Faulty'}</span>
          </button>
        </div>

        <div class="asset-card-specs-box">
          <div style="color: var(--text-muted); font-size: 0.74rem; margin-bottom: 4px; font-weight: 600;">SPECIFICATIONS:</div>
          <div style="font-size: 0.82rem; line-height: 1.4;">${escapeHtml(asset.specs || 'Standard Config')}</div>
          ${asset.faultDetails ? `
            <div style="color: #ef4444; font-size: 0.75rem; margin-top: 6px; font-weight: 600;">
              ⚠️ Fault: ${escapeHtml(asset.faultDetails.issue)}
            </div>
          ` : ''}
        </div>
      </div>

      <div class="asset-card-footer">
        <button class="btn btn-sm ${excelCount > 0 ? 'btn-excel' : 'btn-secondary'}" data-action="manage-excel" data-id="${asset.id}">
          📊 ${excelCount > 0 ? `${excelCount} Audit Sheet` : 'Attach Excel'}
        </button>
        ${excelCount > 0 ? `
          <button class="btn btn-sm btn-danger-outline" data-action="delete-asset-excel" data-id="${asset.id}" title="Delete all Excel files for ${asset.id}" style="padding: 4px 8px; font-size: 0.78rem;">
            🗑️ Delete
          </button>
        ` : ''}
        <button class="btn btn-primary btn-sm" data-action="quick-view" data-id="${asset.id}">
          🔍 View Full Specs
        </button>
      </div>
    `;

    assetCardsContainer.appendChild(card);
  });
}

// Scroll to Inventory List
function scrollToInventory() {
  const target = document.getElementById('inventorySection');
  if (target) {
    target.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

// Setup Event Listeners
function setupEventListeners() {
  // Primary Navigation Tabs
  document.querySelectorAll('.college-nav-tabs .nav-tab-btn[data-nav-tab]').forEach(tabBtn => {
    tabBtn.addEventListener('click', () => {
      const tab = tabBtn.dataset.navTab;

      // Reset active on all tabs
      document.querySelectorAll('.college-nav-tabs .nav-tab-btn').forEach(b => b.classList.remove('active'));
      tabBtn.classList.add('active');

      if (tab === 'computers') {
        inlineAddEquipmentPanel.style.display = 'none';
        activeFilters.type = 'Computer';
        activeFilters.department = 'ALL';
        activeFilters.status = 'ALL';
        activeFilters.search = '';
        activeFilters.hasExcel = false;
        if (deptSelectFilter) deptSelectFilter.value = 'ALL';
        if (searchInput) searchInput.value = '';
        syncFilterPills('Computer');
        renderAll();
        scrollToInventory();
        showToast('💻 सर्व संगणक यादी (Displaying all Computers)');
      } else if (tab === 'printers') {
        inlineAddEquipmentPanel.style.display = 'none';
        activeFilters.type = 'Printer';
        activeFilters.department = 'ALL';
        activeFilters.status = 'ALL';
        activeFilters.search = '';
        activeFilters.hasExcel = false;
        if (deptSelectFilter) deptSelectFilter.value = 'ALL';
        if (searchInput) searchInput.value = '';
        syncFilterPills('Printer');
        renderAll();
        scrollToInventory();
        showToast('🖨️ सर्व प्रिंटर यादी (Displaying all Printers)');
      } else if (tab === 'projectors') {
        // FIXED: PROJECTOR TAB CLICK - RESET DEPT FILTER SO ALL 20 PROJECTORS ARE DISPLAYED
        inlineAddEquipmentPanel.style.display = 'none';
        activeFilters.type = 'Projector';
        activeFilters.department = 'ALL';
        activeFilters.status = 'ALL';
        activeFilters.search = '';
        activeFilters.hasExcel = false;
        if (deptSelectFilter) deptSelectFilter.value = 'ALL';
        if (searchInput) searchInput.value = '';
        syncFilterPills('Projector');
        renderAll();
        scrollToInventory();
        showToast('📽️ सर्व प्रोजेक्टर यादी (Displaying all 20 Projectors)');
      } else if (tab === 'add-item') {
        inlineAddEquipmentPanel.style.display = 'block';
        inlineAddEquipmentPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
        const deadStockInput = document.getElementById('inlineDeadStock');
        if (deadStockInput) deadStockInput.focus();
        showToast('➕ नवीन उपकरण नोंदणी फॉर्म उघडला आहे (Add Equipment Form opened)');
      } else if (tab === 'excel-import') {
        openModal('importModal');
        renderGlobalExcelTable();
      }
    });
  });

  // Inline Form Type Selector (Computer, Printer, Projector)
  if (inlineAssetType) {
    inlineAssetType.addEventListener('change', (e) => {
      const selectedType = e.target.value;
      if (selectedType === 'Computer') {
        inlinePcForm.style.display = 'block';
        inlinePrinterForm.style.display = 'none';
        inlineProjectorForm.style.display = 'none';
      } else if (selectedType === 'Printer') {
        inlinePcForm.style.display = 'none';
        inlinePrinterForm.style.display = 'block';
        inlineProjectorForm.style.display = 'none';
      } else if (selectedType === 'Projector') {
        inlinePcForm.style.display = 'none';
        inlinePrinterForm.style.display = 'none';
        inlineProjectorForm.style.display = 'block';
      }
    });
  }

  // Inline PC Form Submission
  if (inlinePcForm) {
    inlinePcForm.addEventListener('submit', (e) => {
      e.preventDefault();
      saveInlinePc();
    });
  }

  // Inline Printer Form Submission
  if (inlinePrinterForm) {
    inlinePrinterForm.addEventListener('submit', (e) => {
      e.preventDefault();
      saveInlinePrinter();
    });
  }

  // Inline Projector Form Submission
  if (inlineProjectorForm) {
    inlineProjectorForm.addEventListener('submit', (e) => {
      e.preventDefault();
      saveInlineProjector();
    });
  }

  // Cancel buttons for inline form
  document.querySelectorAll('#btnCancelInlineAdd, .btn-cancel-inline').forEach(btn => {
    btn.addEventListener('click', () => {
      inlineAddEquipmentPanel.style.display = 'none';
      document.querySelectorAll('.college-nav-tabs .nav-tab-btn').forEach(b => b.classList.remove('active'));
      const activeTabBtn = document.querySelector(`.college-nav-tabs .nav-tab-btn[data-nav-tab="${activeFilters.type === 'Printer' ? 'printers' : activeFilters.type === 'Projector' ? 'projectors' : 'computers'}"]`);
      if (activeTabBtn) activeTabBtn.classList.add('active');
      scrollToInventory();
    });
  });

  // Search input
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      activeFilters.search = e.target.value.trim();
      renderInventoryList();
      updateBanner();
    });
  }

  // Course search filter inside chooser bar
  if (courseFilterSearch) {
    courseFilterSearch.addEventListener('input', () => {
      renderCourseChips();
    });
  }

  // Toggle All 60 Departments matrix
  const btnToggleDeptMatrix = document.getElementById('btnToggleDeptMatrix');
  const analyticsGridSection = document.getElementById('analyticsGridSection');
  if (btnToggleDeptMatrix && analyticsGridSection) {
    btnToggleDeptMatrix.addEventListener('click', () => {
      const isHidden = analyticsGridSection.style.display === 'none';
      analyticsGridSection.style.display = isHidden ? 'grid' : 'none';
      btnToggleDeptMatrix.textContent = isHidden ? 'Hide Dept Matrix' : '📁 Browse All 60 Depts';
      if (isHidden) {
        analyticsGridSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }
    });
  }

  // Department dropdown
  if (deptSelectFilter) {
    deptSelectFilter.addEventListener('change', (e) => {
      selectCourse(e.target.value);
    });
  }

  // Reset filters
  if (btnResetFilters) btnResetFilters.addEventListener('click', resetAllFilters);
  if (btnClearActiveBanner) btnClearActiveBanner.addEventListener('click', resetAllFilters);

  // KPI Cards Cursor Click-to-Display
  document.querySelectorAll('.kpi-card[data-kpi-filter]').forEach(card => {
    card.addEventListener('click', () => {
      const filter = card.dataset.kpiFilter;
      if (filter === 'ALL') {
        activeFilters.type = 'ALL';
        activeFilters.department = 'ALL';
        activeFilters.hasExcel = false;
        syncFilterPills('ALL');
      } else if (filter === 'Excel') {
        activeFilters.hasExcel = true;
        const excelPill = document.querySelector('.pill-btn[data-filter-excel]');
        if (excelPill) excelPill.classList.add('active');
      } else {
        activeFilters.type = filter;
        activeFilters.department = 'ALL';
        activeFilters.hasExcel = false;
        syncFilterPills(filter);
      }

      if (deptSelectFilter) deptSelectFilter.value = 'ALL';
      renderAll();
      scrollToInventory();
      showToast(`Showing ${filter === 'ALL' ? 'All Equipment' : filter === 'Excel' ? 'Excel Audit Files' : `All ${filter}s`}`);
    });
  });

  // Filter pills
  document.querySelectorAll('.pill-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      if (btn.dataset.filterType) {
        document.querySelectorAll('.pill-btn[data-filter-type]').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        activeFilters.type = btn.dataset.filterType;
        activeFilters.department = 'ALL';
        if (deptSelectFilter) deptSelectFilter.value = 'ALL';
      } else if (btn.dataset.filterStatus) {
        document.querySelectorAll('.pill-btn[data-filter-status]').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        activeFilters.status = btn.dataset.filterStatus;
      } else if (btn.dataset.filterExcel) {
        btn.classList.toggle('active');
        activeFilters.hasExcel = btn.classList.contains('active');
      }
      renderAll();
    });
  });

  // View Mode Switcher
  const viewModeToggle = document.getElementById('viewModeToggle');
  if (viewModeToggle) {
    viewModeToggle.addEventListener('click', (e) => {
      const btn = e.target.closest('.view-btn');
      if (!btn) return;
      document.querySelectorAll('.view-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentViewMode = btn.dataset.view;
      renderInventoryList();
    });
  }

  // Delegated Actions in Inventory (Table & Cards)
  const inventoryWrapper = document.getElementById('inventoryMainWrapper');
  if (inventoryWrapper) {
    inventoryWrapper.addEventListener('click', (e) => {
      const courseTarget = e.target.closest('[data-action="filter-course"]');
      if (courseTarget) {
        e.stopPropagation();
        selectCourse(courseTarget.dataset.dept);
        return;
      }

      const target = e.target.closest('[data-action]');
      if (!target) {
        const row = e.target.closest('tr[data-asset-id]');
        if (row) {
          const asset = assets.find(a => a.id === row.dataset.assetId);
          if (asset) openQuickSpecsModal(asset);
        }
        return;
      }

      e.stopPropagation();
      const action = target.dataset.action;
      const id = target.dataset.id;
      const asset = assets.find(a => a.id === id);
      if (!asset) return;

      if (action === 'toggle-status') {
        asset.status = asset.status === 'Working' ? 'Not Working' : 'Working';
        if (asset.status === 'Working') {
          asset.faultDetails = null;
        } else {
          asset.faultDetails = {
            issue: 'Hardware maintenance required',
            reportedDate: new Date().toISOString().split('T')[0],
            reportedBy: 'Staff',
            priority: 'High',
            technician: 'Maintenance Team'
          };
        }
        saveData();
        renderAll();
        showToast(`${asset.id} status updated to ${asset.status}`);
      } else if (action === 'quick-view') {
        openQuickSpecsModal(asset);
      } else if (action === 'manage-excel') {
        openExcelModal(asset);
      } else if (action === 'edit-asset') {
        openAssetModal(asset);
      } else if (action === 'delete-asset') {
        if (confirm(`Are you sure you want to delete ${asset.name} (${asset.id})?`)) {
          assets = assets.filter(a => a.id !== id);
          saveData();
          renderAll();
          showToast(`Asset ${id} deleted`);
        }
      } else if (action === 'delete-asset-excel') {
        const count = (asset.excelAttachments || []).length;
        if (count === 0) {
          showToast('या उपकरणास कोणतीही एक्सेल फाइल जोडलेली नाही (No Excel file attached).');
          return;
        }
        if (confirm(`या उपकरणाची (${asset.id}) सर्व (${count}) एक्सेल फाइल्स हटवायची आहेत का?\n(Are you sure you want to delete all ${count} Excel file(s) attached to ${asset.name} (${asset.id})?)`)) {
          asset.excelAttachments = [];
          saveData();
          renderAll();
          showToast(`🗑️ ${asset.id} ची सर्व एक्सेल फाइल्स हटवली (Excel files deleted for ${asset.id})`);
        }
      }
    });
  }

  // Chart Switch View
  const btnToggleChartType = document.getElementById('btnToggleChartType');
  if (btnToggleChartType) {
    btnToggleChartType.addEventListener('click', () => {
      chartType = chartType === 'doughnut' ? 'bar' : 'doughnut';
      renderChart();
    });
  }

  // Theme Toggle
  const btnThemeToggle = document.getElementById('btnThemeToggle');
  const savedTheme = localStorage.getItem('assetpulse_theme') || 'light';
  document.documentElement.setAttribute('data-theme', savedTheme);

  if (btnThemeToggle) {
    btnThemeToggle.addEventListener('click', () => {
      const current = document.documentElement.getAttribute('data-theme') || 'light';
      const next = current === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next);
      localStorage.setItem('assetpulse_theme', next);
      renderChart();
    });
  }

  // Modal Close buttons
  document.querySelectorAll('[data-close-modal]').forEach(btn => {
    btn.addEventListener('click', () => {
      closeModal(btn.dataset.closeModal);
    });
  });

  // Modal form toggle
  const formStatus = document.getElementById('formStatus');
  if (formStatus) {
    formStatus.addEventListener('change', () => {
      const faultSection = document.getElementById('formFaultSection');
      if (faultSection) {
        faultSection.style.display = formStatus.value === 'Not Working' ? 'block' : 'none';
      }
    });
  }

  // Excel Export
  const btnExportExcel = document.getElementById('btnExportExcel');
  if (btnExportExcel) {
    btnExportExcel.addEventListener('click', exportInventoryExcel);
  }

  // Excel Bulk Import Modal
  const btnImportExcel = document.getElementById('btnImportExcel');
  if (btnImportExcel) {
    btnImportExcel.addEventListener('click', () => {
      openModal('importModal');
      updateImportModalActiveFile();
      renderGlobalExcelTable();
    });
  }

  // Quick Modal Excel link
  const btnQuickOpenExcel = document.getElementById('btnQuickOpenExcel');
  if (btnQuickOpenExcel) {
    btnQuickOpenExcel.addEventListener('click', () => {
      if (currentQuickAsset) {
        closeModal('specsQuickModal');
        openExcelModal(currentQuickAsset);
      }
    });
  }

  // Toolbar Erase Data button (Reset to 0)
  const btnToolbarEraseData = document.getElementById('btnToolbarEraseData');
  if (btnToolbarEraseData) {
    btnToolbarEraseData.addEventListener('click', eraseAllInventoryData);
  }

  setupExcelAndDriveUpload();
  setupBulkExcelImportAndManager();
  setupInlineFormAttachments();
}

// Sync filter pills state
function syncFilterPills(filterType) {
  document.querySelectorAll('.pill-btn[data-filter-type]').forEach(b => {
    if (b.dataset.filterType === filterType) {
      b.classList.add('active');
    } else {
      b.classList.remove('active');
    }
  });
}

// Save Inline PC Registration (From user screenshot)
function saveInlinePc() {
  const deadStock = document.getElementById('inlineDeadStock').value.trim();
  const make = document.getElementById('inlineMake').value.trim();
  const deviceType = document.getElementById('inlineDeviceType').value;
  const cpu = document.getElementById('inlineCpu').value.trim();
  const dept = document.getElementById('inlineDepartment').value;
  const motherboard = document.getElementById('inlineMotherboard').value.trim();
  const ram = document.getElementById('inlineRam').value.trim();
  const hdd = document.getElementById('inlineHdd').value.trim();
  const ssd = document.getElementById('inlineSsd').value.trim();
  const os = document.getElementById('inlineOs').value.trim();
  const monitor = document.getElementById('inlineMonitor').value.trim();
  const keyboardMouse = document.getElementById('inlineKeyboardMouse').value.trim();
  const status = document.getElementById('inlineStatus').value;
  const remark = document.getElementById('inlineRemark').value.trim();
  const location = document.getElementById('inlineLocation').value.trim();
  const user = document.getElementById('inlineUser').value.trim();

  if (!deadStock || !make || !dept) {
    alert('कृपया Dead Stock Number, Make आणि Department भरा (Please fill required fields).');
    return;
  }

  const specList = [];
  if (cpu) specList.push(cpu);
  if (ram) specList.push(ram);
  if (hdd) specList.push(hdd + (hdd.toLowerCase().includes('hdd') ? '' : ' HDD'));
  if (ssd) specList.push(ssd + (ssd.toLowerCase().includes('ssd') ? '' : ' SSD'));
  if (os) specList.push(os);
  if (motherboard) specList.push('MB: ' + motherboard);
  if (monitor) specList.push('Monitor: ' + monitor);
  if (keyboardMouse) specList.push('Kbd/Mouse: ' + keyboardMouse);

  const pcId = deadStock || `PC-${assets.length + 1}`;

  const newAsset = {
    id: pcId,
    type: 'Computer',
    name: `${make} ${cpu || 'Desktop'}`.trim(),
    department: dept,
    location: location || `${dept} Computer Lab`,
    assignedUser: user || `${dept} Faculty / Students`,
    status: status,
    serialNumber: deadStock,
    purchaseDate: new Date().toISOString().split('T')[0],
    specs: specList.join(', ') || 'Desktop PC Configuration',
    ipAddress: user && user.includes('.') ? user : `192.168.10.${(assets.length % 250) + 1}`,
    faultDetails: status === 'Not Working' ? {
      issue: remark || 'Hardware issue reported',
      reportedDate: new Date().toISOString().split('T')[0],
      reportedBy: 'Staff',
      priority: 'High',
      technician: 'Maintenance Team'
    } : null,
    excelAttachments: (function() {
      if (pendingInlinePcAttachment) {
        const att = pendingInlinePcAttachment;
        pendingInlinePcAttachment = null;
        const label = document.getElementById('inlinePcAttachmentLabel');
        if (label) label.textContent = 'कोणतीही फाईल निवडलेली नाही (No file selected)';
        const rmBtn = document.getElementById('btnRemoveInlinePcAtt');
        if (rmBtn) rmBtn.style.display = 'none';
        return [att];
      }
      return [
        {
          id: `att-${Date.now()}`,
          fileName: `${pcId.replace(/[\/\\?%*:|"<>]/g, '_')}_Audit.xlsx`,
          fileSize: '15.4 KB',
          uploadDate: new Date().toISOString().replace('T', ' ').substring(0, 16),
          notes: `Dead Stock Registration Sheet for ${pcId}`,
          sheetData: [
            ['Attribute', 'Value'],
            ['Dead Stock Register Number', deadStock],
            ['Make (कंपनी)', make],
            ['Type', deviceType],
            ['C.P.U. (Processor)', cpu],
            ['Department (विभाग)', dept],
            ['Motherboard', motherboard],
            ['RAM (स्मृती)', ram],
            ['HDD', hdd],
            ['SSD', ssd],
            ['Operating System (O. S.)', os],
            ['Monitor', monitor],
            ['Keyboard & Mouse', keyboardMouse],
            ['Status', status],
            ['Remark / Formating', remark]
          ]
        }
      ];
    })()
  };

  assets.unshift(newAsset);
  saveData();

  inlinePcForm.reset();
  inlineAddEquipmentPanel.style.display = 'none';

  activeFilters.type = 'Computer';
  activeFilters.department = dept;
  syncFilterPills('Computer');

  renderAll();
  scrollToInventory();
  showToast(`✅ संगणक ${pcId} (${dept}) यशस्वीरित्या नोंदवला गेला! (PC Saved!)`);
}

// Save Inline Printer Registration
function saveInlinePrinter() {
  const assetNo = document.getElementById('inlinePrinterAssetNo').value.trim();
  const dept = document.getElementById('inlinePrinterDept').value;
  const make = document.getElementById('inlinePrinterMake').value.trim();
  const type = document.getElementById('inlinePrinterType').value;
  const port = document.getElementById('inlinePrinterPort').value;
  const status = document.getElementById('inlinePrinterStatus').value;
  const remark = document.getElementById('inlinePrinterRemark').value.trim();

  if (!assetNo || !make || !dept) {
    alert('कृपया Printer Asset Number, Make आणि Department भरा.');
    return;
  }

  const prnId = `PRN-${assetNo.replace(/[\/\\?%*:|"<>]/g, '_')}`;

  const newAsset = {
    id: prnId,
    type: 'Printer',
    name: `${make} (${type})`,
    department: dept,
    location: `${dept} Office / Lab`,
    assignedUser: `${dept} Staff`,
    status: status,
    serialNumber: assetNo,
    purchaseDate: new Date().toISOString().split('T')[0],
    specs: `Type: ${type}, Port: ${port}, Make: ${make}`,
    ipAddress: '',
    faultDetails: status === 'Not Working' ? {
      issue: remark || 'Printer service required',
      reportedDate: new Date().toISOString().split('T')[0],
      reportedBy: 'Staff',
      priority: 'Medium',
      technician: 'Printer Vendor'
    } : null,
    excelAttachments: []
  };

  assets.unshift(newAsset);
  saveData();

  inlinePrinterForm.reset();
  inlineAddEquipmentPanel.style.display = 'none';

  activeFilters.type = 'Printer';
  activeFilters.department = dept;
  syncFilterPills('Printer');

  renderAll();
  scrollToInventory();
  showToast(`✅ प्रिंटर ${assetNo} (${dept}) यशस्वीरित्या नोंदवला गेला! (Printer Saved!)`);
}

// Save Inline Projector Registration
function saveInlineProjector() {
  const serialNo = document.getElementById('inlineProjectorSerial').value.trim();
  const dept = document.getElementById('inlineProjectorDept').value;
  const make = document.getElementById('inlineProjectorMake').value.trim();
  const location = document.getElementById('inlineProjectorLocation').value.trim();
  const status = document.getElementById('inlineProjectorStatus').value;
  const remark = document.getElementById('inlineProjectorRemark').value.trim();

  if (!serialNo || !make || !dept) {
    alert('कृपया Projector Serial Number, Make आणि Department भरा.');
    return;
  }

  const prjId = `PRJ-${serialNo}`;

  const newAsset = {
    id: prjId,
    type: 'Projector',
    name: `${make}`,
    department: dept,
    location: location || `${dept} Lecture Hall`,
    assignedUser: `${dept} Faculty`,
    status: status,
    serialNumber: serialNo,
    purchaseDate: new Date().toISOString().split('T')[0],
    specs: `Model: ${make}, Location: ${location || 'Room'}`,
    ipAddress: '',
    faultDetails: status === 'Not Working' ? {
      issue: remark || 'Projector display/lamp check needed',
      reportedDate: new Date().toISOString().split('T')[0],
      reportedBy: 'Dept Staff',
      priority: 'High',
      technician: 'AV Support Team'
    } : null,
    excelAttachments: []
  };

  assets.unshift(newAsset);
  saveData();

  inlineProjectorForm.reset();
  inlineAddEquipmentPanel.style.display = 'none';

  activeFilters.type = 'Projector';
  activeFilters.department = dept;
  syncFilterPills('Projector');

  renderAll();
  scrollToInventory();
  showToast(`✅ प्रोजेक्टर ${serialNo} (${dept}) यशस्वीरित्या नोंदवला गेला! (Projector Saved!)`);
}

// Reset all filters
function resetAllFilters() {
  activeFilters = {
    search: '',
    department: 'ALL',
    type: 'ALL',
    status: 'ALL',
    hasExcel: false
  };
  if (searchInput) searchInput.value = '';
  if (courseFilterSearch) courseFilterSearch.value = '';
  if (deptSelectFilter) deptSelectFilter.value = 'ALL';
  document.querySelectorAll('.pill-btn').forEach(btn => btn.classList.remove('active'));
  const allPill = document.querySelector('.pill-btn[data-filter-type="ALL"]');
  const allStatusPill = document.querySelector('.pill-btn[data-filter-status="ALL"]');
  if (allPill) allPill.classList.add('active');
  if (allStatusPill) allStatusPill.classList.add('active');
  
  updateActiveBadge();
  renderAll();
  showToast('Filters reset: Displaying all items');
}

// Quick Specs Modal
function openQuickSpecsModal(asset) {
  currentQuickAsset = asset;
  const modal = document.getElementById('specsQuickModal');
  const title = document.getElementById('quickModalDevName');
  const body = document.getElementById('specsQuickBody');

  const isWorking = asset.status === 'Working';
  const typeIcon = asset.type === 'Computer' ? '🖥️' : asset.type === 'Printer' ? '🖨️' : '📽️';
  if (title) title.textContent = `${asset.name} (${asset.id})`;

  let specItems = [];
  if (asset.specs) {
    specItems = asset.specs.split(',').map(s => s.trim());
  }

  if (body) {
    body.innerHTML = `
      <div style="background: var(--bg-tertiary); padding: 14px; border-radius: var(--radius-md); margin-bottom: 16px; border: 1px solid var(--border-color-light);">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
          <div style="font-size: 1.1rem; font-weight: 700;">
            ${typeIcon} ${escapeHtml(asset.name)}
          </div>
          <span class="status-badge ${isWorking ? 'working' : 'faulty'}">
            <span class="status-dot"></span>
            <span>${isWorking ? 'Operational' : 'Faulty'}</span>
          </span>
        </div>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 8px; font-size: 0.82rem;">
          <div><span style="color: var(--text-muted);">Asset ID:</span> <strong class="asset-tag-code">${escapeHtml(asset.id)}</strong></div>
          <div><span style="color: var(--text-muted);">Course / Dept:</span> <strong style="color: #60a5fa;">${escapeHtml(asset.department)}</strong></div>
          <div><span style="color: var(--text-muted);">Location:</span> <strong>${escapeHtml(asset.location || 'Lab')}</strong></div>
          <div><span style="color: var(--text-muted);">Dead Stock / Serial:</span> <strong>${escapeHtml(asset.serialNumber || 'N/A')}</strong></div>
        </div>
      </div>

      <h4 style="font-size: 0.9rem; font-weight: 700; margin-bottom: 10px; color: var(--text-main);">
        Hardware Configuration &amp; Specifications:
      </h4>

      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 10px; margin-bottom: 16px;">
        ${specItems.map(item => `
          <div style="background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-md); padding: 10px;">
            <div style="font-size: 0.72rem; color: var(--text-muted); text-transform: uppercase;">Component</div>
            <div style="font-size: 0.85rem; font-weight: 600; color: var(--text-main); margin-top: 2px;">${escapeHtml(item)}</div>
          </div>
        `).join('')}
      </div>

      ${asset.faultDetails ? `
        <div style="background: rgba(239, 68, 68, 0.1); border: 1px solid rgba(239, 68, 68, 0.3); border-radius: var(--radius-md); padding: 12px; margin-top: 10px;">
          <div style="color: #ef4444; font-weight: 700; font-size: 0.85rem;">⚠️ Reported Malfunction:</div>
          <div style="color: var(--text-main); font-size: 0.82rem; margin-top: 4px;">${escapeHtml(asset.faultDetails.issue)}</div>
          <div style="color: var(--text-muted); font-size: 0.74rem; margin-top: 6px;">Reported: ${asset.faultDetails.reportedDate} &bull; Tech: ${asset.faultDetails.technician}</div>
        </div>
      ` : ''}
    `;
  }

  openModal('specsQuickModal');
}

// Modal Helpers
function openModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) modal.classList.add('active');
}

function closeModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) modal.classList.remove('active');
}

// Open Asset Add/Edit Modal (Popup)
function openAssetModal(asset) {
  const title = document.getElementById('assetModalTitle');
  const formAssetId = document.getElementById('formAssetId');
  const formAssetType = document.getElementById('formAssetType');
  const formAssetTag = document.getElementById('formAssetTag');
  const formAssetName = document.getElementById('formAssetName');
  const formDepartment = document.getElementById('formDepartment');
  const formLocation = document.getElementById('formLocation');
  const formAssignedUser = document.getElementById('formAssignedUser');
  const formIpAddress = document.getElementById('formIpAddress');
  const formSerialNumber = document.getElementById('formSerialNumber');
  const formStatus = document.getElementById('formStatus');
  const formSpecs = document.getElementById('formSpecs');
  const faultSection = document.getElementById('formFaultSection');
  const formFaultIssue = document.getElementById('formFaultIssue');

  if (asset) {
    if (title) title.textContent = `Edit Equipment: ${asset.id}`;
    if (formAssetId) formAssetId.value = asset.id;
    if (formAssetType) formAssetType.value = asset.type || 'Computer';
    if (formAssetTag) {
      formAssetTag.value = asset.id;
      formAssetTag.disabled = true;
    }
    if (formAssetName) formAssetName.value = asset.name || '';
    if (formDepartment) formDepartment.value = asset.department || departments[0];
    if (formLocation) formLocation.value = asset.location || '';
    if (formAssignedUser) formAssignedUser.value = asset.assignedUser || '';
    if (formIpAddress) formIpAddress.value = asset.ipAddress || '';
    if (formSerialNumber) formSerialNumber.value = asset.serialNumber || '';
    if (formStatus) formStatus.value = asset.status || 'Working';
    if (formSpecs) formSpecs.value = asset.specs || '';
    if (faultSection) {
      if (asset.status === 'Not Working' && asset.faultDetails) {
        faultSection.style.display = 'block';
        if (formFaultIssue) formFaultIssue.value = asset.faultDetails.issue || '';
      } else {
        faultSection.style.display = 'none';
        if (formFaultIssue) formFaultIssue.value = '';
      }
    }
  }

  openModal('assetModal');
}

// =========================================================================
// SPREADSHEET PREVIEW MODAL LOGIC (Clean Preview Only)
// =========================================================================
let currentPreviewAttachment = null;

// Helper to generate a complete hardware & dead-stock audit sheet for an asset
function generateDefaultAuditSheet(asset) {
  const rows = [
    ['Field / Parameter', 'Equipment Audit Specification'],
    ['Dead Stock Register No / Asset Tag', asset.id || '-'],
    ['Equipment Category', `${asset.type === 'Computer' ? 'संगणक (Computer / Laptop)' : asset.type === 'Printer' ? 'प्रिंटर (Printer / Copier)' : 'प्रोजेक्टर (Projector)'}`],
    ['Make / Brand & Model', asset.name || '-'],
    ['Department / Course Allocation', asset.department || 'General Administration'],
    ['Physical Location / Lab / Hall', asset.location || `${asset.department} Lab`],
    ['Assigned Custodian / Staff / Lab', asset.assignedUser || 'Department Faculty & Students'],
    ['Operational Working Status', asset.status || 'Working'],
    ['Serial Number / Barcode Tag', asset.serialNumber || asset.id || '-'],
    ['Hardware Specifications', asset.specs || 'Standard Lab Configuration'],
    ['IP Address / Network Config', asset.ipAddress || '192.168.1.100'],
    ['Dead Stock Entry Date', asset.purchaseDate || new Date().toISOString().split('T')[0]],
    ['Physical Audit Verification', 'Verified & Recorded in Register'],
    ['Maintenance / Remark Status', asset.faultDetails ? `Issue: ${asset.faultDetails.issue}` : 'Working Condition - No Faults']
  ];

  return {
    id: `audit-${asset.id}`,
    fileName: `${asset.id}_Hardware_Audit.xlsx`,
    fileSize: '14.5 KB',
    uploadDate: new Date().toISOString().replace('T', ' ').substring(0, 16),
    notes: `Official Equipment Audit Sheet for ${asset.id}`,
    sheetData: rows,
    isGeneratedDefault: true
  };
}

function openExcelModal(asset) {
  if (!asset) return;
  currentExcelTargetAsset = asset;
  const assetNameEl = document.getElementById('excelModalAssetName');
  const assetIdEl = document.getElementById('excelTargetAssetId');
  const deptEl = document.getElementById('excelTargetDept');
  const userEl = document.getElementById('excelTargetUser');
  const statusBadge = document.getElementById('excelTargetStatus');

  if (assetNameEl) assetNameEl.textContent = `${asset.name} (${asset.id})`;
  if (assetIdEl) assetIdEl.textContent = asset.id;
  if (deptEl) deptEl.textContent = asset.department;
  if (userEl) userEl.textContent = asset.assignedUser || 'Shared';
  
  if (statusBadge) {
    statusBadge.textContent = asset.status;
    statusBadge.className = `status-badge ${asset.status === 'Working' ? 'working' : 'faulty'}`;
  }

  const previewSection = document.getElementById('spreadsheetPreviewSection');
  const emptyNotice = document.getElementById('excelEmptyNotice');
  const fileTabs = document.getElementById('excelModalFileTabs');
  
  let attachments = (asset.excelAttachments && asset.excelAttachments.length > 0)
    ? [...asset.excelAttachments]
    : [generateDefaultAuditSheet(asset)];

  // If the first attachment has no sheetData (e.g. cloud link or empty), provide default audit sheet
  if (!attachments[0].sheetData || attachments[0].sheetData.length === 0) {
    attachments[0] = generateDefaultAuditSheet(asset);
  }

  if (emptyNotice) emptyNotice.style.display = 'none';
  if (previewSection) previewSection.style.display = 'block';

  // If multiple sheets attached, show switcher tabs
  if (fileTabs) {
    if (attachments.length > 1) {
      fileTabs.style.display = 'flex';
      fileTabs.innerHTML = '';
      attachments.forEach((att, idx) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = `btn btn-sm ${idx === 0 ? 'btn-excel active' : 'btn-secondary'}`;
        btn.style.fontSize = '0.76rem';
        btn.textContent = `📄 ${att.fileName}`;
        btn.addEventListener('click', () => {
          fileTabs.querySelectorAll('button').forEach(b => b.className = 'btn btn-sm btn-secondary');
          btn.className = 'btn btn-sm btn-excel active';
          previewSheetData(att, asset);
        });
        fileTabs.appendChild(btn);
      });
    } else {
      fileTabs.style.display = 'none';
    }
  }

  previewSheetData(attachments[0], asset);
  openModal('excelModal');
}

let currentSpreadsheetViewMode = 'vertical';

function previewSheetData(attachment, asset, forceMode) {
  if (forceMode) {
    currentSpreadsheetViewMode = forceMode;
  }

  if (!attachment || !attachment.sheetData || attachment.sheetData.length === 0) {
    if (asset || currentExcelTargetAsset) {
      attachment = generateDefaultAuditSheet(asset || currentExcelTargetAsset);
    }
  }
  currentPreviewAttachment = attachment;

  const previewSection = document.getElementById('spreadsheetPreviewSection');
  const fileNameSpan = document.getElementById('previewFileName');
  const sheetCountSpan = document.getElementById('previewSheetCount');
  const table = document.getElementById('previewSpreadsheetTable');
  const emptyNotice = document.getElementById('excelEmptyNotice');
  const toggleGroup = document.getElementById('previewViewToggleGroup');
  const btnVertical = document.getElementById('btnPreviewVerticalView');
  const btnRawRow = document.getElementById('btnPreviewRawRowView');

  if (!attachment || !attachment.sheetData || attachment.sheetData.length === 0) {
    if (previewSection) previewSection.style.display = 'none';
    if (emptyNotice) emptyNotice.style.display = 'block';
    return;
  }

  if (emptyNotice) emptyNotice.style.display = 'none';
  if (previewSection) previewSection.style.display = 'block';
  if (fileNameSpan) fileNameSpan.textContent = attachment.fileName || `${(asset || currentExcelTargetAsset)?.id || 'Equipment'}_Audit.xlsx`;

  const rows = attachment.sheetData;
  if (!table) return;
  table.innerHTML = '';

  // Determine shape of rows
  const isSingleRowWideRecord = (rows.length === 2 && Array.isArray(rows[0]) && rows[0].length > 2);
  const isAlreadyTwoColumn = (rows.length >= 2 && Array.isArray(rows[0]) && rows[0].length === 2);

  if (isSingleRowWideRecord) {
    if (toggleGroup) toggleGroup.style.display = 'inline-flex';
    if (btnVertical && btnRawRow) {
      if (currentSpreadsheetViewMode === 'vertical') {
        btnVertical.classList.add('active');
        btnRawRow.classList.remove('active');
      } else {
        btnVertical.classList.remove('active');
        btnRawRow.classList.add('active');
      }
    }

    if (currentSpreadsheetViewMode === 'vertical') {
      // 2-column Attribute-Value View
      if (sheetCountSpan) sheetCountSpan.textContent = `(${rows[0].length} Parameters)`;

      const thead = document.createElement('thead');
      thead.innerHTML = `
        <tr>
          <th style="width: 34%; min-width: 180px;">घटक / फील्ड (Parameter / Field)</th>
          <th>नोंदणी तपशील (Audit Value Details)</th>
        </tr>
      `;
      table.appendChild(thead);

      const tbody = document.createElement('tbody');
      const headers = rows[0];
      const dataRow = rows[1] || [];

      headers.forEach((header, colIdx) => {
        const val = dataRow[colIdx];
        const valStr = (val !== null && val !== undefined && String(val).trim() !== '') ? String(val).trim() : '—';
        const tr = document.createElement('tr');

        let displayValHtml = escapeHtml(valStr);
        const headerLower = String(header).toLowerCase();

        // Highlight status or empty values cleanly
        if (valStr === '—') {
          displayValHtml = `<span style="color: var(--text-dim); font-style: italic;">— रिकामे (Empty) —</span>`;
        } else if (headerLower.includes('status') || headerLower.includes('स्थिती')) {
          const isWorking = /working|सुरू/i.test(valStr);
          displayValHtml = `<span class="status-badge ${isWorking ? 'working' : 'faulty'}">${escapeHtml(valStr)}</span>`;
        } else if (headerLower.includes('dead stock') || headerLower.includes('asset')) {
          displayValHtml = `<strong class="asset-tag-code">${escapeHtml(valStr)}</strong>`;
        } else if (headerLower.includes('make') || headerLower.includes('कंपनी') || headerLower.includes('cpu') || headerLower.includes('ram')) {
          displayValHtml = `<strong style="color: var(--text-main);">${escapeHtml(valStr)}</strong>`;
        }

        tr.innerHTML = `
          <td class="spec-col-label"><strong>${escapeHtml(String(header || `Column ${colIdx + 1}`))}</strong></td>
          <td class="spec-col-value">${displayValHtml}</td>
        `;
        tbody.appendChild(tr);
      });

      table.appendChild(tbody);
      return;
    } else {
      // Raw horizontal row view
      if (sheetCountSpan) sheetCountSpan.textContent = `(1 Excel Row, ${rows[0].length} Columns)`;
      const thead = document.createElement('thead');
      const tbody = document.createElement('tbody');

      const trHead = document.createElement('tr');
      rows[0].forEach(cell => {
        const th = document.createElement('th');
        th.textContent = cell !== null && cell !== undefined ? String(cell) : '';
        trHead.appendChild(th);
      });
      thead.appendChild(trHead);

      const trBody = document.createElement('tr');
      (rows[1] || []).forEach(cell => {
        const td = document.createElement('td');
        td.textContent = cell !== null && cell !== undefined ? String(cell) : '';
        trBody.appendChild(td);
      });
      tbody.appendChild(trBody);

      table.appendChild(thead);
      table.appendChild(tbody);
      return;
    }
  }

  // If already 2 columns (e.g. standard audit sheet)
  if (toggleGroup) toggleGroup.style.display = 'none';

  if (isAlreadyTwoColumn) {
    if (sheetCountSpan) sheetCountSpan.textContent = `(${rows.length - 1} Parameters)`;
    const thead = document.createElement('thead');
    thead.innerHTML = `
      <tr>
        <th style="width: 34%; min-width: 180px;">${escapeHtml(String(rows[0][0] || 'घटक / फील्ड'))}</th>
        <th>${escapeHtml(String(rows[0][1] || 'नोंदणी तपशील'))}</th>
      </tr>
    `;
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    for (let r = 1; r < rows.length; r++) {
      const tr = document.createElement('tr');
      const paramName = String(rows[r][0] || '');
      const paramVal = String(rows[r][1] || '');
      let valHtml = escapeHtml(paramVal);
      if (/status|स्थिती/i.test(paramName)) {
        valHtml = `<span class="status-badge ${/working/i.test(paramVal) ? 'working' : 'faulty'}">${escapeHtml(paramVal)}</span>`;
      }
      tr.innerHTML = `
        <td class="spec-col-label"><strong>${escapeHtml(paramName)}</strong></td>
        <td class="spec-col-value">${valHtml}</td>
      `;
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    return;
  }

  // Multi-row, multi-column sheet
  if (sheetCountSpan) sheetCountSpan.textContent = `(${rows.length > 1 ? rows.length - 1 : rows.length} Rows, ${rows[0]?.length || 0} Columns)`;
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');

  rows.forEach((row, rowIdx) => {
    const tr = document.createElement('tr');
    const cells = Array.isArray(row) ? row : Object.values(row);
    cells.forEach(cell => {
      const el = document.createElement(rowIdx === 0 ? 'th' : 'td');
      el.textContent = cell !== null && cell !== undefined ? String(cell) : '';
      tr.appendChild(el);
    });

    if (rowIdx === 0) {
      thead.appendChild(tr);
    } else {
      tbody.appendChild(tr);
    }
  });

  table.appendChild(thead);
  if (tbody.children.length > 0) {
    table.appendChild(tbody);
  }
}

function downloadAttachedExcel(attachment, asset) {
  if (typeof XLSX === 'undefined') {
    alert('Excel exporter library is not loaded');
    return;
  }
  const sheetData = (attachment && attachment.sheetData && attachment.sheetData.length > 0)
    ? attachment.sheetData
    : generateDefaultAuditSheet(asset).sheetData;

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  XLSX.utils.book_append_sheet(wb, ws, 'AuditSheet');
  XLSX.writeFile(wb, (attachment && attachment.fileName) || `${asset.id}_audit.xlsx`);
}

// Setup Download & Delete actions on the Spreadsheet Preview Modal
function setupExcelAndDriveUpload() {
  const btnVertical = document.getElementById('btnPreviewVerticalView');
  const btnRawRow = document.getElementById('btnPreviewRawRowView');
  if (btnVertical) {
    btnVertical.addEventListener('click', () => {
      currentSpreadsheetViewMode = 'vertical';
      if (currentPreviewAttachment) {
        previewSheetData(currentPreviewAttachment, currentExcelTargetAsset, 'vertical');
      }
    });
  }
  if (btnRawRow) {
    btnRawRow.addEventListener('click', () => {
      currentSpreadsheetViewMode = 'horizontal';
      if (currentPreviewAttachment) {
        previewSheetData(currentPreviewAttachment, currentExcelTargetAsset, 'horizontal');
      }
    });
  }

  const btnDownloadCurrentPreview = document.getElementById('btnDownloadCurrentPreview');
  if (btnDownloadCurrentPreview) {
    btnDownloadCurrentPreview.addEventListener('click', () => {
      if (currentPreviewAttachment && currentExcelTargetAsset) {
        downloadAttachedExcel(currentPreviewAttachment, currentExcelTargetAsset);
      } else if (currentExcelTargetAsset) {
        const def = generateDefaultAuditSheet(currentExcelTargetAsset);
        downloadAttachedExcel(def, currentExcelTargetAsset);
      } else {
        alert('No previewable sheet to download.');
      }
    });
  }

  const btnDeleteCurrentPreview = document.getElementById('btnDeleteCurrentPreview');
  if (btnDeleteCurrentPreview) {
    btnDeleteCurrentPreview.addEventListener('click', () => {
      if (!currentExcelTargetAsset) return;

      if (!currentPreviewAttachment || currentPreviewAttachment.isGeneratedDefault) {
        showToast('ℹ️ Audit sheet is already displaying the default register specifications.');
        return;
      }

      if (confirm(`Do you want to permanently delete "${currentPreviewAttachment.fileName}" from ${currentExcelTargetAsset.id}?`)) {
        currentExcelTargetAsset.excelAttachments = (currentExcelTargetAsset.excelAttachments || []).filter(a => a.id !== currentPreviewAttachment.id);
        saveData();
        renderAll();
        openExcelModal(currentExcelTargetAsset);
        showToast(`🗑️ "${currentPreviewAttachment.fileName}" deleted`);
      }
    });
  }
}

// =========================================================================
// BULK EXCEL IMPORT & CENTRAL EXCEL FILE MANAGER
// =========================================================================
let pendingBulkImportData = null;

function setupBulkExcelImportAndManager() {
  const tabImportUpload = document.getElementById('tabImportUpload');
  const tabImportManage = document.getElementById('tabImportManage');
  const panelImportUpload = document.getElementById('panelImportUpload');
  const panelImportManage = document.getElementById('panelImportManage');

  const bulkImportFileInput = document.getElementById('bulkImportFileInput');
  const bulkImportDropzone = document.getElementById('bulkImportDropzone');
  const btnBrowseBulkExcel = document.getElementById('btnBrowseBulkExcel');
  const btnConfirmBulkImport = document.getElementById('btnConfirmBulkImport');

  const bulkCloudDriveUrl = document.getElementById('bulkCloudDriveUrl');
  const btnAttachBulkCloudLink = document.getElementById('btnAttachBulkCloudLink');

  const searchGlobalExcelInput = document.getElementById('searchGlobalExcelInput');
  const btnDeleteAllSystemExcel = document.getElementById('btnDeleteAllSystemExcel');

  // Sub-tabs in modal
  if (tabImportUpload && tabImportManage && panelImportUpload && panelImportManage) {
    tabImportUpload.addEventListener('click', () => {
      tabImportUpload.className = 'btn btn-sm btn-excel active';
      tabImportManage.className = 'btn btn-sm btn-secondary';
      panelImportUpload.style.display = 'block';
      panelImportManage.style.display = 'none';
      if (btnConfirmBulkImport) {
        btnConfirmBulkImport.style.display = 'inline-block';
      }
    });

    tabImportManage.addEventListener('click', () => {
      tabImportManage.className = 'btn btn-sm btn-excel active';
      tabImportUpload.className = 'btn btn-sm btn-secondary';
      panelImportUpload.style.display = 'none';
      panelImportManage.style.display = 'block';
      if (btnConfirmBulkImport) {
        btnConfirmBulkImport.style.display = 'none';
      }
      renderGlobalExcelTable();
    });
  }

  // File Picker Trigger
  if (btnBrowseBulkExcel && bulkImportFileInput) {
    btnBrowseBulkExcel.addEventListener('click', (e) => {
      e.stopPropagation();
      bulkImportFileInput.click();
    });
  }

  if (bulkImportFileInput) {
    bulkImportFileInput.addEventListener('click', (e) => e.stopPropagation());
    bulkImportFileInput.addEventListener('change', (e) => {
      const file = e.target.files && e.target.files[0];
      if (file) {
        parseBulkExcelFile(file);
        bulkImportFileInput.value = '';
      }
    });
  }

  if (bulkImportDropzone && bulkImportFileInput) {
    bulkImportDropzone.addEventListener('click', (e) => {
      if (e.target !== bulkImportFileInput && e.target !== btnBrowseBulkExcel && !btnBrowseBulkExcel?.contains(e.target)) {
        bulkImportFileInput.click();
      }
    });

    bulkImportDropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      bulkImportDropzone.classList.add('dragover');
    });

    bulkImportDropzone.addEventListener('dragleave', () => bulkImportDropzone.classList.remove('dragover'));

    bulkImportDropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      bulkImportDropzone.classList.remove('dragover');
      const file = e.dataTransfer.files && e.dataTransfer.files[0];
      if (file) {
        parseBulkExcelFile(file);
      }
    });
  }

  // Confirm Import Handler (Supports both footer button and direct preview button)
  const handleConfirmImport = () => {
    if (!pendingBulkImportData) {
      alert('कृपया प्रथम एक योग्य एक्सेल फाइल निवडा (Please choose an Excel file first).');
      return;
    }
    const hasItems = (pendingBulkImportData.parsedItems && pendingBulkImportData.parsedItems.length > 0) ||
                     (pendingBulkImportData.dataRows && pendingBulkImportData.dataRows.length > 0);
    if (!hasItems) {
      alert('या एक्सेल फाईलमध्ये कोणतीही उपकरणे आढळली नाहीत (No equipment data rows found).');
      return;
    }
    importParsedRows(pendingBulkImportData);
  };

  if (btnConfirmBulkImport) {
    btnConfirmBulkImport.addEventListener('click', handleConfirmImport);
  }

  const btnConfirmStagedImportDirect = document.getElementById('btnConfirmStagedImportDirect');
  if (btnConfirmStagedImportDirect) {
    btnConfirmStagedImportDirect.addEventListener('click', handleConfirmImport);
  }

  // Search in Central Excel File Manager
  if (searchGlobalExcelInput) {
    searchGlobalExcelInput.addEventListener('input', () => {
      renderGlobalExcelTable(searchGlobalExcelInput.value.trim().toLowerCase());
    });
  }

  // Delete ALL Excel attachments (Clear Excel files count to 0)
  if (btnDeleteAllSystemExcel) {
    btnDeleteAllSystemExcel.addEventListener('click', () => {
      deleteAllExcelAttachments();
    });
  }

  // Delete Added Excel File and Erase All Data Button
  const btnDeleteAddedExcelAndEraseData = document.getElementById('btnDeleteAddedExcelAndEraseData');
  if (btnDeleteAddedExcelAndEraseData) {
    btnDeleteAddedExcelAndEraseData.addEventListener('click', () => {
      eraseAllInventoryData();
      closeModal('importModal');
    });
  }

  // Cancel Staged File Import Button
  const btnCancelStagedImport = document.getElementById('btnCancelStagedImport');
  if (btnCancelStagedImport) {
    btnCancelStagedImport.addEventListener('click', () => {
      pendingBulkImportData = null;
      const previewDiv = document.getElementById('bulkImportPreview');
      if (previewDiv) previewDiv.style.display = 'none';
      if (bulkImportFileInput) bulkImportFileInput.value = '';
      if (btnConfirmBulkImport) {
        btnConfirmBulkImport.disabled = true;
        btnConfirmBulkImport.textContent = '💾 यादीमध्ये इंपोर्ट करा (Confirm & Import)';
      }
      showToast('निवडलेली एक्सेल फाईल रद्द करण्यात आली (Selected file removed)');
    });
  }

  // Erase ALL inventory data (Reset everything to 0)
  const btnEraseAllInventoryData = document.getElementById('btnEraseAllInventoryData');
  if (btnEraseAllInventoryData) {
    btnEraseAllInventoryData.addEventListener('click', () => {
      eraseAllInventoryData();
    });
  }

  // Restore default sample data (580 items)
  const btnRestoreDefaultData = document.getElementById('btnRestoreDefaultData');
  if (btnRestoreDefaultData) {
    btnRestoreDefaultData.addEventListener('click', () => {
      restoreDefaultData();
    });
  }

  updateImportModalActiveFile();
}

// Update Active Imported Excel File Card in Import Modal
function updateImportModalActiveFile() {
  const activeSection = document.getElementById('activeImportedFileSection');
  const noActiveSection = document.getElementById('noActiveImportedFileSection');
  const fileNameBadge = document.getElementById('activeImportedFileNameBadge');
  const totalCountSpan = document.getElementById('activeImportedTotalCount');
  const compCountSpan = document.getElementById('activeImportedComputersCount');
  const prnCountSpan = document.getElementById('activeImportedPrintersCount');
  const prjCountSpan = document.getElementById('activeImportedProjectorsCount');

  if (!activeSection || !noActiveSection) return;

  if (assets.length > 0) {
    activeSection.style.display = 'block';
    noActiveSection.style.display = 'none';

    let currentFile = localStorage.getItem('assetpulse_active_excel_filename');
    if (!currentFile) {
      const firstWithFile = assets.find(a => a.excelAttachments && a.excelAttachments.length > 0);
      currentFile = firstWithFile?.excelAttachments[0]?.fileName || 'Revised_Inventory_wdn_updated_05-05-2026.xlsx';
    }

    if (fileNameBadge) fileNameBadge.textContent = currentFile;
    if (totalCountSpan) totalCountSpan.textContent = assets.length;

    const comps = assets.filter(a => a.type === 'Computer').length;
    const prns = assets.filter(a => a.type === 'Printer').length;
    const prjs = assets.filter(a => a.type === 'Projector').length;

    if (compCountSpan) compCountSpan.textContent = `${comps} संगणक`;
    if (prnCountSpan) prnCountSpan.textContent = `${prns} प्रिंटर्स`;
    if (prjCountSpan) prjCountSpan.textContent = `${prjs} प्रोजेक्टर्स`;
  } else {
    activeSection.style.display = 'none';
    noActiveSection.style.display = 'block';
  }
}

function deleteAllExcelAttachments() {
  let totalFiles = 0;
  assets.forEach(a => {
    totalFiles += (a.excelAttachments || []).length;
  });

  if (totalFiles === 0) {
    showToast('कोणतीही संलग्न एक्सेल फाइल अस्तित्वात नाही (No attached Excel files).');
    return;
  }

  if (confirm(`⚠️ सावधान (CONFIRMATION):\n\nसर्व उपकरणांच्या मिळून एकूण ${totalFiles} एक्सेल फाइल्स डिलीट करायच्या आहेत का?\nउपकरणांची संख्या तशीच राहील, फक्त एक्सेल ऑडिट फाइल्सची संख्या शून्य (0) होईल.\n\n(Remove all ${totalFiles} attached Excel sheets? Equipment count remains, only Excel audits become 0.)`)) {
    assets.forEach(a => {
      a.excelAttachments = [];
    });
    saveData();
    renderAll();
    renderGlobalExcelTable();
    showToast(`🗑️ सर्व ${totalFiles} एक्सेल फाइल्स हटवल्या - संख्या 0 झाली (All Excel attachments deleted)`);
  }
}

// Helper to find true table header row (skipping banner/title rows)
function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const row = rows[i] || [];
    const filled = row.filter(c => c !== null && c !== undefined && String(c).trim() !== '');
    if (filled.length >= 3) {
      const str = row.map(c => String(c || '').toLowerCase()).join(' ');
      if (
        str.includes('sr') ||
        str.includes('dept') ||
        str.includes('department') ||
        str.includes('make') ||
        str.includes('dead') ||
        str.includes('asset') ||
        str.includes('model') ||
        str.includes('serial') ||
        str.includes('printer') ||
        str.includes('projector')
      ) {
        return i;
      }
    }
  }
  return 0;
}

// Normalize Course / Department names
function normalizeDeptName(rawDept) {
  if (!rawDept) return 'General Administration';
  let d = String(rawDept).trim();
  const lower = d.toLowerCase();
  if (lower === 'bcs' || (lower.includes('bcs') && !lower.includes('old') && !lower.includes('stat') && !lower.includes('hod'))) return 'BCS';
  if (lower.includes('bcs') && lower.includes('old')) return 'BCS Old PC';
  if (lower.includes('bcs') && lower.includes('hod')) return 'BCS (HOD Office)';
  if (lower.includes('bcs') && lower.includes('stat')) return 'BCS (Stat & Electronics)';
  if (lower === 'language' || lower.includes('language lab')) return 'Language Lab';
  if (lower === 'stat & electronic' || lower === 'stat & electronics') return 'Stat & Electronics';
  if (lower === 'maths' || lower === 'mathematics') return 'Mathematics';
  if (lower.includes('library')) return 'Library';
  if (lower.includes('microbiolology') || lower.includes('microbiology')) return 'Microbiology';
  if (lower.includes('history')) return 'History';
  if (lower === 'srpd') return 'SRPD';
  if (lower.includes('srpd old')) return 'SRPD Old PC';
  if (lower.includes('physics bsc (pg)') || lower.includes('physics pg dept') || lower.includes('pg lab (phy,chem)')) return 'PG Lab (Physics & Chemistry)';
  if (lower === 'physics bsc' || lower.includes('physics bsc') || lower.includes('physics sr wing')) return 'Physics BSc (Sr Wing)';
  if (lower.includes('jr. physics') || lower.includes('physics jr')) return 'Jr. Physics';
  if (lower.includes('bsc cs old')) return 'BSc CS Old PC';
  if (lower.includes('bsc cs (optional)')) return 'BSc CS (Optional)';
  if (lower === 'bsc cs' || lower.includes('bsc cs')) return 'BSc CS';
  if (lower === 'b.sc it + msc  cs' || lower.includes('b.sc it + msc cs')) return 'B.Sc IT + MSc CS';
  if (lower === 'msc cs' || lower === 'msc  cs') return 'MSc CS';
  if (lower === 'msc electronics') return 'MSc Electronics';
  if (lower === 'b.sc it' || lower === 'b.sc it ') return 'B.Sc IT';
  if (lower === 'electronics m.sc') return 'MSc Electronics';
  if (lower === 'electronics' || lower.includes('electronics sr wing') || lower.includes('electronics dept')) return 'Electronics';
  if (lower.includes('geography')) return 'Geography';
  if (lower === 'stat' || lower === 'statistics') return 'Statistics';
  if (lower.includes('website')) return 'Website';
  if (lower.includes('zoology')) return 'Zoology';
  if (lower.includes('botany')) return 'Botany';
  if (lower.includes('principal')) return 'Principal Sir Cabin';
  if (lower.includes('naac')) return 'NAAC';
  if (lower.includes('ncc')) return 'NCC';
  if (lower === 'office') return 'Office';
  if (lower === 'ng office') return 'NG Office';
  if (lower === 'erp') return 'ERP';
  if (lower.includes('girls hostel')) return 'Girls Hostel';
  if (lower === 'hindi') return 'Hindi';
  if (lower === 'english') return 'English';
  if (lower.includes('yogshikshak')) return 'Diploma in Yogshikshak';
  if (lower === 'marathi') return 'Marathi';
  if (lower === 'sro') return 'SRO';
  if (lower.includes('sangli jilha') || lower.includes('rojgar')) return 'Sangli Jilha Rojgar Srujan Kendra';
  if (lower.includes('exam 18')) return 'Exam 18 Room';
  if (lower.includes('economics')) return 'Economics';
  if (lower.includes('sociology')) return 'Sociology';
  if (lower.includes('psychology')) return 'Psychology';
  return d;
}

// Parse Bulk Excel Spreadsheet (Multi-Sheet: Computers, Printers, Projectors, Index)
function parseBulkExcelFile(file) {
  if (typeof XLSX === 'undefined') {
    alert('Excel parser library (xlsx.full.min.js) is not loaded.');
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
        alert('The uploaded Excel workbook contains no readable sheets.');
        return;
      }

      const allParsedItems = [];
      const discoveredDepts = new Set();

      // Check Index / Departments Sheet first to pre-seed course names
      workbook.SheetNames.forEach(sheetName => {
        const sLower = sheetName.toLowerCase().trim();
        if (sLower.includes('index') || sLower.includes('summary')) {
          const ws = workbook.Sheets[sheetName];
          const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
          const hIdx = findHeaderRow(raw);
          const headers = (raw[hIdx] || []).map(h => String(h || '').trim());
          const dCol = headers.findIndex(h => h.toLowerCase().includes('dept') || h.toLowerCase().includes('course') || h.toLowerCase().includes('name'));
          const colIdx = dCol >= 0 ? dCol : 1;
          for (let r = hIdx + 1; r < raw.length; r++) {
            if (raw[r] && raw[r][colIdx]) {
              const val = String(raw[r][colIdx]).trim();
              if (val && !val.toLowerCase().includes('total') && !val.toLowerCase().includes('deccan') && isNaN(Number(val))) {
                discoveredDepts.add(normalizeDeptName(val));
              }
            }
          }
        }
      });

      // Parse each sheet for Computers, Printers, Projectors
      workbook.SheetNames.forEach(sheetName => {
        const sLower = sheetName.toLowerCase().trim();
        if (sLower.includes('index')) return; // handled above

        let defaultCat = 'Computer';
        if (sLower.includes('print')) defaultCat = 'Printer';
        else if (sLower.includes('proj')) defaultCat = 'Projector';

        const ws = workbook.Sheets[sheetName];
        const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1 });
        if (!rawRows || rawRows.length < 2) return;

        const hIdx = findHeaderRow(rawRows);
        const headers = (rawRows[hIdx] || []).map(h => String(h || '').trim());

        const findCol = (terms) => headers.findIndex(h => terms.some(t => h.toLowerCase().includes(t)));
        const idIdx = findCol(['dead stock', 'printer asset', 'serial number', 'asset id', 'tag', 'no']);
        const deptIdx = findCol(['department', 'dept.', 'dept', 'course', 'lab', 'room']);
        const makeIdx = findCol(['make', 'company', 'name', 'model', 'brand']);
        const typeIdx = findCol(['type', 'category']);
        const cpuIdx = findCol(['cpu', 'processor', 'c.p.u.']);
        const ramIdx = findCol(['ram', 'memory']);
        const hddIdx = findCol(['hdd', 'hard disk']);
        const ssdIdx = findCol(['ssd']);
        const osIdx = findCol(['o. s.', 'os', 'operating system']);
        const monitorIdx = findCol(['monitor', 'display']);
        const statusIdx = findCol(['status', 'condition']);
        const remarkIdx = findCol(['remark', 'port', 'issue', 'fault']);

        const dataRows = rawRows.slice(hIdx + 1).filter(r => r && r.some(c => c !== null && c !== undefined && String(c).trim() !== ''));

        dataRows.forEach((row, idx) => {
          const firstCell = String(row[0] || '').toLowerCase().trim();
          if (firstCell.includes('total') || firstCell.includes('deccan') || firstCell.includes('will') || firstCell.includes('sr.')) return;

          // Department / Course extraction
          let rawDept = deptIdx >= 0 && row[deptIdx] ? String(row[deptIdx]).trim() : '';
          if (!rawDept || rawDept === 'null' || rawDept === 'undefined' || rawDept === '-') {
            rawDept = 'General Administration';
          }
          const dept = normalizeDeptName(rawDept);
          discoveredDepts.add(dept);

          // Equipment Category Detection
          let cat = defaultCat;
          const rawType = typeIdx >= 0 && row[typeIdx] ? String(row[typeIdx]).trim().toLowerCase() : '';
          const rawMake = makeIdx >= 0 && row[makeIdx] ? String(row[makeIdx]).trim() : '';

          if (cat === 'Computer') {
            if (rawType.includes('print') || rawMake.toLowerCase().includes('laserjet') || rawMake.toLowerCase().includes('canon lbp') || rawMake.toLowerCase().includes('epson')) {
              cat = 'Printer';
            } else if (rawType.includes('proj') || rawMake.toLowerCase().includes('benq') || rawMake.toLowerCase().includes('dlp') || rawMake.toLowerCase().includes('projector')) {
              cat = 'Projector';
            }
          }

          // Asset ID Detection
          let rawId = idIdx >= 0 && row[idIdx] ? String(row[idIdx]).trim() : '';
          const srNo = row[0] && !isNaN(Number(row[0])) ? Number(row[0]) : (idx + 1);
          const prefix = cat === 'Computer' ? 'WCS-PC' : cat === 'Printer' ? 'WCS-PRN' : 'WCS-PRJ';
          let assetId = rawId;
          if (!assetId || assetId === 'null' || assetId === 'undefined' || !isNaN(Number(assetId))) {
            assetId = `${prefix}-${srNo}`;
          }

          const itemName = rawMake || `${cat} Item ${srNo}`;

          // Status & Operational condition
          const rawStatus = statusIdx >= 0 && row[statusIdx] ? String(row[statusIdx]).trim().toLowerCase() : 'working';
          const isWorking = !rawStatus.includes('not') && !rawStatus.includes('fault') && !rawStatus.includes('dead') && !rawStatus.includes('बंद') && !rawStatus.includes('issue');

          // Assemble detailed specifications
          let specs = '';
          if (cat === 'Computer') {
            const parts = [];
            if (cpuIdx >= 0 && row[cpuIdx]) parts.push(String(row[cpuIdx]).trim());
            if (ramIdx >= 0 && row[ramIdx]) parts.push(String(row[ramIdx]).trim());
            if (hddIdx >= 0 && row[hddIdx]) parts.push(String(row[hddIdx]).trim());
            if (ssdIdx >= 0 && row[ssdIdx]) parts.push(String(row[ssdIdx]).trim());
            if (osIdx >= 0 && row[osIdx]) parts.push(String(row[osIdx]).trim());
            if (monitorIdx >= 0 && row[monitorIdx]) parts.push('Monitor: ' + String(row[monitorIdx]).trim());
            specs = parts.filter(Boolean).join(', ') || 'Standard Lab Configuration';
          } else if (cat === 'Printer') {
            const parts = [];
            if (rawType) parts.push('Type: ' + rawType);
            if (remarkIdx >= 0 && row[remarkIdx]) parts.push(String(row[remarkIdx]).trim());
            specs = parts.filter(Boolean).join(', ') || 'Standard Laser Printer';
          } else if (cat === 'Projector') {
            const parts = [];
            if (rawMake) parts.push('Model: ' + rawMake);
            if (remarkIdx >= 0 && row[remarkIdx]) parts.push(String(row[remarkIdx]).trim());
            specs = parts.filter(Boolean).join(', ') || 'DLP Projector System';
          }

          allParsedItems.push({
            id: assetId,
            type: cat,
            name: itemName,
            department: dept,
            location: `${dept} Lab / Room`,
            assignedUser: `${dept} Faculty & Students`,
            status: isWorking ? 'Working' : 'Not Working',
            serialNumber: assetId,
            purchaseDate: new Date().toISOString().split('T')[0],
            specs: specs,
            ipAddress: `192.168.1.${(allParsedItems.length % 250) + 1}`,
            faultDetails: !isWorking ? {
              issue: (remarkIdx >= 0 && row[remarkIdx]) ? String(row[remarkIdx]).trim() : 'Reported in Excel Audit Sheet',
              reportedDate: new Date().toISOString().split('T')[0],
              reportedBy: 'Excel Stock Audit',
              priority: 'High',
              technician: 'Maintenance Team'
            } : null,
            excelAttachments: [
              {
                id: `att-${Date.now()}-${idx}`,
                fileName: file.name,
                fileSize: `${Math.round(JSON.stringify(row).length / 1024 * 10) / 10} KB`,
                uploadDate: new Date().toISOString().replace('T', ' ').substring(0, 16),
                notes: `Stock Audit Sheet: ${sheetName} (Row ${srNo})`,
                sheetData: [headers, row]
              }
            ]
          });
        });
      });

      if (allParsedItems.length === 0) {
        alert('Could not find any equipment data in the uploaded Excel file. Please ensure it has valid columns.');
        return;
      }

      const compCount = allParsedItems.filter(i => i.type === 'Computer').length;
      const prnCount = allParsedItems.filter(i => i.type === 'Printer').length;
      const prjCount = allParsedItems.filter(i => i.type === 'Projector').length;

      pendingBulkImportData = {
        fileName: file.name,
        parsedItems: allParsedItems,
        discoveredDepts: Array.from(discoveredDepts).sort(),
        counts: { total: allParsedItems.length, computers: compCount, printers: prnCount, projectors: prjCount, depts: discoveredDepts.size }
      };

      // Render Preview
      const previewDiv = document.getElementById('bulkImportPreview');
      const foundCount = document.getElementById('bulkFoundCount');
      const previewFileName = document.getElementById('bulkPreviewFileName');
      const previewList = document.getElementById('bulkPreviewList');
      const btnConfirm = document.getElementById('btnConfirmBulkImport');

      if (previewDiv) previewDiv.style.display = 'block';
      if (foundCount) foundCount.textContent = allParsedItems.length;
      if (previewFileName) {
        previewFileName.textContent = `${file.name} (${compCount} Computers, ${prnCount} Printers, ${prjCount} Projectors across ${discoveredDepts.size} Courses)`;
      }

      if (previewList) {
        previewList.innerHTML = '';
        // Sample preview from each category
        const sampleComputers = allParsedItems.filter(i => i.type === 'Computer').slice(0, 3);
        const samplePrinters = allParsedItems.filter(i => i.type === 'Printer').slice(0, 2);
        const sampleProjectors = allParsedItems.filter(i => i.type === 'Projector').slice(0, 2);
        const sampleItems = [...sampleComputers, ...samplePrinters, ...sampleProjectors];

        sampleItems.forEach(item => {
          const li = document.createElement('li');
          const typeIcon = item.type === 'Computer' ? '💻' : item.type === 'Printer' ? '🖨️' : '📽️';
          li.innerHTML = `<strong>${typeIcon} [${item.type}]</strong> ${escapeHtml(item.name)} &bull; <strong>Course/Dept:</strong> <span style="color:#60a5fa;">${escapeHtml(item.department)}</span> &bull; <em>${escapeHtml(item.specs.substring(0, 60))}</em>`;
          previewList.appendChild(li);
        });

        if (allParsedItems.length > sampleItems.length) {
          const li = document.createElement('li');
          li.style.fontStyle = 'italic';
          li.style.color = '#94a3b8';
          li.textContent = `... आणि इतर ${allParsedItems.length - sampleItems.length} उपकरणे (एकूण ${discoveredDepts.size} कोर्सेस / विभाग)`;
          previewList.appendChild(li);
        }
      }

      if (btnConfirm) {
        btnConfirm.disabled = false;
        btnConfirm.textContent = `💾 सर्व ${allParsedItems.length} उपकरणे इंपोर्ट करा (Import ${allParsedItems.length} Items)`;
      }

      showToast(`📊 "${file.name}" मधून ${allParsedItems.length} उपकरणे (संगणक: ${compCount}, प्रिंटर: ${prnCount}, प्रोजेक्टर: ${prjCount}) लोड केली!`);
    } catch (err) {
      console.error(err);
      alert('Error reading Excel spreadsheet: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// Import parsed rows into live inventory
function importParsedRows(pendingData) {
  if (!pendingData || !pendingData.parsedItems || pendingData.parsedItems.length === 0) return;

  const { fileName, parsedItems, discoveredDepts, counts } = pendingData;

  // Replace inventory with newly parsed complete stock dataset
  assets = parsedItems;
  departments = discoveredDepts;

  localStorage.setItem('assetpulse_active_excel_filename', fileName);
  localStorage.removeItem(STORAGE_KEY_CLEARED);
  saveData();
  renderAll();
  updateImportModalActiveFile();
  closeModal('importModal');
  pendingBulkImportData = null;

  const btnConfirm = document.getElementById('btnConfirmBulkImport');
  if (btnConfirm) btnConfirm.disabled = true;

  showToast(`✅ ${counts.total} उपकरणे (संगणक: ${counts.computers}, प्रिंटर्स: ${counts.printers}, प्रोजेक्टर: ${counts.projectors}) व ${counts.depts} कोर्सेस यशस्वीरित्या इंपोर्ट केली!`);
}

// Render Central All-Excel Files Table
function renderGlobalExcelTable(searchQuery = '') {
  const tbody = document.getElementById('globalExcelTableBody');
  const badge = document.getElementById('allExcelFilesBadge');
  if (!tbody) return;
  tbody.innerHTML = '';

  const allItems = [];
  assets.forEach(asset => {
    (asset.excelAttachments || []).forEach((att, attIdx) => {
      allItems.push({ asset, att, attIdx });
    });
  });

  if (badge) badge.textContent = allItems.length;

  const filtered = allItems.filter(item => {
    if (!searchQuery) return true;
    const q = searchQuery.toLowerCase();
    return (
      item.asset.id.toLowerCase().includes(q) ||
      item.asset.name.toLowerCase().includes(q) ||
      item.asset.department.toLowerCase().includes(q) ||
      item.att.fileName.toLowerCase().includes(q) ||
      (item.att.service || '').toLowerCase().includes(q)
    );
  });

  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="5" style="text-align: center; color: var(--text-muted); padding: 24px;">
          ${allItems.length === 0 ? 'या प्रणालीमध्ये कोणतीही संलग्न एक्सेल फाइल किंवा ड्राइव्ह लिंक नाही (No Excel files attached).' : 'शोधाप्रमाणे कोणतीही फाइल आढळली नाही (No matching files found).'}
        </td>
      </tr>
    `;
    return;
  }

  filtered.forEach(({ asset, att, attIdx }) => {
    const tr = document.createElement('tr');
    const isCloud = !!att.isCloudLink;

    tr.innerHTML = `
      <td>
        <div style="font-weight: 700; color: var(--text-main);">${escapeHtml(asset.name)}</div>
        <div class="asset-tag-code" style="display: inline-block; margin-top: 2px;">${escapeHtml(asset.id)}</div>
      </td>
      <td>
        <span style="color: #60a5fa; font-weight: 600;">${escapeHtml(asset.department)}</span>
      </td>
      <td>
        <div style="font-weight: 600; display: flex; align-items: center; gap: 6px;">
          <span>${isCloud ? '☁️' : '📊'}</span>
          <span>${escapeHtml(att.fileName)}</span>
          ${isCloud ? `<span class="cloud-badge">${escapeHtml(att.service || 'Cloud')}</span>` : ''}
        </div>
        <div style="font-size: 0.74rem; color: var(--text-muted);">
          ${att.fileSize || 'Excel'} &bull; ${escapeHtml(att.notes || '')}
        </div>
      </td>
      <td>
        <span style="font-size: 0.78rem; color: var(--text-muted);">${escapeHtml(att.uploadDate || '')}</span>
      </td>
      <td style="text-align: right; white-space: nowrap;">
        <button class="btn btn-sm btn-secondary" data-preview-global-asset="${escapeHtml(asset.id)}" data-preview-global-idx="${attIdx}" style="font-size: 0.76rem; padding: 4px 8px; margin-right: 4px;" title="Preview this spreadsheet">
          🔍 Preview
        </button>
        ${isCloud ? `
          <a href="${escapeHtml(att.url)}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-secondary" style="font-size: 0.76rem; padding: 4px 8px;">
            🔗 Open Drive
          </a>
        ` : `
          <button class="btn btn-sm btn-excel" data-download-global-asset="${escapeHtml(asset.id)}" data-download-global-idx="${attIdx}" style="font-size: 0.76rem; padding: 4px 8px;">
            ⬇️ Download
          </button>
        `}
        <button class="btn btn-sm btn-danger-outline" data-delete-global-asset="${escapeHtml(asset.id)}" data-delete-global-id="${escapeHtml(att.id)}" style="font-size: 0.76rem; padding: 4px 8px; margin-left: 4px;" title="Delete this Excel file">
          🗑️ Delete
        </button>
      </td>
    `;

    tbody.appendChild(tr);
  });

  // Attach button events
  tbody.querySelectorAll('[data-delete-global-asset]').forEach(btn => {
    btn.addEventListener('click', () => {
      const assetId = btn.dataset.deleteGlobalAsset;
      const attId = btn.dataset.deleteGlobalId;
      const targetAsset = assets.find(a => a.id === assetId);
      if (!targetAsset) return;

      const fileToDelete = (targetAsset.excelAttachments || []).find(a => a.id === attId);
      const fileName = fileToDelete ? fileToDelete.fileName : 'Excel file';

      if (confirm(`तुम्हाला खात्री आहे का?\n\n"${fileName}" ही फाइल "${assetId}" मधून कायमस्वरूपी डिलीट करायची आहे का?\n(Are you sure you want to delete this file?)`)) {
        targetAsset.excelAttachments = (targetAsset.excelAttachments || []).filter(a => a.id !== attId);
        saveData();
        renderAll();
        renderGlobalExcelTable(searchQuery);
        showToast(`🗑️ "${fileName}" डिलीट करण्यात आली (File deleted)`);
      }
    });
  });

  tbody.querySelectorAll('[data-preview-global-asset]').forEach(btn => {
    btn.addEventListener('click', () => {
      const assetId = btn.dataset.previewGlobalAsset;
      const attIdx = parseInt(btn.dataset.previewGlobalIdx, 10);
      const targetAsset = assets.find(a => a.id === assetId);
      if (!targetAsset) return;
      openExcelModal(targetAsset);
      if (targetAsset.excelAttachments && targetAsset.excelAttachments[attIdx]) {
        previewSheetData(targetAsset.excelAttachments[attIdx], targetAsset);
      }
    });
  });

  tbody.querySelectorAll('[data-download-global-asset]').forEach(btn => {
    btn.addEventListener('click', () => {
      const assetId = btn.dataset.downloadGlobalAsset;
      const attIdx = parseInt(btn.dataset.downloadGlobalIdx, 10);
      const targetAsset = assets.find(a => a.id === assetId);
      if (!targetAsset || !targetAsset.excelAttachments || !targetAsset.excelAttachments[attIdx]) return;
      downloadAttachedExcel(targetAsset.excelAttachments[attIdx], targetAsset);
    });
  });
}

// Setup Inline Add Equipment Form Attachments
let pendingInlinePcAttachment = null;

function setupInlineFormAttachments() {
  const inlinePcFileInput = document.getElementById('inlinePcFileInput');
  const btnBrowseInlinePcFile = document.getElementById('btnBrowseInlinePcFile');
  const btnDriveInlinePc = document.getElementById('btnDriveInlinePc');
  const btnRemoveInlinePcAtt = document.getElementById('btnRemoveInlinePcAtt');
  const inlinePcAttachmentLabel = document.getElementById('inlinePcAttachmentLabel');

  if (btnBrowseInlinePcFile && inlinePcFileInput) {
    btnBrowseInlinePcFile.addEventListener('click', () => inlinePcFileInput.click());
  }

  if (inlinePcFileInput) {
    inlinePcFileInput.addEventListener('click', (e) => e.stopPropagation());
    inlinePcFileInput.addEventListener('change', (e) => {
      const file = e.target.files && e.target.files[0];
      if (file) {
        if (typeof XLSX === 'undefined') {
          alert('XLSX library not loaded.');
          return;
        }
        const reader = new FileReader();
        reader.onload = (re) => {
          try {
            const data = new Uint8Array(re.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            const sName = wb.SheetNames[0];
            const sData = XLSX.utils.sheet_to_json(wb.Sheets[sName], { header: 1 });
            pendingInlinePcAttachment = {
              id: `att-${Date.now()}`,
              fileName: file.name,
              fileSize: `${Math.round(file.size / 1024)} KB`,
              uploadDate: new Date().toISOString().replace('T', ' ').substring(0, 16),
              notes: 'Attached during equipment registration',
              sheetData: sData
            };
            if (inlinePcAttachmentLabel) {
              inlinePcAttachmentLabel.innerHTML = `<span style="color: #34d399; font-weight: 700;">✅ जोडली: ${escapeHtml(file.name)}</span> (${Math.round(file.size / 1024)} KB)`;
            }
            if (btnRemoveInlinePcAtt) btnRemoveInlinePcAtt.style.display = 'inline-flex';
            showToast(`📊 "${file.name}" जोडण्यात आली`);
          } catch (err) {
            alert('Error parsing file: ' + err.message);
          }
        };
        reader.readAsArrayBuffer(file);
      }
    });
  }

  if (btnDriveInlinePc) {
    btnDriveInlinePc.addEventListener('click', () => {
      const url = prompt('Google Drive किंवा OneDrive लिंक पेस्ट करा (Paste Drive Share Link):');
      if (!url) return;
      const docName = prompt('फाइलचे नाव (Document Name):', 'Equipment_Audit_Drive.xlsx') || 'Equipment_Audit_Drive.xlsx';
      pendingInlinePcAttachment = {
        id: `att-cloud-${Date.now()}`,
        isCloudLink: true,
        fileName: docName,
        fileSize: url.includes('onedrive') ? 'OneDrive' : 'Google Drive',
        uploadDate: new Date().toISOString().replace('T', ' ').substring(0, 16),
        notes: 'Attached via Cloud Drive Link',
        url: url,
        service: url.includes('onedrive') ? 'OneDrive' : 'Google Drive'
      };
      if (inlinePcAttachmentLabel) {
        inlinePcAttachmentLabel.innerHTML = `<span style="color: #38bdf8; font-weight: 700;">☁️ ${escapeHtml(pendingInlinePcAttachment.service)}: ${escapeHtml(docName)}</span>`;
      }
      if (btnRemoveInlinePcAtt) btnRemoveInlinePcAtt.style.display = 'inline-flex';
      showToast(`☁️ ड्राइव्ह लिंक जोडली`);
    });
  }

  if (btnRemoveInlinePcAtt) {
    btnRemoveInlinePcAtt.addEventListener('click', () => {
      pendingInlinePcAttachment = null;
      if (inlinePcFileInput) inlinePcFileInput.value = '';
      if (inlinePcAttachmentLabel) inlinePcAttachmentLabel.textContent = 'कोणतीही फाईल निवडलेली नाही (No file selected)';
      btnRemoveInlinePcAtt.style.display = 'none';
      showToast('🗑️ निवडलेली फाइल काढण्यात आली (Attachment removed)');
    });
  }
}

function handleUploadForAsset(file, asset) {
  if (typeof XLSX === 'undefined') {
    alert('XLSX parser library not loaded');
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      const noteInput = document.getElementById('excelAttachmentNote');
      const note = noteInput ? noteInput.value.trim() : '';

      if (!asset.excelAttachments) asset.excelAttachments = [];
      asset.excelAttachments.push({
        id: `att-${Date.now()}`,
        fileName: file.name,
        fileSize: `${Math.round(file.size / 1024)} KB`,
        uploadDate: new Date().toISOString().replace('T', ' ').substring(0, 16),
        notes: note || 'Uploaded hardware sheet',
        sheetData: jsonData
      });

      saveData();
      renderExcelAttachmentsList(asset);
      renderKpis();
      renderInventoryList();
      showToast(`📊 Attached ${file.name} to ${asset.id}`);
      if (noteInput) noteInput.value = '';
    } catch (err) {
      alert('Error parsing Excel file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// Master Export to Excel
function exportInventoryExcel() {
  if (typeof XLSX === 'undefined') {
    alert('XLSX library not loaded');
    return;
  }

  const wb = XLSX.utils.book_new();

  const summaryAoa = [
    ['Dead Stock & Hardware Inventory Report'],
    [`Generated: ${new Date().toLocaleString()}`],
    [],
    ['Department / Course', 'Total Equipment', 'Computers', 'Printers', 'Projectors', 'Working', 'Faulty']
  ];

  departments.forEach(dept => {
    const list = assets.filter(a => a.department === dept);
    if (list.length === 0) return;
    const comps = list.filter(a => a.type === 'Computer').length;
    const prns = list.filter(a => a.type === 'Printer').length;
    const prjs = list.filter(a => a.type === 'Projector').length;
    const wrk = list.filter(a => a.status === 'Working').length;
    const flt = list.length - wrk;

    summaryAoa.push([dept, list.length, comps, prns, prjs, wrk, flt]);
  });

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryAoa);
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Department Summary');

  const assetsAoa = [
    ['Asset Tag / Dead Stock', 'Category', 'Model / Name', 'Department', 'Location', 'User', 'Status', 'Serial Number', 'Specs', 'IP Address', 'Fault / Issue']
  ];

  const targetList = activeFilters.department === 'ALL'
    ? assets
    : assets.filter(a => a.department === activeFilters.department);

  targetList.forEach(a => {
    assetsAoa.push([
      a.id,
      a.type,
      a.name,
      a.department,
      a.location,
      a.assignedUser,
      a.status,
      a.serialNumber,
      a.specs,
      a.ipAddress,
      a.faultDetails ? a.faultDetails.issue : ''
    ]);
  });

  const sheetName = activeFilters.department === 'ALL'
    ? 'All Equipment'
    : activeFilters.department.substring(0, 31).replace(/[\/\\?%*:|"<>\[\]]/g, ' ');

  const wsAssets = XLSX.utils.aoa_to_sheet(assetsAoa);
  XLSX.utils.book_append_sheet(wb, wsAssets, sheetName);

  const outFileName = activeFilters.department === 'ALL'
    ? 'Dead_Stock_Inventory_Report.xlsx'
    : `${activeFilters.department.replace(/[\/\\?%*:|"<>]/g, '_')}_Inventory.xlsx`;

  XLSX.writeFile(wb, outFileName);
  showToast(`Downloaded ${outFileName}`);
}

// Toast Notification Helper
function showToast(msg) {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = 'toast';
  toast.textContent = msg;
  container.appendChild(toast);

  setTimeout(() => {
    toast.remove();
  }, 3400);
}

// Escape HTML
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
