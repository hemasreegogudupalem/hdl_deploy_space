/**
 * Oracle Fusion Cloud HCM - HDL Automation
 * Frontend Application Controller (Light and Dark Theme Edition)
 * Supports the existing single-template flow and additive batch generation.
 */

// Application State
const TEMPLATE_CATEGORY_OPTIONS = [
  'Employee Data',
  'Organization Data',
  'Reference Data',
];

const state = {
  currentTab: 'dashboard',
  theme: 'light',
  selectedInput: 'input/saudi_employee_data.xlsx',
  selectedTemplate: 'Worker',
  templateCategoryOverrides: {},
  batch: {
    enabled: false,
    categories: {
      'Employee Data': false,
      'Organization Data': false,
      'Reference Data': false,
    },
    selectedTemplates: [],
    schemas: {},
    columns: [],
    columnMappings: {},
    mappingReady: false,
    refreshToken: 0,
    validation: null,
    isBusy: false,
  },
  inputs: [],
  templates: [],
  dbStatus: null,
  detectedDefaults: {},
  generationResult: null,
  isGenerating: false,

  // Data Grid state for Upload Tab
  dataGrid: {
    page: 1,
    pageSize: 15,
    search: '',
    totalRows: 0,
    totalPages: 1,
    columns: [],
    rows: []
  },

  // Mappings state
  mappings: {
    filter: 'all',
    search: '',
    data: []
  },

  // Outputs state
  outputs: {
    files: [],
    reports: [],
    activeFile: null,
    activeCategory: 'output',
    requestedFilename: null,
    activeRawText: '',
  }
};

// 3D Tab Themes for Sidebar Navigation
const tabThemes = {
  dashboard: {
    activeClass: 'bg-indigo-50 text-indigo-950 border-indigo-300 border-b-2 border-b-indigo-400 shadow-xs font-bold',
    iconBg: 'bg-indigo-100 text-indigo-700 shadow-2xs',
    dotColor: 'bg-indigo-500 shadow-xs'
  },
  upload: {
    activeClass: 'bg-emerald-50 text-emerald-950 border-emerald-300 border-b-2 border-b-emerald-400 shadow-xs font-bold',
    iconBg: 'bg-emerald-100 text-emerald-700 shadow-2xs',
    dotColor: 'bg-emerald-500 shadow-xs'
  },
  generator: {
    activeClass: 'bg-indigo-50 text-indigo-950 border-indigo-300 border-b-2 border-b-indigo-400 shadow-xs font-bold',
    iconBg: 'bg-indigo-100 text-indigo-700 shadow-2xs',
    dotColor: 'bg-indigo-500 shadow-xs'
  },
  mappings: {
    activeClass: 'bg-amber-50 text-amber-950 border-amber-300 border-b-2 border-b-amber-400 shadow-xs font-bold',
    iconBg: 'bg-amber-100 text-amber-700 shadow-2xs',
    dotColor: 'bg-amber-500 shadow-xs'
  },
  outputs: {
    activeClass: 'bg-violet-50 text-violet-950 border-violet-300 border-b-2 border-b-violet-400 shadow-xs font-bold',
    iconBg: 'bg-violet-100 text-violet-700 shadow-2xs',
    dotColor: 'bg-violet-500 shadow-xs'
  },
  database: {
    activeClass: 'bg-teal-50 text-teal-950 border-teal-300 border-b-2 border-b-teal-400 shadow-xs font-bold',
    iconBg: 'bg-teal-100 text-teal-700 shadow-2xs',
    dotColor: 'bg-teal-500 shadow-xs'
  }
};

// Initialize Application
document.addEventListener('DOMContentLoaded', async () => {
  initializeTheme();
  loadTemplateCategoryOverrides();
  setupNavigation();
  setupThemeToggle();
  setupDashboard();
  setupFileUpload();
  setupEventListeners();
  setupBatchGeneration();
  setupTemplateCategoryManager();
  switchTab(state.currentTab);

  // Load initial backend status and assets
  await refreshSystemStatus();
  await loadTemplates();
  await loadInputFiles();
  await triggerAutoDetect();
  await loadOutputFiles();
});

function initializeTheme() {
  state.theme = localStorage.getItem('hdl-theme') === 'dark' ? 'dark' : 'light';
  document.documentElement.dataset.theme = state.theme;
  renderThemeToggle();
}

function renderThemeToggle() {
  const button = document.getElementById('theme-toggle');
  const label = document.getElementById('theme-toggle-label');
  const icon = document.getElementById('theme-toggle-icon');
  if (!button || !label || !icon) return;

  const isDark = state.theme === 'dark';
  button.setAttribute('aria-pressed', String(isDark));
  button.setAttribute('aria-label', `Switch to ${isDark ? 'light' : 'dark'} theme`);
  label.textContent = isDark ? 'Light theme' : 'Dark theme';
  icon.innerHTML = isDark
    ? '<path d="M20.9 13A9 9 0 0 1 11 3.1 9 9 0 1 0 20.9 13Z"></path>'
    : '<path d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.36 6.36-.7-.7M6.34 6.34l-.7-.7m12.72 0-.7.7M6.34 17.66l-.7.7"></path><circle cx="12" cy="12" r="4"></circle>';
}

function setupThemeToggle() {
  document.getElementById('theme-toggle')?.addEventListener('click', () => {
    state.theme = state.theme === 'dark' ? 'light' : 'dark';
    document.documentElement.dataset.theme = state.theme;
    localStorage.setItem('hdl-theme', state.theme);
    renderThemeToggle();
  });
}

function setupDashboard() {
  document.querySelectorAll('[data-dashboard-tab]').forEach(button => {
    button.addEventListener('click', () => {
      switchTab(button.dataset.dashboardTab);
    });
  });
  renderDashboard();
}

function renderDashboard() {
  const workbookCount = document.getElementById('dashboard-workbook-count');
  const activeWorkbook = document.getElementById('dashboard-active-workbook');
  const templateCount = document.getElementById('dashboard-template-count');
  const outputCount = document.getElementById('dashboard-output-count');
  const databaseStatus = document.getElementById('dashboard-database-status');

  if (workbookCount) workbookCount.textContent = state.inputs.length;
  if (activeWorkbook) {
    activeWorkbook.textContent = state.selectedInput
      ? `Selected: ${state.selectedInput.split(/[/\\]/).pop()}`
      : 'No workbook selected';
  }
  if (templateCount) templateCount.textContent = state.templates.length;
  if (outputCount) outputCount.textContent = state.outputs.files.length;
  if (databaseStatus) {
    const connected = state.dbStatus?.status === 'connected';
    databaseStatus.textContent = connected ? 'Connected' : state.dbStatus ? 'Offline' : 'Checking…';
    databaseStatus.classList.toggle('is-online', connected);
    databaseStatus.classList.toggle('is-offline', Boolean(state.dbStatus) && !connected);
  }
}

// Toast Notifications (Clean Enterprise Light UI, No Emojis)
function showToast(message, type = 'info') {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toast = document.createElement('div');
  const borderClass =
    type === 'success' ? 'border-emerald-300 text-emerald-950 bg-emerald-50/90' :
    type === 'error' ? 'border-rose-300 text-rose-950 bg-rose-50/90' :
    type === 'warning' ? 'border-amber-300 text-amber-950 bg-amber-50/90' :
    'border-indigo-300 text-slate-800 bg-white';

  const dotClass =
    type === 'success' ? 'bg-emerald-600' :
    type === 'error' ? 'bg-rose-600' :
    type === 'warning' ? 'bg-amber-600' :
    'bg-indigo-600';

  toast.className = `toast ${borderClass}`;
  toast.innerHTML = `
    <span class="w-2 h-2 rounded-full ${dotClass} shrink-0"></span>
    <div class="flex-1 text-xs font-semibold">${escapeHtml(message)}</div>
    <button onclick="this.parentElement.remove()" class="text-slate-400 hover:text-slate-700 text-sm font-bold ml-2">&times;</button>
  `;

  container.appendChild(toast);
  setTimeout(() => {
    toast.style.transition = 'all 0.25s ease-out';
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(8px)';
    setTimeout(() => toast.remove(), 250);
  }, 4000);
}

// System Status and DB Check
async function refreshSystemStatus() {
  try {
    const res = await fetch('/api/status');
    const data = await res.json();
    state.dbStatus = data.database;
    renderStatusBadge(data.database);
    renderDashboard();
    return data;
  } catch (err) {
    state.dbStatus = { status: 'offline' };
    renderStatusBadge({ status: 'offline', error: err.message });
    renderDashboard();
  }
}

function renderStatusBadge(db) {
  const badge = document.getElementById('mysql-status-badge');
  const details = document.getElementById('mysql-status-text');
  const clientBadge = document.getElementById('client-badge-text');
  if (!badge) return;

  if (db && db.status === 'connected') {
    badge.className = 'w-2 h-2 rounded-full bg-emerald-500 live-dot inline-block';
    details.innerHTML = `MySQL: <span class="font-bold text-emerald-900">${db.database}</span> @ ${db.host}:${db.port}`;
    if (clientBadge) {
      clientBadge.textContent = `Client: ${db.active_client || 'Active'} (${db.source_system_owner || 'GRISMO'})`;
    }
  } else {
    badge.className = 'w-2 h-2 rounded-full bg-rose-600 inline-block';
    details.innerHTML = `<span class="text-rose-700 font-bold">MySQL Offline</span>`;
    if (clientBadge) {
      clientBadge.textContent = 'Client: Offline';
    }
  }
}

// Tab Navigation: upload, generator, mappings, outputs, database
function setupNavigation() {
  const navButtons = document.querySelectorAll('.nav-tab-btn');
  navButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const targetTab = btn.getAttribute('data-tab');
      switchTab(targetTab);
    });
  });
}

function switchTab(tabName) {
  state.currentTab = tabName;
  document.querySelectorAll('.tab-view').forEach(view => {
    view.classList.add('hidden');
  });

  const activeView = document.getElementById(`view-${tabName}`);
  if (activeView) activeView.classList.remove('hidden');

  document.querySelectorAll('.nav-tab-btn').forEach(btn => {
    const t = btn.getAttribute('data-tab');
    const theme = tabThemes[t] || tabThemes.generator;
    const iconSpan = btn.querySelector('span:first-child');
    let indicator = btn.querySelector('.tab-active-dot');

    if (t === tabName) {
      btn.className = `nav-tab-btn w-full flex items-center justify-between px-3 py-2.5 rounded-xl text-xs font-bold transition-all border ${theme.activeClass}`;
      if (iconSpan) iconSpan.className = `w-6 h-6 rounded-lg ${theme.iconBg} flex items-center justify-center`;
      if (!indicator) {
        indicator = document.createElement('span');
        indicator.className = `tab-active-dot w-2 h-2 rounded-full ${theme.dotColor}`;
        btn.appendChild(indicator);
      } else {
        indicator.className = `tab-active-dot w-2 h-2 rounded-full ${theme.dotColor}`;
      }
    } else {
      btn.className = `nav-tab-btn w-full flex items-center justify-between px-3 py-2.5 rounded-xl text-xs font-semibold transition-all border border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50`;
      if (iconSpan) iconSpan.className = 'w-6 h-6 rounded-lg bg-slate-100 text-slate-600 flex items-center justify-center';
      if (indicator) indicator.remove();
    }
  });

  // Lazy tab data loading
  if (tabName === 'upload') loadDataGrid();
  if (tabName === 'mappings') loadFieldMappings();
  if (tabName === 'database') loadDatabaseSchema();
  if (tabName === 'outputs') loadOutputFiles();
}

// Load Templates from API
async function loadTemplates() {
  try {
    const res = await fetch('/api/templates');
    const templates = await res.json();
    if (!res.ok) throw new Error(templates.detail || 'Failed to load templates');
    state.templates = templates;
    setExpectedOutput(state.selectedTemplate);
    renderBatchCategoryCards();
    renderTemplateOptions(templates);
    renderSchemaTemplateSelect(templates);
    updateTemplateBadges();
    renderDashboard();
  } catch (err) {
    showToast('Failed to load templates: ' + err.message, 'error');
  }
}

function renderSchemaTemplateSelect(templates) {
  const select = document.getElementById('db-template-select');
  if (!select) return;
  select.innerHTML = templates.map(t => `
    <option value="${t.template_name}" ${t.template_name === state.selectedTemplate ? 'selected' : ''}>
      ${t.template_name} (${t.objects_count} objects)
    </option>
  `).join('');

  select.onchange = (e) => {
    if (state.isGenerating) {
      e.target.value = state.selectedTemplate;
      return;
    }
    setExpectedOutput(e.target.value);
    state.selectedTemplate = e.target.value;
    updateTemplateBadges();
    loadDatabaseSchema();
  };
}

function renderTemplateOptions(templates) {
  const container = document.getElementById('template-cards-container');
  if (!container) return;
  container.innerHTML = '';

  const heading = document.getElementById('template-list-heading');
  const hint = document.getElementById('template-list-hint');
  if (heading) {
    heading.textContent = state.batch.enabled
      ? 'Select Templates'
      : 'Available Templates';
  }
  if (hint) {
    hint.textContent = state.batch.enabled
      ? 'Select one or more template cards from the active categories.'
      : 'Choose one template for a standard DAT, or select categories above to build a batch.';
  }

  const visibleTemplates = state.batch.enabled
    ? templates.filter(template => (
      state.batch.categories[batchTemplateCategory(template.template_name)]
    ))
    : templates;

  if (visibleTemplates.length === 0) {
    container.innerHTML = `
      <div class="col-span-full p-5 text-center text-slate-500 bg-slate-50 border border-dashed border-slate-300 rounded-xl text-xs">
        No templates are available in the selected categories.
      </div>
    `;
    return;
  }

  visibleTemplates.forEach(t => {
    const isSelected = state.batch.enabled
      ? state.batch.selectedTemplates.includes(t.template_name)
      : t.template_name === state.selectedTemplate;
    const card = document.createElement('div');
    card.className = `template-selection-card p-4 rounded-xl cursor-pointer ${
      isSelected
        ? 'template-selection-card-active border-2 border-indigo-600 border-b-4 border-b-indigo-800 bg-indigo-50/50 shadow-md'
        : 'card-3d hover:border-indigo-300'
    }`;
    card.onclick = () => {
      if (state.batch.enabled) {
        toggleBatchTemplate(t.template_name);
      } else {
        selectTemplate(t.template_name);
      }
    };

    card.innerHTML = `
      <div class="flex items-start justify-between mb-2">
        <div>
          <span class="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-bold bg-indigo-100 text-indigo-800 border border-indigo-200 mb-1">
            ${escapeHtml(t.output_file_name || t.template_name + '.dat')}
          </span>
          <h4 class="text-xs font-bold text-slate-900">${escapeHtml(t.template_name)}</h4>
        </div>
        <div class="w-4 h-4 rounded-full border ${isSelected ? 'bg-indigo-600 border-indigo-600' : 'border-slate-300'} flex items-center justify-center">
          ${isSelected ? '<svg class="w-2.5 h-2.5 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="3" d="M5 13l4 4L19 7"></path></svg>' : ''}
        </div>
      </div>
      <p class="text-[11px] text-slate-500 mb-2.5 line-clamp-2">${escapeHtml(t.description || 'Oracle Fusion HDL template object bundle.')}</p>
      <div class="flex flex-wrap gap-1">
        ${(t.objects || []).slice(0, 5).map(o => `
          <span class="text-[10px] bg-slate-100 text-slate-700 px-1.5 py-0.5 rounded border border-slate-200 font-medium">${escapeHtml(o)}</span>
        `).join('')}
        ${(t.objects || []).length > 5 ? `<span class="text-[10px] text-indigo-600 font-bold">+${t.objects.length - 5} more</span>` : ''}
      </div>
    `;
    container.appendChild(card);
  });
}

function renderBatchCategoryCards() {
  const container = document.getElementById('batch-category-cards');
  if (!container) return;

  const categories = [
    {
      name: 'Employee Data',
      description: 'Worker, contact, and employee records',
      tone: 'emerald',
      icon: '<path d="M16 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="10" cy="7" r="4"/><path d="M20 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/>',
    },
    {
      name: 'Organization Data',
      description: 'Jobs, locations, and organization setup',
      tone: 'blue',
      icon: '<rect x="3" y="7" width="18" height="14" rx="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16M7 11h.01M7 15h.01M12 11h.01M12 15h.01M17 11h.01M17 15h.01"/>',
    },
    {
      name: 'Reference Data',
      description: 'Person types and reusable reference values',
      tone: 'violet',
      icon: '<path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2Z"/><path d="M8 7h8M8 11h8"/>',
    },
  ];

  container.innerHTML = categories.map(category => {
    const count = state.templates.filter(
      template => batchTemplateCategory(template.template_name) === category.name,
    ).length;
    const active = state.batch.categories[category.name];
    return `
      <button
        type="button"
        class="category-card category-card-${category.tone} ${active ? 'category-card-active' : ''}"
        data-batch-category="${escapeHtml(category.name)}"
        aria-pressed="${active}"
      >
        <span class="category-card-icon" aria-hidden="true">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">${category.icon}</svg>
        </span>
        <span class="category-card-copy">
          <span class="category-card-title">${escapeHtml(category.name)}</span>
          <span class="category-card-description">${escapeHtml(category.description)}</span>
        </span>
        <span class="category-card-count">${count} ${count === 1 ? 'template' : 'templates'}</span>
        <span class="category-card-check" aria-hidden="true">${active ? '&#10003;' : '+'}</span>
      </button>
    `;
  }).join('');
}

function batchTemplateCategory(templateName) {
  const override = state.templateCategoryOverrides[templateName];
  if (TEMPLATE_CATEGORY_OPTIONS.includes(override)) return override;

  const template = state.templates.find(
    item => item.template_name === templateName,
  );
  if (TEMPLATE_CATEGORY_OPTIONS.includes(template?.category)) {
    return template.category;
  }
  const normalized = String(templateName || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  if (
    normalized === 'worker'
    || normalized === 'contact'
    || normalized === 'persondisability'
    || normalized.startsWith('person')
    || normalized.includes('employee')
    || normalized.includes('assignment')
  ) return 'Employee Data';
  if (
    normalized.includes('job')
    || normalized.includes('organization')
    || normalized.includes('businessunit')
    || normalized.includes('department')
    || normalized.includes('location')
    || normalized.includes('calendar')
    || normalized.includes('action')
    || normalized.includes('resource')
    || normalized.includes('legislativedatagroup')
  ) return 'Organization Data';
  return 'Reference Data';
}

function loadTemplateCategoryOverrides() {
  const saved = localStorage.getItem('hdl-template-categories');
  if (!saved) return;

  try {
    const overrides = JSON.parse(saved);
    if (!overrides || typeof overrides !== 'object' || Array.isArray(overrides)) {
      throw new Error('Saved template categories are not a valid object.');
    }
    state.templateCategoryOverrides = Object.fromEntries(
      Object.entries(overrides).filter(([, category]) => (
        TEMPLATE_CATEGORY_OPTIONS.includes(category)
      )),
    );
  } catch (error) {
    showToast(`Could not load saved template categories: ${error.message}`, 'error');
  }
}

function renderTemplateCategoryManager() {
  const container = document.getElementById('template-category-list');
  if (!container) return;

  if (state.templates.length === 0) {
    container.innerHTML = '<p class="p-4 text-center text-sm text-slate-500">No templates are currently available.</p>';
    return;
  }

  container.innerHTML = state.templates.map(template => {
    const category = batchTemplateCategory(template.template_name);
    return `
      <label class="template-category-row">
        <span class="template-category-name">
          <span>${escapeHtml(template.template_name)}</span>
          <span>${escapeHtml(template.output_file_name || `${template.template_name}.dat`)}</span>
        </span>
        <select class="template-category-select" data-template-name="${escapeHtml(template.template_name)}" aria-label="Category for ${escapeHtml(template.template_name)}">
          ${TEMPLATE_CATEGORY_OPTIONS.map(option => `
            <option value="${escapeHtml(option)}" ${category === option ? 'selected' : ''}>${escapeHtml(option)}</option>
          `).join('')}
        </select>
      </label>
    `;
  }).join('');
}

function setupTemplateCategoryManager() {
  const dialog = document.getElementById('template-category-dialog');
  const saveButton = document.getElementById('btn-save-template-categories');
  if (!dialog || !saveButton) return;

  document.getElementById('btn-manage-template-categories')?.addEventListener('click', () => {
    renderTemplateCategoryManager();
    dialog.showModal();
  });
  document.getElementById('btn-close-template-category-dialog')?.addEventListener(
    'click',
    () => dialog.close(),
  );
  document.getElementById('btn-cancel-template-categories')?.addEventListener(
    'click',
    () => dialog.close(),
  );
  saveButton.addEventListener('click', () => {
    const overrides = {};
    document.querySelectorAll('#template-category-list select[data-template-name]').forEach(select => {
      if (TEMPLATE_CATEGORY_OPTIONS.includes(select.value)) {
        overrides[select.dataset.templateName] = select.value;
      }
    });

    try {
      localStorage.setItem('hdl-template-categories', JSON.stringify(overrides));
    } catch (error) {
      showToast(`Could not save template categories: ${error.message}`, 'error');
      return;
    }

    state.templateCategoryOverrides = overrides;
    const previousSelection = state.batch.selectedTemplates.join('|');
    state.batch.selectedTemplates = state.batch.selectedTemplates.filter(
      name => !state.batch.enabled || state.batch.categories[batchTemplateCategory(name)],
    );
    if (state.batch.selectedTemplates.length > 0) {
      state.selectedTemplate = state.batch.selectedTemplates[0];
    } else if (previousSelection) {
      state.selectedTemplate = 'Worker';
    }
    setExpectedOutput(state.selectedTemplate);
    updateTemplateBadges();
    renderBatchCategoryCards();
    renderTemplateOptions(state.templates);
    renderBatchSelectionSummary();
    invalidateBatchValidation();
    refreshBatchMappingPanels();
    if (previousSelection !== state.batch.selectedTemplates.join('|')) {
      triggerAutoDetect();
    }

    dialog.close();
    showToast('Template categories saved in this browser.', 'success');
  });
}

function toggleBatchTemplate(templateName) {
  const selected = state.batch.selectedTemplates;
  state.batch.selectedTemplates = selected.includes(templateName)
    ? selected.filter(name => name !== templateName)
    : [...selected, templateName];

  if (state.batch.selectedTemplates.length > 0) {
    state.selectedTemplate = state.batch.selectedTemplates[0];
    setExpectedOutput(state.selectedTemplate);
    updateTemplateBadges();
    triggerAutoDetect();
  } else {
    state.selectedTemplate = 'Worker';
    setExpectedOutput(state.selectedTemplate);
    updateTemplateBadges();
  }
  renderTemplateOptions(state.templates);
  renderBatchSelectionSummary();
  invalidateBatchValidation();
  refreshBatchMappingPanels();
}

function renderBatchSelectionSummary() {
  const summary = document.getElementById('batch-selection-summary');
  const selected = state.batch.selectedTemplates;
  const activeCategories = Object.keys(state.batch.categories).filter(
    category => state.batch.categories[category],
  );
  if (summary) summary.textContent = state.batch.enabled
    ? selected.length
      ? `${selected.length} selected`
      : `${activeCategories.length} categor${activeCategories.length === 1 ? 'y' : 'ies'} active · select templates`
    : 'Single-template mode';
  const actionPanel = document.getElementById('batch-action-panel');
  actionPanel?.classList.toggle(
    'hidden',
    !state.batch.enabled || selected.length === 0,
  );
  const validateButton = document.getElementById('btn-validate-batch');
  const generateButton = document.getElementById('btn-generate-batch');
  if (validateButton) {
    validateButton.disabled = (
      !state.batch.enabled
      || selected.length === 0
      || !state.batch.mappingReady
      || state.batch.isBusy
    );
  }
  if (generateButton) {
    generateButton.disabled = (
      !state.batch.validation?.can_generate || state.batch.isBusy
    );
  }
}

function setupBatchGeneration() {
  const categoryCards = document.getElementById('batch-category-cards');
  const singleGenerateButton = document.getElementById('btn-execute-generate');
  const singleGeneratePanel = document.getElementById('single-generation-trigger');

  categoryCards?.addEventListener('click', event => {
    const button = event.target.closest('[data-batch-category]');
    if (!button) return;
    const category = button.dataset.batchCategory;
    state.batch.categories[category] = !state.batch.categories[category];
    state.batch.enabled = Object.values(state.batch.categories).some(Boolean);
    state.batch.selectedTemplates = state.batch.selectedTemplates.filter(
      name => state.batch.categories[batchTemplateCategory(name)],
    );
    if (state.batch.selectedTemplates.length > 0) {
      state.selectedTemplate = state.batch.selectedTemplates[0];
      setExpectedOutput(state.selectedTemplate);
      updateTemplateBadges();
    }

    renderBatchCategoryCards();
    renderTemplateOptions(state.templates);
    renderBatchSelectionSummary();
    singleGeneratePanel?.classList.toggle('hidden', state.batch.enabled);
    if (singleGenerateButton) {
      singleGenerateButton.disabled = state.batch.enabled;
    }
    invalidateBatchValidation();
    refreshBatchMappingPanels();
  });

  document.getElementById('btn-validate-batch')?.addEventListener(
    'click',
    validateBatchSelections,
  );
  document.getElementById('btn-generate-batch')?.addEventListener(
    'click',
    generateBatchSelections,
  );
  renderBatchSelectionSummary();
}

function normalizeExcelHeader(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

async function refreshBatchMappingPanels() {
  const refreshToken = ++state.batch.refreshToken;
  const selectedTemplates = [...state.batch.selectedTemplates];
  state.batch.mappingReady = false;
  renderBatchSelectionSummary();
  if (!state.batch.enabled || selectedTemplates.length === 0 || !state.selectedInput) {
    state.batch.columnMappings = {};
    return;
  }

  const filePath = state.selectedInput;

  try {
    const columnsResponse = await fetch(
      `/api/excel/preview?file_path=${encodeURIComponent(filePath)}&page=1&page_size=1`,
    );
    const columnsData = await columnsResponse.json();
    if (!columnsResponse.ok) {
      throw new Error(columnsData.detail || 'Failed to read workbook columns.');
    }
    const templateData = await Promise.all(selectedTemplates.map(async templateName => {
      const [schemaResponse, mappingResponse] = await Promise.all([
        fetch(
          `/api/generation/templates/${encodeURIComponent(templateName)}`,
        ),
        fetch(
          `/api/mappings/preview?file_path=${encodeURIComponent(filePath)}&template_name=${encodeURIComponent(templateName)}`,
        ),
      ]);
      const schema = await schemaResponse.json();
      const mappingData = await mappingResponse.json();
      if (!schemaResponse.ok) {
        throw new Error(schema.detail || `Failed to load ${templateName} fields.`);
      }
      if (!mappingResponse.ok) {
        throw new Error(mappingData.detail || `Failed to analyze ${templateName} mappings.`);
      }
      return { schema, mappingData };
    }));

    if (
      refreshToken !== state.batch.refreshToken
      || filePath !== state.selectedInput
      || selectedTemplates.join('|') !== state.batch.selectedTemplates.join('|')
    ) return;

    const columns = columnsData.columns || [];
    state.batch.columns = columns;
    state.batch.columnMappings = {};
    templateData.forEach(({ schema, mappingData }) => {
      state.batch.schemas[schema.template_name] = schema;
      const byObjectAndField = new Map();
      const byField = new Map();
      (mappingData.mappings || []).forEach(mapping => {
        if (!mapping.excel_column) return;
        const fieldKey = normalizeExcelHeader(mapping.field_name);
        const objectKey = `${normalizeExcelHeader(mapping.object_name)}.${fieldKey}`;
        byObjectAndField.set(objectKey, mapping.excel_column);
        const fieldColumns = byField.get(fieldKey) || new Set();
        fieldColumns.add(mapping.excel_column);
        byField.set(fieldKey, fieldColumns);
      });

      const columnMappings = {};
      schema.source_fields.forEach(field => {
        if (field.generated || field.internal_id || field.defaulted) return;
        const normalizedField = normalizeExcelHeader(field.field_name);
        const objectKey = `${normalizeExcelHeader(field.object_name)}.${normalizedField}`;
        const fieldMatches = byField.get(normalizedField);
        const analyzedColumn = field.object_name
          ? byObjectAndField.get(objectKey)
          : fieldMatches && fieldMatches.size === 1
            ? [...fieldMatches][0]
            : null;
        const headerColumn = columns.find(
          column => normalizeExcelHeader(column) === normalizedField,
        );
        const matchedColumn = columns.includes(analyzedColumn)
          ? analyzedColumn
          : headerColumn;
        if (matchedColumn) {
          columnMappings[field.mapping_key] = matchedColumn;
        }
      });
      state.batch.columnMappings[schema.template_name] = columnMappings;
    });
    state.batch.mappingReady = true;
    invalidateBatchValidation();
    renderBatchSelectionSummary();
  } catch (error) {
    if (refreshToken !== state.batch.refreshToken) return;
    showToast(`Automatic batch mapping failed: ${error.message}`, 'error');
  }
}

function invalidateBatchValidation() {
  state.batch.validation = null;
  const summary = document.getElementById('batch-validation-summary');
  const generateButton = document.getElementById('btn-generate-batch');
  if (summary) {
    summary.classList.add('hidden');
    summary.innerHTML = '';
  }
  if (generateButton) generateButton.disabled = true;
}

function batchGenerationPayload() {
  return {
    file_path: state.selectedInput,
    templates: state.batch.selectedTemplates.map(templateName => ({
      template_name: templateName,
      column_mappings: state.batch.columnMappings[templateName] || {},
    })),
    legislation_code: normalizeLegislationCode(
      document.getElementById('input-legislation-code').value,
    ) || null,
    legal_employer_name: document.getElementById('input-legal-employer').value.trim() || null,
    source_system_owner: document.getElementById('input-source-owner').value.trim() || null,
  };
}

function renderBatchValidation(validation) {
  const container = document.getElementById('batch-validation-summary');
  const generateButton = document.getElementById('btn-generate-batch');
  if (!container) return;
  state.batch.validation = validation;
  container.classList.remove('hidden');
  container.innerHTML = `
    <div class="p-3 rounded-lg border ${validation.can_generate ? 'border-emerald-200 bg-emerald-50' : 'border-rose-200 bg-rose-50'}">
      <div class="text-xs font-bold ${validation.can_generate ? 'text-emerald-800' : 'text-rose-800'}">
        ${validation.can_generate ? 'Validation complete. Ready to generate.' : 'Fix required mappings or errors before generating.'}
        ${validation.row_count} workbook rows checked.
      </div>
    </div>
    ${validation.templates.map(item => `
      <div class="p-3 rounded-lg border ${item.ready ? 'border-slate-200 bg-white' : 'border-rose-200 bg-rose-50'} space-y-1">
        <div class="flex flex-wrap justify-between gap-2 text-xs">
          <span class="font-bold text-slate-800">${escapeHtml(item.template_name)}</span>
          <span class="text-slate-600">${item.row_count} rows • ${item.generated_rows} generated</span>
        </div>
        ${item.missing_required_columns.length ? `
          <div class="text-[11px] text-rose-700">Missing required mappings: ${item.missing_required_columns.map(escapeHtml).join(', ')}</div>
        ` : ''}
        ${item.errors.length ? `
          <ul class="list-disc pl-5 text-[11px] text-rose-700">
            ${item.errors.slice(0, 10).map(error => `
              <li>${escapeHtml(typeof error === 'string' ? error : error.Error || JSON.stringify(error))}</li>
            `).join('')}
          </ul>
        ` : ''}
      </div>
    `).join('')}
  `;
  if (generateButton) generateButton.disabled = !validation.can_generate;
}

async function validateBatchSelections() {
  if (!state.batch.selectedTemplates.length) {
    showToast('Select at least one template.', 'warning');
    return;
  }
  if (!state.batch.mappingReady) {
    showToast('Wait for automatic template mapping to finish before validating.', 'warning');
    return;
  }
  const button = document.getElementById('btn-validate-batch');
  const summary = document.getElementById('batch-validation-summary');
  button.disabled = true;
  button.textContent = 'Validating...';
  try {
    const response = await fetch('/api/generate/batch/validate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(batchGenerationPayload()),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.detail || 'Batch validation failed.');
    renderBatchValidation(result);
  } catch (error) {
    if (summary) {
      summary.classList.remove('hidden');
      summary.innerHTML = `<div class="p-3 rounded-lg border border-rose-200 bg-rose-50 text-xs text-rose-800">${escapeHtml(error.message)}</div>`;
    }
    showToast(error.message, 'error');
  } finally {
    button.disabled = false;
    button.textContent = 'Validate selected templates';
    renderBatchSelectionSummary();
  }
}

async function generateBatchSelections() {
  if (!state.batch.validation?.can_generate) {
    showToast('Validate the selected templates before generating.', 'warning');
    return;
  }
  const button = document.getElementById('btn-generate-batch');
  const summary = document.getElementById('batch-generation-summary');
  button.disabled = true;
  state.batch.isBusy = true;
  summary.classList.remove('hidden');
  summary.innerHTML = '<div class="text-xs text-slate-600">Generating selected templates in HDL load order...</div>';
  try {
    const response = await fetch('/api/generate/batch', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(batchGenerationPayload()),
    });
    const result = await response.json();
    if (!response.ok || !result.success || !result.zip_filename) {
      throw new Error(result.detail || 'Batch generation failed.');
    }
    summary.innerHTML = `
      <div class="p-3 rounded-lg border border-emerald-200 bg-emerald-50 space-y-2">
        <div class="text-xs font-bold text-emerald-800">Generated ${result.output_filenames.length} DAT file(s).</div>
        <a class="inline-flex btn-3d-primary px-3 py-1.5 text-xs font-bold" href="${escapeHtml(result.zip_download_url)}">
          Download ${escapeHtml(result.zip_filename)}
        </a>
        <ul class="text-[11px] text-slate-700 list-disc pl-5">
          ${result.templates.map(item => `
            <li>${escapeHtml(item.template_name)}: ${item.generated_rows} rows${item.deduplicated_rows ? `, ${item.deduplicated_rows} duplicates removed` : ''}</li>
          `).join('')}
        </ul>
      </div>
    `;
    showToast('All selected HDL templates generated.', 'success');
    state.outputs.requestedFilename = result.zip_filename;
    state.outputs.activeFile = result.zip_filename;
    state.outputs.activeCategory = 'output';
    await loadOutputFiles(result.zip_filename);
    switchTab('outputs');
  } catch (error) {
    summary.innerHTML = `<div class="p-3 rounded-lg border border-rose-200 bg-rose-50 text-xs text-rose-800">${escapeHtml(error.message)}</div>`;
    showToast(error.message, 'error');
  } finally {
    state.batch.isBusy = false;
    renderBatchSelectionSummary();
  }
}

function updateTemplateBadges() {
  const current = state.templates.find(t => t.template_name === state.selectedTemplate);
  const badge = document.getElementById('current-template-badge');
  const targetMapTmpl = document.getElementById('map-target-template');
  const genTargetHeading = document.getElementById('gen-target-heading');
  const genTargetDatName = document.getElementById('gen-target-datname');

  const datName = current ? (current.output_file_name || `${state.selectedTemplate}.dat`) : `${state.selectedTemplate}.dat`;

  if (badge) {
    if (current) {
      badge.textContent = `${datName} (${current.objects_count} Objects)`;
    } else {
      badge.textContent = datName;
    }
  }

  if (targetMapTmpl) targetMapTmpl.textContent = state.selectedTemplate;
  if (genTargetHeading) genTargetHeading.textContent = `Ready to Generate ${datName}`;
  if (genTargetDatName) genTargetDatName.textContent = datName;
}

function selectTemplate(templateName) {
  if (state.isGenerating) return;
  state.selectedTemplate = templateName;
  setExpectedOutput(templateName);
  clearGenerationFeedback();
  renderTemplateOptions(state.templates);
  updateTemplateBadges();
  triggerAutoDetect();
}

function setExpectedOutput(templateName) {
  const selectedTemplate = state.templates.find(
    template => template.template_name === templateName,
  );
  state.outputs.requestedFilename = selectedTemplate
    ? (selectedTemplate.output_file_name || `${templateName}.dat`)
    : `${templateName}.dat`;
  state.outputs.activeFile = state.outputs.files.some(
    file => file.filename === state.outputs.requestedFilename,
  ) ? state.outputs.requestedFilename : null;
}

function clearGenerationFeedback() {
  state.generationResult = null;
  document.getElementById('generation-results-card')?.classList.add('hidden');
  document.getElementById('validation-errors-inline-card')?.classList.add('hidden');
  const errorsTbody = document.getElementById('validation-errors-table-body');
  if (errorsTbody) errorsTbody.innerHTML = '';
  const progressSection = document.getElementById('generation-progress-box');
  if (progressSection) progressSection.classList.add('hidden');
  const progressBar = document.getElementById('generation-progress-bar');
  if (progressBar) {
    progressBar.classList.remove('bg-rose-600');
    progressBar.classList.add('from-indigo-600', 'to-blue-600');
    progressBar.style.width = '0%';
  }
}

// Load Input Excel Files
async function loadInputFiles() {
  try {
    const res = await fetch('/api/inputs');
    const files = await res.json();
    if (!res.ok) throw new Error(files.detail || 'Failed to load workbooks');
    state.inputs = files;
    if (!files.some(file => file.path === state.selectedInput)) {
      state.selectedInput = files[0]?.path || '';
    }
    renderInputFileOptions(files);
    updateFileBadges();
    loadDataGrid();
    renderDashboard();
  } catch (err) {
    showToast('Failed to load input files: ' + err.message, 'error');
  }
}

function renderInputFileOptions(files) {
  const container = document.getElementById('input-files-list');
  if (!container) return;
  container.innerHTML = '';

  if (files.length === 0) {
    container.innerHTML = `<div class="p-3 text-center text-xs text-slate-400">No Excel workbooks found in input/ directory.</div>`;
    return;
  }

  files.forEach(f => {
    const isSelected = f.path === state.selectedInput;
    const item = document.createElement('div');
    item.className = `flex items-center justify-between p-3 rounded-xl transition-all cursor-pointer ${
      isSelected
        ? 'bg-emerald-50/90 border-2 border-emerald-500 border-b-3 border-b-emerald-600 shadow-sm'
        : 'bg-white border border-slate-200 border-b-2 border-b-slate-300 hover:border-slate-300 hover:bg-slate-50'
    }`;
    item.onclick = event => {
      if (event.target.closest('.btn-delete-input')) return;
      selectInputFile(f.path);
    };

    item.innerHTML = `
      <div class="flex items-center gap-3">
        <div class="w-8 h-8 rounded-lg bg-emerald-100 text-emerald-800 border border-emerald-200 flex items-center justify-center font-black text-[11px] shadow-2xs">
          XLS
        </div>
        <div>
          <div class="text-xs font-bold text-slate-900">${escapeHtml(f.filename)}</div>
          <div class="text-[10px] text-slate-500 font-medium">${f.size_kb} KB • Modified ${f.modified}</div>
        </div>
      </div>
      <div class="flex items-center gap-2">
        ${isSelected ? '<span class="text-[10px] font-bold text-emerald-800 bg-emerald-100 px-2 py-0.5 rounded-full border border-emerald-300">Active Sheet</span>' : ''}
        <div class="w-4 h-4 rounded-full border ${isSelected ? 'bg-emerald-600 border-emerald-600' : 'border-slate-300'} flex items-center justify-center">
          ${isSelected ? '<svg class="w-2.5 h-2.5 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="3" d="M5 13l4 4L19 7"></path></svg>' : ''}
        </div>
        <button type="button" class="btn-delete-input" aria-label="Delete workbook ${escapeHtml(f.filename)}" title="Delete workbook">
          <svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
            <path d="M3 6h18M8 6V4h8v2m3 0-.9 14H5.9L5 6m4 4v6m6-6v6"></path>
          </svg>
        </button>
      </div>
    `;
    item.querySelector('.btn-delete-input').addEventListener('click', event => {
      event.stopPropagation();
      deleteInputWorkbook(f);
    });
    container.appendChild(item);
  });
}

function updateFileBadges() {
  const fileName = state.selectedInput.split(/[/\\]/).pop() || 'No workbook selected';
  const fileBadge = document.getElementById('current-file-badge');
  const targetMapFile = document.getElementById('map-target-file');
  if (fileBadge) fileBadge.textContent = fileName;
  if (targetMapFile) targetMapFile.textContent = fileName;
}

function selectInputFile(path) {
  const fileChanged = path !== state.selectedInput;
  state.selectedInput = path;
  if (fileChanged) {
    state.batch.schemas = {};
    state.batch.columnMappings = {};
    state.batch.mappingReady = false;
    invalidateBatchValidation();
  }
  renderInputFileOptions(state.inputs);
  updateFileBadges();
  renderDashboard();
  state.dataGrid.page = 1;
  loadDataGrid();
  triggerAutoDetect();
  refreshBatchMappingPanels();
}

async function deleteInputWorkbook(inputFile) {
  if (state.isGenerating || state.batch.isBusy) {
    showToast('Wait for the current generation to finish before deleting a workbook.', 'warning');
    return;
  }
  if (!window.confirm(`Delete "${inputFile.filename}" from the input/ folder? This cannot be undone.`)) {
    return;
  }

  try {
    const response = await fetch(`/api/inputs?filename=${encodeURIComponent(inputFile.filename)}`, {
      method: 'DELETE',
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.detail || 'Failed to delete workbook.');

    const selectedFileDeleted = state.selectedInput === inputFile.path;
    state.inputs = state.inputs.filter(file => file.filename !== inputFile.filename);
    if (selectedFileDeleted) {
      selectInputFile(state.inputs[0]?.path || '');
    } else {
      renderInputFileOptions(state.inputs);
      renderDashboard();
    }
    showToast(`Deleted ${inputFile.filename}.`, 'success');
  } catch (error) {
    showToast(`Could not delete ${inputFile.filename}: ${error.message}`, 'error');
  }
}

// Drag & Drop File Upload
function setupFileUpload() {
  const dropzone = document.getElementById('file-dropzone');
  const fileInput = document.getElementById('excel-file-input');
  if (!dropzone || !fileInput) return;

  dropzone.addEventListener('click', (e) => {
    if (e.target !== fileInput) {
      fileInput.click();
    }
  });

  ['dragenter', 'dragover'].forEach(name => {
    dropzone.addEventListener(name, (e) => {
      e.preventDefault();
      dropzone.classList.add('border-emerald-500', 'bg-emerald-50/60');
    });
  });

  ['dragleave', 'drop'].forEach(name => {
    dropzone.addEventListener(name, (e) => {
      e.preventDefault();
      dropzone.classList.remove('border-emerald-500', 'bg-emerald-50/60');
    });
  });

  dropzone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) uploadFile(files[0]);
  });

  fileInput.addEventListener('change', () => {
    if (fileInput.files.length > 0) uploadFile(fileInput.files[0]);
  });
}

async function uploadFile(file) {
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    showToast('Please upload an Excel workbook (.xlsx or .xls)', 'warning');
    return;
  }

  const formData = new FormData();
  formData.append('file', file);

  try {
    showToast(`Uploading ${file.name}...`, 'info');
    const res = await fetch('/api/upload', {
      method: 'POST',
      body: formData,
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.detail || 'Upload failed');

    showToast(`Successfully uploaded ${result.filename}`, 'success');
    await loadInputFiles();
    selectInputFile(result.path);
  } catch (err) {
    showToast(err.message, 'error');
  }
}

// Auto-Detect Parameters
function normalizeLegislationCode(value) {
  const code = String(value || '').trim();
  return code.toUpperCase() === 'SAU' ? 'SA' : code;
}

async function triggerAutoDetect() {
  if (!state.selectedInput) return;
  const btn = document.getElementById('btn-autodetect');
  if (btn) btn.innerHTML = 'Detecting...';

  try {
    const res = await fetch(`/api/defaults?file_path=${encodeURIComponent(state.selectedInput)}&template_name=${encodeURIComponent(state.selectedTemplate)}`);
    const data = await res.json();
    if (res.ok) {
      state.detectedDefaults = data;
      const legislationCode = normalizeLegislationCode(data.legislation_code);
      document.getElementById('input-legislation-code').value = legislationCode;
      document.getElementById('input-legal-employer').value = data.legal_employer_name || '';
      document.getElementById('input-source-owner').value = data.source_system_owner || 'GRISMO';

      document.getElementById('badge-detected-legislation').textContent = legislationCode || 'Auto';
      document.getElementById('badge-detected-employer').textContent = data.legal_employer_name || 'Auto';
      document.getElementById('badge-detected-owner').textContent = data.source_system_owner || 'Auto';
    }
  } catch (err) {
    console.warn('Auto-detect error:', err);
  } finally {
    if (btn) {
      btn.innerHTML = `
        <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"></path></svg>
        <span>Auto-Detect Parameters</span>
      `;
    }
  }
}

// HDL Generation Pipeline Execution
async function executeGeneration() {
  if (state.isGenerating) return;
  if (!state.selectedInput) {
    showToast('Select or upload an Excel workbook before generating.', 'warning');
    return;
  }
  state.isGenerating = true;

  const btn = document.getElementById('btn-execute-generate');
  btn.disabled = true;
  btn.innerHTML = 'Transforming & Generating HDL...';

  const progressSection = document.getElementById('generation-progress-box');
  const progressBar = document.getElementById('generation-progress-bar');
  const progressText = document.getElementById('generation-progress-text');
  clearGenerationFeedback();

  progressSection.classList.remove('hidden');
  progressBar.style.width = '25%';
  progressText.textContent = 'Connecting to database and reading input Excel...';

  const payload = {
    file_path: state.selectedInput,
    template_name: state.selectedTemplate,
    legislation_code: normalizeLegislationCode(document.getElementById('input-legislation-code').value) || null,
    legal_employer_name: document.getElementById('input-legal-employer').value.trim() || null,
    source_system_owner: document.getElementById('input-source-owner').value.trim() || null,
  };

  try {
    progressBar.style.width = '65%';
    progressText.textContent = 'Mapping objects, validating national IDs, generating records...';

    const res = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });

    const result = await res.json();
    if (!res.ok || !result.success || !result.output_filename) {
      throw new Error(result.detail || 'Generation failed to return an output file.');
    }
    if (
      result.template_name
      && result.template_name.toLowerCase() !== payload.template_name.toLowerCase()
    ) {
      throw new Error(
        `Generated output is for ${result.template_name}, not ${payload.template_name}.`,
      );
    }

    progressBar.style.width = '100%';
    progressText.textContent = 'HDL bundle and audit reports generated successfully!';
    state.generationResult = result;

    renderGenerationResults(result);
    showToast(`HDL generation complete: ${result.generated_rows} records generated`, 'success');
    if (result.validation_errors_count > 0) {
      showToast(
        `${result.validation_errors_count} source validation issue(s); affected records were skipped.`,
        'warning',
      );
    }
    state.outputs.requestedFilename = result.output_filename;
    state.outputs.activeFile = result.output_filename;
    state.outputs.activeCategory = 'output';
    await loadOutputFiles(result.output_filename);
    switchTab('outputs');
  } catch (err) {
    progressBar.style.width = '100%';
    progressBar.classList.remove('from-indigo-600', 'to-blue-600');
    progressBar.classList.add('bg-rose-600');
    progressText.textContent = 'Error: ' + err.message;
    showToast(err.message, 'error');
  } finally {
    state.isGenerating = false;
    btn.disabled = false;
    btn.innerHTML = 'Execute Transformation & Generate HDL Bundle';
  }
}

function renderGenerationResults(res) {
  const container = document.getElementById('generation-results-card');
  if (!container) return;
  container.classList.remove('hidden');

  document.getElementById('res-stat-read').textContent = res.total_excel_rows;
  document.getElementById('res-stat-generated').textContent = res.generated_rows;
  document.getElementById('res-stat-skipped').textContent = res.skipped_rows;
  document.getElementById('res-stat-errors').textContent = res.validation_errors_count;
  document.getElementById('res-stat-filesize').textContent = res.output_file_size_kb + ' KB';

  // Quick download DAT button
  const quickDownload = document.getElementById('btn-quick-download-dat');
  if (quickDownload) {
    quickDownload.textContent = `Download ${res.output_filename}`;
    quickDownload.onclick = () => {
      window.location.href = res.output_download_url ||
        `/api/reports/download/output/${encodeURIComponent(res.output_filename)}`;
    };
  }

  // Inline validation errors table
  const errorsCard = document.getElementById('validation-errors-inline-card');
  const errorsTbody = document.getElementById('validation-errors-table-body');
  if (res.validation_errors && res.validation_errors.length > 0) {
    errorsCard.classList.remove('hidden');
    errorsTbody.innerHTML = res.validation_errors.map(err => `
      <tr class="hover:bg-rose-50/50 transition-colors">
        <td class="py-1.5 px-3 font-mono text-slate-500 font-semibold">${err.ExcelRow || '—'}</td>
        <td class="py-1.5 px-3 font-bold text-slate-800">${err.Object || res.template_name || 'HDL'}</td>
        <td class="py-1.5 px-3 font-mono text-indigo-700 font-semibold">${err.Field || '—'}</td>
        <td class="py-1.5 px-3 text-rose-700 font-semibold">${escapeHtml(err.Error || '')}</td>
      </tr>
    `).join('');
  } else {
    errorsCard.classList.add('hidden');
  }
}

// Data Grid Spreadsheet View (Embedded on Upload Tab)
async function loadDataGrid() {
  const tableHead = document.getElementById('grid-table-head');
  const tableBody = document.getElementById('grid-table-body');
  if (!tableHead || !tableBody) return;

  if (!state.selectedInput) {
    state.dataGrid.totalRows = 0;
    state.dataGrid.totalPages = 1;
    state.dataGrid.columns = [];
    state.dataGrid.rows = [];
    tableHead.innerHTML = '';
    tableBody.innerHTML = '<tr><td colspan="10" class="text-center py-6 text-slate-400 text-xs">Upload or select a workbook to inspect its data.</td></tr>';
    document.getElementById('grid-total-rows-badge').textContent = '0 Rows';
    document.getElementById('grid-total-cols-badge').textContent = '0 Columns';
    document.getElementById('grid-filename-badge').textContent = 'No workbook selected';
    document.getElementById('grid-page-indicator').textContent = '0 rows';
    return;
  }

  tableBody.innerHTML = `<tr><td colspan="10" class="text-center py-6 text-slate-400 text-xs">Loading spreadsheet data...</td></tr>`;

  try {
    const query = `/api/excel/preview?file_path=${encodeURIComponent(state.selectedInput)}&page=${state.dataGrid.page}&page_size=${state.dataGrid.pageSize}&search=${encodeURIComponent(state.dataGrid.search)}`;
    const res = await fetch(query);
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail);

    state.dataGrid.totalRows = data.total_rows;
    state.dataGrid.totalPages = data.total_pages;
    state.dataGrid.columns = data.columns;
    state.dataGrid.rows = data.rows;

    document.getElementById('grid-total-rows-badge').textContent = `${data.total_rows} Rows`;
    document.getElementById('grid-total-cols-badge').textContent = `${data.columns.length} Columns`;
    document.getElementById('grid-filename-badge').textContent = data.filename;

    const startRow = (data.current_page - 1) * data.page_size + 1;
    const endRow = Math.min(data.current_page * data.page_size, data.filtered_rows);
    document.getElementById('grid-page-indicator').textContent =
      data.filtered_rows > 0 ? `Showing ${startRow}–${endRow} of ${data.filtered_rows} rows (Page ${data.current_page} of ${data.total_pages})` : '0 rows';

    // Render Headers
    tableHead.innerHTML = `
      <tr>
        <th class="py-2.5 px-3 text-left w-14"># Row</th>
        ${data.columns.map(col => `
          <th class="py-2.5 px-3 text-left whitespace-nowrap">${escapeHtml(col)}</th>
        `).join('')}
      </tr>
    `;

    // Render Rows
    if (data.rows.length === 0) {
      tableBody.innerHTML = `<tr><td colspan="${data.columns.length + 1}" class="text-center py-6 text-slate-400 text-xs">No rows matched your search.</td></tr>`;
      return;
    }

    tableBody.innerHTML = data.rows.map(row => `
      <tr class="transition-colors hover:bg-slate-50">
        <td class="py-1.5 px-3 font-mono text-slate-400 bg-slate-50 font-medium">${row._row_number}</td>
        ${data.columns.map(col => `
          <td class="py-1.5 px-3 text-slate-700 whitespace-nowrap max-w-xs truncate" title="${escapeHtml(row[col] || '')}">
            ${escapeHtml(row[col] || '')}
          </td>
        `).join('')}
      </tr>
    `).join('');

  } catch (err) {
    tableBody.innerHTML = `<tr><td colspan="10" class="text-center py-6 text-rose-600 text-xs">Failed to load data: ${err.message}</td></tr>`;
  }
}

// Field Mappings Tab View
async function loadFieldMappings() {
  const tbody = document.getElementById('mappings-table-body');
  if (!tbody) return;

  if (!state.selectedInput) {
    state.mappings.data = [];
    document.getElementById('map-stat-total').textContent = '0';
    document.getElementById('map-stat-mapped').textContent = '0';
    document.getElementById('map-stat-unmapped').textContent = '0';
    tbody.innerHTML = '<tr><td colspan="6" class="text-center py-6 text-slate-400 text-xs">Upload or select a workbook to view field mappings.</td></tr>';
    return;
  }

  tbody.innerHTML = `<tr><td colspan="6" class="text-center py-6 text-slate-400 text-xs">Analyzing field mappings...</td></tr>`;

  try {
    const res = await fetch(`/api/mappings/preview?file_path=${encodeURIComponent(state.selectedInput)}&template_name=${encodeURIComponent(state.selectedTemplate)}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail);

    state.mappings.data = data.mappings;

    document.getElementById('map-stat-total').textContent = data.total_fields;
    document.getElementById('map-stat-mapped').textContent = data.mapped_count;
    document.getElementById('map-stat-unmapped').textContent = data.unmapped_count;

    renderFilteredMappings();
  } catch (err) {
    tbody.innerHTML = `<tr><td colspan="6" class="text-center py-6 text-rose-600 text-xs">Failed to load mappings: ${err.message}</td></tr>`;
  }
}

function renderFilteredMappings() {
  const tbody = document.getElementById('mappings-table-body');
  if (!tbody) return;

  const filter = state.mappings.filter;
  const search = state.mappings.search.toLowerCase();

  const filtered = state.mappings.data.filter(m => {
    if (filter === 'mapped' && m.status !== 'MAPPED') return false;
    if (filter === 'unmapped' && m.status !== 'UNMAPPED') return false;
    if (filter === 'direct' && m.mapping_type !== 'DIRECT') return false;
    if (filter === 'generated' && m.mapping_type !== 'GENERATED') return false;
    if (filter === 'default' && m.mapping_type !== 'DEFAULT') return false;

    if (search) {
      const matchField = (m.field_name || '').toLowerCase().includes(search);
      const matchObj = (m.object_name || '').toLowerCase().includes(search);
      const matchCol = (m.excel_column || '').toLowerCase().includes(search);
      return matchField || matchObj || matchCol;
    }
    return true;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="text-center py-8 text-slate-400 text-xs">No mappings matched your filter.</td></tr>`;
    return;
  }

  tbody.innerHTML = filtered.map(m => {
    const isMapped = m.status === 'MAPPED';
    const typeBadgeClass =
      m.mapping_type === 'DIRECT' ? 'bg-blue-100 text-blue-800 border-blue-200' :
      m.mapping_type === 'GENERATED' ? 'bg-indigo-100 text-indigo-800 border-indigo-200' :
      m.mapping_type === 'DERIVED' ? 'bg-amber-100 text-amber-800 border-amber-200' :
      m.mapping_type === 'DEFAULT' ? 'bg-purple-100 text-purple-800 border-purple-200' :
      'bg-slate-100 text-slate-600 border-slate-200';

    return `
      <tr class="hover:bg-slate-50 transition-colors">
        <td class="py-2 px-3 font-mono text-slate-600 font-medium">${escapeHtml(m.object_name)}</td>
        <td class="py-2 px-3 font-bold text-slate-900">${escapeHtml(m.field_name)}</td>
        <td class="py-2 px-3 text-slate-700">
          ${m.excel_column ? `<span class="bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200 font-mono text-emerald-900 font-semibold">${escapeHtml(m.excel_column)}</span>` : '<span class="text-slate-400">—</span>'}
        </td>
        <td class="py-2 px-3">
          ${m.mapping_type ? `<span class="px-2 py-0.5 rounded border text-[10px] font-bold ${typeBadgeClass}">${m.mapping_type}</span>` : '<span class="text-slate-400">—</span>'}
        </td>
        <td class="py-2 px-3 text-slate-600 font-mono text-[11px]">
          ${m.default_value ? escapeHtml(m.default_value) : '<span class="text-slate-400">—</span>'}
        </td>
        <td class="py-2 px-3">
          <span class="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-black tracking-wide ${
            isMapped ? 'bg-emerald-100 text-emerald-800 border border-emerald-300' : 'bg-rose-100 text-rose-800 border border-rose-300'
          }">
            ${isMapped ? 'MAPPED' : 'UNMAPPED'}
          </span>
        </td>
      </tr>
    `;
  }).join('');
}

// ================= OUTPUTS TAB: FULL DATA VISIBLE, DOWNLOAD & COPY =================
async function loadOutputFiles(preferredFilename = null) {
  const selectorContainer = document.getElementById('outputs-file-selector');
  if (!selectorContainer) return;

  try {
    const res = await fetch('/api/reports');
    const data = await res.json();
    // Include generated DAT files and batch ZIP bundles.
    state.outputs.files = (data.outputs || []).filter(item => {
      const name = (item.filename || '').toLowerCase();
      return (
        name.endsWith('.dat')
        || name.endsWith('.zip')
      ) && !name.includes('mapping_report') && !name.includes('validation_errors');
    });
    renderDashboard();
    state.outputs.reports = [];

    if (preferredFilename) {
      state.outputs.requestedFilename = preferredFilename;
    }
    if (state.outputs.requestedFilename) {
      state.outputs.activeFile = state.outputs.files.some(
        file => file.filename === state.outputs.requestedFilename,
      ) ? state.outputs.requestedFilename : null;
    } else if (
      state.outputs.activeFile
      && !state.outputs.files.some(file => file.filename === state.outputs.activeFile)
    ) {
      state.outputs.activeFile = null;
    }
    renderOutputsFileSelector();
  } catch (err) {
    showToast('Failed to load output files: ' + err.message, 'error');
  }
}

function renderOutputsFileSelector() {
  const container = document.getElementById('outputs-file-selector');
  if (!container) return;
  container.innerHTML = '';

  // Render generated DAT files and multi-template ZIP bundles.
  const allItems = state.outputs.files
    .filter(item => {
      const name = (item.filename || '').toLowerCase();
      return (
        name.endsWith('.dat')
        || name.endsWith('.zip')
      ) && !name.includes('mapping_report') && !name.includes('validation_errors');
    })
    .map(item => ({
      ...item,
      category: 'output',
      isDat: item.filename.toLowerCase().endsWith('.dat'),
    }));

  if (allItems.length === 0) {
    container.innerHTML = `<div class="p-1 text-xs text-slate-500 font-mono">No generated .dat files yet. Execute in HDL Generator.</div>`;
    document.getElementById('output-data-viewer').innerHTML = `
      <div class="py-12 text-center text-slate-500 text-xs font-mono">
        ${state.outputs.requestedFilename
          ? `No generated ${escapeHtml(state.outputs.requestedFilename)} yet. Execute the selected template in HDL Generator.`
          : 'No HDL .dat files available. Execute transformation in the HDL Generator tab.'}
      </div>
    `;
    return;
  }

  if (
    !state.outputs.activeFile
    && !state.outputs.requestedFilename
    && allItems.length > 0
  ) {
    state.outputs.activeFile = allItems[0].filename;
    state.outputs.activeCategory = allItems[0].category;
  }

  allItems.forEach(item => {
    const isSelected = item.filename === state.outputs.activeFile;
    const btn = document.createElement('button');

    btn.className = `dark-tab-btn ${isSelected ? 'active' : ''}`;
    btn.onclick = () => selectOutputFile(item.filename, item.category);

    btn.innerHTML = `
      <span class="text-[9px] font-mono uppercase px-1.5 py-0.5 rounded font-black ${item.isDat ? 'bg-indigo-950 text-indigo-400 border border-indigo-800' : 'bg-amber-950 text-amber-300 border border-amber-800'}">${item.isDat ? 'DAT' : 'ZIP'}</span>
      <span class="font-mono text-xs font-bold">${escapeHtml(item.filename)}</span>
      <span class="text-[10px] text-slate-500 font-normal">${item.size_kb} KB</span>
    `;
    container.appendChild(btn);
  });

  if (state.outputs.activeFile) {
    displayActiveOutputContent();
  } else {
    const expected = state.outputs.requestedFilename;
    document.getElementById('output-data-viewer').innerHTML = `
      <div class="py-12 text-center text-slate-500 text-xs font-mono">
        ${expected
          ? `No generated ${escapeHtml(expected)} yet. Execute the selected template in HDL Generator.`
          : 'Select a generated HDL file to view its contents.'}
      </div>
    `;
  }
}

function selectOutputFile(filename, category) {
  state.outputs.activeFile = filename;
  state.outputs.activeCategory = category;
  state.outputs.requestedFilename = filename;
  renderOutputsFileSelector();
}

async function displayActiveOutputContent() {
  const viewer = document.getElementById('output-data-viewer');
  const title = document.getElementById('active-file-title');
  const subtitle = document.getElementById('active-file-subtitle');
  const badge = document.getElementById('active-file-badge');
  const downloadBtn = document.getElementById('btn-download-active-file');

  const filename = state.outputs.activeFile;
  const category = state.outputs.activeCategory;
  if (!filename || !viewer) return;

  const isDat = filename.endsWith('.dat');
  const isZip = filename.endsWith('.zip');
  const isXlsx = filename.endsWith('.xlsx');
  const isCsv = filename.endsWith('.csv');

  title.textContent = filename;
  badge.textContent = isDat ? 'DAT' : isZip ? 'ZIP' : isXlsx ? 'XLSX' : 'CSV';
  badge.className = isDat ? 'px-2 py-0.5 rounded bg-indigo-950 text-indigo-400 font-bold font-mono text-[10px] border border-indigo-800' :
                    isZip ? 'px-2 py-0.5 rounded bg-amber-950 text-amber-300 font-bold font-mono text-[10px] border border-amber-800' :
                    isXlsx ? 'px-2 py-0.5 rounded bg-emerald-950 text-emerald-400 font-bold font-mono text-[10px] border border-emerald-800' :
                    'px-2 py-0.5 rounded bg-rose-950 text-rose-400 font-bold font-mono text-[10px] border border-rose-800';

  downloadBtn.href = `/api/reports/download/${category}/${encodeURIComponent(filename)}`;
  const copyButton = document.getElementById('btn-copy-active-file');
  if (copyButton) copyButton.classList.toggle('hidden', isZip);

  viewer.innerHTML = `<div class="py-12 text-center text-slate-500 text-xs font-mono">Loading deliverable records...</div>`;

  try {
    if (isZip) {
      subtitle.textContent = 'Batch HDL bundle archive';
      state.outputs.activeRawText = '';
      viewer.innerHTML = `
        <div class="py-12 px-4 text-center text-slate-400 text-xs font-mono space-y-3">
          <div>This ZIP contains the selected template DAT files in HDL load order.</div>
          <a class="inline-flex btn-dark-glow" href="${downloadBtn.href}">Download ZIP bundle</a>
        </div>
      `;
      return;
    }

    if (isDat) {
      // DAT file preview with DARK THEME and NORMAL compact line spacing
      const res = await fetch(`/api/reports/preview-dat/${encodeURIComponent(filename)}?max_lines=1500`);
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail);

      subtitle.textContent = `${data.metadata_lines} METADATA lines • ${data.merge_lines} MERGE lines • Total ${data.lines.length} lines`;
      state.outputs.activeRawText = data.lines.join('\n');

      viewer.innerHTML = `
        <div class="dat-viewer-dark">
          ${data.lines.map((line, idx) => {
            const isMeta = line.startsWith('METADATA|');
            return `
              <div class="dat-row">
                <span class="dat-linenum">${idx + 1}</span>
                <span class="dat-content ${isMeta ? 'dat-content-metadata' : 'dat-content-merge'}">${escapeHtml(line)}</span>
              </div>
            `;
          }).join('')}
        </div>
      `;

    } else if (isCsv) {
      // CSV report preview with DARK THEME table (e.g. validation_errors.csv)
      const res = await fetch(`/api/reports/preview-csv/${encodeURIComponent(filename)}?max_rows=500`);
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail);

      subtitle.textContent = `${data.total_rows} total records in audit report`;

      // Build TSV for copying to Excel/text
      const headerRow = data.columns.join('\t');
      const rowsText = data.rows.map(r => data.columns.map(c => r[c] ?? '').join('\t')).join('\n');
      state.outputs.activeRawText = headerRow + '\n' + rowsText;

      if (data.rows.length === 0) {
        viewer.innerHTML = `<div class="py-12 text-center text-xs text-slate-500 font-medium">File is empty (0 validation errors reported).</div>`;
        return;
      }

      viewer.innerHTML = `
        <div class="dark-table-container">
          <table class="dark-table">
            <thead>
              <tr>
                ${data.columns.map(c => `<th>${escapeHtml(c)}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
              ${data.rows.map(r => `
                <tr>
                  ${data.columns.map(c => `
                    <td class="${c === 'Error' ? 'text-rose-400 font-bold' : ''}">
                      ${escapeHtml(r[c] ?? '')}
                    </td>
                  `).join('')}
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      `;

    } else if (isXlsx) {
      // Excel mapping report preview with DARK THEME table
      const res = await fetch(`/api/reports/preview-excel/${encodeURIComponent(filename)}?max_rows=500`);
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail);

      subtitle.textContent = `${data.total_rows} total attribute mapping definitions`;

      // Build TSV for copying
      const headerRow = data.columns.join('\t');
      const rowsText = data.rows.map(r => data.columns.map(c => r[c] ?? '').join('\t')).join('\n');
      state.outputs.activeRawText = headerRow + '\n' + rowsText;

      viewer.innerHTML = `
        <div class="dark-table-container">
          <table class="dark-table">
            <thead>
              <tr>
                ${data.columns.map(c => `<th>${escapeHtml(c)}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
              ${data.rows.map(r => `
                <tr>
                  ${data.columns.map(c => `
                    <td class="${c === 'status' && r[c] === 'MAPPED' ? 'text-emerald-400 font-bold' : ''}">
                      ${escapeHtml(r[c] ?? '')}
                    </td>
                  `).join('')}
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      `;
    }
  } catch (err) {
    viewer.innerHTML = `<div class="py-12 text-center text-xs text-rose-400">Error reading file content: ${err.message}</div>`;
  }
}

function copyActiveFileOutput() {
  if (!state.outputs.activeRawText) {
    showToast('No content available to copy.', 'warning');
    return;
  }

  navigator.clipboard.writeText(state.outputs.activeRawText).then(() => {
    const copyBtnText = document.getElementById('btn-copy-text');
    if (copyBtnText) {
      const originalText = copyBtnText.textContent;
      copyBtnText.textContent = 'Copied to Clipboard!';
      setTimeout(() => {
        copyBtnText.textContent = originalText;
      }, 2000);
    }
    showToast(`Copied ${state.outputs.activeFile} contents to clipboard`, 'success');
  }).catch(err => {
    showToast('Failed to copy: ' + err.message, 'error');
  });
}

// Database Schema Viewer
async function loadDatabaseSchema() {
  const container = document.getElementById('db-schema-container');
  if (!container) return;

  container.innerHTML = `<div class="text-center py-6 text-slate-400 text-xs">Loading database structure...</div>`;

  try {
    const res = await fetch(`/api/templates/${encodeURIComponent(state.selectedTemplate)}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail);

    if (data.error || !data.objects || data.objects.length === 0) {
      container.innerHTML = `
        <div class="p-4 rounded-xl bg-amber-50 border border-amber-300 text-amber-950 space-y-1 text-xs">
          <div class="font-bold">Template Notice: ${data.template_name}</div>
          <p class="text-amber-800">${data.error || 'No active objects or attributes configured for this template in MySQL.'}</p>
        </div>
      `;
      return;
    }

    container.innerHTML = `
      <div class="db-schema-summary p-4 rounded-xl bg-gradient-to-r from-teal-50/60 via-cyan-50/40 to-slate-50 border border-teal-200 flex flex-wrap items-center justify-between gap-2.5 shadow-2xs">
        <div>
          <h3 class="text-xs font-bold text-slate-900">${data.template_name} Template Configuration</h3>
          <p class="text-[11px] text-slate-600">Output: <code class="font-mono text-teal-800 font-bold">${data.output_file_name}</code> • Source Owner: <code class="font-mono text-emerald-800 font-bold">${data.source_system_owner}</code></p>
        </div>
        <span class="px-2.5 py-1 rounded-md bg-teal-100 text-teal-800 border border-teal-300 text-xs font-bold">
          ${data.objects.length} Objects
        </span>
      </div>

      <div class="space-y-3 pt-1">
        ${data.objects.map(obj => `
          <div class="card-3d overflow-hidden">
            <div class="db-schema-object-header p-3 bg-slate-50/90 border-b border-slate-200 flex items-center justify-between">
              <div class="flex items-center gap-2">
                <span class="w-2 h-2 rounded-full bg-teal-600"></span>
                <span class="font-black text-xs text-slate-900">${obj.object_name}</span>
                <span class="text-[10px] text-slate-500 font-medium">(${obj.attribute_count} attributes)</span>
              </div>
              <span class="text-[11px] text-teal-800 font-bold bg-teal-50 px-2 py-0.5 rounded border border-teal-200">${obj.mappings_count} mapped</span>
            </div>
            <div class="db-schema-attribute-grid p-3.5 grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 gap-2 bg-white">
              ${obj.attributes.map(attr => `
                <div class="db-schema-attribute ${attr.required ? 'is-required border border-amber-300 border-b-2 border-b-amber-400 bg-amber-50/70 text-amber-950' : 'is-optional border border-slate-200 border-b-2 border-b-slate-300 bg-slate-50/80 text-slate-800'} p-2 rounded-lg text-xs">
                  <div class="font-bold truncate" title="${attr.field_name}">
                    ${attr.field_name}
                  </div>
                  <div class="text-[10px] ${attr.required ? 'text-amber-800' : 'text-slate-500'} flex justify-between mt-1 font-mono font-medium">
                    <span>${attr.data_type}</span>
                    ${attr.required ? '<span class="db-schema-required-badge font-black bg-amber-200/80 px-1 rounded text-[9px]">REQ</span>' : ''}
                  </div>
                </div>
              `).join('')}
            </div>
          </div>
        `).join('')}
      </div>
    `;
  } catch (err) {
    container.innerHTML = `<div class="p-4 text-center text-xs text-rose-600">Error loading schema: ${err.message}</div>`;
  }
}

// Setup Event Listeners
function setupEventListeners() {
  // Execute Generate Button
  document.getElementById('btn-execute-generate')?.addEventListener('click', executeGeneration);

  // Auto-detect button
  document.getElementById('btn-autodetect')?.addEventListener('click', triggerAutoDetect);

  // Copy active output button
  document.getElementById('btn-copy-active-file')?.addEventListener('click', copyActiveFileOutput);

  // Test DB connection button
  document.getElementById('btn-test-db')?.addEventListener('click', async () => {
    showToast('Testing MySQL connection...', 'info');
    const result = await refreshSystemStatus();
    if (result && result.database && result.database.status === 'connected') {
      showToast(`Connected to ${result.database.database} (MySQL ${result.database.version})`, 'success');
    } else {
      showToast('Database connection failed: ' + (result?.database?.error || 'Unknown error'), 'error');
    }
  });

  // Spreadsheet Data Grid Controls
  document.getElementById('grid-search-input')?.addEventListener('input', (e) => {
    state.dataGrid.search = e.target.value;
    state.dataGrid.page = 1;
    loadDataGrid();
  });

  document.getElementById('grid-pagesize-select')?.addEventListener('change', (e) => {
    state.dataGrid.pageSize = parseInt(e.target.value, 10);
    state.dataGrid.page = 1;
    loadDataGrid();
  });

  document.getElementById('grid-first-btn')?.addEventListener('click', () => {
    if (state.dataGrid.page > 1) {
      state.dataGrid.page = 1;
      loadDataGrid();
    }
  });

  document.getElementById('grid-prev-btn')?.addEventListener('click', () => {
    if (state.dataGrid.page > 1) {
      state.dataGrid.page--;
      loadDataGrid();
    }
  });

  document.getElementById('grid-next-btn')?.addEventListener('click', () => {
    if (state.dataGrid.page < state.dataGrid.totalPages) {
      state.dataGrid.page++;
      loadDataGrid();
    }
  });

  document.getElementById('grid-last-btn')?.addEventListener('click', () => {
    if (state.dataGrid.page < state.dataGrid.totalPages) {
      state.dataGrid.page = state.dataGrid.totalPages;
      loadDataGrid();
    }
  });

  // Mappings Filters
  document.querySelectorAll('.mapping-filter-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.mapping-filter-btn').forEach(b => {
        b.className = 'mapping-filter-btn px-3 py-1 rounded-md text-xs font-medium bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 transition-colors';
      });
      btn.className = 'mapping-filter-btn px-3 py-1 rounded-md text-xs font-bold bg-indigo-600 text-white transition-colors';
      state.mappings.filter = btn.getAttribute('data-filter');
      renderFilteredMappings();
    });
  });

  document.getElementById('mappings-search-input')?.addEventListener('input', (e) => {
    state.mappings.search = e.target.value;
    renderFilteredMappings();
  });
}

// Utility: escape HTML
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
