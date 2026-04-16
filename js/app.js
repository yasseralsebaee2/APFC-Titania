    const JSON_URL = 'https://raw.githubusercontent.com/yasseralsebaee2/APFC-Data/refs/heads/main/_apfc-pile-asbuilt.json_';
    const USERS_URL = 'https://raw.githubusercontent.com/yasseralsebaee2/APFC-Data/refs/heads/main/_apfc-users.json_';
    const MANPOWER_URL = 'https://raw.githubusercontent.com/yasseralsebaee2/APFC-Data/refs/heads/main/_apfc_manpower.json_';
    const COMPANY_MANPOWER_URL = 'https://raw.githubusercontent.com/yasseralsebaee2/APFC-Data/refs/heads/main/_apfc_manpowers.json_';
    const ACCESS_REQUEST_EMAIL = 'yasser.alsebaee@granadaeurope.com';
    const DEFAULT_PROJECT = 'Titania';
    const AUTH_STORAGE_KEY = 'apfcDashboardAuth';

    const els = {
      authShell: document.getElementById('authShell'),
      authForm: document.getElementById('authForm'),
      authLoginInput: document.getElementById('authLoginInput'),
      authPasswordInput: document.getElementById('authPasswordInput'),
      authSubmitBtn: document.getElementById('authSubmitBtn'),
      authRequestAccessBtn: document.getElementById('authRequestAccessBtn'),
      authRequestPanel: document.getElementById('authRequestPanel'),
      authRequestNameInput: document.getElementById('authRequestNameInput'),
      authRequestEmailInput: document.getElementById('authRequestEmailInput'),
      authRequestError: document.getElementById('authRequestError'),
      authRequestCancelBtn: document.getElementById('authRequestCancelBtn'),
      authRequestSendBtn: document.getElementById('authRequestSendBtn'),
      authError: document.getElementById('authError'),
      signOutBtn: document.getElementById('signOutBtn'),
      projectSelector: document.getElementById('projectSelector'),
      projectScopeBtn: document.getElementById('projectScopeBtn'),
      dataSourceChip: document.getElementById('dataSourceChip'),
      overviewDateModeButtons: Array.from(document.querySelectorAll('#overviewDateModeToggle button')),
      refreshDashboardBtn: document.getElementById('refreshDashboardBtn'),
      pageSubtitle: document.getElementById('pageSubtitle'),
      kpiTotal: document.getElementById('kpiTotal'),
      kpiExecuted: document.getElementById('kpiExecuted'),
      kpiRemaining: document.getElementById('kpiRemaining'),
      kpiRigs: document.getElementById('kpiRigs'),
      kpiAvg: document.getElementById('kpiAvg'),
      kpiAvgMetaL: document.getElementById('kpiAvgMetaL'),
      kpiAvgMeta: document.getElementById('kpiAvgMeta'),
      kpiEta: document.getElementById('kpiEta'),
      kpiEtaMeta: document.getElementById('kpiEtaMeta'),
      kpiEtaDays: document.getElementById('kpiEtaDays'),
      kpiTotalMetaR: document.getElementById('kpiTotalMetaR'),
      kpiCompletedMetaL: document.getElementById('kpiCompletedMetaL'),
      kpiCompletedMetaR: document.getElementById('kpiCompletedMetaR'),
      kpiRemainingMetaR: document.getElementById('kpiRemainingMetaR'),
      chartTitle: document.getElementById('chartTitle'),
      chartTag: document.getElementById('chartTag'),
      chartGrid: document.getElementById('chartGrid'),
      chartSeries: document.getElementById('chartSeries'),
      chartLabels: document.getElementById('chartLabels'),
      chartXAxis: document.getElementById('chartXAxis'),
      chartWrap: document.getElementById('chartWrap'),
      chartSvg: document.getElementById('chartSvg'),
      chartTooltip: document.getElementById('chartTooltip'),
      overviewSeriesLegend: document.getElementById('overviewSeriesLegend'),

      timelineStartDate: document.getElementById('timelineStartDate'),
      timelineEndDate: document.getElementById('timelineEndDate'),
      timelineApplyBtn: document.getElementById('timelineApplyBtn'),
      timelineClearBtn: document.getElementById('timelineClearBtn'),
      timelinePreset7Btn: document.getElementById('timelinePreset7Btn'),
      timelinePreset14Btn: document.getElementById('timelinePreset14Btn'),
      timelinePreset30Btn: document.getElementById('timelinePreset30Btn'),
      timelinePileList: document.getElementById('timelinePileList'),
      timelineTableBody: document.getElementById('timelineTableBody'),
      timelineScopeCount: document.getElementById('timelineScopeCount'),
      timelineChartWrap: document.getElementById('timelineChartWrap'),
      timelineSvg: document.getElementById('timelineSvg'),
      timelineEmpty: document.getElementById('timelineEmpty'),
      timelineSubtitle: document.getElementById('timelineSubtitle'),
      timelineTag: document.getElementById('timelineTag'),
      timelineZoomOutBtn: document.getElementById('timelineZoomOutBtn'),
      timelineZoomResetBtn: document.getElementById('timelineZoomResetBtn'),
      timelineZoomInBtn: document.getElementById('timelineZoomInBtn'),
      timelineSummarySubtitle: document.getElementById('timelineSummarySubtitle'),
      timelinePieSubtitle: document.getElementById('timelinePieSubtitle'),
      timelinePileTag: document.getElementById('timelinePileTag'),
      timelinePileCount: document.getElementById('timelinePileCount'),
      timelinePileCountMeta: document.getElementById('timelinePileCountMeta'),
      timelinePeriodValue: document.getElementById('timelinePeriodValue'),
      timelinePeriodMeta: document.getElementById('timelinePeriodMeta'),
      timelineGrossAvg: document.getElementById('timelineGrossAvg'),
      timelineCycleAvg: document.getElementById('timelineCycleAvg'),
      timelinePieSvg: document.getElementById('timelinePieSvg'),
      timelinePieTotal: document.getElementById('timelinePieTotal'),

      granularityToggleButtons: Array.from(document.querySelectorAll('#granularityToggleGroup .chart-toggle')),
      modeLabelCumulative: document.getElementById('modeLabelCumulative'),
      metricToggleButtons: Array.from(document.querySelectorAll('#metricToggleGroup .chart-toggle')),
      matrixBody: document.getElementById('matrixBody'),
      kpiRow: document.getElementById('kpiRow'),
      pageOverview: document.getElementById('pageOverview'),
      pageMap: document.getElementById('pageMap'),
      pageProduction: document.getElementById('pageProduction'),
      pageUtilization: document.getElementById('pageUtilization'),
      pageManpower: document.getElementById('pageManpower'),
      pageCompanyManpower: document.getElementById('pageCompanyManpower'),
      pageTimeline: document.getElementById('pageTimeline'),
      pageCost: document.getElementById('pageCost'),
      costTableHeadRow: document.querySelector('.financial-table thead tr'),
      costTableBody: document.getElementById('costTableBody'),
      costTrendWrap: document.getElementById('costTrendWrap'),
      costTrendSvg: document.getElementById('costTrendSvg'),
      costLmWrap: document.getElementById('costLmWrap'),
      costLmSvg: document.getElementById('costLmSvg'),
      manpowerTableBody: document.getElementById('manpowerTableBody'),
      manpowerHistSvg: document.getElementById('manpowerHistSvg'),
      equipmentTableBody: document.getElementById('equipmentTableBody'),
      equipmentHistSvg: document.getElementById('equipmentHistSvg'),
      companyManpowerLayoutSelect: document.getElementById('companyManpowerLayoutSelect'),
      companyManpowerDesignationSelect: document.getElementById('companyManpowerDesignationSelect'),
      companyManpowerScopeButtons: Array.from(document.querySelectorAll('#companyManpowerScopeToggle button')),
      companyManpowerExportBtn: document.getElementById('companyManpowerExportBtn'),
      companyManpowerSubtitle: document.getElementById('companyManpowerSubtitle'),
      companyManpowerSummaryMeta: document.getElementById('companyManpowerSummaryMeta'),
      companyManpowerColGroup: document.getElementById('companyManpowerColGroup'),
      companyManpowerHeadRow: document.getElementById('companyManpowerHeadRow'),
      companyManpowerMatrixBody: document.getElementById('companyManpowerMatrixBody'),
      companyExportPanel: document.getElementById('companyExportPanel'),
      companyExportProjectSelect: document.getElementById('companyExportProjectSelect'),
      companyExportCancelBtn: document.getElementById('companyExportCancelBtn'),
      companyExportConfirmBtn: document.getElementById('companyExportConfirmBtn'),
      utilizationTableBody: document.getElementById('utilizationTableBody'),
      utilizationSvg: document.getElementById('utilizationSvg'),
      utilizationChartTag: document.getElementById('utilizationChartTag'),
      utilizationUtilTicks: document.getElementById('utilizationUtilTicks'),
      utilizationLmTicks: document.getElementById('utilizationLmTicks'),
      utilizationLegend: document.getElementById('utilizationLegend'),
      utilizationChartWrap: document.getElementById('utilizationChartWrap'),
      projectMapFrame: document.getElementById('projectMapFrame'),
      navButtons: Array.from(document.querySelectorAll('.nav-btn')),
      prodSvgs: {
        gross: document.getElementById('prodSvgGross'),
        drilling: document.getElementById('prodSvgDrilling'),
        cage: document.getElementById('prodSvgCage'),
        concrete: document.getElementById('prodSvgConcrete'),
        overconsumption: document.getElementById('prodSvgOverconsumption'),
        overexcavation: document.getElementById('prodSvgOverexcavation')
      },
      prodToolButtons: Array.from(document.querySelectorAll('.mini-toggle[data-prod-key]')),
      utilizationModeButtons: Array.from(document.querySelectorAll('[data-util-mode]'))
    };

    let rawRows = [];
    let manpowerRows = [];
    let companyManpowerRows = [];
    let usersDirectory = [];
    let currentUser = null;
    let selectedProject = DEFAULT_PROJECT;
    let selectedPlot = '';
    let chartMode = 'daily'; // daily or cumulative (toggle)
    let chartMetric = 'piles';
    let chartGranularity = 'day';
    let activePage = 'overview';
    let companyManpowerScopeMode = 'filtered';
    let companyManpowerDesignationFilter = 'all';
    let companyManpowerColumnWidths = {};
    let utilizationMode = 'daily';
    let overviewDateMode = 'shift'; // shared reporting mode for Overview + Production only
    let prodState = {
      gross: 'day',
      drilling: 'day',
      cage: 'day',
      concrete: 'day',
      overconsumption: 'day',
      overexcavation: 'day'
    };
    let productionChartsInitialized = false;

    let timelineState = {
      start: '',
      end: '',
      pile: 'all',
      zoom: 1
    };
    let timelineTooltipEl = null;

    function normalizeText(value) {
      return String(value || '').trim();
    }

    function normalizeLogin(value) {
      return normalizeText(value).toLowerCase();
    }

    function escapeHtml(value) {
      return normalizeText(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    }

    function isAllPlotsValue(value) {
      const normalized = normalizeText(value).toLowerCase();
      return !normalized || normalized === '-' || normalized === '—' || normalized === 'all' || normalized === 'all plots' || normalized === 'all plot';
    }

    function isAllProjectsValue(value) {
      const normalized = normalizeText(value).toLowerCase();
      return !normalized || normalized === 'all';
    }

    function normalizeDateString(value) {
      const raw = String(value || '').trim();
      if (!raw) return '';
      const d = new Date(raw);
      if (Number.isNaN(d.getTime())) return '';
      return d.toISOString().slice(0, 10);
    }

    function getStoredAuthSession() {
      try {
        const raw = window.localStorage.getItem(AUTH_STORAGE_KEY);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        return parsed && typeof parsed === 'object' ? parsed : null;
      } catch {
        return null;
      }
    }

    function storeAuthSession(user) {
      window.localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(user));
    }

    function clearAuthSession() {
      window.localStorage.removeItem(AUTH_STORAGE_KEY);
    }

    function sanitizeUserRecord(record) {
      return {
        username: normalizeText(record.username || record.Username || record.userName || record.UserName),
        password: String(record.password || record.Password || ''),
        email: normalizeText(record.email || record.Email),
        type: normalizeText(record.type || record.Type).toLowerCase() || 'user',
        name: normalizeText(record.username || record.Username || record.email || record.Email),
        project: normalizeText(record.project || record.Project) || DEFAULT_PROJECT,
        plot: normalizeText(record.plot || record.Plot)
      };
    }

    function extractUserList(data) {
      if (Array.isArray(data)) return data;
      if (!data || typeof data !== 'object') return [];

      const directLists = [
        data.body,
        data.Body,
        data.users,
        data.Users,
        data.items,
        data.Items,
        data.records,
        data.Records,
        data.data,
        data.Data
      ];

      for (const candidate of directLists) {
        if (Array.isArray(candidate)) return candidate;
      }

      const objectValues = Object.values(data);
      if (objectValues.length && objectValues.every(value => value && typeof value === 'object' && !Array.isArray(value))) {
        return objectValues;
      }

      return [];
    }

    function extractManpowerList(data) {
      const items = Array.isArray(data) ? data : (!data || typeof data !== 'object' ? [] : [data.body, data.rows, data.items, data.data].find(Array.isArray) || []);
      return items.filter(row => {
        const keys = Object.keys(row || {}).map(key => key.toLowerCase());
        return keys.includes('date') || keys.includes('projectmanager') || keys.includes('siteengineer') || keys.includes('project manager');
      });
    }

    function extractCompanyManpowerList(data) {
      const items = Array.isArray(data) ? data : (!data || typeof data !== 'object' ? [] : [data.body, data.rows, data.items, data.data].find(Array.isArray) || []);
      return items.filter(row => {
        const keys = Object.keys(row || {}).map(key => key.toLowerCase());
        return keys.includes('employee number') || keys.includes('employee name') || keys.includes('designation');
      });
    }

    const COMPANY_PROJECT_ALIASES = {
      Titania: ['titania', '89'],
      Vintage: ['vintage']
    };

    function getCompanyProjectToken(value) {
      const normalized = normalizeText(value).toLowerCase();
      if (!normalized) return '';
      for (const [label, aliases] of Object.entries(COMPANY_PROJECT_ALIASES)) {
        if ([label.toLowerCase(), ...aliases].includes(normalized)) return label.toLowerCase();
      }
      return normalized;
    }

    function getCompanyProjectLabel(value) {
      const normalized = normalizeText(value);
      if (!normalized) return 'Unassigned';
      for (const [label, aliases] of Object.entries(COMPANY_PROJECT_ALIASES)) {
        if ([label.toLowerCase(), ...aliases].includes(normalized.toLowerCase())) return label;
      }
      return normalized;
    }

    function sanitizeCompanyEmployee(row) {
      return {
        employeeNumber: normalizeText(row['Employee Number'] || row.employeeNumber || row.employee_number),
        employeeName: normalizeText(row['Employee Name'] || row.employeeName || row.employee_name),
        designation: normalizeText(row.Designation || row.designation),
        project: getCompanyProjectLabel(row.Project || row.project),
        projectRaw: normalizeText(row.Project || row.project),
        shift: normalizeText(row.Shift || row.shift),
        campNumber: normalizeText(row['Camp Number'] || row.campNumber || row.camp_number),
        roomNumber: normalizeText(row['Room Number'] || row.roomNumber || row.room_number),
        joiningDate: normalizeText(row['Joining Date'] || row.joiningDate || row.joining_date),
        remarks: normalizeText(row.Remarks || row.remarks)
      };
    }

    function getScopeLabel() {
      if (!selectedProject) return 'No Project';
      if (isAllPlotsValue(selectedPlot)) return selectedProject;
      return `${selectedProject} / ${selectedPlot}`;
    }

    function getScopeSubtitle() {
      if (!selectedProject) return 'Project Dashboard';
      if (isAllPlotsValue(selectedPlot)) return `Project ${selectedProject}`;
      return `Project ${selectedProject} • Plot ${selectedPlot}`;
    }

    function isManagerUser() {
      const role = normalizeText(currentUser?.type).toLowerCase();
      return role === 'manager';
    }

    function canSelectProjects() {
      return !!currentUser && (isManagerUser() || isAllProjectsValue(currentUser?.project));
    }

    function canAccessPage(page) {
      if (!currentUser) return false;
      if (page === 'cost') return isManagerUser();
      return true;
    }

    function updateUserContextUi() {
      if (els.pageSubtitle) els.pageSubtitle.textContent = getScopeSubtitle();
      if (els.projectScopeBtn) els.projectScopeBtn.textContent = getScopeLabel();

      if (els.projectSelector) {
        const showProjectSelector = canSelectProjects();
        els.projectSelector.hidden = !showProjectSelector;
        els.projectSelector.style.display = showProjectSelector ? '' : 'none';
      }

      if (els.projectScopeBtn) {
        const showProjectButton = !canSelectProjects();
        els.projectScopeBtn.hidden = !showProjectButton;
        els.projectScopeBtn.style.display = showProjectButton ? '' : 'none';
      }

      els.navButtons.forEach(btn => {
        const label = btn.querySelector('.nav-label')?.textContent?.trim();
        if (label === 'Cost') {
          const allowCost = isManagerUser();
          btn.hidden = !allowCost;
          btn.style.display = allowCost ? '' : 'none';
          btn.classList.toggle('role-hidden', !allowCost);
        }
      });

      if (els.pageCost) {
        const allowCost = isManagerUser();
        els.pageCost.hidden = !allowCost;
        els.pageCost.style.display = allowCost ? '' : 'none';
      }

      if (els.signOutBtn) {
        els.signOutBtn.hidden = !currentUser;
      }

      if (!canAccessPage(activePage)) {
        activePage = 'overview';
      }
    }

    function setAuthLocked(isLocked, errorMessage = '') {
      document.body.classList.toggle('auth-locked', isLocked);
      if (els.authShell) els.authShell.setAttribute('aria-hidden', isLocked ? 'false' : 'true');
      if (els.authError) els.authError.textContent = errorMessage || '';
      if (isLocked) {
        els.dataSourceChip.textContent = 'Sign In Required';
      }
    }

    function toggleRequestAccessPanel(isOpen) {
      if (!els.authRequestPanel) return;
      els.authRequestPanel.hidden = !isOpen;
      if (isOpen && els.authRequestNameInput) {
        els.authRequestNameInput.focus();
      }
    }

    function setRequestAccessError(message = '') {
      if (!els.authRequestError) return;
      els.authRequestError.textContent = message || '';
    }

    function validateEmail(value) {
      const email = normalizeText(value);
      if (!email) return false;
      return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
    }

    function submitAccessRequest() {
      const requestName = normalizeText(els.authRequestNameInput?.value);
      const requestEmail = normalizeText(els.authRequestEmailInput?.value);
      setRequestAccessError('');

      if (!requestName) {
        setRequestAccessError('Please enter your name.');
        els.authRequestNameInput?.focus();
        return;
      }
      if (!requestEmail) {
        setRequestAccessError('Please enter your email address.');
        els.authRequestEmailInput?.focus();
        return;
      }
      if (!validateEmail(requestEmail)) {
        setRequestAccessError('Please enter a valid email address.');
        els.authRequestEmailInput?.focus();
        return;
      }

      const subject = `APFC Dashboard Access Request - ${requestName}`;
      const bodyLines = [
        'Hello,',
        '',
        'Please grant me access to APFC Project Dashboard.',
        '',
        `Name: ${requestName}`,
        `Email: ${requestEmail}`,
        '',
        'Thank you.'
      ];
      const mailtoUrl = `mailto:${encodeURIComponent(ACCESS_REQUEST_EMAIL)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(bodyLines.join('\n'))}`;
      window.location.href = mailtoUrl;
      setAuthLocked(true, 'Request draft opened in your email app. Please send it.');
      toggleRequestAccessPanel(false);
      setRequestAccessError('');
      if (els.authRequestNameInput) els.authRequestNameInput.value = '';
      if (els.authRequestEmailInput) els.authRequestEmailInput.value = '';
    }

    async function loadUsersDirectory() {
      const res = await fetch(USERS_URL, { cache: 'no-store' });
      if (!res.ok) throw new Error(`Users source HTTP ${res.status}`);
      const data = await res.json();
      const list = extractUserList(data);
      usersDirectory = list.map(sanitizeUserRecord).filter(user => user.username && user.password && user.project);
      if (!usersDirectory.length) throw new Error('No users found in the APFC users source');
    }

    function findUserRecord(loginValue) {
      const normalized = normalizeLogin(loginValue);
      return usersDirectory.find(user => normalizeLogin(user.username) === normalized || normalizeLogin(user.email) === normalized) || null;
    }

    function broadcastAuthContext() {
      if (!els.projectMapFrame?.contentWindow) return;
      els.projectMapFrame.contentWindow.postMessage({
        type: 'AUTH_CONTEXT_UPDATED',
        payload: currentUser ? {
          project: selectedProject,
          plot: selectedPlot,
          overviewDateMode,
          type: currentUser.type,
          name: currentUser.name,
          username: currentUser.username,
          email: currentUser.email
        } : null
      }, window.location.origin);
    }

    function triggerMapFocusAnimation() {
      if (!els.projectMapFrame?.contentWindow) return;
      els.projectMapFrame.contentWindow.postMessage({
        type: 'MAP_FOCUS_ANIMATE'
      }, window.location.origin);
    }

    function bindMapFrameSync() {
      if (!els.projectMapFrame || els.projectMapFrame.dataset.syncBound === 'true') return;
      els.projectMapFrame.addEventListener('load', () => {
        broadcastAuthContext();
        setTimeout(() => triggerMapFocusAnimation(), 80);
      });
      els.projectMapFrame.dataset.syncBound = 'true';
    }

    function persistScopedSession() {
      if (!currentUser) return;
      storeAuthSession({
        ...currentUser,
        selectedProject,
        selectedPlot
      });
    }

    function syncProjectScopeFromData() {
      if (!rawRows.length) return;

      if (canSelectProjects()) {
        const projects = getProjectList(rawRows).filter(project => normalizeText(project).toLowerCase() !== 'all');
        if (projects.length) {
          const currentProjectIsValid = projects.includes(selectedProject);
          if (!currentProjectIsValid) {
            selectedProject = projects.includes(DEFAULT_PROJECT) ? DEFAULT_PROJECT : projects[0];
          }
          renderProjectOptions(projects);
        }
        persistScopedSession();
        return;
      }

      selectedProject = currentUser?.project || DEFAULT_PROJECT;
      persistScopedSession();
    }

    function applyUserSession(user) {
      currentUser = user;
      selectedProject = isAllProjectsValue(user?.project) ? DEFAULT_PROJECT : (user?.project || DEFAULT_PROJECT);
      selectedPlot = user?.plot || '';
      timelineState.pile = 'all';
      companyManpowerScopeMode = 'filtered';
      companyManpowerDesignationFilter = 'all';
      updateUserContextUi();
      if (currentUser) {
        persistScopedSession();
        setAuthLocked(false);
      } else {
        clearAuthSession();
      }
      broadcastAuthContext();
    }

    function formatDateKeyInTimezone(value, timeZone = 'Asia/Dubai') {
      const d = value instanceof Date ? value : new Date(value);
      if (!(d instanceof Date) || Number.isNaN(d.getTime())) return '';
      const parts = new Intl.DateTimeFormat('en-CA', {
        timeZone,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
      }).formatToParts(d);
      const year = parts.find(part => part.type === 'year')?.value;
      const month = parts.find(part => part.type === 'month')?.value;
      const day = parts.find(part => part.type === 'day')?.value;
      return year && month && day ? `${year}-${month}-${day}` : '';
    }

    function getProductionDateKey(row) {
      const concreteEndRaw = row?.asbuilt_concreteEnd || row?.asbuilt_concreteend || row?.asbuilt_ConcreteEnd;
      const concreteEnd = concreteEndRaw ? new Date(concreteEndRaw) : null;

      if (concreteEnd instanceof Date && !Number.isNaN(concreteEnd.getTime())) {
        const d = new Date(concreteEnd);
        const hour = d.getUTCHours();
        const minute = d.getUTCMinutes();
        const totalMinutes = hour * 60 + minute;

        // Production/reporting date concept for production/timeline pages:
        // anything completed before 07:00 belongs to the previous production date.
        if (totalMinutes < (7 * 60)) {
          d.setUTCDate(d.getUTCDate() - 1);
        }
        return d.toISOString().slice(0, 10);
      }

      return normalizeDateString(row?.castingDate);
    }

    function getOverviewDateKey(row) {
      const concreteEndRaw = row?.asbuilt_concreteEnd || row?.asbuilt_concreteend || row?.asbuilt_ConcreteEnd;
      const concreteEnd = concreteEndRaw ? new Date(concreteEndRaw) : null;

      if (concreteEnd instanceof Date && !Number.isNaN(concreteEnd.getTime())) {
        const d = new Date(concreteEnd);

        if (overviewDateMode === 'calendar') {
          return formatDateKeyInTimezone(d);
        }

        const totalMinutes = d.getUTCHours() * 60 + d.getUTCMinutes();
        if (totalMinutes < (7 * 60)) {
          d.setUTCDate(d.getUTCDate() - 1);
        }
        return d.toISOString().slice(0, 10);
      }

      return normalizeDateString(row?.castingDate);
    }


    function formatDateLabel(value) {
      if (!value) return 'Ã¢â‚¬â€';
      const d = new Date(value);
      if (Number.isNaN(d.getTime())) return 'Ã¢â‚¬â€';
      return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: '2-digit', timeZone: 'UTC' });
    }

    function formatDateFullLabel(value) {
      if (!value) return 'Ã¢â‚¬â€';
      const d = new Date(value);
      if (Number.isNaN(d.getTime())) return 'Ã¢â‚¬â€';
      return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', timeZone: 'UTC' });
    }

    function formatShortDateLabel(value) {
      if (!value) return 'Ã¢â‚¬â€';
      const d = new Date(value);
      if (Number.isNaN(d.getTime())) return 'Ã¢â‚¬â€';
      return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', timeZone: 'UTC' });
    }

    function formatDayNumberLabel(value) {
      if (!value) return 'Ã¢â‚¬â€';
      const d = new Date(value);
      if (Number.isNaN(d.getTime())) return 'Ã¢â‚¬â€';
      return String(d.getUTCDate());
    }

    function formatMonthShortLabel(value) {
      if (!value) return 'Ã¢â‚¬â€';
      const d = new Date(value);
      if (Number.isNaN(d.getTime())) return 'Ã¢â‚¬â€';
      return d.toLocaleDateString('en-GB', { month: 'short', timeZone: 'UTC' });
    }

    function todayKey() {
      const now = new Date();
      return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
    }

    function previousDayKey() {
      const d = new Date(`${todayKey()}T00:00:00Z`);
      d.setUTCDate(d.getUTCDate() - 1);
      return d.toISOString().slice(0, 10);
    }

    function isWorkingDay(dateKey) {
      const d = new Date(`${dateKey}T00:00:00Z`);
      const day = d.getUTCDay();
      return day >= 1 && day <= 6;
    }

    function isExecutedRow(row) {
      return row.isExecuted === true;
    }

    function getExecutedRows(rows) {
      return rows.filter(isExecutedRow);
    }

    function getProjectList(rows) {
      const projects = Array.from(new Set(rows.map(r => normalizeText(r.project)).filter(Boolean))).sort((a, b) => a.localeCompare(b));
      if (projects.includes(DEFAULT_PROJECT)) {
        return [DEFAULT_PROJECT, ...projects.filter(project => project !== DEFAULT_PROJECT)];
      }
      return projects;
    }

    function getRowsForProject(project) {
      const rawProject = normalizeText(project || selectedProject || DEFAULT_PROJECT);
      const targetProject = isAllProjectsValue(rawProject) ? '' : rawProject;
      return rawRows.filter(r => {
        const projectMatch = !targetProject || normalizeText(r.project) === targetProject;
        const plotMatch = isAllPlotsValue(selectedPlot) || normalizeText(r.plot) === normalizeText(selectedPlot);
        return projectMatch && plotMatch;
      });
    }

    function metricLabel(metric) {
      return metric === 'piles' ? 'Piles' : metric === 'lm' ? 'Lm' : 'm3';
    }

    function metricUnit(metric) {
      return metric === 'piles' ? 'piles' : metric === 'lm' ? 'Lm' : 'mÃ‚Â³';
    }

    function metricValueForRow(row, metric) {
      if (metric === 'piles') return 1;
      if (metric === 'lm') {
        const n = Number(row.asbuilt_depth);
        return Number.isFinite(n) ? n : 0;
      }
      const n = Number(row.asbuilt_concreteQty);
      return Number.isFinite(n) ? n : 0;
    }

    function startOfWeek(dateKey) {
      const d = new Date(dateKey + 'T00:00:00Z');
      const day = d.getUTCDay() || 7;
      d.setUTCDate(d.getUTCDate() - day + 1);
      return d.toISOString().slice(0, 10);
    }

    function startOfMonth(dateKey) {
      const d = new Date(dateKey + 'T00:00:00Z');
      d.setUTCDate(1);
      return d.toISOString().slice(0, 10);
    }

    function periodKey(dateKey, granularity) {
      if (granularity === 'week') return startOfWeek(dateKey);
      if (granularity === 'month') return startOfMonth(dateKey);
      return dateKey;
    }

    function incrementPeriod(dateObj, granularity) {
      const d = new Date(dateObj);
      if (granularity === 'week') d.setUTCDate(d.getUTCDate() + 7);
      else if (granularity === 'month') d.setUTCMonth(d.getUTCMonth() + 1, 1);
      else d.setUTCDate(d.getUTCDate() + 1);
      return d;
    }

    function getActiveRigCount(rows) {
      return new Set(getExecutedRows(rows).map(r => normalizeText(r.machine)).filter(Boolean)).size;
    }

    function getOverviewRigOverride(project) {
      const projectKey = normalizeText(project).toLowerCase();
      if (projectKey === 'titania') {
        return 2;
      }
      if (projectKey === 'vintage') {
        return 1;
      }
      return null;
    }

    function getPlannedRigCountForDate(rows, dateKey, project = selectedProject) {
      const projectKey = normalizeText(project).toLowerCase();
      if (projectKey === 'titania') {
        return dateKey >= '2026-04-03' ? 2 : 1;
      }
      if (projectKey === 'vintage') {
        return 1;
      }
      const activeRigs = getActiveRigCount(rows);
      return activeRigs || 1;
    }

    function getPeriodEndKey(periodStartKey, granularity) {
      const next = incrementPeriod(new Date(periodStartKey + 'T00:00:00Z'), granularity);
      next.setUTCDate(next.getUTCDate() - 1);
      const previousDay = previousDayKey();
      return next.toISOString().slice(0, 10) > previousDay ? previousDay : next.toISOString().slice(0, 10);
    }

    function buildPlannedCurve(rows, data, project = selectedProject) {
      if (!data.length) return [];
      const executedDates = getExecutedRows(rows).map(r => getOverviewDateKey(r)).filter(Boolean).sort();
      if (!executedDates.length) return [];

      const firstDate = executedDates[0];
      let cumulative = 0;

      return data.map(item => {
        const periodEnd = getPeriodEndKey(item.date, chartGranularity);
        let cursor = new Date(firstDate + 'T00:00:00Z');
        let planned = 0;
        while (cursor.toISOString().slice(0, 10) <= periodEnd) {
          const key = cursor.toISOString().slice(0, 10);
          if (isWorkingDay(key)) {
            planned += getPlannedRigCountForDate(rows, key, project) * 3;
          }
          cursor.setUTCDate(cursor.getUTCDate() + 1);
        }
        cumulative = planned;
        return { date: item.date, value: cumulative };
      });
    }

    function formatPeriodLabel(dateKey) {
      if (chartGranularity === 'day') return formatDateLabel(dateKey).replace(/\s/g, '');
      if (chartGranularity === 'week') return getCW(dateKey).replace(' ', '');
      return new Date(dateKey + 'T00:00:00Z').toLocaleDateString('en-GB', { month: 'short', year: '2-digit', timeZone: 'UTC' }).replace(/\s/g, '');
    }

    function periodTitle(dateKey) {
      if (chartGranularity === 'day') return formatDateLabel(dateKey);
      if (chartGranularity === 'week') return getCW(dateKey) + ' Ã‚Â· ' + formatDateLabel(dateKey);
      return new Date(dateKey + 'T00:00:00Z').toLocaleDateString('en-GB', { month: 'long', year: 'numeric', timeZone: 'UTC' });
    }

    function aggregateDailyMetrics(rows, metric) {
      const map = new Map();
      const executedRows = getExecutedRows(rows);
      const dates = executedRows.map(row => getOverviewDateKey(row)).filter(Boolean).sort();
      if (!dates.length) return [];

      const firstKey = periodKey(dates[0], chartGranularity);
      const rangeEndKey = overviewDateMode === 'calendar' ? todayKey() : previousDayKey();
      const lastKey = periodKey(rangeEndKey, chartGranularity);
      let cursor = new Date(firstKey + 'T00:00:00Z');
      const end = new Date(lastKey + 'T00:00:00Z');

      while (cursor <= end) {
        const key = cursor.toISOString().slice(0, 10);
        map.set(key, { date: key, value: 0, executedCount: 0, items: [] });
        cursor = incrementPeriod(cursor, chartGranularity);
      }

      executedRows.forEach(row => {
        const date = getOverviewDateKey(row);
        if (!date) return;
        const key = periodKey(date, chartGranularity);
        if (!map.has(key)) return;
        const item = map.get(key);
        item.value += metricValueForRow(row, metric);
        item.executedCount += 1;
        item.items.push({ id: normalizeText(row.id), machine: normalizeText(row.machine) || '-', date: date });
      });

      return Array.from(map.values()).sort((a, b) => a.date.localeCompare(b.date));
    }

    function lastCalendarDaysMetric(rows, metric, calendarDayCount = 7) {
      const executedRows = getExecutedRows(rows).filter(r => getOverviewDateKey(r));
      if (!executedRows.length) return 0;

      const allDates = Array.from(new Set(
        executedRows
        .map(r => getOverviewDateKey(r))
        .filter(Boolean)
      )).sort();

      const latestDate = allDates[allDates.length - 1];

      if (!latestDate) return 0;

      const endDate = new Date(`${latestDate}T00:00:00Z`);
      const startDate = new Date(endDate);
      startDate.setUTCDate(endDate.getUTCDate() - (calendarDayCount - 1));
      const startKey = startDate.toISOString().slice(0, 10);
      const selectedDates = [];
      const cursor = new Date(startDate);
      while (cursor <= endDate) {
        selectedDates.push(cursor.toISOString().slice(0, 10));
        cursor.setUTCDate(cursor.getUTCDate() + 1);
      }
      const endKey = latestDate;

      const rowsInWindow = executedRows.filter(row => {
        const dateKey = getOverviewDateKey(row);
        return dateKey && dateKey >= startKey && dateKey <= endKey;
      });

      if (metric === 'piles') {
        return rowsInWindow.length / selectedDates.length;
      } else if (metric === 'lm') {
        const total = rowsInWindow.reduce((sum, row) => sum + (Number(row.asbuilt_depth) || 0), 0);
        return total / selectedDates.length;
      } else if (metric === 'm3') {
        const total = rowsInWindow.reduce((sum, row) => sum + (Number(row.asbuilt_concreteQty) || 0), 0);
        return total / selectedDates.length;
      }

      return 0;
    }

    function addWorkingDays(dateKey, workingDays) {
      const d = new Date(`${dateKey}T00:00:00Z`);
      let added = 0;
      while (added < workingDays) {
        d.setUTCDate(d.getUTCDate() + 1);
        const key = d.toISOString().slice(0, 10);
        if (isWorkingDay(key)) added += 1;
      }
      return d.toISOString().slice(0, 10);
    }

    function monthsBetween(dateFrom, dateTo) {
      const a = new Date(`${dateFrom}T00:00:00Z`);
      const b = new Date(`${dateTo}T00:00:00Z`);
      return Math.max(0, (b - a) / 86400000) / 30.44;
    }

    function computeStats(rows, project = selectedProject) {
      const total = rows.length;
      const executedRows = getExecutedRows(rows);
      const completed = executedRows.length;
      const remaining = Math.max(0, total - completed);
      const progress = total ? (completed / total) * 100 : 0;
      const projectKey = normalizeText(project).toLowerCase();

      const executedDates = Array.from(new Set(
        executedRows
          .map(r => getOverviewDateKey(r))
          .filter(Boolean)
      )).sort();
      const workedDaysCount = executedDates.length;

      let avgPilesRaw = 0;
      let avgLmRaw = 0;
      if (workedDaysCount > 0 && workedDaysCount < 6) {
        const totalLmFromStart = executedRows.reduce((sum, row) => sum + (Number(row.asbuilt_depth) || 0), 0);
        avgPilesRaw = completed / workedDaysCount;
        avgLmRaw = totalLmFromStart / workedDaysCount;
      } else {
        avgPilesRaw = lastCalendarDaysMetric(rows, 'piles', 7) * (7 / 6);
        avgLmRaw = lastCalendarDaysMetric(rows, 'lm', 7) * (7 / 6);
      }
      const avgPiles = avgPilesRaw > 0 ? Math.round(avgPilesRaw * 10) / 10 : 0;
      const avgLm = avgLmRaw > 0 ? Math.round(avgLmRaw * 10) / 10 : 0;

      const yesterdayExecuted = aggregateDailyMetrics(rows, 'piles').find(x => x.date === previousDayKey())?.executedCount || 0;
      const latestCasting = executedRows.map(r => normalizeDateString(r.castingDate)).filter(Boolean).sort().pop() || '';
      let activeRigs = new Set(executedRows.map(r => normalizeText(r.machine)).filter(Boolean)).size;
      const fixedRigCount = getOverviewRigOverride(projectKey);
      if (fixedRigCount !== null) activeRigs = fixedRigCount;
      let etaDate = '';
      let etaMonths = null;
      if (remaining > 0 && avgPilesRaw > 0) {
        const workingDaysNeeded = Math.ceil(remaining / avgPilesRaw);
        etaDate = addWorkingDays(todayKey(), workingDaysNeeded);
        etaMonths = monthsBetween(todayKey(), etaDate);
      }
      return { total, completed, remaining, activeRigs, avgPiles, avgLm, progress, yesterdayExecuted, latestCasting, etaDate, etaMonths };
    }

    function getCW(dateKey) {
      const date = new Date(`${dateKey}T00:00:00Z`);
      const tmp = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
      const dayNum = tmp.getUTCDay() || 7;
      tmp.setUTCDate(tmp.getUTCDate() + 4 - dayNum);
      const yearStart = new Date(Date.UTC(tmp.getUTCFullYear(), 0, 1));
      const weekNo = Math.ceil((((tmp - yearStart) / 86400000) + 1) / 7);
      return `CW ${String(weekNo).padStart(2, '0')}`;
    }

    function renderExecutionMatrix(rows) {
      const executed = getExecutedRows(rows)
        .filter(row => getOverviewDateKey(row))
        .sort((a, b) =>
          getOverviewDateKey(b).localeCompare(getOverviewDateKey(a)) ||
          normalizeText(a.id).localeCompare(normalizeText(b.id), undefined, { numeric: true })
        );

      if (!executed.length) {
        els.matrixBody.innerHTML = '<tr><td colspan="4" style="padding:14px;color:var(--muted);">No executed piles available.</td></tr>';
        return;
      }

      const groups = new Map();
      executed.forEach(row => {
        const dateKey = getOverviewDateKey(row);
        const cw = getCW(dateKey);
        if (!groups.has(cw)) groups.set(cw, { rows: [], sumLm: 0, sumM3: 0, count: 0, latestDate: dateKey });
        const group = groups.get(cw);
        const lm = Number(row.asbuilt_depth) || 0;
        const m3 = Number(row.asbuilt_concreteQty) || 0;
        group.rows.push({ id: normalizeText(row.id) || '-', date: dateKey, lm, m3 });
        group.sumLm += lm;
        group.sumM3 += m3;
        group.count += 1;
        if (dateKey > group.latestDate) group.latestDate = dateKey;
      });

      const sortedGroups = Array.from(groups.entries()).sort((a, b) => {
        const aNum = parseInt(String(a[0]).replace(/\D/g, ''), 10) || 0;
        const bNum = parseInt(String(b[0]).replace(/\D/g, ''), 10) || 0;
        if (bNum !== aNum) return bNum - aNum;
        return (b[1].latestDate || '').localeCompare(a[1].latestDate || '');
      });

      let html = '';
      sortedGroups.forEach(([cw, group]) => {
        group.rows.sort((a, b) => b.date.localeCompare(a.date) || a.id.localeCompare(b.id, undefined, { numeric: true }));
        html += `<tr class="cw-row"><td>${cw}</td><td class="id-cell">${group.count}</td><td class="num">${group.sumLm.toFixed(2)}</td><td class="num">${group.sumM3.toFixed(2)}</td></tr>`;
        group.rows.forEach(item => {
          html += `<tr><td>${formatDateFullLabel(item.date)}</td><td class="id-cell">${item.id}</td><td class="num">${item.lm.toFixed(2)}</td><td class="num">${item.m3.toFixed(2)}</td></tr>`;
        });
      });

      els.matrixBody.innerHTML = html;
    }

    function renderProjectOptions(projects) {
      if (!els.projectSelector) return;
      els.projectSelector.innerHTML = projects.map(project => `<option value="${project}" ${project === selectedProject ? 'selected' : ''}>${project}</option>`).join('');
    }

    function setActivePage(page) {
      if (!currentUser) return;
      if (!canAccessPage(page)) {
        page = 'overview';
      }
      activePage = page;
      els.pageOverview.classList.toggle('active', page === 'overview');
      if (els.pageMap) els.pageMap.classList.toggle('active', page === 'map');
      els.pageProduction.classList.toggle('active', page === 'production');
      if (els.pageUtilization) els.pageUtilization.classList.toggle('active', page === 'utilization');
      if (els.pageManpower) els.pageManpower.classList.toggle('active', page === 'manpower');
      if (els.pageCompanyManpower) els.pageCompanyManpower.classList.toggle('active', page === 'companymanpower');
      els.pageTimeline.classList.toggle('active', page === 'timeline');
      if (els.pageCost) els.pageCost.classList.toggle('active', page === 'cost');
      els.kpiRow.style.display = page === 'overview' ? 'grid' : 'none';
      document.querySelector('.content')?.classList.toggle('production-mode', page === 'production');

      const showDateModeToggle = page === 'overview' || page === 'production' || page === 'utilization';
      const dateModeToggle = document.getElementById('overviewDateModeToggle');
      if (dateModeToggle) {
        dateModeToggle.style.display = showDateModeToggle ? 'inline-flex' : 'none';
      }

      els.navButtons.forEach(btn => {
        btn.classList.toggle('active', btn.dataset.page === page);
      });

      if (page === 'overview') renderDashboard(selectedProject);
      if (page === 'map') window.setTimeout(triggerMapFocusAnimation, 80);
      if (page === 'production') renderProductionPage(selectedProject, true);
      if (page === 'utilization') renderUtilizationPage(selectedProject);
      if (page === 'manpower') renderManpowerPage(selectedProject);
      if (page === 'companymanpower') renderCompanyManpowerPage(selectedProject);
      if (page === 'timeline') renderTimelinePage(selectedProject, true);
      if (page === 'cost') renderCostPage(selectedProject, true);
    }

    function aggregateAverageBy(rows, valueField, groupBy) {
      const executed = getExecutedRows(rows).filter(r => getOverviewDateKey(r));
      const map = new Map();

      executed.forEach(row => {
        const raw = parseFloat(row[valueField]);
        if (!Number.isFinite(raw)) return;

        const reportDate = getOverviewDateKey(row);
        let key = '';
        if (groupBy === 'pile') key = normalizeText(row.id) || '-';
        else if (groupBy === 'cw') key = getCW(reportDate);
        else key = reportDate;

        if (!map.has(key)) map.set(key, { key, sum: 0, count: 0, items: [] });
        const item = map.get(key);
        item.sum += raw;
        item.count += 1;
        item.items.push({
          id: normalizeText(row.id) || '-',
          value: raw,
          date: reportDate,
          cw: getCW(reportDate)
        });
      });

      return Array.from(map.values())
        .map(item => ({
          key: item.key,
          avg: item.count ? item.sum / item.count : 0,
          count: item.count,
          items: item.items.sort((a, b) => (a.date || '').localeCompare(b.date || '') || a.id.localeCompare(b.id, undefined, { numeric: true }))
        }))
        .sort((a, b) => a.key.localeCompare(b.key, undefined, { numeric: true }));
    }

    function formatNumberOneDecimal(value) {
      return Number(value || 0).toFixed(1);
    }

    function formatPercentValue(value) {
      const pct = Number(value || 0) * 100;
      return `${pct.toFixed(1)}%`;
    }

    function getProductionLabelFontSize(chartWidth) {
      if (chartWidth < 320) return 9.5;
      if (chartWidth < 460) return 10.5;
      if (chartWidth < 700) return 11.5;
      return 12.5;
    }

    function placeTooltipInScrollWrap(tooltipEl, wrapEl, anchorX, anchorY, tooltipWidth = 300, tooltipHeight = 170) {
      const offset = 12;
      const scrollLeft = wrapEl.scrollLeft || 0;
      const scrollTop = wrapEl.scrollTop || 0;
      const viewLeft = scrollLeft;
      const viewRight = scrollLeft + wrapEl.clientWidth;
      const viewTop = scrollTop;
      const viewBottom = scrollTop + wrapEl.clientHeight;

      let x = anchorX - (tooltipWidth / 2);
      let y = anchorY - tooltipHeight - offset;

      if (y < viewTop + 8) {
        y = anchorY + offset;
      }

      x = Math.max(viewLeft + 8, Math.min(x, viewRight - tooltipWidth - 8));
      y = Math.max(viewTop + 8, Math.min(y, viewBottom - tooltipHeight - 8));

      tooltipEl.style.left = `${x}px`;
      tooltipEl.style.top = `${y}px`;
    }

    function showProductionTooltip(evt, dataPoint, valueFormatter, options = {}) {
      const wrap = evt.currentTarget.closest('.prod-chart-wrap');
      if (!wrap) return;

      const title = options.title || 'Value';
      const periodText = options.groupBy === 'day'
        ? formatDateFullLabel(dataPoint.key)
        : (options.groupBy === 'cw' ? dataPoint.key : dataPoint.key);

      let html = `<div class="tooltip-title">${periodText}</div>`;

      if (options.groupBy === 'day' || options.groupBy === 'cw') {
        html += `<div class="tooltip-row"><span>AVG ${title}</span><strong>${valueFormatter(dataPoint.avg)}</strong></div>`;
      } else {
        html += `<div class="tooltip-row"><span>${title}</span><strong>${valueFormatter(dataPoint.avg)}</strong></div>`;
      }

      if (options.groupBy === 'day' && Array.isArray(dataPoint.items) && dataPoint.items.length) {
        html += `<div class="tooltip-list">`;
        dataPoint.items.forEach(item => {
          html += `<div class="tooltip-list-item"><span>${item.id || '-'}</span><span>${valueFormatter(item.value)}</span></div>`;
        });
        html += `</div>`;
      }

      els.chartTooltip.innerHTML = html;
      els.chartTooltip.classList.add('visible');
      document.body.appendChild(els.chartTooltip);
      els.chartTooltip.style.position = 'fixed';

      const tooltipRect = els.chartTooltip.getBoundingClientRect();
      const targetRect = evt.currentTarget.getBoundingClientRect();
      const gap = 10;

      let left = targetRect.left + (targetRect.width / 2) - (tooltipRect.width / 2);
      let top = targetRect.top - tooltipRect.height - gap;

      if (top < 8) {
        top = targetRect.bottom + gap;
      }

      left = Math.max(8, Math.min(left, window.innerWidth - tooltipRect.width - 8));
      top = Math.max(8, Math.min(top, window.innerHeight - tooltipRect.height - 8));

      els.chartTooltip.style.left = `${left}px`;
      els.chartTooltip.style.top = `${top}px`;
    }

    function renderSmallBarChart(svg, data, valueFormatter, options = {}) {
      const shouldAnimate = options.animate !== false;
      while (svg.firstChild) svg.removeChild(svg.firstChild);
      if (!data.length) return;

      const defs = svgEl('defs');
      const grad = svgEl('linearGradient', { id: 'prodBarGradient', x1: '0', y1: '0', x2: '0', y2: '1' });
      grad.appendChild(svgEl('stop', { offset: '0%', 'stop-color': '#8ef0bf' }));
      grad.appendChild(svgEl('stop', { offset: '55%', 'stop-color': '#6ee7b7', 'stop-opacity': '0.58' }));
      grad.appendChild(svgEl('stop', { offset: '100%', 'stop-color': '#6ee7b7', 'stop-opacity': '0.18' }));
      defs.appendChild(grad);
      svg.appendChild(defs);

      const isPileView = options.groupBy === 'pile';
      const visibleBars = isPileView ? 8 : 12;
      const wrap = svg.closest('.prod-chart-wrap');
      const wrapWidth = wrap?.clientWidth || svg.clientWidth || svg.getBoundingClientRect().width || 420;
      const wrapHeight = wrap?.clientHeight || svg.clientHeight || svg.getBoundingClientRect().height || 250;
      const width = Math.max(Math.round(wrapWidth), 320);
      const height = Math.max(Math.round(wrapHeight), 230);
      const left = 54;
      const right = 18;
      const top = 18;
      const bottom = 64;
      const innerH = Math.max(height - top - bottom, 120);
      const plotTop = top;
      const plotBottom = top + innerH;
      const values = data.map(d => Number(d.avg) || 0);
      const minVal = Math.min(0, ...values);
      const maxVal = Math.max(0, ...values);
      const range = Math.max(maxVal - minVal, 1);
      const visibleInnerW = Math.max(width - left - right, 180);
      const scrollNeeded = data.length > visibleBars;
      const stepX = scrollNeeded
        ? visibleInnerW / visibleBars
        : (data.length > 1 ? visibleInnerW / (data.length - 1) : 0);
      const innerW = scrollNeeded ? stepX * data.length : visibleInnerW;
      const svgWidth = left + innerW + right;
      const barW = isPileView
        ? Math.max(12, Math.min(18, stepX * 0.30))
        : Math.max(16, Math.min(24, stepX * 0.40));
      const chartWidth = width;
      const labelFontSize = Math.max(9, getProductionLabelFontSize(chartWidth) - 1.5);
      const minBarHeightForLabel = chartWidth < 420 ? 24 : 18;
      const yScale = value => plotTop + ((maxVal - value) / range) * innerH;
      const baselineY = minVal >= 0 ? plotBottom : (maxVal <= 0 ? plotTop : yScale(0));

      svg.setAttribute('viewBox', `0 0 ${svgWidth} ${height}`);
      svg.setAttribute('width', svgWidth);
      svg.setAttribute('height', height);
      svg.style.width = scrollNeeded ? `${svgWidth}px` : '100%';
      svg.style.minWidth = scrollNeeded ? `${svgWidth}px` : '100%';
      svg.style.height = '100%';

      for (let i = 0; i < 5; i += 1) {
        const tickValue = maxVal - ((maxVal - minVal) / 4) * i;
        const y = yScale(tickValue);
        svg.appendChild(svgEl('line', { x1: left, y1: y, x2: svgWidth - right, y2: y, stroke: 'rgba(255,255,255,0.06)', 'stroke-width': '1' }));
      }

      svg.appendChild(svgEl('line', { x1: left, y1: baselineY, x2: svgWidth - right, y2: baselineY, stroke: 'rgba(255,255,255,0.16)', 'stroke-width': '1.2' }));

      const monthRuns = [];
      let currentRun = null;
      const animatedBars = [];
      const xStart = left + (scrollNeeded ? (stepX / 2) : (barW / 2));
      const xEnd = left + visibleInnerW - (scrollNeeded ? (stepX / 2) : (barW / 2));

      data.forEach((item, idx) => {
        const x = data.length <= 1
          ? left + visibleInnerW / 2
          : (scrollNeeded
            ? left + stepX * idx + stepX / 2
            : xStart + ((xEnd - xStart) / (data.length - 1)) * idx);
        const finalYValue = yScale(item.avg);
        const finalY = Math.min(finalYValue, baselineY);
        const finalH = Math.max(2, Math.abs(finalYValue - baselineY));

        const bar = svgEl('rect', {
          x: x - barW / 2,
          y: baselineY,
          width: barW,
          height: 0,
          rx: 3,
          fill: 'url(#prodBarGradient)',
          stroke: 'rgba(255,255,255,0.12)',
          'stroke-width': '1'
        });
        svg.appendChild(bar);
        animatedBars.push({ bar, finalY, finalH, avg: item.avg, idx });

        if (finalH >= minBarHeightForLabel) {
          const labelY = item.avg >= 0 ? finalY - 8 : finalY + finalH + 18;
          const label = svgEl('text', {
            x,
            y: labelY,
            class: 'prod-label',
            style: `font-size:${labelFontSize}px; opacity:0; transition: opacity 180ms ease;`
          });
          label.textContent = valueFormatter(item.avg);
          svg.appendChild(label);
          animatedBars[animatedBars.length - 1].label = label;
        }

        const isDayLabel = options.groupBy === 'day' && /^\d{4}-\d{2}-\d{2}$/.test(item.key);
        const axisFontSize = Math.max(labelFontSize - 1, 10);

        if (isDayLabel) {
          const monthText = formatMonthShortLabel(item.key);
          const dayText = formatDayNumberLabel(item.key);

          if (!currentRun || currentRun.month !== monthText) {
            currentRun = { month: monthText, startX: x, endX: x };
            monthRuns.push(currentRun);
          } else {
            currentRun.endX = x;
          }

          const dayAxis = svgEl('text', {
            x,
            y: height - 12,
            class: 'prod-axis',
            style: `font-size:${axisFontSize}px; font-weight:800;`
          });
          dayAxis.textContent = dayText;
          svg.appendChild(dayAxis);
        } else {
          const axis = svgEl('text', {
            x,
            y: height - 14,
            class: 'prod-axis',
            style: `font-size:${axisFontSize}px; font-weight:700;`
          });
          axis.textContent = item.key;
          svg.appendChild(axis);
        }

        const hit = svgEl('rect', {
          x: x - barW / 2,
          y: finalY,
          width: barW,
          height: finalH,
          class: 'hover-target'
        });
        hit.addEventListener('mousemove', evt => showProductionTooltip(evt, item, valueFormatter, options));
        hit.addEventListener('mouseleave', hideTooltip);
        svg.appendChild(hit);
      });

      if (monthRuns.length) {
        monthRuns.forEach(run => {
          const monthAxis = svgEl('text', {
            x: (run.startX + run.endX) / 2,
            y: height - 30,
            class: 'prod-axis',
            style: 'font-size:10px; font-weight:700; letter-spacing:0.4px;'
          });
          monthAxis.textContent = run.month;
          svg.appendChild(monthAxis);
        });

        if (monthRuns.length > 1) {
          for (let i = 1; i < monthRuns.length; i += 1) {
            const prevRun = monthRuns[i - 1];
            const nextRun = monthRuns[i];
            const separatorX = (prevRun.endX + nextRun.startX) / 2;

            svg.appendChild(svgEl('line', {
              x1: separatorX,
              y1: height - 44,
              x2: separatorX,
              y2: height - 4,
              stroke: 'rgba(255,255,255,0.18)',
              'stroke-width': '1.2'
            }));
          }
        }
      }

      const animateBars = () => {
        const duration = 480;
        const stagger = 28;
        const startTime = performance.now();

        function easeOutCubic(t) {
          return 1 - Math.pow(1 - t, 3);
        }

        function frame(now) {
          let running = false;
          animatedBars.forEach(item => {
            const local = Math.min(Math.max((now - startTime - item.idx * stagger) / duration, 0), 1);
            const eased = easeOutCubic(local);
            if (local < 1) running = true;

            const h = item.finalH * eased;
            const y = item.avg >= 0 ? baselineY - h : baselineY;
            item.bar.setAttribute('y', y);
            item.bar.setAttribute('height', h);

            if (item.label) {
              item.label.style.opacity = local > 0.85 ? '1' : '0';
            }
          });
          if (running) requestAnimationFrame(frame);
        }

        requestAnimationFrame(frame);
      };

      if (shouldAnimate) {
        animateBars();
      } else {
        animatedBars.forEach(item => {
          item.bar.setAttribute('y', item.avg >= 0 ? baselineY - item.finalH : baselineY);
          item.bar.setAttribute('height', item.finalH);
          if (item.label) item.label.style.opacity = '1';
        });
      }

      if (wrap) {
        if (!wrap.dataset.interactiveBound) {
          let isDown = false;
          let startX = 0;
          let startScrollLeft = 0;
          wrap.addEventListener('click', evt => {
            if (!evt.target.closest('.hover-target')) hideTooltip();
          });
          wrap.addEventListener('pointerdown', evt => {
            if ((wrap.scrollWidth - wrap.clientWidth) <= 4) return;
            isDown = true;
            startX = evt.clientX;
            startScrollLeft = wrap.scrollLeft;
            wrap.classList.add('dragging');
          });
          window.addEventListener('pointermove', evt => {
            if (!isDown) return;
            wrap.scrollLeft = startScrollLeft - (evt.clientX - startX);
          });
          const stopDrag = () => {
            isDown = false;
            wrap.classList.remove('dragging');
          };
          window.addEventListener('pointerup', stopDrag);
          wrap.addEventListener('pointerleave', stopDrag);
          wrap.addEventListener('wheel', evt => {
            if ((wrap.scrollWidth - wrap.clientWidth) <= 4) return;
            if (Math.abs(evt.deltaY) > Math.abs(evt.deltaX)) {
              evt.preventDefault();
              wrap.scrollLeft += evt.deltaY;
            }
          }, { passive: false });
          wrap.dataset.interactiveBound = '1';
        }
        if (scrollNeeded) {
          const moveToRight = () => { wrap.scrollLeft = Math.max(wrap.scrollWidth - wrap.clientWidth, 0); };
          moveToRight();
          requestAnimationFrame(moveToRight);
          requestAnimationFrame(() => requestAnimationFrame(moveToRight));
        } else {
          wrap.scrollLeft = 0;
        }
      }
    }

function renderProductionMetricChart(project, key, forceAnimate = false) {
      const rows = getRowsForProject(project);
      const configs = {
        gross: {
          svg: els.prodSvgs.gross,
          field: 'asbuilt_durationgross',
          formatter: v => formatNumberOneDecimal(v),
          options: { title: 'Gross Duration', groupBy: prodState.gross }
        },
        drilling: {
          svg: els.prodSvgs.drilling,
          field: 'asbuilt_durationdrilling',
          formatter: v => formatNumberOneDecimal(v),
          options: { title: 'Drilling Duration', groupBy: prodState.drilling }
        },
        cage: {
          svg: els.prodSvgs.cage,
          field: 'asbuilt_durationcage',
          formatter: v => formatNumberOneDecimal(v),
          options: { title: 'Cage Duration', groupBy: prodState.cage }
        },
        concrete: {
          svg: els.prodSvgs.concrete,
          field: 'asbuilt_durationconcrete',
          formatter: v => formatNumberOneDecimal(v),
          options: { title: 'Concrete Duration', groupBy: prodState.concrete }
        },
        overconsumption: {
          svg: els.prodSvgs.overconsumption,
          field: 'asbuilt_overconsumption',
          formatter: formatPercentValue,
          options: { title: 'Over Consumption', isPercent: true, groupBy: prodState.overconsumption }
        },
        overexcavation: {
          svg: els.prodSvgs.overexcavation,
          field: 'asbuilt_overexcavation',
          formatter: v => `${formatNumberOneDecimal(v)} m`,
          options: { title: 'Over Excavation', groupBy: prodState.overexcavation }
        }
      };

      const cfg = configs[key];
      if (!cfg || !cfg.svg) return;

      renderSmallBarChart(
        cfg.svg,
        aggregateAverageBy(rows, cfg.field, cfg.options.groupBy),
        cfg.formatter,
        { ...cfg.options, animate: forceAnimate || !productionChartsInitialized }
      );
    }

    function renderProductionPage(project, forceAnimate = false) {
      ['gross', 'drilling', 'cage', 'concrete', 'overconsumption', 'overexcavation']
        .forEach(key => renderProductionMetricChart(project, key, forceAnimate));
      productionChartsInitialized = true;
    }

    function buildChartDataset(rows, metric, mode) {
      const daily = aggregateDailyMetrics(rows, metric);
      let cumulative = 0;
      return daily.map(item => {
        cumulative += item.value;
        return {
          date: item.date,
          value: mode === 'daily' ? item.value : cumulative,
          dayValue: item.value,
          executedCount: item.executedCount,
          items: item.items,
          cumulative
        };
      });
    }

    function svgEl(tag, attrs = {}) {
      const el = document.createElementNS('http://www.w3.org/2000/svg', tag);
      Object.entries(attrs).forEach(([k, v]) => el.setAttribute(k, v));
      return el;
    }

    function clearSvgGroup(node) {
      while (node.firstChild) node.removeChild(node.firstChild);
    }

    function renderChart(rows, project = selectedProject) {
      const data = buildChartDataset(rows, chartMetric, chartMode);
      const plannedData = (chartMode === 'cumulative' && chartMetric === 'piles') ? buildPlannedCurve(rows, data, project) : [];
      if (els.overviewSeriesLegend) {
        const showLegend = chartMode === 'cumulative' && chartMetric === 'piles' && plannedData.length > 0;
        els.overviewSeriesLegend.hidden = !showLegend;
      }
      const titleMetric = metricLabel(chartMetric);
      const granularityLabel = chartGranularity === 'day' ? 'Day' : (chartGranularity === 'week' ? 'Week' : 'Month');
      const useTwoRowMobileDayAxis = window.matchMedia('(max-width: 767px)').matches && chartGranularity === 'day';
      els.chartTitle.textContent = `${chartMode === 'daily' ? 'Daily' : 'Cumulative'} ${titleMetric} by ${granularityLabel}`;
      els.chartTag.textContent = chartMode === 'daily' ? 'Daily' : 'Cumulative';
      clearSvgGroup(els.chartGrid);
      clearSvgGroup(els.chartSeries);
      clearSvgGroup(els.chartLabels);
      clearSvgGroup(els.chartXAxis);

      if (!data.length) return;

      const wrap = els.chartWrap;
      const visiblePoints = 14;
      const wrapWidth = wrap?.clientWidth || 920;
      const wrapHeight = wrap?.clientHeight || 360;
      const width = Math.max(Math.round(wrapWidth), 520);
      const height = Math.max(Math.round(wrapHeight), 320);
      const left = 58;
      const right = 20;
      const top = 20;
      const bottom = useTwoRowMobileDayAxis ? 74 : 56;
      const innerH = height - top - bottom;
      const visibleInnerW = width - left - right;
      const scrollNeeded = data.length > visiblePoints;
      const stepX = scrollNeeded
        ? visibleInnerW / visiblePoints
        : visibleInnerW / Math.max(data.length, 1);
      const innerW = scrollNeeded ? stepX * data.length : visibleInnerW;
      const svgWidth = left + innerW + right;
      const maxVal = Math.max(1, ...data.map(d => d.value), ...plannedData.map(d => d.value || 0));
      const barW = Math.min(34, Math.max(16, stepX * 0.52));

      els.chartSvg.setAttribute('viewBox', `0 0 ${svgWidth} ${height}`);
      els.chartSvg.setAttribute('width', svgWidth);
      els.chartSvg.setAttribute('height', height);
      els.chartSvg.style.width = scrollNeeded ? `${svgWidth}px` : '100%';
      els.chartSvg.style.minWidth = scrollNeeded ? `${svgWidth}px` : '100%';
      els.chartSvg.style.height = '100%';

      chartHoverGuide = null;
      if (chartMode === 'cumulative') {
        chartHoverGuide = svgEl('line', {
          x1: left,
          y1: top,
          x2: left,
          y2: top + innerH,
          stroke: 'rgba(255,255,255,0.55)',
          'stroke-width': '1.5',
          'stroke-dasharray': '6 6',
          style: 'display:none; pointer-events:none;'
        });
        els.chartSeries.appendChild(chartHoverGuide);
      }

      for (let i = 0; i < 5; i += 1) {
        const y = top + (innerH / 4) * i;
        els.chartGrid.appendChild(svgEl('line', { x1: left, y1: y, x2: svgWidth - right, y2: y, stroke: 'rgba(255,255,255,0.06)', 'stroke-width': '1' }));
      }

      if (chartMode === 'daily') {
        const monthRuns = [];
        let currentRun = null;
        const animatedBars = [];
        data.forEach((d, idx) => {
          const x = left + stepX * idx + stepX / 2;
          const h = (d.value / maxVal) * (innerH - 8);
          const y = top + innerH - h;
          const bar = svgEl('rect', { x: x - barW / 2, y: top + innerH, width: barW, height: 0, rx: 8, class: 'bar-shape' });
          els.chartSeries.appendChild(bar);
          const label = svgEl('text', { x, y: y - 8, class: 'bar-label', style: 'opacity:0; transition:opacity 180ms ease;' });
          label.textContent = chartMetric === 'piles' ? d.executedCount : Number(d.value.toFixed(1)).toString();
          els.chartLabels.appendChild(label);
          animatedBars.push({ bar, label, finalY: y, finalH: h, idx });
          const hit = svgEl('rect', { x: x - barW / 2, y, width: barW, height: h, class: 'hover-target' });
          hit.dataset.anchorX = x;
          hit.dataset.anchorY = y;
          hit.addEventListener('mousemove', evt => showTooltip(evt, d, null, { anchorX: x, anchorY: y }));
          hit.addEventListener('mouseleave', hideTooltip);
          els.chartSeries.appendChild(hit);
          if (useTwoRowMobileDayAxis) {
            const monthText = formatMonthShortLabel(d.date);
            const dayText = formatDayNumberLabel(d.date);
            if (!currentRun || currentRun.month !== monthText) {
              currentRun = { month: monthText, startX: x, endX: x };
              monthRuns.push(currentRun);
            } else {
              currentRun.endX = x;
            }
            const dayAxis = svgEl('text', { x, y: height - 18, class: 'axis-label', 'text-anchor': 'middle' });
            dayAxis.textContent = dayText;
            els.chartXAxis.appendChild(dayAxis);
          } else {
            const xLabel = svgEl('text', { x, y: height - 18, class: 'axis-label', 'text-anchor': 'middle' });
            xLabel.textContent = formatPeriodLabel(d.date);
            els.chartXAxis.appendChild(xLabel);
          }
        });

        if (useTwoRowMobileDayAxis && monthRuns.length) {
          monthRuns.forEach(run => {
            const monthAxis = svgEl('text', {
              x: (run.startX + run.endX) / 2,
              y: height - 38,
              class: 'axis-label',
              'text-anchor': 'middle',
              style: 'font-size:10px; font-weight:700; letter-spacing:0.25px;'
            });
            monthAxis.textContent = run.month;
            els.chartXAxis.appendChild(monthAxis);
          });
        }

        const startTime = performance.now();
        const duration = 520;
        const stagger = 26;
        const easeOutCubic = t => 1 - Math.pow(1 - t, 3);
        const frame = now => {
          let running = false;
          animatedBars.forEach(item => {
            const local = Math.min(Math.max((now - startTime - item.idx * stagger) / duration, 0), 1);
            const eased = easeOutCubic(local);
            if (local < 1) running = true;
            const h = item.finalH * eased;
            const y = (top + innerH) - h;
            item.bar.setAttribute('y', y);
            item.bar.setAttribute('height', h);
            item.label.style.opacity = local > 0.82 ? '1' : '0';
          });
          if (running) requestAnimationFrame(frame);
        };
        requestAnimationFrame(frame);
      } else {
        const monthRuns = [];
        let currentRun = null;
        let path = '';
        let area = `M ${left} ${top + innerH}`;
        let plannedPath = '';
        data.forEach((d, idx) => {
          const x = left + stepX * idx + stepX / 2;
          const y = top + innerH - (d.value / maxVal) * (innerH - 10);
          path += `${idx === 0 ? 'M' : 'L'} ${x} ${y} `;
          area += ` L ${x} ${y}`;

          if (plannedData[idx]) {
            const plannedY = top + innerH - (plannedData[idx].value / maxVal) * (innerH - 10);
            plannedPath += `${idx === 0 ? 'M' : 'L'} ${x} ${plannedY} `;
          }

          els.chartSeries.appendChild(svgEl('circle', { cx: x, cy: y, r: 5.5, class: 'point-dot' }));
          const label = svgEl('text', { x, y: y - 10, class: 'point-label' });
          label.textContent = chartMetric === 'piles' ? Math.round(d.value).toString() : Number(d.value.toFixed(1)).toString();
          els.chartLabels.appendChild(label);
          const hit = svgEl('rect', { x: x - stepX / 2, y: top, width: stepX, height: innerH + bottom, class: 'hover-target' });
          hit.dataset.anchorX = x;
          hit.dataset.anchorY = y;
          hit.addEventListener('mousemove', evt => showTooltip(evt, d, plannedData[idx], { anchorX: x, anchorY: y, showGuide: true }));
          hit.addEventListener('mouseleave', hideTooltip);
          els.chartSeries.appendChild(hit);
          if (useTwoRowMobileDayAxis) {
            const monthText = formatMonthShortLabel(d.date);
            const dayText = formatDayNumberLabel(d.date);
            if (!currentRun || currentRun.month !== monthText) {
              currentRun = { month: monthText, startX: x, endX: x };
              monthRuns.push(currentRun);
            } else {
              currentRun.endX = x;
            }
            const dayAxis = svgEl('text', { x, y: height - 18, class: 'axis-label', 'text-anchor': 'middle' });
            dayAxis.textContent = dayText;
            els.chartXAxis.appendChild(dayAxis);
          } else {
            const xLabel = svgEl('text', { x, y: height - 18, class: 'axis-label', 'text-anchor': 'middle' });
            xLabel.textContent = formatPeriodLabel(d.date);
            els.chartXAxis.appendChild(xLabel);
          }
        });
        if (useTwoRowMobileDayAxis && monthRuns.length) {
          monthRuns.forEach(run => {
            const monthAxis = svgEl('text', {
              x: (run.startX + run.endX) / 2,
              y: height - 38,
              class: 'axis-label',
              'text-anchor': 'middle',
              style: 'font-size:10px; font-weight:700; letter-spacing:0.25px;'
            });
            monthAxis.textContent = run.month;
            els.chartXAxis.appendChild(monthAxis);
          });
        }
        const lastX = left + stepX * (data.length - 1) + stepX / 2;
        area += ` L ${lastX} ${top + innerH} Z`;

        const areaEl = svgEl('path', { d: area, class: 'area-shape', style: 'opacity:0; transition:opacity 260ms ease;' });
        els.chartSeries.appendChild(areaEl);

        let plannedEl = null;
        if (plannedPath) {
          plannedEl = svgEl('path', {
            d: plannedPath.trim(),
            fill: 'none',
            stroke: 'rgba(255,255,255,0.72)',
            'stroke-width': '2',
            'stroke-dasharray': '8 6',
            'stroke-linecap': 'round',
            'stroke-linejoin': 'round'
          });
          els.chartSeries.appendChild(plannedEl);
        }

        const lineEl = svgEl('path', { d: path.trim(), class: 'line-shape' });
        els.chartSeries.appendChild(lineEl);

        const animatePath = pathEl => {
          const length = pathEl.getTotalLength();
          pathEl.style.strokeDasharray = `${length}`;
          pathEl.style.strokeDashoffset = `${length}`;
          pathEl.style.transition = 'stroke-dashoffset 700ms cubic-bezier(.22,.61,.36,1)';
          requestAnimationFrame(() => { pathEl.style.strokeDashoffset = '0'; });
        };

        requestAnimationFrame(() => {
          areaEl.style.opacity = '1';
          animatePath(lineEl);
          if (plannedEl) animatePath(plannedEl);
          Array.from(els.chartLabels.children).forEach((label, idx) => {
            label.style.opacity = '0';
            label.style.transition = 'opacity 180ms ease';
            setTimeout(() => { label.style.opacity = '1'; }, 280 + idx * 18);
          });
          Array.from(els.chartSeries.querySelectorAll('.point-dot')).forEach((dot, idx) => {
            dot.style.opacity = '0';
            dot.style.transformOrigin = 'center';
            dot.style.transition = 'opacity 180ms ease';
            setTimeout(() => { dot.style.opacity = '1'; }, 260 + idx * 18);
          });
        });
      }

      if (wrap) {
        if (!wrap.dataset.interactiveBound) {
          let isDown = false;
          let startX = 0;
          let startScrollLeft = 0;
          wrap.addEventListener('click', evt => {
            if (!evt.target.closest('.hover-target')) hideTooltip();
          });
          wrap.addEventListener('pointerdown', evt => {
            if ((wrap.scrollWidth - wrap.clientWidth) <= 4) return;
            isDown = true;
            startX = evt.clientX;
            startScrollLeft = wrap.scrollLeft;
            wrap.classList.add('dragging');
          });

          window.addEventListener('pointermove', evt => {
            if (!isDown) return;
            wrap.scrollLeft = startScrollLeft - (evt.clientX - startX);
          });

          const stopDrag = () => {
            isDown = false;
            wrap.classList.remove('dragging');
          };

          window.addEventListener('pointerup', stopDrag);
          wrap.addEventListener('pointerleave', stopDrag);
          wrap.addEventListener('wheel', evt => {
            if ((wrap.scrollWidth - wrap.clientWidth) <= 4) return;
            if (Math.abs(evt.deltaY) > Math.abs(evt.deltaX)) {
              evt.preventDefault();
              wrap.scrollLeft += evt.deltaY;
            }
          }, { passive: false });

          wrap.dataset.interactiveBound = '1';
        }

        if (scrollNeeded) {
          wrap.scrollLeft = Math.max(wrap.scrollWidth - wrap.clientWidth, 0);
        } else {
          wrap.scrollLeft = 0;
        }
      }
    }

    function showTooltip(evt, dataPoint, plannedPoint = null, options = {}) {
      const wrap = els.chartWrap;
      let html = `<div class="tooltip-title">${periodTitle(dataPoint.date)}</div>`;
      if (chartMode === 'daily') {
        if (chartMetric === 'piles') {
          html += `<div class="tooltip-row"><span>Piles Executed</span><strong>${dataPoint.executedCount}</strong></div>`;
          if (chartGranularity === 'day' && dataPoint.items.length) {
            html += `<div class="tooltip-list">${dataPoint.items.map(item => `<div class="tooltip-list-item"><span>${item.id || '-'}</span><span>${item.machine || '-'}</span></div>`).join('')}</div>`;
          }
        } else {
          const unit = metricUnit(chartMetric);
          html += `<div class="tooltip-row"><span>${metricLabel(chartMetric)}</span><strong>${Number(dataPoint.value.toFixed(1)).toLocaleString()} ${unit}</strong></div>`;
          html += `<div class="tooltip-row"><span>Piles Executed</span><strong>${dataPoint.executedCount}</strong></div>`;
        }
      } else {
        const unit = metricUnit(chartMetric);
        html += `<div class="tooltip-row"><span>Executed This ${chartGranularity === 'day' ? 'Date' : 'Period'}</span><strong>${Number(dataPoint.dayValue.toFixed(1)).toLocaleString()} ${unit}</strong></div>`;
        html += `<div class="tooltip-row"><span>Cumulative</span><strong>${Number(dataPoint.value.toFixed(1)).toLocaleString()} ${unit}</strong></div>`;
        if (chartMetric === 'piles' && plannedPoint) {
          html += `<div class="tooltip-row"><span>Planned</span><strong>${Number(plannedPoint.value.toFixed(1)).toLocaleString()} piles</strong></div>`;
          html += `<div class="tooltip-row"><span>Variance</span><strong>${Number((dataPoint.value - plannedPoint.value).toFixed(1)).toLocaleString()} piles</strong></div>`;
        }

      }
      els.chartTooltip.innerHTML = html;
      els.chartTooltip.classList.add('visible');
      wrap.appendChild(els.chartTooltip);

      const tooltipHeight = chartMode === 'daily' && chartMetric === 'piles' ? 190 : 150;
      const anchorX = Number(options.anchorX || evt.currentTarget?.dataset?.anchorX || 0);
      const anchorY = Number(options.anchorY || evt.currentTarget?.dataset?.anchorY || 0);
      placeTooltipInScrollWrap(els.chartTooltip, wrap, anchorX, anchorY, 280, tooltipHeight);

      if (chartHoverGuide && options.showGuide) {
        chartHoverGuide.setAttribute('x1', anchorX);
        chartHoverGuide.setAttribute('x2', anchorX);
        chartHoverGuide.style.display = 'block';
      }
    }

    function hideTooltip() {
      els.chartTooltip.classList.remove('visible');
      els.chartTooltip.style.position = '';
      els.chartTooltip.style.left = '';
      els.chartTooltip.style.top = '';
      els.chartTooltip.innerHTML = '';
      if (chartHoverGuide) chartHoverGuide.style.display = 'none';
    }

    function hideChartTooltipsOnOutsideInteraction(evt) {
      const target = evt.target;
      if (!target) return;
      if (target.closest('.hover-target')) return;
      if (target.closest('.chart-tooltip')) return;
      hideTooltip();
      hideTimelineTooltip();
      if (typeof window.hideCostTtp === 'function') window.hideCostTtp();
    }

    function syncModeToggleUI() {
      const btn = document.getElementById('modeToggle');
      const isCumulative = chartMode === 'cumulative';
      btn.classList.toggle('on', isCumulative);
      btn.setAttribute('aria-pressed', isCumulative ? 'true' : 'false');
      els.modeLabelCumulative.classList.toggle('active', isCumulative);
    }

    function toggleMode() {
      chartMode = chartMode === 'daily' ? 'cumulative' : 'daily';
      syncModeToggleUI();
      renderDashboard(selectedProject);
    }

    function setMetric(metric) {
      chartMetric = metric;
      els.metricToggleButtons.forEach(btn => btn.classList.toggle('active', btn.dataset.metric === metric));
      renderDashboard(selectedProject);
    }

    function setGranularity(granularity) {
      chartGranularity = granularity;
      els.granularityToggleButtons.forEach(btn => btn.classList.toggle('active', btn.dataset.granularity === granularity));
      renderDashboard(selectedProject);
    }

    function renderDashboard(project) {
      const rows = getRowsForProject(project);
      const stats = computeStats(rows, project);
      const avgBasisLabel = '7CD';
      els.pageSubtitle.textContent = getScopeSubtitle();
      els.kpiTotal.textContent = stats.total.toLocaleString();
      els.kpiExecuted.textContent = stats.completed.toLocaleString();
      els.kpiRemaining.textContent = stats.remaining.toLocaleString();
      els.kpiRigs.textContent = stats.activeRigs.toLocaleString();
      els.kpiAvg.textContent = stats.avgPiles ? stats.avgPiles.toLocaleString() + ' pile/d' : '0 pile/d';
      if (els.kpiAvgMetaL) els.kpiAvgMetaL.textContent = stats.avgLm ? stats.avgLm.toLocaleString() + ' Lm' : '0 Lm';
      els.kpiAvgMeta.textContent = avgBasisLabel;
      els.kpiEta.textContent = stats.etaDate ? formatDateLabel(stats.etaDate) : 'Ã¢â‚¬â€';
      els.kpiTotalMetaR.textContent = 'Live';
      els.kpiCompletedMetaL.textContent = `${stats.progress.toFixed(1)}% executed`;
      els.kpiCompletedMetaR.textContent = `${stats.yesterdayExecuted} yesterday`;
      els.kpiRemainingMetaR.textContent = `${(100 - stats.progress).toFixed(1)}%`;
      els.kpiEtaDays.textContent = stats.etaMonths !== null ? `${stats.etaMonths.toFixed(1)} mo` : 'Ã¢â‚¬â€';
      renderChart(rows, project);
      renderExecutionMatrix(rows);
      if (activePage === 'production') renderProductionPage(project);
    }

    async function loadDashboardData() {
      els.dataSourceChip.textContent = 'Loading Source';
      const res = await fetch(JSON_URL, { cache: 'no-store' });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();

      const sourceRows = Array.isArray(data) ? data : (Array.isArray(data.piles) ? data.piles : []);
      rawRows = sourceRows;
      if (!rawRows.length) throw new Error('No rows found in project source');
      if (!currentUser) throw new Error('Sign in required');

      try {
        const manpowerRes = await fetch(MANPOWER_URL, { cache: 'no-store' });
        if (!manpowerRes.ok) throw new Error(`HTTP ${manpowerRes.status}`);
        const manpowerData = await manpowerRes.json();
        manpowerRows = extractManpowerList(manpowerData);
      } catch (err) {
        console.error('Unable to load manpower source:', err);
        manpowerRows = [];
      }

      try {
        const companyManpowerRes = await fetch(COMPANY_MANPOWER_URL, { cache: 'no-store' });
        if (!companyManpowerRes.ok) throw new Error(`HTTP ${companyManpowerRes.status}`);
        const companyManpowerData = await companyManpowerRes.json();
        companyManpowerRows = extractCompanyManpowerList(companyManpowerData);
      } catch (err) {
        console.error('Unable to load company manpower source:', err);
        companyManpowerRows = [];
      }

      syncProjectScopeFromData();
      broadcastAuthContext();
      updateCompanyExportProjectOptions();
      updateCompanyDesignationOptions();

      renderDashboard(selectedProject);
      if (activePage === 'utilization') renderUtilizationPage(selectedProject);
      if (activePage === 'manpower') renderManpowerPage(selectedProject);
      if (activePage === 'companymanpower') renderCompanyManpowerPage(selectedProject);
      els.dataSourceChip.textContent = 'Live Data Source';
    }

    async function refreshDashboardData() {
      if (!currentUser) {
        setAuthLocked(true);
        return;
      }
      if (els.refreshDashboardBtn) {
        els.refreshDashboardBtn.disabled = true;
        els.refreshDashboardBtn.textContent = 'Refreshing...';
      }
      try {
        if (typeof window.refreshAppToLatestBuild === 'function') {
          await window.refreshAppToLatestBuild();
          return;
        }
        await loadDashboardData();
      } catch (err) {
        console.error(err);
        els.dataSourceChip.textContent = 'Source Error';
      } finally {
        if (els.refreshDashboardBtn) {
          els.refreshDashboardBtn.disabled = false;
          els.refreshDashboardBtn.textContent = 'Refresh Dashboard';
        }
      }
    }


    function getRowDate(row, keys) {
      for (const key of keys) {
        const value = row?.[key];
        if (!value) continue;
        const d = new Date(value);
        if (!Number.isNaN(d.getTime())) return d;
      }
      return null;
    }

    function hoursBetween(start, end) {
      if (!(start instanceof Date) || !(end instanceof Date)) return null;
      const diff = (end.getTime() - start.getTime()) / 36e5;
      return Number.isFinite(diff) && diff >= 0 ? diff : null;
    }

    function formatDateTimeLabel(value) {
      if (!(value instanceof Date) || Number.isNaN(value.getTime())) return 'Ã¢â‚¬â€';
      return value.toLocaleString('en-GB', {
        day: '2-digit',
        month: 'short',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false,
        timeZone: 'UTC'
      });
    }

    function formatTimelineHeaderDate(value) {
      if (!(value instanceof Date) || Number.isNaN(value.getTime())) return 'Ã¢â‚¬â€';
      return value.toLocaleDateString('en-GB', {
        day: '2-digit',
        month: 'short',
        timeZone: 'UTC'
      });
    }

    function buildPileTimelineRows(rows) {
      const candidates = rows.map(row => {
        const drillStart = getRowDate(row, ['asbuilt_drillingStart', 'asbuilt_drillingstart', 'asbuilt_DrillingStart']);
        const drillEnd = getRowDate(row, ['asbuilt_drillingEnd', 'asbuilt_drillingend', 'asbuilt_DrillingEnd']);
        const cageStart = getRowDate(row, ['asbuilt_cageStart', 'asbuilt_cagestart', 'asbuilt_CageStart']);
        const cageEnd = getRowDate(row, ['asbuilt_cageEnd', 'asbuilt_cageend', 'asbuilt_CageEnd']);
        const pourStart = getRowDate(row, ['asbuilt_concreteStart', 'asbuilt_concretestart', 'asbuilt_ConcreteStart']);
        const pourEnd = getRowDate(row, ['asbuilt_concreteEnd', 'asbuilt_concreteend', 'asbuilt_ConcreteEnd']);
        const id = normalizeText(row.id || row.pileId || row.PileID || row.name);
        if (!id) return null;

        const activities = [
          { key: 'drilling', label: 'Drilling', color: '#4ade80', start: drillStart, end: drillEnd, duration: Number(row.asbuilt_durationdrilling) || hoursBetween(drillStart, drillEnd) || 0 },
          { key: 'cage', label: 'Cage Installation', color: '#c4e45f', start: cageStart, end: cageEnd, duration: Number(row.asbuilt_durationcage) || hoursBetween(cageStart, cageEnd) || 0 },
          { key: 'pouring', label: 'Pouring', color: '#4b92c6', start: pourStart, end: pourEnd, duration: Number(row.asbuilt_durationconcrete) || hoursBetween(pourStart, pourEnd) || 0 }
        ].filter(a => a.start && a.end && a.end >= a.start);

        if (!activities.length) return null;

        const gap1 = drillEnd && cageStart ? hoursBetween(drillEnd, cageStart) : null;
        const gap2 = cageEnd && pourStart ? hoursBetween(cageEnd, pourStart) : null;
        const gross = Number(row.asbuilt_durationgross) || (drillStart && pourEnd ? hoursBetween(drillStart, pourEnd) : null) || 0;
        const designDepth = Number(row.design_depth);
        const asbuiltDepth = Number(row.asbuilt_depth);

        return {
          id,
          row,
          length: Number.isFinite(asbuiltDepth) && asbuiltDepth > 0 ? asbuiltDepth : (Number.isFinite(designDepth) ? designDepth : 0),
          activities,
          drillStart,
          drillEnd,
          cageStart,
          cageEnd,
          pourStart,
          pourEnd,
          drilling: Number(row.asbuilt_durationdrilling) || hoursBetween(drillStart, drillEnd) || 0,
          cage: Number(row.asbuilt_durationcage) || hoursBetween(cageStart, cageEnd) || 0,
          pouring: Number(row.asbuilt_durationconcrete) || hoursBetween(pourStart, pourEnd) || 0,
          gap1,
          gap2,
          gross,
          minStart: new Date(Math.min(...activities.map(a => a.start.getTime()))),
          maxEnd: new Date(Math.max(...activities.map(a => a.end.getTime())))
        };
      }).filter(Boolean);

      return candidates.sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }));
    }

    function timelineRowsForProject(project) {
      return buildPileTimelineRows(getRowsForProject(project));
    }

    function getTimelineRowsForPileFilter(project) {
      const rows = timelineRowsForProject(project);
      const start = timelineState.start ? new Date(`${timelineState.start}T00:00:00Z`) : null;
      const end = timelineState.end ? new Date(`${timelineState.end}T23:59:59Z`) : null;
      return rows.filter(item => {
        if (start && item.maxEnd < start) return false;
        if (end && item.minStart > end) return false;
        return true;
      });
    }


    function getTimelineFilteredRows(project) {
      const rows = timelineRowsForProject(project);
      const pile = timelineState.pile;
      const start = timelineState.start ? new Date(`${timelineState.start}T00:00:00Z`) : null;
      const end = timelineState.end ? new Date(`${timelineState.end}T23:59:59Z`) : null;

      return rows.filter(item => {
        if (pile !== 'all' && item.id !== pile) return false;
        if (start && item.maxEnd < start) return false;
        if (end && item.minStart > end) return false;
        return true;
      });
    }

    function updateTimelinePileList(project) {
      if (!els.timelinePileList) return;
      const rows = getTimelineRowsForPileFilter(project);
      const pileIds = rows.map(r => r.id);

      if (timelineState.pile !== 'all' && !pileIds.includes(timelineState.pile)) {
        timelineState.pile = 'all';
      }

      els.timelinePileList.innerHTML = '';
      const makeBtn = (label, value) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'timeline-chip' + (timelineState.pile === value ? ' active' : '');
        btn.textContent = label;
        btn.addEventListener('click', () => {
          timelineState.pile = value;
          updateTimelinePileList(project);
          renderTimelinePage(project);
        });
        return btn;
      };

      els.timelinePileList.appendChild(makeBtn('All', 'all'));
      pileIds.forEach(id => els.timelinePileList.appendChild(makeBtn(id, id)));
    }

    function renderTimelineTable(rows) {
      if (!els.timelineTableBody) return;
      if (!rows.length) {
        els.timelineTableBody.innerHTML = `<tr><td colspan="8" style="text-align:center; color:rgba(244,247,251,0.62);">No data</td></tr>`;
        return;
      }
      els.timelineTableBody.innerHTML = rows.map(item => `
        <tr>
          <td>${item.id}</td>
          <td>${formatNumberOneDecimal(item.length)}</td>
          <td>${formatNumberOneDecimal(item.drilling)} hr</td>
          <td>${formatNumberOneDecimal(item.cage)} hr</td>
          <td>${formatNumberOneDecimal(item.pouring)} hr</td>
          <td>${Number.isFinite(item.gap1) ? `${formatNumberOneDecimal(item.gap1)} hr` : 'Ã¢â‚¬â€'}</td>
          <td>${Number.isFinite(item.gap2) ? `${formatNumberOneDecimal(item.gap2)} hr` : 'Ã¢â‚¬â€'}</td>
          <td>${formatNumberOneDecimal(item.gross)} hr</td>
        </tr>
      `).join('');
    }

    function renderTimelineSummary(rows, forceAnimate = false) {
      const count = rows.length;
      const avg = key => {
        const vals = rows.map(r => r[key]).filter(v => Number.isFinite(v) && v >= 0);
        return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
      };

      const label = timelineState.start || timelineState.end
        ? `${timelineState.start || 'Ã¢â‚¬Â¦'} Ã¢â€ â€™ ${timelineState.end || 'Ã¢â‚¬Â¦'}`
        : 'Total / All Dates';

      els.timelineSummarySubtitle.textContent = `Current scope: ${label}`;
      els.timelinePieSubtitle.textContent = `Average activities and gaps for ${timelineState.pile === 'all' ? 'all filtered piles' : timelineState.pile}`;
      els.timelinePileTag.textContent = timelineState.pile === 'all' ? 'All Piles' : timelineState.pile;
      els.timelineTag.textContent = timelineState.pile === 'all' ? `${count} piles` : timelineState.pile;
      if (els.timelineScopeCount) {
        els.timelineScopeCount.textContent = `${count} pile${count === 1 ? '' : 's'}`;
      }

      renderTimelineTable(rows);

      const gapColor = '#f1a9a9';
      const items = [
        { label: 'Drilling', value: avg('drilling'), color: '#4ade80' },
        { label: 'Gap 1', value: avg('gap1'), color: gapColor },
        { label: 'Cage', value: avg('cage'), color: '#c4e45f' },
        { label: 'Gap 2', value: avg('gap2'), color: gapColor },
        { label: 'Pouring', value: avg('pouring'), color: '#4b92c6' }
      ].filter(item => Number.isFinite(item.value) && item.value > 0);

      renderTimelinePie(items, forceAnimate);
    }

    function renderTimelinePie(items, forceAnimate = false) {
      const svg = els.timelinePieSvg;
      while (svg.firstChild) svg.removeChild(svg.firstChild);

      const cx = 180;
      const cy = 160;
      const outerR = 108;
      const innerR = 56;
      const total = items.reduce((sum, item) => sum + item.value, 0);
      els.timelinePieTotal.textContent = formatNumberOneDecimal(total);

      if (!items.length || total <= 0) {
        const t = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        t.setAttribute('x', cx);
        t.setAttribute('y', cy);
        t.setAttribute('text-anchor', 'middle');
        t.setAttribute('class', 'timeline-axis-text');
        t.textContent = 'No data';
        svg.appendChild(t);
        return;
      }

      let angle = -Math.PI / 2;
      items.forEach((item, itemIdx) => {
        const slice = (item.value / total) * Math.PI * 2;
        const next = angle + slice;
        const x1 = cx + Math.cos(angle) * outerR;
        const y1 = cy + Math.sin(angle) * outerR;
        const x2 = cx + Math.cos(next) * outerR;
        const y2 = cy + Math.sin(next) * outerR;
        const ix1 = cx + Math.cos(next) * innerR;
        const iy1 = cy + Math.sin(next) * innerR;
        const ix2 = cx + Math.cos(angle) * innerR;
        const iy2 = cy + Math.sin(angle) * innerR;
        const largeArc = slice > Math.PI ? 1 : 0;

        const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        path.setAttribute('d', `M ${x1} ${y1} A ${outerR} ${outerR} 0 ${largeArc} 1 ${x2} ${y2} L ${ix1} ${iy1} A ${innerR} ${innerR} 0 ${largeArc} 0 ${ix2} ${iy2} Z`);
        path.setAttribute('fill', item.color);
        path.setAttribute('opacity', forceAnimate ? '0' : '0.95');
        if (forceAnimate) {
            path.style.transformOrigin = `${cx}px ${cy}px`;
            path.style.transform = 'scale(0.7)';
            path.style.transition = 'all 340ms cubic-bezier(.22,.61,.36,1)';
            setTimeout(() => {
                path.style.opacity = '0.95';
                path.style.transform = 'scale(1)';
            }, 80 + itemIdx * 50);
        }
        svg.appendChild(path);

        const mid = angle + slice / 2;
        const lx1 = cx + Math.cos(mid) * (outerR + 6);
        const ly1 = cy + Math.sin(mid) * (outerR + 6);
        const lx2 = cx + Math.cos(mid) * (outerR + 22);
        const ly2 = cy + Math.sin(mid) * (outerR + 22);
        const lx3 = lx2 + (Math.cos(mid) >= 0 ? 12 : -12);

        const line = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        line.setAttribute('d', `M ${lx1} ${ly1} L ${lx2} ${ly2} L ${lx3} ${ly2}`);
        line.setAttribute('fill', 'none');
        line.setAttribute('stroke', 'rgba(255,255,255,0.36)');
        line.setAttribute('stroke-width', '1.2');
        line.style.opacity = forceAnimate ? '0' : '1';
        if (forceAnimate) {
            line.style.transition = 'opacity 240ms ease';
            setTimeout(() => { line.style.opacity = '1'; }, 240 + itemIdx * 50);
        }
        svg.appendChild(line);

        const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        label.setAttribute('x', lx3 + (Math.cos(mid) >= 0 ? 4 : -4));
        label.setAttribute('y', ly2 + 4);
        label.setAttribute('text-anchor', Math.cos(mid) >= 0 ? 'start' : 'end');
        label.setAttribute('class', 'timeline-axis-text');
        label.style.opacity = forceAnimate ? '0' : '1';
        if (forceAnimate) {
            label.style.transition = 'opacity 240ms ease';
            setTimeout(() => { label.style.opacity = '1'; }, 270 + itemIdx * 50);
        }

        const t1 = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
        t1.textContent = `${item.label} `;
        t1.setAttribute('fill', 'rgba(244,247,251,0.92)');

        const t2 = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
        t2.textContent = `${formatNumberOneDecimal(item.value)} hr (${formatNumberOneDecimal((item.value / total) * 100)}%)`;
        t2.setAttribute('fill', 'rgba(244,247,251,0.62)');

        label.appendChild(t1);
        label.appendChild(t2);
        svg.appendChild(label);

        angle = next;
      });
    }

    function showTimelineTooltip(evt, html) {
      if (!timelineTooltipEl) {
        timelineTooltipEl = document.createElement('div');
        timelineTooltipEl.className = 'timeline-tooltip';
        document.body.appendChild(timelineTooltipEl);
      }
      timelineTooltipEl.innerHTML = html;
      timelineTooltipEl.classList.add('visible');
      const rect = timelineTooltipEl.getBoundingClientRect();
      const targetRect = evt.currentTarget.getBoundingClientRect();
      let left = targetRect.left + targetRect.width / 2 - rect.width / 2;
      let top = targetRect.top - rect.height - 10;
      if (top < 8) top = targetRect.bottom + 10;
      left = Math.max(8, Math.min(left, window.innerWidth - rect.width - 8));
      top = Math.max(8, Math.min(top, window.innerHeight - rect.height - 8));
      timelineTooltipEl.style.left = `${left}px`;
      timelineTooltipEl.style.top = `${top}px`;
    }

    function hideTimelineTooltip() {
      if (timelineTooltipEl) timelineTooltipEl.classList.remove('visible');
    }

    function renderTimelinePage(project, forceAnimate = false) {
      if (!els.pageTimeline) return;
      updateTimelinePileList(project);
      const rows = getTimelineFilteredRows(project);
      renderTimelineSummary(rows, forceAnimate);

      const wrap = els.timelineChartWrap;
      const svg = els.timelineSvg;
      while (svg.firstChild) svg.removeChild(svg.firstChild);
      wrap.scrollTop = 0;

      if (!rows.length) {
        els.timelineEmpty.style.display = 'grid';
        els.timelineSubtitle.textContent = 'No pile activities found for the current filter';
        return;
      }

      els.timelineEmpty.style.display = 'none';
      els.timelineSubtitle.textContent = timelineState.pile === 'all'
        ? 'Pile activities timeline filtered by selected period and pile'
        : `${timelineState.pile} activity timeline`;

      const minStart = new Date(Math.min(...rows.flatMap(r => r.activities.map(a => a.start.getTime()))));
      const maxEnd = new Date(Math.max(...rows.flatMap(r => r.activities.map(a => a.end.getTime()))));
      const axisStart = timelineState.start ? new Date(`${timelineState.start}T00:00:00Z`) : minStart;
      const axisEnd = timelineState.end ? new Date(`${timelineState.end}T23:59:59Z`) : maxEnd;
      const totalHours = Math.max((axisEnd.getTime() - axisStart.getTime()) / 36e5, 24);

      const stickyPaneW = 128;
      const plotStartX = stickyPaneW + 10;
      const right = 24;
      const top = 44;
      const bottom = 28;
      const laneGap = 18;
      const laneHeight = 56;
      const lanes = [
        { key: 'drilling', label: 'Drilling', color: '#4ade80' },
        { key: 'cage', label: 'Cage Installation', color: '#c4e45f' },
        { key: 'pouring', label: 'Pouring', color: '#4b92c6' }
      ];
      const innerH = lanes.length * laneHeight + (lanes.length - 1) * laneGap;

      let pxPerHour = 4.6;
      if (timelineState.pile === 'all' && !timelineState.start && !timelineState.end) pxPerHour = 5.8;
      if (timelineState.pile !== 'all') pxPerHour = 9.5;
      if (totalHours <= 24) pxPerHour = Math.max(pxPerHour, 26);
      else if (totalHours <= 72) pxPerHour = Math.max(pxPerHour, 14);
      else if (totalHours <= 168) pxPerHour = Math.max(pxPerHour, 8.5);
      pxPerHour *= timelineState.zoom || 1;

      const minVisibleWidth = Math.max((wrap?.clientWidth || 900) - plotStartX - right, 520);
      const innerW = Math.max(minVisibleWidth, totalHours * pxPerHour);
      const width = plotStartX + innerW + right;
      const height = top + innerH + bottom;

      svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
      svg.setAttribute('width', width);
      svg.setAttribute('height', height);
      svg.style.width = `${width}px`;
      svg.style.minWidth = `${width}px`;
      svg.style.height = `${height}px`;
      wrap.style.height = `${height + 6}px`;
      wrap.style.overflowY = 'hidden';

      lanes.forEach((lane, idx) => {
        const y = top + idx * (laneHeight + laneGap);
        const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        line.setAttribute('x1', plotStartX);
        line.setAttribute('y1', y + laneHeight / 2);
        line.setAttribute('x2', width - right);
        line.setAttribute('y2', y + laneHeight / 2);
        line.setAttribute('stroke', 'rgba(255,255,255,0.08)');
        line.setAttribute('stroke-width', '1');
        svg.appendChild(line);
      });

      const dayCount = Math.ceil(totalHours / 24);
      for (let i = 0; i <= dayCount; i += 1) {
        const dt = new Date(axisStart.getTime() + i * 24 * 36e5);
        const x = plotStartX + 8 + ((dt.getTime() - axisStart.getTime()) / 36e5) * pxPerHour;

        const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        line.setAttribute('x1', x);
        line.setAttribute('y1', top - 6);
        line.setAttribute('x2', x);
        line.setAttribute('y2', height - bottom + 4);
        line.setAttribute('stroke', 'rgba(255,255,255,0.06)');
        line.setAttribute('stroke-width', '1');
        svg.appendChild(line);

        if (i < dayCount) {
          const txt = document.createElementNS('http://www.w3.org/2000/svg', 'text');
          txt.setAttribute('x', x + 6);
          txt.setAttribute('y', top - 14);
          txt.setAttribute('class', 'timeline-axis-text');
          txt.setAttribute('data-timeline-date', '1');
          txt.textContent = formatTimelineHeaderDate(dt);
          svg.appendChild(txt);
        }
      }

      const laneIndex = { drilling: 0, cage: 1, pouring: 2 };
      const animatedTimelineElements = [];
      rows.forEach(pile => {
        pile.activities.forEach(activity => {
          const idx = laneIndex[activity.key];
          const laneY = top + idx * (laneHeight + laneGap);
          const x = plotStartX + 8 + ((activity.start.getTime() - axisStart.getTime()) / 36e5) * pxPerHour;
          const w = Math.max(6, ((activity.end.getTime() - activity.start.getTime()) / 36e5) * pxPerHour);

          const bar = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
          bar.setAttribute('x', x);
          bar.setAttribute('y', laneY + 8);
          bar.setAttribute('width', forceAnimate ? '0' : String(w));
          bar.setAttribute('height', laneHeight - 16);
          bar.setAttribute('rx', '4');
          bar.setAttribute('fill', activity.color);
          bar.setAttribute('opacity', '0.94');
          bar.setAttribute('stroke', 'rgba(255,255,255,0.12)');
          bar.setAttribute('stroke-width', '1');
          svg.appendChild(bar);

          const duration = hoursBetween(activity.start, activity.end) || activity.duration || 0;
          const hit = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
          hit.setAttribute('x', x);
          hit.setAttribute('y', laneY + 6);
          hit.setAttribute('width', Math.max(10, w));
          hit.setAttribute('height', laneHeight - 12);
          hit.setAttribute('fill', 'transparent');
          hit.style.cursor = 'pointer';
          const html = `<div class="tooltip-title">${pile.id} Ã‚Â· ${activity.label}</div><div class="tooltip-row"><span>Start</span><strong>${formatDateTimeLabel(activity.start)}</strong></div><div class="tooltip-row"><span>End</span><strong>${formatDateTimeLabel(activity.end)}</strong></div><div class="tooltip-row"><span>Duration</span><strong>${formatNumberOneDecimal(duration)} hr</strong></div>`;
          hit.addEventListener('mousemove', evt => showTimelineTooltip(evt, html));
          hit.addEventListener('mouseleave', hideTimelineTooltip);
          svg.appendChild(hit);

          let labelObj = null;
          if ((timelineState.pile !== 'all' || w > 56) && pile.id.length <= 8) {
            labelObj = document.createElementNS('http://www.w3.org/2000/svg', 'text');
            labelObj.setAttribute('x', x + w / 2);
            labelObj.setAttribute('y', laneY + laneHeight / 2 + 4);
            labelObj.setAttribute('class', 'timeline-bar-label');
            labelObj.style.opacity = forceAnimate ? '0' : '1';
            labelObj.textContent = pile.id;
            svg.appendChild(labelObj);
          }
          if (forceAnimate) {
            animatedTimelineElements.push({ bar, finalW: w, label: labelObj });
          }
        });
      });

      if (forceAnimate && animatedTimelineElements.length) {
        const startTime = performance.now();
        const duration = 580;
        const stagger = 12;
        const easeOutCubic = t => 1 - Math.pow(1 - t, 3);
        const frame = now => {
          let running = false;
          animatedTimelineElements.forEach((item, idx) => {
            const local = Math.min(Math.max((now - startTime - idx * stagger) / duration, 0), 1);
            const eased = easeOutCubic(local);
            if (local < 1) running = true;
            item.bar.setAttribute('width', String(item.finalW * eased));
            if (item.label) {
              item.label.style.opacity = local > 0.85 ? '1' : '0';
            }
          });
          if (running) requestAnimationFrame(frame);
        };
        requestAnimationFrame(frame);
      }

      // Foreground sticky pane to keep Y-axis labels isolated from scrolled chart content
      const stickyBg = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      stickyBg.setAttribute('x', '0');
      stickyBg.setAttribute('y', '0');
      stickyBg.setAttribute('width', String(stickyPaneW));
      stickyBg.setAttribute('height', String(height));
      stickyBg.setAttribute('fill', 'rgba(6,8,11,0.985)');
      stickyBg.setAttribute('stroke', 'rgba(255,255,255,0.06)');
      stickyBg.setAttribute('stroke-width', '1');
      stickyBg.setAttribute('data-sticky-lane-bg', '1');
      svg.appendChild(stickyBg);

      lanes.forEach((lane, idx) => {
        const y = top + idx * (laneHeight + laneGap);
        const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        label.setAttribute('x', String(stickyPaneW - 10));
        label.setAttribute('y', String(y + laneHeight / 2 + 4));
        label.setAttribute('text-anchor', 'end');
        label.setAttribute('class', 'timeline-lane-label');
        label.setAttribute('data-base-x', String(stickyPaneW - 10));
        label.setAttribute('data-sticky-lane-label', '1');
        label.textContent = lane.label;
        svg.appendChild(label);
      });

      const stickyDivider = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      stickyDivider.setAttribute('x1', String(stickyPaneW));
      stickyDivider.setAttribute('y1', '0');
      stickyDivider.setAttribute('x2', String(stickyPaneW));
      stickyDivider.setAttribute('y2', String(height));
      stickyDivider.setAttribute('stroke', 'rgba(255,255,255,0.08)');
      stickyDivider.setAttribute('stroke-width', '1');
      stickyDivider.setAttribute('data-sticky-divider', '1');
      svg.appendChild(stickyDivider);

      const syncTimelineStickyLabels = () => {
        const scrollLeft = wrap.scrollLeft || 0;
        const stickyLabels = svg.querySelectorAll('[data-sticky-lane-label="1"]');
        stickyLabels.forEach(label => {
          const baseX = Number(label.getAttribute('data-base-x') || (stickyPaneW - 18));
          label.setAttribute('x', String(baseX + scrollLeft));
        });
        const stickyBgEl = svg.querySelector('[data-sticky-lane-bg="1"]');
        if (stickyBgEl) stickyBgEl.setAttribute('x', String(scrollLeft));
        const stickyDividerEl = svg.querySelector('[data-sticky-divider="1"]');
        if (stickyDividerEl) {
          stickyDividerEl.setAttribute('x1', String(stickyPaneW + scrollLeft));
          stickyDividerEl.setAttribute('x2', String(stickyPaneW + scrollLeft));
        }
      };

      const moveToRight = () => {
        wrap.scrollLeft = Math.max(wrap.scrollWidth - wrap.clientWidth, 0);
        syncTimelineStickyLabels();
      };
      if (wrap.scrollWidth > wrap.clientWidth) {
        moveToRight();
        requestAnimationFrame(moveToRight);
      } else {
        wrap.scrollLeft = 0;
        syncTimelineStickyLabels();
      }

      if (els.timelineZoomResetBtn) {
        els.timelineZoomResetBtn.textContent = `${Math.round((timelineState.zoom || 1) * 100)}%`;
      }

      if (!wrap.dataset.timelineInteractiveBound) {
        let isDown = false;
        let startX = 0;
        let startScrollLeft = 0;
        wrap.addEventListener('pointerdown', evt => {
          if ((wrap.scrollWidth - wrap.clientWidth) <= 4) return;
          isDown = true;
          startX = evt.clientX;
          startScrollLeft = wrap.scrollLeft;
          wrap.classList.add('dragging');
        });
        window.addEventListener('pointermove', evt => {
          if (!isDown) return;
          wrap.scrollLeft = startScrollLeft - (evt.clientX - startX);
          syncTimelineStickyLabels();
        });
        const stopDrag = () => {
          isDown = false;
          wrap.classList.remove('dragging');
        };
        window.addEventListener('pointerup', stopDrag);
        wrap.addEventListener('pointerleave', stopDrag);
        wrap.addEventListener('wheel', evt => {
          if (evt.ctrlKey || evt.metaKey) {
            evt.preventDefault();
            timelineState.zoom = Math.max(0.5, Math.min(4, (timelineState.zoom || 1) * (evt.deltaY < 0 ? 1.1 : 0.9)));
            renderTimelinePage(project);
            return;
          }
          if ((wrap.scrollWidth - wrap.clientWidth) <= 4) return;
          if (Math.abs(evt.deltaY) > Math.abs(evt.deltaX)) {
            evt.preventDefault();
            wrap.scrollLeft += evt.deltaY;
            syncTimelineStickyLabels();
          }
        }, { passive: false });
        wrap.addEventListener('scroll', syncTimelineStickyLabels, { passive: true });
        wrap.dataset.timelineInteractiveBound = '1';
      }
      syncTimelineStickyLabels();
    }




    function getTodayIsoLocal() {
      const d = new Date();
      const local = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      return local.toISOString().slice(0, 10);
    }

    function syncTimelinePresetButtons() {
      const buttons = [
        [els.timelinePreset7Btn, 7],
        [els.timelinePreset14Btn, 14],
        [els.timelinePreset30Btn, 30]
      ];
      let activeDays = null;
      if (timelineState.start && timelineState.end && timelineState.end === getTodayIsoLocal()) {
        const start = new Date(`${timelineState.start}T00:00:00`);
        const end = new Date(`${timelineState.end}T00:00:00`);
        const diffDays = Math.round((end - start) / 86400000) + 1;
        if ([7, 14, 30].includes(diffDays)) activeDays = diffDays;
      }
      buttons.forEach(([btn, days]) => {
        if (!btn) return;
        btn.classList.toggle('primary', activeDays === days);
      });
    }

    function applyTimelinePreset(days) {
      const end = new Date();
      end.setHours(0, 0, 0, 0);
      const start = new Date(end);
      start.setDate(start.getDate() - (days - 1));

      const endStr = end.toISOString().slice(0, 10);
      const startStr = start.toISOString().slice(0, 10);

      timelineState.start = startStr;
      timelineState.end = endStr;

      if (els.timelineStartDate) els.timelineStartDate.value = startStr;
      if (els.timelineEndDate) els.timelineEndDate.value = endStr;

      updateTimelinePileList(selectedProject);
      syncTimelinePresetButtons();
      renderTimelinePage(selectedProject);
    }

    function renderCostPage(project, forceAnimate = false) {
      if (!els.pageCost) return;
      const rows = getRowsForProject(project);
      const isTitaniaCostTemplateProject = normalizeText(project).toLowerCase() === 'titania';

      const executed = getExecutedRows(rows).filter(r => getOverviewDateKey(r));
      const cwLmMap = new Map();
      executed.forEach(r => {
        const date = getOverviewDateKey(r);
        let cwNum = 0;
        try {
          cwNum = parseInt(String(getCW(date)).replace(/\D/g, ''), 10);
        } catch (e) {}
        if (!Number.isFinite(cwNum) || cwNum <= 0) return;
        const asbuilt = Number(r.asbuilt_depth) || Number(r.design_depth) || 0;
        cwLmMap.set(cwNum, (cwLmMap.get(cwNum) || 0) + asbuilt);
      });

      const baseWeeks = isTitaniaCostTemplateProject ? [10, 11, 12, 13] : [];
      const data = isTitaniaCostTemplateProject
        ? {
            10: { pm: 1, se: 2, foreman: 2, op: 3, vb: 1, rig: 5, hl: 7, we: 1, me: 1, rigs: 1 },
            11: { pm: 1, se: 2, foreman: 3, op: 6, vb: 2, rig: 7, hl: 7, we: 2, me: 1, rigs: 2 },
            12: { pm: 1, se: 2, foreman: 3, op: 6, vb: 2, rig: 9, hl: 11, we: 1, me: 1, rigs: 2 },
            13: { pm: 1, se: 3, foreman: 3, op: 6, vb: 2, rig: 9, hl: 11, we: 2, me: 1, rigs: 2 }
          }
        : {};

      function getManpowerValue(row, keys) {
        for (const key of keys) {
          const raw = row?.[key];
          if (raw === undefined || raw === null || raw === '') continue;
          const n = Number(raw);
          if (Number.isFinite(n)) return n;
        }
        return 0;
      }

      const manpowerWeekly = new Map();
      const filteredManpower = manpowerRows.filter(row => {
        const projectMatch = normalizeText(row?.project) === normalizeText(selectedProject);
        const plotMatch = isAllPlotsValue(selectedPlot) || normalizeText(row?.plot) === normalizeText(selectedPlot);
        return projectMatch && plotMatch;
      });

      const latestDailyManpower = new Map();
      filteredManpower.forEach(row => {
        const dateKey = normalizeDateString(row?.date);
        if (!dateKey) return;
        latestDailyManpower.set(dateKey, row);
      });

      const manpowerDates = Array.from(latestDailyManpower.keys()).sort();
      const targetLastDay = previousDayKey();
      let lastDayDateKey = '';
      for (let i = manpowerDates.length - 1; i >= 0; i -= 1) {
        if (manpowerDates[i] <= targetLastDay) {
          lastDayDateKey = manpowerDates[i];
          break;
        }
      }
      if (!lastDayDateKey && manpowerDates.length) {
        lastDayDateKey = manpowerDates[manpowerDates.length - 1];
      }

      const lastDayManpower = lastDayDateKey ? latestDailyManpower.get(lastDayDateKey) : null;
      const lastDayCounts = lastDayManpower ? {
        pm: getManpowerValue(lastDayManpower, ['projectmanager', 'project manager']),
        se: getManpowerValue(lastDayManpower, ['siteengineer', 'site engineer', 'siteenginner']),
        foreman: getManpowerValue(lastDayManpower, ['foreman']),
        op: getManpowerValue(lastDayManpower, ['operator', 'operators']),
        vb: getManpowerValue(lastDayManpower, ['vibro operator', 'vibrooperator', 'vibro_operator']),
        rig: getManpowerValue(lastDayManpower, ['riggers', 'rigger']),
        we: getManpowerValue(lastDayManpower, ['welder', 'welders']),
        me: getManpowerValue(lastDayManpower, ['mechanic', 'mechanics']),
        hl: getManpowerValue(lastDayManpower, ['helpers', 'helper'])
      } : { pm: 0, se: 0, foreman: 0, op: 0, vb: 0, rig: 0, we: 0, me: 0, hl: 0 };

      const minDynamicWeek = isTitaniaCostTemplateProject ? 14 : 1;
      Array.from(latestDailyManpower.values()).forEach(row => {
        const dateKey = normalizeDateString(row?.date);
        if (!dateKey) return;
        let cwNum = 0;
        try {
          cwNum = parseInt(String(getCW(dateKey)).replace(/\D/g, ''), 10);
        } catch (e) {}
        if (!Number.isFinite(cwNum) || cwNum < minDynamicWeek) return;

        const bucket = manpowerWeekly.get(cwNum) || { dailyRows: [] };
        bucket.dailyRows.push({
          date: dateKey,
          pm: getManpowerValue(row, ['projectmanager', 'project manager']),
          se: getManpowerValue(row, ['siteengineer', 'site engineer', 'siteenginner']),
          foreman: getManpowerValue(row, ['foreman']),
          op: getManpowerValue(row, ['operator', 'operators']),
          vb: getManpowerValue(row, ['vibro operator', 'vibrooperator', 'vibro_operator']),
          rig: getManpowerValue(row, ['riggers', 'rigger']),
          we: getManpowerValue(row, ['welder', 'welders']),
          me: getManpowerValue(row, ['mechanic', 'mechanics']),
          hl: getManpowerValue(row, ['helpers', 'helper'])
        });
        manpowerWeekly.set(cwNum, bucket);
      });

      const dynamicWeeks = Array.from(manpowerWeekly.keys()).sort((a, b) => a - b);
      const weekLastDayLabelMap = new Map();
      dynamicWeeks.forEach(week => {
        const bucket = manpowerWeekly.get(week);
        const dailyRows = Array.isArray(bucket?.dailyRows) ? bucket.dailyRows.sort((a, b) => a.date.localeCompare(b.date)) : [];
        const lastDaily = dailyRows[dailyRows.length - 1] || {};
        const days = Math.max(1, Math.min(6, dailyRows.length || 0));
        weekLastDayLabelMap.set(week, lastDaily.date ? `Last day (${formatShortDateLabel(lastDaily.date)})` : '');
        data[week] = {
          source: 'manpower',
          dailyRows,
          pm: Number(lastDaily.pm) || 0,
          se: Number(lastDaily.se) || 0,
          foreman: Number(lastDaily.foreman) || 0,
          op: Number(lastDaily.op) || 0,
          vb: Number(lastDaily.vb) || 0,
          rig: Number(lastDaily.rig) || 0,
          we: Number(lastDaily.we) || 0,
          me: Number(lastDaily.me) || 0,
          hl: Number(lastDaily.hl) || 0,
          rigs: 2,
          daysFromManpower: days
        };
      });

      const weeks = [...baseWeeks, ...dynamicWeeks];

      if (els.costTableHeadRow) {
        const weekHeaders = weeks.map(week => {
          const lastDayLabel = weekLastDayLabelMap.get(week);
          if (lastDayLabel) {
            return `<th class="num"><div>Week ${week}</div><div style="font-size:10px; color:rgba(244,247,251,0.6); margin-top:2px;">${lastDayLabel}</div></th>`;
          }
          return `<th class="num">Week ${week}</th>`;
        }).join('');
        const lastDayHeaderText = lastDayDateKey ? `Last day (${formatShortDateLabel(lastDayDateKey)})` : 'Last day';
        els.costTableHeadRow.innerHTML = `<th class="col-head">CATEGORY / WEEK</th>${weekHeaders}<th class="num total-col">${lastDayHeaderText}</th>`;
      }

      const dailyRates = {
        pm: 461.5,
        se: 461.5,
        foreman: 343.4,
        op: 247.8,
        vb: 101,
        rig: 108.2,
        we: 151.4,
        me: 315.5,
        hl: 101
      };

      const overheadsDaily = 2540; // Updated overheads rate
      const rigRental = 2500; // Updated rig rate
      const vbRental = 3250; // Updated vibro+powerpack rate
      const craneRental = 1100;
      const projectKey = normalizeText(project).toLowerCase();
      const getEquipmentCountsForDate = (dateKey = '') => {
        const normalizedDate = normalizeDateString(dateKey);
        if (projectKey === 'vintage') {
          return {
            rig: 1,
            vb: 0,
            crane: normalizedDate && normalizedDate >= '2026-04-08' ? 1 : 0
          };
        }
        if (projectKey === 'titania') {
          return {
            rig: normalizedDate && normalizedDate >= '2026-04-03' ? 2 : 1,
            vb: 1,
            crane: 1
          };
        }
        return { rig: 1, vb: 0, crane: 0 };
      };
      const equipmentCounts = getEquipmentCountsForDate(lastDayDateKey);

      const lastDayValues = {
        pm: lastDayCounts.pm * dailyRates.pm,
        se: lastDayCounts.se * dailyRates.se,
        foreman: lastDayCounts.foreman * dailyRates.foreman,
        op: lastDayCounts.op * dailyRates.op,
        vb: lastDayCounts.vb * dailyRates.vb,
        rig: lastDayCounts.rig * dailyRates.rig,
        we: lastDayCounts.we * dailyRates.we,
        me: lastDayCounts.me * dailyRates.me,
        hl: lastDayCounts.hl * dailyRates.hl,
        r_rig: lastDayDateKey ? equipmentCounts.rig * rigRental : 0,
        r_vb: lastDayDateKey ? equipmentCounts.vb * vbRental : 0,
        r_crane: lastDayDateKey ? equipmentCounts.crane * craneRental : 0
      };
      lastDayValues.salaries = lastDayValues.pm + lastDayValues.se + lastDayValues.foreman + lastDayValues.op + lastDayValues.vb + lastDayValues.rig + lastDayValues.we + lastDayValues.me + lastDayValues.hl;
      lastDayValues.rental = lastDayValues.r_rig + lastDayValues.r_vb + lastDayValues.r_crane;
      lastDayValues.direct = lastDayValues.salaries + lastDayValues.rental;
      lastDayValues.overheads = lastDayDateKey ? overheadsDaily : 0;
      lastDayValues.total = lastDayValues.direct + lastDayValues.overheads;

      const lastDayLm = lastDayDateKey
        ? executed
            .filter(r => getOverviewDateKey(r) === lastDayDateKey)
            .reduce((sum, r) => sum + (Number(r.asbuilt_depth) || Number(r.design_depth) || 0), 0)
        : 0;
      const lastDayCumLm = lastDayDateKey
        ? executed
            .filter(r => getOverviewDateKey(r) <= lastDayDateKey)
            .reduce((sum, r) => sum + (Number(r.asbuilt_depth) || Number(r.design_depth) || 0), 0)
        : 0;
      lastDayValues.lm = lastDayLm;
      lastDayValues.cumLm = lastDayCumLm;
      lastDayValues.cplm = lastDayLm > 0 ? (lastDayValues.total / lastDayLm) : 0;
      
      const getEquipDays = (w, eq, fallbackDays = 6) => {
        if (isTitaniaCostTemplateProject) {
          if (w === 10) return 1;
          if (eq === 'rig' || eq === 'crane') {
             if (w === 12) return 2; // Rig/Crane offline starting Wed Mar 18
             if (w === 13) return 3; // Rig/Crane offline until Wed Mar 25
             return fallbackDays;
          }
          if (eq === 'vb') {
            if (w === 12) return 2; // Vibrator offline Mar 18-21, 2026 (4 days)
            return fallbackDays;
          }
        }
        return fallbackDays;
      };

      const sums = { salaries: 0, rental: 0, direct: 0, overheads: 0, total: 0, lm: 0, r_crane: 0 };
      
      let runningCumLm = 0;
      const weeklyData = weeks.map(w => {
        const d = data[w];
        const days = Number.isFinite(d?.daysFromManpower) ? d.daysFromManpower : (w === 10 ? 1 : 6);
        const isDynamicManpowerWeek = d?.source === 'manpower' && Array.isArray(d?.dailyRows) && d.dailyRows.length > 0;

        const calcRoleCost = (roleKey, rate) => {
          if (!isDynamicManpowerWeek) {
            return (Number(d?.[roleKey]) || 0) * rate * days;
          }
          let roleTotal = 0;
          d.dailyRows.forEach(dayRow => {
            const count = Number(dayRow?.[roleKey]) || 0;
            roleTotal += count * rate;
          });
          return roleTotal;
        };
        
        const pm = calcRoleCost('pm', dailyRates.pm);
        const se = calcRoleCost('se', dailyRates.se);
        const foreman = calcRoleCost('foreman', dailyRates.foreman);
        const op = calcRoleCost('op', dailyRates.op);
        const vb = calcRoleCost('vb', dailyRates.vb);
        const rig = calcRoleCost('rig', dailyRates.rig);
        const we = calcRoleCost('we', dailyRates.we);
        const me = calcRoleCost('me', dailyRates.me);
        const hl = calcRoleCost('hl', dailyRates.hl);

        const salaries = pm + se + foreman + op + vb + rig + we + me + hl;

        const calcEquipmentRental = (equipmentKey, rate) => {
          return d.dailyRows.reduce((sum, dayRow) => {
            const count = Number(getEquipmentCountsForDate(dayRow?.date)?.[equipmentKey]) || 0;
            return sum + (count * rate);
          }, 0);
        };

        const getStaticEquipmentRental = (equipmentKey, rate) => {
          const count = Number(getEquipmentCountsForDate('')[equipmentKey]) || 0;
          return count * rate * getEquipDays(w, equipmentKey.replace('r_', ''), days);
        };

        const r_rig = isDynamicManpowerWeek ? calcEquipmentRental('rig', rigRental) : getStaticEquipmentRental('rig', rigRental);
        const r_vb = isDynamicManpowerWeek ? calcEquipmentRental('vb', vbRental) : getStaticEquipmentRental('vb', vbRental);
        const r_crane = isDynamicManpowerWeek ? calcEquipmentRental('crane', craneRental) : getStaticEquipmentRental('crane', craneRental);
        const rental = r_rig + r_vb + r_crane;

        const direct = salaries + rental;
        const overheads = overheadsDaily * days;
        const total = direct + overheads;
        const lm = cwLmMap.get(w) || 0;
        
        sums.pm = (sums.pm || 0) + pm;
        sums.se = (sums.se || 0) + se;
        sums.foreman = (sums.foreman || 0) + foreman;
        sums.op = (sums.op || 0) + op;
        sums.vb = (sums.vb || 0) + vb;
        sums.rig = (sums.rig || 0) + rig;
        sums.we = (sums.we || 0) + we;
        sums.me = (sums.me || 0) + me;
        sums.hl = (sums.hl || 0) + hl;

        sums.r_rig = (sums.r_rig || 0) + r_rig;
        sums.r_vb = (sums.r_vb || 0) + r_vb;
        sums.r_crane = (sums.r_crane || 0) + r_crane;

        sums.salaries += salaries;
        sums.rental += rental;
        sums.direct += direct;
        sums.overheads += overheads;
        sums.total += total;
        sums.lm += lm;
        runningCumLm += lm;
        const cumLm = runningCumLm;
        const cumTotal = sums.total;
        const cplm = cumLm > 0 ? cumTotal / cumLm : 0;

        return { w, days, salaries, pm, se, foreman, op, vb, rig, we, me, hl, rental, r_rig, r_vb, r_crane, direct, overheads, total, lm, cumLm, cumTotal, cplm };
      });
      const getEquipmentWeekMeta = (weekNumber, equipmentKey, fallbackDays = 0) => {
        const bucket = data[weekNumber];
        const isDynamicWeek = bucket?.source === 'manpower' && Array.isArray(bucket?.dailyRows) && bucket.dailyRows.length > 0;
        if (isDynamicWeek) {
          const counts = bucket.dailyRows
            .map(dayRow => Number(getEquipmentCountsForDate(dayRow?.date)?.[equipmentKey]) || 0)
            .filter(count => count > 0);
          if (!counts.length) return { hasValue: false, countText: '-', daysText: '-', count: 0, days: 0 };
          const uniqueCounts = Array.from(new Set(counts));
          if (uniqueCounts.length === 1) {
            return {
              hasValue: true,
              countText: `${uniqueCounts[0]} Nos`,
              daysText: `${counts.length} Days`,
              count: uniqueCounts[0],
              days: counts.length
            };
          }
          const minCount = Math.min(...counts);
          const maxCount = Math.max(...counts);
          return {
            hasValue: true,
            countText: `${minCount}-${maxCount} Nos`,
            daysText: `${counts.length} Days`,
            count: counts.reduce((sum, value) => sum + value, 0),
            days: counts.length
          };
        }

        const staticCount = Number(getEquipmentCountsForDate('')[equipmentKey]) || 0;
        if (!staticCount) return { hasValue: false, countText: '-', daysText: '-', count: 0, days: 0 };
        const effectiveDays = getEquipDays(weekNumber, equipmentKey, fallbackDays);
        return {
          hasValue: true,
          countText: `${staticCount} Nos`,
          daysText: `${effectiveDays} Days`,
          count: staticCount,
          days: effectiveDays
        };
      };

      function formatCost(val, isLm = false) {
        if (!val) return '-';
        return Number(val.toFixed(isLm ? 1 : 2)).toLocaleString('en-US');
      }

      const hasRigRental = weeklyData.some(wd => wd.r_rig > 0) || lastDayValues.r_rig > 0;
      const hasVbRental = weeklyData.some(wd => wd.r_vb > 0) || lastDayValues.r_vb > 0;
      const hasCraneRental = weeklyData.some(wd => wd.r_crane > 0) || lastDayValues.r_crane > 0;
      const rentalDetailRows = [];
      if (hasRigRental) rentalDetailRows.push({ targetPivot: 'rental', key: 'r_rig', label: 'Piling Rig' });
      if (hasVbRental) rentalDetailRows.push({ targetPivot: 'rental', key: 'r_vb', label: 'Vibrator & Power Pack' });
      if (hasCraneRental) rentalDetailRows.push({ targetPivot: 'rental', key: 'r_crane', label: 'Crawler Crane' });

      const tableRows = [
        { key: 'directSection', sectionLabel: 'Direct', sectionRow: true },
        { key: 'salaries', label: 'Personnel Cost', groupSub: true, hasPivot: true },
        { targetPivot: 'salaries', key: 'pm', label: 'Project Manager' },
        { targetPivot: 'salaries', key: 'se', label: 'Site Engineers' },
        { targetPivot: 'salaries', key: 'foreman', label: 'Foreman' },
        { targetPivot: 'salaries', key: 'op', label: 'Equipment Operators' },
        { targetPivot: 'salaries', key: 'vb', label: 'Vibro Operators' },
        { targetPivot: 'salaries', key: 'rig', label: 'Riggers' },
        { targetPivot: 'salaries', key: 'we', label: 'Welders' },
        { targetPivot: 'salaries', key: 'me', label: 'Mechanic' },
        { targetPivot: 'salaries', key: 'hl', label: 'Helpers' },

        { key: 'rental', label: 'Equipment Rental', groupSub: true, hasPivot: true },
        ...rentalDetailRows,

        { key: 'direct', label: '', groupHead: true },
        { key: 'overheadSection', sectionLabel: 'Indirect', sectionRow: true },
        { key: 'overheads', label: 'Overheads Share', groupSub: true },
        { key: 'overheads', label: '', groupHead: true },
        { key: 'totalBreak', sectionLabel: '', sectionRow: true },
        { key: 'total', label: 'Total Cost', totalRow: true },
        { key: 'productionSection', sectionLabel: 'Production', sectionRow: true },
        { key: 'lm', label: 'Executed Lm', groupSub: true },
        { key: 'cumLm', label: 'Cumulative Lm', groupHead: true },
        { key: 'unitRateSection', sectionLabel: 'Unit Rate Calculation', sectionRow: true },
        { key: 'cplm', label: 'Cost / Lm', groupHead: true }
      ];

      function getCostRowRate(row) {
        if (!row || !row.targetPivot) return null;
        if (row.targetPivot === 'salaries') return dailyRates[row.key] || null;
        if (row.key === 'r_rig') return rigRental;
        if (row.key === 'r_vb') return vbRental;
        if (row.key === 'r_crane') return craneRental;
        return null;
      }

      window.toggleCostPivot = function(btn, groupKey) {
        btn.classList.toggle('expanded');
        document.querySelectorAll(`.pivot-${groupKey}`).forEach(tr => {
          tr.classList.toggle('expanded');
        });
      };

      let html = '';
      tableRows.forEach(row => {
        if (row.sectionRow) {
          const sectionLabel = row.sectionLabel || '';
          html += `
            <tr class="section-row row-${row.key}${sectionLabel ? '' : ' empty-section-row'}">
              <td colspan="${weeks.length + 2}">
                <div class="section-row-inner">
                  ${sectionLabel ? `<span class="section-row-label">${sectionLabel}</span>` : ''}
                </div>
              </td>
            </tr>`;
          return;
        }
        let totalVal = row.key === 'cplm'
          ? lastDayValues.cplm
          : lastDayValues[row.key];
        if (row.key === 'cumLm') totalVal = null;

        const baseClass = row.totalRow ? 'total-row' : (row.groupHead ? 'group-head' : (row.targetPivot ? `pivot-child pivot-${row.targetPivot}` : 'group-sub'));
        const cssClass = `${baseClass} row-${row.key}`;
        html += `<tr class="${cssClass}">`;

        let labelHtml = row.label;
        if (row.hasPivot) {
          labelHtml = `<div class="cost-pivot-toggle" onclick="toggleCostPivot(this, '${row.key}')">${row.label}</div>`;
        } else if (row.targetPivot) {
          const rateLabel = getCostRowRate(row);
          labelHtml = `<div>${row.label}</div><div class="cost-rate-note">${rateLabel ? `${formatCost(rateLabel, true)} AED` : ''}</div>`;
        }

        html += `<td class="col-head">${labelHtml}</td>`;

        weeklyData.forEach(wd => {
          if (row.targetPivot) {
            let count = 0;
            let targetDays = wd.days;
            let countText = '-';
            let daysText = '-';
            if (row.targetPivot === 'salaries') {
              count = data[wd.w][row.key];
              countText = `${count} Nos`;
              daysText = `${targetDays} Days`;
            } else if (row.key === 'r_rig') {
              const meta = getEquipmentWeekMeta(wd.w, 'rig', wd.days);
              count = meta.count;
              targetDays = meta.days || wd.days;
              countText = meta.countText;
              daysText = meta.daysText;
            } else if (row.key === 'r_vb') {
              const meta = getEquipmentWeekMeta(wd.w, 'vb', wd.days);
              count = meta.count;
              targetDays = meta.days || wd.days;
              countText = meta.countText;
              daysText = meta.daysText;
            } else if (row.key === 'r_crane') {
              const meta = getEquipmentWeekMeta(wd.w, 'crane', wd.days);
              count = meta.count;
              targetDays = meta.days || wd.days;
              countText = meta.countText;
              daysText = meta.daysText;
            }
            if (count > 0 && Number(wd[row.key]) > 0) {
              html += `<td class="num cost-detail-cell">
                 <div class="cost-detail-value">${formatCost(wd[row.key])} AED</div>
                 <div class="cost-detail-meta">
                   <span class="cost-detail-count">${countText}</span>
                   <span class="cost-detail-times">x</span>
                   <span class="cost-detail-days">${daysText}</span>
                 </div>
               </td>`;
              return;
            }
            html += `<td class="num">-</td>`;
            return;
          }
          html += `<td class="num">${formatCost(wd[row.key], row.key === 'lm')}</td>`;
        });

        if (row.targetPivot) {
          let lastDayCount = 0;
          if (row.targetPivot === 'salaries') {
            lastDayCount = Number(lastDayCounts[row.key]) || 0;
          } else if (row.key === 'r_rig') {
            lastDayCount = lastDayDateKey ? equipmentCounts.rig : 0;
          } else if (row.key === 'r_vb') {
            lastDayCount = lastDayDateKey ? equipmentCounts.vb : 0;
          } else if (row.key === 'r_crane') {
            lastDayCount = lastDayDateKey ? equipmentCounts.crane : 0;
          }

          const lastDayValue = Number(lastDayValues[row.key]) || 0;
          if (lastDayCount > 0 || lastDayValue > 0) {
            html += `<td class="num total-col cost-detail-cell">
              <div class="cost-detail-value">${formatCost(lastDayValue)} AED</div>
              <div class="cost-detail-meta">
                <span class="cost-detail-count">${lastDayCount} Nos</span>
              </div>
            </td>`;
          } else {
            html += `<td class="num total-col">-</td>`;
          }
        } else {
          if (row.key === 'cumLm') {
            html += `<td class="num total-col"></td>`;
          } else {
            html += `<td class="num total-col">${formatCost(totalVal, row.key === 'lm')}</td>`;
          }
        }
        html += `</tr>`;
      });
      if (els.costTableBody) els.costTableBody.innerHTML = html;

      // Global Tooltip Context for inline HTML string calls
      let ttp = document.getElementById('costTooltip');
      if (!ttp) {
        ttp = document.createElement('div');
        ttp.id = 'costTooltip';
        ttp.className = 'chart-tooltip';
        ttp.style.position = 'fixed';
        document.body.appendChild(ttp);
      }
      window.showCostTtp = (e, html) => {
        ttp.innerHTML = html;
        ttp.classList.add('visible');
        let x = e.clientX + 14;
        let y = e.clientY + 14;
        if (x + 240 > window.innerWidth) x = e.clientX - 250;
        ttp.style.left = x + 'px';
        ttp.style.top = y + 'px';
      };
      window.hideCostTtp = () => ttp.classList.remove('visible');

      // Render Cost Trend stacked bar chart
      const svg = els.costTrendSvg;
      const wrap = document.getElementById('costTrendWrap') || (svg ? svg.parentElement : null);
      if (svg && wrap) {
        while (svg.firstChild) svg.removeChild(svg.firstChild);
        
        let legend = wrap.querySelector('.cost-chart-legend');
        if (!legend) {
           legend = document.createElement('div');
           legend.className = 'cost-chart-legend';
           legend.style.cssText = 'position:absolute; top:8px; width:100%; display:flex; gap:16px; z-index:5; justify-content:center; padding:6px 12px; pointer-events:none;';
           legend.innerHTML = `
             <div style="display:flex; align-items:center; gap:6px; font-size:11.5px; font-weight:800; color:rgba(255,255,255,0.7);"><span style="width:12px; height:12px; background:#4b92c6; border-radius:3px; opacity:0.9;"></span> Salaries</div>
             <div style="display:flex; align-items:center; gap:6px; font-size:11.5px; font-weight:800; color:rgba(255,255,255,0.7);"><span style="width:12px; height:12px; background:#c4e45f; border-radius:3px; opacity:0.9;"></span> Rental</div>
             <div style="display:flex; align-items:center; gap:6px; font-size:11.5px; font-weight:800; color:rgba(255,255,255,0.7);"><span style="width:12px; height:12px; background:#4ade80; border-radius:3px; opacity:0.9;"></span> Overheads</div>
           `;
           wrap.appendChild(legend);
        }
        const chartH = 320;
        const chartW = 1000;
        const maxTotal = Math.max(...weeklyData.map(d => d.total));
        const scaleY = (chartH - 60) / (maxTotal || 1);
        const stepX = chartW / (weeks.length + 1);
        
        weeklyData.forEach((wd, i) => {
          const x = stepX * (i + 1);
          let currentY = chartH - 30; // base line
          
          const stack = [
            { val: wd.salaries, color: '#4b92c6', label: 'Sal.' },
            { val: wd.rental, color: '#c4e45f', label: 'Rent.' },
            { val: wd.overheads, color: '#4ade80', label: 'Ovhd.' }
          ];

          stack.forEach((st, barIdx) => {
            if (!st.val) return;
            const h = st.val * scaleY;
            const y = currentY - h;
            
            const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
            rect.setAttribute('x', x - 30);
            rect.setAttribute('y', forceAnimate ? currentY : y);
            rect.setAttribute('width', 60);
            rect.setAttribute('height', forceAnimate ? 0 : h);
            rect.setAttribute('fill', st.color);
            rect.setAttribute('opacity', '0.9');
            rect.setAttribute('rx', '4');
            rect.setAttribute('stroke', 'rgba(255,255,255,0.08)');
            rect.setAttribute('stroke-width', '1');
            
            if (forceAnimate) {
               rect.style.transition = 'all 480ms cubic-bezier(0.25, 1, 0.5, 1)';
               setTimeout(() => {
                  rect.setAttribute('height', h);
                  rect.setAttribute('y', y);
               }, 50 + (i * 80) + (barIdx * 30)); 
            }
            svg.appendChild(rect);
            
            if (h > 24) {
               const lbl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
               lbl.setAttribute('x', x);
               lbl.setAttribute('y', y + h/2 + 5);
               lbl.setAttribute('fill', '#000');
               lbl.setAttribute('font-size', '15px');
               lbl.setAttribute('font-family', 'Inter, Arial, sans-serif');
               lbl.setAttribute('font-weight', '900');
               lbl.setAttribute('text-anchor', 'middle');
               lbl.setAttribute('opacity', forceAnimate ? '0' : '0.85');
               const pct = wd.total > 0 ? ((st.val / wd.total) * 100).toFixed(0) : 0;
               lbl.textContent = `${pct}%`;
               
               if (forceAnimate) {
                 lbl.style.transition = 'opacity 300ms ease';
                 setTimeout(() => lbl.setAttribute('opacity', '0.8'), 300 + i * 80);
               }
               svg.appendChild(lbl);
            }
            currentY -= h;
          });
          
          const xLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
          xLabel.setAttribute('x', x);
          xLabel.setAttribute('y', chartH - 10);
          xLabel.setAttribute('class', 'axis-label');
          xLabel.setAttribute('font-size', '14.5px');
          xLabel.setAttribute('font-weight', '700');
          xLabel.setAttribute('text-anchor', 'middle');
          xLabel.textContent = `Week ${wd.w}`;
          svg.appendChild(xLabel);

          // Hover Map
          const hoverRect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
          hoverRect.setAttribute('x', x - 35);
          hoverRect.setAttribute('y', 20);
          hoverRect.setAttribute('width', 70);
          hoverRect.setAttribute('height', chartH - 50);
          hoverRect.setAttribute('class', 'hover-target');
          
          const sPct = wd.total > 0 ? ((wd.salaries/wd.total)*100).toFixed(1) : 0;
          const rPct = wd.total > 0 ? ((wd.rental/wd.total)*100).toFixed(1) : 0;
          const oPct = wd.total > 0 ? ((wd.overheads/wd.total)*100).toFixed(1) : 0;
          
          hoverRect.addEventListener('mousemove', e => window.showCostTtp(e, `
            <div class="tooltip-title">Week ${wd.w} Breakdown</div>
            <div class="tooltip-row"><span>Salaries <span style="color:#60a5fa; font-size:10px;">(${sPct}%)</span></span><strong>${wd.salaries > 0 ? `${formatCost(wd.salaries)} AED` : '-'}</strong></div>
            <div class="tooltip-row"><span>Rental Eq. <span style="color:#a3e635; font-size:10px;">(${rPct}%)</span></span><strong>${wd.rental > 0 ? `${formatCost(wd.rental)} AED` : '-'}</strong></div>
            <div class="tooltip-row"><span>Overheads <span style="color:#4ade80; font-size:10px;">(${oPct}%)</span></span><strong>${wd.overheads > 0 ? `${formatCost(wd.overheads)} AED` : '-'}</strong></div>
            <hr style="border:0; border-top:1px solid rgba(255,255,255,0.1); margin:8px 0;">
            <div class="tooltip-row"><span>Total Cost</span><strong style="color:var(--accent)">${wd.total > 0 ? `${formatCost(wd.total)} AED` : '-'}</strong></div>
          `));
          hoverRect.addEventListener('mouseleave', window.hideCostTtp);
          svg.appendChild(hoverRect);
        });
      }

      // Render Cost per Lm Line Chart
      const lmSvg = els.costLmSvg;
      if (lmSvg) {
        while (lmSvg.firstChild) lmSvg.removeChild(lmSvg.firstChild);
        const chartH = 220;
        const chartW = 620;
        lmSvg.setAttribute('viewBox', `0 0 ${chartW} ${chartH}`);
        const stepX = chartW / (weeks.length + 1);
        const maxLmCost = Math.max(...weeklyData.map(d => d.cplm));
        const maxScaleScaleLmScale = maxLmCost > 0 ? maxLmCost * 1.2 : 1;
        const lmScaleY = (chartH - 72) / maxScaleScaleLmScale;
        let pathD = '';
        
        weeklyData.forEach((wd, i) => {
          const x = stepX * (i + 1);
          const y = chartH - 52 - (wd.cplm * lmScaleY);
          
          pathD += `${i === 0 ? 'M' : 'L'} ${x} ${y} `;
          
          const dot = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
          dot.setAttribute('cx', x);
          dot.setAttribute('cy', y);
          dot.setAttribute('r', '4.5');
          dot.setAttribute('fill', '#f4f7fb');
          dot.setAttribute('class', 'point-dot');
          dot.style.opacity = forceAnimate ? '0' : '1';
          
          const valL = document.createElementNS('http://www.w3.org/2000/svg', 'text');
          valL.setAttribute('x', x);
          valL.setAttribute('y', y - 12);
          valL.setAttribute('class', 'point-label');
          valL.setAttribute('font-size', '16px');
          valL.setAttribute('font-weight', '900');
          valL.setAttribute('fill', '#f4f7fb');
          valL.style.opacity = forceAnimate ? '0' : '1';
          valL.textContent = wd.cplm > 0 ? wd.cplm.toFixed(1) : "-";
          valL.setAttribute('text-anchor', 'middle');
          
          if (forceAnimate) {
             dot.style.transition = 'opacity 300ms ease';
             valL.style.transition = 'opacity 300ms ease';
             setTimeout(() => {
               dot.style.opacity = '1';
               valL.style.opacity = '1';
             }, 300 + i * 100);
          }
          
          lmSvg.appendChild(dot);
          lmSvg.appendChild(valL);
          
          const xLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
          xLabel.setAttribute('x', x);
          xLabel.setAttribute('y', chartH - 14);
          xLabel.setAttribute('class', 'axis-label');
          xLabel.setAttribute('font-size', '15px');
          xLabel.setAttribute('font-weight', '700');
          xLabel.setAttribute('fill', 'rgba(244,247,251,0.92)');
          xLabel.setAttribute('text-anchor', 'middle');
          xLabel.textContent = `Week ${wd.w}`;
          lmSvg.appendChild(xLabel);

          const hoverCircle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
          hoverCircle.setAttribute('cx', x);
          hoverCircle.setAttribute('cy', y);
          hoverCircle.setAttribute('r', '30');
          hoverCircle.setAttribute('class', 'hover-target');
          hoverCircle.addEventListener('mousemove', e => window.showCostTtp(e, `
            <div class="tooltip-title">Week ${wd.w} Efficiency</div>
            <div class="tooltip-row"><span>Cum. Executed</span><strong>${wd.cumLm > 0 ? `${formatCost(wd.cumLm, true)} Lm` : '-'}</strong></div>
            <div class="tooltip-row"><span>Cum. Total Cost</span><strong>${wd.cumTotal > 0 ? `${formatCost(wd.cumTotal)} AED` : '-'}</strong></div>
            <hr style="border:0; border-top:1px solid rgba(255,255,255,0.1); margin:8px 0;">
            <div class="tooltip-row"><span>Cum. Cost / Lm</span><strong style="color:var(--accent)">${wd.cplm > 0 ? `${wd.cplm.toFixed(1)} AED/Lm` : "-"}</strong></div>
          `));
          hoverCircle.addEventListener('mouseleave', window.hideCostTtp);
          lmSvg.appendChild(hoverCircle);
        });
        
        if (pathD) {
          const pathLine = document.createElementNS('http://www.w3.org/2000/svg', 'path');
          pathLine.setAttribute('d', pathD.trim());
          pathLine.setAttribute('fill', 'none');
          pathLine.setAttribute('stroke', 'rgba(142,240,191,0.95)');
          pathLine.setAttribute('stroke-width', '3');
          pathLine.setAttribute('class', 'line-shape');
          lmSvg.insertBefore(pathLine, lmSvg.firstChild);
          
          if (forceAnimate) {
             const len = 1500; // approximation since getTotalLength isn't pre-rendered
             pathLine.style.strokeDasharray = len;
             pathLine.style.strokeDashoffset = len;
             pathLine.style.transition = 'stroke-dashoffset 800ms cubic-bezier(.22,.61,.36,1)';
             requestAnimationFrame(() => {
                 requestAnimationFrame(() => {
                     pathLine.style.strokeDashoffset = '0';
                 });
             });
          }
        }
      }
    }

    function getManpowerCount(row, keys) {
      for (const key of keys) {
        const raw = row?.[key];
        if (raw === undefined || raw === null || raw === '') continue;
        const n = Number(raw);
        if (Number.isFinite(n)) return n;
      }
      return 0;
    }

    function getFilteredManpowerRows(project) {
      const rawProject = normalizeText(project || selectedProject || DEFAULT_PROJECT);
      const targetProject = isAllProjectsValue(rawProject) ? '' : rawProject;
      const filtered = manpowerRows.filter(row => {
        const projectMatch = !targetProject || normalizeText(row?.project) === targetProject;
        const plotMatch = isAllPlotsValue(selectedPlot) || normalizeText(row?.plot) === normalizeText(selectedPlot);
        return projectMatch && plotMatch;
      });

      const latestByDate = new Map();
      filtered.forEach(row => {
        const dateKey = normalizeDateString(row?.date);
        if (!dateKey) return;
        latestByDate.set(dateKey, row);
      });

      const rows = Array.from(latestByDate.entries())
        .sort((a, b) => b[0].localeCompare(a[0]))
        .map(([dateKey, row]) => {
          const item = {
            date: dateKey,
            pm: getManpowerCount(row, ['projectmanager', 'project manager']),
            se: getManpowerCount(row, ['siteengineer', 'site engineer', 'siteenginner']),
            foreman: getManpowerCount(row, ['foreman']),
            op: getManpowerCount(row, ['operator', 'operators']),
            vb: getManpowerCount(row, ['vibro operator', 'vibrooperator', 'vibro_operator']),
            rig: getManpowerCount(row, ['riggers', 'rigger']),
            we: getManpowerCount(row, ['welder', 'welders']),
            me: getManpowerCount(row, ['mechanic', 'mechanics']),
            hl: getManpowerCount(row, ['helpers', 'helper'])
          };
          item.total = item.pm + item.se + item.foreman + item.op + item.vb + item.rig + item.we + item.me + item.hl;
          return item;
        });

      return rows;
    }

    function renderManpowerHistogram(rows) {
      const svg = els.manpowerHistSvg;
      if (!svg) return;
      while (svg.firstChild) svg.removeChild(svg.firstChild);

      if (!rows.length) {
        const empty = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        empty.setAttribute('x', '500');
        empty.setAttribute('y', '160');
        empty.setAttribute('text-anchor', 'middle');
        empty.setAttribute('fill', 'rgba(244,247,251,0.62)');
        empty.setAttribute('font-size', '14');
        empty.setAttribute('font-weight', '700');
        empty.textContent = 'No manpower data available';
        svg.appendChild(empty);
        return;
      }

      const chartW = 1000;
      const chartH = 320;
      const left = 52;
      const right = 18;
      const top = 16;
      const bottom = 62;
      const innerH = chartH - top - bottom;
      const stepX = (chartW - left - right) / Math.max(rows.length, 1);
      const barW = Math.max(16, Math.min(40, stepX * 0.5));
      const maxVal = Math.max(...rows.map(r => r.total), 1);
      const scaleY = innerH / maxVal;
      svg.setAttribute('viewBox', `0 0 ${chartW} ${chartH}`);

      let tooltipEl = document.getElementById('manpowerTooltip');
      if (!tooltipEl) {
        tooltipEl = document.createElement('div');
        tooltipEl.id = 'manpowerTooltip';
        tooltipEl.className = 'chart-tooltip';
        tooltipEl.style.position = 'fixed';
        document.body.appendChild(tooltipEl);
      }

      const showTooltip = (evt, row) => {
        tooltipEl.innerHTML = `
          <div class="tooltip-title">${formatDateFullLabel(row.date)}</div>
          <div class="tooltip-row"><span>PM</span><strong>${row.pm}</strong></div>
          <div class="tooltip-row"><span>Engineers</span><strong>${row.se}</strong></div>
          <div class="tooltip-row"><span>Foreman</span><strong>${row.foreman}</strong></div>
          <div class="tooltip-row"><span>Operator</span><strong>${row.op}</strong></div>
          <div class="tooltip-row"><span>Vibro Op.</span><strong>${row.vb}</strong></div>
          <div class="tooltip-row"><span>Riggers</span><strong>${row.rig}</strong></div>
          <div class="tooltip-row"><span>Welders</span><strong>${row.we}</strong></div>
          <div class="tooltip-row"><span>Mechanic</span><strong>${row.me}</strong></div>
          <div class="tooltip-row"><span>Helpers</span><strong>${row.hl}</strong></div>
          <hr style="border:0; border-top:1px solid rgba(255,255,255,0.1); margin:8px 0;">
          <div class="tooltip-row"><span>Total</span><strong>${row.total}</strong></div>
        `;
        tooltipEl.classList.add('visible');
        const x = Math.min(window.innerWidth - 260, evt.clientX + 14);
        const y = Math.min(window.innerHeight - 220, evt.clientY + 14);
        tooltipEl.style.left = `${Math.max(8, x)}px`;
        tooltipEl.style.top = `${Math.max(8, y)}px`;
      };

      const hideTooltip = () => {
        tooltipEl.classList.remove('visible');
      };

      for (let i = 0; i < 4; i += 1) {
        const y = top + (innerH / 3) * i;
        const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        line.setAttribute('x1', String(left));
        line.setAttribute('x2', String(chartW - right));
        line.setAttribute('y1', String(y));
        line.setAttribute('y2', String(y));
        line.setAttribute('stroke', 'rgba(255,255,255,0.08)');
        line.setAttribute('stroke-width', '1');
        svg.appendChild(line);
      }

      rows.slice().reverse().forEach((row, idx) => {
        const x = left + stepX * idx + stepX / 2;
        const h = row.total * scaleY;
        const y = chartH - bottom - h;

        const bar = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        bar.setAttribute('x', String(x - barW / 2));
        bar.setAttribute('y', String(y));
        bar.setAttribute('width', String(barW));
        bar.setAttribute('height', String(h));
        bar.setAttribute('rx', '4');
        bar.setAttribute('fill', 'rgba(142,240,191,0.86)');
        bar.setAttribute('stroke', 'rgba(255,255,255,0.2)');
        bar.setAttribute('stroke-width', '1');
        bar.addEventListener('mousemove', evt => showTooltip(evt, row));
        bar.addEventListener('mouseleave', hideTooltip);
        svg.appendChild(bar);

        const val = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        val.setAttribute('x', String(x));
        val.setAttribute('y', String(Math.max(y - 6, 12)));
        val.setAttribute('text-anchor', 'middle');
        val.setAttribute('fill', 'rgba(244,247,251,0.92)');
        val.setAttribute('font-size', '16');
        val.setAttribute('font-weight', '800');
        val.textContent = String(row.total);
        svg.appendChild(val);

        const dateLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        dateLabel.setAttribute('x', String(x));
        dateLabel.setAttribute('y', String(chartH - 18));
        dateLabel.setAttribute('text-anchor', 'middle');
        dateLabel.setAttribute('fill', 'rgba(244,247,251,0.72)');
        dateLabel.setAttribute('font-size', '13');
        dateLabel.setAttribute('font-weight', '700');
        dateLabel.textContent = formatShortDateLabel(row.date);
        svg.appendChild(dateLabel);
      });

      svg.onmouseleave = hideTooltip;
    }

    function getProjectEquipmentDefinitions(project) {
      const projectKey = normalizeText(project || selectedProject || DEFAULT_PROJECT).toLowerCase();
      const definitions = {
        titania: [
          { ownership: 'Owned', category: 'rig', label: 'Sany SR285 #3', start: 'project-start', color: '#8ef0bf' },
          { ownership: 'Owned', category: 'rig', label: 'Sany SR285 #4', start: '2026-04-03', color: '#7bb6ff' },
          { ownership: 'Owned', category: 'crane', label: 'Sany CR203', start: 'project-start', color: '#f1d48b' },
          { ownership: 'Rented', category: 'vibrator', label: 'Vibrator', start: 'project-start', color: '#f5a6b8' }
        ],
        vintage: [
          { ownership: 'Owned', category: 'rig', label: 'Sany SR285 #5', start: 'project-start', color: '#8ef0bf' },
          { ownership: 'Owned', category: 'crane', label: 'Sany CR205', start: 'project-start', color: '#f1d48b' }
        ]
      };
      return definitions[projectKey] ? definitions[projectKey].map(item => ({ ...item })) : [];
    }

    function getProjectEquipmentDateRows(project) {
      const equipment = getProjectEquipmentDefinitions(project);
      const projectManpowerRows = getFilteredManpowerRows(project);
      const allDateKeys = new Set();

      projectManpowerRows.forEach(row => {
        const dateKey = normalizeDateString(row?.date);
        if (dateKey) allDateKeys.add(dateKey);
      });

      const sortedDates = Array.from(allDateKeys).sort();
      const projectStartDate = sortedDates[0] || normalizeDateString(new Date().toISOString());

      const resolvedEquipment = equipment.map(item => ({
        ...item,
        start: item.start === 'project-start' ? projectStartDate : item.start
      }));

      resolvedEquipment.forEach(item => {
        const dateKey = normalizeDateString(item.start);
        if (dateKey) allDateKeys.add(dateKey);
      });

      const fullSortedDates = Array.from(allDateKeys).sort();
      if (!fullSortedDates.length || !resolvedEquipment.length) return [];

      const startDate = fullSortedDates[0];
      const endDate = fullSortedDates[fullSortedDates.length - 1];
      const rows = [];

      let cursor = new Date(`${startDate}T00:00:00Z`);
      const end = new Date(`${endDate}T00:00:00Z`);

      while (cursor <= end) {
        const dateKey = cursor.toISOString().slice(0, 10);
        const activeItems = resolvedEquipment.filter(item => normalizeDateString(item.start) <= dateKey);
        const ownedRigs = activeItems.filter(item => item.ownership === 'Owned' && item.category === 'rig').length;
        const ownedCranes = activeItems.filter(item => item.ownership === 'Owned' && item.category === 'crane').length;
        const rentedEq = activeItems.filter(item => item.ownership === 'Rented').length;
        rows.push({
          date: dateKey,
          ownedRigs,
          ownedCranes,
          rentedEq,
          total: ownedRigs + ownedCranes + rentedEq,
          activeItems
        });
        cursor.setUTCDate(cursor.getUTCDate() + 1);
      }

      return rows;
    }

    function getCompanyProjectOrder() {
      const baseProjects = Array.from(new Set(rawRows.map(row => normalizeText(row?.project)).filter(Boolean)));
      const companyProjects = Array.from(new Set(companyManpowerRows.map(row => getCompanyProjectLabel(row.Project || row.project)).filter(Boolean)));
      const ordered = [];
      [...baseProjects, ...companyProjects].forEach(project => {
        if (!ordered.includes(project)) ordered.push(project);
      });
      return ordered;
    }

    function updateCompanyDesignationOptions() {
      if (!els.companyManpowerDesignationSelect) return;
      const designations = Array.from(new Set(
        companyManpowerRows
          .map(sanitizeCompanyEmployee)
          .map(employee => employee.designation)
          .filter(Boolean)
      )).sort((a, b) => a.localeCompare(b));

      const currentValue = companyManpowerDesignationFilter || 'all';
      els.companyManpowerDesignationSelect.innerHTML = [
        '<option value="all">All Designations</option>',
        ...designations.map(designation => `<option value="${escapeHtml(designation)}">${escapeHtml(designation)}</option>`)
      ].join('');

      els.companyManpowerDesignationSelect.value = designations.includes(currentValue) ? currentValue : 'all';
      companyManpowerDesignationFilter = els.companyManpowerDesignationSelect.value;
    }

    function getCompanyColumnWidth(projectName) {
      return companyManpowerColumnWidths[projectName] || 320;
    }

    function ensureCompanyColumnWidths(projects) {
      projects.forEach(projectName => {
        if (!companyManpowerColumnWidths[projectName]) {
          companyManpowerColumnWidths[projectName] = 320;
        }
      });
    }

    function getFilteredCompanyEmployees(project = selectedProject) {
      const targetProjectToken = getCompanyProjectToken(project || selectedProject || DEFAULT_PROJECT);
      return companyManpowerRows
        .map(sanitizeCompanyEmployee)
        .filter(employee => employee.employeeName && employee.designation)
        .filter(employee => {
          if (companyManpowerScopeMode === 'all') return true;
          return getCompanyProjectToken(employee.projectRaw || employee.project) === targetProjectToken;
        })
        .filter(employee => companyManpowerDesignationFilter === 'all' || employee.designation === companyManpowerDesignationFilter);
    }

    function getCompanyProjectColumns(project = selectedProject) {
      const filteredEmployees = getFilteredCompanyEmployees(project);
      const projectOrder = getCompanyProjectOrder();
      const filteredProjects = Array.from(new Set(filteredEmployees.map(employee => employee.project)));
      filteredProjects.sort((a, b) => {
        const ai = projectOrder.indexOf(a);
        const bi = projectOrder.indexOf(b);
        if (ai === -1 && bi === -1) return a.localeCompare(b);
        if (ai === -1) return 1;
        if (bi === -1) return -1;
        return ai - bi;
      });
      return filteredProjects;
    }

    function groupCompanyEmployeesForMatrix(project = selectedProject) {
      const employees = getFilteredCompanyEmployees(project);
      const projects = getCompanyProjectColumns(project);
      const designationMap = new Map();

      employees.forEach(employee => {
        if (!designationMap.has(employee.designation)) designationMap.set(employee.designation, new Map());
        const projectMap = designationMap.get(employee.designation);
        if (!projectMap.has(employee.project)) projectMap.set(employee.project, []);
        projectMap.get(employee.project).push(employee);
      });

      const designations = Array.from(designationMap.keys()).sort((a, b) => a.localeCompare(b));
      return {
        projects,
        designationGroups: designations.map(designation => {
          const projectMap = designationMap.get(designation);
          const counts = projects.map(projectName => (projectMap.get(projectName) || []).length);
          const rowCount = Math.max(...counts, 1);
          const total = counts.reduce((sum, value) => sum + value, 0);
          projects.forEach(projectName => {
            const list = projectMap.get(projectName) || [];
            list.sort((a, b) => a.employeeName.localeCompare(b.employeeName));
          });
          return { designation, projectMap, rowCount, total };
        }),
        totalEmployees: employees.length
      };
    }

    function updateCompanyExportProjectOptions() {
      if (!els.companyExportProjectSelect) return;
      const projects = getCompanyProjectOrder();
      els.companyExportProjectSelect.innerHTML = [
        '<option value="all">All Projects</option>',
        ...projects.map(project => `<option value="${escapeHtml(project)}">${escapeHtml(project)}</option>`)
      ].join('');
    }

    function toggleCompanyExportPanel(show) {
      if (!els.companyExportPanel) return;
      els.companyExportPanel.hidden = !show;
    }

    function exportCompanyManpowerWorkbook(scopeProject) {
      const normalizedScope = normalizeText(scopeProject).toLowerCase();
      const rows = companyManpowerRows
        .map(sanitizeCompanyEmployee)
        .filter(employee => employee.employeeName && employee.designation)
        .filter(employee => normalizedScope === 'all' || employee.project === scopeProject)
        .sort((a, b) => (
          a.project.localeCompare(b.project) ||
          a.designation.localeCompare(b.designation) ||
          a.employeeName.localeCompare(b.employeeName)
        ));

      const title = normalizedScope === 'all' ? 'APFC Company Manpower' : `${scopeProject} Manpower`;
      const generatedAt = new Date().toLocaleString('en-GB', { timeZone: 'Asia/Dubai' });
      const tableRows = rows.map(row => `
        <tr>
          <td>${escapeHtml(row.employeeNumber || '-')}</td>
          <td>${escapeHtml(row.employeeName)}</td>
          <td>${escapeHtml(row.designation)}</td>
          <td>${escapeHtml(row.project)}</td>
          <td>${escapeHtml(row.shift || '-')}</td>
          <td>${escapeHtml(row.campNumber || '-')}</td>
          <td>${escapeHtml(row.roomNumber || '-')}</td>
          <td>${escapeHtml(row.joiningDate || '-')}</td>
          <td>${escapeHtml(row.remarks || '-')}</td>
        </tr>
      `).join('');

      const workbookHtml = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
          <head>
            <meta charset="UTF-8">
            <style>
              body { font-family: Segoe UI, Arial, sans-serif; color: #0e1b24; }
              h1 { margin: 0 0 6px; font-size: 22px; }
              p { margin: 0 0 14px; color: #41515e; font-size: 12px; }
              table { border-collapse: collapse; width: 100%; }
              th, td { border: 1px solid #c6d3dc; padding: 8px 10px; font-size: 11px; text-align: left; }
              th { background: #0f1720; color: #ffffff; font-weight: 700; text-transform: uppercase; letter-spacing: 0.04em; }
              tr:nth-child(even) td { background: #f4f8fb; }
            </style>
          </head>
          <body>
            <h1>${escapeHtml(title)}</h1>
            <p>Generated from APFC Project Dashboard on ${escapeHtml(generatedAt)}</p>
            <table>
              <thead>
                <tr>
                  <th>Employee Number</th>
                  <th>Employee Name</th>
                  <th>Designation</th>
                  <th>Project</th>
                  <th>Shift</th>
                  <th>Camp Number</th>
                  <th>Room Number</th>
                  <th>Joining Date</th>
                  <th>Remarks</th>
                </tr>
              </thead>
              <tbody>${tableRows}</tbody>
            </table>
          </body>
        </html>
      `;

      const blob = new Blob([workbookHtml], { type: 'application/vnd.ms-excel;charset=utf-8;' });
      const link = document.createElement('a');
      const slug = normalizedScope === 'all' ? 'all-projects' : scopeProject.toLowerCase().replace(/[^a-z0-9]+/g, '-');
      link.href = URL.createObjectURL(blob);
      link.download = `apfc-manpower-${slug}.xls`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(link.href);
    }

    function renderCompanyManpowerPage(project = selectedProject) {
      const layout = els.companyManpowerLayoutSelect?.value || 'project-matrix';
      const { projects, designationGroups, totalEmployees } = groupCompanyEmployeesForMatrix(project);

      if (els.companyManpowerSubtitle) {
        els.companyManpowerSubtitle.textContent = layout === 'project-matrix'
          ? 'Employee comparison by designation and project'
          : 'Employee comparison by designation and project';
      }

      if (els.companyManpowerDesignationSelect && els.companyManpowerDesignationSelect.value !== companyManpowerDesignationFilter) {
        els.companyManpowerDesignationSelect.value = companyManpowerDesignationFilter;
      }

      els.companyManpowerScopeButtons.forEach(btn => {
        btn.classList.toggle('active', btn.dataset.companyScope === companyManpowerScopeMode);
      });

      ensureCompanyColumnWidths(projects);

      if (els.companyManpowerColGroup) {
        els.companyManpowerColGroup.innerHTML = [
          '<col style="width:220px;">',
          ...projects.map(projectName => `<col style="width:${getCompanyColumnWidth(projectName)}px;">`)
        ].join('');
      }

      if (els.companyManpowerSummaryMeta) {
        const scopeText = companyManpowerScopeMode === 'all' ? 'All projects' : `Filtered by ${project}`;
        els.companyManpowerSummaryMeta.textContent = `${scopeText} • ${projects.length} project column${projects.length === 1 ? '' : 's'} • ${totalEmployees} employee${totalEmployees === 1 ? '' : 's'}`;
      }

      if (els.companyManpowerHeadRow) {
        els.companyManpowerHeadRow.innerHTML = [
          '<th>Designation</th>',
          ...projects.map(projectName => {
            const count = getFilteredCompanyEmployees(project).filter(employee => employee.project === projectName).length;
            return `<th class="company-project-th" data-company-col="${escapeHtml(projectName)}"><div class="company-project-head"><strong>${escapeHtml(projectName)}</strong><span>${count} Staff</span></div><span class="company-col-resizer" data-company-col="${escapeHtml(projectName)}" aria-hidden="true"></span></th>`;
          })
        ].join('');
      }

      if (!els.companyManpowerMatrixBody) return;

      if (!projects.length || !designationGroups.length) {
        els.companyManpowerMatrixBody.innerHTML = `<tr><td colspan="${Math.max(projects.length + 1, 2)}" class="manpower-empty">No company manpower records available for the current scope.</td></tr>`;
        return;
      }

      els.companyManpowerMatrixBody.innerHTML = designationGroups.map(group => {
        return Array.from({ length: group.rowCount }, (_, rowIndex) => {
          const cells = projects.map(projectName => {
            const employee = (group.projectMap.get(projectName) || [])[rowIndex];
            if (!employee) {
              return '<td data-company-col="' + escapeHtml(projectName) + '"><div class="company-empty-slot">—</div></td>';
            }
            return `
              <td data-company-col="${escapeHtml(projectName)}">
                <div class="company-employee-card">
                  <div class="company-employee-name">${escapeHtml(employee.employeeName)}</div>
                  <div class="company-employee-meta">${escapeHtml(employee.shift || '-')}</div>
                </div>
              </td>
            `;
          }).join('');

          return `
            <tr>
              ${rowIndex === 0 ? `<td rowspan="${group.rowCount}" class="company-designation-cell"><div class="company-designation-name">${escapeHtml(group.designation)}</div><div class="company-designation-meta">${group.total} Staff</div></td>` : ''}
              ${cells}
            </tr>
          `;
        }).join('');
      }).join('');

      bindCompanyManpowerHoverState();
      bindCompanyColumnResize();
    }

    function bindCompanyManpowerHoverState() {
      const table = document.getElementById('companyManpowerTable');
      const body = els.companyManpowerMatrixBody;
      const headRow = els.companyManpowerHeadRow;
      if (!table || !body || !headRow || table.dataset.hoverBound === 'true') return;

      const clearHover = () => {
        table.querySelectorAll('.company-hover-row, .company-hover-col, .company-hover-cell').forEach(el => {
          el.classList.remove('company-hover-row', 'company-hover-col', 'company-hover-cell');
        });
      };

      body.addEventListener('mouseover', evt => {
        const td = evt.target.closest('td[data-company-col]');
        if (!td) return;
        clearHover();
        const tr = td.parentElement;
        const colName = td.dataset.companyCol;
        tr?.querySelectorAll('td').forEach(cell => cell.classList.add('company-hover-row'));
        td.classList.add('company-hover-cell');
        headRow.querySelectorAll('th').forEach(th => {
          const strong = th.querySelector('strong');
          if (strong && strong.textContent === colName) th.classList.add('company-hover-col');
        });
        table.querySelectorAll(`td[data-company-col="${CSS.escape(colName)}"]`).forEach(cell => {
          cell.classList.add('company-hover-col');
        });
      });

      body.addEventListener('mouseleave', clearHover);
      table.dataset.hoverBound = 'true';
    }

    function bindCompanyColumnResize() {
      const table = document.getElementById('companyManpowerTable');
      const headRow = els.companyManpowerHeadRow;
      if (!table || !headRow || table.dataset.resizeBound === 'true') return;

      let dragState = null;

      const stopDrag = () => {
        if (!dragState) return;
        dragState.header.classList.remove('company-col-resizing');
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
        dragState = null;
      };

      const onMove = evt => {
        if (!dragState) return;
        const delta = evt.clientX - dragState.startX;
        const nextWidth = Math.max(220, dragState.startWidth + delta);
        companyManpowerColumnWidths[dragState.project] = nextWidth;
        const col = els.companyManpowerColGroup?.children[dragState.index + 1];
        if (col) col.style.width = `${nextWidth}px`;
      };

      headRow.addEventListener('mousedown', evt => {
        const handle = evt.target.closest('.company-col-resizer');
        if (!handle) return;
        const project = handle.dataset.companyCol;
        const header = handle.closest('th');
        const index = Array.from(headRow.children).indexOf(header) - 1;
        if (!project || !header || index < 0) return;
        evt.preventDefault();
        dragState = {
          project,
          header,
          index,
          startX: evt.clientX,
          startWidth: getCompanyColumnWidth(project)
        };
        header.classList.add('company-col-resizing');
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';
      });

      window.addEventListener('mousemove', onMove);
      window.addEventListener('mouseup', stopDrag);
      table.dataset.resizeBound = 'true';
    }

    function renderEquipmentHistogram(rows) {
      const svg = els.equipmentHistSvg;
      if (!svg) return;
      while (svg.firstChild) svg.removeChild(svg.firstChild);

      const wrap = svg.parentElement;
      if (wrap) {
        let legend = wrap.parentElement.querySelector('.equipment-chart-legend');
        if (!legend) {
          legend = document.createElement('div');
          legend.className = 'equipment-chart-legend';
          legend.innerHTML = `
            <div class="equipment-chart-legend-item"><span class="equipment-chart-legend-swatch" style="background:#8ef0bf;"></span>Owned Rigs</div>
            <div class="equipment-chart-legend-item"><span class="equipment-chart-legend-swatch" style="background:#f1d48b;"></span>Owned Cranes</div>
            <div class="equipment-chart-legend-item"><span class="equipment-chart-legend-swatch" style="background:#f5a6b8;"></span>Others</div>
          `;
          wrap.parentElement.insertBefore(legend, wrap);
        }
      }

      if (!rows.length) {
        if (wrap) {
          wrap.style.overflowX = 'hidden';
          wrap.style.overflowY = 'hidden';
        }
        const empty = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        empty.setAttribute('x', '500');
        empty.setAttribute('y', '160');
        empty.setAttribute('text-anchor', 'middle');
        empty.setAttribute('fill', 'rgba(244,247,251,0.62)');
        empty.setAttribute('font-size', '14');
        empty.setAttribute('font-weight', '700');
        empty.textContent = 'No equipment timeline available';
        svg.appendChild(empty);
        return;
      }

      const maxVisibleBars = 14;
      const chartH = 320;
      const left = 52;
      const right = 18;
      const top = 16;
      const bottom = 62;
      const isScrollable = rows.length > maxVisibleBars;
      const overflowSlotW = 58;
      const overflowBarW = 28;
      const overflowChartW = left + right + ((rows.length - 1) * overflowSlotW) + overflowBarW;
      const chartW = isScrollable ? overflowChartW : 1000;
      const innerH = chartH - top - bottom;
      const stepX = isScrollable ? overflowSlotW : (chartW - left - right) / Math.max(rows.length, 1);
      const barW = isScrollable ? overflowBarW : Math.max(16, Math.min(40, stepX * 0.5));
      const maxVal = Math.max(...rows.map(r => r.total), 1);
      const scaleY = innerH / maxVal;
      svg.setAttribute('viewBox', `0 0 ${chartW} ${chartH}`);
      if (wrap) {
        wrap.classList.toggle('equipment-chart-wrap-scrollable', isScrollable);
        wrap.style.overflowX = isScrollable ? 'auto' : 'hidden';
        wrap.style.overflowY = 'hidden';
        wrap.style.webkitOverflowScrolling = 'touch';
        if (!isScrollable) wrap.scrollLeft = 0;
      }
      svg.style.width = isScrollable ? `${chartW}px` : '100%';
      svg.style.minWidth = isScrollable ? `${chartW}px` : '100%';
      svg.style.height = '100%';

      let tooltipEl = document.getElementById('equipmentTooltip');
      if (!tooltipEl) {
        tooltipEl = document.createElement('div');
        tooltipEl.id = 'equipmentTooltip';
        tooltipEl.className = 'chart-tooltip';
        tooltipEl.style.position = 'fixed';
        document.body.appendChild(tooltipEl);
      }

      const showTooltip = (evt, row) => {
        const formatEquipmentType = value => {
          if (value === 'rig') return 'Rig';
          if (value === 'crane') return 'Crane';
          if (value === 'vibrator') return 'Other';
          return value;
        };
        const fleetHtml = row.activeItems.length
          ? row.activeItems.map(item => `<div class="tooltip-row"><span>${formatEquipmentType(item.category)}</span><strong>${escapeHtml(item.label)} | ${item.ownership}</strong></div>`).join('')
          : '<div class="tooltip-row"><span>Fleet</span><strong>-</strong></div>';
        tooltipEl.innerHTML = `
          <div class="tooltip-title">${formatDateFullLabel(row.date)}</div>
          ${fleetHtml}
        `;
        tooltipEl.classList.add('visible');
        const x = Math.min(window.innerWidth - 280, evt.clientX + 14);
        const y = Math.min(window.innerHeight - 260, evt.clientY + 14);
        tooltipEl.style.left = `${Math.max(8, x)}px`;
        tooltipEl.style.top = `${Math.max(8, y)}px`;
      };

      const hideTooltip = () => tooltipEl.classList.remove('visible');

      for (let i = 0; i < 4; i += 1) {
        const y = top + (innerH / 3) * i;
        const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        line.setAttribute('x1', String(left));
        line.setAttribute('x2', String(chartW - right));
        line.setAttribute('y1', String(y));
        line.setAttribute('y2', String(y));
        line.setAttribute('stroke', 'rgba(255,255,255,0.08)');
        line.setAttribute('stroke-width', '1');
        svg.appendChild(line);
      }

      rows.forEach((row, idx) => {
        const barX = isScrollable ? (left + stepX * idx) : (left + stepX * idx + (stepX - barW) / 2);
        const x = barX + barW / 2;
        const segments = [
          { value: row.ownedRigs, color: '#8ef0bf' },
          { value: row.ownedCranes, color: '#f1d48b' },
          { value: row.rentedEq, color: '#f5a6b8' }
        ];

        let currentY = chartH - bottom;
        segments.forEach(segment => {
          if (!segment.value) return;
          const h = segment.value * scaleY;
          currentY -= h;
          const bar = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
          bar.setAttribute('x', String(barX));
          bar.setAttribute('y', String(currentY));
          bar.setAttribute('width', String(barW));
          bar.setAttribute('height', String(h));
          bar.setAttribute('rx', '4');
          bar.setAttribute('fill', segment.color);
          bar.setAttribute('fill-opacity', '0.92');
          bar.setAttribute('stroke', 'rgba(255,255,255,0.18)');
          bar.setAttribute('stroke-width', '1');
          bar.addEventListener('mousemove', evt => showTooltip(evt, row));
          bar.addEventListener('mouseleave', hideTooltip);
          svg.appendChild(bar);
        });

        const totalLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        totalLabel.setAttribute('x', String(x));
        totalLabel.setAttribute('y', String(Math.max(currentY - 6, 12)));
        totalLabel.setAttribute('text-anchor', 'middle');
        totalLabel.setAttribute('fill', 'rgba(244,247,251,0.94)');
        totalLabel.setAttribute('font-size', '15');
        totalLabel.setAttribute('font-weight', '800');
        totalLabel.textContent = String(row.total);
        svg.appendChild(totalLabel);

        const dateLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        dateLabel.setAttribute('x', String(x));
        dateLabel.setAttribute('y', String(chartH - 18));
        dateLabel.setAttribute('text-anchor', 'middle');
        dateLabel.setAttribute('fill', 'rgba(244,247,251,0.72)');
        dateLabel.setAttribute('font-size', '13');
        dateLabel.setAttribute('font-weight', '700');
        dateLabel.textContent = formatShortDateLabel(row.date);
        svg.appendChild(dateLabel);
      });

      if (wrap && isScrollable) {
        requestAnimationFrame(() => {
          wrap.scrollLeft = Math.max(0, wrap.scrollWidth - wrap.clientWidth);
        });
        if (!wrap.dataset.equipmentWheelBound) {
          wrap.addEventListener('wheel', evt => {
            if (Math.abs(evt.deltaY) <= Math.abs(evt.deltaX)) return;
            evt.preventDefault();
            wrap.scrollLeft += evt.deltaY;
          }, { passive: false });
          wrap.dataset.equipmentWheelBound = 'true';
        }
      }

      svg.onmouseleave = hideTooltip;
    }

    function getUtilizationRigColors(rigKeys) {
      const palette = ['#8ef0bf', '#7bb6ff', '#f0d58e', '#f5a6b8', '#c7b8ff', '#7fe7f0'];
      return rigKeys.reduce((acc, rig, idx) => {
        acc[rig] = palette[idx % palette.length];
        return acc;
      }, {});
    }

    function getUtilizationRows(project) {
      const rows = getRowsForProject(project);
      const grouped = new Map();

      const getRigNameForUtilization = row => normalizeText(
        row?.machine ||
        row?.Machine ||
        row?.rig ||
        row?.Rig ||
        row?.machine_no ||
        row?.machineNo ||
        row?.machineno
      );

      const getDrillingHoursForUtilization = row => {
        const direct = Number(
          row?.asbuilt_durationdrilling ??
          row?.asbuilt_DurationDrilling ??
          row?.asbuilt_drillingduration ??
          row?.durationdrilling
        );
        if (Number.isFinite(direct) && direct >= 0) return direct;
        const drillStart = getRowDate(row, ['asbuilt_drillingStart', 'asbuilt_drillingstart', 'asbuilt_DrillingStart']);
        const drillEnd = getRowDate(row, ['asbuilt_drillingEnd', 'asbuilt_drillingend', 'asbuilt_DrillingEnd']);
        return hoursBetween(drillStart, drillEnd) || 0;
      };

      const utilizationCandidates = rows.filter(row => {
        const dateKey = getOverviewDateKey(row);
        const rig = getRigNameForUtilization(row);
        const drillingHours = getDrillingHoursForUtilization(row);
        const hasExecutionFlag = row?.isExecuted === true;
        const hasConcreteEnd = !!getRowDate(row, ['asbuilt_concreteEnd', 'asbuilt_concreteend', 'asbuilt_ConcreteEnd']);
        return !!dateKey && !!rig && (hasExecutionFlag || hasConcreteEnd || drillingHours > 0);
      });

      utilizationCandidates.forEach(row => {
        const dateKey = getOverviewDateKey(row);
        const rig = getRigNameForUtilization(row);
        if (!dateKey || !rig) return;
        const drillingHours = getDrillingHoursForUtilization(row);
        const drilledLm = Number.isFinite(Number(row?.asbuilt_depth)) ? Number(row.asbuilt_depth) : 0;
        const key = `${dateKey}__${rig}`;
        if (!grouped.has(key)) {
          grouped.set(key, { date: dateKey, rig, drillingHours: 0, drilledLm: 0, shiftHours: 12 });
        }
        grouped.get(key).drillingHours += drillingHours;
        grouped.get(key).drilledLm += drilledLm;
      });

      return Array.from(grouped.values())
        .map(item => ({
          ...item,
          utilization: item.shiftHours > 0 ? (item.drillingHours / item.shiftHours) * 100 : 0
        }))
        .sort((a, b) => b.date.localeCompare(a.date) || a.rig.localeCompare(b.rig, undefined, { numeric: true }));
    }

    function buildUtilizationSeries(rows, mode = utilizationMode) {
      if (!rows.length) return { dates: [], rigs: [], series: [], colors: {} };
      const dates = Array.from(new Set(rows.map(row => row.date))).sort();
      const rigs = Array.from(new Set(rows.map(row => row.rig))).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
      const colors = getUtilizationRigColors(rigs);

      const byRigDate = new Map();
      rows.forEach(row => {
        byRigDate.set(`${row.rig}__${row.date}`, row);
      });

      const series = rigs.map(rig => {
        let cumulativeHours = 0;
        let activeShiftDays = 0;
        let mobilized = false;
        const points = dates.map((dateKey) => {
          const row = byRigDate.get(`${rig}__${dateKey}`);
          const drillingHours = row ? row.drillingHours : 0;
          const drilledLm = row ? row.drilledLm : 0;
          const dailyUtilization = row ? row.utilization : 0;
          if (row) mobilized = true;
          if (mobilized) activeShiftDays += 1;
          cumulativeHours += drillingHours;
          const cumulativeUtilization = mobilized && activeShiftDays > 0
            ? (cumulativeHours / (activeShiftDays * 12)) * 100
            : null;
          return {
            date: dateKey,
            rig,
            drillingHours,
            cumulativeHours,
            drilledLm,
            utilization: dailyUtilization,
            cumulativeUtilization,
            cumulativeLm: drilledLm,
            mobilized
          };
        });
        let runningLm = 0;
        points.forEach(point => {
          runningLm += point.drilledLm;
          point.cumulativeLm = runningLm;
        });
        return {
          rig,
          color: colors[rig],
          points
        };
      });

      return { dates, rigs, series, colors };
    }

    function renderUtilizationLegend(series) {
      if (!els.utilizationLegend) return;
      if (!series.length) {
        els.utilizationLegend.innerHTML = '';
        return;
      }
      els.utilizationLegend.innerHTML = series.map(item => `
        <div class="utilization-legend-item">
          <span class="utilization-legend-swatch" style="background:${item.color};"></span>
          <span>${escapeHtml(item.rig)}</span>
        </div>
      `).join('');
    }

    function getUtilizationDayGroups(rows) {
      const grouped = new Map();
      rows.forEach(row => {
        if (!grouped.has(row.date)) grouped.set(row.date, []);
        grouped.get(row.date).push(row);
      });
      return Array.from(grouped.entries())
        .sort((a, b) => b[0].localeCompare(a[0]))
        .map(([date, items]) => {
          const drillingHours = items.reduce((sum, item) => sum + item.drillingHours, 0);
          const shiftHours = items.reduce((sum, item) => sum + item.shiftHours, 0);
          return {
            date,
            items: items.slice().sort((a, b) => a.rig.localeCompare(b.rig, undefined, { numeric: true })),
            totalUtilization: shiftHours > 0 ? (drillingHours / shiftHours) * 100 : 0
          };
        });
    }

    function getUtilizationChartLayout(wrapEl, dates, rigCount, mode, chartH, options = {}) {
      const left = options.left ?? 78;
      const right = options.right ?? 24;
      const top = options.top ?? 18;
      const bottom = options.bottom ?? 72;
      const overflowThreshold = options.overflowThreshold ?? (mode === 'daily' ? 8 : 10);
      const minDaySpan = mode === 'daily'
        ? Math.max(options.minDailySpan ?? 58, rigCount * (options.dailyRigFactor ?? 16) + (rigCount - 1) * (options.dailyGapFactor ?? 4))
        : (options.cumulativeSpan ?? 72);
      const minWrapWidth = options.minWrapWidth ?? 760;
      const wrapWidth = Math.max((wrapEl?.clientWidth || 0) - 4, minWrapWidth);
      const requiredChartW = left + right + (dates.length * minDaySpan);
      const shouldOverflow = dates.length > overflowThreshold && requiredChartW > wrapWidth;
      const chartW = shouldOverflow ? Math.max(wrapWidth, requiredChartW) : wrapWidth;
      const innerW = chartW - left - right;
      const innerH = chartH - top - bottom;
      const plotSidePad = mode === 'daily'
        ? (options.dailyPlotSidePad ?? 26)
        : (options.cumulativePlotSidePad ?? 14);
      const effectiveW = Math.max(innerW - plotSidePad * 2, 1);
      const stepX = dates.length > 1 ? effectiveW / (dates.length - 1) : 0;
      return { chartW, chartH, left, right, top, bottom, innerW, innerH, plotSidePad, stepX, shouldOverflow };
    }

    function bindUtilizationChartScrolls() {
      const primary = els.utilizationChartWrap;
      const secondary = els.utilizationLmChartWrap;
      if (!primary || !secondary || primary.dataset.syncBound) return;

      let syncing = false;
      const syncScroll = (source, target) => {
        if (syncing) return;
        syncing = true;
        target.scrollLeft = source.scrollLeft;
        requestAnimationFrame(() => { syncing = false; });
      };

      primary.addEventListener('scroll', () => syncScroll(primary, secondary));
      secondary.addEventListener('scroll', () => syncScroll(secondary, primary));

      const wheelScroll = (evt) => {
        const canScrollX = primary.scrollWidth > primary.clientWidth + 4;
        if (!canScrollX) return;
        const delta = Math.abs(evt.deltaY) > Math.abs(evt.deltaX) ? evt.deltaY : evt.deltaX;
        if (!delta) return;
        evt.preventDefault();
        primary.scrollLeft += delta;
        secondary.scrollLeft = primary.scrollLeft;
      };

      primary.addEventListener('wheel', wheelScroll, { passive: false });
      secondary.addEventListener('wheel', wheelScroll, { passive: false });
      primary.dataset.syncBound = 'true';
    }

    function renderUtilizationChart(rows) {
      const svg = els.utilizationSvg;
      if (!svg) return;
      while (svg.firstChild) svg.removeChild(svg.firstChild);

      const { dates, series } = buildUtilizationSeries(rows, utilizationMode);
      renderUtilizationLegend(series);
      if (els.utilizationChartTag) {
        els.utilizationChartTag.textContent = utilizationMode === 'daily' ? 'Daily' : 'Cumulative';
      }

      if (!dates.length || !series.length) {
        const empty = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        empty.setAttribute('x', '500');
        empty.setAttribute('y', '160');
        empty.setAttribute('text-anchor', 'middle');
        empty.setAttribute('fill', 'rgba(244,247,251,0.62)');
        empty.setAttribute('font-size', '14');
        empty.setAttribute('font-weight', '700');
        empty.textContent = 'No rig utilization data available';
        svg.appendChild(empty);
        return;
      }

      const rigCount = Math.max(series.length, 1);
      const layout = getUtilizationChartLayout(els.utilizationChartWrap, dates, rigCount, utilizationMode, 330);
      const { chartW, chartH, left, right, top, bottom, innerW, innerH, plotSidePad, stepX } = layout;
      const maxVal = Math.max(100, ...series.flatMap(item => item.points
        .map(point => (utilizationMode === 'daily' ? point.utilization : point.cumulativeUtilization))
        .filter(value => Number.isFinite(value))));
      const yFor = value => top + innerH - (Math.max(0, value) / maxVal) * innerH;

      svg.setAttribute('viewBox', `0 0 ${chartW} ${chartH}`);
      svg.style.width = `${chartW}px`;
      svg.style.height = `${chartH}px`;

      let tooltipEl = document.getElementById('utilizationTooltip');
      if (!tooltipEl) {
        tooltipEl = document.createElement('div');
        tooltipEl.id = 'utilizationTooltip';
        tooltipEl.className = 'chart-tooltip';
        tooltipEl.style.position = 'fixed';
        document.body.appendChild(tooltipEl);
      }

      let hoverGuide = document.getElementById('utilizationHoverGuide');
      if (!hoverGuide) {
        hoverGuide = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        hoverGuide.id = 'utilizationHoverGuide';
      }
      hoverGuide.setAttribute('y1', String(top));
      hoverGuide.setAttribute('y2', String(chartH - bottom));
      hoverGuide.setAttribute('stroke', 'rgba(255,255,255,0.38)');
      hoverGuide.setAttribute('stroke-width', '1.5');
      hoverGuide.setAttribute('stroke-dasharray', '5 6');
      hoverGuide.setAttribute('opacity', utilizationMode === 'cumulative' ? '1' : '0');
      hoverGuide.style.display = 'none';
      svg.appendChild(hoverGuide);

      const showTooltip = (evt, dateKey, points, guideX) => {
        tooltipEl.innerHTML = `
          <div class="tooltip-title">${formatDateFullLabel(dateKey)}</div>
          ${points.map(point => `
            <div class="tooltip-row">
              <span>${escapeHtml(point.rig)}</span>
              <strong>${Number.isFinite(utilizationMode === 'daily' ? point.utilization : point.cumulativeUtilization) ? `${(utilizationMode === 'daily' ? point.utilization : point.cumulativeUtilization).toFixed(1)}%` : '—'}</strong>
            </div>
            <div class="tooltip-row">
              <span>Drilling Hr</span>
              <strong>${point.drillingHours.toFixed(1)} hr</strong>
            </div>
          `).join('')}
        `;
        tooltipEl.classList.add('visible');
        const tooltipX = Math.min(window.innerWidth - 260, evt.clientX + 14);
        const tooltipY = Math.min(window.innerHeight - 220, evt.clientY + 14);
        tooltipEl.style.left = `${Math.max(8, tooltipX)}px`;
        tooltipEl.style.top = `${Math.max(8, tooltipY)}px`;
        if (utilizationMode === 'cumulative') {
          hoverGuide.setAttribute('x1', String(guideX));
          hoverGuide.setAttribute('x2', String(guideX));
          hoverGuide.style.display = 'block';
        }
      };

      const hideTooltip = () => {
        tooltipEl.classList.remove('visible');
        hoverGuide.style.display = 'none';
      };

      for (let i = 0; i <= 4; i += 1) {
        const value = (maxVal / 4) * i;
        const y = yFor(value);
        const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        line.setAttribute('x1', String(left));
        line.setAttribute('x2', String(chartW - right));
        line.setAttribute('y1', String(y));
        line.setAttribute('y2', String(y));
        line.setAttribute('stroke', 'rgba(255,255,255,0.08)');
        line.setAttribute('stroke-width', '1');
        svg.appendChild(line);

        const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        label.setAttribute('x', String(left - 10));
        label.setAttribute('y', String(y + 4));
        label.setAttribute('text-anchor', 'end');
        label.setAttribute('fill', 'rgba(244,247,251,0.68)');
        label.setAttribute('font-size', '12');
        label.setAttribute('font-weight', '700');
        label.textContent = `${value.toFixed(0)}%`;
        svg.appendChild(label);
      }

      dates.forEach((dateKey, idx) => {
        const x = dates.length === 1 ? left + innerW / 2 : left + plotSidePad + stepX * idx;

        if (utilizationMode === 'daily') {
          const separator = document.createElementNS('http://www.w3.org/2000/svg', 'line');
          separator.setAttribute('x1', String(x));
          separator.setAttribute('x2', String(x));
          separator.setAttribute('y1', String(top));
          separator.setAttribute('y2', String(chartH - bottom));
          separator.setAttribute('stroke', 'rgba(255,255,255,0.12)');
          separator.setAttribute('stroke-width', '1');
          separator.setAttribute('stroke-dasharray', '3 8');
          svg.appendChild(separator);
        }

        const dateLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        dateLabel.setAttribute('x', String(x));
        dateLabel.setAttribute('y', String(chartH - 18));
        dateLabel.setAttribute('text-anchor', 'middle');
        dateLabel.setAttribute('fill', 'rgba(244,247,251,0.72)');
        dateLabel.setAttribute('font-size', '12');
        dateLabel.setAttribute('font-weight', '700');
        dateLabel.textContent = formatShortDateLabel(dateKey);
        svg.appendChild(dateLabel);

        const hover = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        hover.setAttribute('x', String(x - Math.max(stepX / 2, 24)));
        hover.setAttribute('y', String(top));
        hover.setAttribute('width', String(Math.max(stepX, 48)));
        hover.setAttribute('height', String(innerH));
        hover.setAttribute('fill', 'transparent');
        hover.style.cursor = 'pointer';
        hover.addEventListener('mousemove', evt => showTooltip(evt, dateKey, series.map(item => item.points[idx]), x));
        hover.addEventListener('mouseleave', hideTooltip);
        svg.appendChild(hover);
      });

      if (utilizationMode === 'daily') {
        const clusterWidth = Math.max(stepX * 0.72, rigCount * 18 + (rigCount - 1) * 4);
        const barGap = 4;
        const barW = Math.max(10, Math.min(18, (clusterWidth - barGap * (rigCount - 1)) / rigCount));

        series.forEach((item, rigIdx) => {
          item.points.forEach((point, idx) => {
            const value = point.utilization;
            if (!Number.isFinite(value)) return;
            const baseX = dates.length === 1 ? left + innerW / 2 : left + plotSidePad + stepX * idx;
            const clusterStart = baseX - ((barW * rigCount + barGap * (rigCount - 1)) / 2);
            const x = clusterStart + rigIdx * (barW + barGap);
            const y = yFor(value);
            const bar = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
            bar.setAttribute('x', String(x));
            bar.setAttribute('y', String(y));
            bar.setAttribute('width', String(barW));
            bar.setAttribute('height', String(Math.max(0, chartH - bottom - y)));
            bar.setAttribute('rx', '7');
            bar.setAttribute('fill', item.color);
            bar.setAttribute('fill-opacity', '0.86');
            bar.setAttribute('stroke', 'rgba(255,255,255,0.26)');
            bar.setAttribute('stroke-width', '1');
            bar.style.cursor = 'pointer';
            bar.addEventListener('mousemove', evt => showTooltip(evt, point.date, series.map(seriesItem => seriesItem.points[idx]), baseX));
            bar.addEventListener('mouseleave', hideTooltip);
            svg.appendChild(bar);

            const valueLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
            valueLabel.setAttribute('x', String(x + barW / 2));
            valueLabel.setAttribute('y', String(Math.max(top + 14, y - 10)));
            valueLabel.setAttribute('text-anchor', 'middle');
            valueLabel.setAttribute('fill', 'rgba(244,247,251,0.92)');
            valueLabel.setAttribute('font-size', '11');
            valueLabel.setAttribute('font-weight', '800');
            valueLabel.textContent = `${value.toFixed(0)}%`;
            valueLabel.style.pointerEvents = 'none';
            svg.appendChild(valueLabel);
          });
        });
      } else {
        series.forEach(item => {
          const pathParts = [];
          let started = false;
          item.points.forEach((point, idx) => {
            const value = point.cumulativeUtilization;
            if (!Number.isFinite(value)) return;
            const x = dates.length === 1 ? left + innerW / 2 : left + plotSidePad + stepX * idx;
            const y = yFor(value);
            pathParts.push(`${started ? 'L' : 'M'} ${x} ${y}`);
            started = true;
          });
          if (!pathParts.length) return;
          const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
          path.setAttribute('d', pathParts.join(' '));
          path.setAttribute('fill', 'none');
          path.setAttribute('stroke', item.color);
          path.setAttribute('stroke-width', '3');
          path.setAttribute('stroke-linecap', 'round');
          path.setAttribute('stroke-linejoin', 'round');
          path.setAttribute('opacity', '0.96');
          svg.appendChild(path);

          item.points.forEach((point, idx) => {
            const value = point.cumulativeUtilization;
            if (!Number.isFinite(value)) return;
            const x = dates.length === 1 ? left + innerW / 2 : left + plotSidePad + stepX * idx;
            const y = yFor(value);
            const dot = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
            dot.setAttribute('cx', String(x));
            dot.setAttribute('cy', String(y));
            dot.setAttribute('r', '4.8');
            dot.setAttribute('fill', item.color);
            dot.setAttribute('stroke', 'rgba(255,255,255,0.9)');
            dot.setAttribute('stroke-width', '1.6');
            dot.style.cursor = 'pointer';
            dot.addEventListener('mousemove', evt => showTooltip(evt, point.date, series.map(seriesItem => seriesItem.points[idx]), x));
            dot.addEventListener('mouseleave', hideTooltip);
            svg.appendChild(dot);

            const valueLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
            valueLabel.setAttribute('x', String(x));
            valueLabel.setAttribute('y', String(Math.max(top + 14, y - 12)));
            valueLabel.setAttribute('text-anchor', 'middle');
            valueLabel.setAttribute('fill', 'rgba(244,247,251,0.92)');
            valueLabel.setAttribute('font-size', '11');
            valueLabel.setAttribute('font-weight', '800');
            valueLabel.textContent = `${value.toFixed(0)}%`;
            valueLabel.style.pointerEvents = 'none';
            svg.appendChild(valueLabel);
          });
        });
      }

      svg.onmouseleave = hideTooltip;

      if (els.utilizationChartWrap) {
        bindUtilizationChartScrolls();
        requestAnimationFrame(() => {
          const hasOverflow = els.utilizationChartWrap.scrollWidth > els.utilizationChartWrap.clientWidth + 4;
          els.utilizationChartWrap.scrollLeft = hasOverflow
            ? Math.max(0, els.utilizationChartWrap.scrollWidth - els.utilizationChartWrap.clientWidth)
            : 0;
          if (els.utilizationLmChartWrap) {
            els.utilizationLmChartWrap.scrollLeft = els.utilizationChartWrap.scrollLeft;
          }
        });
      }
    }

    function renderUtilizationLmChart(rows) {
      const svg = els.utilizationLmSvg;
      if (!svg) return;
      while (svg.firstChild) svg.removeChild(svg.firstChild);

      const { dates, series } = buildUtilizationSeries(rows, utilizationMode);
      if (els.utilizationLmChartTag) {
        els.utilizationLmChartTag.textContent = utilizationMode === 'daily' ? 'Daily' : 'Cumulative';
      }
      if (!dates.length || !series.length) {
        const empty = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        empty.setAttribute('x', '500');
        empty.setAttribute('y', '130');
        empty.setAttribute('text-anchor', 'middle');
        empty.setAttribute('fill', 'rgba(244,247,251,0.56)');
        empty.setAttribute('font-size', '12');
        empty.setAttribute('font-weight', '700');
        empty.textContent = 'No drilled Lm data available';
        svg.appendChild(empty);
        return;
      }

      const rigCount = Math.max(series.length, 1);
      const layout = getUtilizationChartLayout(els.utilizationLmChartWrap, dates, rigCount, utilizationMode, 250, {
        left: 62,
        right: 20,
        top: 16,
        bottom: 56,
        dailyPlotSidePad: 18,
        cumulativePlotSidePad: 12
      });
      const { chartW, chartH, left, right, top, bottom, innerW, innerH, plotSidePad, stepX } = layout;
      const maxVal = Math.max(1, ...series.flatMap(item => item.points.map(point => (
        utilizationMode === 'daily' ? point.drilledLm : point.cumulativeLm
      ))));
      const yFor = value => top + innerH - (Math.max(0, value) / maxVal) * innerH;

      svg.setAttribute('viewBox', `0 0 ${chartW} ${chartH}`);
      svg.style.width = `${chartW}px`;
      svg.style.height = `${chartH}px`;

      let tooltipEl = document.getElementById('utilizationLmTooltip');
      if (!tooltipEl) {
        tooltipEl = document.createElement('div');
        tooltipEl.id = 'utilizationLmTooltip';
        tooltipEl.className = 'chart-tooltip';
        tooltipEl.style.position = 'fixed';
        document.body.appendChild(tooltipEl);
      }

      let hoverGuide = document.getElementById('utilizationLmHoverGuide');
      if (!hoverGuide) {
        hoverGuide = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        hoverGuide.id = 'utilizationLmHoverGuide';
      }
      hoverGuide.setAttribute('y1', String(top));
      hoverGuide.setAttribute('y2', String(chartH - bottom));
      hoverGuide.setAttribute('stroke', 'rgba(255,255,255,0.28)');
      hoverGuide.setAttribute('stroke-width', '1.2');
      hoverGuide.setAttribute('stroke-dasharray', '5 6');
      hoverGuide.style.display = 'none';
      svg.appendChild(hoverGuide);

      const showTooltip = (evt, dateKey, points, guideX) => {
        tooltipEl.innerHTML = `
          <div class="tooltip-title">${formatDateFullLabel(dateKey)}</div>
          ${points.map(point => `
            <div class="tooltip-row">
              <span>${escapeHtml(point.rig)}</span>
              <strong>${(utilizationMode === 'daily' ? point.drilledLm : point.cumulativeLm).toFixed(1)} Lm</strong>
            </div>
          `).join('')}
        `;
        tooltipEl.classList.add('visible');
        const tooltipX = Math.min(window.innerWidth - 260, evt.clientX + 14);
        const tooltipY = Math.min(window.innerHeight - 220, evt.clientY + 14);
        tooltipEl.style.left = `${Math.max(8, tooltipX)}px`;
        tooltipEl.style.top = `${Math.max(8, tooltipY)}px`;
        hoverGuide.setAttribute('x1', String(guideX));
        hoverGuide.setAttribute('x2', String(guideX));
        hoverGuide.style.display = 'block';
      };

      const hideTooltip = () => {
        tooltipEl.classList.remove('visible');
        hoverGuide.style.display = 'none';
      };

      for (let i = 0; i <= 4; i += 1) {
        const value = (maxVal / 4) * i;
        const y = yFor(value);
        const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        line.setAttribute('x1', String(left));
        line.setAttribute('x2', String(chartW - right));
        line.setAttribute('y1', String(y));
        line.setAttribute('y2', String(y));
        line.setAttribute('stroke', 'rgba(255,255,255,0.07)');
        line.setAttribute('stroke-width', '1');
        svg.appendChild(line);

        const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        label.setAttribute('x', String(left - 8));
        label.setAttribute('y', String(y + 4));
        label.setAttribute('text-anchor', 'end');
        label.setAttribute('fill', 'rgba(244,247,251,0.62)');
        label.setAttribute('font-size', '11');
        label.setAttribute('font-weight', '700');
        label.textContent = `${value.toFixed(0)}`;
        svg.appendChild(label);
      }

      dates.forEach((dateKey, idx) => {
        const x = dates.length === 1 ? left + innerW / 2 : left + plotSidePad + stepX * idx;
        const dateLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        dateLabel.setAttribute('x', String(x));
        dateLabel.setAttribute('y', String(chartH - 16));
        dateLabel.setAttribute('text-anchor', 'middle');
        dateLabel.setAttribute('fill', 'rgba(244,247,251,0.68)');
        dateLabel.setAttribute('font-size', '11');
        dateLabel.setAttribute('font-weight', '700');
        dateLabel.textContent = formatShortDateLabel(dateKey);
        svg.appendChild(dateLabel);

        const hover = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        hover.setAttribute('x', String(x - Math.max(stepX / 2, 22)));
        hover.setAttribute('y', String(top));
        hover.setAttribute('width', String(Math.max(stepX, 44)));
        hover.setAttribute('height', String(innerH));
        hover.setAttribute('fill', 'transparent');
        hover.style.cursor = 'pointer';
        hover.addEventListener('mousemove', evt => showTooltip(evt, dateKey, series.map(item => item.points[idx]), x));
        hover.addEventListener('mouseleave', hideTooltip);
        svg.appendChild(hover);
      });

      if (utilizationMode === 'daily') {
        const clusterWidth = Math.max(stepX * 0.72, rigCount * 18 + (rigCount - 1) * 4);
        const barGap = 4;
        const barW = Math.max(10, Math.min(18, (clusterWidth - barGap * (rigCount - 1)) / rigCount));

        series.forEach((item, rigIdx) => {
          item.points.forEach((point, idx) => {
            const value = point.drilledLm;
            const baseX = dates.length === 1 ? left + innerW / 2 : left + plotSidePad + stepX * idx;
            const clusterStart = baseX - ((barW * rigCount + barGap * (rigCount - 1)) / 2);
            const x = clusterStart + rigIdx * (barW + barGap);
            const y = yFor(value);
            const bar = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
            bar.setAttribute('x', String(x));
            bar.setAttribute('y', String(y));
            bar.setAttribute('width', String(barW));
            bar.setAttribute('height', String(Math.max(0, chartH - bottom - y)));
            bar.setAttribute('rx', '7');
            bar.setAttribute('fill', item.color);
            bar.setAttribute('fill-opacity', '0.82');
            bar.setAttribute('stroke', 'rgba(255,255,255,0.2)');
            bar.setAttribute('stroke-width', '1');
            bar.style.cursor = 'pointer';
            bar.addEventListener('mousemove', evt => showTooltip(evt, point.date, series.map(seriesItem => seriesItem.points[idx]), baseX));
            bar.addEventListener('mouseleave', hideTooltip);
            svg.appendChild(bar);

            const valueLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
            valueLabel.setAttribute('x', String(x + barW / 2));
            valueLabel.setAttribute('y', String(Math.max(top + 12, y - 8)));
            valueLabel.setAttribute('text-anchor', 'middle');
            valueLabel.setAttribute('fill', 'rgba(244,247,251,0.9)');
            valueLabel.setAttribute('font-size', '10');
            valueLabel.setAttribute('font-weight', '800');
            valueLabel.textContent = `${value.toFixed(1)}`;
            valueLabel.style.pointerEvents = 'none';
            svg.appendChild(valueLabel);
          });
        });
      } else {
        series.forEach(item => {
          const pathParts = [];
          item.points.forEach((point, idx) => {
            const value = point.cumulativeLm;
            const x = dates.length === 1 ? left + innerW / 2 : left + plotSidePad + stepX * idx;
            const y = yFor(value);
            pathParts.push(`${idx === 0 ? 'M' : 'L'} ${x} ${y}`);
          });
          const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
          path.setAttribute('d', pathParts.join(' '));
          path.setAttribute('fill', 'none');
          path.setAttribute('stroke', item.color);
          path.setAttribute('stroke-width', '2.6');
          path.setAttribute('stroke-linecap', 'round');
          path.setAttribute('stroke-linejoin', 'round');
          path.setAttribute('opacity', '0.95');
          svg.appendChild(path);

          item.points.forEach((point, idx) => {
            const value = point.cumulativeLm;
            const x = dates.length === 1 ? left + innerW / 2 : left + plotSidePad + stepX * idx;
            const y = yFor(value);
            const dot = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
            dot.setAttribute('cx', String(x));
            dot.setAttribute('cy', String(y));
            dot.setAttribute('r', '4.2');
            dot.setAttribute('fill', item.color);
            dot.setAttribute('stroke', 'rgba(255,255,255,0.88)');
            dot.setAttribute('stroke-width', '1.4');
            dot.style.cursor = 'pointer';
            dot.addEventListener('mousemove', evt => showTooltip(evt, point.date, series.map(seriesItem => seriesItem.points[idx]), x));
            dot.addEventListener('mouseleave', hideTooltip);
            svg.appendChild(dot);
          });
        });
      }

      svg.onmouseleave = hideTooltip;

      if (els.utilizationLmChartWrap) {
        bindUtilizationChartScrolls();
        requestAnimationFrame(() => {
          els.utilizationLmChartWrap.scrollLeft = els.utilizationChartWrap?.scrollLeft || 0;
        });
      }
    }

    function bindUtilizationUnifiedScroll() {
      const wrap = els.utilizationChartWrap;
      if (!wrap || wrap.dataset.unifiedScrollBound) return;
      wrap.addEventListener('wheel', evt => {
        const canScrollX = wrap.scrollWidth > wrap.clientWidth + 4;
        if (!canScrollX) return;
        const delta = Math.abs(evt.deltaY) > Math.abs(evt.deltaX) ? evt.deltaY : evt.deltaX;
        if (!delta) return;
        evt.preventDefault();
        wrap.scrollLeft += delta;
      }, { passive: false });
      wrap.dataset.unifiedScrollBound = 'true';
    }

    function renderUtilizationCombinedChart(rows) {
      const svg = els.utilizationSvg;
      if (!svg) return;
      while (svg.firstChild) svg.removeChild(svg.firstChild);

      const { dates, series } = buildUtilizationSeries(rows, utilizationMode);
      renderUtilizationLegend(series);
      if (els.utilizationChartTag) {
        els.utilizationChartTag.textContent = utilizationMode === 'daily' ? 'Daily' : 'Cumulative';
      }

      if (!dates.length || !series.length) {
        const empty = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        empty.setAttribute('x', '500');
        empty.setAttribute('y', '250');
        empty.setAttribute('text-anchor', 'middle');
        empty.setAttribute('fill', 'rgba(244,247,251,0.62)');
        empty.setAttribute('font-size', '14');
        empty.setAttribute('font-weight', '700');
        empty.textContent = 'No rig performance data available';
        svg.appendChild(empty);
        return;
      }

      const rigCount = Math.max(series.length, 1);
      const layout = getUtilizationChartLayout(els.utilizationChartWrap, dates, rigCount, utilizationMode, 560, {
        left: 84,
        right: 20,
        top: 28,
        bottom: 72,
        minWrapWidth: 0,
        overflowThreshold: utilizationMode === 'daily' ? 8 : 10,
        minDailySpan: 50,
        dailyRigFactor: 13,
        dailyGapFactor: 3,
        dailyPlotSidePad: 28,
        cumulativeSpan: 60,
        cumulativePlotSidePad: 18
      });
      const { chartW, chartH, left, right, top, bottom, innerW, plotSidePad, stepX } = layout;
      const dividerY = 316;
      const topArea = { y1: top, y2: 280 };
      const bottomArea = { y1: dividerY + 30, y2: chartH - bottom };
      const topHeight = topArea.y2 - topArea.y1;
      const bottomHeight = bottomArea.y2 - bottomArea.y1;
      const utilMax = Math.max(100, ...series.flatMap(item => item.points.map(point => utilizationMode === 'daily' ? point.utilization : point.cumulativeUtilization).filter(Number.isFinite)));
      const lmMax = Math.max(1, ...series.flatMap(item => item.points.map(point => utilizationMode === 'daily' ? point.drilledLm : point.cumulativeLm).filter(Number.isFinite)));
      const yForUtil = value => topArea.y1 + topHeight - (Math.max(0, value) / utilMax) * topHeight;
      const yForLm = value => bottomArea.y1 + bottomHeight - (Math.max(0, value) / lmMax) * bottomHeight;
      if (els.utilizationUtilTicks) {
        els.utilizationUtilTicks.innerHTML = Array.from({ length: 5 }, (_, i) => {
          const value = (utilMax / 4) * (4 - i);
          const pct = (i / 4) * 100;
          return `<div class="utilization-axis-tick" style="top:${pct}%">${value.toFixed(0)}%</div>`;
        }).join('');
      }
      if (els.utilizationLmTicks) {
        els.utilizationLmTicks.innerHTML = Array.from({ length: 5 }, (_, i) => {
          const value = (lmMax / 4) * (4 - i);
          const pct = (i / 4) * 100;
          return `<div class="utilization-axis-tick" style="top:${pct}%">${value.toFixed(0)} m</div>`;
        }).join('');
      }

      svg.setAttribute('viewBox', `0 0 ${chartW} ${chartH}`);
      svg.style.width = `${chartW}px`;
      svg.style.height = `${chartH}px`;

      let tooltipEl = document.getElementById('utilizationTooltip');
      if (!tooltipEl) {
        tooltipEl = document.createElement('div');
        tooltipEl.id = 'utilizationTooltip';
        tooltipEl.className = 'chart-tooltip';
        tooltipEl.style.position = 'fixed';
        document.body.appendChild(tooltipEl);
      }

      let hoverGuide = document.getElementById('utilizationHoverGuide');
      if (!hoverGuide) {
        hoverGuide = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        hoverGuide.id = 'utilizationHoverGuide';
      }
      hoverGuide.setAttribute('y1', String(topArea.y1));
      hoverGuide.setAttribute('y2', String(bottomArea.y2));
      hoverGuide.setAttribute('stroke', 'rgba(255,255,255,0.32)');
      hoverGuide.setAttribute('stroke-width', '1.4');
      hoverGuide.setAttribute('stroke-dasharray', '5 6');
      hoverGuide.style.display = 'none';
      svg.appendChild(hoverGuide);

      const formatUtilValue = point => {
        const value = utilizationMode === 'daily' ? point.utilization : point.cumulativeUtilization;
        return Number.isFinite(value) && value > 0 ? `${value.toFixed(1)}%` : '-';
      };

      const formatLmValue = point => {
        const value = utilizationMode === 'daily' ? point.drilledLm : point.cumulativeLm;
        return Number.isFinite(value) && value > 0 ? `${value.toFixed(1)} m` : '-';
      };

      const formatHoursValue = point => {
        const value = utilizationMode === 'daily' ? point.drillingHours : point.cumulativeHours;
        return Number.isFinite(value) && value > 0 ? `${value.toFixed(1)} hr` : '-';
      };
      const pointsForIndex = idx => series.map(item => ({ ...item.points[idx], color: item.color }));

      const showTooltip = (evt, dateKey, points, guideX) => {
        tooltipEl.innerHTML = `
          <div class="tooltip-title">${formatDateFullLabel(dateKey)}</div>
          ${points.map(point => `
            <div class="tooltip-row"><span style="color:${point.color}; font-weight:800;">${escapeHtml(point.rig)}</span><strong style="color:${point.color};">${formatUtilValue(point)}</strong></div>
            <div class="tooltip-row"><span style="color:${point.color};">Drilled Lm</span><strong style="color:${point.color};">${formatLmValue(point)}</strong></div>
            <div class="tooltip-row"><span style="color:${point.color};">Drilling Hr</span><strong style="color:${point.color};">${formatHoursValue(point)}</strong></div>
          `).join('')}
        `;
        tooltipEl.classList.add('visible');
        const tooltipX = Math.min(window.innerWidth - 260, evt.clientX + 14);
        const tooltipY = Math.min(window.innerHeight - 220, evt.clientY + 14);
        tooltipEl.style.left = `${Math.max(8, tooltipX)}px`;
        tooltipEl.style.top = `${Math.max(8, tooltipY)}px`;
        hoverGuide.setAttribute('x1', String(guideX));
        hoverGuide.setAttribute('x2', String(guideX));
        hoverGuide.style.display = 'block';
      };

      const hideTooltip = () => {
        tooltipEl.classList.remove('visible');
        hoverGuide.style.display = 'none';
      };

      const addGrid = (maxValue, yFor, suffix, areaBottom, fontSize) => {
        for (let i = 0; i <= 4; i += 1) {
          const value = (maxValue / 4) * i;
          const y = yFor(value);
          const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
          line.setAttribute('x1', String(left));
          line.setAttribute('x2', String(chartW - right));
          line.setAttribute('y1', String(y));
          line.setAttribute('y2', String(y));
          line.setAttribute('stroke', 'rgba(255,255,255,0.08)');
          line.setAttribute('stroke-width', '1');
          svg.appendChild(line);

        }
      };

      addGrid(utilMax, yForUtil, '%', topArea.y2, 12);
      addGrid(lmMax, yForLm, '', bottomArea.y2, 11);

      const divider = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      divider.setAttribute('x1', String(left));
      divider.setAttribute('x2', String(chartW - right));
        divider.setAttribute('y1', String(dividerY));
        divider.setAttribute('y2', String(dividerY));
      divider.setAttribute('stroke', 'rgba(255,255,255,0.14)');
      divider.setAttribute('stroke-width', '1');
      divider.setAttribute('stroke-dasharray', '4 6');
      svg.appendChild(divider);

      dates.forEach((dateKey, idx) => {
        const x = dates.length === 1 ? left + innerW / 2 : left + plotSidePad + stepX * idx;

        if (utilizationMode === 'daily' && idx < dates.length - 1) {
          const separatorX = x + stepX / 2;
          const separator = document.createElementNS('http://www.w3.org/2000/svg', 'line');
          separator.setAttribute('x1', String(separatorX));
          separator.setAttribute('x2', String(separatorX));
          separator.setAttribute('y1', String(topArea.y1));
          separator.setAttribute('y2', String(bottomArea.y2));
          separator.setAttribute('stroke', 'rgba(255,255,255,0.18)');
          separator.setAttribute('stroke-width', '1');
          separator.setAttribute('stroke-dasharray', '4 8');
          svg.appendChild(separator);
        }

        const dateLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        dateLabel.setAttribute('x', String(x));
        dateLabel.setAttribute('y', String(chartH - 18));
        dateLabel.setAttribute('text-anchor', 'middle');
        dateLabel.setAttribute('fill', 'rgba(244,247,251,0.72)');
        dateLabel.setAttribute('font-size', '9');
        dateLabel.setAttribute('font-weight', '700');
        dateLabel.textContent = formatShortDateLabel(dateKey);
        svg.appendChild(dateLabel);

        const hover = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        hover.setAttribute('x', String(x - Math.max(stepX / 2, 24)));
        hover.setAttribute('y', String(topArea.y1));
        hover.setAttribute('width', String(Math.max(stepX, 48)));
        hover.setAttribute('height', String(bottomArea.y2 - topArea.y1));
        hover.setAttribute('fill', 'transparent');
        hover.style.cursor = 'pointer';
        hover.addEventListener('mousemove', evt => showTooltip(evt, dateKey, pointsForIndex(idx), x));
        hover.addEventListener('mouseleave', hideTooltip);
        svg.appendChild(hover);
      });

      if (utilizationMode === 'daily') {
        const clusterWidth = Math.max(stepX * 0.64, rigCount * 16 + (rigCount - 1) * 4);
        const barGap = 4;
        const barW = Math.max(8, Math.min(15, (clusterWidth - barGap * (rigCount - 1)) / rigCount));

        series.forEach((item, rigIdx) => {
          item.points.forEach((point, idx) => {
            const baseX = dates.length === 1 ? left + innerW / 2 : left + plotSidePad + stepX * idx;
            const clusterStart = baseX - ((barW * rigCount + barGap * (rigCount - 1)) / 2);
            const x = clusterStart + rigIdx * (barW + barGap);

            const utilValue = point.utilization;
            if (Number.isFinite(utilValue) && utilValue > 0) {
              const utilY = yForUtil(utilValue);
              const utilBar = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
              utilBar.setAttribute('x', String(x));
              utilBar.setAttribute('y', String(utilY));
              utilBar.setAttribute('width', String(barW));
              utilBar.setAttribute('height', String(Math.max(0, topArea.y2 - utilY)));
              utilBar.setAttribute('rx', '5');
              utilBar.setAttribute('fill', item.color);
              utilBar.setAttribute('fill-opacity', '0.86');
              utilBar.setAttribute('stroke', 'rgba(255,255,255,0.22)');
              utilBar.setAttribute('stroke-width', '1');
              utilBar.style.cursor = 'pointer';
              utilBar.addEventListener('mousemove', evt => showTooltip(evt, point.date, pointsForIndex(idx), baseX));
              utilBar.addEventListener('mouseleave', hideTooltip);
              svg.appendChild(utilBar);

              const utilLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
              utilLabel.setAttribute('x', String(x + barW / 2));
              utilLabel.setAttribute('y', String(Math.max(topArea.y1 + 14, utilY - 10)));
              utilLabel.setAttribute('text-anchor', 'middle');
              utilLabel.setAttribute('fill', 'rgba(244,247,251,0.92)');
              utilLabel.setAttribute('font-size', '9');
              utilLabel.setAttribute('font-weight', '800');
              utilLabel.textContent = `${utilValue.toFixed(0)}%`;
              utilLabel.style.pointerEvents = 'none';
              svg.appendChild(utilLabel);
            }

            const lmValue = point.drilledLm;
            if (Number.isFinite(lmValue) && lmValue > 0) {
              const lmY = yForLm(lmValue);
              const lmBar = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
              lmBar.setAttribute('x', String(x));
              lmBar.setAttribute('y', String(lmY));
              lmBar.setAttribute('width', String(barW));
              lmBar.setAttribute('height', String(Math.max(0, bottomArea.y2 - lmY)));
              lmBar.setAttribute('rx', '5');
              lmBar.setAttribute('fill', item.color);
              lmBar.setAttribute('fill-opacity', '0.78');
              lmBar.setAttribute('stroke', 'rgba(255,255,255,0.18)');
              lmBar.setAttribute('stroke-width', '1');
              lmBar.style.cursor = 'pointer';
              lmBar.addEventListener('mousemove', evt => showTooltip(evt, point.date, pointsForIndex(idx), baseX));
              lmBar.addEventListener('mouseleave', hideTooltip);
              svg.appendChild(lmBar);

              const lmLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
              lmLabel.setAttribute('x', String(x + barW / 2));
              lmLabel.setAttribute('y', String(Math.max(bottomArea.y1 + 12, lmY - 8)));
              lmLabel.setAttribute('text-anchor', 'middle');
              lmLabel.setAttribute('fill', 'rgba(244,247,251,0.88)');
              lmLabel.setAttribute('font-size', '8');
              lmLabel.setAttribute('font-weight', '800');
              lmLabel.textContent = `${Math.round(lmValue)}m`;
              lmLabel.style.pointerEvents = 'none';
              svg.appendChild(lmLabel);
            }
          });
        });
      } else {
        series.forEach(item => {
          const utilPathParts = [];
          const lmPathParts = [];
          let utilStarted = false;
          let lmStarted = false;

          item.points.forEach((point, idx) => {
            const x = dates.length === 1 ? left + innerW / 2 : left + plotSidePad + stepX * idx;
            const utilValue = point.cumulativeUtilization;
            if (Number.isFinite(utilValue) && utilValue > 0) {
              utilPathParts.push(`${utilStarted ? 'L' : 'M'} ${x} ${yForUtil(utilValue)}`);
              utilStarted = true;
            }
            lmPathParts.push(`${lmStarted ? 'L' : 'M'} ${x} ${yForLm(point.cumulativeLm)}`);
            lmStarted = true;
          });

          if (utilPathParts.length) {
            const utilPath = document.createElementNS('http://www.w3.org/2000/svg', 'path');
            utilPath.setAttribute('d', utilPathParts.join(' '));
            utilPath.setAttribute('fill', 'none');
            utilPath.setAttribute('stroke', item.color);
            utilPath.setAttribute('stroke-width', '3');
            utilPath.setAttribute('stroke-linecap', 'round');
            utilPath.setAttribute('stroke-linejoin', 'round');
            utilPath.setAttribute('opacity', '0.96');
            svg.appendChild(utilPath);
          }

          const lmPath = document.createElementNS('http://www.w3.org/2000/svg', 'path');
          lmPath.setAttribute('d', lmPathParts.join(' '));
          lmPath.setAttribute('fill', 'none');
          lmPath.setAttribute('stroke', item.color);
          lmPath.setAttribute('stroke-width', '2.6');
          lmPath.setAttribute('stroke-linecap', 'round');
          lmPath.setAttribute('stroke-linejoin', 'round');
          lmPath.setAttribute('opacity', '0.95');
          svg.appendChild(lmPath);

          item.points.forEach((point, idx) => {
            const x = dates.length === 1 ? left + innerW / 2 : left + plotSidePad + stepX * idx;
            const utilValue = point.cumulativeUtilization;
            if (Number.isFinite(utilValue) && utilValue > 0) {
              const utilY = yForUtil(utilValue);
              const utilDot = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
              utilDot.setAttribute('cx', String(x));
              utilDot.setAttribute('cy', String(utilY));
              utilDot.setAttribute('r', '4.8');
              utilDot.setAttribute('fill', item.color);
              utilDot.setAttribute('stroke', 'rgba(255,255,255,0.9)');
              utilDot.setAttribute('stroke-width', '1.6');
              utilDot.style.cursor = 'pointer';
              utilDot.addEventListener('mousemove', evt => showTooltip(evt, point.date, pointsForIndex(idx), x));
              utilDot.addEventListener('mouseleave', hideTooltip);
              svg.appendChild(utilDot);

              const utilLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
              utilLabel.setAttribute('x', String(x));
              utilLabel.setAttribute('y', String(Math.max(topArea.y1 + 14, utilY - 12)));
              utilLabel.setAttribute('text-anchor', 'middle');
              utilLabel.setAttribute('fill', 'rgba(244,247,251,0.92)');
              utilLabel.setAttribute('font-size', '9');
              utilLabel.setAttribute('font-weight', '800');
              utilLabel.textContent = `${utilValue.toFixed(0)}%`;
              utilLabel.style.pointerEvents = 'none';
              svg.appendChild(utilLabel);
            }

            if (Number.isFinite(point.cumulativeLm) && point.cumulativeLm > 0) {
              const lmY = yForLm(point.cumulativeLm);
              const lmDot = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
              lmDot.setAttribute('cx', String(x));
              lmDot.setAttribute('cy', String(lmY));
              lmDot.setAttribute('r', '4.2');
              lmDot.setAttribute('fill', item.color);
              lmDot.setAttribute('stroke', 'rgba(255,255,255,0.88)');
              lmDot.setAttribute('stroke-width', '1.4');
              lmDot.style.cursor = 'pointer';
              lmDot.addEventListener('mousemove', evt => showTooltip(evt, point.date, pointsForIndex(idx), x));
              lmDot.addEventListener('mouseleave', hideTooltip);
              svg.appendChild(lmDot);

              const lmLabel = document.createElementNS('http://www.w3.org/2000/svg', 'text');
              lmLabel.setAttribute('x', String(x));
              lmLabel.setAttribute('y', String(Math.max(bottomArea.y1 + 12, lmY - 10)));
              lmLabel.setAttribute('text-anchor', 'middle');
              lmLabel.setAttribute('fill', 'rgba(244,247,251,0.88)');
              lmLabel.setAttribute('font-size', '8');
              lmLabel.setAttribute('font-weight', '800');
              lmLabel.textContent = `${Math.round(point.cumulativeLm)}m`;
              lmLabel.style.pointerEvents = 'none';
              svg.appendChild(lmLabel);
            }
          });
        });
      }

      svg.onmouseleave = hideTooltip;

      if (els.utilizationChartWrap) {
        bindUtilizationUnifiedScroll();
        requestAnimationFrame(() => {
          const hasOverflow = els.utilizationChartWrap.scrollWidth > els.utilizationChartWrap.clientWidth + 4;
          els.utilizationChartWrap.scrollLeft = hasOverflow
            ? Math.max(0, els.utilizationChartWrap.scrollWidth - els.utilizationChartWrap.clientWidth)
            : 0;
        });
      }
    }

    function renderUtilizationPage(project) {
      const rows = getUtilizationRows(project);
      const dayGroups = getUtilizationDayGroups(rows);
      const { colors } = buildUtilizationSeries(rows, utilizationMode);

      if (els.utilizationTableBody) {
        if (!rows.length) {
          els.utilizationTableBody.innerHTML = '<tr><td colspan="7" class="utilization-empty">No utilization data available for current scope.</td></tr>';
        } else {
          els.utilizationTableBody.innerHTML = dayGroups.map(group => group.items.map((row, idx) => {
            const utilClass = row.utilization >= 75 ? 'utilization-high' : (row.utilization >= 45 ? 'utilization-mid' : 'utilization-low');
            const totalClass = group.totalUtilization >= 75 ? 'utilization-high' : (group.totalUtilization >= 45 ? 'utilization-mid' : 'utilization-low');
            const rigColor = colors[row.rig] || '#8ef0bf';
            return `
              <tr class="${idx === 0 ? 'utilization-group-start' : ''}">
                ${idx === 0 ? `<td rowspan="${group.items.length}" class="utilization-date-cell">${formatDateFullLabel(group.date)}</td>` : ''}
                <td>
                  <span class="utilization-rig-chip">
                    <span class="utilization-rig-dot" style="background:${rigColor};"></span>
                    ${escapeHtml(row.rig)}
                  </span>
                </td>
                <td class="num">${row.drillingHours.toFixed(1)}</td>
                <td class="num">${row.drilledLm.toFixed(1)}</td>
                <td class="num">${row.shiftHours.toFixed(1)}</td>
                <td class="num ${utilClass}">${row.utilization.toFixed(1)}%</td>
                ${idx === 0 ? `<td rowspan="${group.items.length}" class="num ${totalClass} utilization-total-cell">${group.totalUtilization.toFixed(1)}%</td>` : ''}
              </tr>
            `;
          }).join('')).join('');
        }
      }

      renderUtilizationCombinedChart(rows);
    }

    function renderManpowerPage(project) {
      const rows = getFilteredManpowerRows(project);
      const equipmentRows = getProjectEquipmentDateRows(project);

      if (els.manpowerTableBody) {
        if (!rows.length) {
          els.manpowerTableBody.innerHTML = '<tr><td colspan="11" class="manpower-empty">No manpower data available for current scope.</td></tr>';
        } else {
          els.manpowerTableBody.innerHTML = rows.map(row => `
            <tr>
              <td>${formatDateFullLabel(row.date)}</td>
              <td class="num">${row.pm}</td>
              <td class="num">${row.se}</td>
              <td class="num">${row.foreman}</td>
              <td class="num">${row.op}</td>
              <td class="num">${row.vb}</td>
              <td class="num">${row.rig}</td>
              <td class="num">${row.we}</td>
              <td class="num">${row.me}</td>
              <td class="num">${row.hl}</td>
              <td class="num total-col">${row.total}</td>
            </tr>
          `).join('');
        }
      }

      renderManpowerHistogram(rows);

      if (els.equipmentTableBody) {
        if (!equipmentRows.length) {
          els.equipmentTableBody.innerHTML = '<tr><td colspan="3" class="manpower-empty">No equipment timeline available for current scope.</td></tr>';
        } else {
          els.equipmentTableBody.innerHTML = equipmentRows.slice().reverse().map(row => {
            const fleetHtml = row.activeItems.length
              ? `<div class="equipment-fleet-list">${row.activeItems.map(item => `
                  <span class="equipment-chip">
                    <span class="equipment-chip-dot equipment-chip-dot-${normalizeText(item.ownership).toLowerCase()}"></span>
                    <span class="equipment-ownership equipment-ownership-${normalizeText(item.ownership).toLowerCase()}">${escapeHtml(item.ownership)}</span>
                    <span class="equipment-name">${escapeHtml(item.label)}</span>
                  </span>
                `).join('')}</div>`
              : '-';
            const summaryHtml = `
              <div class="equipment-summary-stack">
                <div class="equipment-summary-line"><span>Rigs</span><strong>${row.ownedRigs}</strong></div>
                <div class="equipment-summary-line"><span>Cranes</span><strong>${row.ownedCranes}</strong></div>
                <div class="equipment-summary-line"><span>Others</span><strong>${row.rentedEq}</strong></div>
              </div>
            `;
            return `
              <tr>
                <td>${formatDateFullLabel(row.date)}</td>
                <td class="equipment-fleet-cell">${fleetHtml}</td>
                <td class="equipment-summary-cell">${summaryHtml}</td>
              </tr>
            `;
          }).join('');
        }
      }

      renderEquipmentHistogram(equipmentRows);
    }

    async function signInWithDirectory(loginValue, passwordValue) {
      const matchedUser = findUserRecord(loginValue);
      if (!matchedUser) {
        throw new Error('User not found in the APFC users source');
      }
      if (passwordValue !== matchedUser.password) {
        throw new Error('Incorrect password');
      }
      return matchedUser;
    }

    function resetDashboardSessionUi() {
      currentUser = null;
      selectedProject = DEFAULT_PROJECT;
      selectedPlot = '';
      companyManpowerScopeMode = 'filtered';
      companyManpowerDesignationFilter = 'all';
      updateUserContextUi();
      setAuthLocked(true);
      toggleRequestAccessPanel(false);
      setRequestAccessError('');
      if (els.authRequestNameInput) els.authRequestNameInput.value = '';
      if (els.authRequestEmailInput) els.authRequestEmailInput.value = '';
    }

    async function submitSignIn() {
      const loginValue = normalizeText(els.authLoginInput?.value);
      const passwordValue = els.authPasswordInput?.value || '';
      if (!loginValue) {
        setAuthLocked(true, 'Enter your username or email.');
        els.authLoginInput?.focus();
        return;
      }
      if (!passwordValue) {
        setAuthLocked(true, 'Enter the temporary password.');
        els.authPasswordInput?.focus();
        return;
      }

      if (els.authSubmitBtn) {
        els.authSubmitBtn.disabled = true;
        els.authSubmitBtn.textContent = 'Signing In...';
      }

      try {
        const user = await signInWithDirectory(loginValue, passwordValue);
        applyUserSession(user);
        await loadDashboardData();
        updateTimelinePileList(selectedProject);
        syncTimelinePresetButtons();
        renderTimelinePage(selectedProject);
      } catch (err) {
        console.error(err);
        setAuthLocked(true, err.message || 'Unable to sign in.');
      } finally {
        if (els.authSubmitBtn) {
          els.authSubmitBtn.disabled = false;
          els.authSubmitBtn.textContent = 'Sign In';
        }
      }
    }

    async function restoreExistingSession() {
      const stored = getStoredAuthSession();
      if (!stored) {
        resetDashboardSessionUi();
        return false;
      }

      const matchedUser = usersDirectory.find(user => (
        normalizeLogin(user.username) === normalizeLogin(stored.username) &&
        normalizeText(user.project) === normalizeText(stored.project) &&
        normalizeText(user.plot) === normalizeText(stored.plot)
      ));

      if (!matchedUser) {
        resetDashboardSessionUi();
        setAuthLocked(true, 'Your saved session is no longer valid. Sign in again.');
        return false;
      }

      applyUserSession(matchedUser);
      return true;
    }

    async function initDashboard() {
      try {
        bindMapFrameSync();
        await loadUsersDirectory();
        const hasSession = await restoreExistingSession();

        if (hasSession) {
          await loadDashboardData();
        }

          if (els.projectSelector) {
            els.projectSelector.addEventListener('change', e => {
            if (!canSelectProjects()) return;
              selectedProject = e.target.value;
              timelineState.pile = 'all';
              persistScopedSession();
            updateTimelinePileList(selectedProject);
            syncTimelinePresetButtons();
            renderDashboard(selectedProject);
            if (activePage === 'production') renderProductionPage(selectedProject, true);
            if (activePage === 'utilization') renderUtilizationPage(selectedProject);
            if (activePage === 'manpower') renderManpowerPage(selectedProject);
            if (activePage === 'companymanpower') renderCompanyManpowerPage(selectedProject);
            if (activePage === 'timeline') renderTimelinePage(selectedProject, true);
            if (activePage === 'cost') renderCostPage(selectedProject, true);
            broadcastAuthContext();
          });
        }

        if (els.refreshDashboardBtn) {
          els.refreshDashboardBtn.addEventListener('click', refreshDashboardData);
        }

        els.authForm?.addEventListener('submit', evt => {
          evt.preventDefault();
          submitSignIn();
        });

        els.authRequestAccessBtn?.addEventListener('click', () => {
          const shouldOpen = !!els.authRequestPanel?.hidden;
          toggleRequestAccessPanel(shouldOpen);
          if (shouldOpen) setRequestAccessError('');
          if (!shouldOpen && els.authError) els.authError.textContent = '';
        });

        els.authRequestPanel?.addEventListener('click', evt => {
          if (evt.target === els.authRequestPanel) toggleRequestAccessPanel(false);
        });

        els.authRequestCancelBtn?.addEventListener('click', () => {
          toggleRequestAccessPanel(false);
          setRequestAccessError('');
          if (els.authError) els.authError.textContent = '';
        });

        els.authRequestSendBtn?.addEventListener('click', () => {
          submitAccessRequest();
        });

        els.signOutBtn?.addEventListener('click', () => {
          applyUserSession(null);
          setAuthLocked(true);
          broadcastAuthContext();
          if (els.authPasswordInput) els.authPasswordInput.value = '';
          if (els.authRequestNameInput) els.authRequestNameInput.value = '';
          if (els.authRequestEmailInput) els.authRequestEmailInput.value = '';
          setRequestAccessError('');
          toggleRequestAccessPanel(false);
          if (els.authLoginInput) {
            els.authLoginInput.value = '';
            els.authLoginInput.focus();
          }
        });

        els.overviewDateModeButtons.forEach(btn => btn.addEventListener('click', () => {
          overviewDateMode = btn.dataset.mode;
          els.overviewDateModeButtons.forEach(b => b.classList.toggle('active', b.dataset.mode === overviewDateMode));
          renderDashboard(selectedProject);
          if (activePage === 'production') renderProductionPage(selectedProject);
          if (activePage === 'utilization') renderUtilizationPage(selectedProject);
          broadcastAuthContext();
        }));

        const dateModeToggle = document.getElementById('overviewDateModeToggle');
        if (dateModeToggle) {
          dateModeToggle.style.display = (activePage === 'overview' || activePage === 'production' || activePage === 'utilization') ? 'inline-flex' : 'none';
        }

        els.granularityToggleButtons.forEach(btn => btn.addEventListener('click', () => setGranularity(btn.dataset.granularity)));
        document.getElementById('modeToggle').addEventListener('click', toggleMode);
        syncModeToggleUI();
        els.metricToggleButtons.forEach(btn => btn.addEventListener('click', () => setMetric(btn.dataset.metric)));

        els.navButtons.forEach(btn => {
          if (!btn.dataset.page) return;
          btn.addEventListener('click', () => setActivePage(btn.dataset.page));
        });

        els.prodToolButtons.forEach(btn => btn.addEventListener('click', () => {
          const key = btn.dataset.prodKey;
          const group = btn.dataset.group;
          prodState[key] = group;
          els.prodToolButtons
            .filter(b => b.dataset.prodKey === key)
            .forEach(b => b.classList.toggle('active', b.dataset.group === group));
          renderProductionMetricChart(selectedProject, key, true);
        }));

        els.utilizationModeButtons.forEach(btn => btn.addEventListener('click', () => {
          utilizationMode = btn.dataset.utilMode;
          els.utilizationModeButtons.forEach(b => b.classList.toggle('active', b.dataset.utilMode === utilizationMode));
          renderUtilizationPage(selectedProject);
        }));

        els.companyManpowerLayoutSelect?.addEventListener('change', () => {
          renderCompanyManpowerPage(selectedProject);
        });

        els.companyManpowerDesignationSelect?.addEventListener('change', () => {
          companyManpowerDesignationFilter = els.companyManpowerDesignationSelect.value || 'all';
          renderCompanyManpowerPage(selectedProject);
        });

        els.companyManpowerScopeButtons.forEach(btn => btn.addEventListener('click', () => {
          companyManpowerScopeMode = btn.dataset.companyScope || 'filtered';
          els.companyManpowerScopeButtons.forEach(button => button.classList.toggle('active', button === btn));
          renderCompanyManpowerPage(selectedProject);
        }));

        els.companyManpowerExportBtn?.addEventListener('click', () => {
          updateCompanyExportProjectOptions();
          toggleCompanyExportPanel(true);
        });

        els.companyExportCancelBtn?.addEventListener('click', () => {
          toggleCompanyExportPanel(false);
        });

        els.companyExportPanel?.addEventListener('click', evt => {
          if (evt.target === els.companyExportPanel) toggleCompanyExportPanel(false);
        });

        els.companyExportConfirmBtn?.addEventListener('click', () => {
          exportCompanyManpowerWorkbook(els.companyExportProjectSelect?.value || 'all');
          toggleCompanyExportPanel(false);
        });

        if (hasSession) {
          updateTimelinePileList(selectedProject);
          syncTimelinePresetButtons();
        }

        els.timelineApplyBtn?.addEventListener('click', () => {
          timelineState.start = els.timelineStartDate.value || '';
          timelineState.end = els.timelineEndDate.value || '';
          updateTimelinePileList(selectedProject);
          syncTimelinePresetButtons();
          renderTimelinePage(selectedProject);
        });

        els.timelineClearBtn?.addEventListener('click', () => {
          timelineState.start = '';
          timelineState.end = '';
          els.timelineStartDate.value = '';
          els.timelineEndDate.value = '';
          updateTimelinePileList(selectedProject);
          syncTimelinePresetButtons();
          renderTimelinePage(selectedProject);
        });

        els.timelinePreset7Btn?.addEventListener('click', () => applyTimelinePreset(7));
        els.timelinePreset14Btn?.addEventListener('click', () => applyTimelinePreset(14));
        els.timelinePreset30Btn?.addEventListener('click', () => applyTimelinePreset(30));

        els.timelineZoomInBtn?.addEventListener('click', () => {
          timelineState.zoom = Math.min(4, (timelineState.zoom || 1) * 1.2);
          renderTimelinePage(selectedProject);
        });

        els.timelineZoomOutBtn?.addEventListener('click', () => {
          timelineState.zoom = Math.max(0.5, (timelineState.zoom || 1) / 1.2);
          renderTimelinePage(selectedProject);
        });

        els.timelineZoomResetBtn?.addEventListener('click', () => {
          timelineState.zoom = 1;
          renderTimelinePage(selectedProject);
        });

        document.addEventListener('pointerdown', hideChartTooltipsOnOutsideInteraction);
        document.addEventListener('scroll', hideTooltip, true);
        window.addEventListener('resize', () => {
          if (!currentUser) return;
          renderDashboard(selectedProject);
        });

        if (!hasSession) {
          setAuthLocked(true);
          els.authLoginInput?.focus();
        }
      } catch (err) {
        console.error(err);
        setAuthLocked(true, err.message || 'Unable to initialize sign in.');
        els.dataSourceChip.textContent = 'Source Error';
      }
    }

    initDashboard();

/* ==========================================================================
   MOBILE UI — hamburger drawer + timeline filter toggle
   Self-contained; runs after the app initialises.
   ========================================================================== */
(function initMobileUI() {
  const menuBtn     = document.getElementById('mobileActionsBtn');
  const topActions  = document.querySelector('.top-actions');
  const filterBtn   = document.getElementById('timelineFilterToggle');
  const tlSidebar   = document.querySelector('.timeline-sidebar');
  let drawerOverlay = document.querySelector('.mobile-drawer-overlay');

  if (!drawerOverlay) {
    drawerOverlay = document.createElement('button');
    drawerOverlay.type = 'button';
    drawerOverlay.className = 'mobile-drawer-overlay';
    drawerOverlay.setAttribute('aria-hidden', 'true');
    document.body.appendChild(drawerOverlay);
  }

  function closeMobileActionsMenu() {
    if (!topActions || !menuBtn) return;
    topActions.classList.remove('mobile-open');
    menuBtn.setAttribute('aria-expanded', 'false');
    document.body.classList.remove('mobile-controls-open');
    drawerOverlay.classList.remove('visible');
  }

  /* -- Hamburger drawer --------------------------------------------------- */
  if (menuBtn && topActions) {
    menuBtn.addEventListener('click', function (e) {
      e.stopPropagation();
      const isOpen = topActions.classList.toggle('mobile-open');
      menuBtn.setAttribute('aria-expanded', String(isOpen));
      document.body.classList.toggle('mobile-controls-open', isOpen);
      drawerOverlay.classList.toggle('visible', isOpen);
    });

    drawerOverlay.addEventListener('click', closeMobileActionsMenu);

    /* Close only when tapping completely outside the topbar area */
    document.addEventListener('click', function (e) {
      if (!e.target.closest('header.topbar') && !e.target.closest('.mobile-drawer-overlay')) {
        closeMobileActionsMenu();
      }
    });

    /* Close after a button action — but NOT for select/toggle elements
       so the project selector and mode toggle remain fully usable        */
    topActions.addEventListener('click', function (e) {
      const tag = e.target.tagName;
      const isInteractive = tag === 'SELECT' || tag === 'INPUT' ||
                            e.target.closest('.mode-toggle') ||
                            e.target.closest('.mode-switch');
      if (!isInteractive) {
        closeMobileActionsMenu();
      }
    });

    window.addEventListener('resize', () => {
      if (window.innerWidth > 900) closeMobileActionsMenu();
    });
  }

  /* -- Timeline filter toggle --------------------------------------------- */
  if (filterBtn && tlSidebar) {
    filterBtn.addEventListener('click', function () {
      const isOpen = tlSidebar.classList.toggle('mobile-open');
      filterBtn.textContent = isOpen ? 'Hide Filters' : 'Show Filters';
      filterBtn.classList.toggle('is-open', isOpen);
    });
  }
})();

