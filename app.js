// Configuração do Google Sheets
const GOOGLE_APPS_SCRIPT_URL = "COLE_SUA_URL";

// Reactive State Management System
const createReactiveState = (initialState, renderCallback) => {
  return new Proxy(initialState, {
    set(target, property, value) {
      target[property] = value;
      if (property === 'periods') {
        localStorage.setItem('coenergy_data_v2', JSON.stringify(value));
      }
      if (property === 'usinas') {
        localStorage.setItem('coenergy_usinas_data_v2', JSON.stringify(value));
      }
      if (property === 'creditos') {
        localStorage.setItem('coenergy_creditos_data_v2', JSON.stringify(value));
      }
      setTimeout(renderCallback, 0);
      return true;
    }
  });
};

// Utility: Normalize string for search (ignore accents and case)
const normalizeString = (str) => {
  return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
};

// Application State
const state = createReactiveState({
  periods: [
    {
      id: "p-2026-0",
      name: "Janeiro de 2026",
      month: "Janeiro",
      year: 2026,
      records: [
        { id: "1", clientName: "Supermercado Alfa", dateInclusion: "2026-01-05", avgKw: 1450.5, status: "Ativos", tipoCliente: "Pessoa Jurídica", saldo: 2450.00, inadimplente: false, contasAtrasadas: [] },
        { id: "2", clientName: "Indústria Ômega", dateInclusion: "2026-01-12", avgKw: 850.0, status: "Em espera", tipoCliente: "Pessoa Jurídica", saldo: 0, inadimplente: true, contasAtrasadas: [] }
      ]
    },
    {
      id: "p-2026-1",
      name: "Fevereiro de 2026",
      month: "Fevereiro",
      year: 2026,
      records: [
        { id: "3", clientName: "Posto de Combustível Beta", dateInclusion: "2026-02-01", avgKw: 2100.2, status: "Ativos", tipoCliente: "Pessoa Jurídica", saldo: 3100.50, inadimplente: false, contasAtrasadas: [] },
        { id: "4", clientName: "Hotel Delta", dateInclusion: "2026-02-15", avgKw: 560.8, status: "Ativos", tipoCliente: "Pessoa Física", saldo: 140.20, inadimplente: false, contasAtrasadas: [] },
        { id: "5", clientName: "Escola Epsilon", dateInclusion: "2026-02-28", avgKw: 1200.0, status: "Saíram", tipoCliente: "Pessoa Jurídica", saldo: 0, inadimplente: false, contasAtrasadas: [] }
      ]
    }
  ],
  usinas: [
    { 
      id: "u-1", 
      name: "Usina Solar Norte",
      records: [
        { id: "r-1", month: "2026-01", expected: 1500, comp: 1450 },
        { id: "r-2", month: "2026-02", expected: 1600, comp: 1580 }
      ]
    },
    { id: "u-2", name: "Usina Solar Sul", records: [] }
  ],
  creditos: [
    { id: "c-1", date: "2026-01", kwTotal: 5000, historico: "Saldo Inicial" },
    { id: "c-2", date: "2026-02", kwTotal: 4800, historico: "Ajuste mensal" }
  ],
  activePeriodId: "p-2026-1",
  activeView: 'dashboard-view',
  isCloudSynced: false,
  inlineEditingRecordId: null,
  inlineEditingPeriodId: null,
  inlineEditingCreditoId: null
}, () => {
  updateDashboardUI();
  updateDataEntryUI();
  updateUsinasUI();
  updateOrganizacaoUI();
  updateCreditosUI();
  updateInadimplenciaUI();
});

// Chart instances
let barChartInstance = null;
let lineChartInstance = null;
let statusChartInstance = null;
let saidasChartInstance = null;
let newClientsChartInstance = null;
let inadimplenciaChartInstance = null;
let tipoClienteChartInstance = null;
let tipoClienteKwChartInstance = null;
let creditosChartInstance = null;
window.usinaChartInstances = {};
window.usinasInEditMode = new Set();
Chart.register(ChartDataLabels);

// Initialize App
document.addEventListener('DOMContentLoaded', () => {
  setupNavigation();
  setupModals();
  setupForms();
  setupLogoUpload();
  setupInteractiveCards();
  setupLoteEntry();
  setupOrganizacaoFilters();
  setupGlobalSearch();
  setupCreditosForm();
  setupInadimplenciaForm();

  // Se já tiver dados no localStorage, renderiza imediatamente a tela
  const localData = localStorage.getItem('coenergy_data_v2');
  if (localData) {
    state.periods = JSON.parse(localData);
  }
  const localUsinas = localStorage.getItem('coenergy_usinas_data_v2');
  if (localUsinas) {
    state.usinas = JSON.parse(localUsinas);
  }
  const localCreditos = localStorage.getItem('coenergy_creditos_data_v2');
  if (localCreditos) {
    state.creditos = JSON.parse(localCreditos);
  }

  if (state.periods.length > 0 || state.usinas.length > 0) {
    updateDashboardUI();
    updateDataEntryUI();
    updateUsinasUI();
    updateInadimplenciaUI();
  }

  // Tentar buscar do Google Sheets (Sincronização Nuvem Inicial)
  showLoader();
  fetchGoogleSheetsData(true).then(() => {
    // Se a URL não foi configurada e não temos dados, carregamos dummy
    if (state.periods.length === 0 && GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
      loadDemoData();
    }
    hideLoader();
  });

  // Polling a cada 5 segundos para reatividade sem refresh (Google Sheets)
  setInterval(() => fetchGoogleSheetsData(false), 5000);
});

// --- GOOGLE SHEETS INTEGRATION ---
function showLoader() {
  const loader = document.getElementById('global-loader');
  if (loader) loader.style.display = 'flex';
}

function hideLoader() {
  const loader = document.getElementById('global-loader');
  if (loader) loader.style.display = 'none';
}

async function fetchGoogleSheetsData(isInitial = false) {
  if (GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) return;

  try {
    const response = await fetch(GOOGLE_APPS_SCRIPT_URL);
    const data = await response.json();
    processSheetsData(data);
    state.isCloudSynced = true;
  } catch (error) {
    console.error("Erro na leitura do Google Sheets:", error);
  }
}

function processSheetsData(fetchedRows) {
  const periodMap = new Map();
  const monthNames = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

  fetchedRows.forEach(row => {
    if (!row.dateInclusion) return;

    // Pode vir como YYYY-MM-DD (do form/app script) ou DD/MM/YYYY (texto no Google Sheets)
    let year, monthIdx;

    // Tenta formato YYYY-MM-DD
    let dateParts = row.dateInclusion.split('-');
    if (dateParts.length === 3) {
      year = parseInt(dateParts[0]);
      monthIdx = parseInt(dateParts[1]) - 1;
    } else {
      // Tenta formato DD/MM/YYYY
      dateParts = row.dateInclusion.split('/');
      if (dateParts.length === 3) {
        year = parseInt(dateParts[2]);
        monthIdx = parseInt(dateParts[1]) - 1;
      } else {
        return; // formato inválido
      }
    }

    const pName = `${monthNames[monthIdx]} de ${year}`;
    const pId = `p-${year}-${monthIdx}`;

    if (!periodMap.has(pId)) {
      periodMap.set(pId, { id: pId, name: pName, month: monthNames[monthIdx], year, records: [] });
    }

    periodMap.get(pId).records.push({
      id: row.id,
      clientName: row.clientName,
      dateInclusion: row.dateInclusion,
      avgKw: parseFloat(row.avgKw) || 0,
      status: row.status,
      tipoCliente: row.tipoCliente || 'Pessoa Física',
      saldo: parseFloat(row.saldo) || 0,
      dateSaida: row.dateSaida || '',
      inadimplente: row.inadimplente || false,
      contasAtrasadas: row.contasAtrasadas || []
    });
  });

  const newPeriods = Array.from(periodMap.values()).sort((a, b) => {
    if (a.year !== b.year) return a.year - b.year;
    return monthNames.indexOf(a.month) - monthNames.indexOf(b.month);
  });

  // If local empty periods exist (user just created tab but added no row), merge them
  state.periods.forEach(localP => {
    if (localP.records.length === 0 && !newPeriods.find(np => np.id === localP.id)) {
      newPeriods.push(localP);
    }
  });

  // Maintain active period
  const oldActive = state.activePeriodId;
  state.periods = newPeriods;

  if (newPeriods.length > 0 && !oldActive) {
    state.activePeriodId = newPeriods[newPeriods.length - 1].id;
  }
}

async function postToGoogleSheets(actionData) {
  if (GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) return;
  showLoader();
  try {
    const payload = JSON.stringify(actionData);
    await fetch(GOOGLE_APPS_SCRIPT_URL, {
      method: "POST",
      body: payload
    });
  } catch (err) {
    console.error("Erro na mutação Cloud:", err);
  } finally {
    hideLoader();
  }
}

// Logo Upload Logic
function setupLogoUpload() {
  const logoUpload = document.getElementById('logo-upload');
  const companyLogo = document.getElementById('company-logo');

  // Load persistent logo on initialization
  const savedLogo = localStorage.getItem('coenergy_custom_logo');
  if (savedLogo) {
    companyLogo.src = savedLogo;
  }

  logoUpload.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (event) => {
        const result = event.target.result;
        companyLogo.src = result;

        // Parse and save to localStorage
        localStorage.setItem('coenergy_custom_logo', result);

        // Reset file input for clean UX
        logoUpload.value = '';
      };
      reader.readAsDataURL(file);
    }
  });
}

function setupNavigation() {
  const navItems = document.querySelectorAll('.nav-item');
  const views = document.querySelectorAll('.view-section');

  navItems.forEach(item => {
    item.addEventListener('click', (e) => {
      navItems.forEach(n => n.classList.remove('active'));
      const btn = e.currentTarget;
      btn.classList.add('active');

      const targetId = btn.getAttribute('data-target');
      views.forEach(v => {
        v.classList.remove('active');
        if (v.id === targetId) v.classList.add('active');
      });

      state.activeView = targetId;

      if (targetId === 'dashboard-view') {
        if (barChartInstance) barChartInstance.resize();
        if (lineChartInstance) lineChartInstance.resize();
        if (statusChartInstance) statusChartInstance.resize();
        if (saidasChartInstance) saidasChartInstance.resize();
        if (newClientsChartInstance) newClientsChartInstance.resize();
        if (inadimplenciaChartInstance) inadimplenciaChartInstance.resize();
        if (tipoClienteChartInstance) tipoClienteChartInstance.resize();
        if (tipoClienteKwChartInstance) tipoClienteKwChartInstance.resize();
      }
      if (targetId === 'creditos-view') {
        if (creditosChartInstance) creditosChartInstance.resize();
      }
    });
  });
}

function setupModals() {
  const modal = document.getElementById('period-modal');
  const addBtn = document.getElementById('add-period-btn');
  const closeBtn = document.getElementById('close-modal-btn');

  addBtn.addEventListener('click', () => {
    modal.classList.add('active');
  });

  closeBtn.addEventListener('click', () => {
    modal.classList.remove('active');
  });

  window.addEventListener('click', (e) => {
    if (e.target === modal) {
      modal.classList.remove('active');
    }
  });
}

function setupGlobalSearch() {
  const searchInput = document.getElementById('global-client-search');
  if (searchInput) {
    searchInput.addEventListener('input', () => {
      updateDataEntryUI();
    });
  }
}

window.loteParsedData = [];

function setupLoteEntry() {
  const textarea = document.getElementById('lote-textarea');
  const previewArea = document.getElementById('lote-preview-area');
  const inputArea = document.getElementById('lote-input-area');
  const tbody = document.getElementById('lote-preview-tbody');
  const countBadge = document.getElementById('lote-count-badge');
  const btnClear = document.getElementById('btn-clear-lote');
  const btnSave = document.getElementById('btn-save-lote');
  const successMsg = document.getElementById('lote-success-msg');

  if (!textarea) return;

  textarea.addEventListener('input', (e) => {
    const text = e.target.value.trim();
    if (!text) return;

    const rows = text.split('\n');
    window.loteParsedData = [];

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i].trim();
      if (!row) continue;

      const cols = row.split('\t');
      if (cols.length >= 6) {
        window.loteParsedData.push({
          clientName: cols[0].trim(),
          avgKw: parseFloat(cols[1].replace(/\./g, '').replace(',', '.').trim()) || 0,
          saldo: parseFloat((cols[2] || '0').replace(/\./g, '').replace(',', '.').trim()) || 0,
          dateInclusion: cols[3].trim(), // YYYY-MM-DD or DD/MM/YYYY
          status: cols[4].trim(),
          tipoCliente: cols[5].trim(),
          inadimplente: false,
          dateSaida: ''
        });
      }
    }

    if (window.loteParsedData.length > 0) {
      renderLotePreview();
      inputArea.style.display = 'none';
      previewArea.style.display = 'block';
    } else {
      alert("Não foi possível identificar os dados. Verifique se copiou as colunas corretamente (separadas por tabulação).");
      textarea.value = '';
    }
  });

  btnClear.addEventListener('click', () => {
    textarea.value = '';
    window.loteParsedData = [];
    inputArea.style.display = 'block';
    previewArea.style.display = 'none';
    successMsg.style.display = 'none';
    btnSave.disabled = false;
  });

  btnSave.addEventListener('click', async () => {
    if (window.loteParsedData.length === 0) return;

    const monthNames = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
    const tempState = [...state.periods];
    const timestamp = Date.now();

    window.loteParsedData.forEach((record, index) => {
      let year, monthIdx;
      if (record.dateInclusion.includes('/')) {
        const parts = record.dateInclusion.split('/');
        year = parseInt(parts[2]);
        monthIdx = parseInt(parts[1]) - 1;
        record.dateInclusionFormatted = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
      } else {
        const parts = record.dateInclusion.split('-');
        year = parseInt(parts[0]);
        monthIdx = parseInt(parts[1]) - 1;
        record.dateInclusionFormatted = record.dateInclusion;
      }

      const month = monthNames[monthIdx];
      const pName = `${month} de ${year}`;
      const pId = `p-${year}-${monthIdx}`;

      let periodIndex = tempState.findIndex(p => p.id === pId);
      if (periodIndex === -1) {
        tempState.push({ id: pId, name: pName, month: month, year: year, records: [] });
        tempState.sort((a, b) => {
          if (a.year !== b.year) return a.year - b.year;
          return monthNames.indexOf(a.month) - monthNames.indexOf(b.month);
        });
        periodIndex = tempState.findIndex(p => p.id === pId);
      }

      record.id = `r-${timestamp}-${index}`;

        tempState[periodIndex].records.push({
          id: record.id,
          clientName: record.clientName,
          dateInclusion: record.dateInclusionFormatted,
          avgKw: record.avgKw,
          status: record.status,
          tipoCliente: record.tipoCliente,
          saldo: record.saldo,
          dateSaida: record.dateSaida,
          inadimplente: record.inadimplente,
          contasAtrasadas: []
        });
    });

    state.periods = tempState;

    if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
      btnSave.disabled = true;
      document.getElementById('btn-save-lote-text').textContent = "Salvando...";

      const payload = {
        action: 'add_lote',
        clients: window.loteParsedData
      };

      await postToGoogleSheets(payload);
      await fetchGoogleSheetsData(false);

      btnSave.disabled = false;
      document.getElementById('btn-save-lote-text').textContent = "Salvar Lote";
    }

    successMsg.style.display = 'block';
    setTimeout(() => {
      btnClear.click();
    }, 3000);
  });

  function renderLotePreview() {
    tbody.innerHTML = '';
    countBadge.textContent = `${window.loteParsedData.length} registros identificados`;

    window.loteParsedData.forEach(record => {
      const tr = document.createElement('tr');
      let displayDate = record.dateInclusion;
      if (displayDate.includes('-')) {
        const dparts = displayDate.split('-');
        if (dparts.length === 3) displayDate = `${dparts[2]}/${dparts[1]}/${dparts[0]}`;
      }

      let badgeClass = 'ativos';
      if (record.status === 'Em espera') badgeClass = 'espera';
      if (record.status === 'Saíram') badgeClass = 'sairam';

      tr.innerHTML = `
        <td><strong>${record.clientName}</strong></td>
        <td><strong>${record.avgKw.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}</strong> kW</td>
        <td><strong>${record.saldo.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}</strong> kW</td>
        <td>${displayDate}</td>
        <td><span class="status-badge ${badgeClass}" style="font-size: 0.85rem; padding: 4px 10px;">${record.status}</span></td>
        <td><span style="font-size: 0.95rem; color: var(--text-muted);">${record.tipoCliente}</span></td>
      `;
      tbody.appendChild(tr);
    });
  }
}

function setupForms() {
  // New Period Form
  const periodForm = document.getElementById('period-form');
  periodForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const monthNames = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
    const month = document.getElementById('periodMonth').value;
    const year = parseInt(document.getElementById('periodYear').value);

    const periodName = `${month} de ${year}`;
    const exists = state.periods.find(p => p.name === periodName);
    if (exists) {
      alert('Esta aba de período já existe!');
      return;
    }

    const monthIdx = monthNames.indexOf(month);
    const newId = `p-${year}-${monthIdx}`;

    const newPeriod = {
      id: newId,
      name: periodName,
      month: month,
      year: year,
      records: []
    };

    state.periods = [...state.periods, newPeriod];
    state.activePeriodId = newPeriod.id;

    document.getElementById('period-modal').classList.remove('active');
    periodForm.reset();
    document.getElementById('periodYear').value = new Date().getFullYear();
  });

  // Toggle Data de Saída
  const clientStatusSelect = document.getElementById('clientStatus');
  const dataSaidaGroup = document.getElementById('dataSaidaGroup');
  const dateSaidaInput = document.getElementById('dateSaida');
  if (clientStatusSelect) {
    clientStatusSelect.addEventListener('change', (e) => {
      if (e.target.value === 'Saíram') {
        dataSaidaGroup.style.display = 'block';
        if (dateSaidaInput) dateSaidaInput.required = true;
      } else {
        dataSaidaGroup.style.display = 'none';
        if (dateSaidaInput) {
          dateSaidaInput.required = false;
          dateSaidaInput.value = '';
        }
      }
    });
  }

  // New Record Form
  const recordForm = document.getElementById('record-form');
  recordForm.addEventListener('submit', async (e) => {
    e.preventDefault();

    const clientName = document.getElementById('clientName').value;
    const dateInclusion = document.getElementById('dateInclusion').value;
    const avgKw = parseFloat(document.getElementById('avgKw').value);
    const clientStatus = document.getElementById('clientStatus').value;
    const tipoCliente = document.getElementById('tipoCliente').value;
    const clientSaldo = parseFloat(document.getElementById('clientSaldo').value) || 0;
    const dateSaida = document.getElementById('dateSaida') ? document.getElementById('dateSaida').value : '';

    const dateParts = dateInclusion.split('-');
    const year = parseInt(dateParts[0]);
    const monthIdx = parseInt(dateParts[1]) - 1;
    const monthNames = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
    const month = monthNames[monthIdx];

    const pName = `${month} de ${year}`;
    const pId = `p-${year}-${monthIdx}`;

    const tempState = [...state.periods];
    let periodIndex = tempState.findIndex(p => p.id === pId);

    if (periodIndex === -1) {
      tempState.push({ id: pId, name: pName, month: month, year: year, records: [] });
      tempState.sort((a, b) => {
        if (a.year !== b.year) return a.year - b.year;
        return monthNames.indexOf(a.month) - monthNames.indexOf(b.month);
      });
      periodIndex = tempState.findIndex(p => p.id === pId);
    }

    let isEditing = window.editingRecordId !== null;
    let finalRecordId = isEditing ? window.editingRecordId : `r-${Date.now()}`;
    let previousInadimplente = false;

    if (isEditing && window.editingPeriodId) {
      const oldPeriodIdx = tempState.findIndex(p => p.id === window.editingPeriodId);
      if (oldPeriodIdx !== -1) {
        const oldRec = tempState[oldPeriodIdx].records.find(r => r.id === finalRecordId);
        if (oldRec) previousInadimplente = oldRec.inadimplente;
        tempState[oldPeriodIdx].records = tempState[oldPeriodIdx].records.filter(r => r.id !== finalRecordId);
      }
    }

    const newRecord = {
      id: finalRecordId,
      clientName,
      dateInclusion,
      avgKw,
      status: clientStatus,
      tipoCliente,
      dateSaida,
      inadimplente: previousInadimplente,
      contasAtrasadas: []
    };

    tempState[periodIndex].records.push(newRecord);
    state.periods = tempState;
    state.activePeriodId = pId;

    if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
      const payload = {
        action: isEditing ? 'edit_client' : 'add',
        id: finalRecordId,
        clientName,
        dateInclusion,
        avgKw,
        status: clientStatus,
        tipoCliente,
        saldo: clientSaldo,
        dateSaida,
        inadimplente: previousInadimplente
      };
      await postToGoogleSheets(payload);
      if (!isEditing) await fetchGoogleSheetsData(false);
    }

    const successMsg = document.getElementById('form-success-msg');
    const successTxt = document.getElementById('form-success-text');
    if (successMsg && successTxt) {
      successTxt.textContent = isEditing ? "Cliente atualizado com sucesso!" : "Cliente adicionado com sucesso!";
      successMsg.style.display = 'inline-block';
      setTimeout(() => { successMsg.style.display = 'none'; }, 4000);
    }

    if (isEditing) {
      window.cancelEditMode();
    } else {
      recordForm.reset();
      document.getElementById('clientStatus').value = "Ativos";
      document.getElementById('tipoCliente').value = "Pessoa Física";
      document.getElementById('dataSaidaGroup').style.display = 'none';
      if (document.getElementById('dateSaida')) document.getElementById('dateSaida').required = false;
    }
  });

  // Usinas Form
  const usinaForm = document.getElementById('usina-form');
  if (usinaForm) {
    usinaForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      const usinaName = document.getElementById('usinaName').value.trim();
      const usinaMonth = document.getElementById('usinaMonth').value; // Format: YYYY-MM
      const usinaEsperada = parseFloat(document.getElementById('usinaEsperada').value);
      const usinaCompensada = parseFloat(document.getElementById('usinaCompensada').value);

      const newRecordId = `u-${Date.now()}`;
      const newRecord = {
        id: newRecordId,
        month: usinaMonth,
        esperada: usinaEsperada,
        compensada: usinaCompensada
      };

      const tempState = [...state.usinas];
      let usinaIndex = tempState.findIndex(u => u.name.toLowerCase() === usinaName.toLowerCase());

      if (usinaIndex === -1) {
        tempState.push({
          name: usinaName,
          records: []
        });
        usinaIndex = tempState.length - 1;
      }

      const existingRecordIndex = tempState[usinaIndex].records.findIndex(r => r.month === usinaMonth);
      if (existingRecordIndex !== -1) {
        tempState[usinaIndex].records[existingRecordIndex] = newRecord;
      } else {
        tempState[usinaIndex].records.push(newRecord);
      }

      tempState[usinaIndex].records.sort((a, b) => a.month.localeCompare(b.month));
      state.usinas = tempState;

      if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
        await postToGoogleSheets({
          action: 'add_usina',
          id: newRecordId,
          usinaName,
          month: usinaMonth,
          esperada: usinaEsperada,
          compensada: usinaCompensada
        });
      }

      const usinaSuccessMsg = document.getElementById('usina-success-msg');
      if (usinaSuccessMsg) {
        usinaSuccessMsg.style.display = 'inline-block';
        setTimeout(() => { usinaSuccessMsg.style.display = 'none'; }, 4000);
      }

      usinaForm.reset();
    });
  }
}

window.startInlineEdit = function (periodId, recordId) {
  state.inlineEditingRecordId = recordId;
  state.inlineEditingPeriodId = periodId;
};

window.cancelInlineEdit = function () {
  state.inlineEditingRecordId = null;
  state.inlineEditingPeriodId = null;
};

window.saveInlineEdit = async function (periodId, recordId) {
  const row = document.querySelector(`tr[data-record-id="${recordId}"]`);
  if (!row) return;

  const newName = row.querySelector('.edit-name').value;
  const newAvgKw = parseFloat(row.querySelector('.edit-kw').value);
  const newDate = row.querySelector('.edit-date').value;
  const newStatus = row.querySelector('.edit-status').value;
  const newTipo = row.querySelector('.edit-tipo').value;
  const newSaldo = parseFloat(row.querySelector('.edit-saldo').value) || 0;

  const tempState = [...state.periods];
  const periodIndex = tempState.findIndex(p => p.id === periodId);
  if (periodIndex !== -1) {
    const recordIndex = tempState[periodIndex].records.findIndex(r => r.id === recordId);
    if (recordIndex !== -1) {
      const record = tempState[periodIndex].records[recordIndex];
      record.clientName = newName;
      record.avgKw = newAvgKw;
      record.dateInclusion = newDate;
      record.status = newStatus;
      record.tipoCliente = newTipo;
      record.saldo = newSaldo;

      // Handle Data de Saída logic if status changed
      if (newStatus !== 'Saíram') {
        record.dateSaida = '';
      }

      state.periods = tempState;

      // Update backend
      if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
        await postToGoogleSheets({
          action: 'edit_client',
          id: recordId,
          clientName: newName,
          dateInclusion: newDate,
          avgKw: newAvgKw,
          status: newStatus,
          tipoCliente: newTipo,
          saldo: newSaldo,
          dateSaida: record.dateSaida,
          inadimplente: record.inadimplente
        });
      }
    }
  }
  window.cancelInlineEdit();
};

window.deleteRecord = async function (periodId, recordId) {
  // Optmistic Update UI
  const tempState = [...state.periods];
  const periodIndex = tempState.findIndex(p => p.id === periodId);
  if (periodIndex !== -1) {
    tempState[periodIndex].records = tempState[periodIndex].records.filter(r => r.id !== recordId);
    state.periods = tempState;
  }

  // Delete from Cloud
  if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
    await postToGoogleSheets({ action: 'delete', id: recordId });
    fetchGoogleSheetsData();
  }
};

window.toggleInadimplente = async function (periodId, recordId) {
  const tempState = [...state.periods];
  const periodIndex = tempState.findIndex(p => p.id === periodId);
  if (periodIndex !== -1) {
    const recordIndex = tempState[periodIndex].records.findIndex(r => r.id == recordId);
    if (recordIndex !== -1) {
      tempState[periodIndex].records[recordIndex].inadimplente = !tempState[periodIndex].records[recordIndex].inadimplente;
      state.periods = tempState;
    }
  }

  if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
    await postToGoogleSheets({ action: 'toggle_inadimplente', id: recordId });
    fetchGoogleSheetsData();
  }
};

window.deleteCurrentMonth = async function () {
  const periodId = state.activePeriodId;
  if (!periodId) return;
  const p = state.periods.find(x => x.id === periodId);
  if (!p) return;

  if (confirm(`Tem certeza que deseja excluir todos os dados do mês de ${p.name}? Esta ação não pode ser desfeita.`)) {
    // Optimistic Update
    state.periods = state.periods.filter(x => x.id !== periodId);
    state.activePeriodId = state.periods.length > 0 ? state.periods[state.periods.length - 1].id : null;

    // Send to Sheets
    if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
      const monthNames = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
      await postToGoogleSheets({
        action: 'delete_month',
        year: p.year,
        monthIdx: monthNames.indexOf(p.month)
      });
      fetchGoogleSheetsData();
    }
  }
};

window.setActivePeriod = function (periodId) {
  state.activePeriodId = periodId;
};

window.editingRecordId = null;
window.editingPeriodId = null;

window.editClient = function (periodId, recordId) {
  const p = state.periods.find(x => x.id === periodId);
  if (!p) return;
  const record = p.records.find(r => r.id === recordId);
  if (!record) return;

  window.editingRecordId = record.id;
  window.editingPeriodId = periodId;

  document.getElementById('clientName').value = record.clientName;
  document.getElementById('avgKw').value = record.avgKw;

  if (record.dateInclusion.includes('/')) {
    const parts = record.dateInclusion.split('/');
    if (parts.length === 3) {
      document.getElementById('dateInclusion').value = `${parts[2]}-${parts[1]}-${parts[0]}`;
    }
  } else {
    document.getElementById('dateInclusion').value = record.dateInclusion;
  }

  document.getElementById('clientStatus').value = record.status;
  document.getElementById('tipoCliente').value = record.tipoCliente;
  document.getElementById('clientSaldo').value = record.saldo;

  if (record.status === 'Saíram') {
    document.getElementById('dataSaidaGroup').style.display = 'block';
    if (document.getElementById('dateSaida')) {
      document.getElementById('dateSaida').required = true;
      if (record.dateSaida) {
        if (record.dateSaida.includes('/')) {
          const sparts = record.dateSaida.split('/');
          document.getElementById('dateSaida').value = `${sparts[2]}-${sparts[1]}-${sparts[0]}`;
        } else {
          document.getElementById('dateSaida').value = record.dateSaida;
        }
      } else {
        document.getElementById('dateSaida').value = '';
      }
    }
  } else {
    document.getElementById('dataSaidaGroup').style.display = 'none';
    if (document.getElementById('dateSaida')) {
      document.getElementById('dateSaida').required = false;
      document.getElementById('dateSaida').value = '';
    }
  }

  const titleEl = document.getElementById('form-main-title');
  if (titleEl) titleEl.textContent = `Editando Cliente: ${record.clientName}`;
  const btnTxt = document.getElementById('btn-submit-text');
  if (btnTxt) btnTxt.textContent = 'Atualizar Dados';
  const cancelBtn = document.getElementById('btn-cancel-edit');
  if (cancelBtn) cancelBtn.style.display = 'inline-block';

  window.scrollTo({ top: document.getElementById('form-main-title').offsetTop - 50, behavior: 'smooth' });
};

window.cancelEditMode = function () {
  window.editingRecordId = null;
  window.editingPeriodId = null;
  document.getElementById('record-form').reset();

  document.getElementById('clientStatus').value = "Ativos";
  document.getElementById('tipoCliente').value = "Pessoa Física";
  document.getElementById('dataSaidaGroup').style.display = 'none';
  if (document.getElementById('dateSaida')) {
    document.getElementById('dateSaida').required = false;
    document.getElementById('dateSaida').value = '';
  }

  const titleEl = document.getElementById('form-main-title');
  if (titleEl) titleEl.textContent = 'Adicionar / Editar Cliente';
  const btnTxt = document.getElementById('btn-submit-text');
  if (btnTxt) btnTxt.textContent = 'Salvar';
  const cancelBtn = document.getElementById('btn-cancel-edit');
  if (cancelBtn) cancelBtn.style.display = 'none';
};

// UI Updaters
function updateDataEntryUI() {
  const periodsList = document.getElementById('periods-list');
  const noPeriodSelected = document.getElementById('no-period-selected');
  const activePeriodContent = document.getElementById('active-period-content');
  const currentPeriodTitle = document.getElementById('current-period-title');
  const recordsTbody = document.getElementById('records-tbody');
  const noRecords = document.getElementById('no-records');
  const tableContainer = document.querySelector('.table-container');
  const searchInput = document.getElementById('global-client-search');
  const searchTerm = searchInput ? normalizeString(searchInput.value) : '';

  // Render periods list on the left
  periodsList.innerHTML = '';
  state.periods.forEach(period => {
    const btn = document.createElement('button');
    btn.className = `period-item ${period.id === state.activePeriodId ? 'active' : ''}`;
    btn.onclick = () => {
      if (searchInput) searchInput.value = ''; // Clear search when picking a period
      window.setActivePeriod(period.id);
    };
    btn.innerHTML = `
      <span>${period.name}</span>
      <span class="period-badge">${period.records.length}</span>
    `;
    periodsList.appendChild(btn);
  });

  const activePeriod = state.periods.find(p => p.id === state.activePeriodId);
  const deleteMonthBtn = document.getElementById('btn-delete-month');

  // Handle display logic (Search Result vs Active Month)
  if (searchTerm) {
    // SEARCH MODE
    noPeriodSelected.style.display = 'none';
    activePeriodContent.style.display = 'block';
    if (deleteMonthBtn) deleteMonthBtn.style.display = 'none';
    currentPeriodTitle.innerHTML = `Resultados para: "${searchInput.value}" <span class="badge">Busca Global</span>`;

    // Flatten and filter ALL records
    let results = [];
    state.periods.forEach(p => {
      p.records.forEach(r => {
        if (normalizeString(r.clientName).includes(searchTerm)) {
          results.push({ ...r, periodId: p.id });
        }
      });
    });

    // Sorting by date (Newest first)
    results.sort((a, b) => b.dateInclusion.localeCompare(a.dateInclusion));

    document.getElementById('period-clients-badge').textContent = `${results.length} Encontrados`;
    document.getElementById('period-kw-badge').textContent = `Filtro Global`;

    if (results.length === 0) {
      tableContainer.style.display = 'none';
      noRecords.style.display = 'block';
      noRecords.innerHTML = `<p>Nenhum cliente encontrado com o termo "${searchInput.value}" em todo o histórico.</p>`;
    } else {
      tableContainer.style.display = 'block';
      noRecords.style.display = 'none';
      renderRecordsTable(results, recordsTbody);
    }
    return;
  }

  // DEFAULT MONTHLY MODE
  if (!activePeriod) {
    noPeriodSelected.style.display = 'flex';
    if (deleteMonthBtn) deleteMonthBtn.style.display = 'none';
    activePeriodContent.style.display = 'none';
    return;
  }

  noPeriodSelected.style.display = 'none';
  activePeriodContent.style.display = 'block';
  if (deleteMonthBtn) deleteMonthBtn.style.display = 'block';
  currentPeriodTitle.innerHTML = activePeriod.name + (state.isCloudSynced ? ' <i class="fa-solid fa-cloud text-orange" style="font-size: 0.8em; margin-left: 8px;" title="Sincronizado via Google Sheets"></i>' : '');

  const totalKw = activePeriod.records.reduce((sum, r) => sum + r.avgKw, 0);
  document.getElementById('period-clients-badge').textContent = `${activePeriod.records.length} Registros`;
  document.getElementById('period-kw-badge').textContent = `${totalKw.toFixed(1)} kW Total`;

  if (activePeriod.records.length === 0) {
    tableContainer.style.display = 'none';
    noRecords.style.display = 'block';
    noRecords.innerHTML = `<p>Nenhum cliente cadastrado nesta aba.</p>`;
  } else {
    tableContainer.style.display = 'block';
    noRecords.style.display = 'none';
    const recordsWithPeriod = activePeriod.records.map(r => ({ ...r, periodId: activePeriod.id }));
    renderRecordsTable(recordsWithPeriod, recordsTbody);
  }
}

/**
 * Shared function to render table rows
 * Expects records to have .periodId attached
 */
function renderRecordsTable(records, tbody) {
  tbody.innerHTML = '';
  records.forEach(record => {
    const isEditing = state.inlineEditingRecordId === record.id;
    const dateParts = record.dateInclusion.split('-');
    const formattedDate = `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`;

    let badgeClass = 'ativos';
    if (record.status === 'Em espera') badgeClass = 'espera';
    if (record.status === 'Saíram') badgeClass = 'sairam';

    const tr = document.createElement('tr');
    tr.setAttribute('data-record-id', record.id);

    if (isEditing) {
      tr.innerHTML = `
        <td><input type="text" class="edit-name" value="${record.clientName}" style="width: 100%; padding: 8px; border-radius: 4px; border: 1px solid var(--orange);"></td>
        <td>
          <select class="edit-tipo" style="width: 100%; padding: 8px; border-radius: 4px; border: 1px solid var(--orange);">
            <option value="Pessoa Física" ${record.tipoCliente === 'Pessoa Física' ? 'selected' : ''}>PF</option>
            <option value="Pessoa Jurídica" ${record.tipoCliente === 'Pessoa Jurídica' ? 'selected' : ''}>PJ</option>
          </select>
        </td>
        <td><input type="date" class="edit-date" value="${record.dateInclusion}" style="width: 100%; padding: 8px; border-radius: 4px; border: 1px solid var(--orange);"></td>
        <td><input type="number" step="0.1" class="edit-kw" value="${record.avgKw}" style="width: 100%; padding: 8px; border-radius: 4px; border: 1px solid var(--orange);"></td>
        <td>
          <select class="edit-status" style="width: 100%; padding: 8px; border-radius: 4px; border: 1px solid var(--orange);">
            <option value="Ativos" ${record.status === 'Ativos' ? 'selected' : ''}>Ativos</option>
            <option value="Em espera" ${record.status === 'Em espera' ? 'selected' : ''}>Em espera</option>
            <option value="Saíram" ${record.status === 'Saíram' ? 'selected' : ''}>Saíram</option>
          </select>
        </td>
        <td><input type="number" step="0.01" class="edit-saldo" value="${record.saldo}" style="width: 100%; padding: 8px; border-radius: 4px; border: 1px solid var(--orange); text-align: center;"></td>
        <td>-</td>
        <td style="display: flex; gap: 8px; justify-content: center; padding: 16px;">
          <button class="btn-icon text-orange" onclick="saveInlineEdit('${record.periodId}', '${record.id}')" title="Salvar Alterações">
            <i class="fa-solid fa-check" style="font-size: 1.5rem;"></i>
          </button>
          <button class="btn-icon text-gray" onclick="cancelInlineEdit()" title="Cancelar">
            <i class="fa-solid fa-xmark" style="font-size: 1.5rem;"></i>
          </button>
        </td>
      `;
    } else {
      tr.innerHTML = `
        <td>
          <strong style="font-size: 1.15rem;">${record.clientName}</strong>
          ${record.inadimplente ? '<span class="badge-inadimplente" style="font-size: 0.8rem; padding: 4px 8px;">Inadimplente</span>' : ''}
        </td>
        <td><span style="font-size: 0.95rem; color: #6b7280; font-weight: 500;">${record.tipoCliente}</span></td>
        <td style="font-size: 1.05rem;">${formattedDate}</td>
        <td style="font-size: 1.05rem;"><strong style="font-size: 1.15rem;">${record.avgKw.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}</strong> kW</td>
        <td><span class="status-badge ${badgeClass}" style="font-size: 0.95rem; padding: 6px 14px;">${record.status}</span></td>
        <td style="font-size: 1.05rem;" data-numeric="true"><strong>${record.saldo.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}</strong> kW</td>
        <td style="font-size: 1.05rem;">${record.dateSaida ? record.dateSaida.split('-').reverse().join('/') : '-'}</td>
        <td style="display: flex; flex-direction: column; gap: 8px; justify-content: center; align-items: center; text-align: center; min-width: 160px; padding: 16px;">
          <button class="action-btn-stacked" onclick="startInlineEdit('${record.periodId}', '${record.id}')" title="Editar este registro (Inline)">
            <i class="fa-solid fa-pencil"></i> Editar
          </button>
          <button class="action-btn-stacked warning ${record.inadimplente ? 'active' : ''}" onclick="toggleInadimplente('${record.periodId}', '${record.id}')" title="Marcar/Desmarcar Inadimplência">
            <i class="fa-solid fa-file-invoice-dollar"></i> Inadimplente
          </button>
          <button class="action-btn-stacked danger" onclick="deleteRecord('${record.periodId}', '${record.id}')" title="Excluir este registro">
            <i class="fa-solid fa-trash"></i> Excluir
          </button>
        </td>
      `;
    }
    tbody.appendChild(tr);
  });
}

function updateDashboardUI() {
  let totalKw = 0;
  let totalClients = 0;

  let countAtivos = 0; let kwAtivos = 0;
  let countEspera = 0; let kwEspera = 0;
  let countSairam = 0; let kwSairam = 0;
  let countInadimplentes = 0; let kwInadimplentes = 0;
  let countInadimplentesAtivos = 0;
  let countPF = 0; let countPJ = 0;
  let kwPF = 0; let kwPJ = 0;

  const labels = [];
  const barData = [];
  const lineData = [];
  const saidasData = [];
  const newClientsData = [];

  let cumulativeEntries = 0;
  let cumulativeExits = 0;
  const lineNetData = [];

  (state.periods || []).forEach(p => {
    let periodTotalKw = 0;
    let periodClientsCount = p.records?.length || 0;
    let monthSaidas = 0;

    (p.records || []).forEach(r => {
      let rKw = r?.avgKw || 0;
      periodTotalKw += rKw;
      if (r?.status === 'Ativos') { countAtivos++; kwAtivos += rKw; }
      else if (r?.status === 'Em espera') { countEspera++; kwEspera += rKw; }
      else if (r?.status === 'Saíram') { countSairam++; monthSaidas++; kwSairam += rKw; }

      if (r?.inadimplente) {
        countInadimplentes++;
        kwInadimplentes += rKw;
        if (r?.status === 'Ativos') countInadimplentesAtivos++;
      }

      if (r?.tipoCliente === 'Pessoa Jurídica') { countPJ++; kwPJ += rKw; }
      else { countPF++; kwPF += rKw; }
    });

    totalKw += periodTotalKw;
    totalClients += periodClientsCount;
    cumulativeEntries += periodClientsCount;
    cumulativeExits += monthSaidas;

    labels.push(p.name);
    barData.push(periodTotalKw);
    lineData.push(cumulativeEntries); // Gross
    lineNetData.push(cumulativeEntries - cumulativeExits); // Net
    saidasData.push(monthSaidas);
    newClientsData.push(periodClientsCount);
  });

  document.getElementById('stat-total-kw').textContent = totalKw.toLocaleString('pt-BR', { maximumFractionDigits: 1 });

  const statTotalClientes = document.getElementById('stat-total-clientes');
  if (statTotalClientes) statTotalClientes.textContent = totalClients;

  document.getElementById('stat-ativos').textContent = countAtivos;
  const statAtivosKw = document.getElementById('stat-ativos-kw');
  if (statAtivosKw) statAtivosKw.textContent = kwAtivos.toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + ' kW Total';

  document.getElementById('stat-espera').textContent = countEspera;
  const statEsperaKw = document.getElementById('stat-espera-kw');
  if (statEsperaKw) statEsperaKw.textContent = kwEspera.toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + ' kW Total';

  document.getElementById('stat-sairam').textContent = countSairam;
  const statSairamKw = document.getElementById('stat-sairam-kw');
  if (statSairamKw) statSairamKw.textContent = kwSairam.toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + ' kW Total';

  const statInadimplentes = document.getElementById('stat-inadimplentes');
  if (statInadimplentes) statInadimplentes.textContent = countInadimplentes;
  const statInadimplentesKw = document.getElementById('stat-inadimplentes-kw');
  if (statInadimplentesKw) statInadimplentesKw.textContent = kwInadimplentes.toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + ' kW Total';

  renderCharts(labels, barData, lineData, lineNetData, countAtivos, countEspera, countSairam, saidasData, countInadimplentesAtivos, totalClients, countPF, countPJ, kwPF, kwPJ, totalKw, newClientsData);
}

function updateUsinasUI() {
  const container = document.getElementById('usinas-accordion-container');
  const emptyState = document.getElementById('no-usinas-state');
  const datalist = document.getElementById('usinas-list');

  if (!container) return;

  if (datalist) {
    datalist.innerHTML = state.usinas.map(u => `<option value="${u.name}"></option>`).join('');
  }

  if (state.usinas.length === 0) {
    if (emptyState) emptyState.style.display = 'flex';
    Array.from(container.children).forEach(child => {
      if (child.id !== 'no-usinas-state') child.remove();
    });
    return;
  }
  if (emptyState) emptyState.style.display = 'none';

  const openAccordions = new Set();
  document.querySelectorAll('.accordion-item.active').forEach(item => {
    if (item.dataset.usinaName) openAccordions.add(item.dataset.usinaName);
  });

  Array.from(container.children).forEach(child => {
    if (child.id !== 'no-usinas-state') child.remove();
  });

  Object.values(window.usinaChartInstances).forEach(chart => {
    if (chart && typeof chart.destroy === 'function') chart.destroy();
  });
  window.usinaChartInstances = {};

  state.usinas.forEach((usina, index) => {
    const isActive = openAccordions.has(usina.name) ? 'active' : '';
    const isEditModeActive = window.usinasInEditMode.has(usina.name) ? 'usina-edit-mode' : '';
    const safeUsinaId = 'usina-' + index;

    const item = document.createElement('div');
    item.className = `accordion-item ${isActive} ${isEditModeActive}`;
    item.dataset.usinaName = usina.name;

    item.innerHTML = `
      <div class="accordion-header" onclick="toggleUsinaAccordion(this)">
        <div class="accordion-title">
          <i class="fa-solid fa-solar-panel text-orange"></i>
          <span>${usina.name}</span>
        </div>
        <div style="display: flex; align-items: center; gap: 12px;">
          <button class="btn-icon" style="color: var(--text-muted); background: none; border: none; cursor: pointer;" onclick="toggleUsinaEditModeMaster(event, '${usina.name.replace(/'/g, "\\'")}')" title="Modo Edição">
            <i class="fa-solid fa-gear"></i>
          </button>
          <button class="btn-danger-sm btn-delete-usina" style="padding: 4px 8px; font-size: 0.8rem;" onclick="deleteUsinaComplete(event, '${usina.name.replace(/'/g, "\\'")}')" title="Excluir Usina">
            <i class="fa-solid fa-trash"></i>
          </button>
          <i class="fa-solid fa-chevron-down accordion-icon"></i>
        </div>
      </div>
      <div class="accordion-content" style="${isActive ? 'max-height: 1000px;' : ''}">
        <div class="accordion-inner">
          ${(!usina.records || usina.records.length === 0) ? `
            <div style="text-align: center; color: var(--text-muted); padding: 32px 0; background: var(--white); border-radius: 8px; border: 1px solid var(--border-color); margin-bottom: 24px;">
              <i class="fa-solid fa-chart-pie" style="font-size: 2rem; color: var(--border-color); margin-bottom: 12px;"></i>
              <p>Sem dados lançados para esta usina.</p>
            </div>
          ` : `
          <div class="usina-charts-container">
            <div class="usina-chart-box" style="display: flex; flex-direction: column;">
              <h4 class="usina-chart-title" id="title-comp-${safeUsinaId}">Último Mês (Comparativo)</h4>
              <div style="flex: 1; min-height: 210px; position: relative;">
                <canvas id="chart-comp-${safeUsinaId}"></canvas>
              </div>
              <div style="display: flex; justify-content: center; align-items: center; gap: 20px; margin-top: auto; padding-top: 24px;">
                <div style="display: flex; align-items: center; gap: 8px;">
                  <span style="display: inline-block; width: 12px; height: 12px; border-radius: 50%; background-color: #f97316;"></span>
                  <span style="color: #374151; font-weight: 500; font-size: 0.85rem;">Energia Injetada</span>
                </div>
                <div style="display: flex; align-items: center; gap: 8px;">
                  <span style="display: inline-block; width: 12px; height: 12px; border-radius: 50%; background-color: #10b981;"></span>
                  <span style="color: #374151; font-weight: 500; font-size: 0.85rem;">Efetivamente Compensado</span>
                </div>
              </div>
            </div>
            <div class="usina-chart-box">
              <h4 class="usina-chart-title">Histórico Geral</h4>
              <div style="height: 250px; position: relative;">
                <canvas id="chart-hist-${safeUsinaId}"></canvas>
              </div>
            </div>
          </div>
          `}
          <div style="margin-top: 32px;">
            <h4 class="usina-chart-title">Histórico de Lançamentos</h4>
            <div class="table-container">
              <table class="data-table">
                <thead>
                  <tr>
                    <th>Mês / Ano</th>
                    <th>Energia Injetada (kW)</th>
                    <th>Efetivamente Compensado (kW)</th>
                    <th class="usina-acoes-header" style="width: 100px;">Ações</th>
                  </tr>
                </thead>
                <tbody id="tbody-usina-${safeUsinaId}">
                  ${generateUsinaTableRowsHTML(usina.name, usina.records)}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    `;

    container.appendChild(item);

    setTimeout(() => {
      renderUsinaCharts(usina, safeUsinaId);
    }, 50);
  });
}

window.toggleUsinaAccordion = function (element) {
  const item = element.closest('.accordion-item');
  const content = item.querySelector('.accordion-content');

  if (item.classList.contains('active')) {
    item.classList.remove('active');
    content.style.maxHeight = '0';
  } else {
    item.classList.add('active');
    content.style.maxHeight = content.scrollHeight + 'px';
  }
};

function generateUsinaTableRowsHTML(usinaName, records) {
  if (!records || records.length === 0) {
    return `<tr><td colspan="4" style="text-align: center; color: #6b7280;">Nenhum registro encontrado.</td></tr>`;
  }

  const safeUsinaName = usinaName.replace(/'/g, "\\'");

  return records.map(record => {
    let fmtMonth = record.month;
    const parts = record.month.split('-');
    if (parts.length === 2) fmtMonth = `${parts[1]}/${parts[0]}`;

    return `
      <tr id="row-${record.id}">
        <td>${fmtMonth}</td>
        <td>
          <span class="read-only-val">${record.esperada}</span>
          <input type="number" class="edit-input-val" id="edit-esperada-${record.id}" value="${record.esperada}" step="0.1" style="display:none; padding: 4px; width: 100px; border: 1px solid var(--border-color); border-radius: 4px;">
        </td>
        <td>
          <span class="read-only-val">${record.compensada}</span>
          <input type="number" class="edit-input-val" id="edit-compensada-${record.id}" value="${record.compensada}" step="0.1" style="display:none; padding: 4px; width: 100px; border: 1px solid var(--border-color); border-radius: 4px;">
        </td>
        <td class="usina-acoes-cell">
          <div class="action-btns-read" style="display: flex;">
            <button class="btn-icon" style="color: var(--text-muted); background: none; border: none; cursor: pointer; margin-right: 8px;" onclick="editUsinaRecord('${record.id}')" title="Editar">
              <i class="fa-solid fa-pencil"></i>
            </button>
            <button class="btn-icon text-red" style="background: none; border: none; cursor: pointer;" onclick="deleteUsinaRecord('${safeUsinaName}', '${record.id}')" title="Excluir">
              <i class="fa-solid fa-trash"></i>
            </button>
          </div>
          <div class="action-btns-edit" style="display: none;">
            <button class="btn-icon" style="color: #10b981; background: none; border: none; cursor: pointer; margin-right: 8px;" onclick="saveUsinaRecord('${safeUsinaName}', '${record.id}', '${record.month}')" title="Salvar">
              <i class="fa-solid fa-check"></i>
            </button>
            <button class="btn-icon" style="color: var(--text-muted); background: none; border: none; cursor: pointer;" onclick="cancelUsinaRecord('${record.id}')" title="Cancelar">
              <i class="fa-solid fa-times"></i>
            </button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

window.toggleUsinaEditModeMaster = function (event, usinaName) {
  event.stopPropagation();
  if (window.usinasInEditMode.has(usinaName)) {
    window.usinasInEditMode.delete(usinaName);
  } else {
    window.usinasInEditMode.add(usinaName);
  }
  updateUsinasUI();
};

window.deleteUsinaComplete = async function (event, usinaName) {
  event.stopPropagation();
  if (confirm(`Atenção: Isso excluirá a usina "${usinaName}" e TODO o seu histórico. Confirmar?`)) {
    state.usinas = state.usinas.filter(u => u.name !== usinaName);
    if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
      await postToGoogleSheets({ action: 'delete_usina_all', usinaName });
    }
  }
};

window.deleteUsinaRecord = async function (usinaName, recordId) {
  if (confirm("Tem certeza que deseja excluir os dados deste mês?")) {
    const tempState = [...state.usinas];
    const usinaIndex = tempState.findIndex(u => u.name === usinaName);
    if (usinaIndex > -1) {
      tempState[usinaIndex].records = tempState[usinaIndex].records.filter(r => r.id !== recordId);
      state.usinas = tempState;
      if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
        await postToGoogleSheets({ action: 'delete_usina', usinaName, id: recordId });
      }
    }
  }
};

window.editUsinaRecord = function (recordId) {
  const row = document.getElementById(`row-${recordId}`);
  if (!row) return;
  row.querySelectorAll('.read-only-val').forEach(el => el.style.display = 'none');
  row.querySelectorAll('.edit-input-val').forEach(el => el.style.display = 'inline-block');
  row.querySelector('.action-btns-read').style.display = 'none';
  row.querySelector('.action-btns-edit').style.display = 'flex';
};

window.cancelUsinaRecord = function (recordId) {
  const row = document.getElementById(`row-${recordId}`);
  if (!row) return;
  row.querySelectorAll('.read-only-val').forEach(el => el.style.display = 'inline-block');
  row.querySelectorAll('.edit-input-val').forEach(el => {
    el.style.display = 'none';
    el.value = el.getAttribute('value');
  });
  row.querySelector('.action-btns-read').style.display = 'flex';
  row.querySelector('.action-btns-edit').style.display = 'none';
};

window.saveUsinaRecord = async function (usinaName, recordId, usinaMonth) {
  const inputEsperada = document.getElementById(`edit-esperada-${recordId}`);
  const inputCompensada = document.getElementById(`edit-compensada-${recordId}`);
  const usinaEsperada = parseFloat(inputEsperada.value);
  const usinaCompensada = parseFloat(inputCompensada.value);

  if (isNaN(usinaEsperada) || isNaN(usinaCompensada)) {
    alert("Valores numéricos inválidos.");
    return;
  }

  const tempState = [...state.usinas];
  const usinaIndex = tempState.findIndex(u => u.name === usinaName);
  if (usinaIndex > -1) {
    const recordIndex = tempState[usinaIndex].records.findIndex(r => r.id === recordId);
    if (recordIndex > -1) {
      tempState[usinaIndex].records[recordIndex].esperada = usinaEsperada;
      tempState[usinaIndex].records[recordIndex].compensada = usinaCompensada;

      state.usinas = tempState;

      if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
        await postToGoogleSheets({
          action: 'update_usina',
          id: recordId,
          usinaName,
          month: usinaMonth,
          esperada: usinaEsperada,
          compensada: usinaCompensada
        });
      }
    }
  }
};

function renderUsinaCharts(usina, canvasIdSuffix) {
  if (!usina.records || usina.records.length === 0) return;

  Chart.defaults.color = '#6b7280';
  Chart.defaults.font.family = "'Inter', sans-serif";
  Chart.defaults.scale.grid.color = '#e5e7eb';

  const lastRecord = usina.records[usina.records.length - 1];

  let fmtMonth = lastRecord.month;
  const parts = lastRecord.month.split('-');
  if (parts.length === 2) {
    const monthNames = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
    const monthIndex = parseInt(parts[1], 10) - 1;
    if (monthIndex >= 0 && monthIndex < 12) {
      fmtMonth = `${monthNames[monthIndex]}/${parts[0]}`;
    } else {
      fmtMonth = `${parts[1]}/${parts[0]}`;
    }
  }

  const titleEl = document.getElementById(`title-comp-${canvasIdSuffix}`);
  if (titleEl) {
    let newTitle = `${fmtMonth} (Comparativo)`;
    if (lastRecord.esperada > 0) {
      const p = ((lastRecord.compensada / lastRecord.esperada) * 100).toFixed(1);
      newTitle += ` <span class="badge badge-accent" style="float: right;">Eficiência: ${p}%</span>`;
    }
    titleEl.innerHTML = newTitle;
  }

  const ctxComp = document.getElementById(`chart-comp-${canvasIdSuffix}`);
  if (ctxComp) {
    window.usinaChartInstances[`comp_${canvasIdSuffix}`] = new Chart(ctxComp, {
      type: 'doughnut',
      data: {
        labels: ['Energia Injetada', 'Efetivamente Compensado'],
        datasets: [{
          data: [lastRecord.esperada, lastRecord.compensada],
          backgroundColor: ['#f97316', '#10b981'],
          borderWidth: 2,
          borderColor: '#ffffff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '60%',
        layout: {
          padding: { top: 30, bottom: 20, left: 30, right: 30 }
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: function (context) { return context.label + ': ' + context.raw + ' kW'; }
            }
          },
          datalabels: {
            anchor: 'end',
            align: 'end',
            offset: 4,
            color: '#374151',
            font: { weight: 'bold', family: "'Inter', sans-serif" },
            formatter: (value) => value > 0 ? value + ' kW' : ''
          }
        }
      },
      plugins: [ChartDataLabels]
    });
  }

  const ctxHist = document.getElementById(`chart-hist-${canvasIdSuffix}`);
  if (ctxHist) {
    const labels = usina.records.map(r => {
      const parts = r.month.split('-');
      return parts.length === 2 ? `${parts[1]}/${parts[0]}` : r.month;
    });
    const dataEsperada = usina.records.map(r => r.esperada);
    const dataCompensada = usina.records.map(r => r.compensada);

    window.usinaChartInstances[`hist_${canvasIdSuffix}`] = new Chart(ctxHist, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          {
            label: 'Energia Injetada',
            data: dataEsperada,
            backgroundColor: '#f97316',
            borderRadius: 4
          },
          {
            label: 'Compensada',
            data: dataCompensada,
            backgroundColor: '#10b981',
            borderRadius: 4
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: { top: 25 } },
        plugins: {
          legend: { position: 'top' },
          datalabels: {
            anchor: 'end',
            align: 'end',
            offset: 4,
            color: '#374151',
            font: { weight: 'bold', family: "'Inter', sans-serif" },
            formatter: (value) => value > 0 ? value.toLocaleString('pt-BR') : ''
          }
        },
        scales: { y: { beginAtZero: true, grace: '20%' } }
      },
      plugins: [ChartDataLabels]
    });
  }
}

function renderCharts(labels, barData, lineData, lineNetData, cAtivos, cEspera, cSairam, saidasData, cInadimplentesAtivos, totalClients, countPF, countPJ, kwPF, kwPJ, totalKwAll, newClientsData) {
  Chart.defaults.color = '#6b7280';
  Chart.defaults.font.family = "'Inter', sans-serif";
  Chart.defaults.scale.grid.color = '#e5e7eb';

  // BAR CHART
  const barCtx = document.getElementById('barChart').getContext('2d');
  if (barChartInstance) barChartInstance.destroy();
  barChartInstance = new Chart(barCtx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Entrada de kW',
        data: barData,
        backgroundColor: '#f97316',
        borderRadius: 4,
        barPercentage: 0.6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 25 } },
      plugins: {
        legend: { display: false },
        tooltip: { backgroundColor: '#1f2937', padding: 10, cornerRadius: 4 },
        datalabels: {
          anchor: 'end',
          align: 'end',
          offset: 4,
          color: '#374151',
          font: { weight: 'bold', family: "'Inter', sans-serif" },
          formatter: (value) => value > 0 ? value.toLocaleString('pt-BR') : ''
        }
      },
      scales: { y: { beginAtZero: true, grace: '20%' } }
    },
    plugins: [ChartDataLabels]
  });

  // DONUT CHART
  const statusCtx = document.getElementById('statusChart').getContext('2d');
  if (statusChartInstance) statusChartInstance.destroy();
  statusChartInstance = new Chart(statusCtx, {
    type: 'doughnut',
    data: {
      labels: ['Ativos', 'Em espera', 'Saíram'],
      datasets: [{
        data: [cAtivos, cEspera, cSairam],
        backgroundColor: ['#10b981', '#f59e0b', '#ef4444'],
        borderWidth: 2,
        borderColor: '#ffffff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '70%',
      layout: { padding: 40 },
      plugins: {
        legend: { display: false },
        datalabels: {
          anchor: 'end',
          align: 'end',
          offset: 2,
          textAlign: 'center',
          color: '#111827',
          font: { weight: 'bold', family: "'Inter', sans-serif", size: 11 },
          formatter: (value, ctx) => {
            let sum = 0;
            let dataArr = ctx.chart.data.datasets[0].data;
            dataArr.forEach(data => { sum += data; });
            if (sum === 0 || value === 0) return '';
            let percentage = (value * 100 / sum).toFixed(1) + "%";
            
            const labelStr = ctx.chart.data.labels[ctx.dataIndex];
            if (labelStr === 'Ativos') {
              return [value.toString(), value === 1 ? "ativo" : "ativos", `(${percentage})`];
            } else if (labelStr === 'Em espera') {
              return [value.toString(), "em", "espera", `(${percentage})`];
            } else if (labelStr === 'Saíram') {
              return [value.toString(), value === 1 ? "saiu" : "saíram", `(${percentage})`];
            }
            return '';
          }
        }
      }
    },
    plugins: [ChartDataLabels]
  });

  // DOM updates for status percentages
  const pctAtivos = totalClients > 0 ? ((cAtivos / totalClients) * 100).toFixed(1) : 0;
  const pctEspera = totalClients > 0 ? ((cEspera / totalClients) * 100).toFixed(1) : 0;
  const pctSairam = totalClients > 0 ? ((cSairam / totalClients) * 100).toFixed(1) : 0;

    const statusContainer = document.getElementById('status-percentages');
    if (statusContainer) {
      statusContainer.innerHTML = `
        <div class="percent-badge ativos">
          <span class="percent-value">${cAtivos} (${pctAtivos}%)</span>
          <span class="percent-label"><span class="percent-dot" style="background-color: #10b981;"></span>Ativos</span>
        </div>
        <div class="percent-badge espera">
          <span class="percent-value">${cEspera} (${pctEspera}%)</span>
          <span class="percent-label"><span class="percent-dot" style="background-color: #f59e0b;"></span>Em Espera</span>
        </div>
        <div class="percent-badge sairam">
          <span class="percent-value" style="color: #ef4444;">${cSairam} (${pctSairam}%)</span>
          <span class="percent-label"><span class="percent-dot" style="background-color: #ef4444;"></span>Saíram</span>
        </div>
      `;
    }

  // INADIMPLENCIA CHART
  const inadCtx = document.getElementById('inadimplenciaChart');
  if (inadCtx) {
    if (inadimplenciaChartInstance) inadimplenciaChartInstance.destroy();

    const emDia = cAtivos - cInadimplentesAtivos;
    const taxaInadimplencia = cAtivos > 0 ? ((cInadimplentesAtivos / cAtivos) * 100).toFixed(1) : 0;

    inadimplenciaChartInstance = new Chart(inadCtx.getContext('2d'), {
      type: 'doughnut',
      data: {
        labels: ['Em dia (Ativos)', 'Inadimplentes (Ativos)'],
        datasets: [{
          data: [emDia, cInadimplentesAtivos],
          backgroundColor: ['#10b981', '#ef4444'],
          borderWidth: 2,
          borderColor: '#ffffff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '70%',
        layout: { padding: 40 },
        plugins: {
          legend: { display: false },
          datalabels: {
            anchor: 'end',
            align: 'end',
            offset: 2,
            textAlign: 'center',
            color: '#111827',
            font: { weight: 'bold', family: "'Inter', sans-serif", size: 11 },
            formatter: (value, ctx) => {
              let sum = 0;
              let dataArr = ctx.chart.data.datasets[0].data;
              dataArr.forEach(data => { sum += data; });
              if (sum === 0 || value === 0) return '';
              let percentage = (value * 100 / sum).toFixed(1) + "%";
              
              const labelStr = ctx.chart.data.labels[ctx.dataIndex];
              if (labelStr === 'Em dia (Ativos)') {
                return [value.toString(), "em dia", `(${percentage})`];
              } else {
                return [value.toString(), value === 1 ? "inadimplente" : "inadimplentes", `(${percentage})`];
              }
            }
          }
        }
      },
      plugins: [ChartDataLabels]
    });

    const inadContainer = document.getElementById('inadimplencia-percentages');
    if (inadContainer) {
      const pctEmDia = cAtivos > 0 ? ((emDia / cAtivos) * 100).toFixed(1) : 0;
      inadContainer.innerHTML = `
        <div class="percent-badge" style="border-left: 4px solid #10b981;">
          <span class="percent-value" style="color: #10b981;">${emDia} (${pctEmDia}%)</span>
          <span class="percent-label"><span class="percent-dot" style="background-color: #10b981;"></span>Em Dia (Ativos)</span>
        </div>
        <div class="percent-badge inadimplentes">
          <span class="percent-value">${cInadimplentesAtivos} (${taxaInadimplencia}%)</span>
          <span class="percent-label"><span class="percent-dot" style="background-color: #ef4444;"></span>Inadimplentes</span>
        </div>
      `;
    }
  }

  // SAÍDAS CHART
  const saidasCtx = document.getElementById('saidasChart').getContext('2d');
  if (saidasChartInstance) saidasChartInstance.destroy();
  saidasChartInstance = new Chart(saidasCtx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Saídas',
        data: saidasData,
        backgroundColor: '#ef4444',
        borderRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        datalabels: {
          color: '#FFFFFF',
          font: { weight: 'bold' },
          formatter: (value) => value > 0 ? value : ''
        }
      },
      scales: { y: { beginAtZero: true, grace: '20%', ticks: { stepSize: 1 } } }
    },
    plugins: [ChartDataLabels]
  });

  // NEW CLIENTS CHART
  const newClientsCtx = document.getElementById('newClientsChart').getContext('2d');
  if (newClientsChartInstance) newClientsChartInstance.destroy();
  newClientsChartInstance = new Chart(newClientsCtx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Novos Contratos',
        data: newClientsData,
        backgroundColor: '#f97316',
        borderRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        datalabels: {
          color: '#FFFFFF',
          font: { weight: 'bold' },
          formatter: (value) => value > 0 ? value : ''
        }
      },
      scales: { y: { beginAtZero: true, grace: '20%', ticks: { stepSize: 1 } } }
    },
    plugins: [ChartDataLabels]
  });

  // LINE CHART (Growth Trend)
  const lineCtx = document.getElementById('lineChart').getContext('2d');
  if (lineChartInstance) lineChartInstance.destroy();
  lineChartInstance = new Chart(lineCtx, {
    type: 'line',
    data: {
      labels: labels,
      datasets: [
        {
          label: 'Entrada Total de Clientes',
          data: lineData,
          borderColor: '#f97316',
          backgroundColor: 'transparent',
          borderWidth: 3,
          pointBackgroundColor: '#ffffff',
          pointBorderColor: '#f97316',
          pointBorderWidth: 2,
          pointRadius: 5,
          tension: 0.3,
          fill: false,
          datalabels: {
            align: 'top',
            anchor: 'end',
            offset: 8,
            color: '#1f2937',
            font: { weight: 'bold', size: 11, family: "'Inter', sans-serif" }
          }
        },
        {
          label: 'Saldo Líquido (Entradas - Saídas)',
          data: lineNetData,
          borderColor: '#4b5563',
          backgroundColor: 'transparent',
          borderWidth: 3,
          pointBackgroundColor: '#ffffff',
          pointBorderColor: '#4b5563',
          pointBorderWidth: 2,
          pointRadius: 5,
          tension: 0.3,
          fill: false,
          datalabels: {
            align: 'bottom',
            anchor: 'start',
            offset: 8,
            color: '#374151',
            font: { weight: 'bold', size: 11, family: "'Inter', sans-serif" }
          }
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 35, right: 20, bottom: 20 } },
      plugins: {
        legend: {
          display: true,
          position: 'top',
          align: 'end',
          labels: { boxWidth: 12, font: { weight: 600 } }
        },
        datalabels: {
          formatter: (value) => value
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          grace: '15%',
          ticks: { stepSize: 1, color: '#6b7280' },
          grid: { borderDash: [5, 5] }
        },
        x: {
          ticks: { color: '#6b7280' },
          grid: { display: false }
        }
      }
    },
    plugins: [ChartDataLabels]
  });

  // TIPO CLIENTE CHART (QUANTIDADE - PF/PJ)
  const tipoCtx = document.getElementById('tipoClienteChart');
  if (tipoCtx) {
    if (tipoClienteChartInstance) tipoClienteChartInstance.destroy();
    tipoClienteChartInstance = new Chart(tipoCtx.getContext('2d'), {
      type: 'doughnut',
      data: {
        labels: ['Pessoa Física', 'Pessoa Jurídica'],
        datasets: [{
          data: [countPF, countPJ],
          backgroundColor: ['#f97316', '#4b5563'],
          borderWidth: 2,
          borderColor: '#ffffff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '70%',
        layout: { padding: 40 },
        plugins: {
          legend: { display: false },
          datalabels: {
            anchor: 'end',
            align: 'end',
            offset: 2,
            textAlign: 'center',
            color: '#111827',
            font: { weight: 'bold', family: "'Inter', sans-serif", size: 11 },
            formatter: (value, ctx) => {
              const sum = ctx.chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
              if (sum === 0 || value === 0) return '';
              const pct = (value * 100 / sum).toFixed(1) + "%";
              
              const labelStr = ctx.chart.data.labels[ctx.dataIndex];
              if (labelStr === 'Pessoa Física') {
                return [value.toString(), value === 1 ? "pessoa" : "pessoas", value === 1 ? "física" : "físicas", `(${pct})`];
              } else {
                return [value.toString(), value === 1 ? "pessoa" : "pessoas", value === 1 ? "jurídica" : "jurídicas", `(${pct})`];
              }
            }
          }
        }
      },
      plugins: [ChartDataLabels]
    });

    const tipoContainer = document.getElementById('tipoCliente-percentages');
    if (tipoContainer) {
      const pctPF = totalClients > 0 ? ((countPF / totalClients) * 100).toFixed(1) : 0;
      const pctPJ = totalClients > 0 ? ((countPJ / totalClients) * 100).toFixed(1) : 0;
      tipoContainer.innerHTML = `
        <div class="percent-badge pf">
          <span class="percent-value">${countPF} (${pctPF}%)</span>
          <span class="percent-label"><span class="percent-dot" style="background-color: #f97316;"></span>PF</span>
        </div>
        <div class="percent-badge pj">
          <span class="percent-value">${countPJ} (${pctPJ}%)</span>
          <span class="percent-label"><span class="percent-dot" style="background-color: #4b5563;"></span>PJ</span>
        </div>
      `;
    }
  }

  // TIPO CLIENTE KW CHART (VOLUME - PF/PJ)
  const tipoKwCtx = document.getElementById('tipoClienteKwChart');
  if (tipoKwCtx) {
    if (tipoClienteKwChartInstance) tipoClienteKwChartInstance.destroy();
    tipoClienteKwChartInstance = new Chart(tipoKwCtx.getContext('2d'), {
      type: 'doughnut',
      data: {
        labels: ['Volume PF', 'Volume PJ'],
        datasets: [{
          data: [kwPF, kwPJ],
          backgroundColor: ['#f97316', '#4b5563'],
          borderWidth: 2,
          borderColor: '#ffffff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '70%',
        layout: { padding: 40 },
        plugins: {
          legend: { display: false },
          datalabels: {
            anchor: 'end',
            align: 'end',
            offset: 2,
            textAlign: 'center',
            color: '#111827',
            font: { weight: 'bold', family: "'Inter', sans-serif", size: 10 },
            formatter: (value, ctx) => {
              const sum = ctx.chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
              if (sum === 0 || value === 0) return '';
              const pct = (value * 100 / sum).toFixed(1) + "%";
              
              const labelStr = ctx.chart.data.labels[ctx.dataIndex];
              if (labelStr === 'Volume PF') {
                return [`${value.toFixed(0)} kW`, "(PF)", `(${pct})`];
              } else {
                return [`${value.toFixed(0)} kW`, "(PJ)", `(${pct})`];
              }
            }
          }
        }
      },
      plugins: [ChartDataLabels]
    });

    const tipoKwContainer = document.getElementById('tipoClienteKw-percentages');
    if (tipoKwContainer) {
      const pctKwPF = totalKwAll > 0 ? ((kwPF / totalKwAll) * 100).toFixed(1) : 0;
      const pctKwPJ = totalKwAll > 0 ? ((kwPJ / totalKwAll) * 100).toFixed(1) : 0;
      tipoKwContainer.innerHTML = `
        <div class="percent-badge pf">
          <span class="percent-value">${kwPF.toFixed(0)} (${pctKwPF}%)</span>
          <span class="percent-label"><span class="percent-dot" style="background-color: #f97316;"></span>Vol. PF</span>
        </div>
        <div class="percent-badge pj">
          <span class="percent-value">${kwPJ.toFixed(0)} (${pctKwPJ}%)</span>
          <span class="percent-label"><span class="percent-dot" style="background-color: #4b5563;"></span>Vol. PJ</span>
        </div>
      `;
    }
  }
}

function loadDemoData() {
  state.periods = [
    {
      id: "p-2026-0",
      name: "Janeiro de 2026",
      month: "Janeiro",
      year: 2026,
      records: [
        { id: "1", clientName: "Supermercado Alfa", dateInclusion: "2026-01-05", avgKw: 1450.5, status: "Ativos", tipoCliente: "Pessoa Jurídica", saldo: 2450.00, inadimplente: false, contasAtrasadas: [] },
        { id: "2", clientName: "Indústria Ômega", dateInclusion: "2026-01-12", avgKw: 850.0, status: "Em espera", tipoCliente: "Pessoa Jurídica", saldo: 0, inadimplente: true, contasAtrasadas: [] }
      ]
    },
    {
      id: "p-2026-1",
      name: "Fevereiro de 2026",
      month: "Fevereiro",
      year: 2026,
      records: [
        { id: "3", clientName: "Posto de Combustível Beta", dateInclusion: "2026-02-01", avgKw: 2100.2, status: "Ativos", tipoCliente: "Pessoa Jurídica", saldo: 3100.50, inadimplente: false, contasAtrasadas: [] },
        { id: "4", clientName: "Hotel Delta", dateInclusion: "2026-02-15", avgKw: 560.8, status: "Ativos", tipoCliente: "Pessoa Física", saldo: 140.20, inadimplente: false, contasAtrasadas: [] },
        { id: "5", clientName: "Escola Epsilon", dateInclusion: "2026-02-28", avgKw: 1200.0, status: "Saíram", tipoCliente: "Pessoa Jurídica", saldo: 0, inadimplente: false, contasAtrasadas: [] }
      ]
    }
  ];
  
  state.usinas = [
    { id: "u-1", name: "Usina Solar Norte" },
    { id: "u-2", name: "Usina Solar Sul" }
  ];

  state.creditos = [
    { id: "c-1", date: "2026-01", kwTotal: 5000, historico: "Mês inicial" }
  ];

  state.activePeriodId = "p-2026-1";
}

function setupInteractiveCards() {
  const cards = document.querySelectorAll('.clickable-card');
  const modal = document.getElementById('clients-modal');
  const closeBtn = document.getElementById('close-clients-modal-btn');

  cards.forEach(card => {
    card.addEventListener('click', () => {
      const filter = card.getAttribute('data-filter');
      openClientsModal(filter);
    });
  });

  closeBtn.addEventListener('click', () => {
    modal.classList.remove('active');
  });

  // Modal will also close by window click, handled in setupModals or here:
  window.addEventListener('click', (e) => {
    if (e.target === modal) {
      modal.classList.remove('active');
    }
  });
}

function openClientsModal(filterType) {
  const modal = document.getElementById('clients-modal');
  const title = document.getElementById('clients-modal-title');
  const thead = document.getElementById('clients-modal-thead');
  const tbody = document.getElementById('clients-modal-tbody');
  const noRecords = document.getElementById('clients-modal-no-records');
  const tableContainer = modal.querySelector('.table-container');

  // Collect all clients across all periods
  let allClients = [];
  state.periods.forEach(p => {
    p.records.forEach(r => {
      allClients.push(r);
    });
  });

  // Filter
  let filteredClients = allClients;
  if (filterType === 'Inadimplentes') {
    title.textContent = 'Clientes Inadimplentes';
    filteredClients = allClients.filter(c => c.inadimplente === true);
  } else if (filterType === 'Ativos') {
    title.textContent = 'Clientes Ativos';
    filteredClients = allClients.filter(c => c.status === 'Ativos');
  } else if (filterType === 'Em espera') {
    title.textContent = 'Clientes Em Espera';
    filteredClients = allClients.filter(c => c.status === 'Em espera');
  } else if (filterType === 'Saíram') {
    title.textContent = 'Clientes que Saíram';
    filteredClients = allClients.filter(c => c.status === 'Saíram');
  } else {
    title.textContent = 'Todos os Clientes (Entrada)';
  }

  // Build thead
  if (filterType === 'Saíram') {
    thead.innerHTML = `
      <th>Nome do Cliente</th>
      <th>Tipo</th>
      <th>Média de kW</th>
      <th>Data de Inclusão</th>
      <th>Status</th>
      <th>Data de Saída</th>
      <th>Ações</th>
    `;
  } else {
    thead.innerHTML = `
      <th>Nome do Cliente</th>
      <th>Tipo</th>
      <th>Média de kW</th>
      <th>Data de Inclusão</th>
      <th>Status</th>
      <th>Ações</th>
    `;
  }

  // Build tbody
  tbody.innerHTML = '';
  if (filteredClients.length === 0) {
    if (tableContainer.querySelector('.data-table')) {
      tableContainer.querySelector('.data-table').style.display = 'none';
    }
    noRecords.style.display = 'block';
  } else {
    if (tableContainer.querySelector('.data-table')) {
      tableContainer.querySelector('.data-table').style.display = 'table';
    }
    noRecords.style.display = 'none';

    filteredClients.forEach(c => {
      const dateParts = c.dateInclusion.split('-');
      const formattedDate = dateParts.length === 3 ? `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}` : c.dateInclusion;

      let badgeClass = 'ativos';
      if (c.status === 'Em espera') badgeClass = 'espera';
      if (c.status === 'Saíram') badgeClass = 'sairam';

      let rowHtml = `
        <td>
          <strong>${c.clientName}</strong>
          ${c.inadimplente ? '<span class="badge-inadimplente">Inadimplente</span>' : ''}
        </td>
        <td><span style="font-size: 0.85rem; color: #6b7280; font-weight: 500;">${c.tipoCliente}</span></td>
        <td><strong>${c.avgKw.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}</strong> kW</td>
        <td>${formattedDate}</td>
        <td><span class="status-badge ${badgeClass}">${c.status}</span></td>
      `;

      if (filterType === 'Saíram') {
        const dateSaidaFormatted = c.dateSaida ? c.dateSaida.split('-').reverse().join('/') : '-';
        rowHtml += `<td>${dateSaidaFormatted}</td>`;
      }

      let recordPeriodId = '';
      state.periods.forEach(p => {
        if (p.records.find(r => r.id === c.id)) recordPeriodId = p.id;
      });

      rowHtml += `
      <td>
        <button class="btn-toggle-inadimplente ${c.inadimplente ? 'active' : ''}" onclick="toggleInadimplente('${recordPeriodId}', '${c.id}'); setTimeout(() => openClientsModal('${filterType}'), 100);" title="Marcar/Desmarcar Inadimplência">
          <i class="fa-solid fa-file-invoice-dollar"></i> Mudar Status
        </button>
      </td>`;

      const tr = document.createElement('tr');
      tr.innerHTML = rowHtml;
      tbody.appendChild(tr);
    });
  }

  modal.classList.add('active');
}

// Organização Clientes logic
function setupOrganizacaoFilters() {
  const kwInput = document.getElementById('filterKwMin');
  const statusSelect = document.getElementById('filterStatus');
  const nameInput = document.getElementById('filterName');
  const btnClear = document.getElementById('btn-clear-filters');

  if (!kwInput || !statusSelect || !btnClear) return;

  kwInput.addEventListener('input', () => updateOrganizacaoUI());
  statusSelect.addEventListener('change', () => updateOrganizacaoUI());
  if (nameInput) nameInput.addEventListener('input', () => updateOrganizacaoUI());

  btnClear.addEventListener('click', () => {
    kwInput.value = '';
    statusSelect.value = 'Todos';
    if (nameInput) nameInput.value = '';
    updateOrganizacaoUI();
  });
}

function updateOrganizacaoUI() {
  const tbody = document.getElementById('organizacao-tbody');
  const noResults = document.getElementById('organizacao-no-results');
  const summaryContainer = document.getElementById('organizacao-summary-container');
  const tableContainer = document.querySelector('#organizacao-view .table-container');

  if (!tbody || !summaryContainer) return;

  const kwMin = parseFloat(document.getElementById('filterKwMin').value) || 0;
  const statusFilter = document.getElementById('filterStatus').value;
  const nameFilter = (document.getElementById('filterName').value || '').trim();
  const normalizedSearch = normalizeString(nameFilter);

  // Flatten all records from all periods
  let allClients = [];
  (state.periods || []).forEach(p => {
    (p.records || []).forEach(r => {
      allClients.push(r);
    });
  });

  const totalCount = allClients.length;

  // Apply filters
  const filtered = allClients.filter(c => {
    const matchesKw = c.avgKw >= kwMin;
    const matchesName = normalizedSearch === '' || normalizeString(c.clientName).includes(normalizedSearch);
    let matchesStatus = true;

    if (statusFilter === 'Inadimplentes') {
      matchesStatus = c.inadimplente === true;
    } else if (statusFilter !== 'Todos') {
      matchesStatus = c.status === statusFilter;
    }

    return matchesKw && matchesStatus && matchesName;
  });

  const filteredCount = filtered.length;
  const percentage = totalCount > 0 ? ((filteredCount / totalCount) * 100).toFixed(1) : 0;
  const isFiltered = kwMin > 0 || statusFilter !== 'Todos' || nameFilter !== '';

  // Calculate Total Balance for filtered clients
  const totalBalance = filtered.reduce((sum, c) => sum + (c.saldo || 0), 0);

  const footerContainer = document.getElementById('organizacao-footer-container');

  // Render Summary Banners (Existing Results + New Balance Card)
  if (totalCount === 0) {
    summaryContainer.innerHTML = '';
    if (footerContainer) footerContainer.innerHTML = '';
  } else {
    summaryContainer.innerHTML = `
      <div class="summary-banner">
        <div class="metric-main">
          <i class="fa-solid fa-square-poll-vertical"></i>
          Resultados: <span class="highlight">${filteredCount}</span> clientes encontrados
        </div>
        <div class="metric-detail">
          ${isFiltered
        ? `Representa <span class="highlight">${percentage}%</span> do total de <span class="highlight">${totalCount}</span> clientes`
        : `Mostrando todos os <span class="highlight">${totalCount}</span> clientes cadastrados (100%)`
      }
        </div>
      </div>
    `;

    if (footerContainer) {
      footerContainer.innerHTML = `
        <div class="summary-banner balance-banner" style="margin-top: 0; margin-bottom: 24px;">
          <div style="display: flex; align-items: center; gap: 20px;">
            <div class="balance-card-icon">
              <i class="fa-solid fa-money-bill-transfer"></i>
            </div>
            <div>
              <div class="balance-card-title">Saldo Total Acumulado</div>
              <div class="balance-card-value" style="color: #111827;">${totalBalance.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })} kW</div>
            </div>
          </div>
          <div class="metric-detail" style="color: #374151; font-weight: 500;">
            <i class="fa-solid fa-circle-info" style="color: #10b981; margin-right: 4px;"></i>
            Soma de todos os saldos no filtro atual
          </div>
        </div>
      `;
    }
  }

  tbody.innerHTML = '';

  if (filtered.length === 0) {
    if (tableContainer) tableContainer.style.display = 'none';
    noResults.style.display = 'block';
  } else {
    if (tableContainer) tableContainer.style.display = 'block';
    noResults.style.display = 'none';

    filtered.forEach(c => {
      const dateParts = c.dateInclusion.split('-');
      const formattedDate = dateParts.length === 3 ? `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}` : c.dateInclusion;

      let badgeClass = 'ativos';
      if (c.status === 'Em espera') badgeClass = 'espera';
      if (c.status === 'Saíram') badgeClass = 'sairam';

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="text-align: center;"><strong>${c.clientName}</strong> ${c.inadimplente ? '<span class="badge-inadimplente">Inadimplente</span>' : ''}</td>
        <td style="text-align: center;"><span style="font-size: 0.85rem; color: #6b7280; font-weight: 500;">${c.tipoCliente}</span></td>
        <td style="text-align: center;"><strong>${c.avgKw.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}</strong> kW</td>
        <td style="text-align: center;">${formattedDate}</td>
        <td style="text-align: center;"><span class="status-badge ${badgeClass}">${c.status}</span></td>
        <td style="text-align: center;" data-numeric="true"><strong>${c.saldo ? c.saldo.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) : '0,0'}</strong> kW</td>
      `;
      tbody.appendChild(tr);
    });
  }
}
function setupCreditosForm() {
  const form = document.getElementById('credito-form');
  if (!form) return;

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const unidade = document.getElementById('creditoUnidade').value.trim();
    const uc = document.getElementById('creditoUC').value.trim();
    const total = parseFloat(document.getElementById('creditoTotal').value);
    const mes = document.getElementById('creditoMes').value;
    const ano = parseInt(document.getElementById('creditoAno').value);

    const newRecord = {
      id: `c-${Date.now()}`,
      unidade,
      uc,
      total,
      mes,
      ano,
      timestamp: Date.now()
    };

    state.creditos = [...state.creditos, newRecord];

    // Cloud Sync if needed
    if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
      await postToGoogleSheets({
        action: 'add_credito',
        ...newRecord
      });
    }

    const msg = document.getElementById('credito-success-msg');
    if (msg) {
      msg.style.display = 'inline-block';
      setTimeout(() => { msg.style.display = 'none'; }, 3000);
    }

    form.reset();
    document.getElementById('creditoAno').value = ano;
  });
}

function updateCreditosUI() {
  const chartCanvas = document.getElementById('creditosChart');
  const tbody = document.getElementById('creditos-tbody');
  const noCreditos = document.getElementById('no-creditos');

  if (!chartCanvas || !tbody) return;

  const creditos = state.creditos;
  const monthOrder = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

  // Table Rendering (Grouped by Unit)
  tbody.innerHTML = '';
  if (creditos.length === 0) {
    if (creditosChartInstance) creditosChartInstance.destroy();
    noCreditos.style.display = 'block';
    return;
  }
  noCreditos.style.display = 'none';

  // Get unique units
  const units = [...new Set(creditos.map(c => c.unidade))].sort();

  units.forEach(unit => {
    // Add Group Header Row
    const headerTr = document.createElement('tr');
    headerTr.className = 'table-group-header';
    headerTr.innerHTML = `<td colspan="6"><i class="fa-solid fa-layer-group"></i> Unidade: ${unit}</td>`;
    tbody.appendChild(headerTr);

    // Filter and sort records for this unit
    const unitRecords = creditos.filter(c => c.unidade === unit).sort((a, b) => {
      if (a.ano !== b.ano) return b.ano - a.ano;
      return monthOrder.indexOf(b.mes) - monthOrder.indexOf(a.mes);
    });

    unitRecords.forEach(c => {
      const isEditing = state.inlineEditingCreditoId === c.id;
      const tr = document.createElement('tr');
      tr.setAttribute('data-credito-id', c.id);

      if (isEditing) {
        tr.innerHTML = `
          <td><input type="text" class="edit-unidade" value="${c.unidade}" style="width: 100%; text-align: center;"></td>
          <td><input type="text" class="edit-uc" value="${c.uc}" style="width: 100%; text-align: center;"></td>
          <td>
            <select class="edit-mes" style="width: 100%; text-align: center;">
              ${monthOrder.map(m => `<option value="${m}" ${m === c.mes ? 'selected' : ''}>${m}</option>`).join('')}
            </select>
          </td>
          <td><input type="number" class="edit-ano" value="${c.ano}" style="width: 100%; text-align: center;"></td>
          <td><input type="number" step="0.1" class="edit-total" value="${c.total}" style="width: 100%; text-align: center;"></td>
          <td style="display: flex; gap: 8px; justify-content: center;">
            <button class="btn-icon" onclick="saveInlineEditCredito('${c.id}')" title="Salvar" style="color: #10b981;">
              <i class="fa-solid fa-check"></i>
            </button>
            <button class="btn-icon" onclick="cancelInlineEditCredito()" title="Cancelar" style="color: #ef4444;">
              <i class="fa-solid fa-xmark"></i>
            </button>
          </td>
        `;
      } else {
        tr.innerHTML = `
          <td><span style="color: var(--text-muted); font-size: 0.85rem;">${c.unidade}</span></td>
          <td>${c.uc}</td>
          <td><strong>${c.mes}</strong></td>
          <td>${c.ano}</td>
          <td><strong>${c.total.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}</strong> kW</td>
          <td style="display: flex; gap: 8px; justify-content: center;">
            <button class="btn-icon" onclick="startInlineEditCredito('${c.id}')" title="Editar" style="color: var(--orange);">
              <i class="fa-solid fa-pencil"></i>
            </button>
            <button class="btn-icon" onclick="deleteCredito('${c.id}')" title="Excluir" style="color: #ef4444;">
              <i class="fa-solid fa-trash"></i>
            </button>
          </td>
        `;
      }
      tbody.appendChild(tr);
    });
  });

  // Chart Rendering
  const periods = [...new Set(creditos.map(c => `${c.mes} ${c.ano}`))].sort((a, b) => {
    const [m1, y1] = a.split(' ');
    const [m2, y2] = b.split(' ');
    if (y1 !== y2) return parseInt(y1) - parseInt(y2);
    return monthOrder.indexOf(m1) - monthOrder.indexOf(m2);
  });

  const datasets = periods.map((period, index) => {
    // Varied, professional categorical palette
    const colors = [
      '#3b82f6', // Blue
      '#10b981', // Emerald
      '#8b5cf6', // Violet
      '#f59e0b', // Amber
      '#f97316', // Orange
      '#06b6d4', // Cyan
      '#ec4899', // Pink
      '#6366f1', // Indigo
      '#14b8a6', // Teal
      '#f43f5e'  // Rose
    ];

    return {
      label: period,
      data: units.map(unit => {
        const [m, y] = period.split(' ');
        const record = creditos.find(c => c.unidade === unit && c.mes === m && parseInt(c.ano) === parseInt(y));
        return record ? record.total : 0;
      }),
      backgroundColor: colors[index % colors.length],
      borderColor: colors[index % colors.length],
      borderWidth: 1,
      borderRadius: 4,
      barPercentage: 0.8,
      categoryPercentage: 0.8
    };
  });

  if (creditosChartInstance) creditosChartInstance.destroy();

  const ctx = chartCanvas.getContext('2d');
  creditosChartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: units,
      datasets: datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: true,
          position: 'top',
          labels: { usePointStyle: true, padding: 20, font: { weight: 600 } }
        },
        tooltip: {
          mode: 'index',
          intersect: false,
          backgroundColor: '#1f2937',
          padding: 12,
          cornerRadius: 8,
          callbacks: {
            title: (items) => `Unidade: ${items[0].label}`,
            label: (item) => `${item.dataset.label}: ${item.raw.toLocaleString('pt-BR')} kW`
          }
        },
        datalabels: {
          anchor: 'center',
          align: 'center',
          color: '#ffffff',
          font: { weight: 'bold', size: 10 },
          textAlign: 'center',
          formatter: (value, context) => {
            if (value === 0) return '';
            const monthName = context.dataset.label.split(' ')[0].substring(0, 3);
            const valFormatted = value.toLocaleString('pt-BR', { maximumFractionDigits: 0 });
            return `${monthName}\n${valFormatted}`;
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          grace: '15%',
          title: { display: true, text: 'Créditos (kW)', font: { weight: 'bold' } },
          grid: { borderDash: [5, 5] }
        },
        x: {
          grid: { display: false },
          ticks: { font: { weight: 700, size: 12 } }
        }
      }
    },
    plugins: [ChartDataLabels]
  });
}

window.startInlineEditCredito = function (id) {
  state.inlineEditingCreditoId = id;
};

window.cancelInlineEditCredito = function () {
  state.inlineEditingCreditoId = null;
};

window.saveInlineEditCredito = async function (id) {
  const row = document.querySelector(`tr[data-credito-id="${id}"]`);
  if (!row) return;

  const newUnidade = row.querySelector('.edit-unidade').value.trim();
  const newUC = row.querySelector('.edit-uc').value.trim();
  const newMes = row.querySelector('.edit-mes').value;
  const newAno = parseInt(row.querySelector('.edit-ano').value);
  const newTotal = parseFloat(row.querySelector('.edit-total').value);

  const tempCreditos = [...state.creditos];
  const idx = tempCreditos.findIndex(c => c.id === id);
  if (idx !== -1) {
    tempCreditos[idx] = { ...tempCreditos[idx], unidade: newUnidade, uc: newUC, mes: newMes, ano: newAno, total: newTotal };
    state.creditos = tempCreditos;

    if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
      await postToGoogleSheets({
        action: 'edit_credito',
        id, unidade: newUnidade, uc: newUC, mes: newMes, ano: newAno, total: newTotal
      });
    }
  }
  window.cancelInlineEditCredito();
};

window.deleteCredito = async function (id) {
  if (confirm("Deseja realmente excluir este lançamento de crédito?")) {
    const tempCreditos = state.creditos.filter(c => c.id !== id);
    state.creditos = tempCreditos;

    if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
      await postToGoogleSheets({ action: 'delete_credito', id });
    }
  }
};

// --- Inadimplencia Module Logic ---
function setupInadimplenciaForm() {
  const form = document.getElementById('inadimplencia-form');
  if (!form) return;

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const select = document.getElementById('inadCliSelect');
    if (!select.value) return;

    const [periodId, recordId] = select.value.split('|');
    const dataEmissao = document.getElementById('inadData').value;
    const valor = parseFloat(document.getElementById('inadValor').value) || 0;

    const tempState = [...state.periods];
    const pIdx = tempState.findIndex(p => p.id === periodId);
    if (pIdx !== -1) {
      const rIdx = tempState[pIdx].records.findIndex(r => r.id === recordId);
      if (rIdx !== -1) {
        const c = tempState[pIdx].records[rIdx];
        if (!c.contasAtrasadas) c.contasAtrasadas = [];
        c.contasAtrasadas.push({
          id: `conta-${Date.now()}`,
          dataEmissao,
          valor
        });
        state.periods = tempState;

        if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
          await postToGoogleSheets({ action: 'add_conta_atrasada', periodId, recordId, dataEmissao, valor });
        }
      }
    }

    document.getElementById('inadData').value = '';
    document.getElementById('inadValor').value = '';
    select.focus();
  });
}

function updateInadimplenciaUI() {
  const select = document.getElementById('inadCliSelect');
  const grid = document.getElementById('inadimplencia-cards-grid');
  const noRecords = document.getElementById('no-inadimplencia-records');
  if (!select || !grid) return;

  const inadimplentes = [];
  (state.periods || []).forEach(p => {
    (p.records || []).forEach(r => {
      if (r.inadimplente === true) inadimplentes.push({ periodId: p.id, record: r });
    });
  });

  const selectCurrentVal = select.value;
  select.innerHTML = '<option value="" disabled selected>Selecione um cliente da lista...</option>';
  inadimplentes.forEach(item => {
    const opt = document.createElement('option');
    opt.value = `${item.periodId}|${item.record.id}`;
    opt.textContent = item.record.clientName;
    select.appendChild(opt);
  });
  if (selectCurrentVal) select.value = selectCurrentVal;

  const defaultingClients = inadimplentes;

  grid.innerHTML = '';
  if (defaultingClients.length === 0) {
    grid.style.display = 'none';
    if(noRecords) noRecords.style.display = 'flex';
  } else {
    grid.style.display = 'grid';
    if(noRecords) noRecords.style.display = 'none';

    defaultingClients.forEach(item => {
      const { periodId, record } = item;
      const hasBills = record.contasAtrasadas && record.contasAtrasadas.length > 0;
      const totalDevido = hasBills ? record.contasAtrasadas.reduce((sum, conta) => sum + conta.valor, 0) : 0;

      const card = document.createElement('div');
      card.className = 'white-panel shadow';
      card.style.padding = '20px';
      card.style.display = 'flex';
      card.style.flexDirection = 'column';

      let bodyContent = '';
      if (!hasBills) {
        bodyContent = `
          <div style="padding: 24px; text-align: center; color: #6b7280; background: #f9fafb; border-radius: 8px; font-size: 0.9rem;">
            <i class="fa-solid fa-file-invoice" style="display: block; font-size: 1.5rem; margin-bottom: 8px; opacity: 0.5;"></i>
            Nenhuma cobrança específica cadastrada.
          </div>
        `;
      } else {
        let rowsHtml = record.contasAtrasadas.map(conta => {
          let displayDate = conta.dataEmissao;
          if (displayDate && displayDate.includes('-')) {
            displayDate = displayDate.split('-').reverse().join('/');
          }
           return `
            <tr>
              <td style="padding: 8px 0; border-bottom: 1px solid #e5e7eb;">${displayDate || '-'}</td>
              <td style="padding: 8px 0; border-bottom: 1px solid #e5e7eb; font-weight: 500;">R$ ${conta.valor.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
              <td style="padding: 8px 0; border-bottom: 1px solid #e5e7eb; text-align: right;">
                <button class="btn-icon" onclick="editContaAtrasada('${periodId}', '${record.id}', '${conta.id}')" title="Editar" style="color: var(--orange); margin-right: 8px;">
                  <i class="fa-solid fa-pencil"></i>
                </button>
                <button class="btn-icon text-red" onclick="deleteContaAtrasada('${periodId}', '${record.id}', '${conta.id}')" title="Excluir">
                  <i class="fa-solid fa-trash"></i>
                </button>
              </td>
            </tr>
          `;
        }).join('');

        bodyContent = `
          <table style="width: 100%; border-collapse: collapse; font-size: 0.95rem;">
            <thead>
              <tr>
                <th style="padding-bottom: 8px; text-align: left; color: #6b7280; font-weight: 600; font-size: 0.85rem;">Data Emissão</th>
                <th style="padding-bottom: 8px; text-align: left; color: #6b7280; font-weight: 600; font-size: 0.85rem;">Valor</th>
                <th style="padding-bottom: 8px; text-align: right; color: #6b7280; font-weight: 600; font-size: 0.85rem;">Ações</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
            </tbody>
          </table>
        `;
      }

      card.innerHTML = `
        <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 16px; padding-bottom: 16px; border-bottom: 1px solid #e5e7eb;">
          <h4 style="margin: 0; color: #1f2937; font-size: 1.1rem;">${record.clientName}</h4>
          <div style="text-align: right;">
            <span style="font-size: 0.8rem; color: #6b7280; text-transform: uppercase; font-weight: 600;">Total Devido</span>
            <div style="color: #ef4444; font-weight: 700; font-size: 1.25rem;">R$ ${totalDevido.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</div>
          </div>
        </div>
        ${bodyContent}
      `;
      grid.appendChild(card);
    });
  }
}

window.deleteContaAtrasada = async function(periodId, recordId, contaId) {
  if (confirm("Deseja realmente excluir esta conta atrasada?")) {
    const tempState = [...state.periods];
    const pIdx = tempState.findIndex(p => p.id === periodId);
    if (pIdx > -1) {
      const rIdx = tempState[pIdx].records.findIndex(r => r.id === recordId);
      if (rIdx > -1) {
        tempState[pIdx].records[rIdx].contasAtrasadas = tempState[pIdx].records[rIdx].contasAtrasadas.filter(c => c.id !== contaId);
        state.periods = tempState;
        if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
          await postToGoogleSheets({ action: 'delete_conta_atrasada', periodId, recordId, contaId });
        }
      }
    }
  }
};

window.editContaAtrasada = async function(periodId, recordId, contaId) {
  const tempState = [...state.periods];
  const pIdx = tempState.findIndex(p => p.id === periodId);
  if (pIdx > -1) {
    const rIdx = tempState[pIdx].records.findIndex(r => r.id === recordId);
    if (rIdx > -1) {
       const conta = tempState[pIdx].records[rIdx].contasAtrasadas.find(c => c.id === contaId);
       if (conta) {
          const newData = prompt("Data de Emissão (AAAA-MM-DD ou DD/MM/AAAA):", conta.dataEmissao);
          if (newData === null) return;
          const newValStr = prompt("Valor (apenas números e ponto/vírgula):", conta.valor);
          if (newValStr === null) return;
          const newVal = parseFloat(newValStr.toString().replace(',', '.'));
          if (!isNaN(newVal)) {
            conta.dataEmissao = newData;
            conta.valor = newVal;
            state.periods = tempState;
            if (!GOOGLE_APPS_SCRIPT_URL.includes("COLE_SUA_URL")) {
               await postToGoogleSheets({ action: 'edit_conta_atrasada', periodId, recordId, contaId, dataEmissao: newData, valor: newVal });
            }
          } else {
            alert("Valor numérico inválido!");
          }
       }
    }
  }
};
