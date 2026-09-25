// Global System State
let currentModule = 'dashboard';

/* Helper Utilities & Toast System */
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function escapeJsString(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/"/g, '\\"')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r');
}

function setElementText(id, text) {
  const el = document.getElementById(id);
  if (el) el.innerText = text;
}

function setElementValue(id, val) {
  const el = document.getElementById(id);
  if (el) el.value = val;
}

function setElementHTML(id, html) {
  const el = document.getElementById(id);
  if (el) el.innerHTML = html;
}

function showToast(message, type = 'info', duration = null) {
  const container = document.getElementById('toast-container');
  if (!container) {
    console.log(`[Toast ${type}]:`, message);
    return;
  }

  // Se for erro, o padrão é NÃO fechar sozinho (permanece na tela até o usuário fechar no X)
  const isError = (type === 'error' || type === 'danger');
  const actualDuration = (duration !== null) ? duration : (isError ? 0 : 4000);

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  
  let iconClass = 'fa-info-circle';
  if (type === 'success') iconClass = 'fa-circle-check';
  else if (isError) iconClass = 'fa-circle-exclamation';
  else if (type === 'warning') iconClass = 'fa-triangle-exclamation';

  toast.innerHTML = `
    <i class="fa-solid ${iconClass}" style="margin-top: 3px; font-size: 1.1rem; flex-shrink: 0;"></i>
    <div class="toast-body" style="flex: 1; font-size: 0.88rem; line-height: 1.4; white-space: pre-wrap; word-break: break-word;">
      ${escapeHtml(message)}
    </div>
    <button type="button" class="toast-close-btn" title="Fechar notificação" onclick="dismissToast(this)">
      <i class="fa-solid fa-xmark"></i>
    </button>
  `;

  container.appendChild(toast);

  if (actualDuration > 0) {
    let timer = setTimeout(() => {
      dismissToastEl(toast);
    }, actualDuration);

    toast.addEventListener('mouseenter', () => clearTimeout(timer));
    toast.addEventListener('mouseleave', () => {
      timer = setTimeout(() => dismissToastEl(toast), 2000);
    });
  }
}

function dismissToast(btn) {
  const toast = btn ? btn.closest('.toast') : null;
  if (toast) dismissToastEl(toast);
}

function dismissToastEl(toast) {
  if (!toast || toast.classList.contains('fade-out')) return;
  toast.classList.add('fade-out');
  setTimeout(() => {
    if (toast && toast.parentNode) toast.parentNode.removeChild(toast);
  }, 300);
}

/* ==================== THEME (DIA / NOITE) SYSTEM ==================== */
function initTheme() {
  const savedTheme = localStorage.getItem('theme_preference') || 
    (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  applyTheme(savedTheme, false);
}

function applyTheme(theme, showNotification = true) {
  const isDark = theme === 'dark';
  document.documentElement.setAttribute('data-theme', isDark ? 'dark' : 'light');
  localStorage.setItem('theme_preference', isDark ? 'dark' : 'light');

  const iconEl = document.getElementById('theme-toggle-icon');
  const labelEl = document.getElementById('theme-toggle-label');
  const toggleBtn = document.getElementById('btn-theme-toggle');

  if (iconEl) {
    iconEl.className = isDark ? 'fa-solid fa-sun' : 'fa-solid fa-moon';
  }
  if (labelEl) {
    labelEl.innerText = isDark ? 'Dia' : 'Noite';
  }
  if (toggleBtn) {
    toggleBtn.setAttribute('title', isDark ? 'Alternar para Modo Dia (Claro)' : 'Alternar para Modo Noite (Escuro)');
  }

  if (showNotification) {
    showToast(isDark ? 'Modo Noite ativado 🌙' : 'Modo Dia ativado ☀️', 'info', 2000);
  }
}

function toggleTheme() {
  const currentTheme = document.documentElement.getAttribute('data-theme') || 'light';
  const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
  applyTheme(newTheme, true);
}

/* ==================== SIDEBAR (MENU LATERAL) COLLAPSE / EXPAND ==================== */
function initSidebar() {
  const savedSidebar = localStorage.getItem('sidebar_collapsed');
  applySidebarState(savedSidebar === 'true', false);
}

function applySidebarState(isCollapsed, notify = false) {
  if (isCollapsed) {
    document.documentElement.classList.add('sidebar-collapsed');
    localStorage.setItem('sidebar_collapsed', 'true');
  } else {
    document.documentElement.classList.remove('sidebar-collapsed');
    localStorage.setItem('sidebar_collapsed', 'false');
  }

  const asideIcon = document.getElementById('sidebar-toggle-icon-aside');
  if (asideIcon) {
    asideIcon.className = isCollapsed ? 'fa-solid fa-angles-right' : 'fa-solid fa-angles-left';
  }
  const topbarBtn = document.getElementById('btn-sidebar-toggle');
  if (topbarBtn) {
    topbarBtn.setAttribute('title', isCollapsed ? 'Expandir Menu Lateral' : 'Recolher Menu Lateral');
  }

  // Trigger resize event for dynamic layout recalculations
  window.dispatchEvent(new Event('resize'));

  if (notify) {
    showToast(isCollapsed ? 'Menu lateral recolhido' : 'Menu lateral expandido', 'info', 1500);
  }
}

function toggleSidebar() {
  const isCurrentlyCollapsed = document.documentElement.classList.contains('sidebar-collapsed');
  applySidebarState(!isCurrentlyCollapsed, true);
}

// Entidades State
let entCurrentPage = 1;
let entTotalPages = 1;
let entFilters = { tipo: 'all', ativo: '1', q: '' };

// Produtos State
let prdCurrentPage = 1;
let prdTotalPages = 1;
let prdFilters = { grupo: 'all', ativo: '1', estoque_baixo: 'all', q: '' };

// Grupos Cache & State
let groupsList = [];
let gruSearchQuery = '';

function onReady(fn) {
  if (document.readyState === 'complete' || document.readyState === 'interactive') {
    setTimeout(fn, 1);
  } else {
    document.addEventListener('DOMContentLoaded', fn);
  }
}

onReady(() => {
  initTheme();
  initSidebar();
  fetchStats();
  loadGroupOptions();
  loadActiveEmpresasSelector();
  setupPdvKeyboardNavigation();

  // Search input debounce listeners
  setupDebounce('input-search-ent', (val) => {
    entFilters.q = val;
    entCurrentPage = 1;
    fetchEntities();
  });

  setupDebounce('input-search-prd', (val) => {
    prdFilters.q = val;
    prdCurrentPage = 1;
    fetchProducts();
  });

  setupDebounce('input-search-gru', (val) => {
    gruSearchQuery = val;
    fetchGroups();
  });

  setupDebounce('rel-filtro-busca', (val) => {
    relFiltros.q = val;
    relCurrentPage = 1;
    fetchRelatorioVendasProduto(1);
  });
});

function setupDebounce(inputId, callback) {
  let timeout;
  const el = document.getElementById(inputId);
  if (el) {
    el.addEventListener('input', (e) => {
      clearTimeout(timeout);
      timeout = setTimeout(() => callback(e.target.value), 300);
    });
  }
}

// Sidebar Dropdown Submenu Toggle
function toggleSidebarSubmenu(menuKey) {
  const group = document.getElementById(`nav-group-${menuKey}`);
  if (!group) return;

  const isCollapsed = document.documentElement.classList.contains('sidebar-collapsed');
  if (isCollapsed) {
    applySidebarState(false, false);
    group.classList.add('open');
    return;
  }

  group.classList.toggle('open');
}

// Module Navigation Switcher
function switchModule(moduleName) {
  currentModule = moduleName;

  // Update Nav Items & Sub-items active class
  document.querySelectorAll('.nav-item, .nav-sub-item').forEach(item => {
    item.classList.toggle('active', item.id === `nav-${moduleName}`);
  });

  // Auto-expand parent dropdown if navigating to a report sub-item
  const isReportModule = moduleName.startsWith('rel-');
  const relGroup = document.getElementById('nav-group-relatorios');
  const parentToggle = document.getElementById('nav-relatorios-toggle');

  if (isReportModule) {
    if (relGroup) relGroup.classList.add('open');
    if (parentToggle) parentToggle.classList.add('parent-active');
  } else {
    if (parentToggle) parentToggle.classList.remove('parent-active');
  }

  // Update View Sections active class
  document.querySelectorAll('.module-view').forEach(view => {
    view.classList.toggle('active', view.id === `view-${moduleName}`);
  });

  // Update Header Action Buttons & Title
  const titleEl = document.getElementById('page-title');
  if (moduleName === 'dashboard') {
    titleEl.innerHTML = `<i class="fa-solid fa-chart-pie"></i> <span>Dashboard - Visão Geral do Sistema</span>`;
  } else if (moduleName === 'pdv') {
    titleEl.innerHTML = `<i class="fa-solid fa-cash-register"></i> <span>PDV - Frente de Caixa</span>`;
    loadPdvPhotoCatalog();
    setTimeout(() => {
      const bc = document.getElementById('pdv-barcode-input');
      if (bc) bc.focus();
    }, 100);
  } else if (moduleName === 'entidades') {
    titleEl.innerHTML = `<i class="fa-solid fa-users"></i> <span>Cadastro de Entidades</span>`;
  } else if (moduleName === 'produtos') {
    titleEl.innerHTML = `<i class="fa-solid fa-boxes-stacked"></i> <span>Cadastro de Produtos</span>`;
  } else if (moduleName === 'grupos') {
    titleEl.innerHTML = `<i class="fa-solid fa-layer-group"></i> <span>Grupos de Produtos</span>`;
  } else if (moduleName === 'empresas') {
    titleEl.innerHTML = `<i class="fa-solid fa-building"></i> <span>Cadastro e Alteração de Empresas</span>`;
  } else if (moduleName === 'pedidos') {
    titleEl.innerHTML = `<i class="fa-solid fa-file-invoice-dollar"></i> <span>Emissão de Pedidos de Vendas (FrmPedidos.frm)</span>`;
  } else if (moduleName === 'cfop') {
    titleEl.innerHTML = `<i class="fa-solid fa-file-contract"></i> <span>Tabela CFOP (CFO) - Naturezas de Operação</span>`;
  } else if (moduleName === 'formas-pgto') {
    titleEl.innerHTML = `<i class="fa-solid fa-credit-card"></i> <span>Formas e Condições de Pagamento</span>`;
  } else if (moduleName === 'natureza-op') {
    titleEl.innerHTML = `<i class="fa-solid fa-route"></i> <span>Cadastro de Naturezas de Operação</span>`;
  } else if (moduleName === 'usuarios') {
    titleEl.innerHTML = `<i class="fa-solid fa-user-shield"></i> <span>Gestão de Usuários & Perfis de Acesso</span>`;
  } else if (moduleName === 'config-db') {
    titleEl.innerHTML = `<i class="fa-solid fa-database"></i> <span>Configuração e Diagnóstico do Banco PostgreSQL 18</span>`;
  } else if (moduleName === 'rel-vendas-produto') {
    titleEl.innerHTML = `<i class="fa-solid fa-chart-column"></i> <span>Relatório de Vendas Agrupado por Produto</span>`;
  }

  document.querySelectorAll('.module-btn').forEach(btn => {
    let show = (btn.id === `btn-new-${moduleName}` || btn.id === `btn-new-${moduleName.slice(0, -1)}`);
    if (moduleName === 'rel-vendas-produto' && (btn.id === 'btn-print-rel-vendas-produto' || btn.id === 'btn-export-rel-vendas-produto')) {
      show = true;
    }
    btn.style.display = show ? 'inline-flex' : 'none';
  });

  // Refresh stats & load module data
  fetchStats();
  if (moduleName === 'entidades') fetchEntities();
  else if (moduleName === 'produtos') {
    loadGroupOptions();
    fetchProducts();
  } else if (moduleName === 'grupos') fetchGroups();
  else if (moduleName === 'empresas') fetchEmpresas();
  else if (moduleName === 'pedidos') fetchOrders();
  else if (moduleName === 'rel-vendas-produto') {
    loadGroupOptions();
    fetchRelatorioVendasProduto(1);
  } else if (moduleName === 'cfop') fetchCfops();
  else if (moduleName === 'formas-pgto') fetchFormasPgto();
  else if (moduleName === 'natureza-op') fetchNaturezasOperacao();
  else if (moduleName === 'usuarios') fetchUsuarios();
  else if (moduleName === 'config-db') {
    loadDbConfig();
    loadPdvConfigAll();
  }
}

function switchModuleWithEstoque() {
  prdFilters.estoque_baixo = '1';
  document.getElementById('select-estoque-prd').value = '1';
  switchModule('produtos');
}

function refreshCurrentModule() {
  fetchStats();
  if (currentModule === 'entidades') fetchEntities();
  else if (currentModule === 'produtos') fetchProducts();
  else if (currentModule === 'grupos') fetchGroups();
  else if (currentModule === 'empresas') fetchEmpresas();
  else if (currentModule === 'pedidos') fetchOrders();
  else if (currentModule === 'rel-vendas-produto') fetchRelatorioVendasProduto(relCurrentPage);
}

// Stats & Dashboard Loader
async function fetchStats() {
  try {
    const res = await fetch('/api/stats');
    const data = await res.json();
    if (res.ok) {
      // Entidades
      if (data.entidades) {
        setElementText('stat-ent-total', data.entidades.total.toLocaleString());
        setElementText('stat-ent-ativos', data.entidades.ativos.toLocaleString());
        setElementText('stat-ent-clientes', data.entidades.clientes.toLocaleString());
        setElementText('stat-ent-fornecedores', data.entidades.fornecedores.toLocaleString());
        setElementText('dash-ent-total', data.entidades.total.toLocaleString());
      }
      // Produtos
      if (data.produtos) {
        setElementText('stat-prd-total', data.produtos.total.toLocaleString());
        setElementText('stat-prd-ativos', data.produtos.ativos.toLocaleString());
        setElementText('stat-prd-baixo', data.produtos.estoque_baixo.toLocaleString());
        setElementText('dash-prd-total', data.produtos.total.toLocaleString());
        setElementText('dash-prd-baixo', data.produtos.estoque_baixo.toLocaleString());

        const dashValor = document.getElementById('dash-valor-estoque');
        if (dashValor && data.produtos.valor_total_estoque !== undefined) {
          dashValor.innerText = `R$ ${data.produtos.valor_total_estoque.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
        }
      }
      // Grupos
      if (data.grupos) {
        setElementText('stat-prd-grupos', data.grupos.total.toLocaleString());
        setElementText('dash-gru-total', data.grupos.total.toLocaleString());
      }

      // Dashboard Widgets
      if (data.dashboard) {
        renderDashboardWidgets(data.dashboard);
      }
    }
  } catch (err) {
    console.error('Erro ao buscar estatísticas:', err);
  }
}

function renderDashboardWidgets(dashData) {
  // Top Groups
  const topGruposEl = document.getElementById('dash-top-grupos');
  if (topGruposEl && dashData.top_grupos && dashData.top_grupos.length) {
    const maxCount = Math.max(...dashData.top_grupos.map(g => (g.TotalProdutos !== undefined ? g.TotalProdutos : (g.totalprodutos || 0))), 1);
    topGruposEl.innerHTML = dashData.top_grupos.map(g => {
      const cnt = g.TotalProdutos !== undefined ? g.TotalProdutos : (g.totalprodutos || 0);
      const name = g.Nome_Grupo || g.nome_grupo || 'Sem Nome';
      const pct = Math.min(100, Math.round((cnt / maxCount) * 100));
      return `
        <div class="top-group-row">
          <div class="top-group-info">
            <span>${escapeHtml(name)}</span>
            <span style="color: var(--accent-purple);">${cnt} produtos</span>
          </div>
          <div class="group-progress-bar-bg">
            <div class="group-progress-bar-fill" style="width: ${pct}%;"></div>
          </div>
        </div>
      `;
    }).join('');
  } else if (topGruposEl) {
    topGruposEl.innerHTML = '<div style="color: var(--text-muted); padding: 1rem; text-align: center;">Nenhum grupo com produtos encontrado.</div>';
  }

  // Recent Products
  const recPrdEl = document.getElementById('dash-recentes-produtos');
  if (recPrdEl && dashData.ultimos_produtos && dashData.ultimos_produtos.length) {
    recPrdEl.innerHTML = dashData.ultimos_produtos.map(p => {
      const cod = p.CodPrd || p.codprd || p.Codigo || '';
      const desc = p.Descricao_Produto || p.descricao_produto || '';
      const preco = floatOrZero(p.Venda !== undefined ? p.Venda : p.venda);
      const est = floatOrZero(p.Estoque !== undefined ? p.Estoque : p.estoque);
      return `
        <tr>
          <td><strong>#${cod}</strong></td>
          <td>${escapeHtml(desc)}</td>
          <td style="text-align: right; font-weight: 600; color: var(--accent-emerald);">R$ ${preco.toFixed(2)}</td>
          <td style="text-align: right;">${est}</td>
        </tr>
      `;
    }).join('');
  } else if (recPrdEl) {
    recPrdEl.innerHTML = '<tr><td colspan="4" class="loading-td">Nenhum produto cadastrado.</td></tr>';
  }

  // Recent Entities
  const recEntEl = document.getElementById('dash-recentes-entidades');
  if (recEntEl && dashData.ultimas_entidades && dashData.ultimas_entidades.length) {
    recEntEl.innerHTML = dashData.ultimas_entidades.map(e => {
      const cod = e.CodEntidade || e.codentidade || '';
      const nome = e.Nome || e.nome || '';
      const cidade = e.Cidade || e.cidade || '-';
      const tipo = (e.Tipo || e.tipo) === 3 ? 'Vend' : ((e.Tipo || e.tipo) === 2 ? 'Forn' : 'Cli');
      return `
        <tr>
          <td><strong>#${cod}</strong></td>
          <td>${escapeHtml(nome)}</td>
          <td>${escapeHtml(cidade)}</td>
          <td><span class="badge badge-cliente">${tipo}</span></td>
        </tr>
      `;
    }).join('');
  } else if (recEntEl) {
    recEntEl.innerHTML = '<tr><td colspan="4" class="loading-td">Nenhuma entidade cadastrada.</td></tr>';
  }
}



// ==========================================
// 1. ENTIDADES MODULE LOGIC
// ==========================================

async function fetchEntities() {
  const tbody = document.getElementById('table-body-entidades');
  tbody.innerHTML = `
    <tr>
      <td colspan="9" class="loading-td">
        <i class="fa-solid fa-spinner fa-spin fa-2x"></i>
        <div style="margin-top: 0.5rem;">Carregando entidades...</div>
      </td>
    </tr>
  `;

  try {
    const queryParams = new URLSearchParams({
      page: entCurrentPage,
      limit: 15,
      tipo: entFilters.tipo,
      ativo: entFilters.ativo,
      q: entFilters.q
    });

    const res = await fetch(`/api/entidades?${queryParams.toString()}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Falha ao carregar');

    entTotalPages = data.pages;
    renderEntitiesTable(data.items);
    renderEntitiesPagination(data.total, data.page, data.limit);
  } catch (err) {
    tbody.innerHTML = `
      <tr>
        <td colspan="9" style="text-align: center; color: var(--accent-rose); padding: 2rem;">
          <i class="fa-solid fa-circle-exclamation fa-2x"></i>
          <div style="margin-top: 0.5rem;">Erro ao carregar dados: ${escapeHtml(err.message)}</div>
        </td>
      </tr>
    `;
  }
}

function renderEntitiesTable(entities) {
  const tbody = document.getElementById('table-body-entidades');
  if (!entities || entities.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="9" class="loading-td">Nenhuma entidade encontrada.</td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = entities.map(e => {
    const doc = e.CPF || e.CGC || '-';
    const cityUf = [e.Cidade, e.Uf].filter(Boolean).join(' / ') || '-';
    const phone = e.Fone || e.Celular || '-';

    let tipoBadge = '<span class="badge badge-outro">Outro</span>';
    if (e.Tipo === 1) tipoBadge = '<span class="badge badge-cliente">Cliente</span>';
    else if (e.Tipo === 2) tipoBadge = '<span class="badge badge-fornecedor">Fornecedor</span>';
    else if (e.Tipo === 3) tipoBadge = '<span class="badge badge-vendedor">Vendedor</span>';
    else if (e.Tipo === 4) tipoBadge = '<span class="badge badge-transp">Transportadora</span>';
    else if (e.Tipo === 5) tipoBadge = '<span class="badge badge-cliente">Cli / Forn</span>';

    const isAtivo = e.Ativo !== 0;
    const ativoBadge = isAtivo
      ? '<span class="badge badge-active"><i class="fa-solid fa-check"></i> Ativo</span>'
      : '<span class="badge badge-inactive"><i class="fa-solid fa-xmark"></i> Inativo</span>';

    return `
      <tr>
        <td><strong>#${e.CodEntidade}</strong></td>
        <td><div style="font-weight: 600;">${escapeHtml(e.Nome || '')}</div></td>
        <td>${escapeHtml(e.Fantasia || '-')}</td>
        <td>${tipoBadge}</td>
        <td>${escapeHtml(doc)}</td>
        <td>${escapeHtml(cityUf)}</td>
        <td>${escapeHtml(phone)}</td>
        <td>${ativoBadge}</td>
        <td style="text-align: center;">
          <div style="display: flex; justify-content: center; gap: 0.35rem;">
            <button class="btn btn-secondary btn-sm" onclick="editEntity(${e.CodEntidade})" title="Editar">
              <i class="fa-solid fa-pen"></i>
            </button>
            <button class="btn btn-danger btn-sm" onclick="deleteEntity(${e.CodEntidade}, '${escapeJsString(e.Nome || '')}')" title="Excluir">
              <i class="fa-solid fa-trash"></i>
            </button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

function renderEntitiesPagination(total, page, limit) {
  const start = total === 0 ? 0 : (page - 1) * limit + 1;
  const end = Math.min(page * limit, total);

  setElementText('pagination-info-ent', `Exibindo ${start} a ${end} de ${total} registros`);
  setElementText('page-number-ent', `Página ${page} de ${entTotalPages || 1}`);

  const btnPrev = document.getElementById('btn-prev-ent');
  if (btnPrev) btnPrev.disabled = page <= 1;
  const btnNext = document.getElementById('btn-next-ent');
  if (btnNext) btnNext.disabled = page >= entTotalPages;
}

function onEntityFilterChange() {
  entFilters.tipo = document.getElementById('select-tipo-ent').value;
  entFilters.ativo = document.getElementById('select-ativo-ent').value;
  entCurrentPage = 1;
  fetchEntities();
}

function changeEntityPage(delta) {
  const newPage = entCurrentPage + delta;
  if (newPage >= 1 && newPage <= entTotalPages) {
    entCurrentPage = newPage;
    fetchEntities();
  }
}

function openEntityModal(data = null) {
  const form = document.getElementById('entity-form');
  form.reset();

  resetModalTabs('entity-modal');

  if (data) {
    document.getElementById('modal-title-ent').innerHTML = `<i class="fa-solid fa-user-pen"></i> Editar Entidade #${data.CodEntidade}`;
    Object.keys(data).forEach(key => {
      const field = form.elements[key];
      if (field) {
        if (field.type === 'checkbox') field.checked = data[key] !== 0;
        else field.value = data[key] !== null ? data[key] : '';
      }
    });
  } else {
    document.getElementById('modal-title-ent').innerHTML = `<i class="fa-solid fa-user-plus"></i> Nova Entidade`;
    document.getElementById('CodEntidade').value = '';
    document.getElementById('Ativo').checked = true;
  }
  document.getElementById('entity-modal').classList.add('active');
}

function closeEntityModal() {
  document.getElementById('entity-modal').classList.remove('active');
}

async function editEntity(id) {
  try {
    const res = await fetch(`/api/entidades/${id}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Erro ao carregar entidade');
    openEntityModal(data);
  } catch (err) {
    showToast(`Erro ao abrir edição: ${err.message}`, 'error');
  }
}

async function saveEntity(event) {
  event.preventDefault();
  const form = document.getElementById('entity-form');
  const formData = new FormData(form);
  const data = {};

  formData.forEach((value, key) => { data[key] = value; });
  data['Ativo'] = document.getElementById('Ativo').checked ? -1 : 0;

  ['Tipo', 'Classificacao', 'Status', 'DiaBomPgto'].forEach(k => {
    if (data[k] !== '' && data[k] !== undefined && data[k] !== null) {
      const v = parseInt(data[k], 10);
      data[k] = isNaN(v) ? null : v;
    } else {
      data[k] = null;
    }
  });
  ['Credito', 'Salario', 'Comissao'].forEach(k => {
    if (data[k] !== '' && data[k] !== undefined && data[k] !== null) {
      const v = parseFloat(String(data[k]).replace(',', '.'));
      data[k] = isNaN(v) ? null : v;
    } else {
      data[k] = null;
    }
  });

  const id = data['CodEntidade'];
  const url = id ? `/api/entidades/${id}` : '/api/entidades';
  const method = id ? 'PUT' : 'POST';

  try {
    const res = await fetch(url, {
      method: method,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao salvar');

    showToast(result.message || 'Salvo com sucesso!', 'success');
    closeEntityModal();
    fetchStats();
    fetchEntities();
  } catch (err) {
    showToast(`Erro ao salvar: ${err.message}`, 'error');
  }
}

async function deleteEntity(id, name) {
  if (!confirm(`Tem certeza que deseja excluir a entidade "${name}" (#${id})?`)) return;

  try {
    const res = await fetch(`/api/entidades/${id}`, { method: 'DELETE' });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao remover');

    showToast(result.message, 'success');
    fetchStats();
    fetchEntities();
  } catch (err) {
    showToast(`Erro ao excluir: ${err.message}`, 'error');
  }
}


// ==========================================
// 2. PRODUTOS MODULE LOGIC
// ==========================================

async function loadGroupOptions() {
  try {
    const res = await fetch('/api/grupos?limit=500');
    const data = await res.json();
    if (res.ok) {
      groupsList = data.items || [];
      
      // Populate Product Filter Dropdown
      const filterSelect = document.getElementById('select-grupo-prd');
      const filterVal = filterSelect.value;
      filterSelect.innerHTML = '<option value="all">Todos os Grupos</option>' + 
        groupsList.map(g => `<option value="${g.CodGru}">${escapeHtml(g.Descricao_Grupo || 'Sem Nome')}</option>`).join('');
      filterSelect.value = filterVal;

      // Populate Product Form Dropdown
      const formSelect = document.getElementById('Prd_Grupo');
      if (formSelect) {
        formSelect.innerHTML = '<option value="">Selecione um grupo...</option>' + 
          groupsList.map(g => `<option value="${g.CodGru}">${escapeHtml(g.Descricao_Grupo || 'Sem Nome')}</option>`).join('');
      }

      // Populate Relatorio Filter Dropdown
      const relGroupSelect = document.getElementById('rel-filtro-grupo');
      if (relGroupSelect) {
        const curVal = relGroupSelect.value;
        relGroupSelect.innerHTML = '<option value="all">Todos os Grupos</option>' + 
          groupsList.map(g => `<option value="${g.CodGru}">${escapeHtml(g.Descricao_Grupo || 'Sem Nome')}</option>`).join('');
        relGroupSelect.value = curVal || 'all';
      }
    }
  } catch (err) {
    console.error('Erro ao carregar opções de grupos:', err);
  }
}

async function fetchProducts() {
  const tbody = document.getElementById('table-body-produtos');
  tbody.innerHTML = `
    <tr>
      <td colspan="10" class="loading-td">
        <i class="fa-solid fa-spinner fa-spin fa-2x"></i>
        <div style="margin-top: 0.5rem;">Carregando produtos...</div>
      </td>
    </tr>
  `;

  try {
    const queryParams = new URLSearchParams({
      page: prdCurrentPage,
      limit: 15,
      grupo: prdFilters.grupo,
      ativo: prdFilters.ativo,
      estoque_baixo: prdFilters.estoque_baixo,
      q: prdFilters.q
    });

    const res = await fetch(`/api/produtos?${queryParams.toString()}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Falha ao carregar produtos');

    prdTotalPages = data.pages;
    renderProductsTable(data.items);
    renderProductsPagination(data.total, data.page, data.limit);
  } catch (err) {
    tbody.innerHTML = `
      <tr>
        <td colspan="10" style="text-align: center; color: var(--accent-rose); padding: 2rem;">
          <i class="fa-solid fa-circle-exclamation fa-2x"></i>
          <div style="margin-top: 0.5rem;">Erro ao carregar dados: ${escapeHtml(err.message)}</div>
        </td>
      </tr>
    `;
  }
}

function renderProductsTable(products) {
  const tbody = document.getElementById('table-body-produtos');
  if (!products || products.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="10" class="loading-td">Nenhum produto encontrado.</td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = products.map(p => {
    const custoStr = p.Custo !== null && p.Custo !== undefined ? `R$ ${p.Custo.toFixed(2)}` : 'R$ 0.00';
    const vendaStr = p.Venda !== null && p.Venda !== undefined ? `R$ ${p.Venda.toFixed(2)}` : 'R$ 0.00';
    const margemStr = p.Margem !== null && p.Margem !== undefined ? `${p.Margem.toFixed(1)}%` : '0%';

    // Group Name
    const groupName = p.Nome_Grupo || (groupsList.find(g => g.CodGru === p.Grupo)?.Descricao_Grupo) || `Grupo #${p.Grupo}`;
    const groupBadge = `<span class="badge badge-group">${escapeHtml(groupName)}</span>`;

    // Stock Badge
    const est = p.Estoque || 0;
    const min = p.Minimo || 0;
    let stockBadge = `<span class="badge badge-stock-ok">${est} ${escapeHtml(p.Embalagem || 'UN')}</span>`;
    if (est <= 0) {
      stockBadge = `<span class="badge badge-stock-out"><i class="fa-solid fa-triangle-exclamation"></i> Sem Estoque (${est})</span>`;
    } else if (est <= min) {
      stockBadge = `<span class="badge badge-stock-low"><i class="fa-solid fa-circle-exclamation"></i> Baixo (${est})</span>`;
    }

    // Status Badge
    const isAtivo = p.Ativo !== 0;
    const ativoBadge = isAtivo 
      ? '<span class="badge badge-active"><i class="fa-solid fa-check"></i> Ativo</span>'
      : '<span class="badge badge-inactive"><i class="fa-solid fa-xmark"></i> Inativo</span>';

    // Thumbnail da Foto do Produto
    const hasPhoto = isValidPhotoUrl(p.Foto);
    const photoThumbnail = hasPhoto
      ? `<div style="width: 38px; height: 38px; border-radius: var(--radius-sm); overflow: hidden; border: 1px solid var(--border-color); background: var(--bg-surface); display: flex; align-items: center; justify-content: center; cursor: pointer;" onclick="editProduct(${p.CodPrd})" title="Foto do produto - clique para editar">
           <img src="${escapeHtml(p.Foto)}" alt="" style="width: 100%; height: 100%; object-fit: cover;" onerror="this.onerror=null; this.parentElement.innerHTML='<i class=\\'fa-solid fa-image\\' style=\\'color: var(--text-muted); opacity: 0.5;\\'></i>';">
         </div>`
      : `<div style="width: 38px; height: 38px; border-radius: var(--radius-sm); border: 1px dashed var(--border-color); background: rgba(255,255,255,0.02); display: flex; align-items: center; justify-content: center; color: var(--text-muted);" title="Sem foto cadastrada">
           <i class="fa-solid fa-box" style="opacity: 0.35; font-size: 0.85rem;"></i>
         </div>`;

    return `
      <tr>
        <td><strong>#${p.CodPrd}</strong></td>
        <td style="text-align: center;">${photoThumbnail}</td>
        <td>
          <div style="font-weight: 600;">${escapeHtml(p.Descricao_Produto || '')}</div>
          ${p.Marca ? `<div style="font-size: 0.75rem; color: var(--text-muted);">${escapeHtml(p.Marca)}</div>` : ''}
        </td>
        <td>${groupBadge}</td>
        <td><code>${escapeHtml(p.CodBar || '-')}</code></td>
        <td style="text-align: right; font-family: monospace;">${custoStr}</td>
        <td style="text-align: right; font-family: monospace; font-weight: 600; color: var(--accent-emerald);">${vendaStr}</td>
        <td style="text-align: right; color: var(--accent-cyan); font-weight: 500;">${margemStr}</td>
        <td style="text-align: right;">${stockBadge}</td>
        <td>${ativoBadge}</td>
        <td style="text-align: center;">
          <div style="display: flex; justify-content: center; gap: 0.35rem;">
            <button class="btn btn-secondary btn-sm" onclick="editProduct(${p.CodPrd})" title="Editar">
              <i class="fa-solid fa-pen"></i>
            </button>
            <button class="btn btn-danger btn-sm" onclick="deleteProduct(${p.CodPrd}, '${escapeJsString(p.Descricao_Produto || '')}')" title="Excluir">
              <i class="fa-solid fa-trash"></i>
            </button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

function renderProductsPagination(total, page, limit) {
  const start = total === 0 ? 0 : (page - 1) * limit + 1;
  const end = Math.min(page * limit, total);

  setElementText('pagination-info-prd', `Exibindo ${start} a ${end} de ${total} registros`);
  setElementText('page-number-prd', `Página ${page} de ${prdTotalPages || 1}`);

  const btnPrev = document.getElementById('btn-prev-prd');
  if (btnPrev) btnPrev.disabled = page <= 1;
  const btnNext = document.getElementById('btn-next-prd');
  if (btnNext) btnNext.disabled = page >= prdTotalPages;
}

function onProductFilterChange() {
  prdFilters.grupo = document.getElementById('select-grupo-prd').value;
  prdFilters.ativo = document.getElementById('select-ativo-prd').value;
  prdFilters.estoque_baixo = document.getElementById('select-estoque-prd').value;
  prdCurrentPage = 1;
  fetchProducts();
}

function changeProductPage(delta) {
  const newPage = prdCurrentPage + delta;
  if (newPage >= 1 && newPage <= prdTotalPages) {
    prdCurrentPage = newPage;
    fetchProducts();
  }
}

// Auto margin & sale price calculation
function calcMargin() {
  const custo = parseFloat(document.getElementById('Prd_Custo').value) || 0;
  const venda = parseFloat(document.getElementById('Prd_Venda').value) || 0;
  if (custo > 0 && venda > 0) {
    const margin = ((venda - custo) / custo) * 100;
    document.getElementById('Prd_Margem').value = margin.toFixed(2);
  }
}

function calcSalePrice() {
  const custo = parseFloat(document.getElementById('Prd_Custo').value) || 0;
  const margem = parseFloat(document.getElementById('Prd_Margem').value) || 0;
  if (custo > 0) {
    const venda = custo * (1 + (margem / 100));
    document.getElementById('Prd_Venda').value = venda.toFixed(2);
  }
}

function openProductModal(data = null) {
  const form = document.getElementById('product-form');
  form.reset();
  resetModalTabs('product-modal');

  const fileInput = document.getElementById('Prd_FileInput');
  if (fileInput) fileInput.value = '';

  if (data) {
    document.getElementById('modal-title-prd').innerHTML = `<i class="fa-solid fa-pen-to-square"></i> Editar Produto #${data.CodPrd}`;
    Object.keys(data).forEach(key => {
      const field = form.elements[key];
      if (field) {
        if (field.type === 'checkbox') field.checked = data[key] !== 0;
        else field.value = data[key] !== null ? data[key] : '';
      }
    });
    setProductPhotoPreview(data.Foto || null);
  } else {
    document.getElementById('modal-title-prd').innerHTML = `<i class="fa-solid fa-box"></i> Novo Produto`;
    document.getElementById('Prd_CodPrd').value = '';
    document.getElementById('Prd_Ativo').checked = true;
    document.getElementById('Prd_Custo').value = '0.00';
    document.getElementById('Prd_Venda').value = '0.00';
    document.getElementById('Prd_Margem').value = '0.00';
    document.getElementById('Prd_Estoque').value = '0';
    document.getElementById('Prd_Minimo').value = '0';
    document.getElementById('Prd_Icm').value = '18.00';
    setProductPhotoPreview(null);
  }
  document.getElementById('product-modal').classList.add('active');
}

function closeProductModal() {
  document.getElementById('product-modal').classList.remove('active');
}

// Funções de Gestão de Foto de Produto
function triggerProductPhotoUpload() {
  const fileInput = document.getElementById('Prd_FileInput');
  if (fileInput) fileInput.click();
}

async function handleProductPhotoFileSelect(event) {
  const file = event.target.files[0];
  if (!file) return;
  await uploadProductPhotoFile(file);
}

async function uploadProductPhotoFile(file) {
  if (!file.type.startsWith('image/')) {
    showToast('Por favor, selecione um arquivo de imagem válido (JPG, PNG, WEBP, GIF, SVG).', 'warning');
    return;
  }

  showToast('Enviando foto do produto...', 'info');
  const formData = new FormData();
  formData.append('file', file);

  try {
    const res = await fetch('/api/produtos/upload-foto', {
      method: 'POST',
      body: formData
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Erro ao enviar foto');

    setProductPhotoPreview(data.url);
    showToast('Foto do produto carregada com sucesso!', 'success');
  } catch (err) {
    showToast(`Falha no upload da foto: ${err.message}`, 'error');
  }
}

function handleProductPhotoUrlInput(url) {
  setProductPhotoPreview(url, false);
}

function applyPhotoUrlManual() {
  const textInput = document.getElementById('Prd_Foto_Input_Text');
  const url = textInput ? textInput.value.trim() : '';
  setProductPhotoPreview(url);
  if (url) showToast('URL da foto aplicada!', 'info');
}

function clearProductPhoto() {
  setProductPhotoPreview(null);
  const fileInput = document.getElementById('Prd_FileInput');
  if (fileInput) fileInput.value = '';
  showToast('Foto removida.', 'info');
}

function isValidPhotoUrl(url) {
  if (!url || typeof url !== 'string') return false;
  const s = url.trim();
  if (s === '' || s === '0' || s === '-1' || s === '1' || s.toLowerCase() === 'null') return false;
  return s.startsWith('/') || s.startsWith('http://') || s.startsWith('https://') || s.startsWith('data:image/') || /\.(jpg|jpeg|png|webp|gif|svg|bmp)(\?.*)?$/i.test(s);
}

function setProductPhotoPreview(url, syncTextInput = true) {
  const cleanUrl = isValidPhotoUrl(url) ? url.trim() : '';
  
  const hiddenInput = document.getElementById('Prd_Foto');
  if (hiddenInput) hiddenInput.value = cleanUrl;

  const textInput = document.getElementById('Prd_Foto_Input_Text');
  if (textInput && syncTextInput) textInput.value = cleanUrl;

  const imgMini = document.getElementById('prd-photo-preview-mini');
  const placeholderMini = document.getElementById('prd-photo-placeholder-mini');
  const imgLarge = document.getElementById('prd-photo-preview-large');
  const placeholderLarge = document.getElementById('prd-photo-placeholder-large');

  if (cleanUrl) {
    if (imgMini) {
      imgMini.src = cleanUrl;
      imgMini.style.display = 'block';
      imgMini.onerror = () => { imgMini.style.display = 'none'; if (placeholderMini) placeholderMini.style.display = 'flex'; };
    }
    if (placeholderMini) placeholderMini.style.display = 'none';

    if (imgLarge) {
      imgLarge.src = cleanUrl;
      imgLarge.style.display = 'block';
      imgLarge.onerror = () => { imgLarge.style.display = 'none'; if (placeholderLarge) placeholderLarge.style.display = 'flex'; };
    }
    if (placeholderLarge) placeholderLarge.style.display = 'none';
  } else {
    if (imgMini) { imgMini.src = ''; imgMini.style.display = 'none'; }
    if (placeholderMini) placeholderMini.style.display = 'flex';

    if (imgLarge) { imgLarge.src = ''; imgLarge.style.display = 'none'; }
    if (placeholderLarge) placeholderLarge.style.display = 'flex';
  }
}

function handlePhotoDragOver(e) {
  e.preventDefault();
  const dz = document.getElementById('prd-dropzone');
  if (dz) dz.style.background = 'rgba(59,130,246,0.15)';
}

function handlePhotoDragLeave(e) {
  e.preventDefault();
  const dz = document.getElementById('prd-dropzone');
  if (dz) dz.style.background = 'rgba(59,130,246,0.05)';
}

function handlePhotoDrop(e) {
  e.preventDefault();
  const dz = document.getElementById('prd-dropzone');
  if (dz) dz.style.background = 'rgba(59,130,246,0.05)';

  if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
    uploadProductPhotoFile(e.dataTransfer.files[0]);
  }
}

async function editProduct(id) {
  try {
    const res = await fetch(`/api/produtos/${id}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Erro ao carregar produto');
    openProductModal(data);
  } catch (err) {
    showToast(`Erro ao abrir edição: ${err.message}`, 'error');
  }
}

async function saveProduct(event) {
  event.preventDefault();
  const form = document.getElementById('product-form');
  const formData = new FormData(form);
  const data = {};

  formData.forEach((value, key) => { data[key] = value; });

  data['Ativo'] = document.getElementById('Prd_Ativo').checked ? -1 : 0;
  data['Promocao'] = document.getElementById('Prd_Promocao').checked ? -1 : 0;
  
  const fotoVal = document.getElementById('Prd_Foto') ? document.getElementById('Prd_Foto').value.trim() : '';
  data['Foto'] = fotoVal || null;

  ['Grupo'].forEach(k => {
    if (data[k] !== '' && data[k] !== undefined) data[k] = parseInt(data[k], 10);
  });
  ['Custo', 'Venda', 'Margem', 'PcoPrazo', 'Comissao', 'Venda_Promocao', 'Estoque', 'Minimo', 'Est_Ideal', 'PesoLiquido', 'PesoBruto', 'Icm', 'Ipi', 'Aliquota'].forEach(k => {
    if (data[k] !== '' && data[k] !== undefined) data[k] = parseFloat(data[k]);
  });

  const id = data['CodPrd'];
  const url = id ? `/api/produtos/${id}` : '/api/produtos';
  const method = id ? 'PUT' : 'POST';

  try {
    const res = await fetch(url, {
      method: method,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao salvar');

    showToast(result.message || 'Produto salvo com sucesso!', 'success');
    closeProductModal();
    fetchStats();
    fetchProducts();
  } catch (err) {
    showToast(`Erro ao salvar: ${err.message}`, 'error');
  }
}

async function deleteProduct(id, name) {
  if (!confirm(`Tem certeza que deseja excluir o produto "${name}" (#${id})?`)) return;

  try {
    const res = await fetch(`/api/produtos/${id}`, { method: 'DELETE' });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao remover');

    showToast(result.message, 'success');
    fetchStats();
    fetchProducts();
  } catch (err) {
    showToast(`Erro ao excluir: ${err.message}`, 'error');
  }
}


// ==========================================
// 3. GRUPOS MODULE LOGIC
// ==========================================

async function fetchGroups() {
  const tbody = document.getElementById('table-body-grupos');
  tbody.innerHTML = `
    <tr>
      <td colspan="6" class="loading-td">
        <i class="fa-solid fa-spinner fa-spin fa-2x"></i>
        <div style="margin-top: 0.5rem;">Carregando grupos...</div>
      </td>
    </tr>
  `;

  try {
    const queryParams = new URLSearchParams({
      limit: 200,
      q: gruSearchQuery
    });

    const res = await fetch(`/api/grupos?${queryParams.toString()}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Falha ao carregar grupos');

    groupsList = data.items || [];
    renderGroupsTable(groupsList);
  } catch (err) {
    tbody.innerHTML = `
      <tr>
        <td colspan="6" style="text-align: center; color: var(--accent-rose); padding: 2rem;">
          <i class="fa-solid fa-circle-exclamation fa-2x"></i>
          <div style="margin-top: 0.5rem;">Erro ao carregar grupos: ${escapeHtml(err.message)}</div>
        </td>
      </tr>
    `;
  }
}

function renderGroupsTable(groups) {
  const tbody = document.getElementById('table-body-grupos');
  if (!groups || groups.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="6" class="loading-td">Nenhum grupo encontrado.</td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = groups.map(g => {
    const baixarBadge = g.BaixarEstoque !== 0
      ? '<span class="badge badge-active"><i class="fa-solid fa-check"></i> Sim</span>'
      : '<span class="badge badge-inactive"><i class="fa-solid fa-xmark"></i> Não</span>';

    const descVal = parseFloat(g.Desconto || 0);
    const descBadge = descVal > 0
      ? `<span class="badge badge-active"><i class="fa-solid fa-percent"></i> ${descVal.toFixed(2)}%</span>`
      : `<span class="badge badge-inactive">0.00%</span>`;

    return `
      <tr>
        <td><strong>#${g.CodGru}</strong></td>
        <td>
          <div style="font-weight: 600; color: var(--text-primary);">${escapeHtml(g.Descricao_Grupo || '')}</div>
        </td>
        <td>${g.Familia ? `Família #${g.Familia}` : '-'}</td>
        <td style="text-align: right;">${descBadge}</td>
        <td>${baixarBadge}</td>
        <td style="text-align: center;">
          <span class="badge badge-group"><i class="fa-solid fa-boxes-stacked"></i> ${g.TotalProdutos || 0} produtos</span>
        </td>
        <td style="text-align: center;">
          <div style="display: flex; justify-content: center; gap: 0.35rem;">
            <button class="btn btn-secondary btn-sm" onclick="editGroup(${g.CodGru})" title="Editar">
              <i class="fa-solid fa-pen"></i>
            </button>
            <button class="btn btn-danger btn-sm" onclick="deleteGroup(${g.CodGru}, '${escapeJsString(g.Descricao_Grupo || '')}')" title="Excluir">
              <i class="fa-solid fa-trash"></i>
            </button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

function openGroupModal(data = null) {
  const form = document.getElementById('group-form');
  form.reset();

  if (data) {
    document.getElementById('modal-title-gru').innerHTML = `<i class="fa-solid fa-layer-group"></i> Editar Grupo #${data.CodGru}`;
    document.getElementById('Gru_CodGru').value = data.CodGru;
    document.getElementById('Gru_Descricao_Grupo').value = data.Descricao_Grupo || '';
    document.getElementById('Gru_Familia').value = data.Familia || 1;
    document.getElementById('Gru_Desconto').value = data.Desconto ? parseFloat(data.Desconto).toFixed(2) : '0.00';
    document.getElementById('Gru_BaixarEstoque').checked = data.BaixarEstoque !== 0;
  } else {
    document.getElementById('modal-title-gru').innerHTML = `<i class="fa-solid fa-layer-group"></i> Novo Grupo`;
    document.getElementById('Gru_CodGru').value = '';
    document.getElementById('Gru_Familia').value = '1';
    document.getElementById('Gru_Desconto').value = '0.00';
    document.getElementById('Gru_BaixarEstoque').checked = true;
  }
  document.getElementById('group-modal').classList.add('active');
}

function closeGroupModal() {
  document.getElementById('group-modal').classList.remove('active');
}

async function editGroup(id) {
  try {
    const res = await fetch(`/api/grupos/${id}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Erro ao carregar grupo');
    openGroupModal(data);
  } catch (err) {
    showToast(`Erro ao abrir edição: ${err.message}`, 'error');
  }
}

async function saveGroup(event) {
  event.preventDefault();
  const form = document.getElementById('group-form');
  const formData = new FormData(form);
  const data = {};

  formData.forEach((value, key) => { data[key] = value; });

  data['BaixarEstoque'] = document.getElementById('Gru_BaixarEstoque').checked ? -1 : 0;
  if (data['Familia']) data['Familia'] = parseInt(data['Familia'], 10);
  data['Desconto'] = parseFloat(data['Desconto'] || 0.0);

  const id = data['CodGru'];
  const url = id ? `/api/grupos/${id}` : '/api/grupos';
  const method = id ? 'PUT' : 'POST';

  try {
    const res = await fetch(url, {
      method: method,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao salvar');

    showToast(result.message || 'Grupo salvo com sucesso!', 'success');
    closeGroupModal();
    loadGroupOptions();
    fetchStats();
    fetchGroups();
  } catch (err) {
    showToast(`Erro ao salvar: ${err.message}`, 'error');
  }
}

async function deleteGroup(id, name) {
  if (!confirm(`Tem certeza que deseja excluir o grupo "${name}" (#${id})?`)) return;

  try {
    const res = await fetch(`/api/grupos/${id}`, { method: 'DELETE' });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao remover');

    showToast(result.message, 'success');
    loadGroupOptions();
    fetchStats();
    fetchGroups();
  } catch (err) {
    showToast(`Erro ao excluir: ${err.message}`, 'error');
  }
}


// ==========================================
// UTILITY FUNCTIONS
// ==========================================

function switchTab(tabId, btnElement) {
  const modal = btnElement.closest('.modal-container');
  modal.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  modal.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

  btnElement.classList.add('active');
  const target = modal.querySelector(`#${tabId}`);
  if (target) target.classList.add('active');
}

function resetModalTabs(modalId) {
  const modal = document.getElementById(modalId);
  if (!modal) return;
  modal.querySelectorAll('.tab-btn').forEach((btn, idx) => {
    btn.classList.toggle('active', idx === 0);
  });
  modal.querySelectorAll('.tab-content').forEach((content, idx) => {
    content.classList.toggle('active', idx === 0);
  });
}

// ==========================================
// 4. PDV / FRENTE DE CAIXA LOGIC
// ==========================================

let pdvCart = [];
let pdvSelectedProduct = null;
let pdvCatalogProducts = [];
let pdvCurrentGroupFilter = 'all';
let pdvCurrentPhotoSearch = '';
let pdvViewMode = 'split';

onReady(() => {
  // Live POS Clock
  setInterval(updatePdvClock, 1000);
  updatePdvClock();
  loadPdvPhotoCatalog();




  // Global Hotkeys Listener (F2, F4, F8, F9, F10)
  document.addEventListener('keydown', (e) => {
    if (currentModule !== 'pdv') return;
    if (e.key === 'F2') {
      e.preventDefault();
      document.getElementById('pdv-barcode-input').focus();
    } else if (e.key === 'F4') {
      e.preventDefault();
      const cliInp = document.getElementById('pdv-client-search-input');
      if (cliInp) { cliInp.focus(); cliInp.select(); }
    } else if (e.key === 'F8') {
      e.preventDefault();
      cancelPdvSale();
    } else if (e.key === 'F9' || e.key === 'F10') {
      e.preventDefault();
      openPaymentModal();
    }
  });

  // Client Search Input Listener
  const clientInput = document.getElementById('pdv-client-search-input');
  if (clientInput) {
    clientInput.addEventListener('input', (e) => {
      debouncePdvClientSearch(e.target.value);
    });
    clientInput.addEventListener('focus', () => {
      debouncePdvClientSearch(clientInput.value);
    });
  }

  // Close overlays on outside click
  document.addEventListener('click', (e) => {
    if (!e.target.closest('#pdv-client-search-input') && !e.target.closest('#pdv-client-results-overlay')) {
      const overlay = document.getElementById('pdv-client-results-overlay');
      if (overlay) overlay.style.display = 'none';
    }
  });

  loadPdvEntities();
});

function updatePdvClock() {
  const clockEl = document.getElementById('pdv-clock-display');
  if (clockEl) {
    const now = new Date();
    clockEl.innerText = now.toLocaleString('pt-BR');
  }
}

let pdvClientSearchTimeout = null;
function debouncePdvClientSearch(val) {
  clearTimeout(pdvClientSearchTimeout);
  pdvClientSearchTimeout = setTimeout(() => searchPdvClients(val), 200);
}

async function searchPdvClients(q) {
  try {
    const res = await fetch(`/api/pdv/clientes?q=${encodeURIComponent(q || '')}`);
    const clients = await res.json();
    const overlay = document.getElementById('pdv-client-results-overlay');
    if (!res.ok || !overlay) return;

    if (!clients || !clients.length) {
      overlay.innerHTML = `<div style="padding: 0.75rem; color: var(--text-muted); font-size: 0.85rem;">Nenhum cliente encontrado</div>`;
      overlay.style.display = 'block';
      return;
    }

    overlay.innerHTML = `
      <div class="pos-result-item" onclick="selectPdvClient(1, 'CONSUMIDOR FINAL', '')">
        <div><strong>1 - CONSUMIDOR FINAL</strong></div>
      </div>
      ` + clients.filter(c => c.CodEntidade !== 1).map((c, idx) => {
        const displayName = c.Fantasia && c.Fantasia.toUpperCase() !== c.Nome.toUpperCase() ? `${c.Nome} (${c.Fantasia})` : c.Nome;
        return `
          <div class="pos-result-item" onclick="selectPdvClient(${c.CodEntidade}, '${escapeJsString(c.Nome)}', '${escapeJsString(c.CPF || c.CGC || '')}')">
            <div>
              <div class="pos-result-name">#${c.CodEntidade} - ${escapeHtml(displayName)}</div>
              <div class="pos-result-meta">${c.CPF || c.CGC ? 'CPF/CNPJ: ' + (c.CPF || c.CGC) : ''} ${c.Cidade ? ' | Cidade: ' + c.Cidade : ''}</div>
            </div>
          </div>
        `;
      }).join('');
    overlay.style.display = 'block';
  } catch (err) {
    console.error('Erro na busca de clientes:', err);
  }
}

function selectPdvClient(id, name, doc) {
  document.getElementById('pdv-select-entidade').value = id;
  const input = document.getElementById('pdv-client-search-input');
  if (input) {
    input.value = id === 1 ? '' : `#${id} - ${name}`;
    input.placeholder = id === 1 ? '1 - CONSUMIDOR FINAL (digite para buscar...)' : `#${id} - ${name}`;
  }
  const overlay = document.getElementById('pdv-client-results-overlay');
  if (overlay) overlay.style.display = 'none';
  showToast(`Cliente selecionado: ${name}`, 'success');
}

function resetPdvClient() {
  selectPdvClient(1, 'CONSUMIDOR FINAL', '');
}

async function loadPdvEntities() {
  try {
    // Vendedores (Tipo = 3)
    const resVend = await fetch('/api/pdv/vendedores');
    const vendedores = await resVend.json();
    const selectVend = document.getElementById('pdv-select-vendedor');
    if (selectVend && Array.isArray(vendedores)) {
      selectVend.innerHTML = '';
      if (!vendedores.length) {
        selectVend.innerHTML = `<option value="1">1 - OPERADOR CAIXA</option>`;
      } else {
        vendedores.forEach((v, idx) => {
          const isDefault = v.CodEntidade === 1 || v.CodEntidade === 2 || idx === 0;
          selectVend.innerHTML += `<option value="${v.CodEntidade}" ${isDefault ? 'selected' : ''}>${v.CodEntidade} - ${escapeHtml(v.Nome)}</option>`;
        });
      }
    }
  } catch (err) {
    console.error('Erro ao carregar vendedores no PDV:', err);
  }
}

let pdvSearchTimeout = null;
function debouncePdvSearch(val) {
  clearTimeout(pdvSearchTimeout);
  if (!val || val.trim().length < 2) {
    document.getElementById('pdv-results-overlay').style.display = 'none';
    return;
  }
  pdvSearchTimeout = setTimeout(() => searchPdvProducts(val), 250);
}

let pdvSearchResults = [];

async function searchPdvProducts(q) {
  try {
    const res = await fetch(`/api/pdv/produtos?q=${encodeURIComponent(q)}`);
    const results = await res.json();
    const overlay = document.getElementById('pdv-results-overlay');
    if (!res.ok || !results.length) {
      overlay.style.display = 'none';
      pdvSearchResults = [];
      return;
    }

    pdvSearchResults = results;
    overlay.innerHTML = results.map((p, idx) => {
      const descGrupo = floatOrZero(p.DescontoGrupo || p.descontogrupo || p.Desconto_Grupo);
      const descBadge = descGrupo > 0 ? `<span class="badge badge-active" style="margin-left: 6px;"><i class="fa-solid fa-percent"></i> ${descGrupo}% desc. grupo</span>` : '';
      const priceDisplay = descGrupo > 0 
        ? `<div style="text-align: right;"><s style="color: var(--text-muted); font-size: 0.8rem;">R$ ${(p.Venda || 0).toFixed(2)}</s><br><strong style="color: var(--accent-emerald);">R$ ${(p.PrecoComDesconto || (p.Venda * (1 - descGrupo/100))).toFixed(2)}</strong></div>`
        : `<strong style="color: var(--accent-emerald);">R$ ${(p.Venda || 0).toFixed(2)}</strong>`;

      return `
        <div class="pos-result-item ${idx === 0 ? 'selected' : ''}" onclick="selectPdvProductByIndex(${idx})">
          <div>
            <div class="pos-result-name">${escapeHtml(p.Descricao_Produto)} ${descBadge}</div>
            <div class="pos-result-meta">Cód: #${p.CodPrd} | EAN: ${p.CodBar || 'SEM EAN'} | Estoque: ${p.Estoque} ${p.Embalagem || 'UN'}</div>
          </div>
          ${priceDisplay}
        </div>
      `;
    }).join('');
    overlay.style.display = 'block';
  } catch (err) {
    console.error('Erro na busca de produtos PDV:', err);
  }
}

function setupPdvKeyboardNavigation() {
  const barcodeInp = document.getElementById('pdv-barcode-input');
  const qtyInp = document.getElementById('pdv-qty-input');
  const priceInp = document.getElementById('pdv-price-input');
  const addBtn = document.getElementById('pdv-add-btn');

  if (barcodeInp) {
    barcodeInp.addEventListener('input', (e) => {
      debouncePdvSearch(e.target.value);
    });

    barcodeInp.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const overlay = document.getElementById('pdv-results-overlay');
        if (overlay && overlay.style.display !== 'none') {
          const selectedItem = overlay.querySelector('.pos-result-item.selected') || overlay.querySelector('.pos-result-item');
          if (selectedItem) {
            selectedItem.click();
            return;
          }
        }
        if (qtyInp) {
          qtyInp.focus();
          qtyInp.select();
        }
      }
    });
  }

  if (qtyInp) {
    qtyInp.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        if (priceInp) {
          priceInp.focus();
          priceInp.select();
        } else {
          addPdvItemFromInput();
        }
      }
    });
    qtyInp.addEventListener('focus', () => {
      qtyInp.select();
    });
  }

  if (priceInp) {
    priceInp.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        addPdvItemFromInput();
      }
    });
    priceInp.addEventListener('focus', () => {
      priceInp.select();
    });
  }
}

function selectPdvProductByIndex(index) {
  if (pdvSearchResults && pdvSearchResults[index]) {
    const prd = pdvSearchResults[index];
    setElementValue('pdv-barcode-input', prd.Descricao_Produto);
    setElementValue('pdv-price-input', floatOrZero(prd.Venda).toFixed(2));
    const overlay = document.getElementById('pdv-results-overlay');
    if (overlay) overlay.style.display = 'none';
    pdvSelectedProduct = prd;
    const qtyInp = document.getElementById('pdv-qty-input');
    if (qtyInp) {
      qtyInp.focus();
      qtyInp.select();
    }
  }
}

function selectPdvProduct(codPrd, desc, price) {
  const found = (pdvSearchResults || []).find(p => p.CodPrd === codPrd);
  if (found) {
    pdvSelectedProduct = found;
  } else {
    pdvSelectedProduct = { CodPrd: codPrd, Descricao_Produto: desc, Venda: floatOrZero(price) };
  }
  setElementValue('pdv-barcode-input', desc);
  setElementValue('pdv-price-input', floatOrZero(price).toFixed(2));
  const overlay = document.getElementById('pdv-results-overlay');
  if (overlay) overlay.style.display = 'none';
  const qtyInp = document.getElementById('pdv-qty-input');
  if (qtyInp) {
    qtyInp.focus();
    qtyInp.select();
  }
}

async function addPdvItemFromInput() {
  const barcodeInput = document.getElementById('pdv-barcode-input');
  const qtyInput = document.getElementById('pdv-qty-input');
  const priceInput = document.getElementById('pdv-price-input');
  if (!barcodeInput) return;

  const val = barcodeInput.value.trim();
  const qty = parseFloat(qtyInput ? qtyInput.value : 1.0) || 1.0;
  const customPrice = (priceInput && priceInput.value.trim() !== '') ? parseFloat(priceInput.value) : null;

  if (!val) {
    showToast('Digite ou escaneie um produto', 'error');
    return;
  }

  const overlay = document.getElementById('pdv-results-overlay');
  if (overlay) overlay.style.display = 'none';

  if (pdvSelectedProduct && pdvSelectedProduct.CodPrd) {
    addPdvItemToCart(pdvSelectedProduct, qty, customPrice);
    pdvSelectedProduct = null;
    barcodeInput.value = '';
    if (qtyInput) qtyInput.value = '1';
    if (priceInput) priceInput.value = '';
    barcodeInput.focus();
    return;
  }

  try {
    const res = await fetch(`/api/pdv/produtos?q=${encodeURIComponent(val)}`);
    const results = await res.json();
    if (!res.ok || !results.length) {
      showToast('Produto não encontrado', 'error');
      return;
    }

    const prd = results[0];
    addPdvItemToCart(prd, qty, customPrice);
    
    // Reset inputs & refocus
    barcodeInput.value = '';
    if (qtyInput) qtyInput.value = '1';
    if (priceInput) priceInput.value = '';
    barcodeInput.focus();
  } catch (err) {
    showToast(`Erro ao buscar item: ${err.message}`, 'error');
  }
}

function addPdvItemToCart(product, qty = 1.0, customPrice = null) {
  const descGrupo = floatOrZero(product.DescontoGrupo || product.descontogrupo || product.Desconto_Grupo);
  const unitPrice = (customPrice !== null && !isNaN(customPrice) && customPrice >= 0)
    ? customPrice
    : floatOrZero(product.Venda);

  const existingIdx = pdvCart.findIndex(item => item.CodPrd === product.CodPrd && Math.abs(item.ValorUnit - unitPrice) < 0.001);

  if (existingIdx >= 0) {
    pdvCart[existingIdx].Qtd += qty;
    pdvCart[existingIdx].Valor = pdvCart[existingIdx].Qtd * pdvCart[existingIdx].ValorUnit;
    if (descGrupo > 0) {
      pdvCart[existingIdx].Desconto = Math.round(pdvCart[existingIdx].Valor * (descGrupo / 100.0) * 100) / 100;
    }
  } else {
    const itemTotal = qty * unitPrice;
    const autoDiscount = descGrupo > 0 ? Math.round(itemTotal * (descGrupo / 100.0) * 100) / 100 : 0.0;
    pdvCart.push({
      CodPrd: product.CodPrd,
      Descricao_Produto: product.Descricao_Produto,
      CodBar: product.CodBar || '',
      Embalagem: product.Embalagem || 'UN',
      Qtd: qty,
      ValorUnit: unitPrice,
      Valor: itemTotal,
      Desconto: autoDiscount,
      DescontoGrupo: descGrupo
    });
  }

  renderPdvCart();
  const descMsg = descGrupo > 0 ? ` (${descGrupo}% desc. grupo)` : '';
  showToast(`Item adicionado: ${product.Descricao_Produto} - R$ ${unitPrice.toFixed(2)}${descMsg}`, 'success');
}

function removePdvCartItem(index) {
  pdvCart.splice(index, 1);
  renderPdvCart();
}

function updatePdvItemQty(index, newQty) {
  const q = parseFloat(newQty);
  if (q > 0 && pdvCart[index]) {
    pdvCart[index].Qtd = q;
    pdvCart[index].Valor = q * pdvCart[index].ValorUnit;
    const descGrupo = floatOrZero(pdvCart[index].DescontoGrupo);
    if (descGrupo > 0) {
      pdvCart[index].Desconto = Math.round(pdvCart[index].Valor * (descGrupo / 100.0) * 100) / 100;
    }
  } else {
    pdvCart.splice(index, 1);
  }
  renderPdvCart();
}

function updatePdvItemPrice(index, newPrice) {
  const p = parseFloat(newPrice);
  if (!isNaN(p) && p >= 0 && pdvCart[index]) {
    pdvCart[index].ValorUnit = p;
    pdvCart[index].Valor = pdvCart[index].Qtd * p;
    const descGrupo = floatOrZero(pdvCart[index].DescontoGrupo);
    if (descGrupo > 0) {
      pdvCart[index].Desconto = Math.round(pdvCart[index].Valor * (descGrupo / 100.0) * 100) / 100;
    }
  }
  renderPdvCart();
}

function renderPdvCart() {
  const tbody = document.getElementById('pdv-cart-tbody');
  if (!pdvCart.length) {
    tbody.innerHTML = `
      <tr>
        <td colspan="7" class="loading-td" style="padding: 4rem !important;">
          <i class="fa-solid fa-cart-shopping fa-3x" style="color: var(--border-color-strong); margin-bottom: 0.75rem;"></i>
          <div>Caixa Aberto. Digite ou escaneie o código de barras do produto para iniciar.</div>
        </td>
      </tr>
    `;
    document.getElementById('pdv-items-count').innerText = '0 itens';
    calcPdvTotals();
    return;
  }

  tbody.innerHTML = pdvCart.map((item, idx) => {
    const descGrupo = floatOrZero(item.DescontoGrupo);
    const itemDesc = floatOrZero(item.Desconto);
    const descInfo = itemDesc > 0
      ? `<div style="font-size: 0.75rem; color: var(--accent-emerald); font-weight: 600; margin-top: 2px;"><i class="fa-solid fa-tag"></i> Desconto: -R$ ${itemDesc.toFixed(2)}${descGrupo > 0 ? ` (${descGrupo}% grupo)` : ''}</div>`
      : '';

    const finalVal = item.Valor - itemDesc;

    return `
      <tr>
        <td><strong>${idx + 1}</strong></td>
        <td>#${item.CodPrd}</td>
        <td>
          <strong>${escapeHtml(item.Descricao_Produto)}</strong>
          <div style="font-size: 0.75rem; color: var(--text-muted);">${item.CodBar ? 'EAN: ' + item.CodBar : ''}</div>
          ${descInfo}
        </td>
        <td style="text-align: right;">
          <input type="number" class="form-control" style="width: 75px; text-align: right; display: inline-block; padding: 0.25rem 0.4rem; font-size: 0.85rem;" value="${item.Qtd}" step="1" min="1" onchange="updatePdvItemQty(${idx}, this.value)">
        </td>
        <td style="text-align: right;">
          <input type="number" class="form-control" style="width: 85px; text-align: right; display: inline-block; padding: 0.25rem 0.4rem; font-size: 0.85rem;" value="${item.ValorUnit.toFixed(2)}" step="0.01" min="0" onchange="updatePdvItemPrice(${idx}, this.value)">
        </td>
        <td style="text-align: right; font-weight: 700; color: var(--accent-emerald);">
          ${itemDesc > 0 ? `<s style="color: var(--text-muted); font-size: 0.8rem;">R$ ${item.Valor.toFixed(2)}</s><br>R$ ${finalVal.toFixed(2)}` : `R$ ${item.Valor.toFixed(2)}`}
        </td>
        <td style="text-align: center;">
          <button class="btn btn-danger btn-sm" onclick="removePdvCartItem(${idx})"><i class="fa-solid fa-xmark"></i></button>
        </td>
      </tr>
    `;
  }).join('');

  document.getElementById('pdv-items-count').innerText = `${pdvCart.length} item(ns)`;
  calcPdvTotals();
}

function calcPdvTotals() {
  const subtotal = pdvCart.reduce((sum, item) => sum + item.Valor, 0.0);
  const itemsDiscountSum = pdvCart.reduce((sum, item) => sum + (item.Desconto || 0.0), 0.0);
  const descInput = document.getElementById('pdv-desconto-input');
  
  if (descInput) {
    descInput.value = itemsDiscountSum.toFixed(2);
  }

  const desconto = parseFloat(descInput ? descInput.value : 0.0) || 0.0;
  const total = Math.max(0.0, subtotal - desconto);

  setElementText('pdv-subtotal-display', `R$ ${subtotal.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
  setElementText('pdv-total-display', `R$ ${total.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
  setElementText('pay-modal-total-display', `R$ ${total.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
  
  calcPdvChange();
}

function calcPdvChange() {
  const subtotal = pdvCart.reduce((sum, item) => sum + item.Valor, 0.0);
  const descInput = document.getElementById('pdv-desconto-input');
  const desconto = parseFloat(descInput ? descInput.value : 0.0) || 0.0;
  const total = Math.max(0.0, subtotal - desconto);

  const valRecEl = document.getElementById('pay-valor-recebido');
  const valRec = parseFloat(valRecEl ? valRecEl.value : 0.0) || 0.0;
  const troco = Math.max(0.0, valRec - total);

  setElementValue('pay-troco-display', `R$ ${troco.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
}

function cancelPdvSale() {
  if (!pdvCart.length) return;
  if (!confirm('Deseja realmente cancelar a venda e limpar o carrinho de compras?')) return;
  pdvCart = [];
  document.getElementById('pdv-desconto-input').value = '0.00';
  renderPdvCart();
  showToast('Venda cancelada', 'error');
}

let pdvSelectedPayMethod = 'DINHEIRO';
let pdvSelectedPayMethodName = 'DINHEIRO';
let pdvSelectedPayMethodTipo = 'DINHEIRO';
let pdvAllFormasPgto = [];

async function loadPdvPaymentMethods() {
  try {
    const res = await fetch('/api/formas-pgto?all=1');
    if (!res.ok) return;
    pdvAllFormasPgto = await res.json();
    
    const select = document.getElementById('pay-select-outras-formas');
    if (!select || !Array.isArray(pdvAllFormasPgto)) return;

    select.innerHTML = '<option value="">-- Selecione outra forma de pagamento cadastrada --</option>';
    
    pdvAllFormasPgto.forEach(f => {
      if (f.Ativo === 0) return;
      const opt = document.createElement('option');
      opt.value = f.id_forma;
      const parcels = f.Parcelas && f.Parcelas > 1 ? ` (${f.Parcelas}x)` : '';
      const days = f.DiasEntreParcelas && f.DiasEntreParcelas > 0 ? ` [${f.DiasEntreParcelas} dias]` : '';
      opt.textContent = `${f.Nome || f.Codigo}${parcels}${days}`;
      opt.dataset.codigo = f.Codigo || '';
      opt.dataset.nome = f.Nome || f.Codigo || '';
      opt.dataset.tipo = f.Tipo || 'A_VISTA';
      select.appendChild(opt);
    });
  } catch (err) {
    console.error('Erro ao carregar formas de pagamento no PDV:', err);
  }
}

function selectPdvPaymentMethod(code, name, tipo) {
  pdvSelectedPayMethod = code;
  pdvSelectedPayMethodName = name;
  pdvSelectedPayMethodTipo = tipo;

  // Atualiza classes ativas nos cards principais
  document.querySelectorAll('.payment-methods-grid .pay-method-card').forEach(card => {
    card.classList.remove('active');
  });

  const activeCard = document.getElementById(`pay-card-${code}`);
  if (activeCard) {
    activeCard.classList.add('active');
  }

  // Reseta select de outras formas
  const select = document.getElementById('pay-select-outras-formas');
  if (select && activeCard) {
    select.value = '';
  }

  // Atualiza badge visual
  const badge = document.getElementById('pay-selected-badge');
  if (badge) {
    badge.innerHTML = `<i class="fa-solid fa-check"></i> ${escapeHtml(name)}`;
  }

  // Exibe/oculta campos de dinheiro e troco
  const isCash = tipo === 'DINHEIRO' || code === 'DINHEIRO' || code === 'A_VISTA';
  const cashInputs = document.getElementById('pay-cash-inputs');
  const changePanel = document.getElementById('pay-change-panel');
  if (cashInputs && changePanel) {
    cashInputs.style.display = isCash ? 'flex' : 'none';
    changePanel.style.display = isCash ? 'flex' : 'none';
  }

  if (isCash) {
    const valRecEl = document.getElementById('pay-valor-recebido');
    if (valRecEl) {
      valRecEl.focus();
      valRecEl.select();
    }
  }
}

function onOtherPaymentMethodChange(selectEl) {
  const selectedId = selectEl.value;
  if (!selectedId) {
    selectPdvPaymentMethod('DINHEIRO', 'DINHEIRO', 'DINHEIRO');
    return;
  }

  const selectedOpt = selectEl.options[selectEl.selectedIndex];
  const code = selectedOpt.dataset.codigo || selectedOpt.textContent;
  const name = selectedOpt.dataset.nome || selectedOpt.textContent;
  const tipo = selectedOpt.dataset.tipo || 'OUTRO';

  pdvSelectedPayMethod = code;
  pdvSelectedPayMethodName = name;
  pdvSelectedPayMethodTipo = tipo;

  // Desmarca os 4 cards principais
  document.querySelectorAll('.payment-methods-grid .pay-method-card').forEach(card => {
    card.classList.remove('active');
  });

  // Atualiza badge visual
  const badge = document.getElementById('pay-selected-badge');
  if (badge) {
    badge.innerHTML = `<i class="fa-solid fa-check"></i> ${escapeHtml(name)}`;
  }

  // Oculta troco se não for dinheiro
  const isCash = tipo === 'DINHEIRO' || code === 'DINHEIRO' || code === 'A_VISTA';
  const cashInputs = document.getElementById('pay-cash-inputs');
  const changePanel = document.getElementById('pay-change-panel');
  if (cashInputs && changePanel) {
    cashInputs.style.display = isCash ? 'flex' : 'none';
    changePanel.style.display = isCash ? 'flex' : 'none';
  }
}

function openPaymentModal() {
  if (!pdvCart.length) {
    showToast('O carrinho está vazio. Adicione produtos antes de finalizar.', 'error');
    return;
  }
  const subtotal = pdvCart.reduce((sum, item) => sum + item.Valor, 0.0);
  const desconto = parseFloat(document.getElementById('pdv-desconto-input').value) || 0.0;
  const total = Math.max(0.0, subtotal - desconto);

  const valRecEl = document.getElementById('pay-valor-recebido');
  if (valRecEl) {
    valRecEl.value = total.toFixed(2);
  }

  calcPdvTotals();
  loadPdvPaymentMethods();
  selectPdvPaymentMethod('DINHEIRO', 'DINHEIRO', 'DINHEIRO');

  document.getElementById('pdv-payment-modal').classList.add('active');
  if (valRecEl) {
    valRecEl.focus();
    valRecEl.select();
  }
}

function closePaymentModal() {
  document.getElementById('pdv-payment-modal').classList.remove('active');
}

let lastSaleData = null;

let isPdvSubmitting = false;

async function submitPdvSale(emitirNfceFlag = true) {
  if (isPdvSubmitting) {
    console.warn('Operação de finalização já em andamento. Aguarde...');
    return;
  }
  if (!pdvCart || !pdvCart.length) {
    showToast('O carrinho de compras está vazio!', 'warning');
    return;
  }

  // Button references
  const btnNfce = document.getElementById('btn-pdv-pay-nfce') || document.querySelector('#pdv-payment-modal .btn-emerald');
  const btnSimple = document.getElementById('btn-pdv-pay-simple') || document.querySelector('#pdv-payment-modal .btn-secondary[onclick*="false"]');
  const btnCancel = document.getElementById('btn-pdv-pay-cancel') || document.querySelector('#pdv-payment-modal .btn-secondary[onclick*="closePaymentModal"]');

  const origNfceHtml = btnNfce ? btnNfce.innerHTML : '';
  const origSimpleHtml = btnSimple ? btnSimple.innerHTML : '';

  // Lock UI & disable buttons to prevent duplicate sales/NFC-e
  isPdvSubmitting = true;
  if (btnNfce) {
    btnNfce.disabled = true;
    if (emitirNfceFlag) {
      btnNfce.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Validando e Emitindo NFC-e...';
    }
  }
  if (btnSimple) {
    btnSimple.disabled = true;
    if (!emitirNfceFlag) {
      btnSimple.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Processando Venda...';
    }
  }
  if (btnCancel) btnCancel.disabled = true;

  const subtotal = pdvCart.reduce((sum, item) => sum + item.Valor, 0.0);
  const desconto = parseFloat(document.getElementById('pdv-desconto-input').value) || 0.0;
  const total = Math.max(0.0, subtotal - desconto);

  const selectedEnt = parseInt(document.getElementById('pdv-select-entidade').value, 10) || 1;
  const selectedVend = parseInt(document.getElementById('pdv-select-vendedor').value, 10) || 1;

  const payMethod = pdvSelectedPayMethodName || pdvSelectedPayMethod || 'DINHEIRO';

  const payload = {
    CodEntidade: selectedEnt,
    Vendedor: selectedVend,
    CondPgto: payMethod,
    SubTotal: subtotal,
    Desconto: desconto,
    Total: total,
    EmitirNFCe: emitirNfceFlag,
    items: pdvCart
  };

  try {
    const res = await fetch('/api/pdv/venda', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const result = await res.json();
    if (!res.ok) {
      throw new Error(result.error || 'Erro ao validar e processar venda');
    }

    showToast(result.message || 'Venda finalizada com sucesso!', 'success');
    closePaymentModal();
    
    // Clear Cart
    pdvCart = [];
    document.getElementById('pdv-desconto-input').value = '0.00';
    document.getElementById('pay-valor-recebido').value = '0.00';
    renderPdvCart();

    // Render & Open Receipt Modal and trigger print
    if (result.data) {
      lastSaleData = result.data;
      renderThermalReceipt(result.data, emitirNfceFlag ? 'nfce' : 'venda', true);
      
      // Auto-trigger printing on successful validation and emission
      if (emitirNfceFlag) {
        printDanfeNfce(true);
      } else {
        printSimpleSaleReceipt(true);
      }
    }
  } catch (err) {
    showToast(`Erro na Validação/Emissão da NFC-e: ${err.message}`, 'error');
  } finally {
    // Unlock UI and restore button states
    isPdvSubmitting = false;
    if (btnNfce) {
      btnNfce.disabled = false;
      btnNfce.innerHTML = origNfceHtml || '<i class="fa-solid fa-check"></i> Emitir NFC-e (F9)';
    }
    if (btnSimple) {
      btnSimple.disabled = false;
      btnSimple.innerHTML = origSimpleHtml || '<i class="fa-solid fa-receipt"></i> Venda Simples';
    }
    if (btnCancel) btnCancel.disabled = false;
  }
}

function formatCpfCnpjReceipt(val) {
  if (!val) return '';
  const d = String(val).replace(/\D/g, '');
  if (d.length === 11) {
    return d.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, '$1.$2.$3-$4');
  } else if (d.length === 14) {
    return d.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5');
  }
  return val;
}

function renderThermalReceipt(saleData, mode = 'nfce', openModal = true) {
  if (!saleData) return;
  lastSaleData = saleData;

  const venda = saleData.venda || saleData.pedido || {};
  const itens = saleData.itens || [];
  const nfe = saleData.nfe || {};
  const empresa = saleData.empresa || {};

  // Logomarca da Empresa no Comprovante
  const logoWrap = document.getElementById('rec-company-logo-wrap');
  const logoImg = document.getElementById('rec-company-logo');
  const logoUrl = empresa.Logo || empresa.logo;
  if (logoWrap && logoImg) {
    if (logoUrl) {
      logoImg.src = logoUrl;
      logoWrap.style.display = 'block';
    } else {
      logoWrap.style.display = 'none';
      logoImg.src = '';
    }
  }

  document.getElementById('rec-company-name').innerText = empresa.NomeEmpresa || empresa.RazaoSocial || 'SIDCOMP VENDAS LTDA';
  document.getElementById('rec-company-sub').innerText = empresa.Cabecalho1 || empresa.Fantasia || 'FRENTE DE CAIXA';
  document.getElementById('rec-company-cnpj').innerText = `CNPJ: ${empresa.CNPJ || '11.054.174/0001-53'} | IE: ${empresa.InscEst || empresa.IE || '123456789110'}`;

  const clientName = (venda.NomeCliente || venda.nomecliente || venda.Nome || '').trim();
  const rawDoc = (venda.CPF || venda.CGC || '').trim();
  const street = (venda.Endereco || '').trim();
  const nro = (venda.Nro || '').trim();
  const bairro = (venda.Bairro || '').trim();
  const cidade = (venda.Cidade || '').trim();
  const uf = (venda.Uf || '').trim();
  const cep = (venda.Cep || '').trim();
  const fone = (venda.Fone || '').trim();

  const isGeneric = !clientName || 
    ['CONSUMIDOR', 'CONSUMIDOR FINAL', 'CLIENTE CONSUMIDOR', 'PADRAO', 'PADRÃO'].includes(clientName.toUpperCase());

  // Bloco de Dados do Cliente entre Cabeçalho e Produtos
  const clientBlock = document.getElementById('rec-client-block');
  if (clientBlock) {
    if (!isGeneric) {
      clientBlock.style.display = 'block';
      document.getElementById('rec-cli-nome').innerText = clientName;

      // CPF / CNPJ
      const docEl = document.getElementById('rec-cli-doc-row');
      if (rawDoc) {
        document.getElementById('rec-cli-doc').innerText = formatCpfCnpjReceipt(rawDoc);
        docEl.style.display = 'block';
      } else {
        docEl.style.display = 'none';
      }

      // Endereço (Logradouro, Nº - Bairro)
      let endStr = street;
      if (nro && nro !== '0' && nro.toUpperCase() !== 'SN') {
        endStr += endStr ? `, ${nro}` : `Nº ${nro}`;
      }
      if (bairro) {
        endStr += endStr ? ` - ${bairro}` : bairro;
      }

      const endEl = document.getElementById('rec-cli-end-row');
      if (endStr) {
        document.getElementById('rec-cli-end').innerText = endStr;
        endEl.style.display = 'block';
      } else {
        endEl.style.display = 'none';
      }

      // Cidade / UF - CEP
      let cidStr = cidade;
      if (uf) {
        cidStr += cidStr ? `/${uf}` : uf;
      }
      if (cep) {
        const cleanCep = cep.replace(/\D/g, '');
        const formattedCep = cleanCep.length === 8 ? cleanCep.replace(/^(\d{5})(\d{3})$/, '$1-$2') : cep;
        cidStr += cidStr ? ` - CEP: ${formattedCep}` : `CEP: ${formattedCep}`;
      }

      const cidEl = document.getElementById('rec-cli-cid-row');
      if (cidStr) {
        document.getElementById('rec-cli-cid').innerText = cidStr;
        cidEl.style.display = 'block';
      } else {
        cidEl.style.display = 'none';
      }

      // Telefone
      const foneEl = document.getElementById('rec-cli-fone-row');
      if (fone) {
        document.getElementById('rec-cli-fone').innerText = fone;
        foneEl.style.display = 'block';
      } else {
        foneEl.style.display = 'none';
      }
    } else {
      clientBlock.style.display = 'none';
    }
  }

  document.getElementById('rec-cliente-nome').innerText = isGeneric ? 'CONSUMIDOR FINAL' : clientName;
  const rawPedId = venda.CodPed || venda.Pedido || 1;
  const pedNumStr = String(rawPedId).padStart(6, '0');
  const fullDateTime = `${venda.DataEmiss || ''} ${venda.Hora || ''}`.trim() || (new Date()).toLocaleString('pt-BR');

  const recPedNumVal = document.getElementById('rec-ped-num-val');
  if (recPedNumVal) {
    recPedNumVal.innerText = `#${rawPedId}`;
  }
  const recPedDateVal = document.getElementById('rec-ped-date-val');
  if (recPedDateVal) {
    recPedDateVal.innerText = fullDateTime;
  }

  document.getElementById('rec-nfe-num').innerText = pedNumStr;
  document.getElementById('rec-nfe-data').innerText = fullDateTime;

  const vSub = floatOrZero(venda.SubTotal || venda.Total);
  const vDesc = floatOrZero(venda.Desconto);
  const vTot = floatOrZero(venda.Total || (vSub - vDesc));

  document.getElementById('rec-subtotal').innerText = vSub.toFixed(2);
  document.getElementById('rec-desconto').innerText = vDesc.toFixed(2);
  document.getElementById('rec-total-final').innerText = vTot.toFixed(2);
  document.getElementById('rec-forma-pgto').innerText = venda.CondPgto || 'DINHEIRO';

  const valRecEl = document.getElementById('pay-valor-recebido');
  const valRec = (valRecEl && parseFloat(valRecEl.value)) ? parseFloat(valRecEl.value) : vTot;
  const troco = Math.max(0.0, valRec - vTot);

  document.getElementById('rec-vlr-pago').innerText = valRec.toFixed(2);
  document.getElementById('rec-troco').innerText = troco.toFixed(2);

  // Items table (2 linhas por produto: Linha 1 = Código e Descrição, Linha 2 = Qtd x Valor Unitário e Subtotal)
  const tbody = document.getElementById('rec-items-tbody');
  tbody.innerHTML = itens.map((it, idx) => {
    const prdCode = it.Produto || it.CodPrd || '';
    const prdDesc = it.Descricao_Produto || it.Descricao || '';
    const qtd = floatOrZero(it.Qtd || 1);
    const embalagem = it.Embalagem || 'UN';
    const vUnit = it.ValorUnit !== undefined && it.ValorUnit !== null && floatOrZero(it.ValorUnit) > 0
      ? floatOrZero(it.ValorUnit)
      : (it.Valor && qtd ? floatOrZero(it.Valor) / qtd : floatOrZero(it.PrecoTabela || 0));
    const vTotItem = it.Valor !== undefined && it.Valor !== null ? floatOrZero(it.Valor) : (qtd * vUnit);

    return `
      <tr class="rec-item-row-title">
        <td colspan="2" style="font-weight: 700; padding-top: 4px; padding-bottom: 1px; word-break: break-word;">
          ${idx + 1}. ${prdCode ? '#' + prdCode + ' ' : ''}${escapeHtml(prdDesc)}
        </td>
      </tr>
      <tr class="rec-item-row-calc">
        <td style="padding-bottom: 4px; padding-left: 10px; font-size: 10.5px;">
          ${qtd} ${embalagem} &nbsp;x&nbsp; ${vUnit.toFixed(2)}
        </td>
        <td style="text-align: right; font-weight: 700; padding-bottom: 4px;">
          ${vTotItem.toFixed(2)}
        </td>
      </tr>
    `;
  }).join('');
  document.getElementById('rec-total-items').innerText = itens.length;

  const isNfceMode = mode === 'nfce' && (nfe.NroChave || venda.Sat);
  const homologBannerEl = document.getElementById('rec-homolog-banner');
  const watermarkEl = document.getElementById('rec-watermark');
  const watermarkTextEl = document.getElementById('rec-watermark-text');

  // Document Title & Blocks
  if (isNfceMode) {
    document.getElementById('rec-modal-header-title').innerHTML = `<i class="fa-solid fa-file-invoice"></i> Comprovante DANFE NFC-e`;
    document.getElementById('rec-doc-title').innerText = 'DANFE NFC-e';
    document.getElementById('rec-doc-subtitle').innerText = 'Nota Fiscal de Consumidor Eletrônica';
    document.getElementById('rec-doc-extra').innerText = 'Não Permite Crédito de ICMS';
    document.getElementById('rec-nfce-block').style.display = 'block';
    document.getElementById('rec-qrcode-block').style.display = 'flex';

    const chave = nfe.NroChave || venda.Sat || '35260811054174000153650010000126591599973082';
    const proto = (nfe.Protocolo || venda.Protocolo || '').trim();
    const statusNfe = String(nfe.Status || nfe.status || venda.StatusNFe || venda.Status || '').toUpperCase();
    const ambiente = String(nfe.Ambiente || venda.Ambiente || (window.currentPdvConfig && window.currentPdvConfig.Nfce && window.currentPdvConfig.Nfce.Ambiente) || '2');

    const isHomologacao = (ambiente !== '1' && ambiente !== 1);
    const isAuthorized = (proto.length > 0 || statusNfe.includes('AUTORIZ'));
    const semValidadeFiscal = isHomologacao || !isAuthorized;

    if (semValidadeFiscal) {
      if (homologBannerEl) {
        homologBannerEl.style.display = 'block';
        homologBannerEl.innerText = isHomologacao 
          ? 'EMITIDA EM AMBIENTE DE HOMOLOGAÇÃO - SEM VALOR FISCAL' 
          : 'DOCUMENTO NÃO VALIDADO NA SEFAZ - SEM VALOR FISCAL';
      }
      if (watermarkEl) watermarkEl.style.display = 'none';
      document.getElementById('rec-protocolo').innerText = isHomologacao
        ? `HOMOLOGAÇÃO - SEM VALOR FISCAL (${proto || 'SIMULAÇÃO'})`
        : `NÃO AUTORIZADA NA SEFAZ (${proto || 'PENDENTE'})`;
    } else {
      if (homologBannerEl) homologBannerEl.style.display = 'none';
      if (watermarkEl) watermarkEl.style.display = 'none';
      document.getElementById('rec-protocolo').innerText = proto || '1352608071205011705';
    }

    document.getElementById('rec-chave-nfe').innerText = chave.replace(/(.{4})/g, '$1 ').trim();

    // QR Code Rendering
    const qrContainer = document.getElementById('rec-qrcode-canvas');
    qrContainer.innerHTML = '';
    const qrUrl = `https://www.nfce.fazenda.sp.gov.br/qrcode?p=${chave}|2|2|000001`;

    if (window.QRCode) {
      try {
        new QRCode(qrContainer, {
          text: qrUrl,
          width: 140,
          height: 140
        });
      } catch (e) {
        renderFallbackQrCode(qrContainer);
      }
    } else {
      renderFallbackQrCode(qrContainer);
    }
  } else {
    document.getElementById('rec-modal-header-title').innerHTML = `<i class="fa-solid fa-receipt"></i> Recibo de Venda`;
    document.getElementById('rec-doc-title').innerText = 'CUPOM DE VENDA';
    document.getElementById('rec-doc-subtitle').innerText = 'Comprovante de Venda';
    document.getElementById('rec-doc-extra').innerText = 'Documento de Simples Conferência';
    document.getElementById('rec-nfce-block').style.display = 'none';
    document.getElementById('rec-qrcode-block').style.display = 'none';

    // Para comprovante de venda simples / não-fiscal
    if (homologBannerEl) homologBannerEl.style.display = 'none';
    if (watermarkEl) watermarkEl.style.display = 'none';
  }

  if (openModal) {
    document.getElementById('pdv-receipt-modal').classList.add('active');
  } else {
    document.getElementById('pdv-receipt-modal').classList.remove('active');
  }
}

function renderFallbackQrCode(container) {
  container.innerHTML = `
    <div style="border: 2px solid #000; padding: 10px; text-align: center; width: 130px; margin: 0 auto; background: #fff;">
      <i class="fa-solid fa-qrcode fa-4x" style="color: #000;"></i>
      <div style="font-size: 8px; font-weight: bold; margin-top: 4px;">SEFAZ QR CODE</div>
    </div>
  `;
}

function printSimpleSaleReceipt(openModal = false) {
  if (lastSaleData) {
    renderThermalReceipt(lastSaleData, 'venda', openModal);
  }
  document.body.classList.add('print-receipt-mode');
  setTimeout(() => {
    window.print();
    setTimeout(() => {
      document.body.classList.remove('print-receipt-mode');
      if (!openModal) {
        document.getElementById('pdv-receipt-modal').classList.remove('active');
      }
    }, 800);
  }, 150);
}

function printDanfeNfce(openModal = false) {
  if (lastSaleData) {
    renderThermalReceipt(lastSaleData, 'nfce', openModal);
  }
  document.body.classList.add('print-receipt-mode');
  setTimeout(() => {
    window.print();
    setTimeout(() => {
      document.body.classList.remove('print-receipt-mode');
      if (!openModal) {
        document.getElementById('pdv-receipt-modal').classList.remove('active');
      }
    }, 800);
  }, 150);
}

function closeReceiptModal() {
  document.getElementById('pdv-receipt-modal').classList.remove('active');
}

function printReceipt() {
  document.body.classList.add('print-receipt-mode');
  setTimeout(() => {
    window.print();
    setTimeout(() => document.body.classList.remove('print-receipt-mode'), 800);
  }, 150);
}

/* ==========================================================================
   PAINEL DE VENDA POR FOTOS (CATÁLOGO TOUCH & CLIQUE NO PDV)
   ========================================================================== */

function setPdvViewMode(mode) {
  pdvViewMode = mode;
  
  // Atualiza botões seletores de modo
  const btnSplit = document.getElementById('btn-pdv-mode-split');
  const btnCat = document.getElementById('btn-pdv-mode-catalog');
  const btnTab = document.getElementById('btn-pdv-mode-table');
  
  if (btnSplit) btnSplit.classList.toggle('active', mode === 'split');
  if (btnCat) btnCat.classList.toggle('active', mode === 'catalog');
  if (btnTab) btnTab.classList.toggle('active', mode === 'table');

  const photoPanel = document.getElementById('pdv-photo-panel');
  const barcodeCard = document.getElementById('pdv-barcode-entry-card');
  const cartCard = document.getElementById('pdv-cart-table-card');
  const photoGrid = document.getElementById('pdv-photo-grid');

  if (mode === 'split') {
    if (photoPanel) photoPanel.style.display = 'flex';
    if (barcodeCard) barcodeCard.style.display = 'block';
    if (cartCard) {
      cartCard.style.display = 'flex';
      cartCard.style.minHeight = '180px';
      cartCard.style.maxHeight = '280px';
    }
    if (photoGrid) photoGrid.style.maxHeight = '380px';
  } else if (mode === 'catalog') {
    if (photoPanel) photoPanel.style.display = 'flex';
    if (barcodeCard) barcodeCard.style.display = 'none';
    if (cartCard) {
      cartCard.style.display = 'flex';
      cartCard.style.minHeight = '140px';
      cartCard.style.maxHeight = '220px';
    }
    if (photoGrid) photoGrid.style.maxHeight = '560px';
  } else if (mode === 'table') {
    if (photoPanel) photoPanel.style.display = 'none';
    if (barcodeCard) barcodeCard.style.display = 'block';
    if (cartCard) {
      cartCard.style.display = 'flex';
      cartCard.style.minHeight = '420px';
      cartCard.style.maxHeight = '580px';
    }
  }
}

let pdvPhotoSearchDebounceTimer = null;
function onPdvPhotoSearchInput(val) {
  clearTimeout(pdvPhotoSearchDebounceTimer);
  pdvPhotoSearchDebounceTimer = setTimeout(() => {
    pdvCurrentPhotoSearch = (val || '').trim();
    renderPdvPhotoGrid();
  }, 180);
}

async function loadPdvPhotoCatalog() {
  try {
    // 1. Carregar grupos para a barra de categorias/chips
    const resGrupos = await fetch('/api/grupos');
    if (resGrupos.ok) {
      const grupos = await resGrupos.json();
      renderPdvCategoryChips(grupos);
    }

    // 2. Carregar catálogo inicial de produtos com prioridade de fotos
    const resPrd = await fetch('/api/pdv/produtos?limit=120');
    if (resPrd.ok) {
      pdvCatalogProducts = await resPrd.json();
      renderPdvPhotoGrid();
    }
  } catch (err) {
    console.error('Erro ao carregar catálogo de fotos no PDV:', err);
    const grid = document.getElementById('pdv-photo-grid');
    if (grid) {
      grid.innerHTML = `
        <div style="grid-column: 1 / -1; text-align: center; padding: 2rem; color: var(--accent-rose);">
          <i class="fa-solid fa-triangle-exclamation fa-2x"></i>
          <div style="margin-top: 0.5rem;">Falha ao carregar produtos no catálogo.</div>
        </div>
      `;
    }
  }
}

function renderPdvCategoryChips(grupos = []) {
  const container = document.getElementById('pdv-category-chips-container');
  if (!container) return;

  let html = `
    <button type="button" class="pdv-category-chip ${pdvCurrentGroupFilter === 'all' ? 'active' : ''}" onclick="filterPdvCatalogByGroup('all', this)">
      <i class="fa-solid fa-boxes-stacked"></i> Todos
    </button>
  `;

  if (Array.isArray(grupos)) {
    grupos.forEach(g => {
      const nome = g.Descricao_Grupo || g.Nome_Grupo || g.Descricao || `Grupo #${g.CodGru}`;
      const isActive = String(pdvCurrentGroupFilter) === String(g.CodGru);
      html += `
        <button type="button" class="pdv-category-chip ${isActive ? 'active' : ''}" onclick="filterPdvCatalogByGroup(${g.CodGru}, this)">
          <i class="fa-solid fa-tag"></i> ${escapeHtml(nome)}
        </button>
      `;
    });
  }

  container.innerHTML = html;
}

async function filterPdvCatalogByGroup(groupId, element) {
  pdvCurrentGroupFilter = groupId;
  
  // Atualizar estilo ativo nos chips
  const container = document.getElementById('pdv-category-chips-container');
  if (container) {
    container.querySelectorAll('.pdv-category-chip').forEach(chip => chip.classList.remove('active'));
    if (element) {
      element.classList.add('active');
    }
  }

  try {
    const url = groupId === 'all'
      ? `/api/pdv/produtos?limit=120`
      : `/api/pdv/produtos?grupo=${encodeURIComponent(groupId)}&limit=120`;
    const res = await fetch(url);
    if (res.ok) {
      pdvCatalogProducts = await res.json();
    }
  } catch (e) {
    console.warn('Erro ao filtrar catálogo por grupo:', e);
  }

  renderPdvPhotoGrid();
}

function isPdvValidPhoto(foto) {
  if (!foto) return false;
  const str = String(foto).trim();
  if (!str || str === '0' || str === '-1' || str === 'null' || str === 'undefined') return false;
  return true;
}

function renderPdvPhotoGrid() {
  const grid = document.getElementById('pdv-photo-grid');
  const countBadge = document.getElementById('pdv-catalog-count');
  if (!grid) return;

  let filtered = pdvCatalogProducts || [];

  // Filtragem por grupo local se aplicável
  if (pdvCurrentGroupFilter && pdvCurrentGroupFilter !== 'all') {
    filtered = filtered.filter(p => String(p.Grupo) === String(pdvCurrentGroupFilter));
  }

  // Filtragem instantânea por texto (nome, código, ean, marca)
  if (pdvCurrentPhotoSearch) {
    const term = pdvCurrentPhotoSearch.toLowerCase();
    filtered = filtered.filter(p => {
      const nome = (p.Descricao_Produto || '').toLowerCase();
      const cod = String(p.CodPrd || '');
      const ean = (p.CodBar || '').toLowerCase();
      const marca = (p.Marca || '').toLowerCase();
      return nome.includes(term) || cod.includes(term) || ean.includes(term) || marca.includes(term);
    });
  }

  if (countBadge) {
    countBadge.innerText = `${filtered.length} produto(s)`;
  }

  if (!filtered.length) {
    grid.innerHTML = `
      <div style="grid-column: 1 / -1; text-align: center; padding: 2.5rem; color: var(--text-muted);">
        <i class="fa-solid fa-box-open fa-3x" style="opacity: 0.35; margin-bottom: 0.5rem;"></i>
        <div style="font-weight: 500; font-size: 0.95rem;">Nenhum produto encontrado neste filtro.</div>
        <div style="font-size: 0.8rem; margin-top: 0.25rem;">Tente outro termo de busca ou selecione a categoria "Todos".</div>
      </div>
    `;
    return;
  }

  grid.innerHTML = filtered.map(p => {
    const temFoto = isPdvValidPhoto(p.Foto);
    const preco = floatOrZero(p.Venda);
    const descGrupo = floatOrZero(p.DescontoGrupo || p.descontogrupo || p.Desconto_Grupo);
    const precoFinal = descGrupo > 0 ? (preco * (1 - descGrupo / 100)) : preco;
    const estoque = floatOrZero(p.Estoque);
    const stockClass = estoque <= 0 ? 'badge-stock-out' : (estoque <= 5 ? 'badge-stock-low' : 'badge-stock-ok');

    const imgHtml = temFoto
      ? `<img src="${escapeHtml(p.Foto)}" alt="${escapeHtml(p.Descricao_Produto)}" loading="lazy" onerror="this.onerror=null; this.parentElement.innerHTML='<div class=\\'pdv-product-placeholder\\'><i class=\\'fa-solid fa-box-open fa-2x\\'></i><span style=\\'font-size:0.7rem;\\'>Sem foto</span></div>';">`
      : `<div class="pdv-product-placeholder">
           <i class="fa-solid fa-box fa-2x"></i>
           <span style="font-size: 0.7rem; font-weight: 500;">Sem foto</span>
         </div>`;

    const priceBadge = descGrupo > 0
      ? `<div class="pdv-product-price-badge" title="Desconto de Grupo ${descGrupo}%">R$ ${precoFinal.toFixed(2)}</div>`
      : `<div class="pdv-product-price-badge">R$ ${preco.toFixed(2)}</div>`;

    return `
      <div class="pdv-product-card" id="pdv-card-prd-${p.CodPrd}" onclick="onPdvPhotoCardClick(${p.CodPrd}, this)" title="Clique para adicionar ao carrinho">
        ${priceBadge}
        <div class="pdv-product-img-wrapper">
          ${imgHtml}
        </div>
        <div class="pdv-product-info">
          <div class="pdv-product-name" title="${escapeHtml(p.Descricao_Produto)}">${escapeHtml(p.Descricao_Produto)}</div>
          <div class="pdv-product-meta-row">
            <span style="font-family: monospace; font-weight: 600; color: var(--text-secondary);">#${p.CodPrd}</span>
            <span class="badge ${stockClass} pdv-product-stock-tag" title="Estoque disponível">${estoque} ${escapeHtml(p.Embalagem || 'UN')}</span>
          </div>
        </div>
      </div>
    `;
  }).join('');
}

function onPdvPhotoCardClick(codPrd, cardElement) {
  const prod = (pdvCatalogProducts || []).find(p => p.CodPrd === codPrd);
  if (!prod) {
    showToast('Produto não localizado no catálogo', 'error');
    return;
  }

  // Animação de feedback táctil e visual
  if (cardElement) {
    cardElement.style.transform = 'scale(0.92)';
    cardElement.style.borderColor = 'var(--accent-emerald)';
    setTimeout(() => {
      cardElement.style.transform = '';
      cardElement.style.borderColor = '';
    }, 180);
  }

  // Se o operador informou quantidade ou preço personalizado nos campos do topo, utiliza-os
  const qtyInp = document.getElementById('pdv-qty-input');
  const priceInp = document.getElementById('pdv-price-input');
  const qty = parseFloat(qtyInp ? qtyInp.value : 1.0) || 1.0;
  const customPrice = (priceInp && priceInp.value.trim() !== '') ? parseFloat(priceInp.value) : null;

  // Adiciona item ao carrinho do PDV
  addPdvItemToCart(prod, qty, customPrice);

  // Reseta inputs superiores
  if (qtyInp) qtyInp.value = '1';
  if (priceInp) priceInp.value = '';
}

function floatOrZero(v) {
  const f = parseFloat(v);
  return isNaN(f) ? 0.0 : f;
}

/* ==================== PDV CONFIG (config.ini) FUNCTIONS ==================== */

async function openPdvConfigModal() {
  document.getElementById('pdv-config-modal').classList.add('active');
  await loadPdvConfig();
}

function closePdvConfigModal() {
  document.getElementById('pdv-config-modal').classList.remove('active');
}

async function loadPdvConfig() {
  try {
    const res = await fetch('/api/pdv/config');
    const cfg = await res.json();
    if (!res.ok) throw new Error(cfg.error || 'Erro ao carregar configurações');

    // [Config]
    const c = cfg.Config || {};
    setElVal('cfg_Config_Porta', c.Porta || 'TELA');
    setElVal('cfg_Config_PortaNf', c.PortaNf || '');
    setElVal('cfg_Config_Leitor', c.Leitor || 'TECLADO');
    setElVal('cfg_Config_LocalEst', c.LocalEst || 'LOJA');
    setElVal('cfg_Config_Copias', c.Copias || '2');
    setElVal('cfg_Config_QtdLinhas', c.QtdLinhas || '7');

    setElCheck('cfg_Config_BaixaAuto', c.BaixaAuto === '1');
    setElCheck('cfg_Config_MostraEstoque', c.MostraEstoque === '1');
    setElCheck('cfg_Config_MostraCusto', c.MostraCusto === '1');
    setElCheck('cfg_Config_Desconto', c.Desconto === '1');

    // [Sat]
    const s = cfg.Sat || {};
    setElVal('cfg_Sat_NroCaixa', s.NroCaixa || '1');
    setElVal('cfg_Sat_CnpjDesenvolvedor', s.CnpjDesenvolvedor || '');
    setElVal('cfg_Sat_LocalAcbr', s.LocalAcbr || '');

    // [Nfce]
    const n = cfg.Nfce || {};
    setElVal('cfg_Nfce_SerieNfce', n.SerieNfce || '1');
    setElVal('cfg_Nfce_NroNfce', n.NroNfce || '1');
    setElVal('cfg_Nfce_CodigoID', n.CodigoID || '');
    setElVal('cfg_Nfce_CodigoSeg', n.CodigoSeg || '');

    // [boleto]
    const b = cfg.boleto || {};
    setElVal('cfg_boleto_Banco', b.Banco || '');
    setElVal('cfg_boleto_AGENCIA', b.AGENCIA || '');
    setElVal('cfg_boleto_CONTA', b.CONTA || '');
    setElVal('cfg_boleto_Convenio', b.Convenio || '');
  } catch (err) {
    showToast(`Erro ao carregar configurações: ${err.message}`, 'error');
  }
}

async function savePdvConfig(e) {
  e.preventDefault();
  const payload = {
    Config: {
      Porta: getElVal('cfg_Config_Porta'),
      PortaNf: getElVal('cfg_Config_PortaNf'),
      Leitor: getElVal('cfg_Config_Leitor'),
      LocalEst: getElVal('cfg_Config_LocalEst'),
      Copias: getElVal('cfg_Config_Copias'),
      QtdLinhas: getElVal('cfg_Config_QtdLinhas'),
      BaixaAuto: getElCheck('cfg_Config_BaixaAuto') ? '1' : '0',
      MostraEstoque: getElCheck('cfg_Config_MostraEstoque') ? '1' : '0',
      MostraCusto: getElCheck('cfg_Config_MostraCusto') ? '1' : '0',
      Desconto: getElCheck('cfg_Config_Desconto') ? '1' : '0'
    },
    Sat: {
      NroCaixa: getElVal('cfg_Sat_NroCaixa'),
      CnpjDesenvolvedor: getElVal('cfg_Sat_CnpjDesenvolvedor'),
      LocalAcbr: getElVal('cfg_Sat_LocalAcbr')
    },
    Nfce: {
      SerieNfce: getElVal('cfg_Nfce_SerieNfce'),
      NroNfce: getElVal('cfg_Nfce_NroNfce'),
      CodigoID: getElVal('cfg_Nfce_CodigoID'),
      CodigoSeg: getElVal('cfg_Nfce_CodigoSeg')
    },
    boleto: {
      Banco: getElVal('cfg_boleto_Banco'),
      AGENCIA: getElVal('cfg_boleto_AGENCIA'),
      CONTA: getElVal('cfg_boleto_CONTA'),
      Convenio: getElVal('cfg_boleto_Convenio')
    }
  };

  try {
    const res = await fetch('/api/pdv/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao salvar');

    showToast(result.message || 'Configurações salvas com sucesso no banco de dados (tabela CfgPdv)!', 'success');
    closePdvConfigModal();
  } catch (err) {
    showToast(`Erro ao salvar configurações: ${err.message}`, 'error');
  }
}

function setElVal(id, val) {
  const el = document.getElementById(id);
  if (el) el.value = val;
}

function setElCheck(id, checked) {
  const el = document.getElementById(id);
  if (el) el.checked = Boolean(checked);
}

function getElVal(id) {
  const el = document.getElementById(id);
  return el ? el.value : '';
}

function getElCheck(id) {
  const el = document.getElementById(id);
  return el ? el.checked : false;
}

/* ==================== MULTIEMPRESA & EMPRESAS FUNCTIONS ==================== */

let currentEmpresaId = 1;
let empresasList = [];

onReady(() => {
  loadActiveEmpresasSelector();
});

async function loadActiveEmpresasSelector() {
  try {
    const res = await fetch('/api/empresas?all=1');
    const data = await res.json();
    if (res.ok && Array.isArray(data)) {
      empresasList = data;
      const select = document.getElementById('select-active-empresa');
      if (select && data.length) {
        select.innerHTML = data.map(e => `
          <option value="${e.id_empresa}" ${e.id_empresa === currentEmpresaId ? 'selected' : ''}>
            ${e.id_empresa} - ${escapeHtml(e.Fantasia || e.RazaoSocial)}
          </option>
        `).join('');
      }
    }
  } catch (err) {
    console.error('Erro ao carregar lista de empresas:', err);
  }
}

function onActiveEmpresaChange(id) {
  currentEmpresaId = parseInt(id, 10) || 1;
  const emp = empresasList.find(e => e.id_empresa === currentEmpresaId);
  showToast(`Empresa ativa alterada: ${emp ? emp.RazaoSocial : id}`, 'info');
  refreshCurrentModule();
}

async function fetchEmpresas() {
  try {
    const res = await fetch('/api/empresas?all=1');
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Erro ao carregar empresas');
    empresasList = data;
    renderEmpresasTable(data);
  } catch (err) {
    showToast(`Erro ao carregar empresas: ${err.message}`, 'error');
  }
}

function renderEmpresasTable(empresas) {
  const tbody = document.getElementById('table-body-empresas');
  if (!tbody) return;

  if (!empresas || !empresas.length) {
    tbody.innerHTML = `<tr><td colspan="9" class="empty-td">Nenhuma empresa cadastrada</td></tr>`;
    return;
  }

  tbody.innerHTML = empresas.map(e => `
    <tr>
      <td><strong>#${e.id_empresa}</strong></td>
      <td><strong>${escapeHtml(e.RazaoSocial)}</strong></td>
      <td>${escapeHtml(e.Fantasia || '-')}</td>
      <td>${escapeHtml(e.CNPJ || '-')}</td>
      <td>${escapeHtml(e.InscEst || '-')}</td>
      <td>${escapeHtml(e.Cidade || '-')}/${escapeHtml(e.UF || 'SP')}</td>
      <td>${escapeHtml(e.Fone || '-')}</td>
      <td>
        <span class="badge ${e.Ativo === 1 ? 'badge-active' : 'badge-inactive'}">
          ${e.Ativo === 1 ? 'Ativa' : 'Inativa'}
        </span>
      </td>
      <td style="text-align: center;">
        <button class="btn btn-secondary btn-sm" onclick="openEmpresaModal(${e.id_empresa})" title="Editar / Alterar Dados da Empresa">
          <i class="fa-solid fa-pen-to-square"></i> Alterar
        </button>
      </td>
    </tr>
  `).join('');
}

async function openEmpresaModal(id_empresa = null) {
  const form = document.getElementById('empresa-form');
  if (form) form.reset();
  document.getElementById('Emp_id_empresa').value = '';
  setEmpresaLogoPreview('');

  const titleEl = document.getElementById('modal-title-emp');

  if (id_empresa) {
    titleEl.innerHTML = `<i class="fa-solid fa-building-pen"></i> Alterar Dados da Empresa #${id_empresa}`;
    try {
      const res = await fetch(`/api/empresas/${id_empresa}`);
      const emp = await res.json();
      if (!res.ok) throw new Error(emp.error || 'Empresa não encontrada');

      document.getElementById('Emp_id_empresa').value = emp.id_empresa;
      document.getElementById('Emp_RazaoSocial').value = emp.RazaoSocial || '';
      document.getElementById('Emp_Fantasia').value = emp.Fantasia || '';
      document.getElementById('Emp_CNPJ').value = emp.CNPJ || '';
      document.getElementById('Emp_InscEst').value = emp.InscEst || '';
      document.getElementById('Emp_Logradouro').value = emp.Logradouro || '';
      document.getElementById('Emp_Nro').value = emp.Nro || '';
      document.getElementById('Emp_Bairro').value = emp.Bairro || '';
      document.getElementById('Emp_Cidade').value = emp.Cidade || '';
      document.getElementById('Emp_UF').value = emp.UF || 'SP';
      document.getElementById('Emp_CEP').value = emp.CEP || '';
      document.getElementById('Emp_Fone').value = emp.Fone || '';
      document.getElementById('Emp_CodigoIBGE').value = emp.CodigoIBGE || '';
      document.getElementById('Emp_RegimeTrib').value = emp.RegimeTrib || '3';
      document.getElementById('Emp_CFOP').value = emp.CFOP || '';
      document.getElementById('Emp_PIS').value = emp.PIS || '01';
      document.getElementById('Emp_AliqPIS').value = emp.AliqPIS || '0,65';
      document.getElementById('Emp_COFINS').value = emp.COFINS || '01';
      document.getElementById('Emp_AliqCOFINS').value = emp.AliqCOFINS || '3';
      document.getElementById('Emp_SitTrib').value = emp.SitTrib || '';
      document.getElementById('Emp_SitTribFixo').value = emp.SitTribFixo || '';
      document.getElementById('Emp_Deducao').value = emp.Deducao || '33,33';
      document.getElementById('Emp_Ativo').checked = emp.Ativo === 1;

      // Logomarca Preview
      setEmpresaLogoPreview(emp.Logo || emp.logo || '');
    } catch (err) {
      showToast(`Erro ao carregar dados da empresa: ${err.message}`, 'error');
      return;
    }
  } else {
    titleEl.innerHTML = `<i class="fa-solid fa-building-circle-check"></i> Cadastrar Nova Empresa`;
    setEmpresaLogoPreview('');
  }

  document.getElementById('empresa-modal').classList.add('active');
}

function closeEmpresaModal() {
  document.getElementById('empresa-modal').classList.remove('active');
}

function setEmpresaLogoPreview(url) {
  const previewImg = document.getElementById('emp-logo-preview-img');
  const placeholder = document.getElementById('emp-logo-placeholder');
  const btnRemove = document.getElementById('btn-remove-emp-logo');
  const inputLogo = document.getElementById('Emp_Logo');

  if (inputLogo) inputLogo.value = url || '';

  if (url) {
    if (previewImg) {
      previewImg.src = url;
      previewImg.style.display = 'block';
    }
    if (placeholder) placeholder.style.display = 'none';
    if (btnRemove) btnRemove.style.display = 'inline-flex';
  } else {
    if (previewImg) {
      previewImg.src = '';
      previewImg.style.display = 'none';
    }
    if (placeholder) placeholder.style.display = 'block';
    if (btnRemove) btnRemove.style.display = 'none';
  }
}

function triggerEmpresaLogoUpload() {
  const fileInput = document.getElementById('emp-logo-file-input');
  if (fileInput) fileInput.click();
}

async function handleEmpresaLogoFileSelect(e) {
  const file = e.target.files[0];
  if (!file) return;
  await uploadEmpresaLogoFile(file);
  e.target.value = '';
}

async function uploadEmpresaLogoFile(file) {
  try {
    const formData = new FormData();
    formData.append('file', file);
    showToast('Enviando logomarca da empresa...', 'info');

    const res = await fetch('/api/empresas/upload-logo', {
      method: 'POST',
      body: formData
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Falha no upload do logo');

    setEmpresaLogoPreview(data.url);
    showToast('Logomarca enviada com sucesso!', 'success');
  } catch (err) {
    showToast(`Erro no upload da logomarca: ${err.message}`, 'error');
  }
}

function removeEmpresaLogo() {
  setEmpresaLogoPreview('');
  showToast('Logomarca removida. Clique em Salvar para confirmar.', 'info');
}

async function saveEmpresa(e) {
  e.preventDefault();
  const form = e.target;
  const formData = new FormData(form);
  const data = Object.fromEntries(formData.entries());

  try {
    const res = await fetch('/api/empresas', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao salvar empresa');

    showToast(result.message || 'Dados da empresa salvos com sucesso!', 'success');
    closeEmpresaModal();
    fetchEmpresas();
    loadActiveEmpresasSelector();
  } catch (err) {
    showToast(`Erro ao salvar empresa: ${err.message}`, 'error');
  }
}


// ==========================================
// 5. PEDIDOS DE VENDAS MODULE LOGIC
// ==========================================

let ordersCurrentPage = 1;
let ordersTotalPages = 1;
let ordersFilters = { status: 'all', q: '' };
let currentOrderItems = [];
let ordSelectedClient = null;
let ordSelectedProduct = null;
let ordClientSearchTimeout = null;
let ordProductSearchTimeout = null;

function debounceOrdersSearch(val) {
  ordersFilters.q = val;
  ordersCurrentPage = 1;
  fetchOrders();
}

function applyOrdersFilters() {
  const statusEl = document.getElementById('select-status-pedidos');
  if (statusEl) ordersFilters.status = statusEl.value;
  ordersCurrentPage = 1;
  fetchOrders();
}

async function fetchOrders() {
  const tbody = document.getElementById('table-body-pedidos');
  if (!tbody) return;

  tbody.innerHTML = `
    <tr>
      <td colspan="8" class="loading-td">
        <i class="fa-solid fa-spinner fa-spin fa-2x"></i>
        <div style="margin-top: 0.5rem;">Carregando pedidos de vendas...</div>
      </td>
    </tr>
  `;

  try {
    const queryParams = new URLSearchParams({
      page: ordersCurrentPage,
      limit: 15,
      q: ordersFilters.q,
      status: ordersFilters.status
    });

    const res = await fetch(`/api/pedidos?${queryParams.toString()}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Falha ao carregar pedidos');

    ordersTotalPages = data.pages;
    renderOrdersTable(data.items);
    renderOrdersPagination(data.total, data.page, data.limit);
  } catch (err) {
    tbody.innerHTML = `
      <tr>
        <td colspan="8" style="text-align: center; color: var(--accent-rose); padding: 2rem;">
          <i class="fa-solid fa-circle-exclamation fa-2x"></i>
          <div style="margin-top: 0.5rem;">Erro ao carregar pedidos: ${escapeHtml(err.message)}</div>
        </td>
      </tr>
    `;
  }
}

function renderOrdersTable(orders) {
  const tbody = document.getElementById('table-body-pedidos');
  if (!tbody) return;

  if (!orders || orders.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="8" class="loading-td">Nenhum pedido de vendas encontrado.</td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = orders.map(o => {
    const totalStr = `R$ ${floatOrZero(o.Total).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    const nfeBadge = o.NroChave
      ? `<span class="badge badge-active" title="Chave: ${o.NroChave}"><i class="fa-solid fa-check-double"></i> NF-e Emitida</span>`
      : `<span class="badge badge-warning"><i class="fa-solid fa-clock"></i> Sem NF-e</span>`;

    const clientName = o.NomeCliente || o.nomecliente || 'CONSUMIDOR FINAL';
    return `
      <tr>
        <td><strong>#${o.CodPed}</strong></td>
        <td>${escapeHtml(o.DataEmiss || '')}</td>
        <td>
          <div style="font-weight: 600; color: var(--text-primary);">${escapeHtml(clientName)}</div>
          <div style="font-size: 0.75rem; color: var(--text-muted);">${o.CPF || o.CGC || ''} ${o.Cidade ? ' - ' + o.Cidade : ''}</div>
        </td>
        <td>Vendedor #${o.Vendedor || 1}</td>
        <td><span class="badge badge-group">${escapeHtml(o.CondPgto || 'A VISTA')}</span></td>
        <td style="text-align: right; font-weight: 700; color: var(--accent-emerald);">${totalStr}</td>
        <td style="text-align: center;">${nfeBadge}</td>
        <td style="text-align: center;">
          <div style="display: flex; justify-content: center; gap: 0.35rem; flex-wrap: wrap;">
            <button class="btn btn-secondary btn-sm" onclick="editOrder(${o.CodPed})" title="Editar Pedido">
              <i class="fa-solid fa-pen"></i>
            </button>
            <button class="btn btn-secondary btn-sm" onclick="imprimirCupomOrder(${o.CodPed})" title="Imprimir Comprovante / Cupom do Pedido">
              <i class="fa-solid fa-receipt"></i>
            </button>
            <button class="btn btn-emerald btn-sm" onclick="openNfeEmissaoModal(${o.CodPed})" title="Abrir Tela de Emissão de NF-e 55 (FrmNota)">
              <i class="fa-solid fa-file-invoice"></i> NF-e
            </button>
            ${o.NroChave ? `<button class="btn btn-secondary btn-sm" onclick="imprimirDanfeOrder(${o.CodPed})" title="Imprimir DANFE NF-e (PDF)"><i class="fa-solid fa-file-pdf" style="color: var(--accent-rose);"></i></button>` : ''}
            <button class="btn btn-danger btn-sm" onclick="deleteOrder(${o.CodPed})" title="Excluir Pedido">
              <i class="fa-solid fa-trash"></i>
            </button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

function renderOrdersPagination(total, page, limit) {
  const container = document.getElementById('pedidos-pagination');
  if (!container) return;

  if (total === 0) {
    container.innerHTML = '';
    return;
  }

  const start = (page - 1) * limit + 1;
  const end = Math.min(page * limit, total);

  container.innerHTML = `
    <div class="pagination-info">
      Exibindo ${start}-${end} de ${total} pedidos
    </div>
    <div class="pagination-buttons">
      <button class="btn btn-secondary btn-sm" ${page <= 1 ? 'disabled' : ''} onclick="changeOrdersPage(${page - 1})">
        <i class="fa-solid fa-chevron-left"></i> Anterior
      </button>
      <span style="font-size: 0.875rem; font-weight: 600; padding: 0 0.5rem;">Página ${page} de ${ordersTotalPages}</span>
      <button class="btn btn-secondary btn-sm" ${page >= ordersTotalPages ? 'disabled' : ''} onclick="changeOrdersPage(${page + 1})">
        Próxima <i class="fa-solid fa-chevron-right"></i>
      </button>
    </div>
  `;
}

function changeOrdersPage(newPage) {
  if (newPage >= 1 && newPage <= ordersTotalPages) {
    ordersCurrentPage = newPage;
    fetchOrders();
  }
}

// Order Creation & Editing Modal Logic
async function openOrderModal(data = null) {
  const form = document.getElementById('order-form');
  if (form) form.reset();

  currentOrderItems = [];
  ordSelectedClient = null;
  ordSelectedProduct = null;

  const now = new Date();
  const dateStr = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`;

  // Load dynamic combo options first
  await Promise.all([
    loadFormasPgtoOptions(),
    loadOrdSellersOptions(),
    loadOrdCarriersOptions(),
    loadOrdNaturezasOptions()
  ]);

  const printBtn = document.getElementById('btn-print-order-modal');

  if (data && data.pedido) {
    const p = data.pedido;
    document.getElementById('modal-title-ped').innerHTML = `<i class="fa-solid fa-file-pen"></i> Editar Pedido de Vendas #${p.CodPed}`;
    document.getElementById('Ord_CodPed').value = p.CodPed;
    const clientName = p.NomeCliente || p.nomecliente;
    document.getElementById('ord-client-search-input').value = clientName ? `#${p.Entidade} - ${clientName}` : (p.Entidade && p.Entidade !== 1 ? `#${p.Entidade}` : '1 - CONSUMIDOR FINAL');
    document.getElementById('Ord_CodEntidade').value = p.Entidade || 1;
    document.getElementById('Ord_Vendedor').value = p.Vendedor || 1;
    document.getElementById('Ord_Transportadora').value = p.Transportadora || 1;
    document.getElementById('Ord_Operacao').value = p.Operacao || 1;
    document.getElementById('Ord_DataEmiss').value = p.DataEmiss || dateStr;
    document.getElementById('Ord_CondPgto').value = p.CondPgto || 'A VISTA';
    document.getElementById('Ord_Cfo').value = p.Cfo || '5102 - VENDA DE MERCADORIAS';
    document.getElementById('Ord_Obs').value = p.Obs || '';
    document.getElementById('Ord_Desconto').value = floatOrZero(p.Desconto).toFixed(2);
    document.getElementById('Ord_ValorFrete').value = floatOrZero(p.ValorFrete).toFixed(2);

    if (printBtn) {
      printBtn.style.display = 'inline-flex';
    }

    if (data.itens && data.itens.length) {
      currentOrderItems = data.itens.map(it => ({
        CodPrd: it.Produto || it.CodPrd,
        Descricao_Produto: it.Descricao_Produto || `PRODUTO #${it.Produto}`,
        Qtd: floatOrZero(it.Qtd),
        ValorUnit: floatOrZero(it.ValorUnit),
        Desconto: floatOrZero(it.Desconto),
        Valor: floatOrZero(it.Valor)
      }));
    }
  } else {
    document.getElementById('modal-title-ped').innerHTML = `<i class="fa-solid fa-file-circle-plus"></i> Novo Pedido de Vendas`;
    document.getElementById('Ord_CodPed').value = '';
    document.getElementById('ord-client-search-input').value = '1 - CONSUMIDOR FINAL';
    document.getElementById('Ord_CodEntidade').value = '1';
    document.getElementById('Ord_DataEmiss').value = dateStr;
    document.getElementById('Ord_Cfo').value = '5102 - VENDA DE MERCADORIAS';
    document.getElementById('Ord_Obs').value = '';
    document.getElementById('Ord_Desconto').value = '0.00';
    document.getElementById('Ord_ValorFrete').value = '0.00';

    if (printBtn) {
      printBtn.style.display = 'none';
    }
  }

  setupOrderAutocomplete();
  renderOrderItemsTable();
  document.getElementById('order-modal').classList.add('active');
}

async function loadOrdSellersOptions() {
  const sel = document.getElementById('Ord_Vendedor');
  if (!sel) return;
  try {
    const res = await fetch('/api/pdv/vendedores');
    const sellers = await res.json();
    if (!res.ok || !sellers.length) return;
    sel.innerHTML = sellers.map(s => `<option value="${s.CodEntidade}">#${s.CodEntidade} - ${escapeHtml(s.Nome)}</option>`).join('');
  } catch (err) {
    console.error('Erro ao carregar vendedores:', err);
  }
}

async function loadOrdCarriersOptions() {
  const sel = document.getElementById('Ord_Transportadora');
  if (!sel) return;
  try {
    const res = await fetch('/api/pdv/transportadoras');
    const carriers = await res.json();
    if (!res.ok || !carriers.length) return;
    sel.innerHTML = carriers.map(c => `<option value="${c.CodEntidade}">#${c.CodEntidade} - ${escapeHtml(c.Nome)}</option>`).join('');
  } catch (err) {
    console.error('Erro ao carregar transportadoras:', err);
  }
}

async function loadOrdNaturezasOptions() {
  const sel = document.getElementById('Ord_Operacao');
  if (!sel) return;
  try {
    const res = await fetch('/api/natureza-operacao');
    const naturezas = await res.json();
    if (!res.ok || !naturezas.length) return;
    sel.innerHTML = naturezas.map(n => `<option value="${n.id_natureza}" data-cfop="${n.CFOP || '5102'}">${escapeHtml(n.Descricao)} (CFOP ${n.CFOP || '5102'})</option>`).join('');
  } catch (err) {
    console.error('Erro ao carregar naturezas de operação:', err);
  }
}

function onNaturezaOperacaoChange(val) {
  const sel = document.getElementById('Ord_Operacao');
  if (!sel) return;
  const opt = sel.options[sel.selectedIndex];
  if (opt && opt.dataset && opt.dataset.cfop) {
    const cfoInp = document.getElementById('Ord_Cfo');
    if (cfoInp) cfoInp.value = `${opt.dataset.cfop} - ${opt.text.split('(')[0].trim()}`;
  }
}

function closeOrderModal() {
  document.getElementById('order-modal').classList.remove('active');
}

async function editOrder(id) {
  try {
    const res = await fetch(`/api/pedidos/${id}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Erro ao carregar pedido');
    openOrderModal(data);
  } catch (err) {
    showToast(`Erro ao abrir pedido: ${err.message}`, 'error');
  }
}

async function deleteOrder(id) {
  if (!confirm(`Tem certeza que deseja excluir o Pedido de Vendas #${id}?`)) return;

  try {
    const res = await fetch(`/api/pedidos/${id}`, { method: 'DELETE' });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao excluir pedido');

    showToast(result.message, 'success');
    fetchOrders();
  } catch (err) {
    showToast(`Erro ao excluir: ${err.message}`, 'error');
  }
}

let ordSellerSearchTimeout = null;
let ordCfoSearchTimeout = null;

function setupOrderAutocomplete() {
  // Cliente Search (Tipo 1)
  const clientInput = document.getElementById('ord-client-search-input');
  if (clientInput) {
    clientInput.oninput = (e) => {
      clearTimeout(ordClientSearchTimeout);
      const val = e.target.value.trim();
      if (val.length < 2) {
        document.getElementById('ord-client-results-overlay').style.display = 'none';
        return;
      }
      ordClientSearchTimeout = setTimeout(() => searchOrdClients(val), 250);
    };
  }

  // Produto Search (suporte a codigo de barras / EAN e tecla Enter)
  const productInput = document.getElementById('ord-product-search-input');
  if (productInput) {
    productInput.oninput = (e) => {
      clearTimeout(ordProductSearchTimeout);
      const val = e.target.value.trim();
      if (val.length < 2) {
        document.getElementById('ord-product-results-overlay').style.display = 'none';
        return;
      }
      ordProductSearchTimeout = setTimeout(() => searchOrdProducts(val), 250);
    };

    productInput.onkeydown = async (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const val = productInput.value.trim();
        if (!val) return;
        try {
          const res = await fetch(`/api/pdv/produtos?q=${encodeURIComponent(val)}`);
          const results = await res.json();
          if (res.ok && results.length) {
            const p = results[0];
            selectOrdProduct(p.CodPrd, p.Descricao_Produto, p.Venda);
          }
        } catch (err) {
          console.error('Erro na leitura de código de barras:', err);
        }
      }
    };
  }
}

async function searchOrdClients(q) {
  try {
    const res = await fetch(`/api/pdv/clientes?q=${encodeURIComponent(q)}`);
    const clients = await res.json();
    const overlay = document.getElementById('ord-client-results-overlay');
    if (!res.ok || !clients.length) {
      overlay.style.display = 'none';
      return;
    }
    overlay.innerHTML = clients.map(c => {
      const displayName = c.Fantasia && c.Fantasia.toUpperCase() !== c.Nome.toUpperCase() ? `${c.Nome} (${c.Fantasia})` : c.Nome;
      return `
        <div class="pos-result-item" onclick="selectOrdClient(${c.CodEntidade}, '${escapeJsString(c.Nome)}')">
          <div class="pos-result-name">#${c.CodEntidade} - ${escapeHtml(displayName)}</div>
          <div class="pos-result-meta">${c.CPF || c.CGC || ''} ${c.Cidade ? ' - ' + c.Cidade : ''}</div>
        </div>
      `;
    }).join('');
    overlay.style.display = 'block';
  } catch (err) {
    console.error('Erro na busca de clientes do pedido:', err);
  }
}

function selectOrdClient(id, name) {
  document.getElementById('Ord_CodEntidade').value = id;
  document.getElementById('ord-client-search-input').value = `#${id} - ${name}`;
  document.getElementById('ord-client-results-overlay').style.display = 'none';
}

async function searchOrdProducts(q) {
  try {
    const res = await fetch(`/api/pdv/produtos?q=${encodeURIComponent(q)}`);
    const products = await res.json();
    const overlay = document.getElementById('ord-product-results-overlay');
    if (!res.ok || !products.length) {
      overlay.style.display = 'none';
      return;
    }

    overlay.innerHTML = products.map(p => `
      <div class="pos-result-item" onclick="selectOrdProduct(${p.CodPrd}, '${escapeJsString(p.Descricao_Produto)}', ${p.Venda})">
        <div>
          <div class="pos-result-name">#${p.CodPrd} - ${escapeHtml(p.Descricao_Produto)}</div>
          <div class="pos-result-meta">EAN: ${p.CodBar || 'SEM EAN'} | Preço: R$ ${floatOrZero(p.Venda).toFixed(2)}</div>
        </div>
        <strong style="color: var(--accent-emerald);">R$ ${floatOrZero(p.Venda).toFixed(2)}</strong>
      </div>
    `).join('');
    overlay.style.display = 'block';
  } catch (err) {
    console.error('Erro na busca de produtos do pedido:', err);
  }
}

function selectOrdProduct(id, desc, price) {
  ordSelectedProduct = { CodPrd: id, Descricao_Produto: desc, Venda: price };
  document.getElementById('ord-product-search-input').value = `#${id} - ${desc}`;
  document.getElementById('ord-item-price').value = floatOrZero(price).toFixed(2);
  document.getElementById('ord-product-results-overlay').style.display = 'none';
  document.getElementById('ord-item-qty').focus();
}

function addOrderItemFromInputs() {
  const prdInput = document.getElementById('ord-product-search-input');
  const qtyInput = document.getElementById('ord-item-qty');
  const priceInput = document.getElementById('ord-item-price');

  if (!ordSelectedProduct || !ordSelectedProduct.CodPrd) {
    showToast('Selecione um produto para adicionar ao pedido', 'error');
    return;
  }

  const qty = parseFloat(qtyInput.value) || 1.0;
  const unitPrice = parseFloat(priceInput.value) || 0.0;

  currentOrderItems.push({
    CodPrd: ordSelectedProduct.CodPrd,
    Descricao_Produto: ordSelectedProduct.Descricao_Produto,
    Qtd: qty,
    ValorUnit: unitPrice,
    Desconto: 0.0,
    Valor: qty * unitPrice
  });

  prdInput.value = '';
  qtyInput.value = '1';
  priceInput.value = '0.00';
  ordSelectedProduct = null;

  renderOrderItemsTable();
  showToast('Item adicionado ao pedido', 'success');
}

function removeOrderItem(idx) {
  currentOrderItems.splice(idx, 1);
  renderOrderItemsTable();
}

function renderOrderItemsTable() {
  const tbody = document.getElementById('ord-items-tbody');
  const badge = document.getElementById('ord-items-badge');
  if (badge) badge.innerText = `${currentOrderItems.length} item(ns)`;

  if (!currentOrderItems.length) {
    tbody.innerHTML = `
      <tr>
        <td colspan="8" class="loading-td">Nenhum item adicionado ao pedido.</td>
      </tr>
    `;
    calcOrderTotals();
    return;
  }

  tbody.innerHTML = currentOrderItems.map((item, idx) => {
    const sub = item.Qtd * item.ValorUnit - (item.Desconto || 0.0);
    return `
      <tr>
        <td><strong>${idx + 1}</strong></td>
        <td>#${item.CodPrd}</td>
        <td><strong>${escapeHtml(item.Descricao_Produto)}</strong></td>
        <td style="text-align: right;">${item.Qtd}</td>
        <td style="text-align: right;">R$ ${item.ValorUnit.toFixed(2)}</td>
        <td style="text-align: right;">R$ ${(item.Desconto || 0.0).toFixed(2)}</td>
        <td style="text-align: right; font-weight: 700; color: var(--accent-emerald);">R$ ${sub.toFixed(2)}</td>
        <td style="text-align: center;">
          <button type="button" class="btn btn-danger btn-sm" onclick="removeOrderItem(${idx})"><i class="fa-solid fa-xmark"></i></button>
        </td>
      </tr>
    `;
  }).join('');

  calcOrderTotals();
}

function calcOrderTotals() {
  const subtotal = currentOrderItems.reduce((sum, item) => sum + (item.Qtd * item.ValorUnit - (item.Desconto || 0.0)), 0.0);
  const desc = parseFloat(document.getElementById('Ord_Desconto').value) || 0.0;
  const frete = parseFloat(document.getElementById('Ord_ValorFrete').value) || 0.0;
  const total = Math.max(0.0, subtotal - desc + frete);

  const displayEl = document.getElementById('Ord_TotalDisplay');
  if (displayEl) {
    displayEl.value = `R$ ${total.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  }
}

async function saveOrder(event) {
  event.preventDefault();

  if (!currentOrderItems.length) {
    showToast('Adicione pelo menos 1 produto ao pedido', 'error');
    return;
  }

  const codPed = document.getElementById('Ord_CodPed').value;

  let entId = parseInt(document.getElementById('Ord_CodEntidade').value, 10);
  const clientInputVal = document.getElementById('ord-client-search-input').value.trim();
  if ((!entId || entId === 1) && clientInputVal) {
    const match = clientInputVal.match(/^#?(\d+)/);
    if (match) {
      entId = parseInt(match[1], 10);
    }
  }
  if (!entId) entId = 1;

  const payload = {
    CodEntidade: entId,
    Vendedor: parseInt(document.getElementById('Ord_Vendedor').value, 10) || 1,
    Transportadora: parseInt(document.getElementById('Ord_Transportadora').value, 10) || 1,
    DataEmiss: document.getElementById('Ord_DataEmiss').value,
    CondPgto: document.getElementById('Ord_CondPgto').value,
    Cfo: document.getElementById('Ord_Cfo').value,
    Obs: document.getElementById('Ord_Obs').value,
    Desconto: parseFloat(document.getElementById('Ord_Desconto').value) || 0.0,
    ValorFrete: parseFloat(document.getElementById('Ord_ValorFrete').value) || 0.0,
    items: currentOrderItems
  };

  const url = codPed ? `/api/pedidos/${codPed}` : '/api/pedidos';
  const method = codPed ? 'PUT' : 'POST';

  try {
    const res = await fetch(url, {
      method: method,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao salvar pedido');

    showToast(result.message || 'Pedido salvo com sucesso!', 'success');
    closeOrderModal();
    fetchOrders();
  } catch (err) {
    showToast(`Erro ao salvar pedido: ${err.message}`, 'error');
  }
}

async function printCurrentOrderFromModal() {
  const codPed = document.getElementById('Ord_CodPed').value;
  if (!codPed) {
    showToast('Salve o pedido antes de imprimir', 'warning');
    return;
  }
  await imprimirCupomOrder(codPed);
}

async function imprimirCupomOrder(codPed) {
  try {
    const id = codPed || (document.getElementById('Ord_CodPed') ? document.getElementById('Ord_CodPed').value : null);
    if (!id) {
      showToast('Nenhum pedido selecionado para impressão', 'warning');
      return;
    }

    const res = await fetch(`/api/pedidos/${id}`);
    const data = await res.json();
    if (!res.ok || !data || !data.pedido) {
      throw new Error((data && data.error) || 'Pedido não encontrado');
    }

    const saleData = {
      venda: data.pedido,
      pedido: data.pedido,
      itens: data.itens || [],
      nfe: data.nfe || {},
      empresa: data.empresa || {}
    };

    lastSaleData = saleData;
    renderThermalReceipt(saleData, 'venda', false);
    printSimpleSaleReceipt(false);
  } catch (err) {
    console.error('Erro ao imprimir pedido:', err);
    showToast(`Erro ao imprimir pedido: ${err.message}`, 'error');
  }
}

async function emitNfeFromOrder(codPed) {
  if (!confirm(`Deseja emitir a NF-e (Modelo 55 - SEFAZ 4.00) para o Pedido #${codPed}?`)) return;

  try {
    const res = await fetch(`/api/pedidos/${codPed}/emitir-nfe`, { method: 'POST' });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao emitir NF-e');

    showToast(result.message || 'NF-e emitida com sucesso!', 'success');
    fetchOrders();
    viewNfeDetails(codPed);
  } catch (err) {
    showToast(`Erro ao emitir NF-e: ${err.message}`, 'error');
  }
}

async function viewNfeDetails(codPed) {
  try {
    const res = await fetch(`/api/pedidos/${codPed}`);
    const data = await res.json();
    if (!res.ok || !data.nfe) throw new Error('Registro de NF-e não encontrado');

    const nfe = data.nfe;
    document.getElementById('nfe-modal-chave').innerText = (nfe.NroChave || '').replace(/(.{4})/g, '$1 ').trim();
    document.getElementById('nfe-modal-protocolo').innerText = nfe.Protocolo || 'N/A';
    document.getElementById('nfe-modal-status').innerText = `${nfe.Status || '100'} - ${nfe.Mensagem || 'Autorizada'}`;
    
    document.getElementById('nfe-modal-xml-text').value = data.nfe.Xml || `<?xml version="1.0" encoding="UTF-8"?>\n<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00">\n  <NFe>\n    <infNFe Id="NFe${nfe.NroChave}" versao="4.00">\n      <!-- Dados do Pedido #${codPed} -->\n    </infNFe>\n  </NFe>\n  <protNFe>\n    <infProt>\n      <chNFe>${nfe.NroChave}</chNFe>\n      <nProt>${nfe.Protocolo}</nProt>\n      <cStat>100</cStat>\n    </infProt>\n  </protNFe>\n</nfeProc>`;

    document.getElementById('nfe-view-modal').classList.add('active');
  } catch (err) {
    showToast(`Erro ao visualizar NF-e: ${err.message}`, 'error');
  }
}

function closeNfeViewModal() {
  document.getElementById('nfe-view-modal').classList.remove('active');
}


// ==========================================
// 6. CFOP (CFO) MODULE LOGIC
// ==========================================

let cfopSearchTimeout = null;
let cfopQuery = '';

function debounceCfopSearch(val) {
  clearTimeout(cfopSearchTimeout);
  cfopQuery = val;
  cfopSearchTimeout = setTimeout(() => fetchCfops(), 250);
}

async function fetchCfops() {
  const tbody = document.getElementById('table-body-cfop');
  if (!tbody) return;

  tbody.innerHTML = `
    <tr>
      <td colspan="3" class="loading-td">
        <i class="fa-solid fa-spinner fa-spin fa-2x"></i>
        <div style="margin-top: 0.5rem;">Carregando CFOPs...</div>
      </td>
    </tr>
  `;

  try {
    const res = await fetch(`/api/cfop?q=${encodeURIComponent(cfopQuery)}`);
    const cfops = await res.json();
    if (!res.ok) throw new Error(cfops.error || 'Erro ao carregar CFOPs');
    renderCfopTable(cfops);
  } catch (err) {
    tbody.innerHTML = `
      <tr>
        <td colspan="3" style="text-align: center; color: var(--accent-rose); padding: 2rem;">
          <i class="fa-solid fa-circle-exclamation fa-2x"></i>
          <div style="margin-top: 0.5rem;">Erro ao carregar CFOPs: ${escapeHtml(err.message)}</div>
        </td>
      </tr>
    `;
  }
}

function renderCfopTable(items) {
  const tbody = document.getElementById('table-body-cfop');
  if (!tbody) return;

  if (!items || !items.length) {
    tbody.innerHTML = `
      <tr>
        <td colspan="3" class="loading-td">Nenhum CFOP encontrado.</td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = items.map(c => `
    <tr>
      <td><strong style="color: var(--accent-blue);">CFOP ${escapeHtml(c.Codigo)}</strong></td>
      <td>${escapeHtml(c.Descricao)}</td>
      <td style="text-align: center;">
        <div style="display: flex; justify-content: center; gap: 0.35rem;">
          <button class="btn btn-secondary btn-sm" onclick="editCfop('${escapeJsString(c.Codigo)}')" title="Editar CFOP">
            <i class="fa-solid fa-pen"></i> Alterar
          </button>
          <button class="btn btn-danger btn-sm" onclick="deleteCfop('${escapeJsString(c.Codigo)}')" title="Excluir CFOP">
            <i class="fa-solid fa-trash"></i>
          </button>
        </div>
      </td>
    </tr>
  `).join('');
}

function openCfopModal(item = null) {
  const form = document.getElementById('cfop-form');
  if (form) form.reset();

  const titleEl = document.getElementById('modal-title-cfop');
  const codeInp = document.getElementById('Cfop_Codigo');

  if (item) {
    titleEl.innerHTML = `<i class="fa-solid fa-file-pen"></i> Editar CFOP #${item.Codigo}`;
    codeInp.value = item.Codigo || '';
    codeInp.readOnly = true;
    document.getElementById('Cfop_Descricao').value = item.Descricao || '';
  } else {
    titleEl.innerHTML = `<i class="fa-solid fa-file-circle-plus"></i> Cadastro de CFOP (Tabela CFO)`;
    codeInp.readOnly = false;
  }

  document.getElementById('cfop-modal').classList.add('active');
}

function closeCfopModal() {
  document.getElementById('cfop-modal').classList.remove('active');
}

async function editCfop(code) {
  try {
    const res = await fetch(`/api/cfop/${code}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'CFOP não encontrado');
    openCfopModal(data);
  } catch (err) {
    showToast(`Erro ao abrir CFOP: ${err.message}`, 'error');
  }
}

async function saveCfop(e) {
  e.preventDefault();
  const form = e.target;
  const formData = new FormData(form);
  const data = Object.fromEntries(formData.entries());

  try {
    const res = await fetch('/api/cfop', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao salvar CFOP');

    showToast(result.message || 'CFOP salvo com sucesso!', 'success');
    closeCfopModal();
    fetchCfops();
  } catch (err) {
    showToast(`Erro ao salvar CFOP: ${err.message}`, 'error');
  }
}

async function deleteCfop(code) {
  if (!confirm(`Deseja realmente excluir o CFOP #${code}?`)) return;

  try {
    const res = await fetch(`/api/cfop/${code}`, { method: 'DELETE' });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao excluir CFOP');

    showToast(result.message, 'success');
    fetchCfops();
  } catch (err) {
    showToast(`Erro ao excluir: ${err.message}`, 'error');
  }
}


// ==========================================
// 7. FORMAS DE PAGAMENTO MODULE LOGIC
// ==========================================

let formasPgtoSearchTimeout = null;
let formasPgtoQuery = '';

function debounceFormasPgtoSearch(val) {
  clearTimeout(formasPgtoSearchTimeout);
  formasPgtoQuery = val;
  formasPgtoSearchTimeout = setTimeout(() => fetchFormasPgto(), 250);
}

async function fetchFormasPgto() {
  const tbody = document.getElementById('table-body-formas-pgto');
  if (!tbody) return;

  tbody.innerHTML = `
    <tr>
      <td colspan="8" class="loading-td">
        <i class="fa-solid fa-spinner fa-spin fa-2x"></i>
        <div style="margin-top: 0.5rem;">Carregando formas de pagamento...</div>
      </td>
    </tr>
  `;

  try {
    const res = await fetch(`/api/formas-pgto?all=1&q=${encodeURIComponent(formasPgtoQuery)}`);
    const formas = await res.json();
    if (!res.ok) throw new Error(formas.error || 'Erro ao carregar formas de pagamento');
    renderFormasPgtoTable(formas);
  } catch (err) {
    tbody.innerHTML = `
      <tr>
        <td colspan="8" style="text-align: center; color: var(--accent-rose); padding: 2rem;">
          <i class="fa-solid fa-circle-exclamation fa-2x"></i>
          <div style="margin-top: 0.5rem;">Erro ao carregar formas de pagamento: ${escapeHtml(err.message)}</div>
        </td>
      </tr>
    `;
  }
}

function renderFormasPgtoTable(items) {
  const tbody = document.getElementById('table-body-formas-pgto');
  if (!tbody) return;

  if (!items || !items.length) {
    tbody.innerHTML = `
      <tr>
        <td colspan="8" class="loading-td">Nenhuma forma de pagamento cadastrada.</td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = items.map(f => `
    <tr>
      <td>#${f.id_forma}</td>
      <td><strong>${escapeHtml(f.Codigo)}</strong></td>
      <td><strong style="color: var(--text-primary);">${escapeHtml(f.Nome)}</strong></td>
      <td><span class="badge badge-group">${escapeHtml(f.Tipo || 'A_VISTA')}</span></td>
      <td style="text-align: right;">${f.Parcelas || 1}x</td>
      <td style="text-align: right;">${f.DiasEntreParcelas || 0} dias</td>
      <td style="text-align: center;">
        <span class="badge ${f.Ativo === 1 ? 'badge-active' : 'badge-inactive'}">
          ${f.Ativo === 1 ? 'Ativo' : 'Inativo'}
        </span>
      </td>
      <td style="text-align: center;">
        <div style="display: flex; justify-content: center; gap: 0.35rem;">
          <button class="btn btn-secondary btn-sm" onclick="editFormaPgto(${f.id_forma})" title="Editar Forma de Pagamento">
            <i class="fa-solid fa-pen"></i> Alterar
          </button>
          <button class="btn btn-danger btn-sm" onclick="deleteFormaPgto(${f.id_forma})" title="Excluir">
            <i class="fa-solid fa-trash"></i>
          </button>
        </div>
      </td>
    </tr>
  `).join('');
}

function openFormaPgtoModal(item = null) {
  const form = document.getElementById('forma-pgto-form');
  if (form) form.reset();

  const titleEl = document.getElementById('modal-title-forma');

  if (item) {
    titleEl.innerHTML = `<i class="fa-solid fa-credit-card"></i> Alterar Forma de Pagamento #${item.id_forma}`;
    document.getElementById('Fpg_id_forma').value = item.id_forma;
    document.getElementById('Fpg_Codigo').value = item.Codigo || '';
    document.getElementById('Fpg_Nome').value = item.Nome || '';
    document.getElementById('Fpg_Tipo').value = item.Tipo || 'A_VISTA';
    document.getElementById('Fpg_Parcelas').value = item.Parcelas || 1;
    document.getElementById('Fpg_DiasEntreParcelas').value = item.DiasEntreParcelas || 0;
    document.getElementById('Fpg_Ativo').checked = item.Ativo === 1;
  } else {
    titleEl.innerHTML = `<i class="fa-solid fa-credit-card"></i> Cadastrar Nova Forma de Pagamento`;
    document.getElementById('Fpg_id_forma').value = '';
    document.getElementById('Fpg_Ativo').checked = true;
  }

  document.getElementById('forma-pgto-modal').classList.add('active');
}

function closeFormaPgtoModal() {
  document.getElementById('forma-pgto-modal').classList.remove('active');
}

async function editFormaPgto(id) {
  try {
    const res = await fetch(`/api/formas-pgto/${id}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Forma de pagamento não encontrada');
    openFormaPgtoModal(data);
  } catch (err) {
    showToast(`Erro ao abrir forma de pagamento: ${err.message}`, 'error');
  }
}

async function saveFormaPgto(e) {
  e.preventDefault();
  const form = e.target;
  const formData = new FormData(form);
  const data = Object.fromEntries(formData.entries());
  data.Ativo = document.getElementById('Fpg_Ativo').checked ? 1 : 0;

  try {
    const res = await fetch('/api/formas-pgto', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao salvar forma de pagamento');

    showToast(result.message || 'Forma de pagamento salva com sucesso!', 'success');
    closeFormaPgtoModal();
    fetchFormasPgto();
    loadFormasPgtoOptions();
  } catch (err) {
    showToast(`Erro ao salvar forma de pagamento: ${err.message}`, 'error');
  }
}

async function deleteFormaPgto(id) {
  if (!confirm(`Deseja realmente excluir a Forma de Pagamento #${id}?`)) return;

  try {
    const res = await fetch(`/api/formas-pgto/${id}`, { method: 'DELETE' });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao excluir forma de pagamento');

    showToast(result.message, 'success');
    fetchFormasPgto();
    loadFormasPgtoOptions();
  } catch (err) {
    showToast(`Erro ao excluir: ${err.message}`, 'error');
  }
}

async function loadFormasPgtoOptions() {
  const selectOrder = document.getElementById('Ord_CondPgto');
  if (!selectOrder) return;

  try {
    const res = await fetch('/api/formas-pgto');
    const formas = await res.json();
    if (!res.ok || !formas.length) return;

    selectOrder.innerHTML = formas.map(f => `
      <option value="${escapeHtml(f.Nome)}">${escapeHtml(f.Nome)}</option>
    `).join('');
  } catch (err) {
    console.error('Erro ao carregar opções de pagamento:', err);
  }
}


// ==========================================
// 8. NATUREZAS DE OPERAÇÃO MODULE LOGIC
// ==========================================

let naturezaOpSearchTimeout = null;
let naturezaOpQuery = '';

function debounceNaturezaOpSearch(val) {
  clearTimeout(naturezaOpSearchTimeout);
  naturezaOpQuery = val;
  naturezaOpSearchTimeout = setTimeout(() => fetchNaturezasOperacao(), 250);
}

async function fetchNaturezasOperacao() {
  const tbody = document.getElementById('table-body-natureza-op');
  if (!tbody) return;

  tbody.innerHTML = `
    <tr>
      <td colspan="6" class="loading-td">
        <i class="fa-solid fa-spinner fa-spin fa-2x"></i>
        <div style="margin-top: 0.5rem;">Carregando naturezas de operação...</div>
      </td>
    </tr>
  `;

  try {
    const res = await fetch(`/api/natureza-operacao?all=1&q=${encodeURIComponent(naturezaOpQuery)}`);
    const naturezas = await res.json();
    if (!res.ok) throw new Error(naturezas.error || 'Erro ao carregar naturezas de operação');
    renderNaturezaOpTable(naturezas);
  } catch (err) {
    tbody.innerHTML = `
      <tr>
        <td colspan="6" style="text-align: center; color: var(--accent-rose); padding: 2rem;">
          <i class="fa-solid fa-circle-exclamation fa-2x"></i>
          <div style="margin-top: 0.5rem;">Erro ao carregar naturezas de operação: ${escapeHtml(err.message)}</div>
        </td>
      </tr>
    `;
  }
}

function renderNaturezaOpTable(items) {
  const tbody = document.getElementById('table-body-natureza-op');
  if (!tbody) return;

  if (!items || !items.length) {
    tbody.innerHTML = `
      <tr>
        <td colspan="6" class="loading-td">Nenhuma natureza de operação cadastrada.</td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = items.map(n => `
    <tr>
      <td>#${n.id_natureza}</td>
      <td><strong>${escapeHtml(n.Codigo)}</strong></td>
      <td><strong style="color: var(--text-primary);">${escapeHtml(n.Descricao)}</strong></td>
      <td><span class="badge badge-group">CFOP ${escapeHtml(n.CFOP || '5102')}</span></td>
      <td style="text-align: center;">
        <span class="badge ${n.Ativo === 1 ? 'badge-active' : 'badge-inactive'}">
          ${n.Ativo === 1 ? 'Ativa' : 'Inativa'}
        </span>
      </td>
      <td style="text-align: center;">
        <div style="display: flex; justify-content: center; gap: 0.35rem;">
          <button class="btn btn-secondary btn-sm" onclick="editNaturezaOp(${n.id_natureza})" title="Editar Natureza de Operação">
            <i class="fa-solid fa-pen"></i> Alterar
          </button>
          <button class="btn btn-danger btn-sm" onclick="deleteNaturezaOp(${n.id_natureza})" title="Excluir">
            <i class="fa-solid fa-trash"></i>
          </button>
        </div>
      </td>
    </tr>
  `).join('');
}

function openNaturezaOpModal(item = null) {
  const form = document.getElementById('natureza-op-form');
  if (form) form.reset();

  const titleEl = document.getElementById('modal-title-nat');

  if (item) {
    titleEl.innerHTML = `<i class="fa-solid fa-route"></i> Alterar Natureza de Operação #${item.id_natureza}`;
    document.getElementById('Nat_id_natureza').value = item.id_natureza;
    document.getElementById('Nat_Codigo').value = item.Codigo || '';
    document.getElementById('Nat_Descricao').value = item.Descricao || '';
    document.getElementById('Nat_CFOP').value = item.CFOP || '5102';
    document.getElementById('Nat_Ativo').checked = item.Ativo === 1;
  } else {
    titleEl.innerHTML = `<i class="fa-solid fa-route"></i> Cadastrar Nova Natureza de Operação`;
    document.getElementById('Nat_id_natureza').value = '';
    document.getElementById('Nat_Ativo').checked = true;
  }

  document.getElementById('natureza-op-modal').classList.add('active');
}

function closeNaturezaOpModal() {
  document.getElementById('natureza-op-modal').classList.remove('active');
}

async function editNaturezaOp(id) {
  try {
    const res = await fetch(`/api/natureza-operacao/${id}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Natureza de operação não encontrada');
    openNaturezaOpModal(data);
  } catch (err) {
    showToast(`Erro ao abrir natureza de operação: ${err.message}`, 'error');
  }
}

async function saveNaturezaOp(e) {
  e.preventDefault();
  const form = e.target;
  const formData = new FormData(form);
  const data = Object.fromEntries(formData.entries());
  data.Ativo = document.getElementById('Nat_Ativo').checked ? 1 : 0;

  try {
    const res = await fetch('/api/natureza-operacao', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao salvar natureza de operação');

    showToast(result.message || 'Natureza de operação salva com sucesso!', 'success');
    closeNaturezaOpModal();
    fetchNaturezasOperacao();
    loadOrdNaturezasOptions();
  } catch (err) {
    showToast(`Erro ao salvar natureza de operação: ${err.message}`, 'error');
  }
}

async function deleteNaturezaOp(id) {
  if (!confirm(`Deseja realmente excluir a Natureza de Operação #${id}?`)) return;

  try {
    const res = await fetch(`/api/natureza-operacao/${id}`, { method: 'DELETE' });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao excluir natureza de operação');

    showToast(result.message, 'success');
    fetchNaturezasOperacao();
    loadOrdNaturezasOptions();
  } catch (err) {
    showToast(`Erro ao excluir: ${err.message}`, 'error');
  }
}


// ==========================================
// 9. AUTHENTICATION & USER MANAGEMENT LOGIC
// ==========================================

let currentUser = null;

async function checkAuth() {
  try {
    const localUser = localStorage.getItem('vendas_user');
    if (localUser) {
      try {
        currentUser = JSON.parse(localUser);
      } catch (e) {}
    }

    const res = await fetch('/api/auth/me');
    const data = await res.json();
    if (data.authenticated && data.user) {
      currentUser = data.user;
      localStorage.setItem('vendas_user', JSON.stringify(currentUser));
    }

    if (currentUser) {
      const loginModal = document.getElementById('login-modal');
      if (loginModal) {
        loginModal.classList.remove('active');
        loginModal.style.display = 'none';
      }
      setElementText('user-display-name', currentUser.Nome || currentUser.Login);
      setElementText('user-display-role', currentUser.Perfil || 'OPERADOR');
      applyRolePermissions();
    } else {
      const loginModal = document.getElementById('login-modal');
      if (loginModal) {
        loginModal.classList.add('active');
        loginModal.style.display = 'flex';
      }
    }
  } catch (err) {
    console.error('Erro na checagem de sessão:', err);
    const localUser = localStorage.getItem('vendas_user');
    if (localUser) {
      try {
        currentUser = JSON.parse(localUser);
        const loginModal = document.getElementById('login-modal');
        if (loginModal) {
          loginModal.classList.remove('active');
          loginModal.style.display = 'none';
        }
        setElementText('user-display-name', currentUser.Nome || currentUser.Login);
        setElementText('user-display-role', currentUser.Perfil || 'OPERADOR');
        applyRolePermissions();
        return;
      } catch (e) {}
    }
    const loginModal = document.getElementById('login-modal');
    if (loginModal) {
      loginModal.classList.add('active');
      loginModal.style.display = 'flex';
    }
  }
}

async function login(e) {
  if (e) {
    e.preventDefault();
    e.stopPropagation();
  }
  const loginVal = document.getElementById('login-input-username').value.trim();
  const senhaVal = document.getElementById('login-input-password').value.trim();
  const errBox = document.getElementById('login-error-msg');
  if (errBox) errBox.style.display = 'none';

  try {
    const res = await fetch('/api/auth/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ login: loginVal, senha: senhaVal })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Erro ao realizar login');

    currentUser = data.user;
    localStorage.setItem('vendas_user', JSON.stringify(currentUser));
    
    showToast(data.message || 'Login realizado com sucesso!', 'success');

    const loginModal = document.getElementById('login-modal');
    if (loginModal) {
      loginModal.classList.remove('active');
      loginModal.style.display = 'none';
    }
    setElementText('user-display-name', currentUser.Nome || currentUser.Login);
    setElementText('user-display-role', currentUser.Perfil || 'OPERADOR');
    applyRolePermissions();
    refreshCurrentModule();
  } catch (err) {
    if (errBox) {
      errBox.innerText = err.message;
      errBox.style.display = 'block';
    } else {
      showToast(err.message, 'error');
    }
  }
  return false;
}

async function logout() {
  try {
    await fetch('/api/auth/logout', { method: 'POST' });
  } catch (err) {
    console.error('Erro ao encerrar sessão:', err);
  }
  currentUser = null;
  localStorage.removeItem('vendas_user');
  const loginModal = document.getElementById('login-modal');
  if (loginModal) {
    loginModal.classList.add('active');
    loginModal.style.display = 'flex';
  }
  showToast('Sessão encerrada com sucesso.', 'info');
}

function applyRolePermissions() {
  if (!currentUser) return;
  const role = (currentUser.Perfil || 'OPERADOR').toUpperCase();

  const navUsuarios = document.getElementById('nav-usuarios');
  const navConfigDb = document.getElementById('nav-config-db');

  if (role === 'ADMIN') {
    if (navUsuarios) navUsuarios.style.display = 'flex';
    if (navConfigDb) navConfigDb.style.display = 'flex';
  } else if (role === 'GERENTE') {
    if (navUsuarios) navUsuarios.style.display = 'flex';
    if (navConfigDb) navConfigDb.style.display = 'none';
  } else {
    if (navUsuarios) navUsuarios.style.display = 'none';
    if (navConfigDb) navConfigDb.style.display = 'none';
  }
}

// User CRUD
let usuariosSearchTimeout = null;
let usuariosQuery = '';

function debounceUsuariosSearch(val) {
  clearTimeout(usuariosSearchTimeout);
  usuariosQuery = val;
  usuariosSearchTimeout = setTimeout(() => fetchUsuarios(), 250);
}

async function fetchUsuarios() {
  const tbody = document.getElementById('table-body-usuarios');
  if (!tbody) return;

  tbody.innerHTML = `
    <tr>
      <td colspan="7" class="loading-td">
        <i class="fa-solid fa-spinner fa-spin fa-2x"></i>
        <div style="margin-top: 0.5rem;">Carregando usuários...</div>
      </td>
    </tr>
  `;

  try {
    const res = await fetch(`/api/usuarios?q=${encodeURIComponent(usuariosQuery)}`);
    const users = await res.json();
    if (!res.ok) throw new Error(users.error || 'Erro ao carregar usuários');
    renderUsuariosTable(users);
  } catch (err) {
    tbody.innerHTML = `
      <tr>
        <td colspan="7" style="text-align: center; color: var(--accent-rose); padding: 2rem;">
          <i class="fa-solid fa-circle-exclamation fa-2x"></i>
          <div style="margin-top: 0.5rem;">Erro ao carregar usuários: ${escapeHtml(err.message)}</div>
        </td>
      </tr>
    `;
  }
}

function renderUsuariosTable(items) {
  const tbody = document.getElementById('table-body-usuarios');
  if (!tbody) return;

  if (!items || !items.length) {
    tbody.innerHTML = `
      <tr>
        <td colspan="7" class="loading-td">Nenhum usuário cadastrado.</td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = items.map(u => `
    <tr>
      <td>#${u.id_usuario}</td>
      <td><strong>${escapeHtml(u.Login)}</strong></td>
      <td><strong style="color: var(--text-primary);">${escapeHtml(u.Nome)}</strong></td>
      <td style="text-align: center;"><span class="badge badge-group">${escapeHtml(u.Perfil || 'OPERADOR')}</span></td>
      <td>${escapeHtml(u.DtCadastro || '-')}</td>
      <td style="text-align: center;">
        <span class="badge ${u.Ativo === 1 ? 'badge-active' : 'badge-inactive'}">
          ${u.Ativo === 1 ? 'Ativo' : 'Inativo'}
        </span>
      </td>
      <td style="text-align: center;">
        <div style="display: flex; justify-content: center; gap: 0.35rem;">
          <button class="btn btn-secondary btn-sm" onclick="editUsuario(${u.id_usuario})" title="Editar Usuário">
            <i class="fa-solid fa-pen"></i> Alterar
          </button>
          <button class="btn btn-danger btn-sm" onclick="deleteUsuario(${u.id_usuario})" title="Excluir">
            <i class="fa-solid fa-trash"></i>
          </button>
        </div>
      </td>
    </tr>
  `).join('');
}

function openUsuarioModal(item = null) {
  const form = document.getElementById('usuario-form');
  if (form) form.reset();

  const titleEl = document.getElementById('modal-title-usr');

  if (item) {
    titleEl.innerHTML = `<i class="fa-solid fa-user-shield"></i> Alterar Usuário #${item.id_usuario}`;
    document.getElementById('Usr_id_usuario').value = item.id_usuario;
    document.getElementById('Usr_Login').value = item.Login || '';
    document.getElementById('Usr_Nome').value = item.Nome || '';
    document.getElementById('Usr_Perfil').value = item.Perfil || 'OPERADOR';
    document.getElementById('Usr_Senha').value = '';
    document.getElementById('Usr_Ativo').checked = item.Ativo === 1;
  } else {
    titleEl.innerHTML = `<i class="fa-solid fa-user-plus"></i> Cadastrar Novo Usuário`;
    document.getElementById('Usr_id_usuario').value = '';
    document.getElementById('Usr_Ativo').checked = true;
  }

  document.getElementById('usuario-modal').classList.add('active');
}

function closeUsuarioModal() {
  document.getElementById('usuario-modal').classList.remove('active');
}

async function editUsuario(id) {
  try {
    const res = await fetch(`/api/usuarios/${id}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Usuário não encontrado');
    openUsuarioModal(data);
  } catch (err) {
    showToast(`Erro ao abrir usuário: ${err.message}`, 'error');
  }
}

async function saveUsuario(e) {
  e.preventDefault();
  const form = e.target;
  const formData = new FormData(form);
  const data = Object.fromEntries(formData.entries());
  data.Ativo = document.getElementById('Usr_Ativo').checked ? 1 : 0;

  try {
    const res = await fetch('/api/usuarios', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao salvar usuário');

    showToast(result.message || 'Usuário salvo com sucesso!', 'success');
    closeUsuarioModal();
    fetchUsuarios();
  } catch (err) {
    showToast(`Erro ao salvar usuário: ${err.message}`, 'error');
  }
}

async function deleteUsuario(id) {
  if (!confirm(`Deseja realmente excluir o Usuário #${id}?`)) return;

  try {
    const res = await fetch(`/api/usuarios/${id}`, { method: 'DELETE' });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao excluir usuário');

    showToast(result.message, 'success');
    fetchUsuarios();
  } catch (err) {
    showToast(`Erro ao excluir: ${err.message}`, 'error');
  }
}


// ==========================================
// 10. DATABASE CONFIGURATION LOGIC
// ==========================================

async function loadDbConfig() {
  const cardBody = document.getElementById('db-status-card-body');
  if (!cardBody) return;

  try {
    const res = await fetch('/api/config/db');
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Erro ao consultar banco de dados');

    document.getElementById('Cfg_Host').value = data.host || 'localhost';
    document.getElementById('Cfg_Port').value = data.port || 5432;
    document.getElementById('Cfg_DbName').value = data.dbname || 'vendas_db';
    document.getElementById('Cfg_User').value = data.user || 'postgres';

    cardBody.innerHTML = `
      <div style="display: flex; flex-direction: column; gap: 1rem;">
        <div style="background: rgba(16, 185, 129, 0.1); border: 1px solid rgba(16, 185, 129, 0.3); padding: 1rem; border-radius: var(--radius-md); display: flex; align-items: center; gap: 0.75rem;">
          <i class="fa-solid fa-circle-check fa-2x" style="color: var(--accent-emerald);"></i>
          <div>
            <strong style="color: var(--accent-emerald); font-size: 1.05rem;">PostgreSQL 18 Operacional</strong>
            <div style="font-size: 0.85rem; color: var(--text-secondary); margin-top: 0.25rem;">${escapeHtml(data.version || '')}</div>
          </div>
        </div>

        <div class="form-grid">
          <div class="form-group col-6" style="background: var(--bg-surface-elevated); padding: 0.75rem; border-radius: var(--radius-md); border: 1px solid var(--border-color);">
            <div style="font-size: 0.75rem; color: var(--text-muted);">HOST / SERVIDOR</div>
            <strong style="font-size: 0.95rem;">${escapeHtml(data.host)}:${data.port}</strong>
          </div>

          <div class="form-group col-6" style="background: var(--bg-surface-elevated); padding: 0.75rem; border-radius: var(--radius-md); border: 1px solid var(--border-color);">
            <div style="font-size: 0.75rem; color: var(--text-muted);">BANCO DE DADOS</div>
            <strong style="font-size: 0.95rem;">${escapeHtml(data.dbname)}</strong>
          </div>

          <div class="form-group col-6" style="background: var(--bg-surface-elevated); padding: 0.75rem; border-radius: var(--radius-md); border: 1px solid var(--border-color);">
            <div style="font-size: 0.75rem; color: var(--text-muted);">TOTAL DE TABELAS</div>
            <strong style="font-size: 1.1rem; color: var(--accent-blue);">${data.tables_count} tabelas</strong>
          </div>

          <div class="form-group col-6" style="background: var(--bg-surface-elevated); padding: 0.75rem; border-radius: var(--radius-md); border: 1px solid var(--border-color);">
            <div style="font-size: 0.75rem; color: var(--text-muted);">TOTAL DE PRODUTOS</div>
            <strong style="font-size: 1.1rem; color: var(--accent-emerald);">${(data.records_summary ? data.records_summary.produtos : 0).toLocaleString()} itens</strong>
          </div>
        </div>
      </div>
    `;
  } catch (err) {
    cardBody.innerHTML = `
      <div style="color: var(--accent-rose); background: rgba(244, 63, 94, 0.1); border: 1px solid rgba(244, 63, 94, 0.3); padding: 1rem; border-radius: var(--radius-md);">
        <i class="fa-solid fa-circle-exclamation fa-2x"></i>
        <div style="margin-top: 0.5rem; font-weight: 600;">Falha na Conexão: ${escapeHtml(err.message)}</div>
      </div>
    `;
  }
}

async function testDbConnection() {
  const host = document.getElementById('Cfg_Host').value.trim();
  const port = document.getElementById('Cfg_Port').value.trim();
  const dbname = document.getElementById('Cfg_DbName').value.trim();
  const user = document.getElementById('Cfg_User').value.trim();
  const password = document.getElementById('Cfg_Password').value.trim();

  try {
    showToast('Testando conexão com PostgreSQL 18...', 'info');
    const res = await fetch('/api/config/db/test', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ host, port, dbname, user, password })
    });
    const data = await res.json();
    if (!res.ok || !data.success) throw new Error(data.message || 'Erro no teste de conexão');

    showToast(`Sucesso! ${data.message} (${data.tables_count} tabelas encontradas)`, 'success');
  } catch (err) {
    showToast(`Erro no teste de conexão: ${err.message}`, 'error');
  }
}

async function saveDbConfig(e) {
  e.preventDefault();
  const serieEl = document.getElementById('Cfg_Card_Nfe_Serie');
  const numEl = document.getElementById('Cfg_Card_Nfe_Numero');
  if (serieEl && numEl) {
    await postPdvConfigSave({
      Nfe: {
        SerieNfe: serieEl.value,
        NroNfe: numEl.value
      }
    }, 'Parâmetros de conexão do banco e NF-e salvos com sucesso!');
  } else {
    showToast('Parâmetros de conexão salvos e ativos no sistema!', 'success');
  }
  loadDbConfig();
}

// ==========================================
// 10.1 DATABASE SYNC (TEST <-> PROD) LOGIC
// ==========================================

let cachedDbTables = [];
let lastSyncDiffs = null;

function openDbSyncModal() {
  const modal = document.getElementById('db-sync-modal');
  if (!modal) return;

  // Preencher configurações de origem com os dados atuais do form
  const srcHost = document.getElementById('Cfg_Host') ? document.getElementById('Cfg_Host').value.trim() : 'localhost';
  const srcPort = document.getElementById('Cfg_Port') ? document.getElementById('Cfg_Port').value.trim() : '5432';
  const srcDb = document.getElementById('Cfg_DbName') ? document.getElementById('Cfg_DbName').value.trim() : 'vendas_db';

  if (document.getElementById('Sync_Src_Host')) document.getElementById('Sync_Src_Host').value = srcHost;
  if (document.getElementById('Sync_Src_Port')) document.getElementById('Sync_Src_Port').value = srcPort;
  if (document.getElementById('Sync_Src_DbName')) document.getElementById('Sync_Src_DbName').value = srcDb;

  if (document.getElementById('Sync_Tgt_Host') && !document.getElementById('Sync_Tgt_Host').value) {
    document.getElementById('Sync_Tgt_Host').value = srcHost;
  }
  if (document.getElementById('Sync_Tgt_Port') && !document.getElementById('Sync_Tgt_Port').value) {
    document.getElementById('Sync_Tgt_Port').value = srcPort;
  }
  if (document.getElementById('Sync_Tgt_DbName') && (!document.getElementById('Sync_Tgt_DbName').value || document.getElementById('Sync_Tgt_DbName').value === 'vendas_prod_db')) {
    document.getElementById('Sync_Tgt_DbName').value = srcDb.endsWith('_db') ? srcDb.replace('_db', '_prod_db') : `${srcDb}_prod`;
  }

  // Resetar resultados anteriores
  lastSyncDiffs = null;
  const resCard = document.getElementById('sync-results-card');
  if (resCard) resCard.style.display = 'none';
  const execLog = document.getElementById('sync-execution-log');
  if (execLog) execLog.style.display = 'none';

  modal.classList.add('active');
  loadSyncTablesList();
}

function closeDbSyncModal() {
  const modal = document.getElementById('db-sync-modal');
  if (modal) modal.classList.remove('active');
}

function toggleSyncTableScope(mode) {
  const wrapper = document.getElementById('sync-tables-selection-wrapper');
  if (!wrapper) return;
  if (mode === 'selected') {
    wrapper.style.display = 'block';
  } else {
    wrapper.style.display = 'none';
  }
}

async function loadSyncTablesList() {
  const container = document.getElementById('sync-tables-list-container');
  if (!container) return;

  const host = document.getElementById('Sync_Src_Host') ? document.getElementById('Sync_Src_Host').value.trim() : '';
  const port = document.getElementById('Sync_Src_Port') ? document.getElementById('Sync_Src_Port').value.trim() : '';
  const dbname = document.getElementById('Sync_Src_DbName') ? document.getElementById('Sync_Src_DbName').value.trim() : '';

  try {
    container.innerHTML = '<div style="color: var(--text-muted); font-size: 0.85rem; padding: 0.5rem;"><i class="fa-solid fa-spinner fa-spin"></i> Carregando tabelas do banco de dados...</div>';
    
    const params = new URLSearchParams();
    if (host) params.append('host', host);
    if (port) params.append('port', port);
    if (dbname) params.append('dbname', dbname);

    const res = await fetch(`/api/config/db/tables?${params.toString()}`);
    const data = await res.json();
    if (!res.ok || !data.success) throw new Error(data.message || 'Erro ao listar tabelas');

    cachedDbTables = data.tables || [];
    renderSyncTablesList(cachedDbTables);
  } catch (err) {
    container.innerHTML = `<div style="color: var(--accent-rose); font-size: 0.85rem; padding: 0.5rem;"><i class="fa-solid fa-circle-exclamation"></i> Falha ao listar tabelas: ${escapeHtml(err.message)}</div>`;
  }
}

function renderSyncTablesList(tables) {
  const container = document.getElementById('sync-tables-list-container');
  if (!container) return;

  if (!tables || tables.length === 0) {
    container.innerHTML = '<div style="color: var(--text-muted); font-size: 0.85rem; padding: 0.5rem;">Nenhuma tabela encontrada no banco de dados.</div>';
    return;
  }

  container.innerHTML = tables.map(t => `
    <label style="display: flex; align-items: center; gap: 0.5rem; background: var(--bg-surface); padding: 0.4rem 0.6rem; border-radius: var(--radius-sm); border: 1px solid var(--border-color); font-size: 0.85rem; cursor: pointer; user-select: none;">
      <input type="checkbox" class="sync-table-checkbox" value="${escapeHtml(t.name)}" checked>
      <span style="font-weight: 600; color: var(--text-primary); flex: 1;">${escapeHtml(t.name)}</span>
      <span style="font-size: 0.75rem; color: var(--text-muted);" title="${t.columns_count} colunas | ${t.rows_count} registros">${t.rows_count} reg.</span>
    </label>
  `).join('');
}

function filterSyncTablesList(query) {
  const cleanQ = (query || '').toLowerCase().trim();
  const checkboxes = document.querySelectorAll('.sync-table-checkbox');
  checkboxes.forEach(cb => {
    const label = cb.closest('label');
    if (!label) return;
    const name = cb.value.toLowerCase();
    if (!cleanQ || name.includes(cleanQ)) {
      label.style.display = 'flex';
    } else {
      label.style.display = 'none';
    }
  });
}

function selectAllSyncTables(checked) {
  const checkboxes = document.querySelectorAll('.sync-table-checkbox');
  checkboxes.forEach(cb => {
    const label = cb.closest('label');
    if (label && label.style.display !== 'none') {
      cb.checked = checked;
    }
  });
}

function getSelectedSyncTables() {
  const modeEl = document.querySelector('input[name="sync_scope_mode"]:checked');
  const mode = modeEl ? modeEl.value : 'all';
  if (mode === 'all') return null;

  const selected = [];
  const checkboxes = document.querySelectorAll('.sync-table-checkbox:checked');
  checkboxes.forEach(cb => selected.push(cb.value));
  return selected;
}

async function analyzeDbSync() {
  const btn = document.getElementById('btn-sync-analyze');
  const resCard = document.getElementById('sync-results-card');
  const badgesBox = document.getElementById('sync-summary-badges');
  const sqlCodeBox = document.getElementById('sync-sql-code');
  const execLog = document.getElementById('sync-execution-log');

  const srcHost = document.getElementById('Sync_Src_Host').value.trim();
  const srcPort = document.getElementById('Sync_Src_Port').value.trim();
  const srcDb = document.getElementById('Sync_Src_DbName').value.trim();
  const srcUser = document.getElementById('Cfg_User') ? document.getElementById('Cfg_User').value.trim() : 'postgres';
  const srcPass = document.getElementById('Cfg_Password') ? document.getElementById('Cfg_Password').value.trim() : 'masterkey';

  const tgtHost = document.getElementById('Sync_Tgt_Host').value.trim();
  const tgtPort = document.getElementById('Sync_Tgt_Port').value.trim();
  const tgtDb = document.getElementById('Sync_Tgt_DbName').value.trim();

  const selectedTables = getSelectedSyncTables();

  try {
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Comparando Estruturas...';
    }
    showToast('Analisando diferenças entre base de Teste e Produção...', 'info');

    const payload = {
      source_host: srcHost,
      source_port: srcPort,
      source_dbname: srcDb,
      source_user: srcUser,
      source_password: srcPass,
      target_host: tgtHost,
      target_port: tgtPort,
      target_dbname: tgtDb,
      target_user: srcUser,
      target_password: srcPass,
      tables: selectedTables
    };

    const res = await fetch('/api/config/db/sync/compare', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();
    if (!res.ok || !data.success) throw new Error(data.message || 'Falha ao analisar bancos');

    lastSyncDiffs = data;

    if (resCard) resCard.style.display = 'block';
    if (execLog) execLog.style.display = 'none';

    const diffs = data.diffs || {};
    const missingTables = diffs.missing_tables || [];
    const missingCols = diffs.missing_columns || [];
    const typeMismatches = diffs.type_mismatches || [];
    const missingIdxs = diffs.missing_indexes || [];
    const sqlStmts = diffs.sql_statements || [];

    if (!data.target_exists) {
      badgesBox.innerHTML = `
        <span class="badge" style="background: rgba(245, 158, 11, 0.15); color: var(--accent-amber); font-size: 0.85rem; padding: 0.4rem 0.75rem; border-radius: var(--radius-sm); border: 1px solid rgba(245, 158, 11, 0.3);">
          <i class="fa-solid fa-triangle-exclamation"></i> Banco de Produção Inexistente (Será Criado)
        </span>
      `;
      sqlCodeBox.textContent = `-- O banco de dados '${tgtDb}' será criado automaticamente e todas as tabelas e dados sincronizados.`;
      showToast(`Atenção: Banco '${tgtDb}' será criado na sincronização.`, 'warning');
      return;
    }

    badgesBox.innerHTML = `
      <span class="badge" style="background: ${missingTables.length ? 'rgba(239, 68, 68, 0.15)' : 'rgba(16, 185, 129, 0.15)'}; color: ${missingTables.length ? 'var(--accent-rose)' : 'var(--accent-emerald)'}; font-size: 0.85rem; padding: 0.4rem 0.75rem; border-radius: var(--radius-sm); border: 1px solid var(--border-color);">
        <i class="fa-solid fa-table"></i> ${missingTables.length} Tabelas Faltantes
      </span>
      <span class="badge" style="background: ${missingCols.length ? 'rgba(245, 158, 11, 0.15)' : 'rgba(16, 185, 129, 0.15)'}; color: ${missingCols.length ? 'var(--accent-amber)' : 'var(--accent-emerald)'}; font-size: 0.85rem; padding: 0.4rem 0.75rem; border-radius: var(--radius-sm); border: 1px solid var(--border-color);">
        <i class="fa-solid fa-columns"></i> ${missingCols.length} Colunas Faltantes
      </span>
      <span class="badge" style="background: ${typeMismatches.length ? 'rgba(168, 85, 247, 0.15)' : 'rgba(16, 185, 129, 0.15)'}; color: ${typeMismatches.length ? '#c084fc' : 'var(--accent-emerald)'}; font-size: 0.85rem; padding: 0.4rem 0.75rem; border-radius: var(--radius-sm); border: 1px solid var(--border-color);">
        <i class="fa-solid fa-code-compare"></i> ${typeMismatches.length} Tipos Divergentes
      </span>
      <span class="badge" style="background: ${missingIdxs.length ? 'rgba(59, 130, 246, 0.15)' : 'rgba(16, 185, 129, 0.15)'}; color: ${missingIdxs.length ? 'var(--accent-blue)' : 'var(--accent-emerald)'}; font-size: 0.85rem; padding: 0.4rem 0.75rem; border-radius: var(--radius-sm); border: 1px solid var(--border-color);">
        <i class="fa-solid fa-key"></i> ${missingIdxs.length} Índices
      </span>
    `;

    if (sqlStmts.length === 0) {
      sqlCodeBox.textContent = `-- PARABÉNS! As estruturas das bases de Teste e Produção já estão 100% IDÊNTICAS e sincronizadas!`;
      showToast('Excelente! Bases já estão 100% sincronizadas.', 'success');
    } else {
      const formattedSql = sqlStmts.map((s, idx) => `-- [${idx + 1}] ${s.description}\n${s.sql}\n`).join('\n');
      sqlCodeBox.textContent = formattedSql;
      showToast(`Diagnóstico concluído: ${sqlStmts.length} alterações identificadas.`, 'info');
    }
  } catch (err) {
    showToast(`Erro ao analisar sincronização: ${err.message}`, 'error');
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<i class="fa-solid fa-magnifying-glass-chart"></i> 1. Analisar Diferenças (Preview)';
    }
  }
}

async function executeDbSync() {
  const btn = document.getElementById('btn-sync-execute');
  const resCard = document.getElementById('sync-results-card');
  const execLog = document.getElementById('sync-execution-log');
  const logContent = document.getElementById('sync-log-content');

  const srcHost = document.getElementById('Sync_Src_Host').value.trim();
  const srcPort = document.getElementById('Sync_Src_Port').value.trim();
  const srcDb = document.getElementById('Sync_Src_DbName').value.trim();
  const srcUser = document.getElementById('Cfg_User') ? document.getElementById('Cfg_User').value.trim() : 'postgres';
  const srcPass = document.getElementById('Cfg_Password') ? document.getElementById('Cfg_Password').value.trim() : 'masterkey';

  const tgtHost = document.getElementById('Sync_Tgt_Host').value.trim();
  const tgtPort = document.getElementById('Sync_Tgt_Port').value.trim();
  const tgtDb = document.getElementById('Sync_Tgt_DbName').value.trim();

  const selectedTables = getSelectedSyncTables();
  const syncData = document.getElementById('Sync_Opt_SyncData') ? document.getElementById('Sync_Opt_SyncData').checked : false;
  const createTarget = document.getElementById('Sync_Opt_CreateTarget') ? document.getElementById('Sync_Opt_CreateTarget').checked : true;

  if (!confirm(`Confirma a sincronização das alterações da base de Teste [${srcDb}] para a base de Produção [${tgtDb}]?\n\nEsta operação aplicará as migrações estruturais diretamente no banco de destino.`)) {
    return;
  }

  try {
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Sincronizando...';
    }
    showToast('Executando sincronização no banco de Produção...', 'info');

    const payload = {
      source_host: srcHost,
      source_port: srcPort,
      source_dbname: srcDb,
      source_user: srcUser,
      source_password: srcPass,
      target_host: tgtHost,
      target_port: tgtPort,
      target_dbname: tgtDb,
      target_user: srcUser,
      target_password: srcPass,
      tables: selectedTables,
      sync_data: syncData,
      create_target: createTarget
    };

    const res = await fetch('/api/config/db/sync/apply', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();
    if (!res.ok || !data.success) throw new Error(data.message || 'Falha ao executar sincronização');

    if (resCard) resCard.style.display = 'block';
    if (execLog) execLog.style.display = 'block';

    const logItems = data.executed_log || [];
    if (logItems.length === 0) {
      logContent.innerHTML = `<div style="color: var(--accent-emerald); font-weight: 600;"><i class="fa-solid fa-circle-check"></i> Tudo pronto! Nenhuma alteração DDL pendente no banco de destino.</div>`;
    } else {
      logContent.innerHTML = logItems.map(item => `
        <div style="display: flex; align-items: center; gap: 0.5rem; color: ${item.status === 'OK' ? 'var(--accent-emerald)' : 'var(--accent-rose)'};">
          <i class="fa-solid ${item.status === 'OK' ? 'fa-circle-check' : 'fa-circle-xmark'}"></i>
          <span>${escapeHtml(item.description)}</span>
        </div>
      `).join('');
    }

    if (syncData && data.data_sync_summary) {
      const summaryItems = Object.entries(data.data_sync_summary).map(([tbl, info]) => `
        <div style="color: var(--text-secondary); font-size: 0.78rem; margin-left: 1.5rem;">
          • Tabela <strong>${escapeHtml(tbl)}</strong>: ${info.status} (${info.rows || 0} registros sincronizados)
        </div>
      `).join('');
      logContent.innerHTML += `<div style="margin-top: 0.5rem; font-weight: 600; color: var(--accent-blue);"><i class="fa-solid fa-arrows-split-up-and-left"></i> Dados Sincronizados:</div>${summaryItems}`;
    }

    showToast(data.message || 'Sincronização aplicada com sucesso!', 'success');
  } catch (err) {
    showToast(`Erro ao sincronizar: ${err.message}`, 'error');
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<i class="fa-solid fa-bolt"></i> 2. Executar Sincronização';
    }
  }
}


// ==========================================
// NF-E VALIDATION & FISCAL CONFIG FUNCTIONS
// ==========================================

let currentValidationCodPed = null;
let currentValidationXml = '';

async function validarNfeOrder(codPed) {
  currentValidationCodPed = codPed;
  showToast(`Consultando WebService SEFAZ e auditando Pedido #${codPed}...`, 'info', 2500);

  try {
    const res = await fetch(`/api/pedidos/${codPed}/validar-nfe`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Falha ao validar NF-e na SEFAZ');

    renderValidationReport(codPed, data);
    document.getElementById('nfe-validation-modal').classList.add('active');
  } catch (err) {
    showToast(`Erro ao validar NF-e: ${err.message}`, 'error');
  }
}

function renderValidationReport(codPed, data) {
  const banner = document.getElementById('nfe-val-banner');
  const icon = document.getElementById('nfe-val-status-icon');
  const title = document.getElementById('nfe-val-title');
  const subtitle = document.getElementById('nfe-val-subtitle');

  const badgeOk = document.getElementById('nfe-val-badge-ok');
  const badgeWarn = document.getElementById('nfe-val-badge-warn');
  const badgeErr = document.getElementById('nfe-val-badge-err');

  const envLabel = document.getElementById('nfe-val-env-label');
  const wsLabel = document.getElementById('nfe-val-ws-label');
  const chaveLabel = document.getElementById('nfe-val-chave-label');
  const xmlContent = document.getElementById('nfe-val-xml-content');
  const xmlBox = document.getElementById('nfe-val-xml-box');

  const checklistContainer = document.getElementById('nfe-val-checklist-container');
  const alertsBox = document.getElementById('nfe-val-alerts-box');
  const alertsList = document.getElementById('nfe-val-alerts-list');

  currentValidationXml = data.xml_preview || '';
  if (xmlContent) xmlContent.innerText = currentValidationXml;
  if (xmlBox) xmlBox.style.display = 'none'; // closed by default

  // Populate Environment & WebService Diagnostics
  if (envLabel) {
    const ambCode = data.ambiente_codigo || '2';
    const isProd = ambCode === '1';
    envLabel.innerHTML = `<i class="fa-solid fa-server"></i> <span style="color: ${isProd ? 'var(--accent-emerald)' : 'var(--accent-blue)'};">${escapeHtml(data.ambiente_nome || (isProd ? '1 - Produção Oficial SEFAZ' : '2 - Homologação / Simulação'))}</span>`;
  }

  if (wsLabel) {
    const ws = data.webservice || {};
    if (ws.online) {
      wsLabel.innerHTML = `<i class="fa-solid fa-circle-check" style="color: var(--accent-emerald);"></i> <span style="color: var(--accent-emerald);">Online (${ws.elapsed_ms || 0}ms - SEFAZ ${escapeHtml(data.uf_sefaz || 'SP')})</span>`;
    } else {
      wsLabel.innerHTML = `<i class="fa-solid fa-circle-exclamation" style="color: var(--accent-amber);"></i> <span style="color: var(--accent-amber);">${escapeHtml(ws.mensagem || 'WebService Indisponível')}</span>`;
    }
  }

  if (chaveLabel) {
    chaveLabel.innerText = data.chave_nfe || 'Chave gerada automaticamente na transmissão';
  }

  if (badgeOk) badgeOk.innerText = `${data.total_checks || 0} Regras Auditadas`;
  if (badgeWarn) badgeWarn.innerText = `${data.total_warnings || 0} Alertas`;
  if (badgeErr) badgeErr.innerText = `${data.total_errors || 0} Erros`;

  const btnEmit = document.getElementById('btn-nfe-val-emit-now');
  const btnOpenFrm = document.getElementById('btn-nfe-val-open-frmnota');

  if (data.valid) {
    if (banner) {
      banner.style.background = 'rgba(16, 185, 129, 0.1)';
      banner.style.borderColor = 'rgba(16, 185, 129, 0.3)';
    }
    if (icon) {
      icon.className = 'fa-solid fa-circle-check fa-2x';
      icon.style.color = 'var(--accent-emerald)';
    }
    if (title) title.innerText = 'NF-e Totalmente Aprovada e Válida!';
    if (subtitle) subtitle.innerText = `Pedido #${codPed} - Todos os parâmetros SEFAZ 4.00, dados do destinatário e comunicação validados com sucesso no ambiente ${data.ambiente_codigo == '1' ? 'de Produção' : 'de Homologação'}.`;
    if (badgeErr) badgeErr.style.display = 'none';

    if (btnEmit) {
      btnEmit.disabled = false;
      btnEmit.style.opacity = '1';
      btnEmit.style.cursor = 'pointer';
      btnEmit.title = 'Transmitir NF-e autorizada';
    }
    if (btnOpenFrm) {
      btnOpenFrm.innerHTML = '<i class="fa-solid fa-file-invoice"></i> Abrir no FrmNota';
    }
  } else {
    if (banner) {
      banner.style.background = 'rgba(244, 63, 94, 0.1)';
      banner.style.borderColor = 'rgba(244, 63, 94, 0.3)';
    }
    if (icon) {
      icon.className = 'fa-solid fa-circle-xmark fa-2x';
      icon.style.color = 'var(--accent-rose)';
    }
    if (title) title.innerText = 'Impedimentos Fiscais Identificados';
    if (subtitle) subtitle.innerText = `Pedido #${codPed} - A NF-e (Modelo 55) não pode ser emitida sem os dados completos do Destinatário (CPF/CNPJ, Nome e Endereço).`;
    if (badgeErr) badgeErr.style.display = 'inline-block';

    if (btnEmit) {
      btnEmit.disabled = true;
      btnEmit.style.opacity = '0.5';
      btnEmit.style.cursor = 'not-allowed';
      btnEmit.title = 'Corrija as pendências do destinatário antes de emitir a NF-e';
    }
    if (btnOpenFrm) {
      btnOpenFrm.innerHTML = '<i class="fa-solid fa-user-pen"></i> Preencher Destinatário no FrmNota';
    }
  }

  // Checklist items
  if (data.checks && data.checks.length) {
    checklistContainer.innerHTML = data.checks.map(c => {
      let iconHtml = '<i class="fa-solid fa-check" style="color: var(--accent-emerald);"></i>';
      let bgStyle = '';
      if (c.status === 'ERROR') {
        iconHtml = '<i class="fa-solid fa-xmark" style="color: var(--accent-rose);"></i>';
        bgStyle = 'background: rgba(244, 63, 94, 0.05);';
      } else if (c.status === 'WARN') {
        iconHtml = '<i class="fa-solid fa-triangle-exclamation" style="color: var(--accent-amber);"></i>';
        bgStyle = 'background: rgba(245, 158, 11, 0.05);';
      }

      return `
        <div style="display: flex; align-items: center; justify-content: space-between; padding: 0.5rem 0.75rem; border-bottom: 1px solid var(--border-color); font-size: 0.85rem; ${bgStyle}">
          <div style="display: flex; align-items: center; gap: 0.6rem;">
            ${iconHtml}
            <div>
              <strong style="color: var(--text-primary);">${escapeHtml(c.categoria)} - ${escapeHtml(c.campo)}</strong>
              <div style="color: var(--text-muted); font-size: 0.78rem;">${escapeHtml(c.detalhe)}</div>
            </div>
          </div>
          <span class="badge ${c.status === 'OK' ? 'badge-active' : (c.status === 'WARN' ? 'badge-warning' : 'badge-danger')}">${c.status}</span>
        </div>
      `;
    }).join('');
  } else {
    checklistContainer.innerHTML = '<div style="text-align: center; color: var(--text-muted); padding: 1rem;">Sem itens no checklist</div>';
  }

  // Errors & Warnings list
  const allAlerts = [...(data.errors || []), ...(data.warnings || [])];
  if (allAlerts.length) {
    if (alertsBox) alertsBox.style.display = 'block';
    if (alertsList) alertsList.innerHTML = allAlerts.map(a => `<li style="margin-bottom: 0.25rem;">${escapeHtml(a)}</li>`).join('');
  } else {
    if (alertsBox) alertsBox.style.display = 'none';
  }
}

function toggleNfeXmlPreview() {
  const box = document.getElementById('nfe-val-xml-box');
  if (box) {
    box.style.display = box.style.display === 'none' ? 'block' : 'none';
  }
}

function copyNfeXmlPreview() {
  if (!currentValidationXml) {
    showToast('Nenhum XML disponível para cópia.', 'warning');
    return;
  }
  navigator.clipboard.writeText(currentValidationXml).then(() => {
    showToast('XML copiado para a área de transferência!', 'success');
  }).catch(() => {
    showToast('Não foi possível copiar o XML automaticamente.', 'warning');
  });
}

function closeNfeValidationModal() {
  const modal = document.getElementById('nfe-validation-modal');
  if (modal) modal.classList.remove('active');
}

function openFrmNotaFromValidationModal() {
  if (!currentValidationCodPed) return;
  const cod = currentValidationCodPed;
  closeNfeValidationModal();
  openNfeEmissaoModal(cod);
}

async function emitNfeFromValidationModal() {
  if (!currentValidationCodPed) return;
  const btnEmit = document.getElementById('btn-nfe-val-emit-now');
  if (btnEmit && btnEmit.disabled) {
    showToast('Existem dados pendentes no Destinatário. Abra o FrmNota para preencher os dados do cliente.', 'warning', 4000);
    return;
  }
  const cod = currentValidationCodPed;
  closeNfeValidationModal();
  emitNfeFromOrder(cod);
}

// Database Config Page Tabs
function switchDbConfigTab(tabId, btnEl) {
  const tabs = document.querySelectorAll('.db-config-tab-content');
  tabs.forEach(t => t.style.display = 'none');

  const target = document.getElementById(tabId);
  if (target) target.style.display = 'block';

  if (btnEl && btnEl.parentElement) {
    btnEl.parentElement.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    btnEl.classList.add('active');
  }
}

async function loadPdvConfigAll() {
  try {
    const res = await fetch('/api/pdv/config');
    const cfg = await res.json();
    if (!res.ok) return;

    // NF-e
    if (cfg.Nfe) {
      if (document.getElementById('Cfg_Page_Nfe_Serie')) document.getElementById('Cfg_Page_Nfe_Serie').value = cfg.Nfe.SerieNfe || '1';
      if (document.getElementById('Cfg_Page_Nfe_Numero')) document.getElementById('Cfg_Page_Nfe_Numero').value = cfg.Nfe.NroNfe || '1';
      if (document.getElementById('Cfg_Page_Nfe_Ambiente')) document.getElementById('Cfg_Page_Nfe_Ambiente').value = cfg.Nfe.Ambiente || '2';
      if (document.getElementById('Cfg_Page_Nfe_Chave')) document.getElementById('Cfg_Page_Nfe_Chave').value = cfg.Nfe.Chave || '';

      if (document.getElementById('Cfg_Main_Nfe_Serie')) document.getElementById('Cfg_Main_Nfe_Serie').value = cfg.Nfe.SerieNfe || '1';
      if (document.getElementById('Cfg_Main_Nfe_Numero')) document.getElementById('Cfg_Main_Nfe_Numero').value = cfg.Nfe.NroNfe || '1';
      if (document.getElementById('Cfg_Main_Nfe_Ambiente')) document.getElementById('Cfg_Main_Nfe_Ambiente').value = cfg.Nfe.Ambiente || '2';

      if (document.getElementById('Cfg_Card_Nfe_Serie')) document.getElementById('Cfg_Card_Nfe_Serie').value = cfg.Nfe.SerieNfe || '1';
      if (document.getElementById('Cfg_Card_Nfe_Numero')) document.getElementById('Cfg_Card_Nfe_Numero').value = cfg.CardNfe ? (cfg.Nfe.NroNfe || '1') : (cfg.Nfe.NroNfe || '1');
    }

    // NFC-e
    if (cfg.Nfce) {
      if (document.getElementById('Cfg_Page_Nfce_Serie')) document.getElementById('Cfg_Page_Nfce_Serie').value = cfg.Nfce.SerieNfce || '1';
      if (document.getElementById('Cfg_Page_Nfce_Numero')) document.getElementById('Cfg_Page_Nfce_Numero').value = cfg.Nfce.NroNfce || '16';
      if (document.getElementById('Cfg_Page_Nfce_Ambiente')) document.getElementById('Cfg_Page_Nfce_Ambiente').value = cfg.Nfce.Ambiente || '2';
      if (document.getElementById('Cfg_Page_Nfce_CodigoID')) document.getElementById('Cfg_Page_Nfce_CodigoID').value = cfg.Nfce.CodigoID || '000001';
      if (document.getElementById('Cfg_Page_Nfce_CodigoSeg')) document.getElementById('Cfg_Page_Nfce_CodigoSeg').value = cfg.Nfce.CodigoSeg || '';
      if (document.getElementById('Cfg_Page_Nfce_Chave')) document.getElementById('Cfg_Page_Nfce_Chave').value = cfg.Nfce.Chave || '';
    }

    // Certificado
    if (cfg.Certificado) {
      if (document.getElementById('Cfg_Page_Cert_Tipo')) document.getElementById('Cfg_Page_Cert_Tipo').value = cfg.Certificado.Tipo || 'A1';
      if (document.getElementById('Cfg_Page_Cert_Caminho')) document.getElementById('Cfg_Page_Cert_Caminho').value = cfg.Certificado.Caminho || '';
      if (document.getElementById('Cfg_Page_Cert_Senha')) document.getElementById('Cfg_Page_Cert_Senha').value = cfg.Certificado.Senha || '';
      if (document.getElementById('Cfg_Page_Cert_Validade')) document.getElementById('Cfg_Page_Cert_Validade').value = cfg.Certificado.Validade || '';

      if (document.getElementById('Cfg_Main_Cert_Tipo')) document.getElementById('Cfg_Main_Cert_Tipo').value = cfg.Certificado.Tipo || 'A1';
      if (document.getElementById('Cfg_Main_Cert_Caminho')) document.getElementById('Cfg_Main_Cert_Caminho').value = cfg.Certificado.Caminho || '';
      if (document.getElementById('Cfg_Main_Cert_Senha')) document.getElementById('Cfg_Main_Cert_Senha').value = cfg.Certificado.Senha || '';
    }
  } catch (e) {
    console.error('Erro ao carregar configurações do PDV/NFe:', e);
  }
}

async function saveNfeAndCertMain(e) {
  if (e && e.preventDefault) e.preventDefault();
  const getVal = (id, fallback = '') => {
    const el = document.getElementById(id);
    return el ? el.value : fallback;
  };

  const payload = {
    Nfe: {
      SerieNfe: getVal('Cfg_Main_Nfe_Serie', '1'),
      NroNfe: getVal('Cfg_Main_Nfe_Numero', '1'),
      Ambiente: getVal('Cfg_Main_Nfe_Ambiente', '2')
    },
    Certificado: {
      Tipo: getVal('Cfg_Main_Cert_Tipo', 'A1'),
      Caminho: getVal('Cfg_Main_Cert_Caminho', 'C:\\Certificados\\certificado.pfx'),
      Senha: getVal('Cfg_Main_Cert_Senha', '')
    }
  };
  await postPdvConfigSave(payload, 'Parâmetros da NF-e (Modelo 55) & Certificado Digital salvos com sucesso!');
}

async function testDigitalCertMain() {
  const caminho = document.getElementById('Cfg_Main_Cert_Caminho') ? document.getElementById('Cfg_Main_Cert_Caminho').value : '';
  const senha = document.getElementById('Cfg_Main_Cert_Senha') ? document.getElementById('Cfg_Main_Cert_Senha').value : '';

  showToast('Testando Certificado Digital...', 'info');

  try {
    const res = await fetch('/api/config/certificado/test', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ caminho, senha })
    });
    const result = await res.json();

    const titleEl = document.getElementById('cert-main-status-title');
    const descEl = document.getElementById('cert-main-status-desc');

    if (result.success) {
      showToast(result.message, 'success');
      if (titleEl) titleEl.innerText = `Certificado Digital Válido (${result.tipo})`;
      if (descEl) descEl.innerText = result.message;
    } else {
      showToast(result.message, 'warning');
      if (titleEl) titleEl.innerText = 'Certificado Digital Não Encontrado';
      if (descEl) descEl.innerText = result.message;
    }
  } catch (err) {
    showToast(`Erro no teste do certificado: ${err.message}`, 'error');
  }
}

function openCertificadoModal() {
  const modal = document.getElementById('certificado-modal');
  if (!modal) return;

  const caminho = document.getElementById('Cfg_Main_Cert_Caminho') ? document.getElementById('Cfg_Main_Cert_Caminho').value : '';
  const senha = document.getElementById('Cfg_Main_Cert_Senha') ? document.getElementById('Cfg_Main_Cert_Senha').value : '';

  if (document.getElementById('Cfg_Modal_Cert_Caminho') && caminho) document.getElementById('Cfg_Modal_Cert_Caminho').value = caminho;
  if (document.getElementById('Cfg_Modal_Cert_Senha') && senha) document.getElementById('Cfg_Modal_Cert_Senha').value = senha;

  modal.classList.add('active');
}

function closeCertificadoModal() {
  const modal = document.getElementById('certificado-modal');
  if (modal) modal.classList.remove('active');
}

async function saveCertModal(e) {
  e.preventDefault();
  const tipo = document.getElementById('Cfg_Modal_Cert_Tipo').value;
  const caminho = document.getElementById('Cfg_Modal_Cert_Caminho').value;
  const senha = document.getElementById('Cfg_Modal_Cert_Senha').value;
  const validade = document.getElementById('Cfg_Modal_Cert_Validade').value;

  if (document.getElementById('Cfg_Main_Cert_Tipo')) document.getElementById('Cfg_Main_Cert_Tipo').value = tipo;
  if (document.getElementById('Cfg_Main_Cert_Caminho')) document.getElementById('Cfg_Main_Cert_Caminho').value = caminho;
  if (document.getElementById('Cfg_Main_Cert_Senha')) document.getElementById('Cfg_Main_Cert_Senha').value = senha;

  const payload = {
    Certificado: {
      Tipo: tipo,
      Caminho: caminho,
      Senha: senha,
      Validade: validade
    }
  };

  await postPdvConfigSave(payload, 'Certificado Digital configurado e salvo com sucesso!');
  closeCertificadoModal();
}

async function testCertModal() {
  const caminho = document.getElementById('Cfg_Modal_Cert_Caminho').value;
  const senha = document.getElementById('Cfg_Modal_Cert_Senha').value;

  showToast('Testando Certificado Digital...', 'info');

  try {
    const res = await fetch('/api/config/certificado/test', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ caminho, senha })
    });
    const result = await res.json();

    const titleEl = document.getElementById('cert-modal-status-title');
    const descEl = document.getElementById('cert-modal-status-desc');

    if (result.success) {
      showToast(result.message, 'success');
      if (titleEl) titleEl.innerText = `Certificado Válido (${result.tipo})`;
      if (descEl) descEl.innerText = result.message;
    } else {
      showToast(result.message, 'warning');
      if (titleEl) titleEl.innerText = 'Atenção no Certificado';
      if (descEl) descEl.innerText = result.message;
    }
  } catch (err) {
    showToast(`Erro no teste: ${err.message}`, 'error');
  }
}

async function saveNfeConfigPage(e) {
  if (e && e.preventDefault) e.preventDefault();
  const getVal = (id, fallback = '') => {
    const el = document.getElementById(id);
    return el ? el.value : fallback;
  };
  const payload = {
    Nfe: {
      SerieNfe: getVal('Cfg_Page_Nfe_Serie', '1'),
      NroNfe: getVal('Cfg_Page_Nfe_Numero', '1'),
      Ambiente: getVal('Cfg_Page_Nfe_Ambiente', '2'),
      Chave: getVal('Cfg_Page_Nfe_Chave', '')
    }
  };
  await postPdvConfigSave(payload, 'Parâmetros de NF-e (Modelo 55) salvos com sucesso!');
}

async function saveNfceConfigPage(e) {
  if (e && e.preventDefault) e.preventDefault();
  const getVal = (id, fallback = '') => {
    const el = document.getElementById(id);
    return el ? el.value : fallback;
  };
  const payload = {
    Nfce: {
      SerieNfce: getVal('Cfg_Page_Nfce_Serie', '1'),
      NroNfce: getVal('Cfg_Page_Nfce_Numero', '16'),
      Ambiente: getVal('Cfg_Page_Nfce_Ambiente', '2'),
      CodigoID: getVal('Cfg_Page_Nfce_CodigoID', '000001'),
      CodigoSeg: getVal('Cfg_Page_Nfce_CodigoSeg', ''),
      Chave: getVal('Cfg_Page_Nfce_Chave', '')
    }
  };
  await postPdvConfigSave(payload, 'Parâmetros de NFC-e (Modelo 65) salvos com sucesso!');
}

async function saveCertConfigPage(e) {
  if (e && e.preventDefault) e.preventDefault();
  const getVal = (id, fallback = '') => {
    const el = document.getElementById(id);
    return el ? el.value : fallback;
  };
  const payload = {
    Certificado: {
      Tipo: getVal('Cfg_Page_Cert_Tipo', 'A1'),
      Caminho: getVal('Cfg_Page_Cert_Caminho', 'C:\\Certificados\\certificado.pfx'),
      Senha: getVal('Cfg_Page_Cert_Senha', ''),
      Validade: getVal('Cfg_Page_Cert_Validade', '')
    }
  };
  await postPdvConfigSave(payload, 'Parâmetros do Certificado Digital salvos com sucesso!');
}

async function postPdvConfigSave(data, successMsg) {
  try {
    const payload = {
      ...data,
      id_empresa: typeof currentEmpresaId !== 'undefined' ? currentEmpresaId : 1
    };
    const res = await fetch('/api/pdv/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao salvar preferências');
    showToast(successMsg, 'success');
  } catch (err) {
    showToast(`Erro ao salvar: ${err.message}`, 'error');
  }
}

async function testDigitalCert() {
  const caminho = document.getElementById('Cfg_Page_Cert_Caminho') ? document.getElementById('Cfg_Page_Cert_Caminho').value : '';
  const senha = document.getElementById('Cfg_Page_Cert_Senha') ? document.getElementById('Cfg_Page_Cert_Senha').value : '';

  showToast('Testando Certificado Digital...', 'info');

  try {
    const res = await fetch('/api/config/certificado/test', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ caminho, senha })
    });
    const result = await res.json();

    const titleEl = document.getElementById('cert-status-title-text');
    const descEl = document.getElementById('cert-status-desc-text');

    if (result.success) {
      showToast(result.message, 'success');
      if (titleEl) titleEl.innerText = `Certificado Configurado (${result.tipo})`;
      if (descEl) descEl.innerText = result.message;
    } else {
      showToast(result.message, 'warning');
      if (titleEl) titleEl.innerText = 'Atenção no Certificado Digital';
      if (descEl) descEl.innerText = result.message;
    }
  } catch (err) {
    showToast(`Erro no teste do certificado: ${err.message}`, 'error');
  }
}

// Check session on page load
document.addEventListener('DOMContentLoaded', () => {
  checkAuth();
});

/* ==================== EMISSÃO DE NF-E MODELO 55 (FRMNOTA.FRM) ==================== */
let currentFrmNotaCodPed = null;

async function openNfeEmissaoModal(codPed, autoValidar = false) {
  currentFrmNotaCodPed = codPed;
  const modal = document.getElementById('nfe-emissao-modal');
  if (!modal) return;

  modal.classList.add('active');
  showToast(`Carregando parâmetros da NF-e para o Pedido #${codPed}...`, 'info');

  try {
    const res = await fetch(`/api/nfe/${codPed}/detalhes`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Erro ao carregar NF-e');

    const ped = data.pedido || {};
    const items = data.itens || [];
    const nfe = data.nfe || {};
    const company = data.empresa || {};
    const cfgNfe = data.config_nfe || {};

    // Populate Cabeçalho & Cliente
    setElVal('FrmNota_CodPed', ped.CodPed);
    setElVal('FrmNota_CodPed_Disp', `#${ped.CodPed}`);
    setElVal('FrmNota_NroNfe', ped.NroNfe || ped.CodPed || cfgNfe.NroNfe || '1');
    setElVal('FrmNota_SerieNfe', ped.SerieNfe || cfgNfe.SerieNfe || '1');
    setElVal('FrmNota_Cfo', ped.Cfo || '5102 - VENDA DE MERCADORIA ADQUIRIDA OU RECEBIDA DE TERCEIROS');
    setElVal('FrmNota_Emissao', ped.DataEmiss || new Date().toLocaleDateString('pt-BR'));
    setElVal('FrmNota_Saida', ped.DtSaida || ped.DataEmiss || new Date().toLocaleDateString('pt-BR'));
    setElVal('FrmNota_Hora', ped.Hora || new Date().toLocaleTimeString('pt-BR'));

    setElVal('FrmNota_CodEntidade', ped.Entidade || ped.CodEntidade || '');
    setElVal('FrmNota_NomeCliente', ped.NomeCliente || ped.Nome || '');
    const docCli = (ped.CPF || ped.CGC || '').trim();
    setElVal('FrmNota_CpfCnpj', docCli);
    setElVal('FrmNota_InscEst', ped.InscEst || ped.IE || '');
    setElVal('FrmNota_Endereco', ped.Endereco || ped.Logradouro || '');
    setElVal('FrmNota_Nro', ped.Nro || '');
    setElVal('FrmNota_Bairro', ped.Bairro || '');
    setElVal('FrmNota_Cidade', ped.Cidade || '');
    setElVal('FrmNota_Uf', ped.Uf || ped.UF || '');
    setElVal('FrmNota_Cep', ped.Cep || ped.CEP || '');

    // Populate Transporte
    setElVal('FrmNota_ModFrete', ped.ModFrete || '9');
    setElVal('FrmNota_NomeTransp', ped.NomeTransp || '');
    setElVal('FrmNota_CnpjsTransp', ped.CnpjsTransp || '');
    setElVal('FrmNota_IeTransp', ped.IeTransp || '');
    setElVal('FrmNota_CidadeTransp', ped.CidadeTransp || '');
    setElVal('FrmNota_UfTransp', ped.UfTransp || '');
    setElVal('FrmNota_PlacaVeic', ped.PlacaVeic || '');
    setElVal('FrmNota_UfVeic', ped.UfVeic || '');
    setElVal('FrmNota_QtdVol', ped.QtdVol || '1');
    setElVal('FrmNota_EspecieVol', ped.EspecieVol || 'VOLUMES');
    setElVal('FrmNota_MarcaVol', ped.MarcaVol || 'DIVERSAS');
    setElVal('FrmNota_PesoBruto', ped.PesoBruto || '0.000');
    setElVal('FrmNota_PesoLiquido', ped.PesoLiquido || '0.000');

    // Populate Totais
    const subtotal = floatOrZero(ped.SubTotal || ped.Total);
    const desconto = floatOrZero(ped.Desconto);
    const frete = floatOrZero(ped.ValorFrete);
    const totalNf = Math.max(0, subtotal - desconto + frete);

    setElVal('FrmNota_BaseIcms', (0.00).toFixed(2));
    setElVal('FrmNota_ValorIcms', (0.00).toFixed(2));
    setElVal('FrmNota_BaseSt', (0.00).toFixed(2));
    setElVal('FrmNota_ValorSt', (0.00).toFixed(2));
    setElVal('FrmNota_TotalProdutos', subtotal.toFixed(2));
    setElVal('FrmNota_ValorFrete', frete.toFixed(2));
    setElVal('FrmNota_Desconto', desconto.toFixed(2));
    setElVal('FrmNota_TotalNota', totalNf.toFixed(2));
    setElVal('FrmNota_ObsRodape', ped.Obs || `PEDIDO DE VENDA #${ped.CodPed} - EMITIDO VIA SISTEMA VENDAS SEFAZ NF-E MODELO 55`);
    setElVal('FrmNota_ObsCorpo', ped.ObsCorpo || '');

    // Status Banner
    const statusText = document.getElementById('frmnota-status-text');
    const statusIcon = document.getElementById('frmnota-status-icon');
    const chaveBadge = document.getElementById('frmnota-chave-badge');

    if (nfe && nfe.NroChave) {
      if (statusText) statusText.innerText = `Status: NF-e Transmitida & Autorizada SEFAZ! (Prot: ${nfe.Protocolo || '100'})`;
      if (statusIcon) { statusIcon.className = 'fa-solid fa-circle-check'; statusIcon.style.color = 'var(--accent-emerald)'; }
      if (chaveBadge) chaveBadge.innerText = `Chave: ${nfe.NroChave}`;
    } else {
      if (statusText) statusText.innerText = 'Status: Digitação da NF-e (Aguardando Emissão Direta SEFAZ)';
      if (statusIcon) { statusIcon.className = 'fa-solid fa-circle-info'; statusIcon.style.color = 'var(--accent-blue)'; }
      if (chaveBadge) chaveBadge.innerText = 'Chave: Gerada na Emissão';
    }

    // Populate Items Table
    renderFrmNotaItems(items);

    if (autoValidar) {
      setTimeout(() => {
        validarNfeFrmNota();
      }, 250);
    }
  } catch (err) {
    showToast(`Erro ao carregar dados da nota: ${err.message}`, 'error');
  }
}

function renderFrmNotaItems(items) {
  const tbody = document.getElementById('frmnota-items-tbody');
  const countBadge = document.getElementById('frmnota-total-items-badge');
  if (countBadge) countBadge.innerText = `${items ? items.length : 0} Itens`;
  if (!tbody) return;

  if (!items || !items.length) {
    tbody.innerHTML = `<tr><td colspan="9" class="loading-td">Nenhum item anexado ao pedido.</td></tr>`;
    return;
  }

  tbody.innerHTML = items.map((it, idx) => {
    const q = floatOrZero(it.Qtd || 1);
    const v = floatOrZero(it.ValorUnit || 0);
    const sub = q * v;
    return `
      <tr>
        <td><strong>#${idx + 1}</strong></td>
        <td>${it.CodPrd || it.Produto || ''}</td>
        <td><strong>${escapeHtml(it.Descricao_Produto || 'PRODUTO')}</strong></td>
        <td>${it.NCM || '62034200'}</td>
        <td>${it.CFOP || '5102'}</td>
        <td>${it.Embalagem || 'UN'}</td>
        <td style="text-align: right;">${q.toFixed(2)}</td>
        <td style="text-align: right;">R$ ${v.toFixed(2)}</td>
        <td style="text-align: right; font-weight: 700; color: var(--accent-emerald);">R$ ${sub.toFixed(2)}</td>
      </tr>
    `;
  }).join('');
}

function closeNfeEmissaoModal() {
  const modal = document.getElementById('nfe-emissao-modal');
  if (modal) modal.classList.remove('active');
}

function switchFrmNotaTab(tabId, btnEl) {
  const tabs = document.querySelectorAll('.frmnota-tab-content');
  tabs.forEach(t => t.style.display = 'none');

  const target = document.getElementById(tabId);
  if (target) target.style.display = 'block';

  if (btnEl && btnEl.parentElement) {
    btnEl.parentElement.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    btnEl.classList.add('active');
  }
}

async function validarNfeFrmNota() {
  if (!currentFrmNotaCodPed) return;
  showToast('Validando estrutura e regras da NF-e SEFAZ 4.00...', 'info');

  try {
    const res = await fetch(`/api/pedidos/${currentFrmNotaCodPed}/validar-nfe`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Falha ao validar NF-e');

    renderFrmNotaValOverlay(currentFrmNotaCodPed, data);
    const overlay = document.getElementById('frmnota-val-overlay');
    if (overlay) overlay.classList.add('active');
  } catch (err) {
    showToast(`Erro ao validar NF-e: ${err.message}`, 'error');
  }
}

function closeFrmNotaValOverlay() {
  const overlay = document.getElementById('frmnota-val-overlay');
  if (overlay) overlay.classList.remove('active');
}

function renderFrmNotaValOverlay(codPed, data) {
  const banner = document.getElementById('frmnota-val-banner');
  const icon = document.getElementById('frmnota-val-icon');
  const title = document.getElementById('frmnota-val-title');
  const subtitle = document.getElementById('frmnota-val-subtitle');

  const badgeOk = document.getElementById('frmnota-val-badge-ok');
  const badgeWarn = document.getElementById('frmnota-val-badge-warn');
  const badgeErr = document.getElementById('frmnota-val-badge-err');

  const checklistContainer = document.getElementById('frmnota-val-checklist-container');
  const alertsBox = document.getElementById('frmnota-val-alerts-box');
  const alertsList = document.getElementById('frmnota-val-alerts-list');

  if (badgeOk) badgeOk.innerText = `${data.total_checks || 0} Regras Auditadas`;
  if (badgeWarn) badgeWarn.innerText = `${data.total_warnings || 0} Alertas`;
  if (badgeErr) badgeErr.innerText = `${data.total_errors || 0} Erros`;

  if (data.valid) {
    if (banner) {
      banner.style.background = 'rgba(16, 185, 129, 0.1)';
      banner.style.borderColor = 'rgba(16, 185, 129, 0.3)';
    }
    if (icon) {
      icon.className = 'fa-solid fa-circle-check fa-2x';
      icon.style.color = 'var(--accent-emerald)';
    }
    if (title) title.innerText = 'NF-e Totalmente Válida para Transmissão!';
    if (subtitle) subtitle.innerText = `Pedido #${codPed} - Todos os parâmetros SEFAZ v4.00 foram aprovados com sucesso.`;
    if (badgeErr) badgeErr.style.display = 'none';
  } else {
    if (banner) {
      banner.style.background = 'rgba(244, 63, 94, 0.1)';
      banner.style.borderColor = 'rgba(244, 63, 94, 0.3)';
    }
    if (icon) {
      icon.className = 'fa-solid fa-circle-xmark fa-2x';
      icon.style.color = 'var(--accent-rose)';
    }
    if (title) title.innerText = 'Impedimentos de Emissão Identificados';
    if (subtitle) subtitle.innerText = `Pedido #${codPed} - Corrija os erros destacados abaixo antes de emitir a NF-e.`;
    if (badgeErr) badgeErr.style.display = 'inline-block';
  }

  // Checklist items
  if (data.checks && data.checks.length) {
    checklistContainer.innerHTML = data.checks.map(c => {
      let iconHtml = '<i class="fa-solid fa-check" style="color: var(--accent-emerald);"></i>';
      let bgStyle = '';
      if (c.status === 'ERROR') {
        iconHtml = '<i class="fa-solid fa-xmark" style="color: var(--accent-rose);"></i>';
        bgStyle = 'background: rgba(244, 63, 94, 0.05);';
      } else if (c.status === 'WARN') {
        iconHtml = '<i class="fa-solid fa-triangle-exclamation" style="color: var(--accent-amber);"></i>';
        bgStyle = 'background: rgba(245, 158, 11, 0.05);';
      }

      return `
        <div style="display: flex; align-items: center; justify-content: space-between; padding: 0.5rem 0.75rem; border-bottom: 1px solid var(--border-color); font-size: 0.85rem; ${bgStyle}">
          <div style="display: flex; align-items: center; gap: 0.6rem;">
            ${iconHtml}
            <div>
              <strong style="color: var(--text-primary);">${escapeHtml(c.categoria)} - ${escapeHtml(c.campo)}</strong>
              <div style="color: var(--text-muted); font-size: 0.78rem;">${escapeHtml(c.detalhe)}</div>
            </div>
          </div>
          <span class="badge ${c.status === 'OK' ? 'badge-active' : (c.status === 'WARN' ? 'badge-warning' : 'badge-danger')}">${c.status}</span>
        </div>
      `;
    }).join('');
  } else {
    checklistContainer.innerHTML = '<div style="text-align: center; color: var(--text-muted); padding: 1rem;">Sem itens no checklist</div>';
  }

  // Errors & Warnings list
  const allAlerts = [...(data.errors || []), ...(data.warnings || [])];
  if (allAlerts.length) {
    if (alertsBox) alertsBox.style.display = 'block';
    if (alertsList) alertsList.innerHTML = allAlerts.map(a => `<li style="margin-bottom: 0.25rem;">${escapeHtml(a)}</li>`).join('');
  } else {
    if (alertsBox) alertsBox.style.display = 'none';
  }
}

async function transmitirNfeFrmNota() {
  if (!currentFrmNotaCodPed) return;
  showToast('Conectando ao WebService SEFAZ 4.00 e transmitindo NF-e...', 'info');

  const payload = getFrmNotaFormData();

  try {
    const res = await fetch(`/api/nfe/${currentFrmNotaCodPed}/emitir-sefaz-direto`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const result = await res.json();
    if (!res.ok || !result.success) throw new Error(result.error || result.message || 'Erro na transmissão SEFAZ');

    showToast(result.message, 'success');
    openNfeEmissaoModal(currentFrmNotaCodPed);
    if (typeof fetchOrders === 'function') fetchOrders();
  } catch (err) {
    showToast(`Falha na transmissão SEFAZ: ${err.message}`, 'error');
  }
}

function getFrmNotaFormData() {
  return {
    NomeCliente: document.getElementById('FrmNota_NomeCliente') ? document.getElementById('FrmNota_NomeCliente').value : '',
    CpfCnpj: document.getElementById('FrmNota_CpfCnpj') ? document.getElementById('FrmNota_CpfCnpj').value : '',
    InscEst: document.getElementById('FrmNota_InscEst') ? document.getElementById('FrmNota_InscEst').value : '',
    Endereco: document.getElementById('FrmNota_Endereco') ? document.getElementById('FrmNota_Endereco').value : '',
    Nro: document.getElementById('FrmNota_Nro') ? document.getElementById('FrmNota_Nro').value : '',
    Bairro: document.getElementById('FrmNota_Bairro') ? document.getElementById('FrmNota_Bairro').value : '',
    Cidade: document.getElementById('FrmNota_Cidade') ? document.getElementById('FrmNota_Cidade').value : '',
    Uf: document.getElementById('FrmNota_Uf') ? document.getElementById('FrmNota_Uf').value : '',
    Cep: document.getElementById('FrmNota_Cep') ? document.getElementById('FrmNota_Cep').value : '',
    NroNfe: document.getElementById('FrmNota_NroNfe') ? document.getElementById('FrmNota_NroNfe').value : '',
    SerieNfe: document.getElementById('FrmNota_SerieNfe') ? document.getElementById('FrmNota_SerieNfe').value : '1',
    Cfo: document.getElementById('FrmNota_Cfo') ? document.getElementById('FrmNota_Cfo').value : '',
    Emissao: document.getElementById('FrmNota_Emissao') ? document.getElementById('FrmNota_Emissao').value : '',
    Saida: document.getElementById('FrmNota_Saida') ? document.getElementById('FrmNota_Saida').value : '',
    ModFrete: document.getElementById('FrmNota_ModFrete') ? document.getElementById('FrmNota_ModFrete').value : '9',
    NomeTransp: document.getElementById('FrmNota_NomeTransp') ? document.getElementById('FrmNota_NomeTransp').value : '',
    CnpjsTransp: document.getElementById('FrmNota_CnpjsTransp') ? document.getElementById('FrmNota_CnpjsTransp').value : '',
    IeTransp: document.getElementById('FrmNota_IeTransp') ? document.getElementById('FrmNota_IeTransp').value : '',
    CidadeTransp: document.getElementById('FrmNota_CidadeTransp') ? document.getElementById('FrmNota_CidadeTransp').value : '',
    UfTransp: document.getElementById('FrmNota_UfTransp') ? document.getElementById('FrmNota_UfTransp').value : '',
    PlacaVeic: document.getElementById('FrmNota_PlacaVeic') ? document.getElementById('FrmNota_PlacaVeic').value : '',
    UfVeic: document.getElementById('FrmNota_UfVeic') ? document.getElementById('FrmNota_UfVeic').value : '',
    QtdVol: document.getElementById('FrmNota_QtdVol') ? document.getElementById('FrmNota_QtdVol').value : '1',
    EspecieVol: document.getElementById('FrmNota_EspecieVol') ? document.getElementById('FrmNota_EspecieVol').value : '',
    MarcaVol: document.getElementById('FrmNota_MarcaVol') ? document.getElementById('FrmNota_MarcaVol').value : '',
    PesoBruto: document.getElementById('FrmNota_PesoBruto') ? document.getElementById('FrmNota_PesoBruto').value : '0.000',
    PesoLiquido: document.getElementById('FrmNota_PesoLiquido') ? document.getElementById('FrmNota_PesoLiquido').value : '0.000',
    ValorFrete: document.getElementById('FrmNota_ValorFrete') ? document.getElementById('FrmNota_ValorFrete').value : '0.00',
    Desconto: document.getElementById('FrmNota_Desconto') ? document.getElementById('FrmNota_Desconto').value : '0.00',
    ObsRodape: document.getElementById('FrmNota_ObsRodape') ? document.getElementById('FrmNota_ObsRodape').value : '',
    ObsCorpo: document.getElementById('FrmNota_ObsCorpo') ? document.getElementById('FrmNota_ObsCorpo').value : ''
  };
}

async function salvarEDeclararFrmNota(e) {
  if (e && e.preventDefault) e.preventDefault();
  if (!currentFrmNotaCodPed) return;

  const payload = getFrmNotaFormData();
  try {
    const res = await fetch(`/api/nfe/${currentFrmNotaCodPed}/salvar-detalhes`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Erro ao salvar NF-e');

    showToast('Parâmetros da Nota Fiscal salvos com sucesso!', 'success');
  } catch (err) {
    showToast(`Erro ao salvar: ${err.message}`, 'error');
  }
}

function imprimirDanfeOrder(codPed) {
  if (!codPed) return;
  window.open(`/api/nfe/${codPed}/danfe-html`, '_blank');
}

function imprimirDanfeFrmNota() {
  if (!currentFrmNotaCodPed) return;
  window.open(`/api/nfe/${currentFrmNotaCodPed}/danfe-html`, '_blank');
}

function baixarXmlFrmNota() {
  if (!currentFrmNotaCodPed) return;
  window.location.href = `/api/nfe/${currentFrmNotaCodPed}/xml-download`;
}

async function cancelarNfeFrmNota() {
  if (!currentFrmNotaCodPed) return;
  const motivo = prompt('Digite o motivo/justificativa para o Cancelamento da NF-e na SEFAZ (Mínimo 15 caracteres):', 'Cancelamento solicitado pelo emitente');
  if (!motivo) return;
  if (motivo.length < 15) {
    showToast('O motivo do cancelamento deve possuir pelo menos 15 caracteres.', 'warning');
    return;
  }

  showToast('Enviando evento de Cancelamento para a SEFAZ...', 'info');
  try {
    const res = await fetch(`/api/nfe/${currentFrmNotaCodPed}/cancelar-sefaz`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ justificativa: motivo })
    });
    const result = await res.json();
    if (!res.ok || !result.success) throw new Error(result.error || 'Erro ao cancelar');

    showToast(result.message, 'success');
    openNfeEmissaoModal(currentFrmNotaCodPed);
    if (typeof fetchOrders === 'function') fetchOrders();
  } catch (err) {
    showToast(`Erro ao cancelar NF-e: ${err.message}`, 'error');
  }
}

async function cceNfeFrmNota() {
  if (!currentFrmNotaCodPed) return;
  const correcao = prompt('Digite a Carta de Correção Eletrônica (CC-e) para enviar à SEFAZ (Mínimo 15 caracteres):');
  if (!correcao) return;

  showToast('Enviando Carta de Correção Eletrônica (CC-e) para a SEFAZ...', 'info');
  try {
    const res = await fetch(`/api/nfe/${currentFrmNotaCodPed}/cce-sefaz`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ correcao })
    });
    const result = await res.json();
    if (!res.ok || !result.success) throw new Error(result.error || 'Erro ao enviar CC-e');

    showToast(result.message, 'success');
  } catch (err) {
    showToast(`Erro no envio da CC-e: ${err.message}`, 'error');
  }
}

/* --- CNPJ CONSULTA & AUTOPREENCHIMENTO --- */
async function consultarCNPJ(inputId) {
  const inputEl = document.getElementById(inputId);
  if (!inputEl) return;

  const rawVal = inputEl.value || '';
  const cleanCnpj = rawVal.replace(/\D/g, '');

  if (cleanCnpj.length !== 14) {
    showToast('Informe um CNPJ válido com 14 dígitos numéricos para realizar a busca.', 'warning');
    inputEl.focus();
    return;
  }

  const container = inputEl.parentElement;
  const btnEl = container ? container.querySelector('button') : null;
  const originalBtnHtml = btnEl ? btnEl.innerHTML : '';

  if (btnEl) {
    btnEl.disabled = true;
    btnEl.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Buscando...';
  }

  try {
    let data = null;

    try {
      const resp = await fetch(`/api/cnpj/${cleanCnpj}`);
      if (resp.ok) {
        data = await resp.json();
      }
    } catch (e) {
      console.warn('Falha no proxy backend de CNPJ, buscando via cliente:', e);
    }

    if (!data || !data.success) {
      try {
        const resp = await fetch(`https://publica.cnpj.ws/cnpj/${cleanCnpj}`);
        if (resp.ok) {
          const cWs = await resp.json();
          const est = cWs.estabelecimento || {};
          const ies = est.inscricoes_estaduais || [];
          const activeIes = ies.filter(i => i.ativo);
          const ieVal = activeIes.length ? activeIes[0].inscricao_estadual : (ies.length ? ies[0].inscricao_estadual : '');

          const tipoLog = est.tipo_logradouro || '';
          const log = est.logradouro || '';
          const fullLog = `${tipoLog} ${log}`.trim();
          const ddd = est.ddd1 || '';
          const tel = est.telefone1 || '';
          const foneStr = (ddd || tel) ? `(${ddd}) ${tel}`.trim() : '';

          data = {
            success: true,
            cnpj: cleanCnpj,
            razao_social: cWs.razao_social || '',
            nome_fantasia: est.nome_fantasia || cWs.razao_social || '',
            inscricao_estadual: ieVal || '',
            logradouro: fullLog,
            numero: est.numero || '',
            complemento: est.complemento || '',
            bairro: est.bairro || '',
            cidade: est.cidade ? est.cidade.nome : '',
            codigo_ibge: est.cidade ? String(est.cidade.ibge_id || '') : '',
            uf: est.estado ? est.estado.sigla : '',
            cep: est.cep || '',
            telefone: foneStr,
            email: est.email || ''
          };
        }
      } catch (e2) {
        console.warn('Falha no CNPJ.ws direto do frontend:', e2);
      }
    }

    if (!data || !data.success) {
      try {
        const resp = await fetch(`https://brasilapi.com.br/api/cnpj/v1/${cleanCnpj}`);
        if (resp.ok) {
          const bApi = await resp.json();
          data = {
            success: true,
            cnpj: bApi.cnpj || cleanCnpj,
            razao_social: bApi.razao_social || '',
            nome_fantasia: bApi.nome_fantasia || bApi.razao_social || '',
            inscricao_estadual: '',
            logradouro: `${bApi.descricao_tipo_de_logradouro || ''} ${bApi.logradouro || ''}`.trim(),
            numero: bApi.numero || '',
            complemento: bApi.complemento || '',
            bairro: bApi.bairro || '',
            cidade: bApi.municipio || '',
            codigo_ibge: bApi.codigo_municipio_ibge ? String(bApi.codigo_municipio_ibge) : '',
            uf: bApi.uf || '',
            cep: bApi.cep || '',
            telefone: bApi.ddd_telefone_1 || bApi.ddd_telefone_2 || '',
            email: bApi.email || ''
          };
        }
      } catch (e3) {
        console.warn('Falha na BrasilAPI direta:', e3);
      }
    }

    if (!data || !data.success) {
      const errMsg = (data && data.error) ? data.error : 'CNPJ não encontrado ou indisponível nas bases públicas.';
      showToast(errMsg, 'error');
      return;
    }

    inputEl.value = cleanCnpj.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, "$1.$2.$3/$4-$5");

    if (inputId === 'CPF') {
      setElVal('Nome', data.razao_social);
      setElVal('Fantasia', data.nome_fantasia || data.razao_social);
      setElVal('FisicaJuridica', '2');
      if (data.inscricao_estadual) setElVal('RG', data.inscricao_estadual);
      setElVal('Endereco', data.logradouro);
      setElVal('Nro', data.numero || 'SN');
      setElVal('Complemento', data.complemento || '');
      setElVal('Bairro', data.bairro || '');
      setElVal('Cidade', data.cidade || '');
      setElVal('Uf', data.uf ? data.uf.toUpperCase() : '');
      setElVal('Cep', formatCEPStr(data.cep));
      if (data.telefone) setElVal('Fone', data.telefone);
      if (data.email) setElVal('Email', data.email);

      showToast(`Dados de "${data.razao_social}" preenchidos com sucesso!`, 'success');
    } else if (inputId === 'Emp_CNPJ') {
      setElVal('Emp_RazaoSocial', data.razao_social);
      setElVal('Emp_Fantasia', data.nome_fantasia || data.razao_social);
      if (data.inscricao_estadual) setElVal('Emp_InscEst', data.inscricao_estadual);
      setElVal('Emp_Logradouro', data.logradouro);
      setElVal('Emp_Nro', data.numero || 'SN');
      setElVal('Emp_Bairro', data.bairro || '');
      setElVal('Emp_Cidade', data.cidade || '');
      setElVal('Emp_UF', data.uf ? data.uf.toUpperCase() : '');
      setElVal('Emp_CEP', formatCEPStr(data.cep));
      if (data.telefone) setElVal('Emp_Fone', data.telefone);
      if (data.email) setElVal('Emp_Email', data.email);
      if (data.codigo_ibge) setElVal('Emp_CodigoIBGE', data.codigo_ibge);
      if (data.regime_tributario) setElVal('Emp_RegimeTrib', data.regime_tributario);

      showToast(`Dados da Empresa "${data.razao_social}" preenchidos com sucesso!`, 'success');
    } else if (inputId === 'FrmNota_CpfCnpj') {
      setElVal('FrmNota_NomeCliente', data.razao_social);
      if (data.inscricao_estadual) setElVal('FrmNota_InscEst', data.inscricao_estadual);
      setElVal('FrmNota_Endereco', data.logradouro);
      setElVal('FrmNota_Nro', data.numero || 'SN');
      setElVal('FrmNota_Bairro', data.bairro || '');
      setElVal('FrmNota_Cidade', data.cidade || '');
      setElVal('FrmNota_Uf', data.uf ? data.uf.toUpperCase() : '');
      setElVal('FrmNota_Cep', formatCEPStr(data.cep));

      showToast(`Dados do destinatário "${data.razao_social}" preenchidos!`, 'success');
    } else {
      showToast(`Consulta concluída: ${data.razao_social}`, 'success');
    }
  } catch (err) {
    console.error('Erro na consulta do CNPJ:', err);
    showToast(`Erro ao consultar CNPJ: ${err.message || err}`, 'error');
  } finally {
    if (btnEl) {
      btnEl.disabled = false;
      btnEl.innerHTML = originalBtnHtml;
    }
  }
}

function formatCEPStr(cep) {
  if (!cep) return '';
  const clean = String(cep).replace(/\D/g, '');
  if (clean.length === 8) {
    return clean.replace(/^(\d{5})(\d{3})$/, '$1-$2');
  }
  return cep;
}

/* --- CERTIFICADO DIGITAL A1 / A3 INTERACTION HANDLERS --- */
let activeCertTargetPrefix = 'Cfg_Page';

function onCertTipoChange(prefix) {
  const tipoEl = document.getElementById(`${prefix}_Cert_Tipo`);
  if (!tipoEl) return;
  const tipo = tipoEl.value;

  const lblEl = document.getElementById(`lbl_${prefix}_Cert_Caminho`);
  const inputEl = document.getElementById(`${prefix}_Cert_Caminho`);
  const iconEl = document.getElementById(`icon_${prefix}_Cert_Browse`);
  const txtEl = document.getElementById(`txt_${prefix}_Cert_Browse`);
  const btnEl = document.getElementById(`btn_${prefix}_Cert_Browse`);

  if (tipo === 'A3') {
    if (lblEl) lblEl.innerText = 'Número de Série / Nome (Opções da Internet / Windows)';
    if (inputEl) inputEl.placeholder = 'Ex: 1A2B3C4D5E6F7G8H ou Nome do Titular nas Opções da Internet';
    if (iconEl) iconEl.className = 'fa-solid fa-globe';
    if (txtEl) txtEl.innerText = 'Opções da Internet';
    if (btnEl) btnEl.title = 'Buscar certificados A3 instalados nas Opções da Internet / Loja do Windows';
  } else {
    if (lblEl) lblEl.innerText = 'Caminho do Arquivo (.pfx / .p12) nas Pastas';
    if (inputEl) inputEl.placeholder = 'Ex: C:\\Certificados\\empresa.pfx';
    if (iconEl) iconEl.className = 'fa-solid fa-folder-open';
    if (txtEl) txtEl.innerText = 'Procurar Pasta...';
    if (btnEl) btnEl.title = 'Buscar arquivo de certificado (.pfx/.p12) em pastas no computador';
  }
}

function handleCertBrowse(prefix) {
  activeCertTargetPrefix = prefix;
  const tipoEl = document.getElementById(`${prefix}_Cert_Tipo`);
  const tipo = tipoEl ? tipoEl.value : 'A1';

  if (tipo === 'A3') {
    openWin32CertModal(prefix);
  } else {
    const fileInput = document.getElementById('global-cert-file-input');
    if (fileInput) {
      fileInput.value = '';
      fileInput.click();
    }
  }
}

async function onGlobalCertFileSelected(input) {
  if (!input.files || input.files.length === 0) return;
  const file = input.files[0];

  const inputEl = document.getElementById(`${activeCertTargetPrefix}_Cert_Caminho`);
  let currentVal = inputEl ? inputEl.value.trim() : '';

  let folderDir = '';
  if (currentVal) {
    const isOldProjectDefault = currentVal.toLowerCase().includes('projetos\\vendas\\certificados') || 
                                currentVal.toLowerCase().includes('projetos/vendas/certificados');
    if (!isOldProjectDefault) {
      const lastSlash = Math.max(currentVal.lastIndexOf('\\'), currentVal.lastIndexOf('/'));
      if (lastSlash !== -1) {
        folderDir = currentVal.substring(0, lastSlash + 1);
      }
    }
  }

  let desiredPath = folderDir ? (folderDir + file.name) : ('C:\\Certificados\\' + file.name);

  showToast(`Processando arquivo '${file.name}'...`, 'info');

  const formData = new FormData();
  formData.append('file', file);
  formData.append('desired_path', desiredPath);

  try {
    const res = await fetch('/api/config/certificado/upload', {
      method: 'POST',
      body: formData
    });
    const result = await res.json();
    if (!res.ok || !result.success) throw new Error(result.error || 'Erro no envio do certificado');

    const savePath = result.caminho || desiredPath;
    
    ['Cfg_Page', 'Cfg_Modal', 'Cfg_Main'].forEach(p => {
      const el = document.getElementById(`${p}_Cert_Caminho`);
      if (el) el.value = savePath;
      const tipoEl = document.getElementById(`${p}_Cert_Tipo`);
      if (tipoEl) {
        tipoEl.value = 'A1';
        onCertTipoChange(p);
      }
    });

    showToast(result.message || `Certificado salvo em '${savePath}' com sucesso!`, 'success');
  } catch (err) {
    showToast(`Erro ao carregar arquivo do certificado: ${err.message}`, 'error');
  }
}

async function openWin32CertModal(prefix) {
  if (prefix) activeCertTargetPrefix = prefix;
  const modal = document.getElementById('win32-cert-modal');
  const listEl = document.getElementById('win32-cert-list');
  if (!modal || !listEl) return;

  listEl.innerHTML = `
    <div style="padding: 1.5rem; text-align: center; color: var(--text-muted);">
      <i class="fa-solid fa-spinner fa-spin fa-2x"></i>
      <div style="margin-top: 0.5rem;">Buscando certificados instalados no Windows (Opções da Internet)...</div>
    </div>
  `;

  modal.classList.add('active');

  try {
    const res = await fetch('/api/config/certificado/win32-store');
    const data = await res.json();

    if (!res.ok || !data.success) throw new Error(data.error || 'Erro ao ler loja do Windows');

    const certs = data.certificates || [];
    if (certs.length === 0) {
      listEl.innerHTML = `
        <div style="padding: 1.25rem; text-align: center; background: var(--bg-surface-elevated); border-radius: var(--radius-md); border: 1px solid var(--border-color);">
          <i class="fa-solid fa-triangle-exclamation fa-2x" style="color: var(--accent-amber);"></i>
          <div style="font-weight: 600; margin-top: 0.5rem; color: var(--text-primary);">Nenhum certificado A3 encontrado nas Opções da Internet</div>
          <div style="font-size: 0.8rem; color: var(--text-muted); margin-top: 0.25rem;">
            Certifique-se de que o cartão/token A3 está inserido e instalado no Windows (Painel de Controle &gt; Opções da Internet &gt; Conteúdo &gt; Certificados).
          </div>
        </div>
      `;
      return;
    }

    listEl.innerHTML = certs.map(c => `
      <div style="padding: 0.85rem; border-radius: var(--radius-md); background: var(--bg-surface-elevated); border: 1px solid var(--border-color); display: flex; align-items: center; justify-content: space-between; gap: 0.5rem;">
        <div style="overflow: hidden;">
          <div style="font-weight: 600; color: var(--text-primary); text-overflow: ellipsis; overflow: hidden; white-space: nowrap;">
            <i class="fa-solid fa-shield-halved" style="color: var(--accent-blue);"></i> ${escapeHtml(c.subject)}
          </div>
          <div style="font-size: 0.775rem; color: var(--text-muted); margin-top: 2px;">
            Série: <code>${escapeHtml(c.serial_number || 'Sem Série')}</code> | Val: ${escapeHtml(c.validade || '-')}
          </div>
        </div>
        <button type="button" class="btn btn-primary btn-sm" onclick="selectWin32Cert('${escapeJsString(c.serial_number || c.subject)}', '${escapeJsString(c.subject)}')" style="white-space: nowrap;">
          <i class="fa-solid fa-check"></i> Selecionar
        </button>
      </div>
    `).join('');
  } catch (err) {
    listEl.innerHTML = `
      <div style="padding: 1rem; color: var(--accent-rose); text-align: center;">
        Erro ao buscar certificados: ${escapeHtml(err.message)}
      </div>
    `;
  }
}

function selectWin32Cert(serialOrSubject, name) {
  const targetVal = serialOrSubject || name;
  ['Cfg_Page', 'Cfg_Modal', 'Cfg_Main'].forEach(p => {
    const el = document.getElementById(`${p}_Cert_Caminho`);
    if (el) el.value = targetVal;
    const tipoEl = document.getElementById(`${p}_Cert_Tipo`);
    if (tipoEl) {
      tipoEl.value = 'A3';
      onCertTipoChange(p);
    }
  });

  showToast(`Certificado A3 selecionado: ${name}`, 'success');
  closeWin32CertModal();
}

function closeWin32CertModal() {
  const modal = document.getElementById('win32-cert-modal');
  if (modal) modal.classList.remove('active');
}


// ==========================================
// RELATÓRIO DE VENDAS AGRUPADO POR PRODUTO
// ==========================================

let relCurrentPage = 1;
let relTotalPages = 1;
let relReportData = null;
let relFiltros = {
  data_inicio: '',
  data_fim: '',
  grupo: 'all',
  status: 'ativos',
  sort_by: 'total_valor',
  sort_order: 'DESC',
  q: '',
  limit: 50
};

function setRelPeriodo(tipo) {
  const dtInicioEl = document.getElementById('rel-filtro-data-inicio');
  const dtFimEl = document.getElementById('rel-filtro-data-fim');
  if (!dtInicioEl || !dtFimEl) return;

  const now = new Date();
  const formatYMD = (d) => {
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  };

  if (tipo === 'hoje') {
    dtInicioEl.value = formatYMD(now);
    dtFimEl.value = formatYMD(now);
  } else if (tipo === '7dias') {
    const d7 = new Date();
    d7.setDate(d7.getDate() - 7);
    dtInicioEl.value = formatYMD(d7);
    dtFimEl.value = formatYMD(now);
  } else if (tipo === 'mes') {
    const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
    dtInicioEl.value = formatYMD(firstDay);
    dtFimEl.value = formatYMD(now);
  } else if (tipo === 'ano') {
    const firstDayYear = new Date(now.getFullYear(), 0, 1);
    dtInicioEl.value = formatYMD(firstDayYear);
    dtFimEl.value = formatYMD(now);
  } else if (tipo === 'tudo') {
    dtInicioEl.value = '';
    dtFimEl.value = '';
  }

  fetchRelatorioVendasProduto(1);
}

function clearRelFiltroBusca() {
  const el = document.getElementById('rel-filtro-busca');
  if (el) {
    el.value = '';
    relFiltros.q = '';
    fetchRelatorioVendasProduto(1);
  }
}

function limparFiltrosRelatorioVendasProduto() {
  const dtInicioEl = document.getElementById('rel-filtro-data-inicio');
  const dtFimEl = document.getElementById('rel-filtro-data-fim');
  const grupoEl = document.getElementById('rel-filtro-grupo');
  const statusEl = document.getElementById('rel-filtro-status');
  const ordEl = document.getElementById('rel-filtro-ordenacao');
  const buscaEl = document.getElementById('rel-filtro-busca');
  const limitEl = document.getElementById('rel-limit-select');

  if (dtInicioEl) dtInicioEl.value = '';
  if (dtFimEl) dtFimEl.value = '';
  if (grupoEl) grupoEl.value = 'all';
  if (statusEl) statusEl.value = 'ativos';
  if (ordEl) ordEl.value = 'total_valor|DESC';
  if (buscaEl) buscaEl.value = '';
  if (limitEl) limitEl.value = '50';

  relFiltros.q = '';
  relCurrentPage = 1;
  fetchRelatorioVendasProduto(1);
  showToast('Filtros restaurados para o padrão', 'info', 1500);
}

function changeRelatorioPage(delta) {
  const newPage = relCurrentPage + delta;
  if (newPage >= 1 && newPage <= relTotalPages) {
    fetchRelatorioVendasProduto(newPage);
  }
}

async function fetchRelatorioVendasProduto(page = 1) {
  relCurrentPage = page;

  const dtInicio = document.getElementById('rel-filtro-data-inicio')?.value || '';
  const dtFim = document.getElementById('rel-filtro-data-fim')?.value || '';
  const grupo = document.getElementById('rel-filtro-grupo')?.value || 'all';
  const status = document.getElementById('rel-filtro-status')?.value || 'ativos';
  const ordVal = document.getElementById('rel-filtro-ordenacao')?.value || 'total_valor|DESC';
  const qVal = document.getElementById('rel-filtro-busca')?.value || '';
  const limitVal = parseInt(document.getElementById('rel-limit-select')?.value || '50', 10);
  const idEmpresa = document.getElementById('select-active-empresa')?.value || 'all';

  const [sort_by, sort_order] = ordVal.split('|');

  const tbody = document.getElementById('tbody-rel-vendas-produto');
  if (tbody) {
    tbody.innerHTML = `
      <tr>
        <td colspan="13" class="loading-td">
          <i class="fa-solid fa-spinner fa-spin fa-2x" style="color: var(--accent-blue);"></i>
          <div style="margin-top: 0.5rem; font-weight: 500;">Consolidando vendas por produto...</div>
        </td>
      </tr>
    `;
  }

  try {
    const params = new URLSearchParams({
      page: relCurrentPage,
      limit: limitVal,
      sort_by: sort_by || 'total_valor',
      sort_order: sort_order || 'DESC',
      status: status
    });

    if (dtInicio) params.append('data_inicio', dtInicio);
    if (dtFim) params.append('data_fim', dtFim);
    if (grupo && grupo !== 'all') params.append('grupo', grupo);
    if (qVal && qVal.trim()) params.append('q', qVal.trim());
    if (idEmpresa && idEmpresa !== 'all') params.append('id_empresa', idEmpresa);

    const res = await fetch(`/api/relatorios/vendas-por-produto?${params.toString()}`);
    const data = await res.json();

    if (!res.ok) throw new Error(data.error || 'Falha ao consultar relatório');

    relReportData = data;
    relTotalPages = data.pages || 1;

    renderRelatorioKPIs(data.summary);
    renderRelatorioRanking(data.items, data.summary);
    renderRelatorioTabela(data.items, data.summary);
    renderRelatorioPaginacao(data.total, data.page, data.limit);

  } catch (err) {
    console.error('Erro no relatório de vendas:', err);
    if (tbody) {
      tbody.innerHTML = `
        <tr>
          <td colspan="13" style="text-align: center; color: var(--accent-rose); padding: 2rem;">
            <i class="fa-solid fa-circle-exclamation fa-2x"></i>
            <div style="margin-top: 0.5rem;">Erro ao gerar relatório: ${escapeHtml(err.message)}</div>
          </td>
        </tr>
      `;
    }
  }
}

function renderRelatorioKPIs(sum) {
  if (!sum) return;
  const faturamento = sum.total_faturamento_geral || 0;
  const bruto = sum.total_bruto_geral || 0;
  const descontos = sum.total_desconto_geral || 0;
  const qtd = sum.total_qtd_geral || 0;
  const pedidos = sum.total_pedidos_geral || 0;
  const distintos = sum.total_produtos_distintos || 0;
  const precoMedio = sum.preco_medio_geral || 0;
  const lucro = sum.lucro_bruto_geral || 0;
  const margem = sum.margem_lucro_pct_geral || 0;

  setElementText('rel-kpi-faturamento', `R$ ${faturamento.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
  setElementText('rel-kpi-bruto-sub', `Total Bruto: R$ ${bruto.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
  setElementText('rel-kpi-qtd', qtd.toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 2 }));
  
  const comNota = sum.total_pedidos_com_nota || 0;
  const semNota = sum.total_pedidos_sem_nota || 0;
  if (comNota > 0 || semNota > 0) {
    setElementText('rel-kpi-pedidos-sub', `Em ${pedidos.toLocaleString()} pedidos (${comNota.toLocaleString()} com NF • ${semNota.toLocaleString()} sem NF)`);
  } else {
    setElementText('rel-kpi-pedidos-sub', `Em ${pedidos.toLocaleString()} pedidos`);
  }
  
  setElementText('rel-kpi-produtos-distintos', distintos.toLocaleString());
  setElementText('rel-kpi-preco-medio', `R$ ${precoMedio.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
  setElementText('rel-kpi-descontos-sub', `Descontos: R$ ${descontos.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
  setElementText('rel-kpi-lucro', `R$ ${lucro.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
  setElementText('rel-kpi-margem-sub', `Margem Média: ${margem.toFixed(2)}%`);
  setElementText('rel-badge-total-registros', `${distintos.toLocaleString()} produtos vendidos`);
}

function renderRelatorioRanking(items, sum) {
  const container = document.getElementById('rel-top-ranking-container');
  if (!container) return;

  if (!items || items.length === 0) {
    container.innerHTML = '<div style="color: var(--text-muted); font-size: 0.85rem; padding: 0.5rem;">Nenhum dado para o período selecionado.</div>';
    return;
  }

  // Pegar os top 5 ordenados por total_valor
  const top5 = [...items].sort((a, b) => (b.total_valor || 0) - (a.total_valor || 0)).slice(0, 5);
  const maxVal = Math.max(...top5.map(t => t.total_valor || 0), 1);

  const medalColors = ['#f59e0b', '#94a3b8', '#d97706', '#3b82f6', '#8b5cf6'];
  const medalIcons = ['1º', '2º', '3º', '4º', '5º'];

  container.innerHTML = top5.map((item, idx) => {
    const val = item.total_valor || 0;
    const part = item.participacao_pct || 0;
    const barWidth = Math.min(Math.max((val / maxVal) * 100, 4), 100);
    const color = medalColors[idx] || '#3b82f6';
    const rankLabel = medalIcons[idx] || `${idx + 1}º`;

    return `
      <div style="display: flex; flex-direction: column; gap: 0.25rem;">
        <div style="display: flex; justify-content: space-between; align-items: center; font-size: 0.85rem;">
          <div style="display: flex; align-items: center; gap: 0.5rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">
            <span style="font-weight: 700; color: ${color}; min-width: 24px;">${rankLabel}</span>
            <strong style="color: var(--text-primary); cursor: pointer;" onclick="openProdutoVendasDetalhesModal(${item.CodPrd})" title="Clique para ver vendas detalhadas">
              #${item.CodPrd} - ${escapeHtml(item.Descricao_Produto)}
            </strong>
            <span class="badge badge-group" style="padding: 0.1rem 0.35rem; font-size: 0.7rem;">${escapeHtml(item.Descricao_Grupo || 'SEM GRUPO')}</span>
          </div>
          <div style="display: flex; align-items: center; gap: 0.75rem; font-size: 0.85rem; white-space: nowrap;">
            <span style="color: var(--text-secondary);">${item.total_qtd} ${escapeHtml(item.Embalagem || 'UN')}</span>
            <strong style="color: var(--accent-emerald);">R$ ${val.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</strong>
            <span class="badge" style="background: rgba(59, 130, 246, 0.1); color: var(--accent-blue); font-size: 0.75rem; padding: 0.1rem 0.4rem;">${part.toFixed(2)}%</span>
          </div>
        </div>
        <div style="width: 100%; height: 8px; background: var(--bg-surface-elevated); border-radius: 4px; overflow: hidden; border: 1px solid var(--border-color);">
          <div style="width: ${barWidth}%; height: 100%; background: linear-gradient(90deg, ${color} 0%, var(--accent-emerald) 100%); border-radius: 4px; transition: width 0.4s ease;"></div>
        </div>
      </div>
    `;
  }).join('');
}

function renderRelatorioTabela(items, sum) {
  const tbody = document.getElementById('tbody-rel-vendas-produto');
  if (!tbody) return;

  if (!items || items.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="13" class="loading-td">Nenhuma venda encontrada para os filtros selecionados.</td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = items.map(p => {
    const faturamentoStr = `R$ ${p.total_valor.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    const precoMedioStr = `R$ ${p.preco_medio.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    const descontoStr = p.total_desconto > 0 ? `R$ ${p.total_desconto.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : '-';
    const lucroStr = `R$ ${p.lucro_bruto.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    const margemClass = p.margem_pct >= 40 ? 'color: var(--accent-emerald); font-weight: 600;' : (p.margem_pct >= 20 ? 'color: var(--accent-blue); font-weight: 600;' : 'color: var(--accent-amber); font-weight: 600;');

    const hasPhoto = isValidPhotoUrl(p.Foto);
    const photoThumbnail = hasPhoto
      ? `<div style="width: 32px; height: 32px; border-radius: var(--radius-sm); overflow: hidden; border: 1px solid var(--border-color); background: var(--bg-surface); display: inline-flex; align-items: center; justify-content: center; cursor: pointer;" onclick="openProdutoVendasDetalhesModal(${p.CodPrd})" title="Ver detalhes do produto">
           <img src="${escapeHtml(p.Foto)}" alt="" style="width: 100%; height: 100%; object-fit: cover;" onerror="this.onerror=null; this.parentElement.innerHTML='<i class=\\'fa-solid fa-image\\' style=\\'color: var(--text-muted); opacity: 0.5;\\'></i>';">
         </div>`
      : `<div style="width: 32px; height: 32px; border-radius: var(--radius-sm); border: 1px dashed var(--border-color); background: rgba(255,255,255,0.02); display: inline-flex; align-items: center; justify-content: center; color: var(--text-muted);" title="Sem foto">
           <i class="fa-solid fa-box" style="opacity: 0.35; font-size: 0.75rem;"></i>
         </div>`;

    const barWidth = Math.min(Math.max(p.participacao_pct, 2), 100);

    return `
      <tr>
        <td style="text-align: center;">
          <div style="display: flex; align-items: center; gap: 0.4rem; justify-content: center;">
            ${photoThumbnail}
            <strong>#${p.CodPrd}</strong>
          </div>
        </td>
        <td>
          <div style="font-weight: 600; color: var(--text-primary); cursor: pointer;" onclick="openProdutoVendasDetalhesModal(${p.CodPrd})" title="Clique para ver vendas detalhadas">
            ${escapeHtml(p.Descricao_Produto)}
          </div>
          ${p.CodBar ? `<div style="font-size: 0.75rem; color: var(--text-muted);">EAN: ${escapeHtml(p.CodBar)}</div>` : ''}
        </td>
        <td><span class="badge badge-group">${escapeHtml(p.Descricao_Grupo || 'SEM GRUPO')}</span></td>
        <td style="text-align: center;"><span class="badge" style="background: var(--bg-surface-elevated); border: 1px solid var(--border-color); font-size: 0.75rem;">${escapeHtml(p.Embalagem || 'UN')}</span></td>
        <td style="text-align: right; font-weight: 600;">${p.total_qtd.toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 2 })}</td>
        <td style="text-align: right; color: var(--text-secondary);">${precoMedioStr}</td>
        <td style="text-align: right; color: var(--text-muted); font-size: 0.85rem;">${descontoStr}</td>
        <td style="text-align: right; font-weight: 700; color: var(--accent-emerald);">${faturamentoStr}</td>
        <td>
          <div style="display: flex; align-items: center; gap: 0.4rem;">
            <div style="flex: 1; height: 6px; background: var(--bg-surface-elevated); border-radius: 3px; overflow: hidden; border: 1px solid var(--border-color);">
              <div style="width: ${barWidth}%; height: 100%; background: linear-gradient(90deg, #3b82f6, #10b981); border-radius: 3px;"></div>
            </div>
            <span style="font-size: 0.75rem; font-weight: 600; min-width: 38px; text-align: right; color: var(--text-secondary);">${p.participacao_pct.toFixed(1)}%</span>
          </div>
        </td>
        <td style="text-align: right; color: var(--text-primary);">${lucroStr}</td>
        <td style="text-align: right; ${margemClass}">${p.margem_pct.toFixed(1)}%</td>
        <td style="text-align: right; color: ${p.EstoqueAtual <= 0 ? 'var(--accent-rose)' : 'var(--text-secondary)'}; font-size: 0.85rem;">
          ${p.EstoqueAtual.toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 2 })}
        </td>
        <td style="text-align: center;">
          <button type="button" class="btn btn-secondary btn-sm" onclick="openProdutoVendasDetalhesModal(${p.CodPrd})" title="Ver histórico de vendas do produto" style="padding: 0.2rem 0.5rem; font-size: 0.75rem; display: inline-flex; align-items: center; gap: 0.3rem;">
            <i class="fa-solid fa-list"></i> Vendas
          </button>
        </td>
      </tr>
    `;
  }).join('');
}

function renderRelatorioPaginacao(total, page, limit) {
  const infoEl = document.getElementById('rel-pagination-info');
  const indicatorEl = document.getElementById('rel-page-indicator');
  const btnPrev = document.getElementById('rel-btn-prev');
  const btnNext = document.getElementById('rel-btn-next');
  const footerEl = document.getElementById('rel-table-footer');

  if (!total || limit === 0) {
    if (footerEl) footerEl.style.display = limit === 0 ? 'none' : 'flex';
    if (infoEl) infoEl.innerText = `Total: ${total || 0} produtos`;
    return;
  }

  if (footerEl) footerEl.style.display = 'flex';

  const start = ((page - 1) * limit) + 1;
  const end = Math.min(page * limit, total);

  if (infoEl) infoEl.innerText = `Mostrando ${start} - ${end} de ${total} produtos`;
  if (indicatorEl) indicatorEl.innerText = `${page} / ${relTotalPages}`;

  if (btnPrev) btnPrev.disabled = (page <= 1);
  if (btnNext) btnNext.disabled = (page >= relTotalPages);
}

async function openProdutoVendasDetalhesModal(codPrd) {
  const modal = document.getElementById('rel-prd-detalhes-modal');
  if (!modal) return;

  modal.classList.add('active');

  const dtInicio = document.getElementById('rel-filtro-data-inicio')?.value || '';
  const dtFim = document.getElementById('rel-filtro-data-fim')?.value || '';
  const status = document.getElementById('rel-filtro-status')?.value || 'ativos';
  const idEmpresa = document.getElementById('select-active-empresa')?.value || 'all';

  const tbody = document.getElementById('tbody-rel-detalhes-vendas');
  if (tbody) {
    tbody.innerHTML = `
      <tr>
        <td colspan="10" class="loading-td">
          <i class="fa-solid fa-spinner fa-spin"></i> Carregando histórico de vendas...
        </td>
      </tr>
    `;
  }

  try {
    const params = new URLSearchParams({ status });
    if (dtInicio) params.append('data_inicio', dtInicio);
    if (dtFim) params.append('data_fim', dtFim);
    if (idEmpresa && idEmpresa !== 'all') params.append('id_empresa', idEmpresa);

    const res = await fetch(`/api/relatorios/vendas-por-produto/detalhes/${codPrd}?${params.toString()}`);
    const data = await res.json();

    if (!res.ok) throw new Error(data.error || 'Falha ao carregar detalhes');

    const prd = data.produto || {};
    const resumo = data.resumo_produto || {};
    const vendas = data.vendas || [];

    // Header info
    setElementText('rel-det-badge-cod', `Cód: #${prd.CodPrd || codPrd}`);
    setElementText('rel-det-badge-grupo', prd.Nome_Grupo || 'SEM GRUPO');
    setElementText('rel-det-badge-un', prd.Embalagem || 'UN');
    setElementText('rel-det-prd-nome', prd.Descricao_Produto || 'Produto');
    setElementText('rel-det-prd-codbar', prd.CodBar || '-');
    setElementText('rel-det-prd-venda', `R$ ${(prd.Venda || 0).toFixed(2)}`);
    setElementText('rel-det-prd-custo', `R$ ${(prd.Custo || 0).toFixed(2)}`);
    setElementText('rel-det-prd-estoque', `${prd.Estoque || 0} ${prd.Embalagem || 'UN'}`);

    const fotoWrap = document.getElementById('rel-det-foto-wrapper');
    if (fotoWrap) {
      if (isValidPhotoUrl(prd.Foto)) {
        fotoWrap.innerHTML = `<img src="${escapeHtml(prd.Foto)}" alt="" style="width: 100%; height: 100%; object-fit: cover;">`;
      } else {
        fotoWrap.innerHTML = `<i class="fa-solid fa-box fa-2x" style="color: var(--text-muted); opacity: 0.5;"></i>`;
      }
    }

    // KPI Badges
    setElementText('rel-det-kpi-pedidos', (resumo.total_pedidos || vendas.length).toLocaleString());
    setElementText('rel-det-kpi-qtd', (resumo.total_qtd || 0).toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 2 }));
    setElementText('rel-det-kpi-preco-medio', `R$ ${(resumo.preco_medio || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
    setElementText('rel-det-kpi-faturamento', `R$ ${(resumo.total_valor || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);

    // Transactions Table
    if (tbody) {
      if (vendas.length === 0) {
        tbody.innerHTML = `
          <tr>
            <td colspan="11" class="loading-td">Nenhuma venda registrada para este produto no período.</td>
          </tr>
        `;
      } else {
        tbody.innerHTML = vendas.map(v => {
          const vlrTotStr = `R$ ${(v.ValorTotal || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
          const vlrUnitStr = `R$ ${(v.ValorUnit || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
          const descStr = v.Desconto > 0 ? `R$ ${(v.Desconto || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : '-';

          let statusBadge = '<span class="badge" style="background: rgba(245, 158, 11, 0.15); color: #d97706; border: 1px solid rgba(245, 158, 11, 0.3); font-size: 0.72rem; padding: 0.15rem 0.4rem; white-space: nowrap;"><i class="fa-solid fa-clock"></i> Sem Nota</span>';
          if (v.is_cancelado) {
            statusBadge = '<span class="badge" style="background: rgba(239, 68, 68, 0.15); color: #ef4444; border: 1px solid rgba(239, 68, 68, 0.3); font-size: 0.72rem; padding: 0.15rem 0.4rem; white-space: nowrap;"><i class="fa-solid fa-ban"></i> Cancelado</span>';
          } else if (v.has_nota) {
            const nro = v.NroNt > 0 ? `NF #${v.NroNt}` : 'NF-e Emitida';
            statusBadge = `<span class="badge" style="background: rgba(16, 185, 129, 0.15); color: #059669; border: 1px solid rgba(16, 185, 129, 0.3); font-size: 0.72rem; padding: 0.15rem 0.4rem; white-space: nowrap;"><i class="fa-solid fa-receipt"></i> ${escapeHtml(nro)}</span>`;
          }

          return `
            <tr>
              <td><strong style="color: var(--accent-blue);">#${v.CodPed}</strong></td>
              <td>${statusBadge}</td>
              <td>${escapeHtml(v.DataEmiss || '-')}</td>
              <td style="color: var(--text-muted); font-size: 0.8rem;">${escapeHtml(v.Hora || '-')}</td>
              <td><div style="font-weight: 500;">${escapeHtml(v.NomeCliente || 'CONSUMIDOR')}</div></td>
              <td style="color: var(--text-muted); font-size: 0.8rem;">${escapeHtml(v.CidadeCliente || '-')}</td>
              <td style="text-align: right; font-weight: 600;">${(v.Qtd || 0).toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 2 })}</td>
              <td style="text-align: right; color: var(--text-secondary);">${vlrUnitStr}</td>
              <td style="text-align: right; color: var(--text-muted); font-size: 0.8rem;">${descStr}</td>
              <td style="text-align: right; font-weight: 700; color: var(--accent-emerald);">${vlrTotStr}</td>
              <td><span class="badge" style="background: var(--bg-surface); border: 1px solid var(--border-color); font-size: 0.75rem;">${escapeHtml(v.CondPgto || 'A VISTA')}</span></td>
            </tr>
          `;
        }).join('');
      }
    }

  } catch (err) {
    showToast(`Erro ao carregar detalhes: ${err.message}`, 'error');
  }
}

function closeProdutoVendasDetalhesModal() {
  const modal = document.getElementById('rel-prd-detalhes-modal');
  if (modal) modal.classList.remove('active');
}

async function imprimirRelatorioVendasProduto() {
  const dtInicio = document.getElementById('rel-filtro-data-inicio')?.value || '';
  const dtFim = document.getElementById('rel-filtro-data-fim')?.value || '';
  const grupo = document.getElementById('rel-filtro-grupo')?.value || 'all';
  const status = document.getElementById('rel-filtro-status')?.value || 'ativos';
  const ordVal = document.getElementById('rel-filtro-ordenacao')?.value || 'total_valor|DESC';
  const qVal = document.getElementById('rel-filtro-busca')?.value || '';
  const idEmpresa = document.getElementById('select-active-empresa')?.value || '1';

  const [sort_by, sort_order] = ordVal.split('|');

  showToast('Preparando relatório formatado para impressão...', 'info', 2000);

  try {
    const params = new URLSearchParams({
      sort_by: sort_by || 'total_valor',
      sort_order: sort_order || 'DESC',
      status: status,
      limit: 0 // Fetch all rows for complete printing
    });

    if (dtInicio) params.append('data_inicio', dtInicio);
    if (dtFim) params.append('data_fim', dtFim);
    if (grupo && grupo !== 'all') params.append('grupo', grupo);
    if (qVal && qVal.trim()) params.append('q', qVal.trim());
    if (idEmpresa && idEmpresa !== 'all') params.append('id_empresa', idEmpresa);

    const [resRel, resEmp] = await Promise.all([
      fetch(`/api/relatorios/vendas-por-produto?${params.toString()}`),
      fetch(`/api/empresas/${idEmpresa}`).catch(() => null)
    ]);

    const data = await resRel.json();
    if (!resRel.ok) throw new Error(data.error || 'Falha ao obter dados para impressão');

    let empData = { RazaoSocial: 'SISTEMA DE VENDAS', CNPJ: '', InscrEst: '', Fone: '', Cidade: '', Uf: '' };
    if (resEmp && resEmp.ok) {
      try {
        const empJson = await resEmp.json();
        if (empJson) empData = { ...empData, ...empJson };
      } catch (e) {}
    }

    const items = data.items || [];
    const sum = data.summary || {};

    if (items.length === 0) {
      showToast('Nenhum registro encontrado para impressão.', 'warning');
      return;
    }

    const now = new Date();
    const dataHoraEmissao = now.toLocaleDateString('pt-BR') + ' ' + now.toLocaleTimeString('pt-BR');
    const periodoFormatado = (dtInicio || dtFim) 
      ? `De ${dtInicio ? dtInicio.split('-').reverse().join('/') : 'Início'} até ${dtFim ? dtFim.split('-').reverse().join('/') : 'Hoje'}`
      : 'Todo o Histórico de Vendas';

    const statusLabelsMap = {
      'ativos': 'Todos os Pedidos (Faturados e Pendentes de Nota)',
      'com_nota': 'Apenas com Nota Fiscal (Faturados)',
      'sem_nota': 'Apenas Pendentes de Nota (Sem NF)',
      'cancelados': 'Apenas Cancelados',
      'todos': 'Todos os Pedidos (Inclusive Cancelados)'
    };
    const statusLabel = statusLabelsMap[status] || 'Todos os Pedidos';
    const grupoLabel = grupo && grupo !== 'all' ? (document.getElementById('rel-filtro-grupo')?.selectedOptions[0]?.text || `Grupo #${grupo}`) : 'Todos os Grupos';

    const faturamentoStr = `R$ ${(sum.total_faturamento_geral || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    const brutoStr = `R$ ${(sum.total_bruto_geral || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    const descontoStr = `R$ ${(sum.total_desconto_geral || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    const lucroStr = `R$ ${(sum.lucro_bruto_geral || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    const precoMedioStr = `R$ ${(sum.preco_medio_geral || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

    const rowsHtml = items.map((p, idx) => {
      const vlrTot = `R$ ${p.total_valor.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
      const pMed = `R$ ${p.preco_medio.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
      const desc = p.total_desconto > 0 ? `R$ ${p.total_desconto.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : '-';
      const luc = `R$ ${p.lucro_bruto.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

      return `
        <tr style="background-color: ${idx % 2 === 0 ? '#ffffff' : '#f9fafb'};">
          <td style="text-align: center; font-weight: 600;">${p.CodPrd}</td>
          <td>
            <strong>${escapeHtml(p.Descricao_Produto)}</strong>
            ${p.CodBar ? `<div style="font-size: 8pt; color: #64748b;">EAN: ${escapeHtml(p.CodBar)}</div>` : ''}
          </td>
          <td>${escapeHtml(p.Descricao_Grupo || 'SEM GRUPO')}</td>
          <td style="text-align: center;">${escapeHtml(p.Embalagem || 'UN')}</td>
          <td style="text-align: right; font-weight: 600;">${p.total_qtd.toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 2 })}</td>
          <td style="text-align: right;">${pMed}</td>
          <td style="text-align: right; color: #64748b;">${desc}</td>
          <td style="text-align: right; font-weight: 700; color: #047857;">${vlrTot}</td>
          <td style="text-align: right; font-weight: 600;">${p.participacao_pct.toFixed(2)}%</td>
          <td style="text-align: right;">${luc}</td>
          <td style="text-align: right; font-weight: 600;">${p.margem_pct.toFixed(1)}%</td>
          <td style="text-align: right;">${p.EstoqueAtual.toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 2 })}</td>
        </tr>
      `;
    }).join('');

    const printDoc = `
      <!DOCTYPE html>
      <html lang="pt-BR">
      <head>
        <meta charset="UTF-8">
        <title>Relatório de Vendas por Produto - ${empData.RazaoSocial || 'Sistema Vendas'}</title>
        <style>
          @page {
            size: A4 portrait;
            margin: 10mm 10mm 12mm 10mm;
          }
          * {
            box-sizing: border-box;
            font-family: Arial, Helvetica, sans-serif;
            color: #1e293b;
          }
          body {
            margin: 0;
            padding: 0;
            background: #ffffff;
            font-size: 8.5pt;
            line-height: 1.3;
          }
          .header-box {
            border-bottom: 2px solid #0f172a;
            padding-bottom: 6px;
            margin-bottom: 10px;
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
          }
          .company-name {
            font-size: 13pt;
            font-weight: 700;
            color: #0f172a;
            margin: 0 0 2px 0;
            text-transform: uppercase;
          }
          .company-sub {
            font-size: 8pt;
            color: #475569;
            margin: 0;
          }
          .report-badge {
            text-align: right;
          }
          .report-title {
            font-size: 11pt;
            font-weight: 700;
            color: #1e3a8a;
            margin: 0 0 3px 0;
          }
          .report-meta {
            font-size: 7.5pt;
            color: #64748b;
          }
          .filter-box {
            background: #f1f5f9;
            border: 1px solid #cbd5e1;
            border-radius: 4px;
            padding: 6px 10px;
            margin-bottom: 10px;
            font-size: 8pt;
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
          }
          .filter-item strong {
            color: #0f172a;
          }
          .kpi-grid {
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 6px;
            margin-bottom: 12px;
          }
          .kpi-card {
            background: #f8fafc;
            border: 1px solid #cbd5e1;
            border-radius: 4px;
            padding: 6px 8px;
            text-align: center;
          }
          .kpi-card .kpi-label {
            font-size: 7pt;
            font-weight: 700;
            text-transform: uppercase;
            color: #475569;
            margin-bottom: 2px;
          }
          .kpi-card .kpi-val {
            font-size: 10pt;
            font-weight: 700;
            color: #0f172a;
          }
          .kpi-card .kpi-sub {
            font-size: 6.5pt;
            color: #64748b;
            margin-top: 2px;
          }
          table.report-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 8pt;
          }
          table.report-table thead {
            display: table-header-group;
          }
          table.report-table tr {
            page-break-inside: avoid;
          }
          table.report-table th {
            background-color: #0f172a;
            color: #ffffff;
            font-weight: 700;
            font-size: 7.5pt;
            padding: 6px 5px;
            text-align: left;
            border: 1px solid #0f172a;
          }
          table.report-table td {
            padding: 5px 5px;
            border: 1px solid #e2e8f0;
            vertical-align: middle;
          }
          .footer-box {
            margin-top: 14px;
            padding-top: 6px;
            border-top: 1px solid #cbd5e1;
            font-size: 7pt;
            color: #64748b;
            display: flex;
            justify-content: space-between;
          }
          @media print {
            body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
            .no-print { display: none !important; }
          }
        </style>
      </head>
      <body>
        <div class="header-box">
          <div>
            <div class="company-name">${escapeHtml(empData.RazaoSocial || 'EMPRESA MATRIZ')}</div>
            <div class="company-sub">
              ${empData.CNPJ ? `CNPJ: ${escapeHtml(empData.CNPJ)}` : ''} 
              ${empData.InscrEst ? ` | IE: ${escapeHtml(empData.InscrEst)}` : ''}
              ${empData.Fone ? ` | Fone: ${escapeHtml(empData.Fone)}` : ''}
              ${empData.Cidade ? ` | ${escapeHtml(empData.Cidade)}/${escapeHtml(empData.Uf || '')}` : ''}
            </div>
          </div>
          <div class="report-badge">
            <div class="report-title">RELATÓRIO DE VENDAS POR PRODUTO</div>
            <div class="report-meta">Emissão: ${dataHoraEmissao}</div>
          </div>
        </div>

        <div class="filter-box">
          <div class="filter-item"><strong>Período:</strong> ${escapeHtml(periodoFormatado)}</div>
          <div class="filter-item"><strong>Grupo:</strong> ${escapeHtml(grupoLabel)}</div>
          <div class="filter-item"><strong>Status:</strong> ${escapeHtml(statusLabel)}</div>
          ${qVal ? `<div class="filter-item"><strong>Busca:</strong> "${escapeHtml(qVal)}"</div>` : ''}
          <div class="filter-item"><strong>Total Itens:</strong> ${items.length} produtos listados</div>
        </div>

        <div class="kpi-grid">
          <div class="kpi-card" style="border-top: 3px solid #059669;">
            <div class="kpi-label">Faturamento Total</div>
            <div class="kpi-val" style="color: #047857;">${faturamentoStr}</div>
            <div class="kpi-sub">Bruto: ${brutoStr}</div>
          </div>
          <div class="kpi-card" style="border-top: 3px solid #2563eb;">
            <div class="kpi-label">Qtd Total Vendida</div>
            <div class="kpi-val">${(sum.total_qtd_geral || 0).toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 2 })}</div>
            <div class="kpi-sub">${(sum.total_pedidos_geral || 0).toLocaleString()} pedidos (${(sum.total_pedidos_com_nota || 0).toLocaleString()} c/ NF • ${(sum.total_pedidos_sem_nota || 0).toLocaleString()} s/ NF)</div>
          </div>
          <div class="kpi-card" style="border-top: 3px solid #8b5cf6;">
            <div class="kpi-label">Produtos Distintos</div>
            <div class="kpi-val">${(sum.total_produtos_distintos || 0).toLocaleString()}</div>
            <div class="kpi-sub">Itens comercializados</div>
          </div>
          <div class="kpi-card" style="border-top: 3px solid #d97706;">
            <div class="kpi-label">Preço Médio / Desconto</div>
            <div class="kpi-val">${precoMedioStr}</div>
            <div class="kpi-sub">Desc: ${descontoStr}</div>
          </div>
          <div class="kpi-card" style="border-top: 3px solid #0284c7;">
            <div class="kpi-label">Lucro Bruto Estimado</div>
            <div class="kpi-val" style="color: #0369a1;">${lucroStr}</div>
            <div class="kpi-sub">Margem: ${(sum.margem_lucro_pct_geral || 0).toFixed(2)}%</div>
          </div>
        </div>

        <table class="report-table">
          <thead>
            <tr>
              <th style="width: 45px; text-align: center;">CÓD.</th>
              <th>DESCRIÇÃO DO PRODUTO</th>
              <th>GRUPO</th>
              <th style="width: 30px; text-align: center;">UN</th>
              <th style="width: 60px; text-align: right;">QTD</th>
              <th style="width: 65px; text-align: right;">P. MÉDIO</th>
              <th style="width: 55px; text-align: right;">DESC.</th>
              <th style="width: 75px; text-align: right;">TOTAL (R$)</th>
              <th style="width: 50px; text-align: right;">% PART.</th>
              <th style="width: 65px; text-align: right;">LUCRO</th>
              <th style="width: 45px; text-align: right;">MARG.</th>
              <th style="width: 50px; text-align: right;">ESTOQUE</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
          <tfoot>
            <tr style="background: #e2e8f0; font-weight: 700;">
              <td colspan="4" style="text-align: right;">TOTAIS GERAIS:</td>
              <td style="text-align: right;">${(sum.total_qtd_geral || 0).toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 2 })}</td>
              <td style="text-align: right;">${precoMedioStr}</td>
              <td style="text-align: right;">${descontoStr}</td>
              <td style="text-align: right; color: #047857;">${faturamentoStr}</td>
              <td style="text-align: right;">100.00%</td>
              <td style="text-align: right;">${lucroStr}</td>
              <td style="text-align: right;">${(sum.margem_lucro_pct_geral || 0).toFixed(1)}%</td>
              <td style="text-align: right;">-</td>
            </tr>
          </tfoot>
        </table>

        <div class="footer-box">
          <span>Sistema Vendas - Módulo Gestão Comercial & PDV Multiempresa</span>
          <span>Página 1 de 1</span>
        </div>

        <script>
          window.onload = function() {
            setTimeout(function() {
              window.print();
            }, 250);
          };
        </script>
      </body>
      </html>
    `;

    // Open clean print window
    const printWindow = window.open('', '_blank', 'width=1000,height=800,menubar=no,toolbar=no,location=no,status=no');
    if (!printWindow) {
      // Fallback if popup blocker active: use hidden iframe
      let printFrame = document.getElementById('report-print-iframe');
      if (!printFrame) {
        printFrame = document.createElement('iframe');
        printFrame.id = 'report-print-iframe';
        printFrame.style.position = 'fixed';
        printFrame.style.right = '0';
        printFrame.style.bottom = '0';
        printFrame.style.width = '0';
        printFrame.style.height = '0';
        printFrame.style.border = '0';
        document.body.appendChild(printFrame);
      }
      const frameDoc = printFrame.contentWindow.document;
      frameDoc.open();
      frameDoc.write(printDoc);
      frameDoc.close();
      setTimeout(() => {
        printFrame.contentWindow.focus();
        printFrame.contentWindow.print();
      }, 500);
      return;
    }

    printWindow.document.open();
    printWindow.document.write(printDoc);
    printWindow.document.close();

  } catch (err) {
    console.error('Erro ao imprimir relatório:', err);
    showToast(`Erro ao preparar impressão: ${err.message}`, 'error');
  }
}

function exportarRelatorioVendasProdutoCSV() {
  const dtInicio = document.getElementById('rel-filtro-data-inicio')?.value || '';
  const dtFim = document.getElementById('rel-filtro-data-fim')?.value || '';
  const grupo = document.getElementById('rel-filtro-grupo')?.value || 'all';
  const status = document.getElementById('rel-filtro-status')?.value || 'ativos';
  const ordVal = document.getElementById('rel-filtro-ordenacao')?.value || 'total_valor|DESC';
  const qVal = document.getElementById('rel-filtro-busca')?.value || '';
  const idEmpresa = document.getElementById('select-active-empresa')?.value || 'all';

  const [sort_by, sort_order] = ordVal.split('|');

  const params = new URLSearchParams({
    sort_by: sort_by || 'total_valor',
    sort_order: sort_order || 'DESC',
    status: status
  });

  if (dtInicio) params.append('data_inicio', dtInicio);
  if (dtFim) params.append('data_fim', dtFim);
  if (grupo && grupo !== 'all') params.append('grupo', grupo);
  if (qVal && qVal.trim()) params.append('q', qVal.trim());
  if (idEmpresa && idEmpresa !== 'all') params.append('id_empresa', idEmpresa);

  showToast('Iniciando download do relatório em CSV/Excel...', 'info', 2000);
  window.location.href = `/api/relatorios/vendas-por-produto/export/csv?${params.toString()}`;
}






