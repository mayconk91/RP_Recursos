/* tour.js — guided tour per first access of each tab
 * - Vanilla JS, offline-friendly
 * - Shows once per tab (localStorage)
 * - Can be relaunched via "? Guia" button
 */
(function(){
  const LS_KEY = 'rp_tour_seen_v1';

  const TAB_LABEL = {
    plan: 'Planejamento',
    exec: 'Visão Executiva',
    avail: 'Disponibilidade',
    cap: 'Capacidade',
    util: 'Exportações & Backup'
  };

  // Steps per tab. Each step: selector + title + text.
  const STEPS = {
    plan: [
      { sel: '.panel.filters', title: 'Filtros e período', text: 'Aqui você filtra por status, tipo, senioridade, tags e define o período (Início/Fim da visão).' },
      { sel: '#btnNovoRecurso', title: 'Adicionar recurso', text: 'Crie um recurso (pessoa/fornecedor) com tipo, senioridade e capacidade.' },
      { sel: '#btnNovaAtividade', title: 'Adicionar atividade', text: 'Cadastre uma atividade e vincule ao recurso. O sistema valida datas, alocação e histórico.' },
      { sel: '#gantt', title: 'Gantt e carga', text: 'Aqui você vê atividades por recurso. O heatmap indica ocupação diária e possíveis sobrecargas.' },
      { sel: '#aggPanel', title: 'Capacidade agregada', text: 'Visão semanal/mensal de ocupação agregada por recurso no período.' },
      { sel: '#recursos-container', title: 'Lista de recursos', text: 'Cards de recursos: editar/excluir e checar dados principais.' },
      { sel: '#atividades-container', title: 'Lista de atividades', text: 'Cards de atividades: editar, excluir, ver histórico e tags.' }
    ],
    exec: [
      { sel: '#kpiPanel', title: 'Indicadores', text: 'Resumo do período: % execução, recursos ativos e quantos estão sobrecarregados.' },
      { sel: '#exec-overload-panel', title: 'Detalhe de sobrecargas', text: 'Mostra períodos de sobreposição e atividades concorrentes por recurso.' },
      { sel: '#riskPanel', title: 'Risco por recurso', text: 'Score baseado em dias de sobrecarga (com +5 para externos). Use o filtro para ver somente quem tem risco.' }
    ],
    avail: [
      { sel: '#tab-avail .panel', title: 'Buscar disponibilidade', text: 'Simule alocação por dias necessários, dias úteis e % de capacidade requerida. Encontre recursos elegíveis no período.' }
    ],
    cap: [
      { sel: '#tab-cap iframe', title: 'Calculadora de capacidade', text: 'Ferramenta auxiliar para estimar capacidade/consumo. Útil para padronizar cálculos antes de cadastrar atividades.' }
    ],
    util: [
      { sel: '#bdControls', title: 'Banco de dados', text: 'Escolha/importe BD, defina padrão e exporte modelos. Este painel controla leitura/gravação do seu BD.' },
      { sel: '#exportFiltered', title: 'Exportação filtrada', text: 'Gere CSV/PDF de uma visão geral recortando por período, recurso e tipo.' },
      { sel: '#btnBackup', title: 'Backup', text: 'Crie um backup JSON para restaurar depois (boa prática antes de mudanças grandes).' }
    ]
  };

  let state = {
    tab: null,
    steps: [],
    idx: 0,
    backdrop: null,
    tooltip: null,
    arrow: null,
    highlighted: null
  };

  function loadSeen(){
    try {
      const raw = localStorage.getItem(LS_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch { return {}; }
  }
  function saveSeen(obj){
    try { localStorage.setItem(LS_KEY, JSON.stringify(obj)); } catch {}
  }

  function clearHighlight(){
    if(state.highlighted){
      state.highlighted.classList.remove('tour-highlight');
      state.highlighted = null;
    }
  }

  function teardown(){
    clearHighlight();
    if(state.backdrop){ state.backdrop.remove(); state.backdrop = null; }
    if(state.tooltip){ state.tooltip.remove(); state.tooltip = null; }
    if(state.arrow){ state.arrow.remove(); state.arrow = null; }
    state.tab = null; state.steps = []; state.idx = 0;
  }

  function ensureUI(){
    if(state.backdrop) return;

    const backdrop = document.createElement('div');
    backdrop.className = 'tour-backdrop';
    backdrop.addEventListener('click', () => teardown());

    const tooltip = document.createElement('div');
    tooltip.className = 'tour-tooltip';
    tooltip.setAttribute('role', 'dialog');
    tooltip.setAttribute('aria-modal', 'true');

    const arrow = document.createElement('div');
    arrow.className = 'tour-arrow';

    document.body.appendChild(backdrop);
    document.body.appendChild(arrow);
    document.body.appendChild(tooltip);

    state.backdrop = backdrop;
    state.tooltip = tooltip;
    state.arrow = arrow;

    // ESC to close
    document.addEventListener('keydown', onKeyDown);
  }

  function onKeyDown(e){
    if(!state.backdrop) return;
    if(e.key === 'Escape') teardown();
    if(e.key === 'ArrowRight') next();
    if(e.key === 'ArrowLeft') prev();
  }

  function getTarget(step){
    if(!step || !step.sel) return null;
    const el = document.querySelector(step.sel);
    return el || null;
  }

  function scrollIntoViewIfNeeded(el){
    try { el.scrollIntoView({ block: 'center', inline: 'nearest', behavior: 'smooth' }); } catch {}
  }

  function positionTooltip(target){
    const rect = target.getBoundingClientRect();
    const pad = 12;

    // Prefer right side, fallback to bottom
    const vw = window.innerWidth;
    const vh = window.innerHeight;

    const tt = state.tooltip;
    const ar = state.arrow;

    // Temporarily set offscreen to measure
    tt.style.left = '-9999px';
    tt.style.top = '-9999px';
    ar.style.left = '-9999px';
    ar.style.top = '-9999px';

    const ttRect = tt.getBoundingClientRect();
    const ttW = ttRect.width || 360;
    const ttH = ttRect.height || 140;

    let left = rect.right + pad;
    let top = rect.top;
    let placement = 'right';

    if(left + ttW + pad > vw){
      // move to left
      left = rect.left - ttW - pad;
      placement = 'left';
    }
    if(left < pad){
      // fallback bottom
      left = Math.min(Math.max(pad, rect.left), vw - ttW - pad);
      top = rect.bottom + pad;
      placement = 'bottom';
    }

    // clamp vertically
    top = Math.min(Math.max(pad, top), vh - ttH - pad);

    tt.style.left = `${left}px`;
    tt.style.top = `${top}px`;

    // Arrow positioning
    const arrowSize = 14;
    let ax = 0, ay = 0;
    if(placement === 'right'){
      ax = rect.right + pad - (arrowSize/2);
      ay = Math.min(Math.max(pad, rect.top + 10), vh - arrowSize - pad);
    } else if(placement === 'left'){
      ax = rect.left - pad - (arrowSize/2);
      ay = Math.min(Math.max(pad, rect.top + 10), vh - arrowSize - pad);
    } else {
      ax = Math.min(Math.max(pad, rect.left + 18), vw - arrowSize - pad);
      ay = rect.bottom + pad - (arrowSize/2);
    }
    ar.style.left = `${ax}px`;
    ar.style.top = `${ay}px`;
  }

  function renderStep(){
    const step = state.steps[state.idx];
    if(!step){ teardown(); return; }

    const target = getTarget(step);
    if(!target){
      // skip missing selectors
      if(state.idx < state.steps.length-1){ state.idx++; renderStep(); }
      else teardown();
      return;
    }

    ensureUI();
    clearHighlight();

    scrollIntoViewIfNeeded(target);
    target.classList.add('tour-highlight');
    state.highlighted = target;

    const tabName = TAB_LABEL[state.tab] || state.tab;
    const isLast = state.idx === state.steps.length - 1;

    state.tooltip.innerHTML = `
      <h4>${tabName} — ${step.title}</h4>
      <p>${step.text}</p>
      <div class="tour-actions">
        <button class="btn" id="tourPrev" ${state.idx===0?'disabled':''}>Voltar</button>
        <button class="btn" id="tourNext">${isLast ? 'Concluir' : 'Próximo'}</button>
        <button class="btn danger" id="tourClose">Fechar</button>
      </div>
    `;

    // wire buttons
    state.tooltip.querySelector('#tourPrev').addEventListener('click', prev);
    state.tooltip.querySelector('#tourNext').addEventListener('click', () => isLast ? finish() : next());
    state.tooltip.querySelector('#tourClose').addEventListener('click', teardown);

    // position after DOM update
    requestAnimationFrame(() => {
      positionTooltip(target);
    });
  }

  function next(){
    if(!state.steps.length) return;
    state.idx = Math.min(state.idx + 1, state.steps.length - 1);
    renderStep();
  }

  function prev(){
    if(!state.steps.length) return;
    state.idx = Math.max(state.idx - 1, 0);
    renderStep();
  }

  function finish(){
    const seen = loadSeen();
    seen[state.tab] = true;
    saveSeen(seen);
    teardown();
  }

  function startTour(tab, force){
    const steps = STEPS[tab] || [];
    if(!steps.length) return;

    const seen = loadSeen();
    if(!force && seen[tab]) return;

    // Only run if the panel exists
    const panel = document.getElementById('tab-' + tab);
    if(!panel) return;

    teardown();
    state.tab = tab;
    state.steps = steps;
    state.idx = 0;
    renderStep();
  }

  // Expose a small API
  window.rpTour = {
    start: (tab, force=true) => startTour(tab, !!force),
    reset: () => saveSeen({})
  };

  function addHelpButton(){
    // Adds a small help button near the tabs to relaunch tours
    if(document.getElementById('btnTour')) return;

    const header = document.querySelector('.topbar');
    if(!header) return;

    const btn = document.createElement('button');
    btn.id = 'btnTour';
    btn.className = 'btn';
    btn.textContent = '❓ Guia';
    btn.title = 'Abrir guia desta aba';
    btn.addEventListener('click', () => {
      const active = document.querySelector('.tab.active');
      const tab = active ? active.dataset.tab : 'plan';
      startTour(tab, true);
    });

    // Insert next to banner if possible
    const banner = document.getElementById('dbStatusBanner');
    if(banner && banner.parentElement){
      banner.parentElement.insertBefore(btn, banner);
    } else {
      header.appendChild(btn);
    }
  }

  document.addEventListener('DOMContentLoaded', () => {
    addHelpButton();

    // Run on first load for the initial active tab
    const initial = document.querySelector('.tab.active');
    startTour(initial ? initial.dataset.tab : 'plan', false);

    // Hook tab changes
    const originalActivate = window.activateTab;
    if(typeof originalActivate === 'function'){
      window.activateTab = function(name){
        originalActivate(name);
        startTour(name, false);
      };
    } else {
      // Fallback: listen to tab clicks
      document.addEventListener('click', (ev) => {
        const b = ev.target.closest('.tab');
        if(b && b.dataset && b.dataset.tab){
          setTimeout(() => startTour(b.dataset.tab, false), 0);
        }
      });
    }
  });
})();
