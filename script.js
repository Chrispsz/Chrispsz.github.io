(() => {
  'use strict';

  // UTILS
  const U = {
    qs: (s) => document.querySelector(s),
    qsa: (s) => [...document.querySelectorAll(s)],
    esc: (s = '') => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'),
    renderIcons: () => window.lucide?.createIcons(),
    digits: (s) => String(s || '').replace(/\D/g, ''),
    norm: (s) => String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim(),
    escapeCSV: (str) => {
      if (str == null) return '';
      str = String(str);
      if (/[",;\n\r]/.test(str)) return '"' + str.replace(/"/g, '""') + '"';
      return str;
    },
    formatBR: (ddd, tel) => {
      const d = U.digits(ddd), n = U.digits(tel);
      if (!n) return '';
      if (!d) return n;
      if (n.length === 9) return `(${d}) ${n.slice(0, 5)}-${n.slice(5)}`;
      if (n.length === 8) return `(${d}) ${n.slice(0, 4)}-${n.slice(4)}`;
      return `(${d}) ${n}`;
    },
    colToIndex: (col) => {
      col = col.toUpperCase();
      if (!/^[A-Z]{1,3}$/.test(col)) throw new Error('Coluna inválida');
      return col.split('').reduce((a, c) => a * 26 + (c.charCodeAt(0) - 64), 0) - 1;
    },
    toColName: (n) => {
      if (n <= 0) return 'A';
      let s = '';
      while (n > 0) {
        const t = (n - 1) % 26;
        s = String.fromCharCode(65 + t) + s;
        n = Math.floor((n - t - 1) / 26);
      }
      return s;
    }
  };

  // LOCAL STORAGE
  const L = {
    get: (k, d = null) => {
      try { return JSON.parse(localStorage.getItem('planilista_' + k)) || d; }
      catch { return d; }
    },
    set: (k, v) => localStorage.setItem('planilista_' + k, JSON.stringify(v))
  };

  // STATE
  const S = {
    file: null,
    wb: null,
    sheet: null,
    data: [],
    detected: new Map(),
    colorMap: L.get('colorMap', {}),
    excluded: L.get('excluded', {}),
    phonePairs: L.get('phonePairs', []),
    applyMask: L.get('applyMask', true),
    theme: L.get('theme', 'dark'),
    autoExclude: L.get('autoExclude', true),
    stats: {},
    processing: false
  };

  // DOM
  const D = {
    status: U.qs('#status'),
    statusText: U.qs('#statusText'),
    drop: U.qs('#drop'),
    fileInput: U.qs('#fileInput'),
    fileName: U.qs('#fileName'),
    panel: U.qs('#panel'),
    tabs: U.qs('#tabs'),
    colInput: U.qs('#colInput'),
    dedup: U.qs('#dedup'),
    autoExcludeChk: U.qs('#autoExcludeChk'),
    processBtn: U.qs('#processBtn'),
    mapBtn: U.qs('#mapBtn'),
    mapCount: U.qs('#mapCount'),
    output: U.qs('#output'),
    copyBtn: U.qs('#copyBtn'),
    clearBtn: U.qs('#clearBtn'),
    csvBtn: U.qs('#csvBtn'),
    txtBtn: U.qs('#txtBtn'),
    colorList: U.qs('#colorList'),
    excludeInfo: U.qs('#excludeInfo'),
    excludedList: U.qs('#excludedList'),
    unmappedBadge: U.qs('#unmappedBadge'),
    themeBtn: U.qs('#themeBtn'),
    phoneBtn: U.qs('#phoneBtn'),
    mapModal: U.qs('#mapModal'),
    closeMap: U.qs('#closeMap'),
    mapBody: U.qs('#mapBody'),
    clearMap: U.qs('#clearMap'),
    saveMap: U.qs('#saveMap'),
    expCfgBtn: U.qs('#expCfgBtn'),
    impCfgBtn: U.qs('#impCfgBtn'),
    impCfgFile: U.qs('#impCfgFile'),
    statsPanel: U.qs('#statsPanel'),
    phoneModal: U.qs('#phoneModal'),
    closePhone: U.qs('#closePhone'),
    pairs: U.qs('#pairs'),
    addPair: U.qs('#addPair'),
    maskChk: U.qs('#maskChk'),
    applyPhone: U.qs('#applyPhone')
  };

  // INIT
  setTheme(S.theme);
  U.renderIcons();
  D.dedup.checked = L.get('dedup', false);
  D.colInput.value = L.get('column', 'A');
  D.autoExcludeChk.checked = S.autoExclude;

  // VERSÃO CORRIGIDA (substitua a função em script.js)
function setTheme(theme) {
  // Define o tema no HTML
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('theme', theme);

  // 1. Encontra o botão do tema
  const themeBtn = document.getElementById('themeBtn');
  if (!themeBtn) return; // Segurança

  // 2. Remove o ícone SVG antigo de dentro do botão
  const oldIcon = themeBtn.querySelector('svg');
  if (oldIcon) {
    oldIcon.remove();
  }

  // 3. Cria um NOVO elemento <i> para o novo ícone
  const newIcon = document.createElement('i');
  const iconName = theme === 'dark' ? 'sun' : 'moon';
  newIcon.setAttribute('data-lucide', iconName);
  
  // 4. Adiciona o novo <i> ao botão
  themeBtn.appendChild(newIcon);

  // 5. Pede para a biblioteca Lucide converter o novo <i> em um <svg>
  lucide.createIcons();
}

  function status(m, t = '') {
    D.statusText.textContent = m;
    D.status.className = 'status' + (t ? ' ' + t : '');
    const icon = D.status.querySelector('i');
    if (icon) {
      icon.setAttribute('data-lucide', t === 'error' ? 'x-circle' : t === 'warn' ? 'alert-triangle' : t === 'ok' ? 'check-circle' : 'info');
      U.renderIcons();
    }
  }

  // FILE HANDLING
  D.drop.addEventListener('click', () => D.fileInput.click());
  D.fileInput.addEventListener('change', e => e.target.files[0] && loadFile(e.target.files[0]));

  ['dragover', 'dragenter'].forEach(ev => D.drop.addEventListener(ev, e => {
    e.preventDefault();
    D.drop.classList.add('active');
  }));

  ['dragleave', 'drop'].forEach(ev => D.drop.addEventListener(ev, e => {
    e.preventDefault();
    D.drop.classList.remove('active');
    if (e.type === 'drop' && e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]);
  }));

  async function loadFile(file) {
    if (file.size > 50 * 1024 * 1024) {
      status('Arquivo muito grande. Máx: 50MB', 'error');
      return;
    }

    try {
      status('Processando arquivo...', 'warn');

      const ab = await file.arrayBuffer();

      S.wb = XLSX.read(ab, {
        type: 'array',
        cellStyles: true,
        raw: false
      });

      S.file = file;
      D.fileName.textContent = file.name;
      D.panel.style.display = 'grid';

      buildTabs();

      const lastSheet = L.get('activeSheet');
      const sheetIndex = S.wb.SheetNames.indexOf(lastSheet);
      const activeIndex = sheetIndex > -1 ? sheetIndex : 0;
      S.sheet = S.wb.SheetNames[activeIndex];

      U.qsa('.tab').forEach((t, i) => t.classList.toggle('active', i === activeIndex));

      loadAOA();

      status(`"${U.esc(file.name)}" carregado.`, 'ok');
      calculateStats();
      process();
    } catch (e) {
      console.error(e);
      status('Erro: ' + e.message, 'error');
    }
  }

  function buildTabs() {
    D.tabs.innerHTML = '';
    S.wb.SheetNames.forEach((n, i) => {
      const b = document.createElement('button');
      b.className = 'tab';
      if (i === 0) b.classList.add('active');
      b.textContent = n;
      b.addEventListener('click', () => selectTab(i));
      D.tabs.appendChild(b);
    });
  }

  function selectTab(index) {
    U.qsa('.tab').forEach((t, i) => t.classList.toggle('active', i === index));
    S.sheet = S.wb.SheetNames[index];
    L.set('activeSheet', S.sheet);
    loadAOA();
    calculateStats();
    process();
  }

  function loadAOA() {
    const ws = S.wb.Sheets[S.sheet];
    S.data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  }

  // DETECÇÃO DE COR OTIMIZADA
  function getCellColor(cell) {
    if (!cell?.s) return null;

    // Caminho direto baseado no que descobrimos
    if (cell.s.fgColor?.rgb) {
      const rgb = cell.s.fgColor.rgb;
      return '#' + rgb.toUpperCase();
    }

    // Fallback para outros possíveis caminhos
    if (cell.s.bgColor?.rgb) {
      const rgb = cell.s.bgColor.rgb;
      return '#' + rgb.toUpperCase();
    }

    // Paleta indexada (caso seja XLS antigo)
    const XLS_PALETTE = {
      0: '#000000', 1: '#FFFFFF', 2: '#FF0000', 3: '#00FF00',
      4: '#0000FF', 5: '#FFFF00', 6: '#FF00FF', 7: '#00FFFF',
      16: '#800000', 17: '#008000', 18: '#000080', 19: '#808000',
      20: '#800080', 21: '#008080', 22: '#C0C0C0', 23: '#808080'
    };

    if (cell.s.fgColor?.indexed !== undefined) {
      return XLS_PALETTE[cell.s.fgColor.indexed];
    }

    return null;
  }

  function shouldExclude(name) {
    if (!name) return false;
    if (S.excluded[name]) return true;

    if (S.autoExclude) {
      const words = ['ok', 'sim', 'yes', 'confirmado', 'pronto', 'feito', 'concluído', 'concluido'];
      const lowerName = name.toLowerCase();
      return words.some(w => new RegExp(`\\b${w}\\b`).test(lowerName));
    }
    return false;
  }

  function calculateStats() {
    if (!S.data.length) return;

    const stats = {
      totalRows: S.data.length - 1,
      totalCols: S.data[0]?.length || 0,
      coloredCells: 0,
      uniqueNumbers: new Set(),
      categories: {}
    };

    const ws = S.wb.Sheets[S.sheet];
    if (!ws || !ws['!ref']) return;

    const range = XLSX.utils.decode_range(ws['!ref']);

    for (let r = range.s.r; r <= Math.min(range.e.r, 10000); r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (cell) {
          const color = getCellColor(cell);
          if (color) {
            stats.coloredCells++;
            const cat = S.colorMap[color];
            if (cat) {
              if (!stats.categories[cat]) stats.categories[cat] = 0;
              stats.categories[cat]++;
            }
          }
          const val = String(cell.w ?? cell.v ?? '');
          const num = U.digits(val);
          if (num) stats.uniqueNumbers.add(num);
        }
      }
    }

    S.stats = stats;
    renderStatsPanel();
  }

  function renderStatsPanel() {
    if (!S.stats || !Object.keys(S.stats).length) return;

    D.statsPanel.innerHTML = `
      <div class="stat-card">
        <div class="stat-value">${S.stats.totalRows.toLocaleString('pt-BR')}</div>
        <div class="stat-label">Linhas</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${S.stats.coloredCells.toLocaleString('pt-BR')}</div>
        <div class="stat-label">Células coloridas</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${S.stats.uniqueNumbers.size.toLocaleString('pt-BR')}</div>
        <div class="stat-label">Números únicos</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${Object.keys(S.stats.categories).length}</div>
        <div class="stat-label">Categorias</div>
      </div>
    `;
    D.statsPanel.style.display = 'grid';
  }

  async function process() {
    if (!S.wb || !S.sheet || S.processing) return;

    S.processing = true;

    const col = (D.colInput.value || 'A').toUpperCase();
    let colIndex;
    try {
      colIndex = U.colToIndex(col);
    } catch (e) {
      status(e.message, 'error');
      S.processing = false;
      return;
    }

    S.detected.clear();
    const grouped = new Map(), excludedG = new Map(), unmapped = new Set();
    const dedup = !!D.dedup.checked;

    const ws = S.wb.Sheets[S.sheet];
    if (!ws || !ws['!ref']) {
      status('Planilha sem dados', 'warn');
      S.processing = false;
      return;
    }

    const range = XLSX.utils.decode_range(ws['!ref']);
    const maxRows = Math.min(range.e.r, 10000);

    for (let r = range.s.r; r <= maxRows; r++) {
      if (r % 1000 === 0) await new Promise(r => setTimeout(r, 1));

      const addr = XLSX.utils.encode_cell({ r, c: colIndex });
      const cell = ws[addr];
      if (!cell) continue;

      const color = getCellColor(cell);
      if (color) S.detected.set(color, (S.detected.get(color) || 0) + 1);

      const val = String(cell.w ?? cell.v ?? '');
      const num = U.digits(val);
      if (!num || !color) continue;

      const cat = S.colorMap[color];
      if (!cat) { unmapped.add(color); continue; }

      const tgt = shouldExclude(cat) ? excludedG : grouped;
      if (!tgt.has(cat)) tgt.set(cat, dedup ? new Set() : []);
      if (dedup) tgt.get(cat).add(num);
      else tgt.get(cat).push(num);
    }

    // Gerar output
    let out = '';
    const sortedCategories = [...grouped.keys()].sort((a, b) => a.localeCompare(b));
    for (const cat of sortedCategories) {
      const list = grouped.get(cat);
      const items = (Array.isArray(list) ? list : Array.from(list))
        .sort((a, b) => {
          const na = BigInt(a), nb = BigInt(b);
          return na < nb ? -1 : na > nb ? 1 : 0;
        });
      out += `${cat} (${items.length}):\n${items.join('\n')}\n\n`;
    }

    D.output.value = out.trim();
    D.copyBtn.disabled = !D.output.value;

    // Mostrar exclusões
    const excl = [...excludedG.keys()];
    if (excl.length) {
      D.excludeInfo.style.display = 'flex';
      D.excludedList.textContent = excl.join(', ');
    } else {
      D.excludeInfo.style.display = 'none';
    }

    renderColorList();

    const unmappedCount = [...S.detected.keys()].filter(hex => !S.colorMap[hex]).length;
    if (unmappedCount > 0) {
      D.mapCount.style.display = 'inline-flex';
      D.mapCount.textContent = unmappedCount;
    } else {
      D.mapCount.style.display = 'none';
    }

    D.unmappedBadge.textContent = unmapped.size ? `${unmapped.size} não mapeada(s)` : '';

    const msg = unmapped.size ? `${unmapped.size} cor(es) sem mapeamento.` : 'Processamento concluído.';
    status(msg, 'ok');

    S.processing = false;
  }

  function renderColorList() {
    D.colorList.innerHTML = '';
    if (S.detected.size === 0) {
      D.colorList.innerHTML = '<div class="chip">Nenhuma cor detectada</div>';
      return;
    }
    S.detected.forEach((count, hex) => {
      const name = S.colorMap[hex] || 'Não mapeado';
      const el = document.createElement('div');
      el.className = 'badge-color';
      el.innerHTML = `
        <span class="sw" style="background:${hex}"></span>
        <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${U.esc(name)}</span>
        <span class="chip">${count}</span>
      `;
      D.colorList.appendChild(el);
    });
  }

  // CONTROLS
  D.processBtn.addEventListener('click', process);
  D.colInput.addEventListener('input', () => { setTimeout(() => { L.set('column', D.colInput.value); process(); }, 500); });
  D.dedup.addEventListener('change', () => { L.set('dedup', D.dedup.checked); process(); });
  D.autoExcludeChk.addEventListener('change', () => { S.autoExclude = D.autoExcludeChk.checked; L.set('autoExclude', S.autoExclude); process(); });
  D.themeBtn.addEventListener('click', () => setTheme(S.theme === 'dark' ? 'light' : 'dark'));

  // RESULT ACTIONS
  D.copyBtn.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(D.output.value || '');
      status('Copiado!', 'ok');
    } catch {
      D.output.select();
      document.execCommand('copy');
      status('Copiado.', 'ok');
    }
  });

  D.clearBtn.addEventListener('click', () => {
    D.output.value = '';
    D.copyBtn.disabled = true;
    status('Resultado limpo.');
  });

  D.csvBtn.addEventListener('click', exportCSV);
  D.txtBtn.addEventListener('click', downloadTXT);

  // MAP MODAL
  D.mapBtn.addEventListener('click', () => { renderMapBody(); toggleModal(D.mapModal, true); });
  D.closeMap.addEventListener('click', () => toggleModal(D.mapModal, false));
  D.mapModal.addEventListener('click', e => { if (e.target === D.mapModal) toggleModal(D.mapModal, false); });
  D.clearMap.addEventListener('click', () => {
    if (confirm('Limpar todos os mapeamentos?')) {
      S.colorMap = {};
      S.excluded = {};
      L.set('colorMap', S.colorMap);
      L.set('excluded', S.excluded);
      renderMapBody();
      process();
    }
  });
  D.saveMap.addEventListener('click', saveMappings);

  // PHONE MODAL
  D.phoneBtn.addEventListener('click', openPhoneModal);
  D.closePhone.addEventListener('click', () => toggleModal(D.phoneModal, false));
  D.phoneModal.addEventListener('click', e => { if (e.target === D.phoneModal) toggleModal(D.phoneModal, false); });
  D.addPair.addEventListener('click', () => D.pairs.appendChild(makePair()));
  D.applyPhone.addEventListener('click', applyPhoneMerge);
  D.maskChk.addEventListener('change', () => { S.applyMask = !!D.maskChk.checked; L.set('applyMask', S.applyMask); });

  // CONFIG
  D.expCfgBtn.addEventListener('click', exportConfig);
  D.impCfgBtn.addEventListener('click', () => D.impCfgFile.click());
  D.impCfgFile.addEventListener('change', e => {
    const f = e.target.files?.[0];
    if (f) importConfig(f);
    e.target.value = '';
  });

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      U.qsa('.modal.show').forEach(m => toggleModal(m, false));
    }
  });

  function toggleModal(m, show) {
    m.classList.toggle('show', !!show);
    if (show) U.renderIcons();
  }

  function renderMapBody() {
    const frag = document.createDocumentFragment();
    const list = Array.from(S.detected.entries()).sort((a, b) => (b[1] - a[1]));
    if (!list.length) {
      D.mapBody.innerHTML = '<div class="chip">Nenhuma cor a configurar</div>';
      return;
    }
    list.forEach(([hex]) => {
      const row = document.createElement('div');
      row.className = 'map-row';
      const name = S.colorMap[hex] || '';
      const excl = S.excluded[name] ? 'checked' : '';
      row.innerHTML = `
        <span class="sw" style="background:${hex}"></span>
        <span style="font-family:monospace;font-size:12px;">${hex}</span>
        <input class="input" data-color="${hex}" value="${U.esc(name)}" placeholder="Nome da categoria" />
        <label class="hint" style="display:flex;align-items:center;gap:6px;">
          <input type="checkbox" data-exclude="${hex}" ${excl} /> Excluir
        </label>
      `;
      frag.appendChild(row);
    });
    D.mapBody.innerHTML = '';
    D.mapBody.appendChild(frag);
  }

  function saveMappings() {
    const inputs = D.mapBody.querySelectorAll('input[data-color]');
    const excls = D.mapBody.querySelectorAll('input[data-exclude]');

    const newExcl = { ...S.excluded };

    inputs.forEach((inp, i) => {
      const hex = inp.dataset.color, name = inp.value.trim();
      if (name) S.colorMap[hex] = name;
      else delete S.colorMap[hex];

      const chk = excls[i];
      if (chk && name) {
        if (chk.checked) newExcl[name] = true;
        else delete newExcl[name];
      }
    });

    S.excluded = newExcl;

    L.set('colorMap', S.colorMap);
    L.set('excluded', S.excluded);
    toggleModal(D.mapModal, false);
    process();
  }

  // PHONE FUNCTIONS
  function guessIndex(headers, re, fb) {
    const idx = headers.findIndex(h => re.test(U.norm(h)));
    return idx >= 0 ? idx : fb;
  }

  function makePair(dddIndex, phoneIndex) {
    const p = document.createElement('div');
    p.style = 'display:flex;gap:8px;align-items:center';
    const headers = S.data[0] || [];
    const mk = (sel) => headers.map((h, i) => `<option value="${i}" ${i === sel ? 'selected' : ''}>${U.esc(h || ('Coluna ' + U.toColName(i + 1)))}</option>`).join('');
    const dSel = document.createElement('select');
    dSel.className = 'input';
    dSel.innerHTML = mk(dddIndex ?? guessIndex(headers, /(ddd|area)/i, 0));
    const tSel = document.createElement('select');
    tSel.className = 'input';
    tSel.innerHTML = mk(phoneIndex ?? guessIndex(headers, /(tel|fone|cel)/i, 1));
    const del = document.createElement('button');
    del.className = 'btn danger';
    del.innerHTML = '<i data-lucide="trash-2"></i>';
    del.onclick = () => p.remove();
    p.appendChild(dSel);
    p.appendChild(tSel);
    p.appendChild(del);
    U.renderIcons();
    return p;
  }

  function openPhoneModal() {
    if (!S.data.length) {
      status('Carregue uma planilha.', 'warn');
      return;
    }
    D.pairs.innerHTML = '';
    if (S.phonePairs.length) {
      S.phonePairs.forEach(pr => D.pairs.appendChild(makePair(pr.dddIndex, pr.phoneIndex)));
    } else {
      D.pairs.appendChild(makePair());
    }
    D.maskChk.checked = !!S.applyMask;
    toggleModal(D.phoneModal, true);
  }

  function applyPhoneMerge() {
    if (!S.data.length) {
      toggleModal(D.phoneModal, false);
      return;
    }

    const pairs = [];
    D.pairs.querySelectorAll('div').forEach(p => {
      const d = parseInt(p.children[0].value, 10);
      const t = parseInt(p.children[1].value, 10);
      if (!isNaN(d) && !isNaN(t) && d !== t) {
        pairs.push({ dddIndex: d, phoneIndex: t });
      }
    });

    if (!pairs.length) {
      status('Nenhum par válido.', 'warn');
      return;
    }

    S.phonePairs = pairs;
    L.set('phonePairs', S.phonePairs);
    S.applyMask = !!D.maskChk.checked;
    L.set('applyMask', S.applyMask);

    toggleModal(D.phoneModal, false);
    status('Configuração de telefones salva.', 'ok');
  }

  function downloadTXT() {
    const text = (D.output.value || '').trim();
    if (!text) {
      status('Nada para baixar.', 'warn');
      return;
    }
    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const a = document.createElement('a');
    const base = (S.file?.name || 'planilha').replace(/\.[^.]+$/, '');
    a.href = URL.createObjectURL(blob);
    a.download = `${base}_lista.txt`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    status('TXT exportado.', 'ok');
  }

  // FUNÇÃO EXPORTCSV ATUALIZADA - FORMATO EXATO
  function exportCSV() {
    if (!S.data.length) {
      status('Carregue uma planilha.', 'warn');
      return;
    }

    const headers = S.data[0] || [];

    // Função auxiliar para encontrar colunas
    const findExact = (searchTerms) => {
      for (let i = 0; i < headers.length; i++) {
        const normalized = U.norm(headers[i]);
        for (const term of searchTerms) {
          if (normalized === U.norm(term)) return i;
        }
      }
      // Segunda tentativa - busca parcial
      for (let i = 0; i < headers.length; i++) {
        const normalized = U.norm(headers[i]);
        for (const term of searchTerms) {
          if (normalized.includes(U.norm(term))) return i;
        }
      }
      return -1;
    };

    // Mapeamento dos índices principais
    const idx = {
      cpf: findExact(['cpf', 'cpf/cnpj', 'cnpj', 'documento']),
      codigo: findExact(['codigo', 'código', 'codigo do cliente', 'código do cliente', 'cod cliente', 'id']),
      nome: findExact(['nome', 'nome do cliente', 'cliente', 'razao social', 'razão social']),
      endereco: findExact(['endereco', 'endereço', 'logradouro', 'rua', 'avenida']),
      bairro: findExact(['bairro']),
      cidade: findExact(['cidade', 'municipio', 'município']),
      debitoTotal: findExact(['debito total', 'débito total', 'valor total', 'total']),
      debitoPendente: findExact(['debito pendente', 'débito pendente', 'valor pendente', 'pendente', 'saldo']),
      vencimento: findExact(['vencimento', 'data vencimento', 'data de vencimento', 'dt vencimento'])
    };

    // Identificar pares de DDD + Telefone
    const phonePairs = [];
    const dddColumns = [];
    const phoneColumns = [];

    headers.forEach((header, i) => {
      const h = U.norm(header);

      // É uma coluna de DDD?
      if (h.includes('ddd')) {
        dddColumns.push(i);
      }
      // É uma coluna de telefone/celular?
      else if (h.match(/(telefone|celular|tel|cel|fone|whats|contato)/) && !h.includes('ddd')) {
        phoneColumns.push(i);
      }
    });

    // Criar pares DDD -> Telefone baseado na proximidade
    dddColumns.forEach(dddIdx => {
      let closestPhone = -1;
      let minDistance = Infinity;

      phoneColumns.forEach(phoneIdx => {
        if (phoneIdx > dddIdx) {
          const distance = phoneIdx - dddIdx;
          if (distance < minDistance) {
            minDistance = distance;
            closestPhone = phoneIdx;
          }
        }
      });

      if (closestPhone !== -1) {
        phonePairs.push({ ddd: dddIdx, phone: closestPhone });
        const phoneIndex = phoneColumns.indexOf(closestPhone);
        if (phoneIndex > -1) phoneColumns.splice(phoneIndex, 1);
      }
    });

    // Headers de saída
    const OUT = [
      'CPF',
      'Codigo do Cliente',
      'Nome do Cliente',
      'Endereco',
      'Bairro',
      'Cidade',
      'Debito total',
      'Debito pendente',
      'Vencimento',
      'Telefone1',
      'Telefone2',
      'Telefone3',
      'Telefone4',
      'Telefone5'
    ];

    const rows = [];

    for (let r = 1; r < S.data.length; r++) {
      const row = S.data[r];
      if (!row || row.length === 0) continue;

      const v = (i) => i > -1 && i < row.length ? String(row[i] || '') : '';

      // Processar telefones
      const phones = [];
      const seen = new Set();

      // Processar pares DDD + Telefone
      phonePairs.forEach(pair => {
        const dddVal = v(pair.ddd);
        const phoneVal = v(pair.phone);

        if (!phoneVal || phoneVal === '0') return;

        const dddDigits = U.digits(dddVal) || '83';
        const phoneDigits = U.digits(phoneVal);

        if (phoneDigits && phoneDigits.length >= 8) {
          const fullNumber = dddDigits + phoneDigits;

          if (!seen.has(fullNumber)) {
            seen.add(fullNumber);
            phones.push(fullNumber);
          }
        }
      });

      // Processar colunas de telefone isoladas
      phoneColumns.forEach(colIdx => {
        const val = v(colIdx);
        if (!val) return;

        let digits = U.digits(val);
        if (!digits) return;

        let fullNumber = '';

        if (digits.length === 11 || digits.length === 10) {
          fullNumber = digits;
        } else if (digits.length === 9 || digits.length === 8) {
          fullNumber = '83' + digits;
        } else {
          return;
        }

        if (!seen.has(fullNumber) && fullNumber.length >= 10) {
          seen.add(fullNumber);
          phones.push(fullNumber);
        }
      });

      // Garantir que temos 5 campos de telefone
      while (phones.length < 5) phones.push('');

      // Pegar valores corretos
      const codigo = v(idx.codigo);
      let nome = v(idx.nome);

      // Se nome está vazio ou igual ao código, tentar encontrar o nome correto
      if (!nome || nome === codigo) {
        for (let i = idx.codigo + 1; i < row.length && i < idx.codigo + 3; i++) {
          const possibleName = v(i);
          if (possibleName && possibleName !== codigo && !(/^\d+$/.test(possibleName))) {
            nome = possibleName;
            break;
          }
        }
      }

      // Processar CPF - remover formatação
      let cpf = U.digits(v(idx.cpf));

      // Processar valores decimais
      let debitoTotal = v(idx.debitoTotal);
      let debitoPendente = v(idx.debitoPendente);

      // Montar linha
      rows.push([
        cpf,
        codigo,
        nome,
        v(idx.endereco),
        v(idx.bairro),
        v(idx.cidade),
        debitoTotal,
        debitoPendente,
        v(idx.vencimento),
        phones[0],
        phones[1],
        phones[2],
        phones[3],
        phones[4]
      ]);
    }

    // Gerar CSV com PONTO E VÍRGULA
    let csv = OUT.join(';') + '\r\n';
    rows.forEach(r => {
      csv += r.map(cell => {
        if (!cell) return '';
        const str = String(cell);
        if (str.includes(';') || str.includes('"') || str.includes('\n')) {
          return '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
      }).join(';') + '\r\n';
    });

    // Download com BOM para UTF-8
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    const base = (S.file?.name || 'planilha').replace(/\.[^.]+$/, '');
    a.download = `${base}_cobranca.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();

    status('CSV exportado com sucesso!', 'ok');
  }

  function exportConfig() {
    const cfg = {
      version: '2.5',
      theme: S.theme,
      column: (D.colInput.value || 'A').toUpperCase(),
      dedup: !!D.dedup.checked,
      autoExclude: S.autoExclude,
      colorMap: S.colorMap,
      excluded: S.excluded,
      phonePairs: S.phonePairs,
      applyMask: S.applyMask
    };

    const blob = new Blob([JSON.stringify(cfg, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `planilista_config.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    status('Configuração exportada.', 'ok');
  }

  function importConfig(file) {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const cfg = JSON.parse(reader.result || '{}');

        if (cfg.theme) setTheme(cfg.theme);
        if (cfg.column) D.colInput.value = cfg.column;
        if (typeof cfg.dedup === 'boolean') D.dedup.checked = cfg.dedup;
        if (typeof cfg.autoExclude === 'boolean') {
          D.autoExcludeChk.checked = cfg.autoExclude;
          S.autoExclude = cfg.autoExclude;
        }
        if (cfg.colorMap) S.colorMap = cfg.colorMap;
        if (cfg.excluded) S.excluded = cfg.excluded;
        if (Array.isArray(cfg.phonePairs)) S.phonePairs = cfg.phonePairs;
        if (typeof cfg.applyMask === 'boolean') S.applyMask = cfg.applyMask;

        L.set('column', D.colInput.value);
        L.set('dedup', D.dedup.checked);
        L.set('autoExclude', S.autoExclude);
        L.set('colorMap', S.colorMap);
        L.set('excluded', S.excluded);
        L.set('phonePairs', S.phonePairs);
        L.set('applyMask', S.applyMask);

        renderMapBody();
        process();
        status('Configuração importada.', 'ok');
        toggleModal(D.mapModal, false);
      } catch (e) {
        console.error(e);
        status('Configuração inválida.', 'error');
      }
    };
    reader.readAsText(file);
  }

  status('Planilista v2.5 - Pronto!', 'ok');
})();
