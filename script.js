(() => {
  const $ = (s, el = document) => el.querySelector(s);

  // ===== Elements =====
  const els = {
    grid: $('#grid'),
    empty: $('#empty'),
    q: $('#q'),
    clearQ: $('#clearQ'),
    countInfo: $('#countInfo'),
    chips: $('#activeChips'),

    dialog: $('#filtersDialog'),
    btnFilters: $('#btnFilters'),
    clearFilters: $('#clearFilters'),
    fUnidade: $('#fUnidade'),
    fArea: $('#fArea'),

    detailsDialog: $('#detailsDialog'),
    detailsTitle: $('#detailsTitle'),
    dColab: $('#dColab'),
    dArea: $('#dArea'),
    dUnidade: $('#dUnidade'),
    dTema: $('#dTema'),
    dDesc: $('#dDesc'),

    // viewer (opcional no HTML)
    dViewer: $('#dViewer'),
    dViewerLabel: $('#dViewerLabel'),
    dOpenNewTab: $('#dOpenNewTab'),
    dFrame: $('#dFrame'),
    dImg: $('#dImg'),
  };

  // ===== State =====
  const state = {
    rows: [],
    filtered: [],
    filters: { q: '', unidade: 'Todos', area: 'Todos' }
  };

  // ===== Utils =====
  const norm = (v) => (v ?? '').toString().trim();
  const normHeader = (h) => norm(h).replace(/\s+/g, ' ').replace(/\n/g, ' ').toLowerCase();
  const pad3 = (n) => String(n).padStart(3, '0');

  const escapeHtml = (str) =>
    (str ?? '').toString()
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#039;');

  const show = (el) => el && el.classList.remove('hidden');
  const hide = (el) => el && el.classList.add('hidden');

  function pickCol(row, candidates){
    for (const c of candidates) {
      if (row[c] != null && norm(row[c]) !== '') return norm(row[c]);
    }
    return '';
  }

  function makeTitleFromText(text){
    const t = norm(text);
    if (!t) return 'Iniciativa';
    const first = t.split(/\n|\r|\./)[0].trim();
    if (first.length >= 10 && first.length <= 64) return first;
    const words = t.split(/\s+/);
    return words.slice(0, 7).join(' ') + (words.length > 7 ? '…' : '');
  }

  function inferTema(text){
    const t = norm(text).toLowerCase();
    if (t.includes('power automate') || t.includes('automat')) return 'Automação / IA';
    if (t.includes('dashboard') || t.includes('bi')) return 'Dados / IA';
    if (t.includes('padron') || t.includes('process')) return 'Processos / IA';
    if (t.includes('trein') || t.includes('capacit')) return 'Pessoas / IA';
    return 'IA / Inovação';
  }

  function isImage(url){ return /\.(png|jpg|jpeg|gif|webp|svg)(\?.*)?$/i.test(url); }
  function isPdf(url){ return /\.pdf(\?.*)?$/i.test(url); }

  function driveFileId(url){
    const m = url.match(/drive\.google\.com\/file\/d\/([^/]+)/i);
    return m ? m[1] : '';
  }

  function toDrivePreviewUrl(url){
    const id = driveFileId(url);
    return id ? `https://drive.google.com/file/d/${id}/preview` : url;
  }

  function classifyAttachment(raw){
    const url = norm(raw);
    if (!url) return { type: 'none', url: '' };

    const u = url.toLowerCase();
    if (u.includes('drive.google.com/file/d/')) {
      return { type: 'iframe', url: toDrivePreviewUrl(url), openUrl: url };
    }
    if (isPdf(url)) return { type: 'pdf', url };
    if (isImage(url)) return { type: 'image', url };
    return { type: 'link', url };
  }

  function resetViewer(){
    hide(els.dViewer);
    hide(els.dFrame);
    hide(els.dImg);
    if (els.dFrame) els.dFrame.removeAttribute('src');
    if (els.dImg) els.dImg.removeAttribute('src');
    if (els.dOpenNewTab) els.dOpenNewTab.setAttribute('href', '#');
    if (els.dViewerLabel) els.dViewerLabel.textContent = 'Anexo';
  }

  // ===== Render =====
  function renderChips(){
    const chips = [];
    if (state.filters.unidade !== 'Todos') chips.push(`Filial: ${state.filters.unidade}`);
    if (state.filters.area !== 'Todos') chips.push(`Área: ${state.filters.area}`);
    if (state.filters.q) chips.push(`Busca: "${state.filters.q}"`);
    els.chips.innerHTML = chips.map(c => `<span class="chip">${escapeHtml(c)}</span>`).join('');
  }

  function render(){
    els.grid.innerHTML = '';
    els.countInfo.textContent = `Iniciativas publicadas: ${state.filtered.length}`;

    if (!state.filtered.length){
      els.empty.classList.remove('hidden');
      return;
    }
    els.empty.classList.add('hidden');

    const frag = document.createDocumentFragment();
    state.filtered.forEach(item => frag.appendChild(cardEl(item)));
    els.grid.appendChild(frag);
  }

  function cardEl(item){
    const card = document.createElement('article');
    card.className = 'card';

    // ✅ aqui removi a seta duplicada (tiramos o <span class="arrow">→</span>)
    card.innerHTML = `
      <div class="card__id">Iniciativa #${pad3(item.id)}</div>
      <h3 class="card__title">${escapeHtml(item.titulo)}</h3>
      <p class="card__desc"><b>Descrição:</b> ${escapeHtml(item.descricao)}</p>

      <div class="meta">
        <div><b>Colaborador:</b> <span>${escapeHtml(item.colaborador || '-')}</span></div>
        <div><b>Área:</b> <span>${escapeHtml(item.area || '-')}</span></div>
        <div><b>Filial:</b> <span>${escapeHtml(item.unidade || '-')}</span></div>
        <div><b>Tema:</b> <span>${escapeHtml(item.tema || 'IA / Inovação')}</span></div>
      </div>

      <div class="card__actions">
        <a class="link" href="#" data-action="details">Ver detalhes →</a>
      </div>
    `;

    const detailsLink = card.querySelector('[data-action="details"]');

    detailsLink.addEventListener('click', (e) => {
      e.preventDefault();
      openDetails(item);
    });

    // clique no card abre também (exceto no link)
    card.addEventListener('click', (e) => {
      if (e.target.closest('a')) return;
      openDetails(item);
    });

    return card;
  }

  // ===== Details =====
  function openDetails(item){
    resetViewer();

    els.detailsTitle.textContent = `Iniciativa #${pad3(item.id)} — ${item.titulo}`;
    els.dColab.textContent = item.colaborador || '-';
    els.dArea.textContent = item.area || '-';
    els.dUnidade.textContent = item.unidade || '-';
    els.dTema.textContent = item.tema || 'IA / Inovação';
    els.dDesc.textContent = item.descricao || '-';

    const att = classifyAttachment(item.anexo);

    if (att.type !== 'none' && els.dViewer){
      show(els.dViewer);
      if (els.dOpenNewTab) els.dOpenNewTab.href = att.openUrl || att.url;

      if (att.type === 'image' && els.dImg){
        if (els.dViewerLabel) els.dViewerLabel.textContent = 'Imagem';
        show(els.dImg);
        els.dImg.src = att.url;
      } else if ((att.type === 'pdf' || att.type === 'iframe') && els.dFrame){
        if (els.dViewerLabel) els.dViewerLabel.textContent = att.type === 'pdf' ? 'PDF' : 'Anexo (Drive)';
        show(els.dFrame);
        els.dFrame.src = att.url;
      } else {
        if (els.dViewerLabel) els.dViewerLabel.textContent = 'Link';
      }
    }

    els.detailsDialog.showModal();
  }

  // ===== Filters =====
  function applyFilters(){
    const q = state.filters.q.toLowerCase();

    state.filtered = state.rows.filter(r => {
      const okUnidade = state.filters.unidade === 'Todos' || r.unidade === state.filters.unidade;
      const okArea = state.filters.area === 'Todos' || r.area === state.filters.area;
      const hay = `${r.id} ${r.titulo} ${r.descricao} ${r.colaborador} ${r.area} ${r.unidade} ${r.tema}`.toLowerCase();
      const okQ = !q || hay.includes(q);
      return okUnidade && okArea && okQ;
    });

    renderChips();
    render();
  }

  function fillSelect(selectEl, values){
    selectEl.innerHTML = '';
    const optAll = document.createElement('option');
    optAll.value = 'Todos';
    optAll.textContent = 'Todos';
    selectEl.appendChild(optAll);

    values.forEach(v => {
      const o = document.createElement('option');
      o.value = v;
      o.textContent = v;
      selectEl.appendChild(o);
    });
  }

  // ===== Load XLSX =====
  async function loadXlsx(){
    const url = 'iniciativas.xlsx';

    const res = await fetch(url);
    if (!res.ok) throw new Error(`Não consegui carregar ${url} (${res.status})`);
    const buf = await res.arrayBuffer();

    const wb = XLSX.read(buf, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: '' });

    const normalized = raw.map((row) => {
      const out = {};
      for (const k of Object.keys(row)) out[normHeader(k)] = row[k];
      return out;
    });

    const COL_COLAB = ['nome colaborador(a)', 'colaborador', 'nome', 'nome do colaborador'];
    const COL_UNIDADE = ['unidade', 'filial'];
    const COL_AREA = ['area', 'área', 'setor'];
    const COL_DESC = [
      'descreva a solução realizada pelo seu(sua) colaborador(a) (em 2025)',
      'descricao', 'descrição', 'iniciativa', 'solucao', 'solução'
    ];
    const COL_ANEXO = ['anexo', 'pdf', 'arquivo', 'link', 'detalhes', 'detalhes_url'];

    const rows = normalized.map((r, idx) => {
      const descricao = pickCol(r, COL_DESC);
      return {
        id: idx + 1,
        colaborador: pickCol(r, COL_COLAB),
        unidade: pickCol(r, COL_UNIDADE),
        area: pickCol(r, COL_AREA),
        titulo: makeTitleFromText(descricao),
        descricao: norm(descricao),
        tema: inferTema(descricao),
        anexo: norm(pickCol(r, COL_ANEXO)),
      };
    });

    state.rows = rows;
    state.filtered = [...rows];

    const unidades = [...new Set(rows.map(r => r.unidade).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'pt-BR'));
    const areas = [...new Set(rows.map(r => r.area).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'pt-BR'));

    fillSelect(els.fUnidade, unidades);
    fillSelect(els.fArea, areas);

    renderChips();
    render();
  }

  // ===== Events =====
  function bind(){
    els.q.addEventListener('input', () => {
      state.filters.q = els.q.value.trim();
      applyFilters();
    });

    els.clearQ.addEventListener('click', () => {
      els.q.value = '';
      state.filters.q = '';
      applyFilters();
      els.q.focus();
    });

    els.btnFilters.addEventListener('click', () => {
      els.fUnidade.value = state.filters.unidade;
      els.fArea.value = state.filters.area;
      els.dialog.showModal();
    });

    els.clearFilters.addEventListener('click', () => {
      state.filters.unidade = 'Todos';
      state.filters.area = 'Todos';
      applyFilters();
      els.dialog.close();
    });

    els.dialog.addEventListener('close', () => {
      if (els.dialog.returnValue === 'ok'){
        state.filters.unidade = els.fUnidade.value || 'Todos';
        state.filters.area = els.fArea.value || 'Todos';
        applyFilters();
      }
    });
  }

  // ===== Init =====
  bind();
  loadXlsx().catch(err => {
    console.error(err);
    els.countInfo.textContent = 'Erro ao carregar iniciativas';
    els.empty.classList.remove('hidden');
    els.empty.innerHTML = `
      <h2>Erro ao carregar o arquivo Excel</h2>
      <p>Confira se o arquivo está em <b>iniciativas.xlsx</b> e se o GitHub Pages está servindo a pasta corretamente.</p>
      <p style="color:#777;font-size:12px">${escapeHtml(err.message)}</p>
    `;
  });
})();