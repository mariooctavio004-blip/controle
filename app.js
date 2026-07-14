const STORAGE_KEY = 'status-envio-fazendas-v2';

const PANEL_DEFS = [
  { key:'campo', title:'Operações em Campo', type:'daily' },
  { key:'abastecimento', title:'Abastecimento', type:'daily' },
  { key:'diario', title:'Diário de Campo (Manejo/Rebanho)', type:'daily' },
  { key:'mensal', title:'Mapa Mensal do Rebanho', type:'monthly' },
  { key:'divergencias', title:'Divergências entre Diário e Sistema', type:'divergencias' }
];

const PANEL_ICONS = {
  campo: '🚜',
  abastecimento: '⛽',
  diario: '📋',
  mensal: '🐄',
  divergencias: '⚖️'
};

function panelIconHTML(key){
  const icon = PANEL_ICONS[key];
  return icon ? `<span class="rep-card-icon">${icon}</span>` : '';
}

const defaultState = {
  company: 'Nelore Cometa',
  title: 'STATUS DE ENVIO DOS CONTROLES - FAZENDAS',
  period: '01/07 A 04/07',
  monthLabel: 'FECHAMENTO – JUNHO/2025',
  banner: 'Contamos com todos para mantermos a disciplina e a qualidade das informações!',
  farms: ['Furna Linda','Estância Cometa','Vale do Jaurú','Pantanal 1','Pantanal 2','São Lucas','Estância Maristela','Santa Rita','Liberdade','São Sebastião'],
  days: ['01/07','02/07','03/07','04/07'],
  selected: ['campo','abastecimento','diario','mensal'],
  data: {},       // data[panelKey][farmIdx][dayIdx] = bool  (daily panels)
  monthly: {},    // monthly[farmIdx] = bool
  diverg: {},     // diverg[farmIdx] = { nd, ns, md, ms } numbers
  sharepoint: {

  campo: {
    name: "Operações em Campo",
    url: "",
    connected: false,
    lastSync: ""
  },

  abastecimento: {
    name: "Abastecimento",
    url: "",
    connected: false,
    lastSync: ""
  },

  diario: {
    name: "Diário de Campo",
    url: "",
    connected: false,
    lastSync: ""
  },

  mensal: {
    name: "Mapa Mensal do Rebanho",
    url: "",
    connected: false,
    lastSync: ""
  },

  divergencias: {
    name: "Divergências entre Diário e Sistema",
    url: "",
    connected: false,
    lastSync: ""
  }

},
  filterStart: '',     // data inicial (YYYY-MM-DD) do período mostrado nos painéis
  filterEnd: ''         // data final (YYYY-MM-DD) do período mostrado nos painéis
};

let state = null;
// =========================
// IMPORTAÇÃO DE PLANILHAS
// =========================

const importedData = {
    campo: null,
    abastecimento: null,
    diario: null,
    divergencias: null
};

const PANEL_LABELS = {
    campo: 'Operações em Campo',
    abastecimento: 'Abastecimento',
    diario: 'Diário de Campo',
    divergencias: 'Divergências entre Diário e Sistema'
};

// Remove acentos e caixa para comparar nomes de abas/fazendas com segurança.
function normalizeName(v){
    return String(v == null ? '' : v).trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

// Decide a que painel uma aba (ou, na falta de aba reconhecível, o nome do
// arquivo) corresponde, pelo nome.
function classifySheetName(name){
    const n = normalizeName(name);
    if (n.includes('DIVERG')) return 'divergencias';
    if (n.includes('DIARIO')) return 'diario';
    if (n.includes('OPERA')) return 'campo';
    if (n.includes('ABASTE')) return 'abastecimento';
    return null;
}

// Lê todas as abas de um workbook já carregado e joga o conteúdo reconhecido
// dentro de importedData. Usado tanto para arquivo local quanto para planilha
// baixada de um link online (SharePoint/OneDrive).
function processWorkbook(workbook, fallbackName){
    workbook.SheetNames.forEach(sheetName => {
        let key = classifySheetName(sheetName);
        // Se o arquivo tiver uma única aba com nome genérico, tenta
        // identificar o painel pelo nome do próprio arquivo.
        if (!key && workbook.SheetNames.length === 1) {
            key = classifySheetName(fallbackName);
        }
        if (!key) return;

        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            raw: false,
            dateNF: 'dd/mm'
        });
        importedData[key] = json;
    });
}

document.getElementById("excelFiles")?.addEventListener("change", async (e) => {

    const files = [...e.target.files];

    if (!files.length) return;

    for (const file of files) {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, {
            type: "array",
            cellDates: true
        });
        // Lê TODAS as abas do arquivo (a sua planilha tem as 4 tabelas em
        // abas separadas dentro de um único arquivo .xlsx).
        processWorkbook(workbook, file.name);
    }

    e.target.value = ''; // permite importar o mesmo arquivo de novo depois

    preencherSistema();

});

// =========================
// PLANILHA ONLINE (SHAREPOINT/ONEDRIVE)
// =========================

// Tenta transformar um link de compartilhamento do SharePoint/OneDrive em um
// link de download direto (adiciona "download=1" na URL). Isso só funciona
// se o arquivo estiver compartilhado como "Qualquer pessoa com o link" e se o
// SharePoint permitir o acesso direto do navegador (CORS); alguns tenants
// bloqueiam esse acesso, e nesse caso a sincronização automática falha e um
// aviso aparece na tela.
function toDirectDownloadUrl(url){
    try{
        const u = new URL(url);
        const host = u.hostname.toLowerCase();
        if(host.includes('sharepoint.com') || host.includes('1drv.ms') || host.includes('onedrive.live.com')){
            if(!u.searchParams.has('download')) u.searchParams.set('download', '1');
        }
        return u.toString();
    }catch(e){
        return url;
    }
}

function renderAttachStatus(isError){
    const icon = document.getElementById('attachStatusIcon');
    const text = document.getElementById('attachStatusText');
    const removeBtn = document.getElementById('removeAttachBtn');
    const urlInput = document.getElementById('sharepointUrl');
    const info = document.querySelector('.attach-info');
    if(!icon || !text || !removeBtn || !urlInput) return;

    if(state.sharepointUrl){
        if(document.activeElement !== urlInput) urlInput.value = state.sharepointUrl;
        removeBtn.style.display = '';
        info.classList.toggle('is-error', !!isError);
        icon.textContent = isError ? '⚠️' : '🔗';
        if(isError){
            text.textContent = 'Não foi possível sincronizar a planilha anexada agora. Tentando de novo automaticamente.';
        } else {
            text.textContent = 'Planilha anexada' + (state.lastSync ? (' — última sincronização: ' + state.lastSync) : ' — sincronizando...') + '.';
        }
    } else {
        removeBtn.style.display = 'none';
        info.classList.remove('is-error');
        icon.textContent = '📎';
        text.textContent = 'Nenhuma planilha online anexada.';
    }
}

// Baixa e reprocessa a planilha anexada. silent=true evita mostrar aviso de
// erro repetido nas sincronizações automáticas de fundo.
async function syncSharepointSheet(silent){
    const url = state.sharepointUrl;
    if(!url) return;
    try{
        const direct = toDirectDownloadUrl(url);
        const res = await fetch(direct, { mode: 'cors' });
        if(!res.ok) throw new Error('HTTP ' + res.status);
        const buffer = await res.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
        processWorkbook(workbook, 'planilha');
        state.lastSync = new Date().toLocaleTimeString('pt-BR', { hour:'2-digit', minute:'2-digit' });
        preencherSistema(); // já faz ensureData + renderAll + saveState + aviso do que foi importado
        renderAttachStatus(false);
    }catch(e){
        console.error(e);
        renderAttachStatus(true);
        if(!silent){
            showInlineWarning('Não foi possível baixar a planilha do link informado. Confira se ela está compartilhada como "Qualquer pessoa com o link pode visualizar" — alguns links do SharePoint bloqueiam o acesso direto do navegador (CORS) e, nesse caso, não é possível sincronizar automaticamente. Detalhe: ' + e.message);
        }
    }
}

let sharepointSyncTimer = null;
function startSharepointAutoSync(){
    clearInterval(sharepointSyncTimer);
    sharepointSyncTimer = setInterval(() => {
        if(state.sharepointUrl) syncSharepointSheet(true);
    }, 3 * 60 * 1000); // ressincroniza a cada 3 minutos
}

document.getElementById('attachSharepointBtn')?.addEventListener('click', async () => {
    const val = document.getElementById('sharepointUrl').value.trim();
    if(!val){ showInlineWarning('Cole o link da planilha antes de anexar.'); return; }
    state.sharepointUrl = val;
    state.lastSync = '';
    saveState();
    renderAttachStatus(false);
    await syncSharepointSheet(false);
});

document.getElementById('removeAttachBtn')?.addEventListener('click', () => {
    state.sharepointUrl = '';
    state.lastSync = '';
    document.getElementById('sharepointUrl').value = '';
    renderAttachStatus(false);
    saveState();
});
const STATUS_CYCLE = { ok:'no', no:'blank', blank:'ok' };
function normalizeStatus(v){
  if(v === true) return 'ok';
  if(v === false) return 'no';
  if(v === 'ok' || v === 'no' || v === 'blank') return v;
  return 'ok';
}

function ensureData(){
  PANEL_DEFS.filter(p=>p.type==='daily').forEach(p=>{
    if(!state.data[p.key]) state.data[p.key] = {};
    state.farms.forEach((f,fi)=>{
      if(!state.data[p.key][fi]) state.data[p.key][fi] = {};
      state.days.forEach((d,di)=>{
        state.data[p.key][fi][di] = normalizeStatus(state.data[p.key][fi][di]);
      });
    });
  });
  state.farms.forEach((f,fi)=>{
    state.monthly[fi] = normalizeStatus(state.monthly[fi]);
    if(!state.diverg[fi]) state.diverg[fi] = { nd:'', ns:'', md:'', ms:'' };
  });
  if(!Array.isArray(state.selected) || state.selected.length===0){
    state.selected = ['campo','abastecimento','diario','mensal'];
  }
  state.selected = state.selected.slice(0,4);
}

// =========================
// FILTRO DE PERÍODO (mostrar só os últimos N dias / um intervalo)
// =========================

// Converte um rótulo de coluna tipo "01/07" ou "01/07/2026" em Date. Se não
// conseguir entender o formato, retorna null (a coluna some do cálculo de
// filtro, mas continua sendo mostrada por segurança).
function parseDayLabel(label){
    const m = String(label == null ? '' : label).trim().match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
    if(!m) return null;
    const day = parseInt(m[1], 10), month = parseInt(m[2], 10) - 1;
    let year = m[3] ? parseInt(m[3], 10) : new Date().getFullYear();
    if(year < 100) year += 2000;
    const d = new Date(year, month, day);
    return isNaN(d.getTime()) ? null : d;
}

function isoToDate(iso){
    if(!iso) return null;
    const d = new Date(iso + 'T00:00:00');
    return isNaN(d.getTime()) ? null : d;
}

function dateToISO(d){
    const y = d.getFullYear(), m = String(d.getMonth()+1).padStart(2,'0'), day = String(d.getDate()).padStart(2,'0');
    return `${y}-${m}-${day}`;
}

// Retorna os índices de state.days que devem aparecer nos painéis, já
// aplicando o filtro de período (campos de data ao lado de "Importar
// Planilhas"). Os dados de todos os dias continuam salvos normalmente —
// só a exibição é filtrada aqui.
function getVisibleDayIndexes(){
    const allIdx = state.days.map((_, i) => i);
    let endDate = isoToDate(state.filterEnd);
    let startDate = isoToDate(state.filterStart);

    if(!endDate && !startDate) return allIdx; // sem filtro definido: mostra tudo

    const parsedDays = state.days.map(parseDayLabel);
    const allParsed = parsedDays.every(d => d !== null);

    // Se as colunas não estiverem no formato DD/MM, não dá pra comparar datas
    // com segurança. Nesse caso, só a data final é usada para pegar as
    // últimas 7 colunas (por posição).
    if(!allParsed){
        if(endDate && !startDate) return allIdx.slice(Math.max(0, allIdx.length - 7));
        return allIdx;
    }

    if(endDate && !startDate){
        startDate = new Date(endDate);
        startDate.setDate(startDate.getDate() - 6); // janela padrão de 7 dias
    }
    if(startDate && !endDate){
        endDate = new Date(startDate);
        endDate.setDate(endDate.getDate() + 6);
    }

    const filtered = allIdx.filter(i => {
        const d = parsedDays[i];
        if(!d) return true;
        return d >= startDate && d <= endDate;
    });
    return filtered.length ? filtered : allIdx;
}

let lastKnownJSON = null;

async function loadState(){
  try{
    const res = await window.storage.get(STORAGE_KEY, true);
    state = (res && res.value) ? JSON.parse(res.value) : JSON.parse(JSON.stringify(defaultState));
  }catch(e){
    state = JSON.parse(JSON.stringify(defaultState));
  }
  ensureData();
  lastKnownJSON = JSON.stringify(state);
  renderAll();
  startPolling();
  startSharepointAutoSync();
  if(state.sharepointUrl) syncSharepointSheet(true); // sincroniza assim que a página abre
}

let saveTimer = null;
function saveState(){
  document.getElementById('saveStatus').textContent = 'Salvando...';
  clearTimeout(saveTimer);
  saveTimer = setTimeout(async ()=>{
    try{
      const json = JSON.stringify(state);
      await window.storage.set(STORAGE_KEY, json, true);
      lastKnownJSON = json;
      const now = new Date().toLocaleTimeString('pt-BR', {hour:'2-digit', minute:'2-digit'});
      document.getElementById('saveStatus').textContent = 'Salvo às ' + now + ' (compartilhado)';
    }catch(e){
      document.getElementById('saveStatus').textContent = 'Erro ao salvar';
    }
  }, 400);
}

function isTypingNow(){
  const el = document.activeElement;
  if(!el) return false;
  const tag = el.tagName;
  return tag === 'INPUT' || tag === 'TEXTAREA';
}

let pollTimer = null;
function startPolling(){
  clearInterval(pollTimer);
  pollTimer = setInterval(async ()=>{
    if(isTypingNow()) return; // evita sobrescrever enquanto alguém digita
    try{
      const res = await window.storage.get(STORAGE_KEY, true);
      const remoteJSON = (res && res.value) ? res.value : null;
      if(remoteJSON && remoteJSON !== lastKnownJSON){
        state = JSON.parse(remoteJSON);
        ensureData();
        lastKnownJSON = JSON.stringify(state);
        renderAll();
        showInlineWarning('Os dados foram atualizados por outra pessoa agora mesmo.');
      }
    }catch(e){ /* silencioso: só tenta de novo no próximo ciclo */ }
  }, 6000);
}

function renderAll(){ renderConfig(); renderReport(); renderAttachStatus(); syncDateInputs(); }

// Mantém os campos de data (topo) sincronizados com o state, sem atrapalhar
// enquanto a pessoa está digitando neles.
function syncDateInputs(){
  const startInp = document.getElementById('reportStartDate');
  const endInp = document.getElementById('reportEndDate');
  if(startInp && document.activeElement !== startInp) startInp.value = state.filterStart || '';
  if(endInp && document.activeElement !== endInp) endInp.value = state.filterEnd || '';
}

// Calcula distância de edição entre duas strings, usada para casar nomes de
// fazenda com pequenas diferenças de digitação (ex.: "ESTACIA" x "ESTANCIA").
function levenshtein(a, b){
    const m = a.length, n = b.length;
    const dp = Array.from({length: m + 1}, () => new Array(n + 1).fill(0));
    for (let i = 0; i <= m; i++) dp[i][0] = i;
    for (let j = 0; j <= n; j++) dp[0][j] = j;
    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            dp[i][j] = a[i-1] === b[j-1] ? dp[i-1][j-1] : 1 + Math.min(dp[i-1][j-1], dp[i-1][j], dp[i][j-1]);
        }
    }
    return dp[m][n];
}

// Encontra o índice da fazenda em state.farms que melhor corresponde ao nome
// vindo da planilha (tolera acentuação, abreviações e pequenos erros de digitação).
function matchFarmIndex(name){
    const norm = normalizeName(name);
    let idx = state.farms.findIndex(f => normalizeName(f) === norm);
    if (idx !== -1) return idx;
    idx = state.farms.findIndex(f => {
        const fn = normalizeName(f);
        return fn.includes(norm) || norm.includes(fn);
    });
    if (idx !== -1) return idx;
    let best = -1, bestDist = Infinity;
    state.farms.forEach((f, i) => {
        const dist = levenshtein(normalizeName(f), norm);
        if (dist < bestDist) { bestDist = dist; best = i; }
    });
    return (best !== -1 && bestDist <= 3) ? best : -1;
}

function cellToStatus(v){
    const n = normalizeName(v);
    if (!n) return 'blank';
    if (n.includes('PENDENTE')) return 'no';
    return 'ok'; // RECEBIDO, COMPRA, etc. contam como enviado/tratado
}

function toNumberOrEmpty(v){
    if (v === undefined || v === null || v === '') return '';
    const n = Number(String(v).replace(',', '.'));
    return isNaN(n) ? '' : n;
}

// Lê uma aba "diária" (Operações em Campo / Abastecimento / Diário de Campo):
// localiza a linha "PERÍODO:" para achar as datas das colunas, depois lê uma
// linha por fazenda com o status de cada dia.
function parseStatusSheet(json){
    let headerRow = -1, farmCol = -1;
    for (let r = 0; r < json.length && headerRow === -1; r++) {
        const row = json[r] || [];
        for (let c = 0; c < row.length; c++) {
            if (normalizeName(row[c]).includes('PERIODO')) { headerRow = r; farmCol = c; break; }
        }
    }
    if (headerRow === -1) return null;

    const days = [];
    const hRow = json[headerRow] || [];
    for (let c = farmCol + 1; c < hRow.length; c++) {
        const v = hRow[c];
        if (v === undefined || v === null || String(v).trim() === '') break;
        days.push(String(v).trim());
    }

    const rows = [];
    for (let r = headerRow + 1; r < json.length; r++) {
        const row = json[r] || [];
        const farm = row[farmCol];
        if (!farm || !String(farm).trim()) continue;
        const statuses = days.map((d, i) => cellToStatus(row[farmCol + 1 + i]));
        rows.push({ farm: String(farm).trim(), statuses });
    }
    return { days, rows };
}

// Lê a aba de Divergências: cabeçalho "Data" | "Fazenda" | ... e uma linha
// por fazenda.
function parseDivergSheet(json){
    let headerRow = -1, dataCol = -1, farmCol = -1;
    for (let r = 0; r < json.length && headerRow === -1; r++) {
        const row = json[r] || [];
        for (let c = 0; c < row.length; c++) {
            if (normalizeName(row[c]) === 'DATA' && normalizeName(row[c+1]).includes('FAZENDA')) {
                headerRow = r; dataCol = c; farmCol = c + 1; break;
            }
        }
    }
    if (headerRow === -1) return null;

    const rows = [];
    for (let r = headerRow + 1; r < json.length; r++) {
        const row = json[r] || [];
        const farm = row[farmCol];
        if (!farm || !String(farm).trim()) continue;
        rows.push({
            farm: String(farm).trim(),
            nd: toNumberOrEmpty(row[farmCol + 1]),
            ns: toNumberOrEmpty(row[farmCol + 2]),
            md: toNumberOrEmpty(row[farmCol + 4]),
            ms: toNumberOrEmpty(row[farmCol + 5])
        });
    }
    return rows;
}

function preencherSistema(){
    const parsed = {
        campo: importedData.campo ? parseStatusSheet(importedData.campo) : null,
        abastecimento: importedData.abastecimento ? parseStatusSheet(importedData.abastecimento) : null,
        diario: importedData.diario ? parseStatusSheet(importedData.diario) : null,
        divergencias: importedData.divergencias ? parseDivergSheet(importedData.divergencias) : null
    };

    const found = [], missing = [];
    const unmatched = new Set();

    // As datas da planilha viram o eixo de dias do relatório (fonte da verdade).
    const canonical = parsed.diario || parsed.campo || parsed.abastecimento;
    if (canonical && canonical.days.length) {
        state.days = canonical.days.slice();
        PANEL_DEFS.filter(p => p.type === 'daily').forEach(p => { state.data[p.key] = {}; });
    }

    ['campo', 'abastecimento', 'diario'].forEach(key => {
        const sheet = parsed[key];
        if (!sheet) { missing.push(PANEL_LABELS[key]); return; }
        found.push(PANEL_LABELS[key]);
        if (!state.data[key]) state.data[key] = {};
        sheet.rows.forEach(row => {
            const fi = matchFarmIndex(row.farm);
            if (fi === -1) { unmatched.add(row.farm); return; }
            if (!state.data[key][fi]) state.data[key][fi] = {};
            row.statuses.forEach((status, i) => {
                const di = state.days.indexOf(sheet.days[i]);
                if (di === -1) return;
                state.data[key][fi][di] = status;
            });
        });
    });

    if (parsed.divergencias) {
        found.push(PANEL_LABELS.divergencias);
        parsed.divergencias.forEach(row => {
            const fi = matchFarmIndex(row.farm);
            if (fi === -1) { unmatched.add(row.farm); return; }
            state.diverg[fi] = { nd: String(row.nd), ns: String(row.ns), md: String(row.md), ms: String(row.ms) };
        });
    } else {
        missing.push(PANEL_LABELS.divergencias);
    }

    ensureData();
    renderAll();
    saveState();

    let msg = 'Importado: ' + (found.join(', ') || 'nada reconhecido') + '.';
    if (missing.length) msg += ' Não encontrado na planilha: ' + missing.join(', ') + '.';
    if (unmatched.size) msg += ' Fazenda(s) da planilha sem correspondência: ' + [...unmatched].join(', ') + '.';
    showInlineWarning(msg);
}

function renderConfig(){
  document.getElementById('companyName').value = state.company;
  document.getElementById('reportTitle').value = state.title;
  document.getElementById('periodLabel').value = state.period;
  document.getElementById('bannerText').value = state.banner;

  const list = document.getElementById('panelSelectList');
  list.innerHTML = '';
  PANEL_DEFS.forEach(p=>{
    const idx = state.selected.indexOf(p.key);
    const isChecked = idx !== -1;
    const item = document.createElement('label');
    item.className = 'panel-select-item' + (isChecked ? ' checked' : '');
    const slotName = idx===0?'Topo esquerda':idx===1?'Topo direita':idx===2?'Baixo esquerda':idx===3?'Baixo direita':'';
    item.innerHTML = `
      <span class="order-badge ${isChecked?'':'empty'}">${isChecked ? (idx+1) : ''}</span>
      <input type="checkbox" data-key="${p.key}" ${isChecked?'checked':''} style="width:16px;height:16px;">
      <span class="name">${p.title}</span>
      <span class="slot-hint">${slotName}</span>
    `;
    list.appendChild(item);
  });
  list.querySelectorAll('input[type=checkbox]').forEach(cb=>{
    cb.addEventListener('change', e=>{
      const key = e.target.dataset.key;
      const warning = document.getElementById('selectWarning');
      if(e.target.checked){
        if(state.selected.length >= 4){
          e.target.checked = false;
          warning.style.display = 'block';
          setTimeout(()=>warning.style.display='none', 2500);
          return;
        }
        state.selected.push(key);
      } else {
        state.selected = state.selected.filter(k=>k!==key);
      }
      renderAll();
      saveState();
    });
  });

  const farmList = document.getElementById('farmList');
  farmList.innerHTML = '';
  state.farms.forEach((f,i)=>{
    const chip = document.createElement('div');
    chip.className = 'chip';
    chip.innerHTML = `<input type="text" data-idx="${i}" value="${f.replace(/"/g,'&quot;')}"><button data-idx="${i}" title="Remover">×</button>`;
    farmList.appendChild(chip);
  });
  farmList.querySelectorAll('input').forEach(inp=>{
    inp.addEventListener('input', e=>{
      state.farms[+e.target.dataset.idx] = e.target.value;
      renderReport(); saveState();
    });
  });
  farmList.querySelectorAll('button').forEach(btn=>{
    btn.addEventListener('click', e=>{
      const idx = +e.target.dataset.idx;
      state.farms.splice(idx,1);
      PANEL_DEFS.filter(p=>p.type==='daily').forEach(p=>{ delete state.data[p.key][idx]; reindexObj(state.data[p.key], idx); });
      delete state.monthly[idx]; reindexObj(state.monthly, idx);
      delete state.diverg[idx]; reindexObj(state.diverg, idx);
      renderAll(); saveState();
    });
  });

  const dayList = document.getElementById('dayList');
  dayList.innerHTML = '';
  state.days.forEach((d,i)=>{
    const chip = document.createElement('div');
    chip.className = 'chip';
    chip.innerHTML = `<input type="text" data-idx="${i}" value="${d.replace(/"/g,'&quot;')}"><button data-idx="${i}" title="Remover">×</button>`;
    dayList.appendChild(chip);
  });
  dayList.querySelectorAll('input').forEach(inp=>{
    inp.addEventListener('input', e=>{
      state.days[+e.target.dataset.idx] = e.target.value;
      renderReport(); saveState();
    });
  });
  dayList.querySelectorAll('button').forEach(btn=>{
    btn.addEventListener('click', e=>{
      const idx = +e.target.dataset.idx;
      state.days.splice(idx,1);
      PANEL_DEFS.filter(p=>p.type==='daily').forEach(p=>{
        Object.keys(state.data[p.key]).forEach(fi=>{
          delete state.data[p.key][fi][idx];
          reindexObj(state.data[p.key][fi], idx);
        });
      });
      renderAll(); saveState();
    });
  });
}

function reindexObj(obj, removedIdx){
  const keys = Object.keys(obj).map(Number).sort((a,b)=>a-b);
  const newObj = {};
  keys.forEach(k=>{
    if(k < removedIdx) newObj[k] = obj[k];
    else if(k > removedIdx) newObj[k-1] = obj[k];
  });
  Object.keys(obj).forEach(k=>delete obj[k]);
  Object.assign(obj, newObj);
}

function iconHTML(status, attrs){
  const cls = status === 'ok' ? 'ok' : (status === 'no' ? 'no' : 'blank');
  const glyph = status === 'ok' ? '✔' : (status === 'no' ? '✕' : '');
  return `<span class="rep-icon ${cls}" data-status="${status}" ${attrs}>${glyph}</span>`;
}

function escapeHtml(s){
  return String(s).replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function buildDailyCard(panel){
  const visible = getVisibleDayIndexes(); // índices de state.days a exibir, já filtrados pelo período escolhido
  let rows = '';
  const perDayPend = visible.map(()=>0);
  let totalPendAll = 0;
  state.farms.forEach((farm, fi)=>{
    let cells = '', totalPend = 0;
    visible.forEach((di, col)=>{
      const status = normalizeStatus(state.data[panel.key][fi] && state.data[panel.key][fi][di]);
      if(status === 'no'){ totalPend++; perDayPend[col]++; }
      cells += `<td>${iconHTML(status, `data-panel="${panel.key}" data-farm="${fi}" data-day="${di}"`)}</td>`;
    });
    totalPendAll += totalPend;
    rows += `<tr><td class="farm-col">${escapeHtml(farm)}</td>${cells}<td class="rep-total pend-col">${totalPend}</td></tr>`;
  });
  const dayHeaders = visible.map(di=>`<th>${escapeHtml(state.days[di])}</th>`).join('');
  const dayFooterCells = perDayPend.map(n=>`<td>${n}</td>`).join('');
  return `
    <div class="rep-card-head">${panelIconHTML(panel.key)}<span>${escapeHtml(panel.title)}</span></div>
    <table class="rep-table">
      <thead><tr><th class="farm-col">Fazenda</th>${dayHeaders}<th class="pend-col">Pendências</th></tr></thead>
      <tbody>${rows}</tbody>
      <tfoot><tr class="pend-row">
        <td class="farm-col">PENDÊNCIAS:</td>${dayFooterCells}<td>${totalPendAll}</td>
      </tr></tfoot>
    </table>
  `;
}

function buildMonthlyCard(panel){
  let mPend = 0, mRows = '';
  state.farms.forEach((farm, fi)=>{
    const status = normalizeStatus(state.monthly[fi]);
    if(status === 'no') mPend++;
    mRows += `<tr><td class="farm-col">${escapeHtml(farm)}</td><td>${iconHTML(status, `data-monthly="${fi}"`)}</td></tr>`;
  });
  return `
    <div class="rep-card-head">${panelIconHTML(panel.key)}<span>${escapeHtml(panel.title)} — ${escapeHtml(state.monthLabel)}</span></div>
    <table class="rep-table">
      <thead><tr><th class="farm-col">Fazenda</th><th class="pend-col">Status</th></tr></thead>
      <tbody>${mRows}</tbody>
      <tfoot><tr class="pend-row">
        <td class="farm-col">PENDÊNCIAS:</td><td>${mPend}</td>
      </tr></tfoot>
    </table>
  `;
}

function buildDivergCard(panel){
  let rows = '';
  let somaNasc = 0, somaMortes = 0;
  state.farms.forEach((farm, fi)=>{
    const d = state.diverg[fi];
    const nd = d.nd === '' ? '' : Number(d.nd);
    const ns = d.ns === '' ? '' : Number(d.ns);
    const md = d.md === '' ? '' : Number(d.md);
    const ms = d.ms === '' ? '' : Number(d.ms);
    const difNasc = (nd==='' || ns==='') ? '' : (nd - ns);
    const difMortes = (md==='' || ms==='') ? '' : (md - ms);
    if(difNasc !== '') somaNasc += difNasc;
    if(difMortes !== '') somaMortes += difMortes;
    const difNascClass = difNasc==='' ? 'zero' : (difNasc===0 ? 'zero':'nonzero');
    const difMortesClass = difMortes==='' ? 'zero' : (difMortes===0 ? 'zero':'nonzero');
    rows += `<tr>
      <td class="farm-col">${escapeHtml(farm)}</td>
      <td><input class="num-input" type="text" inputmode="numeric" data-fi="${fi}" data-field="nd" value="${escapeHtml(d.nd)}"></td>
      <td><input class="num-input" type="text" inputmode="numeric" data-fi="${fi}" data-field="ns" value="${escapeHtml(d.ns)}"></td>
      <td class="diff-val ${difNascClass}">${difNasc}</td>
      <td><input class="num-input" type="text" inputmode="numeric" data-fi="${fi}" data-field="md" value="${escapeHtml(d.md)}"></td>
      <td><input class="num-input" type="text" inputmode="numeric" data-fi="${fi}" data-field="ms" value="${escapeHtml(d.ms)}"></td>
      <td class="diff-val ${difMortesClass}">${difMortes}</td>
    </tr>`;
  });
  return `
    <div class="rep-card-head">${panelIconHTML(panel.key)}<span>${escapeHtml(panel.title)}</span></div>
    <table class="rep-table">
      <thead><tr>
        <th class="farm-col">Fazenda</th><th>Nasc. Diário</th><th>Nasc. Sistema</th><th>Dif. Nasc.</th>
        <th>Mortes Diário</th><th>Mortes Sistema</th><th>Dif. Mortes</th>
      </tr></thead>
      <tbody>${rows}</tbody>
      <tfoot><tr class="pend-row pend-row-navy">
        <td class="farm-col">TOTAL:</td><td></td><td></td><td>${somaNasc}</td><td></td><td></td><td>${somaMortes}</td>
      </tr></tfoot>
    </table>
  `;
}

function renderReport(){
  document.getElementById('titleDisplay').textContent = state.title;
  document.getElementById('periodDisplay').textContent = state.period;
  document.getElementById('bannerDisplay').textContent = state.banner;

  const grid = document.getElementById('reportGrid');
  grid.innerHTML = '';

  state.selected.forEach(key=>{
    const panel = PANEL_DEFS.find(p=>p.key===key);
    if(!panel) return;
    const card = document.createElement('div');
    card.className = 'rep-card';
    if(panel.type === 'daily') card.innerHTML = buildDailyCard(panel);
    else if(panel.type === 'monthly') card.innerHTML = buildMonthlyCard(panel);
    else if(panel.type === 'divergencias') card.innerHTML = buildDivergCard(panel);
    grid.appendChild(card);
  });

  grid.querySelectorAll('.rep-icon[data-panel]').forEach(el=>{
    el.addEventListener('click', ()=>{
      const panel = el.dataset.panel, fi = el.dataset.farm, di = el.dataset.day;
      const cur = normalizeStatus(state.data[panel][fi][di]);
      state.data[panel][fi][di] = STATUS_CYCLE[cur];
      renderReport(); saveState();
    });
  });
  grid.querySelectorAll('.rep-icon[data-monthly]').forEach(el=>{
    el.addEventListener('click', ()=>{
      const fi = el.dataset.monthly;
      const cur = normalizeStatus(state.monthly[fi]);
      state.monthly[fi] = STATUS_CYCLE[cur];
      renderReport(); saveState();
    });
  });
  grid.querySelectorAll('.num-input').forEach(inp=>{
    inp.addEventListener('input', e=>{
      const fi = e.target.dataset.fi, field = e.target.dataset.field;
      // permite apenas dígitos e um sinal de menos no início
      let clean = e.target.value.replace(/[^0-9-]/g,'');
      clean = clean.replace(/(?!^)-/g,'');
      state.diverg[fi][field] = clean;
      renderReport(); saveState();
      // restaura o foco e o cursor após o re-render
      const again = grid.querySelector(`.num-input[data-fi="${fi}"][data-field="${field}"]`);
      if(again){
        again.focus();
        try{
          const v = again.value;
          again.setSelectionRange(v.length, v.length);
        }catch(err){ /* alguns navegadores/tipos não suportam, ignora */ }
      }
    });
  });
}

function bind(id, event, callback) {
    const el = document.getElementById(id);

    if (!el) {
        console.warn("Elemento não encontrado:", id);
        return;
    }

    el.addEventListener(event, callback);
}

bind("companyName","input",e=>{
    state.company=e.target.value;
    renderReport();
    saveState();
});

bind("reportTitle","input",e=>{
    state.title=e.target.value;
    renderReport();
    saveState();
});

bind("periodLabel","input",e=>{
    state.period=e.target.value;
    renderReport();
    saveState();
});

bind("bannerText","input",e=>{
    state.banner=e.target.value;
    renderReport();
    saveState();
});

bind("addFarmBtn","click",()=>{
    state.farms.push("Nova fazenda");
    ensureData();
    renderAll();
    saveState();
});

bind("addDayBtn","click",()=>{
    state.days.push("--/--");
    ensureData();
    renderAll();
    saveState();
});

bind("resetBtn","click",()=>{
    document.getElementById("confirmOverlay").style.display="flex";
});

bind("confirmCancel","click",()=>{
    document.getElementById("confirmOverlay").style.display="none";
});

bind("confirmOk","click",()=>{
    state=JSON.parse(JSON.stringify(defaultState));
    ensureData();
    renderAll();
    saveState();
    document.getElementById("confirmOverlay").style.display="none";
});

bind("toggleConfigBtn","click",(e)=>{
    const panel=document.getElementById("configPanel");
    const hidden=panel.style.display==="none";
    panel.style.display=hidden?"block":"none";
    e.target.textContent=hidden?"Ocultar edição":"Mostrar edição";
});
function isIOS(){
  return /iP(hone|od|ad)/.test(navigator.platform) ||
    (navigator.userAgent.includes('Mac') && 'ontouchend' in document);
}

// html2canvas e a impressão não capturam bem o valor digitado dentro de
// <input>. Por isso, na hora de gerar a imagem ou imprimir, trocamos
// temporariamente cada campo numérico por um texto simples com o mesmo valor.
function swapInputsForCapture(){
  const report = document.getElementById('report');
  report.querySelectorAll('.num-input').forEach(inp=>{
    const span = document.createElement('span');
    span.className = 'num-print-value';
    span.textContent = inp.value === '' ? '' : inp.value;
    span.style.width = inp.style.width || getComputedStyle(inp).width;
    inp.dataset.wasHidden = '1';
    inp.style.display = 'none';
    inp.insertAdjacentElement('afterend', span);
  });
}
function restoreInputsAfterCapture(){
  const report = document.getElementById('report');
  report.querySelectorAll('.num-print-value').forEach(span=>span.remove());
  report.querySelectorAll('.num-input[data-wasHidden]').forEach(inp=>{
    inp.style.display = '';
    delete inp.dataset.wasHidden;
  });
}

async function generateCanvas(){
  return await html2canvas(document.getElementById('report'), { scale:2, backgroundColor:'#ffffff', useCORS:true });
}

function showInlineWarning(msg){
  let box = document.getElementById('inlineWarningBox');
  if(!box){
    box = document.createElement('div');
    box.id = 'inlineWarningBox';
    box.className = 'inline-warning-box';
    document.querySelector('.app').appendChild(box);
  }
  box.textContent = msg;
  box.style.display = 'block';
  clearTimeout(box._timer);
  box._timer = setTimeout(()=>{ box.style.display = 'none'; }, 5000);
}

document.getElementById('exportBtn').addEventListener('click', async ()=>{
  const btn = document.getElementById('exportBtn');
  btn.textContent = 'Gerando...';
  swapInputsForCapture();
  try{
    const canvas = await generateCanvas();
    const filename = 'status_envio_' + (state.period||'periodo').replace(/[^a-zA-Z0-9]/g,'_') + '.png';

    if(isIOS()){
      // iOS Safari ignora o atributo download: abrimos a imagem numa aba
      // para o usuário segurar o dedo na imagem e escolher "Salvar imagem".
      const dataUrl = canvas.toDataURL('image/png');
      const win = window.open();
      if(win){
        win.document.write(
          '<html><head><title>' + filename + '</title></head>' +
          '<body style="margin:0;background:#111;display:flex;align-items:center;justify-content:center;">' +
          '<img src="' + dataUrl + '" style="max-width:100%;height:auto;" />' +
          '</body></html>'
        );
      } else {
        showInlineWarning('Seu navegador bloqueou a nova aba. Permita pop-ups para ver a imagem, ou use o botão Imprimir.');
      }
    } else {
      canvas.toBlob(blob=>{
        if(!blob){ showInlineWarning('Não foi possível gerar a imagem. Tente o botão Imprimir.'); return; }
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.download = filename;
        link.href = url;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(()=>URL.revokeObjectURL(url), 4000);
      }, 'image/png');
    }
  }catch(e){
    console.error(e);
    showInlineWarning('Não foi possível gerar a imagem. Tente o botão Imprimir, ou baixar de novo.');
  }
  restoreInputsAfterCapture();
  btn.textContent = 'Baixar imagem';
});

window.addEventListener('beforeprint', swapInputsForCapture);
window.addEventListener('afterprint', restoreInputsAfterCapture);

loadState();
/* ============================================================
   MODAL DAS PLANILHAS
============================================================ */

const workbookModal = document.getElementById("workbookModal");
const openWorkbookModal = document.getElementById("openWorkbookModal");
const closeWorkbookModal = document.getElementById("closeWorkbookModal");

const connectionIndicator = document.getElementById("connectionIndicator");
const currentWorkbook = document.getElementById("currentWorkbook");
const disconnectWorkbook = document.getElementById("disconnectWorkbook");

/* input oculto para abrir o seletor de arquivos */
const workbookInput = document.createElement("input");
workbookInput.type = "file";
workbookInput.accept = ".xlsx,.xls";
workbookInput.style.display = "none";
document.body.appendChild(workbookInput);

const connectedWorkbooks = {
    campo: null,
    abastecimento: null,
    diario: null,
    mensal: null,
    divergencias: null
};
let selectedWorkbookType = null;
function updateConnectionStatus() {

    const list = Object.values(connectedWorkbooks).filter(Boolean);

    if (list.length === 0) {
        connectionIndicator.innerHTML = "🔴 Nenhuma planilha conectada";
        currentWorkbook.innerHTML = "Selecione uma planilha";
        return;
    }

    connectionIndicator.innerHTML = `🟢 ${list.length} planilha(s) conectada(s)`;

    currentWorkbook.innerHTML = list.map((w, index) => `
        <div class="connected-item">
            ✔ ${w.title}<br>
            <small>${w.fileName}</small>
        </div>
    `).join("");

}
/* ===========================
   Abrir Modal
=========================== */

openWorkbookModal?.addEventListener("click", () => {

    workbookModal.style.display = "flex";

});

/* ===========================
   Fechar Modal
=========================== */

closeWorkbookModal?.addEventListener("click", () => {

    workbookModal.style.display = "none";

});

/* ===========================
   Escolha da planilha
=========================== */

document.querySelectorAll(".workbook-option").forEach(btn => {

    btn.addEventListener("click", () => {

        selectedWorkbookType = btn.dataset.id;

        workbookInput.click();

    });

});

/* ===========================
   Arquivo escolhido
=========================== */

workbookInput.addEventListener("change", e => {

    const file = e.target.files[0];

    if (!file) return;

    const option = document.querySelector(
        `.workbook-option[data-id="${selectedWorkbookType}"]`
    );

    currentWorkbook.innerHTML = `
    <strong>${option.innerText}</strong><br>
    ${file.name}
`;

workbookModal.style.display = "none";

    

/* ===========================
   Desvincular
=========================== */

disconnectWorkbook?.addEventListener("click", () => {

    if (!selectedWorkbookType) {
        alert("Selecione uma planilha primeiro.");
        return;
    }

    if (!connectedWorkbooks[selectedWorkbookType]) {
        alert("Nenhuma planilha conectada neste painel.");
        return;
    }

    if (!confirm("Deseja remover esta planilha?")) return;

    connectedWorkbooks[selectedWorkbookType] = null;

    workbookInput.value = "";

    updateConnectionStatus();

});

/* ===========================
   Fechar clicando fora
=========================== */

window.addEventListener("click", e => {

    if (e.target === workbookModal) {

        workbookModal.style.display = "none";

    }

});
