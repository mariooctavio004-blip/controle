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
  diverg: {}      // diverg[farmIdx] = { nd, ns, md, ms } numbers
};

let state = null;
// =========================
// IMPORTAÇÃO DE PLANILHAS
// Suporta os layouts das planilhas usadas no Excel:
// - Diário, Operação, Abastecimento e Mapa Mensal: fazendas nas linhas e datas nas colunas.
// - Divergências: cabeçalhos de nascimento/mortes diário x sistema.
// =========================

const importedData = {
  campo: null,
  abastecimento: null,
  diario: null,
  mensal: null,
  divergencias: null
};

const DAILY_IMPORT_KEYS = ['campo', 'abastecimento', 'diario'];

function normalizeText(value){
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function normalizeCompact(value){
  return normalizeText(value).replace(/\s+/g, '');
}

function levenshteinDistance(a, b){
  const matrix = Array.from({ length:a.length + 1 }, () => Array(b.length + 1).fill(0));
  for(let i = 0; i <= a.length; i++) matrix[i][0] = i;
  for(let j = 0; j <= b.length; j++) matrix[0][j] = j;
  for(let i = 1; i <= a.length; i++){
    for(let j = 1; j <= b.length; j++){
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost
      );
    }
  }
  return matrix[a.length][b.length];
}

function isSimilarFarmName(cell, farm){
  const cellNorm = normalizeCompact(cell);
  const farmNorm = normalizeCompact(farm);
  if(!cellNorm || !farmNorm) return false;
  if(cellNorm === farmNorm || cellNorm.includes(farmNorm) || farmNorm.includes(cellNorm)) return true;

  const cellWords = normalizeText(cell).split(' ').filter(Boolean);
  const farmWords = normalizeText(farm).split(' ').filter(Boolean);
  const matchingWords = cellWords.filter(word => farmWords.includes(word));
  if(matchingWords.length && matchingWords.length >= Math.min(2, farmWords.length)) return true;

  const distance = levenshteinDistance(cellNorm, farmNorm);
  return distance <= Math.max(2, Math.floor(farmNorm.length * 0.18));
}

function classifyImportFile(filename){
  const name = normalizeText(filename);
  if(name.includes('abaste')) return 'abastecimento';
  if(name.includes('diario')) return 'diario';
  if(name.includes('diverg')) return 'divergencias';
  if(name.includes('mensal') || name.includes('rebanho')) return 'mensal';
  if(name.includes('opera') || name.includes('campo')) return 'campo';
  return null;
}

function sheetRowsToJsonRows(rows){
  return rows.map(row => row.map(cell => String(cell ?? '').trim()));
}

function parseDateLike(value){
  const text = String(value ?? '').trim();
  if(!text) return null;

  const br = text.match(/\b(\d{1,2})[\/.-](\d{1,2})(?:[\/.-]\d{2,4})?\b/);
  if(br){
    return `${br[1].padStart(2, '0')}/${br[2].padStart(2, '0')}`;
  }

  const monthMap = {
    jan:'01', january:'01', fev:'02', feb:'02', february:'02', fevereiro:'02',
    mar:'03', march:'03', marco:'03', março:'03', abr:'04', apr:'04', april:'04', abril:'04',
    mai:'05', may:'05', maio:'05', jun:'06', june:'06', junho:'06', jul:'07', july:'07', julho:'07',
    ago:'08', aug:'08', august:'08', agosto:'08', set:'09', sep:'09', sept:'09', september:'09', setembro:'09',
    out:'10', oct:'10', october:'10', outubro:'10', nov:'11', november:'11', novembro:'11',
    dez:'12', dec:'12', december:'12', dezembro:'12'
  };
  const monthName = normalizeText(text).match(/\b(\d{1,2})\s*([a-z]+)\b|\b([a-z]+)\s*(\d{1,2})\b/);
  if(monthName){
    const day = monthName[1] || monthName[4];
    const month = monthName[2] || monthName[3];
    const monthNumber = monthMap[month];
    if(monthNumber) return `${String(day).padStart(2, '0')}/${monthNumber}`;
  }

  const parsed = new Date(text);
  if(!Number.isNaN(parsed.getTime())){
    return parsed.toLocaleDateString('pt-BR', { day:'2-digit', month:'2-digit' });
  }

  return null;
}

function getFarmMatchFromRow(row){
  for(let cellIdx = 0; cellIdx < row.length; cellIdx++){
    const cell = row[cellIdx];
    const farmIdx = state.farms.findIndex(farm => isSimilarFarmName(cell, farm));
    if(farmIdx >= 0) return { farmIdx, cellIdx };
  }
  return { farmIdx:-1, cellIdx:-1 };
}

function getFarmIndexFromRow(row){
  return getFarmMatchFromRow(row).farmIdx;
}

function cellToStatus(value){
  const text = normalizeText(value);
  if(!text) return 'no';
  if(['ok', 'sim', 's', 'recebido', 'enviado', 'entregue', 'lancado', 'feito', 'concluido', '1', 'true', 'x'].includes(text)) return 'ok';
  if(text.includes('recebid') || text.includes('enviad') || text.includes('concluid') || text.includes('lancad')) return 'ok';
  if(text.includes('compra') || text.includes('n a') || text.includes('nao se aplica') || text === '-') return 'blank';
  if(text.includes('pend') || text.includes('nao') || text.includes('falta') || text.includes('atras')) return 'no';
  const num = Number(String(value).replace(',', '.'));
  if(!Number.isNaN(num)) return num > 0 ? 'ok' : 'no';
  return 'ok';
}

function findDateColumns(rows){
  const columns = {};
  rows.slice(0, 12).forEach(row => {
    row.forEach((cell, idx) => {
      const label = parseDateLike(cell);
      if(label && state.days.includes(label)) columns[label] = idx;
    });
  });
  return columns;
}

function updateDailyPanelFromRows(panelKey, rows){
  if(!rows || !rows.length) return 0;
  const dateColumns = findDateColumns(rows);
  let updates = 0;

  rows.forEach(row => {
    const farmIdx = getFarmIndexFromRow(row);
    if(farmIdx < 0) return;

    state.days.forEach((day, dayIdx) => {
      const colIdx = dateColumns[day];
      if(colIdx === undefined) return;
      state.data[panelKey][farmIdx][dayIdx] = cellToStatus(row[colIdx]);
      updates++;
    });
  });

  return updates;
}

function updateMonthlyFromRows(rows){
  if(!rows || !rows.length) return 0;
  let updates = 0;
  rows.forEach(row => {
    const { farmIdx, cellIdx } = getFarmMatchFromRow(row);
    if(farmIdx < 0) return;

    const statuses = row
      .slice(cellIdx + 1)
      .map(cellToStatus)
      .filter((status, idx) => String(row[cellIdx + 1 + idx] ?? '').trim() !== '');

    if(!statuses.length) return;
    state.monthly[farmIdx] = statuses.includes('no') ? 'no' : (statuses.includes('ok') ? 'ok' : 'blank');
    updates++;
  });
  return updates;
}

function findHeaderColumns(rows){
  const aliases = {
    nd: ['nasc diario', 'nascimento diario', 'nascimentos diario'],
    ns: ['nasc sistema', 'nascimento sistema', 'nascimentos sistema'],
    md: ['mortes diario', 'morte diario', 'mort diario'],
    ms: ['mortes sistema', 'morte sistema', 'mort sistema']
  };
  const columns = {};
  rows.slice(0, 15).forEach(row => {
    row.forEach((cell, idx) => {
      const text = normalizeText(cell);
      Object.entries(aliases).forEach(([field, names]) => {
        if(columns[field] === undefined && names.some(name => text.includes(name))) columns[field] = idx;
      });
    });
  });
  return columns;
}

function updateDivergenciasFromRows(rows){
  if(!rows || !rows.length) return 0;
  const columns = findHeaderColumns(rows);
  let updates = 0;
  rows.forEach(row => {
    const farmIdx = getFarmIndexFromRow(row);
    if(farmIdx < 0) return;
    ['nd', 'ns', 'md', 'ms'].forEach(field => {
      const colIdx = columns[field];
      if(colIdx === undefined) return;
      const raw = String(row[colIdx] ?? '').replace(/[^0-9-]/g, '');
      state.diverg[farmIdx][field] = raw;
      updates++;
    });
  });
  return updates;
}

function applyImportedData(){
  ensureData();
  const counts = [];
  DAILY_IMPORT_KEYS.forEach(key => {
    const updated = updateDailyPanelFromRows(key, importedData[key]);
    if(updated) counts.push(`${PANEL_DEFS.find(p => p.key === key).title}: ${updated}`);
  });

  const monthlyUpdated = updateMonthlyFromRows(importedData.mensal);
  if(monthlyUpdated) counts.push(`Mapa Mensal: ${monthlyUpdated}`);

  const divergUpdated = updateDivergenciasFromRows(importedData.divergencias);
  if(divergUpdated) counts.push(`Divergências: ${divergUpdated}`);

  renderAll();
  saveState();
  showInlineWarning(counts.length ? `Planilhas importadas. Campos atualizados: ${counts.join(' | ')}` : 'Nenhum dado compatível foi encontrado nas planilhas importadas.');
}

async function importExcelFiles(files){
  for(const file of files){
    const key = classifyImportFile(file.name);
    if(!key){
      showInlineWarning(`Não reconheci o tipo da planilha: ${file.name}`);
      continue;
    }

    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type:'array', cellDates:true });
    const rows = [];
    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { header:1, raw:false, defval:'' });
      rows.push(...sheetRowsToJsonRows(json));
    });
    importedData[key] = rows;
  }

  applyImportedData();
}

function preencherSistema(){
  applyImportedData();
}

document.getElementById('importExcelBtn')?.addEventListener('click', () => {
  document.getElementById('excelFiles').click();
});

document.getElementById('excelFiles')?.addEventListener('change', async e => {
  const files = [...e.target.files];
  if(!files.length) return;
  await importExcelFiles(files);
  e.target.value = '';
});

document.getElementById('processBtn')?.addEventListener('click', preencherSistema);
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

function renderAll(){ renderConfig(); renderReport(); }

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
  let rows = '';
  const perDayPend = state.days.map(()=>0);
  let totalPendAll = 0;
  state.farms.forEach((farm, fi)=>{
    let cells = '', totalPend = 0;
    state.days.forEach((day, di)=>{
      const status = normalizeStatus(state.data[panel.key][fi] && state.data[panel.key][fi][di]);
      if(status === 'no'){ totalPend++; perDayPend[di]++; }
      cells += `<td>${iconHTML(status, `data-panel="${panel.key}" data-farm="${fi}" data-day="${di}"`)}</td>`;
    });
    totalPendAll += totalPend;
    rows += `<tr><td class="farm-col">${escapeHtml(farm)}</td>${cells}<td class="rep-total pend-col">${totalPend}</td></tr>`;
  });
  const dayHeaders = state.days.map(d=>`<th>${escapeHtml(d)}</th>`).join('');
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

document.getElementById('companyName').addEventListener('input', e=>{ state.company = e.target.value; renderReport(); saveState(); });
document.getElementById('reportTitle').addEventListener('input', e=>{ state.title = e.target.value; renderReport(); saveState(); });
document.getElementById('periodLabel').addEventListener('input', e=>{ state.period = e.target.value; renderReport(); saveState(); });
document.getElementById('bannerText').addEventListener('input', e=>{ state.banner = e.target.value; renderReport(); saveState(); });

document.getElementById('addFarmBtn').addEventListener('click', ()=>{
  state.farms.push('Nova fazenda'); ensureData(); renderAll(); saveState();
});
document.getElementById('addDayBtn').addEventListener('click', ()=>{
  state.days.push('--/--'); ensureData(); renderAll(); saveState();
});

document.getElementById('resetBtn').addEventListener('click', ()=>{
  document.getElementById('confirmOverlay').style.display = 'flex';
});
document.getElementById('confirmCancel').addEventListener('click', ()=>{
  document.getElementById('confirmOverlay').style.display = 'none';
});
document.getElementById('confirmOk').addEventListener('click', ()=>{
  state = JSON.parse(JSON.stringify(defaultState));
  ensureData(); renderAll(); saveState();
  document.getElementById('confirmOverlay').style.display = 'none';
});

document.getElementById('toggleConfigBtn').addEventListener('click', (e)=>{
  const panel = document.getElementById('configPanel');
  const hidden = panel.style.display === 'none';
  panel.style.display = hidden ? 'block' : 'none';
  e.target.textContent = hidden ? 'Ocultar edição' : 'Mostrar edição';
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
// ===========================================
// VALIDAÇÃO DAS PLANILHAS
// ===========================================

function validarImportacao(){
  const labels = {
    campo: 'Operações em Campo',
    abastecimento: 'Abastecimento',
    diario: 'Diário',
    mensal: 'Mapa Mensal',
    divergencias: 'Divergências'
  };
  const importadas = Object.entries(importedData)
    .filter(([, rows]) => Array.isArray(rows) && rows.length)
    .map(([key]) => labels[key]);

  if(!importadas.length){
    alert('Nenhuma planilha foi importada ainda.');
    return;
  }

  alert('Planilhas importadas:\n\n' + importadas.join('\n'));
  prepararDados();
}

function prepararDados(){
  console.clear();
  Object.entries(importedData).forEach(([key, rows]) => {
    if(Array.isArray(rows) && rows.length){
      console.log(key);
      console.table(rows);
    }
  });
}
loadState();
