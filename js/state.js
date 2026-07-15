/* ============================================================
   STATE.JS
   Estado padrão, estado atual, dados importados e funções que
   garantem/normalizam a integridade do estado, além de
   carregamento/salvamento (storage) e polling de atualizações.
============================================================ */

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
// PERSISTÊNCIA (storage compartilhado) E SINCRONIZAÇÃO ENTRE ABAS/USUÁRIOS
// =========================

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
