/* ============================================================
   UTILS.JS
   Funções utilitárias genéricas: normalização de texto,
   comparação de nomes de fazenda, escape de HTML, ícones,
   atalho de addEventListener, reindexação de objetos e
   conversões de células de planilha.
============================================================ */

// Remove acentos e caixa para comparar nomes de abas/fazendas com segurança.
function normalizeName(v){
    return String(v == null ? '' : v).trim().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
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

function escapeHtml(s){
  return String(s).replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function iconHTML(status, attrs){
  const cls = status === 'ok' ? 'ok' : (status === 'no' ? 'no' : 'blank');
  const glyph = status === 'ok' ? '✔' : (status === 'no' ? '✕' : '');
  return `<span class="rep-icon ${cls}" data-status="${status}" ${attrs}>${glyph}</span>`;
}

function bind(id, event, callback) {
    const el = document.getElementById(id);

    if (!el) {
        console.warn("Elemento não encontrado:", id);
        return;
    }

    el.addEventListener(event, callback);
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
