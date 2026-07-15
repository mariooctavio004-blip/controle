/* ============================================================
   EXCEL.JS
   Leitura de workbooks (.xlsx), reconhecimento das abas,
   parsing das tabelas de status/divergências e preenchimento
   do state a partir dos dados importados.
============================================================ */

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

// =========================
// IMPORTAR EXCEL (input de arquivo local)
// =========================

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
