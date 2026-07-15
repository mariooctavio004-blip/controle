/* ============================================================
   EXCEL.JS
   Importação de arquivos Excel locais e dados vindos do
   SharePoint/OneDrive.

   Recursos:
   - lê todas as abas de cada arquivo;
   - reconhece cada aba pelo nome e pelo conteúdo;
   - permite várias abas do mesmo tipo;
   - respeita o filtro Hoje / Últimos 7 dias / intervalo;
   - processa Campo, Abastecimento, Diário, Mensal e Divergências;
   - mantém compatibilidade com o sharepoint.js.
============================================================ */

const EXCEL_PANEL_KEYS = [
    "campo",
    "abastecimento",
    "diario",
    "mensal",
    "divergencias"
];

const EXCEL_PANEL_LABELS = {
    campo: "Operações em Campo",
    abastecimento: "Abastecimento",
    diario: "Diário de Campo",
    mensal: "Mapa Mensal do Rebanho",
    divergencias: "Divergências entre Diário e Sistema"
};




/* ============================================================
   REGISTRO DOS ARQUIVOS IMPORTADOS LOCALMENTE
============================================================ */

function ensureImportedFilesState() {
    if (!state) return [];

    if (!Array.isArray(state.importedFiles)) {
        state.importedFiles = [];
    }

    return state.importedFiles;
}

function registerImportedLocalFile(fileName, sheets = [], ignoredSheets = []) {
    const files = ensureImportedFilesState();
    const normalizedFileName = String(fileName || "Planilha sem nome").trim();

    const normalizedSheets = (Array.isArray(sheets) ? sheets : [])
        .map(item => ({
            name: String(item?.name || item?.sheetName || "").trim(),
            key: item?.key || null
        }))
        .filter(item => item.name && EXCEL_PANEL_KEYS.includes(item.key));

    const uniqueSheets = normalizedSheets.filter((item, index, array) =>
        array.findIndex(other =>
            normalizeName(other.name) === normalizeName(item.name) &&
            other.key === item.key
        ) === index
    );

    const uniquePanels = [...new Set(uniqueSheets.map(item => item.key))];

    const existingIndex = files.findIndex(item =>
        normalizeName(item?.fileName) === normalizeName(normalizedFileName)
    );

    const record = {
        id: existingIndex >= 0
            ? files[existingIndex].id
            : `local-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        source: "local",
        fileName: normalizedFileName,
        sheets: uniqueSheets,
        panels: uniquePanels,
        ignoredSheets: [...new Set(ignoredSheets || [])],
        importedAt: new Date().toISOString(),
        lastSync: new Date().toLocaleTimeString("pt-BR", {
            hour: "2-digit",
            minute: "2-digit"
        })
    };

    if (existingIndex >= 0) {
        files[existingIndex] = record;
    } else {
        files.push(record);
    }
}

function getImportedLocalFiles() {
    return ensureImportedFilesState().filter(item =>
        item && item.source === "local" && item.fileName
    );
}

function renderImportedWorkbookStatus() {
    if (!state) return;

    const indicator = document.getElementById("connectionIndicator");
    const list = document.getElementById("currentWorkbook");
    if (!indicator || !list) return;

    const localFiles = getImportedLocalFiles();
    const urlConnections = EXCEL_PANEL_KEYS
        .map(key => ({ key, ...(state.sharepoint?.[key] || {}) }))
        .filter(item => item.url);

    const totalSources = localFiles.length + urlConnections.length;

    if (!totalSources) {
        indicator.textContent = "Nenhuma planilha conectada";
        list.textContent = "Selecione ou conecte uma planilha";
        return;
    }

    indicator.textContent = `🟢 ${totalSources} planilha(s) disponível(is)`;

    const localHTML = localFiles.map(file => {
        const exactSheetNames = Array.isArray(file.sheets) && file.sheets.length
            ? file.sheets.map(sheet => sheet.name).filter(Boolean)
            : (file.panels || [])
                .map(key => EXCEL_PANEL_LABELS[key])
                .filter(Boolean);

        return `
            <div class="connected-item connected-item-local">
                <span>📄</span>

                <div class="connected-item-content">
                    <strong>${escapeHtml(file.fileName)}</strong>
                    <small>Abas: ${escapeHtml(exactSheetNames.join(" • ") || "Nenhuma aba reconhecida")}</small>
                </div>

                <button
                    type="button"
                    class="disconnect-local-workbook"
                    data-local-file-id="${escapeHtml(file.id)}"
                    title="Desvincular ${escapeHtml(file.fileName)}"
                    aria-label="Desvincular ${escapeHtml(file.fileName)}"
                >
                    Desvincular
                </button>
            </div>
        `;
    }).join("");

    const urlHTML = urlConnections.map(connection => `
        <div class="connected-item connected-item-url">
            <span>${connection.connected ? "🔗" : "⚠️"}</span>
            <div>
                <strong>${escapeHtml(EXCEL_PANEL_LABELS[connection.key])}</strong>
                <small>${connection.connected
                    ? `URL conectada${connection.lastSync ? ` • ${escapeHtml(connection.lastSync)}` : ""}`
                    : "URL cadastrada, aguardando sincronização"}</small>
            </div>
        </div>
    `).join("");

    list.innerHTML = localHTML + urlHTML;
}

function disconnectImportedLocalFile(fileId) {
    const files = ensureImportedFilesState();
    const index = files.findIndex(file => file?.id === fileId);

    if (index === -1) return;

    const file = files[index];
    const confirmed = window.confirm(
        `Deseja desvincular a planilha ${file.fileName}? Os dados já importados permanecerão no relatório.`
    );

    if (!confirmed) return;

    files.splice(index, 1);

    if (typeof saveState === "function") {
        saveState();
    }

    renderImportedWorkbookStatus();

    if (typeof showInlineWarning === "function") {
        showInlineWarning(`Planilha ${file.fileName} desvinculada. Os dados importados foram mantidos.`);
    }
}

function bindImportedWorkbookStatusEvents() {
    const list = document.getElementById("currentWorkbook");

    if (!list || list.dataset.localDisconnectBound === "true") return;

    list.dataset.localDisconnectBound = "true";

    list.addEventListener("click", event => {
        const button = event.target.closest(".disconnect-local-workbook");
        if (!button) return;

        event.preventDefault();
        event.stopPropagation();
        disconnectImportedLocalFile(button.dataset.localFileId);
    });
}

function installImportedWorkbookStatusIntegration() {
    if (window.__importedWorkbookStatusInstalled) return;
    window.__importedWorkbookStatusInstalled = true;

    const originalRenderAttachStatus =
        typeof window.renderAttachStatus === "function"
            ? window.renderAttachStatus
            : null;

    if (originalRenderAttachStatus) {
        window.renderAttachStatus = function(...args) {
            originalRenderAttachStatus.apply(this, args);
            renderImportedWorkbookStatus();
        };
    }

    renderImportedWorkbookStatus();
    bindImportedWorkbookStatusEvents();
}


/* ============================================================
   RECONHECIMENTO DAS ABAS
============================================================ */

function classifySheetName(name) {
    const normalized = normalizeName(name);

    if (!normalized) return null;

    if (normalized.includes("DIVERG")) {
        return "divergencias";
    }

    if (normalized.includes("ABASTE")) {
        return "abastecimento";
    }

    if (
        normalized.includes("DIARIO") ||
        normalized.includes("MANEJO")
    ) {
        return "diario";
    }

    if (
        normalized.includes("MAPA MENSAL") ||
        normalized.includes("FECHAMENTO MENSAL") ||
        normalized.includes("MENSAL")
    ) {
        return "mensal";
    }

    if (
        normalized.includes("OPERAC") ||
        normalized.includes("SISFROTA") ||
        normalized.includes("HORIMETRO") ||
        normalized === "CAMPO"
    ) {
        return "campo";
    }

    return null;
}

function sheetToMatrix(sheet) {
    if (!sheet) return [];

    return XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: true,
        defval: "",
        dateNF: "dd/mm/yyyy"
    });
}

function detectSheetKeyByContent(matrix) {
    if (!Array.isArray(matrix) || !matrix.length) {
        return null;
    }

    const sample = matrix
        .slice(0, 50)
        .flat()
        .map(value => normalizeName(value))
        .filter(Boolean)
        .join(" | ");

    if (
        sample.includes("DIVERGENCIA") ||
        sample.includes("DIF NASC") ||
        sample.includes("NASC SISTEMA") ||
        sample.includes("MORTES SISTEMA") ||
        (
            sample.includes("NASCIMENTO") &&
            sample.includes("SISTEMA") &&
            sample.includes("MORTE")
        )
    ) {
        return "divergencias";
    }

    if (
        sample.includes("ABASTECIMENTO") ||
        sample.includes("COMBUSTIVEL") ||
        sample.includes("LITROS ABASTECIDOS")
    ) {
        return "abastecimento";
    }

    if (
        sample.includes("DIARIO DE CAMPO") ||
        sample.includes("MANEJO") ||
        sample.includes("MOVIMENTACAO DO REBANHO")
    ) {
        return "diario";
    }

    if (
        sample.includes("MAPA MENSAL") ||
        sample.includes("FECHAMENTO MENSAL") ||
        sample.includes("SALDO FINAL") ||
        sample.includes("SALDO DO REBANHO")
    ) {
        return "mensal";
    }

    if (
        sample.includes("OPERACOES EM CAMPO") ||
        sample.includes("OPERACAO EM CAMPO") ||
        sample.includes("SISFROTA") ||
        sample.includes("HORIMETRO")
    ) {
        return "campo";
    }

    return null;
}


/* ============================================================
   ARMAZENAMENTO TEMPORÁRIO DE VÁRIAS ABAS
============================================================ */

function isMatrix(value) {
    return (
        Array.isArray(value) &&
        (
            value.length === 0 ||
            Array.isArray(value[0])
        )
    );
}

function addImportedSheet(key, matrix, sheetName = "") {
    if (!EXCEL_PANEL_KEYS.includes(key)) return;
    if (!isMatrix(matrix) || !matrix.length) return;

    const current = importedData[key];

    /*
     * Formato novo: lista de abas.
     * Cada item guarda nome e matriz.
     */
    if (!Array.isArray(current) || !current.length) {
        importedData[key] = [];
    } else if (isMatrix(current)) {
        /*
         * Compatibilidade: o SharePoint pode ter colocado uma matriz
         * diretamente em importedData[key].
         */
        importedData[key] = [
            {
                sheetName: EXCEL_PANEL_LABELS[key],
                matrix: current
            }
        ];
    }

    importedData[key].push({
        sheetName,
        matrix
    });
}

function getImportedMatrices(key) {
    const value = importedData[key];

    if (!value) return [];

    /* Uma matriz direta, usada pelo SharePoint. */
    if (isMatrix(value)) {
        return value.length ? [value] : [];
    }

    /* Lista de abas criada pela importação local. */
    if (Array.isArray(value)) {
        return value
            .map(item => item?.matrix)
            .filter(matrix => isMatrix(matrix) && matrix.length);
    }

    return [];
}

function hasImportedData(key) {
    return getImportedMatrices(key).length > 0;
}

function clearImportedData() {
    EXCEL_PANEL_KEYS.forEach(key => {
        importedData[key] = null;
    });
}


/* ============================================================
   LEITURA DE TODAS AS ABAS DE UM WORKBOOK
============================================================ */

function processWorkbook(workbook, fallbackName = "") {
    if (
        !workbook ||
        !Array.isArray(workbook.SheetNames) ||
        !workbook.SheetNames.length
    ) {
        throw new Error("O arquivo não possui abas válidas.");
    }

    const recognized = [];
    const recognizedSheets = [];
    const ignored = [];
    const hidden = [];
    const fallbackKey = classifySheetName(fallbackName);
    const sheetMetadata = workbook.Workbook?.Sheets || [];

    workbook.SheetNames.forEach((sheetName, sheetIndex) => {
        const metadata = sheetMetadata[sheetIndex];

        /*
         * SheetJS usa Hidden = 0 para visível, 1 para oculta e 2 para
         * muito oculta. Abas ocultas nunca são importadas.
         */
        if (Number(metadata?.Hidden || 0) !== 0) {
            hidden.push(sheetName);
            console.info(`Aba oculta ignorada: ${sheetName}`);
            return;
        }

        const sheet = workbook.Sheets[sheetName];

        if (!sheet) {
            ignored.push(sheetName);
            return;
        }

        const matrix = sheetToMatrix(sheet);

        if (!matrix.length) {
            ignored.push(sheetName);
            return;
        }

        let key = classifySheetName(sheetName);

        if (!key) {
            key = detectSheetKeyByContent(matrix);
        }

        if (!key && fallbackKey) {
            key = fallbackKey;
        }

        if (!key) {
            ignored.push(sheetName);
            console.warn(`Aba não reconhecida: ${sheetName}`);
            return;
        }

        addImportedSheet(key, matrix, sheetName);

        if (!recognized.includes(key)) {
            recognized.push(key);
        }

        recognizedSheets.push({
            name: sheetName,
            key
        });

        console.info(
            `Aba reconhecida: ${sheetName} → ${EXCEL_PANEL_LABELS[key]}`
        );
    });

    return {
        recognized,
        recognizedSheets,
        ignored,
        hidden
    };
}


/* ============================================================
   FUNÇÕES DE CABEÇALHO
============================================================ */

function findHeaderCell(matrix, acceptedTerms) {
    for (let rowIndex = 0; rowIndex < matrix.length; rowIndex++) {
        const row = matrix[rowIndex] || [];

        for (let columnIndex = 0; columnIndex < row.length; columnIndex++) {
            const value = normalizeName(row[columnIndex]);

            if (
                acceptedTerms.some(term =>
                    value.includes(normalizeName(term))
                )
            ) {
                return {
                    rowIndex,
                    columnIndex
                };
            }
        }
    }

    return null;
}

function findColumnInRow(row, acceptedTerms) {
    if (!Array.isArray(row)) return -1;

    return row.findIndex(cell => {
        const normalizedCell = normalizeName(cell);

        return acceptedTerms.some(term =>
            normalizedCell.includes(normalizeName(term))
        );
    });
}

function shouldIgnoreFarmRow(farm) {
    const normalized = normalizeName(farm);

    return (
        !normalized ||
        normalized === "TOTAL" ||
        normalized.includes("PENDENCIA") ||
        normalized.includes("OBSERVACAO") ||
        normalized.includes("LEGENDA")
    );
}


/* ============================================================
   DATAS IMPORTADAS
============================================================ */

function normalizeImportedDay(value) {
    if (value === undefined || value === null || value === "") {
        return "";
    }

    /* Data real produzida pelo SheetJS com cellDates:true. */
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        const day = String(value.getDate()).padStart(2, "0");
        const month = String(value.getMonth() + 1).padStart(2, "0");
        return `${day}/${month}`;
    }

    /* Número serial do Excel, caso a célula não tenha vindo como Date. */
    if (typeof value === "number" && Number.isFinite(value)) {
        const parsed = XLSX?.SSF?.parse_date_code?.(value);
        if (parsed && parsed.d && parsed.m) {
            return `${String(parsed.d).padStart(2, "0")}/${String(parsed.m).padStart(2, "0")}`;
        }
    }

    const text = String(value).trim();
    if (!text) return "";

    /* ISO: 2026-07-14 ou 2026-07-14T00:00:00. */
    let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
    if (match) {
        return `${String(match[3]).padStart(2, "0")}/${String(match[2]).padStart(2, "0")}`;
    }

    /* Formatos brasileiros: DD/MM, DD/MM/AAAA, DD-MM-AAAA. */
    match = text.match(/^(\d{1,2})[\/-](\d{1,2})(?:[\/-](\d{2,4}))?(?:\s+.*)?$/);
    if (match) {
        return `${String(match[1]).padStart(2, "0")}/${String(match[2]).padStart(2, "0")}`;
    }

    return text;
}

function uniqueDays(days) {
    const result = [];

    days.forEach(day => {
        const normalized = normalizeImportedDay(day);

        if (normalized && !result.includes(normalized)) {
            result.push(normalized);
        }
    });

    return result;
}

function isImportDateFilterActive() {
    return Boolean(
        state?.filterStart ||
        state?.filterEnd
    );
}

function getImportDateRange() {
    let startDate = isoToDate(state?.filterStart);
    let endDate = isoToDate(state?.filterEnd);

    if (endDate && !startDate) {
        startDate = new Date(endDate);
        startDate.setDate(startDate.getDate() - 6);
    }

    if (startDate && !endDate) {
        endDate = new Date(startDate);
        endDate.setDate(endDate.getDate() + 6);
    }

    if (startDate && endDate && startDate > endDate) {
        [startDate, endDate] = [endDate, startDate];
    }

    return {
        startDate,
        endDate
    };
}

function getFilteredImportedDays(days) {
    const normalizedDays = uniqueDays(days);

    if (!isImportDateFilterActive()) {
        return normalizedDays;
    }

    const {
        startDate,
        endDate
    } = getImportDateRange();

    return normalizedDays.filter(dayLabel => {
        const date = parseDayLabel(dayLabel);

        if (!date) return false;
        if (startDate && date < startDate) return false;
        if (endDate && date > endDate) return false;

        return true;
    });
}


/* ============================================================
   PARSER DOS PAINÉIS DIÁRIOS
============================================================ */

function parseStatusSheet(matrix) {
    if (!isMatrix(matrix) || !matrix.length) {
        return null;
    }

    let headerRow = -1;
    let farmColumn = -1;

    const periodHeader = findHeaderCell(matrix, ["PERIODO"]);

    if (periodHeader) {
        headerRow = periodHeader.rowIndex;
        farmColumn = periodHeader.columnIndex;
    }

    if (headerRow === -1) {
        const farmHeader = findHeaderCell(matrix, ["FAZENDA"]);

        if (farmHeader) {
            headerRow = farmHeader.rowIndex;
            farmColumn = farmHeader.columnIndex;
        }
    }

    if (headerRow === -1 || farmColumn === -1) {
        return null;
    }

    const header = matrix[headerRow] || [];
    const days = [];

    for (
        let columnIndex = farmColumn + 1;
        columnIndex < header.length;
        columnIndex++
    ) {
        const value = header[columnIndex];
        const normalizedDay = normalizeImportedDay(value);

        if (!normalizedDay) {
            /* Permite células vazias no meio, mas para após duas vazias. */
            const nextValue = normalizeImportedDay(header[columnIndex + 1]);
            if (!nextValue) break;
            continue;
        }

        /* Para antes de colunas de total/pendência. */
        const normalizedHeader = normalizeName(value);
        if (
            normalizedHeader.includes("PENDENCIA") ||
            normalizedHeader.includes("PENDENTE") ||
            normalizedHeader.includes("DIAS RECEBIDOS") ||
            normalizedHeader.includes("RECEBIDOS") ||
            normalizedHeader.includes("TOTAL") ||
            normalizedHeader.includes("STATUS")
        ) {
            break;
        }

        days.push(normalizedDay);
    }

    if (!days.length) {
        return null;
    }

    const rows = [];

    for (
        let rowIndex = headerRow + 1;
        rowIndex < matrix.length;
        rowIndex++
    ) {
        const row = matrix[rowIndex] || [];
        const farm = String(row[farmColumn] ?? "").trim();

        if (shouldIgnoreFarmRow(farm)) continue;

        const statuses = days.map((day, dayIndex) =>
            cellToStatus(row[farmColumn + 1 + dayIndex])
        );

        rows.push({
            farm,
            statuses
        });
    }

    if (!rows.length) {
        return null;
    }

    return {
        days,
        rows
    };
}

function mergeParsedDailySheets(key) {
    const parsedSheets = getImportedMatrices(key)
        .map(matrix => parseStatusSheet(matrix))
        .filter(Boolean);

    if (!parsedSheets.length) {
        return null;
    }

    const days = uniqueDays(
        parsedSheets.flatMap(sheet => sheet.days)
    );

    const farms = new Map();

    parsedSheets.forEach(sheet => {
        sheet.rows.forEach(row => {
            const farmKey = normalizeName(row.farm);

            if (!farms.has(farmKey)) {
                farms.set(farmKey, {
                    farm: row.farm,
                    statusByDay: {}
                });
            }

            const farmData = farms.get(farmKey);

            sheet.days.forEach((day, dayIndex) => {
                farmData.statusByDay[
                    normalizeImportedDay(day)
                ] = normalizeStatus(row.statuses[dayIndex]);
            });
        });
    });

    return {
        days,
        rows: [...farms.values()].map(farmData => ({
            farm: farmData.farm,
            statuses: days.map(day =>
                farmData.statusByDay[day] ?? "blank"
            )
        }))
    };
}


/* ============================================================
   PARSER DO MAPA MENSAL
============================================================ */

function parseMonthlySheet(matrix) {
    if (!isMatrix(matrix) || !matrix.length) {
        return null;
    }

    let headerRow = -1;
    let farmColumn = -1;
    let statusColumn = -1;

    for (let rowIndex = 0; rowIndex < matrix.length; rowIndex++) {
        const row = matrix[rowIndex] || [];
        const possibleFarmColumn = findColumnInRow(row, ["FAZENDA"]);

        if (possibleFarmColumn === -1) continue;

        headerRow = rowIndex;
        farmColumn = possibleFarmColumn;
        statusColumn = findColumnInRow(row, [
            "STATUS",
            "SITUACAO",
            "ENVIO",
            "RECEBIDO",
            "FECHAMENTO"
        ]);

        if (statusColumn === -1) {
            statusColumn = farmColumn + 1;
        }

        break;
    }

    if (
        headerRow === -1 ||
        farmColumn === -1 ||
        statusColumn === -1
    ) {
        return null;
    }

    const rows = [];

    for (
        let rowIndex = headerRow + 1;
        rowIndex < matrix.length;
        rowIndex++
    ) {
        const row = matrix[rowIndex] || [];
        const farm = String(row[farmColumn] ?? "").trim();

        if (shouldIgnoreFarmRow(farm)) continue;

        rows.push({
            farm,
            status: cellToStatus(row[statusColumn])
        });
    }

    return rows.length ? rows : null;
}

function mergeParsedMonthlySheets() {
    const rows = getImportedMatrices("mensal")
        .flatMap(matrix => parseMonthlySheet(matrix) || []);

    if (!rows.length) return null;

    const farms = new Map();

    rows.forEach(row => {
        farms.set(normalizeName(row.farm), row);
    });

    return [...farms.values()];
}


/* ============================================================
   PARSER DE DIVERGÊNCIAS
============================================================ */

function parseDivergSheet(matrix) {
    if (!isMatrix(matrix) || !matrix.length) {
        return null;
    }

    let headerRow = -1;
    let farmColumn = -1;

    for (let rowIndex = 0; rowIndex < matrix.length; rowIndex++) {
        const row = matrix[rowIndex] || [];
        const possibleFarmColumn = findColumnInRow(row, ["FAZENDA"]);

        if (possibleFarmColumn !== -1) {
            headerRow = rowIndex;
            farmColumn = possibleFarmColumn;
            break;
        }
    }

    if (headerRow === -1 || farmColumn === -1) {
        return null;
    }

    const header = matrix[headerRow] || [];

    let nascDiarioColumn = findColumnInRow(header, [
        "NASC DIARIO",
        "NASCIMENTO DIARIO"
    ]);

    let nascSistemaColumn = findColumnInRow(header, [
        "NASC SISTEMA",
        "NASCIMENTO SISTEMA"
    ]);

    let mortesDiarioColumn = findColumnInRow(header, [
        "MORTES DIARIO",
        "MORTE DIARIO"
    ]);

    let mortesSistemaColumn = findColumnInRow(header, [
        "MORTES SISTEMA",
        "MORTE SISTEMA"
    ]);

    /* Compatibilidade com a estrutura antiga por posição. */
    if (nascDiarioColumn === -1) nascDiarioColumn = farmColumn + 1;
    if (nascSistemaColumn === -1) nascSistemaColumn = farmColumn + 2;
    if (mortesDiarioColumn === -1) mortesDiarioColumn = farmColumn + 4;
    if (mortesSistemaColumn === -1) mortesSistemaColumn = farmColumn + 5;

    const rows = [];

    for (
        let rowIndex = headerRow + 1;
        rowIndex < matrix.length;
        rowIndex++
    ) {
        const row = matrix[rowIndex] || [];
        const farm = String(row[farmColumn] ?? "").trim();

        if (shouldIgnoreFarmRow(farm)) continue;

        rows.push({
            farm,
            nd: toNumberOrEmpty(row[nascDiarioColumn]),
            ns: toNumberOrEmpty(row[nascSistemaColumn]),
            md: toNumberOrEmpty(row[mortesDiarioColumn]),
            ms: toNumberOrEmpty(row[mortesSistemaColumn])
        });
    }

    return rows.length ? rows : null;
}

function mergeParsedDivergenceSheets() {
    const rows = getImportedMatrices("divergencias")
        .flatMap(matrix => parseDivergSheet(matrix) || []);

    if (!rows.length) return null;

    const farms = new Map();

    rows.forEach(row => {
        farms.set(normalizeName(row.farm), row);
    });

    return [...farms.values()];
}


/* ============================================================
   AJUSTE DO EIXO DE DIAS
============================================================ */

function remapDailyDataToDays(newDays) {
    const previousDays = Array.isArray(state.days)
        ? [...state.days]
        : [];

    const normalizedNewDays = uniqueDays(newDays);

    PANEL_DEFS
        .filter(panel => panel.type === "daily")
        .forEach(panel => {
            const previousPanelData = state.data?.[panel.key] || {};
            const remappedPanelData = {};

            state.farms.forEach((farm, farmIndex) => {
                remappedPanelData[farmIndex] = {};

                normalizedNewDays.forEach((newDay, newDayIndex) => {
                    const previousDayIndex = previousDays.findIndex(
                        previousDay =>
                            normalizeImportedDay(previousDay) === newDay
                    );

                    const previousStatus =
                        previousDayIndex === -1
                            ? "blank"
                            : previousPanelData?.[farmIndex]?.[
                                previousDayIndex
                            ];

                    remappedPanelData[farmIndex][newDayIndex] =
                        previousStatus === undefined
                            ? "blank"
                            : normalizeStatus(previousStatus);
                });
            });

            state.data[panel.key] = remappedPanelData;
        });

    state.days = normalizedNewDays;
}


/* ============================================================
   APLICAÇÃO DOS DADOS NO STATE
============================================================ */

function applyDailySheet(key, parsed, selectedDays, unmatched) {
    if (!parsed || !selectedDays.length) return false;

    state.data[key] = {};

    state.farms.forEach((farm, farmIndex) => {
        state.data[key][farmIndex] = {};

        state.days.forEach((day, dayIndex) => {
            state.data[key][farmIndex][dayIndex] = "blank";
        });
    });

    parsed.rows.forEach(row => {
        const farmIndex = matchFarmIndex(row.farm);

        if (farmIndex === -1) {
            unmatched.add(row.farm);
            return;
        }

        parsed.days.forEach((importedDay, importedDayIndex) => {
            const normalizedDay = normalizeImportedDay(importedDay);

            if (!selectedDays.includes(normalizedDay)) return;

            const stateDayIndex = state.days.findIndex(
                day => normalizeImportedDay(day) === normalizedDay
            );

            if (stateDayIndex === -1) return;

            state.data[key][farmIndex][stateDayIndex] = normalizeStatus(
                row.statuses[importedDayIndex]
            );
        });
    });

    return true;
}

function applyMonthlySheet(parsed, unmatched) {
    if (!parsed?.length) return false;

    const newMonthly = {};

    state.farms.forEach((farm, farmIndex) => {
        newMonthly[farmIndex] = "blank";
    });

    parsed.forEach(row => {
        const farmIndex = matchFarmIndex(row.farm);

        if (farmIndex === -1) {
            unmatched.add(row.farm);
            return;
        }

        newMonthly[farmIndex] = normalizeStatus(row.status);
    });

    state.monthly = newMonthly;

    return true;
}

function applyDivergSheet(parsed, unmatched) {
    if (!parsed?.length) return false;

    parsed.forEach(row => {
        const farmIndex = matchFarmIndex(row.farm);

        if (farmIndex === -1) {
            unmatched.add(row.farm);
            return;
        }

        state.diverg[farmIndex] = {
            nd: row.nd === "" ? "" : String(row.nd),
            ns: row.ns === "" ? "" : String(row.ns),
            md: row.md === "" ? "" : String(row.md),
            ms: row.ms === "" ? "" : String(row.ms)
        };
    });

    return true;
}


/* ============================================================
   PREENCHIMENTO DO SISTEMA
============================================================ */

function preencherSistema(options = {}) {
    const {
        silent = false
    } = options;

    const importedKeys = EXCEL_PANEL_KEYS.filter(hasImportedData);

    if (!importedKeys.length) {
        const result = {
            found: [],
            failed: [],
            unmatched: [],
            noDatesInFilter: false
        };

        if (!silent) {
            showInlineWarning(
                "Nenhuma aba reconhecida foi encontrada no arquivo."
            );
        }

        return result;
    }

    const parsed = {
        campo: hasImportedData("campo")
            ? mergeParsedDailySheets("campo")
            : null,

        abastecimento: hasImportedData("abastecimento")
            ? mergeParsedDailySheets("abastecimento")
            : null,

        diario: hasImportedData("diario")
            ? mergeParsedDailySheets("diario")
            : null,

        mensal: hasImportedData("mensal")
            ? mergeParsedMonthlySheets()
            : null,

        divergencias: hasImportedData("divergencias")
            ? mergeParsedDivergenceSheets()
            : null
    };

    const found = [];
    const failed = [];
    const unmatched = new Set();

    const dailyKeys = [
        "campo",
        "abastecimento",
        "diario"
    ];

    const dailyParsedSheets = dailyKeys
        .map(key => parsed[key])
        .filter(Boolean);

    const availableDailyDays = uniqueDays(
        dailyParsedSheets.flatMap(sheet => sheet.days)
    );

    const selectedDays = getFilteredImportedDays(availableDailyDays);
    const noDatesInFilter = (
        isImportDateFilterActive() &&
        availableDailyDays.length > 0 &&
        selectedDays.length === 0
    );

    if (selectedDays.length) {
        remapDailyDataToDays(selectedDays);
    }

    dailyKeys.forEach(key => {
        if (!hasImportedData(key)) return;

        if (noDatesInFilter) {
            failed.push(key);
            return;
        }

        if (
            applyDailySheet(
                key,
                parsed[key],
                selectedDays,
                unmatched
            )
        ) {
            found.push(key);
        } else {
            failed.push(key);
        }
    });

    if (hasImportedData("mensal")) {
        if (applyMonthlySheet(parsed.mensal, unmatched)) {
            found.push("mensal");
        } else {
            failed.push("mensal");
        }
    }

    if (hasImportedData("divergencias")) {
        if (applyDivergSheet(parsed.divergencias, unmatched)) {
            found.push("divergencias");
        } else {
            failed.push("divergencias");
        }
    }

    ensureData();

    if (typeof renderAll === "function") {
        renderAll();
    }

    if (typeof saveState === "function") {
        saveState();
    }

    clearImportedData();

    const result = {
        found,
        failed,
        unmatched: [...unmatched],
        noDatesInFilter
    };

    if (!silent) {
        showInlineWarning(buildImportMessage(result));
    }

    return result;
}

function buildImportMessage(result, extraErrors = [], ignoredSheets = []) {
    let message = result.found.length
        ? "Importado: " +
          result.found
              .map(key => EXCEL_PANEL_LABELS[key])
              .join(", ") +
          "."
        : "Nenhuma planilha foi importada.";

    if (result.noDatesInFilter) {
        message +=
            " As abas diárias não possuem datas dentro do período selecionado.";
    }

    if (result.failed.length) {
        message +=
            " Estrutura não reconhecida ou sem dados válidos em: " +
            [...new Set(result.failed)]
                .map(key => EXCEL_PANEL_LABELS[key])
                .join(", ") +
            ".";
    }

    if (result.unmatched.length) {
        message +=
            " Fazendas sem correspondência: " +
            result.unmatched.join(", ") +
            ".";
    }

    if (ignoredSheets.length) {
        message +=
            " Abas ignoradas: " +
            [...new Set(ignoredSheets)].join(", ") +
            ".";
    }

    if (extraErrors.length) {
        message += " Erros: " + extraErrors.join(" | ");
    }

    return message;
}


/* ============================================================
   IMPORTAÇÃO LOCAL
============================================================ */

async function importLocalExcelFiles(event) {
    const input = event.target;
    const files = [...(input.files || [])];

    if (!files.length) return;

    clearImportedData();

    const errors = [];
    const ignoredSheets = [];
    let recognizedSomething = false;

    for (const file of files) {
        try {
            const buffer = await file.arrayBuffer();

            const workbook = XLSX.read(buffer, {
                type: "array",
                cellDates: true
            });

            const result = processWorkbook(workbook, file.name);

            if (result.recognized.length) {
                recognizedSomething = true;

                registerImportedLocalFile(
                    file.name,
                    result.recognizedSheets,
                    result.ignored
                );
            }

            ignoredSheets.push(
                ...result.ignored.map(sheetName =>
                    `${file.name} → ${sheetName}`
                )
            );
        } catch (error) {
            console.error(`Erro ao importar ${file.name}:`, error);

            errors.push(
                `${file.name}: ${error?.message || "erro desconhecido"}`
            );
        }
    }

    /* Permite escolher novamente o mesmo arquivo. */
    input.value = "";

    if (!recognizedSomething) {
        clearImportedData();

        let message =
            "Nenhuma aba reconhecida foi encontrada nos arquivos selecionados.";

        if (ignoredSheets.length) {
            message +=
                " Abas encontradas: " +
                ignoredSheets.join(", ") +
                ".";
        }

        if (errors.length) {
            message += " Erros: " + errors.join(" | ");
        }

        showInlineWarning(message);
        return;
    }

    const result = preencherSistema({
        silent: true
    });

    if (typeof saveState === "function") {
        saveState();
    }

    renderImportedWorkbookStatus();

    showInlineWarning(
        buildImportMessage(result, errors, ignoredSheets)
    );
}


/* ============================================================
   EVENTO DO INPUT LOCAL
============================================================ */

function bindExcelImportEvents() {
    const input = document.getElementById("excelFiles");

    if (!input || input.dataset.bound === "true") {
        return;
    }

    input.dataset.bound = "true";
    input.addEventListener("change", importLocalExcelFiles);
}

document.addEventListener("DOMContentLoaded", () => {
    bindExcelImportEvents();
    installImportedWorkbookStatusIntegration();
});
