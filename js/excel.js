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

const LOCAL_WORKBOOK_DB = "controle-workbooks-v1";
const LOCAL_WORKBOOK_STORE = "files";
let importedRefreshPromise = null;

function openLocalWorkbookDatabase() {
    return new Promise((resolve, reject) => {
        if (!("indexedDB" in window)) { resolve(null); return; }
        const request = indexedDB.open(LOCAL_WORKBOOK_DB, 1);
        request.onupgradeneeded = () => {
            const db = request.result;
            if (!db.objectStoreNames.contains(LOCAL_WORKBOOK_STORE)) {
                db.createObjectStore(LOCAL_WORKBOOK_STORE, { keyPath: "id" });
            }
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function saveLocalWorkbookBuffer(record) {
    const db = await openLocalWorkbookDatabase();
    if (!db) return;
    await new Promise((resolve, reject) => {
        const tx = db.transaction(LOCAL_WORKBOOK_STORE, "readwrite");
        tx.objectStore(LOCAL_WORKBOOK_STORE).put(record);
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
    });
    db.close();
}

async function readLocalWorkbookBuffers() {
    const db = await openLocalWorkbookDatabase();
    if (!db) return [];
    const rows = await new Promise((resolve, reject) => {
        const tx = db.transaction(LOCAL_WORKBOOK_STORE, "readonly");
        const request = tx.objectStore(LOCAL_WORKBOOK_STORE).getAll();
        request.onsuccess = () => resolve(request.result || []);
        request.onerror = () => reject(request.error);
    });
    db.close();
    return rows;
}

async function deleteLocalWorkbookBuffer(id) {
    const db = await openLocalWorkbookDatabase();
    if (!db) return;
    await new Promise((resolve, reject) => {
        const tx = db.transaction(LOCAL_WORKBOOK_STORE, "readwrite");
        tx.objectStore(LOCAL_WORKBOOK_STORE).delete(id);
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
    });
    db.close();
}





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
    deleteLocalWorkbookBuffer(file.id).catch(error =>
        console.warn("Não foi possível remover o arquivo do cache:", error)
    );

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

const EXCEL_ALLOWED_SHEET_NAMES = {
    campo: [
        "OPERACOES EM CAMPO",
        "OPERACAO EM CAMPO",
        "CAMPO",
        "SISFROTA",
        "RESUMO OPERACOES SISFROTA"
    ],

    abastecimento: [
        "ABASTECIMENTO",
        "CONTROLE DE ABASTECIMENTO"
    ],

    diario: [
        "DIARIO DE CAMPO",
        "DIARIO",
        "MANEJO REBANHO",
        "DIARIO DE CAMPO MANEJO REBANHO"
    ],

    mensal: [
        "MAPA MENSAL DO REBANHO",
        "MAPA MENSAL",
        "FECHAMENTO MENSAL",
        "MENSAL"
    ],

    divergencias: [
        "DIVERGENCIAS",
        "DIVERGENCIA",
        "DIVERGENCIAS ENTRE DIARIO E SISTEMA"
    ]
};

function classifySheetName(name) {
    const normalized = normalizeName(name);

    if (!normalized) return null;

    for (const [key, allowedNames] of Object.entries(
        EXCEL_ALLOWED_SHEET_NAMES
    )) {
        if (allowedNames.includes(normalized)) {
            return key;
        }
    }

    return null;
}

function isAllowedExcelSheetName(name) {
    return Boolean(classifySheetName(name));
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

        /*
         * Importação estrita: somente nomes de abas autorizados.
         * Não tenta adivinhar pelo conteúdo e não usa o nome do arquivo
         * como atalho. Isso impede que abas auxiliares, como VJ, EC,
         * LISTA ou RESUMO ENTRADA E SAIDA, sejam importadas por engano.
         */
        const key = classifySheetName(sheetName);

        if (!key) {
            ignored.push(sheetName);
            console.info(`Aba ignorada por nome: ${sheetName}`);
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

function parseExcelDateValue(value) {
    if (
        value === undefined ||
        value === null ||
        value === ""
    ) {
        return null;
    }

    if (
        value instanceof Date &&
        !Number.isNaN(value.getTime())
    ) {
        return new Date(
            value.getFullYear(),
            value.getMonth(),
            value.getDate()
        );
    }

    if (
        typeof value === "number" &&
        Number.isFinite(value)
    ) {
        const parsed =
            XLSX?.SSF?.parse_date_code?.(value);

        if (
            parsed &&
            parsed.y &&
            parsed.m &&
            parsed.d
        ) {
            return new Date(
                parsed.y,
                parsed.m - 1,
                parsed.d
            );
        }
    }

    const text =
        String(value).trim();

    if (!text) return null;

    let match = text.match(
        /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/
    );

    if (match) {
        return new Date(
            Number(match[1]),
            Number(match[2]) - 1,
            Number(match[3])
        );
    }

    match = text.match(
        /^(\d{1,2})[\/-](\d{1,2})(?:[\/-](\d{2,4}))?$/
    );

    if (match) {
        let year = match[3]
            ? Number(match[3])
            : new Date().getFullYear();

        if (year < 100) {
            year += 2000;
        }

        return new Date(
            year,
            Number(match[2]) - 1,
            Number(match[1])
        );
    }

    return null;
}

function getMonthlyReferenceDate() {
    const filterEnd =
        isoToDate(state?.filterEnd);

    if (filterEnd) {
        return filterEnd;
    }

    const filterStart =
        isoToDate(state?.filterStart);

    if (filterStart) {
        return filterStart;
    }

    const periodText =
        String(state?.period || "");

    const matches = [
        ...periodText.matchAll(
            /(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?/g
        )
    ];

    if (matches.length) {
        const last =
            matches[matches.length - 1];

        let year = last[3]
            ? Number(last[3])
            : new Date().getFullYear();

        if (year < 100) {
            year += 2000;
        }

        return new Date(
            year,
            Number(last[2]) - 1,
            Number(last[1])
        );
    }

    return new Date();
}

function formatMonthlyLabel(date) {
    if (
        !(date instanceof Date) ||
        Number.isNaN(date.getTime())
    ) {
        return "";
    }

    const monthName =
        date
            .toLocaleDateString(
                "pt-BR",
                {
                    month: "long"
                }
            )
            .toUpperCase();

    return `FECHAMENTO – ${monthName}/${date.getFullYear()}`;
}

function countMeaningfulMonthlyStatuses(
    matrix,
    firstDataRow,
    farmColumn,
    statusColumn
) {
    let score = 0;

    for (
        let rowIndex = firstDataRow;
        rowIndex < matrix.length;
        rowIndex++
    ) {
        const row =
            matrix[rowIndex] || [];

        const farm =
            String(
                row[farmColumn] ?? ""
            ).trim();

        if (shouldIgnoreFarmRow(farm)) {
            continue;
        }

        const normalized =
            normalizeName(
                row[statusColumn]
            );

        if (
            normalized.includes("RECEB") ||
            normalized.includes("PEND") ||
            normalized.includes("ENVIAD") ||
            normalized.includes("CONCLUID") ||
            normalized.includes("OK")
        ) {
            score++;
        }
    }

    return score;
}

function parseMonthlyPeriodLayout(matrix) {
    const periodHeader =
        findHeaderCell(
            matrix,
            ["PERIODO"]
        );

    if (!periodHeader) {
        return null;
    }

    const headerRow =
        periodHeader.rowIndex;

    const farmColumn =
        periodHeader.columnIndex;

    const header =
        matrix[headerRow] || [];

    const referenceDate =
        getMonthlyReferenceDate();

    const targetYear =
        referenceDate.getFullYear();

    const targetMonth =
        referenceDate.getMonth();

    const candidateColumns = [];

    for (
        let columnIndex = farmColumn + 1;
        columnIndex < header.length;
        columnIndex++
    ) {
        const date =
            parseExcelDateValue(
                header[columnIndex]
            );

        if (!date) continue;

        if (
            date.getFullYear() !== targetYear ||
            date.getMonth() !== targetMonth
        ) {
            continue;
        }

        candidateColumns.push({
            columnIndex,
            date,
            score:
                countMeaningfulMonthlyStatuses(
                    matrix,
                    headerRow + 1,
                    farmColumn,
                    columnIndex
                )
        });
    }

    if (!candidateColumns.length) {
        return {
            rows: [],
            referenceDate,
            monthFound: false
        };
    }

    /*
     * Quando existem várias datas do mesmo mês, usa a coluna
     * que realmente contém mais status preenchidos.
     */
    candidateColumns.sort(
        (first, second) =>
            second.score - first.score ||
            first.columnIndex - second.columnIndex
    );

    const selected =
        candidateColumns[0];

    const rows = [];

    for (
        let rowIndex = headerRow + 1;
        rowIndex < matrix.length;
        rowIndex++
    ) {
        const row =
            matrix[rowIndex] || [];

        const farm =
            String(
                row[farmColumn] ?? ""
            ).trim();

        if (shouldIgnoreFarmRow(farm)) {
            continue;
        }

        rows.push({
            farm,
            status:
                cellToStatus(
                    row[selected.columnIndex]
                )
        });
    }

    return {
        rows,
        referenceDate:
            selected.date,
        monthFound: true,
        selectedColumn:
            selected.columnIndex
    };
}

function parseMonthlySimpleLayout(matrix) {
    let headerRow = -1;
    let farmColumn = -1;
    let statusColumn = -1;

    for (
        let rowIndex = 0;
        rowIndex < matrix.length;
        rowIndex++
    ) {
        const row =
            matrix[rowIndex] || [];

        const possibleFarmColumn =
            findColumnInRow(
                row,
                ["FAZENDA"]
            );

        if (
            possibleFarmColumn === -1
        ) {
            continue;
        }

        headerRow = rowIndex;
        farmColumn =
            possibleFarmColumn;

        statusColumn =
            findColumnInRow(
                row,
                [
                    "STATUS",
                    "SITUACAO",
                    "ENVIO",
                    "RECEBIDO",
                    "FECHAMENTO"
                ]
            );

        if (statusColumn === -1) {
            statusColumn =
                farmColumn + 1;
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
        const row =
            matrix[rowIndex] || [];

        const farm =
            String(
                row[farmColumn] ?? ""
            ).trim();

        if (shouldIgnoreFarmRow(farm)) {
            continue;
        }

        rows.push({
            farm,
            status:
                cellToStatus(
                    row[statusColumn]
                )
        });
    }

    return rows.length
        ? {
            rows,
            referenceDate:
                getMonthlyReferenceDate(),
            monthFound: true
        }
        : null;
}

function parseMonthlySheet(matrix) {
    if (
        !isMatrix(matrix) ||
        !matrix.length
    ) {
        return null;
    }

    const periodLayout =
        parseMonthlyPeriodLayout(
            matrix
        );

    if (
        periodLayout &&
        periodLayout.monthFound
    ) {
        return periodLayout;
    }

    const simpleLayout =
        parseMonthlySimpleLayout(
            matrix
        );

    if (simpleLayout) {
        return simpleLayout;
    }

    return periodLayout;
}

function mergeParsedMonthlySheets() {
    const parsedSheets =
        getImportedMatrices("mensal")
            .map(matrix =>
                parseMonthlySheet(matrix)
            )
            .filter(Boolean);

    if (!parsedSheets.length) {
        return null;
    }

    const validSheets =
        parsedSheets.filter(
            parsed =>
                parsed.monthFound &&
                Array.isArray(parsed.rows) &&
                parsed.rows.length
        );

    if (!validSheets.length) {
        return {
            rows: [],
            referenceDate:
                getMonthlyReferenceDate(),
            monthFound: false
        };
    }

    const farms =
        new Map();

    validSheets.forEach(parsed => {
        parsed.rows.forEach(row => {
            farms.set(
                normalizeName(row.farm),
                row
            );
        });
    });

    const referenceDate =
        validSheets.find(
            parsed =>
                parsed.referenceDate
        )?.referenceDate ||
        getMonthlyReferenceDate();

    return {
        rows:
            [...farms.values()],
        referenceDate,
        monthFound: true
    };
}


/* ============================================================
   PARSER DE DIVERGÊNCIAS
============================================================ */

function getDivergenceFilterMonthKeys() {
    const keys = new Set();

    let startDate =
        isoToDate(state?.filterStart);

    let endDate =
        isoToDate(state?.filterEnd);

    if (!startDate && !endDate) {
        return keys;
    }

    if (!startDate) {
        startDate =
            new Date(endDate);
    }

    if (!endDate) {
        endDate =
            new Date(startDate);
    }

    if (startDate > endDate) {
        [startDate, endDate] =
            [endDate, startDate];
    }

    /*
     * Cria uma chave para cada mês atravessado pelo filtro.
     * Exemplo: 28/06 a 04/07 aceita JUNHO e JULHO.
     */
    const cursor =
        new Date(
            startDate.getFullYear(),
            startDate.getMonth(),
            1
        );

    const lastMonth =
        new Date(
            endDate.getFullYear(),
            endDate.getMonth(),
            1
        );

    while (cursor <= lastMonth) {
        keys.add(
            `${cursor.getFullYear()}-${String(
                cursor.getMonth() + 1
            ).padStart(2, "0")}`
        );

        cursor.setMonth(
            cursor.getMonth() + 1
        );
    }

    return keys;
}

function parseDivergenceDate(value) {
    if (
        value === undefined ||
        value === null ||
        value === ""
    ) {
        return null;
    }

    if (
        value instanceof Date &&
        !Number.isNaN(value.getTime())
    ) {
        return new Date(
            value.getFullYear(),
            value.getMonth(),
            value.getDate()
        );
    }

    /*
     * Datas do Excel chegam como número serial.
     * Aceita somente a faixa plausível de datas modernas para não
     * confundir quantidades como 27, 59, 230 e 239 com datas.
     */
    if (
        typeof value === "number" &&
        Number.isFinite(value) &&
        value >= 20000 &&
        value <= 80000
    ) {
        const parsed =
            XLSX?.SSF?.parse_date_code?.(value);

        if (
            parsed &&
            parsed.y &&
            parsed.m &&
            parsed.d
        ) {
            return new Date(
                parsed.y,
                parsed.m - 1,
                parsed.d
            );
        }
    }

    const text =
        String(value).trim();

    if (!text) return null;

    let match =
        text.match(
            /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/
        );

    if (match) {
        return new Date(
            Number(match[1]),
            Number(match[2]) - 1,
            Number(match[3])
        );
    }

    match =
        text.match(
            /^(\d{1,2})[\/.-](\d{1,2})(?:[\/.-](\d{2,4}))?$/
        );

    if (match) {
        let year =
            match[3]
                ? Number(match[3])
                : (
                    isoToDate(state?.filterEnd) ||
                    isoToDate(state?.filterStart) ||
                    new Date()
                ).getFullYear();

        if (year < 100) {
            year += 2000;
        }

        return new Date(
            year,
            Number(match[2]) - 1,
            Number(match[1])
        );
    }

    return null;
}

function getDivergenceMonthKey(date) {
    if (
        !(date instanceof Date) ||
        Number.isNaN(date.getTime())
    ) {
        return "";
    }

    return `${date.getFullYear()}-${String(
        date.getMonth() + 1
    ).padStart(2, "0")}`;
}

function parseDivergenceDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    if (typeof value === "number" && Number.isFinite(value) && value >= 20000 && value <= 80000) {
        const parsed = XLSX?.SSF?.parse_date_code?.(value);
        if (parsed?.y && parsed?.m && parsed?.d) return new Date(parsed.y, parsed.m - 1, parsed.d);
    }
    const text = String(value ?? "").trim();
    let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
    if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    match = text.match(/^(\d{1,2})[\/.\-](\d{1,2})(?:[\/.\-](\d{2,4}))?$/);
    if (!match) return null;
    let year = match[3] ? Number(match[3]) : (isoToDate(state?.filterEnd) || isoToDate(state?.filterStart) || new Date()).getFullYear();
    if (year < 100) year += 2000;
    const date = new Date(year, Number(match[2]) - 1, Number(match[1]));
    return Number.isNaN(date.getTime()) ? null : date;
}

function divergenceDateMatchesFilter(date) {
    if (!state?.filterStart && !state?.filterEnd) return true;
    if (!date) return false;
    const { startDate, endDate } = getImportDateRange();
    const monthStart = new Date(date.getFullYear(), date.getMonth(), 1);
    const monthEnd = new Date(date.getFullYear(), date.getMonth() + 1, 0);
    return (!startDate || monthEnd >= startDate) && (!endDate || monthStart <= endDate);
}

function parseDivergSheet(matrix) {
    if (!isMatrix(matrix) || !matrix.length) return null;
    let headerRow = -1;
    let farmColumn = -1;
    for (let rowIndex = 0; rowIndex < matrix.length; rowIndex++) {
        const possibleFarmColumn = findColumnInRow(matrix[rowIndex] || [], ["FAZENDA"]);
        if (possibleFarmColumn !== -1) { headerRow = rowIndex; farmColumn = possibleFarmColumn; break; }
    }
    if (headerRow === -1) return null;
    const header = matrix[headerRow] || [];
    let dateColumn = findColumnInRow(header, ["DATA", "REFERENCIA", "REFERÊNCIA", "MES", "MÊS"]);
    if (dateColumn === -1 && farmColumn > 0) dateColumn = farmColumn - 1;
    let nd = findColumnInRow(header, ["NASC DIARIO", "NASCIMENTO DIARIO", "NASCIMENTOS DIARIO"]);
    let ns = findColumnInRow(header, ["NASC SISTEMA", "NASCIMENTO SISTEMA", "NASCIMENTOS SISTEMA"]);
    let md = findColumnInRow(header, ["MORTES DIARIO", "MORTE DIARIO"]);
    let ms = findColumnInRow(header, ["MORTES SISTEMA", "MORTE SISTEMA"]);
    if (nd === -1) nd = farmColumn + 1;
    if (ns === -1) ns = farmColumn + 2;
    if (md === -1) md = farmColumn + 4;
    if (ms === -1) ms = farmColumn + 5;

    const rows = [];
    for (let rowIndex = headerRow + 1; rowIndex < matrix.length; rowIndex++) {
        const row = matrix[rowIndex] || [];
        const farm = String(row[farmColumn] ?? "").trim();
        if (shouldIgnoreFarmRow(farm)) continue;
        const referenceDate = dateColumn >= 0 ? parseDivergenceDate(row[dateColumn]) : null;
        if (!divergenceDateMatchesFilter(referenceDate)) continue;
        rows.push({ farm, nd: toNumberOrEmpty(row[nd]), ns: toNumberOrEmpty(row[ns]), md: toNumberOrEmpty(row[md]), ms: toNumberOrEmpty(row[ms]) });
    }
    return rows.length ? rows : null;
}

function mergeParsedDivergenceSheets() {
    const rows = getImportedMatrices("divergencias").flatMap(matrix => parseDivergSheet(matrix) || []);
    if (!rows.length) return null;
    const farms = new Map();
    rows.forEach(row => farms.set(normalizeName(row.farm), row));
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

    if (
        !state.dailyDataReady ||
        typeof state.dailyDataReady !== "object"
    ) {
        state.dailyDataReady = {};
    }

    state.dailyDataReady[key] = true;
    state.manualEntryEnabled = false;

    return true;
}

function applyMonthlySheet(parsed, unmatched) {
    if (
        !parsed ||
        !parsed.monthFound ||
        !Array.isArray(parsed.rows) ||
        !parsed.rows.length
    ) {
        return false;
    }

    const newMonthly = {};

    state.farms.forEach(
        (farm, farmIndex) => {
            newMonthly[farmIndex] =
                "blank";
        }
    );

    parsed.rows.forEach(row => {
        const farmIndex =
            matchFarmIndex(row.farm);

        if (farmIndex === -1) {
            unmatched.add(row.farm);
            return;
        }

        newMonthly[farmIndex] =
            normalizeStatus(
                row.status
            );
    });

    /*
     * Atualiza somente o painel mensal.
     * Nunca altera state.days nem os dados diários.
     */
    state.monthly =
        newMonthly;

    if (parsed.referenceDate) {
        state.monthLabel =
            formatMonthlyLabel(
                parsed.referenceDate
            );
    }

    return true;
}

function applyDivergSheet(parsed, unmatched) {
    /*
     * Sempre limpa primeiro para não manter valores do mês anterior.
     */
    state.diverg = {};

    state.farms.forEach((farm, farmIndex) => {
        state.diverg[farmIndex] = {
            nd: "",
            ns: "",
            md: "",
            ms: ""
        };
    });

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

function clearDailyPanel(key) {
    state.data[key] = {};
    state.farms.forEach((farm, farmIndex) => { state.data[key][farmIndex] = {}; });
    if (state.dailyDataReady) state.dailyDataReady[key] = false;
}

function clearMonthlyPanel() {
    state.monthly = {};
    state.farms.forEach((farm, farmIndex) => { state.monthly[farmIndex] = "blank"; });
    state.monthLabel = formatMonthlyLabel(getMonthlyReferenceDate());
}

function clearDivergencePanel() {
    state.diverg = {};
    state.farms.forEach((farm, farmIndex) => {
        state.diverg[farmIndex] = { nd: "", ns: "", md: "", ms: "" };
    });
}

function preencherSistema(options = {}) {
    const { silent = false, save = true, render = true } = options;
    const importedKeys = EXCEL_PANEL_KEYS.filter(hasImportedData);
    if (!importedKeys.length) {
        const result = { found: [], failed: [], unmatched: [], noDatesInFilter: true };
        if (!silent) showInlineWarning("Nenhuma fonte de planilha está disponível.");
        return result;
    }

    const parsed = {
        campo: hasImportedData("campo") ? mergeParsedDailySheets("campo") : null,
        abastecimento: hasImportedData("abastecimento") ? mergeParsedDailySheets("abastecimento") : null,
        diario: hasImportedData("diario") ? mergeParsedDailySheets("diario") : null,
        mensal: hasImportedData("mensal") ? mergeParsedMonthlySheets() : null,
        divergencias: hasImportedData("divergencias") ? mergeParsedDivergenceSheets() : null
    };

    const found = [];
    const failed = [];
    const unmatched = new Set();
    const dailyKeys = ["campo", "abastecimento", "diario"];
    const availableDailyDays = uniqueDays(dailyKeys.map(key => parsed[key]).filter(Boolean).flatMap(sheet => sheet.days));
    const selectedDays = getFilteredImportedDays(availableDailyDays);
    const noDatesInFilter = isImportDateFilterActive() && availableDailyDays.length > 0 && selectedDays.length === 0;

    state.days = selectedDays;
    dailyKeys.forEach(key => {
        if (!hasImportedData(key)) return;
        if (!parsed[key] || !selectedDays.length) { clearDailyPanel(key); failed.push(key); return; }
        if (applyDailySheet(key, parsed[key], selectedDays, unmatched)) found.push(key);
        else { clearDailyPanel(key); failed.push(key); }
    });

    if (hasImportedData("mensal")) {
        if (applyMonthlySheet(parsed.mensal, unmatched)) found.push("mensal");
        else { clearMonthlyPanel(); failed.push("mensal"); }
    }

    if (hasImportedData("divergencias")) {
        clearDivergencePanel();
        if (applyDivergSheet(parsed.divergencias, unmatched)) found.push("divergencias");
        else failed.push("divergencias");
    }

    ensureData();
    if (render && typeof renderAll === "function") renderAll();
    if (save && typeof saveState === "function") saveState();

    const result = { found, failed, unmatched: [...unmatched], noDatesInFilter };
    if (!silent) showInlineWarning(buildImportMessage(result));
    return result;
}

async function loadLocalWorkbooksIntoImportedData() {
    const records = await readLocalWorkbookBuffers();
    if (!records.length) return 0;
    clearImportedData();
    let count = 0;
    for (const record of records) {
        try {
            const workbook = XLSX.read(record.buffer, { type: "array", cellDates: true });
            processWorkbook(workbook, record.fileName || "");
            count++;
        } catch (error) {
            console.warn(`Não foi possível restaurar ${record.fileName || "planilha"}:`, error);
        }
    }
    return count;
}

async function reapplyImportedDataForCurrentFilter(options = {}) {
    if (importedRefreshPromise) return importedRefreshPromise;
    importedRefreshPromise = (async () => {
        if (!EXCEL_PANEL_KEYS.some(hasImportedData)) await loadLocalWorkbooksIntoImportedData();
        return preencherSistema({ silent: options.silent !== false, save: options.save !== false, render: true });
    })();
    try { return await importedRefreshPromise; }
    finally { importedRefreshPromise = null; }
}

async function refreshAllImportedSources(options = {}) {
    clearImportedData();
    await loadLocalWorkbooksIntoImportedData();
    return preencherSistema({ silent: options.silent !== false, save: options.save !== false, render: true });
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

    if (
        result.failed.includes("mensal")
    ) {
        message +=
            " O Mapa Mensal não possui dados do mesmo mês e ano do período selecionado.";
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

    const errors = [];
    const ignoredSheets = [];
    clearImportedData();
    let recognizedSomething = false;

    for (const file of files) {
        try {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
            const result = processWorkbook(workbook, file.name);
            if (result.recognized.length) {
                recognizedSomething = true;
                const existing = ensureImportedFilesState().find(item => normalizeName(item.fileName) === normalizeName(file.name));
                const id = existing?.id || `local-${Date.now()}-${Math.random().toString(36).slice(2,8)}`;
                registerImportedLocalFile(file.name, result.recognizedSheets, result.ignored);
                const registered = ensureImportedFilesState().find(item => normalizeName(item.fileName) === normalizeName(file.name));
                if (registered) registered.id = id;
                await saveLocalWorkbookBuffer({ id, fileName: file.name, lastModified: file.lastModified, size: file.size, buffer });
            }
            ignoredSheets.push(...result.ignored.map(sheetName => `${file.name} → ${sheetName}`));
        } catch (error) {
            console.error(`Erro ao importar ${file.name}:`, error);
            errors.push(`${file.name}: ${error?.message || "erro desconhecido"}`);
        }
    }

    input.value = "";
    if (!recognizedSomething) {
        showInlineWarning("Nenhuma aba reconhecida foi encontrada. " + errors.join(" | "));
        return;
    }

    const result = preencherSistema({ silent: true, save: true, render: true });
    renderImportedWorkbookStatus();
    showInlineWarning(buildImportMessage(result, errors, ignoredSheets));
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

document.addEventListener("DOMContentLoaded", async () => {
    bindExcelImportEvents();
    installImportedWorkbookStatusIntegration();
    try {
        if (!EXCEL_PANEL_KEYS.some(hasImportedData)) {
            await loadLocalWorkbooksIntoImportedData();
        }
    } catch (error) {
        console.warn("Não foi possível restaurar as planilhas locais:", error);
    }
});
