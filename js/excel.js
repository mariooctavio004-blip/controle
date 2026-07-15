/* ============================================================
   EXCEL.JS
   Leitura, reconhecimento e processamento das planilhas
   de Campo, Abastecimento, Diário, Mensal e Divergências.
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
   RECONHECIMENTO DAS PLANILHAS
============================================================ */

function classifySheetName(name) {
    const normalized = normalizeName(name);

    if (normalized.includes("DIVERG")) {
        return "divergencias";
    }

    if (
        normalized.includes("MAPA MENSAL") ||
        normalized.includes("MENSAL") ||
        normalized.includes("REBANHO")
    ) {
        return "mensal";
    }

    if (normalized.includes("DIARIO")) {
        return "diario";
    }

    if (
        normalized.includes("OPERA") ||
        normalized.includes("CAMPO")
    ) {
        return "campo";
    }

    if (normalized.includes("ABASTE")) {
        return "abastecimento";
    }

    return null;
}


/* ============================================================
   CONVERSÃO DE UMA ABA PARA MATRIZ
============================================================ */

function sheetToMatrix(sheet) {
    if (!sheet) return [];

    return XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
        defval: "",
        dateNF: "dd/mm/yyyy"
    });
}


/* ============================================================
   IDENTIFICAÇÃO DA ABA PELO CONTEÚDO
============================================================ */

function detectSheetKeyByContent(json) {
    if (!Array.isArray(json) || !json.length) {
        return null;
    }

    const sample = json
        .slice(0, 30)
        .flat()
        .map(normalizeName)
        .filter(Boolean)
        .join(" | ");

    if (
        sample.includes("DIF NASC") ||
        sample.includes("NASC SISTEMA") ||
        sample.includes("MORTES SISTEMA") ||
        (
            sample.includes("NASCIMENTO") &&
            sample.includes("SISTEMA")
        )
    ) {
        return "divergencias";
    }

    if (
        sample.includes("MAPA MENSAL") ||
        sample.includes("FECHAMENTO MENSAL") ||
        sample.includes("SALDO FINAL")
    ) {
        return "mensal";
    }

    if (
        sample.includes("ABASTECIMENTO") ||
        sample.includes("COMBUSTIVEL") ||
        sample.includes("LITROS")
    ) {
        return "abastecimento";
    }

    if (
        sample.includes("DIARIO DE CAMPO") ||
        sample.includes("MANEJO") ||
        sample.includes("REBANHO")
    ) {
        return "diario";
    }

    if (
        sample.includes("OPERACOES EM CAMPO") ||
        sample.includes("OPERACAO EM CAMPO")
    ) {
        return "campo";
    }

    return null;
}


/* ============================================================
   LEITURA DE TODAS AS ABAS
============================================================ */

function processWorkbook(workbook, fallbackName = "") {
    if (
        !workbook ||
        !Array.isArray(workbook.SheetNames) ||
        !workbook.SheetNames.length
    ) {
        throw new Error(
            "O arquivo informado não possui abas válidas."
        );
    }

    const recognized = [];

    workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];

        if (!sheet) return;

        const matrix = sheetToMatrix(sheet);

        if (!matrix.length) return;

        /*
         * Primeiro tenta identificar pelo nome da aba.
         */
        let key = classifySheetName(sheetName);

        /*
         * Depois tenta identificar pelo conteúdo da aba.
         * Isso permite ler duas ou mais abas dentro do mesmo arquivo,
         * mesmo quando elas possuem nomes genéricos.
         */
        if (!key) {
            key = detectSheetKeyByContent(matrix);
        }

        /*
         * Se o arquivo tiver apenas uma aba, ainda tenta usar
         * o nome do próprio arquivo.
         */
        if (
            !key &&
            workbook.SheetNames.length === 1
        ) {
            key = classifySheetName(fallbackName);
        }

        if (!key) {
            console.warn(
                `Aba não reconhecida: ${sheetName}`
            );

            return;
        }

        importedData[key] = matrix;

        if (!recognized.includes(key)) {
            recognized.push(key);
        }
    });

    return recognized;
}


/* ============================================================
   PROCURA DE CABEÇALHOS
============================================================ */

function findHeaderCell(json, acceptedTerms) {
    for (let rowIndex = 0; rowIndex < json.length; rowIndex++) {
        const row = json[rowIndex] || [];

        for (
            let columnIndex = 0;
            columnIndex < row.length;
            columnIndex++
        ) {
            const value = normalizeName(row[columnIndex]);

            const matched = acceptedTerms.some(term =>
                value.includes(term)
            );

            if (matched) {
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
        const normalized = normalizeName(cell);

        return acceptedTerms.some(term =>
            normalized.includes(term)
        );
    });
}


/* ============================================================
   NORMALIZAÇÃO DE DATAS DO CABEÇALHO
============================================================ */

function normalizeImportedDay(value) {
    if (
        value === undefined ||
        value === null ||
        String(value).trim() === ""
    ) {
        return "";
    }

    const text = String(value).trim();

    /*
     * Datas já formatadas como DD/MM ou DD/MM/AAAA.
     */
    const dateMatch = text.match(
        /^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/
    );

    if (dateMatch) {
        const day = String(dateMatch[1]).padStart(2, "0");
        const month = String(dateMatch[2]).padStart(2, "0");

        return `${day}/${month}`;
    }

    return text;
}


/* ============================================================
   PAINÉIS DIÁRIOS
============================================================ */

function parseStatusSheet(json) {
    if (!Array.isArray(json) || !json.length) {
        return null;
    }

    let headerRow = -1;
    let farmColumn = -1;

    /*
     * Primeiro procura a célula "PERÍODO".
     */
    const periodHeader = findHeaderCell(
        json,
        ["PERIODO"]
    );

    if (periodHeader) {
        headerRow = periodHeader.rowIndex;
        farmColumn = periodHeader.columnIndex;
    }

    /*
     * Caso não encontre "PERÍODO", procura um cabeçalho Fazenda.
     */
    if (headerRow === -1) {
        const farmHeader = findHeaderCell(
            json,
            ["FAZENDA"]
        );

        if (farmHeader) {
            headerRow = farmHeader.rowIndex;
            farmColumn = farmHeader.columnIndex;
        }
    }

    if (headerRow === -1 || farmColumn === -1) {
        return null;
    }

    const header = json[headerRow] || [];
    const days = [];

    for (
        let column = farmColumn + 1;
        column < header.length;
        column++
    ) {
        const value = header[column];

        if (
            value === undefined ||
            value === null ||
            String(value).trim() === ""
        ) {
            break;
        }

        const normalizedDay = normalizeImportedDay(value);

        if (normalizedDay) {
            days.push(normalizedDay);
        }
    }

    if (!days.length) {
        return null;
    }

    const rows = [];

    for (
        let rowIndex = headerRow + 1;
        rowIndex < json.length;
        rowIndex++
    ) {
        const row = json[rowIndex] || [];
        const farmValue = row[farmColumn];
        const farm = String(farmValue || "").trim();

        if (!farm) continue;

        /*
         * Ignora linhas de totais, observações e rodapés.
         */
        const normalizedFarm = normalizeName(farm);

        if (
            normalizedFarm === "TOTAL" ||
            normalizedFarm.includes("PENDENCIA") ||
            normalizedFarm.includes("OBSERVACAO")
        ) {
            continue;
        }

        const statuses = days.map((day, index) => {
            const value = row[farmColumn + 1 + index];

            return cellToStatus(value);
        });

        rows.push({
            farm,
            statuses
        });
    }

    return {
        days,
        rows
    };
}


/* ============================================================
   MAPA MENSAL
============================================================ */

function parseMonthlySheet(json) {
    if (!Array.isArray(json) || !json.length) {
        return null;
    }

    let headerRow = -1;
    let farmColumn = -1;
    let statusColumn = -1;

    for (
        let rowIndex = 0;
        rowIndex < json.length;
        rowIndex++
    ) {
        const row = json[rowIndex] || [];

        const possibleFarmColumn = findColumnInRow(
            row,
            ["FAZENDA"]
        );

        if (possibleFarmColumn === -1) continue;

        headerRow = rowIndex;
        farmColumn = possibleFarmColumn;

        statusColumn = findColumnInRow(
            row,
            [
                "STATUS",
                "SITUACAO",
                "ENVIO",
                "RECEBIDO",
                "FECHAMENTO"
            ]
        );

        /*
         * Caso não haja um nome claro para a coluna de status,
         * usa a primeira coluna após Fazenda.
         */
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
        rowIndex < json.length;
        rowIndex++
    ) {
        const row = json[rowIndex] || [];
        const farm = String(row[farmColumn] || "").trim();

        if (!farm) continue;

        const normalizedFarm = normalizeName(farm);

        if (
            normalizedFarm === "TOTAL" ||
            normalizedFarm.includes("PENDENCIA") ||
            normalizedFarm.includes("OBSERVACAO")
        ) {
            continue;
        }

        rows.push({
            farm,
            status: cellToStatus(row[statusColumn])
        });
    }

    return rows.length ? rows : null;
}


/* ============================================================
   DIVERGÊNCIAS
============================================================ */

function parseDivergSheet(json) {
    if (!Array.isArray(json) || !json.length) {
        return null;
    }

    let headerRow = -1;
    let farmColumn = -1;

    for (
        let rowIndex = 0;
        rowIndex < json.length;
        rowIndex++
    ) {
        const row = json[rowIndex] || [];

        const possibleFarmColumn = findColumnInRow(
            row,
            ["FAZENDA"]
        );

        if (possibleFarmColumn !== -1) {
            headerRow = rowIndex;
            farmColumn = possibleFarmColumn;
            break;
        }
    }

    if (headerRow === -1 || farmColumn === -1) {
        return null;
    }

    const header = json[headerRow] || [];

    let nascDiarioColumn = findColumnInRow(
        header,
        ["NASC DIARIO", "NASCIMENTO DIARIO"]
    );

    let nascSistemaColumn = findColumnInRow(
        header,
        ["NASC SISTEMA", "NASCIMENTO SISTEMA"]
    );

    let mortesDiarioColumn = findColumnInRow(
        header,
        ["MORTES DIARIO", "MORTE DIARIO"]
    );

    let mortesSistemaColumn = findColumnInRow(
        header,
        ["MORTES SISTEMA", "MORTE SISTEMA"]
    );

    /*
     * Compatibilidade com o formato antigo por posição.
     */
    if (nascDiarioColumn === -1) {
        nascDiarioColumn = farmColumn + 1;
    }

    if (nascSistemaColumn === -1) {
        nascSistemaColumn = farmColumn + 2;
    }

    if (mortesDiarioColumn === -1) {
        mortesDiarioColumn = farmColumn + 4;
    }

    if (mortesSistemaColumn === -1) {
        mortesSistemaColumn = farmColumn + 5;
    }

    const rows = [];

    for (
        let rowIndex = headerRow + 1;
        rowIndex < json.length;
        rowIndex++
    ) {
        const row = json[rowIndex] || [];
        const farm = String(row[farmColumn] || "").trim();

        if (!farm) continue;

        const normalizedFarm = normalizeName(farm);

        if (
            normalizedFarm === "TOTAL" ||
            normalizedFarm.includes("OBSERVACAO")
        ) {
            continue;
        }

        rows.push({
            farm,

            nd: toNumberOrEmpty(
                row[nascDiarioColumn]
            ),

            ns: toNumberOrEmpty(
                row[nascSistemaColumn]
            ),

            md: toNumberOrEmpty(
                row[mortesDiarioColumn]
            ),

            ms: toNumberOrEmpty(
                row[mortesSistemaColumn]
            )
        });
    }

    return rows.length ? rows : null;
}

/* ============================================================
   FILTRO APLICADO DURANTE A IMPORTAÇÃO
============================================================ */

function isImportDateFilterActive() {
    return Boolean(
        state?.filterStart ||
        state?.filterEnd
    );
}

function getFilteredImportedDays(days) {
    if (!Array.isArray(days)) {
        return [];
    }

    const normalizedDays = days.map(
        normalizeImportedDay
    );

    /*
     * Sem filtro ativo, importa todas as datas.
     */
    if (!isImportDateFilterActive()) {
        return normalizedDays.filter(Boolean);
    }

    let startDate = isoToDate(
        state.filterStart
    );

    let endDate = isoToDate(
        state.filterEnd
    );

    if (endDate && !startDate) {
        startDate = new Date(endDate);

        startDate.setDate(
            startDate.getDate() - 6
        );
    }

    if (startDate && !endDate) {
        endDate = new Date(startDate);

        endDate.setDate(
            endDate.getDate() + 6
        );
    }

    if (
        startDate &&
        endDate &&
        startDate > endDate
    ) {
        [startDate, endDate] = [
            endDate,
            startDate
        ];
    }

    return normalizedDays.filter(dayLabel => {
        const date = parseDayLabel(dayLabel);

        if (!date) return false;

        if (
            startDate &&
            date < startDate
        ) {
            return false;
        }

        if (
            endDate &&
            date > endDate
        ) {
            return false;
        }

        return true;
    });
}

/* ============================================================
   PRESERVAÇÃO DOS DADOS AO TROCAR O EIXO DE DATAS
============================================================ */

function remapDailyDataToDays(newDays) {
    const previousDays = Array.isArray(state.days)
        ? [...state.days]
        : [];

    const normalizedNewDays =
        newDays.map(normalizeImportedDay);

    PANEL_DEFS
        .filter(panel => panel.type === "daily")
        .forEach(panel => {
            const previousPanelData =
                state.data?.[panel.key] || {};

            const remappedPanelData = {};

            state.farms.forEach((farm, farmIndex) => {
                remappedPanelData[farmIndex] = {};

                normalizedNewDays.forEach(
                    (newDay, newDayIndex) => {
                        const previousDayIndex =
                            previousDays.findIndex(
                                previousDay =>
                                    normalizeImportedDay(
                                        previousDay
                                    ) === newDay
                            );

                        const previousStatus =
                            previousDayIndex === -1
                                ? undefined
                                : previousPanelData?.[farmIndex]?.[
                                    previousDayIndex
                                ];

                        remappedPanelData[farmIndex][newDayIndex] =
                            normalizeStatus(previousStatus);
                    }
                );
            });

            state.data[panel.key] =
                remappedPanelData;
        });

    state.days = normalizedNewDays;
}


/* ============================================================
   APLICAÇÃO DOS PAINÉIS DIÁRIOS
============================================================ */

function applyDailySheet(key, parsed, unmatched) {
    if (!parsed) return false;

    /*
     * Limpa somente o painel que está sendo atualizado.
     */
    state.data[key] = {};

    state.farms.forEach((farm, farmIndex) => {
        state.data[key][farmIndex] = {};
    });

    parsed.rows.forEach(row => {
        const farmIndex =
            matchFarmIndex(row.farm);

        if (farmIndex === -1) {
            unmatched.add(row.farm);
            return;
        }

        row.statuses.forEach(
            (status, importedDayIndex) => {
                const importedDay =
                    normalizeImportedDay(
                        parsed.days[importedDayIndex]
                    );

                const stateDayIndex =
                    state.days.findIndex(
                        stateDay =>
                            normalizeImportedDay(stateDay) ===
                            importedDay
                    );

                if (stateDayIndex === -1) return;

                state.data[key][farmIndex][stateDayIndex] =
                    normalizeStatus(status);
            }
        );
    });

    return true;
}


/* ============================================================
   APLICAÇÃO DO MAPA MENSAL
============================================================ */

function applyMonthlySheet(parsed, unmatched) {
    if (!parsed) return false;

    const newMonthly = {};

    state.farms.forEach((farm, farmIndex) => {
        newMonthly[farmIndex] =
            normalizeStatus(state.monthly?.[farmIndex]);
    });

    parsed.forEach(row => {
        const farmIndex =
            matchFarmIndex(row.farm);

        if (farmIndex === -1) {
            unmatched.add(row.farm);
            return;
        }

        newMonthly[farmIndex] =
            normalizeStatus(row.status);
    });

    state.monthly = newMonthly;

    return true;
}


/* ============================================================
   APLICAÇÃO DAS DIVERGÊNCIAS
============================================================ */

function applyDivergSheet(parsed, unmatched) {
    if (!parsed) return false;

    parsed.forEach(row => {
        const farmIndex =
            matchFarmIndex(row.farm);

        if (farmIndex === -1) {
            unmatched.add(row.farm);
            return;
        }

        state.diverg[farmIndex] = {
            nd:
                row.nd === ""
                    ? ""
                    : String(row.nd),

            ns:
                row.ns === ""
                    ? ""
                    : String(row.ns),

            md:
                row.md === ""
                    ? ""
                    : String(row.md),

            ms:
                row.ms === ""
                    ? ""
                    : String(row.ms)
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

    const parsed = {
        campo:
            importedData.campo
                ? parseStatusSheet(importedData.campo)
                : null,

        abastecimento:
            importedData.abastecimento
                ? parseStatusSheet(
                    importedData.abastecimento
                )
                : null,

        diario:
            importedData.diario
                ? parseStatusSheet(importedData.diario)
                : null,

        mensal:
            importedData.mensal
                ? parseMonthlySheet(importedData.mensal)
                : null,

        divergencias:
            importedData.divergencias
                ? parseDivergSheet(
                    importedData.divergencias
                )
                : null
    };

    const importedKeys = EXCEL_PANEL_KEYS.filter(
        key => importedData[key]
    );

    if (!importedKeys.length) {
        if (!silent) {
            showInlineWarning(
                "Nenhuma planilha foi carregada para processamento."
            );
        }

        return {
            found: [],
            failed: [],
            unmatched: []
        };
    }

    const found = [];
    const failed = [];
    const unmatched = new Set();

    /*
     * O eixo de datas é atualizado somente quando uma planilha
     * diária estiver sendo processada.
     */
    const canonical =
    parsed.diario ||
    parsed.campo ||
    parsed.abastecimento;

if (
    canonical &&
    Array.isArray(canonical.days) &&
    canonical.days.length
) {
    const importedDays =
        getFilteredImportedDays(
            canonical.days
        );

    if (importedDays.length) {
        remapDailyDataToDays(
            importedDays
        );
    } else if (
        isImportDateFilterActive()
    ) {
        showInlineWarning(
            "A planilha não possui dados dentro do período selecionado."
        );

        return {
            found: [],
            failed: [],
            unmatched: []
        };
    }
}
    ["campo", "abastecimento", "diario"]
        .forEach(key => {
            if (!importedData[key]) return;

            if (
                applyDailySheet(
                    key,
                    parsed[key],
                    unmatched
                )
            ) {
                found.push(key);
            } else {
                failed.push(key);
            }
        });

    if (importedData.mensal) {
        if (
            applyMonthlySheet(
                parsed.mensal,
                unmatched
            )
        ) {
            found.push("mensal");
        } else {
            failed.push("mensal");
        }
    }

    if (importedData.divergencias) {
        if (
            applyDivergSheet(
                parsed.divergencias,
                unmatched
            )
        ) {
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

    /*
     * Limpa os dados temporários processados para evitar que
     * uma sincronização futura reaplique planilhas antigas.
     */
    importedKeys.forEach(key => {
        importedData[key] = null;
    });

    if (!silent) {
        let message = found.length
            ? "Importado: " +
              found
                  .map(key => EXCEL_PANEL_LABELS[key])
                  .join(", ") +
              "."
            : "Nenhuma planilha reconhecida foi importada.";

        if (failed.length) {
            message +=
                " Não foi possível reconhecer a estrutura de: " +
                failed
                    .map(key => EXCEL_PANEL_LABELS[key])
                    .join(", ") +
                ".";
        }

        if (unmatched.size) {
            message +=
                " Fazendas sem correspondência: " +
                [...unmatched].join(", ") +
                ".";
        }

        showInlineWarning(message);
    }

    return {
        found,
        failed,
        unmatched: [...unmatched]
    };
}


/* ============================================================
   IMPORTAÇÃO DE ARQUIVO LOCAL
============================================================ */

async function importLocalExcelFiles(event) {
    const input = event.target;
    const files = [...(input.files || [])];

    if (!files.length) return;

    const importedFileNames = [];
    const errors = [];

    for (const file of files) {
        try {
            const buffer =
                await file.arrayBuffer();

            const workbook = XLSX.read(buffer, {
                type: "array",
                cellDates: true
            });

            const recognized =
                processWorkbook(
                    workbook,
                    file.name
                );

            if (recognized?.length) {
                importedFileNames.push(file.name);
            } else {
                errors.push(
                    `${file.name}: nenhuma aba reconhecida`
                );
            }
        } catch (error) {
            console.error(
                `Erro ao importar ${file.name}:`,
                error
            );

            errors.push(
                `${file.name}: ${
                    error?.message ||
                    "erro desconhecido"
                }`
            );
        }
    }

    /*
     * Permite selecionar o mesmo arquivo novamente.
     */
    input.value = "";

    const result = preencherSistema({
        silent: true
    });

    let message = result.found.length
        ? "Planilhas importadas: " +
          result.found
              .map(key => EXCEL_PANEL_LABELS[key])
              .join(", ") +
          "."
        : "Nenhuma planilha foi importada.";

    if (result.failed.length) {
        message +=
            " Estrutura não reconhecida em: " +
            result.failed
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

    if (errors.length) {
        message += " Erros: " + errors.join(" | ");
    }

    showInlineWarning(message);
}


/* ============================================================
   REGISTRO DO INPUT LOCAL
============================================================ */

function bindExcelImportEvents() {
    const input =
        document.getElementById("excelFiles");

    if (!input) return;

    if (input.dataset.bound === "true") {
        return;
    }

    input.dataset.bound = "true";

    input.addEventListener(
        "change",
        importLocalExcelFiles
    );
}

document.addEventListener(
    "DOMContentLoaded",
    bindExcelImportEvents
);
