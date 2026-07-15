/* ============================================================
   STATE.JS
   Estado padrão, normalização dos dados, carregamento,
   salvamento e sincronização entre usuários/abas.
============================================================ */

/* ============================================================
   ESTADO PADRÃO
============================================================ */

const defaultState = {
    company: "Nelore Cometa",

    title: "STATUS DE ENVIO DOS CONTROLES - FAZENDAS",

    period: "01/07 A 04/07",

    monthLabel: "FECHAMENTO – JUNHO/2025",

    banner:
        "Contamos com todos para mantermos a disciplina e a qualidade das informações!",

    farms: [
        "Furna Linda",
        "Estância Cometa",
        "Vale do Jaurú",
        "Pantanal 1",
        "Pantanal 2",
        "São Lucas",
        "Estância Maristela",
        "Santa Rita",
        "Liberdade",
        "São Sebastião"
    ],

    days: [
        "01/07",
        "02/07",
        "03/07",
        "04/07"
    ],

    selected: [
        "campo",
        "abastecimento",
        "diario",
        "mensal"
    ],

    /*
     * Painéis diários:
     * data[painel][fazenda][dia] = "ok" | "no" | "blank"
     */
    data: {},

    /*
     * Painel mensal:
     * monthly[fazenda] = "ok" | "no" | "blank"
     */
    monthly: {},

    /*
     * Divergências:
     * diverg[fazenda] = { nd, ns, md, ms }
     */
    diverg: {},

    /*
     * Cada planilha possui uma URL e sincronização independente.
     */
    sharepoint: {
        campo: {
            name: "Operações em Campo",
            url: "",
            connected: false,
            lastSync: "",
            error: ""
        },

        abastecimento: {
            name: "Abastecimento",
            url: "",
            connected: false,
            lastSync: "",
            error: ""
        },

        diario: {
            name: "Diário de Campo",
            url: "",
            connected: false,
            lastSync: "",
            error: ""
        },

        mensal: {
            name: "Mapa Mensal do Rebanho",
            url: "",
            connected: false,
            lastSync: "",
            error: ""
        },

        divergencias: {
            name: "Divergências entre Diário e Sistema",
            url: "",
            connected: false,
            lastSync: "",
            error: ""
        }
    },

    filterStart: "",

    filterEnd: "",

manualEntryEnabled: false,

dailyDataReady: {
    campo: false,
    abastecimento: false,
    diario: false
}
};

let state = null;


/* ============================================================
   DADOS TEMPORÁRIOS IMPORTADOS
============================================================ */

const importedData = {
    campo: null,
    abastecimento: null,
    diario: null,
    mensal: null,
    divergencias: null
};


/* ============================================================
   VARIÁVEIS DE CONTROLE
============================================================ */

let lastKnownJSON = null;
let saveTimer = null;
let pollTimer = null;


/* ============================================================
   FUNÇÕES AUXILIARES DO ESTADO
============================================================ */

function cloneDefaultState() {
    return JSON.parse(JSON.stringify(defaultState));
}

function normalizeStatus(value) {
    if (value === true) return "ok";
    if (value === false) return "no";

    if (
        value === "ok" ||
        value === "no" ||
        value === "blank"
    ) {
        return value;
    }

    return "ok";
}

function mergeObject(defaultObject, savedObject) {
    if (
        !savedObject ||
        typeof savedObject !== "object" ||
        Array.isArray(savedObject)
    ) {
        return JSON.parse(JSON.stringify(defaultObject));
    }

    const result = {
        ...defaultObject,
        ...savedObject
    };

    Object.keys(defaultObject).forEach(key => {
        const defaultValue = defaultObject[key];
        const savedValue = savedObject[key];

        if (
            defaultValue &&
            typeof defaultValue === "object" &&
            !Array.isArray(defaultValue)
        ) {
            result[key] = mergeObject(
                defaultValue,
                savedValue
            );
        }
    });

    return result;
}

function normalizeLoadedState(savedState) {
    const normalized = mergeObject(
        cloneDefaultState(),
        savedState
    );

    normalized.farms = Array.isArray(normalized.farms)
        ? normalized.farms
        : cloneDefaultState().farms;

    normalized.days = Array.isArray(normalized.days)
        ? normalized.days
        : cloneDefaultState().days;

    normalized.selected = Array.isArray(normalized.selected)
        ? normalized.selected
        : cloneDefaultState().selected;

    normalized.data =
        normalized.data &&
        typeof normalized.data === "object"
            ? normalized.data
            : {};

    normalized.monthly =
        normalized.monthly &&
        typeof normalized.monthly === "object"
            ? normalized.monthly
            : {};

    normalized.diverg =
        normalized.diverg &&
        typeof normalized.diverg === "object"
            ? normalized.diverg
            : {};

    /*
     * Migração do formato antigo, que possuía apenas
     * state.sharepointUrl e state.lastSync.
     */
    if (
        savedState?.sharepointUrl &&
        !normalized.sharepoint.campo.url
    ) {
        normalized.sharepoint.campo.url =
            savedState.sharepointUrl;

        normalized.sharepoint.campo.connected = false;

        normalized.sharepoint.campo.lastSync =
            savedState.lastSync || "";
    }

    delete normalized.sharepointUrl;
    delete normalized.lastSync;

    return normalized;
}


/* ============================================================
   GARANTIA DA ESTRUTURA DOS DADOS
============================================================ */

function ensureData() {
    if (!state) {
        state = cloneDefaultState();
    }

    if (!Array.isArray(state.farms)) {
        state.farms = [];
    }

    if (!Array.isArray(state.days)) {
        state.days = [];
    }

    if (!state.data || typeof state.data !== "object") {
        state.data = {};
    }

    if (
        !state.monthly ||
        typeof state.monthly !== "object"
    ) {
        state.monthly = {};
    }

    if (
        !state.diverg ||
        typeof state.diverg !== "object"
    ) {
        state.diverg = {};
    }

    /*
     * Painéis diários.
     */
    PANEL_DEFS
        .filter(panel => panel.type === "daily")
        .forEach(panel => {
            if (!state.data[panel.key]) {
                state.data[panel.key] = {};
            }

            state.farms.forEach((farm, farmIndex) => {
                if (!state.data[panel.key][farmIndex]) {
                    state.data[panel.key][farmIndex] = {};
                }

                state.days.forEach((day, dayIndex) => {
                    const current =
                        state.data[panel.key][farmIndex][dayIndex];

                    state.data[panel.key][farmIndex][dayIndex] =
                        normalizeStatus(current);
                });
            });
        });

    /*
     * Mapa mensal e divergências.
     */
    state.farms.forEach((farm, farmIndex) => {
        state.monthly[farmIndex] =
            normalizeStatus(state.monthly[farmIndex]);

        const currentDivergence =
            state.diverg[farmIndex];

        if (
            !currentDivergence ||
            typeof currentDivergence !== "object"
        ) {
            state.diverg[farmIndex] = {
                nd: "",
                ns: "",
                md: "",
                ms: ""
            };
        } else {
            state.diverg[farmIndex] = {
                nd:
                    currentDivergence.nd === undefined
                        ? ""
                        : String(currentDivergence.nd),

                ns:
                    currentDivergence.ns === undefined
                        ? ""
                        : String(currentDivergence.ns),

                md:
                    currentDivergence.md === undefined
                        ? ""
                        : String(currentDivergence.md),

                ms:
                    currentDivergence.ms === undefined
                        ? ""
                        : String(currentDivergence.ms)
            };
        }
    });

    /*
     * Painéis selecionados.
     */
    const validPanelKeys =
        PANEL_DEFS.map(panel => panel.key);

    if (
        !Array.isArray(state.selected) ||
        !state.selected.length
    ) {
        state.selected = [
            "campo",
            "abastecimento",
            "diario",
            "mensal"
        ];
    }

    state.selected = state.selected
        .filter(
            (key, index, array) =>
                validPanelKeys.includes(key) &&
                array.indexOf(key) === index
        )
        .slice(0, 4);

    /*
     * Estrutura das planilhas online.
     */
    if (
        !state.sharepoint ||
        typeof state.sharepoint !== "object"
    ) {
        state.sharepoint =
            cloneDefaultState().sharepoint;
    }

    Object.keys(defaultState.sharepoint).forEach(key => {
        const defaultConnection =
            defaultState.sharepoint[key];

        const currentConnection =
            state.sharepoint[key];

        state.sharepoint[key] = {
            ...defaultConnection,
            ...(
                currentConnection &&
                typeof currentConnection === "object"
                    ? currentConnection
                    : {}
            )
        };
    });

    state.filterStart =
        typeof state.filterStart === "string"
            ? state.filterStart
            : "";

    state.filterEnd =
        typeof state.filterEnd === "string"
            ? state.filterEnd
            : "";
}


/* ============================================================
   INDICADOR DE SALVAMENTO
============================================================ */

function updateSaveStatus(message) {
    const element =
        document.getElementById("saveStatus");

    /*
     * O layout novo não possui obrigatoriamente esse elemento.
     * Por isso, sua ausência não gera erro no console.
     */
    if (element) {
        element.textContent = message;
    }
}


/* ============================================================
   ADAPTADOR DE ARMAZENAMENTO
============================================================ */

async function storageGet() {
    /*
     * Usa o storage compartilhado quando ele estiver disponível.
     */
    if (
        window.storage &&
        typeof window.storage.get === "function"
    ) {
        const response =
            await window.storage.get(STORAGE_KEY, true);

        return response?.value || null;
    }

    /*
     * Fallback para navegadores comuns e GitHub Pages.
     */
    return localStorage.getItem(STORAGE_KEY);
}

async function storageSet(value) {
    if (
        window.storage &&
        typeof window.storage.set === "function"
    ) {
        await window.storage.set(
            STORAGE_KEY,
            value,
            true
        );

        return;
    }

    localStorage.setItem(STORAGE_KEY, value);
}


/* ============================================================
   CARREGAMENTO
============================================================ */

async function loadState() {
    try {
        const savedJSON = await storageGet();

        if (savedJSON) {
            const savedState = JSON.parse(savedJSON);

            state = normalizeLoadedState(savedState);
        } else {
            state = cloneDefaultState();
        }
    } catch (error) {
        console.error(
            "Erro ao carregar o estado:",
            error
        );

        state = cloneDefaultState();
    }

    ensureData();

    lastKnownJSON = JSON.stringify(state);

    if (typeof renderAll === "function") {
        renderAll();
    }

    startPolling();

    /*
     * A inicialização do SharePoint agora ocorre no app.js.
     * Não iniciamos aqui para evitar dois timers e duas
     * sincronizações simultâneas.
     */

    return state;
}


/* ============================================================
   SALVAMENTO
============================================================ */

function saveState() {
    updateSaveStatus("Salvando...");

    clearTimeout(saveTimer);

    saveTimer = setTimeout(async () => {
        try {
            ensureData();

            const json = JSON.stringify(state);

            await storageSet(json);

            lastKnownJSON = json;

            const now =
                new Date().toLocaleTimeString(
                    "pt-BR",
                    {
                        hour: "2-digit",
                        minute: "2-digit"
                    }
                );

            updateSaveStatus(
                `Salvo às ${now}`
            );
        } catch (error) {
            console.error(
                "Erro ao salvar o estado:",
                error
            );

            updateSaveStatus("Erro ao salvar");

            if (
                typeof showInlineWarning === "function"
            ) {
                showInlineWarning(
                    "Não foi possível salvar as alterações."
                );
            }
        }
    }, 400);
}


/* ============================================================
   SINCRONIZAÇÃO ENTRE ABAS E USUÁRIOS
============================================================ */

function isTypingNow() {
    const element = document.activeElement;

    if (!element) return false;

    const tagName =
        element.tagName?.toUpperCase();

    return (
        tagName === "INPUT" ||
        tagName === "TEXTAREA" ||
        tagName === "SELECT" ||
        element.isContentEditable
    );
}

function startPolling() {
    clearInterval(pollTimer);

    pollTimer = setInterval(async () => {
        if (isTypingNow()) return;

        try {
            const remoteJSON = await storageGet();

            if (
                !remoteJSON ||
                remoteJSON === lastKnownJSON
            ) {
                return;
            }

            const remoteState =
                JSON.parse(remoteJSON);

            state =
                normalizeLoadedState(remoteState);

            ensureData();

            lastKnownJSON =
                JSON.stringify(state);

            if (typeof renderAll === "function") {
                renderAll();
            }

            if (
                typeof renderAttachStatus === "function"
            ) {
                renderAttachStatus();
            }

            if (
                typeof showInlineWarning === "function"
            ) {
                showInlineWarning(
                    "Os dados foram atualizados por outra pessoa."
                );
            }
        } catch (error) {
            /*
             * O polling continuará tentando no próximo ciclo.
             */
            console.warn(
                "Não foi possível verificar atualizações:",
                error
            );
        }
    }, 6000);
}

function stopPolling() {
    clearInterval(pollTimer);
    pollTimer = null;
}


/* ============================================================
   RESET DO ESTADO
============================================================ */

function resetState() {
    state = cloneDefaultState();

    Object.keys(importedData).forEach(key => {
        importedData[key] = null;
    });

    ensureData();

    lastKnownJSON = JSON.stringify(state);

    if (typeof renderAll === "function") {
        renderAll();
    }

    if (
        typeof renderAttachStatus === "function"
    ) {
        renderAttachStatus();
    }

    saveState();
}
