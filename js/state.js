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

    return "blank";
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
    ensureImportedWorkbookState();
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


    if (
        !state.dailyDataReady ||
        typeof state.dailyDataReady !== "object"
    ) {
        state.dailyDataReady = {
            campo: false,
            abastecimento: false,
            diario: false
        };
    }

    ["campo", "abastecimento", "diario"].forEach(key => {
        state.dailyDataReady[key] =
            Boolean(state.dailyDataReady[key]);
    });

    state.manualEntryEnabled =
        Boolean(state.manualEntryEnabled);

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
   CACHE DOS ARQUIVOS IMPORTADOS
============================================================ */

if (!state?.importedWorkbooks) {
    defaultState.importedWorkbooks = {};
}

/*
 * Guarda informações suficientes para reprocessar as planilhas
 * quando o filtro mudar ou o botão Atualizar for acionado.
 */
function ensureImportedWorkbookState() {
    if (!state.importedWorkbooks ||
        typeof state.importedWorkbooks !== "object") {
        state.importedWorkbooks = {};
    }
}

function registerImportedWorkbookCache(key, metadata) {
    ensureImportedWorkbookState();
    state.importedWorkbooks[key] = {
        ...(state.importedWorkbooks[key] || {}),
        ...metadata,
        updatedAt: new Date().toISOString()
    };
}

function clearImportedWorkbookCache() {
    state.importedWorkbooks = {};
}

window.registerImportedWorkbookCache =
    registerImportedWorkbookCache;

window.clearImportedWorkbookCache =
    clearImportedWorkbookCache;

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
   ARMAZENAMENTO LOCAL DE SEGURANÇA
============================================================ */

function getLocalBackup() {
    try {
        return localStorage.getItem(STORAGE_KEY);
    } catch (error) {
        console.warn(
            "Não foi possível ler o backup local:",
            error
        );

        return null;
    }
}

function setLocalBackup(value) {
    try {
        localStorage.setItem(
            STORAGE_KEY,
            value
        );
    } catch (error) {
        console.warn(
            "Não foi possível salvar o backup local:",
            error
        );
    }
}


/* ============================================================
   APLICAÇÃO DE ESTADO RECEBIDO
============================================================ */

let pendingRemoteState = null;
let pendingRemoteTimer = null;
let realtimeUnsubscribe = null;

function applyReceivedState(
    receivedState,
    {
        showMessage = false,
        force = false
    } = {}
) {
    if (
        !receivedState ||
        typeof receivedState !== "object" ||
        Array.isArray(receivedState)
    ) {
        return false;
    }

    if (!force && isTypingNow()) {
        pendingRemoteState = receivedState;

        clearTimeout(pendingRemoteTimer);

        pendingRemoteTimer = setTimeout(() => {
            if (!pendingRemoteState) return;

            if (isTypingNow()) {
                applyReceivedState(
                    pendingRemoteState,
                    {
                        showMessage,
                        force: false
                    }
                );

                return;
            }

            const queuedState =
                pendingRemoteState;

            pendingRemoteState = null;

            applyReceivedState(
                queuedState,
                {
                    showMessage,
                    force: true
                }
            );
        }, 900);

        return false;
    }

    const normalized =
        normalizeLoadedState(
            receivedState
        );

    const remoteJSON =
        JSON.stringify(normalized);

    if (
        remoteJSON === lastKnownJSON
    ) {
        return false;
    }

    state = normalized;
    ensureData();
    ensureImportedWorkbookState();

    lastKnownJSON =
        JSON.stringify(state);

    setLocalBackup(lastKnownJSON);

    if (
        typeof renderAll === "function"
    ) {
        renderAll();
    }

    if (
        typeof renderAttachStatus === "function"
    ) {
        renderAttachStatus();
    }

    if (
        showMessage &&
        typeof showInlineWarning === "function"
    ) {
        showInlineWarning(
            "Os dados foram atualizados por outra pessoa."
        );
    }

    return true;
}


/* ============================================================
   CARREGAMENTO
============================================================ */

async function loadState() {
    let localState = null;

    try {
        const localJSON =
            getLocalBackup();

        if (localJSON) {
            localState =
                normalizeLoadedState(
                    JSON.parse(localJSON)
                );
        }
    } catch (error) {
        console.warn(
            "O backup local estava inválido:",
            error
        );
    }

    state =
        localState ||
        cloneDefaultState();

    ensureData();

    /*
     * Primeiro renderiza o backup local para o sistema abrir
     * rapidamente. Em seguida busca a versão compartilhada.
     */
    lastKnownJSON =
        JSON.stringify(state);

    if (
        typeof renderAll === "function"
    ) {
        renderAll();
    }

    try {
        if (
            typeof initializeSupabaseClient === "function"
        ) {
            initializeSupabaseClient();
        }

        if (
            typeof fetchSharedAppState === "function"
        ) {
            const remoteState =
                await fetchSharedAppState();

            const remoteHasContent =
                remoteState &&
                typeof remoteState === "object" &&
                !Array.isArray(remoteState) &&
                Object.keys(remoteState).length > 0;

            if (remoteHasContent) {
                applyReceivedState(
                    remoteState,
                    {
                        force: true
                    }
                );
            } else if (
                localState &&
                typeof saveSharedAppState === "function"
            ) {
                /*
                 * Primeira migração: se o Supabase ainda estiver
                 * vazio, envia a versão que já existia no navegador.
                 */
                await saveSharedAppState(
                    state
                );
            }
        }

        if (
            typeof subscribeToSharedAppState === "function"
        ) {
            realtimeUnsubscribe =
                subscribeToSharedAppState(
                    remoteState => {
                        applyReceivedState(
                            remoteState,
                            {
                                showMessage: true
                            }
                        );
                    }
                );
        }
    } catch (error) {
        console.error(
            "Não foi possível conectar ao Supabase:",
            error
        );

        if (
            typeof showInlineWarning === "function"
        ) {
            showInlineWarning(
                "O sistema abriu com o backup local. A sincronização online será tentada novamente."
            );
        }
    }

    startPolling();

    return state;
}


/* ============================================================
   SALVAMENTO
============================================================ */

function saveState() {
    updateSaveStatus(
        "Salvando..."
    );

    clearTimeout(saveTimer);

    saveTimer = setTimeout(async () => {
        try {
            ensureData();

            const json =
                JSON.stringify(state);

            /*
             * O backup local é atualizado antes da rede.
             */
            setLocalBackup(json);
            lastKnownJSON = json;

            if (
                typeof saveSharedAppState === "function"
            ) {
                await saveSharedAppState(
                    state
                );
            }

            const now =
                new Date().toLocaleTimeString(
                    "pt-BR",
                    {
                        hour: "2-digit",
                        minute: "2-digit"
                    }
                );

            updateSaveStatus(
                `Salvo às ${now} — sincronizado`
            );
        } catch (error) {
            console.error(
                "Erro ao salvar o estado compartilhado:",
                error
            );

            updateSaveStatus(
                "Salvo localmente — sem sincronização"
            );

            if (
                typeof showInlineWarning === "function"
            ) {
                showInlineWarning(
                    "As alterações foram salvas neste navegador, mas não foi possível sincronizar com os outros usuários."
                );
            }
        }
    }, 500);
}


/* ============================================================
   SINCRONIZAÇÃO EM TEMPO REAL E POLLING DE SEGURANÇA
============================================================ */

function isTypingNow() {
    const element =
        document.activeElement;

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

    /*
     * O Realtime normalmente atualiza imediatamente.
     * Este polling de 20 segundos serve apenas como segurança
     * caso o WebSocket seja interrompido.
     */
    pollTimer = setInterval(async () => {
        if (isTypingNow()) return;

        try {
            if (
                typeof fetchSharedAppState !== "function"
            ) {
                return;
            }

            const remoteState =
                await fetchSharedAppState();

            if (!remoteState) return;

            applyReceivedState(
                remoteState,
                {
                    showMessage: false
                }
            );
        } catch (error) {
            console.warn(
                "Não foi possível verificar atualizações no Supabase:",
                error
            );
        }
    }, 20000);
}

function stopPolling() {
    clearInterval(pollTimer);
    pollTimer = null;

    clearTimeout(
        pendingRemoteTimer
    );

    pendingRemoteTimer = null;
    pendingRemoteState = null;

    if (
        typeof realtimeUnsubscribe === "function"
    ) {
        realtimeUnsubscribe();
    }

    realtimeUnsubscribe = null;
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

    lastKnownJSON =
        JSON.stringify(state);

    setLocalBackup(
        lastKnownJSON
    );

    if (
        typeof renderAll === "function"
    ) {
        renderAll();
    }

    if (
        typeof renderAttachStatus === "function"
    ) {
        renderAttachStatus();
    }

    saveState();
}

window.addEventListener(
    "beforeunload",
    stopPolling
);
