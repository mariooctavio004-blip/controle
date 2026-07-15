/* ============================================================
   SHAREPOINT.JS
   Gerencia as URLs, conexões e sincronizações independentes
   das planilhas do SharePoint/OneDrive.
============================================================ */

const SHAREPOINT_SYNC_INTERVAL = 3 * 60 * 1000;

const SHAREPOINT_KEYS = [
    "campo",
    "abastecimento",
    "diario",
    "mensal",
    "divergencias"
];

const SHAREPOINT_LABELS = {
    campo: "Operações em Campo",
    abastecimento: "Abastecimento",
    diario: "Diário de Campo",
    mensal: "Mapa Mensal do Rebanho",
    divergencias: "Divergências Diário × Sistema"
};

let sharepointSyncTimer = null;
let sharepointSyncRunning = false;

/* ============================================================
   ESTADO DAS CONEXÕES
============================================================ */

function createDefaultSharepointConnection(key) {
    return {
        name: SHAREPOINT_LABELS[key] || key,
        url: "",
        connected: false,
        lastSync: "",
        error: ""
    };
}

function ensureSharepointState() {
    if (!state) return;

    if (!state.sharepoint || typeof state.sharepoint !== "object") {
        state.sharepoint = {};
    }

    SHAREPOINT_KEYS.forEach(key => {
        const current = state.sharepoint[key];

        if (!current || typeof current !== "object") {
            state.sharepoint[key] = createDefaultSharepointConnection(key);
            return;
        }

        state.sharepoint[key] = {
            ...createDefaultSharepointConnection(key),
            ...current,
            name: current.name || SHAREPOINT_LABELS[key]
        };
    });

    /*
     * Migração do formato antigo:
     * state.sharepointUrl e state.lastSync.
     */
    if (state.sharepointUrl && !state.sharepoint.campo.url) {
        state.sharepoint.campo.url = state.sharepointUrl;
        state.sharepoint.campo.connected = true;
        state.sharepoint.campo.lastSync = state.lastSync || "";
    }
}

/* ============================================================
   ELEMENTOS DA INTERFACE
============================================================ */

function getSharepointElements(key) {
    return {
        input: document.getElementById(`${key}Url`),
        status: document.getElementById(`${key}Status`),
        lastSync: document.getElementById(`${key}LastSync`),
        card: document.querySelector(
            `.workbook-card[data-workbook="${key}"]`
        ),
        connectButton: document.querySelector(
            `.connect-workbook-btn[data-id="${key}"]`
        ),
        removeButton: document.querySelector(
            `.remove-workbook-btn[data-id="${key}"]`
        )
    };
}

function setButtonLoading(button, loading, loadingText = "Conectando...") {
    if (!button) return;

    if (loading) {
        if (!button.dataset.originalText) {
            button.dataset.originalText = button.textContent.trim();
        }

        button.disabled = true;
        button.textContent = loadingText;
        return;
    }

    button.disabled = false;

    if (button.dataset.originalText) {
        button.textContent = button.dataset.originalText;
        delete button.dataset.originalText;
    }
}

/* ============================================================
   CONVERSÃO DE URL
============================================================ */

function toDirectDownloadUrl(url) {
    const value = String(url || "").trim();

    if (!value) return "";

    try {
        const parsedUrl = new URL(value);
        const host = parsedUrl.hostname.toLowerCase();

        const isMicrosoftFile =
            host.includes("sharepoint.com") ||
            host.includes("1drv.ms") ||
            host.includes("onedrive.live.com");

        if (isMicrosoftFile) {
            parsedUrl.searchParams.set("download", "1");
        }

        return parsedUrl.toString();
    } catch (error) {
        return value;
    }
}

function isValidSharepointUrl(url) {
    try {
        const parsedUrl = new URL(String(url || "").trim());

        return (
            parsedUrl.protocol === "https:" ||
            parsedUrl.protocol === "http:"
        );
    } catch (error) {
        return false;
    }
}

/* ============================================================
   STATUS VISUAL INDIVIDUAL
============================================================ */

function renderSharepointItemStatus(key) {
    ensureSharepointState();

    const connection = state?.sharepoint?.[key];

    if (!connection) return;

    const elements = getSharepointElements(key);

    if (
        elements.input &&
        document.activeElement !== elements.input
    ) {
        elements.input.value = connection.url || "";
    }

    if (elements.status) {
        elements.status.classList.remove(
            "online",
            "offline",
            "error",
            "syncing"
        );

        if (connection.error) {
            elements.status.classList.add("error");
            elements.status.textContent = "Erro na sincronização";
        } else if (connection.connected) {
            elements.status.classList.add("online");
            elements.status.textContent = "Conectado";
        } else if (connection.url) {
            elements.status.classList.add("syncing");
            elements.status.textContent = "Aguardando sincronização";
        } else {
            elements.status.classList.add("offline");
            elements.status.textContent = "Não conectado";
        }
    }

    if (elements.lastSync) {
        if (connection.error) {
            elements.lastSync.textContent = connection.error;
        } else if (connection.lastSync) {
            elements.lastSync.textContent =
                `Última sincronização: ${connection.lastSync}`;
        } else if (connection.url) {
            elements.lastSync.textContent =
                "Planilha cadastrada. Aguardando sincronização.";
        } else {
            elements.lastSync.textContent = "Nunca sincronizada";
        }
    }

    if (elements.card) {
        elements.card.classList.toggle(
            "is-connected",
            Boolean(connection.connected)
        );

        elements.card.classList.toggle(
            "is-error",
            Boolean(connection.error)
        );
    }

    if (elements.removeButton) {
        elements.removeButton.disabled = !connection.url;
    }
}

/*
 * Esta função continua existindo porque renderAll() pode chamá-la.
 */
function renderAttachStatus() {
    ensureSharepointState();

    SHAREPOINT_KEYS.forEach(renderSharepointItemStatus);

    updateGeneralConnectionStatus();
}

/* ============================================================
   STATUS GERAL NO CABEÇALHO
============================================================ */

function updateGeneralConnectionStatus() {
    ensureSharepointState();

    const indicator = document.getElementById("connectionIndicator");
    const workbookList = document.getElementById("currentWorkbook");

    const connections = SHAREPOINT_KEYS
        .map(key => ({
            key,
            ...state.sharepoint[key]
        }))
        .filter(item => item.url);

    const connected = connections.filter(item => item.connected);
    const errors = connections.filter(item => item.error);

    if (indicator) {
        if (!connections.length) {
            indicator.textContent = "Nenhuma planilha conectada";
        } else if (errors.length) {
            indicator.textContent =
                `⚠️ ${connected.length} conectada(s), ` +
                `${errors.length} com erro`;
        } else {
            indicator.textContent =
                `🟢 ${connected.length} de ` +
                `${SHAREPOINT_KEYS.length} planilhas conectadas`;
        }
    }

    if (!workbookList) return;

    if (!connections.length) {
        workbookList.textContent = "Selecione uma planilha";
        return;
    }

    workbookList.innerHTML = connections
        .map(connection => {
            let icon = "🟡";

            if (connection.connected) icon = "✔";
            if (connection.error) icon = "⚠️";

            return `
                <div class="connected-item">
                    <span>${icon}</span>

                    <div>
                        <strong>
                            ${escapeHtml(
                                SHAREPOINT_LABELS[connection.key]
                            )}
                        </strong>

                        <small>
                            ${
                                connection.lastSync
                                    ? `Atualizada às ${escapeHtml(
                                        connection.lastSync
                                    )}`
                                    : "Aguardando sincronização"
                            }
                        </small>
                    </div>
                </div>
            `;
        })
        .join("");
}

/* ============================================================
   DOWNLOAD DA PLANILHA
============================================================ */

async function fetchSharepointWorkbook(url) {
    const directUrl = toDirectDownloadUrl(url);

    const response = await fetch(directUrl, {
        method: "GET",
        mode: "cors",
        cache: "no-store"
    });

    if (!response.ok) {
        throw new Error(
            `Não foi possível baixar a planilha. HTTP ${response.status}.`
        );
    }

    const contentType =
        response.headers.get("content-type") || "";

    if (
        contentType.includes("text/html") &&
        !contentType.includes("spreadsheet")
    ) {
        throw new Error(
            "O link retornou uma página HTML, e não um arquivo Excel. " +
            "Confira se o link permite baixar a planilha."
        );
    }

    const buffer = await response.arrayBuffer();

    if (!buffer.byteLength) {
        throw new Error("A planilha baixada está vazia.");
    }

    return XLSX.read(buffer, {
        type: "array",
        cellDates: true
    });
}

/* ============================================================
   PROCESSAMENTO POR PAINEL
============================================================ */

function processSharepointWorkbook(workbook, key) {
    if (
        !workbook ||
        !Array.isArray(workbook.SheetNames) ||
        !workbook.SheetNames.length
    ) {
        throw new Error("O arquivo não possui nenhuma aba válida.");
    }

    /*
     * Cada URL pertence a um painel específico. Por isso, usamos
     * a primeira aba do arquivo como fonte daquele painel.
     */
    const firstSheetName = workbook.SheetNames[0];
    const firstSheet = workbook.Sheets[firstSheetName];

    if (!firstSheet) {
        throw new Error("Não foi possível ler a primeira aba da planilha.");
    }

    const json = XLSX.utils.sheet_to_json(firstSheet, {
        header: 1,
        raw: false,
        dateNF: "dd/mm"
    });

    if (!Array.isArray(json) || !json.length) {
        throw new Error("A planilha não possui dados reconhecíveis.");
    }

    /*
     * importedData já é utilizado pelo excel.js.
     * O mensal também é armazenado aqui e será tratado no excel.js.
     */
    importedData[key] = json;
}

/* ============================================================
   SINCRONIZAÇÃO INDIVIDUAL
============================================================ */

async function syncSharepointSheet(key, silent = false) {
    ensureSharepointState();

    if (!SHAREPOINT_KEYS.includes(key)) {
        throw new Error(`Painel inválido: ${key}`);
    }

    const connection = state.sharepoint[key];
    const elements = getSharepointElements(key);
    const url = String(connection.url || "").trim();

    if (!url) {
        if (!silent) {
            showInlineWarning(
                `Informe a URL de ${SHAREPOINT_LABELS[key]}.`
            );
        }

        return false;
    }

    if (!isValidSharepointUrl(url)) {
        connection.connected = false;
        connection.error = "A URL informada não é válida.";

        renderSharepointItemStatus(key);
        updateGeneralConnectionStatus();

        if (!silent) {
            showInlineWarning(connection.error);
        }

        return false;
    }

    connection.error = "";

    if (elements.status) {
        elements.status.className = "workbook-status syncing";
        elements.status.textContent = "Sincronizando...";
    }

    setButtonLoading(
        elements.connectButton,
        true,
        "Sincronizando..."
    );

    try {
        const workbook = await fetchSharepointWorkbook(url);

        processSharepointWorkbook(workbook, key);

        connection.connected = true;
        connection.error = "";
        connection.lastSync =
            new Date().toLocaleTimeString("pt-BR", {
                hour: "2-digit",
                minute: "2-digit"
            });

        /*
         * preencherSistema processa os dados que foram colocados
         * em importedData e atualiza o relatório.
         */
        if (typeof preencherSistema === "function") {
            preencherSistema();
        } else {
            if (typeof ensureData === "function") ensureData();
            if (typeof renderAll === "function") renderAll();
            if (typeof saveState === "function") saveState();
        }

        renderSharepointItemStatus(key);
        updateGeneralConnectionStatus();

        if (!silent) {
            showInlineWarning(
                `${SHAREPOINT_LABELS[key]} sincronizada com sucesso.`
            );
        }

        return true;
    } catch (error) {
        console.error(
            `Erro ao sincronizar ${SHAREPOINT_LABELS[key]}:`,
            error
        );

        connection.connected = false;
        connection.error =
            error?.message ||
            "Não foi possível sincronizar esta planilha.";

        if (typeof saveState === "function") {
            saveState();
        }

        renderSharepointItemStatus(key);
        updateGeneralConnectionStatus();

        if (!silent) {
            showInlineWarning(
                `${SHAREPOINT_LABELS[key]}: ${connection.error}`
            );
        }

        return false;
    } finally {
        setButtonLoading(elements.connectButton, false);
    }
}

/* ============================================================
   CONECTAR UMA PLANILHA
============================================================ */

async function connectSharepointWorkbook(key) {
    ensureSharepointState();

    const elements = getSharepointElements(key);
    const url = elements.input?.value.trim() || "";

    if (!url) {
        showInlineWarning(
            `Cole a URL de ${SHAREPOINT_LABELS[key]}.`
        );

        elements.input?.focus();
        return;
    }

    if (!isValidSharepointUrl(url)) {
        showInlineWarning(
            `A URL de ${SHAREPOINT_LABELS[key]} não é válida.`
        );

        elements.input?.focus();
        return;
    }

    state.sharepoint[key].url = url;
    state.sharepoint[key].connected = false;
    state.sharepoint[key].lastSync = "";
    state.sharepoint[key].error = "";

    if (typeof saveState === "function") {
        saveState();
    }

    renderSharepointItemStatus(key);
    updateGeneralConnectionStatus();

    await syncSharepointSheet(key, false);
}

/* ============================================================
   REMOVER UMA PLANILHA
============================================================ */

function removeSharepointWorkbook(key) {
    ensureSharepointState();

    const connection = state.sharepoint[key];

    if (!connection?.url) {
        showInlineWarning(
            `${SHAREPOINT_LABELS[key]} não possui uma URL cadastrada.`
        );

        return;
    }

    const confirmed = window.confirm(
        `Deseja remover a conexão de ${SHAREPOINT_LABELS[key]}?`
    );

    if (!confirmed) return;

    state.sharepoint[key] = createDefaultSharepointConnection(key);

    /*
     * Remove somente os dados importados daquele painel.
     * Os dados editados manualmente no relatório permanecem.
     */
    if (
        typeof importedData === "object" &&
        importedData !== null &&
        Object.prototype.hasOwnProperty.call(importedData, key)
    ) {
        importedData[key] = null;
    }

    const elements = getSharepointElements(key);

    if (elements.input) {
        elements.input.value = "";
    }

    if (typeof saveState === "function") {
        saveState();
    }

    renderSharepointItemStatus(key);
    updateGeneralConnectionStatus();

    showInlineWarning(
        `${SHAREPOINT_LABELS[key]} foi desconectada.`
    );
}

/* ============================================================
   SINCRONIZAR TODAS
============================================================ */

async function syncAllSharepointSheets(silent = false) {
    ensureSharepointState();

    if (sharepointSyncRunning) return;

    const configuredKeys = SHAREPOINT_KEYS.filter(
        key => state.sharepoint[key]?.url
    );

    if (!configuredKeys.length) {
        if (!silent) {
            showInlineWarning(
                "Cadastre pelo menos uma URL antes de sincronizar."
            );
        }

        return;
    }

    sharepointSyncRunning = true;

    const syncAllButton =
        document.getElementById("syncAllWorkbooksBtn");

    const processButton =
        document.getElementById("processBtn");

    setButtonLoading(
        syncAllButton,
        true,
        "Sincronizando..."
    );

    setButtonLoading(
        processButton,
        true,
        "Atualizando..."
    );

    try {
        const results = [];

        /*
         * Fazemos em sequência para evitar várias renderizações
         * e downloads simultâneos pesados.
         */
        for (const key of configuredKeys) {
            const success =
                await syncSharepointSheet(key, true);

            results.push({
                key,
                success
            });
        }

        const successCount =
            results.filter(item => item.success).length;

        const errorCount =
            results.length - successCount;

        if (!silent) {
            if (!errorCount) {
                showInlineWarning(
                    `${successCount} planilha(s) sincronizada(s) com sucesso.`
                );
            } else {
                showInlineWarning(
                    `${successCount} planilha(s) sincronizada(s) e ` +
                    `${errorCount} com erro.`
                );
            }
        }
    } finally {
        sharepointSyncRunning = false;

        setButtonLoading(syncAllButton, false);
        setButtonLoading(processButton, false);

        renderAttachStatus();
    }
}

/* ============================================================
   SINCRONIZAÇÃO AUTOMÁTICA
============================================================ */

function startSharepointAutoSync() {
    clearInterval(sharepointSyncTimer);

    sharepointSyncTimer = setInterval(() => {
        if (!state?.sharepoint) return;

        const hasConfiguredUrl =
            SHAREPOINT_KEYS.some(
                key => Boolean(state.sharepoint[key]?.url)
            );

        if (hasConfiguredUrl) {
            syncAllSharepointSheets(true);
        }
    }, SHAREPOINT_SYNC_INTERVAL);
}

function stopSharepointAutoSync() {
    clearInterval(sharepointSyncTimer);
    sharepointSyncTimer = null;
}

/* ============================================================
   EVENTOS
============================================================ */

function bindSharepointEvents() {
    document
        .querySelectorAll(".connect-workbook-btn")
        .forEach(button => {
            if (button.dataset.bound === "true") return;

            button.dataset.bound = "true";

            button.addEventListener("click", () => {
                connectSharepointWorkbook(button.dataset.id);
            });
        });

    document
        .querySelectorAll(".remove-workbook-btn")
        .forEach(button => {
            if (button.dataset.bound === "true") return;

            button.dataset.bound = "true";

            button.addEventListener("click", () => {
                removeSharepointWorkbook(button.dataset.id);
            });
        });

    document
        .querySelectorAll(".workbook-url-input")
        .forEach(input => {
            if (input.dataset.bound === "true") return;

            input.dataset.bound = "true";

            input.addEventListener("keydown", event => {
                if (event.key !== "Enter") return;

                event.preventDefault();

                const card =
                    input.closest(".workbook-card");

                const key =
                    card?.dataset.workbook;

                if (key) {
                    connectSharepointWorkbook(key);
                }
            });
        });

    const syncAllButton =
        document.getElementById("syncAllWorkbooksBtn");

    if (
        syncAllButton &&
        syncAllButton.dataset.bound !== "true"
    ) {
        syncAllButton.dataset.bound = "true";

        syncAllButton.addEventListener("click", () => {
            syncAllSharepointSheets(false);
        });
    }

    /*
     * O botão Atualizar do cabeçalho sincroniza todas as URLs.
     */
    const processButton =
        document.getElementById("processBtn");

    if (
        processButton &&
        processButton.dataset.sharepointBound !== "true"
    ) {
        processButton.dataset.sharepointBound = "true";

        processButton.addEventListener("click", () => {
            syncAllSharepointSheets(false);
        });
    }
}

/* ============================================================
   INICIALIZAÇÃO
============================================================ */

function initializeSharepoint() {
    if (!state) return;

    ensureSharepointState();
    bindSharepointEvents();
    renderAttachStatus();
    startSharepointAutoSync();
}
