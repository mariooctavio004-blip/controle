/* ============================================================
   EVENTS.JS
   Eventos gerais da interface:
   - configurações do relatório;
   - fazendas e dias;
   - confirmação de limpeza;
   - modal das planilhas;
   - menu de configurações;
   - exportação da imagem;
   - impressão.
============================================================ */


/* ============================================================
   UTILITÁRIOS DE EVENTOS
============================================================ */

function getElement(id) {
    return document.getElementById(id);
}

function bindEvent(id, eventName, callback) {
    const element = getElement(id);

    if (!element) {
        return false;
    }

    /*
     * Evita registrar o mesmo evento duas vezes caso alguma
     * inicialização seja executada novamente.
     */
    const bindingKey = `bound${eventName}`;

    if (element.dataset[bindingKey] === "true") {
        return true;
    }

    element.dataset[bindingKey] = "true";
    element.addEventListener(eventName, callback);

    return true;
}


/* ============================================================
   AVISOS DA INTERFACE
============================================================ */

function showInlineWarning(message, duration = 5000) {
    let box = getElement("inlineWarningBox");

    if (!box) {
        box = document.createElement("div");
        box.id = "inlineWarningBox";
        box.className = "inline-warning-box";

        const app = document.querySelector(".app");

        if (app) {
            app.appendChild(box);
        } else {
            document.body.appendChild(box);
        }
    }

    box.textContent = String(message || "");
    box.style.display = "block";

    clearTimeout(box._timer);

    box._timer = setTimeout(() => {
        box.style.display = "none";
    }, duration);
}


/* ============================================================
   CAMPOS DAS CONFIGURAÇÕES
============================================================ */

function bindConfigurationFields() {
    bindEvent("companyName", "input", event => {
        state.company = event.target.value;

        if (typeof renderReport === "function") {
            renderReport();
        }

        saveState();
    });

    bindEvent("reportTitle", "input", event => {
        state.title = event.target.value;

        if (typeof renderReport === "function") {
            renderReport();
        }

        saveState();
    });

    bindEvent("periodLabel", "input", event => {
        state.period = event.target.value;

        if (typeof renderReport === "function") {
            renderReport();
        }

        saveState();
    });

    bindEvent("bannerText", "input", event => {
        state.banner = event.target.value;

        if (typeof renderReport === "function") {
            renderReport();
        }

        saveState();
    });
}


/* ============================================================
   FAZENDAS E DIAS
============================================================ */

function createUniqueFarmName() {
    const baseName = "Nova fazenda";

    if (!state.farms.includes(baseName)) {
        return baseName;
    }

    let number = 2;

    while (state.farms.includes(`${baseName} ${number}`)) {
        number++;
    }

    return `${baseName} ${number}`;
}

function createUniqueDayLabel() {
    const baseLabel = "--/--";

    if (!state.days.includes(baseLabel)) {
        return baseLabel;
    }

    let number = 2;

    while (state.days.includes(`${baseLabel} ${number}`)) {
        number++;
    }

    return `${baseLabel} ${number}`;
}

function bindFarmAndDayButtons() {
    bindEvent("addFarmBtn", "click", () => {
        if (!Array.isArray(state.farms)) {
            state.farms = [];
        }

        state.farms.push(createUniqueFarmName());

        ensureData();
        renderAll();
        saveState();

        /*
         * Move o foco para o último campo de fazenda.
         */
        requestAnimationFrame(() => {
            const inputs = document.querySelectorAll(
                "#farmList input"
            );

            const lastInput = inputs[inputs.length - 1];

            if (lastInput) {
                lastInput.focus();
                lastInput.select();
            }
        });
    });

    bindEvent("addDayBtn", "click", () => {
        if (!Array.isArray(state.days)) {
            state.days = [];
        }

        state.days.push(createUniqueDayLabel());

        ensureData();
        renderAll();
        saveState();

        requestAnimationFrame(() => {
            const inputs = document.querySelectorAll(
                "#dayList input"
            );

            const lastInput = inputs[inputs.length - 1];

            if (lastInput) {
                lastInput.focus();
                lastInput.select();
            }
        });
    });
}


/* ============================================================
   CONFIRMAÇÃO DE LIMPEZA
============================================================ */

function openResetConfirmation() {
    const overlay = getElement("confirmOverlay");

    if (!overlay) return;

    overlay.style.display = "flex";
    document.body.classList.add("modal-open");
}

function closeResetConfirmation() {
    const overlay = getElement("confirmOverlay");

    if (!overlay) return;

    overlay.style.display = "none";
    document.body.classList.remove("modal-open");
}

function confirmReset() {
    if (typeof resetState === "function") {
        resetState();
    } else {
        state = JSON.parse(JSON.stringify(defaultState));

        Object.keys(importedData || {}).forEach(key => {
            importedData[key] = null;
        });

        ensureData();
        renderAll();
        saveState();
    }

    closeResetConfirmation();

    if (typeof renderAttachStatus === "function") {
        renderAttachStatus();
    }

    showInlineWarning(
        "Todos os dados foram restaurados para o padrão."
    );
}

function bindResetEvents() {
    bindEvent("resetBtn", "click", event => {
        event.stopPropagation();

        closeConfigMenu();
        openResetConfirmation();
    });

    bindEvent("confirmCancel", "click", () => {
        closeResetConfirmation();
    });

    bindEvent("confirmOk", "click", () => {
        confirmReset();
    });

    const overlay = getElement("confirmOverlay");

    if (
        overlay &&
        overlay.dataset.boundOverlay !== "true"
    ) {
        overlay.dataset.boundOverlay = "true";

        overlay.addEventListener("click", event => {
            if (event.target === overlay) {
                closeResetConfirmation();
            }
        });
    }
}


/* ============================================================
   PAINEL DE EDIÇÃO
============================================================ */

function isConfigPanelHidden() {
    const panel = getElement("configPanel");

    if (!panel) return true;

    return (
        panel.hidden ||
        panel.style.display === "none" ||
        panel.classList.contains("is-hidden")
    );
}

function updateToggleConfigButton() {
    const button = getElement("toggleConfigBtn");

    if (!button) return;

    button.innerHTML = isConfigPanelHidden()
        ? "👁 Mostrar edição"
        : "🙈 Ocultar edição";
}

function toggleConfigPanel() {
    const panel = getElement("configPanel");

    if (!panel) {
        showInlineWarning(
            "O painel de configurações não foi encontrado."
        );

        return;
    }

    const shouldShow = isConfigPanelHidden();

    panel.hidden = !shouldShow;
    panel.style.display = shouldShow ? "" : "none";
    panel.classList.toggle("is-hidden", !shouldShow);

    updateToggleConfigButton();

    closeConfigMenu();
}

function bindConfigPanelEvents() {
    bindEvent("toggleConfigBtn", "click", event => {
        event.stopPropagation();
        toggleConfigPanel();
    });

    updateToggleConfigButton();
}


/* ============================================================
   MENU DE CONFIGURAÇÕES
============================================================ */

function openConfigMenu() {
    const menu = getElement("configMenu");

    if (!menu) return;

    menu.classList.add("show");
}

function closeConfigMenu() {
    const menu = getElement("configMenu");

    if (!menu) return;

    menu.classList.remove("show");
}

function toggleConfigMenu() {
    const menu = getElement("configMenu");

    if (!menu) return;

    menu.classList.toggle("show");
}

function bindConfigMenuEvents() {
    bindEvent("configMenuBtn", "click", event => {
        event.preventDefault();
        event.stopPropagation();

        toggleConfigMenu();
    });

    const menu = getElement("configMenu");

    if (
        menu &&
        menu.dataset.boundClick !== "true"
    ) {
        menu.dataset.boundClick = "true";

        menu.addEventListener("click", event => {
            event.stopPropagation();
        });
    }

    if (
        document.body.dataset.boundConfigMenu !== "true"
    ) {
        document.body.dataset.boundConfigMenu = "true";

        document.addEventListener("click", () => {
            closeConfigMenu();
        });
    }
}


/* ============================================================
   MODAL DAS PLANILHAS
============================================================ */

function openWorkbookConnectionModal() {
    const modal = getElement("workbookModal");

    if (!modal) {
        showInlineWarning(
            "O modal das planilhas não foi encontrado."
        );

        return;
    }

    if (typeof renderAttachStatus === "function") {
        renderAttachStatus();
    }

    modal.style.display = "flex";
    modal.setAttribute("aria-hidden", "false");

    document.body.classList.add("modal-open");

    requestAnimationFrame(() => {
        const firstInput = modal.querySelector(
            ".workbook-url-input"
        );

        firstInput?.focus();
    });
}

function closeWorkbookConnectionModal() {
    const modal = getElement("workbookModal");

    if (!modal) return;

    modal.style.display = "none";
    modal.setAttribute("aria-hidden", "true");

    document.body.classList.remove("modal-open");
}

function bindWorkbookModalEvents() {
    bindEvent("openWorkbookModal", "click", () => {
        openWorkbookConnectionModal();
    });

    bindEvent("closeWorkbookModal", "click", () => {
        closeWorkbookConnectionModal();
    });

    bindEvent(
        "closeWorkbookModalBottom",
        "click",
        () => {
            closeWorkbookConnectionModal();
        }
    );

    const modal = getElement("workbookModal");

    if (
        modal &&
        modal.dataset.boundOverlay !== "true"
    ) {
        modal.dataset.boundOverlay = "true";

        modal.addEventListener("click", event => {
            if (event.target === modal) {
                closeWorkbookConnectionModal();
            }
        });
    }
}


/* ============================================================
   EXPORTAÇÃO E IMPRESSÃO
============================================================ */

function isIOS() {
    return (
        /iP(hone|od|ad)/.test(navigator.platform) ||
        (
            navigator.userAgent.includes("Mac") &&
            "ontouchend" in document
        )
    );
}

function swapInputsForCapture() {
    const report = getElement("report");

    if (!report) return;

    report
        .querySelectorAll(".num-input")
        .forEach(input => {
            if (input.dataset.wasHidden === "1") {
                return;
            }

            const span = document.createElement("span");

            span.className = "num-print-value";
            span.textContent = input.value || "";

            const computedStyle =
                window.getComputedStyle(input);

            span.style.display = "inline-flex";
            span.style.alignItems = "center";
            span.style.justifyContent = "center";
            span.style.width = computedStyle.width;
            span.style.minHeight = computedStyle.height;

            input.dataset.wasHidden = "1";
            input.style.display = "none";

            input.insertAdjacentElement(
                "afterend",
                span
            );
        });
}

function restoreInputsAfterCapture() {
    const report = getElement("report");

    if (!report) return;

    report
        .querySelectorAll(".num-print-value")
        .forEach(span => span.remove());

    report
        .querySelectorAll(
            ".num-input[data-was-hidden], " +
            ".num-input[data-wasHidden]"
        )
        .forEach(input => {
            input.style.display = "";
            delete input.dataset.wasHidden;
        });

    /*
     * querySelector não trata camelCase do dataset como atributo.
     * Este bloco garante a restauração em todos os navegadores.
     */
    report
        .querySelectorAll(".num-input")
        .forEach(input => {
            if (input.dataset.wasHidden === "1") {
                input.style.display = "";
                delete input.dataset.wasHidden;
            }
        });
}

async function generateCanvas() {
    const report = getElement("report");

    if (!report) {
        throw new Error(
            "A área do relatório não foi encontrada."
        );
    }

    if (typeof html2canvas !== "function") {
        throw new Error(
            "A biblioteca html2canvas não foi carregada."
        );
    }

    /*
     * Espera as fontes e imagens terminarem de carregar.
     */
    if (document.fonts?.ready) {
        await document.fonts.ready;
    }

    const images = [...report.querySelectorAll("img")];

    await Promise.all(
        images.map(image => {
            if (image.complete) {
                return Promise.resolve();
            }

            return new Promise(resolve => {
                image.addEventListener(
                    "load",
                    resolve,
                    { once: true }
                );

                image.addEventListener(
                    "error",
                    resolve,
                    { once: true }
                );
            });
        })
    );

    return html2canvas(report, {
        scale: 2,
        backgroundColor: "#ffffff",
        useCORS: true,
        allowTaint: false,
        logging: false,
        scrollX: 0,
        scrollY: -window.scrollY,
        windowWidth: report.scrollWidth,
        windowHeight: report.scrollHeight
    });
}

function createExportFilename() {
    const period =
        String(state?.period || "periodo")
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .replace(/[^a-zA-Z0-9]+/g, "_")
            .replace(/^_+|_+$/g, "");

    return `status_envio_${period || "periodo"}.png`;
}

function downloadCanvas(canvas, filename) {
    return new Promise((resolve, reject) => {
        canvas.toBlob(blob => {
            if (!blob) {
                reject(
                    new Error(
                        "Não foi possível criar o arquivo de imagem."
                    )
                );

                return;
            }

            const url = URL.createObjectURL(blob);
            const link = document.createElement("a");

            link.href = url;
            link.download = filename;
            link.style.display = "none";

            document.body.appendChild(link);

            link.click();
            link.remove();

            setTimeout(() => {
                URL.revokeObjectURL(url);
            }, 4000);

            resolve();
        }, "image/png");
    });
}

function openCanvasOnIOS(canvas, filename) {
    const dataUrl = canvas.toDataURL("image/png");
    const newWindow = window.open("", "_blank");

    if (!newWindow) {
        throw new Error(
            "O navegador bloqueou a nova aba. " +
            "Permita pop-ups e tente novamente."
        );
    }

    newWindow.document.write(`
        <!DOCTYPE html>

        <html lang="pt-BR">

        <head>
            <meta charset="UTF-8">

            <meta
                name="viewport"
                content="width=device-width, initial-scale=1.0"
            >

            <title>${escapeHtml(filename)}</title>
        </head>

        <body
            style="
                margin:0;
                padding:20px;
                background:#111827;
                display:flex;
                align-items:flex-start;
                justify-content:center;
            "
        >
            <img
                src="${dataUrl}"
                alt="Status de envio"
                style="
                    display:block;
                    max-width:100%;
                    height:auto;
                    background:white;
                    box-shadow:0 20px 60px rgba(0,0,0,.35);
                "
            >
        </body>

        </html>
    `);

    newWindow.document.close();
}

async function exportReportImage() {
    const button = getElement("exportBtn");

    if (!button || button.disabled) return;

    const originalText = button.textContent;

    button.disabled = true;
    button.textContent = "⏳ Gerando...";

    swapInputsForCapture();

    try {
        const canvas = await generateCanvas();
        const filename = createExportFilename();

        if (isIOS()) {
            openCanvasOnIOS(canvas, filename);
        } else {
            await downloadCanvas(canvas, filename);
        }

        showInlineWarning(
            "Imagem gerada com sucesso."
        );
    } catch (error) {
        console.error(
            "Erro ao exportar o relatório:",
            error
        );

        showInlineWarning(
            error?.message ||
            "Não foi possível gerar a imagem."
        );
    } finally {
        restoreInputsAfterCapture();

        button.disabled = false;
        button.textContent = originalText;
    }
}

function bindExportEvents() {
    bindEvent("exportBtn", "click", () => {
        exportReportImage();
    });

    if (
        window.dataset?.boundPrintEvents !== "true"
    ) {
        /*
         * window não possui dataset. Usamos uma propriedade própria.
         */
    }

    if (!window.__reportPrintEventsBound) {
        window.__reportPrintEventsBound = true;

        window.addEventListener(
            "beforeprint",
            swapInputsForCapture
        );

        window.addEventListener(
            "afterprint",
            restoreInputsAfterCapture
        );
    }
}


/* ============================================================
   TECLADO
============================================================ */

function bindKeyboardEvents() {
    if (document.body.dataset.boundKeyboard === "true") {
        return;
    }

    document.body.dataset.boundKeyboard = "true";

    document.addEventListener("keydown", event => {
        if (event.key !== "Escape") return;

        closeConfigMenu();
        closeWorkbookConnectionModal();
        closeResetConfirmation();
    });
}


/* ============================================================
   INICIALIZAÇÃO DOS EVENTOS
============================================================ */

function initializeEvents() {
    bindConfigurationFields();
    bindFarmAndDayButtons();
    bindResetEvents();
    bindConfigPanelEvents();
    bindConfigMenuEvents();
    bindWorkbookModalEvents();
    bindExportEvents();
    bindKeyboardEvents();
}

document.addEventListener(
    "DOMContentLoaded",
    initializeEvents
);
