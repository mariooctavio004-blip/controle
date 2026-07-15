/* ============================================================
   RENDER.JS
   Renderização geral do painel de configurações e do relatório.
============================================================ */


/* ============================================================
   RENDERIZAÇÃO GERAL
============================================================ */

function renderAll() {
    if (!state) return;

    ensureData();

    renderConfig();
    renderReport();

    if (typeof renderAttachStatus === "function") {
        renderAttachStatus();
    }

    if (typeof syncDateInputs === "function") {
        syncDateInputs();
    }
}


/* ============================================================
   CAMPOS DAS CONFIGURAÇÕES
============================================================ */

function setInputValue(id, value) {
    const input = document.getElementById(id);

    if (!input) return;

    /*
     * Não sobrescreve o campo enquanto a pessoa está digitando.
     */
    if (document.activeElement === input) {
        return;
    }

    input.value = value ?? "";
}

function renderGeneralConfigFields() {
    setInputValue(
        "companyName",
        state.company
    );

    setInputValue(
        "reportTitle",
        state.title
    );

    setInputValue(
        "periodLabel",
        state.period
    );

    setInputValue(
        "bannerText",
        state.banner
    );
}


/* ============================================================
   PAINÉIS SELECIONADOS
============================================================ */

function getPanelSlotName(index) {
    const slots = [
        "Topo esquerdo",
        "Topo direito",
        "Baixo esquerdo",
        "Baixo direito"
    ];

    return slots[index] || "";
}

function createPanelSelectionItem(panel) {
    const selectedIndex =
        state.selected.indexOf(panel.key);

    const isSelected =
        selectedIndex !== -1;

    const label =
        document.createElement("label");

    label.className =
        "panel-select-item" +
        (isSelected ? " checked" : "");

    label.dataset.key = panel.key;

    const orderBadge =
        document.createElement("span");

    orderBadge.className =
        "order-badge" +
        (isSelected ? "" : " empty");

    orderBadge.textContent =
        isSelected
            ? String(selectedIndex + 1)
            : "";

    const checkbox =
        document.createElement("input");

    checkbox.type = "checkbox";
    checkbox.dataset.key = panel.key;
    checkbox.checked = isSelected;
    checkbox.setAttribute(
        "aria-label",
        `Exibir ${panel.title}`
    );

    const name =
        document.createElement("span");

    name.className = "name";
    name.textContent = panel.title;

    const slot =
        document.createElement("span");

    slot.className = "slot-hint";
    slot.textContent =
        isSelected
            ? getPanelSlotName(selectedIndex)
            : "Não exibido";

    label.append(
        orderBadge,
        checkbox,
        name,
        slot
    );

    return label;
}

function showPanelSelectionWarning() {
    const warning =
        document.getElementById(
            "selectWarning"
        );

    if (!warning) return;

    warning.textContent =
        "Você pode selecionar no máximo quatro painéis.";

    warning.style.display = "block";

    clearTimeout(warning._timer);

    warning._timer = setTimeout(() => {
        warning.style.display = "none";
    }, 3000);
}

function handlePanelSelectionChange(event) {
    const checkbox = event.target;
    const key = checkbox.dataset.key;

    if (!key) return;

    if (checkbox.checked) {
        if (state.selected.length >= 4) {
            checkbox.checked = false;

            showPanelSelectionWarning();

            return;
        }

        if (!state.selected.includes(key)) {
            state.selected.push(key);
        }
    } else {
        state.selected =
            state.selected.filter(
                selectedKey =>
                    selectedKey !== key
            );
    }

    renderConfig();
    renderReport();
    saveState();
}

function renderPanelSelection() {
    const list =
        document.getElementById(
            "panelSelectList"
        );

    if (!list) return;

    list.innerHTML = "";

    const fragment =
        document.createDocumentFragment();

    PANEL_DEFS.forEach(panel => {
        fragment.appendChild(
            createPanelSelectionItem(panel)
        );
    });

    list.appendChild(fragment);

    list
        .querySelectorAll(
            'input[type="checkbox"]'
        )
        .forEach(checkbox => {
            checkbox.addEventListener(
                "change",
                handlePanelSelectionChange
            );
        });
}


/* ============================================================
   FAZENDAS
============================================================ */

function createEditableChip({
    value,
    index,
    type,
    placeholder
}) {
    const chip =
        document.createElement("div");

    chip.className = "chip";
    chip.dataset.index = String(index);
    chip.dataset.type = type;

    const input =
        document.createElement("input");

    input.type = "text";
    input.value = value ?? "";
    input.dataset.idx = String(index);
    input.dataset.type = type;
    input.placeholder = placeholder;
    input.autocomplete = "off";

    const removeButton =
        document.createElement("button");

    removeButton.type = "button";
    removeButton.dataset.idx =
        String(index);

    removeButton.dataset.type =
        type;

    removeButton.className =
        "chip-remove";

    removeButton.title =
        type === "farm"
            ? "Remover fazenda"
            : "Remover dia";

    removeButton.setAttribute(
        "aria-label",
        removeButton.title
    );

    removeButton.textContent = "×";

    chip.append(
        input,
        removeButton
    );

    return chip;
}

function handleFarmInput(event) {
    const index =
        Number(event.target.dataset.idx);

    if (
        !Number.isInteger(index) ||
        index < 0 ||
        index >= state.farms.length
    ) {
        return;
    }

    state.farms[index] =
        event.target.value;

    renderReport();
    saveState();
}

function removeFarm(index) {
    if (
        !Number.isInteger(index) ||
        index < 0 ||
        index >= state.farms.length
    ) {
        return;
    }

    if (state.farms.length <= 1) {
        showInlineWarning(
            "O relatório precisa ter pelo menos uma fazenda."
        );

        return;
    }

    const farmName =
        state.farms[index] ||
        "esta fazenda";

    const confirmed =
        window.confirm(
            `Deseja remover ${farmName}?`
        );

    if (!confirmed) return;

    state.farms.splice(index, 1);

    PANEL_DEFS
        .filter(
            panel =>
                panel.type === "daily"
        )
        .forEach(panel => {
            const panelData =
                state.data?.[panel.key];

            if (!panelData) return;

            delete panelData[index];

            reindexObj(
                panelData,
                index
            );
        });

    delete state.monthly[index];

    reindexObj(
        state.monthly,
        index
    );

    delete state.diverg[index];

    reindexObj(
        state.diverg,
        index
    );

    ensureData();
    renderAll();
    saveState();
}

function renderFarmList() {
    const farmList =
        document.getElementById(
            "farmList"
        );

    if (!farmList) return;

    farmList.innerHTML = "";

    const fragment =
        document.createDocumentFragment();

    state.farms.forEach(
        (farm, index) => {
            fragment.appendChild(
                createEditableChip({
                    value: farm,
                    index,
                    type: "farm",
                    placeholder:
                        "Nome da fazenda"
                })
            );
        }
    );

    farmList.appendChild(fragment);

    farmList
        .querySelectorAll("input")
        .forEach(input => {
            input.addEventListener(
                "input",
                handleFarmInput
            );
        });

    farmList
        .querySelectorAll("button")
        .forEach(button => {
            button.addEventListener(
                "click",
                () => {
                    removeFarm(
                        Number(
                            button.dataset.idx
                        )
                    );
                }
            );
        });
}


/* ============================================================
   DIAS
============================================================ */

function normalizeDayInput(value) {
    return String(value ?? "")
        .trim()
        .replace(/\s+/g, "");
}

function handleDayInput(event) {
    const index =
        Number(event.target.dataset.idx);

    if (
        !Number.isInteger(index) ||
        index < 0 ||
        index >= state.days.length
    ) {
        return;
    }

    state.days[index] =
        event.target.value;

    renderReport();
    saveState();
}

function removeDay(index) {
    if (
        !Number.isInteger(index) ||
        index < 0 ||
        index >= state.days.length
    ) {
        return;
    }

    if (state.days.length <= 1) {
        showInlineWarning(
            "O relatório precisa ter pelo menos um dia."
        );

        return;
    }

    const dayLabel =
        state.days[index] ||
        "este dia";

    const confirmed =
        window.confirm(
            `Deseja remover o dia ${dayLabel}?`
        );

    if (!confirmed) return;

    state.days.splice(index, 1);

    PANEL_DEFS
        .filter(
            panel =>
                panel.type === "daily"
        )
        .forEach(panel => {
            const panelData =
                state.data?.[panel.key];

            if (!panelData) return;

            Object.keys(
                panelData
            ).forEach(farmIndex => {
                const farmData =
                    panelData[farmIndex];

                if (!farmData) return;

                delete farmData[index];

                reindexObj(
                    farmData,
                    index
                );
            });
        });

    ensureData();
    renderAll();
    saveState();
}

function renderDayList() {
    state.days = Array.isArray(state.days)
        ? state.days.filter(day =>
            String(day ?? "").trim()
        )
        : [];

    const dayList =
        document.getElementById(
            "dayList"
        );

    if (!dayList) return;

    dayList.innerHTML = "";

    const fragment =
        document.createDocumentFragment();

    state.days.forEach(
        (day, index) => {
            fragment.appendChild(
                createEditableChip({
                    value: day,
                    index,
                    type: "day",
                    placeholder: "DD/MM"
                })
            );
        }
    );

    dayList.appendChild(fragment);

    dayList
        .querySelectorAll("input")
        .forEach(input => {
            input.addEventListener(
                "input",
                handleDayInput
            );

            input.addEventListener(
                "blur",
                event => {
                    const index =
                        Number(
                            event.target.dataset.idx
                        );

                    const normalized =
                        normalizeDayInput(
                            event.target.value
                        );

                    if (!normalized) {
                        event.target.value =
                            state.days[index] ||
                            "--/--";

                        return;
                    }

                    const duplicateIndex =
                        state.days.findIndex(
                            (day, dayIndex) =>
                                dayIndex !== index &&
                                normalizeDayInput(day) ===
                                normalized
                        );

                    if (
                        duplicateIndex !== -1
                    ) {
                        showInlineWarning(
                            "Este dia já existe no relatório."
                        );

                        event.target.value =
                            state.days[index];

                        return;
                    }

                    state.days[index] =
                        normalized;

                    event.target.value =
                        normalized;

                    renderReport();
                    saveState();
                }
            );
        });

    dayList
        .querySelectorAll("button")
        .forEach(button => {
            button.addEventListener(
                "click",
                () => {
                    removeDay(
                        Number(
                            button.dataset.idx
                        )
                    );
                }
            );
        });
}


/* ============================================================
   PAINEL DE CONFIGURAÇÕES
============================================================ */

function renderConfig() {
    if (!state) return;

    renderGeneralConfigFields();
    renderPanelSelection();
    renderFarmList();
    renderDayList();
}


/* ============================================================
   CABEÇALHO DO RELATÓRIO
============================================================ */

function renderReportHeader() {
    const titleDisplay =
        document.getElementById(
            "titleDisplay"
        );

    const periodDisplay =
        document.getElementById(
            "periodDisplay"
        );

    const bannerDisplay =
        document.getElementById(
            "bannerDisplay"
        );

    if (titleDisplay) {
        titleDisplay.textContent =
            state.title || "";
    }

    if (periodDisplay) {
        periodDisplay.textContent =
            state.period || "";
    }

    if (bannerDisplay) {
        bannerDisplay.textContent =
            state.banner || "";

        bannerDisplay.style.display =
            state.banner
                ? ""
                : "none";
    }
}


/* ============================================================
   ESTADOS VAZIOS DO RELATÓRIO
============================================================ */

function createEmptyReportState({
    icon,
    title,
    message
}) {
    const container =
        document.createElement("div");

    container.className =
        "report-empty-state";

    container.innerHTML = `
        <div class="report-empty-icon">
            ${icon}
        </div>

        <strong>
            ${escapeHtml(title)}
        </strong>

        <p>
            ${escapeHtml(message)}
        </p>
    `;

    return container;
}


/* ============================================================
   EVENTOS DO RELATÓRIO
============================================================ */

function cycleDailyStatus(element) {
    const panelKey =
        element.dataset.panel;

    const farmIndex =
        Number(element.dataset.farm);

    const dayIndex =
        Number(element.dataset.day);

    if (
        !panelKey ||
        !Number.isInteger(farmIndex) ||
        !Number.isInteger(dayIndex)
    ) {
        return;
    }

    if (
        !state.data?.[panelKey]?.[farmIndex]
    ) {
        return;
    }

    const currentStatus =
        normalizeStatus(
            state.data[panelKey][farmIndex][dayIndex]
        );

    state.data[panelKey][farmIndex][dayIndex] =
        STATUS_CYCLE[currentStatus];

    renderReport();
    saveState();
}

function cycleMonthlyStatus(element) {
    const farmIndex =
        Number(element.dataset.monthly);

    if (!Number.isInteger(farmIndex)) {
        return;
    }

    const currentStatus =
        normalizeStatus(
            state.monthly[farmIndex]
        );

    state.monthly[farmIndex] =
        STATUS_CYCLE[currentStatus];

    renderReport();
    saveState();
}

function sanitizeNumericInput(value) {
    let clean =
        String(value ?? "")
            .replace(/[^0-9-]/g, "");

    clean =
        clean.replace(/(?!^)-/g, "");

    return clean;
}

function handleDivergenceInput(event) {
    const input = event.target;

    const farmIndex =
        Number(input.dataset.fi);

    const field =
        input.dataset.field;

    const allowedFields = [
        "nd",
        "ns",
        "md",
        "ms"
    ];

    if (
        !Number.isInteger(farmIndex) ||
        !allowedFields.includes(field)
    ) {
        return;
    }

    const cleanValue =
        sanitizeNumericInput(
            input.value
        );

    input.value = cleanValue;

    if (!state.diverg[farmIndex]) {
        state.diverg[farmIndex] = {
            nd: "",
            ns: "",
            md: "",
            ms: ""
        };
    }

    state.diverg[farmIndex][field] =
        cleanValue;

    /*
     * Salva imediatamente, mas evita recriar toda
     * a tabela enquanto o usuário está digitando.
     */
    saveState();
}

function handleDivergenceBlur(event) {
    const input = event.target;

    const farmIndex =
        input.dataset.fi;

    const field =
        input.dataset.field;

    renderReport();

    requestAnimationFrame(() => {
        const restoredInput =
            document.querySelector(
                `.num-input[data-fi="${farmIndex}"]` +
                `[data-field="${field}"]`
            );

        if (restoredInput) {
            restoredInput.focus();

            const length =
                restoredInput.value.length;

            try {
                restoredInput.setSelectionRange(
                    length,
                    length
                );
            } catch (error) {
                /*
                 * Alguns navegadores não suportam
                 * setSelectionRange nesse contexto.
                 */
            }
        }
    });
}

function bindReportEvents(grid) {
    grid
        .querySelectorAll(
            ".rep-icon[data-panel]"
        )
        .forEach(element => {
            element.addEventListener(
                "click",
                () => {
                    cycleDailyStatus(element);
                }
            );
        });

    grid
        .querySelectorAll(
            ".rep-icon[data-monthly]"
        )
        .forEach(element => {
            element.addEventListener(
                "click",
                () => {
                    cycleMonthlyStatus(element);
                }
            );
        });

    grid
        .querySelectorAll(".num-input")
        .forEach(input => {
            input.addEventListener(
                "input",
                handleDivergenceInput
            );

            input.addEventListener(
                "change",
                () => {
                    renderReport();
                }
            );
        });
}



/* ============================================================
   REAPLICAÇÃO AUTOMÁTICA DOS DADOS IMPORTADOS
============================================================ */

let renderRefreshTimer = null;

function refreshReportFromImportedSources() {
    clearTimeout(renderRefreshTimer);

    renderRefreshTimer = setTimeout(async () => {
        try {
            if (typeof reapplyImportedDataForCurrentFilter === "function") {
                await reapplyImportedDataForCurrentFilter({
                    silent:true,
                    save:false
                });
            }

            renderReport();
        } catch(error){
            console.error(error);
        }
    },80);
}

window.refreshReportFromImportedSources =
    refreshReportFromImportedSources;

/* ============================================================
   RELATÓRIO
============================================================ */

function renderReport() {
    if (!state) return;

    renderReportHeader();

    const grid =
        document.getElementById(
            "reportGrid"
        );

    if (!grid) return;

    grid.innerHTML = "";

    if (
        !Array.isArray(state.selected) ||
        !state.selected.length
    ) {
        grid.appendChild(
            createEmptyReportState({
                icon: "📊",
                title:
                    "Nenhum painel selecionado",
                message:
                    "Selecione pelo menos um painel nas configurações."
            })
        );

        return;
    }

    const selectedPanels =
        state.selected
            .map(key =>
                PANEL_DEFS.find(
                    panel =>
                        panel.key === key
                )
            )
            .filter(Boolean);

    if (!selectedPanels.length) {
        grid.appendChild(
            createEmptyReportState({
                icon: "⚠️",
                title:
                    "Painéis indisponíveis",
                message:
                    "Não foi possível localizar os painéis selecionados."
            })
        );

        return;
    }

    if (typeof updateMonthlyPeriodFromFilter === "function") {
        updateMonthlyPeriodFromFilter();
    }

    const visibleDayIndexes =
        typeof getVisibleDayIndexes ===
        "function"
            ? getVisibleDayIndexes()
            : state.days.map(
                (_, index) => index
            );

    const fragment =
        document.createDocumentFragment();

    selectedPanels.forEach(panel => {
        const card =
            document.createElement("article");

        card.className =
            `rep-card rep-card-${panel.type}`;

        card.dataset.panel = panel.key;

        if (
            panel.type === "daily" &&
            visibleDayIndexes.length === 0
        ) {
            card.innerHTML = `
                <div class="rep-card-head">
                    ${panelIconHTML(panel.key)}

                    <span>
                        ${escapeHtml(panel.title)}
                    </span>
                </div>

                <div class="panel-empty-state">
                    <strong>
                        Nenhum dia encontrado
                    </strong>

                    <p>
                        Não existem dados dentro do período selecionado.
                    </p>
                </div>
            `;
        } else if (
            panel.type === "daily"
        ) {
            card.innerHTML =
                buildDailyCard(panel);
        } else if (
            panel.type === "monthly"
        ) {
            card.innerHTML =
                buildMonthlyCard(panel);
        } else if (
            panel.type ===
            "divergencias"
        ) {
            card.innerHTML =
                buildDivergCard(panel);
        }

        fragment.appendChild(card);
    });

    grid.appendChild(fragment);

    /*
     * Adiciona uma classe baseada na quantidade
     * de painéis para o CSS organizar o grid.
     */
    grid.dataset.count =
        String(selectedPanels.length);

    bindReportEvents(grid);
}
