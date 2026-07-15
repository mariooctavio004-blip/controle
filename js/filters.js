/* ============================================================
   FILTERS.JS
   Controle dos filtros de período do relatório.
============================================================ */

let filtersInitialized = false;


/* ============================================================
   CONVERSÃO DE DATAS
============================================================ */

function parseDayLabel(label) {
    const value = String(label ?? "").trim();

    const match = value.match(
        /^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/
    );

    if (!match) return null;

    const day = Number(match[1]);
    const month = Number(match[2]) - 1;

    let year = match[3]
        ? Number(match[3])
        : new Date().getFullYear();

    if (year < 100) {
        year += 2000;
    }

    const date = new Date(year, month, day);

    /*
     * Evita datas inválidas como 31/02.
     */
    if (
        Number.isNaN(date.getTime()) ||
        date.getDate() !== day ||
        date.getMonth() !== month ||
        date.getFullYear() !== year
    ) {
        return null;
    }

    date.setHours(0, 0, 0, 0);

    return date;
}

function isoToDate(iso) {
    if (!iso) return null;

    const match = String(iso).match(
        /^(\d{4})-(\d{2})-(\d{2})$/
    );

    if (!match) return null;

    const year = Number(match[1]);
    const month = Number(match[2]) - 1;
    const day = Number(match[3]);

    const date = new Date(year, month, day);

    if (
        Number.isNaN(date.getTime()) ||
        date.getDate() !== day ||
        date.getMonth() !== month ||
        date.getFullYear() !== year
    ) {
        return null;
    }

    date.setHours(0, 0, 0, 0);

    return date;
}

function dateToISO(date) {
    if (!(date instanceof Date)) {
        return "";
    }

    if (Number.isNaN(date.getTime())) {
        return "";
    }

    const year = date.getFullYear();

    const month = String(
        date.getMonth() + 1
    ).padStart(2, "0");

    const day = String(
        date.getDate()
    ).padStart(2, "0");

    return `${year}-${month}-${day}`;
}

function formatDateToDayLabel(date) {
    if (!(date instanceof Date)) {
        return "";
    }

    const day = String(
        date.getDate()
    ).padStart(2, "0");

    const month = String(
        date.getMonth() + 1
    ).padStart(2, "0");

    return `${day}/${month}`;
}


/* ============================================================
   AUXILIARES DO FILTRO
============================================================ */

function getAllDayIndexes() {
    if (!Array.isArray(state?.days)) {
        return [];
    }

    return state.days.map((_, index) => index);
}

function normalizeDateRange(startDate, endDate) {
    let start = startDate;
    let end = endDate;

    if (start && end && start > end) {
        [start, end] = [end, start];
    }

    return {
        startDate: start,
        endDate: end
    };
}

function getLatestAvailableDayDate() {
    if (!Array.isArray(state?.days)) {
        return null;
    }

    const parsedDates = state.days
        .map(parseDayLabel)
        .filter(Boolean)
        .sort((a, b) => a - b);

    if (!parsedDates.length) {
        return null;
    }

    return new Date(
        parsedDates[parsedDates.length - 1]
    );
}


/* ============================================================
   ÍNDICES VISÍVEIS NO RELATÓRIO
============================================================ */

function getVisibleDayIndexes() {
    const allIndexes = getAllDayIndexes();

    if (!allIndexes.length) {
        return [];
    }

    let startDate = isoToDate(
        state?.filterStart
    );

    let endDate = isoToDate(
        state?.filterEnd
    );

    if (!startDate && !endDate) {
        return allIndexes;
    }

    const parsedDays = state.days.map(
        parseDayLabel
    );

    const allDaysAreValid =
        parsedDays.every(Boolean);

    /*
     * Caso algum rótulo não seja uma data válida,
     * preserva a exibição por posição.
     */
    if (!allDaysAreValid) {
        if (endDate && !startDate) {
            return allIndexes.slice(
                Math.max(0, allIndexes.length - 7)
            );
        }

        return allIndexes;
    }

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

    const normalizedRange =
        normalizeDateRange(
            startDate,
            endDate
        );

    startDate =
        normalizedRange.startDate;

    endDate =
        normalizedRange.endDate;

    const filteredIndexes =
        allIndexes.filter(index => {
            const date = parsedDays[index];

            return (
                date >= startDate &&
                date <= endDate
            );
        });

    /*
     * Quando não houver datas dentro do intervalo,
     * retorna vazio para deixar claro que o filtro
     * não encontrou resultados.
     */
    return filteredIndexes;
}


/* ============================================================
   SINCRONIZAÇÃO DOS CAMPOS
============================================================ */

function syncDateInputs() {
    const startInput =
        document.getElementById(
            "reportStartDate"
        );

    const endInput =
        document.getElementById(
            "reportEndDate"
        );

    if (
        startInput &&
        document.activeElement !== startInput
    ) {
        startInput.value =
            state?.filterStart || "";
    }

    if (
        endInput &&
        document.activeElement !== endInput
    ) {
        endInput.value =
            state?.filterEnd || "";
    }
}


/* ============================================================
   APLICAÇÃO DO FILTRO
============================================================ */

function applyDateFilter(
    startISO,
    endISO,
    options = {}
) {
    const {
        showMessage = true,
        updatePeriodLabel = true
    } = options;

    let startDate = isoToDate(startISO);
    let endDate = isoToDate(endISO);

    if (startISO && !startDate) {
        showInlineWarning(
            "A data inicial informada é inválida."
        );

        return false;
    }

    if (endISO && !endDate) {
        showInlineWarning(
            "A data final informada é inválida."
        );

        return false;
    }

    const normalizedRange =
        normalizeDateRange(
            startDate,
            endDate
        );

    startDate =
        normalizedRange.startDate;

    endDate =
        normalizedRange.endDate;

    state.filterStart =
        startDate
            ? dateToISO(startDate)
            : "";

    state.filterEnd =
        endDate
            ? dateToISO(endDate)
            : "";

    if (
        updatePeriodLabel &&
        startDate &&
        endDate
    ) {
        state.period =
            `${formatDateToDayLabel(startDate)} ` +
            `A ${formatDateToDayLabel(endDate)}`;
    }

    syncDateInputs();

    if (
        typeof renderReport === "function"
    ) {
        renderReport();
    }

    if (
        typeof renderConfig === "function"
    ) {
        renderConfig();
    }

    saveState();

    if (showMessage) {
        const visibleIndexes =
            getVisibleDayIndexes();

        if (!visibleIndexes.length) {
            showInlineWarning(
                "Nenhum dia foi encontrado nesse período."
            );
        } else {
            showInlineWarning(
                `Filtro aplicado: ` +
                `${visibleIndexes.length} dia(s) exibido(s).`
            );
        }
    }

    return true;
}

function clearDateFilter(
    options = {}
) {
    const {
        showMessage = true
    } = options;

    state.filterStart = "";
    state.filterEnd = "";

    syncDateInputs();

    if (
        typeof renderReport === "function"
    ) {
        renderReport();
    }

    saveState();

    if (showMessage) {
        showInlineWarning(
            "Filtro de período removido."
        );
    }
}


/* ============================================================
   AÇÕES DOS BOTÕES
============================================================ */

function filterToday() {
    const today = new Date();

    today.setHours(0, 0, 0, 0);

    const iso = dateToISO(today);

    applyDateFilter(
        iso,
        iso,
        {
            showMessage: true,
            updatePeriodLabel: true
        }
    );
}

function filterLastSevenDays() {
    const endDate = new Date();

    endDate.setHours(0, 0, 0, 0);

    const startDate = new Date(endDate);

    startDate.setDate(
        startDate.getDate() - 6
    );

    applyDateFilter(
        dateToISO(startDate),
        dateToISO(endDate),
        {
            showMessage: true,
            updatePeriodLabel: true
        }
    );
}

function handleStartDateChange(event) {
    const startISO =
        event.target.value;

    const endISO =
        document.getElementById(
            "reportEndDate"
        )?.value || state.filterEnd;

    applyDateFilter(
        startISO,
        endISO,
        {
            showMessage: false,
            updatePeriodLabel: Boolean(
                startISO && endISO
            )
        }
    );
}

function handleEndDateChange(event) {
    const endISO =
        event.target.value;

    const startISO =
        document.getElementById(
            "reportStartDate"
        )?.value || state.filterStart;

    applyDateFilter(
        startISO,
        endISO,
        {
            showMessage: false,
            updatePeriodLabel: Boolean(
                startISO && endISO
            )
        }
    );
}


/* ============================================================
   REGISTRO DOS EVENTOS
============================================================ */

function bindFilterEvent(
    element,
    eventName,
    callback
) {
    if (!element) return;

    const key =
        `filterBound${eventName}`;

    if (element.dataset[key] === "true") {
        return;
    }

    element.dataset[key] = "true";

    element.addEventListener(
        eventName,
        callback
    );
}

function initializeFilters() {
    if (filtersInitialized) {
        syncDateInputs();
        return;
    }

    filtersInitialized = true;

    const startInput =
        document.getElementById(
            "reportStartDate"
        );

    const endInput =
        document.getElementById(
            "reportEndDate"
        );

    const todayButton =
        document.getElementById(
            "todayBtn"
        );

    const lastSevenButton =
        document.getElementById(
            "last7Btn"
        );

    bindFilterEvent(
        startInput,
        "change",
        handleStartDateChange
    );

    bindFilterEvent(
        endInput,
        "change",
        handleEndDateChange
    );

    bindFilterEvent(
        todayButton,
        "click",
        filterToday
    );

    bindFilterEvent(
        lastSevenButton,
        "click",
        filterLastSevenDays
    );

    syncDateInputs();
}


/* ============================================================
   INICIALIZAÇÃO
============================================================ */

document.addEventListener(
    "DOMContentLoaded",
    initializeFilters
);
