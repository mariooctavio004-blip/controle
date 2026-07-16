/* ============================================================
   FILTERS.JS
   O filtro altera somente o período selecionado. Os dados brutos
   permanecem preservados e são reaplicados pelo excel.js.
============================================================ */

let filtersInitialized = false;
let filterApplyToken = 0;

function parseDayLabel(label) {
    const value = String(label ?? "").trim();
    const match = value.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
    if (!match) return null;

    const day = Number(match[1]);
    const month = Number(match[2]) - 1;
    let year = match[3] ? Number(match[3]) : getFilterReferenceYear();
    if (year < 100) year += 2000;

    const date = new Date(year, month, day);
    if (
        Number.isNaN(date.getTime()) ||
        date.getDate() !== day ||
        date.getMonth() !== month ||
        date.getFullYear() !== year
    ) return null;

    date.setHours(0, 0, 0, 0);
    return date;
}

function isoToDate(iso) {
    if (!iso) return null;
    const match = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) return null;

    const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    if (
        Number.isNaN(date.getTime()) ||
        date.getFullYear() !== Number(match[1]) ||
        date.getMonth() !== Number(match[2]) - 1 ||
        date.getDate() !== Number(match[3])
    ) return null;

    date.setHours(0, 0, 0, 0);
    return date;
}

function dateToISO(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatDateToDayLabel(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
    return `${String(date.getDate()).padStart(2, "0")}/${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function getFilterReferenceYear() {
    return (
        isoToDate(state?.filterEnd)?.getFullYear() ||
        isoToDate(state?.filterStart)?.getFullYear() ||
        new Date().getFullYear()
    );
}

function normalizeDateRange(startDate, endDate) {
    let start = startDate;
    let end = endDate;
    if (start && end && start > end) [start, end] = [end, start];
    return { startDate: start, endDate: end };
}

function buildDayLabelsBetweenDates(startDate, endDate) {
    if (!startDate || !endDate) return [];
    const result = [];
    const current = new Date(startDate);
    const limit = new Date(endDate);
    current.setHours(0,0,0,0);
    limit.setHours(0,0,0,0);
    while (current <= limit) {
        result.push(formatDateToDayLabel(current));
        current.setDate(current.getDate() + 1);
    }
    return result;
}

function getCurrentFilterRange() {
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

    return normalizeDateRange(startDate, endDate);
}

function getVisibleDayIndexes() {
    if (!Array.isArray(state?.days) || !state.days.length) return [];
    const { startDate, endDate } = getCurrentFilterRange();
    if (!startDate && !endDate) return state.days.map((_, index) => index);

    const year = getFilterReferenceYear();
    return state.days.reduce((indexes, label, index) => {
        const match = String(label).trim().match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
        if (!match) return indexes;
        let itemYear = match[3] ? Number(match[3]) : year;
        if (itemYear < 100) itemYear += 2000;
        const date = new Date(itemYear, Number(match[2]) - 1, Number(match[1]));
        date.setHours(0,0,0,0);
        if ((!startDate || date >= startDate) && (!endDate || date <= endDate)) indexes.push(index);
        return indexes;
    }, []);
}

function syncDateInputs() {
    const startInput = document.getElementById("reportStartDate");
    const endInput = document.getElementById("reportEndDate");
    if (startInput && document.activeElement !== startInput) startInput.value = state?.filterStart || "";
    if (endInput && document.activeElement !== endInput) endInput.value = state?.filterEnd || "";
}

async function refreshDataAfterFilterChange({ showMessage = false } = {}) {
    const token = ++filterApplyToken;
    try {
        if (typeof reapplyImportedDataForCurrentFilter === "function") {
            const result = await reapplyImportedDataForCurrentFilter({ silent: true, save: false });
            if (token !== filterApplyToken) return false;
            if (showMessage && typeof showInlineWarning === "function") {
                if (result?.noDatesInFilter) showInlineWarning("Nenhum dado foi encontrado nesse período.");
                else showInlineWarning("Período atualizado com os arquivos importados.");
            }
        } else if (typeof renderAll === "function") {
            renderAll();
        }
        if (typeof saveState === "function") saveState();
        return true;
    } catch (error) {
        console.error("Erro ao atualizar o período:", error);
        if (typeof showInlineWarning === "function") showInlineWarning("Não foi possível atualizar os dados para esse período.");
        return false;
    }
}

async function applyDateFilter(startISO, endISO, options = {}) {
    const { showMessage = true, updatePeriodLabel = true } = options;
    let startDate = isoToDate(startISO);
    let endDate = isoToDate(endISO);

    if (startISO && !startDate) { showInlineWarning("A data inicial informada é inválida."); return false; }
    if (endISO && !endDate) { showInlineWarning("A data final informada é inválida."); return false; }

    ({ startDate, endDate } = normalizeDateRange(startDate, endDate));
    state.filterStart = startDate ? dateToISO(startDate) : "";
    state.filterEnd = endDate ? dateToISO(endDate) : "";

    if (updatePeriodLabel && startDate && endDate) {
        state.period = `${formatDateToDayLabel(startDate)} A ${formatDateToDayLabel(endDate)}`;
    }

    state.manualEntryEnabled = false;
    syncDateInputs();
    await refreshDataAfterFilterChange({ showMessage });
    return true;
}

async function clearDateFilter(options = {}) {
    state.filterStart = "";
    state.filterEnd = "";
    syncDateInputs();
    await refreshDataAfterFilterChange({ showMessage: false });
    if (options.showMessage !== false) showInlineWarning("Filtro de período removido.");
}

function filterToday() {
    const today = new Date();
    today.setHours(0,0,0,0);
    const iso = dateToISO(today);
    return applyDateFilter(iso, iso, { showMessage: true, updatePeriodLabel: true });
}

function filterLastSevenDays() {
    const endDate = new Date();
    endDate.setHours(0,0,0,0);
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 6);
    return applyDateFilter(dateToISO(startDate), dateToISO(endDate), { showMessage: true, updatePeriodLabel: true });
}

function handleStartDateChange(event) {
    const endISO = document.getElementById("reportEndDate")?.value || state.filterEnd;
    return applyDateFilter(event.target.value, endISO, { showMessage: false, updatePeriodLabel: Boolean(event.target.value && endISO) });
}

function handleEndDateChange(event) {
    const startISO = document.getElementById("reportStartDate")?.value || state.filterStart;
    return applyDateFilter(startISO, event.target.value, { showMessage: false, updatePeriodLabel: Boolean(startISO && event.target.value) });
}

function bindFilterEvent(element, eventName, callback) {
    if (!element) return;
    const key = `filterBound${eventName}`;
    if (element.dataset[key] === "true") return;
    element.dataset[key] = "true";
    element.addEventListener(eventName, callback);
}

function initializeFilters() {
    if (filtersInitialized) { syncDateInputs(); return; }
    filtersInitialized = true;
    bindFilterEvent(document.getElementById("reportStartDate"), "change", handleStartDateChange);
    bindFilterEvent(document.getElementById("reportEndDate"), "change", handleEndDateChange);
    bindFilterEvent(document.getElementById("todayBtn"), "click", filterToday);
    bindFilterEvent(document.getElementById("last7Btn"), "click", filterLastSevenDays);
    syncDateInputs();
}

document.addEventListener("DOMContentLoaded", initializeFilters);
