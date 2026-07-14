// ==========================================
// UTILITÁRIOS
// ==========================================

function panelIconHTML(key) {

    const icon = PANEL_ICONS[key];

    return icon
        ? `<span class="rep-card-icon">${icon}</span>`
        : "";

}

function normalizeName(value) {

    return String(value ?? "")
        .trim()
        .toUpperCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "");

}

function classifySheetName(name) {

    const value = normalizeName(name);

    if (value.includes("DIVERG"))
        return "divergencias";

    if (value.includes("DIARIO"))
        return "diario";

    if (value.includes("OPERA"))
        return "campo";

    if (value.includes("ABASTE"))
        return "abastecimento";

    return null;

}
