// ==========================================
// CONFIGURAÇÕES GERAIS DO SISTEMA
// ==========================================

const STORAGE_KEY = "status-envio-fazendas-v2";

const PANEL_DEFS = [
    {
        key: "campo",
        title: "Operações em Campo",
        type: "daily"
    },
    {
        key: "abastecimento",
        title: "Abastecimento",
        type: "daily"
    },
    {
        key: "diario",
        title: "Diário de Campo (Manejo/Rebanho)",
        type: "daily"
    },
    {
        key: "mensal",
        title: "Mapa Mensal do Rebanho",
        type: "monthly"
    },
    {
        key: "divergencias",
        title: "Divergências entre Diário e Sistema",
        type: "divergencias"
    }
];

const PANEL_ICONS = {
    campo: "🚜",
    abastecimento: "⛽",
    diario: "📋",
    mensal: "🐄",
    divergencias: "⚖️"
};

const PANEL_LABELS = {
    campo: "Operações em Campo",
    abastecimento: "Abastecimento",
    diario: "Diário de Campo",
    divergencias: "Divergências entre Diário e Sistema"
};
const STATUS_CYCLE = { ok:'no', no:'blank', blank:'ok' };

function panelIconHTML(key) {

    const icon = PANEL_ICONS[key];

    return icon
        ? `<span class="rep-card-icon">${icon}</span>`
        : "";

}
