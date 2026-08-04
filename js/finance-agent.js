/* ============================================================
   FINANCE-AGENT.JS
   Assistente financeiro em formato de chat.
============================================================ */

const FINANCE_CHAT_STORAGE_KEY = "financeChatStateV2";

const entryLabels = {
    income: "Renda",
    spent: "Gasto realizado",
    planned: "Gasto previsto",
    investment: "Investimento"
};

const defaultChatState = {
    messages: [
        {
            role: "assistant",
            text: "Olá! Me conte seus gastos em linguagem simples. Ex.: “gastei 80 no mercado”, “ainda vou gastar 1200 de aluguel”, “recebi 5000” ou “investi 300”."
        }
    ],
    entries: []
};

let chatState = null;

function formatCurrency(value) {
    return Number(value || 0).toLocaleString("pt-BR", {
        style: "currency",
        currency: "BRL"
    });
}

function escapeHTML(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function parseMoney(value) {
    const normalized = String(value || "")
        .replace(/\s/g, "")
        .replace(/r\$/gi, "")
        .replace(/\.(?=\d{3}(\D|$))/g, "")
        .replace(",", ".")
        .replace(/[^0-9.]/g, "");

    return Number(normalized) || 0;
}

function loadChatState() {
    try {
        const savedState = JSON.parse(localStorage.getItem(FINANCE_CHAT_STORAGE_KEY));

        chatState = {
            messages: Array.isArray(savedState?.messages)
                ? savedState.messages
                : defaultChatState.messages,
            entries: Array.isArray(savedState?.entries)
                ? savedState.entries
                : defaultChatState.entries
        };
    } catch (error) {
        chatState = JSON.parse(JSON.stringify(defaultChatState));
    }
}

function saveChatState() {
    localStorage.setItem(FINANCE_CHAT_STORAGE_KEY, JSON.stringify(chatState));
}

function detectEntryType(text) {
    const normalized = text.toLowerCase();

    if (/\b(recebi|ganhei|sal[aá]rio|renda|entrada)\b/.test(normalized)) {
        return "income";
    }

    if (/\b(investi|investir|investimento|apliquei|aplicar|aporte)\b/.test(normalized)) {
        return "investment";
    }

    if (/\b(vou gastar|irei gastar|ainda vou|previsto|planejo|futuro|conta que vem|vou pagar|irei pagar)\b/.test(normalized)) {
        return "planned";
    }

    if (/\b(gastei|paguei|comprei|despesa|sa[ií]da|custou)\b/.test(normalized)) {
        return "spent";
    }

    return "spent";
}

function extractDescription(text, amountText) {
    const withoutAmount = text.replace(amountText, "");
    const cleaned = withoutAmount
        .replace(/\b(eu|já|ja|ainda|vou|irei|gastei|gastar|paguei|pagar|comprei|recebi|ganhei|sal[aá]rio|renda|investi|investir|apliquei|aplicar|de|do|da|no|na|em|com|para|previsto|planejo|hoje|amanh[ãa])\b/gi, " ")
        .replace(/\s+/g, " ")
        .trim();

    return cleaned || "Lançamento informado no chat";
}

function parseEntriesFromMessage(message) {
    const normalizedMessage = message.replace(
        /\s+e\s+(?=(gastei|paguei|comprei|vou|irei|ainda|recebi|ganhei|investi|apliquei))/gi,
        "; "
    );
    const parts = normalizedMessage
        .split(/\n|;| e também | também /i)
        .map(part => part.trim())
        .filter(Boolean);

    return parts.flatMap(part => {
        const moneyMatches = [...part.matchAll(/(?:r\$\s*)?\d{1,3}(?:\.\d{3})*(?:,\d{2})?|(?:r\$\s*)?\d+(?:[,.]\d{1,2})?/gi)];

        return moneyMatches
            .map(match => {
                const amount = parseMoney(match[0]);

                if (!amount) return null;

                return {
                    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
                    type: detectEntryType(part),
                    amount,
                    description: extractDescription(part, match[0]),
                    source: part,
                    createdAt: new Date().toISOString()
                };
            })
            .filter(Boolean);
    });
}

function getTotals() {
    return chatState.entries.reduce(
        (totals, entry) => {
            totals[entry.type] += Number(entry.amount) || 0;
            return totals;
        },
        {
            income: 0,
            spent: 0,
            planned: 0,
            investment: 0
        }
    );
}

function getProjectedBalance(totals = getTotals()) {
    return totals.income - totals.spent - totals.planned - totals.investment;
}

function buildInsights(totals = getTotals()) {
    const projectedBalance = getProjectedBalance(totals);
    const totalCommitments = totals.spent + totals.planned + totals.investment;
    const spentRatio = totals.income > 0 ? (totals.spent + totals.planned) / totals.income : 0;
    const insights = [];

    if (totals.income === 0) {
        insights.push("Informe sua renda do mês para o assistente calcular se os gastos cabem no orçamento.");
    }

    if (projectedBalance < 0) {
        insights.push(`Atenção: seu saldo projetado está negativo em ${formatCurrency(Math.abs(projectedBalance))}.`);
    } else {
        insights.push(`Seu saldo projetado é ${formatCurrency(projectedBalance)} depois dos gastos e investimentos informados.`);
    }

    if (spentRatio > 0.8) {
        insights.push("Gastos realizados + previstos já passam de 80% da renda. Revise gastos variáveis antes de assumir novas despesas.");
    } else if (totals.income > 0) {
        insights.push("Os gastos informados ainda estão dentro de uma margem administrável da renda.");
    }

    if (totals.planned > 0) {
        insights.push(`Você ainda planeja gastar ${formatCurrency(totals.planned)}. Reserve esse valor para não confundir com saldo livre.`);
    }

    if (totals.investment > 0 && totalCommitments > 0) {
        insights.push(`Investimentos representam ${Math.round((totals.investment / totalCommitments) * 100)}% dos compromissos informados.`);
    }

    return insights;
}

function buildAssistantReply(parsedEntries) {
    const totals = getTotals();
    const insights = buildInsights(totals);

    if (parsedEntries.length === 0) {
        return "Não encontrei nenhum valor nessa mensagem. Tente algo como: “gastei 50 no almoço”, “vou gastar 900 de aluguel” ou “recebi 5000”.";
    }

    const understood = parsedEntries
        .map(entry => `• ${entryLabels[entry.type]}: ${formatCurrency(entry.amount)} — ${entry.description}`)
        .join("\n");

    return `Entendi e registrei:\n${understood}\n\nResumo agora: ${insights[0]}`;
}

function renderMessages() {
    const chatMessages = document.getElementById("chatMessages");

    chatMessages.innerHTML = chatState.messages
        .map(message => `
            <article class="message ${message.role}">
                <small>${message.role === "user" ? "Você" : "Assistente"}</small>
                <div>${escapeHTML(message.text)}</div>
            </article>
        `)
        .join("");

    chatMessages.scrollTop = chatMessages.scrollHeight;
}

function renderSummary() {
    const totals = getTotals();
    const projectedBalance = getProjectedBalance(totals);
    const insights = buildInsights(totals);

    document.getElementById("incomeTotal").textContent = formatCurrency(totals.income);
    document.getElementById("spentTotal").textContent = formatCurrency(totals.spent);
    document.getElementById("plannedTotal").textContent = formatCurrency(totals.planned);
    document.getElementById("investmentTotal").textContent = formatCurrency(totals.investment);
    document.getElementById("projectedBalance").textContent = formatCurrency(projectedBalance);

    document.getElementById("insightList").innerHTML = insights
        .map(insight => `<li>${escapeHTML(insight)}</li>`)
        .join("");

    document.getElementById("entryCount").textContent = `${chatState.entries.length} ${chatState.entries.length === 1 ? "item" : "itens"}`;
    document.getElementById("entryList").innerHTML = chatState.entries
        .slice()
        .reverse()
        .map(entry => `
            <li class="entry-item">
                <span>
                    <strong>${escapeHTML(entry.description)}</strong>
                    <small>${entryLabels[entry.type]} • ${new Date(entry.createdAt).toLocaleDateString("pt-BR")}</small>
                    <small>Origem: ${escapeHTML(entry.source)}</small>
                </span>
                <span class="entry-value ${entry.type}">${formatCurrency(entry.amount)}</span>
                <button type="button" data-remove-entry="${entry.id}" aria-label="Remover lançamento">×</button>
            </li>
        `)
        .join("");
}

function renderApp() {
    renderMessages();
    renderSummary();
}

function handleChatSubmit(event) {
    event.preventDefault();

    const input = document.getElementById("chatInput");
    const text = input.value.trim();

    if (!text) return;

    chatState.messages.push({
        role: "user",
        text
    });

    const parsedEntries = parseEntriesFromMessage(text);
    chatState.entries.push(...parsedEntries);
    chatState.messages.push({
        role: "assistant",
        text: buildAssistantReply(parsedEntries)
    });

    input.value = "";
    saveChatState();
    renderApp();
}

function clearChat() {
    chatState = JSON.parse(JSON.stringify(defaultChatState));
    saveChatState();
    renderApp();
}

function bindEvents() {
    document.getElementById("chatForm").addEventListener("submit", handleChatSubmit);
    document.getElementById("clearChatBtn").addEventListener("click", clearChat);
    document.getElementById("entryList").addEventListener("click", event => {
        const entryId = event.target.dataset.removeEntry;

        if (!entryId) return;

        chatState.entries = chatState.entries.filter(entry => entry.id !== entryId);
        saveChatState();
        renderSummary();
    });
}

document.addEventListener("DOMContentLoaded", () => {
    loadChatState();
    bindEvents();
    renderApp();
});
