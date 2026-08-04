/* ============================================================
   FINANCE-AGENT.JS
   Assistente financeiro em formato de chat.

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
   Pierre: agente local para controle de gastos e investimentos.
   Agente local para controle de gastos, orçamento e investimentos.

const FINANCE_AGENT_STORAGE_KEY = "financeAgentStateV1";

const financeAgentDefaultState = {
    monthlyIncome: 5000,
    monthlyBudget: 3500,
    investmentGoal: 1000,
    transactions: [
        {
            description: "Mercado",
            category: "Alimentação",
            type: "expense",
            amount: 620,
            date: "2026-08-01"
        },
        {
            description: "Tesouro Direto",
            category: "Investimentos",
            type: "investment",
            amount: 850,
            date: "2026-08-02"
        }
    ]
};

let financeAgentState = null;

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
function parseMoneyInput(value) {
    const normalized = String(value || "")
        .replace(/\./g, "")
        .replace(",", ".")
        .replace(/[^0-9.-]/g, "");

    return Math.max(0, Number(normalized) || 0);
}

function loadFinanceAgentState() {
    try {
        const savedState = JSON.parse(
            localStorage.getItem(FINANCE_AGENT_STORAGE_KEY)
        );

        financeAgentState = {
            ...financeAgentDefaultState,
            ...savedState,
            transactions: Array.isArray(savedState?.transactions)
                ? savedState.transactions
                : financeAgentDefaultState.transactions
        };
    } catch (error) {
        financeAgentState = JSON.parse(
            JSON.stringify(financeAgentDefaultState)
        );
    }
}

function saveFinanceAgentState() {
    localStorage.setItem(
        FINANCE_AGENT_STORAGE_KEY,
        JSON.stringify(financeAgentState)
    );
}

function getFinanceTotals() {
    return financeAgentState.transactions.reduce(
        (totals, transaction) => {
            const amount = Number(transaction.amount) || 0;

            if (transaction.type === "income") {
                totals.income += amount;
            } else if (transaction.type === "investment") {
                totals.investments += amount;
            } else {
                totals.expenses += amount;
                totals.byCategory[transaction.category] =
                    (totals.byCategory[transaction.category] || 0) + amount;
            }

            return totals;
        },
        {
            income: Number(financeAgentState.monthlyIncome) || 0,
            expenses: 0,
            investments: 0,
            byCategory: {}
        }
    );
}

function getFinanceAgentAdvice(totals) {
    const balance = totals.income - totals.expenses - totals.investments;
    const budgetUsage = financeAgentState.monthlyBudget > 0
        ? totals.expenses / financeAgentState.monthlyBudget
        : 0;
    const investmentProgress = financeAgentState.investmentGoal > 0
        ? totals.investments / financeAgentState.investmentGoal
        : 0;
    const topCategory = Object.entries(totals.byCategory)
        .sort((a, b) => b[1] - a[1])[0];

    const advice = [];

    if (budgetUsage >= 1) {
        advice.push("Pierre alerta: o limite de gastos passou do combinado. Pause compras não essenciais e revise a maior categoria.");
    } else if (budgetUsage >= 0.8) {
        advice.push("Pierre percebeu que mais de 80% do limite já foi usado. Priorize contas obrigatórias e adie gastos variáveis.");
    } else {
        advice.push("Pierre vê um orçamento saudável. Continue registrando tudo para manter a previsão confiável.");
    }

    if (investmentProgress < 0.5) {
        advice.push("Pierre recomenda antecipar o aporte: a meta de investimento ainda está abaixo do ritmo ideal.");
    } else {
        advice.push("Pierre gostou do ritmo de investimento. Mantenha a constância e diversifique conforme seu perfil de risco.");
    }

    if (balance < 0) {
        advice.push("Pierre calculou saldo negativo. Corte despesas recorrentes antes de assumir novos compromissos.");
    } else if (topCategory) {
        advice.push(`Pierre encontrou o maior foco de economia: ${topCategory[0]}, com ${formatCurrency(topCategory[1])} em gastos.`);
        advice.push("Seu orçamento de gastos foi ultrapassado. Trave compras não essenciais até revisar as categorias mais caras.");
    } else if (budgetUsage >= 0.8) {
        advice.push("Você já usou mais de 80% do orçamento. Priorize contas obrigatórias e adie gastos variáveis.");
    } else {
        advice.push("O orçamento está saudável. Continue registrando os gastos para manter a previsão do mês confiável.");
    }

    if (investmentProgress < 0.5) {
        advice.push("A meta de investimento está atrasada. Considere separar o aporte assim que a renda entrar.");
    } else {
        advice.push("Seu ritmo de investimento está bom. Mantenha a reserva e diversifique conforme seu perfil de risco.");
    }

    if (balance < 0) {
        advice.push("O saldo projetado está negativo. Reduza despesas recorrentes antes de assumir novos compromissos.");
    } else if (topCategory) {
        advice.push(`Maior foco de economia: ${topCategory[0]}, com ${formatCurrency(topCategory[1])} em gastos.`);
    }

    return advice;
}

function renderFinanceAgent() {
    if (!financeAgentState) return;

    const totals = getFinanceTotals();
    const balance = totals.income - totals.expenses - totals.investments;
    const budgetUsage = financeAgentState.monthlyBudget > 0
        ? Math.min(100, (totals.expenses / financeAgentState.monthlyBudget) * 100)
        : 0;
    const investmentProgress = financeAgentState.investmentGoal > 0
        ? Math.min(100, (totals.investments / financeAgentState.investmentGoal) * 100)
        : 0;

    document.getElementById("financeIncomeValue").textContent = formatCurrency(totals.income);
    document.getElementById("financeExpenseValue").textContent = formatCurrency(totals.expenses);
    document.getElementById("financeInvestmentValue").textContent = formatCurrency(totals.investments);
    document.getElementById("financeBalanceValue").textContent = formatCurrency(balance);
    document.getElementById("financeBudgetProgress").style.width = `${budgetUsage}%`;
    document.getElementById("financeInvestmentProgress").style.width = `${investmentProgress}%`;
    document.getElementById("financeBudgetText").textContent = `${budgetUsage.toFixed(0)}% do orçamento usado`;
    document.getElementById("financeInvestmentText").textContent = `${investmentProgress.toFixed(0)}% da meta investida`;

    const mainAdvice = document.getElementById("mainAdvice");
    if (mainAdvice) {
        mainAdvice.textContent = getFinanceAgentAdvice(totals)[0] ||
            "Registre seus valores para eu acompanhar orçamento, saldo e aportes.";
    }

    const adviceList = document.getElementById("financeAdviceList");
    adviceList.innerHTML = getFinanceAgentAdvice(totals)
        .map(item => `<li>${escapeHTML(item)}</li>`)
        .join("");

    const transactionList = document.getElementById("financeTransactionList");
    transactionList.innerHTML = financeAgentState.transactions
        .slice()
        .reverse()
        .map((transaction, index) => {
            const originalIndex = financeAgentState.transactions.length - 1 - index;
            return `
                <li>
                    <span>
                        <strong>${escapeHTML(transaction.description)}</strong>
                        <small>${escapeHTML(transaction.category)} • ${escapeHTML(transaction.date)}</small>
                    </span>
                    <span>${formatCurrency(transaction.amount)}</span>
                    <button type="button" data-finance-remove="${originalIndex}" aria-label="Remover lançamento">×</button>
                </li>
            `;
        })
        .join("");
}

function bindFinanceAgentEvents() {
    const form = document.getElementById("financeTransactionForm");
    const settingsForm = document.getElementById("financeSettingsForm");

    form?.addEventListener("submit", event => {
        event.preventDefault();

        const formData = new FormData(form);

        financeAgentState.transactions.push({
            description: formData.get("description") || "Lançamento",
            category: formData.get("category") || "Outros",
            type: formData.get("type") || "expense",
            amount: parseMoneyInput(formData.get("amount")),
            date: formData.get("date") || new Date().toISOString().slice(0, 10)
        });

        form.reset();
        document.getElementById("financeDate").value = new Date().toISOString().slice(0, 10);
        saveFinanceAgentState();
        renderFinanceAgent();
    });

    settingsForm?.addEventListener("submit", event => {
        event.preventDefault();

        const formData = new FormData(settingsForm);
        financeAgentState.monthlyIncome = parseMoneyInput(formData.get("monthlyIncome"));
        financeAgentState.monthlyBudget = parseMoneyInput(formData.get("monthlyBudget"));
        financeAgentState.investmentGoal = parseMoneyInput(formData.get("investmentGoal"));

        saveFinanceAgentState();
        renderFinanceAgent();
    });

    document.getElementById("financeTransactionList")?.addEventListener("click", event => {
        const removeIndex = event.target.dataset.financeRemove;

        if (removeIndex === undefined) return;

        financeAgentState.transactions.splice(Number(removeIndex), 1);
        saveFinanceAgentState();
        renderFinanceAgent();
    });
}

document.addEventListener("DOMContentLoaded", () => {
    loadChatState();
    bindEvents();
    renderApp();
    loadFinanceAgentState();

    document.getElementById("financeDate").value = new Date().toISOString().slice(0, 10);
    document.getElementById("financeMonthlyIncome").value = financeAgentState.monthlyIncome;
    document.getElementById("financeMonthlyBudget").value = financeAgentState.monthlyBudget;
    document.getElementById("financeInvestmentGoal").value = financeAgentState.investmentGoal;

    bindFinanceAgentEvents();
    renderFinanceAgent();
});
