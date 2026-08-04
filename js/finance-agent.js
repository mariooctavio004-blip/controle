/* ============================================================
   FINANCE-AGENT.JS
   Agente local para controle de gastos, orçamento e investimentos.
============================================================ */

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
    loadFinanceAgentState();

    document.getElementById("financeDate").value = new Date().toISOString().slice(0, 10);
    document.getElementById("financeMonthlyIncome").value = financeAgentState.monthlyIncome;
    document.getElementById("financeMonthlyBudget").value = financeAgentState.monthlyBudget;
    document.getElementById("financeInvestmentGoal").value = financeAgentState.investmentGoal;

    bindFinanceAgentEvents();
    renderFinanceAgent();
});
