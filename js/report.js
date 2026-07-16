/* ============================================================
   REPORT.JS
   Construção dos cards do relatório:
   - painéis diários;
   - mapa mensal;
   - divergências.
============================================================ */


/* ============================================================
   CABEÇALHO DOS CARDS
============================================================ */

function buildCardHeader(panel, subtitle = "") {
    const subtitleHTML = subtitle
        ? `<small class="rep-card-subtitle">${escapeHtml(subtitle)}</small>`
        : "";

    return `
        <div class="rep-card-head">

            <div class="rep-card-title-group">

                ${panelIconHTML(panel.key)}

                <div class="rep-card-title-text">

                    <span>
                        ${escapeHtml(panel.title)}
                    </span>

                    ${subtitleHTML}

                </div>

            </div>

        </div>
    `;
}


/* ============================================================
   STATUS
============================================================ */

function getStatusLabel(status) {
    const labels = {
        ok: "Recebido",
        no: "Pendente",
        blank: "Não informado"
    };

    return labels[status] || "Não informado";
}

function buildStatusIcon(status, attributes = "") {
    const normalized =
        normalizeStatus(status);

    const className =
        normalized === "ok"
            ? "ok"
            : normalized === "no"
                ? "no"
                : "blank";

    const glyph =
        normalized === "ok"
            ? "✔"
            : normalized === "no"
                ? "✕"
                : "—";

    return `
        <button
            type="button"
            class="rep-icon ${className}"
            data-status="${normalized}"
            title="${getStatusLabel(normalized)}"
            aria-label="${getStatusLabel(normalized)}"
            ${attributes}
        >
            ${glyph}
        </button>
    `;
}


/* ============================================================
   PAINÉIS DIÁRIOS
============================================================ */

function buildDailyCard(panel) {
    const visibleDayIndexes =
        getVisibleDayIndexes();

    const pendingByDay =
        visibleDayIndexes.map(() => 0);

    let totalPending = 0;
    let bodyRows = "";

    state.farms.forEach(
        (farm, farmIndex) => {
            let cells = "";
            let farmPending = 0;

            visibleDayIndexes.forEach(
                (dayIndex, visibleColumnIndex) => {
                    const status =
                        normalizeStatus(
                            state.data?.[panel.key]?.[
                                farmIndex
                            ]?.[dayIndex]
                        );

                    if (status === "no") {
                        farmPending++;
                        pendingByDay[
                            visibleColumnIndex
                        ]++;
                    }

                    const attributes = [
                        `data-panel="${escapeHtml(panel.key)}"`,
                        `data-farm="${farmIndex}"`,
                        `data-day="${dayIndex}"`
                    ].join(" ");

                    cells += `
                        <td class="status-cell">
                            ${buildStatusIcon(
                                status,
                                attributes
                            )}
                        </td>
                    `;
                }
            );

            totalPending += farmPending;

            bodyRows += `
                <tr>

                    <td class="farm-col">
                        ${escapeHtml(farm)}
                    </td>

                    ${cells}

                    <td
                        class="
                            rep-total
                            pend-col
                            ${
                                farmPending > 0
                                    ? "has-pending"
                                    : "no-pending"
                            }
                        "
                    >
                        ${farmPending}
                    </td>

                </tr>
            `;
        }
    );

    const dayHeaders =
        visibleDayIndexes
            .map(dayIndex => `
                <th class="day-col">
                    ${escapeHtml(
                        state.days[dayIndex]
                    )}
                </th>
            `)
            .join("");

    const footerCells =
        pendingByDay
            .map(value => `
                <td
                    class="${
                        value > 0
                            ? "has-pending"
                            : "no-pending"
                    }"
                >
                    ${value}
                </td>
            `)
            .join("");

    return `
        ${buildCardHeader(panel)}

        <div class="rep-table-wrapper">

            <table class="rep-table rep-table-daily">

                <thead>

                    <tr>

                        <th class="farm-col">
                            Fazenda
                        </th>

                        ${dayHeaders}

                        <th class="pend-col">
                            Pendências
                        </th>

                    </tr>

                </thead>

                <tbody>

                    ${bodyRows}

                </tbody>

                <tfoot>

                    <tr class="pend-row">

                        <td class="farm-col">
                            PENDÊNCIAS
                        </td>

                        ${footerCells}

                        <td
                            class="
                                total-pending-cell
                                ${
                                    totalPending > 0
                                        ? "has-pending"
                                        : "no-pending"
                                }
                            "
                        >
                            ${totalPending}
                        </td>

                    </tr>

                </tfoot>

            </table>

        </div>
    `;
}


/* ============================================================
   MAPA MENSAL
============================================================ */

function buildMonthlyCard(panel) {
    let pendingCount = 0;
    let bodyRows = "";

    state.farms.forEach(
        (farm, farmIndex) => {
            const status =
                normalizeStatus(
                    state.monthly?.[farmIndex]
                );

            if (status === "no") {
                pendingCount++;
            }

            bodyRows += `
                <tr>

                    <td class="farm-col">
                        ${escapeHtml(farm)}
                    </td>

                    <td class="status-cell">

                        ${buildStatusIcon(
                            status,
                            `data-monthly="${farmIndex}"`
                        )}

                    </td>

                </tr>
            `;
        }
    );

    return `
        ${buildCardHeader(
            panel,
            state.monthLabel || ""
        )}

        <div class="rep-table-wrapper">

            <table class="rep-table rep-table-monthly">

                <thead>

                    <tr>

                        <th class="farm-col">
                            Fazenda
                        </th>

                        <th class="pend-col">
                            Status
                        </th>

                    </tr>

                </thead>

                <tbody>

                    ${bodyRows}

                </tbody>

                <tfoot>

                    <tr class="pend-row">

                        <td class="farm-col">
                            PENDÊNCIAS
                        </td>

                        <td
                            class="${
                                pendingCount > 0
                                    ? "has-pending"
                                    : "no-pending"
                            }"
                        >
                            ${pendingCount}
                        </td>

                    </tr>

                </tfoot>

            </table>

        </div>
    `;
}


/* ============================================================
   DIVERGÊNCIAS
============================================================ */

function toNumericValue(value) {
    if (
        value === "" ||
        value === null ||
        value === undefined
    ) {
        return null;
    }

    const number =
        Number(
            String(value)
                .replace(",", ".")
        );

    return Number.isFinite(number)
        ? number
        : null;
}

function calculateDifference(firstValue, secondValue) {
    const first =
        toNumericValue(firstValue);

    const second =
        toNumericValue(secondValue);

    if (
        first === null ||
        second === null
    ) {
        return null;
    }

    return first - second;
}

function getDifferenceClass(value) {
    if (value === null) {
        return "empty";
    }

    if (value === 0) {
        return "zero";
    }

    return value > 0
        ? "positive"
        : "negative";
}

function formatDifference(value) {
    if (value === null) {
        return "—";
    }

    return String(value);
}

function buildNumberInput({
    farmIndex,
    field,
    value,
    label
}) {
    return `
        <input
            class="num-input"
            type="text"
            inputmode="numeric"
            autocomplete="off"
            data-fi="${farmIndex}"
            data-field="${field}"
            value="${escapeHtml(value ?? "")}"
            aria-label="${escapeHtml(label)}"
        >
    `;
}

function buildDivergCard(panel) {
    let bodyRows = "";

    let totalBirthDifference = 0;
    let totalDeathDifference = 0;

    let hasBirthDifference = false;
    let hasDeathDifference = false;

    state.farms.forEach(
        (farm, farmIndex) => {
            const data =
                state.diverg?.[farmIndex] || {
                    nd: "",
                    ns: "",
                    md: "",
                    ms: ""
                };

            const birthDifference =
                calculateDifference(
                    data.nd,
                    data.ns
                );

            const deathDifference =
                calculateDifference(
                    data.md,
                    data.ms
                );

            if (birthDifference !== null) {
                totalBirthDifference +=
                    birthDifference;

                hasBirthDifference = true;
            }

            if (deathDifference !== null) {
                totalDeathDifference +=
                    deathDifference;

                hasDeathDifference = true;
            }

            bodyRows += `
                <tr>

                    <td class="farm-col">
                        ${escapeHtml(farm)}
                    </td>

                    <td>
                        ${buildNumberInput({
                            farmIndex,
                            field: "nd",
                            value: data.nd,
                            label:
                                `Nascimentos no diário — ${farm}`
                        })}
                    </td>

                    <td>
                        ${buildNumberInput({
                            farmIndex,
                            field: "ns",
                            value: data.ns,
                            label:
                                `Nascimentos no sistema — ${farm}`
                        })}
                    </td>

                    <td
                        class="
                            diff-val
                            ${getDifferenceClass(
                                birthDifference
                            )}
                        "
                    >
                        ${formatDifference(
                            birthDifference
                        )}
                    </td>

                    <td>
                        ${buildNumberInput({
                            farmIndex,
                            field: "md",
                            value: data.md,
                            label:
                                `Mortes no diário — ${farm}`
                        })}
                    </td>

                    <td>
                        ${buildNumberInput({
                            farmIndex,
                            field: "ms",
                            value: data.ms,
                            label:
                                `Mortes no sistema — ${farm}`
                        })}
                    </td>

                    <td
                        class="
                            diff-val
                            ${getDifferenceClass(
                                deathDifference
                            )}
                        "
                    >
                        ${formatDifference(
                            deathDifference
                        )}
                    </td>

                </tr>
            `;
        }
    );

    const formattedBirthTotal =
        hasBirthDifference
            ? totalBirthDifference
            : "—";

    const formattedDeathTotal =
        hasDeathDifference
            ? totalDeathDifference
            : "—";

    return `
        ${buildCardHeader(panel)}

        <div class="rep-table-wrapper">

            <table class="rep-table rep-table-diverg">

                <thead>

                    <tr>

                        <th
                            class="farm-col"
                            rowspan="2"
                        >
                            Fazenda
                        </th>

                        <th
                            class="group-header"
                            colspan="3"
                        >
                            Nascimentos
                        </th>

                        <th
                            class="group-header"
                            colspan="3"
                        >
                            Mortes
                        </th>

                    </tr>

                    <tr>

                        <th>
                            Diário
                        </th>

                        <th>
                            Sistema
                        </th>

                        <th>
                            Diferença
                        </th>

                        <th>
                            Diário
                        </th>

                        <th>
                            Sistema
                        </th>

                        <th>
                            Diferença
                        </th>

                    </tr>

                </thead>

                <tbody>

                    ${bodyRows}

                </tbody>

                <tfoot>

                    <tr class="pend-row pend-row-navy">

                        <td class="farm-col">
                            TOTAL
                        </td>

                        <td></td>

                        <td></td>

                        <td
                            class="
                                diff-val
                                ${getDifferenceClass(
                                    hasBirthDifference
                                        ? totalBirthDifference
                                        : null
                                )}
                            "
                        >
                            ${formattedBirthTotal}
                        </td>

                        <td></td>

                        <td></td>

                        <td
                            class="
                                diff-val
                                ${getDifferenceClass(
                                    hasDeathDifference
                                        ? totalDeathDifference
                                        : null
                                )}
                            "
                        >
                            ${formattedDeathTotal}
                        </td>

                    </tr>

                </tfoot>

            </table>

        </div>
    `;
}
