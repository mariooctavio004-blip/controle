// ==========================================
// UTILITÁRIOS
// ==========================================

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

function levenshtein(a, b){
    const m = a.length, n = b.length;
    const dp = Array.from({length: m + 1}, () => new Array(n + 1).fill(0));
    for (let i = 0; i <= m; i++) dp[i][0] = i;
    for (let j = 0; j <= n; j++) dp[0][j] = j;
    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            dp[i][j] = a[i-1] === b[j-1] ? dp[i-1][j-1] : 1 + Math.min(dp[i-1][j-1], dp[i-1][j], dp[i][j-1]);
        }
    }
    return dp[m][n];
}

function matchFarmIndex(name){
    const norm = normalizeName(name);
    let idx = state.farms.findIndex(f => normalizeName(f) === norm);
    if (idx !== -1) return idx;
    idx = state.farms.findIndex(f => {
        const fn = normalizeName(f);
        return fn.includes(norm) || norm.includes(fn);
    });
    if (idx !== -1) return idx;
    let best = -1, bestDist = Infinity;
    state.farms.forEach((f, i) => {
        const dist = levenshtein(normalizeName(f), norm);
        if (dist < bestDist) { bestDist = dist; best = i; }
    });
    return (best !== -1 && bestDist <= 3) ? best : -1;
}
