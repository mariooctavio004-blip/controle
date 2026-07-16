/* ============================================================
   APP.JS
   Ponto de entrada da aplicação.
============================================================ */

document.addEventListener("DOMContentLoaded", async () => {
    await loadState();
    if (typeof initializeSharepoint === "function") initializeSharepoint();
    try {
        if (typeof reapplyImportedDataForCurrentFilter === "function") {
            await reapplyImportedDataForCurrentFilter({ silent: true, save: false });
        }
    } catch (error) {
        console.warn("Não foi possível restaurar os dados importados:", error);
    }
});
