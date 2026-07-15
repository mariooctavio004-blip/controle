/* ============================================================
   APP.JS
   Ponto de entrada da aplicação.
============================================================ */

document.addEventListener("DOMContentLoaded", async () => {
    await loadState();

    if (typeof initializeSharepoint === "function") {
        initializeSharepoint();
    }
});
