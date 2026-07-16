/* ============================================================
   APP.JS
   Inicialização única e reconstrução segura dos dados.
============================================================ */

document.addEventListener("DOMContentLoaded", async () => {
    await loadState();

    if (typeof initializeSharepoint === "function") {
        initializeSharepoint();
    }

    try {
        /*
         * Nunca confia nos dados derivados salvos no state.
         * Reconstrói tudo exclusivamente a partir das fontes ativas:
         * arquivos locais realmente existentes no IndexedDB e URLs.
         */
        if (typeof clearImportedData === "function") {
            clearImportedData();
        }

        if (typeof loadLocalWorkbooksIntoImportedData === "function") {
            await loadLocalWorkbooksIntoImportedData({
                purgeOrphans: true
            });
        }

        if (typeof syncAllSharepointSheets === "function") {
            await syncAllSharepointSheets(true);
        }

        if (typeof preencherSistema === "function") {
            preencherSistema({
                silent: true,
                save: true,
                render: true
            });
        }

        if (typeof renderImportedWorkbookStatus === "function") {
            renderImportedWorkbookStatus();
        }
    } catch (error) {
        console.error(
            "Não foi possível reconstruir os dados a partir das fontes ativas:",
            error
        );

        if (typeof clearAllDerivedPanels === "function") {
            clearAllDerivedPanels();
        }

        if (typeof renderAll === "function") {
            renderAll();
        }
    }
});
