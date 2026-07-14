// ==========================================
// IMPORTAÇÃO DE PLANILHAS
// ==========================================

function processWorkbook(workbook, fallbackName) {

    workbook.SheetNames.forEach(sheetName => {

        let key = classifySheetName(sheetName);

        if (!key && workbook.SheetNames.length === 1) {

            key = classifySheetName(fallbackName);

        }

        if (!key) return;

        const sheet = workbook.Sheets[sheetName];

        importedData[key] = XLSX.utils.sheet_to_json(sheet, {

            header: 1,

            raw: false,

            dateNF: "dd/mm"

        });

    });

}

document
    .getElementById("excelFiles")
    ?.addEventListener("change", importarPlanilhas);

async function importarPlanilhas(event) {

    const files = [...event.target.files];

    if (!files.length)
        return;

    for (const file of files) {

        const data = await file.arrayBuffer();

        const workbook = XLSX.read(data, {

            type: "array",

            cellDates: true

        });

        processWorkbook(workbook, file.name);

    }

    event.target.value = "";

    preencherSistema();

}
