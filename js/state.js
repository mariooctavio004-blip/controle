// ==========================================
// ESTADO PADRÃO DO SISTEMA
// ==========================================

const defaultState = {

    company: "Nelore Cometa",

    title: "STATUS DE ENVIO DOS CONTROLES - FAZENDAS",

    period: "01/07 A 04/07",

    monthLabel: "FECHAMENTO – JUNHO/2025",

    banner:
        "Contamos com todos para mantermos a disciplina e a qualidade das informações!",

    farms: [
        "Furna Linda",
        "Estância Cometa",
        "Vale do Jaurú",
        "Pantanal 1",
        "Pantanal 2",
        "São Lucas",
        "Estância Maristela",
        "Santa Rita",
        "Liberdade",
        "São Sebastião"
    ],

    days: [
        "01/07",
        "02/07",
        "03/07",
        "04/07"
    ],

    selected: [
        "campo",
        "abastecimento",
        "diario",
        "mensal"
    ],

    data: {},

    monthly: {},

    diverg: {},

    sharepoint: {

        campo: {
            name: "Operações em Campo",
            url: "",
            connected: false,
            lastSync: ""
        },

        abastecimento: {
            name: "Abastecimento",
            url: "",
            connected: false,
            lastSync: ""
        },

        diario: {
            name: "Diário de Campo",
            url: "",
            connected: false,
            lastSync: ""
        },

        mensal: {
            name: "Mapa Mensal do Rebanho",
            url: "",
            connected: false,
            lastSync: ""
        },

        divergencias: {
            name: "Divergências entre Diário e Sistema",
            url: "",
            connected: false,
            lastSync: ""
        }

    },

    filterStart: "",

    filterEnd: ""

};

let state = null;

const importedData = {

    campo: null,

    abastecimento: null,

    diario: null,

    divergencias: null

};

function normalizeStatus(v){
  if(v === true) return 'ok';
  if(v === false) return 'no';
  if(v === 'ok' || v === 'no' || v === 'blank') return v;
  return 'ok';
}
