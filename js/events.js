/* ============================================================
   EVENTS.JS
   Todos os addEventListener de botões/inputs de configuração,
   exportação de imagem (PNG/impressão), modal de planilhas e
   menu de configurações.
============================================================ */

function showInlineWarning(msg){
  let box = document.getElementById('inlineWarningBox');
  if(!box){
    box = document.createElement('div');
    box.id = 'inlineWarningBox';
    box.className = 'inline-warning-box';
    document.querySelector('.app').appendChild(box);
  }
  box.textContent = msg;
  box.style.display = 'block';
  clearTimeout(box._timer);
  box._timer = setTimeout(()=>{ box.style.display = 'none'; }, 5000);
}

// =========================
// CAMPOS DE CONFIGURAÇÃO (empresa, título, período, banner)
// =========================

bind("companyName","input",e=>{
    state.company=e.target.value;
    renderReport();
    saveState();
});

bind("reportTitle","input",e=>{
    state.title=e.target.value;
    renderReport();
    saveState();
});

bind("periodLabel","input",e=>{
    state.period=e.target.value;
    renderReport();
    saveState();
});

bind("bannerText","input",e=>{
    state.banner=e.target.value;
    renderReport();
    saveState();
});

// =========================
// BOTÕES DE FAZENDA / DIA / RESET / CONFIG
// =========================

bind("addFarmBtn","click",()=>{
    state.farms.push("Nova fazenda");
    ensureData();
    renderAll();
    saveState();
});

bind("addDayBtn","click",()=>{
    state.days.push("--/--");
    ensureData();
    renderAll();
    saveState();
});

bind("resetBtn","click",()=>{
    document.getElementById("confirmOverlay").style.display="flex";
});

bind("confirmCancel","click",()=>{
    document.getElementById("confirmOverlay").style.display="none";
});

bind("confirmOk","click",()=>{
    state=JSON.parse(JSON.stringify(defaultState));
    ensureData();
    renderAll();
    saveState();
    document.getElementById("confirmOverlay").style.display="none";
});

bind("toggleConfigBtn","click",(e)=>{
    const panel=document.getElementById("configPanel");
    const hidden=panel.style.display==="none";
    panel.style.display=hidden?"block":"none";
    e.target.textContent=hidden?"Ocultar edição":"Mostrar edição";
});

// =========================
// EXPORTAR IMAGEM / IMPRESSÃO
// =========================

function isIOS(){
  return /iP(hone|od|ad)/.test(navigator.platform) ||
    (navigator.userAgent.includes('Mac') && 'ontouchend' in document);
}

// html2canvas e a impressão não capturam bem o valor digitado dentro de
// <input>. Por isso, na hora de gerar a imagem ou imprimir, trocamos
// temporariamente cada campo numérico por um texto simples com o mesmo valor.
function swapInputsForCapture(){
  const report = document.getElementById('report');
  report.querySelectorAll('.num-input').forEach(inp=>{
    const span = document.createElement('span');
    span.className = 'num-print-value';
    span.textContent = inp.value === '' ? '' : inp.value;
    span.style.width = inp.style.width || getComputedStyle(inp).width;
    inp.dataset.wasHidden = '1';
    inp.style.display = 'none';
    inp.insertAdjacentElement('afterend', span);
  });
}
function restoreInputsAfterCapture(){
  const report = document.getElementById('report');
  report.querySelectorAll('.num-print-value').forEach(span=>span.remove());
  report.querySelectorAll('.num-input[data-wasHidden]').forEach(inp=>{
    inp.style.display = '';
    delete inp.dataset.wasHidden;
  });
}

async function generateCanvas(){
  return await html2canvas(document.getElementById('report'), { scale:2, backgroundColor:'#ffffff', useCORS:true });
}

document.getElementById('exportBtn').addEventListener('click', async ()=>{
  const btn = document.getElementById('exportBtn');
  btn.textContent = 'Gerando...';
  swapInputsForCapture();
  try{
    const canvas = await generateCanvas();
    const filename = 'status_envio_' + (state.period||'periodo').replace(/[^a-zA-Z0-9]/g,'_') + '.png';

    if(isIOS()){
      // iOS Safari ignora o atributo download: abrimos a imagem numa aba
      // para o usuário segurar o dedo na imagem e escolher "Salvar imagem".
      const dataUrl = canvas.toDataURL('image/png');
      const win = window.open();
      if(win){
        win.document.write(
          '<html><head><title>' + filename + '</title></head>' +
          '<body style="margin:0;background:#111;display:flex;align-items:center;justify-content:center;">' +
          '<img src="' + dataUrl + '" style="max-width:100%;height:auto;" />' +
          '</body></html>'
        );
      } else {
        showInlineWarning('Seu navegador bloqueou a nova aba. Permita pop-ups para ver a imagem, ou use o botão Imprimir.');
      }
    } else {
      canvas.toBlob(blob=>{
        if(!blob){ showInlineWarning('Não foi possível gerar a imagem. Tente o botão Imprimir.'); return; }
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.download = filename;
        link.href = url;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(()=>URL.revokeObjectURL(url), 4000);
      }, 'image/png');
    }
  }catch(e){
    console.error(e);
    showInlineWarning('Não foi possível gerar a imagem. Tente o botão Imprimir, ou baixar de novo.');
  }
  restoreInputsAfterCapture();
  btn.textContent = 'Baixar imagem';
});

window.addEventListener('beforeprint', swapInputsForCapture);
window.addEventListener('afterprint', restoreInputsAfterCapture);

/* ============================================================
   MODAL DAS PLANILHAS
============================================================ */

const workbookModal = document.getElementById("workbookModal");
const openWorkbookModal = document.getElementById("openWorkbookModal");
const closeWorkbookModal = document.getElementById("closeWorkbookModal");

const connectionIndicator = document.getElementById("connectionIndicator");
const currentWorkbook = document.getElementById("currentWorkbook");
const disconnectWorkbook = document.getElementById("disconnectWorkbook");

/* input oculto para abrir o seletor de arquivos */
const workbookInput = document.createElement("input");
workbookInput.type = "file";
workbookInput.accept = ".xlsx,.xls";
workbookInput.style.display = "none";
document.body.appendChild(workbookInput);

const connectedWorkbooks = {
    campo: null,
    abastecimento: null,
    diario: null,
    mensal: null,
    divergencias: null
};
let selectedWorkbookType = null;
function updateConnectionStatus() {

    const list = Object.values(connectedWorkbooks).filter(Boolean);

    if (list.length === 0) {
        connectionIndicator.innerHTML = "🔴 Nenhuma planilha conectada";
        currentWorkbook.innerHTML = "Selecione uma planilha";
        return;
    }

    connectionIndicator.innerHTML = `🟢 ${list.length} planilha(s) conectada(s)`;

    currentWorkbook.innerHTML = list.map((w, index) => `
        <div class="connected-item">
            ✔ ${w.title}<br>
            <small>${w.fileName}</small>
        </div>
    `).join("");

}
/* ===========================
   Abrir Modal
=========================== */

openWorkbookModal?.addEventListener("click", () => {

    workbookModal.style.display = "flex";

});

/* ===========================
   Fechar Modal
=========================== */

closeWorkbookModal?.addEventListener("click", () => {

    workbookModal.style.display = "none";

});

/* ===========================
   Escolha da planilha
=========================== */

document.querySelectorAll(".workbook-option").forEach(btn => {

    btn.addEventListener("click", () => {

        selectedWorkbookType = btn.dataset.id;

        workbookInput.click();

    });

});

/* ===========================
   Arquivo escolhido
=========================== */

workbookInput.addEventListener("change", e => {

    const file = e.target.files[0];

    if (!file) return;

    const option = document.querySelector(
        `.workbook-option[data-id="${selectedWorkbookType}"]`
    );

    connectedWorkbooks[selectedWorkbookType] = {
        id: selectedWorkbookType,
        title: option.innerText,
        fileName: file.name,
        lastUpdate: new Date().toLocaleTimeString(),
        file
    };

    updateConnectionStatus();

    workbookModal.style.display = "none";

});
/* ===========================
   Desvincular
=========================== */

disconnectWorkbook?.addEventListener("click", () => {

    if (!selectedWorkbookType) {
        alert("Selecione uma planilha primeiro.");
        return;
    }

    if (!connectedWorkbooks[selectedWorkbookType]) {
        alert("Nenhuma planilha conectada neste painel.");
        return;
    }

    if (!confirm("Deseja remover esta planilha?")) return;

    connectedWorkbooks[selectedWorkbookType] = null;

    workbookInput.value = "";

    updateConnectionStatus();

});

/* ===========================
   Fechar clicando fora
=========================== */

window.addEventListener("click", e => {

    if (e.target === workbookModal) {

        workbookModal.style.display = "none";

    }

});

/* ============================================================
   MENU DE CONFIGURAÇÕES
============================================================ */

const configBtn = document.getElementById("configMenuBtn");
const configMenu = document.getElementById("configMenu");

configBtn?.addEventListener("click", (e)=>{

    e.stopPropagation();

    configMenu.classList.toggle("show");

});

window.addEventListener("click", ()=>{

    configMenu.classList.remove("show");

});
