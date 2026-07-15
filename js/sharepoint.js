/* ============================================================
   SHAREPOINT.JS
   Anexar/sincronizar planilha online do SharePoint/OneDrive:
   conversão de link, download, status visual e sincronização
   automática periódica.
============================================================ */

// Tenta transformar um link de compartilhamento do SharePoint/OneDrive em um
// link de download direto (adiciona "download=1" na URL). Isso só funciona
// se o arquivo estiver compartilhado como "Qualquer pessoa com o link" e se o
// SharePoint permitir o acesso direto do navegador (CORS); alguns tenants
// bloqueiam esse acesso, e nesse caso a sincronização automática falha e um
// aviso aparece na tela.
function toDirectDownloadUrl(url){
    try{
        const u = new URL(url);
        const host = u.hostname.toLowerCase();
        if(host.includes('sharepoint.com') || host.includes('1drv.ms') || host.includes('onedrive.live.com')){
            if(!u.searchParams.has('download')) u.searchParams.set('download', '1');
        }
        return u.toString();
    }catch(e){
        return url;
    }
}

function renderAttachStatus(isError){
    const icon = document.getElementById('attachStatusIcon');
    const text = document.getElementById('attachStatusText');
    const removeBtn = document.getElementById('removeAttachBtn');
    const urlInput = document.getElementById('sharepointUrl');
    const info = document.querySelector('.attach-info');
    if(!icon || !text || !removeBtn || !urlInput) return;

    if(state.sharepointUrl){
        if(document.activeElement !== urlInput) urlInput.value = state.sharepointUrl;
        removeBtn.style.display = '';
        info.classList.toggle('is-error', !!isError);
        icon.textContent = isError ? '⚠️' : '🔗';
        if(isError){
            text.textContent = 'Não foi possível sincronizar a planilha anexada agora. Tentando de novo automaticamente.';
        } else {
            text.textContent = 'Planilha anexada' + (state.lastSync ? (' — última sincronização: ' + state.lastSync) : ' — sincronizando...') + '.';
        }
    } else {
        removeBtn.style.display = 'none';
        info.classList.remove('is-error');
        icon.textContent = '📎';
        text.textContent = 'Nenhuma planilha online anexada.';
    }
}

// Baixa e reprocessa a planilha anexada. silent=true evita mostrar aviso de
// erro repetido nas sincronizações automáticas de fundo.
async function syncSharepointSheet(silent){
    const url = state.sharepointUrl;
    if(!url) return;
    try{
        const direct = toDirectDownloadUrl(url);
        const res = await fetch(direct, { mode: 'cors' });
        if(!res.ok) throw new Error('HTTP ' + res.status);
        const buffer = await res.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
        processWorkbook(workbook, 'planilha');
        state.lastSync = new Date().toLocaleTimeString('pt-BR', { hour:'2-digit', minute:'2-digit' });
        preencherSistema(); // já faz ensureData + renderAll + saveState + aviso do que foi importado
        renderAttachStatus(false);
    }catch(e){
        console.error(e);
        renderAttachStatus(true);
        if(!silent){
            showInlineWarning('Não foi possível baixar a planilha do link informado. Confira se ela está compartilhada como "Qualquer pessoa com o link pode visualizar" — alguns links do SharePoint bloqueiam o acesso direto do navegador (CORS) e, nesse caso, não é possível sincronizar automaticamente. Detalhe: ' + e.message);
        }
    }
}

let sharepointSyncTimer = null;
function startSharepointAutoSync(){
    clearInterval(sharepointSyncTimer);
    sharepointSyncTimer = setInterval(() => {
        if(state.sharepointUrl) syncSharepointSheet(true);
    }, 3 * 60 * 1000); // ressincroniza a cada 3 minutos
}

document.getElementById('attachSharepointBtn')?.addEventListener('click', async () => {
    const val = document.getElementById('sharepointUrl').value.trim();
    if(!val){ showInlineWarning('Cole o link da planilha antes de anexar.'); return; }
    state.sharepointUrl = val;
    state.lastSync = '';
    saveState();
    renderAttachStatus(false);
    await syncSharepointSheet(false);
});

document.getElementById('removeAttachBtn')?.addEventListener('click', () => {
    state.sharepointUrl = '';
    state.lastSync = '';
    document.getElementById('sharepointUrl').value = '';
    renderAttachStatus(false);
    saveState();
});
