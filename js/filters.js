/* ============================================================
   FILTERS.JS
   Filtro de período: mostrar só os últimos N dias / um
   intervalo escolhido, sem alterar os dados salvos.
============================================================ */

// Converte um rótulo de coluna tipo "01/07" ou "01/07/2026" em Date. Se não
// conseguir entender o formato, retorna null (a coluna some do cálculo de
// filtro, mas continua sendo mostrada por segurança).
function parseDayLabel(label){
    const m = String(label == null ? '' : label).trim().match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
    if(!m) return null;
    const day = parseInt(m[1], 10), month = parseInt(m[2], 10) - 1;
    let year = m[3] ? parseInt(m[3], 10) : new Date().getFullYear();
    if(year < 100) year += 2000;
    const d = new Date(year, month, day);
    return isNaN(d.getTime()) ? null : d;
}

function isoToDate(iso){
    if(!iso) return null;
    const d = new Date(iso + 'T00:00:00');
    return isNaN(d.getTime()) ? null : d;
}

function dateToISO(d){
    const y = d.getFullYear(), m = String(d.getMonth()+1).padStart(2,'0'), day = String(d.getDate()).padStart(2,'0');
    return `${y}-${m}-${day}`;
}

// Retorna os índices de state.days que devem aparecer nos painéis, já
// aplicando o filtro de período (campos de data ao lado de "Importar
// Planilhas"). Os dados de todos os dias continuam salvos normalmente —
// só a exibição é filtrada aqui.
function getVisibleDayIndexes(){
    const allIdx = state.days.map((_, i) => i);
    let endDate = isoToDate(state.filterEnd);
    let startDate = isoToDate(state.filterStart);

    if(!endDate && !startDate) return allIdx; // sem filtro definido: mostra tudo

    const parsedDays = state.days.map(parseDayLabel);
    const allParsed = parsedDays.every(d => d !== null);

    // Se as colunas não estiverem no formato DD/MM, não dá pra comparar datas
    // com segurança. Nesse caso, só a data final é usada para pegar as
    // últimas 7 colunas (por posição).
    if(!allParsed){
        if(endDate && !startDate) return allIdx.slice(Math.max(0, allIdx.length - 7));
        return allIdx;
    }

    if(endDate && !startDate){
        startDate = new Date(endDate);
        startDate.setDate(startDate.getDate() - 6); // janela padrão de 7 dias
    }
    if(startDate && !endDate){
        endDate = new Date(startDate);
        endDate.setDate(endDate.getDate() + 6);
    }

    const filtered = allIdx.filter(i => {
        const d = parsedDays[i];
        if(!d) return true;
        return d >= startDate && d <= endDate;
    });
    return filtered.length ? filtered : allIdx;
}
