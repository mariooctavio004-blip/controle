/* ============================================================
   REPORT.JS
   Construção do HTML de cada tipo de card do relatório: diário,
   mensal e de divergências.
============================================================ */

function buildDailyCard(panel){
  const visible = getVisibleDayIndexes(); // índices de state.days a exibir, já filtrados pelo período escolhido
  let rows = '';
  const perDayPend = visible.map(()=>0);
  let totalPendAll = 0;
  state.farms.forEach((farm, fi)=>{
    let cells = '', totalPend = 0;
    visible.forEach((di, col)=>{
      const status = normalizeStatus(state.data[panel.key][fi] && state.data[panel.key][fi][di]);
      if(status === 'no'){ totalPend++; perDayPend[col]++; }
      cells += `<td>${iconHTML(status, `data-panel="${panel.key}" data-farm="${fi}" data-day="${di}"`)}</td>`;
    });
    totalPendAll += totalPend;
    rows += `<tr><td class="farm-col">${escapeHtml(farm)}</td>${cells}<td class="rep-total pend-col">${totalPend}</td></tr>`;
  });
  const dayHeaders = visible.map(di=>`<th>${escapeHtml(state.days[di])}</th>`).join('');
  const dayFooterCells = perDayPend.map(n=>`<td>${n}</td>`).join('');
  return `
    <div class="rep-card-head">${panelIconHTML(panel.key)}<span>${escapeHtml(panel.title)}</span></div>
    <table class="rep-table">
      <thead><tr><th class="farm-col">Fazenda</th>${dayHeaders}<th class="pend-col">Pendências</th></tr></thead>
      <tbody>${rows}</tbody>
      <tfoot><tr class="pend-row">
        <td class="farm-col">PENDÊNCIAS:</td>${dayFooterCells}<td>${totalPendAll}</td>
      </tr></tfoot>
    </table>
  `;
}

function buildMonthlyCard(panel){
  let mPend = 0, mRows = '';
  state.farms.forEach((farm, fi)=>{
    const status = normalizeStatus(state.monthly[fi]);
    if(status === 'no') mPend++;
    mRows += `<tr><td class="farm-col">${escapeHtml(farm)}</td><td>${iconHTML(status, `data-monthly="${fi}"`)}</td></tr>`;
  });
  return `
    <div class="rep-card-head">${panelIconHTML(panel.key)}<span>${escapeHtml(panel.title)} — ${escapeHtml(state.monthLabel)}</span></div>
    <table class="rep-table">
      <thead><tr><th class="farm-col">Fazenda</th><th class="pend-col">Status</th></tr></thead>
      <tbody>${mRows}</tbody>
      <tfoot><tr class="pend-row">
        <td class="farm-col">PENDÊNCIAS:</td><td>${mPend}</td>
      </tr></tfoot>
    </table>
  `;
}

function buildDivergCard(panel){
  let rows = '';
  let somaNasc = 0, somaMortes = 0;
  state.farms.forEach((farm, fi)=>{
    const d = state.diverg[fi];
    const nd = d.nd === '' ? '' : Number(d.nd);
    const ns = d.ns === '' ? '' : Number(d.ns);
    const md = d.md === '' ? '' : Number(d.md);
    const ms = d.ms === '' ? '' : Number(d.ms);
    const difNasc = (nd==='' || ns==='') ? '' : (nd - ns);
    const difMortes = (md==='' || ms==='') ? '' : (md - ms);
    if(difNasc !== '') somaNasc += difNasc;
    if(difMortes !== '') somaMortes += difMortes;
    const difNascClass = difNasc==='' ? 'zero' : (difNasc===0 ? 'zero':'nonzero');
    const difMortesClass = difMortes==='' ? 'zero' : (difMortes===0 ? 'zero':'nonzero');
    rows += `<tr>
      <td class="farm-col">${escapeHtml(farm)}</td>
      <td><input class="num-input" type="text" inputmode="numeric" data-fi="${fi}" data-field="nd" value="${escapeHtml(d.nd)}"></td>
      <td><input class="num-input" type="text" inputmode="numeric" data-fi="${fi}" data-field="ns" value="${escapeHtml(d.ns)}"></td>
      <td class="diff-val ${difNascClass}">${difNasc}</td>
      <td><input class="num-input" type="text" inputmode="numeric" data-fi="${fi}" data-field="md" value="${escapeHtml(d.md)}"></td>
      <td><input class="num-input" type="text" inputmode="numeric" data-fi="${fi}" data-field="ms" value="${escapeHtml(d.ms)}"></td>
      <td class="diff-val ${difMortesClass}">${difMortes}</td>
    </tr>`;
  });
  return `
    <div class="rep-card-head">${panelIconHTML(panel.key)}<span>${escapeHtml(panel.title)}</span></div>
    <table class="rep-table">
      <thead><tr>
        <th class="farm-col">Fazenda</th><th>Nasc. Diário</th><th>Nasc. Sistema</th><th>Dif. Nasc.</th>
        <th>Mortes Diário</th><th>Mortes Sistema</th><th>Dif. Mortes</th>
      </tr></thead>
      <tbody>${rows}</tbody>
      <tfoot><tr class="pend-row pend-row-navy">
        <td class="farm-col">TOTAL:</td><td></td><td></td><td>${somaNasc}</td><td></td><td></td><td>${somaMortes}</td>
      </tr></tfoot>
    </table>
  `;
}
