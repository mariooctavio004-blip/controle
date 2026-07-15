/* ============================================================
   RENDER.JS
   Renderização geral: relatório (grid de cards) e painel de
   configuração (empresa, painéis selecionados, fazendas, dias).
============================================================ */

function renderAll(){ renderConfig(); renderReport(); renderAttachStatus(); syncDateInputs(); }

// Mantém os campos de data (topo) sincronizados com o state, sem atrapalhar
// enquanto a pessoa está digitando neles.
function syncDateInputs(){
  const startInp = document.getElementById('reportStartDate');
  const endInp = document.getElementById('reportEndDate');
  if(startInp && document.activeElement !== startInp) startInp.value = state.filterStart || '';
  if(endInp && document.activeElement !== endInp) endInp.value = state.filterEnd || '';
}

function renderConfig(){
  document.getElementById('companyName').value = state.company;
  document.getElementById('reportTitle').value = state.title;
  document.getElementById('periodLabel').value = state.period;
  document.getElementById('bannerText').value = state.banner;

  const list = document.getElementById('panelSelectList');
  list.innerHTML = '';
  PANEL_DEFS.forEach(p=>{
    const idx = state.selected.indexOf(p.key);
    const isChecked = idx !== -1;
    const item = document.createElement('label');
    item.className = 'panel-select-item' + (isChecked ? ' checked' : '');
    const slotName = idx===0?'Topo esquerda':idx===1?'Topo direita':idx===2?'Baixo esquerda':idx===3?'Baixo direita':'';
    item.innerHTML = `
      <span class="order-badge ${isChecked?'':'empty'}">${isChecked ? (idx+1) : ''}</span>
      <input type="checkbox" data-key="${p.key}" ${isChecked?'checked':''} style="width:16px;height:16px;">
      <span class="name">${p.title}</span>
      <span class="slot-hint">${slotName}</span>
    `;
    list.appendChild(item);
  });
  list.querySelectorAll('input[type=checkbox]').forEach(cb=>{
    cb.addEventListener('change', e=>{
      const key = e.target.dataset.key;
      const warning = document.getElementById('selectWarning');
      if(e.target.checked){
        if(state.selected.length >= 4){
          e.target.checked = false;
          warning.style.display = 'block';
          setTimeout(()=>warning.style.display='none', 2500);
          return;
        }
        state.selected.push(key);
      } else {
        state.selected = state.selected.filter(k=>k!==key);
      }
      renderAll();
      saveState();
    });
  });

  const farmList = document.getElementById('farmList');
  farmList.innerHTML = '';
  state.farms.forEach((f,i)=>{
    const chip = document.createElement('div');
    chip.className = 'chip';
    chip.innerHTML = `<input type="text" data-idx="${i}" value="${f.replace(/"/g,'&quot;')}"><button data-idx="${i}" title="Remover">×</button>`;
    farmList.appendChild(chip);
  });
  farmList.querySelectorAll('input').forEach(inp=>{
    inp.addEventListener('input', e=>{
      state.farms[+e.target.dataset.idx] = e.target.value;
      renderReport(); saveState();
    });
  });
  farmList.querySelectorAll('button').forEach(btn=>{
    btn.addEventListener('click', e=>{
      const idx = +e.target.dataset.idx;
      state.farms.splice(idx,1);
      PANEL_DEFS.filter(p=>p.type==='daily').forEach(p=>{ delete state.data[p.key][idx]; reindexObj(state.data[p.key], idx); });
      delete state.monthly[idx]; reindexObj(state.monthly, idx);
      delete state.diverg[idx]; reindexObj(state.diverg, idx);
      renderAll(); saveState();
    });
  });

  const dayList = document.getElementById('dayList');
  dayList.innerHTML = '';
  state.days.forEach((d,i)=>{
    const chip = document.createElement('div');
    chip.className = 'chip';
    chip.innerHTML = `<input type="text" data-idx="${i}" value="${d.replace(/"/g,'&quot;')}"><button data-idx="${i}" title="Remover">×</button>`;
    dayList.appendChild(chip);
  });
  dayList.querySelectorAll('input').forEach(inp=>{
    inp.addEventListener('input', e=>{
      state.days[+e.target.dataset.idx] = e.target.value;
      renderReport(); saveState();
    });
  });
  dayList.querySelectorAll('button').forEach(btn=>{
    btn.addEventListener('click', e=>{
      const idx = +e.target.dataset.idx;
      state.days.splice(idx,1);
      PANEL_DEFS.filter(p=>p.type==='daily').forEach(p=>{
        Object.keys(state.data[p.key]).forEach(fi=>{
          delete state.data[p.key][fi][idx];
          reindexObj(state.data[p.key][fi], idx);
        });
      });
      renderAll(); saveState();
    });
  });
}

function renderReport(){
  document.getElementById('titleDisplay').textContent = state.title;
  document.getElementById('periodDisplay').textContent = state.period;
  document.getElementById('bannerDisplay').textContent = state.banner;

  const grid = document.getElementById('reportGrid');
  grid.innerHTML = '';

  state.selected.forEach(key=>{
    const panel = PANEL_DEFS.find(p=>p.key===key);
    if(!panel) return;
    const card = document.createElement('div');
    card.className = 'rep-card';
    if(panel.type === 'daily') card.innerHTML = buildDailyCard(panel);
    else if(panel.type === 'monthly') card.innerHTML = buildMonthlyCard(panel);
    else if(panel.type === 'divergencias') card.innerHTML = buildDivergCard(panel);
    grid.appendChild(card);
  });

  grid.querySelectorAll('.rep-icon[data-panel]').forEach(el=>{
    el.addEventListener('click', ()=>{
      const panel = el.dataset.panel, fi = el.dataset.farm, di = el.dataset.day;
      const cur = normalizeStatus(state.data[panel][fi][di]);
      state.data[panel][fi][di] = STATUS_CYCLE[cur];
      renderReport(); saveState();
    });
  });
  grid.querySelectorAll('.rep-icon[data-monthly]').forEach(el=>{
    el.addEventListener('click', ()=>{
      const fi = el.dataset.monthly;
      const cur = normalizeStatus(state.monthly[fi]);
      state.monthly[fi] = STATUS_CYCLE[cur];
      renderReport(); saveState();
    });
  });
  grid.querySelectorAll('.num-input').forEach(inp=>{
    inp.addEventListener('input', e=>{
      const fi = e.target.dataset.fi, field = e.target.dataset.field;
      // permite apenas dígitos e um sinal de menos no início
      let clean = e.target.value.replace(/[^0-9-]/g,'');
      clean = clean.replace(/(?!^)-/g,'');
      state.diverg[fi][field] = clean;
      renderReport(); saveState();
      // restaura o foco e o cursor após o re-render
      const again = grid.querySelector(`.num-input[data-fi="${fi}"][data-field="${field}"]`);
      if(again){
        again.focus();
        try{
          const v = again.value;
          again.setSelectionRange(v.length, v.length);
        }catch(err){ /* alguns navegadores/tipos não suportam, ignora */ }
      }
    });
  });
}
