/* כרמל רהיטים — קטלוג */

const NAV = [
  { key:'all', label:'הכל' },
  { key:'bar', label:'כסאות בר' },
  { key:'chair', label:'כסאות אוכל' },
  { key:'gardenset', label:'פינות ישיבה ואוכל לחצר' },
  { key:'table_group', label:'שולחנות' },
  { key:'office', label:'כיסא משרדי' },
  { key:'other', label:'אביזרים ומטריות' },
];

const MATERIAL_LABELS = {
  iron:'ברזל', aluminum:'אלומיניום', upholstered:'מרופד', plastic:'פלסטיק',
  wood:'עץ', office:'משרדי', umbrella:'מטריה', garden:'ריהוט גן'
};

const CAT_LABELS = {
  chair:'כיסא אוכל', bar:'כיסא בר', office:'כיסא משרדי', table:'שולחן',
  gardenset:'פינת ישיבה/אוכל לחצר', umbrella:'מטריית שמש', accessory:'אביזר', set:'סט שולחן וכיסאות'
};

function matchesNav(p, navKey){
  if(navKey==='all') return true;
  if(navKey==='table_group') return p.category==='table' || p.category==='set';
  if(navKey==='other') return p.category==='umbrella' || p.category==='accessory';
  return p.category===navKey;
}

let state = { nav:'all', material:'all', q:'' };

function uniqueMaterials(navKey){
  const mats = new Set();
  PRODUCTS.forEach(p=>{
    if(matchesNav(p, navKey) && p.material && MATERIAL_LABELS[p.material] && p.category!=='office'){
      mats.add(p.material);
    }
  });
  return [...mats];
}

function render(){
  const grid = document.getElementById('grid');
  const q = state.q.trim();
  let list = PRODUCTS.filter(p=>{
    if(!matchesNav(p, state.nav)) return false;
    if(state.material!=='all' && p.material!==state.material) return false;
    if(q && !(p.name.includes(q) || (p.code||'').toLowerCase().includes(q.toLowerCase()))) return false;
    return true;
  });

  document.getElementById('resultCount').textContent = `${list.length} מוצרים`;

  if(list.length===0){
    grid.innerHTML = `<div class="empty-state" style="grid-column:1/-1"><h4>לא נמצאו מוצרים</h4><p>נסו לשנות את הסינון או את מילות החיפוש</p></div>`;
    return;
  }

  grid.innerHTML = list.map(p=>cardHTML(p)).join('');
  grid.querySelectorAll('.card').forEach(el=>{
    el.addEventListener('click', ()=> openModal(el.dataset.id));
  });
}

function cardHTML(p){
  const cols = p.colors.slice(0,6);
  const extra = p.colors.length>6 ? `<span class="swatch-more">+${p.colors.length-6}</span>` : '';
  return `
  <div class="card" data-id="${p.id}">
    <div class="thumb">
      <span class="badge-cat">${CAT_LABELS[p.category]||''}</span>
      <img src="${p.img}" alt="${p.name}" loading="lazy">
      <span class="price-tag">${p.price}&nbsp;₪</span>
    </div>
    <div class="body">
      <h4>${p.name}</h4>
      <div class="code">${p.code || ''}</div>
      <div class="dims">${p.dims || ''}</div>
      ${cols.length? `<div class="swatches">${cols.map(c=>`<span class="swatch" style="background:${c.hex}" title="${c.name}"></span>`).join('')}${extra}</div>` : ''}
    </div>
  </div>`;
}

function openModal(id){
  const p = PRODUCTS.find(x=>x.id===id);
  if(!p) return;
  document.getElementById('mImg').src = p.img;
  document.getElementById('mImg').alt = p.name;
  document.getElementById('mName').textContent = p.name;
  document.getElementById('mCode').textContent = [p.code, p.dims].filter(Boolean).join(' · ');
  document.getElementById('mPrice').textContent = p.price + ' ₪';
  document.getElementById('mNote').textContent = p.note || '';
  document.getElementById('mNoteRow').style.display = p.note ? 'block' : 'none';

  const swWrap = document.getElementById('mSwatches');
  if(p.colors.length){
    swWrap.innerHTML = p.colors.map(c=>`<span class="m-swatch"><span class="dot" style="background:${c.hex}"></span>${c.name}</span>`).join('');
    swWrap.parentElement.style.display = 'block';
  } else {
    swWrap.parentElement.style.display = 'none';
  }

  const waMsg = encodeURIComponent(`שלום, אשמח לפרטים על ${p.name} (${p.code || ''}) ממחירון כרמל רהיטים - ${p.price} ₪`);
  document.getElementById('mWhatsapp').href = `https://wa.me/972500000000?text=${waMsg}`;

  document.getElementById('modalBackdrop').classList.add('open');
}
function closeModal(){
  document.getElementById('modalBackdrop').classList.remove('open');
}

function buildNav(){
  const nav = document.getElementById('catNav');
  nav.innerHTML = NAV.map(n=>`<button data-key="${n.key}" class="${n.key===state.nav?'active':''}">${n.label}</button>`).join('');
  nav.querySelectorAll('button').forEach(b=>{
    b.addEventListener('click', ()=>{
      state.nav = b.dataset.key;
      state.material = 'all';
      buildNav();
      buildMaterialChips();
      updateSectionTitle();
      render();
      document.getElementById('gridTop').scrollIntoView({behavior:'smooth', block:'start'});
    });
  });
}

function buildMaterialChips(){
  const wrap = document.getElementById('materialChips');
  const mats = uniqueMaterials(state.nav);
  if(mats.length<2){ wrap.innerHTML=''; wrap.style.display='none'; return; }
  wrap.style.display='flex';
  wrap.innerHTML = `<span class="chip ${state.material==='all'?'active':''}" data-m="all">הכל</span>` +
    mats.map(m=>`<span class="chip ${state.material===m?'active':''}" data-m="${m}">${MATERIAL_LABELS[m]}</span>`).join('');
  wrap.querySelectorAll('.chip').forEach(c=>{
    c.addEventListener('click', ()=>{
      state.material = c.dataset.m;
      buildMaterialChips();
      render();
    });
  });
}

function updateSectionTitle(){
  const navItem = NAV.find(n=>n.key===state.nav);
  document.getElementById('sectionTitle').textContent = navItem ? navItem.label : '';
}

function buildFootNav(){
  const el = document.getElementById('footNav');
  if(!el) return;
  el.innerHTML = NAV.filter(n=>n.key!=='all').map(n=>`<li><a href="#gridTop" data-key="${n.key}">${n.label}</a></li>`).join('');
  el.querySelectorAll('a').forEach(a=>{
    a.addEventListener('click', ()=>{
      state.nav = a.dataset.key;
      state.material='all';
      buildNav(); buildMaterialChips(); updateSectionTitle(); render();
    });
  });
}

document.addEventListener('DOMContentLoaded', ()=>{
  buildNav();
  buildMaterialChips();
  updateSectionTitle();
  buildFootNav();
  render();

  document.getElementById('searchInput').addEventListener('input', (e)=>{
    state.q = e.target.value;
    render();
  });
  document.getElementById('modalBackdrop').addEventListener('click', (e)=>{
    if(e.target.id==='modalBackdrop') closeModal();
  });
  document.getElementById('closeModalBtn').addEventListener('click', closeModal);
  document.addEventListener('keydown', (e)=>{ if(e.key==='Escape') closeModal(); });
});
