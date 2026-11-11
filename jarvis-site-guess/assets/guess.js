/* JARVIS v3a — «Угадайка»: посетитель выбирает победителя матча КХЛ/НХЛ, модель — команду с большей вероятностью победы (w1/w2).
   Выборы — только в localStorage, ключ jv_guess_v1 = {ключ матча:{u:'h'|'a',m:'h'|'a',mp:0.55,t:'ГГГГ-ММ-ДД ЧЧ:ММ' (МСК),lg,h,a,hn,an}}.
   Итог — по счёту из assets/data/teams-lite.js (игры клуба хозяев). Без денег и кэфов. Страница guess.html. */
(()=>{
const KEY='jv_guess_v1',T=window.JV_TODAY||{},M=window.JV_MODEL||{},TEAMS=(window.JV_TEAMS||{}).teams||{};
const $=id=>document.getElementById(id);
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'})[c]);
const MG=['','января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря'];
const ms=t=>Date.parse(String(t).replace(' ','T')+':00+03:00');
const mskDay=x=>new Date(x+3*36e5).toISOString().slice(0,10);
const dayName=d=>{const n=Date.now();return d===mskDay(n)?'Сегодня':d===mskDay(n+864e5)?'Завтра':d===mskDay(n-864e5)?'Вчера':+d.slice(8,10)+' '+MG[+d.slice(5,7)]};
const LG={khl:'КХЛ',nhl:'НХЛ'};
let mem={};try{mem=JSON.parse(localStorage.getItem(KEY)||'{}')||{}}catch(e){}
let canSave=true;const save=()=>{try{localStorage.setItem(KEY,JSON.stringify(mem))}catch(e){canSave=false}};
const em=(id,nm)=>`<img class="gs-em" src="assets/em/${esc(id)}.webp" alt="" width="44" height="44" loading="lazy" onerror="this.style.visibility='hidden'">`;

/* итог сыгранного матча: игра хозяев дома против гостей в пределах 36 ч от начала */
function final(g){
  const tm=TEAMS[g.h],t0=ms(g.t);if(!tm||!tm.games)return null;
  for(const x of tm.games){if(x.o!==g.a||!x.h||x.gf==null||x.ga==null||x.gf===x.ga)continue;
    if(Math.abs(ms(x.d)-t0)<36*36e5)return{gf:x.gf,ga:x.ga,x:x.x||'',w:x.gf>x.ga?'h':'a'}}
  return null}

const open=(T.matches||[]).filter(m=>LG[m.lg]&&M[m.key]&&M[m.key].w1!=null&&M[m.key].w2!=null&&M[m.key].w1!==M[m.key].w2)
  .sort((a,b)=>ms(a.t)-ms(b.t));

function pick(m,side){
  if(ms(m.t)<=Date.now())return;const md=M[m.key];
  mem[m.key]={u:side,m:md.w1>md.w2?'h':'a',mp:Math.max(md.w1,md.w2),t:m.t,lg:m.lg,h:m.h,a:m.a,hn:m.hn,an:m.an};save();render()}

function card(m){
  const g=mem[m.key],live=ms(m.t)<=Date.now(),md=M[m.key],ms_=md.w1>md.w2?'h':'a',nm=s=>s==='h'?m.hn:m.an;
  const btn=s=>`<button type="button" class="gs-t${g&&g.u===s?' on':''}" data-k="${esc(m.key)}" data-s="${s}" aria-pressed="${g&&g.u===s?'true':'false'}"${live?' disabled':''}>${em(s==='h'?m.h:m.a)}<b>${esc(nm(s))}</b><small>${s==='h'?'хозяева':'гости'}</small></button>`;
  let foot;
  if(g){const same=g.u===g.m;foot=`<p class="gs-vs ${same?'same':'diff'}"><span>Модель: <b>${esc(nm(g.m))}</b> · ${Math.round(g.mp*100)} %</span><em>${same?'вы заодно с моделью':'вы против модели'}</em></p>`}
  else foot=`<p class="gs-vs wait"><span>${live?'Матч начался — выбор закрыт':'Выберите команду — покажем выбор модели'}</span></p>`;
  return `<article class="gs-card${g?' picked':''}${live?' live':''}"><header><b>${LG[m.lg]}</b><span>${esc(m.t.slice(11))} МСК${live?' · идёт или сыгран':''}</span></header><div class="gs-btns">${btn('h')}<i aria-hidden="true">—</i>${btn('a')}</div>${foot}</article>`}

const pl=n=>{const a=n%10,b=n%100;return a===1&&b!==11?'выбор ждёт итога':a>=2&&a<=4&&(b<12||b>14)?'выбора ждут итога':'выборов ждут итога'};
let shown=10,all=false;const phone=window.matchMedia&&matchMedia('(max-width:680px)').matches,LIM=6;
function render(){
  /* сыгранные */
  const rows=Object.keys(mem).map(k=>Object.assign({key:k},mem[k],{f:final(mem[k])})).sort((a,b)=>ms(b.t)-ms(a.t));
  const done=rows.filter(r=>r.f),wait=rows.filter(r=>!r.f&&ms(r.t)<=Date.now()),lost=r=>!r.f&&Date.now()-ms(r.t)>4*864e5;
  const u=done.filter(r=>r.u===r.f.w).length,mo=done.filter(r=>r.m===r.f.w).length,n=done.length;
  const pend=rows.filter(r=>!r.f&&!lost(r)).length;
  $('gs-kpi').innerHTML=`<li class="${n&&u>mo?'win':''}"><b>${u}<small>из ${n}</small></b><span>вы</span></li><li class="mdl"><b>${mo}<small>из ${n}</small></b><span>модель</span></li><li><b>${pend}</b><span>${pl(pend)}</span></li>`;
  $('gs-form').innerHTML=n?(u>mo?`Вы впереди на ${u-mo}.`:u<mo?`Модель впереди на ${mo-u}.`:'Счёт равный.')+` <span class="rs-dots" aria-label="последние матчи">${done.slice(0,12).reverse().map(r=>`<i class="${r.u===r.f.w?'':'l'}"></i>`).join('')}</span>`
    :(canSave?'Счёт появится после первого сыгранного матча.':'Браузер не даёт сохранять выборы — счёт не сохранится.');
  /* ближайшие */
  const box=$('gs-play');
  if(!open.length)box.innerHTML='<p class="rs-note">Матчей с расчётом сейчас нет. Загляните позже — данные обновляются несколько раз в день.</p>';
  else{const list=phone&&!all?open.slice(0,LIM):open,days=[];list.forEach(m=>{const d=mskDay(ms(m.t));let g=days.find(x=>x.d===d);if(!g)days.push(g={d,l:[]});g.l.push(m)});
    box.innerHTML=days.map(g=>`<h3 class="gs-day">${dayName(g.d)}<span>${+g.d.slice(8,10)} ${MG[+g.d.slice(5,7)]}</span></h3><div class="gs-grid">${g.l.map(card).join('')}</div>`).join('')+(list.length<open.length?`<button class="rs-more" id="gs-all" type="button">Показать ещё ${open.length-list.length}</button>`:'')}
  /* история */
  const hist=done.concat(wait).sort((a,b)=>ms(b.t)-ms(a.t)),sec=$('mine');sec.hidden=!hist.length;
  $('gs-done').innerHTML=hist.slice(0,shown).map(r=>{const f=r.f,nm=s=>s==='h'?r.hn:r.an,d=mskDay(ms(r.t));
    const mark=(s,who)=>f?`<span class="gs-m ${s===f.w?'w':'l'}"><i aria-hidden="true">${s===f.w?'✓':'✕'}</i>${who}: ${esc(nm(s))}</span>`:`<span class="gs-m">${who}: ${esc(nm(s))}</span>`;
    return `<div class="gs-row"><time>${+d.slice(8,10)} ${MG[+d.slice(5,7)]}<small>${LG[r.lg]||''}</small></time><p><b>${esc(r.hn)} — ${esc(r.an)}</b>${f?`<em>${f.gf}:${f.ga}${f.x?' '+esc(f.x):''}</em>`:`<em class="p">${lost(r)?'счёт не найден — в зачёт не идёт':'ждём счёт'}</em>`}</p><div>${mark(r.u,'Вы')}${mark(r.m,'Модель')}</div></div>`}).join('');
  $('gs-more').hidden=hist.length<=shown;
  $('gs-reset').parentNode.hidden=!rows.length}

document.addEventListener('click',e=>{
  const b=e.target.closest&&e.target.closest('.gs-t');
  if(b){const m=open.find(x=>x.key===b.dataset.k);if(m)pick(m,b.dataset.s);
    const nb=document.querySelector(`.gs-t[data-k="${CSS.escape(b.dataset.k)}"][data-s="${b.dataset.s}"]`);if(nb)nb.focus({preventScroll:true});return}
  if(e.target.id==='gs-more'){shown+=20;render()}
  if(e.target.id==='gs-all'){all=true;render()}
  if(e.target.id==='gs-reset'&&confirm('Удалить все ваши выборы и счёт в этом браузере?')){mem={};save();render()}});
render();
/* начало матча во время просмотра — закрыть выбор */
setInterval(()=>{if(open.some(m=>{const d=ms(m.t)-Date.now();return d<=0&&d>-65e3}))render()},6e4);
})();
