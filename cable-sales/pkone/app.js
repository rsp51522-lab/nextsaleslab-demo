import {parseXls} from '../parser.js';
const $=s=>document.querySelector(s), yen=n=>new Intl.NumberFormat('ja-JP',{style:'currency',currency:'JPY',maximumFractionDigits:0}).format(n);
const esc=s=>String(s??'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;');
const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const xml=s=>new DOMParser().parseFromString(s,'application/xml');
const serialize=x=>new XMLSerializer().serializeToString(x);
const files={sales:null,confirmed:null,annual:null};
const month=$('#month'),year=$('#year');
for(const n of [6,7,8,10,11,12,1,2,3,4,5])month.add(new Option(`${n}月`,n));
for(let y=2026;y<=Math.max(2028,new Date().getFullYear()+1);y++)year.add(new Option(`${y}年度`,y));
let db,loaded=false,revision=0,storedSummaries={},detailAvailable=false;
function storageKey(kind,n=Number(month.value),y=Number(year.value)){return `${y}:${n}:${kind}`;}
function openStore(){return new Promise((resolve,reject)=>{const req=indexedDB.open('pkone-monthly-v1',1);req.onupgradeneeded=()=>req.result.createObjectStore('items');req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);});}
function readItem(key){return new Promise((resolve,reject)=>{const req=db.transaction('items','readonly').objectStore('items').get(key);req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);});}
function writeItem(key,value){return new Promise((resolve,reject)=>{const tx=db.transaction('items','readwrite');tx.objectStore('items').put(value,key);tx.oncomplete=()=>resolve();tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error||Error('保存が中断されました。'));});}
function deleteItem(key){return new Promise((resolve,reject)=>{const tx=db.transaction('items','readwrite');tx.objectStore('items').delete(key);tx.oncomplete=()=>resolve();tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error||Error('削除が中断されました。'));});}
function fileLabel(key){$(`#${key}-name`).textContent=files[key]?.name||'未選択';}
async function loadSelection(){
 const id=++revision;loaded=false;const n=Number(month.value),y=Number(year.value);
 const entries=await Promise.all(['sales','confirmed'].map(k=>readItem(storageKey(k,n,y))));
 const annual=await readItem(`${y}:annual`);if(id!==revision)return;
 [files.sales,files.confirmed]=entries;files.annual=annual||null;
 for(const k of Object.keys(files)){fileLabel(k);$(`#${k}`).value='';}
 storedSummaries={};for(const m of fiscalMonths){const d=await readItem(storageKey('summary',m,y));if(d)storedSummaries[m]=d;}
 if(id!==revision)return;loaded=true;render();
}
async function initialize(){try{db=await openStore();const selected=await readItem('selection');year.value=String(selected?.year||2026);month.value=String(selected?.month||9);await loadSelection();}catch(e){console.error(e);$('#status').textContent='ブラウザ内の保管領域を開けません。設定を確認してください。';$('#status').className='notice error';}}

function monthCol(n,base){return String.fromCharCode(base+[6,7,8,9,10,11,12,1,2,3,4,5].indexOf(Number(n)));}
function value(map,ref){return Number(map[ref])||0;}
function repName(name){const raw=String(name||'').replace(/\s+/g,'');return ['浅野','小林','西沢','遠藤'].find(n=>raw===n||raw.startsWith(n)&&/^[ァ-ヶーぁ-ゖ]+$/.test(raw.slice(n.length)))||raw;}
async function loadBook(file){
 const zip=await JSZip.loadAsync(await file.arrayBuffer());
 const book=xml(await zip.file('xl/workbook.xml').async('string'));
 const rel=xml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
 const relMap=Object.fromEntries([...rel.getElementsByTagName('Relationship')].map(x=>[x.getAttribute('Id'),x.getAttribute('Target')]));
 const shared=zip.file('xl/sharedStrings.xml')?xml(await zip.file('xl/sharedStrings.xml').async('string')):null;
 const strings=shared?[...shared.getElementsByTagNameNS(ns,'si')].map(x=>[...x.getElementsByTagNameNS(ns,'t')].map(t=>t.textContent).join('')):[];
 const sheets={};for(const s of book.getElementsByTagNameNS(ns,'sheet')){
  const target=relMap[s.getAttribute('r:id')||s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')];
  if(!target)continue;
  const path=target.startsWith('/')?target.slice(1):'xl/'+target.replace(/^\.\//,'');
  const entry=zip.file(path);if(!entry)continue;
  const doc=xml(await entry.async('string')), map={};
  for(const c of doc.getElementsByTagNameNS(ns,'c')){
   const v=c.getElementsByTagNameNS(ns,'v')[0], inline=c.getElementsByTagNameNS(ns,'is')[0];
   map[c.getAttribute('r')]=c.getAttribute('t')==='s'?strings[Number(v?.textContent)]:inline?inline.textContent:v?.textContent??'';
  }
  sheets[s.getAttribute('name')]={path,doc,map};
 }
 return {zip,book,sheets};
}
function salesData(s){
 const head=s.cells.get(2)||[];
 if(String(head[1]||'').trim()!=='契約日'||String(head[4]||'').trim()!=='担当'||String(head[5]||'').trim()!=='名前')throw Error('営業実績表の列配置が異なります。');
 const rows=[...s.cells.entries()].filter(([i,r])=>i>=5&&Number.isFinite(r[1])&&r[4]&&r[5]);
 const reps={};for(const [,r] of rows){const k=repName(r[4]),x=reps[k]??={fresh:0,cancel:0,op:0};x.fresh+=Number(r[64])||0;for(let col=46;col<=62;col++)x.op+=Number(r[col])||0;}
 for(const [i,r] of s.cancellations)if(i>=5&&Number.isFinite(r[1])&&r[2]&&r[3]){const k=repName(r[2]),x=reps[k]??={fresh:0,cancel:0};x.cancel+=Number(r[61])||0;}
 const serial=rows[0]?.[1]?.[1];if(!serial)throw Error('営業の契約行が見つかりません。');
 const date=new Date(Date.UTC(1899,11,30)+Math.round(serial)*86400000);
 return {reps,rows:rows.length,sourceMonth:date.getUTCMonth()+1};
}
function confirmedData(book){const s=book.sheets['実績']?.map;if(!s)throw Error('確定ファイルに「実績」シートがありません。');
 const g=(ref)=>value(s,ref);
 const sales={shared:g('D30'),existing:g('D31'),tenant:g('D32'),op:g('D33'),fixed:g('D34'),cancel:g('D35'),total:g('D13')};
 const m={tochigi:[g('E48'),g('E49'),g('E50'),g('E51')],koga:[g('F48'),g('F49'),g('F50'),g('F51')],total:g('D14')};
 const counts={result:g('G26'),gain:g('G28'),cancel:g('G35'),shared:g('G30'),existing:g('G31'),tenant:g('G32'),op:g('G33'),tochigi:[g('H48'),g('H49'),g('H50'),g('H51')],koga:[g('I48'),g('I49'),g('I50'),g('I51')]};
 if(!sales.total||!m.total||Math.abs(sales.shared+sales.existing+sales.tenant+sales.op+sales.fixed-sales.cancel-sales.total)>1||Math.abs([...m.tochigi,...m.koga].reduce((a,b)=>a+b,0)-m.total)>1)throw Error('確定ファイルの営業・メンテ内訳が合計と一致しません。');
 const cancelSheet=book.sheets['栃木キャン']?.map,pk=book.sheets['栃木PK']?.map;
 if(!cancelSheet||!pk||value(cancelSheet,'H66')!==sales.cancel)throw Error('確定ファイルの栃木キャンと実績の金額が一致しません。');
 const repCancel={};for(let r=61;r<=64;r++)if(cancelSheet[`B${r}`])repCancel[repName(cancelSheet[`B${r}`])]={fresh:value(cancelSheet,`G${r}`),sales:value(cancelSheet,`H${r}`)};
 if(Object.values(repCancel).reduce((a,x)=>a+x.sales,0)!==sales.cancel)throw Error('キャンセル担当別の合計が一致しません。');
 const repSales={};for(let r=180;r<=183;r++)if(pk[`C${r}`])repSales[repName(pk[`C${r}`])]=value(pk,`J${r}`);
 return {sales,m,counts,overall:g('D9'),repCancel,repSales};
}
function setCell(sheet,ref,n){
 const doc=sheet.doc,data=doc.getElementsByTagNameNS(ns,'sheetData')[0];if(!data)throw Error('Excelのシート構造を読み取れません。');
 const rowNum=Number(ref.match(/\d+/)[0]);let row=[...data.getElementsByTagNameNS(ns,'row')].find(x=>Number(x.getAttribute('r'))===rowNum);
 if(!row){row=doc.createElementNS(ns,'row');row.setAttribute('r',rowNum);const next=[...data.children].find(x=>x.localName==='row'&&Number(x.getAttribute('r'))>rowNum);data.insertBefore(row,next||null);}
 let cell=[...row.children].find(x=>x.localName==='c'&&x.getAttribute('r')===ref);
 if(!cell){cell=doc.createElementNS(ns,'c');cell.setAttribute('r',ref);const next=[...row.children].find(x=>x.localName==='c'&&columnIndex(x.getAttribute('r'))>columnIndex(ref));row.insertBefore(cell,next||null);}
 cell.removeAttribute('t');for(const child of [...cell.children])if(child.localName==='f'||child.localName==='v'||child.localName==='is')cell.removeChild(child);
 const v=doc.createElementNS(ns,'v');v.textContent=String(n);cell.appendChild(v);
}
function setFormulaCache(sheet,ref,n){const cell=[...sheet.doc.getElementsByTagNameNS(ns,'c')].find(x=>x.getAttribute('r')===ref);if(!cell||![...cell.children].some(x=>x.localName==='f'))return;let v=[...cell.children].find(x=>x.localName==='v');if(!v){v=sheet.doc.createElementNS(ns,'v');cell.appendChild(v);}v.textContent=String(n);}
function columnIndex(ref){return [...ref.match(/^[A-Z]+/)[0]].reduce((a,c)=>a*26+c.charCodeAt(0)-64,0);}
const fiscalMonths=[6,7,8,9,10,11,12,1,2,3,4,5];
function showView(){
 const view=location.hash==='#sales'?'sales-view':location.hash==='#maintenance'?'maintenance':'dashboard';
 for(const id of ['dashboard','sales-view','maintenance'])$(`#${id}`).hidden=id!==view||(id!=='dashboard'&&!detailAvailable);
 $('#empty-view').hidden=view==='dashboard'||detailAvailable;
 if(!$('#result').hidden)window.scrollTo({top:0,behavior:'auto'});
}
const monthlyTarget=4500000;
function ensureDashboardCharts(){
 const dashboard=$('#dashboard');
 if(!dashboard)throw Error('ダッシュボードの表示領域が見つかりません。ページを再読み込みしてください。');
 if(!$('#this-month-summary')){
  const summary=document.createElement('div');summary.id='this-month-summary';summary.className='this-month';
  summary.style.cssText='display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px;margin:16px 0';
  const style=document.createElement('style');style.textContent='.this-month>div{background:#fff;border:1px solid #c5d7e5;border-radius:12px;padding:14px}.this-month span{display:block;color:#58758e;font-size:13px}.this-month strong{display:block;color:#111;font-size:20px;margin-top:5px}';dashboard.append(style);
  (dashboard.querySelector('.progress-layout')||dashboard.firstElementChild).before(summary);
 }
 const oldDepartment=$('#department-chart');if(oldDepartment)oldDepartment.closest('.card')?.remove();
 const heading=$('#monthly-progress')?.closest('.card')?.querySelector('h3');if(heading)heading.textContent='全体売上・営業・メンテの月別推移';
 const missing=[['fresh-chart','営業の月別真水実績'],['pt-chart','メンテの月別PT']].filter(([id])=>!$(`#${id}`));
 if(missing.length){
  const stack=document.createElement('div');stack.className='chart-stack';stack.style.cssText='display:grid;gap:16px;margin:16px 0';
  for(const [id,title] of missing){const card=document.createElement('div');card.className='card';card.innerHTML=`<h3>${title}</h3><div id="${id}"></div>${id==='pt-chart'&&!$('#pt-note')?'<p id="pt-note" class="sub"></p>':''}`;stack.append(card);}
  (dashboard.querySelector('.progress-layout')||dashboard).after(stack);
 }
 if(!$('#pt-note')){const note=document.createElement('p');note.id='pt-note';note.className='sub';$('#pt-chart').after(note);}
}
function chartData(annual,d,n){
 const dept=annual.sheets['部署売上']?.map,maint=annual.sheets['メンテ']?.map;
 if(!dept||!maint)throw Error('年間ファイルに「部署売上」「メンテ」がありません。');
 return fiscalMonths.map(m=>{
  const col=monthCol(m,69),mc=monthCol(m,68),summary=storedSummaries[m];
  const total=value(dept,`${col}14`),sales=value(dept,`${col}15`),maintenance=value(dept,`${col}16`);
  const annualPt=value(maint,`${mc}14`),entered=total>0||sales>0||maintenance>0||annualPt>0||Boolean(summary);
  const selected=m===n&&Boolean(d);
  return {month:m,actual:entered?(selected?d.overall:summary?.overall??total):null,
   sales:entered?(selected?d.sales.total:summary?.sales??sales):null,
   maintenance:entered?(selected?d.m.total:summary?.maintenance??maintenance):null,
   fresh:entered?(selected?d.counts.result:summary?.fresh??value(dept,`${col}75`)):null,
   pt:entered?(annualPt||((selected?d.m.total:summary?.maintenance??maintenance)/4000)):null,
   estimatedPt:entered&&!annualPt&&Boolean(selected?d.m.total:summary?.maintenance??maintenance)};
 });
}
function barChart(data,series,unit){
 const width=900,height=290,left=55,right=24,top=28,bottom=42;
 const maxValue=Math.max(1,...data.flatMap(item=>series.map(s=>item[s.key]||0)));
 const max=Math.ceil(maxValue*1.15/(unit==='円'?1000000:20))*(unit==='円'?1000000:20);
 const x=i=>left+i*(width-left-right)/12+(width-left-right)/24;
 const y=v=>height-bottom-v/max*(height-top-bottom),group=(width-left-right)/12,bar=Math.min(23,(group-10)/series.length);
 const format=v=>unit==='円'?`${Math.round(v/10000)}万`:Number(v.toFixed(1)).toLocaleString('ja-JP');
 const guide=[0,max/2,max].map(v=>`<line x1="${left}" y1="${y(v)}" x2="${width-right}" y2="${y(v)}" stroke="#dce7ef"/><text x="${left-7}" y="${y(v)+4}" text-anchor="end" fill="#58758e" font-size="11">${format(v)}</text>`).join('');
 const bars=data.flatMap((item,i)=>series.map((s,j)=>{const v=item[s.key];if(v===null)return '';const bx=x(i)+(j-(series.length-1)/2)*bar,by=y(v),h=Math.max(0,height-bottom-by);return `<rect x="${bx-bar/2}" y="${by}" width="${bar-2}" height="${h}" rx="3" fill="${s.color}"><title>${item.month}月 ${s.name} ${unit==='円'?yen(v):format(v)+unit}</title></rect>`;})).join('');
 const labels=data.map((item,i)=>`<text x="${x(i)}" y="${height-12}" text-anchor="middle" fill="#42627d" font-size="13">${item.month}月</text>`).join('');
 const legend=series.map(s=>`<span><i style="background:${s.color}"></i>${s.name}</span>`).join('');
 return `<svg class="line-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="${series.map(s=>s.name).join('・')}の月別縦棒グラフ">${guide}${bars}${labels}</svg><div class="chart-legend">${legend}</div>`;
}
function renderProgress(annual,d,n){
 ensureDashboardCharts();
 const data=chartData(annual,d,n);
 const max=Math.ceil(Math.max(monthlyTarget,...data.map(v=>Math.max(v.actual||0,(v.sales||0)+(v.maintenance||0))))*1.15/1000000)*1000000;
 const left=52,right=24,top=22,bottom=42,width=700,height=310;
 const x=i=>left+i*(width-left-right)/11,y=v=>height-bottom-v/max*(height-top-bottom);
 const segments=[];let current=[];
 for(let i=0;i<data.length;i++){
  if(data[i].actual===null){if(current.length)segments.push(current);current=[];}
  else current.push(`${x(i)},${y(data[i].actual)}`);
 }if(current.length)segments.push(current);
 const axis=[0,monthlyTarget,max].filter((v,i,a)=>a.indexOf(v)===i).map(v=>`<line x1="${left}" y1="${y(v)}" x2="${width-right}" y2="${y(v)}" stroke="#dce7ef"/><text x="${left-7}" y="${y(v)+4}" text-anchor="end" fill="#58758e" font-size="11">${(v/10000).toLocaleString('ja-JP')}万</text>`).join('');
 const bars=data.map((v,i)=>{
  if(v.actual===null)return '';
  const sales=v.sales||0,maintenance=v.maintenance||0,bx=x(i)-14,base=y(0);
  return `${sales?`<rect x="${bx}" y="${y(sales)}" width="28" height="${base-y(sales)}" fill="#2876a7"><title>${v.month}月 営業 ${yen(sales)}</title></rect>`:''}${maintenance?`<rect x="${bx}" y="${y(sales+maintenance)}" width="28" height="${y(sales)-y(sales+maintenance)}" fill="#1e806d"><title>${v.month}月 メンテ ${yen(maintenance)}</title></rect>`:''}`;
 }).join('');
 $('#monthly-progress').innerHTML=`<svg class="line-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="全体売上の折れ線と、営業・メンテ売上の積み上げ棒グラフ。目標は毎月450万円。">
 ${axis}<line x1="${x(0)}" y1="${y(monthlyTarget)}" x2="${x(11)}" y2="${y(monthlyTarget)}" stroke="#8a9ba8" stroke-width="2" stroke-dasharray="7 5"/>
 ${bars}
 ${segments.map(points=>points.length>1?`<polyline points="${points.join(' ')}" fill="none" stroke="#d64051" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>`:'').join('')}
 ${data.map((v,i)=>v.actual===null?'':`<circle cx="${x(i)}" cy="${y(v.actual)}" r="5" fill="#d64051" stroke="#fff" stroke-width="1.5"><title>${v.month}月 全体売上 ${yen(v.actual)}</title></circle>`).join('')}
 ${data.map((v,i)=>`<text x="${x(i)}" y="${height-12}" text-anchor="middle" fill="#42627d" font-size="12">${v.month}月</text>`).join('')}</svg><div class="chart-legend"><span><i style="background:#d64051"></i>全体売上（折れ線）</span><span><i style="background:#2876a7"></i>営業</span><span><i style="background:#1e806d"></i>メンテ</span><span><i style="background:#8a9ba8"></i>目標 450万円</span></div>`;
 const periods=[['上期（6〜11月）',data.slice(0,6)],['下期（12〜5月）',data.slice(6)],['年間（6〜5月）',data]];
 $('#period-progress').innerHTML=periods.map(([label,items])=>{const actual=items.reduce((sum,v)=>sum+(v.actual||0),0),target=items.length*monthlyTarget,count=items.filter(v=>v.actual!==null).length;return `<div class="progress-card"><strong class="period-title">${label}</strong><div class="progress-metrics"><div class="progress-percent"><span>進捗率</span><strong>${Math.round(actual/target*100)}%</strong></div><div class="progress-facts"><div><span>稼働</span><strong>${count}/${items.length}か月</strong></div><div class="target"><span>目標</span><strong>${yen(target)}</strong></div><div><span>実績</span><strong>${yen(actual)}</strong></div></div></div></div>`;}).join('');
 $('#fresh-chart').innerHTML=barChart(data,[{key:'fresh',name:'戸建真水実績',color:'#2876a7'}],'件');
 $('#pt-chart').innerHTML=barChart(data,[{key:'pt',name:'メンテPT',color:'#1e806d'}],'PT');
 $('#pt-note').textContent=data.some(x=>x.estimatedPt)?'※ 年間ファイルにPT未入力の月は、確定したメンテ売上 ÷ 4,000円で換算しています。':'';
 const now=data.find(x=>x.month===n),diff=(now?.actual||0)-monthlyTarget;
 $('#this-month-summary').innerHTML=[['目標達成率',`${Math.round((now?.actual||0)/monthlyTarget*100)}%`],['目標との差額',`${diff>=0?'+':''}${yen(diff)}`],['営業真水実績',`${now?.fresh??0}件`],['メンテPT',`${Number((now?.pt??0).toFixed(1))} PT`]].map(([label,val])=>`<div><span>${n}月 ${label}</span><strong>${val}</strong></div>`).join('');
}
async function updateAnnual(book,d,s,n){
 const annual=book.sheets['部署売上'],sales=book.sheets['営業'];if(!annual||!sales||!book.sheets['メンテ'])throw Error('2026年ファイルに「部署売上」「営業」「メンテ」が揃っていません。');
 const c=monthCol(n,69),sc=monthCol(n,68),put=(r,v)=>setCell(annual,`${c}${r}`,v);
 for(const [r,k] of [[40,'shared'],[41,'existing'],[42,'tenant'],[43,'op'],[44,'fixed'],[45,'cancel']])put(r,d.sales[k]);
 [d.m.tochigi,d.m.koga].forEach((list,i)=>list.forEach((v,j)=>put((i?59:53)+j,v)));
 for(const [r,v] of [[75,d.counts.result],[76,d.counts.gain],[77,d.counts.cancel],[78,d.counts.shared+d.counts.existing+d.counts.tenant],[79,d.counts.shared],[80,d.counts.existing],[81,d.counts.tenant],[82,d.counts.op],...[85,86,87,88].map((r,i)=>[r,d.counts.tochigi[i]]),...[90,91,92,93].map((r,i)=>[r,d.counts.koga[i]])])put(r,v);
 for(const [r,v] of [[14,d.overall],[15,d.sales.total],[16,d.m.total],[20,d.overall],[37,d.sales.total],[47,d.m.total]])setFormulaCache(annual,`${c}${r}`,v);
 const reps=['浅野','小林','西沢','遠藤'];reps.forEach((name,i)=>{const x=s.reps[name]||{fresh:0,cancel:0};setCell(sales,`${sc}${13+i}`,x.fresh);setCell(sales,`${sc}${21+i}`,d.repCancel[name]?.fresh||0);});
 // Formula cells and unrelated months stay intact. Excel refreshes dependent cached values on open.
 let calc=book.book.getElementsByTagNameNS(ns,'calcPr')[0];if(!calc){calc=book.book.createElementNS(ns,'calcPr');book.book.documentElement.appendChild(calc);}calc.setAttribute('fullCalcOnLoad','1');calc.setAttribute('forceFullCalc','1');calc.setAttribute('calcMode','auto');
 book.zip.file('xl/workbook.xml',serialize(book.book));
 for(const sh of [annual,sales])book.zip.file(sh.path,serialize(sh.doc));
 return book.zip.generateAsync({type:'blob',compression:'DEFLATE'});
}
async function syncAnnual(){
 if(!files.annual)return 0;
 const y=Number(year.value),book=await loadBook(files.annual);let output,count=0;
 for(const m of fiscalMonths){
  const [sf,cf]=await Promise.all([readItem(storageKey('sales',m,y)),readItem(storageKey('confirmed',m,y))]);
  if(!sf||!cf)continue;
  const sales=salesData(parseXls(await sf.arrayBuffer())),confirmed=confirmedData(await loadBook(cf));
  if(sales.sourceMonth!==m)throw Error(`${m}月の営業実績表は${sales.sourceMonth}月のデータです。選択した月を確認してください。`);
  output=await updateAnnual(book,confirmed,sales,m);count++;
 }
 if(count){
  const updated=new File([output],`${y}年度_更新済.xlsx`,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  await writeItem(`${y}:annual`,updated);files.annual=updated;fileLabel('annual');
 }
 return count;
}
async function render(){if(!loaded)return;const id=++revision,status=$('#status');$('#save').disabled=true;$('#result').hidden=true;detailAvailable=false;status.className='notice';try{
 if(!files.annual){status.textContent='目次から年間ファイルを選んでください。選んだファイルはこのブラウザに保管されます。';showView();return;}
 status.textContent='読み込み中…';const n=Number(month.value),annual=await loadBook(files.annual);if(id!==revision)return;
 if(!files.sales||!files.confirmed){
  const current=chartData(annual,null,n).find(x=>x.month===n);
  $('#warning').hidden=true;
  $('#total').textContent=yen(current?.actual??0);$('#sales-total').textContent=yen(current?.sales??0);$('#maintenance-total').textContent=yen(current?.maintenance??0);
  renderProgress(annual,null,n);
  $('#result').hidden=false;status.textContent=`${files.annual.name}は保存済みです。年間ファイルの入力済み月を表示中です。`;status.className='notice ok';showView();$('#save').disabled=false;return;
 }
 const sales=salesData(parseXls(await files.sales.arrayBuffer()));
 const confirmed=await loadBook(files.confirmed);
 const d=confirmedData(confirmed);if(id!==revision)return;
 const summary={overall:d.overall,sales:d.sales.total,maintenance:d.m.total,fresh:d.counts.result};
 await writeItem(storageKey('summary'),summary);storedSummaries[n]=summary;
 const warnings=[];if(sales.sourceMonth!==n)warnings.push(`営業実績表は${sales.sourceMonth}月の契約日です。反映先は${n}月です。`);
 const unallocated=d.sales.shared+d.sales.existing+d.sales.tenant+d.sales.op-Object.values(d.repSales).reduce((a,b)=>a+b,0);
 if(unallocated)warnings.push(`担当者別元売上と営業区分売上の差額 ${yen(unallocated)} は担当者に配賦していません。固定費も担当者別には含めません。`);
 if(Math.abs(d.overall-d.sales.total-d.m.total)>1)warnings.push('確定ファイルの合計に営業・メンテ以外の売上が含まれます。');
 const existing=annual.sheets['部署売上'].map;if(existing[`${monthCol(n,69)}40`]||existing[`${monthCol(n,69)}53`])warnings.push('年間ファイルのこの月には既存値があります。ダウンロード時は保管した月の対象項目を置き換えます。');
 $('#warning').hidden=!warnings.length;$('#warning').textContent=warnings.join('　');
 $('#total').textContent=yen(d.overall);$('#sales-total').textContent=yen(d.sales.total);$('#maintenance-total').textContent=yen(d.m.total);
 const freshGain=d.counts.shared+d.counts.existing;
 $('#sales-kpi').textContent=yen(d.sales.total);$('#shared-kpi').textContent=`${freshGain-d.counts.cancel}件`;$('#existing-kpi').textContent=`${freshGain}件`;
 $('#maintenance-kpi').textContent=yen(d.m.total);$('#tochigi-kpi').textContent=yen(d.m.tochigi.reduce((a,b)=>a+b,0));$('#koga-kpi').textContent=yen(d.m.koga.reduce((a,b)=>a+b,0));
 renderProgress(annual,d,n);
 $('#sales-rows').innerHTML=[['シェアド',d.counts.shared,d.sales.shared],['既存',d.counts.existing,d.sales.existing],['店子',d.counts.tenant,d.sales.tenant],['OP',d.counts.op,d.sales.op],['固定費',0,d.sales.fixed],['キャンセル控除',d.counts.cancel,-d.sales.cancel]].map(([name,count,amount])=>`<tr${name==='キャンセル控除'?' class="cancel-row"':''}><th>${name}</th><td>${count}</td><td>${yen(amount)}</td></tr>`).join('');
 const oldMaintenance=$('#maintenance-rows');
 if(oldMaintenance)oldMaintenance.closest('.card').outerHTML='<div class="maintenance-details"><div class="card"><h3>栃木</h3><div class="scroll"><table><thead><tr><th>区分</th><th>件数・時間</th><th>確定売上</th></tr></thead><tbody id="tochigi-rows"></tbody></table></div></div><div class="card"><h3>古河</h3><div class="scroll"><table><thead><tr><th>区分</th><th>件数・時間</th><th>確定売上</th></tr></thead><tbody id="koga-rows"></tbody></table></div></div></div>';
 for(const [place,amounts,counts] of [['tochigi',d.m.tochigi,d.counts.tochigi],['koga',d.m.koga,d.counts.koga]])
  $(`#${place}-rows`).innerHTML=['訪問','時間','延長','OP'].map((label,i)=>`<tr><th>${label}</th><td>${counts[i]}</td><td>${yen(amounts[i])}</td></tr>`).join('');
 $('#reps').innerHTML=Object.entries(sales.reps).map(([name,x])=>`<tr><th>${esc(name)}</th><td>${d.repSales[name]===undefined?'要確認':yen(d.repSales[name]-(d.repCancel[name]?.sales||0))}</td><td>${x.fresh-(d.repCancel[name]?.fresh||0)}</td><td>${x.fresh}</td><td>${d.repCancel[name]?.fresh||0}</td><td>${x.op||0}</td></tr>`).join('');
 detailAvailable=true;$('#result').hidden=false;status.textContent=`${year.value}年度 ${n}月は保存済みです。確定ファイルの営業 ${yen(d.sales.total)}・メンテ ${yen(d.m.total)} を表示中です。`;status.className='notice ok';showView();$('#save').disabled=false;
 }catch(e){console.error(e);status.textContent=e.message||'ファイルを読み込めませんでした。';status.className='notice error';showView();}}
$('#save').onclick=async()=>{const status=$('#status');try{
 if(!files.annual)throw Error('先に年間ファイルを選んでください。');
 status.textContent='年間ファイルを保存中…';$('#save').disabled=true;
 const count=await syncAnnual(),output=files.annual;
 const url=URL.createObjectURL(output),a=document.createElement('a');a.href=url;a.download=output.name;a.click();setTimeout(()=>URL.revokeObjectURL(url),30000);
 status.textContent=count?`${count}か月分を反映した年間ファイルを保存し、ダウンロードしました。`:'保存済みの年間ファイルをダウンロードしました。';status.className='notice ok';
 }catch(e){status.textContent=e.message;status.className='notice error';}finally{$('#save').disabled=false;}};
for(const key of Object.keys(files))$(`#${key}`).addEventListener('change',async e=>{
 const picked=e.target.files[0];if(!picked||!db)return;
 const y=Number(year.value),n=Number(month.value);year.disabled=true;month.disabled=true;$('#save').disabled=true;
 const status=$('#status');status.textContent=`${picked.name}を確認して保存中…`;status.className='notice';
 try{
  if(key==='sales'){const parsed=salesData(parseXls(await picked.arrayBuffer()));if(parsed.sourceMonth!==n)throw Error(`この営業実績表は${parsed.sourceMonth}月です。画面右上で${parsed.sourceMonth}月を選んでください。`);}
  else if(key==='confirmed')confirmedData(await loadBook(picked));
  else {const book=await loadBook(picked);if(!book.sheets['部署売上']||!book.sheets['営業']||!book.sheets['メンテ'])throw Error('年間ファイルに「部署売上」「営業」「メンテ」がありません。');}
  await writeItem(key==='annual'?`${y}:annual`:storageKey(key,n,y),picked);
  if(key==='confirmed'){await deleteItem(storageKey('summary',n,y));delete storedSummaries[n];}
  files[key]=picked;fileLabel(key);
  let syncError;try{await syncAnnual();}catch(err){syncError=err;}
  await render();
  if(syncError){status.textContent=`ファイルは保存しましたが、年間ファイルへの反映に失敗しました: ${syncError.message}`;status.className='notice error';}
 }catch(err){status.textContent=`保存できませんでした: ${err.message}`;status.className='notice error';}
 finally{year.disabled=false;month.disabled=false;e.target.value='';}
});
async function selectPeriod(){if(!db)return;try{await writeItem('selection',{year:Number(year.value),month:Number(month.value)});await loadSelection();}catch(e){$('#status').textContent=e.message;$('#status').className='notice error';}}
month.addEventListener('change',selectPeriod);year.addEventListener('change',selectPeriod);
window.addEventListener('hashchange',showView);showView();initialize();
