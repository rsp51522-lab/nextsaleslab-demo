import {readSource,targets} from './source-reader.js';
const $=s=>document.querySelector(s);
let base=[];
try{const saved=JSON.parse(localStorage.getItem('cable-price-source')||'null');if(validSource(saved))base=saved;}catch{}
function validSource(rows){return Array.isArray(rows)&&rows.length>0&&rows.every(p=>p&&typeof p.col==='string'&&typeof p.name==='string'&&typeof p.category==='string'&&typeof p.group==='string'&&Number.isSafeInteger(p.price)&&p.price>=0)&&new Set(rows.map(p=>p.col)).size===rows.length;}
const yen=n=>new Intl.NumberFormat('ja-JP',{style:'currency',currency:'JPY',maximumFractionDigits:0}).format(n);
const esc=s=>String(s??'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;');
const price=p=>Number(localStorage.getItem('cable-price-'+p.col)??p.price);
function render(){if(!base.length){$('#rows').innerHTML='';$('#status').textContent='最初に「単価の元ファイル」を選択してください。読み込んだデータはこの端末に保存されます。';return;}const search=$('#search').value.trim().toLowerCase();const shown=base.filter(p=>`${p.col} ${p.category} ${p.group} ${p.name}`.toLowerCase().includes(search));
 $('#rows').innerHTML=shown.map(p=>`<tr class="${price(p)!==p.price?'changed':''}"><th>${esc(p.col)}</th><td>${esc(p.category)}・${esc(p.group)}</td><td>${esc(p.name)}</td><td>${yen(p.price)}</td><td><input type="number" min="0" step="1" value="${price(p)}" data-col="${esc(p.col)}" aria-label="${esc(p.name)}の単価"> 円</td></tr>`).join('');
 $('#status').textContent=`${shown.length}項目を表示中。変更した単価はこの端末に自動保存します。`;
}
$('#search').oninput=render;
const preview=document.createElement('section');preview.className='card import-preview';preview.hidden=true;preview.setAttribute('aria-label','単価の取り込み確認');$('.controls').after(preview);
let pending=[],sourceRevision=0;
function showPreview(rows,fileName){
 pending=rows;
 preview.innerHTML=`<h2>取り込み内容の確認</h2><p>${esc(fileName)}：${rows.length}行。単価が空欄の行と重複する行は選択していません。必要に応じて単価・元列を修正し、使用する行にチェックを入れてください。</p><p>取り込むと単価表を置き換え、取り込む商品の手動変更も元ファイルの単価に更新します。</p><div class="scroll"><table><thead><tr><th>使用</th><th>ファイルの行</th><th>種別・区分</th><th>商品</th><th>単価（円）</th><th>集計表の元列</th><th>確認事項</th></tr></thead><tbody>${rows.map((p,i)=>`<tr><td><input type="checkbox" data-use="${i}" aria-label="${esc(p.name)}を取り込む" ${p.selected?'checked':''}></td><th>${esc(p.source)}</th><td>${esc(p.category)}・${esc(p.group)}</td><td>${esc(p.name)}</td><td><input type="number" min="0" step="1" data-price="${i}" aria-label="${esc(p.name)}の取り込み単価" value="${p.price??''}"></td><td><select data-target="${i}" aria-label="${esc(p.name)}の元列"><option value="">元列を選択</option>${targets.map(t=>`<option value="${t.col}" ${t.col===p.col?'selected':''}>${t.col}：${esc(t.category)}・${esc(t.group)}・${esc(t.name)}</option>`).join('')}</select></td><td>${esc(p.issue||'')}</td></tr>`).join('')}</tbody></table></div><p id="preview-status" role="status"></p><button type="button" class="secondary" id="apply-source">選択した行を取り込む</button> <button type="button" class="secondary" id="cancel-source">キャンセル</button>`;
 preview.hidden=false;
 $('#preview-status').textContent=`${rows.filter(p=>p.selected).length}行を取り込みます。${rows.filter(p=>!p.selected).length}行は取り込みません。`;
 $('#apply-source').onclick=()=>{
  const rows=pending.filter(p=>p.selected).map(p=>({...p,index:targets.find(t=>t.col===p.col)?.index}));
  if(!validSource(rows)||rows.some(p=>!Number.isInteger(p.index))){$('#preview-status').textContent='使用する行を1つ以上選び、元列の重複・空欄、単価の空欄を解消してください。';return;}
  const excluded=pending.length-rows.length;
  try{localStorage.setItem('cable-price-source',JSON.stringify(rows));for(const p of rows)localStorage.removeItem('cable-price-'+p.col);base=rows;preview.hidden=true;pending=[];render();$('#status').textContent=`${fileName} から${rows.length}項目を保存しました。選択しなかった${excluded}行は取り込んでいません。`;}catch(err){$('#preview-status').textContent=`保存できませんでした: ${err.message}`;}
 };
 $('#cancel-source').onclick=()=>{preview.hidden=true;pending=[];render();};
 $('#status').textContent='取り込み内容を確認し、「選択した行を取り込む」を押してください。まだ保存していません。';
}
preview.onchange=e=>{
 const input=e.target;
 if(input.matches('[data-use]'))pending[Number(input.dataset.use)].selected=input.checked;
 if(input.matches('[data-price]'))pending[Number(input.dataset.price)].price=input.value.trim()===''?null:Number(input.value);
 if(input.matches('[data-target]'))pending[Number(input.dataset.target)].col=input.value;
 const selected=pending.filter(p=>p.selected);$('#preview-status').textContent=`${selected.length}行を取り込みます。${pending.length-selected.length}行は取り込みません。`;
};
$('#source').onchange=async e=>{const file=e.target.files?.[0];if(!file)return;const revision=++sourceRevision;preview.hidden=true;pending=[];$('#status').textContent='単価ファイルを読み込み中…';try{const result=await readSource(file);if(revision!==sourceRevision)return;if(result.json){const rows=result.json;if(!validSource(rows))throw Error('単価の元ファイルのJSON形式が異なります。');localStorage.setItem('cable-price-source',JSON.stringify(rows));base=rows;render();$('#status').textContent=`${rows.length}項目の単価を保存しました。`;}else showPreview(result.rows,file.name);}catch(err){if(revision===sourceRevision)$('#status').textContent=`読み込めませんでした: ${err.message}`;}finally{e.target.value='';}};
$('#rows').onchange=e=>{if(!e.target.matches('[data-col]'))return;const p=base.find(x=>x.col===e.target.dataset.col),value=Number(e.target.value);
 if(!p||!Number.isSafeInteger(value)||value<0){e.target.value=price(p);return;}localStorage.setItem('cable-price-'+p.col,String(value));render();};
$('#reset').onclick=()=>{if(!confirm('変更した商品単価をすべて初期値に戻しますか？'))return;for(const p of base)localStorage.removeItem('cable-price-'+p.col);render();};
$('#export').onclick=()=>{const data={format:'nsl-cable-price-master-v1',prices:Object.fromEntries(base.map(p=>[p.col,price(p)]))};const url=URL.createObjectURL(new Blob([JSON.stringify(data,null,2)],{type:'application/json'})),a=document.createElement('a');a.href=url;a.download='営業単価マスター.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),30000);};
$('#import').onchange=async e=>{try{const data=JSON.parse(await e.target.files[0].text());if(data.format!=='nsl-cable-price-master-v1'||!data.prices||!base.every(p=>Number.isSafeInteger(data.prices[p.col])&&data.prices[p.col]>=0))throw Error('マスターファイルの形式または単価が異なります。');for(const p of base)localStorage.setItem('cable-price-'+p.col,String(data.prices[p.col]));render();$('#status').textContent='マスターを読み込みました。';}catch(err){$('#status').textContent=err.message;}e.target.value='';};
render();

