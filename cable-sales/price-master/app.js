const $=s=>document.querySelector(s),base=await fetch('../prices.json').then(x=>x.json());
const yen=n=>new Intl.NumberFormat('ja-JP',{style:'currency',currency:'JPY',maximumFractionDigits:0}).format(n);
const esc=s=>String(s??'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;');
const price=p=>Number(localStorage.getItem('cable-price-'+p.col)??p.price);
function render(){const search=$('#search').value.trim().toLowerCase();const shown=base.filter(p=>`${p.col} ${p.category} ${p.group} ${p.name}`.toLowerCase().includes(search));
 $('#rows').innerHTML=shown.map(p=>`<tr class="${price(p)!==p.price?'changed':''}"><th>${esc(p.col)}</th><td>${esc(p.category)}・${esc(p.group)}</td><td>${esc(p.name)}</td><td>${yen(p.price)}</td><td><input type="number" min="0" step="1" value="${price(p)}" data-col="${esc(p.col)}" aria-label="${esc(p.name)}の単価"> 円</td></tr>`).join('');
 $('#status').textContent=`${shown.length}項目を表示中。変更した単価はこの端末に自動保存します。`;
}
$('#search').oninput=render;
$('#rows').onchange=e=>{if(!e.target.matches('[data-col]'))return;const p=base.find(x=>x.col===e.target.dataset.col),value=Number(e.target.value);
 if(!p||!Number.isSafeInteger(value)||value<0){e.target.value=price(p);return;}localStorage.setItem('cable-price-'+p.col,String(value));render();};
$('#reset').onclick=()=>{if(!confirm('変更した商品単価をすべて初期値に戻しますか？'))return;for(const p of base)localStorage.removeItem('cable-price-'+p.col);render();};
$('#export').onclick=()=>{const data={format:'nsl-cable-price-master-v1',prices:Object.fromEntries(base.map(p=>[p.col,price(p)]))};const url=URL.createObjectURL(new Blob([JSON.stringify(data,null,2)],{type:'application/json'})),a=document.createElement('a');a.href=url;a.download='営業単価マスター.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),30000);};
$('#import').onchange=async e=>{try{const data=JSON.parse(await e.target.files[0].text());if(data.format!=='nsl-cable-price-master-v1'||!data.prices||!base.every(p=>Number.isSafeInteger(data.prices[p.col])&&data.prices[p.col]>=0))throw Error('マスターファイルの形式または単価が異なります。');for(const p of base)localStorage.setItem('cable-price-'+p.col,String(data.prices[p.col]));render();$('#status').textContent='マスターを読み込みました。';}catch(err){$('#status').textContent=err.message;}e.target.value='';};
render();
