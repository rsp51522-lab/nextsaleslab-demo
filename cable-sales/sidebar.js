const root=new URL('./',document.currentScript.src);
const links=[['ダッシュボード','pkone/#dashboard'],['営業','pkone/#sales'],['メンテ','pkone/#maintenance'],['単価マスター','price-master/'],['スタッフ申請','staff/']];
const path=location.pathname,hash=location.hash;
const active=path.includes('/staff/')?'スタッフ申請':path.includes('/price-master/')?'単価マスター':path.includes('/pkone/')?(hash==='#maintenance'?'メンテ':hash==='#sales'?'営業':'ダッシュボード'):'営業';
const menu=document.createElement('aside');menu.className='side-menu';menu.setAttribute('aria-label','目次');
const title=document.createElement('div');title.className='side-title';title.textContent='PKONE 目次';menu.appendChild(title);
for(const [label,target] of links){const a=document.createElement('a');a.href=new URL(target,root).href;a.textContent=label;if(label===active)a.setAttribute('aria-current','page');menu.appendChild(a);}
const note=document.createElement('small');note.textContent='月末実績管理';menu.appendChild(note);document.body.appendChild(menu);
const toggle=document.createElement('button');toggle.type='button';toggle.className='side-toggle';toggle.setAttribute('aria-label','目次を開く');toggle.textContent='☰ 目次';toggle.onclick=()=>{const open=menu.classList.toggle('open');toggle.setAttribute('aria-label',open?'目次を閉じる':'目次を開く');toggle.textContent=open?'× 閉じる':'☰ 目次';};document.body.appendChild(toggle);
window.addEventListener('hashchange',()=>{if(!location.pathname.includes('/pkone/'))return;const label=location.hash==='#maintenance'?'メンテ':location.hash==='#sales'?'営業':'ダッシュボード';for(const a of menu.querySelectorAll('a')){a.removeAttribute('aria-current');if(a.textContent===label)a.setAttribute('aria-current','page');}});
