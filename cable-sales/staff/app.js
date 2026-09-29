const form=document.querySelector('#request'),status=document.querySelector('#status');
form.onsubmit=e=>{e.preventDefault();const data=Object.fromEntries(new FormData(form));data.createdAt=new Date().toISOString();const blob=new Blob([JSON.stringify(data,null,2)],{type:'application/json'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=`スタッフ申請_${data.name.replace(/[\\/<>:"|?*]/g,'')}.json`;a.click();setTimeout(()=>URL.revokeObjectURL(url),30000);status.textContent='申請書を保存しました。社内の担当者へ渡してください。';};
document.querySelector('#print').onclick=()=>window.print();
