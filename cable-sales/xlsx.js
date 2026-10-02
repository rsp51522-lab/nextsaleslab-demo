const utf=s=>new TextEncoder().encode(s);
const esc=s=>String(s??'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;');
const col=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
function sheetXml(rows){
 const body=rows.map((row,i)=>`<row r="${i+1}">${row.map((v,j)=>v===null||v===undefined||v===''?'':Number.isFinite(v)&&typeof v==='number'?`<c r="${col(j+1)}${i+1}"><v>${v}</v></c>`:`<c r="${col(j+1)}${i+1}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`).join('')}</row>`).join('');
 return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews><sheetData>${body}</sheetData></worksheet>`;
}
const put16=(a,p,n)=>{a[p]=n&255;a[p+1]=n>>>8&255;},put32=(a,p,n)=>{put16(a,p,n);put16(a,p+2,n>>>16);};
function crc32(data){let c=0xffffffff;for(const b of data){c^=b;for(let i=0;i<8;i++)c=c&1?(c>>>1)^0xedb88320:c>>>1;}return(c^0xffffffff)>>>0;}
function zip(files){let locals=[],centrals=[],offset=0;
 for(const [name,content] of files){const n=utf(name),data=utf(content),crc=crc32(data),l=new Uint8Array(30+n.length+data.length);put32(l,0,0x04034b50);put16(l,4,20);put16(l,6,0x800);put16(l,8,0);put32(l,14,crc);put32(l,18,data.length);put32(l,22,data.length);put16(l,26,n.length);l.set(n,30);l.set(data,30+n.length);locals.push(l);
  const c=new Uint8Array(46+n.length);put32(c,0,0x02014b50);put16(c,4,20);put16(c,6,20);put16(c,8,0x800);put32(c,16,crc);put32(c,20,data.length);put32(c,24,data.length);put16(c,28,n.length);put32(c,42,offset);c.set(n,46);centrals.push(c);offset+=l.length;
 }
 const size=centrals.reduce((x,a)=>x+a.length,0),end=new Uint8Array(22);put32(end,0,0x06054b50);put16(end,8,files.length);put16(end,10,files.length);put32(end,12,size);put32(end,16,offset);
 return new Blob([...locals,...centrals,end],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
export function makeWorkbook({period,products,records,cancellations,byRep,byCategory,total,gross,cancelSales,fileName,confirmed}){
 const summary=[['ケーブルテレビ売上管理表',period],['確定・営業売上',total],['キャンセル控除前',gross],['キャンセル売上',-cancelSales],[],['営業・売上区分','戸建真水数／件数','確定売上'],...Object.entries(byCategory).map(([k,v])=>[k,v.count,v.sales]),[],['担当者','確定表の担当売上（キャン控除後）','実績','戸建真水数','確定キャンセル真水','OP件数'],...Object.entries(byRep).map(([k,v])=>[k,v.gross-v.cancelSales,v.fresh-v.cancelFresh,v.fresh,v.cancelFresh,v.op]),['担当者に未配賦の区分調整',confirmed.adjustment],['固定費',confirmed.amount.fixed],[],['営業実績表',fileName],['確定ファイル',confirmed.sourceName],['計算','実績 = 戸建真水数 - 確定キャンセル真水']];
 const item=[['商品別・単価マスターによる参考計算',period],['確定営業売上は売上サマリーを参照'],[],['元列','区分','種別','商品','件数','単価','参考売上'],...products.map(p=>[p.col,p.category,p.group,p.name,p.count,p.price,p.sales]),['参考合計','','','',products.reduce((a,p)=>a+p.count,0),'',products.reduce((a,p)=>a+p.sales,0)]];
 const detail=[['契約明細',period],[],['実績表の行','契約日','工事日','担当','お客様','エリア',...products.map(p=>p.col+' '+p.name),'参考商品売上'],...records.map(r=>[r.sourceRow,r.date,r.workDate,r.rep,r.customer,r.area,...r.items,r.sales])];
 const cancelDetail=[['キャンセル明細',period],['確定表の控除合計',-cancelSales],[],['確定表の担当','キャンセル真水','戻入売上'],...Object.entries(confirmed.reps).map(([name,x])=>[name,x.cancelFresh,x.cancelSales]),[],['営業実績表のキャンセル明細（照合用）'],['元行','担当','お客様','元表の真水','備考'],...cancellations.map(c=>[c.sourceRow,c.rep,c.customer,c.fresh,c.note])];
 const names=['売上サマリー','商品別売上','契約明細','キャンセル明細'],sheets=[summary,item,detail,cancelDetail];
 const content=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${names.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`;
 const workbook=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${names.map((n,i)=>`<sheet name="${n}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets></workbook>`;
 const files=[['[Content_Types].xml',content],['_rels/.rels',`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],['xl/workbook.xml',workbook],['xl/_rels/workbook.xml.rels',`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${names.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}</Relationships>`],...sheets.map((r,i)=>[`xl/worksheets/sheet${i+1}.xml`,sheetXml(r)])];
 return zip(files);
}
