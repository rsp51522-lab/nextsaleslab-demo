// Read cached results from the supplied fixed-format monthly confirmation workbook.
const MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export async function readConfirmation(file){
 if(!/\.xlsx$/i.test(file?.name||''))throw Error('確定ファイル（.xlsx）を選んでください。');
 const zip=await JSZip.loadAsync(await file.arrayBuffer());
 const parse=s=>new DOMParser().parseFromString(s,'application/xml');
 const book=parse(await zip.file('xl/workbook.xml').async('string'));
 const rel=parse(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
 const targets=Object.fromEntries([...rel.getElementsByTagName('Relationship')].map(e=>[e.getAttribute('Id'),e.getAttribute('Target')]));
 const strings=zip.file('xl/sharedStrings.xml')?[...parse(await zip.file('xl/sharedStrings.xml').async('string')).getElementsByTagNameNS(MAIN,'si')].map(x=>[...x.getElementsByTagNameNS(MAIN,'t')].map(t=>t.textContent).join('')):[];
 const sheets={};for(const sheet of book.getElementsByTagNameNS(MAIN,'sheet')){
  const id=sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
  const target=targets[id];if(!target)continue;
  const path=target.startsWith('/')?target.slice(1):'xl/'+target.replace(/^\.\//,'');
  const entry=zip.file(path);if(!entry)continue;
  const doc=parse(await entry.async('string')),cells={};
  for(const cell of doc.getElementsByTagNameNS(MAIN,'c')){
   const v=cell.getElementsByTagNameNS(MAIN,'v')[0],inline=cell.getElementsByTagNameNS(MAIN,'is')[0];
   if(v)cells[cell.getAttribute('r')]=cell.getAttribute('t')==='s'?strings[Number(v.textContent)]:v.textContent;
   else if(inline)cells[cell.getAttribute('r')]=inline.textContent;
  }
  sheets[sheet.getAttribute('name')]=cells;
 }
 if(!sheets['実績']||!sheets['栃木キャン']||!sheets['栃木PK'])throw Error('確定ファイルに必要な「実績」「栃木キャン」「栃木PK」がありません。');
 const s=sheets['実績'],c=sheets['栃木キャン'],p=sheets['栃木PK'];
 const n=(map,ref)=>Number(map[ref])||0;
 const amount={shared:n(s,'D30'),existing:n(s,'D31'),tenant:n(s,'D32'),op:n(s,'D33'),fixed:n(s,'D34'),cancel:n(s,'D35'),sales:n(s,'D13'),maintenance:n(s,'D14'),overall:n(s,'D9')};
 if(!amount.sales||Math.abs(amount.shared+amount.existing+amount.tenant+amount.op+amount.fixed-amount.cancel-amount.sales)>1||Math.abs(amount.sales+amount.maintenance-amount.overall)>1)throw Error('確定ファイルの営業・メンテ売上が内訳と一致しません。');
 if(Math.abs(n(c,'H66')-amount.cancel)>1)throw Error('実績と栃木キャンのキャンセル金額が一致しません。');
 const count={shared:n(s,'G30'),existing:n(s,'G31'),tenant:n(s,'G32'),op:n(s,'G33'),cancel:n(s,'G35')};
 const reps={};for(let i=0;i<4;i++){
  const name=p[`C${180+i}`];if(!name)continue;
  const crow=[61,62,63,64].find(r=>c[`B${r}`]===name);
  reps[name]={gross:n(p,`J${180+i}`),cancelSales:crow?n(c,`H${crow}`):0,cancelFresh:crow?n(c,`G${crow}`):0};
 }
 const repCancel=Object.values(reps).reduce((a,x)=>a+x.cancelSales,0);
 if(repCancel!==amount.cancel)throw Error('キャンセル担当者別の合計と確定売上が一致しません。');
 const fixedByFilename=file.name.match(/(?:^|[^0-9])(1[0-2]|[1-9])月/);
 return {amount,count,reps,sourceName:file.name,month:fixedByFilename?Number(fixedByFilename[1]):null,
  adjustment:amount.shared+amount.existing+amount.tenant+amount.op-Object.values(reps).reduce((a,x)=>a+x.gross,0)};
}
