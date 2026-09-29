// Reads the fixed Excel 97-2003 (.xls / BIFF8) report entirely in the browser.
const U16 = (d,p)=>d.getUint16(p,true), U32=(d,p)=>d.getUint32(p,true);
const cat = (parts)=>{ const n=parts.reduce((a,b)=>a+b.length,0), out=new Uint8Array(n);let p=0;for(const x of parts){out.set(x,p);p+=x.length;}return out;};
function oleStream(buffer){
 const bytes=new Uint8Array(buffer),d=new DataView(buffer);
 if(Array.from(bytes.slice(0,8)).join()!=='208,207,17,224,161,177,26,225') throw Error('Excel 97-2003形式（.xls）のファイルを選んでください。');
 const sectorSize=1<<U16(d,30), miniSize=1<<U16(d,32), off=s=>(s+1)*sectorSize;
 const fatIds=[];for(let i=0;i<109;i++){let v=U32(d,76+4*i);if(v<0xfffffffa)fatIds.push(v);}
 let dif=U32(d,68);for(let k=0;k<U32(d,72)&&dif<0xfffffffa;k++){let base=off(dif);for(let i=0;i<sectorSize/4-1;i++){let v=U32(d,base+4*i);if(v<0xfffffffa)fatIds.push(v);}dif=U32(d,base+sectorSize-4);}
 const fat=[];for(const id of fatIds.slice(0,U32(d,44)))for(let i=0;i<sectorSize/4;i++)fat.push(U32(d,off(id)+4*i));
 const chain=(start,count=10000)=>{let parts=[],seen=new Set(),cur=start;while(cur<0xfffffffa&&parts.length<count){if(seen.has(cur)||off(cur)+sectorSize>bytes.length)throw Error('ファイルの内部構造を読み取れません。');seen.add(cur);parts.push(bytes.slice(off(cur),off(cur)+sectorSize));cur=fat[cur];}return cat(parts);};
 const dir=chain(U32(d,48));let root=null,book=null;
 for(let p=0;p+128<=dir.length;p+=128){const dv=new DataView(dir.buffer,dir.byteOffset+p,128),len=U16(dv,64),name= len>=2 ? new TextDecoder('utf-16le').decode(dir.slice(p,p+len-2)):'';
  const ent={name,type:dir[p+66],start:U32(dv,116),size:U32(dv,120)};
  if(ent.type===5)root=ent;if(ent.type===2&&(name==='Workbook'||name==='Book'))book=ent;
 }
 if(!book||!root)throw Error('実績表のブックが見つかりません。');
 if(book.size>=U32(d,56))return chain(book.start).slice(0,book.size);
 const miniFat=chain(U32(d,60),U32(d,64)), mfd=new DataView(miniFat.buffer);
 const miniStream=chain(root.start).slice(0,root.size), parts=[],seen=new Set();let cur=book.start;
 while(cur<0xfffffffa){if(seen.has(cur))throw Error('ファイルを読み取れません。');seen.add(cur);parts.push(miniStream.slice(cur*miniSize,(cur+1)*miniSize));cur=U32(mfd,cur*4);}
 return cat(parts).slice(0,book.size);
}
function sstStrings(segments){
 let si=0,pos=8, strings=[];const dec=new TextDecoder('utf-16le');
 const go=()=>{while(si<segments.length&&pos>=segments[si].length){si++;pos=0;}if(si>=segments.length)throw Error('文字データの解析に失敗しました。');};
 const byte=()=>{go();return segments[si][pos++];};
 const word=()=>byte() | byte()<<8;
 const int=()=>word() | word()<<16;
 const unique=new DataView(segments[0].buffer,segments[0].byteOffset,segments[0].length).getUint32(4,true);
 for(let i=0;i<unique;i++){
  let chars=word(),opt=byte(),rich=(opt&8)?word():0,ext=(opt&4)?int():0,wide=!!(opt&1),text='';
  for(let c=0;c<chars;){
   if(pos>=segments[si].length){si++;pos=0;wide=!!(byte()&1);}
   if(wide){const available=Math.floor((segments[si].length-pos)/2),take=Math.min(chars-c,available);if(!take)throw Error('文字データを読み取れません。');text+=dec.decode(segments[si].slice(pos,pos+take*2));pos+=take*2;c+=take;}
   else{const take=Math.min(chars-c,segments[si].length-pos);if(!take)throw Error('文字データを読み取れません。');text+=String.fromCharCode(...segments[si].slice(pos,pos+take));pos+=take;c+=take;}
  }
  for(let k=0;k<rich*4+ext;k++)byte();strings.push(text);
 }
 return strings;
}
function rk(v){let n=(v&2)?v>>2:new DataView(new Uint32Array([0,v&0xfffffffc]).buffer).getFloat64(0,true);return (v&1)?n/100:n;}
export function parseXls(buffer){
 const b=oleStream(buffer),d=new DataView(b.buffer,b.byteOffset,b.byteLength),records=[];let p=0;
 while(p+4<=b.length){const type=U16(d,p),len=U16(d,p+2);if(p+4+len>b.length)break;records.push({type,start:p+4,len});p+=4+len;}
 let sst=[],sheets=[];
 for(let i=0;i<records.length;i++){
  const r=records[i],v=new DataView(b.buffer,b.byteOffset+r.start,r.len);
  if(r.type===0x85){const len=b[r.start+6],opt=b[r.start+7],name=(opt&1)?new TextDecoder('utf-16le').decode(b.slice(r.start+8,r.start+8+len*2)):new TextDecoder('latin1').decode(b.slice(r.start+8,r.start+8+len));sheets.push({offset:U32(v,0),name:name.trim()});}
  if(r.type===0xfc){const seg=[b.slice(r.start,r.start+r.len)];while(records[i+1]?.type===0x3c){i++;const x=records[i];seg.push(b.slice(x.start,x.start+x.len));}sst=sstStrings(seg);}
 }
 const sheet=sheets.find(x=>/月/.test(x.name)&&!x.name.includes('キャンセル'))||sheets[0];
 if(!sheet)throw Error('実績のシートが見つかりません。');
 const readSheet=(target)=>{if(!target)return new Map();const cells=new Map();let inSheet=false;
 const put=(row,col,value)=>{let a=cells.get(row);if(!a){a=[];cells.set(row,a);}a[col]=value;};
 for(const r of records){if(r.start-4===target.offset)inSheet=true;if(!inSheet)continue;
  if(r.type===0x0a)break;
  if(r.len<6)continue;
  const v=new DataView(b.buffer,b.byteOffset+r.start,r.len),row=U16(v,0),col=U16(v,2);
  if(r.type===0xfd)put(row,col,sst[U32(v,6)]??'');
  else if(r.type===0x203)put(row,col,v.getFloat64(6,true));
  else if(r.type===0x27e)put(row,col,rk(U32(v,6)));
  else if(r.type===0xbd){const end=U16(v,r.len-2);for(let c=col;c<=end;c++)put(row,c,rk(U32(v,4+(c-col)*6+2)));}
  else if(r.type===0x06 && r.len>=14){if(!(b[r.start+12]===255&&b[r.start+13]===255))put(row,col,v.getFloat64(6,true));}
  else if(r.type===0x204){const n=U16(v,6);put(row,col,new TextDecoder('latin1').decode(b.slice(r.start+8,r.start+8+n)));}
 }
 return cells;};
 return {name:sheet.name,cells:readSheet(sheet),cancellations:readSheet(sheets.find(x=>x.name.includes('キャンセル')))};
}
