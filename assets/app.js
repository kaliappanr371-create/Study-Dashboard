/* ===================== ZIP + XLSX (read & write, no libraries) ===================== */
const te=new TextEncoder(), td=new TextDecoder();
async function inflateRaw(u8){
  if(typeof DecompressionStream==='undefined')throw new Error('This browser cannot unzip .xlsx files (needs a recent Chrome, Edge, Firefox or Safari 16.4+).');
  const ds=new DecompressionStream('deflate-raw');const w=ds.writable.getWriter();w.write(u8);w.close();
  return new Uint8Array(await new Response(ds.readable).arrayBuffer());
}
async function unzip(buf){
  const u8=new Uint8Array(buf),dv=new DataView(buf);
  let e=u8.length-22;while(e>=0&&dv.getUint32(e,true)!==0x06054b50)e--;
  if(e<0)throw new Error('This is not a valid .xlsx file (not a zip archive). Old .xls files must be re-saved as .xlsx.');
  const n=dv.getUint16(e+10,true);let p=dv.getUint32(e+16,true);const files={};
  for(let i=0;i<n;i++){
    if(dv.getUint32(p,true)!==0x02014b50)break;
    const method=dv.getUint16(p+10,true),csize=dv.getUint32(p+20,true),nl=dv.getUint16(p+28,true),el=dv.getUint16(p+30,true),cl=dv.getUint16(p+32,true),off=dv.getUint32(p+42,true);
    files[td.decode(u8.subarray(p+46,p+46+nl))]={method,csize,off};p+=46+nl+el+cl;
  }
  const get=async name=>{const f=files[name];if(!f)return null;const l=dv.getUint16(f.off+26,true)+dv.getUint16(f.off+28,true);const d=u8.subarray(f.off+30+l,f.off+30+l+f.csize);return f.method===0?d:await inflateRaw(d)};
  return{names:Object.keys(files),text:async n=>{const d=await get(n);return d?td.decode(d):null}};
}
function xml(s){return new DOMParser().parseFromString(s,'application/xml')}
function colIdx(ref){let c=0;const m=ref.match(/^[A-Z]+/i)[0].toUpperCase();for(const ch of m)c=c*26+ch.charCodeAt(0)-64;return c-1}
function colName(i){let s='';i++;while(i>0){const m=(i-1)%26;s=String.fromCharCode(65+m)+s;i=Math.floor((i-1)/26)}return s}
const BUILTIN_DATE={14:'d',15:'d',16:'d',17:'d',18:'t',19:'t',20:'t',21:'t',22:'dt',45:'t',46:'t',47:'t'};
function fmtKind(code){
  const c=code.replace(/"[^"]*"/g,'').replace(/\\./g,'').replace(/\[(?!h\]|m\]|s\])[^\]]*\]/gi,'');
  const t=/[hs]/i.test(c)||/\[h\]|\[m\]/i.test(c),d=/[dy]/i.test(c);
  return t&&d?'dt':t?'t':d?'d':null;
}
function isoFromSerial(v){const d=new Date(Math.round((Math.floor(v)-25569)*864e5));return d.getUTCFullYear()+'-'+String(d.getUTCMonth()+1).padStart(2,'0')+'-'+String(d.getUTCDate()).padStart(2,'0')}
/* Returns {sheets:[{name,rows:[ [cell|undefined,...] ]}]}.  cell = {v, k:'s'|'n'|'d'|'t'|'b'|'e'}  (d = date serial, t = time/duration as fraction of a day) */
async function readXlsx(buf){
  const z=await unzip(buf);
  const wb=await z.text('xl/workbook.xml');if(!wb)throw new Error('No workbook found inside this file.');
  const wbd=xml(wb),rels=xml(await z.text('xl/_rels/workbook.xml.rels')||'<x/>');
  const relMap={};for(const r of rels.getElementsByTagName('Relationship'))relMap[r.getAttribute('Id')]=r.getAttribute('Target');
  const ss=[];const sst=await z.text('xl/sharedStrings.xml');
  if(sst){for(const si of xml(sst).getElementsByTagName('si')){let t='';for(const x of si.getElementsByTagName('t'))t+=x.textContent;ss.push(t)}}
  const xfKind=[];const st=await z.text('xl/styles.xml');
  if(st){const sd=xml(st),custom={};for(const nf of sd.getElementsByTagName('numFmt'))custom[nf.getAttribute('numFmtId')]=nf.getAttribute('formatCode');
    const cx=sd.getElementsByTagName('cellXfs')[0];if(cx)for(const xf of cx.getElementsByTagName('xf')){const id=xf.getAttribute('numFmtId');xfKind.push(custom[id]!==undefined?fmtKind(custom[id]):(BUILTIN_DATE[id]||null))}}
  const sheets=[];
  for(const s of wbd.getElementsByTagName('sheet')){
    const rid=s.getAttribute('r:id')||s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
    let target=relMap[rid]||'';target=target.replace(/^\//,'');if(!target.startsWith('xl/'))target='xl/'+target;
    const sx=await z.text(target);if(!sx)continue;const rows=[];
    for(const row of xml(sx).getElementsByTagName('row')){
      const rn=parseInt(row.getAttribute('r'))-1,arr=[];let any=false;
      for(const c of row.getElementsByTagName('c')){
        const ref=c.getAttribute('r');const ci=colIdx(ref);const t=c.getAttribute('t'),sidx=c.getAttribute('s');
        const ve=c.getElementsByTagName('v')[0];let cell=null;
        if(t==='inlineStr'){let x='';for(const e of c.getElementsByTagName('t'))x+=e.textContent;cell={v:x,k:'s'}}
        else if(ve){const raw=ve.textContent;
          if(t==='s')cell={v:ss[+raw]??'',k:'s'};else if(t==='str')cell={v:raw,k:'s'};else if(t==='b')cell={v:raw==='1',k:'b'};else if(t==='e')cell={v:raw,k:'e'};
          else if(t==='d')cell={v:raw.slice(0,10),k:'s'};
          else{const n=parseFloat(raw);const kd=sidx!=null?xfKind[+sidx]:null;cell={v:n,k:kd==='d'||kd==='dt'?'d':kd==='t'?'t':'n'}}}
        if(cell&&!(cell.k==='s'&&cell.v==='')){arr[ci]=cell;any=true}
      }
      if(any)rows[rn]=arr;
    }
    sheets.push({name:s.getAttribute('name'),rows});
  }
  return{sheets};
}
/* minimal CSV reader -> same shape */
function readCsv(text){
  const rows=[];let row=[],f='',q=false;text=text.replace(/^﻿/,'');
  for(let i=0;i<text.length;i++){const ch=text[i];
    if(q){if(ch==='"'){if(text[i+1]==='"'){f+='"';i++}else q=false}else f+=ch}
    else if(ch==='"')q=true;else if(ch===','){row.push(f);f=''}else if(ch==='\n'||ch==='\r'){if(ch==='\r'&&text[i+1]==='\n')i++;row.push(f);f='';rows.push(row);row=[]}else f+=ch}
  if(f!==''||row.length){row.push(f);rows.push(row)}
  const out=[];rows.forEach((r,i)=>{const a=[];let any=false;r.forEach((v,j)=>{v=v.trim();if(v!==''){a[j]={v,k:'s'};any=true}});if(any)out[i]=a});
  return{sheets:[{name:'CSV',rows:out}]};
}
async function readAnyFile(file){
  const buf=await file.arrayBuffer();
  if(/\.csv$/i.test(file.name))return readCsv(td.decode(buf));
  if(/\.xls$/i.test(file.name))throw new Error('Old .xls format is not supported. Open it in Excel / Google Sheets and save as .xlsx.');
  return readXlsx(buf);
}

/* ---------- writer ---------- */
function crc32(u8){if(!crc32.t){crc32.t=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;crc32.t[n]=c>>>0}}let c=0xFFFFFFFF;for(let i=0;i<u8.length;i++)c=crc32.t[(c^u8[i])&255]^(c>>>8);return(c^0xFFFFFFFF)>>>0}
function zipStore(files){
  const parts=[],cd=[];let off=0;
  for(const f of files){const nm=te.encode(f.name),crc=crc32(f.data),sz=f.data.length;
    const lh=new DataView(new ArrayBuffer(30));lh.setUint32(0,0x04034b50,true);lh.setUint16(4,20,true);lh.setUint16(6,0x0800,true);lh.setUint16(12,0x21,true);lh.setUint32(14,crc,true);lh.setUint32(18,sz,true);lh.setUint32(22,sz,true);lh.setUint16(26,nm.length,true);
    parts.push(new Uint8Array(lh.buffer),nm,f.data);
    const ch=new DataView(new ArrayBuffer(46));ch.setUint32(0,0x02014b50,true);ch.setUint16(4,20,true);ch.setUint16(6,20,true);ch.setUint16(8,0x0800,true);ch.setUint16(14,0x21,true);ch.setUint32(16,crc,true);ch.setUint32(20,sz,true);ch.setUint32(24,sz,true);ch.setUint16(28,nm.length,true);ch.setUint32(42,off,true);
    cd.push(new Uint8Array(ch.buffer),nm);off+=30+nm.length+sz}
  const cds=cd.reduce((a,x)=>a+x.length,0);const e=new DataView(new ArrayBuffer(22));
  e.setUint32(0,0x06054b50,true);e.setUint16(8,files.length,true);e.setUint16(10,files.length,true);e.setUint32(12,cds,true);e.setUint32(16,off,true);
  return new Blob([...parts,...cd,new Uint8Array(e.buffer)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
const xe=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'');
/* sheet = {name, widths:[], rows:[[cell]]}; cell = string | number | {dur:minutes} | {date:'YYYY-MM-DD'} | {h:'header text'} | null */
function buildXlsx(sheets){
  const enc=s=>te.encode(s);const files=[];
  files.push({name:'[Content_Types].xml',data:enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'+sheets.map((s,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')+'</Types>')});
  files.push({name:'_rels/.rels',data:enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')});
  const used=new Set();
  files.push({name:'xl/workbook.xml',data:enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+sheets.map((s,i)=>{let n=s.name.replace(/[\[\]:*?\/\\]/g,'-').slice(0,31)||'Sheet';let b=n,k=2;while(used.has(n.toLowerCase())){n=b.slice(0,28)+' '+k++}used.add(n.toLowerCase());return`<sheet name="${xe(n)}" sheetId="${i+1}" r:id="rId${i+1}"/>`}).join('')+'</sheets></workbook>')});
  files.push({name:'xl/_rels/workbook.xml.rels',data:enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+sheets.map((s,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')+`<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`)});
  files.push({name:'xl/styles.xml',data:enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="[h]:mm:ss"/><numFmt numFmtId="165" formatCode="yyyy-mm-dd"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8ECFF"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>')});
  sheets.forEach((s,i)=>{
    let x='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
    if(s.widths)x+='<cols>'+s.widths.map((w,j)=>`<col min="${j+1}" max="${j+1}" width="${w}" customWidth="1"/>`).join('')+'</cols>';
    x+='<sheetData>';
    s.rows.forEach((r,ri)=>{x+=`<row r="${ri+1}">`;
      (r||[]).forEach((c,ci)=>{if(c===null||c===undefined||c==='')return;const ref=colName(ci)+(ri+1);
        if(typeof c==='number')x+=`<c r="${ref}"><v>${c}</v></c>`;
        else if(c.dur!==undefined)x+=`<c r="${ref}" s="2"><v>${c.dur/1440}</v></c>`;
        else if(c.date){const d=Date.UTC(+c.date.slice(0,4),+c.date.slice(5,7)-1,+c.date.slice(8,10))/864e5+25569;x+=`<c r="${ref}" s="3"><v>${d}</v></c>`}
        else if(c.h!==undefined)x+=`<c r="${ref}" s="1" t="inlineStr"><is><t>${xe(c.h)}</t></is></c>`;
        else x+=`<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xe(c)}</t></is></c>`});
      x+='</row>'});
    x+='</sheetData></worksheet>';files.push({name:`xl/worksheets/sheet${i+1}.xml`,data:enc(x)});
  });
  return zipStore(files);
}

/* ===================== SHEET UNDERSTANDING (pure functions, no app state) ===================== */
const WEEKDAYS=/^(mon|tue|wed|thu|fri|sat|sun)[a-z]*$/i;
const cellText=c=>{if(!c)return'';if(c.k==='d')return isoFromSerial(c.v);if(c.k==='t')return fmtMin(c.v*1440);if(c.k==='b')return c.v?'TRUE':'FALSE';return String(c.v).replace(/\s+/g,' ').trim()};
function fmtMin(m){m=Math.round(m);return Math.floor(m/60)+':'+String(m%60).padStart(2,'0')}
function parseDurText(t){
  t=String(t).trim().toLowerCase();if(!t||t==='-')return null;let m;
  if(m=t.match(/^(\d+):(\d{1,2})(?::(\d{1,2}))?$/))return +m[1]*60+ +m[2]+(+m[3]||0)/60;
  if(m=t.match(/^(?:(\d+(?:\.\d+)?)\s*h(?:ours?|rs?)?)?\s*(?:(\d+)\s*m(?:in(?:ute)?s?)?)?$/)&&(m[1]||m[2]))return (+m[1]||0)*60+(+m[2]||0);
  if(m=t.match(/^\d+(\.\d+)?$/))return parseFloat(t)*60;
  return null;
}
function durOf(c,hdr){
  if(!c)return null;
  if(c.k==='t')return Math.round(c.v*86400)/60;
  if(c.k==='n'){if(/min/i.test(hdr))return c.v;if(c.v>0&&c.v<1)return Math.round(c.v*86400)/60;return c.v*60}
  if(c.k==='s')return parseDurText(c.v);
  return null;
}
function dateOf(c){
  if(!c)return null;if(c.k==='d')return isoFromSerial(c.v);
  if(c.k==='s'){let m=c.v.trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);if(m)return m[1]+'-'+m[2].padStart(2,'0')+'-'+m[3].padStart(2,'0');
    m=c.v.trim().match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);if(m)return m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0')}
  return null;
}
function roleOf(h){
  const t=h.toLowerCase().trim();if(!t)return null;
  if(/download|offline/.test(t))return'dl';
  if(/^(status|watched|completion|state|done\??|completed\??)$|status/.test(t))return'status';
  if(/duration|length|runtime|run time|^hours?$|^hrs?$|^time$/.test(t))return'dur';
  if(/priority|importance/.test(t))return'pri';
  if(/note|remark|comment/.test(t))return'note';
  if(/^(date|scheduled|schedule|planned|due|watch date|date watched)/.test(t))return'date';
  if(/^(topic|topics|chapter|chapters|title|particulars?|subject|unit)$/.test(t))return'topic';
  if(/^(day|days|lecture|lectures|lec|class|#|no\.?|sr\.? ?no\.?|s\.? ?no\.?|sl\.? ?no\.?)$/.test(t))return'lab';
  return'x';
}
/* find the header row: the row with the most recognised headers */
function findHeader(rows){
  let best=-1,bs=0;
  for(let r=0;r<Math.min(rows.length,40);r++){const row=rows[r];if(!row)continue;
    const roles=new Set();row.forEach(c=>{if(c&&c.k==='s'){const ro=roleOf(c.v);if(ro&&ro!=='x')roles.add(ro)}});
    const sc=roles.size+(roles.has('dur')?1:0)+(roles.has('topic')||roles.has('lab')?1:0);
    if((roles.has('dur')||roles.has('status')||roles.has('topic'))&&roles.size>=2&&sc>bs){bs=sc;best=r}}
  return best;
}
function classifySheet(sh){
  const rows=sh.rows,h=findHeader(rows);
  if(h>=0){const heads=rows[h].map(c=>c&&c.k==='s'?c.v.toLowerCase():'');
    const hasDur=heads.some(t=>roleOf(t)==='dur'),hasDone=heads.some(t=>/^done\??$|^completed?\??$/.test(t)),hasChange=heads.some(t=>/change|item|task|point|description/.test(t));
    if(!hasDur&&hasDone&&hasChange)return'checklist';
    if(hasDur||heads.some(t=>roleOf(t)==='status'))return'lectures';
  }
  return'info';
}
const normTopic=t=>{
  let s=String(t||'').toLowerCase().replace(/^\d+_\d+[a-z]?_+/,'').replace(/_+/g,' ');
  const had=/^part\s*[a-z0-9]+\s*[-–:]\s*/.test(s);
  s=s.replace(/^part\s*[a-z0-9]+\s*[-–:]\s*/,'').replace(/\s*[-–]\s*(part\s*[a-z0-9]+|\d+|raw)\s*$/,'').replace(/\s*part\s*-?\s*[a-z0-9]+\s*$/,'').replace(/\s+raw\s*$/,'');
  if(had)s=s.replace(/\s*-?\s*\d+$/,'');
  return s.replace(/\s+/g,' ').trim();
};
function parseLectureSheet(sh){
  const rows=sh.rows,hr=findHeader(rows),H=rows[hr];
  const cols=[];H.forEach((c,i)=>{if(c&&c.k==='s')cols.push({i,text:c.v.trim(),role:roleOf(c.v)})});
  const byRole=r=>cols.filter(c=>c.role===r);
  const col={dur:byRole('dur')[0],status:byRole('status')[0],dl:byRole('dl')[0],pri:byRole('pri')[0],note:byRole('note')[0],date:byRole('date')[0],topic:byRole('topic')[0]};
  let labs=byRole('lab');
  const lastRow=rows.length-1;
  const textiness=c=>{let n=0,t=0;for(let r=hr+1;r<=lastRow;r++){const x=rows[r]&&rows[r][c.i];if(x&&x.k==='s'){t++;if(x.v.length>14||/\s\S+\s/.test(x.v))n++}}return t?n/t:0};
  if(!col.topic&&labs.length){const lec=labs.filter(c=>/lectures$/i.test(c.text))[0];const pick=lec||labs.slice().sort((a,b)=>textiness(b)-textiness(a))[0];col.topic=pick;labs=labs.filter(c=>c!==pick)}
  const used=new Set(Object.values(col).filter(Boolean).concat(labs).map(c=>c.i));
  const extraCols=cols.filter(c=>!used.has(c.i));
  const lab=labs[0],lab2=labs[1];extraCols.forEach(c=>used.add(c.i));if(lab)used.add(lab.i);if(lab2)used.add(lab2.i);
  /* unlabeled columns: detect a run of consecutive dates, weekday names and day tags */
  const unl=new Set();rows.forEach((r,ri)=>{if(r&&ri>hr)r.forEach((c,ci)=>{if(c&&!used.has(ci)&&!cols.some(x=>x.i===ci))unl.add(ci)})});
  const lecRowIdx=[];
  /* first pass: which rows are lecture rows */
  const isLecRow=r=>{const row=rows[r];if(!row)return false;const g=i=>i==null?'':cellText(row[i]);return !!(g(lab&&lab.i)||g(col.topic&&col.topic.i)||g(col.status&&col.status.i)||(col.dur&&row[col.dur.i]&&durOf(row[col.dur.i],col.dur.text)>0&&(g(lab&&lab.i)||g(col.topic&&col.topic.i))))};
  for(let r=hr+1;r<=lastRow;r++)if(isLecRow(r))lecRowIdx.push(r);
  let dateCol=col.date?col.date.i:null,dateRun=new Map();
  if(dateCol===null){let best=null;for(const ci of unl){const run=new Map();let prev=null;for(const r of lecRowIdx){const d=dateOf(rows[r]&&rows[r][ci]);if(!d)continue;if(prev===null||(0<=diffDays(prev,d)&&diffDays(prev,d)<=3)){run.set(r,d);prev=d}else break}
      if(run.size>=3&&(!best||run.size>best.run.size))best={ci,run}}
    if(best){dateCol=best.ci;dateRun=best.run}}
  else for(const r of lecRowIdx){const d=dateOf(rows[r][dateCol]);if(d)dateRun.set(r,d)}
  const out=[],report={headerRow:hr+1,scanned:lastRow-hr,blank:0,totals:[],ignored:0,ignoredEx:[],panel:[]};
  const occ={};
  for(let r=hr+1;r<=lastRow;r++){
    const row=rows[r];if(!row||!row.some(Boolean)){report.blank++;continue}
    const g=i=>i==null?'':cellText(row[i]);
    let label=g(lab&&lab.i),topic=g(col.topic&&col.topic.i),statusT=g(col.status&&col.status.i);
    const dm=col.dur?durOf(row[col.dur.i],col.dur.text):null;
    if(!label&&!topic&&!statusT){ if(dm>0)report.totals.push({row:r+1,min:dm});
      /* side-panel / stray cells */ row.forEach((c,ci)=>{if(c){report.ignored++;if(report.ignoredEx.length<6)report.ignoredEx.push(cellText(c))}}); continue}
    const x={};extraCols.forEach(c=>{const t=g(c.i);if(t!=='')x[c.text]=t});
    let tag='';
    row.forEach((c,ci)=>{if(!c||used.has(ci)||ci===dateCol)return;const t=cellText(c);
      if(WEEKDAYS.test(t)&&unl.has(ci))x.Weekday=t;else if(/^(holiday|off|leave|rest( day)?|break)$/i.test(t)&&unl.has(ci))tag=t;
      else{report.ignored++;if(report.ignoredEx.length<6)report.ignoredEx.push(t)}});
    const base=label+'|'+(lab2?g(lab2.i):'')+'|'+topic.slice(0,40);occ[base]=(occ[base]||0)+1;
    const xs=/^miss/i.test(statusT)?'missed':/^(complete|completed|done|watched|finished|yes|y|✓|✔)$/i.test(statusT)||/^complet/i.test(statusT)?'done':'pending';
    let dl=null;if(col.dl){const t=g(col.dl.i);dl=/^(complete|completed|downloaded|done|yes|y|✓|✔)$/i.test(t)||/^(download|complet)/i.test(t)&&!/not|pending|no/i.test(t)}
    const date=dateRun.get(r)||null;
    let L=label;if(L.length>16){const m=L.match(/^(day[\s-]*\d+|lecture[\s-]*\d+)/i);if(m){x[lab.text]=L;L=m[1]}}
    out.push({key:base+'#'+occ[base],l:L,n:lab2?g(lab2.i):'',t:topic,m:dm>0?dm:0,hasDur:dm!==null&&dm>0,xs,dl,date,tag,pri:col.pri?g(col.pri.i).toLowerCase():'',note:col.note?g(col.note.i):'',x,srcRow:r+1});
  }
  /* de-duplicate display labels (7, 7, 7 -> 7a, 7b, 7c) */
  const cnt={};out.forEach(o=>cnt[o.l]=(cnt[o.l]||0)+1);const seen={};
  out.forEach(o=>{if(o.l&&cnt[o.l]>1){seen[o.l]=(seen[o.l]||0)+1;o.l=o.l+String.fromCharCode(96+seen[o.l]%26||122)}});
  /* side panel label/value pairs (informational) */
  for(let r=0;r<=lastRow;r++){const row=rows[r];if(!row)continue;row.forEach((c,ci)=>{if(c&&c.k==='s'&&/hours per day|total hours|completed hours|remaining hours|expected date|weekly target|by (nov|dec|jan)/i.test(c.v)&&row[ci+1])report.panel.push(c.v+': '+cellText(row[ci+1]))})}
  report.columns={label:lab&&lab.text,lecture2:lab2&&lab2.text,topic:col.topic&&col.topic.text,duration:col.dur&&col.dur.text,status:col.status&&col.status.text,download:col.dl&&col.dl.text,date:col.date?col.date.text:(dateCol!==null?'(unlabelled date column '+colName(dateCol)+')':null),extras:extraCols.map(c=>c.text)};
  return{type:'lectures',name:sh.name,rows:out,report};
}
function parseChecklistSheet(sh){
  const rows=sh.rows,hr=findHeader(rows)>=0?findHeader(rows):0;const H=rows[hr]||[];
  const idx=re=>H.findIndex(c=>c&&c.k==='s'&&re.test(c.v.toLowerCase()));
  const iT=idx(/change|item|task|point|description/),iG=idx(/chapter|group|section|categor|unit|subject/),iR=idx(/ref|page/),iD=idx(/^done|^completed?/),iN=idx(/^#$|^no\.?$|^sr/);
  const items=[];
  for(let r=hr+1;r<rows.length;r++){const row=rows[r];if(!row)continue;const g=i=>i<0?'':cellText(row[i]);const t=g(iT);if(!t)continue;
    items.push({key:g(iN)||String(items.length+1),g:g(iG),t,ref:g(iR),done:/^(yes|y|done|true|complete|completed|✓|✔)/i.test(g(iD))})}
  return{type:'checklist',name:sh.name,items,report:{headerRow:hr+1,imported:items.length,scanned:rows.length-hr-1}};
}
function parseInfoSheet(sh){const lines=[];sh.rows.forEach(r=>{if(r){const t=r.map(cellText).filter(Boolean).join(' · ');if(t)lines.push(t)}});return{type:'info',name:sh.name,lines}}
function parseWorkbook(wb){return wb.sheets.map(sh=>{const t=classifySheet(sh);return t==='lectures'?parseLectureSheet(sh):t==='checklist'?parseChecklistSheet(sh):parseInfoSheet(sh)})}

/* ===================== CORE: dates, state, storage ===================== */
const pad=n=>String(n).padStart(2,'0');
function dstr(d){return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())}
function todayStr(){return dstr(new Date())}
function pd(s){return new Date(s+'T00:00:00')}
function addDays(s,n){const d=pd(s);d.setDate(d.getDate()+n);return dstr(d)}
function diffDays(a,b){return Math.round((pd(b)-pd(a))/864e5)}
function nowMin(){const d=new Date();return d.getHours()*60+d.getMinutes()}
const tmin=t=>{const m=String(t||'0:0').match(/(\d+):(\d+)/);return m?+m[1]*60+ +m[2]:0};
const mtime=m=>pad(Math.floor(m/60)%24)+':'+pad(Math.round(m%60));
const $=s=>document.querySelector(s);
const esc=s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const pct=(a,b)=>b>0?Math.round(Math.min(a,b)/b*100):0;
const fh=m=>{m=Math.round(m);if(m<=0)return'0m';const h=Math.floor(m/60),r=m%60;return h?(r?h+'h '+r+'m':h+'h'):r+'m'};
const fhd=m=>(Math.round(m/6)/10)+'h';
const fd=(s,short)=>{if(!s)return'—';const d=pd(s);return d.toLocaleDateString('en-IN',short?{day:'numeric',month:'short'}:{day:'numeric',month:'short',year:'numeric'})};
const fdw=s=>pd(s).toLocaleDateString('en-IN',{weekday:'short',day:'numeric',month:'short'});
let _id=Date.now();const uid=p=>(p||'x')+(_id++).toString(36);
const PALETTE=['#4f46e5','#16a34a','#ea580c','#db2777','#0891b2','#7c3aed','#ca8a04','#0d9488','#e11d48','#475569'];

const PROFILES_KEY='cft3_profiles';
let PROF={active:'p1',list:[{id:'p1',name:'My dashboard'}]};
try{const p=JSON.parse(localStorage.getItem(PROFILES_KEY));if(p&&p.list&&p.list.length)PROF=p}catch(e){}
const dataKey=()=>'cft3_data_'+PROF.active;
let S=null,V=1,_cache={},storageOK=true;
function defaults(){
  return{v:3,
    settings:{goal:6,lecPerDay:5,lightDay:0,testDay:6,dayStart:'07:00',dayEnd:'22:00',lunchAt:'13:00',lunchMin:45,sessionMin:100,breakMin:15,warmup:true,streakMin:60,syncSec:10,revSteps:[3,7,21],revRepeat:21,notify:true,backupDays:7,
      theme:{mode:'auto',accent:'#4f46e5',radius:14,font:'sans',density:'cozy',preset:'Indigo'}},
    exams:[{id:'e1',name:'CA Final – May 2027 attempt',date:'2027-05-02',note:'Date is approximate – update when ICAI announces the schedule.'}],
    milestones:[],subjects:[],lectures:{},chapters:{},tasks:[],logs:[],practice:[],tests:[],notes:[],reminders:[],checklists:[],fixed:[],
    layout:{preset:'Everything',widgets:LAYOUTS.Everything.map(w=>({...w}))},
    sync:{files:[]},imports:[],pomo:{work:25,short:5,long:15,every:4,subj:'',task:'',run:null,done:0},
    meta:{created:todayStr(),lastBackup:null,nid:1}};
}
const LAYOUTS={
  Everything:[['today',1],['plan',2],['classes',2],['goals',1],['alerts',1],['schedule',2],['milestones',1],['exams',1],['revision',1],['missed',1],['pomodoro',1],['heatmap',2],['trend',1],['reminders',1],['notes',1],['tests',1],['downloads',1],['checklists',1]],
  Focus:[['today',1],['pomodoro',1],['alerts',1],['plan',2],['classes',1],['reminders',1]],
  Planner:[['plan',2],['today',1],['schedule',2],['goals',1],['revision',1],['missed',1],['reminders',1],['notes',1]],
  Analytics:[['goals',1],['heatmap',2],['trend',1],['classes',2],['milestones',1],['tests',1],['exams',1]],
  Minimal:[['today',1],['plan',2],['classes',3]]
};
Object.keys(LAYOUTS).forEach(k=>LAYOUTS[k]=LAYOUTS[k].map(([id,size])=>({id,size,hidden:false})));
const APP_VER=window.APP_VER||'dev';
const MIGRATIONS={/* add {2:fn} style upgrade steps here when the schema version (v) is raised */};
function migrate(x){while(x&&x.v<3&&MIGRATIONS[x.v]){x=MIGRATIONS[x.v](x)}return x}
function load(){
  let raw=null;
  try{raw=localStorage.getItem(dataKey())}catch(e){storageOK=false;return false}
  if(!raw)return false;
  try{
    const x=migrate(JSON.parse(raw));
    if(x&&x.v===3){
      /* keep ONE snapshot of the data from before an app-version change, so an upgrade can always be undone */
      try{const sk='cft3_snap_'+PROF.active;if(localStorage.getItem('cft3_ver')!==APP_VER){localStorage.setItem(sk,raw);localStorage.setItem('cft3_ver',APP_VER)}}catch(e){}
      S=x;return true}
    throw new Error('unsupported data version '+(x&&x.v));
  }catch(e){
    /* never overwrite data we cannot read: park it under a recovery key first */
    try{localStorage.setItem('cft3_recover_'+PROF.active+'_'+Date.now(),raw)}catch(_){}
    LOAD_PROBLEM='Your saved data could not be read ('+e.message+'). A copy was kept for recovery; starting from a fresh dashboard.';
    return false}
}
let LOAD_PROBLEM='';
let _saveT=null;
function save(){V++;_cache={};clearTimeout(_saveT);_saveT=setTimeout(flush,250)}
function flush(){try{localStorage.setItem(dataKey(),JSON.stringify(S));storageOK=true}catch(e){storageOK=false}}
window.addEventListener('beforeunload',flush);
const memo=(k,fn)=>(k in _cache)?_cache[k]:(_cache[k]=fn());
const sj=id=>S.subjects.find(s=>s.id===id);
const sjn=id=>{const s=sj(id);return s?s.name:'—'};
const scol=id=>{const s=sj(id);return s?s.color:'#888'};

/* ===================== lecture & chapter maths ===================== */
const LL=sid=>(S.lectures[sid]||[]).filter(l=>!l.gone);
const sumM=(L,f)=>{let t=0;for(const l of L)if(!f||f(l))t+=l.m;return t};
function released(sid){return sumM(LL(sid))}
function totalM(sid){const s=sj(sid),r=released(sid);return s&&s.batchH?Math.max(r,s.batchH*60):r}
function doneM(sid){return sumM(LL(sid),l=>l.d)}
function remM(sid){return totalM(sid)-doneM(sid)}
function hasLectures(sid){return LL(sid).length>0}
function pendingSubjects(){return S.subjects.filter(s=>!s.hide&&remM(s.id)>0&&hasLectures(s.id))}
const chLecs=(sid,c)=>LL(sid).filter(l=>l.ch===c.id);
function lecDone(sid,c){const A=chLecs(sid,c);return A.length?A.every(l=>l.d):!!c.cls}
function score(sid,c){return(lecDone(sid,c)?25:0)+(c.r1?20:0)+(c.r2?20:0)+(c.r3?10:0)+(c.pr?15:0)+(c.conf>=4?10:0)}
function isReady(sid,c){return lecDone(sid,c)&&c.r1&&c.r2&&c.pr&&c.conf>=4}
function chStatus(sid,c){if(!lecDone(sid,c))return['Lectures pending',''];if(isReady(sid,c))return['Exam-ready','good'];if(c.r1||c.r2||c.r3)return['Revising','acc'];return['Lectures done','warn']}
function subjScore(sid){const C=S.chapters[sid]||[];return C.length?Math.round(C.reduce((a,c)=>a+score(sid,c),0)/C.length):0}
function learnedOn(sid,c){let m=null;for(const l of chLecs(sid,c)){if(l.doneOn&&(!m||l.doneOn>m))m=l.doneOn}return m}
function revCount(c){return(c.r1?1:0)+(c.r2?1:0)+(c.r3?1:0)}
function nextRevDue(sid,c){const st=S.settings.revSteps,k=revCount(c),base=c.last||learnedOn(sid,c),gap=k<st.length?st[k]:S.settings.revRepeat;return base?addDays(base,gap):todayStr()}
function revQueue(){return memo('rq',()=>{const out=[],T=todayStr();
  S.subjects.forEach(s=>{if(s.hide)return;(S.chapters[s.id]||[]).forEach((c,i)=>{if(!lecDone(s.id,c))return;const due=nextRevDue(s.id,c);out.push({s,c,due,k:revCount(c),idx:i,overdue:due<=T})})});
  return out.sort((a,b)=>(a.due<b.due?-1:a.due>b.due?1:a.k-b.k||a.idx-b.idx))})}

/* ===================== study time accounting ===================== */
function dayIdx(){return memo('di',()=>{const by={};
  const add=(d,sid,m)=>{if(!d||!(m>0))return;const o=by[d]||(by[d]={t:0,s:{}});o.t+=m;o.s[sid]=(o.s[sid]||0)+m};
  S.tasks.forEach(t=>{if(t.type!=='break')add(t.date,t.subj,+t.act||0)});
  S.logs.forEach(l=>add(l.date,l.subj,l.min));
  for(const sid in S.lectures)S.lectures[sid].forEach(l=>{if(l.d&&l.doneOn&&l.src==='import'&&!l.gone)add(l.doneOn,sid,l.m)});
  return by})}
const minOn=d=>(dayIdx()[d]||{t:0}).t;
const goalFor=d=>pd(d).getDay()===S.settings.lightDay?S.settings.goal/2:S.settings.goal;
const goalMin=d=>goalFor(d)*60;
function streak(){let n=0,d=todayStr();const min=S.settings.streakMin;if(minOn(d)<min)d=addDays(d,-1);while(minOn(d)>=min){n++;d=addDays(d,-1)}return n}
function longestStreak(){const by=dayIdx(),ds=Object.keys(by).filter(d=>by[d].t>=S.settings.streakMin).sort();let best=0,cur=0,prev=null;for(const d of ds){cur=prev&&diffDays(prev,d)===1?cur+1:1;best=Math.max(best,cur);prev=d}return best}
function rangeMin(a,b){let t=0;for(let d=a;d<=b;d=addDays(d,1))t+=minOn(d);return t}
function weekRange(date){const d=pd(date),k=(d.getDay()+6)%7;const s=addDays(date,-k);return[s,addDays(s,6)]}
function monthRange(date){const d=pd(date);return[dstr(new Date(d.getFullYear(),d.getMonth(),1)),dstr(new Date(d.getFullYear(),d.getMonth()+1,0))]}
function goalSum(a,b){let t=0;for(let d=a;d<=b;d=addDays(d,1))t+=goalMin(d);return t}

/* ===================== subjects, deadlines & the lecture schedule ===================== */
function ensureSubject(name,extra){
  let s=S.subjects.find(x=>x.name.toLowerCase()===name.toLowerCase());
  if(!s){s={id:uid('s'),name,full:name,fac:'',res:'',color:PALETTE[S.subjects.length%PALETTE.length],pri:2,due:'',batchH:0,hide:false};S.subjects.push(s)}
  Object.assign(s,extra||{});if(!S.lectures[s.id])S.lectures[s.id]=[];if(!S.chapters[s.id])S.chapters[s.id]=[];return s;
}
function subjDue(sid){
  const s=sj(sid);if(s&&s.due)return s.due;
  let best=null;S.milestones.forEach(m=>{if(m.kind==='lec'&&(m.subjects==='all'||(m.subjects||[]).includes(sid))&&m.date&&(!best||m.date<best))best=m.date});
  if(best)return best;const ex=S.exams[0];return ex?addDays(ex.date,-60):addDays(todayStr(),90);
}
const isMissed=l=>!l.d&&!l.gone&&(l.missed||(memo('ms',()=>{const st=new Set();S.tasks.forEach(t=>{if(t.lec&&!t.done&&t.date<todayStr())st.add(t.lec)});return st}).has(l.id)));
function lecCap(date){const c=Math.min(S.settings.lecPerDay,S.settings.goal*0.72)*60;return pd(date).getDay()===S.settings.lightDay?c/2:c}
/* Deficit round-robin: each day's lecture time is shared between subjects in proportion to how urgent they are. */
function schedule(){return memo('sch',()=>{
  const T=todayStr(),days={},planned={},finish={};
  const subs=S.subjects.filter(s=>!s.hide).map(s=>{const q=LL(s.id).filter(l=>!l.d&&l.m>0&&!l.arch);q.sort((a,b)=>(isMissed(b)?1:0)-(isMissed(a)?1:0));
    const extra=Math.max(0,totalM(s.id)-released(s.id));return{s,q,i:0,credit:0,due:subjDue(s.id),extra,w:[1.4,1,0.7][(s.pri||2)-1]||1}}).filter(x=>x.q.length||x.extra>0);
  let day=T,guard=0;
  const doneToday=sumM(Object.values(S.lectures).flat(),l=>l.d&&l.doneOn===T&&l.src==='app'&&!l.gone);
  while(subs.some(x=>x.i<x.q.length)&&guard++<700){
    let cap=lecCap(day);if(day===T)cap=Math.max(0,cap-doneToday);
    const act=subs.filter(x=>x.i<x.q.length);
    const rate=x=>{const rem=sumM(x.q.slice(x.i))+x.extra;return x.w*rem/Math.max(7,diffDays(day,x.due)+1)};
    const tw=act.reduce((a,x)=>a+rate(x),0)||1;act.forEach(x=>x.credit+=cap*rate(x)/tw);
    let used=0;const todayList=[];
    while(true){
      const cand=act.filter(x=>x.i<x.q.length).sort((a,b)=>b.credit-a.credit);let pick=null;
      for(const x of cand){const l=x.q[x.i];if(todayList.length===0||used+l.m<=cap+30){pick=x;break}}
      if(!pick||(used>=cap&&todayList.length))break;
      const l=pick.q[pick.i++];pick.credit-=l.m;used+=l.m;todayList.push(l);planned[l.id]=day;
      if(pick.i>=pick.q.length&&!finish[pick.s.id])finish[pick.s.id]=day;
    }
    if(todayList.length)days[day]=todayList;
    day=addDays(day,1);
  }
  subs.forEach(x=>{if(!x.q.length&&!finish[x.s.id])finish[x.s.id]=null});
  /* hours/day needed to hit each deadline (lectures only) */
  const need=(()=>{let worst=0;const arr=subs.map(x=>({due:x.due,rem:sumM(x.q.slice(x.i>=0?0:0))+x.extra}));
    arr.forEach(a=>{const tot=arr.filter(b=>b.due<=a.due).reduce((s,b)=>s+b.rem,0);const d=Math.max(1,diffDays(T,a.due)+1);worst=Math.max(worst,tot/60/d*(7/6.5))});return worst})();
  const last=Object.keys(days).sort().pop()||T;
  return{days,planned,finish,need,last,subs:subs.map(x=>({id:x.s.id,due:x.due,extra:x.extra}))};
})}
function subjProjection(sid){const sc=schedule(),fin=sc.finish[sid],rem=remM(sid),due=subjDue(sid);
  const extra=Math.max(0,totalM(sid)-released(sid));
  return{rem,fin,due,slack:fin?diffDays(fin,due):null,extra}}

/* ===================== import: apply parsed sheets to state ===================== */
function autoGroup(rows){
  let g=[];
  rows.forEach((r,i)=>{const k=normTopic(r.t),last=g[g.length-1];
    const filler=!r.m&&r.t&&/rtp|extra que|general|study plan|raw/i.test(r.t);
    if(last&&(last.k===k||(filler&&last.k)))last.idx.push(i);else g.push({k,name:r.t,idx:[i]})});
  if(g.length>30||(rows.length/Math.max(1,g.length))<1.7){
    const chunk=480;g=[];let cur=null,acc=0;
    rows.forEach((r,i)=>{if(!cur||acc>=chunk){cur={k:'',name:'',idx:[]};g.push(cur);acc=0}cur.idx.push(i);acc+=r.m});
    g.forEach((x,n)=>{const a=rows[x.idx[0]],b=rows[x.idx[x.idx.length-1]];x.name='Part '+(n+1)+' · '+(a.l||'')+(b!==a?' – '+(b.l||''):'')+' · '+(a.t||'').slice(0,40)});
  }
  g.forEach(x=>{if(!x.name||!x.k)x.name=x.name||'Untitled / not yet released'});
  return g;
}
function newChapter(sid,name){const c={id:uid('c'),name,cls:false,r1:false,r2:false,r3:false,pr:false,conf:0,last:''};S.chapters[sid].push(c);return c}
function applyLectureSheet(subject,parsed,opt){
  opt=opt||{};const sid=subject.id,old=S.lectures[sid]||[];const first=old.length===0;
  const byKey=new Map(old.filter(l=>l.k).map(l=>[l.k,l]));
  const out=[];let added=0,updated=0,statusChanged=0;
  const chMap=opt.chList;                      // optional array of chapter names aligned with rows (seed presets)
  const groups=first&&!chMap?autoGroup(parsed.rows):null;const chOf=new Map();
  if(groups)groups.forEach(g=>{const c=newChapter(sid,g.name);g.idx.forEach(i=>chOf.set(i,c.id))});
  const chByName=new Map(S.chapters[sid].map(c=>[c.name,c]));
  let prevCh=S.chapters[sid].length?S.chapters[sid][S.chapters[sid].length-1].id:null;
  parsed.rows.forEach((r,i)=>{
    let l=byKey.get(r.key);
    if(l){updated++;l.l=r.l;l.n=r.n;l.t=r.t;l.m=r.m;l.x=r.x;l.row=r.srcRow;l.gone=false;l.tag=r.tag;byKey.delete(r.key);
      if(l.xs!==r.xs){if(l.xs!==undefined)statusChanged++;l.d=r.xs==='done';l.xs=r.xs;if(l.d&&!l.doneOn){l.doneOn=r.date||todayStr();l.src='import'}if(!l.d)l.doneOn=null;if(r.xs==='missed')l.missed=true}
      if(r.dl!==null&&l.xdl!==r.dl){l.dl=r.dl;l.xdl=r.dl}
      if(r.date&&l.d)l.doneOn=r.date;else if(r.date)l.sdate=r.date;
      if(r.pri&&!l.pri)l.pri=r.pri;if(r.note&&!l.note)l.note=r.note}
    else{added++;
      let ch=chMap?(chByName.get(chMap[i])||(chByName.set(chMap[i],newChapter(sid,chMap[i])),chByName.get(chMap[i]))).id:chOf.get(i)||prevCh;
      if(!ch){ch=newChapter(sid,'Imported lectures').id}
      l={id:uid('l'),k:r.key,l:r.l,n:r.n,t:r.t,m:r.m,d:r.xs==='done',xs:r.xs,dl:r.dl,xdl:r.dl,ch,pri:/high|urgent|^1$/.test(r.pri)?'high':/low/.test(r.pri)?'low':'',note:r.note||'',doneOn:r.xs==='done'?r.date:null,sdate:r.xs!=='done'?r.date:null,missed:r.xs==='missed',arch:false,src:'import',x:r.x,row:r.srcRow,tag:r.tag};
    }
    prevCh=l.ch;out.push(l)});
  const gone=[...byKey.values()];gone.forEach(l=>{if(l.src==='import'){if(opt.removeGone){}else{l.gone=true;out.push(l)}}else out.push(l)});
  S.lectures[sid]=out;
  return{added,updated,gone:gone.length,statusChanged};
}
function applyChecklist(name,parsed,info){
  let cl=S.checklists.find(c=>c.name===name);if(!cl){cl={id:uid('k'),name,notes:'',items:[]};S.checklists.push(cl)}
  const byKey=new Map(cl.items.map(i=>[i.key,i]));
  cl.items=parsed.items.map(it=>{const ex=byKey.get(it.key);return ex?Object.assign(ex,{g:it.g,t:it.t,ref:it.ref}):{id:uid('i'),key:it.key,g:it.g,t:it.t,ref:it.ref,done:it.done}});
  if(info)cl.notes=info;return cl;
}
/* the full import, from parsed workbooks with a mapping chosen in the preview dialog */
function commitImport(fileName,parsedSheets,plan){
  const log={at:new Date().toISOString(),file:fileName,sheets:[]};const info=parsedSheets.filter(p=>p.type==='info').flatMap(p=>p.lines).join('\n');
  parsedSheets.forEach((p,i)=>{const pl=plan[i];if(!pl||pl.skip)return;
    if(p.type==='lectures'){
      const s=ensureSubject(pl.subject,pl.extra||{});const r=applyLectureSheet(s,p,{chList:pl.chList});
      const sum=sumM(p.rows.map(x=>({m:x.m})),()=>true),tot=p.report.totals.length?p.report.totals[p.report.totals.length-1].min:null;
      log.sheets.push({sheet:p.name,type:'lectures',subject:s.name,rows:p.rows.length,scanned:p.report.scanned,blank:p.report.blank,min:sum,sheetTotal:tot,ignored:p.report.ignored,panel:p.report.panel,cols:p.report.columns,res:r,headerRow:p.report.headerRow,done:p.rows.filter(x=>x.xs==='done').length});
      s.src=fileName;
    }else if(p.type==='checklist'){const cl=applyChecklist(pl.subject||p.name,p,info);log.sheets.push({sheet:p.name,type:'checklist',subject:cl.name,rows:p.items.length,scanned:p.report.scanned})}
  });
  S.imports.unshift(log);S.imports=S.imports.slice(0,30);save();return log;
}
/* auto-plan per file name -> subject name */
function guessSubject(fileName,sheetName){
  let b=fileName.replace(/\.[a-z0-9]+$/i,'').replace(/^\d{6,}[-_]/,'');
  const first=b.split(/[_\- ]+/)[0];
  if(first&&first.length<=6)return first.toUpperCase()===first?first:(first[0].toUpperCase()+first.slice(1));
  return b.replace(/[_]+/g,' ').trim()||sheetName;
}

/* ===================== AUTO PLANNER, MILESTONES, ALERTS ===================== */
function fixSid(){for(const sid in S.lectures)S.lectures[sid].forEach(l=>{l.sid=sid})}
const round5=m=>Math.max(5,Math.round(m/5)*5);
function fixedFor(date){const wd=pd(date).getDay();return S.fixed.filter(f=>!f.days||!f.days.length||f.days.includes(wd)).map(f=>({name:f.name,s:tmin(f.start),e:tmin(f.end)}))}
function splitMin(total,max){if(total<20)return[];const n=Math.ceil(total/max),each=round5(total/n);return Array.from({length:n},()=>each)}
function pickRoundRobin(items,n,subjOrder){
  const by={};items.forEach(it=>(by[it.s.id]=by[it.s.id]||[]).push(it));
  const order=(subjOrder||Object.keys(by)).filter(k=>by[k]);Object.keys(by).forEach(k=>{if(!order.includes(k))order.push(k)});
  const out=[];let guard=0;while(out.length<n&&guard++<500){let any=false;for(const k of order){if(out.length>=n)break;if(by[k].length){out.push(by[k].shift());any=true}}if(!any)break}
  return out;
}
const REV_TEXT=[
  'first pass – skim notes / chart book, redo 4–5 key illustrations, list doubts',
  'second pass – close-book recall of provisions, then redo the sums you got wrong',
  'final pass – one-page summary + rapid-fire questions',
  'maintenance – quick recall and 3 mixed questions'];
function planDay(date){
  const st=S.settings,T=todayStr();
  const kept=S.tasks.filter(t=>t.date===date&&!(t.auto&&!t.done&&!t.lock));
  const keptLec=new Set(kept.filter(t=>t.lec).map(t=>t.lec));
  const keptMin=kept.reduce((a,t)=>a+(t.type==='break'?0:(+t.plan||0)),0);
  let dayMin=Math.round(goalMin(date));
  if(date===T)dayMin-=minOn(T);dayMin-=keptMin;
  const sc=schedule(),lecList=(sc.days[date]||[]).filter(l=>!keptLec.has(l.id));
  const B=[];
  let lecMin=0;lecList.forEach(l=>{lecMin+=l.m;B.push({type:'lecture',subj:l.sid,text:`${l.l?l.l+' · ':''}${l.t||'Lecture'}`,plan:Math.round(l.m),lec:l.id,ch:l.ch})});
  const subjOfLec=[...new Set(lecList.map(l=>l.sid))];
  let recap=0;
  if(lecMin>0){recap=Math.min(45,Math.max(15,round5(lecMin*0.15)));
    B.push({type:'recap',subj:subjOfLec[0],text:`Recap & 1-page notes: ${lecList.map(l=>sjn(l.sid)+' '+(l.l||'')).join(', ')}`,plan:recap,after:true})}
  let rest=Math.max(0,dayMin-lecMin-recap);
  const isTest=pd(date).getDay()===st.testDay&&rest>=120;let testMin=0;
  if(isTest){testMin=90;rest-=90}
  const lecDay=lecMin>0;const revMin=Math.round(rest*(lecDay?0.55:0.5)),prMin=rest-revMin;
  const q=revQueue();const usedCh=new Set(kept.filter(t=>t.ch).map(t=>t.ch));
  const revChunks=splitMin(revMin,45),prChunks=splitMin(prMin,45);
  const pendIds=pendingSubjects().map(s=>s.id);
  const revPicks=pickRoundRobin(q.filter(x=>!usedCh.has(x.c.id)),revChunks.length,S.subjects.map(s=>s.id).filter(i=>!pendIds.includes(i)).concat(pendIds));
  const revB=revChunks.map((d,i)=>{const p=revPicks[i];return p?{type:'revision',subj:p.s.id,ch:p.c.id,plan:d,text:`Revise ${p.s.name}: ${p.c.name} – ${REV_TEXT[Math.min(p.k,3)]}`}:{type:'revision',subj:'',plan:d,text:'Revise: pick any learnt chapter and recall it closed-book'}});
  revPicks.forEach(p=>usedCh.add(p.c.id));
  const prItems=[];S.subjects.forEach(s=>{if(s.hide)return;(S.chapters[s.id]||[]).forEach((c,i)=>{if(lecDone(s.id,c)&&!c.pr&&!usedCh.has(c.id))prItems.push({s,c,k:c.r1?0:1,idx:i})})});
  prItems.sort((a,b)=>a.k-b.k);
  const prPicks=pickRoundRobin(prItems,prChunks.length);
  const prB=prChunks.map((d,i)=>{const p=prPicks[i];const n=Math.max(3,Math.round(d/6));return p?{type:'practice',subj:p.s.id,ch:p.c.id,plan:d,text:`Practice ${p.s.name}: ${p.c.name} – ~${n} questions (RTP / MTP / past papers first); mark errors, log accuracy`}:{type:'practice',subj:'',plan:d,text:`Mixed practice – ~${n} questions from any learnt chapter; log accuracy`}});
  const testB=isTest?[{type:'test',subj:'',plan:testMin,text:'Timed self-test – attempt a full section under exam conditions, then review mistakes'}]:[];
  const warm=st.warmup&&revB.length?[revB.shift()]:[];
  const seq=[...warm,...B.filter(b=>!b.after),...B.filter(b=>b.after),...revB,...prB,...testB];
  /* lay the blocks out on the clock, with breaks, lunch and fixed commitments */
  const busy=fixedFor(date).concat(kept.filter(t=>t.start&&t.end).map(t=>({name:t.text,s:tmin(t.start),e:tmin(t.end)})));
  let cur=tmin(st.dayStart);if(date===T)cur=Math.max(cur,Math.ceil(nowMin()/5)*5);
  let since=0,lunch=date===T&&cur>tmin(st.lunchAt)+60;const out=[];
  const mk=(o)=>({id:uid('t'),date,start:mtime(o.s),end:mtime(o.e),subj:o.subj||'',type:o.type,text:o.text,plan:o.type==='break'?0:o.plan,act:0,done:false,lec:o.lec||null,ch:o.ch||null,auto:true,lock:false});
  const clash=(s,e)=>busy.find(b=>s<b.e&&e>b.s);
  for(const b of seq){
    if(!lunch&&cur>=tmin(st.lunchAt)){out.push(mk({s:cur,e:cur+st.lunchMin,type:'break',text:'Lunch'}));cur+=st.lunchMin;lunch=true;since=0}
    let c;while(c=clash(cur,cur+b.plan)){cur=c.e;since=0}
    if(since>=st.sessionMin){out.push(mk({s:cur,e:cur+st.breakMin,type:'break',text:'Short break'}));cur+=st.breakMin;since=0;while(c=clash(cur,cur+b.plan)){cur=c.e}}
    out.push(mk({s:cur,e:cur+b.plan,...b}));cur+=b.plan;since+=b.plan;
  }
  return{tasks:out,overflow:Math.max(0,cur-tmin(st.dayEnd)),lecMin,total:seq.reduce((a,b)=>a+b.plan,0)};
}
function autoPlan(date){
  const r=planDay(date);
  S.tasks=S.tasks.filter(t=>!(t.date===date&&t.auto&&!t.done&&!t.lock)).concat(r.tasks);save();return r;
}

/* ===================== milestones ===================== */
function msProgress(m){
  const ids=m.subjects==='all'||!m.subjects?S.subjects.filter(s=>!s.hide).map(s=>s.id):m.subjects;
  if(m.kind==='lec'){let d=0,t=0;ids.forEach(i=>{d+=doneM(i);t+=totalM(i)});return pct(d,t)}
  if(m.kind==='rev1'||m.kind==='rev2'||m.kind==='rev3'){const k='r'+m.kind.slice(3);let n=0,t=0;ids.forEach(i=>(S.chapters[i]||[]).forEach(c=>{t++;if(c[k])n++}));return pct(n,t)}
  if(m.kind==='tests')return pct(S.tests.length,m.target||10);
  return m.done?100:0;
}
/* ===================== reminders, alerts ===================== */
function missedLecs(){const out=[];for(const sid in S.lectures)S.lectures[sid].forEach(l=>{if(isMissed(l))out.push(l)});return out}
function alertsList(){
  const A=[],T=todayStr(),sc=schedule(),nowM=nowMin();
  S.reminders.forEach(r=>{if(!r.done&&r.at&&new Date(r.at)<=new Date())A.push({sev:'bad',t:'Reminder: '+r.text,b:'Due '+new Date(r.at).toLocaleString('en-IN',{day:'numeric',month:'short',hour:'2-digit',minute:'2-digit'}),act:'rem-done',p:r.id,btn:'Done'})});
  const gm=goalMin(T),dm=minOn(T);
  if(dm>=gm)A.push({sev:'good',t:'Daily target reached 🎉',b:`${fh(dm)} studied today (goal ${fhd(gm)}).`});
  else if(nowM>=18*60&&dm<gm*0.6)A.push({sev:'warn',t:'Behind on today\'s target',b:`${fh(dm)} of ${fhd(gm)} done – ${fh(gm-dm)} to go.`});
  pendingSubjects().forEach(s=>{const p=subjProjection(s.id);if(p.fin&&p.slack<0)A.push({sev:p.slack<-14?'bad':'warn',t:`${s.name} is ${-p.slack} day${p.slack===-1?'':'s'} behind schedule`,b:`Projected finish ${fd(p.fin)} vs target ${fd(p.due)}. Raise the daily goal or lecture hours, or move the target.`})});
  if(sc.need>S.settings.lecPerDay*1.02)A.push({sev:'warn',t:`Deadlines need ~${sc.need.toFixed(1)}h of lectures a day`,b:`Your lecture cap is ${S.settings.lecPerDay}h/day. Settings → Study plan lets you raise it.`});
  const ms=missedLecs();if(ms.length)A.push({sev:'warn',t:`${ms.length} missed lecture${ms.length>1?'s':''}`,b:'Planned earlier but not watched – they are scheduled first in the next plan.'});
  const rq=revQueue().filter(x=>x.overdue&&x.k>0);if(rq.length)A.push({sev:'info',t:`${rq.length} repeat revision${rq.length>1?'s':''} due`,b:rq.slice(0,3).map(x=>x.s.name+': '+x.c.name).join(' · ')});
  for(const sid of [...new Set(pendingSubjects().map(s=>s.id))]){const l=(sc.days[T]||[]).concat(sc.days[addDays(T,1)]||[]).filter(l=>l.sid===sid&&l.dl===false);
    if(l.length)A.push({sev:'info',t:`Download ${sjn(sid)} lectures before you need them`,b:`${l.length} lecture${l.length>1?'s':''} planned for the next 2 days are not downloaded (${fh(sumM(l))}).`})}
  const act=pendingSubjects().map(s=>s.id);const by=dayIdx();
  act.forEach(sid=>{let gap=0,d=T;while(gap<14&&!((by[d]||{s:{}}).s[sid]>0)){gap++;d=addDays(d,-1)}if(gap>=4&&Object.keys(by).length>3)A.push({sev:'info',t:`No ${sjn(sid)} study for ${gap} days`,b:'Keep every pending class moving a little each day.'})});
  S.exams.forEach(e=>{const d=diffDays(T,e.date);if(d>=0&&d<=45)A.push({sev:d<=14?'bad':'warn',t:`${e.name}: ${d} day${d===1?'':'s'} left`,b:fd(e.date)})});
  if(!S.meta.lastBackup||diffDays(S.meta.lastBackup,T)>=S.settings.backupDays)A.push({sev:'info',t:'Time for a backup',b:'Data & sync → Export backup keeps your progress safe.',act:'tab',p:'data',btn:'Open'});
  if(!storageOK)A.push({sev:'bad',t:'Browser storage is unavailable',b:'Progress cannot be saved in this window. Export a backup before closing.'});
  const order={bad:0,warn:1,info:2,good:3};return A.sort((a,b)=>order[a.sev]-order[b.sev]);
}

/* ===================== UI: icons, helpers, widgets ===================== */
const IC={grid:'<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/>',
 book:'<path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/>',
 cal:'<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>',
 layers:'<path d="M12 2 2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5M2 12l10 5 10-5"/>',
 chart:'<path d="M12 20V10M18 20V4M6 20v-4"/>',target:'<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>',
 check:'<path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>',
 sync:'<path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.5 9a9 9 0 0 1 14.9-3.4L23 10M1 14l4.6 4.4A9 9 0 0 0 20.5 15"/>',
 sliders:'<path d="M4 21v-7M4 10V3M12 21v-9M12 8V3M20 21v-5M20 12V3M1 14h6M9 8h6M17 16h6"/>',
 plus:'<path d="M12 5v14M5 12h14"/>',search:'<circle cx="11" cy="11" r="8"/><path d="M21 21l-4.3-4.3"/>',
 play:'<path d="M6 3l14 9-14 9z"/>',pause:'<path d="M6 4h4v16H6zM14 4h4v16h-4z"/>',moon:'<path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/>',
 x:'<path d="M18 6 6 18M6 6l12 12"/>',grip:'<circle cx="9" cy="6" r="1"/><circle cx="15" cy="6" r="1"/><circle cx="9" cy="12" r="1"/><circle cx="15" cy="12" r="1"/><circle cx="9" cy="18" r="1"/><circle cx="15" cy="18" r="1"/>',
 edit:'<path d="M12 20h9M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/>',dl:'<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/>',
 bolt:'<path d="M13 2 3 14h9l-1 8 10-12h-9z"/>',lock:'<rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',more:'<circle cx="5" cy="12" r="1.2"/><circle cx="12" cy="12" r="1.2"/><circle cx="19" cy="12" r="1.2"/>',
 left:'<path d="M15 18l-6-6 6-6"/>',right:'<path d="M9 18l6-6-6-6"/>',eye:'<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8S1 12 1 12z"/><circle cx="12" cy="12" r="3"/>',trash:'<path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6"/>'};
const icon=n=>`<svg class="ic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${IC[n]||''}</svg>`;
const bar=(p,c,cls)=>`<div class="bar ${cls||''}"><i style="width:${Math.max(0,Math.min(100,p))}%;${c?'background:'+c:''}"></i></div>`;
const chip=(t,k)=>`<span class="chip ${k||''}">${t}</span>`;
const sdot=id=>`<span class="dot" style="background:${scol(id)}"></span>`;
const subjChip=id=>`<span class="chip" style="background:color-mix(in srgb,${scol(id)} 16%,transparent);color:${scol(id)}">${esc(sjn(id))}</span>`;
const typeLabel={lecture:'Lecture',recap:'Recap',revision:'Revision',practice:'Practice',test:'Test',custom:'Task',break:'Break',focus:'Focus'};
const typeCol={lecture:'#4f46e5',recap:'#0891b2',revision:'#ea580c',practice:'#16a34a',test:'#db2777',custom:'#64748b',break:'#cbd5e1'};
function slackChip(p){if(p.rem<=0)return chip('Done','good');if(p.slack===null)return chip('—');return p.slack>=0?chip(p.slack>=3?`On track · ${p.slack}d spare`:'On track','good'):chip(`${-p.slack}d behind`,'bad')}

const W={};
const WDEF={today:['Today','1'],plan:["Today's plan",'2'],classes:['Class progress','2'],goals:['Goals','1'],alerts:['Smart alerts','1'],schedule:['Lecture schedule','2'],milestones:['Milestones','1'],exams:['Exam countdown','1'],revision:['Revision planner','1'],missed:['Missed lectures','1'],pomodoro:['Focus timer','1'],heatmap:['Consistency','2'],trend:['Weekly trend','1'],reminders:['Reminders','1'],notes:['Notes','1'],tests:['Tests','1'],downloads:['Offline downloads','1'],checklists:['Checklists','1']};

W.today=()=>{
  const T=todayStr(),m=minOn(T),g=goalMin(T),p=pct(m,g),r=46,C=2*Math.PI*r,lec=sumM(Object.values(S.lectures).flat(),l=>l.d&&l.doneOn===T&&!l.gone);
  return`<div style="position:relative"><svg class="ring" width="124" height="124" viewBox="0 0 124 124"><circle cx="62" cy="62" r="${r}" fill="none" stroke="var(--surface2)" stroke-width="10"/><circle cx="62" cy="62" r="${r}" fill="none" stroke="${m>=g?'var(--good)':'var(--accent)'}" stroke-width="10" stroke-linecap="round" stroke-dasharray="${C*p/100} ${C}" transform="rotate(-90 62 62)"/></svg>
   <div style="position:absolute;inset:0;display:grid;place-items:center;text-align:center"><div><div class="big" style="font-size:24px">${fh(m)}</div><div class="muted sm">of ${fhd(g)}</div></div></div></div>
   <div class="row sb sm"><span class="muted">Streak</span><b>${streak()} day${streak()===1?'':'s'} 🔥</b></div>
   <div class="row sb sm"><span class="muted">Lecture time today</span><b>${fh(lec)}</b></div>
   <div class="row" style="gap:6px"><button class="btn sm" data-a="quick" data-p="log">${icon('plus')}Log time</button><button class="btn sm" data-a="autoplan" data-p="${T}">${icon('bolt')}Plan my day</button></div>`};

W.plan=()=>{
  const T=todayStr(),L=S.tasks.filter(t=>t.date===T).sort((a,b)=>(a.start||'99')<(b.start||'99')?-1:1),real=L.filter(t=>t.type!=='break'),dn=real.filter(t=>t.done).length;
  if(!L.length)return`<div class="empty">Nothing planned for today.<br><button class="btn pri mt" data-a="autoplan" data-p="${T}">${icon('bolt')}Auto-plan today</button></div>`;
  return`<div class="row sb sm"><span class="muted">${dn} of ${real.length} done · ${fh(real.reduce((a,t)=>a+(+t.plan||0),0))} planned</span><span class="row"><button class="btn sm" data-a="autoplan" data-p="${T}">${icon('bolt')}Re-plan</button><button class="btn sm" data-a="tab" data-p="planner">Open planner</button></span></div>${bar(pct(dn,real.length))}
  <div>${L.slice(0,14).map(t=>taskRow(t,true)).join('')}${L.length>14?`<div class="muted sm mt">+${L.length-14} more in the planner</div>`:''}</div>`};
function taskRow(t,compact){
  const brk=t.type==='break';
  return`<div class="task ${t.done?'done':''} ${brk?'brk':''}" style="padding:7px 0;border-bottom:1px dashed var(--line)">${brk?'<span style="width:17px"></span>':`<input type="checkbox" data-c="taskdone" data-p="${t.id}" ${t.done?'checked':''}>`}
   <span class="muted sm" style="width:84px;font-variant-numeric:tabular-nums">${t.start||''}${t.end?'–'+t.end:''}</span>
   <span class="grow tx">${t.subj&&!brk?subjChip(t.subj)+' ':''}${esc(t.text)}</span>${brk?'':`<span class="muted sm">${fh(t.plan)}</span>`}</div>`}

W.classes=()=>{
  const subs=S.subjects.filter(s=>!s.hide).map(s=>({s,p:subjProjection(s.id),has:hasLectures(s.id)}));
  subs.sort((a,b)=>(b.p.rem>0)-(a.p.rem>0)||(a.p.slack??999)-(b.p.slack??999));
  return subs.map(({s,p,has})=>{const t=totalM(s.id),d=doneM(s.id),nx=LL(s.id).find(l=>!l.d&&l.m>0);
   return`<div style="cursor:pointer" data-a="goto-subj" data-p="${s.id}"><div class="row sb"><span class="row" style="gap:7px">${sdot(s.id)}<b>${esc(s.name)}</b><span class="muted sm">${esc(s.fac)}</span></span>${has?slackChip(p):chip(subjScore(s.id)+'% ready','acc')}</div>
   ${has?bar(pct(d,t),s.color):bar(subjScore(s.id),s.color)}
   <div class="muted sm">${has?`${fh(d)} of ${fh(t)} watched · <b>${fh(p.rem)}</b> left${p.fin?` · done by ~${fd(p.fin,1)} (target ${fd(p.due,1)})`:''}${nx?` · next: ${esc(nx.l)} ${esc(nx.t).slice(0,34)}`:''}`:'No lecture index imported – tracked by chapter in Chapters'}</div></div>`}).join('')||'<div class="empty">No classes yet. Import your Excel files in Data &amp; sync.</div>'};

W.goals=()=>{
  const T=todayStr(),[w0,w1]=weekRange(T),[m0,m1]=monthRange(T);
  const row=(lbl,a,b,s,e)=>{const g=goalSum(s,e),pace=goalSum(s,T>e?e:T);return`<div><div class="row sb"><b>${lbl}</b><span class="muted sm">${fh(a)} / ${fh(g)}</span></div>${bar(pct(a,g),a>=pace?'var(--good)':'var(--accent)')}<div class="muted sm">${a>=pace?'On pace':`${fh(pace-a)} behind pace`}${T<=e?' · '+(diffDays(T,e)+1)+' day(s) left':''}</div></div>`};
  return row('Today',minOn(T),0,T,T)+row('This week',rangeMin(w0,w1),0,w0,w1)+row('This month',rangeMin(m0,m1),0,m0,m1)};

W.alerts=()=>{const A=alertsList();return A.length?A.slice(0,6).map(a=>`<div class="alert ${a.sev}"><i></i><div class="grow"><b>${esc(a.t)}</b><div class="muted sm">${esc(a.b||'')}</div></div>${a.act?`<button class="btn sm" data-a="${a.act}" data-p="${a.p}">${a.btn||'Go'}</button>`:''}</div>`).join(''):'<div class="empty">All clear ✨</div>'};

W.schedule=()=>{
  const sc=schedule(),ds=Object.keys(sc.days).sort().slice(0,7);
  if(!ds.length)return'<div class="empty">No pending lectures – every class is complete 🎉</div>';
  return ds.map(d=>{const L=sc.days[d],bs={};L.forEach(l=>bs[l.sid]=(bs[l.sid]||0)+l.m);
   return`<div class="row sb" style="padding:7px 0;border-bottom:1px dashed var(--line)"><span style="width:84px"><b>${d===todayStr()?'Today':fdw(d)}</b></span><span class="grow row" style="gap:5px">${Object.keys(bs).map(k=>`<span class="chip" style="background:color-mix(in srgb,${scol(k)} 16%,transparent);color:${scol(k)}">${esc(sjn(k))} ${fh(bs[k])}</span>`).join('')}</span><span class="muted sm">${L.length} lec · ${fh(sumM(L))}</span></div>`}).join('')+
   `<div class="row sb sm mt"><span class="muted">All pending lectures finish by <b>${fd(sc.last)}</b></span><button class="btn sm" data-a="lecview" data-p="date">Full schedule</button></div>`};

W.milestones=()=>S.milestones.length?S.milestones.slice().sort((a,b)=>a.date<b.date?-1:1).map(m=>{const p=msProgress(m),d=diffDays(todayStr(),m.date);
  return`<div><div class="row sb"><b class="sm">${esc(m.name)}</b>${p>=100?chip('Done','good'):d<0?chip('overdue','bad'):chip(d+'d','')}</div>${bar(p,p>=100?'var(--good)':'')}<div class="muted sm">${p}% · ${fd(m.date)}</div></div>`}).join(''):'<div class="empty">No milestones yet – add them in Customize.</div>';

W.exams=()=>S.exams.length?S.exams.slice().sort((a,b)=>a.date<b.date?-1:1).map(e=>{const d=diffDays(todayStr(),e.date);return`<div class="row sb"><div><b>${esc(e.name)}</b><div class="muted sm">${fd(e.date)}${e.note?' · '+esc(e.note):''}</div></div><div style="text-align:right"><div class="big" style="font-size:26px">${d<0?'–':d}</div><div class="muted sm">${d<0?'passed':'days · '+Math.round(d/7)+' wks'}</div></div></div>`}).join(''):'<div class="empty">Add an exam date in Customize.</div>';

W.revision=()=>{
  const q=revQueue().slice(0,6);
  return q.length?q.map(x=>`<div class="row sb" style="padding:5px 0;border-bottom:1px dashed var(--line)"><span class="grow sm">${sdot(x.s.id)} <b>${esc(x.s.name)}</b> ${esc(x.c.name).slice(0,46)}<br><span class="muted">${x.k===0?'first revision':'revision '+(x.k+1)} · ${x.overdue?'due now':'due '+fd(x.due,1)}</span></span><button class="btn sm" data-a="addrev" data-p="${x.s.id}|${x.c.id}">+ Today</button></div>`).join('')+`<button class="btn sm mt" data-a="tab" data-p="chapters">All chapters</button>`:'<div class="empty">Finish lectures to build your revision queue.</div>'};

W.missed=()=>{const L=missedLecs();return L.length?L.slice(0,6).map(l=>`<div class="row sb" style="padding:5px 0;border-bottom:1px dashed var(--line)"><span class="grow sm">${sdot(l.sid)} <b>${esc(sjn(l.sid))}</b> ${esc(l.l)} · ${esc(l.t).slice(0,36)}<br><span class="muted">${fh(l.m)}</span></span><button class="btn sm" data-a="lecdone" data-p="${l.id}">Watched</button><button class="btn sm ghost" data-a="unmiss" data-p="${l.id}" title="Clear flag">${icon('x')}</button></div>`).join('')+(L.length>6?`<div class="muted sm mt">+${L.length-6} more</div>`:''):'<div class="empty">No missed lectures 👌</div>'};

W.pomodoro=()=>{
  const P=S.pomo,run=P.run;
  return`<div class="muted sm" id="pomo-mode" style="text-align:center">${run?(run.mode==='work'?'Focus':'Break'):'Ready'}</div><div class="pomo-time" id="pomo-time">${fmtClock(run?remainingMs():P.work*60000)}</div>
  <div class="row" style="justify-content:center"><button class="btn pri" data-a="pomo-toggle">${icon(run&&!run.paused?'pause':'play')}${run?(run.paused?'Resume':'Pause'):'Start'}</button><button class="btn" data-a="pomo-reset">Reset</button><button class="btn ghost" data-a="pomo-skip" title="Skip">${icon('right')}</button></div>
  <div class="row"><select data-c="pomo-subj" class="grow"><option value="">Subject (optional)</option>${S.subjects.map(s=>`<option value="${s.id}" ${P.subj===s.id?'selected':''}>${esc(s.name)}</option>`).join('')}</select></div>
  <div class="row sm muted" style="justify-content:center">${[[25,5],[50,10],[90,15]].map(([w,b])=>`<button class="btn sm ${P.work===w?'pri':''}" data-a="pomo-preset" data-p="${w}|${b}">${w}/${b}</button>`).join('')}<span>${P.done} session${P.done===1?'':'s'} today</span></div>`};

W.heatmap=()=>{
  const weeks=26,T=todayStr(),start=addDays(T,-((pd(T).getDay()+6)%7)-7*(weeks-1));let cells='',months='';
  for(let w=0;w<weeks;w++)for(let d=0;d<7;d++){const day=addDays(start,w*7+d);if(day>T)continue;const m=minOn(day),g=goalMin(day);const lv=m<=0?0:m<g*.4?1:m<g*.8?2:m<g?3:4;
    cells+=`<rect x="${w*15+24}" y="${d*15+14}" width="12" height="12" fill="${lv?`color-mix(in srgb,var(--accent) ${[0,25,50,75,100][lv]}%,var(--surface2))`:'var(--surface2)'}"><title>${fd(day)} · ${fh(m)}</title></rect>`;
    if(d===0&&pd(day).getDate()<=7)months+=`<text x="${w*15+24}" y="9">${pd(day).toLocaleDateString('en-IN',{month:'short'})}</text>`}
  const lbl=['Mon','Wed','Fri'].map((n,i)=>`<text x="0" y="${i*30+25}">${n}</text>`).join('');
  const by=dayIdx(),days30=Array.from({length:30},(_,i)=>addDays(T,-i)),met=days30.filter(d=>minOn(d)>=goalMin(d)).length,avg=days30.reduce((a,d)=>a+minOn(d),0)/30;
  return`<div class="row" style="gap:18px"><div class="stat"><b class="big" style="font-size:22px">${streak()}</b><span class="muted sm">current streak</span></div><div class="stat"><b class="big" style="font-size:22px">${longestStreak()}</b><span class="muted sm">longest</span></div><div class="stat"><b class="big" style="font-size:22px">${met}/30</b><span class="muted sm">days goal met</span></div><div class="stat"><b class="big" style="font-size:22px">${fhd(avg)}</b><span class="muted sm">avg / day</span></div></div>
  <div class="scroll"><svg class="heat" viewBox="0 0 ${weeks*15+26} 122" width="100%" style="min-width:520px">${months}${lbl}${cells}</svg></div>`};

W.trend=()=>{
  const T=todayStr(),[w0]=weekRange(T),N=10,vals=[];for(let i=N-1;i>=0;i--){const s=addDays(w0,-7*i);vals.push({s,m:rangeMin(s,addDays(s,6)),g:goalSum(s,addDays(s,6))})}
  const mx=Math.max(...vals.map(v=>Math.max(v.m,v.g)),60);const bs=vals.map((v,i)=>{const h=v.m/mx*90;return`<rect x="${i*32+6}" y="${104-h}" width="22" height="${h}" rx="4" fill="${v.m>=v.g?'var(--good)':'var(--accent)'}"><title>${fd(v.s)} · ${fh(v.m)}</title></rect><text x="${i*32+17}" y="118" text-anchor="middle">${pd(v.s).getDate()}/${pd(v.s).getMonth()+1}</text>`}).join('');
  const by={};for(let i=0;i<30;i++){const o=dayIdx()[addDays(T,-i)];if(o)for(const k in o.s)by[k]=(by[k]||0)+o.s[k]}const tot=Object.values(by).reduce((a,b)=>a+b,0)||1;
  return`<svg viewBox="0 0 ${N*32+8} 124" width="100%"><line x1="0" x2="${N*32+8}" y1="104" y2="104" stroke="var(--line)"/>${bs}</svg><div class="muted sm">Hours per week (green = weekly goal met)</div>
  <div class="mt">${Object.keys(by).sort((a,b)=>by[b]-by[a]).map(k=>`<div class="row sb sm"><span>${sdot(k)} ${esc(sjn(k))}</span><span class="muted">${fh(by[k])} · ${Math.round(by[k]/tot*100)}%</span></div>`).join('')||'<span class="muted sm">No study logged in the last 30 days.</span>'}</div>`};

W.reminders=()=>{const L=S.reminders.filter(r=>!r.done).sort((a,b)=>(a.at||'z')<(b.at||'z')?-1:1);
  return`<div>${L.slice(0,6).map(r=>`<div class="row sb" style="padding:5px 0;border-bottom:1px dashed var(--line)"><span class="sm grow">${esc(r.text)}<br><span class="muted">${r.at?new Date(r.at).toLocaleString('en-IN',{day:'numeric',month:'short',hour:'2-digit',minute:'2-digit'}):'no time'}</span></span><button class="btn sm" data-a="rem-done" data-p="${r.id}">Done</button></div>`).join('')||'<div class="empty">No reminders</div>'}</div>
  <div class="row"><input id="rem-text" class="grow" placeholder="Remind me to…"><input id="rem-at" type="datetime-local" style="width:170px"><button class="btn sm pri" data-a="rem-add">Add</button></div>`};

W.notes=()=>{const L=S.notes.slice().sort((a,b)=>(b.pin?1:0)-(a.pin?1:0)||b.ts-a.ts);
  return`<div class="row"><input id="note-text" class="grow" placeholder="Quick note…"><button class="btn sm pri" data-a="note-add">Add</button></div>
  ${L.slice(0,6).map(n=>`<div class="row sb sm" style="padding:6px 0;border-bottom:1px dashed var(--line)"><span class="grow">${n.pin?'📌 ':''}${esc(n.text)}</span><button class="btn sm ghost" data-a="note-pin" data-p="${n.id}">${n.pin?'Unpin':'Pin'}</button><button class="btn sm ghost" data-a="note-del" data-p="${n.id}">${icon('x')}</button></div>`).join('')||'<div class="empty">No notes</div>'}`};

W.tests=()=>{const L=S.tests.slice().sort((a,b)=>b.date.localeCompare(a.date)).slice(0,5);const avg=S.tests.length?Math.round(S.tests.reduce((a,t)=>a+pct(t.score,t.max),0)/S.tests.length):null;
  const P=S.practice,a=P.reduce((x,p)=>x+p.att,0),c=P.reduce((x,p)=>x+p.cor,0);
  return`<div class="row" style="gap:18px"><div class="stat"><b class="big" style="font-size:22px">${avg===null?'–':avg+'%'}</b><span class="muted sm">avg test score</span></div><div class="stat"><b class="big" style="font-size:22px">${a?pct(c,a)+'%':'–'}</b><span class="muted sm">practice accuracy</span></div></div>
  ${L.map(t=>`<div class="row sb sm"><span>${sdot(t.subj)} ${esc(t.name||t.type)} <span class="muted">${fd(t.date,1)}</span></span><b>${pct(t.score,t.max)}%</b></div>`).join('')||'<div class="empty">No tests logged</div>'}<button class="btn sm" data-a="tab" data-p="tests">Open tests</button>`};

W.downloads=()=>{
  const rows=S.subjects.filter(s=>!s.hide).map(s=>{const L=LL(s.id).filter(l=>!l.d&&l.m>0&&l.dl!==null&&l.dl!==undefined);return{s,got:sumM(L,l=>l.dl),need:sumM(L,l=>!l.dl),nl:L.filter(l=>!l.dl)}}).filter(r=>r.got+r.need>0);
  if(!rows.length)return'<div class="empty">No download status in your sheets.</div>';
  const sc=schedule();const order=Object.keys(sc.days).sort().flatMap(d=>sc.days[d]).filter(l=>l.dl===false).slice(0,5);
  return rows.map(r=>`<div class="row sb sm">${`<span>${sdot(r.s.id)} <b>${esc(r.s.name)}</b></span>`}<span class="muted">⬇ ${fh(r.got)} ready · ${fh(r.need)} to download</span></div>`).join('')+
   (order.length?`<div class="muted sm mt">Download next (in study order):</div>`+order.map(l=>`<div class="row sb sm"><span>${sdot(l.sid)} ${esc(l.l)} · ${esc(l.t).slice(0,32)}</span><button class="btn sm" data-a="dl-mark" data-p="${l.id}">Downloaded</button></div>`).join(''):'<div class="muted sm mt">Everything planned is downloaded ✓</div>')};

W.checklists=()=>S.checklists.length?S.checklists.map(c=>{const d=c.items.filter(i=>i.done).length;return`<div style="cursor:pointer" data-a="open-chk" data-p="${c.id}"><div class="row sb"><b class="sm">${esc(c.name)}</b><span class="muted sm">${d}/${c.items.length}</span></div>${bar(pct(d,c.items.length))}</div>`}).join(''):'<div class="empty">Import a checklist sheet (e.g. syllabus changes) in Data &amp; sync.</div>';

function widgetHTML(w,i,n){
  const [title]=WDEF[w.id];
  return`<section class="w" data-size="${w.size}" data-wid="${w.id}" ${UI.custom?'draggable="true"':''}><header><h3>${title}</h3><span class="ctl"><span class="btn ghost sm iconbtn" style="cursor:grab" title="Drag to reorder">${icon('grip')}</span><button class="btn ghost sm" data-a="w-move" data-p="${w.id}|-1" ${i===0?'disabled':''}>${icon('left')}</button><button class="btn ghost sm" data-a="w-move" data-p="${w.id}|1" ${i===n-1?'disabled':''}>${icon('right')}</button><button class="btn ghost sm" data-a="w-size" data-p="${w.id}" title="Change width">${['S','M','L'][w.size-1]}</button><button class="btn ghost sm" data-a="w-hide" data-p="${w.id}" title="Hide">${icon('eye')}</button></span></header>${W[w.id]()}</section>`;
}

/* ===================== VIEWS ===================== */
const UI={tab:'dashboard',custom:false,lec:{view:'chapters',sid:'all',status:'pending',pri:'',q:'',from:'',to:'',sort:'sheet',page:0},plDate:todayStr(),revSid:'',chk:'',chkG:'all',chkQ:'',open:new Set(),imp:null,tq:'',dpage:0};
const TABS=[['dashboard','Dashboard','grid'],['lectures','Lectures','book'],['planner','Planner','cal'],['chapters','Chapters','layers'],['analytics','Analytics','chart'],['tests','Practice & tests','target'],['checklists','Checklists','check'],['data','Data & sync','sync'],['settings','Customize','sliders']];
const subjOpts=(sel,all)=>(all?`<option value="">${all}</option>`:'')+S.subjects.map(s=>`<option value="${s.id}" ${sel===s.id?'selected':''}>${esc(s.name)}</option>`).join('');
const subjPills=(cur,act,all)=>`<div class="pills">${all?`<button class="${cur===''||cur==='all'?'on':''}" data-a="${act}" data-p="${all}">All</button>`:''}${S.subjects.filter(s=>!s.hide).map(s=>`<button class="${cur===s.id?'on':''}" data-a="${act}" data-p="${s.id}">${esc(s.name)}</button>`).join('')}</div>`;

/* ---------- dashboard ---------- */
function vDashboard(){
  const L=S.layout.widgets.filter(w=>!w.hidden&&W[w.id]),hidden=S.layout.widgets.filter(w=>w.hidden||!W[w.id]);
  const missing=Object.keys(WDEF).filter(k=>!S.layout.widgets.some(w=>w.id===k));
  let h='';
  if(UI.custom)h+=`<div class="card"><div class="row sb"><div><b>Customizing your dashboard</b><div class="muted sm">Drag cards by the handle (or use ◀ ▶), change width with S/M/L, hide what you don't need.</div></div><div class="row"><select data-c="layout-preset"><option value="">Layout preset…</option>${Object.keys(LAYOUTS).map(k=>`<option ${S.layout.preset===k?'selected':''}>${k}</option>`).join('')}</select><button class="btn pri" data-a="customize">Done</button></div></div>
   ${(hidden.length||missing.length)?`<div class="row mt"><span class="muted sm">Add:</span>${hidden.map(w=>`<button class="btn sm" data-a="w-show" data-p="${w.id}">${icon('plus')}${WDEF[w.id][0]}</button>`).join('')}${missing.map(k=>`<button class="btn sm" data-a="w-add" data-p="${k}">${icon('plus')}${WDEF[k][0]}</button>`).join('')}</div>`:''}</div>`;
  const pend=pendingSubjects();
  h+=`<div class="${UI.custom?'custom ':''}grid12" id="grid">${L.map((w,i)=>widgetHTML(w,i,L.length)).join('')}</div>`;
  if(!S.subjects.length)h=`<div class="card" style="text-align:center;padding:40px"><h2 style="font-size:20px;color:var(--text);text-transform:none;letter-spacing:0">Welcome 👋</h2><p class="muted">Start by importing your lecture index Excel files – one per class. Every row will be read and tracked.</p><button class="btn pri" data-a="tab" data-p="data">Import Excel files</button></div>`+h;
  return h;
}

/* ---------- lectures ---------- */
function lecDate(l){return l.doneOn||l.sdate||schedule().planned[l.id]||''}
function lecMatch(l){
  const f=UI.lec,T=todayStr();
  if(f.sid!=='all'&&l.sid!==f.sid)return false;
  const pend=!l.d&&!l.arch;
  switch(f.status){case'pending':if(!pend)return false;break;case'done':if(!l.d||l.arch)return false;break;case'missed':if(!isMissed(l))return false;break;
    case'due':if(!pend||!((schedule().planned[l.id]||'9')<=T))return false;break;case'upcoming':if(!pend||(schedule().planned[l.id]||'0')<=T)return false;break;
    case'archived':if(!l.arch)return false;break;default:if(l.arch)return false}
  if(f.pri==='high'&&l.pri!=='high')return false;if(f.pri==='low'&&l.pri!=='low')return false;if(f.pri==='normal'&&l.pri)return false;
  const d=lecDate(l);if(f.from&&(!d||d<f.from))return false;if(f.to&&(!d||d>f.to))return false;
  if(f.q){const hay=(l.t+' '+l.l+' '+l.n+' '+(l.note||'')+' '+Object.values(l.x||{}).join(' ')+' '+sjn(l.sid)+' '+((S.chapters[l.sid]||[]).find(c=>c.id===l.ch)||{name:''}).name).toLowerCase();
    if(!f.q.toLowerCase().split(/\s+/).every(w=>hay.includes(w)))return false}
  return true;
}
const filterActive=()=>{const f=UI.lec;return f.q||f.pri||f.from||f.to};
function lecStatusChip(l){const T=todayStr();if(l.gone)return chip('not in sheet','warn');if(l.arch)return chip('archived');if(l.d)return chip('Watched','good');if(isMissed(l))return chip('Missed','bad');const p=schedule().planned[l.id];return p?(p<=T?chip('Due','acc'):chip('Upcoming')):chip('Pending')}
function lecRow(l,cols){
  const ch=(S.chapters[l.sid]||[]).find(c=>c.id===l.ch),p=schedule().planned[l.id],date=l.d?l.doneOn:(p||l.sdate);
  const x=Object.entries(l.x||{}).map(([k,v])=>`${k}: ${v}`).join(' · ');
  return`<tr class="${l.d?'done':''} ${l.gone?'gone':''}"><td class="c"><input type="checkbox" data-c="lecdone" data-p="${l.id}" ${l.d?'checked':''}></td>
  ${cols.subj?`<td>${subjChip(l.sid)}</td>`:''}<td class="sm muted" style="white-space:nowrap">${esc(l.l)}${l.n?`<br>${esc(l.n)}`:''}</td>
  <td class="t">${esc(l.t)||'<span class="muted">(no topic yet)</span>'} ${l.pri==='high'?chip('⚑ high','bad'):l.pri==='low'?chip('low'):''}${l.tag?' '+chip(esc(l.tag),'warn'):''}${l.note?` <span class="muted sm">· ${esc(l.note)}</span>`:''}${x?`<div class="muted sm">${esc(x)}${l.row?` · sheet row ${l.row}`:''}</div>`:l.row?`<div class="muted sm">sheet row ${l.row}</div>`:''}</td>
  ${cols.ch?`<td><select data-c="lec-ch" data-p="${l.id}" style="max-width:170px;font-size:12.5px">${(S.chapters[l.sid]||[]).map(c=>`<option value="${c.id}" ${c.id===l.ch?'selected':''}>${esc(c.name).slice(0,40)}</option>`).join('')}</select></td>`:''}
  <td class="c" data-a="editdur" data-p="${l.id}" style="cursor:pointer;white-space:nowrap" title="Click to edit length">${l.m?fh(l.m):'<span class="chip warn">—</span>'}</td>
  ${cols.dl?`<td class="c">${l.dl===null||l.dl===undefined?'':`<input type="checkbox" data-c="lecdl" data-p="${l.id}" ${l.dl?'checked':''}>`}</td>`:''}
  <td class="sm" style="white-space:nowrap">${lecStatusChip(l)}${date?`<div class="muted sm">${l.d?'watched ':''}${fd(date,1)}</div>`:''}</td>
  <td><details class="menu"><summary>${icon('more')}</summary><div class="pop"><button data-a="lec-pri" data-p="${l.id}">Priority: ${l.pri||'normal'} → next</button><button data-a="lec-miss" data-p="${l.id}">${l.missed?'Clear missed flag':'Mark as missed'}</button><button data-a="lec-note" data-p="${l.id}">Add / edit note</button>${l.d?`<button data-a="lec-arch" data-p="${l.id}">${l.arch?'Restore from archive':'Archive'}</button>`:''}<button data-a="lec-del" data-p="${l.id}" class="danger" style="color:var(--bad)">Remove lecture</button></div></details></td></tr>`;
}
function lecTable(L,cols){return`<div class="scroll"><table><tr><th class="c">✓</th>${cols.subj?'<th>Class</th>':''}<th>Day / lecture</th><th>Topic</th>${cols.ch?'<th>Chapter</th>':''}<th class="c">Length</th>${cols.dl?'<th class="c">⬇</th>':''}<th>Status</th><th></th></tr>${L.map(l=>lecRow(l,cols)).join('')||'<tr><td colspan="9" class="empty">No lectures match these filters.</td></tr>'}</table></div>`}
function chapterBody(sid,cid){
  const L=LL(sid).filter(l=>l.ch===cid&&lecMatch(l));const anyDl=LL(sid).some(l=>l.dl!==null&&l.dl!==undefined);
  return lecTable(L,{subj:false,ch:false,dl:anyDl});
}
function lecSort(L){const s=UI.lec.sort,pr={high:0,'':1,low:2};
  const f={sheet:null,date:(a,b)=>(lecDate(a)||'9').localeCompare(lecDate(b)||'9'),subject:(a,b)=>sjn(a.sid).localeCompare(sjn(b.sid)),topic:(a,b)=>a.t.localeCompare(b.t),duration:(a,b)=>b.m-a.m,
    urgency:(a,b)=>(isMissed(b)-isMissed(a))||(pr[a.pri||'']-pr[b.pri||''])||((schedule().planned[a.id]||'9').localeCompare(schedule().planned[b.id]||'9'))}[s];
  return f?L.slice().sort(f):L}
function vLectures(){
  const f=UI.lec;let h='';
  h+=`<div class="kpis">${S.subjects.filter(s=>!s.hide&&hasLectures(s.id)).map(s=>{const t=totalM(s.id),d=doneM(s.id),n=LL(s.id).length,dn=LL(s.id).filter(l=>l.d).length;return`<div class="card" style="cursor:pointer;${f.sid===s.id?'outline:2px solid var(--accent)':''}" data-a="lec-subj" data-p="${s.id}"><div class="row sb"><b>${sdot(s.id)} ${esc(s.name)}</b><span class="chip acc">${pct(d,t)}%</span></div>${bar(pct(d,t),s.color)}<div class="muted sm mt">${dn}/${n} lectures · ${fh(t-d)} left</div></div>`}).join('')}</div>`;
  h+=`<div class="card"><div class="row sb"><div class="seg">${[['chapters','By chapter'],['table','Table'],['date','By date'],['archive','Archive']].map(([k,l])=>`<button class="${f.view===k?'on':''}" data-a="lecview" data-p="${k}">${l}</button>`).join('')}</div>
   <div class="row"><button class="btn" data-a="quick" data-p="lecture">${icon('plus')}Add lecture</button><button class="btn" data-a="arch-all">Archive completed</button></div></div>
   <div class="row mt"><select data-c="lf" data-p="sid"><option value="all">All classes</option>${S.subjects.filter(s=>hasLectures(s.id)).map(s=>`<option value="${s.id}" ${f.sid===s.id?'selected':''}>${esc(s.name)}</option>`).join('')}</select>
   <select data-c="lf" data-p="status">${[['pending','Pending'],['due','Due now'],['upcoming','Upcoming'],['missed','Missed'],['done','Completed'],['all','All'],['archived','Archived']].map(([k,l])=>`<option value="${k}" ${f.status===k?'selected':''}>${l}</option>`).join('')}</select>
   <select data-c="lf" data-p="pri"><option value="">Any priority</option><option value="high" ${f.pri==='high'?'selected':''}>High</option><option value="normal" ${f.pri==='normal'?'selected':''}>Normal</option><option value="low" ${f.pri==='low'?'selected':''}>Low</option></select>
   <label class="row sm muted">From <input type="date" data-c="lf" data-p="from" value="${f.from}"></label><label class="row sm muted">To <input type="date" data-c="lf" data-p="to" value="${f.to}"></label>
   <select data-c="lf" data-p="sort">${[['sheet','Sheet order'],['date','Date'],['subject','Class'],['urgency','Urgency'],['duration','Longest first'],['topic','Topic A–Z']].map(([k,l])=>`<option value="${k}" ${f.sort===k?'selected':''}>Sort: ${l}</option>`).join('')}</select>
   <div class="search grow" style="flex-basis:200px"><span>${icon('search')}</span><input data-i="lsearch" placeholder="Search topic, lecture, chapter, notes…" value="${esc(f.q)}"></div></div></div><div id="lec-results">${lecResults()}</div>`;
  return h;
}
function lecResults(){
  const f=UI.lec,all=Object.values(S.lectures).flat().filter(l=>lecMatch(l)),n=all.length;
  const count=`<div class="muted sm" style="margin:0 2px 10px">${n} lecture${n===1?'':'s'} · ${fh(sumM(all))}</div>`;
  if(f.view==='table'||f.view==='archive'){const PS=100,L=lecSort(f.view==='archive'?Object.values(S.lectures).flat().filter(l=>l.arch&&(f.sid==='all'||l.sid===f.sid)):all),pages=Math.max(1,Math.ceil(L.length/PS));f.page=Math.min(f.page,pages-1);
    return count+`<div class="card" style="padding:6px">${lecTable(L.slice(f.page*PS,(f.page+1)*PS),{subj:true,ch:true,dl:L.some(l=>l.dl!==null&&l.dl!==undefined)})}</div>`+(pages>1?`<div class="row" style="justify-content:center"><button class="btn sm" data-a="lpage" data-p="-1">Prev</button><span class="muted sm">Page ${f.page+1} of ${pages}</span><button class="btn sm" data-a="lpage" data-p="1">Next</button></div>`:'')}
  if(f.view==='date'){
    const g={};all.forEach(l=>{const d=lecDate(l)||'Undated';(g[d]=g[d]||[]).push(l)});
    const keys=Object.keys(g).sort((a,b)=>a==='Undated'?1:b==='Undated'?-1:a<b?-1:1),PS=20,pages=Math.max(1,Math.ceil(keys.length/PS));f.page=Math.min(f.page,pages-1);
    return count+keys.slice(f.page*PS,(f.page+1)*PS).map(d=>{const L=g[d];return`<div class="card"><div class="row sb"><b>${d==='Undated'?'Undated (no date in sheet, not scheduled)':fdw(d)+(d===todayStr()?' · today':'')}</b><span class="muted sm">${L.length} lecture${L.length>1?'s':''} · ${fh(sumM(L))}</span></div><div style="margin-top:6px">${lecTable(L,{subj:true,ch:false,dl:L.some(l=>l.dl!==null&&l.dl!==undefined)})}</div></div>`}).join('')+(pages>1?`<div class="row" style="justify-content:center"><button class="btn sm" data-a="lpage" data-p="-1">Prev</button><span class="muted sm">Days ${f.page*PS+1}–${Math.min(keys.length,(f.page+1)*PS)} of ${keys.length}</span><button class="btn sm" data-a="lpage" data-p="1">Next</button></div>`:'')}
  /* chapters */
  let h=count;
  S.subjects.filter(s=>(f.sid==='all'||f.sid===s.id)&&hasLectures(s.id)).forEach(s=>{
    const chs=S.chapters[s.id];let sec='';let opened=false;
    chs.forEach(c=>{const cl=LL(s.id).filter(l=>l.ch===c.id);const m=cl.filter(lecMatch);if(!m.length&&(filterActive()||f.status!=='all'&&f.status!=='pending'))return;if(!m.length&&f.status==='pending'&&cl.length)return;
      const cd=sumM(cl,l=>l.d),ct=sumM(cl),key=s.id+c.id;const open=UI.open.has(key)||(!opened&&!UI.open.size&&cl.some(l=>!l.d)&&(opened=true));
      sec+=`<details class="ch" data-key="${key}" data-sid="${s.id}" data-cid="${c.id}" ${open?'open':''}><summary><div class="row sb"><span><b>${esc(c.name)}</b></span><span class="muted sm">${cl.filter(l=>l.d).length}/${cl.length} · ${fh(cd)} / ${fh(ct)}</span></div>${bar(pct(cd,ct),s.color)}</summary><div class="in">${open?chapterBody(s.id,c.id):''}</div></details>`});
    if(sec)h+=`<div class="row" style="margin:16px 2px 8px"><h3 style="font-size:15px;color:var(--text)">${sdot(s.id)} ${esc(s.full||s.name)}</h3><span class="muted sm">${esc(s.fac)}</span></div>`+sec});
  return h+(h===count?'<div class="card empty">No lectures match. Try “All” status or clear the search.</div>':'');
}

/* ---------- planner ---------- */
function vPlanner(){
  const d=UI.plDate,T=todayStr(),L=S.tasks.filter(t=>t.date===d).sort((a,b)=>(a.start||'99')<(b.start||'99')?-1:a.start===b.start?0:1);
  const real=L.filter(t=>t.type!=='break'),plan=real.reduce((a,t)=>a+(+t.plan||0),0),act=minOn(d),gm=goalMin(d);
  const fx=fixedFor(d);const [w0]=weekRange(d);
  const backlog=S.tasks.filter(t=>t.date<T&&!t.done&&t.type!=='break');
  let h=`<div class="card"><div class="row sb"><div class="row"><button class="btn iconbtn" data-a="pl-nav" data-p="-1">${icon('left')}</button><input type="date" data-c="pl-date" value="${d}"><button class="btn iconbtn" data-a="pl-nav" data-p="1">${icon('right')}</button><button class="btn" data-a="pl-nav" data-p="0">Today</button><b>${fdw(d)}</b>${pd(d).getDay()===S.settings.lightDay?chip('light day','acc'):''}</div>
   <div class="row"><button class="btn pri" data-a="autoplan" data-p="${d}">${icon('bolt')}Auto-plan this day</button><button class="btn" data-a="autoplan-week">Plan next 7 days</button><button class="btn danger" data-a="clear-auto" data-p="${d}">Clear auto-plan</button></div></div>
   <div class="row mt sm"><span class="muted">Studied <b>${fh(act)}</b> · planned <b>${fh(plan)}</b> · goal <b>${fhd(gm)}</b></span></div>${bar(pct(act,gm),act>=gm?'var(--good)':'',  'lg')}</div>`;
  h+=`<div class="card"><div class="row" style="gap:8px">${Array.from({length:7},(_,i)=>{const x=addDays(w0,i),a=minOn(x),p=S.tasks.filter(t=>t.date===x&&t.type!=='break').reduce((s,t)=>s+(+t.plan||0),0);return`<button class="btn grow" style="flex-direction:column;gap:2px;${x===d?'border-color:var(--accent);background:var(--accent-soft)':''}" data-a="pl-day" data-p="${x}"><b class="sm">${pd(x).toLocaleDateString('en-IN',{weekday:'short'})} ${pd(x).getDate()}</b><span class="muted sm">${fhd(a)} / ${fhd(Math.max(p,gm))}</span>${bar(pct(a,gm),x===T?'':'')}</button>`}).join('')}</div></div>`;
  if(backlog.length)h+=`<div class="card"><div class="row sb"><div><b>Backlog</b> <span class="muted sm">${backlog.length} unfinished task${backlog.length>1?'s':''} from earlier days</span></div><button class="btn pri" data-a="move-backlog">Move all to ${d===T?'today':fd(d,1)}</button></div></div>`;
  h+=`<div class="card"><h2>Schedule</h2>`;
  if(!L.length&&!fx.length)h+=`<div class="empty">No tasks yet. Auto-plan builds a full day from your pending lectures, revisions and practice, or add your own below.</div>`;
  else{
    const rows=L.map(t=>({k:t.start||'99',h:plTask(t)})).concat(fx.map(f=>({k:mtime(f.s),h:`<div class="task brk" style="padding:7px 0;border-bottom:1px dashed var(--line)"><span style="width:17px"></span><span class="muted sm" style="width:84px">${mtime(f.s)}–${mtime(f.e)}</span><span class="grow tx">📌 ${esc(f.name)} (fixed)</span></div>`}))).sort((a,b)=>a.k<b.k?-1:1);
    h+=rows.map(r=>r.h).join('')}
  h+=`<div class="row mt2"><select id="nt-subj">${subjOpts('','General')}</select><select id="nt-type">${['custom','lecture','revision','practice','test','recap'].map(k=>`<option value="${k}">${typeLabel[k]}</option>`).join('')}</select><input id="nt-text" class="grow" placeholder="Task…" style="min-width:160px"><input id="nt-start" type="time" style="width:110px"><input id="nt-min" type="number" min="5" step="5" placeholder="min" style="width:76px"><button class="btn pri" data-a="task-add">Add</button></div>
   <div class="muted sm mt">Auto-planner settings (day start, lunch, session length, lecture hours, fixed commitments) live in Customize → Study plan. Lock 🔒 a task to keep it when you re-plan.</div></div>`;
  return h;
}
function plTask(t){
  const brk=t.type==='break';
  return`<div class="task ${t.done?'done':''} ${brk?'brk':''}" style="padding:8px 0;border-bottom:1px dashed var(--line)"><div class="tag" style="background:${typeCol[t.type]||'#94a3b8'}"></div>
   ${brk?'<span style="width:17px"></span>':`<input type="checkbox" data-c="taskdone" data-p="${t.id}" ${t.done?'checked':''}>`}
   <input type="time" data-c="task-f" data-p="${t.id}|start" value="${t.start||''}" style="width:104px;padding:4px 6px"><span class="muted sm">${t.end?'→ '+t.end:''}</span>
   <span class="grow tx" style="min-width:160px">${t.subj&&!brk?subjChip(t.subj)+' ':''}${brk?'':`<span class="chip">${typeLabel[t.type]||t.type}</span> `}${esc(t.text)}${t.auto?' <span class="muted sm">· auto</span>':''}</span>
   ${brk?'':`<label class="muted sm">plan <input type="number" data-c="task-f" data-p="${t.id}|plan" value="${t.plan}" style="width:64px;padding:4px 6px"> min</label><label class="muted sm">done <input type="number" data-c="task-f" data-p="${t.id}|act" value="${t.act||''}" style="width:64px;padding:4px 6px"> min</label>`}
   <span class="row" style="gap:0">${brk?'':`<button class="btn ghost sm" data-a="task-lock" data-p="${t.id}" title="Keep when re-planning" style="${t.lock?'color:var(--accent)':''}">${icon('lock')}</button><button class="btn ghost sm" data-a="task-edit" data-p="${t.id}" title="Edit text">${icon('edit')}</button><button class="btn ghost sm" data-a="task-next" data-p="${t.id}" title="Move to next day">${icon('right')}</button>`}<button class="btn ghost sm" data-a="task-del" data-p="${t.id}">${icon('x')}</button></span></div>`}

/* ---------- chapters / revision ---------- */
function vChapters(){
  const sid=UI.revSid||(S.subjects.find(s=>!s.hide)||{}).id;if(!sid)return'<div class="card empty">No classes yet.</div>';
  const su=sj(sid),C=S.chapters[sid]||[];
  let h=subjPills(sid,'rev-subj')+`<div class="card"><div class="row sb"><div><b>${esc(su.full||su.name)}</b> <span class="muted sm">${esc(su.fac)}</span>${su.res?`<div class="muted sm">Resources: ${esc(su.res)}</div>`:''}</div><span class="chip good">${subjScore(sid)}% exam-ready score</span></div>
   <div class="note mt">A chapter is <b>exam-ready</b> only when lectures are done <b>and</b> Revision 1 &amp; 2 are done <b>and</b> you've practised questions <b>and</b> recall is rated 4–5. Lectures alone are 25% of the score. Ticking a revision stamps “last revised” and schedules the next one.</div></div>
   <div class="card scroll"><table><tr><th>Chapter</th><th class="c">Lectures</th><th class="c">Rev 1</th><th class="c">Rev 2</th><th class="c">Final</th><th class="c">Practice</th><th class="c">Recall</th><th>Last / next</th><th>Status</th><th></th></tr>`;
  C.forEach((c,i)=>{const L=chLecs(sid,c),ld=lecDone(sid,c),st=chStatus(sid,c),nx=ld?nextRevDue(sid,c):null;
    h+=`<tr><td>${esc(c.name)}</td><td class="c">${L.length?chip(L.filter(l=>l.d).length+'/'+L.length,ld?'good':''):`<input type="checkbox" data-c="cf" data-p="${c.id}|cls" ${c.cls?'checked':''}>`}</td>
    ${['r1','r2','r3','pr'].map(k=>`<td class="c"><input type="checkbox" data-c="cf" data-p="${c.id}|${k}" ${c[k]?'checked':''}></td>`).join('')}
    <td class="c"><select data-c="conf" data-p="${c.id}">${[0,1,2,3,4,5].map(n=>`<option value="${n}" ${c.conf===n?'selected':''}>${n||'–'}</option>`).join('')}</select></td>
    <td class="sm muted" style="white-space:nowrap">${c.last?fd(c.last,1):'—'}${nx?`<br>next ${nx<=todayStr()?'<b style="color:var(--warn)">now</b>':fd(nx,1)}`:''}<br><button class="btn sm" data-a="rev-today" data-p="${c.id}">Revised today</button></td>
    <td>${chip(st[0],st[1])}</td><td><details class="menu"><summary>${icon('more')}</summary><div class="pop"><button data-a="ch-ren" data-p="${c.id}">Rename</button>${i>0?`<button data-a="ch-merge" data-p="${c.id}">Merge into previous</button>`:''}<button data-a="ch-del" data-p="${c.id}" style="color:var(--bad)">Delete chapter</button></div></details></td></tr>`});
  h+=`</table><div class="row mt"><input id="newch" class="grow" placeholder="Add chapter / unit"><button class="btn pri" data-a="ch-add" data-p="${sid}">Add</button></div></div>`;return h;
}

/* ---------- analytics ---------- */
function vAnalytics(){
  const T=todayStr(),by=dayIdx(),all=Object.values(by).reduce((a,o)=>a+o.t,0),dys=Object.keys(by).length;
  const wd=[0,0,0,0,0,0,0],wc=[0,0,0,0,0,0,0];Object.keys(by).forEach(d=>{const k=(pd(d).getDay()+6)%7;wd[k]+=by[d].t;wc[k]++});
  const mx=Math.max(...wd.map((v,i)=>wc[i]?v/wc[i]:0),60);
  const bars=wd.map((v,i)=>{const a=wc[i]?v/wc[i]:0,h=a/mx*80;return`<rect x="${i*44+8}" y="${90-h}" width="28" height="${h}" rx="5" fill="var(--accent)"><title>${fh(a)} avg</title></rect><text x="${i*44+22}" y="104" text-anchor="middle">${['Mon','Tue','Wed','Thu','Fri','Sat','Sun'][i]}</text>`}).join('');
  const best=Object.keys(by).sort((a,b)=>by[b].t-by[a].t)[0];
  const last14=Array.from({length:14},(_,i)=>addDays(T,-i)),lecPace=last14.reduce((a,d)=>a+sumM(Object.values(S.lectures).flat(),l=>l.d&&l.doneOn===d&&!l.gone),0)/14;
  return`<div class="kpis"><div class="card stat"><span class="muted sm">Total logged</span><b class="big">${fh(all)}</b></div><div class="card stat"><span class="muted sm">Active days</span><b class="big">${dys}</b></div><div class="card stat"><span class="muted sm">Best day</span><b class="big">${best?fh(by[best].t):'–'}</b><span class="muted sm">${best?fd(best,1):''}</span></div><div class="card stat"><span class="muted sm">Lecture pace (14d)</span><b class="big">${fh(lecPace)}</b><span class="muted sm">per day</span></div></div>
  <div class="grid12"><div class="w" data-size="3"><header><h3>Consistency</h3></header>${W.heatmap()}</div>
  <div class="w" data-size="2"><header><h3>Average study time by weekday</h3></header><svg viewBox="0 0 316 110" width="100%">${bars}</svg></div><div class="w" data-size="1"><header><h3>Weekly trend</h3></header>${W.trend()}</div>
  <div class="w" data-size="3"><header><h3>Class overview</h3></header><div class="scroll"><table><tr><th>Class</th><th>Lectures</th><th>Hours studied (all time)</th><th>Last 7 days</th><th>Readiness</th><th>Projected finish</th></tr>${S.subjects.filter(s=>!s.hide).map(s=>{let all=0,w=0;Object.keys(by).forEach(d=>{const v=by[d].s[s.id]||0;all+=v;if(d>addDays(T,-7))w+=v});const p=subjProjection(s.id);return`<tr><td>${sdot(s.id)} <b>${esc(s.name)}</b></td><td>${hasLectures(s.id)?pct(doneM(s.id),totalM(s.id))+'%':'–'}</td><td>${fh(all)}</td><td>${fh(w)}</td><td>${subjScore(s.id)}%</td><td>${p.rem>0&&p.fin?fd(p.fin):hasLectures(s.id)?'Complete':'–'}</td></tr>`}).join('')}</table></div></div></div>`;
}

/* ---------- practice & tests ---------- */
function vTests(){
  const tl=S.tests.slice().sort((a,b)=>a.date.localeCompare(b.date));let chart='';
  if(tl.length>1){const Wd=560,H=150,n=tl.length,pts=tl.map((t,i)=>[20+i*(Wd-40)/(n-1),H-20-pct(t.score,t.max)/100*(H-40),pct(t.score,t.max)]),py=H-20-.5*(H-40);
    chart=`<svg viewBox="0 0 ${Wd} ${H}" width="100%"><line x1="20" x2="${Wd-20}" y1="${py}" y2="${py}" stroke="var(--warn)" stroke-dasharray="4 4"/><text x="22" y="${py-3}">50% pass mark</text><polyline fill="none" stroke="var(--accent)" stroke-width="2" points="${pts.map(p=>p[0]+','+p[1]).join(' ')}"/>${pts.map(p=>`<circle cx="${p[0]}" cy="${p[1]}" r="4" fill="var(--accent)"/><text x="${p[0]}" y="${p[1]-8}" text-anchor="middle">${p[2]}%</text>`).join('')}</svg>`}
  const acc=S.subjects.map(s=>{const P=S.practice.filter(p=>p.subj===s.id),a=P.reduce((x,p)=>x+p.att,0),c=P.reduce((x,p)=>x+p.cor,0);return{s,a,c}}).filter(x=>x.a);
  return`<div class="card"><h2>Log question practice</h2><div class="row mt"><input type="date" id="pdt" value="${todayStr()}"><select id="psub">${subjOpts('')}</select><input id="pnote" class="grow" placeholder="Chapter / source (optional)"><input id="patt" type="number" min="0" placeholder="Attempted" style="width:100px"><input id="pcor" type="number" min="0" placeholder="Correct" style="width:90px"><button class="btn pri" data-a="prac-add">Add</button></div>
   ${acc.map(x=>`<div class="mt"><div class="row sb sm"><span>${sdot(x.s.id)} ${esc(x.s.name)}</span><span class="muted">${x.c}/${x.a} · <b>${pct(x.c,x.a)}%</b></span></div>${bar(pct(x.c,x.a),x.s.color)}</div>`).join('')}
   ${S.practice.length?`<table class="mt">${S.practice.slice().reverse().slice(0,10).map(p=>`<tr><td>${fd(p.date)}</td><td>${esc(sjn(p.subj))} <span class="muted">${esc(p.note)}</span></td><td>${p.cor}/${p.att}</td><td><button class="btn ghost sm" data-a="prac-del" data-p="${p.id}">${icon('x')}</button></td></tr>`).join('')}</table>`:''}</div>
  <div class="card"><h2>Mock tests, MTP, RTP &amp; past papers</h2><div class="row mt"><input type="date" id="tdt" value="${todayStr()}"><select id="tsub">${subjOpts('')}</select><select id="ttype"><option>Mock</option><option>MTP</option><option>RTP</option><option>PYQ</option></select><input id="tname" class="grow" placeholder="Name (e.g. MTP 1 Series 1)"><input id="tsc" type="number" step="0.5" placeholder="Marks" style="width:80px"><input id="tmx" type="number" value="100" style="width:75px"><input id="tmin" type="number" placeholder="Minutes" style="width:85px"><button class="btn pri" data-a="test-add">Add</button></div></div>
  ${chart?`<div class="card"><h2>Score trend</h2>${chart}</div>`:''}
  <div class="card"><h2>All tests</h2>${tl.length?`<table class="mt">${tl.slice().reverse().map(t=>`<tr><td>${fd(t.date)}</td><td>${esc(t.name||t.type)} ${chip(t.type)} <span class="muted">${esc(sjn(t.subj))}</span></td><td><b>${t.score}/${t.max}</b> (${pct(t.score,t.max)}%)</td><td class="muted">${t.min?t.min+' min':''}</td><td><button class="btn ghost sm" data-a="test-del" data-p="${t.id}">${icon('x')}</button></td></tr>`).join('')}</table>`:'<div class="empty">No tests logged yet.</div>'}</div>`;
}

/* ---------- checklists ---------- */
function vChecklists(){
  if(!S.checklists.length)return`<div class="card empty">No checklists yet. Import a sheet with a “Done?” column (for example your syllabus-changes sheet) from Data &amp; sync, or create one here.<br><button class="btn pri mt" data-a="chk-new">New checklist</button></div>`;
  const c=S.checklists.find(x=>x.id===UI.chk)||S.checklists[0];UI.chk=c.id;
  const groups=[...new Set(c.items.map(i=>i.g).filter(Boolean))],q=UI.chkQ.toLowerCase();
  const L=c.items.filter(i=>(UI.chkG==='all'||(UI.chkG==='open'?!i.done:i.g===UI.chkG))&&(!q||(i.t+i.g+i.ref).toLowerCase().includes(q)));const d=c.items.filter(i=>i.done).length;
  return`<div class="pills" style="margin-bottom:12px">${S.checklists.map(x=>`<button class="${x.id===c.id?'on':''}" data-a="open-chk" data-p="${x.id}">${esc(x.name)}</button>`).join('')}<button data-a="chk-new">+ New</button></div>
  <div class="card"><div class="row sb"><div><b>${esc(c.name)}</b><div class="muted sm">${d} of ${c.items.length} done · ${pct(d,c.items.length)}%</div></div><div class="row"><button class="btn sm" data-a="chk-ren" data-p="${c.id}">Rename</button><button class="btn sm danger" data-a="chk-del" data-p="${c.id}">Delete</button></div></div>${bar(pct(d,c.items.length))}
   ${c.notes?`<div class="note mt" style="white-space:pre-line">${esc(c.notes)}</div>`:''}
   <div class="row mt"><select data-c="chk-g"><option value="all">All groups</option><option value="open" ${UI.chkG==='open'?'selected':''}>Not done yet</option>${groups.map(g=>`<option ${UI.chkG===g?'selected':''}>${esc(g)}</option>`).join('')}</select><input data-c="chk-q" class="grow" placeholder="Search (press Enter)" value="${esc(UI.chkQ)}"></div></div>
  <div class="card scroll"><table><tr><th></th><th>#</th><th>Group</th><th>Item</th><th>Ref.</th><th></th></tr>${L.map(i=>`<tr class="${i.done?'done':''}"><td><input type="checkbox" data-c="chk-item" data-p="${c.id}|${i.id}" ${i.done?'checked':''}></td><td class="muted">${esc(i.key)}</td><td class="sm">${esc(i.g)}</td><td class="t">${esc(i.t)}</td><td class="sm muted">${esc(i.ref)}</td><td><button class="btn ghost sm" data-a="chk-item-del" data-p="${c.id}|${i.id}">${icon('x')}</button></td></tr>`).join('')||'<tr><td colspan="6" class="empty">No items.</td></tr>'}</table>
  <div class="row mt"><input id="chk-new-g" placeholder="Group" style="width:140px"><input id="chk-new-t" class="grow" placeholder="Add item…"><button class="btn pri" data-a="chk-add" data-p="${c.id}">Add</button></div></div>`;
}

/* ---------- data & sync ---------- */
function vData(){
  const fs=S.sync.files;
  return`<div class="card"><h2>Import from Excel / CSV</h2><p class="muted sm" style="margin:6px 0 12px">Drop your lecture-index files here (or choose them). The reader finds the header row itself, then imports <b>every row</b> of every sheet – including blank-duration rows, placeholder days, dates, weekday/holiday notes and any extra columns – and shows a reconciliation so you can see nothing was dropped. Re-importing the same file merges: status changes made in Excel are picked up, while your priority, notes, chapter assignments and revision progress are kept.</p>
   <div class="drop" id="drop"><div style="font-size:15px;color:var(--text)"><b>Drag &amp; drop .xlsx / .csv files</b></div><div class="mt"><label class="btn pri" style="cursor:pointer">${icon('plus')}Choose files<input type="file" id="file-in" accept=".xlsx,.csv" multiple style="display:none"></label></div></div></div>
  <div class="card"><h2>Live link to your spreadsheets</h2><p class="muted sm" style="margin:6px 0 10px">${window.showOpenFilePicker?'Link the files once; the dashboard re-reads them every few seconds and merges changes as you save in Excel. (Excel → dashboard only; nothing is written to your spreadsheet.)':'<b>This browser cannot keep a live file link</b> (needs Chrome or Edge on desktop). Use “Refresh from files” – pick the updated file(s) and they merge instantly.'}</p>
   <div class="row">${window.showOpenFilePicker?`<button class="btn pri" data-a="link-files">${icon('sync')}Link Excel files</button>`:''}<button class="btn" data-a="refresh-files">Refresh from files…</button>
   ${window.showOpenFilePicker?`<label class="row sm muted">Check every <select data-c="set" data-p="settings.syncSec" data-t="n">${[5,10,30,60].map(n=>`<option value="${n}" ${S.settings.syncSec===n?'selected':''}>${n}s</option>`).join('')}</select></label>`:''}</div>
   ${fs.length?`<table class="mt">${fs.map(f=>`<tr><td><b>${esc(f.name)}</b><div class="muted sm">${Object.entries(f.map||{}).map(([a,b])=>esc(a)+' → '+esc(b)).join(' · ')}</div></td><td class="sm muted">${f.linked?(SYNC.state[f.name]||'linked'):'manual'}${f.lastSync?'<br>last sync '+new Date(f.lastSync).toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit'}):''}</td><td><button class="btn ghost sm" data-a="sync-unlink" data-p="${esc(f.name)}">${icon('x')}</button></td></tr>`).join('')}</table>`:''}</div>
  <div class="card"><h2>Export</h2><div class="row mt"><button class="btn pri" data-a="export-xlsx">Excel workbook (everything)</button><button class="btn" data-a="export-xlsx-subj">Excel per class (re-importable)</button><button class="btn" data-a="export-csv">Lectures CSV</button><button class="btn" data-a="export-log">Study log CSV</button><button class="btn" data-a="print">PDF report (print)</button></div></div>
  ${pwaCard()}
  <div class="card"><h2>Backup &amp; restore</h2><div class="muted sm" style="margin:6px 0 10px">Last backup: ${S.meta.lastBackup?fd(S.meta.lastBackup):'never'} · stored in this browser only.</div><div class="row"><button class="btn pri" data-a="backup">Download backup (.json)</button><label class="btn" style="cursor:pointer">Restore backup<input type="file" id="restore-in" accept=".json" style="display:none"></label></div></div>
  <div class="card"><h2>Share with friends</h2><p class="muted sm" style="margin:6px 0 10px">Send them this same HTML file (or the blank copy), plus a template. A template carries your layout, theme, targets and subject setup – <b>never your progress, notes or study log</b>.</p><div class="row"><button class="btn pri" data-a="tpl-export" data-p="0">Template (settings &amp; layout)</button><button class="btn" data-a="tpl-export" data-p="1">Template + lecture lists (all pending)</button><label class="btn" style="cursor:pointer">Import a template<input type="file" id="tpl-in" accept=".json" style="display:none"></label></div></div>
  <div class="card"><h2>Profiles</h2><p class="muted sm" style="margin:6px 0 10px">Separate dashboards on this device (for siblings or different attempts). Each profile has its own data.</p><div class="row">${PROF.list.map(p=>`<button class="btn ${p.id===PROF.active?'pri':''}" data-a="prof-switch" data-p="${p.id}">${esc(p.name)}</button>`).join('')}<button class="btn" data-a="prof-new">+ New profile</button><button class="btn" data-a="prof-ren">Rename</button>${PROF.list.length>1?'<button class="btn danger" data-a="prof-del">Delete current</button>':''}</div></div>
  <div class="card"><h2>Import log</h2>${S.imports.length?S.imports.map(l=>`<div class="mt"><b>${esc(l.file)}</b> <span class="muted sm">${new Date(l.at).toLocaleString('en-IN')}</span>${l.sheets.map(s=>s.type==='lectures'?`<div class="sm" style="padding:6px 0;border-bottom:1px dashed var(--line)">${chip(esc(s.subject),'acc')} sheet “${esc(s.sheet)}” · header at row ${s.headerRow} · <b>${s.rows}</b> lecture rows imported of ${s.scanned} rows scanned (${s.blank} blank) · ${fh(s.min)} total ${s.sheetTotal!==null?(Math.abs(s.sheetTotal-s.min)<1?chip('= sheet total ✓','good'):chip('sheet total '+fh(s.sheetTotal),'warn')):''} · ${s.done} marked watched${s.res?` · new ${s.res.added}, updated ${s.res.updated}${s.res.gone?', not in sheet '+s.res.gone:''}${s.res.statusChanged?', '+s.res.statusChanged+' status changes from Excel':''}`:''}${s.ignored?`<div class="muted sm">${s.ignored} side-panel cell(s) ignored${s.panel&&s.panel.length?' · found: '+esc(s.panel.slice(0,4).join('; ')):''}</div>`:''}</div>`:`<div class="sm" style="padding:6px 0;border-bottom:1px dashed var(--line)">${chip('checklist')} “${esc(s.subject)}” · ${s.rows} items imported</div>`).join('')}</div>`).join(''):'<div class="muted sm">Nothing imported yet.</div>'}</div>`;
}
function importModal(){
  const I=UI.imp;if(!I)return'';
  return`<h2>Review import</h2><div class="muted sm" style="margin-bottom:10px">Nothing is changed until you confirm. “Class” is where the lectures go – pick an existing class to merge, or type a new name.</div>
  ${I.files.map((f,fi)=>f.parsed.map((p,si)=>{const pl=f.plan[si];
    if(p.type==='info')return`<div class="card" style="box-shadow:none"><b>${esc(f.name)} · ${esc(p.name)}</b> ${chip('notes')}<div class="muted sm mt" style="white-space:pre-line">${esc(p.lines.slice(0,6).join('\n'))}</div><div class="muted sm">Attached to the checklist from this file (if any).</div></div>`;
    const r=p.report,tot=r&&r.totals&&r.totals.length?r.totals[r.totals.length-1].min:null,sum=p.rows?p.rows.reduce((a,x)=>a+x.m,0):0;
    return`<div class="card" style="box-shadow:none"><div class="row sb"><b>${esc(f.name)} · ${esc(p.name)}</b>${chip(p.type==='lectures'?'lecture index':'checklist','acc')}</div>
    <div class="row mt"><label class="f grow">${p.type==='lectures'?'Class':'Checklist name'}<input data-c="imp-subj" data-p="${fi}|${si}" value="${esc(pl.subject)}" list="subjlist"></label><label class="row sm" style="align-self:end"><input type="checkbox" data-c="imp-skip" data-p="${fi}|${si}" ${pl.skip?'checked':''}> skip this sheet</label></div>
    ${p.type==='lectures'?`<div class="sm mt"><b>${p.rows.length}</b> lecture rows found (header on row ${r.headerRow}, ${r.scanned} rows scanned, ${r.blank} blank) · ${fh(sum)} total ${tot!==null?(Math.abs(tot-sum)<1?chip('matches sheet total ✓','good'):chip('sheet total says '+fh(tot),'warn')):''}<br><span class="muted">Columns: ${Object.entries(r.columns).filter(([k,v])=>v&&!(Array.isArray(v)&&!v.length)).map(([k,v])=>k+' = '+(Array.isArray(v)?v.join(', '):esc(v))).join(' · ')}</span>${p.rows.filter(x=>x.date).length?`<br><span class="muted">${p.rows.filter(x=>x.date).length} rows carry a date</span>`:''}${r.ignored?`<br><span class="muted">${r.ignored} side-panel cells ignored</span>`:''}</div>
     <div class="scroll mt"><table class="sm"><tr><th>#</th><th>Day</th><th>Topic</th><th>Length</th><th>Status</th></tr>${p.rows.slice(0,4).map((x,i)=>`<tr><td>${x.srcRow}</td><td>${esc(x.l)}</td><td>${esc(x.t)}</td><td>${fh(x.m)}</td><td>${x.xs}</td></tr>`).join('')}<tr><td colspan="5" class="muted">… and ${Math.max(0,p.rows.length-4)} more rows, all imported</td></tr></table></div>`:`<div class="sm mt"><b>${p.items.length}</b> checklist items found.</div>`}</div>`}).join('')).join('')}
  <datalist id="subjlist">${S.subjects.map(s=>`<option value="${esc(s.name)}">`).join('')}</datalist>
  <div class="row" style="justify-content:flex-end"><button class="btn" data-a="modal-close">Cancel</button><button class="btn pri" data-a="imp-confirm">Import everything</button></div>`;
}

/* ---------- customize / settings ---------- */
const PRESETS={Indigo:{accent:'#4f46e5',mode:'auto',radius:14,font:'sans'},Midnight:{accent:'#7c96ff',mode:'dark',radius:14,font:'sans'},Paper:{accent:'#b45309',mode:'light',radius:8,font:'serif'},Forest:{accent:'#15803d',mode:'auto',radius:16,font:'rounded'},Sunset:{accent:'#e11d48',mode:'auto',radius:20,font:'rounded'},Ocean:{accent:'#0891b2',mode:'auto',radius:12,font:'sans'},Mono:{accent:'#334155',mode:'auto',radius:4,font:'mono'}};
function vSettings(){
  const s=S.settings,t=s.theme;
  const num=(lbl,path,val,extra)=>`<label class="f">${lbl}<input type="number" data-c="set" data-p="${path}" data-t="n" value="${val}" ${extra||''}></label>`;
  const tim=(lbl,path,val)=>`<label class="f">${lbl}<input type="time" data-c="set" data-p="${path}" value="${val}"></label>`;
  return`<div class="card"><h2>Look &amp; feel</h2><div class="row mt">${Object.keys(PRESETS).map(k=>`<button class="btn ${t.preset===k?'pri':''}" data-a="theme-preset" data-p="${k}"><span class="dot" style="background:${PRESETS[k].accent}"></span>${k}</button>`).join('')}</div>
   <div class="row mt2"><label class="f">Mode<select data-c="set" data-p="settings.theme.mode">${['auto','light','dark'].map(m=>`<option ${t.mode===m?'selected':''}>${m}</option>`).join('')}</select></label>
   <label class="f">Accent colour<input type="color" data-c="set" data-p="settings.theme.accent" value="${t.accent}"></label>
   <label class="f">Font<select data-c="set" data-p="settings.theme.font">${[['sans','Clean sans'],['rounded','Rounded'],['serif','Serif'],['mono','Mono']].map(([k,l])=>`<option value="${k}" ${t.font===k?'selected':''}>${l}</option>`).join('')}</select></label>
   <label class="f">Corner radius (${t.radius}px)<input type="range" min="2" max="26" data-c="set" data-p="settings.theme.radius" data-t="n" value="${t.radius}"></label>
   <label class="f">Density<select data-c="set" data-p="settings.theme.density">${['compact','cozy','roomy'].map(m=>`<option ${t.density===m?'selected':''}>${m}</option>`).join('')}</select></label></div>
   <div class="row mt2"><button class="btn" data-a="customize">${icon('grid')}Customize dashboard layout</button><select data-c="layout-preset"><option value="">Apply a layout preset…</option>${Object.keys(LAYOUTS).map(k=>`<option>${k}</option>`).join('')}</select></div></div>
  <div class="card"><h2>Study plan</h2><div class="muted sm" style="margin:6px 0 12px">These drive the auto-planner, the lecture schedule/projections and your daily, weekly and monthly goals.</div><div class="row">
   ${num('Daily study target (hours, min 6)','settings.goal',s.goal,'min="6" max="16" step="0.5"')}${num('Max lecture hours / day (capped at ~70% of target)','settings.lecPerDay',s.lecPerDay,'min="1" max="14" step="0.5"')}
   <label class="f">Lighter day (half target)<select data-c="set" data-p="settings.lightDay" data-t="n">${['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'].map((n,i)=>`<option value="${i}" ${s.lightDay===i?'selected':''}>${n}</option>`).join('')}</select></label>
   <label class="f">Weekly timed-test day<select data-c="set" data-p="settings.testDay" data-t="n">${['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'].map((n,i)=>`<option value="${i}" ${s.testDay===i?'selected':''}>${n}</option>`).join('')}</select></label>
   ${tim('Day starts','settings.dayStart',s.dayStart)}${tim('Day ends','settings.dayEnd',s.dayEnd)}${tim('Lunch at','settings.lunchAt',s.lunchAt)}${num('Lunch (min)','settings.lunchMin',s.lunchMin)}${num('Break after (min of work)','settings.sessionMin',s.sessionMin)}${num('Break length (min)','settings.breakMin',s.breakMin)}${num('Streak counts from (min/day)','settings.streakMin',s.streakMin)}
   <label class="f">Warm-up revision first<select data-c="set" data-p="settings.warmup" data-t="b"><option value="true" ${s.warmup?'selected':''}>Yes</option><option value="false" ${!s.warmup?'selected':''}>No</option></select></label></div>
   <div class="row mt2"><label class="f">Revision gaps in days (1st, 2nd, 3rd)<input data-c="set" data-p="settings.revSteps" data-t="list" value="${s.revSteps.join(', ')}"></label>${num('Then repeat every (days)','settings.revRepeat',s.revRepeat)}</div></div>
  <div class="card"><h2>Classes</h2><div class="scroll"><table><tr><th>Name</th><th>Full name</th><th>Faculty</th><th>Colour</th><th>Priority</th><th>Lectures due</th><th>Batch total (h)</th><th>Hide</th><th></th></tr>${S.subjects.map(x=>`<tr><td><input data-c="subj" data-p="${x.id}|name" value="${esc(x.name)}" style="width:90px"></td><td><input data-c="subj" data-p="${x.id}|full" value="${esc(x.full)}" style="width:170px"></td><td><input data-c="subj" data-p="${x.id}|fac" value="${esc(x.fac)}" style="width:130px"></td><td><input type="color" data-c="subj" data-p="${x.id}|color" value="${x.color}"></td>
   <td><select data-c="subj" data-p="${x.id}|pri"><option value="1" ${x.pri===1?'selected':''}>High</option><option value="2" ${x.pri===2?'selected':''}>Normal</option><option value="3" ${x.pri===3?'selected':''}>Low</option></select></td><td><input type="date" data-c="subj" data-p="${x.id}|due" value="${x.due||''}"></td><td><input type="number" data-c="subj" data-p="${x.id}|batchH" value="${x.batchH||''}" style="width:80px" placeholder="e.g. 240" title="Total course hours – counts not-yet-released lectures"></td><td><input type="checkbox" data-c="subj" data-p="${x.id}|hide" ${x.hide?'checked':''}></td><td><button class="btn ghost sm danger" data-a="subj-del" data-p="${x.id}">${icon('trash')}</button></td></tr>`).join('')}</table></div><div class="row mt"><button class="btn" data-a="subj-add">+ Add class</button><span class="muted sm">“Batch total” (e.g. DT 240h) lets the schedule count recordings not yet released. “Lectures due” overrides the milestone date for urgency.</span></div></div>
  <div class="card"><h2>Exams &amp; milestones</h2>${S.exams.map(e=>`<div class="row mt"><input data-c="exam" data-p="${e.id}|name" value="${esc(e.name)}" class="grow"><input type="date" data-c="exam" data-p="${e.id}|date" value="${e.date}"><button class="btn ghost sm danger" data-a="exam-del" data-p="${e.id}">${icon('trash')}</button></div>`).join('')}<button class="btn sm mt" data-a="exam-add">+ Exam</button>
   <h2 class="mt2">Milestones</h2>${S.milestones.map(m=>`<div class="row mt"><input data-c="ms" data-p="${m.id}|name" value="${esc(m.name)}" class="grow" style="min-width:180px"><select data-c="ms" data-p="${m.id}|kind">${[['lec','Lectures complete'],['rev1','First revision'],['rev2','Second revision'],['rev3','Final revision'],['tests','Tests logged'],['manual','Manual']].map(([k,l])=>`<option value="${k}" ${m.kind===k?'selected':''}>${l}</option>`).join('')}</select>
   <select data-c="ms" data-p="${m.id}|subjects"><option value="all" ${m.subjects==='all'?'selected':''}>All classes</option>${S.subjects.map(x=>`<option value="${x.id}" ${Array.isArray(m.subjects)&&m.subjects.length===1&&m.subjects[0]===x.id?'selected':''}>${esc(x.name)} only</option>`).join('')}</select>
   <input type="date" data-c="ms" data-p="${m.id}|date" value="${m.date}">${m.kind==='manual'?`<label class="row sm"><input type="checkbox" data-c="ms" data-p="${m.id}|done" ${m.done?'checked':''}> done</label>`:''}<button class="btn ghost sm danger" data-a="ms-del" data-p="${m.id}">${icon('trash')}</button></div>`).join('')}<button class="btn sm mt" data-a="ms-add">+ Milestone</button></div>
  <div class="card"><h2>Fixed commitments</h2><div class="muted sm" style="margin:6px 0">The auto-planner schedules around these (classes, gym, travel…).</div>${S.fixed.map(f=>`<div class="row mt"><input data-c="fixed" data-p="${f.id}|name" value="${esc(f.name)}" class="grow"><input type="time" data-c="fixed" data-p="${f.id}|start" value="${f.start}"><input type="time" data-c="fixed" data-p="${f.id}|end" value="${f.end}"><input data-c="fixed" data-p="${f.id}|days" value="${(f.days||[]).join(',')}" placeholder="days 0-6 (blank = daily)" style="width:150px"><button class="btn ghost sm danger" data-a="fixed-del" data-p="${f.id}">${icon('trash')}</button></div>`).join('')}<button class="btn sm mt" data-a="fixed-add">+ Commitment</button></div>
  <div class="card"><h2>Notifications</h2><div class="row mt"><button class="btn" data-a="notif">Enable browser notifications</button><span class="muted sm">${'Notification' in window?'Permission: '+Notification.permission:'Not supported in this browser'} · reminders and timers also show inside the app.</span></div></div>`;
}

function pwaCard(){
  const secure=/^https?:$/.test(location.protocol),mb=n=>(n/1048576).toFixed(1)+' MB';
  const st=PWA.installed?'Installed – running as an app window':PWA.ev?'Ready to install':secure?'Open in Edge/Chrome and use the browser’s install icon (or the menu → Apps → Install)':'Not installable from a file – serve it from localhost or https (see README)';
  const off=PWA.sw==='ready'?'Yes – everything is saved on this device, no internet needed':PWA.sw==='unavailable'?'Not active (needs localhost/https)':PWA.sw;
  return`<div class="card"><h2>App &amp; offline</h2><table class="mt sm">
   <tr><td>Install</td><td>${st}</td></tr><tr><td>Works offline</td><td>${esc(off)}</td></tr>
   <tr><td>Data protection</td><td>${PWA.persisted===true?'Browser has promised not to clear this data automatically':PWA.persisted===false?'Not guaranteed – the browser may clear data under disk pressure. Keep a JSON backup.':'Unknown'}</td></tr>
   <tr><td>Storage used</td><td>${PWA.est?mb(PWA.est.usage||0)+' of '+mb(PWA.est.quota||0):'–'}</td></tr></table>
   <div class="row mt">${PWA.ev?`<button class="btn pri" data-a="pwa-install">Install app</button>`:''}${PWA.waiting?`<button class="btn pri" data-a="pwa-update">Update to new version</button>`:''}</div>
   <p class="muted sm" style="margin-top:8px">Your progress lives in this browser profile on this computer. Updating the app never touches it; a copy from before each update is also kept automatically.</p></div>`}

/* ===================== APP: shell, actions, sync, timers, exports ===================== */
function toast(msg,kind){const t=document.createElement('div');t.className='toast '+(kind||'');t.textContent=msg;$('#toasts').appendChild(t);setTimeout(()=>t.remove(),kind==='bad'?7000:4200)}
function openModal(html){const m=$('#modal');m.innerHTML='<div>'+html+'</div>';m.classList.add('open')}
function closeModal(){$('#modal').classList.remove('open');$('#modal').innerHTML='';UI.imp=null}
const lecMap=()=>memo('lm',()=>{const m=new Map();for(const sid in S.lectures)S.lectures[sid].forEach(l=>m.set(l.id,l));return m});
const lecById=id=>lecMap().get(id);
const chById=id=>{for(const s in S.chapters){const c=S.chapters[s].find(x=>x.id===id);if(c)return[s,c]}return[]};
function inkFor(hex){const n=parseInt(hex.slice(1),16),r=n>>16,g=n>>8&255,b=n&255;return(0.299*r+0.587*g+0.114*b)>160?'#101828':'#fff'}
function applyTheme(){const t=S.settings.theme,r=document.documentElement;r.dataset.mode=t.mode;r.dataset.font=t.font;r.dataset.density=t.density;r.style.setProperty('--accent',t.accent);r.style.setProperty('--accent-ink',inkFor(t.accent));r.style.setProperty('--radius',t.radius+'px');const m=document.querySelector('meta[name=theme-color]');if(m)m.content=getComputedStyle(r).getPropertyValue('--bg').trim()||m.content}

function syncPill(){const fs=S.sync.files.filter(f=>f.linked);if(!fs.length)return`<div class="pill-sync" data-a="tab" data-p="data" style="cursor:pointer"><span class="dot" style="background:var(--muted)"></span><span class="lbl">Excel: manual</span></div>`;
  const bad=fs.some(f=>/reconnect|error/i.test(SYNC.state[f.name]||''));
  return`<div class="pill-sync" data-a="${bad?'sync-reconnect':'tab'}" data-p="data" style="cursor:pointer"><span class="dot" style="background:${bad?'var(--warn)':'var(--good)'}"></span><span class="lbl">${bad?'Reconnect Excel':'Live sync · '+fs.length+' file'+(fs.length>1?'s':'')}</span></div>`}
function shell(){
  const t=TABS.find(x=>x[0]===UI.tab),view=({dashboard:vDashboard,lectures:vLectures,planner:vPlanner,chapters:vChapters,analytics:vAnalytics,tests:vTests,checklists:vChecklists,data:vData,settings:vSettings})[UI.tab]();
  const T=todayStr(),ex=S.exams[0],dl=ex?diffDays(T,ex.date):null;
  const sub=UI.tab==='dashboard'?`<span class="muted sm" style="white-space:nowrap">${fdw(T)}${dl!==null&&dl>=0?` · ${dl} days to ${esc(ex.name.split('–').pop().trim())}`:''}</span>`:'';
  return`<div class="shell"><aside class="side"><div class="brand"><i>🎓</i><span>${esc((PROF.list.find(p=>p.id===PROF.active)||{name:'Study'}).name)}</span></div><nav class="nav">${TABS.map(([k,l,i])=>`<button class="${k===UI.tab?'on':''}" data-a="tab" data-p="${k}">${icon(i)}<span>${l}</span></button>`).join('')}</nav><div class="foot">${syncPill()}<button class="btn ghost" data-a="mode" style="justify-content:flex-start">${icon('moon')}<span class="lbl">Light / dark</span></button></div></aside>
  <div class="main"><header class="top"><h1>${t[1]}</h1>${sub}<span class="sp"></span><div class="search"><span>${icon('search')}</span><input id="gsearch" placeholder="Search all lectures & topics…"></div>
   <span id="pomo-chip">${pomoChip()}</span>${UI.tab==='dashboard'?`<button class="btn ${UI.custom?'pri':''}" data-a="customize">${icon('sliders')}<span>${UI.custom?'Done':'Customize'}</span></button>`:''}</header><div id="view">${view}</div></div></div>
  <nav class="bottom">${TABS.map(([k,l,i])=>`<button class="${k===UI.tab?'on':''}" data-a="tab" data-p="${k}">${icon(i)}<span>${l.split(' ')[0]}</span></button>`).join('')}</nav><button class="fab" data-a="quick" data-p="menu" aria-label="Quick add">${icon('plus')}</button>`;
}
function render(){const y=window.scrollY;applyTheme();$('#app').innerHTML=shell();window.scrollTo(0,y);if(typeof syncTimers==='function')syncTimers()}
function commit(){save();render()}

/* ---------- lecture state changes ---------- */
function setLecDone(l,val){
  l.d=val;S.logs=S.logs.filter(x=>x.lec!==l.id);
  if(val){l.doneOn=todayStr();l.src='app';l.missed=false;
    const t=S.tasks.find(x=>x.lec===l.id&&!x.done);if(t){t.done=true;if(!t.act)t.act=Math.round(t.plan)}else if(l.m>0)S.logs.push({id:uid('g'),date:l.doneOn,subj:l.sid,min:Math.round(l.m),src:'lecture',lec:l.id,note:l.t})}
  else{l.doneOn=null;S.tasks.forEach(t=>{if(t.lec===l.id&&t.done){t.done=false;t.act=0}})}
}
function setTaskDone(t,val){
  t.done=val;if(val&&!(+t.act))t.act=+t.plan||0;if(!val&&+t.act===+t.plan)t.act=0;
  if(t.lec){const l=lecById(t.lec);if(l){l.d=val;if(val){l.doneOn=t.date;l.src='app';l.missed=false;S.logs=S.logs.filter(x=>x.lec!==l.id)}else l.doneOn=null}}
  if(t.ch&&val&&(t.type==='revision')){const[s,c]=chById(t.ch);if(c){const k=c.r1?c.r2?'r3':'r2':'r1';c[k]=true;c.last=t.date}}
  if(t.ch&&val&&t.type==='practice'){const[s,c]=chById(t.ch);if(c)c.pr=true}
}

/* ---------- Excel import flow ---------- */
const baseName=n=>n.replace(/\.[a-z0-9]+$/i,'').replace(/^[0-9a-f]{6,10}[-_]/i,'').replace(/[_]+/g,' ').trim();
function defaultPlan(fileName,parsed){
  const rec=S.sync.files.find(f=>f.name===fileName),nLec=parsed.filter(p=>p.type==='lectures').length;
  return parsed.map(p=>{
    const known=rec&&rec.map&&rec.map[p.name];
    if(p.type==='lectures')return{subject:known||(nLec>1?p.name:guessSubject(fileName,p.name)),skip:false};
    if(p.type==='checklist')return{subject:known||baseName(fileName),skip:false};
    return{skip:true}});
}
async function handleFiles(files,opts){
  opts=opts||{};const list=[];
  for(const f of files){
    try{const wb=await readAnyFile(f);const parsed=parseWorkbook(wb);
      if(!parsed.some(p=>p.type!=='info')){toast(`“${f.name}”: no lecture table found (needs a header row with Topic / Duration / Status columns).`,'bad');continue}
      list.push({name:f.name.replace(/^[0-9a-f]{8}-/i,''),file:f,handle:opts.handle&&opts.handle[f.name],parsed,plan:null})}
    catch(e){toast(`“${f.name}”: ${e.message}`,'bad')}}
  if(!list.length)return;
  list.forEach(x=>x.plan=defaultPlan(x.name,x.parsed));
  const allKnown=list.every(x=>S.sync.files.some(f=>f.name===x.name));
  if(allKnown&&!opts.preview){finishImport(list,true);return}
  UI.imp={files:list};openModal(importModal());
}
function finishImport(list,silent){
  let lec=0,cls=new Set(),chg=0;
  list.forEach(x=>{
    const log=commitImport(x.name,x.parsed,x.plan);
    log.sheets.forEach(s=>{if(s.type==='lectures'){lec+=s.rows;cls.add(s.subject);if(s.res)chg+=s.res.added+s.res.statusChanged}});
    let rec=S.sync.files.find(f=>f.name===x.name);if(!rec){rec={name:x.name,map:{},linked:false};S.sync.files.push(rec)}
    x.parsed.forEach((p,i)=>{if(x.plan[i]&&!x.plan[i].skip&&p.type!=='info')rec.map[p.name]=x.plan[i].subject});
    rec.lastSync=Date.now();rec.lastMod=x.file.lastModified;
    if(x.handle){rec.linked=true;SYNC.handles[x.name]=x.handle;SYNC.state[x.name]='live';idbSet(PROF.active+':'+x.name,x.handle)}
  });
  fixSid();save();closeModal();render();
  if(silent){if(chg||true)toast(chg?`Synced – ${chg} change${chg>1?'s':''} picked up from Excel.`:'Spreadsheet refreshed – no changes.')}
  else toast(`Imported ${lec} lecture rows into ${[...cls].join(', ')||'your dashboard'}. See the Import log for the row-by-row check.`);
}

/* ---------- live link (File System Access API) ---------- */
const SYNC={state:{},handles:{},timer:null,busy:false};
const idb={open(){return new Promise((res,rej)=>{const r=indexedDB.open('cft3',1);r.onupgradeneeded=()=>r.result.createObjectStore('h');r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error)})}};
async function idbSet(k,v){try{const db=await idb.open();db.transaction('h','readwrite').objectStore('h').put(v,k)}catch(e){}}
async function idbGet(k){try{const db=await idb.open();return await new Promise(res=>{const r=db.transaction('h').objectStore('h').get(k);r.onsuccess=()=>res(r.result);r.onerror=()=>res(null)})}catch(e){return null}}
async function idbDel(k){try{const db=await idb.open();db.transaction('h','readwrite').objectStore('h').delete(k)}catch(e){}}
async function restoreHandles(){
  for(const f of S.sync.files.filter(x=>x.linked)){const h=await idbGet(PROF.active+':'+f.name);if(!h){SYNC.state[f.name]='needs reconnect';continue}
    SYNC.handles[f.name]=h;try{SYNC.state[f.name]=(await h.queryPermission({mode:'read'}))==='granted'?'live':'needs reconnect'}catch(e){SYNC.state[f.name]='needs reconnect'}}
}
async function pollSync(){
  if(SYNC.busy)return;SYNC.busy=true;
  try{for(const rec of S.sync.files.filter(f=>f.linked)){const h=SYNC.handles[rec.name];if(!h||SYNC.state[rec.name]!=='live')continue;
    try{const file=await h.getFile();if(file.lastModified===rec.lastMod)continue;
      const parsed=parseWorkbook(await readAnyFile(file));const x={name:rec.name,file,parsed,plan:defaultPlan(rec.name,parsed)};rec.lastMod=file.lastModified;finishImport([x],true)}
    catch(e){SYNC.state[rec.name]='error: '+e.message}}}
  finally{SYNC.busy=false}
}
function startSyncLoop(){clearInterval(SYNC.timer);SYNC.timer=null;
  /* poll only when files are actually linked AND the window is visible; otherwise no timer exists at all */
  if(document.hidden||!S.sync.files.some(f=>f.linked&&SYNC.handles[f.name]))return;
  SYNC.timer=setInterval(pollSync,Math.max(3,S.settings.syncSec)*1000)}

/* ---------- pomodoro ---------- */
const fmtClock=ms=>{ms=Math.max(0,Math.ceil(ms/1000));return pad(Math.floor(ms/60))+':'+pad(ms%60)};
function remainingMs(){const r=S.pomo.run;return r?(r.paused?r.left:r.endAt-Date.now()):0}
function pomoChip(){const r=S.pomo.run;return r?`<button class="btn sm" data-a="tab" data-p="dashboard" title="Focus timer">${r.mode==='work'?'🍅':'☕'} <span id="pomo-chip-t">${fmtClock(remainingMs())}</span></button>`:''}
function pomoStart(mode){const P=S.pomo,mins=mode==='work'?P.work:mode==='short'?P.short:P.long;P.run={mode,total:mins*60000,endAt:Date.now()+mins*60000,paused:false,left:mins*60000};save()}
function beep(){try{const a=new (window.AudioContext||window.webkitAudioContext)(),o=a.createOscillator(),g=a.createGain();o.connect(g);g.connect(a.destination);o.frequency.value=880;g.gain.setValueAtTime(.15,a.currentTime);g.gain.exponentialRampToValueAtTime(.001,a.currentTime+1.2);o.start();o.stop(a.currentTime+1.2)}catch(e){}}
function notify(title,body){if(S.settings.notify&&'Notification' in window&&Notification.permission==='granted'){try{new Notification(title,{body})}catch(e){}}}
function pomoComplete(){
  const P=S.pomo,r=P.run;beep();
  if(r.mode==='work'){const T=todayStr();if(P.doneDate!==T){P.doneDate=T;P.done=0}P.done++;S.logs.push({id:uid('g'),date:T,subj:P.subj||'',min:P.work,src:'focus',note:'Focus session'});
    const long=P.done%P.every===0;toast(`Focus session done – ${P.work} min logged. Take a ${long?P.long:P.short} min break.`);notify('Focus session complete',`${P.work} minutes logged.`);pomoStart(long?'long':'short')}
  else{toast('Break over – ready for the next focus session.');notify('Break over','Time to focus.');P.run=null;save()}
  render();
}
/* One light tick, only while it is needed and the window is visible:
   - 500 ms while a focus/break timer is running (to keep the clock accurate)
   - 30 s while there is an unfired reminder. With nothing to watch, there is NO timer. */
let _tick=null,_tickMs=0;
function pomoTick(){
  const r=S.pomo&&S.pomo.run;
  if(r&&!r.paused){if(remainingMs()<=0)pomoComplete();else{const a=$('#pomo-time'),b=$('#pomo-chip-t'),t=fmtClock(remainingMs());if(a)a.textContent=t;if(b)b.textContent=t}}
}
function remTick(){
  let hit=false;S.reminders.forEach(r=>{if(!r.done&&!r.fired&&r.at&&new Date(r.at)<=new Date()){r.fired=true;hit=true;toast('⏰ '+r.text);beep();notify('Reminder',r.text)}});
  if(hit){save();if(UI.tab==='dashboard')render()}
}
function needTick(){
  if(!S)return 0;
  const r=S.pomo&&S.pomo.run;if(r&&!r.paused)return document.hidden?30000:500;
  return S.reminders.some(x=>!x.done&&!x.fired&&x.at)?30000:0}
function syncTimers(){
  const ms=needTick();
  if(ms!==_tickMs){clearInterval(_tick);_tick=null;_tickMs=ms;if(ms)_tick=setInterval(()=>{pomoTick();remTick();if(needTick()!==_tickMs)syncTimers()},ms)}
}
document.addEventListener('visibilitychange',()=>{
  if(!S)return;
  if(!document.hidden){pomoTick();remTick();pollSync();flush()}else flush();
  syncTimers();startSyncLoop()});

/* ---------- exports ---------- */
function download(name,blob){const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},500)}
const csvCell=v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"';
const toCSV=rows=>'﻿'+rows.map(r=>r.map(csvCell).join(',')).join('\r\n');
function lecRows(sid){
  const xk=[...new Set(LL(sid).flatMap(l=>Object.keys(l.x||{})))];const chn=id=>((S.chapters[sid]||[]).find(c=>c.id===id)||{name:''}).name;
  const hdr=[{h:'Day'},{h:'Lecture'},{h:'Topic'},{h:'Duration'},{h:'Status'},{h:'Download Status'},{h:'Date'},{h:'Priority'},{h:'Chapter / unit'},{h:'Notes'}].concat(xk.map(k=>({h:k})));
  const rows=[hdr];
  LL(sid).forEach(l=>rows.push([l.l,l.n,l.t,l.m?{dur:l.m}:null,l.d?'Completed':isMissed(l)?'Missed':'Pending',l.dl===null||l.dl===undefined?null:(l.dl?'Completed':'Pending'),(l.doneOn||l.sdate)?{date:l.doneOn||l.sdate}:null,l.pri||'',chn(l.ch),l.note||''].concat(xk.map(k=>(l.x||{})[k]||''))));
  return rows;
}
function studyLogRows(){
  const rows=[];S.tasks.forEach(t=>{if(t.type!=='break'&&+t.act>0)rows.push([t.date,sjn(t.subj),t.text,+t.act,'planner'])});
  S.logs.forEach(l=>rows.push([l.date,sjn(l.subj),l.note||'',l.min,l.src]));
  for(const sid in S.lectures)S.lectures[sid].forEach(l=>{if(l.d&&l.doneOn&&l.src==='import'&&!l.gone&&l.m)rows.push([l.doneOn,sjn(sid),(l.l?l.l+' – ':'')+l.t,Math.round(l.m),'excel history'])});
  return rows.sort((a,b)=>a[0]<b[0]?-1:1);
}
function exportXlsx(perSubject){
  const sheets=[];
  if(!perSubject){
    sheets.push({name:'Summary',widths:[14,24,22,10,10,12,12,14,12],rows:[['Class','Full name','Faculty','Lectures','Watched','Total (h)','Left (h)','Finish by','Readiness %'].map(h=>({h}))].concat(S.subjects.map(s=>{const p=subjProjection(s.id);return[s.name,s.full,s.fac,LL(s.id).length,LL(s.id).filter(l=>l.d).length,Math.round(totalM(s.id)/6)/10,Math.round(remM(s.id)/6)/10,p.fin?{date:p.fin}:'',subjScore(s.id)]}))});
  }
  S.subjects.forEach(s=>{if(hasLectures(s.id))sheets.push({name:s.name,widths:[10,12,46,10,12,14,12,9,34,24],rows:lecRows(s.id)})});
  if(!perSubject){
    sheets.push({name:'Study log',widths:[12,14,50,10,14],rows:[['Date','Class','What','Minutes','Source'].map(h=>({h}))].concat(studyLogRows().map(r=>[{date:r[0]},r[1],r[2],r[3],r[4]]))});
    sheets.push({name:'Plan',widths:[12,8,8,14,10,60,9,9,7],rows:[['Date','Start','End','Class','Type','Task','Planned min','Done min','Done'].map(h=>({h}))].concat(S.tasks.slice().sort((a,b)=>(a.date+a.start)<(b.date+b.start)?-1:1).map(t=>[{date:t.date},t.start,t.end,sjn(t.subj),t.type,t.text,t.plan,t.act||0,t.done?'yes':'no']))});
    sheets.push({name:'Tests',widths:[12,14,8,30,8,8,9],rows:[['Date','Class','Type','Name','Marks','Out of','Minutes'].map(h=>({h}))].concat(S.tests.map(t=>[{date:t.date},sjn(t.subj),t.type,t.name,t.score,t.max,t.min||0]))});
    S.checklists.forEach(c=>sheets.push({name:('Checklist '+c.name).slice(0,31),widths:[6,22,70,16,8],rows:[['#','Group','Item','Reference','Done?'].map(h=>({h}))].concat(c.items.map(i=>[i.key,i.g,i.t,i.ref,i.done?'Yes':'No']))}));
  }
  if(!sheets.length)return toast('Nothing to export yet.','bad');
  download(perSubject?'lecture-indexes-'+todayStr()+'.xlsx':'study-dashboard-'+todayStr()+'.xlsx',buildXlsx(sheets));
}
function backup(){S.meta.lastBackup=todayStr();save();download('study-dashboard-backup-'+todayStr()+'.json',new Blob([JSON.stringify(S)],{type:'application/json'}));render()}
function templateExport(withLectures){
  const t={type:'cft-template',v:1,settings:JSON.parse(JSON.stringify(S.settings)),layout:S.layout,exams:S.exams,milestones:S.milestones,fixed:S.fixed,
    subjects:S.subjects.map(s=>({name:s.name,full:s.full,fac:s.fac,color:s.color,pri:s.pri,due:s.due,batchH:s.batchH,res:s.res}))};
  if(withLectures){t.lectures={};t.chapters={};S.subjects.forEach(s=>{t.chapters[s.name]=(S.chapters[s.id]||[]).map(c=>c.name);
    t.lectures[s.name]=LL(s.id).map(l=>({l:l.l,n:l.n,t:l.t,m:l.m,ch:((S.chapters[s.id]||[]).find(c=>c.id===l.ch)||{name:''}).name,dl:l.dl===null?null:false,x:l.x||{}}))})}
  download('study-template-'+(withLectures?'with-lectures-':'')+todayStr()+'.json',new Blob([JSON.stringify(t,null,1)],{type:'application/json'}));
}
function templateApply(t){
  if(!t||t.type!=='cft-template')throw new Error('Not a template file from this dashboard.');
  Object.assign(S.settings,{...t.settings,theme:{...S.settings.theme,...t.settings.theme}});
  if(t.layout)S.layout=t.layout;if(t.exams)S.exams=t.exams;if(t.milestones)S.milestones=t.milestones;if(t.fixed)S.fixed=t.fixed;
  (t.subjects||[]).forEach(x=>{const s=ensureSubject(x.name,{full:x.full,fac:x.fac,color:x.color,pri:x.pri,due:x.due,batchH:x.batchH,res:x.res});
    if(t.lectures&&t.lectures[x.name]&&!S.lectures[s.id].length){const cm={};(t.chapters[x.name]||[]).forEach(n=>cm[n]=newChapter(s.id,n).id);
      S.lectures[s.id]=t.lectures[x.name].map((l,i)=>{if(l.ch&&!cm[l.ch])cm[l.ch]=newChapter(s.id,l.ch).id;return{id:uid('l'),k:'t'+i,l:l.l,n:l.n,t:l.t,m:l.m,d:false,xs:'pending',dl:l.dl,xdl:l.dl,ch:cm[l.ch]||(S.chapters[s.id][0]||newChapter(s.id,'Lectures')).id,pri:'',note:'',doneOn:null,sdate:null,missed:false,arch:false,src:'import',x:l.x||{}}})}});
  fixSid();
}
function printReport(){
  const T=todayStr(),p=$('#print');
  p.innerHTML=`<h1>Study report – ${esc(PROF.list.find(x=>x.id===PROF.active).name)}</h1><div>${fdw(T)} · daily target ${S.settings.goal}h · streak ${streak()} days · today ${fh(minOn(T))}</div>
  <h2>Classes</h2><table><tr><th>Class</th><th>Lectures</th><th>Watched</th><th>Left</th><th>Finish by</th><th>Target</th><th>Readiness</th></tr>${S.subjects.map(s=>{const pr=subjProjection(s.id);return`<tr><td>${esc(s.name)}</td><td>${LL(s.id).filter(l=>l.d).length}/${LL(s.id).length}</td><td>${pct(doneM(s.id),totalM(s.id))}%</td><td>${fh(pr.rem)}</td><td>${pr.fin?fd(pr.fin):'–'}</td><td>${fd(pr.due)}</td><td>${subjScore(s.id)}%</td></tr>`}).join('')}</table>
  <h2>Milestones &amp; exams</h2><table>${S.milestones.map(m=>`<tr><td>${esc(m.name)}</td><td>${fd(m.date)}</td><td>${msProgress(m)}%</td></tr>`).join('')}${S.exams.map(e=>`<tr><td>${esc(e.name)}</td><td>${fd(e.date)}</td><td>${diffDays(T,e.date)} days</td></tr>`).join('')}</table>
  ${S.subjects.filter(s=>LL(s.id).some(l=>!l.d)).map(s=>`<h2>Pending lectures – ${esc(s.name)}</h2><table><tr><th>Day</th><th>Topic</th><th>Length</th><th>Planned</th></tr>${LL(s.id).filter(l=>!l.d).map(l=>`<tr><td>${esc(l.l)}</td><td>${esc(l.t)}</td><td>${fh(l.m)}</td><td>${fd(schedule().planned[l.id],1)}</td></tr>`).join('')}</table>`).join('')}`;
  window.print();
}

/* ---------- quick add ---------- */
function quick(kind){
  const F=(id,lbl,inner)=>`<label class="f">${lbl}${inner}</label>`;
  const forms={
   menu:`<h2>Quick add</h2><div class="pills">${[['log','Log study time'],['task',"Add task to today"],['note','Note'],['reminder','Reminder'],['lecture','Lecture'],['focus','Start focus session']].map(([k,l])=>`<button data-a="quick" data-p="${k}">${l}</button>`).join('')}</div>`,
   log:`<h2>Log study time</h2><div class="row">${F('','Class',`<select id="q-subj">${subjOpts('','General')}</select>`)}${F('','Minutes',`<input id="q-min" type="number" min="5" step="5" value="60" style="width:100px">`)}${F('','Note',`<input id="q-note" placeholder="optional" style="width:240px">`)}</div>`,
   task:`<h2>Add task to today</h2><div class="row">${F('','Class',`<select id="q-subj">${subjOpts('','General')}</select>`)}${F('','Task',`<input id="q-text" style="width:300px">`)}${F('','Minutes',`<input id="q-min" type="number" value="45" style="width:90px">`)}</div>`,
   note:`<h2>Note</h2><textarea id="q-text" placeholder="Write a note…"></textarea>`,
   reminder:`<h2>Reminder</h2><div class="row">${F('','Remind me to',`<input id="q-text" style="width:300px">`)}${F('','When',`<input id="q-at" type="datetime-local">`)}</div>`,
   lecture:`<h2>Add a lecture</h2><div class="row">${F('','Class',`<select id="q-subj" data-c="q-subj">${subjOpts('')}</select>`)}${F('','Day / lecture',`<input id="q-l" style="width:110px">`)}${F('','Topic',`<input id="q-text" style="width:260px">`)}${F('','Length (h:mm)',`<input id="q-dur" placeholder="1:30" style="width:100px">`)}</div>`,
   focus:`<h2>Focus session</h2><div class="row">${F('','Class',`<select id="q-subj">${subjOpts(S.pomo.subj,'General')}</select>`)}</div>`};
  openModal(forms[kind]+(kind==='menu'?'':`<div class="row mt2" style="justify-content:flex-end"><button class="btn" data-a="modal-close">Cancel</button><button class="btn pri" data-a="quick-save" data-p="${kind}">Save</button></div>`));
}
function quickSave(kind){
  const v=id=>($('#q-'+id)||{}).value||'',T=todayStr();
  if(kind==='log'){const m=+v('min');if(!(m>0))return;S.logs.push({id:uid('g'),date:T,subj:v('subj'),min:m,src:'manual',note:v('note')})}
  else if(kind==='task'){if(!v('text'))return;S.tasks.push({id:uid('t'),date:T,start:'',end:'',subj:v('subj'),type:'custom',text:v('text'),plan:+v('min')||30,act:0,done:false})}
  else if(kind==='note'){if(!v('text').trim())return;S.notes.push({id:uid('n'),text:v('text').trim(),pin:false,ts:Date.now()})}
  else if(kind==='reminder'){if(!v('text'))return;S.reminders.push({id:uid('r'),text:v('text'),at:v('at'),done:false})}
  else if(kind==='lecture'){const sid=v('subj');const m=parseDurText(v('dur'));if(!sid||!v('text'))return toast('Pick a class and enter a topic.','bad');const chs=S.chapters[sid];const ch=(chs[chs.length-1]||newChapter(sid,'Added lectures')).id;
    S.lectures[sid].push({id:uid('l'),k:uid('k'),l:v('l'),n:'',t:v('text'),m:m||0,d:false,xs:'pending',dl:null,ch,pri:'',note:'',src:'app',x:{},sid})}
  else if(kind==='focus'){S.pomo.subj=v('subj');pomoStart('work');closeModal();UI.tab='dashboard';return commit()}
  closeModal();commit();toast('Saved.');
}

/* ---------- actions ---------- */
const ACT={
 tab:(e,p)=>{UI.tab=p;if(p!=='dashboard')UI.custom=false;commit();window.scrollTo(0,0)},
 mode:()=>{const t=S.settings.theme,m=document.documentElement.dataset.mode,dark=m==='dark'||(m==='auto'&&matchMedia('(prefers-color-scheme:dark)').matches);t.mode=dark?'light':'dark';t.preset='Custom';commit()},
 quick:(e,p)=>quick(p),'quick-save':(e,p)=>quickSave(p),'modal-close':()=>closeModal(),
 customize:()=>{UI.custom=!UI.custom;if(UI.tab!=='dashboard'){UI.tab='dashboard'}render()},
 autoplan:(e,p)=>{const r=autoPlan(p);commit();toast(`Planned ${fh(r.total)} for ${p===todayStr()?'today':fdw(p)}${r.overflow?` – runs ${fh(r.overflow)} past your day end`:''}.`)},
 'autoplan-week':()=>{let ov=0;for(let i=0;i<7;i++){const d=addDays(todayStr(),i);ov+=autoPlan(d).overflow>0?1:0}commit();toast('Planned the next 7 days.'+(ov?` ${ov} day(s) overflow your day end.`:''))},
 'clear-auto':(e,p)=>{S.tasks=S.tasks.filter(t=>!(t.date===p&&t.auto&&!t.done&&!t.lock));commit()},
 'move-backlog':()=>{S.tasks.filter(t=>t.date<todayStr()&&!t.done&&t.type!=='break').forEach(t=>{t.date=UI.plDate;t.start='';t.end='';t.auto=false});commit()},
 'goto-subj':(e,p)=>{UI.lec.sid=p;UI.lec.status='all';UI.tab='lectures';commit()},
 'lec-subj':(e,p)=>{UI.lec.sid=UI.lec.sid===p?'all':p;UI.lec.page=0;commit()},
 lecview:(e,p)=>{UI.lec.view=p;UI.lec.page=0;if(p==='date'&&UI.tab!=='lectures'){UI.tab='lectures';UI.lec.status='all'}commit()},
 lpage:(e,p)=>{UI.lec.page=Math.max(0,UI.lec.page+ +p);$('#lec-results').innerHTML=lecResults()},
 lecdone:(e,p)=>{const l=lecById(p);if(l){setLecDone(l,true);commit()}},unmiss:(e,p)=>{const l=lecById(p);l.missed=false;S.tasks=S.tasks.filter(t=>!(t.lec===p&&!t.done&&t.date<todayStr()));commit()},
 'dl-mark':(e,p)=>{const l=lecById(p);l.dl=true;commit()},
 addrev:(e,p)=>{const[sid,cid]=p.split('|'),[,c]=chById(cid);S.tasks.push({id:uid('t'),date:todayStr(),start:'',end:'',subj:sid,type:'revision',text:`Revise ${sjn(sid)}: ${c.name}`,plan:40,act:0,done:false,ch:cid});commit();toast('Added to today\'s plan.')},
 'rem-add':()=>{const t=$('#rem-text').value.trim();if(!t)return;S.reminders.push({id:uid('r'),text:t,at:$('#rem-at').value,done:false});commit()},
 'rem-done':(e,p)=>{S.reminders.find(r=>r.id===p).done=true;commit()},
 'note-add':()=>{const t=$('#note-text').value.trim();if(!t)return;S.notes.push({id:uid('n'),text:t,pin:false,ts:Date.now()});commit()},
 'note-pin':(e,p)=>{const n=S.notes.find(x=>x.id===p);n.pin=!n.pin;commit()},'note-del':(e,p)=>{S.notes=S.notes.filter(n=>n.id!==p);commit()},
 'w-move':(e,p)=>{const[id,d]=p.split('|'),A=S.layout.widgets,vis=A.filter(w=>!w.hidden),i=vis.findIndex(w=>w.id===id),j=i+ +d;if(j<0||j>=vis.length)return;const a=A.indexOf(vis[i]),b=A.indexOf(vis[j]);[A[a],A[b]]=[A[b],A[a]];S.layout.preset='Custom';commit()},
 'w-size':(e,p)=>{const w=S.layout.widgets.find(x=>x.id===p);w.size=w.size%3+1;S.layout.preset='Custom';commit()},
 'w-hide':(e,p)=>{S.layout.widgets.find(x=>x.id===p).hidden=true;commit()},'w-show':(e,p)=>{S.layout.widgets.find(x=>x.id===p).hidden=false;commit()},
 'w-add':(e,p)=>{S.layout.widgets.push({id:p,size:+WDEF[p][1],hidden:false});commit()},
 'pomo-toggle':()=>{const P=S.pomo,r=P.run;if(!r)pomoStart('work');else if(r.paused){r.paused=false;r.endAt=Date.now()+r.left}else{r.paused=true;r.left=r.endAt-Date.now()}commit()},
 'pomo-reset':()=>{S.pomo.run=null;commit()},'pomo-skip':()=>{const r=S.pomo.run;if(r)pomoComplete();else pomoStart('short');commit()},
 'pomo-preset':(e,p)=>{const[w,b]=p.split('|').map(Number);Object.assign(S.pomo,{work:w,short:b,run:null});commit()},
 'open-chk':(e,p)=>{UI.chk=p;UI.tab='checklists';commit()},
 editdur:(e,p)=>{const l=lecById(p),v=prompt('Lecture length (h:mm, e.g. 1:45)',l.m?Math.floor(l.m/60)+':'+pad(Math.round(l.m%60)):'');if(v===null)return;const m=parseDurText(v);if(m===null)return alert('Could not read that length.');l.m=m;commit()},
 'lec-pri':(e,p)=>{const l=lecById(p);l.pri=l.pri===''||!l.pri?'high':l.pri==='high'?'low':'';commit()},'lec-miss':(e,p)=>{const l=lecById(p);l.missed=!l.missed;commit()},
 'lec-note':(e,p)=>{const l=lecById(p),v=prompt('Note for this lecture',l.note||'');if(v===null)return;l.note=v;commit()},'lec-arch':(e,p)=>{const l=lecById(p);l.arch=!l.arch;commit()},
 'lec-del':(e,p)=>{if(!confirm('Remove this lecture from the dashboard? (If it is still in your Excel file it will come back on the next sync.)'))return;const l=lecById(p);S.lectures[l.sid]=S.lectures[l.sid].filter(x=>x!==l);commit()},
 'arch-all':()=>{if(!confirm('Archive every completed lecture? They stay counted in progress but disappear from the lists (see Archive view).'))return;let n=0;Object.values(S.lectures).flat().forEach(l=>{if(l.d&&!l.arch){l.arch=true;n++}});commit();toast(n+' lectures archived.')},
 'pl-nav':(e,p)=>{UI.plDate=p==='0'?todayStr():addDays(UI.plDate,+p);commit()},'pl-day':(e,p)=>{UI.plDate=p;commit()},
 'task-add':()=>{const t=$('#nt-text').value.trim();if(!t)return;const st=$('#nt-start').value,m=+$('#nt-min').value||30;S.tasks.push({id:uid('t'),date:UI.plDate,start:st,end:st?mtime(tmin(st)+m):'',subj:$('#nt-subj').value,type:$('#nt-type').value,text:t,plan:m,act:0,done:false});commit()},
 'task-lock':(e,p)=>{const t=S.tasks.find(x=>x.id===p);t.lock=!t.lock;commit()},'task-edit':(e,p)=>{const t=S.tasks.find(x=>x.id===p),v=prompt('Task',t.text);if(v){t.text=v;t.auto=false;commit()}},
 'task-next':(e,p)=>{const t=S.tasks.find(x=>x.id===p);t.date=addDays(t.date,1);t.start='';t.end='';commit()},'task-del':(e,p)=>{S.tasks=S.tasks.filter(x=>x.id!==p);commit()},
 'rev-subj':(e,p)=>{UI.revSid=p;commit()},'rev-today':(e,p)=>{const[,c]=chById(p);c.last=todayStr();commit()},
 'ch-ren':(e,p)=>{const[,c]=chById(p),v=prompt('Chapter name',c.name);if(v){c.name=v;commit()}},
 'ch-merge':(e,p)=>{const[sid,c]=chById(p),A=S.chapters[sid],i=A.indexOf(c),prev=A[i-1];LL(sid).concat(S.lectures[sid]).forEach(l=>{if(l.ch===c.id)l.ch=prev.id});['r1','r2','r3','pr'].forEach(k=>prev[k]=prev[k]&&c[k]);A.splice(i,1);commit()},
 'ch-del':(e,p)=>{const[sid,c]=chById(p);if(!confirm('Delete this chapter? Its lectures move to the previous chapter.'))return;const A=S.chapters[sid],i=A.indexOf(c),tgt=A[i-1]||A[i+1];if(tgt)S.lectures[sid].forEach(l=>{if(l.ch===c.id)l.ch=tgt.id});A.splice(i,1);commit()},
 'ch-add':(e,p)=>{const v=$('#newch').value.trim();if(!v)return;newChapter(p,v);commit()},
 'prac-add':()=>{const a=+$('#patt').value,c=+$('#pcor').value;if(!(a>0)||c>a)return alert('Enter attempted and correct (correct cannot exceed attempted).');S.practice.push({id:uid('p'),date:$('#pdt').value,subj:$('#psub').value,note:$('#pnote').value.trim(),att:a,cor:c});commit()},'prac-del':(e,p)=>{S.practice=S.practice.filter(x=>x.id!==p);commit()},
 'test-add':()=>{const sc=parseFloat($('#tsc').value),mx=parseFloat($('#tmx').value);if(isNaN(sc)||!(mx>0))return alert('Enter marks and maximum marks.');S.tests.push({id:uid('q'),date:$('#tdt').value,subj:$('#tsub').value,type:$('#ttype').value,name:$('#tname').value.trim(),score:sc,max:mx,min:+$('#tmin').value||0});commit()},'test-del':(e,p)=>{S.tests=S.tests.filter(x=>x.id!==p);commit()},
 'chk-new':()=>{const n=prompt('Checklist name');if(!n)return;const c={id:uid('k'),name:n,notes:'',items:[]};S.checklists.push(c);UI.chk=c.id;commit()},
 'chk-ren':(e,p)=>{const c=S.checklists.find(x=>x.id===p),n=prompt('Name',c.name);if(n){c.name=n;commit()}},'chk-del':(e,p)=>{if(confirm('Delete this checklist?')){S.checklists=S.checklists.filter(x=>x.id!==p);UI.chk='';commit()}},
 'chk-add':(e,p)=>{const t=$('#chk-new-t').value.trim();if(!t)return;const c=S.checklists.find(x=>x.id===p);c.items.push({id:uid('i'),key:String(c.items.length+1),g:$('#chk-new-g').value.trim(),t,ref:'',done:false});commit()},
 'chk-item-del':(e,p)=>{const[c,i]=p.split('|'),k=S.checklists.find(x=>x.id===c);k.items=k.items.filter(x=>x.id!==i);commit()},
 'link-files':async()=>{try{const hs=await showOpenFilePicker({multiple:true,types:[{description:'Excel / CSV',accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx'],'text/csv':['.csv']}}]});const files=[],map={};for(const h of hs){const f=await h.getFile();files.push(f);map[f.name]=h}await handleFiles(files,{handle:map,preview:true})}catch(e){if(e.name!=='AbortError')toast('Could not link: '+e.message,'bad')}},
 'refresh-files':()=>{const i=document.createElement('input');i.type='file';i.multiple=true;i.accept='.xlsx,.csv';i.onchange=()=>handleFiles([...i.files]);i.click()},
 'sync-unlink':(e,p)=>{S.sync.files=S.sync.files.filter(f=>f.name!==p);idbDel(PROF.active+':'+p);delete SYNC.handles[p];commit()},
 'sync-reconnect':async()=>{for(const f of S.sync.files.filter(x=>x.linked)){const h=SYNC.handles[f.name];if(!h)continue;try{if(await h.requestPermission({mode:'read'})==='granted')SYNC.state[f.name]='live'}catch(e){}}commit();pollSync()},
 'export-xlsx':()=>exportXlsx(false),'export-xlsx-subj':()=>exportXlsx(true),
 'export-csv':()=>{const rows=[['Class','Day','Lecture','Topic','Minutes','Status','Downloaded','Date','Priority','Chapter','Notes']];S.subjects.forEach(s=>LL(s.id).forEach(l=>rows.push([s.name,l.l,l.n,l.t,Math.round(l.m*10)/10,l.d?'Completed':'Pending',l.dl===null||l.dl===undefined?'':l.dl?'Yes':'No',l.doneOn||l.sdate||schedule().planned[l.id]||'',l.pri||'',((S.chapters[s.id]||[]).find(c=>c.id===l.ch)||{}).name||'',l.note||''])));download('lectures-'+todayStr()+'.csv',new Blob([toCSV(rows)],{type:'text/csv'}))},
 'export-log':()=>download('study-log-'+todayStr()+'.csv',new Blob([toCSV([['Date','Class','What','Minutes','Source']].concat(studyLogRows()))],{type:'text/csv'})),
 print:()=>printReport(),backup:()=>backup(),'tpl-export':(e,p)=>templateExport(p==='1'),
 'prof-switch':(e,p)=>{flush();PROF.active=p;localStorage.setItem(PROFILES_KEY,JSON.stringify(PROF));bootProfile()},
 'prof-new':()=>{const n=prompt('Profile name (e.g. a friend\'s name)');if(!n)return;flush();const id='p'+Date.now().toString(36);PROF.list.push({id,name:n});PROF.active=id;localStorage.setItem(PROFILES_KEY,JSON.stringify(PROF));bootProfile(true)},
 'prof-ren':()=>{const p=PROF.list.find(x=>x.id===PROF.active),n=prompt('Rename profile',p.name);if(n){p.name=n;localStorage.setItem(PROFILES_KEY,JSON.stringify(PROF));render()}},
 'prof-del':()=>{if(!confirm('Delete this profile and all its data on this device?'))return;localStorage.removeItem(dataKey());PROF.list=PROF.list.filter(x=>x.id!==PROF.active);PROF.active=PROF.list[0].id;localStorage.setItem(PROFILES_KEY,JSON.stringify(PROF));bootProfile()},
 'imp-confirm':()=>{const I=UI.imp;if(!I)return;finishImport(I.files,false)},
 'theme-preset':(e,p)=>{Object.assign(S.settings.theme,PRESETS[p],{preset:p});commit()},
 'subj-add':()=>{const n=prompt('Class name');if(n){ensureSubject(n);commit()}},
 'subj-del':(e,p)=>{if(!confirm('Delete this class and all its lectures and chapters?'))return;S.subjects=S.subjects.filter(s=>s.id!==p);delete S.lectures[p];delete S.chapters[p];S.tasks=S.tasks.filter(t=>t.subj!==p);commit()},
 'exam-add':()=>{S.exams.push({id:uid('e'),name:'New exam',date:addDays(todayStr(),60),note:''});commit()},'exam-del':(e,p)=>{S.exams=S.exams.filter(x=>x.id!==p);commit()},
 'ms-add':()=>{S.milestones.push({id:uid('m'),name:'New milestone',kind:'manual',subjects:'all',date:addDays(todayStr(),30),done:false});commit()},'ms-del':(e,p)=>{S.milestones=S.milestones.filter(x=>x.id!==p);commit()},
 'fixed-add':()=>{S.fixed.push({id:uid('f'),name:'Commitment',start:'17:00',end:'18:00',days:[]});commit()},'fixed-del':(e,p)=>{S.fixed=S.fixed.filter(x=>x.id!==p);commit()},
 notif:async()=>{if(!('Notification' in window))return;await Notification.requestPermission();render()}
};
document.addEventListener('click',e=>{const el=e.target.closest('[data-a]');if(!el)return;const f=ACT[el.dataset.a];if(!f)return;
  if(el.closest('summary')||el.tagName==='BUTTON'||el.dataset.a)e.stopPropagation();
  const d=el.closest('details.menu');if(d)d.open=false;
  try{f(el,el.dataset.p,e)}catch(err){console.error(err);toast('Something went wrong: '+err.message,'bad')}});
const setPath=(obj,path,val)=>{const k=path.split('.');let o=obj;for(let i=0;i<k.length-1;i++)o=o[k[i]];o[k[k.length-1]]=val};
const CHG={
 taskdone:(el,p)=>{setTaskDone(S.tasks.find(t=>t.id===p),el.checked);commit()},
 lecdone:(el,p)=>{setLecDone(lecById(p),el.checked);commit()},lecdl:(el,p)=>{lecById(p).dl=el.checked;commit()},
 'lec-ch':(el,p)=>{lecById(p).ch=el.value;commit()},
 lf:(el,p)=>{UI.lec[p]=el.value;UI.lec.page=0;commit()},
 'pl-date':(el)=>{UI.plDate=el.value||todayStr();commit()},
 'task-f':(el,p)=>{const[id,k]=p.split('|'),t=S.tasks.find(x=>x.id===id);t[k]=k==='start'?el.value:Math.max(0,+el.value||0);if(k==='start'&&el.value)t.end=mtime(tmin(el.value)+(+t.plan||0));if(k==='plan'&&t.start)t.end=mtime(tmin(t.start)+t.plan);if(k==='act'&&t.act>0&&t.done===false&&t.act>=t.plan)t.done=false;t.auto=t.auto&&k==='act';commit()},
 cf:(el,p)=>{const[id,k]=p.split('|'),[,c]=chById(id);c[k]=el.checked;if(el.checked&&k!=='cls'&&k!=='pr')c.last=todayStr();commit()},
 conf:(el,p)=>{chById(p)[1].conf=+el.value;commit()},
 set:(el)=>{const t=el.dataset.t;let v=el.value;if(t==='n')v=+v;else if(t==='b')v=v==='true';else if(t==='list')v=v.split(/[,\s]+/).map(Number).filter(n=>n>0);
   if(el.dataset.p==='settings.goal')v=Math.max(6,v);setPath(S,el.dataset.p,v);if(el.dataset.p.startsWith('settings.theme'))S.settings.theme.preset='Custom';if(el.dataset.p==='settings.syncSec')startSyncLoop();commit()},
 'layout-preset':(el)=>{if(!el.value)return;S.layout={preset:el.value,widgets:LAYOUTS[el.value].map(w=>({...w}))};commit()},
 'pomo-subj':(el)=>{S.pomo.subj=el.value;save()},
 'chk-item':(el,p)=>{const[c,i]=p.split('|');S.checklists.find(x=>x.id===c).items.find(x=>x.id===i).done=el.checked;commit()},
 'chk-g':(el)=>{UI.chkG=el.value;commit()},'chk-q':(el)=>{UI.chkQ=el.value;commit()},
 'imp-subj':(el,p)=>{const[f,s]=p.split('|');UI.imp.files[f].plan[s].subject=el.value.trim()},'imp-skip':(el,p)=>{const[f,s]=p.split('|');UI.imp.files[f].plan[s].skip=el.checked},
 subj:(el,p)=>{const[id,k]=p.split('|'),s=sj(id);s[k]=k==='hide'?el.checked:k==='pri'?+el.value:k==='batchH'?(+el.value||0):el.value;commit()},
 exam:(el,p)=>{const[id,k]=p.split('|');S.exams.find(x=>x.id===id)[k]=el.value;commit()},
 ms:(el,p)=>{const[id,k]=p.split('|'),m=S.milestones.find(x=>x.id===id);m[k]=k==='done'?el.checked:k==='subjects'?(el.value==='all'?'all':[el.value]):el.value;commit()},
 fixed:(el,p)=>{const[id,k]=p.split('|'),f=S.fixed.find(x=>x.id===id);f[k]=k==='days'?el.value.split(/[,\s]+/).map(Number).filter(n=>n>=0&&n<=6&&!isNaN(n)):el.value;commit()}
};
document.addEventListener('change',e=>{const el=e.target;
  if(el.id==='file-in'){handleFiles([...el.files],{preview:true});el.value='';return}
  if(el.id==='restore-in'&&el.files[0]){el.files[0].text().then(t=>{try{const x=JSON.parse(t);if(x.v!==3)throw 0;S=x;fixSid();save();toast('Backup restored.');render()}catch(_){toast('That is not a backup from this dashboard.','bad')}});return}
  if(el.id==='tpl-in'&&el.files[0]){el.files[0].text().then(t=>{try{templateApply(JSON.parse(t));commit();toast('Template applied.')}catch(err){toast(err.message,'bad')}});return}
  const c=el.dataset&&el.dataset.c;if(c&&CHG[c]){try{CHG[c](el,el.dataset.p)}catch(err){console.error(err);toast('Something went wrong: '+err.message,'bad')}}});
let _ls=null;
document.addEventListener('input',e=>{const el=e.target;
  if(el.dataset.i==='lsearch'){clearTimeout(_ls);_ls=setTimeout(()=>{UI.lec.q=el.value;UI.lec.page=0;$('#lec-results').innerHTML=lecResults()},140)}});
document.addEventListener('keydown',e=>{if(e.key!=='Enter')return;const id=e.target.id;
  if(id==='gsearch'){UI.lec.q=e.target.value;UI.lec.status='all';UI.lec.sid='all';UI.lec.view='table';UI.tab='lectures';commit()}
  else if(id==='rem-text')ACT['rem-add']();else if(id==='note-text')ACT['note-add']();else if(id==='newch'){const b=document.querySelector('[data-a="ch-add"]');b&&b.click()}});
document.addEventListener('toggle',e=>{const d=e.target;if(d.matches&&d.matches('details.ch[data-key]')){if(d.open){UI.open.add(d.dataset.key);const b=d.querySelector('.in');if(b&&!b.innerHTML.trim())b.innerHTML=chapterBody(d.dataset.sid,d.dataset.cid)}else UI.open.delete(d.dataset.key)}},true);
/* drag & drop: widgets + files */
let _drag=null;
document.addEventListener('dragstart',e=>{const w=e.target.closest&&e.target.closest('.w[data-wid]');if(!w||!UI.custom)return;_drag=w.dataset.wid;w.classList.add('drag');e.dataTransfer.effectAllowed='move';try{e.dataTransfer.setData('text/plain',_drag)}catch(_){}});
document.addEventListener('dragend',()=>{_drag=null;document.querySelectorAll('.w.drag,.w.over').forEach(x=>x.classList.remove('drag','over'))});
document.addEventListener('dragover',e=>{const w=e.target.closest&&e.target.closest('.w[data-wid]');if(_drag&&w){e.preventDefault();document.querySelectorAll('.w.over').forEach(x=>x.classList.remove('over'));if(w.dataset.wid!==_drag)w.classList.add('over')}
  const dz=e.target.closest&&e.target.closest('#drop');if(dz&&e.dataTransfer&&[...e.dataTransfer.types].includes('Files')){e.preventDefault();dz.classList.add('on')}});
document.addEventListener('dragleave',e=>{const dz=e.target.closest&&e.target.closest('#drop');if(dz)dz.classList.remove('on')});
document.addEventListener('drop',e=>{
  const w=e.target.closest&&e.target.closest('.w[data-wid]');
  if(_drag&&w){e.preventDefault();const A=S.layout.widgets,from=A.findIndex(x=>x.id===_drag),to=A.findIndex(x=>x.id===w.dataset.wid);if(from>=0&&to>=0&&from!==to){const[m]=A.splice(from,1);A.splice(to,0,m);S.layout.preset='Custom';commit()}return}
  if(e.dataTransfer&&e.dataTransfer.files&&e.dataTransfer.files.length){e.preventDefault();UI.tab='data';handleFiles([...e.dataTransfer.files],{preview:true})}});
window.addEventListener('dragover',e=>{if(e.dataTransfer&&[...e.dataTransfer.types].includes('Files'))e.preventDefault()});

/* ---------- boot ---------- */
function b64buf(b){const s=atob(b),u=new Uint8Array(s.length);for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i);return u.buffer}
async function seedFromFiles(){
  const seed=window.SEED;S=defaults();
  if(!seed){fixSid();return}
  S.exams=seed.exams||S.exams;if(seed.settings)Object.assign(S.settings,seed.settings);
  for(const f of seed.files){
    const wb=await readXlsx(b64buf(f.b64)),parsed=parseWorkbook(wb),plan=parsed.map((p,i)=>{const sp=f.sheets&&f.sheets[p.name];if(p.type==='info')return{skip:true};
      return p.type==='lectures'?{subject:sp.subject,extra:sp.extra,chList:sp.chList}:{subject:sp.subject}});
    const log=commitImport(f.name,parsed,plan);
    const rec={name:f.name,map:{},linked:false,lastSync:Date.now()};parsed.forEach((p,i)=>{if(plan[i]&&!plan[i].skip)rec.map[p.name]=plan[i].subject});S.sync.files.push(rec);
  }
  S.milestones=(seed.milestones||[]).map(m=>({...m,id:uid('m'),subjects:m.subjects==='all'?'all':m.subjects.map(n=>(S.subjects.find(s=>s.name===n)||{}).id).filter(Boolean)}));
  (seed.extraSubjects||[]).forEach(x=>{const s=ensureSubject(x.name,x.extra);(x.chapters||[]).forEach(n=>newChapter(s.id,n))});
  S.subjects.sort((a,b)=>(seed.order||[]).indexOf(a.name)-(seed.order||[]).indexOf(b.name));
  fixSid();S.meta.seeded=true;save();
}
async function bootProfile(fresh){
  let ok=!fresh&&load();
  if(!ok){document.getElementById('app').innerHTML='<div style="padding:60px;text-align:center;color:var(--muted)">Setting up your dashboard…</div>';try{await seedFromFiles()}catch(e){console.error(e);S=S||defaults();toast('Setup problem: '+e.message,'bad')}}
  else fixSid();
  if(!S.layout||!S.layout.widgets)S.layout={preset:'Everything',widgets:LAYOUTS.Everything.map(w=>({...w}))};
  UI.plDate=todayStr();UI.lec.sid='all';UI.revSid='';
  applyTheme();render();flush();
  if(LOAD_PROBLEM){toast(LOAD_PROBLEM,'bad');LOAD_PROBLEM=''}
  await restoreHandles();render();startSyncLoop();
  if(!ok&&window.SEED){const n=Object.values(S.lectures).flat().length;toast(`Loaded ${n} lecture rows from ${S.sync.files.length} Excel files – see Data & sync → Import log for the row-by-row check.`)}
}
/* ---------- PWA: install, offline, storage, updates ---------- */
const PWA={ev:null,sw:'none',waiting:null,persisted:null,installed:matchMedia('(display-mode: standalone)').matches,est:null};
window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();PWA.ev=e;if(S)render()});
window.addEventListener('appinstalled',()=>{PWA.ev=null;PWA.installed=true;toast('Installed. You can now open it from the Start menu.');if(S)render()});
async function pwaInit(){
  try{if(navigator.storage&&navigator.storage.persist){PWA.persisted=await navigator.storage.persisted()||await navigator.storage.persist()}}catch(e){}
  try{if(navigator.storage&&navigator.storage.estimate){const e=await navigator.storage.estimate();PWA.est=e}}catch(e){}
  if(!('serviceWorker' in navigator)||!/^https?:$/.test(location.protocol)){PWA.sw='unavailable';return}
  try{
    const reg=await navigator.serviceWorker.register('./sw.js',{scope:'./'});
    PWA.sw=reg.active?'ready':'installing';
    const watch=w=>w.addEventListener('statechange',()=>{if(w.state==='installed'&&navigator.serviceWorker.controller){PWA.waiting=w;toast('A new version is ready. Open Data & sync → App & offline to update.');if(S)render()}else if(w.state==='activated'){PWA.sw='ready';if(S)render()}});
    if(reg.waiting&&navigator.serviceWorker.controller)PWA.waiting=reg.waiting;
    if(reg.installing)watch(reg.installing);reg.addEventListener('updatefound',()=>watch(reg.installing));
    navigator.serviceWorker.ready.then(()=>{PWA.sw='ready';if(S)render()});
    let reloading=false;navigator.serviceWorker.addEventListener('controllerchange',()=>{if(reloading)return;reloading=true;flush();location.reload()});
  }catch(e){PWA.sw='error: '+e.message}
}
Object.assign(ACT,{'pwa-install':async()=>{if(!PWA.ev)return;PWA.ev.prompt();await PWA.ev.userChoice;PWA.ev=null;render()},
 'pwa-update':()=>{flush();if(PWA.waiting)PWA.waiting.postMessage('skip-waiting')}});
document.addEventListener('DOMContentLoaded',async()=>{
  const q=new URLSearchParams(location.search).get('tab');
  await bootProfile();
  if(q&&TABS&&TABS.some&&TABS.some(t=>t[0]===q)){UI.tab=q;render()}
  pwaInit()});

