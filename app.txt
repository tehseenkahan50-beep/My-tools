
const $=s=>document.querySelector(s);
const kb=n=>n>1e6?(n/1e6).toFixed(1)+' MB':Math.max(1,Math.round(n/1e3))+' KB';
const base=n=>n.replace(/\.[^.]+$/,'');
const loadImg=f=>new Promise((ok,no)=>{const i=new Image();i.onload=()=>ok(i);i.onerror=()=>no(new Error("Can't open "+f.name+'. Use a JPG, PNG or WebP image.'));i.src=URL.createObjectURL(f)});
const toJpg=async(f,max,q)=>{const i=await loadImg(f),s=Math.min(1,max/Math.max(i.width,i.height)),c=document.createElement('canvas');
 c.width=Math.round(i.width*s);c.height=Math.round(i.height*s);const x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);x.drawImage(i,0,0,c.width,c.height);
 return new Promise(r=>c.toBlob(r,'image/jpeg',q))};

const CD='https://cdnjs.cloudflare.com/ajax/libs/',libs={};
const lib=u=>libs[u]||(libs[u]=new Promise((ok,no)=>{const s=document.createElement('script');s.src=u;s.onload=ok;s.onerror=()=>{delete libs[u];no(new Error('Could not load the converter. Check your internet connection.'))};document.head.appendChild(s)}));
const esc=s=>s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const X='<?xml version="1.0" encoding="UTF-8"?>';

const arab=s=>(s.match(/[\u0600-\u06FF\u0750-\u077F]/g)||[]).length>s.length*.3;
const docxFrom=async ps=>{
 const body=ps.map(p=>{if(p===null)return'<w:p><w:r><w:br w:type="page"/></w:r></w:p>';const a=arab(p);
  return'<w:p>'+(a?'<w:pPr><w:bidi/></w:pPr>':'')+'<w:r>'+(a?'<w:rPr><w:rtl/></w:rPr>':'')+'<w:t xml:space="preserve">'+esc(p)+'</w:t></w:r></w:p>'}).join(''),z=new JSZip();
 z.file('[Content_Types].xml',X+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
 z.file('_rels/.rels',X+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
 z.file('word/document.xml',X+'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'+body+'</w:body></w:document>');
 return z.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'})};

const TOOLS={
 jpg2pdf:{from:'JPG',to:'PDF',name:'JPG to PDF',desc:'Combine photos into one PDF.',accept:'image/*',min:1,hint:'JPG, PNG or WebP. Pages follow the order you pick.',
  async run(fs){const{PDFDocument}=PDFLib,d=await PDFDocument.create();
   for(const f of fs){const b=await toJpg(f,2400,.9),im=await d.embedJpg(await b.arrayBuffer()),p=d.addPage([im.width,im.height]);p.drawImage(im,{x:0,y:0,width:im.width,height:im.height})}
   return[{blob:new Blob([await d.save()],{type:'application/pdf'}),name:'fileflip-images.pdf'}]}},
 merge:{from:'PDF',to:'PDF',name:'Merge PDF',desc:'Join several PDFs into one file.',accept:'application/pdf',min:2,hint:'Pick at least two PDFs. They are joined in the order you pick.',
  async run(fs){const{PDFDocument}=PDFLib,d=await PDFDocument.create();
   for(const f of fs){const s=await PDFDocument.load(await f.arrayBuffer());(await d.copyPages(s,s.getPageIndices())).forEach(p=>d.addPage(p))}
   return[{blob:new Blob([await d.save()],{type:'application/pdf'}),name:'fileflip-merged.pdf'}]}},
 compress:{from:'IMG',to:'Smaller',name:'Compress image',desc:'Cut photo size, keep it sharp.',accept:'image/*',min:1,hint:'JPG, PNG or WebP. Output is a lighter JPG.',
  async run(fs){const r=[];for(const f of fs){const b=await toJpg(f,2000,.72);r.push({blob:b,name:base(f.name)+'-small.jpg',note:kb(f.size)+' → '+kb(b.size)})}return r}},
 xlsx:{from:'XLSX',to:'CSV',name:'Excel to CSV',desc:'Turn a spreadsheet into CSV.',accept:'.xlsx,.xls',min:1,hint:'XLSX or XLS. Each sheet becomes its own CSV.',
  async run(fs){const r=[];for(const f of fs){const w=XLSX.read(await f.arrayBuffer());
   w.SheetNames.forEach(n=>r.push({blob:new Blob([XLSX.utils.sheet_to_csv(w.Sheets[n])],{type:'text/csv'}),name:base(f.name)+(w.SheetNames.length>1?'-'+n:'')+'.csv'}))}return r}},
 word2pdf:{from:'WORD',to:'PDF',name:'Word to PDF',desc:'Turn a DOCX file into a PDF.',accept:'.docx',min:1,hint:'DOCX only. Simple documents convert best; the PDF text is not selectable.',
  async run(fs){await lib(CD+'mammoth/1.6.0/mammoth.browser.min.js');await lib(CD+'html2pdf.js/0.10.1/html2pdf.bundle.min.js');const r=[];
   for(const f of fs){const h=(await mammoth.convertToHtml({arrayBuffer:await f.arrayBuffer()})).value,d=document.createElement('div');
    d.style.cssText='width:680px;padding:24px;font:14px/1.5 Arial,sans-serif;color:#000;background:#fff;position:fixed;left:-9999px;top:0';
    d.innerHTML='<style>table{border-collapse:collapse}td,th{border:1px solid #999;padding:4px}img{max-width:100%}</style>'+h;document.body.appendChild(d);
    try{const b=await html2pdf().set({margin:12,html2canvas:{scale:2},jsPDF:{unit:'mm',format:'a4'}}).from(d).outputPdf('blob');r.push({blob:b,name:base(f.name)+'.pdf'})}finally{d.remove()}}
   return r}},
 pdf2word:{from:'PDF',to:'WORD',name:'PDF to Word',desc:'Get the text of a PDF as DOCX.',accept:'application/pdf',min:1,hint:'PDFs with selectable text. Only text is kept, not layout, tables or images. For scans or Urdu PDFs use Scan to Word (OCR).',
  async run(fs){await lib(CD+'pdf.js/3.11.174/pdf.min.js');await lib(CD+'jszip/3.10.1/jszip.min.js');
   pdfjsLib.GlobalWorkerOptions.workerSrc=CD+'pdf.js/3.11.174/pdf.worker.min.js';const r=[];
   for(const f of fs){const pdf=await pdfjsLib.getDocument({data:await f.arrayBuffer()}).promise,ps=[];
    for(let n=1;n<=pdf.numPages;n++){const tc=await(await pdf.getPage(n)).getTextContent();let y=null,ln='';
     for(const it of tc.items){const yy=it.transform[5];if(y!==null&&Math.abs(yy-y)>3){if(ln.trim())ps.push(ln.trim());ln=''}ln+=it.str;y=yy}
     if(ln.trim())ps.push(ln.trim());if(n<pdf.numPages)ps.push(null)}
    const all=ps.filter(Boolean).join('');
    if(!all)throw new Error(f.name+' has no selectable text. Use Scan to Word (OCR) instead.');
    const odd=(all.match(/[\u00A1-\u00BF\u00C0-\u00FF]/g)||[]).length;
    if(all.length>40&&odd/all.length>.15&&!/[\u0600-\u06FF]/.test(all))throw new Error(f.name+' uses a custom font (common in Urdu PDFs), so its text would come out garbled. Use Scan to Word (OCR) instead.');
    r.push({blob:await docxFrom(ps),name:base(f.name)+'.docx'})}
   return r}},
 ocr:{from:'SCAN',to:'WORD',name:'Scan to Word (OCR)',desc:'Read text from scans and photos.',accept:'application/pdf,image/*',min:1,opts:true,hint:'Scanned PDFs, photos and Urdu PDFs. Try one page first, it is slow on phones.',
  async run(fs){const L=$('#lang').value.split('+');
   await lib('https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js');await lib(CD+'pdf.js/3.11.174/pdf.min.js');await lib(CD+'jszip/3.10.1/jszip.min.js');
   pdfjsLib.GlobalWorkerOptions.workerSrc=CD+'pdf.js/3.11.174/pdf.worker.min.js';
   const items=[];
   for(const f of fs){if(f.type==='application/pdf'){const pdf=await pdfjsLib.getDocument({data:await f.arrayBuffer()}).promise;
     for(let n=1;n<=pdf.numPages;n++){const pg=await pdf.getPage(n),vp=pg.getViewport({scale:2.5}),c=document.createElement('canvas');c.width=vp.width;c.height=vp.height;await pg.render({canvasContext:c.getContext('2d'),viewport:vp}).promise;items.push({f,src:c})}}
    else items.push({f,src:f})}
   if(items.length>15)throw new Error('Too many pages ('+items.length+'). Use 15 pages or fewer at a time.');
   const w=await Tesseract.createWorker(L),r=[];let done=0;
   try{for(const f of fs){const ps=[];
     for(const it of items.filter(i=>i.f===f)){$('#msg').textContent='Reading page '+(++done)+' of '+items.length+'…';
      const t=(await w.recognize(it.src)).data.text;
      if(ps.length)ps.push(null);
      ps.push(...t.split(/\n\s*\n/).map(s=>s.replace(/\s*\n\s*/g,' ').trim()).filter(Boolean))}
     if(!ps.some(Boolean))throw new Error('No text found in '+f.name+'.');
     r.push({blob:await docxFrom(ps),name:base(f.name)+'-ocr.docx'})}
   }finally{await w.terminate()}
   return r}}
};

const SLUG={jpg2pdf:'jpg-to-pdf',merge:'merge-pdf',compress:'compress-image',xlsx:'excel-to-csv',word2pdf:'word-to-pdf',pdf2word:'pdf-to-word',ocr:'scan-to-word'};
let cur=null,files=[];
const grid=$('#tools');
if(grid)for(const[k,t]of Object.entries(TOOLS)){const a=document.createElement('a');a.className='tile';a.href='/'+SLUG[k];a.innerHTML='<div class="fmt"><span>'+t.from+'</span><i aria-hidden="true">⇄</i><span>'+t.to+'</span></div><h2>'+t.name+'</h2><p>'+t.desc+'</p>';grid.appendChild(a)}
if($('#ws')){
 cur=TOOLS[document.body.dataset.tool];
 const msg=(t,err)=>{const m=$('#msg');m.textContent=t||'';m.className=err?'err':''};
 const render=()=>{
  $('#fl').innerHTML=files.map(f=>'<li><span>'+f.name.replace(/</g,'&lt;')+'</span><span>'+kb(f.size)+'</span></li>').join('');
  $('#go').disabled=files.length<cur.min;$('#out').innerHTML='';
  msg(files.length&&files.length<cur.min?'Add at least '+cur.min+' files to continue.':'');
 };
 const add=l=>{files=files.concat([...l]);render()};
 const fi=$('#fi'),dz=$('#dz'),go=$('#go');
 fi.accept=cur.accept;fi.multiple=true;$('#hint').textContent=cur.hint;
 if(cur.opts)dz.insertAdjacentHTML('beforebegin','<label class="lang">Language <select id="lang"><option value="urd+eng">Urdu + English</option><option value="urd">Urdu</option><option value="eng">English</option><option value="ara">Arabic</option></select></label>');
 fi.onchange=e=>{add(e.target.files);fi.value=''};
 ['dragenter','dragover'].forEach(v=>dz.addEventListener(v,e=>{e.preventDefault();dz.classList.add('on')}));
 ['dragleave','drop'].forEach(v=>dz.addEventListener(v,e=>{e.preventDefault();dz.classList.remove('on')}));
 dz.addEventListener('drop',e=>add(e.dataTransfer.files));
 go.onclick=async()=>{
  go.disabled=true;msg('Working…');$('#out').innerHTML='';
  try{
   const res=await cur.run(files);
   $('#out').innerHTML=res.map(r=>{const u=URL.createObjectURL(r.blob);return'<a class="btn alt" href="'+u+'" download="'+r.name+'">Download '+r.name+(r.note?' ('+r.note+')':'')+'</a>'}).join('');
   msg('Done. Your files never left this device.');
  }catch(e){msg('Something went wrong: '+(e.message||'unknown error')+'. Check the file and try again.',true)}
  go.disabled=files.length<cur.min;
 };
 render();
}
