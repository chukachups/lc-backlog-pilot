(() => {
  "use strict";
  const model = globalThis.Backlog;
  const enc = new TextEncoder(), dec = new TextDecoder();
  const xmlHead = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const officeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const pkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";
  const spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const types = "http://schemas.openxmlformats.org/package/2006/content-types";
  const XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  const workbookLimit = 450 * 1024;
  const tag = (root, name) => [...root.getElementsByTagNameNS("*", name)];
  const esc = value => String(value).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]/g,"")
    .replace(/[&<>"']/g, char => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&apos;"})[char]);

  function encode64(bytes) {
    let binary = "";
    for (let i=0; i<bytes.length; i+=8192) binary += String.fromCharCode(...bytes.subarray(i,i+8192));
    return btoa(binary);
  }
  const decode64 = value => Uint8Array.from(atob(value), char => char.charCodeAt(0));
  const crcTable = Uint32Array.from({length:256}, (_, n) => {
    for (let i=0; i<8; i++) n = (n >>> 1) ^ ((n & 1) ? 0xedb88320 : 0);
    return n;
  });
  function crc32(bytes) {
    let n = 0xffffffff;
    for (const byte of bytes) n = crcTable[(n ^ byte) & 255] ^ (n >>> 8);
    return (n ^ 0xffffffff) >>> 0;
  }
  function block(length) {
    const bytes = new Uint8Array(length);
    return {bytes, view:new DataView(bytes.buffer)};
  }
  const stored = text => {
    const bytes = typeof text === "string" ? enc.encode(text) : text;
    return {bytes, method:0, size:bytes.length, crc:crc32(bytes)};
  };
  function pack(files, mime) {
    const chunks = [], central = [];
    let offset = 0, centralSize = 0;
    for (const [name, file] of files) {
      const entry = typeof file === "string" ? stored(file) : file;
      const nameData = enc.encode(name);
      const local = block(30+nameData.length), index = block(46+nameData.length);
      local.view.setUint32(0,0x04034b50,true);
      index.view.setUint32(0,0x02014b50,true);
      index.view.setUint16(4,20,true);
      for (const [header, pos] of [[local,4],[index,6]]) {
        header.view.setUint16(pos,20,true);
        header.view.setUint16(pos+2,0x800,true);
        header.view.setUint16(pos+4,entry.method,true);
        header.view.setUint16(pos+8,33,true);
        header.view.setUint32(pos+10,entry.crc,true);
        header.view.setUint32(pos+14,entry.bytes.length,true);
        header.view.setUint32(pos+18,entry.size,true);
        header.view.setUint16(pos+22,nameData.length,true);
      }
      local.bytes.set(nameData,30);
      index.bytes.set(nameData,46);
      index.view.setUint32(42,offset,true);
      chunks.push(local.bytes,entry.bytes);
      central.push(index.bytes);
      offset += local.bytes.length + entry.bytes.length;
      centralSize += index.bytes.length;
    }
    const end = block(22);
    end.view.setUint32(0,0x06054b50,true);
    end.view.setUint16(8,central.length,true);
    end.view.setUint16(10,central.length,true);
    end.view.setUint32(12,centralSize,true);
    end.view.setUint32(16,offset,true);
    return new Blob([...chunks,...central,end.bytes],{type:mime});
  }
  function archive(bytes) {
    const view = new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);
    const short = pos => view.getUint16(pos,true), long = pos => view.getUint32(pos,true);
    let end = -1;
    for (let i=bytes.length-22; i>=Math.max(0,bytes.length-65557); i--) {
      if (long(i)===0x06054b50 && i+22+short(i+20)===bytes.length) { end=i; break; }
    }
    if (end<0 || short(end+4) || short(end+6)) throw Error("Выберите обычный файл Excel в формате .xlsx.");
    const entries = new Map(), count = short(end+10);
    let cursor = long(end+16);
    if (count>4000) throw Error("В архиве Excel слишком много файлов.");
    for (let n=0; n<count; n++) {
      if (cursor+46>end || long(cursor)!==0x02014b50) throw Error("Повреждён архив Excel.");
      const method=short(cursor+10), size=long(cursor+24), packed=long(cursor+20), crc=long(cursor+16);
      const start=long(cursor+42), nameSize=short(cursor+28), next=cursor+46+nameSize+short(cursor+30)+short(cursor+32);
      const name=dec.decode(bytes.subarray(cursor+46,cursor+46+nameSize));
      if (next>end || start+30>bytes.length || long(start)!==0x04034b50 ||
          short(cursor+8)&1 || ![0,8].includes(method) || entries.has(name)) throw Error("Архив Excel повреждён или защищён.");
      const data=start+30+short(start+26)+short(start+28);
      if (data+packed>bytes.length) throw Error("Лист Excel загрузился не полностью.");
      entries.set(name,{bytes:bytes.subarray(data,data+packed),method,size,crc});
      cursor=next;
    }
    return entries;
  }
  async function readEntry(entries, name) {
    const entry=entries.get(name);
    if (!entry) throw Error("В книге Excel отсутствует часть " + name + ".");
    if (entry.size>12*1024*1024) throw Error("Одна из частей книги слишком большая.");
    if (entry.method===0) {
      if (entry.bytes.length!==entry.size || crc32(entry.bytes)!==entry.crc) throw Error("Архив Excel повреждён.");
      return dec.decode(entry.bytes);
    }
    let inflate;
    try { inflate=new DecompressionStream("deflate-raw"); }
    catch { throw Error("Импорт Excel поддерживается в актуальных Edge и Chrome."); }
    const reader=new Blob([entry.bytes]).stream().pipeThrough(inflate).getReader();
    const bytes=new Uint8Array(entry.size);
    let offset=0;
    try {
      while (true) {
        const chunk=await reader.read();
        if (chunk.done) break;
        if (offset+chunk.value.length>bytes.length) throw Error("Архив Excel повреждён.");
        bytes.set(chunk.value,offset);
        offset+=chunk.value.length;
      }
    } finally { await reader.cancel(); }
    if (offset!==bytes.length || crc32(bytes)!==entry.crc) throw Error("Архив Excel повреждён.");
    return dec.decode(bytes);
  }
  function documentPart(raw) {
    const xml=new DOMParser().parseFromString(raw,"application/xml");
    if (tag(xml,"parsererror").length || xml.doctype) throw Error("В книге Excel повреждён XML.");
    return xml;
  }
  async function loadSheets(bytes) {
    const entries=archive(bytes);
    const xml=async path=>documentPart(await readEntry(entries,path));
    const [book,relations,shared] = await Promise.all([
      xml("xl/workbook.xml"),xml("xl/_rels/workbook.xml.rels"),
      entries.has("xl/sharedStrings.xml") ? xml("xl/sharedStrings.xml") : null
    ]);
    const links=new Map(tag(relations,"Relationship").map(r=>[r.getAttribute("Id"),r.getAttribute("Target")]));
    const named=new Map(tag(book,"sheet").map(sheet=>{
      const id=[...sheet.attributes].find(attr=>attr.localName==="id")?.value;
      const target=links.get(id) || "";
      const base=new URL("https://local.invalid/xl/workbook.xml");
      const path=new URL(target,base);
      if (!target || path.origin!==base.origin) throw Error("В книге найдена внешняя ссылка на лист.");
      return [sheet.getAttribute("name"),decodeURIComponent(path.pathname.slice(1))];
    }));
    const rich=node=>node?[...node.children].filter(c=>["t","r"].includes(c.localName))
      .map(c=>c.localName==="t"?c.textContent:tag(c,"t").map(t=>t.textContent).join("")).join(""):"";
    const strings=shared ? tag(shared,"si").map(rich) : [];
    const cellText = cell => {
      const value=tag(cell,"v")[0]?.textContent;
      const type=cell.getAttribute("t");
      return model.text(type==="s" ? strings[Number(value)] : type==="inlineStr" ? rich(tag(cell,"is")[0]) : value);
    };
    const sheets=new Map();
    await Promise.all([...model.sheets,"Гант"].map(async name=>{
      const path=named.get(name);
      if (!path) throw Error("В Excel отсутствует лист «"+name+"».");
      const doc=await xml(path);
      const rows=tag(doc,"row").map(node=>{
        const cells=new Map(tag(node,"c").map(cell=>[/^[A-Z]+/.exec(cell.getAttribute("r") || "")?.[0],cell]));
        return {node,num:Number(node.getAttribute("r")),cells,get:column=>cells.has(column)?cellText(cells.get(column)):""};
      });
      sheets.set(name,{doc,path,rows});
    }));
    return {entries,sheets};
  }
  function tableInfo(rows, sheet) {
    const header=rows.find(r=>r.num===1);
    const value=column=>model.normalize(header?.get(column));
    const isClarification=sheet==="Уточняю";
    if (value("B")!==(isClarification?"идея":"задача")) throw Error("Проверьте заголовок листа «"+sheet+"».");
    const steps=[...header.cells.keys()].find(column=>column && value(column).startsWith("шаги") && !value(column).includes("jira"));
    if (!steps) throw Error("На листе «"+sheet+"» не найдена колонка «Шаги».");
    return {steps,owner:isClarification?"F":"E",status:isClarification?"K":"J",
      priority:isClarification?"J":"G",jira:isClarification?"M":"L",plan:isClarification?"":"K"};
  }
  function ganttRows(rows) {
    const header=rows.find(r=>r.get("A")==="№" && model.normalize(r.get("B"))==="задача");
    if (!header) throw Error("На листе «Гант» не найдены кварталы.");
    const columns=[...header.cells.keys()].filter(col=>col && /^[1-4]\s*кв\.?\s*(?:20)?\d{2}$/i.test(header.get(col)));
    if (!columns.length) throw Error("На листе «Гант» не найдены кварталы.");
    const legend={...model.legend};
    rows.filter(r=>r.num<header.num).forEach(row=>{
      if (/^[ОМРИ]$/.test(row.get("A")) && row.get("B")) legend[row.get("A")]=row.get("B");
    });
    const marks=new Map();
    rows.filter(r=>r.num>header.num).forEach(row=>{
      const match=/^(\d+)\s+(FTE|ТМЗ)$/i.exec(row.get("A"));
      if (match) marks.set(match[2].toUpperCase()+" "+match[1],columns.map(c=>row.get(c)));
    });
    return {quarters:columns.map(c=>header.get(c)),legend,marks};
  }
  function workbookIdentity(row, sheet) {
    const number=/^\d+/.exec(row.get("A"))?.[0] || "";
    const title=row.get("B");
    if (!title || (!number && sheet!=="Уточняю")) return null;
    const id=sheet+(number?" "+number:" · "+model.normalize(title));
    return {id,title,number,code:sheet+(number?" "+number:"")};
  }
  async function importExcel(file) {
    if (file.size>workbookLimit) throw Error("Книга превышает 450 КБ. Выберите файл беклога из четырёх листов.");
    const bytes=new Uint8Array(await file.arrayBuffer()), {sheets}=await loadSheets(bytes);
    const gantt=ganttRows(sheets.get("Гант").rows);
    const anchor=model.today();
    const modified=file.lastModified && new Date(file.lastModified).toISOString().slice(0,10);
    const tasks=[];
    for (const name of model.sheets) {
      const rows=sheets.get(name).rows, fields=tableInfo(rows,name);
      for (const row of rows.filter(r=>r.num>1)) {
        const identity=workbookIdentity(row,name);
        if (!identity) continue;
        tasks.push({...identity,sheet:name,owner:row.get(fields.owner),status:row.get(fields.status),
          priority:row.get(fields.priority),jira:row.get(fields.jira),plan:row.get(fields.plan),
          result:"",resultOn:"",excelRow:row.num,marks:gantt.marks.get(identity.id)||[],
          steps:model.parseSteps(row.get(fields.steps),model.validDate(modified)?modified:anchor)});
      }
    }
    if (!tasks.length) throw Error("В книге нет строк с задачами.");
    return model.validate({schema:1,kind:"excel",name:file.name,tasks,quarters:gantt.quarters,
      legend:gantt.legend,workbook:{name:file.name,base64:encode64(bytes)},report:null});
  }
  const columnNumber=letters=>[...letters].reduce((sum,letter)=>sum*26+letter.charCodeAt(0)-64,0);
  function stepCell(sheet, row, column, value) {
    const ns=sheet.doc.documentElement.namespaceURI;
    const cell=row.cells.get(column) || sheet.doc.createElementNS(ns,"c");
    if (!cell.parentNode) {
      cell.setAttribute("r",column+row.num);
      const next=[...row.cells].find(([name])=>name && columnNumber(name)>columnNumber(column));
      row.node.insertBefore(cell,next?.[1] ?? null);
    }
    if (tag(cell,"f").length) throw Error("В ячейке "+column+row.num+" есть формула. Нужна текстовая колонка для шагов.");
    for (const child of [...cell.children]) if (["v","is"].includes(child.localName)) child.remove();
    cell.setAttribute("t","inlineStr");
    const raw=sheet.doc.createElementNS(ns,"is"), text=sheet.doc.createElementNS(ns,"t");
    text.setAttributeNS("http://www.w3.org/XML/1998/namespace","xml:space","preserve");
    text.textContent=value.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]/g,"");
    raw.appendChild(text);
    cell.insertBefore(raw,cell.firstElementChild);
  }
  async function exportExcel(state) {
    if (state.kind==="demo") return demoExcel(state);
    const {entries,sheets}=await loadSheets(decode64(state.workbook.base64));
    for (const name of model.sheets) {
      const sheet=sheets.get(name), {steps}=tableInfo(sheet.rows,name);
      const rows=new Map(sheet.rows.map(row=>[row.num,row]));
      let changed=false;
      for (const task of state.tasks.filter(t=>t.sheet===name)) {
        const row=rows.get(task.excelRow);
        if (!row || workbookIdentity(row,name)?.id!==task.id) throw Error("Строка задачи "+task.code+" не совпала с исходной книгой.");
        if (!task.steps.length && !row.get(steps)) continue;
        const value=model.formatSteps(task.steps);
        if (value.length>32767) throw Error("Шаги задачи "+task.code+" превышают предел одной ячейки Excel.");
        if (row.get(steps)===value) continue;
        stepCell(sheet,row,steps,value); changed=true;
      }
      if (changed) entries.set(sheet.path,stored(xmlHead+new XMLSerializer().serializeToString(sheet.doc.documentElement)));
    }
    return pack(entries,XLSX);
  }

  const rel=(id,kind,target)=>'<Relationship Id="'+id+'" Type="'+officeRel+'/'+kind+'" Target="'+target+'"/>';
  const wrapRels=items=>xmlHead+'<Relationships xmlns="'+pkgRel+'">'+items.join("")+'</Relationships>';
  const wrapTypes=overrides=>xmlHead+'<Types xmlns="'+types+'"><Default Extension="xml" ContentType="application/xml"/>'+
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+overrides.join("")+'</Types>';
  const override=(part,type)=>'<Override PartName="/'+part+'" ContentType="'+type+'"/>';
  function demoExcel(state) {
    const nameList=["FTE","ТМЗ","Гант","Уточняю"];
    const fields=["№","Задача","Идея","Автор идеи","Исполнитель ОАТЗ","Сложность","Приоритет","Экономика","Сроки","Статус","План","Задача в Jira"];
    const clar=["№ в конфе","Идея","Кем предложена","Автор идеи","Комментарий","Исполнитель ОАТЗ","Приоритет из конфы","Сложность","Экономика","Приоритет","Статус","Сроки","Задача в Jira","Шаги"];
    const rowsBySheet=nameList.map(name=>{
      const tasks=state.tasks.filter(t=>t.sheet===name);
      if (name==="Гант") return [...Object.entries(state.legend),[],["№","Задача","Приоритет",...state.quarters],
        ...state.tasks.filter(t=>t.marks.length).map(t=>[t.number+" "+t.sheet,t.title,t.priority,...t.marks])];
      if (name==="Уточняю") return [clar,...tasks.map(t=>[t.number,t.title,"","","",t.owner,"","","",t.priority,t.status,"",t.jira,model.formatSteps(t.steps)])];
      const columns=name==="FTE"?[...fields,"Записи в задаче Jira","Шаги, что реально делается"]:[...fields,"Шаги, что реально делается"];
      return [columns,...tasks.map(t=>[t.number,t.title,"","",t.owner,"",t.priority,"","",t.status,t.plan,t.jira,
        ...(name==="FTE"?[""]:[]),model.formatSteps(t.steps)])];
    });
    const table=rows=>xmlHead+'<worksheet xmlns="'+spreadsheet+'"><sheetData>'+rows.map((row,i)=>
      '<row r="'+(i+1)+'">'+row.map((v,j)=>v===""?"":'<c r="'+String.fromCharCode(65+j)+(i+1)+'" t="inlineStr"><is><t xml:space="preserve">'+esc(v)+'</t></is></c>').join("")+'</row>').join("")+'</sheetData></worksheet>';
    const files=new Map([
      ["[Content_Types].xml",wrapTypes([override("xl/workbook.xml","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
        ...nameList.map((_,i)=>override("xl/worksheets/sheet"+(i+1)+".xml","application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"))])],
      ["_rels/.rels",wrapRels([rel("rId1","officeDocument","xl/workbook.xml")])],
      ["xl/workbook.xml",xmlHead+'<workbook xmlns="'+spreadsheet+'" xmlns:r="'+officeRel+'"><sheets>'+
        nameList.map((n,i)=>'<sheet name="'+esc(n)+'" sheetId="'+(i+1)+'" r:id="rId'+(i+1)+'"/>').join("")+'</sheets></workbook>'],
      ["xl/_rels/workbook.xml.rels",wrapRels(nameList.map((_,i)=>rel("rId"+(i+1),"worksheet","worksheets/sheet"+(i+1)+".xml")))],
      ...rowsBySheet.map((rows,i)=>["xl/worksheets/sheet"+(i+1)+".xml",table(rows)])
    ]);
    return pack(files,XLSX);
  }
  function exportWord(report) {
    const headings=new Set(Object.values(model.sections));
    let first=true;
    const paragraphs=report.split(/\r?\n/).filter(line=>line.trim()).map(line=>{
      const title=first, heading=headings.has(line.trim());
      first=false;
      const emphasis=title||heading;
      return '<w:p><w:pPr>'+(emphasis?'<w:keepNext/>':'')+'<w:spacing w:before="'+(heading?220:0)+'" w:after="150" w:line="276" w:lineRule="auto"/></w:pPr>'+
        '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'+(emphasis?'<w:b/>':'')+
        '<w:sz w:val="'+(title?36:heading?28:22)+'"/></w:rPr><w:t xml:space="preserve">'+esc(line)+'</w:t></w:r></w:p>';
    }).join("");
    return pack(new Map([
      ["[Content_Types].xml",wrapTypes([override("word/document.xml","application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")])],
      ["_rels/.rels",wrapRels([rel("rId1","officeDocument","word/document.xml")])],
      ["word/document.xml",xmlHead+'<w:document xmlns:w="'+word+'"><w:body>'+paragraphs+
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="450" w:footer="450" w:gutter="0"/></w:sectPr></w:body></w:document>']
    ]),DOCX);
  }
  globalThis.BacklogOffice={importExcel,exportExcel,exportWord,encode64,decode64};
})();
