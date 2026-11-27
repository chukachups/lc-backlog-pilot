(() => {
  "use strict";
  const model=globalThis.Backlog, office=globalThis.BacklogOffice, github=globalThis.BacklogGitHub;
  const cacheKey="lc-backlog-pilot-v1";
  const byId=id=>document.getElementById(id);
  const view=byId("view"), dialog=byId("dialog"), dialogBody=byId("dialog-body");
  const views=["deadlines","registry","gantt","report"];
  const plural=new Intl.PluralRules("ru");
  const count=(n,forms)=>n+" "+forms[plural.select(n)];
  const stepCount=n=>count(n,{one:"шаг",few:"шага",many:"шагов",other:"шага"});
  const taskCount=n=>count(n,{one:"задача",few:"задачи",many:"задач",other:"задачи"});
  const emptySlot=data=>({data,repo:"",sha:"",dirty:false});
  let slots={mode:"demo",demo:emptySlot(model.demo()),excel:null};
  let cacheError=false, cloud=null, syncTimer, revision=0, toastTimer;
  let deadlineFilter="overdue", registryQuery="", registrySheet="", reportUI=null;
  const slot=()=>slots[slots.mode];
  const state=()=>slot().data;
  const section=()=>views.includes(location.hash.slice(1))?location.hash.slice(1):"deadlines";
  const el=(tag,cls,text)=>{
    const node=document.createElement(tag);
    if (cls) node.className=cls;
    if (text!==undefined) node.textContent=text;
    return node;
  };
  const button=(label,cls,action)=>{
    const node=el("button",cls,label); node.type="button"; node.onclick=action; return node;
  };
  const input=(type,label,value="")=>{
    const node=el("input"); node.type=type; node.value=value; node.setAttribute("aria-label",label); return node;
  };
  function field(label,control) {
    const node=el("label"); node.append(el("span","field-title",label),control); return node;
  }
  function getTextArea(label,value="",rows=3) {
    const area=el("textarea"); area.rows=rows; area.value=value; area.setAttribute("aria-label",label); return area;
  }
  function restoreCache() {
    try {
      const saved=JSON.parse(localStorage.getItem(cacheKey));
      if (!saved) return;
      for (const mode of ["demo","excel"]) {
        if (!saved[mode]) continue;
        const row=saved[mode], data=model.validate(row.data);
        if (data.kind!==mode) continue;
        slots[mode]={data,repo:String(row.repo||""),sha:String(row.sha||""),dirty:row.dirty===true};
      }
      if (saved.mode==="excel" && slots.excel) slots.mode="excel";
    } catch { cacheError=true; }
  }
  function cache() {
    try { localStorage.setItem(cacheKey,JSON.stringify(slots)); cacheError=false; }
    catch { cacheError=true; }
    paintStatus();
  }
  function notify(text) {
    const toast=byId("toast");
    toast.textContent=text; toast.hidden=false; clearTimeout(toastTimer);
    toastTimer=setTimeout(()=>{ toast.hidden=true; },4500);
  }
  function changed() {
    slot().dirty=true;
    if (slots.mode==="demo") { revision++; queueSync(); }
    cache();
  }
  function download(blob,name) {
    const url=URL.createObjectURL(blob), link=el("a");
    link.href=url; link.download=name;
    document.body.appendChild(link); link.click(); link.remove();
    setTimeout(()=>URL.revokeObjectURL(url),1000);
  }
  function downloadBackup(data=state()) {
    download(new Blob([JSON.stringify(data,null,2)],{type:"application/json"}),"Беклог — копия "+model.today()+".json");
  }
  function restoreControl(panel) {
    const chooser=input("file","Копия беклога"); chooser.accept=".json"; chooser.hidden=true;
    panel.append(chooser,button("Открыть копию JSON","secondary",()=>chooser.click()));
    chooser.onchange=async()=>{
      const file=chooser.files[0]; chooser.value="";
      if (!file) return;
      try {
        if (file.size>12000000) throw Error("Копия беклога слишком большая.");
        const data=model.validate(JSON.parse(await file.text()));
        if (!confirm("Заменить локальную копию выбранным файлом?")) return;
        if (cloud?.busy) throw Error("Дождитесь завершения синхронизации и попробуйте снова.");
        if (data.kind==="demo") {cloud?.client.close();cloud=null;clearTimeout(syncTimer);revision++;}
        slots[data.kind]=emptySlot(data); slots[data.kind].dirty=true;
        slots.mode=data.kind; reportUI=null;
        dialog.close();cache();render();notify("Копия открыта в браузере");
      } catch (error) { notify(error instanceof SyntaxError?"Не удалось прочитать копию JSON.":error.message); }
    };
  }
  function switchMode(mode) {
    if (mode===slots.mode) return;
    slots.mode=mode; reportUI=null; registryQuery=""; registrySheet="";
    cache(); render();
    if (mode==="demo") queueSync();
  }
  function paintStatus() {
    const control=byId("connect");
    const online=Boolean(cloud?.ready)&&slots.mode==="demo";
    control.classList.toggle("connected",online && !cloud.error);
    byId("connect-caption").textContent=slots.mode==="excel"?"Локальная книга Excel":cloud?.busy?"Синхронизация…":
      cloud?.error?"Проверить синхронизацию":online?(slots.demo.dirty?"Есть изменения":"GitHub подключён"):
      cloud?"Настроить хранилище":"Подключить GitHub";
    const source=byId("source");
    source.replaceChildren(el("span",null,state().name+(cacheError?" · браузер не сохранил копию":"")));
    if (slots.excel) source.append(document.createTextNode(" · "),
      button(slots.mode==="excel"?"Учебные данные":"Локальная книга","link-button",
        ()=>switchMode(slots.mode==="excel"?"demo":"excel")));
  }
  const stagedReport=()=>reportUI?.dirty;
  function queueSync() {
    clearTimeout(syncTimer);
    if (slots.mode==="demo" && cloud?.ready && !cloud.conflict && slots.demo.dirty)
      syncTimer=setTimeout(()=>sync(false),800);
  }
  async function sync(pull) {
    const connection=cloud;
    if (slots.mode!=="demo" || !connection?.ready || connection.busy || connection.conflict) return;
    const local=slots.demo, before=revision;
    if (!local.dirty && (!pull || dialog.open || stagedReport())) return;
    connection.busy=true; connection.error=""; paintStatus();
    let success=false;
    try {
      if (local.dirty) {
        local.sha=await connection.client.write(model.copy(local.data),local.sha);
        if (revision===before) local.dirty=false;
      } else {
        const remote=await connection.client.read();
        if (!remote.data) throw Error("Файл беклога исчез из репозитория. Проверьте подключение.");
        if (remote.data.kind!=="demo") throw Error("Для облачного пилота используйте вымышленные данные.");
        if (remote.sha!==local.sha && before===revision && !local.dirty && !dialog.open && !stagedReport()) {
          local.sha=remote.sha; local.data=remote.data; reportUI=null;
          if (slots.mode==="demo") render();
        }
      }
      success=true;
    } catch (error) {
      connection.error=error.message;
      if (error instanceof github.Conflict) connection.conflict=true;
      if (pull || connection.conflict) notify(error.message);
    } finally {
      connection.busy=false; cache();
      if (success && local.dirty) queueSync();
    }
  }
  function openDialog(title,content) {
    byId("dialog-title").textContent=title; dialogBody.replaceChildren(content);
    if (!dialog.open) dialog.showModal();
  }
  byId("close-dialog").onclick=()=>dialog.close();
  function cloudPanel() {
    const panel=el("div","form-grid");
    if (slots.mode==="excel") {
      panel.append(el("p","context-box","Книга Excel открыта локально. Синхронизация в пилоте доступна в учебном беклоге."),
        button("Открыть учебный беклог","primary",()=>{dialog.close();switchMode("demo");}));
      restoreControl(panel);
      openDialog("Хранилище GitHub",panel); return;
    }
    if (!cloud) {
      const requestedRepo=new URLSearchParams(location.search).get("repo") || "";
      const repo=input("text","Закрытый репозиторий",slots.demo.repo || requestedRepo.slice(0,140));
      repo.placeholder="владелец/репозиторий";
      const token=input("password","Fine-grained токен"); token.autocomplete="off";
      const error=el("p","form-error");
      error.setAttribute("role","status");
      const connect=button("Подключиться","primary",async()=>{
        let client, connected=false;
        connect.disabled=true; error.textContent="";
        try {
          client=github.connect(repo.value,token.value); token.value="";
          await client.verify();
          const remote=await client.read(), local=slots.demo;
          if (!dialog.open || !dialogBody.contains(panel)) return;
          if (remote.data && remote.data.kind!=="demo") throw Error("Для облачного пилота используйте вымышленные данные.");
          const conflict=Boolean(remote.data && local.dirty && (local.repo!==client.key || local.sha!==remote.sha));
          cloud={client,ready:Boolean(remote.data),busy:false,conflict,error:conflict?"В локальной копии и на GitHub есть разные правки. Скачайте копию перед загрузкой данных GitHub.":""};
          if (remote.data && !conflict) {
            if (!local.dirty) { local.data=remote.data; reportUI=null; }
            local.repo=client.key; local.sha=remote.sha;
          }
          connected=true; cache(); render();
          if (remote.data && !conflict) { dialog.close(); notify("Беклог подключён к GitHub"); queueSync(); }
          else cloudPanel();
        } catch (exception) { error.textContent=exception.message; }
        finally { connect.disabled=false; if (!connected) client?.close(); }
      });
      panel.append(el("p","muted","Подключите отдельный репозиторий Private для учебных задач."),
        field("Владелец / репозиторий",repo),field("Fine-grained токен",token),
        el("p","muted small","Разрешение Contents: Read and write только для этого репозитория. Токен действует в открытой вкладке."),
        error,el("div","actions"));
      panel.lastChild.appendChild(connect);
      const keyLink=el("a","small","Создать ключ на GitHub");
      keyLink.href="https://github.com/settings/personal-access-tokens/new";
      keyLink.target="_blank";keyLink.rel="noopener noreferrer";
      panel.lastChild.appendChild(keyLink);
    } else {
      const connection=cloud, info=el("p","context-box",connection.client.name), error=el("p","form-error",connection.error);
      const actions=el("div","actions");
      panel.append(info,error,actions);
      if (!connection.ready) {
        panel.insertBefore(el("p","muted","В репозитории пока нет backlog.json. Создайте файл из текущего учебного беклога."),error);
        const create=button("Сохранить учебный беклог на GitHub","primary",async()=>{
          create.disabled=true; connection.busy=true; paintStatus();
          try {
            const local=slots.demo, before=revision;
            local.sha=await connection.client.write(model.copy(local.data),"");
            local.repo=connection.client.key;
            if (before===revision) local.dirty=false;
            connection.ready=true; connection.error="";
            dialog.close(); notify("Учебный беклог сохранён на GitHub");
          } catch (exc) {
            connection.error=exc.message; error.textContent=exc.message;
            if (exc instanceof github.Conflict) { connection.ready=true; connection.conflict=true; cloudPanel(); }
          } finally { connection.busy=false; create.disabled=false; cache(); queueSync(); }
        });
        actions.appendChild(create);
      } else if (connection.conflict) {
        actions.append(button("Скачать локальную копию","secondary",()=>downloadBackup(slots.demo.data)),
          button("Загрузить версию GitHub","primary",async()=>{
            if (!confirm("Заменить учебную копию в браузере версией GitHub? Сохраните свою копию перед заменой.")) return;
            try {
              const remote=await connection.client.read();
              if (!remote.data || remote.data.kind!=="demo") throw Error("На GitHub не найдена учебная версия.");
              slots.demo={data:remote.data,repo:connection.client.key,sha:remote.sha,dirty:false};
              connection.conflict=false; connection.error=""; revision++; reportUI=null;
              dialog.close(); cache(); render(); notify("Загружена свежая версия GitHub");
            } catch (exc) { error.textContent=exc.message; }
          }));
      } else {
        actions.append(button("Синхронизировать","primary",()=>{ dialog.close(); sync(true); }),
          button("Скачать копию","secondary",()=>downloadBackup(slots.demo.data)));
      }
      actions.appendChild(button("Отключить","secondary",()=>{
        if (connection.busy) return;
        connection.client.close(); cloud=null; clearTimeout(syncTimer);
        cache(); cloudPanel();
      }));
    }
    restoreControl(panel);
    openDialog("Хранилище GitHub",panel);
  }
  byId("connect").onclick=cloudPanel;
  window.addEventListener("focus",()=>sync(true));

  function heading(title,detail) {
    const top=el("div","view-head"), text=el("div");
    text.appendChild(el("h1",null,title));
    if (detail) text.appendChild(el("p","muted small",detail));
    top.appendChild(text); view.appendChild(top);
  }
  const deadlineNames={overdue:["Просрочено","danger"],today:["Сегодня","amber"],week:["Ближайшие 7 дней","good"],none:["Без срока",""]};
  function deadlines() {
    heading("Сроки",model.dateLabel(model.today()));
    const groups={overdue:[],today:[],week:[],none:[]};
    state().tasks.forEach(task=>task.steps.forEach(step=>{
      const key=model.deadlineGroup(step);
      if (groups[key]) groups[key].push({task,step});
    }));
    const counts=el("div","counts");
    Object.entries(deadlineNames).forEach(([key,[label,tone]])=>{
      const tile=button("","metric "+tone,()=>{deadlineFilter=key;render();});
      tile.setAttribute("aria-pressed",String(deadlineFilter===key));
      tile.append(el("strong",null,groups[key].length),el("span",null,label));
      counts.appendChild(tile);
    });
    view.appendChild(counts);
    const panel=el("section","panel"), bar=el("div","panel-head"), list=el("div","step-list");
    bar.append(el("h2",null,deadlineNames[deadlineFilter][0]),el("span","muted small",stepCount(groups[deadlineFilter].length)));
    const rows=groups[deadlineFilter].slice().sort((a,b)=>a.step.due.localeCompare(b.step.due));
    if (!rows.length) list.appendChild(el("p","empty","В этой группе шагов нет."));
    for (const {task,step} of rows) {
      const row=el("div","deadline"), toggle=input("checkbox","Выполнить "+step.text), info=el("div","deadline-info");
      toggle.onchange=()=>{step.done=true;step.completedOn=model.today();changed();render();notify("Шаг отмечен выполненным");};
      info.append(el("p",null,step.text),button(task.code+" · "+task.title,"link-button small",()=>taskDialog(task.id)));
      row.append(toggle,info,el("span","date-badge "+deadlineNames[deadlineFilter][1],step.due?model.dateLabel(step.due):"Без срока"));
      list.appendChild(row);
    }
    panel.append(bar,list);view.appendChild(panel);
  }
  function table(labels,cls) {
    const grid=el("table","table "+cls), head=grid.createTHead().insertRow();
    labels.forEach(label=>{const col=el("th",null,label);col.scope="col";head.appendChild(col);});
    return [grid,grid.createTBody()];
  }
  function sheetSelect(value,onChange) {
    const select=el("select"); select.setAttribute("aria-label","Направление");
    select.append(new Option("Все направления",""),...model.sheets.map(sheet=>new Option(sheet,sheet)));
    select.value=value;select.onchange=()=>onChange(select.value);return select;
  }
  function registry() {
    heading("Реестр задач",taskCount(state().tasks.length));
    const panel=el("section","panel"), tools=el("div","toolbar"), results=el("div","table-scroll");
    const search=input("search","Поиск по задаче, исполнителю или Jira",registryQuery);
    search.placeholder=search.getAttribute("aria-label");
    search.oninput=()=>{registryQuery=search.value;draw();};
    tools.append(search,sheetSelect(registrySheet,value=>{registrySheet=value;draw();}));
    function draw() {
      const tasks=state().tasks.filter(task=>(!registrySheet||task.sheet===registrySheet) &&
        model.normalize([task.title,task.code,task.jira,task.owner].join(" ")).includes(model.normalize(registryQuery)));
      results.replaceChildren();
      if (!tasks.length) { results.appendChild(el("p","empty","Задачи по фильтру не найдены.")); return; }
      const [grid,body]=table(["Задача","Исполнитель","Статус","Шаги"],"registry");
      for (const task of tasks) {
        const tr=body.insertRow();
        tr.insertCell().append(el("div","muted small",task.code+(task.jira && task.jira!=="—"?" · "+task.jira:"")),
          button(task.title,"link-button",()=>taskDialog(task.id)));
        tr.insertCell().textContent=task.owner||"—";
        tr.insertCell().appendChild(el("span","tag",task.status||"—"));
        tr.insertCell().textContent=task.steps.filter(step=>step.done).length+" / "+task.steps.length;
      }
      results.appendChild(grid);
    }
    panel.append(tools,results); view.appendChild(panel); draw();
  }
  function gantt() {
    heading("Гант");
    const data=state(), panel=el("section","panel"), legend=el("div","gantt-legend"), wrap=el("div","table-scroll");
    Object.entries(data.legend).forEach(([mark,title])=>{
      const item=el("span","small");
      item.append(el("span","tag stage-"+mark,mark),document.createTextNode(" "+title));legend.appendChild(item);
    });
    const [grid,body]=table(["Задача",...data.quarters],"gantt");
    data.tasks.filter(task=>task.marks.length).forEach(task=>{
      const row=body.insertRow();
      row.insertCell().append(el("div","muted small",task.code),button(task.title,"link-button",()=>taskDialog(task.id)));
      data.quarters.forEach((_,i)=>{
        const cell=row.insertCell(), mark=task.marks[i];
        if (mark) {const tag=el("span","tag stage-"+mark,mark);tag.title=data.legend[mark]||mark;cell.appendChild(tag);}
      });
    });
    wrap.appendChild(grid);panel.append(legend,wrap);view.appendChild(panel);
  }
  function taskDialog(id) {
    const current=state().tasks.find(task=>task.id===id);
    if (!current) return;
    const draft=model.copy(current), panel=el("div","form-grid"), steps=el("div"), error=el("p","form-error");
    const result=getTextArea("Результат для руководства",draft.result);
    result.placeholder="Что проверили и какой результат получили";
    result.maxLength=12000;
    const heading=el("h3",null,"Шаги задачи");
    const renderSteps=()=>{
      steps.replaceChildren();
      draft.steps.forEach((step,i)=>{
        const card=el("div","editor-card"), label="Шаг "+(i+1), text=getTextArea(label,step.text,2);
        text.maxLength=4000; text.oninput=()=>{step.text=text.value;};
        const date=input("date","Срок "+label,step.due); date.onchange=()=>{step.due=date.value;step.review=false;review.hidden=true;};
        const checkbox=input("checkbox","Выполнено "+label); checkbox.checked=step.done;
        checkbox.onchange=()=>{step.done=checkbox.checked;step.completedOn=step.done?model.today():"";};
        const toggle=el("label","inline-field"); toggle.append(checkbox,document.createTextNode("Выполнено"));
        const due=el("label","inline-field"); due.append(document.createTextNode("Срок"),date);
        const controls=el("div","step-meta");
        controls.append(toggle,due,button("Удалить шаг","subtle",()=>{draft.steps.splice(i,1);renderSteps();}));
        const review=el("p","small form-error","Год определён по дате файла. Проверьте срок."); review.hidden=!step.review;
        card.append(text,controls,review);steps.appendChild(card);
      });
      if (!draft.steps.length) steps.appendChild(el("p","muted small","Шагов пока нет."));
    };
    const add=button("Добавить шаг","secondary",()=>{
      draft.steps.push({id:model.uid(),text:"",due:"",done:false,completedOn:"",review:false});
      renderSteps(); steps.lastChild.querySelector("textarea").focus();
    });
    const bulk=el("details"), paste=getTextArea("Список шагов","",4);
    paste.placeholder="Вставьте список из итогов встречи. Укажите срок рядом с шагом, если он известен.";
    const append=button("Добавить из списка","secondary",()=>{
      const added=model.parseSteps(paste.value);
      if (!added.length) return;
      draft.steps.push(...added);paste.value="";renderSteps();notify("Добавлено шагов: "+added.length);
    });
    bulk.append(el("summary",null,"Вставить список шагов"),paste,append);
    const controls=el("div","actions");
    controls.appendChild(button("Сохранить задачу","primary",()=>{
      const cleanSteps=draft.steps.filter(step=>step.text.trim()).map(step=>({...step,text:step.text.trim()}));
      if (cleanSteps.length>300) { error.textContent="У одной задачи может быть до 300 шагов.";return; }
      const newResult=result.value.trim();
      if (newResult!==current.result) { current.result=newResult;current.resultOn=newResult?model.today():""; }
      current.steps=cleanSteps;changed();dialog.close();render();notify("Задача сохранена");
    }));
    panel.append(el("p","muted small",[current.code,current.status,current.owner].filter(Boolean).join(" · ")));
    if (current.plan && current.plan!=="—") {
      const plan=el("details");
      const text=el("p","muted preserve",current.plan);
      plan.append(el("summary",null,"План из Excel"),text);panel.appendChild(plan);
    }
    panel.append(heading,steps,add,bulk,field("Результат для руководства",result),error,controls);
    openDialog(current.title,panel);renderSteps();
  }
  function reportView() {
    const data=state(), saved=data.report;
    if (!reportUI) reportUI={selected:new Set(saved?.selected??data.tasks.filter(t=>t.sheet!=="Уточняю").map(t=>t.id)),
      from:saved?.from||model.addDays(model.today(),-30),to:saved?.to||model.today(),text:saved?.text||"",dirty:false};
    heading("Отчёт для руководства");
    const wrap=el("div","report-layout"), choices=el("div","report-panel panel"), editor=el("div","report-panel panel");
    const from=input("date","Начало периода",reportUI.from), to=input("date","Конец периода",reportUI.to);
    const list=el("div","report-choices"), error=el("p","form-error"), checkboxes=[];
    choices.append(el("h2",null,"Период и задачи"),field("С",from),field("По",to));
    from.onchange=()=>{reportUI.from=from.value;};
    to.onchange=()=>{reportUI.to=to.value;};
    const selectAll=el("div","actions");
    const select=enabled=>checkboxes.forEach(([toggle,id])=>{
      toggle.checked=enabled;
      if (enabled) reportUI.selected.add(id); else reportUI.selected.delete(id);
    });
    selectAll.append(button("Выбрать все","secondary",()=>select(true)),button("Снять выбор","secondary",()=>select(false)));
    data.tasks.forEach(task=>{
      const row=el("label","checkrow"), toggle=input("checkbox",task.code+" "+task.title);
      toggle.checked=reportUI.selected.has(task.id);
      toggle.onchange=()=>{if (toggle.checked) reportUI.selected.add(task.id); else reportUI.selected.delete(task.id);};
      row.append(toggle,el("span","small",task.code+" · "+task.title));list.appendChild(row);checkboxes.push([toggle,task.id]);
    });
    const unknown=el("label","checkrow"), include=input("checkbox","Включить выполненные шаги без даты");
    unknown.append(include,el("span","small","Включить выполненные шаги без даты"));
    const area=getTextArea("Текст отчёта",reportUI.text,18);
    area.className="report-editor";area.maxLength=160000;area.placeholder="Выберите задачи и нажмите «Собрать текст».";
    area.oninput=()=>{reportUI.text=area.value;reportUI.dirty=true;};
    choices.append(selectAll,list,unknown,error,button("Собрать текст","primary",()=>{
      error.textContent="";
      if (!reportUI.selected.size) { error.textContent="Выберите хотя бы одну задачу.";return; }
      if (from.value && to.value && from.value>to.value) { error.textContent="Дата начала позже даты окончания.";return; }
      if (area.value.trim() && reportUI.dirty && !confirm("Заменить текущий текст новым черновиком?")) return;
      area.value=reportUI.text=model.report(data.tasks,reportUI.selected,from.value,to.value,include.checked);
      reportUI.dirty=true;
    }));
    const actions=el("div","actions");
    actions.append(button("Сохранить черновик","secondary",()=>{
      if (!area.value.trim()) { notify("Сначала соберите или напишите текст отчёта.");return; }
      data.report={text:area.value,from:from.value,to:to.value,selected:[...reportUI.selected]};
      reportUI.dirty=false;changed();notify("Черновик сохранён");
    }),button("Выгрузить Word","primary",()=>{
      if (!area.value.trim()) { notify("Сначала соберите или напишите текст отчёта.");return; }
      download(office.exportWord(area.value),"Отчёт по беклогу — "+(to.value||model.today())+".docx");
    }));
    editor.append(el("h2",null,"Текст отчёта"),area,actions);wrap.append(choices,editor);view.appendChild(wrap);
  }

  const workbook=byId("workbook-file");
  byId("import").onclick=()=>workbook.click();
  workbook.onchange=async()=>{
    const file=workbook.files[0]; workbook.value="";
    if (!file) return;
    const trigger=byId("import");trigger.disabled=true;
    try {
      const incoming=await office.importExcel(file), panel=el("div","form-grid");
      const inferred=incoming.tasks.flatMap(task=>task.steps).filter(step=>step.review).length;
      panel.append(el("p",null,incoming.name),el("p",null,"Загружено задач: "+incoming.tasks.length+"."));
      if (inferred) panel.appendChild(el("p","context-box warning","Проверьте год у "+inferred+" сроков, где в Excel был указан только день и месяц."));
      if (slots.excel) panel.append(el("p","muted","Эта книга заменит локальную копию Excel."),
        button("Скачать текущую копию","secondary",()=>downloadBackup(slots.excel.data)));
      panel.appendChild(button("Открыть книгу локально","primary",()=>{
        slots.excel=emptySlot(incoming);
        switchMode("excel");reportUI=null;cache();dialog.close();
        location.hash="registry";render();
      }));
      openDialog("Проверка импорта",panel);
    } catch (error) {
      openDialog("Ошибка импорта",el("p","form-error",error.message));
    } finally { trigger.disabled=false; }
  };
  byId("export").onclick=async()=>{
    const trigger=byId("export");trigger.disabled=true;
    try { download(await office.exportExcel(state()),"Беклог — "+model.today()+".xlsx"); }
    catch (error) { openDialog("Ошибка выгрузки",el("p","form-error",error.message)); }
    finally { trigger.disabled=false; }
  };
  function render() {
    const page=section();
    document.querySelectorAll("nav [data-view]").forEach(a=>{
      if (a.dataset.view===page) a.setAttribute("aria-current","page");else a.removeAttribute("aria-current");
    });
    view.replaceChildren();
    if (page==="registry") registry();
    else if (page==="gantt") gantt();
    else if (page==="report") reportView();
    else deadlines();
    paintStatus();
  }
  window.addEventListener("hashchange",render);
  restoreCache();render();
})();
