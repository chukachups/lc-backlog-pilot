(() => {
  "use strict";
  const model=globalThis.Backlog, office=globalThis.BacklogOffice, github=globalThis.BacklogGitHub;
  const credentials=globalThis.BacklogCredentials;
  const params=new URLSearchParams(location.search);
  const workspace=params.get("workspace")==="work"?"excel":"demo";
  const workRepo="chukachups/lc-backlog-work-data";
  const cacheKey=workspace==="excel"?"lc-backlog-work-v1":"lc-backlog-pilot-v1";
  const byId=id=>document.getElementById(id);
  const view=byId("view"), dialog=byId("dialog"), dialogBody=byId("dialog-body");
  const views=["deadlines","registry","gantt","report"];
  const plural=new Intl.PluralRules("ru");
  const count=(n,forms)=>n+" "+forms[plural.select(n)];
  const taskCount=n=>count(n,{one:"задача",few:"задачи",many:"задач",other:"задачи"});
  const emptySlot=data=>({data,repo:"",sha:"",dirty:false});
  let slots={mode:workspace,demo:emptySlot(model.demo()),excel:workspace==="excel"?emptySlot(null):null};
  let cacheError=false, cloud=null, syncTimer, revision=0, toastTimer;
  let resuming=false, resumeAvailable=false, resumeDeferred=false, connectionMessage="";
  let deadlineFilter=params.get("show")==="done"?"done":"overdue";
  let registryQuery="", registrySheet="", reportUI=null;
  const slot=()=>slots[slots.mode];
  const state=()=>slot().data || {name:"Рабочий беклог",tasks:[],quarters:[],legend:{},report:null};
  const cloudMode=()=>slots.mode===workspace;
  const storageRepo=()=>workspace==="excel"?workRepo:params.get("repo") || "";
  function checkRemote(data) {
    if (data && data.kind!==workspace) throw Error("Файл из GitHub относится к другому беклогу.");
  }
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
        if (!saved[mode]?.data) continue;
        const row=saved[mode], data=model.validate(row.data);
        if (data.kind!==mode) continue;
        slots[mode]={data,repo:String(row.repo||""),sha:String(row.sha||""),dirty:row.dirty===true};
      }
      if (workspace==="demo" && saved.mode==="excel" && slots.excel) slots.mode="excel";
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
    if (cloudMode()) { revision++; queueSync(); }
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
        if (workspace==="excel" && data.kind!=="excel") throw Error("Для учебной копии откройте учебный беклог.");
        if (!confirm("Заменить локальную копию выбранным файлом?")) return;
        if (cloud?.busy) throw Error("Дождитесь завершения синхронизации и попробуйте снова.");
        if (data.kind===workspace) {cloud?.client.close();cloud=null;clearTimeout(syncTimer);revision++;}
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
    if (cloudMode()) {
      if (cloud) queueSync();
      else resumeConnection();
    }
  }
  function paintStatus() {
    const control=byId("connect");
    const online=Boolean(cloud?.ready)&&cloudMode();
    control.classList.toggle("connected",online && !cloud.error);
    control.disabled=cloudMode()&&resuming;
    byId("connect-caption").textContent=!cloudMode()?"Локальная книга Excel":resuming?"Подключение…":cloud?.busy?"Синхронизация…":
      cloud?.error?"Проверить синхронизацию":online?(slots[workspace].dirty?"Есть изменения":"GitHub подключён"):
      cloud?"Настроить хранилище":resumeAvailable?"Переподключить GitHub":"Подключить GitHub";
    const source=byId("source");
    source.replaceChildren(el("span",null,state().name+(cacheError?" · браузер не сохранил копию":"")));
    if (workspace==="demo" && slots.excel) source.append(document.createTextNode(" · "),
      button(slots.mode==="excel"?"Учебные данные":"Локальная книга","link-button",
        ()=>switchMode(slots.mode==="excel"?"demo":"excel")));
  }
  const stagedReport=()=>reportUI?.dirty;
  function attachCloud(client,remote) {
    checkRemote(remote.data);
    if (workspace==="excel" && !remote.data) throw Error("В рабочем репозитории отсутствует файл беклога.");
    const local=slots[workspace];
    const conflict=Boolean(remote.data && local.dirty && (local.repo!==client.key || local.sha!==remote.sha));
    cloud={client,ready:Boolean(remote.data),busy:false,conflict,
      error:conflict?"В браузере и на GitHub есть разные правки. Скачайте копию перед загрузкой данных GitHub.":""};
    if (remote.data && !conflict) {
      if (!local.dirty) { local.data=remote.data;if (!stagedReport()) reportUI=null; }
      local.repo=client.key;local.sha=remote.sha;
    }
  }
  async function clearRemembered() {
    await credentials.clear(workspace);
    resumeAvailable=false;
  }
  async function resumeConnection() {
    if (cloud || resuming || !cloudMode() || !credentials.available()) return;
    resuming=true;connectionMessage="";paintStatus();
    let client;
    try {
      const saved=await credentials.load(workspace);
      resumeAvailable=Boolean(saved);
      if (!saved || !cloudMode()) return;
      const requested=storageRepo();
      if (requested && requested.toLowerCase()!==saved.repo.toLowerCase()) return;
      client=github.connect(saved.repo,saved.token);
      await client.verify();
      const remote=await client.read();
      if (!cloudMode()) return;
      if (dialog.open) { resumeDeferred=true;return; }
      attachCloud(client,remote);client=null;
      cache();render();queueSync();
    } catch (error) {
      connectionMessage=error.message;
      if (error instanceof github.AuthenticationError) {
        try { await clearRemembered(); } catch { connectionMessage+=" Браузер не смог удалить старый ключ."; }
      }
    } finally {
      client?.close();
      resuming=false;paintStatus();
    }
  }
  function queueSync() {
    clearTimeout(syncTimer);
    if (cloudMode() && cloud?.ready && !cloud.conflict && slots[workspace].dirty)
      syncTimer=setTimeout(()=>sync(false),800);
  }
  async function sync(pull) {
    const connection=cloud;
    if (!cloudMode() || !connection?.ready || connection.busy || connection.conflict) return;
    const local=slots[workspace], before=revision;
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
        checkRemote(remote.data);
        if (remote.sha!==local.sha && before===revision && !local.dirty && !dialog.open && !stagedReport()) {
          local.sha=remote.sha; local.data=remote.data; reportUI=null;
          if (cloudMode()) render();
        }
      }
      success=true;
    } catch (error) {
      connection.error=error.message;
      if (error instanceof github.Conflict) connection.conflict=true;
      if (error instanceof github.AuthenticationError) {
        connection.client.close();cloud=null;connectionMessage=error.message;
        try { await clearRemembered(); } catch { connectionMessage+=" Браузер не смог удалить старый ключ."; }
        notify(connectionMessage);
      }
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
  dialog.addEventListener("close",()=>{
    if (resumeDeferred) { resumeDeferred=false;resumeConnection(); }
  });
  function cloudPanel() {
    if (resuming && cloudMode()) return;
    const panel=el("div","form-grid");
    if (!cloudMode()) {
      panel.append(el("p","context-box","Книга Excel открыта локально. Синхронизация в пилоте доступна в учебном беклоге."),
        button("Открыть учебный беклог","primary",()=>{dialog.close();switchMode("demo");}));
      restoreControl(panel);
      openDialog("Хранилище GitHub",panel); return;
    }
    if (!cloud) {
      const repo=input("text","Закрытый репозиторий",workspace==="excel"?workRepo:slots.demo.repo || storageRepo().slice(0,140));
      repo.placeholder="владелец/репозиторий";
      repo.readOnly=workspace==="excel";
      const token=input("password","Fine-grained токен"); token.autocomplete="off";
      const remember=input("checkbox","Запомнить на этом устройстве");
      remember.checked=credentials.available();
      remember.disabled=!credentials.available();
      const keep=el("label","checkrow remember");
      keep.append(remember,el("span","small","Запомнить на этом устройстве"));
      const error=el("p","form-error",connectionMessage);
      error.setAttribute("role","status");
      const connect=button("Подключиться","primary",async()=>{
        let client, connected=false;
        let accessKey=token.value;
        connect.disabled=true; error.textContent="";
        try {
          client=github.connect(workspace==="excel"?workRepo:repo.value,accessKey);token.value="";
          await client.verify();
          const remote=await client.read();
          if (!dialog.open || !dialogBody.contains(panel)) return;
          attachCloud(client,remote);
          connectionMessage="";
          let notice="Беклог подключён к GitHub";
          try {
            if (remember.checked) {
              await credentials.save(client.name,accessKey.trim(),workspace);
              resumeAvailable=true;
            } else if (resumeAvailable) await clearRemembered();
          } catch { notice="GitHub подключён. Браузер не смог запомнить ключ."; }
          connected=true; cache(); render();
          if (cloud.ready && !cloud.conflict) { dialog.close(); notify(notice); queueSync(); }
          else cloudPanel();
        } catch (exception) { error.textContent=exception.message; }
        finally { accessKey="";connect.disabled=false; if (!connected) client?.close(); }
      });
      token.onkeydown=event=>{ if (event.key==="Enter" && !connect.disabled) connect.click(); };
      panel.append(field("Репозиторий",repo),field("Ключ доступа",token),keep,error,el("div","actions"));
      panel.lastChild.appendChild(connect);
      const keyLink=el("a","small","Создать ключ на GitHub");
      keyLink.href="https://github.com/settings/personal-access-tokens/new";
      keyLink.target="_blank";keyLink.rel="noopener noreferrer";
      panel.lastChild.appendChild(keyLink);
      if (resumeAvailable) {
        panel.appendChild(button("Повторить сохранённое подключение","secondary",()=>{dialog.close();resumeConnection();}));
      }
    } else {
      const connection=cloud, info=el("p",null,connection.client.name), error=el("p","form-error",connection.error);
      const actions=el("div","actions");
      panel.append(info);
      if (resumeAvailable) panel.appendChild(el("p","muted small","Подключение сохранено в этом браузере."));
      panel.append(error,actions);
      if (!connection.ready) {
        panel.insertBefore(el("p","muted","В репозитории пока нет backlog.json. Создайте файл из текущего учебного беклога."),error);
        const create=button("Сохранить учебный беклог на GitHub","primary",async()=>{
          create.disabled=true; connection.busy=true; paintStatus();
          try {
            const local=slots[workspace], before=revision;
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
        actions.append(button("Скачать локальную копию","secondary",()=>downloadBackup(slots[workspace].data)),
          button("Загрузить версию GitHub","primary",async()=>{
            if (!confirm("Заменить копию в браузере версией GitHub? Сохраните свою копию перед заменой.")) return;
            try {
              const remote=await connection.client.read();
              if (!remote.data) throw Error("На GitHub не найден файл беклога.");
              checkRemote(remote.data);
              slots[workspace]={data:remote.data,repo:connection.client.key,sha:remote.sha,dirty:false};
              connection.conflict=false; connection.error=""; revision++; reportUI=null;
              dialog.close(); cache(); render(); notify("Загружена свежая версия GitHub");
            } catch (exc) { error.textContent=exc.message; }
          }));
      } else {
        actions.append(button("Синхронизировать","primary",()=>{ dialog.close(); sync(true); }),
          button("Скачать копию","secondary",()=>downloadBackup(slots[workspace].data)));
      }
      actions.appendChild(button("Отключить","secondary",async()=>{
        if (connection.busy) return;
        try { await clearRemembered(); }
        catch { error.textContent="Браузер не смог удалить сохранённый ключ. Попробуйте ещё раз.";return; }
        connection.client.close();cloud=null;connectionMessage="";clearTimeout(syncTimer);
        cache();cloudPanel();
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
  const deadlineNames={overdue:["Просрочено","danger"],today:["Сегодня","amber"],week:["Ближайшие 7 дней","good"],
    none:["Без срока",""],done:["Выполнено","good"]};
  function deadlines() {
    heading("Сроки",model.dateLabel(model.today()));
    const groups={overdue:[],today:[],week:[],none:[],done:[]};
    state().tasks.forEach(task=>task.steps.forEach(step=>{
      const key=step.done?"done":model.deadlineGroup(step);
      if (groups[key]) groups[key].push({task,step});
    }));
    const completed=deadlineFilter==="done";
    const panel=el("section","panel"), counts=el("div","filters deadline-filters"), list=el("div","step-list");
    counts.setAttribute("aria-label","Фильтр шагов");
    Object.entries(deadlineNames).forEach(([key,[label,tone]])=>{
      const tile=button("","filter "+tone,()=>{deadlineFilter=key;render();});
      tile.setAttribute("aria-pressed",String(deadlineFilter===key));
      tile.append(el("span",null,label),el("span","filter-count",groups[key].length));
      counts.appendChild(tile);
    });
    const rows=groups[deadlineFilter].slice().sort((a,b)=>completed?
      b.step.completedOn.localeCompare(a.step.completedOn):a.step.due.localeCompare(b.step.due));
    if (!rows.length) list.appendChild(el("p","empty",completed?"Выполненных шагов пока нет.":"В этой группе шагов нет."));
    if (rows.length) {
      const header=el("div","deadline deadline-head");
      header.append(el("span"),el("span",null,"Шаг"),el("span",null,"Задача"),el("span",null,completed?"Дата выполнения":"Срок"));
      list.appendChild(header);
    }
    for (const {task,step} of rows) {
      const row=el("div","deadline");
      if (completed) {
        const mark=el("span","done-mark","✓");
        mark.setAttribute("aria-label","Выполнено");
        row.appendChild(mark);
      } else {
        const toggle=input("checkbox","Выполнить "+step.text);
        toggle.onchange=()=>{step.done=true;step.completedOn=model.today();changed();render();notify("Шаг отмечен выполненным");};
        row.appendChild(toggle);
      }
      const taskLink=button(task.title,"link-button deadline-task",()=>taskDialog(task.id,completed?"done":"all"));
      taskLink.title=task.code+" · "+task.title;
      const date=completed?step.completedOn:step.due;
      const dateCell=el("span","date-badge deadline-date "+deadlineNames[deadlineFilter][1],date?model.dateLabel(date):"—");
      if (completed && !date) dateCell.title="Дата выполнения не указана";
      row.append(el("p","deadline-step",step.text),taskLink,dateCell);
      list.appendChild(row);
    }
    panel.append(counts,list);view.appendChild(panel);
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
      const [grid,body]=table(["Задача","Исполнитель","Статус"],"registry");
      for (const task of tasks) {
        const tr=body.insertRow();
        const done=task.steps.filter(step=>step.done).length;
        const doneLink=button("Выполнено "+done+" из "+task.steps.length,"link-button registry-completed",()=>taskDialog(task.id,"done"));
        doneLink.setAttribute("aria-label","Выполнено "+done+" из "+task.steps.length+" · "+task.title);
        tr.insertCell().append(el("div","muted small",task.code+(task.jira && task.jira!=="—"?" · "+task.jira:"")),
          button(task.title,"link-button",()=>taskDialog(task.id)),doneLink);
        tr.insertCell().textContent=task.owner||"—";
        tr.insertCell().appendChild(el("span","tag",task.status||"—"));
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
  function taskDialog(id,initialFilter="all") {
    const current=state().tasks.find(task=>task.id===id);
    if (!current) return;
    const draft=model.copy(current), panel=el("div","form-grid"), filters=el("div","filters task-filters"),
      steps=el("div","task-steps"), error=el("p","form-error");
    let filter=initialFilter;
    const result=getTextArea("Результат для руководства",draft.result);
    result.placeholder="Что проверили и какой результат получили";
    result.maxLength=12000;
    const renderSteps=(focusId)=>{
      steps.replaceChildren();filters.replaceChildren();
      const open=draft.steps.filter(step=>!step.done).length, done=draft.steps.length-open;
      for (const [key,title,size] of [["all","Все",draft.steps.length],["open","Нужно выполнить",open],["done","Выполнено",done]]) {
        const tab=button("","filter",()=>{filter=key;renderSteps();});
        tab.setAttribute("aria-pressed",String(filter===key));tab.dataset.stepFilter=key;
        tab.append(el("span",null,title),el("span","filter-count",size));
        filters.appendChild(tab);
      }
      const renderStep=(step,i,group)=>{
        const card=el("div","editor-card"+(step.done?" is-done":"")), label="Шаг "+(i+1), text=getTextArea(label,step.text,2);
        text.maxLength=4000; text.oninput=()=>{step.text=text.value;};
        const date=input("date","Срок "+label,step.due); date.onchange=()=>{step.due=date.value;step.review=false;review.hidden=true;};
        const checkbox=input("checkbox","Выполнено "+label); checkbox.checked=step.done;
        checkbox.dataset.stepToggle=step.id;
        checkbox.onchange=()=>{
          step.done=checkbox.checked;step.completedOn=step.done?model.today():"";
          renderSteps(step.id);
        };
        const toggle=el("label","inline-field"); toggle.append(checkbox,document.createTextNode("Выполнено"));
        const due=el("label","inline-field"); due.append(document.createTextNode("Срок"),date);
        const controls=el("div","step-meta");
        controls.append(toggle,due,button("Удалить шаг","subtle",()=>{draft.steps.splice(i,1);renderSteps();}));
        if (step.completedOn) controls.appendChild(el("span","muted small","Готово "+model.dateLabel(step.completedOn)));
        const review=el("p","small form-error","Год определён по дате файла. Проверьте срок."); review.hidden=!step.review;
        card.append(text,controls,review);group.appendChild(card);
      };
      for (const [key,title,done] of [["open","Нужно выполнить",false],["done","Выполнено",true]]) {
        if (filter!=="all" && filter!==key) continue;
        const group=el("section","step-section");
        if (filter==="all") group.appendChild(el("h3",null,title));
        draft.steps.forEach((step,i)=>{ if (step.done===done) renderStep(step,i,group); });
        if (!draft.steps.some(step=>step.done===done))
          group.appendChild(el("p","muted small",done?"Выполненных шагов пока нет.":"Открытых шагов пока нет."));
        steps.appendChild(group);
      }
      if (focusId) {
        const toggle=[...steps.querySelectorAll("input[data-step-toggle]")].find(node=>node.dataset.stepToggle===focusId);
        (toggle ?? filters.querySelector('[aria-pressed="true"]'))?.focus();
      }
    };
    const add=button("Добавить шаг","secondary",()=>{
      draft.steps.push({id:model.uid(),text:"",due:"",done:false,completedOn:"",review:false});
      filter="open";renderSteps();steps.querySelector(".editor-card:last-child textarea")?.focus();
    });
    const bulk=el("details"), paste=getTextArea("Список шагов","",4);
    paste.placeholder="Вставьте список из итогов встречи. Укажите срок рядом с шагом, если он известен.";
    const append=button("Добавить из списка","secondary",()=>{
      const added=model.parseSteps(paste.value);
      if (!added.length) return;
      draft.steps.push(...added);paste.value="";filter="all";renderSteps();notify("Добавлено шагов: "+added.length);
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
    const stepActions=el("div","actions step-actions");
    stepActions.append(add,bulk);
    panel.append(filters,steps,stepActions,field("Результат для руководства",result),error,controls);
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
  byId("import").hidden=workspace==="excel";
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
    byId("export").disabled=!slot().data;
    if (!slot().data) {
      heading("Рабочий беклог");
      const panel=el("div","panel empty");
      panel.append(el("p",null,"Подключите рабочий репозиторий GitHub, чтобы открыть задачи."));
      view.appendChild(panel);paintStatus();return;
    }
    if (page==="registry") registry();
    else if (page==="gantt") gantt();
    else if (page==="report") reportView();
    else deadlines();
    paintStatus();
  }
  window.addEventListener("hashchange",render);
  restoreCache();render();resumeConnection();
})();
