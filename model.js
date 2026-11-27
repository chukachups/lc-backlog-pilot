(() => {
  "use strict";

  const sheets = ["FTE", "ТМЗ", "Уточняю"];
  const sections = {"FTE":"Сокращение ручной работы","ТМЗ":"Запас и доступность","Уточняю":"Уточняю"};
  const legend = {"О":"Обследование","М":"MVP","Р":"Развитие","И":"Итоговый срок"};
  const datePattern = /^\d{4}-\d{2}-\d{2}$/;
  const uid = () => globalThis.crypto?.randomUUID?.() ?? Math.random().toString(36).slice(2) + Date.now().toString(36);
  const text = value => String(value ?? "").replace(/\u00a0/g, " ").trim();
  const normalize = value => text(value).toLocaleLowerCase("ru").replaceAll("ё", "е").replace(/\s+/g, " ");
  const copy = value => structuredClone(value);

  function today() {
    const parts = new Intl.DateTimeFormat("en-CA", {timeZone:"Europe/Moscow", year:"numeric",month:"2-digit",day:"2-digit"})
      .formatToParts(new Date());
    return ["year","month","day"].map(key => parts.find(part => part.type === key).value).join("-");
  }
  function addDays(date, n) {
    const time = new Date(date + "T00:00:00Z");
    time.setUTCDate(time.getUTCDate() + n);
    return time.toISOString().slice(0, 10);
  }
  function validDate(value) {
    if (typeof value !== "string" || !datePattern.test(value)) return false;
    const time = new Date(value + "T00:00:00Z");
    return !Number.isNaN(+time) && time.toISOString().slice(0, 10) === value;
  }
  const dateLabel = date => date ? date.split("-").reverse().join(".") : "";

  function detectDue(value, anchor) {
    const match = /(?:[,;]\s*|[ \t]+[-–—][ \t]+|[ \t]+(?=срок[ \t])|^)(?:(срок(?:[ \t]+до)?|до)[ \t]+)(\d{1,2})[./](\d{1,2})(?:[./](\d{4}|\d{2}))?\.?\s*$/i.exec(value);
    if (!match || !value.slice(0, match.index).trim()) return {text:value, due:"", review:false};
    const day = match[2].padStart(2,"0"), month = match[3].padStart(2,"0");
    const year = Number(anchor.slice(0,4));
    const explicit = match[4] ? Number(match[4]) + (match[4].length === 2 ? 2000 : 0) : null;
    const choices = (explicit ? [explicit] : [year,year + 1,year - 1]).map(y => y + "-" + month + "-" + day)
      .filter(validDate).sort((a,b) => Math.abs(new Date(a) - new Date(anchor)) - Math.abs(new Date(b) - new Date(anchor)));
    if (!choices.length) return {text:value, due:"", review:false};
    return {text:value.slice(0, match.index).trim(), due:choices[0], review:!explicit};
  }
  function parseSteps(value, anchor=today()) {
    const lines = String(value ?? "").split(/\r?\n/).map(text).filter(Boolean);
    const bullet = /^(?:[-–—•*]\s+|\d+[.)]\s+)/;
    const heading = /^(?:нужно\s+выполнить|не\s+выполнено|выполнено|выполенено)\s*:?\s*$/i;
    const structured = lines.some(line => bullet.test(line) || heading.test(line));
    const parts = [];
    let done = false, previous = null;
    for (const line of lines) {
      if (heading.test(line)) {
        done = /^выпо/i.test(line);
        previous = null;
      } else if ((bullet.test(line) || !structured || !previous) && !/^[-–—]+$/.test(line)) {
        previous = {value:line.replace(bullet, ""), done};
        parts.push(previous);
      } else if (previous) {
        previous.value += "\n" + line;
      }
    }
    return parts.map(part => ({id:uid(), ...detectDue(part.value, anchor), done:part.done, completedOn:""}));
  }
  function formatSteps(steps) {
    const block = (title, done) => [title, ...steps.filter(step => step.done === done).map(step =>
      "- " + step.text + (step.due ? ", срок до " + dateLabel(step.due) : ""))].join("\n");
    return block("Нужно выполнить:", false) + "\n\n" + block("Выполнено:", true);
  }
  function deadlineGroup(step, day=today()) {
    if (step.done) return "";
    if (!step.due) return "none";
    if (step.due < day) return "overdue";
    if (step.due === day) return "today";
    return step.due <= addDays(day,7) ? "week" : "later";
  }

  function demo() {
    const day = today(), y = Number(day.slice(0,4)), current = Math.ceil(Number(day.slice(5,7))/3);
    const quarters = Array.from({length:6}, (_, i) => ((current + i - 1) % 4 + 1) + " кв " + (y + Math.floor((current+i-1)/4)));
    const rows = [
      ["FTE","Проверка документов при приёмке","Алексей М.","В работе","A",[
        ["Сверить результаты на второй выборке",-3,false],["Подготовить перечень расхождений",0,false],["Проверить документы первой выборки",-8,true]],
        "На первой вымышленной выборке обработали 120 документов. Для 96 нашли подходящий заказ."],
      ["FTE","Единый канал обращений","Марина Л.","В работе","A",[
        ["Согласовать текст уведомления для пилота",0,false],["Получить обратную связь участников",3,false],["Собрать типовые обращения",-5,true]],""],
      ["FTE","Шаблоны заказов","Алексей М.","Подготовка","B",[
        ["Проверить один шаблон у поставщика",-1,false],["Подготовить инструкцию",6,false]],""],
      ["ТМЗ","График заказов на пилоте","Елена К.","В работе","A",[
        ["Проверить доступность к концу недели",2,false],["Собрать исходные данные по запасу",-4,true]],
        "В демонстрационной группе пять поставщиков. Исходные данные собраны, результат пилота предстоит измерить."],
      ["ТМЗ","Прогноз даты прихода","Марина Л.","Подготовка","B",[
        ["Уточнить, кто обновляет плановую дату",null,false],["Сверить прогноз и факт по пилотным поставкам",7,false]],""],
      ["ТМЗ","Проверка резерва перед заказом","Елена К.","В работе","B",[
        ["Описать два случая с чужим резервом",null,false],["Проверить расчёт на копии реестра",12,false]],""],
      ["Уточняю","Отчёт по отклонениям доставки","—","Уточняю","C",[
        ["Собрать вопросы к пользователям отчёта",null,false]],""]
    ];
    const numbers = {"FTE":0,"ТМЗ":0,"Уточняю":0};
    const tasks = rows.map(([sheet,title,owner,status,priority,items,result], i) => {
      const number = String(++numbers[sheet]), code = sheet + " " + number;
      return {id:code,sheet,code,number,title,owner,status,priority,jira:"",plan:"",result,
        resultOn:result ? addDays(day,-2) : "", excelRow:Number(number)+1,
        marks:sheet==="Уточняю" ? [] : (i<3 ? ["О","М","Р","","",""] : ["","О","М","","И",""]),
        steps:items.map(([name,offset,done]) => ({id:uid(),text:name,due:offset===null?"":addDays(day,offset),
          done,completedOn:done ? addDays(day,offset) : "",review:false}))};
    });
    return {schema:1,kind:"demo",name:"Учебный беклог",tasks,quarters,legend:{...legend},workbook:null,report:null};
  }
  function validate(value) {
    const string = (v, size=32000) => typeof v === "string" && v.length <= size;
    const validOptionalDate = date => date === "" || validDate(date);
    if (value?.schema !== 1 || !["demo","excel"].includes(value.kind) || !string(value.name,256) ||
        !Array.isArray(value.tasks) || value.tasks.length > 500 || !Array.isArray(value.quarters) ||
        value.quarters.length > 30 || value.quarters.some(v=>!string(v,50)) || !value.legend ||
        Object.values(value.legend).some(v=>!string(v,200))) throw Error("Неподдерживаемый формат данных беклога.");
    const ids = new Set();
    for (const task of value.tasks) {
      if (!task || !string(task.id,400) || ids.has(task.id) || !sheets.includes(task.sheet) ||
          !["code","number","title","owner","status","priority","jira","plan","result"].every(k=>string(task[k])) ||
          !validOptionalDate(task.resultOn) || !Number.isInteger(task.excelRow) || task.excelRow < 2 ||
          !Array.isArray(task.marks) || task.marks.some(v=>!string(v,60)) ||
          !Array.isArray(task.steps) || task.steps.length > 300) throw Error("Некорректная карточка задачи.");
      ids.add(task.id);
      const stepIds = new Set();
      for (const step of task.steps) {
        if (!step || !string(step.id,100) || stepIds.has(step.id) || !string(step.text) || !step.text ||
            !validOptionalDate(step.due) || !validOptionalDate(step.completedOn) ||
            typeof step.done !== "boolean" || typeof step.review !== "boolean") throw Error("Некорректный шаг задачи.");
        stepIds.add(step.id);
      }
    }
    if (value.kind === "excel" && (!value.workbook || !string(value.workbook.base64,700000) || !string(value.workbook.name,256)))
      throw Error("Исходная книга Excel отсутствует или слишком большая.");
    if (value.kind === "demo" && value.workbook !== null) throw Error("В учебном беклоге обнаружено вложение.");
    if (value.report !== null && (!value.report || !string(value.report.text,160000) ||
        !validOptionalDate(value.report.from) || !validOptionalDate(value.report.to) ||
        !Array.isArray(value.report.selected) || value.report.selected.some(id=>!string(id,400)))) throw Error("Некорректный черновик отчёта.");
    return value;
  }

  const past = {"проверить":"Проверили","сверить":"Сверили","подготовить":"Подготовили","собрать":"Собрали",
    "согласовать":"Согласовали","получить":"Получили","уточнить":"Уточнили","провести":"Провели","описать":"Описали",
    "настроить":"Настроили","добавить":"Добавили","изучить":"Изучили","обсудить":"Обсудили","запустить":"Запустили"};
  const sentence = value => /[.!?…]$/.test(value) ? value : value + ".";
  function completedSentence(value) {
    const match = /^([а-яё]+)(\s+.+)$/is.exec(value);
    const verb = match && past[normalize(match[1])];
    if (verb) return sentence(verb + match[2]);
    if (match && /(?:ли|лся|лась|ено|ана|ано)$/i.test(match[1])) return sentence(value);
    return sentence("Завершили шаг «" + value.replace(/[.!?…]$/, "") + "»");
  }
  function report(tasks, selected, from, to, includeUnknown) {
    const start = from || "0000-00-00", end = to || "9999-99-99";
    const active = date => date >= start && date <= end;
    const lines = ["Что сделали по беклогу — итоги на " + dateLabel(to || today())];
    if (from) lines.push("Отчётный период с " + dateLabel(from) + " по " + dateLabel(to || today()));
    for (const sheet of sheets) {
      const chosen = tasks.filter(task => task.sheet === sheet && selected.has(task.id));
      if (!chosen.length) continue;
      lines.push("", sections[sheet]);
      for (const task of chosen) {
        const pieces = task.steps.filter(step => step.done && (step.completedOn ? active(step.completedOn) : includeUnknown))
          .map(step => completedSentence(step.text.replace(/\s+/g," ")));
        if (task.result && (!task.resultOn || active(task.resultOn))) pieces.push(sentence(task.result.replace(/\s+/g, " ")));
        const pending = task.steps.filter(step => !step.done).sort((a,b)=>(a.due || "9999").localeCompare(b.due || "9999"))[0];
        if (pending) {
          const action = pending.text.replace(/\s+/g," ").replace(/[.!?…]$/, "");
          pieces.push("Следующий шаг — " + action.charAt(0).toLocaleLowerCase("ru") + action.slice(1) +
            (pending.due ? ", срок до " + dateLabel(pending.due) : "") + ".");
        }
        const heading = (task.number ? task.number + ". " : "") + sentence(task.title);
        lines.push(heading + (pieces.length ? " " + pieces.join(" ") : ""));
      }
    }
    return lines.join("\n");
  }

  globalThis.Backlog = {sheets,sections,legend,uid,text,normalize,copy,today,addDays,validDate,dateLabel,
    parseSteps,formatSteps,deadlineGroup,demo,validate,report};
})();
