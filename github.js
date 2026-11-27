(() => {
  "use strict";
  const file = "backlog.json";
  const MAX_BYTES = 900000;
  const encoder = new TextEncoder();
  const decoder = new TextDecoder("utf-8",{fatal:true});
  const model = globalThis.Backlog;
  const office = globalThis.BacklogOffice;

  class Conflict extends Error {
    constructor() { super("На GitHub уже есть более свежая версия. Локальные правки сохранены в этом браузере."); }
  }
  function repository(value) {
    const name=value.trim();
    if (!/^[a-zA-Z0-9][a-zA-Z0-9-]{0,38}\/[a-zA-Z0-9_.-]{1,100}$/.test(name) ||
        [".",".."].includes(name.split("/")[1])) throw Error("Введите репозиторий в формате владелец/название.");
    return name;
  }
  function connect(value, accessToken) {
    const name=repository(value);
    let token=accessToken.trim();
    if (!/^github_pat_[A-Za-z0-9_]{10,250}$/.test(token)) throw Error("Нужен fine-grained токен GitHub.");
    const root="https://api.github.com/repos/"+name.split("/").map(encodeURIComponent).join("/");

    async function request(path, options={}) {
      try {
        return await fetch(root+path,{cache:"no-store",redirect:"error",referrerPolicy:"no-referrer",
          ...options,headers:{Accept:"application/vnd.github+json",Authorization:"Bearer "+token,
            ...(options.body?{"Content-Type":"application/json"}:{})}});
      } catch {
        throw Error("GitHub API недоступен. Проверьте соединение и повторите синхронизацию.");
      }
    }
    function failure(response, action) {
      if (response.status===401) return Error("GitHub отклонил ключ. Проверьте срок действия токена.");
      if ([403,404].includes(response.status)) return Error("Нет доступа к закрытому репозиторию. Проверьте имя репозитория и разрешение Contents: Read and write.");
      if (response.status===429) return Error("Достигнут лимит запросов GitHub. Повторите синхронизацию позже.");
      return Error(action+" не удалось (GitHub "+response.status+").");
    }
    async function verify() {
      const response=await request("");
      if (!response.ok) throw failure(response,"Проверить репозиторий");
      const repo=await response.json();
      if (!repo.private || (repo.visibility && repo.visibility!=="private"))
        throw Error("Для синхронизации выберите отдельный репозиторий с видимостью Private.");
    }
    async function read() {
      const response=await request("/contents/"+file);
      if (response.status===404) return {sha:"",data:null};
      if (!response.ok) throw failure(response,"Загрузить беклог");
      const record=await response.json();
      if (record.type!=="file" || record.encoding!=="base64" || !record.sha ||
          typeof record.content!=="string" || record.size>MAX_BYTES)
        throw Error("Размер или формат backlog.json не поддерживается.");
      try {
        return {sha:record.sha,data:model.validate(JSON.parse(decoder.decode(office.decode64(record.content))))};
      } catch {
        throw Error("Файл backlog.json не удалось прочитать. Проверьте, что его создало приложение.");
      }
    }
    async function write(data, sha) {
      model.validate(data);
      const bytes=encoder.encode(JSON.stringify(data));
      if (bytes.length>MAX_BYTES) throw Error("Беклог слишком большой для синхронизации через GitHub. Скачайте локальную копию.");
      await verify();
      const response=await request("/contents/"+file,{method:"PUT",
        body:JSON.stringify({message:"Обновить беклог",content:office.encode64(bytes),...(sha?{sha}:{})})});
      if ([409,422].includes(response.status)) {
        const latest=await read();
        if (latest.sha!==sha) throw new Conflict();
      }
      if (!response.ok) throw failure(response,"Сохранить беклог");
      const result=await response.json();
      if (!result.content?.sha) throw Error("GitHub принял запрос без версии файла. Переподключите хранилище.");
      return result.content.sha;
    }
    return {name,key:name.toLowerCase(),verify,read,write,close:()=>{ token=""; }};
  }
  globalThis.BacklogGitHub={connect,Conflict};
})();
