(() => {
  "use strict";

  const name="lc-backlog-connection-v1";
  const encoder=new TextEncoder(), decoder=new TextDecoder();
  const context=encoder.encode("lc-backlog/github/v1");

  function database() {
    return new Promise((resolve,reject)=>{
      const request=indexedDB.open(name,1);
      request.onupgradeneeded=()=>request.result.createObjectStore("credentials");
      request.onsuccess=()=>resolve(request.result);
      request.onerror=()=>reject(request.error);
      request.onblocked=()=>reject(Error("Хранилище ключа занято другой вкладкой."));
    });
  }
  async function store(mode,operation) {
    const db=await database();
    try {
      return await new Promise((resolve,reject)=>{
        const transaction=db.transaction("credentials",mode);
        const request=operation(transaction.objectStore("credentials"));
        transaction.oncomplete=()=>resolve(request.result);
        transaction.onabort=()=>reject(transaction.error);
        transaction.onerror=()=>reject(transaction.error);
      });
    } finally { db.close(); }
  }
  function available() {
    return Boolean(globalThis.isSecureContext && globalThis.crypto?.subtle && globalThis.indexedDB);
  }
  async function save(repo,token) {
    const data=encoder.encode(JSON.stringify({repo,token}));
    const key=await crypto.subtle.generateKey({name:"AES-GCM",length:256},false,["encrypt","decrypt"]);
    const iv=crypto.getRandomValues(new Uint8Array(12));
    try {
      const sealed=await crypto.subtle.encrypt({name:"AES-GCM",iv,additionalData:context},key,data);
      await store("readwrite",table=>table.put({key,iv,sealed},"active"));
    } finally { data.fill(0); }
  }
  async function load() {
    if (!available()) return null;
    const entry=await store("readonly",table=>table.get("active"));
    if (!entry) return null;
    const data=new Uint8Array(await crypto.subtle.decrypt({name:"AES-GCM",iv:entry.iv,additionalData:context},entry.key,entry.sealed));
    try { return JSON.parse(decoder.decode(data)); }
    finally { data.fill(0); }
  }
  async function clear() {
    if (available()) await store("readwrite",table=>table.delete("active"));
  }

  globalThis.BacklogCredentials={available,save,load,clear};
})();
