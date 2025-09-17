export const $  = (s, root=document) => root.querySelector(s);
export const $$ = (s, root=document) => Array.from(root.querySelectorAll(s));
export const clear = el => { while (el.firstChild) el.removeChild(el.firstChild); };
export const div = cls => { const d=document.createElement('div'); d.className=cls||''; return d; };
export const tdText = v => { const t=document.createElement('td'); t.textContent = v==null ? '' : String(v); return t; };
export const debounce = (fn, delay=200) => { let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a), delay); }; };

let busyEl=null;
export function setBusy(on, text='작업 중…'){
  if (on){
    if (busyEl) return;
    busyEl=document.createElement('div'); busyEl.className='busy';
    const sp=document.createElement('div'); sp.className='spinner'; sp.textContent=text;
    busyEl.append(sp); document.body.append(busyEl);
  } else { if (busyEl){ busyEl.remove(); busyEl=null; } }
}

export function toast(msg, kind='info', ms=2600){
  let wrap = $('.toast-wrap');
  if (!wrap){ wrap=document.createElement('div'); wrap.className='toast-wrap'; document.body.append(wrap); }
  const t = document.createElement('div');
  t.className = 'toast' + (kind==='ok'?' ok':kind==='err'?' err':'');
  t.textContent = msg; wrap.append(t);
  setTimeout(()=>{ t.remove(); if(!wrap.children.length) wrap.remove(); }, ms);
}

window.addEventListener('error', e => toast(`에러: ${e.message}`,'err',4200));
window.addEventListener('unhandledrejection', e => toast(`에러: ${e.reason}`,'err',4200));
