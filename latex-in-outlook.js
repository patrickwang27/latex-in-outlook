// ==UserScript==
// @name         LaTeX-in-Outlook
// @namespace    owa-local-tex
// @version      1.0.0
// @description  Local LaTeX rendering in OWA: Auto-Paste (hold Ctrl/Cmd+V) or Auto (data-URI). Correct display sizing + reliable Undo via alt-tag encoding.
// @match        https://outlook.office.com/mail/*
// @match        https://outlook.office365.com/mail/*
// @match        https://outlook.live.com/mail/*
// @grant        GM_addStyle
// @grant        GM_xmlhttpRequest
// @connect      127.0.0.1
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  // --- Config ---
  const SERVER = "http://127.0.0.1:8765/render"; // your local Flask endpoint
  const DPI    = 350;                              // render quality

  // --- UI ---
  GM_addStyle(`
    .texk-toolbar{position:fixed;right:12px;bottom:12px;z-index:2147483646;display:flex;gap:6px;flex-wrap:wrap;padding:6px;border-radius:12px;background:rgba(255,255,255,.98);box-shadow:0 2px 10px rgba(0,0,0,.12);border:1px solid rgba(0,0,0,.12);font:12px/1.2 system-ui,Segoe UI,Arial}
    .texk-btn{padding:6px 10px;border-radius:999px;border:1px solid rgba(0,0,0,.15);background:#fff;cursor:pointer;user-select:none}
    .texk-btn:hover{background:#f7f7f7}
    .texk-badge{position:fixed;left:10px;bottom:10px;z-index:2147483646;padding:4px 8px;border-radius:6px;background:rgba(0,0,0,.65);color:#fff;font:11px/1.2 system-ui,Segoe UI,Arial;opacity:.9}
    .texk-progress{position:fixed;left:10px;bottom:34px;z-index:2147483646;padding:4px 8px;border-radius:6px;background:rgba(0,0,0,.35);color:#fff;font:11px/1.2 system-ui,Segoe UI,Arial}
    .texk-hl{background:rgba(255,235,59,.35);padding:0 4px}
  `);

  const badge    = mkBadge('TeX: ready (local)');
  const progress = mkBadge('', 'texk-progress');

  function mkBadge(text, cls='texk-badge'){
    const el = document.createElement('div');
    el.className = cls; el.textContent = text;
    document.documentElement.appendChild(el);
    return t => (el.textContent = t);
  }

  const tb = document.createElement('div');
  tb.className = 'texk-toolbar';
  document.documentElement.appendChild(tb);

  mkBtn('Auto-Paste',   'Render via local LaTeX; hold Ctrl/Cmd+V while it walks equations', autoPasteMode);
  mkBtn('Auto (data-URI)','Fully automatic (data: images). May be stripped on send.',        autoDataUriMode);
  mkBtn('Undo',         'Restore images back to TeX (keeps $$ for display)',                 undoImages);

  function mkBtn(label, title, fn){
    const b=document.createElement('button');
    b.className='texk-btn'; b.textContent=label; b.title=title;
    b.addEventListener('click', fn); tb.appendChild(b);
  }

  // --- Helpers ---
  const getBody=()=>{
    let el=document.activeElement;
    if(el&&el.isContentEditable) return el;
    const cs=[...document.querySelectorAll('[contenteditable="true"]')];
    if(cs.length) return cs.sort((a,b)=>b.getBoundingClientRect().width-a.getBoundingClientRect().width)[0];
    for (const f of document.querySelectorAll('iframe')){
      try{
        const d=f.contentDocument; if(!d) continue;
        const e=d.activeElement; if(e&&e.isContentEditable) return e;
        const fs=[...d.querySelectorAll('[contenteditable="true"]')];
        if(fs.length) return fs.sort((a,b)=>b.getBoundingClientRect().width-a.getBoundingClientRect().width)[0];
      }catch{}
    }
    return null;
  };

  function wrapTeXInBody(body){
    // avoid double wrap
    if (body.querySelector('.texk-eqn')) return;
    let s = body.innerHTML;
    // display $$...$$ → neutral token (no $$ inside)
    s = s.replace(/\$\$([\s\S]*?)\$\$/g, (_, inner) =>
      `<span class="texk-eqn" data-tex="${encodeURIComponent(inner.trim())}" data-display="true" contenteditable="false">[eqn]</span>`
    );
    // inline $...$ (avoid $$)
    s = s.replace(/(^|[^$])\$([^\n][\s\S]*?)\$(?!\$)/g, (m, pre, inner) =>
      `${pre}<span class="texk-eqn" data-tex="${encodeURIComponent(inner.trim())}" data-display="false" contenteditable="false">[eqn]</span>`
    );
    body.innerHTML = s;
  }

  function httpRender(tex, display){
    return new Promise((resolve, reject)=>{
      GM_xmlhttpRequest({
        method: 'POST', url: SERVER, headers: {'Content-Type':'application/json'},
        data: JSON.stringify({ tex, display, dpi: DPI }), responseType: 'arraybuffer',
        onload: r => {
          if (r.status>=200 && r.status<300) resolve(new Blob([r.response], {type:'image/png'}));
          else reject(new Error(`HTTP ${r.status}`));
        },
        onerror: () => reject(new Error('network error to local renderer')),
        ontimeout: () => reject(new Error('timeout to local renderer')),
      });
    });
  }

  const blobToDataURL = blob => new Promise((res,rej)=>{ const fr=new FileReader(); fr.onload=()=>res(fr.result); fr.onerror=rej; fr.readAsDataURL(blob); });

  // --- Image sizing ---
  function applyImageStyle(img, display){
    if (display){
      // balanced display size (~two text lines), centered
      img.style.height      = '2.0em';
      img.style.width       = 'auto';
      img.style.maxWidth    = '90%';
      img.style.display     = 'block';
      img.style.margin      = '0.5em auto';
      img.style.verticalAlign = 'middle';
    } else {
      // inline: match text x-height and baseline
      img.style.height        = '1em';
      img.style.width         = 'auto';
      img.style.maxWidth      = 'none';
      img.style.display       = 'inline-block';
      img.style.verticalAlign = '-0.15em';
    }
  }

  // --- Alt tag encoding for reliable Undo ---
  const altMark = (tex, display) => `texk:${display?'d':'i'}:${tex}`;
  function parseAlt(alt){
    // texk:d:<tex> or texk:i:<tex>
    if (!alt || typeof alt !== 'string') return null;
    const m = /^texk:(d|i):(.*)$/s.exec(alt);
    if (!m) return null;
    return { display: m[1]==='d', tex: m[2] };
  }

  // ============================================================
  // MODE 1: Auto (data-URI) — zero gestures; may be stripped on send
  // ============================================================
  async function autoDataUriMode(){
    const body=getBody(); if(!body){ alert('Click in the message body first.'); return; }
    wrapTeXInBody(body);
    const eqs=[...body.querySelectorAll('.texk-eqn')];
    if(!eqs.length){ alert('No $…$ / $$…$$ found.'); return; }

    badge(`TeX: Auto (data-URI) ${eqs.length}`);
    progress(`0 / ${eqs.length}`);

    for (let i=0;i<eqs.length;i++){
      const n        = eqs[i];
      const tex      = decodeURIComponent(n.dataset.tex||'');
      const display  = n.dataset.display === 'true';

      let blob;
      try { blob = await httpRender(tex, display); }
      catch(e){ alert(`Render failed for equation ${i+1}: ${e.message}`); return; }

      const dataUrl = await blobToDataURL(blob);

      const img = document.createElement('img');
      img.src = dataUrl;
      img.alt = altMark(tex, display);          // <<< encode display/inline in alt
      applyImageStyle(img, display);

      const wrap = document.createElement('span');
      wrap.className = 'texk-png-wrap';
      wrap.dataset.tex = encodeURIComponent(tex);
      wrap.dataset.display = display ? 'true' : 'false';
      wrap.appendChild(img);

      n.replaceWith(wrap);
      progress(`${i+1} / ${eqs.length}`);
    }
    badge('TeX: Auto (data-URI) done ✓');
    progress('');
  }

  // ============================================================
  // MODE 2: Auto-Paste — you hold/tap Ctrl/Cmd+V; images get uploaded inline
  // ============================================================
  async function autoPasteMode(){
    const body=getBody(); if(!body){ alert('Click in the message body first.'); return; }
    wrapTeXInBody(body);
    const doc = body.ownerDocument;
    const eqs = [...body.querySelectorAll('.texk-eqn')];
    if(!eqs.length){ alert('No $…$ / $$…$$ found.'); return; }

    badge(`TeX: Auto-Paste (${eqs.length}). Hold Ctrl/Cmd+V now…`);
    progress(`0 / ${eqs.length}`);

    const wait = ms => new Promise(r=>setTimeout(r,ms));

    // After user paste, wrap the nearest <img>, style it, and stamp alt with display info
    function tagNearestPastedImage(anchorSpan, tex, display){
      const parent = anchorSpan.parentNode || body;
      let candidate = null;

      if (anchorSpan.nextSibling instanceof Element && anchorSpan.nextSibling.tagName === 'IMG') candidate = anchorSpan.nextSibling;
      else if (anchorSpan.previousSibling instanceof Element && anchorSpan.previousSibling.tagName === 'IMG') candidate = anchorSpan.previousSibling;
      else {
        const imgs = [...parent.querySelectorAll('img')];
        if (imgs.length) candidate = imgs[imgs.length - 1];
      }
      if (!candidate) return false;

      applyImageStyle(candidate, display);
      candidate.alt = altMark(tex, display);     // <<< encode here too

      const wrap = doc.createElement('span');
      wrap.className = 'texk-png-wrap';
      wrap.dataset.tex = encodeURIComponent(tex);
      wrap.dataset.display = display ? 'true' : 'false';
      candidate.replaceWith(wrap); wrap.appendChild(candidate);
      return true;
    }

    for (let i=0;i<eqs.length;i++){
      const n        = eqs[i];
      const tex      = decodeURIComponent(n.dataset.tex||'');
      const display  = n.dataset.display === 'true';

      let blob;
      try { blob = await httpRender(tex, display); }
      catch(e){ alert(`Render failed for equation ${i+1}: ${e.message}`); return; }

      try { await navigator.clipboard.write([new ClipboardItem({'image/png': blob})]); }
      catch { alert('Clipboard write blocked. Keep focus, then hold/tap Ctrl/Cmd+V.'); return; }

      // placeholder + caret
      const ph = doc.createElement('span');
      ph.className = 'texk-hl';
      ph.textContent = '[paste image here]';
      ph.dataset.tex = encodeURIComponent(tex);
      ph.dataset.display = display ? 'true' : 'false';
      n.replaceWith(ph);

      const sel = doc.getSelection(); const r = doc.createRange();
      r.setStartAfter(ph); r.collapse(true);
      sel.removeAllRanges(); sel.addRange(r); body.focus();

      progress(`${i+1} / ${eqs.length} — paste now`);
      await wait(1600);                                 // brief window for your paste
      tagNearestPastedImage(ph, tex, display);
      if (ph.isConnected) ph.remove();
    }

    badge('TeX: Auto-Paste done ✓');
    progress('');
  }

  // ------------------------------------------------------------
  // Undo (now robust): use wrapper metadata; else parse alt "texk:d:" / "texk:i:"
  // ------------------------------------------------------------
  async function undoImages(){
    const body = getBody(); if(!body){ alert('Click in the message body first.'); return; }

    // Preferred path: our wrapped images with preserved metadata
    const wraps = [...body.querySelectorAll('span.texk-png-wrap')];
    let restored = 0;

    if (wraps.length){
      for (const w of wraps){
        let tex      = decodeURIComponent(w.dataset.tex || '');
        let display  = (w.dataset.display === 'true');

        // If wrapper attrs were stripped, try image alt
        if (!tex || w.dataset.display == null){
          const img = w.querySelector('img');
          const info = img && parseAlt(img.alt);
          if (info){ tex = info.tex; display = info.display; }
        }

        w.replaceWith(document.createTextNode(display ? `$$${tex}$$` : `$${tex}$`));
        restored++;
      }
      alert(`Restored ${restored} equation(s) to TeX.`);
      return;
    }

    // Fallback: plain <img> nodes (no wrapper). Parse alt marker.
    const imgs = [...body.querySelectorAll('img')];
    for (const img of imgs){
      const info = parseAlt(img.alt);
      if (!info) continue;
      img.replaceWith(document.createTextNode(info.display ? `$$${info.tex}$$` : `$${info.tex}$`));
      restored++;
    }
    alert(restored ? `Restored ${restored} equation(s) to TeX.` : 'No TeX-like images found.');
  }

})();
