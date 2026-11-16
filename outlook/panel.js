'use strict';
/* Enviar a Kanban · complemento de Outlook (Office.js).
   Crea una tarjeta en el tablero de Seguimiento de Proyectos a partir del correo abierto.
   Principio: la tarjeta SIEMPRE se crea. El texto largo se recorta; los archivos que no caben o fallan quedan anotados. */
(function () {
  var CFG_KEY = 'gloria_ol_cfg_v1', LAST_KEY = 'gloria_ol_last_v1';
  var $app = document.getElementById('app'), $cfgBtn = document.getElementById('btn-cfg');
  var cfg = null, cat = null, mail = null, busy = false, pendingItem = false, staged = null;

  /* ---------- utilidades ---------- */
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function store(k, v) { try { if (v === undefined) return JSON.parse(localStorage.getItem(k) || 'null'); if (v === null) localStorage.removeItem(k); else localStorage.setItem(k, JSON.stringify(v)); } catch (e) { return null; } }
  function fmtSz(n) { n = +n || 0; return n < 1024 ? n + ' B' : n < 1048576 ? Math.round(n / 1024) + ' KB' : (n / 1048576).toFixed(n < 10485760 ? 1 : 0) + ' MB'; }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function b64FromBytes(u8) { var s = '', i; for (i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s); }
  function b64FromText(t) { return b64FromBytes(new TextEncoder().encode(t)); }
  function nombreArchivo(s, ext) { s = String(s || 'correo').replace(/[\\\/:*?"<>|\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80); return (s || 'correo') + ext; }
  function titulo(asunto) { return String(asunto || '').replace(/^\s*(\[[^\]]{1,30}\]\s*)?((re|rv|fw|fwd|enc|reenv)\s*:\s*)+/i, '').replace(/^\s*\[(externo|external)\]\s*/i, '').trim(); }
  function cut(s, n) { return s.length > n ? s.slice(0, n - 1).replace(/\s+\S*$/, '') + '…' : s; }
  function limpiaTexto(t) {
    t = String(t || '').replace(/\r\n?/g, '\n');
    var m = t.search(/\n\s*(-{2,}\s*(Original Message|Mensaje original)|De:\s.+\n\s*(Enviado|Fecha|Sent)\s*:|From:\s.+\n\s*Sent\s*:|El .{5,90} escribi[óo]:|On .{5,90} wrote:)/i);
    if (m > 40) t = t.slice(0, m);
    t = t.split('\n').filter(function (l) { return !/^\s*>/.test(l); }).join('\n');
    return t.replace(/[ \t]+/g, ' ').replace(/\n[ \t]+/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  }
  function oAsync(fn) { return new Promise(function (res, rej) { try { fn(function (r) { if (r.status === Office.AsyncResultStatus.Succeeded) res(r.value); else rej(new Error((r.error && r.error.message) || 'Outlook no pudo completar la operación.')); }); } catch (e) { rej(e); } }); }
  function ensureDate(d) { try { return d ? new Date(d).toISOString() : ''; } catch (e) { return ''; } }

  /* ---------- servicio ---------- */
  var MSG = { 'bad-key': 'La clave no es válida.', 'forbidden': 'Esta clave no permite crear tarjetas.', 'locked': 'Demasiados intentos con clave incorrecta: el servicio se bloqueó unos 10 minutos.', 'sin-base': 'La base compartida aún está vacía: carga primero el Excel en la aplicación.', 'proyecto': 'El proyecto ya no existe.', 'too-big': 'La base compartida está llena.', 'archivo-grande': 'El archivo supera el máximo permitido.' };
  function rpc(a, extra, tmo) {
    var ctl = typeof AbortController !== 'undefined' ? new AbortController() : null, to = ctl ? setTimeout(function () { ctl.abort(); }, tmo || 60000) : 0;
    var body = JSON.stringify(Object.assign({ a: a, k: cfg.key, who: cfg.who }, extra || {}));
    return fetch(cfg.url, { method: 'POST', body: body, headers: { 'Content-Type': 'text/plain;charset=utf-8' }, redirect: 'follow', cache: 'no-store', credentials: 'omit', referrerPolicy: 'no-referrer', signal: ctl ? ctl.signal : undefined })
      .catch(function (e) { var er = new Error('Sin conexión con el servicio. Revisa tu red y la dirección.'); er.network = true; throw er; })
      .then(function (res) { if (to) clearTimeout(to); if (!res.ok) { var er = new Error('El servicio respondió con error (' + res.status + ').'); er.network = true; throw er; } return res.json().catch(function () { throw new Error('El servicio no respondió como se esperaba: revisa que la dirección termine en /exec.'); }); })
      .then(function (j) { if (j && j.ok) return j; var er = new Error(MSG[j && j.code] || (j && j.msg) || 'El servicio devolvió un error.'); er.code = j && j.code; er.data = j; throw er; });
  }
  function urlOk(u) { return /^https:\/\/script\.google\.com\/(a\/[^/]+\/)?macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(u); }

  /* ---------- lectura del correo abierto ---------- */
  function readMail() {
    var it = Office.context.mailbox.item;
    if (!it) return Promise.resolve(null);
    var m = { item: it, asunto: it.subject || '', de: '', fecha: ensureDate(it.dateTimeCreated), para: '', adj: [], msgId: it.internetMessageId || it.itemId || '', enlace: '' };
    try { m.de = it.from ? (it.from.displayName ? it.from.displayName + ' <' + it.from.emailAddress + '>' : it.from.emailAddress) : ''; } catch (e) { /* */ }
    try { m.para = (it.to || []).map(function (x) { return x.displayName || x.emailAddress; }).join('; '); } catch (e) { /* */ }
    try { if (it.itemId && Office.context.mailbox.convertToRestId) m.enlace = 'https://outlook.office.com/owa/?ItemID=' + encodeURIComponent(Office.context.mailbox.convertToRestId(it.itemId, Office.MailboxEnums.RestVersion.v2_0)) + '&exvsurl=1&viewmodel=ReadMessageItem'; } catch (e) { /* */ }
    m.adj = (it.attachments || []).map(function (a) { return { id: a.id, nombre: a.name || 'archivo', tam: +a.size || 0, inline: !!a.isInline, tipo: a.attachmentType || 'file', contentId: a.contentId || '' }; });
    return Promise.all([
      oAsync(function (cb) { it.body.getAsync(Office.CoercionType.Text, cb); }).catch(function () { return ''; }),
      oAsync(function (cb) { it.body.getAsync(Office.CoercionType.Html, cb); }).catch(function () { return ''; })
    ]).then(function (r) { m.texto = r[0] || ''; m.html = r[1] || ''; return m; });
  }
  var MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' };
  function mimeDe(n) { var e = String(n).split('.').pop().toLowerCase(); return MIME[e] || 'application/octet-stream'; }
  function esImagen(a) { return /\.(png|jpe?g|gif|bmp|webp)$/i.test(a.nombre); }
  function contenido(m, id) { return oAsync(function (cb) { m.item.getAttachmentContentAsync(id, cb); }); }

  /* HTML del correo: sin scripts ni objetos activos; las imágenes incrustadas (cid:) se convierten en datos para que el archivo sea autónomo. */
  function htmlLimpio(m, imgs) {
    var doc = new DOMParser().parseFromString(m.html || '<pre></pre>', 'text/html');
    ['script', 'iframe', 'object', 'embed', 'link', 'base', 'form', 'meta'].forEach(function (t) { Array.prototype.slice.call(doc.querySelectorAll(t)).forEach(function (n) { n.remove(); }); });
    Array.prototype.slice.call(doc.querySelectorAll('*')).forEach(function (n) { Array.prototype.slice.call(n.attributes).forEach(function (at) { if (/^on/i.test(at.name)) n.removeAttribute(at.name); }); });
    var usadas = {}, sinImg = 0;
    Array.prototype.slice.call(doc.querySelectorAll('img')).forEach(function (img) {
      var src = img.getAttribute('src') || '', cid = /^cid:/i.test(src) ? src.slice(4) : '', k = cid.split('@')[0].toLowerCase(), hit = null;
      if (cid) imgs.forEach(function (x) { if (!hit && ((x.contentId && x.contentId.replace(/^<|>$/g, '').toLowerCase() === cid.toLowerCase()) || x.nombre.toLowerCase() === k)) hit = x; });
      if (hit && hit.dataUri) { img.setAttribute('src', hit.dataUri); usadas[hit.id] = 1; }
      else if (!/^data:/i.test(src)) { var alt = img.getAttribute('alt') || 'imagen'; img.removeAttribute('src'); img.setAttribute('alt', '[' + alt + ' · no incluida]'); sinImg++; }
    });
    return { html: doc.body.innerHTML, usadas: usadas, sinImg: sinImg };
  }
  function fechaLarga(iso) { try { return new Date(iso).toLocaleString('es-PE', { dateStyle: 'full', timeStyle: 'short' }); } catch (e) { return iso || ''; } }
  function docCorreo(m, cuerpo, aviso) {
    return '<!doctype html><html lang="es"><head><meta charset="utf-8"><title>' + esc(m.asunto) + '</title><style>body{font:14px Segoe UI,Arial,sans-serif;color:#222;margin:14px}.h{border-bottom:1px solid #ccc;padding-bottom:8px;margin-bottom:12px;font-size:13px}.h b{font-size:15px}img{max-width:100%;height:auto}</style></head><body><div class="h"><b>' + esc(m.asunto) + '</b><br>De: ' + esc(m.de) + '<br>' + (m.para ? 'Para: ' + esc(m.para) + '<br>' : '') + 'Fecha: ' + esc(fechaLarga(m.fecha)) + (aviso ? '<br><i>' + esc(aviso) + '</i>' : '') + '</div>' + cuerpo + '</body></html>';
  }

  /* ---------- pantallas ---------- */
  function setApp(html) { $app.innerHTML = html; }
  function showConnect(err) {
    $cfgBtn.hidden = true;
    var d = cfg || {};
    setApp('<form id="f-cfg" novalidate><p class="hint">Pide al administrador la <b>dirección del servicio</b> (termina en /exec) y la <b>clave de captura</b>. Se guardan solo en este equipo.</p>' +
      (err ? '<div class="msg err" role="alert">' + esc(err) + '</div>' : '') +
      '<div class="fld"><label for="u">Dirección del servicio</label><input id="u" type="url" autocomplete="off" placeholder="https://script.google.com/macros/s/…/exec" value="' + esc(d.url || '') + '"></div>' +
      '<div class="fld"><label for="k">Clave de captura</label><input id="k" type="password" autocomplete="off" value="' + esc(d.key || '') + '"></div>' +
      '<div class="fld"><label for="w">Tu nombre</label><input id="w" type="text" autocomplete="name" value="' + esc(d.who || userName()) + '"><span class="hint">Queda como autor de las tarjetas.</span></div>' +
      '<button class="btn pri" type="submit">Probar y conectar</button></form>');
    document.getElementById('f-cfg').onsubmit = function (e) {
      e.preventDefault();
      var u = document.getElementById('u').value.trim(), k = document.getElementById('k').value, w = document.getElementById('w').value.trim();
      cfg = { url: u, key: k, who: w };                          // lo escrito se conserva si hay que corregir algo
      if (!urlOk(u)) return showConnect('La dirección debe ser la URL de la implementación de Apps Script y terminar en /exec.');
      if (k.length < 12) return showConnect('La clave tiene al menos 12 caracteres.');
      if (!w) return showConnect('Escribe tu nombre.');
      var b = e.target.querySelector('button'); b.disabled = true; b.textContent = 'Conectando…';
      loadCatalog().then(function () { store(CFG_KEY, cfg); start(true); }).catch(function (er) { showConnect(er.message); });
    };
  }
  function userName() { try { return Office.context.mailbox.userProfile.displayName || ''; } catch (e) { return ''; } }
  function loadCatalog() { return rpc('catalogo', null, 45000).then(function (c) { cat = c; return c; }); }

  function optProj(p, sel) { return '<option value="' + esc(p.id) + '"' + (p.id === sel ? ' selected' : '') + '>' + esc(p.nombre + (p.codigo ? ' · ' + p.codigo : '')) + '</option>'; }
  function optFases(p, selId) {
    if (!p) return '<option value="">(Elige primero el proyecto)</option>';
    if (!p.fases.length) return '<option value="">(El proyecto no tiene etapas abiertas)</option>';
    var sel = selId || (p.fases.filter(function (f) { return f.sug; })[0] || p.fases[0]).id;
    return p.fases.map(function (f) { return '<option value="' + esc(f.id) + '"' + (f.id === sel ? ' selected' : '') + '>' + esc(f.etapa + ' · ' + f.estado) + '</option>'; }).join('');
  }

  function showForm(res) {
    $cfgBtn.hidden = false;
    if (!mail) { setApp('<p class="muted pad">Abre un correo para enviarlo al tablero.</p>'); return; }
    if (cat.vacio) { setApp('<div class="msg warn">La base compartida aún está vacía. Carga primero el Excel de proyectos en la aplicación.</div>'); return; }
    var last = store(LAST_KEY) || {}, lim = cat.limites, act = cat.proyectos.filter(function (p) { return p.activo; }), oth = cat.proyectos.filter(function (p) { return !p.activo; });
    var pid = (last.proyectoId && cat.proyectos.some(function (p) { return p.id === last.proyectoId; })) ? last.proyectoId : '';
    var proj = cat.proyectos.filter(function (p) { return p.id === pid; })[0] || null;
    var notas = cut(limpiaTexto(mail.texto), 5000);
    var pri = cat.prioridades.length ? cat.prioridades : [{ nombre: 'Media' }];
    var archivos = mail.adj.filter(function (a) { return !a.inline; }), inl = mail.adj.filter(function (a) { return a.inline; });
    var aviso = cat.uso >= 85 ? '<div class="msg err" role="alert"><b>La base está al ' + cat.uso + ' % de su capacidad.</b>Avisa al administrador: si se llena, dejará de guardar.</div>' : (cat.uso >= 60 ? '<div class="msg warn"><b>La base está al ' + cat.uso + ' % de su capacidad.</b>Avisa al administrador para depurar la base.</div>' : '');
    setApp(
      aviso + (res || '') +
      '<div class="src"><span class="muted">Correo de origen</span><b id="o-asunto">' + esc(mail.asunto || '(sin asunto)') + '</b><span>' + esc(mail.de) + (mail.fecha ? ' · ' + esc(new Date(mail.fecha).toLocaleString('es-PE')) : '') + '</span></div>' +
      '<form id="f" novalidate>' +
      '<div class="fld"><label for="t">Actividad a realizar</label><input id="t" type="text" autocomplete="off" value="' + esc(cut(titulo(mail.asunto), lim.titulo)) + '"><span class="hint">Se propone el asunto: edítalo para definir el alcance de la actividad. El asunto original queda guardado como origen.</span></div>' +
      '<div class="fld"><label for="p">Proyecto</label><select id="p">' + (pid ? '' : '<option value="" selected>— Elige el proyecto —</option>') + (act.length ? '<optgroup label="Activos">' + act.map(function (p) { return optProj(p, pid); }).join('') + '</optgroup>' : '') + (oth.length ? '<optgroup label="Otros">' + oth.map(function (p) { return optProj(p, pid); }).join('') + '</optgroup>' : '') + '</select></div>' +
      '<div class="fld"><label for="e">Etapa del proyecto</label><select id="e">' + optFases(proj, last.proyectoId === pid ? last.faseId : '') + '</select></div>' +
      '<div class="row2"><div class="fld"><label for="r">Responsable</label><select id="r"></select></div><div class="fld"><label for="l">Fecha límite</label><input id="l" type="date"></div></div>' +
      '<fieldset><legend>Prioridad</legend><div class="chips">' + pri.map(function (x, i) { return '<label class="chip"><input type="radio" name="pri" value="' + esc(x.nombre) + '"' + (x.nombre === 'Media' || (pri.length === 1) ? ' checked' : '') + '><span>' + esc(x.nombre) + '</span></label>'; }).join('') + '</div></fieldset>' +
      '<div class="fld"><label for="n">Notas / alcance</label><textarea id="n" rows="5">' + esc(notas) + '</textarea><div class="cnt" id="n-c"><span></span><span><button type="button" class="lnk" style="color:#1c3f94" id="n-fill">Usar inicio del correo</button> · <button type="button" class="lnk" style="color:#1c3f94" id="n-clr">Vaciar</button></span></div></div>' +
      '<div class="box"><h2>Qué se guarda con la tarjeta</h2>' +
      '<label class="chk"><input type="checkbox" id="c-mail" checked><span>Correo completo (texto e imágenes)' + (inl.length ? ' · <span class="sz">' + inl.length + ' imagen' + (inl.length === 1 ? '' : 'es') + ' incrustada' + (inl.length === 1 ? '' : 's') + '</span>' : '') + '</span></label>' +
      (archivos.length ? archivos.map(function (a, i) { var big = a.tam > lim.adjMB * 1048576, nube = a.tipo === 'cloud'; return '<label class="chk' + (big || nube ? ' dis' : '') + '"><input type="checkbox" class="c-adj" data-i="' + mail.adj.indexOf(a) + '"' + (big || nube ? ' disabled' : ' checked') + '><span>' + esc(a.nombre) + ' <span class="sz">' + fmtSz(a.tam) + '</span>' + (big ? '<br><span class="sz">No cabe (máx. ' + lim.adjMB + ' MB): quedará anotado sin archivo.</span>' : '') + (nube ? '<br><span class="sz">Está en la nube: no se copia.</span>' : '') + '</span></label>'; }).join('') : '<p class="hint" style="margin:4px 0 0">Este correo no tiene adjuntos.</p>') + '</div>' +
      '<div id="msg" role="alert"></div>' +
      '<button class="btn pri" type="submit" id="go">Crear tarea en el tablero</button></form>');
    fillResp(proj, last.resp);
    var $p = document.getElementById('p'), $n = document.getElementById('n');
    $p.onchange = function () { var pp = cat.proyectos.filter(function (x) { return x.id === $p.value; })[0]; document.getElementById('e').innerHTML = optFases(pp); fillResp(pp); };
    function count() { var n = $n.value.length, c = document.getElementById('n-c').firstChild; c.textContent = n + ' / ' + lim.notas + (n > lim.notas ? ' · se recortará al crear' : ''); document.getElementById('n-c').className = 'cnt' + (n > lim.notas ? ' over' : ''); }
    $n.oninput = count; count();
    document.getElementById('n-fill').onclick = function () { $n.value = notas; count(); };
    document.getElementById('n-clr').onclick = function () { $n.value = ''; count(); $n.focus(); };
    document.getElementById('f').onsubmit = function (e) { e.preventDefault(); submit(false); };
  }
  function fillResp(p, pref) {
    var $r = document.getElementById('r'); if (!$r) return;
    var me = (cfg.who || '').toLowerCase(), names = cat.responsables, hit = names.filter(function (n) { return n.toLowerCase() === me; })[0];
    var def = pref && names.indexOf(pref) >= 0 ? pref : (hit || (p && p.pm && names.indexOf(p.pm) >= 0 ? p.pm : ''));
    $r.innerHTML = '<option value="">(Sin asignar)</option>' + names.map(function (n) { return '<option' + (n === def ? ' selected' : '') + '>' + esc(n) + '</option>'; }).join('');
  }
  function msg(html, cls) { var m = document.getElementById('msg'); if (m) m.innerHTML = html ? '<div class="msg ' + cls + '">' + html + '</div>' : ''; }

  /* ---------- crear la tarjeta ---------- */
  function progress(txt, pct) {
    var m = document.getElementById('msg'); if (!m) return;
    m.innerHTML = '<div class="msg warn"><span id="pg-t"></span><div class="prog"><i id="pg-b"></i></div></div>';
    document.getElementById('pg-t').textContent = txt; document.getElementById('pg-b').style.width = Math.max(4, Math.min(100, pct || 0)) + '%';
  }
  function lee(sel) { var e = document.querySelector(sel); return e ? e.value : ''; }
  function submit(forzar) {
    if (busy) return;
    var tit = lee('#t').trim(), pid = lee('#p');
    if (!pid) { msg('<b>Elige el proyecto.</b>', 'err'); document.getElementById('p').focus(); return; }
    if (!tit) { msg('<b>Escribe la actividad a realizar.</b>', 'err'); document.getElementById('t').focus(); return; }
    var pri = document.querySelector('input[name=pri]:checked'), datos = { proyectoId: pid, faseId: lee('#e'), titulo: tit, resp: lee('#r'), limite: lee('#l'), prioridad: pri ? pri.value : 'Media', notas: lee('#n') };
    var sel = Array.prototype.map.call(document.querySelectorAll('.c-adj:checked'), function (c) { return mail.adj[+c.dataset.i]; });
    var noAdj = mail.adj.filter(function (a) { return !a.inline && sel.indexOf(a) < 0; });
    var conCorreo = document.getElementById('c-mail').checked;
    busy = true; document.getElementById('go').disabled = true; msg('', '');
    store(LAST_KEY, { proyectoId: pid, faseId: datos.faseId, resp: datos.resp });
    var m = mail;
    (forzar ? Promise.resolve({ existentes: [] }) : rpc('capturar', { proyectoId: pid, msgId: m.msgId, verificar: true }, 45000))
      .then(function (v) {
        if (v.existentes && v.existentes.length) {
          busy = false; document.getElementById('go').disabled = false;
          var ul = v.existentes.map(function (x) { return '<li><span class="st">•</span><span>' + esc(x.titulo) + '</span></li>'; }).join('');
          document.getElementById('msg').innerHTML = '<div class="msg warn"><b>Este correo ya tiene tarjeta en ese proyecto</b><ul class="res">' + ul + '</ul><div class="actions"><button type="button" class="btn sm" id="dup-go">Crear otra actividad</button><button type="button" class="btn sm" id="dup-no">Cancelar</button></div></div>';
          document.getElementById('dup-go').onclick = function () { submit(true); };
          document.getElementById('dup-no').onclick = function () { msg('', ''); };
          return null;
        }
        return subirTodo(m, sel, noAdj, conCorreo).then(function (fichas) { return crear(m, datos, fichas, forzar); });
      })
      .catch(function (e) { fin(); msg('<b>No se pudo crear la tarjeta.</b>' + esc(e.message || e) + (staged ? '<div class="actions"><button type="button" class="btn sm" id="retry">Reintentar</button></div>' : ''), 'err'); var r = document.getElementById('retry'); if (r) r.onclick = function () { reintentarCrear(datos, forzar); }; });
  }
  function fin() { busy = false; var g = document.getElementById('go'); if (g) g.disabled = false; if (pendingItem) { pendingItem = false; start(); } }

  /* sube correo completo y adjuntos uno por uno; si algo falla lo anota y sigue */
  function subirTodo(m, sel, noAdj, conCorreo) {
    var fichas = [], total = (conCorreo ? 1 : 0) + sel.length, hecho = 0, lim = cat.limites, chain = Promise.resolve();
    noAdj.forEach(function (a) { fichas.push({ nombre: a.nombre, tam: a.tam, rol: 'adjunto', estado: 'omitido', motivo: a.tam > lim.adjMB * 1048576 ? 'Supera ' + lim.adjMB + ' MB: no se guardó el archivo.' : (a.tipo === 'cloud' ? 'Archivo en la nube: no se copia.' : 'No se guardó el archivo.') }); });
    var imgs = [];
    function paso(txt) { progress(txt, total ? (hecho / total) * 100 : 5); }
    function subir(nombre, tipo, b64, extra) {
      function una() { return rpc('subir', { nombre: nombre, tipo: tipo, b64: b64 }, 180000); }
      return una().catch(function (e) { if (e.network) return sleep(800).then(una); throw e; }).then(function (j) { return Object.assign({ nombre: j.nombre, tam: j.tam, tipo: tipo, fileId: j.fileId, estado: 'ok' }, extra || {}); });
    }
    function falla(a, e, rol) { return { nombre: a.nombre, tam: a.tam || 0, tipo: a.tipo || '', rol: rol || 'adjunto', estado: e && e.code === 'archivo-grande' ? 'omitido' : 'error', motivo: cut(e && e.code === 'archivo-grande' ? 'Supera ' + lim.adjMB + ' MB: no se guardó el archivo.' : 'No se pudo subir: ' + (e && e.message || 'error'), 160) }; }
    /* 1) adjuntos (los inline sin referenciar se tratan después) */
    sel.forEach(function (a) {
      chain = chain.then(function () {
        paso('Guardando ' + a.nombre + ' (' + fmtSz(a.tam) + ')…');
        return contenido(m, a.id).then(function (c) {
          var b64, nombre = a.nombre, tipo = mimeDe(a.nombre);
          if (c.format === 'base64' || c.format === 'Base64') b64 = c.content;
          else if (c.format === 'eml' || c.format === 'Eml') { b64 = b64FromText(c.content); nombre = /\.eml$/i.test(nombre) ? nombre : nombre + '.eml'; tipo = 'message/rfc822'; }
          else throw new Error('Formato de adjunto no admitido (' + c.format + ').');
          return subir(nombre, a.nombre.match(/\.pdf$/i) ? 'application/pdf' : tipo, b64, { rol: 'adjunto' });
        }).then(function (f) { fichas.push(f); }, function (e) { fichas.push(falla(a, e)); }).then(function () { hecho++; });
      });
    });
    /* 2) correo completo */
    if (conCorreo) chain = chain.then(function () {
      paso('Guardando el correo completo…');
      var inl = m.adj.filter(function (a) { return a.inline && esImagen(a); }), pend = inl.slice(), cuota = 8 * 1048576, usado = 0;
      function sig() {
        var a = pend.shift(); if (!a) return Promise.resolve();
        if (usado + a.tam * 1.4 > cuota) return sig();
        return contenido(m, a.id).then(function (c) { if (/base64/i.test(c.format)) { a.dataUri = 'data:' + mimeDe(a.nombre) + ';base64,' + c.content; usado += c.content.length; } }, function () { /* sin esa imagen */ }).then(sig);
      }
      return sig().then(function () {
        var r = htmlLimpio(m, inl), aviso = r.sinImg ? r.sinImg + ' imagen(es) de internet o no disponibles no se incluyeron.' : '', doc = docCorreo(m, r.html, aviso);
        var tamMax = lim.adjMB * 1048576;
        if (doc.length * 1.4 > tamMax) { var r2 = htmlLimpio(m, []); doc = docCorreo(m, r2.html, 'El correo era muy grande: se guardó sin imágenes.'); }
        if (doc.length * 1.4 > tamMax) doc = docCorreo(m, '<pre style="white-space:pre-wrap">' + esc(m.texto) + '</pre>', 'El correo era muy grande: se guardó solo su texto.');
        if (doc.length * 1.4 > tamMax) doc = docCorreo(m, '<pre style="white-space:pre-wrap">' + esc(m.texto.slice(0, Math.floor(tamMax / 2.5))) + '</pre>', 'El correo era muy grande: se guardó solo el inicio de su texto.');
        var nm = nombreArchivo('Correo - ' + titulo(m.asunto), '.html');
        return subir(nm, 'text/html', b64FromText(doc), { rol: 'correo' }).then(function (f) { fichas.unshift(f); }, function (e) { fichas.unshift(falla({ nombre: nm, tam: doc.length }, e, 'correo')); }).then(function () { hecho++; });
      });
    });
    return chain.then(function () { return fichas; });
  }
  function crear(m, datos, fichas, forzar) {
    staged = { m: m, fichas: fichas };
    progress('Creando la tarjeta…', 96);
    var ex = cut(limpiaTexto(m.texto), 5000);
    var body = Object.assign({}, datos, { msgId: m.msgId, forzar: !!forzar, adjuntos: fichas, correo: { asunto: m.asunto, de: m.de, fecha: m.fecha, enlace: m.enlace, extracto: ex, caracteres: (m.texto || '').length } });
    return rpc('capturar', body, 90000).then(function (r) { staged = null; resultado(r, datos, fichas, m); }, function (e) {
      if (e.code === 'duplicado') { staged = null; throw new Error('Ya existe una tarjeta de este correo en ese proyecto.'); }
      throw e;
    });
  }
  function reintentarCrear(datos, forzar) {
    if (!staged || busy) return;
    busy = true; progress('Creando la tarjeta…', 96);
    crear(staged.m, datos, staged.fichas, forzar).catch(function (e) { fin(); msg('<b>No se pudo crear la tarjeta.</b>' + esc(e.message || e) + '<div class="actions"><button type="button" class="btn sm" id="retry">Reintentar</button></div>', 'err'); document.getElementById('retry').onclick = function () { reintentarCrear(datos, forzar); }; });
  }
  function resultado(r, datos, fichas, m) {
    var p = cat.proyectos.filter(function (x) { return x.id === datos.proyectoId; })[0] || {}, rec = r.recortes || {};
    var hayCorreo = fichas.some(function (f) { return f.rol === 'correo' && f.estado === 'ok'; });
    var items = [];
    if (rec.titulo) items.push('El título era más largo que la tarjeta y se recortó.');
    if (rec.notas) items.push('Las notas se recortaron al límite de la tarjeta' + (hayCorreo ? '; el texto completo está en el correo guardado.' : '.'));
    var malos = (r.adjuntos || []).filter(function (a) { return a.estado !== 'ok'; });
    var lista = (r.adjuntos || []).map(function (a) { return '<li><span class="st">' + (a.estado === 'ok' ? '✔' : '⚠') + '</span><span>' + esc(a.nombre) + (a.estado === 'ok' ? '' : ' — ' + esc(a.motivo || 'no se guardó')) + '</span></li>'; }).join('');
    var html = '<div class="msg ' + (malos.length ? 'warn' : 'ok') + '" role="status"><b>Tarea creada en ' + esc(p.nombre || 'el proyecto') + '</b>' + esc((r.fase || '') + (r.fase ? ' · ' : '') + 'To-Do') + '<br><span class="small">Aparece en el tablero en menos de un minuto.</span>' +
      (r.reabrioEtapa ? '<br><span class="small">La etapa «' + esc(r.reabrioEtapa) + '» estaba cerrada y se reabrió.</span>' : '') +
      (items.length ? '<ul class="res">' + items.map(function (x) { return '<li><span class="st">ℹ</span><span>' + esc(x) + '</span></li>'; }).join('') + '</ul>' : '') +
      (lista ? '<ul class="res">' + lista + '</ul>' : '') +
      '<div class="actions">' + (malos.some(function (a) { return a.estado === 'error'; }) ? '<button type="button" class="btn sm" id="rt-adj">Reintentar archivos fallidos</button>' : '') + '<button type="button" class="btn sm" id="otra">Crear otra actividad de este correo</button></div></div>';
    var cambio = pendingItem;                                   // mientras se guardaba, el usuario eligió otro correo (panel anclado)
    busy = false; staged = null; pendingItem = false;
    if (cambio) { html = html.replace(/<div class="actions">.*<\/div><\/div>$/, '</div>'); readMail().then(function (m2) { mail = m2; showForm(html); }); return; }
    showForm(html); document.getElementById('otra').onclick = function () { showForm(''); };
    var rb = document.getElementById('rt-adj'); if (rb) rb.onclick = function () { reintentarAdjuntos(m, r.tareaId, malos); };
    window.scrollTo(0, 0);
  }
  function reintentarAdjuntos(m, tareaId, malos) {
    if (busy) return; busy = true;
    var nombres = malos.filter(function (a) { return a.estado === 'error'; }).map(function (a) { return a.nombre; });
    var sel = m.adj.filter(function (a) { return nombres.indexOf(a.nombre) >= 0 && !a.inline; });
    progress('Reintentando ' + sel.length + ' archivo(s)…', 30);
    subirTodo(m, sel, [], false).then(function (fichas) { return rpc('vincular', { tareaId: tareaId, adjuntos: fichas }, 90000); }).then(function (r) {
      busy = false; var ok = (r.adjuntos || []).filter(function (a) { return a.estado === 'ok'; }).length;
      msg('<b>' + ok + ' de ' + (r.adjuntos || []).length + ' archivo(s) guardados.</b>', ok === (r.adjuntos || []).length ? 'ok' : 'warn');
    }).catch(function (e) { busy = false; msg('<b>No se pudo completar.</b>' + esc(e.message || e), 'err'); });
  }

  /* ---------- arranque ---------- */
  function start(tengoCatalogo) {
    cfg = store(CFG_KEY);
    if (!cfg || !cfg.url || !cfg.key) { showConnect(''); return; }
    setApp('<p class="muted pad">Cargando proyectos…</p>');
    $cfgBtn.hidden = false;
    (tengoCatalogo && cat ? Promise.resolve(cat) : loadCatalog()).then(function () { return readMail(); }).then(function (m) { mail = m; showForm(''); })
      .catch(function (e) {
        if (e.code === 'bad-key' || e.code === 'forbidden') { showConnect(e.message); return; }
        setApp('<div class="msg err" role="alert"><b>No se pudo cargar el catálogo.</b>' + esc(e.message || e) + '<div class="actions"><button type="button" class="btn sm" id="re">Reintentar</button></div></div>');
        document.getElementById('re').onclick = start;
      });
  }
  $cfgBtn.onclick = function () { if (busy) return; showConnect(''); };
  function onItem() {
    if (busy) { pendingItem = true; return; }
    if (!cfg || !cat) return;
    readMail().then(function (m) { mail = m; showForm(''); });
  }
  Office.onReady(function (info) {
    if (info && info.host && info.host !== Office.HostType.Outlook) { setApp('<p class="msg err">Este complemento es solo para Outlook.</p>'); return; }
    try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItem); } catch (e) { /* el panel no está anclado */ }
    start();
  });
})();
