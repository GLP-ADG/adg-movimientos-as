// ADG Movimientos — réplica administrada por AS SYSTEMS.
// Usa la misma base de datos de la app ADG Movimientos actual (usuarios/{uid}/movs) y respeta sus reglas.
import { initializeApp, deleteApp } from 'firebase/app';
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut, sendPasswordResetEmail, createUserWithEmailAndPassword } from 'firebase/auth';
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager, collection, doc, getDoc, getDocs, setDoc, updateDoc, deleteDoc, onSnapshot, serverTimestamp } from 'firebase/firestore';
import * as XLSX from 'xlsx';

const cfg = window.FIREBASE_CONFIG;
const app = initializeApp(cfg);
const auth = getAuth(app);
const db = initializeFirestore(app, { localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) });

const TIPOS = ['Gasto', 'Ingreso', 'Traslado'];
const RUBROS_BASE = ['Nómina y prestaciones', 'Combustible', 'Materiales', 'Servicios públicos', 'Arriendo', 'Transporte y viáticos', 'Impuestos', 'Mantenimiento', 'Papelería y oficina', 'Compra de GLP', 'Ventas', 'Préstamos', 'Otros'];
const S = { user: null, perfil: null, movs: [], usuarios: [], ver: null, verMovs: [], vista: 'movs', f: { mes: '', tipo: '', q: '' }, unsub: [] };

const $ = (s, el = document) => el.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pesos = n => '$' + Math.round(Number(n) || 0).toLocaleString('es-CO');
const hoy = () => { const d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 10); };
const ffecha = s => { const [y, m, d] = String(s || '').split('-'); return d ? `${d}/${m}/${y}` : s; };
const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
const esAdmin = () => S.perfil?.rol === 'admin';
const uuid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));

function toast(t, err) { const d = document.createElement('div'); d.className = 'toast' + (err ? ' error' : ''); d.textContent = t; document.body.appendChild(d); setTimeout(() => d.classList.add('ver'), 10); setTimeout(() => d.remove(), err ? 5000 : 2800); }
function errTxt(e) {
  const c = e?.code || '';
  if (/invalid-credential|wrong-password|user-not-found/.test(c)) return 'Correo o contraseña incorrectos.';
  if (/permission-denied/.test(c)) return 'Su usuario no tiene permiso para esta acción.';
  if (/email-already-in-use/.test(c)) return 'Ese correo ya tiene cuenta.';
  if (/weak-password/.test(c)) return 'La contraseña debe tener mínimo 6 caracteres.';
  if (/too-many-requests/.test(c)) return 'Demasiados intentos. Espere unos minutos.';
  return e?.message || String(e);
}
function modal(titulo, html, pie) {
  cerrar(); const m = document.createElement('div'); m.className = 'fondo'; m.id = 'modal';
  m.innerHTML = `<div class="caja"><div class="cab"><h3>${titulo}</h3><button class="x" data-c>×</button></div><div class="cuerpo">${html}</div>${pie ? `<div class="pie">${pie}</div>` : ''}</div>`;
  m.onclick = e => { if (e.target === m || e.target.closest('[data-c]')) cerrar(); };
  document.body.appendChild(m); return m;
}
function cerrar() { $('#modal')?.remove(); }

/* ---------------- sesión ---------------- */
function login(msg = '') {
  S.unsub.forEach(f => f()); S.unsub = [];
  $('#app').innerHTML = `<div class="login"><img src="img/logo-systems.png" alt="AS SYSTEMS" class="logo"><h1>ADG Movimientos</h1>
  <p class="sub">Control de movimientos financieros</p>${msg ? `<div class="aviso">${esc(msg)}</div>` : ''}
  <form id="fl"><label>Correo<input type="email" name="e" required autocomplete="username"></label><label>Contraseña<input type="password" name="p" required autocomplete="current-password"></label>
  <button class="btn grande">Ingresar</button></form><a href="#" id="olvido">Olvidé mi contraseña</a><p class="marca">Administrado por AS SYSTEMS</p></div>`;
  $('#fl').onsubmit = async e => { e.preventDefault(); const f = new FormData(e.target); const b = $('button', e.target); b.disabled = true;
    try { await signInWithEmailAndPassword(auth, f.get('e').trim(), f.get('p')); } catch (er) { toast(errTxt(er), 1); b.disabled = false; } };
  $('#olvido').onclick = async e => { e.preventDefault(); const em = $('[name=e]').value.trim(); if (!em) return toast('Escriba su correo arriba.', 1);
    try { await sendPasswordResetEmail(auth, em); toast('Le enviamos un correo para cambiar la contraseña.'); } catch (er) { toast(errTxt(er), 1); } };
}
onAuthStateChanged(auth, async u => {
  S.user = u; S.perfil = null;
  if (!u) return login();
  try { const d = await getDoc(doc(db, 'usuarios', u.uid)); S.perfil = d.exists() ? d.data() : null; } catch { }
  if (!S.perfil || S.perfil.activo !== true) { await signOut(auth); return login(S.perfil ? 'Su usuario está desactivado.' : 'Su usuario no está habilitado en ADG Movimientos.'); }
  S.unsub.push(onSnapshot(collection(db, 'usuarios', u.uid, 'movs'), s => { S.movs = s.docs.map(d => ({ id: d.id, ...d.data() })); pintar(); }));
  if (esAdmin()) S.unsub.push(onSnapshot(collection(db, 'usuarios'), s => { S.usuarios = s.docs.map(d => ({ id: d.id, ...d.data() })); pintar(); }));
  marco();
});

/* ---------------- interfaz ---------------- */
function marco() {
  $('#app').innerHTML = `<header class="barra"><div class="tit"><b>ADG Movimientos</b><small>AS SYSTEMS</small></div>
    <div class="quien">${esc(S.perfil.nombre || S.user.email)}<small>${esAdmin() ? 'Administrador' : 'Usuario'}</small></div><button class="btn sec peq" id="salir">Salir</button></header>
    <nav class="tabs"><button data-v="movs">Mis movimientos</button><button data-v="resumen">Resumen</button>${esAdmin() ? '<button data-v="todos">Todos los usuarios</button><button data-v="usuarios">Usuarios</button>' : ''}</nav>
    <main id="vista"></main>`;
  $('#salir').onclick = () => signOut(auth);
  $('.tabs').onclick = e => { const v = e.target.dataset.v; if (v) { S.vista = v; pintar(); } };
  pintar();
}
const filtrar = l => l.filter(m => (!S.f.mes || (m.fecha || '').startsWith(S.f.mes)) && (!S.f.tipo || m.tipo === S.f.tipo)
  && (!S.f.q || norm([m.concepto, m.rubro, m.sale, m.corresponde].join(' ')).includes(norm(S.f.q)))).sort((a, b) => (b.fecha || '').localeCompare(a.fecha || ''));
const totales = l => { const t = { Ingreso: 0, Gasto: 0, Traslado: 0 }; l.forEach(m => t[m.tipo] = (t[m.tipo] || 0) + (Number(m.valor) || 0)); t.saldo = t.Ingreso - t.Gasto; return t; };
function filtrosHtml() {
  return `<div class="filtros"><input type="month" id="fmes" value="${S.f.mes}"><select id="ftipo"><option value="">Todos los tipos</option>${TIPOS.map(t => `<option ${S.f.tipo === t ? 'selected' : ''}>${t}</option>`).join('')}</select>
  <input type="search" id="fq" placeholder="Concepto, rubro, cuenta…" value="${esc(S.f.q)}"></div>`;
}
function tarjetasTot(t) {
  return `<div class="tots"><div class="tot ing"><small>Ingresos</small><b>${pesos(t.Ingreso)}</b></div><div class="tot gas"><small>Gastos</small><b>${pesos(t.Gasto)}</b></div>
  <div class="tot tra"><small>Traslados</small><b>${pesos(t.Traslado)}</b></div><div class="tot sal ${t.saldo < 0 ? 'neg' : ''}"><small>Ingresos − gastos</small><b>${pesos(t.saldo)}</b></div></div>`;
}
function tabla(l, editable) {
  if (!l.length) return '<p class="vacio">No hay movimientos con ese filtro.</p>';
  return `<div class="tcont"><table><thead><tr><th>Fecha</th><th>Tipo</th><th>Concepto</th><th>Rubro</th><th>Sale de</th><th>Corresponde a</th><th class="n">Valor</th>${editable ? '<th></th>' : ''}</tr></thead><tbody>
  ${l.map(m => `<tr><td>${ffecha(m.fecha)}</td><td><span class="tag ${m.tipo}">${m.tipo}</span></td><td>${esc(m.concepto)}</td><td>${esc(m.rubro)}</td><td>${esc(m.sale)}</td><td>${esc(m.corresponde)}</td>
  <td class="n">${pesos(m.valor)}</td>${editable ? `<td><button class="lnk" data-ed="${m.id}">Editar</button></td>` : ''}</tr>`).join('')}</tbody></table></div>`;
}
function pintar() {
  const v = $('#vista'); if (!v) return;
  document.querySelectorAll('.tabs button').forEach(b => b.classList.toggle('act', b.dataset.v === S.vista));
  const act = document.activeElement?.id;
  if (S.vista === 'movs') {
    const l = filtrar(S.movs);
    v.innerHTML = `<div class="cabv"><h2>Mis movimientos</h2><div><button class="btn" id="nuevo">+ Movimiento</button> <button class="btn sec" id="exp">Excel</button></div></div>${filtrosHtml()}${tarjetasTot(totales(l))}${tabla(l, true)}`;
    $('#nuevo').onclick = () => form(); $('#exp').onclick = () => exportar(l, 'Mis_movimientos');
    v.onclick = e => { const id = e.target.dataset.ed; if (id) form(S.movs.find(m => m.id === id)); };
  } else if (S.vista === 'resumen') {
    const l = filtrar(S.movs); const por = {};
    l.forEach(m => { const k = m.rubro || 'Sin rubro'; por[k] = por[k] || { Ingreso: 0, Gasto: 0, Traslado: 0 }; por[k][m.tipo] += Number(m.valor) || 0; });
    v.innerHTML = `<div class="cabv"><h2>Resumen por rubro</h2></div>${filtrosHtml()}${tarjetasTot(totales(l))}
    <div class="tcont"><table><thead><tr><th>Rubro</th><th class="n">Ingresos</th><th class="n">Gastos</th><th class="n">Traslados</th></tr></thead><tbody>
    ${Object.entries(por).sort((a, b) => b[1].Gasto - a[1].Gasto).map(([k, t]) => `<tr><td>${esc(k)}</td><td class="n">${pesos(t.Ingreso)}</td><td class="n">${pesos(t.Gasto)}</td><td class="n">${pesos(t.Traslado)}</td></tr>`).join('') || '<tr><td colspan="4" class="vacio">Sin datos</td></tr>'}</tbody></table></div>`;
  } else if (S.vista === 'todos' && esAdmin()) {
    const us = S.usuarios.slice().sort((a, b) => (a.nombre || '').localeCompare(b.nombre || ''));
    const l = filtrar(S.verMovs);
    v.innerHTML = `<div class="cabv"><h2>Movimientos de</h2><div><select id="ver"><option value="">— Seleccione usuario —</option>${us.map(u => `<option value="${u.id}" ${S.ver === u.id ? 'selected' : ''}>${esc(u.nombre || u.email)}</option>`).join('')}</select>
    ${S.ver ? '<button class="btn sec" id="exp">Excel</button>' : ''}</div></div>${S.ver ? filtrosHtml() + tarjetasTot(totales(l)) + tabla(l, false) : '<p class="vacio">Elija un usuario para ver sus movimientos (solo lectura).</p>'}`;
    $('#ver').onchange = async e => { S.ver = e.target.value; S.verMovs = []; if (S.ver) { try { const s = await getDocs(collection(db, 'usuarios', S.ver, 'movs')); S.verMovs = s.docs.map(d => ({ id: d.id, ...d.data() })); } catch (er) { toast(errTxt(er), 1); } } pintar(); };
    const x = $('#exp'); if (x) x.onclick = () => exportar(l, 'Movimientos_' + (S.usuarios.find(u => u.id === S.ver)?.nombre || 'usuario'));
  } else if (S.vista === 'usuarios' && esAdmin()) {
    v.innerHTML = `<div class="cabv"><h2>Usuarios</h2><button class="btn" id="nu">+ Usuario</button></div>
    <div class="tcont"><table><thead><tr><th>Nombre</th><th>Correo</th><th>Rol</th><th>Alta</th><th>Estado</th><th></th></tr></thead><tbody>
    ${S.usuarios.map(u => `<tr><td>${esc(u.nombre)}</td><td>${esc(u.email)}</td><td>${u.rol === 'admin' ? 'Administrador' : 'Usuario'}</td><td>${ffecha(u.alta)}</td><td>${u.activo ? '<span class="tag Ingreso">Activo</span>' : '<span class="tag Gasto">Inactivo</span>'}</td>
    <td>${u.id !== S.user.uid ? `<button class="lnk" data-act="${u.id}">${u.activo ? 'Desactivar' : 'Activar'}</button>` : ''}</td></tr>`).join('')}</tbody></table></div>`;
    $('#nu').onclick = nuevoUsuario;
    v.onclick = async e => { const id = e.target.dataset.act; if (!id) return; const u = S.usuarios.find(x => x.id === id);
      try { await updateDoc(doc(db, 'usuarios', id), { activo: !u.activo }); toast(u.activo ? 'Usuario desactivado' : 'Usuario activado'); } catch (er) { toast(errTxt(er), 1); } };
  }
  ['fmes', 'ftipo', 'fq'].forEach(id => { const el = $('#' + id); if (!el) return; el.oninput = el.onchange = () => { S.f.mes = $('#fmes')?.value || ''; S.f.tipo = $('#ftipo')?.value || ''; S.f.q = $('#fq')?.value || ''; pintar(); }; });
  if (act && $('#' + act)) { const el = $('#' + act); el.focus(); if (el.type === 'search') el.setSelectionRange(el.value.length, el.value.length); }
}
function lista(campo, base = []) { return [...new Set([...base, ...S.movs.map(m => m[campo]).filter(Boolean)])].sort((a, b) => a.localeCompare(b)); }
function form(m) {
  const md = modal(m ? 'Editar movimiento' : 'Nuevo movimiento', `<form id="fm">
    <div class="g2"><label>Fecha<input type="date" name="fecha" required value="${m?.fecha || hoy()}"></label>
    <label>Tipo<select name="tipo">${TIPOS.map(t => `<option ${m?.tipo === t ? 'selected' : ''}>${t}</option>`).join('')}</select></label></div>
    <label>Valor ($)<input type="number" name="valor" min="1" step="1" required inputmode="numeric" value="${m?.valor ?? ''}"></label>
    <label>Concepto<input name="concepto" required maxlength="300" value="${esc(m?.concepto || '')}"></label>
    <label>Rubro<input name="rubro" list="dlr" required maxlength="80" value="${esc(m?.rubro || '')}"></label>
    <div class="g2"><label>Sale de (cuenta / caja / persona)<input name="sale" list="dls" required maxlength="80" value="${esc(m?.sale || '')}"></label>
    <label>Corresponde a<input name="corresponde" list="dlc" required maxlength="80" value="${esc(m?.corresponde || '')}"></label></div>
    <datalist id="dlr">${lista('rubro', RUBROS_BASE).map(x => `<option>${esc(x)}</option>`).join('')}</datalist>
    <datalist id="dls">${lista('sale').map(x => `<option>${esc(x)}</option>`).join('')}</datalist>
    <datalist id="dlc">${lista('corresponde', ['ADG']).map(x => `<option>${esc(x)}</option>`).join('')}</datalist></form>`,
    `${m ? '<button class="btn rojo izq" id="borrar">Eliminar</button>' : ''}<button class="btn sec" data-c>Cancelar</button><button class="btn" id="guardar">Guardar</button>`);
  $('#guardar', md).onclick = async () => {
    const f = new FormData($('#fm', md)); const d = {
      fecha: f.get('fecha'), tipo: f.get('tipo'), valor: Math.round(Number(f.get('valor'))), concepto: f.get('concepto').trim(),
      rubro: f.get('rubro').trim(), sale: f.get('sale').trim(), corresponde: f.get('corresponde').trim(), actualizado: serverTimestamp() };
    if (!d.fecha || !(d.valor > 0) || !d.concepto || !d.rubro || !d.sale || !d.corresponde) return toast('Complete todos los campos; el valor debe ser mayor que cero.', 1);
    try { await setDoc(doc(db, 'usuarios', S.user.uid, 'movs', m?.id || uuid()), d); cerrar(); toast('Movimiento guardado'); } catch (er) { toast(errTxt(er), 1); }
  };
  const b = $('#borrar', md); if (b) b.onclick = async () => {
    if (b.dataset.ok !== '1') { b.dataset.ok = '1'; b.textContent = '¿Seguro? Toque de nuevo'; return; }
    try { await deleteDoc(doc(db, 'usuarios', S.user.uid, 'movs', m.id)); cerrar(); toast('Movimiento eliminado'); } catch (er) { toast(errTxt(er), 1); }
  };
}
function nuevoUsuario() {
  const md = modal('Nuevo usuario', `<form id="fu"><label>Nombre<input name="n" required></label><label>Correo<input type="email" name="e" required></label>
    <label>Contraseña inicial (mín. 6)<input name="p" minlength="6" required></label><label>Rol<select name="r"><option value="usuario">Usuario (ve solo lo suyo)</option><option value="admin">Administrador</option></select></label></form>`,
    `<button class="btn sec" data-c>Cancelar</button><button class="btn" id="cu">Crear</button>`);
  $('#cu', md).onclick = async () => {
    const f = new FormData($('#fu', md)); const sec = initializeApp(cfg, 'alta-' + Date.now());
    try {
      const c = await createUserWithEmailAndPassword(getAuth(sec), f.get('e').trim(), f.get('p'));
      await setDoc(doc(db, 'usuarios', c.user.uid), { nombre: f.get('n').trim(), email: f.get('e').trim(), rol: f.get('r'), activo: true, alta: hoy() });
      await signOut(getAuth(sec)); cerrar(); toast('Usuario creado');
    } catch (er) { toast(errTxt(er), 1); } finally { deleteApp(sec); }
  };
}
function exportar(l, nombre) {
  if (window.Capacitor?.isNativePlatform?.()) return toast('La descarga de Excel se hace desde el PC o el navegador.', 1);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(l.map(m => ({ Fecha: m.fecha, Tipo: m.tipo, Concepto: m.concepto, Rubro: m.rubro, 'Sale de': m.sale, 'Corresponde a': m.corresponde, Valor: Number(m.valor) || 0 }))), 'Movimientos');
  XLSX.writeFile(wb, `${nombre.replace(/\s+/g, '_')}_${hoy()}.xlsx`);
}
login();
