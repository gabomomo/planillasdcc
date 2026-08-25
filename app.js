/* ============================================================
   Generador de Planilla Bisemanal — versión web
   Ingeniería Estrella S.A.

   Todo corre dentro del navegador:
     - Pyodide (Python compilado a WebAssembly) ejecuta planilla.py
     - Los archivos viven en el sistema de archivos virtual de Pyodide
     - Se respaldan en IndexedDB para que sobrevivan al cerrar la pestaña

   Ningún archivo se sube a ningún servidor.
   ============================================================ */

const PYODIDE_URL = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";
const RAIZ = "/trabajo";
const DIR_PLANTILLAS = RAIZ + "/Plantillas";
const DIR_GENERADAS = RAIZ + "/Planillas generadas";

let py = null;
let P = null;                 // el módulo planilla.py
let madreNombre = null;

const $ = (s) => document.querySelector(s);
const $$ = (s) => Array.prototype.slice.call(document.querySelectorAll(s));

function fmt(v) {
  if (typeof v === "number") return v.toLocaleString("es-CR", { maximumFractionDigits: 2 });
  return v === null || v === undefined ? "" : String(v);
}
const esNum = (v) => typeof v === "number";

/* ---------------------------------------------------------------
   Guardado en el navegador (IndexedDB)
   --------------------------------------------------------------- */
const DB_NOMBRE = "planilla-bisemanal";
let db = null;

function abrirDB() {
  return new Promise((ok, mal) => {
    const r = indexedDB.open(DB_NOMBRE, 1);
    r.onupgradeneeded = () => {
      const d = r.result;
      if (!d.objectStoreNames.contains("archivos")) d.createObjectStore("archivos");
      if (!d.objectStoreNames.contains("ajustes")) d.createObjectStore("ajustes");
    };
    r.onsuccess = () => { db = r.result; ok(db); };
    r.onerror = () => mal(r.error);
  });
}

function idbGuardar(almacen, clave, valor) {
  return new Promise((ok, mal) => {
    const t = db.transaction(almacen, "readwrite");
    t.objectStore(almacen).put(valor, clave);
    t.oncomplete = ok;
    t.onerror = () => mal(t.error);
  });
}
function idbBorrar(almacen, clave) {
  return new Promise((ok, mal) => {
    const t = db.transaction(almacen, "readwrite");
    t.objectStore(almacen).delete(clave);
    t.oncomplete = ok;
    t.onerror = () => mal(t.error);
  });
}
function idbTodo(almacen) {
  return new Promise((ok, mal) => {
    const t = db.transaction(almacen, "readonly");
    const s = t.objectStore(almacen);
    const claves = s.getAllKeys(), vals = s.getAll();
    t.oncomplete = () => {
      const m = {};
      claves.result.forEach((k, i) => { m[k] = vals.result[i]; });
      ok(m);
    };
    t.onerror = () => mal(t.error);
  });
}

/* ---------------------------------------------------------------
   Puente entre IndexedDB y el sistema de archivos de Pyodide
   --------------------------------------------------------------- */
function asegurarDir(ruta) {
  const partes = ruta.split("/").filter(Boolean);
  let acum = "";
  for (const p of partes) {
    acum += "/" + p;
    try { py.FS.mkdir(acum); } catch (e) { /* ya existe */ }
  }
}

async function restaurarArchivos() {
  asegurarDir(DIR_PLANTILLAS);
  asegurarDir(DIR_GENERADAS);
  const todos = await idbTodo("archivos");
  let n = 0;
  for (const rel in todos) {
    const destino = RAIZ + "/" + rel;
    asegurarDir(destino.substring(0, destino.lastIndexOf("/")));
    py.FS.writeFile(destino, new Uint8Array(todos[rel]));
    n++;
  }
  return n;
}

async function persistir(rel, bytes) {
  await idbGuardar("archivos", rel, bytes.buffer ? bytes.buffer : bytes);
}

/** Copia a IndexedDB todo lo que haya en "Planillas generadas". */
async function persistirGeneradas() {
  let nombres = [];
  try { nombres = py.FS.readdir(DIR_GENERADAS).filter((n) => n.endsWith(".xlsx")); }
  catch (e) { return; }
  for (const n of nombres) {
    const b = py.FS.readFile(DIR_GENERADAS + "/" + n);
    await persistir("Planillas generadas/" + n, b);
  }
}

function descargar(rutaFS, nombre) {
  const datos = py.FS.readFile(rutaFS);
  const blob = new Blob([datos], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = nombre;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}

/* ---------------------------------------------------------------
   Arranque
   --------------------------------------------------------------- */
function paso(t) { $("#arranque-detalle").textContent = t; }

async function arrancar() {
  try {
    await abrirDB();

    paso("Cargando el motor… (unos 10 MB la primera vez)");
    py = await loadPyodide({ indexURL: PYODIDE_URL });

    paso("Instalando el lector de Excel…");
    await py.loadPackage("micropip");
    const mp = py.pyimport("micropip");
    await mp.install("openpyxl");
    // Pillow hace falta para que las imágenes de la Madre (el logo)
    // sobrevivan al volver a guardar el archivo.
    try { await mp.install("pillow"); } catch (e) { console.warn("pillow:", e); }

    paso("Cargando el programa…");
    const codigo = await (await fetch("planilla.py?v=" + Date.now())).text();
    asegurarDir(RAIZ);
    py.FS.writeFile(RAIZ + "/planilla.py", new TextEncoder().encode(codigo));
    py.runPython(`import sys\nsys.path.insert(0, "${RAIZ}")`);
    P = py.pyimport("planilla");

    paso("Recuperando sus archivos…");
    await restaurarArchivos();

    const guardado = (await idbTodo("ajustes")).madre;
    if (guardado) {
      try { py.FS.stat(DIR_PLANTILLAS + "/" + guardado); madreNombre = guardado; }
      catch (e) { madreNombre = null; }
    }

    $("#arranque").style.display = "none";
    pintarMadre();
    iniciarUI();
  } catch (e) {
    console.error(e);
    $("#arranque-detalle").style.display = "none";
    document.querySelector("#arranque .rueda").style.display = "none";
    const c = $("#arranque-error");
    c.classList.remove("oculto");
    c.innerHTML =
      "<strong>No se pudo iniciar el programa.</strong><br>" + (e.message || e) +
      "<br><br><span style='font-size:13px'>Revise que tenga conexión a internet y " +
      "recargue la página. Si el problema sigue, pruebe con otro navegador.</span>";
  }
}

/* ---------------------------------------------------------------
   La Planilla Madre (se sube una vez y queda guardada)
   --------------------------------------------------------------- */
function pintarMadre() {
  if (madreNombre) {
    $("#madre-nombre").textContent = madreNombre;
    $("#madre-guardada").classList.remove("oculto");
    $("#madre-falta").classList.add("oculto");
    const sel = $("#madre");
    sel.innerHTML = "<option>" + madreNombre + "</option>";
    sel.value = madreNombre;
  } else {
    $("#madre-guardada").classList.add("oculto");
    $("#madre-falta").classList.remove("oculto");
  }
}

async function tomarMadre(f) {
  if (!/\.xls[xm]$/i.test(f.name)) { alert("Debe ser un archivo de Excel (.xlsx)"); return; }
  const b = new Uint8Array(await f.arrayBuffer());
  // se borra la anterior para que no queden dos plantillas
  if (madreNombre && madreNombre !== f.name) {
    try { py.FS.unlink(DIR_PLANTILLAS + "/" + madreNombre); } catch (e) {}
    await idbBorrar("archivos", "Plantillas/" + madreNombre);
  }
  py.FS.writeFile(DIR_PLANTILLAS + "/" + f.name, b);
  await persistir("Plantillas/" + f.name, b);
  await idbGuardar("ajustes", "madre", f.name);
  madreNombre = f.name;
  pintarMadre();
  inspeccionar();
}

/* ---------------------------------------------------------------
   Llamadas al motor (Python) — reemplazan al servidor
   --------------------------------------------------------------- */
function jsonPy(codigo) {
  const r = py.runPython("import json\n" + codigo);
  return JSON.parse(r);
}

function inspeccionarPy(rutaHoras, pagoISO) {
  return jsonPy(`
import planilla as P, datetime, os, openpyxl
_h = ${JSON.stringify(rutaHoras)}
_pago = ${pagoISO ? `datetime.date.fromisoformat(${JSON.stringify(pagoISO)})` : "None"}
_madre = ${JSON.stringify(DIR_PLANTILLAS + "/" + (madreNombre || ""))}
_d = P.detectar_periodo(_h)
if not _d:
    _out = {"ok": False, "error": "Ese archivo no tiene la hoja 'Horas de trabajo' con las fechas del periodo en la fila 2. Revise que sea el archivo de horas correcto."}
else:
    _ruta, _nom, _prim = P.elegir_estado(${JSON.stringify(RAIZ)}, _madre, _pago)
    _e = P.leer_estado(_ruta)
    _cont = None
    if _e["fin"]:
        _esp = _e["fin"] + datetime.timedelta(days=1)
        _dif = (_d[0] - _esp).days
        if _dif > 0:
            _cont = "Quedarian %d dias sin pagar entre el periodo anterior (cerro el %s) y este." % (_dif, _e["fin"].strftime("%d/%m/%Y"))
        elif _dif < 0:
            _cont = "Se estarian pagando %d dias dos veces: el periodo anterior cerro el %s." % (abs(_dif), _e["fin"].strftime("%d/%m/%Y"))
    _sug = (_e["pago"] + datetime.timedelta(days=P.DIAS_CICLO)) if _e["pago"] else (_d[-1] + datetime.timedelta(days=12))
    _fp = _pago or _sug
    _sal = "%s%s.xlsx" % (P.PREFIJO_SALIDA, _fp.strftime("%d-%m-%Y"))
    _out = {
        "ok": True, "dias": len(_d),
        "inicio": _d[0].strftime("%d/%m/%Y"), "fin": _d[-1].strftime("%d/%m/%Y"),
        "ciclo_ok": len(_d) == P.DIAS_CICLO,
        "periodo_anterior_fin": _e["fin"].strftime("%d/%m/%Y") if _e["fin"] else None,
        "pago_sugerido": _sug.isoformat(), "continuidad": _cont,
        "ventana": _e["posicion"], "estado": _nom, "primera": _prim,
        "salida": _sal,
        "ya_existe": os.path.isfile(os.path.join(${JSON.stringify(DIR_GENERADAS)}, _sal)),
    }
json.dumps(_out, default=str)
`);
}

function generarPy(rutaHoras, pagoISO, blancos, ingresos, mixtas) {
  return jsonPy(`
import planilla as P, datetime, os
_sal, _rep, _res, _av = P.procesar(
    ${JSON.stringify(rutaHoras)},
    ${JSON.stringify(DIR_PLANTILLAS + "/" + (madreNombre || ""))},
    ${JSON.stringify(RAIZ)},
    datetime.date.fromisoformat(${JSON.stringify(pagoISO)}),
    dias_blanco_como_ausencia=${blancos ? "True" : "False"},
    prorratear_ingreso=${ingresos ? "True" : "False"},
    jornada_mixta=${mixtas ? '"neutro"' : '"madre"'},
    log=lambda *a: None)
_sec = []
for _clave, _tit, _enc, _nota in P.SECCIONES:
    _f = _av.items.get(_clave, [])
    _sec.append({"titulo": __import__("re").sub(r"^\\d+b? ", "", _tit),
                 "encabezados": list(_enc), "nota": _nota, "cantidad": len(_f),
                 "filas": [["" if v is None else v for v in x] for x in _f]})
json.dumps({"ok": True, "planilla": os.path.basename(_sal),
            "reporte": os.path.basename(_rep),
            "resumen": [[k, v] for k, v in _res], "sesiones": _sec}, default=str)
`);
}

function unificarPy(nombres, etiqueta) {
  return jsonPy(`
import planilla as P, os
_rutas = [os.path.join(${JSON.stringify(DIR_GENERADAS)}, n) for n in ${JSON.stringify(nombres)}]
_sal, _filas, _per, _av = P.unificar(_rutas, ${JSON.stringify(RAIZ)},
                                     ${JSON.stringify(etiqueta || "")} or None,
                                     log=lambda *a: None)
_SEC = [
  ("movimiento", "Movimientos detectados", ["Carné","Nombre","Estado","Fecha IC","Fecha EX","De dónde salió"]),
  ("sin_fecha", "Salidas sin fecha", ["Carné","Nombre","Estado","Detalle","Nota"]),
  ("ingreso_dudoso", "Fechas de ingreso a confirmar", ["Carné","Nombre","Fecha usada","De dónde salió","Qué hacer"]),
  ("multiple", "Casos con más de una condición", ["Carné","Nombre","Estado asignado","Detalle","Nota"]),
  ("parcial", "Aparecen en algunos períodos", ["Carné","Nombre","En cuántos","De cuántos","Faltó en"]),
]
_sec = []
for _c, _t, _e in _SEC:
    _f = _av.items.get(_c, [])
    _sec.append({"titulo": _t, "encabezados": _e, "nota": "", "cantidad": len(_f),
                 "filas": [["" if v is None else v for v in x] for x in _f]})
_est = {}
for _x in _filas:
    _est[_x["estado"]] = _est.get(_x["estado"], 0) + 1
json.dumps({"ok": True, "archivo": os.path.basename(_sal), "personas": len(_filas),
  "estados": _est, "bruto": round(sum(x["bruto"] for x in _filas), 2),
  "dias": round(sum(x["dias"] for x in _filas), 2),
  "periodos": [{"archivo": p["archivo"], "pago": p["pago"].strftime("%d/%m/%Y")} for p in _per],
  "secciones": _sec,
  "muestra": [[x["carne"], x["dias"], x["puesto"], x["nombre"], x["cedula"], x["bruto"],
               x["estado"], x["f_ic"].strftime("%d/%m/%Y") if x["f_ic"] else "",
               x["f_ex"].strftime("%d/%m/%Y") if x["f_ex"] else ""] for x in _filas]}, default=str)
`);
}

function listarGeneradasPy() {
  return jsonPy(`
import planilla as P
_g = P.listar_generadas(${JSON.stringify(RAIZ)})
_g.reverse()
json.dumps([{"nombre": x["nombre"], "pago": x["pago"].strftime("%d/%m/%Y"),
             "periodo": "%s al %s" % (x["ini"].strftime("%d/%m/%Y") if x["ini"] else "?",
                                      x["fin"].strftime("%d/%m/%Y") if x["fin"] else "?"),
             "mes": "%02d/%d" % (x["pago"].month, x["pago"].year)} for x in _g], default=str)
`);
}

/* ---------------------------------------------------------------
   Respaldo / restauración
   --------------------------------------------------------------- */
async function exportarRespaldo() {
  const arch = await idbTodo("archivos");
  const aj = await idbTodo("ajustes");
  const paquete = { version: 1, ajustes: aj, archivos: {} };
  for (const k in arch) {
    const u8 = new Uint8Array(arch[k]);
    let s = "";
    for (let i = 0; i < u8.length; i += 8192)
      s += String.fromCharCode.apply(null, u8.subarray(i, i + 8192));
    paquete.archivos[k] = btoa(s);
  }
  const blob = new Blob([JSON.stringify(paquete)], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "Respaldo planilla " + new Date().toISOString().slice(0, 10) + ".json";
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

async function importarRespaldo(f) {
  const p = JSON.parse(await f.text());
  if (!p || !p.archivos) { alert("Ese archivo no es un respaldo válido."); return; }
  for (const k in p.archivos) {
    const bin = atob(p.archivos[k]);
    const u8 = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    await persistir(k, u8);
  }
  for (const k in (p.ajustes || {})) await idbGuardar("ajustes", k, p.ajustes[k]);
  alert("Respaldo restaurado. La página se va a recargar.");
  location.reload();
}
