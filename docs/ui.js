/* ============================================================
   Pantalla — versión web. Llama al motor de Python directamente
   (ver app.js); no hay servidor de por medio.
   ============================================================ */

let archivoHoras = null;      // File elegido por la persona
let rutaHoras = null;         // dónde quedó dentro de Pyodide
let fechaTocada = false;
let generadas = [];

function iniciarUI() {
  /* ---- pestañas ---- */
  $$(".pest").forEach((b) => {
    b.addEventListener("click", () => {
      $$(".pest").forEach((x) => x.classList.remove("activa"));
      b.classList.add("activa");
      const v = b.dataset.vista;
      $("#vista-planilla").classList.toggle("oculto", v !== "planilla");
      $("#vista-ccss").classList.toggle("oculto", v !== "ccss");
      if (v === "ccss") cargarGeneradas();
    });
  });

  /* ---- Planilla Madre ---- */
  const zm = $("#zona-madre"), em = $("#entrada-madre");
  zm.addEventListener("click", () => em.click());
  ["dragenter", "dragover"].forEach((e) =>
    zm.addEventListener(e, (ev) => { ev.preventDefault(); zm.classList.add("encima"); }));
  ["dragleave", "drop"].forEach((e) =>
    zm.addEventListener(e, (ev) => { ev.preventDefault(); zm.classList.remove("encima"); }));
  zm.addEventListener("drop", (ev) => { if (ev.dataTransfer.files.length) tomarMadre(ev.dataTransfer.files[0]); });
  em.addEventListener("change", () => { if (em.files.length) tomarMadre(em.files[0]); });
  $("#cambiar-madre").addEventListener("click", () => {
    $("#madre-guardada").classList.add("oculto");
    $("#madre-falta").classList.remove("oculto");
  });

  /* ---- archivo de horas ---- */
  const zona = $("#zona"), entrada = $("#entrada");
  zona.addEventListener("click", () => entrada.click());
  ["dragenter", "dragover"].forEach((e) =>
    zona.addEventListener(e, (ev) => { ev.preventDefault(); zona.classList.add("encima"); }));
  ["dragleave", "drop"].forEach((e) =>
    zona.addEventListener(e, (ev) => { ev.preventDefault(); zona.classList.remove("encima"); }));
  zona.addEventListener("drop", (ev) => { if (ev.dataTransfer.files.length) tomarHoras(ev.dataTransfer.files[0]); });
  entrada.addEventListener("change", () => { if (entrada.files.length) tomarHoras(entrada.files[0]); });
  $("#cambiar").addEventListener("click", () => {
    archivoHoras = null; rutaHoras = null; estado("");
    $("#archivo").classList.add("oculto");
    zona.classList.remove("oculto");
    $("#detectado").classList.add("oculto");
    $("#generar").disabled = true;
  });

  $("#pago").addEventListener("input", () => { fechaTocada = true; });
  $("#pago").addEventListener("change", inspeccionar);
  $("#generar").addEventListener("click", generar);
  $("#unir").addEventListener("click", unificar);

  /* ---- respaldo ---- */
  $("#btn-exportar").addEventListener("click", (e) => { e.preventDefault(); exportarRespaldo(); });
  $("#btn-importar").addEventListener("click", (e) => { e.preventDefault(); $("#entrada-respaldo").click(); });
  const salir = $("#btn-salir");
  if (salir) salir.addEventListener("click", (e) => { e.preventDefault(); cerrarSesion(); });
  $("#entrada-respaldo").addEventListener("change", (e) => {
    if (e.target.files.length) importarRespaldo(e.target.files[0]);
  });
}

function estado(t) {
  const e = $("#estado");
  if (!t) { e.classList.add("oculto"); return; }
  e.classList.remove("oculto"); e.textContent = t;
}

function fallo(msg) {
  const c = $("#detectado");
  c.classList.remove("oculto");
  c.innerHTML = '<div class="alerta mal"><strong>No se pudo revisar el archivo.</strong><br>' +
    msg + '<div style="margin-top:10px"><button class="secundario" id="reintentar">Reintentar</button></div></div>';
  const b = document.getElementById("reintentar");
  if (b) b.addEventListener("click", inspeccionar);
  $("#generar").disabled = !rutaHoras;
}

async function tomarHoras(f) {
  if (!/\.xls[xm]$/i.test(f.name)) { alert("Debe ser un archivo de Excel (.xlsx)"); return; }
  archivoHoras = f;
  $("#archivo-nombre").textContent = f.name;
  $("#archivo-meta").textContent = Math.round(f.size / 1024).toLocaleString("es-CR") + " KB";
  $("#archivo").classList.remove("oculto");
  $("#zona").classList.add("oculto");
  estado("Leyendo el archivo…");
  try {
    const b = new Uint8Array(await f.arrayBuffer());
    rutaHoras = RAIZ + "/horas_del_periodo.xlsx";
    py.FS.writeFile(rutaHoras, b);
    estado("");
    inspeccionar();
  } catch (e) {
    estado(""); fallo(String(e.message || e));
  }
}

function inspeccionar() {
  if (!rutaHoras || !madreNombre) return;
  estado("Revisando el período y la Planilla Madre…");
  setTimeout(() => {
    try {
      const d = inspeccionarPy(rutaHoras, $("#pago").value || null);
      estado("");
      const c = $("#detectado");
      c.classList.remove("oculto");
      if (!d.ok) { c.innerHTML = '<div class="alerta mal">' + d.error + "</div>"; $("#generar").disabled = true; return; }
      let h = '<div class="detectado">' +
        '<div class="dato"><div class="et">Período detectado</div><div class="vl">' + d.inicio + " al " + d.fin + "</div></div>" +
        '<div class="dato"><div class="et">Duración</div><div class="vl">' + d.dias + " días</div></div>" +
        '<div class="dato"><div class="et">Ventana de renta</div><div class="vl">' + (d.ventana === 1 ? "Abre" : "Cierra") + "</div></div></div>" +
        '<div class="alerta bien" style="background:#f7f9fc;border-color:var(--linea);color:var(--texto)">' +
        (d.primera
          ? "<strong>Primera corrida.</strong> No hay planillas generadas todavía, así que se arranca desde la plantilla."
          : "<strong>Continúa desde:</strong> " + d.estado) +
        (d.salida ? '<br><span style="color:var(--suave);font-size:13px">Va a generar: <strong>' + d.salida + "</strong>" +
          (d.ya_existe ? " — ya existe uno con ese nombre y se va a reemplazar" : "") + "</span>" : "") + "</div>";
      if (!d.ciclo_ok) h += '<div class="alerta mal"><strong>El período no tiene 14 días.</strong> El ciclo es bisemanal: revise la fila 2 de la hoja «Horas de trabajo».</div>';
      if (d.continuidad) h += '<div class="alerta mal"><strong>Atención con la continuidad.</strong> ' + d.continuidad + "</div>";
      else if (d.periodo_anterior_fin) h += '<div class="alerta bien">El período anterior cerró el ' + d.periodo_anterior_fin + " y este arranca justo al día siguiente. El ciclo va bien.</div>";
      h += '<div class="alerta ojo">' + (d.ventana === 1
        ? "<strong>Este pago ABRE la ventana de renta.</strong> La retención queda en cero y se cobrará completa en el pago siguiente."
        : "<strong>Este pago CIERRA la ventana de renta.</strong> Aquí se retiene la renta de los dos pagos.") + "</div>";
      c.innerHTML = h;
      if (d.pago_sugerido && !fechaTocada) $("#pago").value = d.pago_sugerido;
      $("#generar").disabled = false;
    } catch (e) { estado(""); fallo(String(e.message || e)); }
  }, 30);
}

async function generar() {
  if (!$("#pago").value) { alert("Indique la fecha de pago"); return; }
  $("#cargando").classList.remove("oculto");
  $("#salida").innerHTML = "";
  $("#generar").disabled = true;
  await new Promise((r) => setTimeout(r, 50));
  try {
    const d = generarPy(rutaHoras, $("#pago").value,
      $("#blancos").checked, $("#ingresos").checked);
    await persistirGeneradas();
    $("#cargando").classList.add("oculto");
    $("#generar").disabled = false;
    pintar(d);
    $("#salida").scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (e) {
    $("#cargando").classList.add("oculto");
    $("#generar").disabled = false;
    $("#salida").innerHTML = '<div class="tarjeta"><div class="alerta mal"><strong>No se pudo generar.</strong><br>' +
      (e.message || e) + "</div></div>";
  }
}

const DESTACAR = {
  "Incluidos en la planilla": 0, "Sin carne (quedaron fuera)": 1,
  "Avisos del ciclo bisemanal": 1, "Diferencias de salario": 1,
  "Enlaces rotos en el archivo de horas": 1,
  "Diferencias de bruto vs archivo de horas": 1,
};

function pintar(d) {
  let h = '<div class="tarjeta">' +
    '<div class="alerta bien" style="margin-top:0"><strong>Listo.</strong> Se generó <strong>' + d.planilla +
    "</strong> y quedó guardada en este navegador. La Planilla Madre no se tocó.</div>" +
    '<div class="descargas" style="margin:16px 0 4px">' +
    '<button class="descargar" data-baja="' + d.planilla + '">Descargar la planilla</button>' +
    '<button class="descargar alt" data-baja="' + d.reporte + '">Descargar el reporte en Excel</button></div>' +
    '<div style="color:var(--suave);font-size:13px;margin-top:12px">Al abrirla en Excel se recalcula sola. ' +
    "La hoja <strong>Verificación</strong> le confirma que los números dieron correctos.</div></div>";

  h += '<div class="tarjeta"><h2 style="margin-top:0">Resumen del período</h2><div class="cifras">';
  d.resumen.forEach((p) => {
    const k = p[0], v = p[1];
    if (!(k in DESTACAR)) return;
    h += '<div class="cifra' + (DESTACAR[k] && v > 0 ? " ojo" : "") + '"><div class="et">' + k +
      '</div><div class="vl">' + fmt(v) + "</div></div>";
  });
  h += "</div><table style='margin-top:20px;min-width:0'><tbody>";
  d.resumen.forEach((p) => {
    const k = p[0], v = p[1];
    if (k === "") { h += '<tr><td colspan="2" style="padding:6px 0;border:0"></td></tr>'; return; }
    if (v === "") { h += '<tr><th colspan="2" style="background:transparent;padding:10px 0 4px;color:var(--azul);border-bottom:1px solid var(--linea)">' + k + "</th></tr>"; return; }
    h += '<tr><td style="border:0">' + k + '</td><td class="n" style="border:0;font-weight:600">' + fmt(v) + "</td></tr>";
  });
  h += "</tbody></table></div><h2>Revisión</h2>";

  d.sesiones.forEach((s, i) => {
    const ok = s.cantidad === 0;
    h += '<div class="grupo' + (!ok && i < 4 ? " abierto" : "") + '"><div class="cabeza">' +
      '<span class="flecha">&#9654;</span><h3>' + s.titulo + "</h3>" +
      '<span class="pastilla' + (ok ? "" : " ojo") + '">' + (ok ? "sin novedad" : s.cantidad + (s.cantidad === 1 ? " caso" : " casos")) +
      '</span></div><div class="cuerpo"><div class="nota">' + s.nota + "</div>";
    if (ok) h += '<div style="color:var(--verde)">Nada que revisar aquí.</div>';
    else {
      h += '<div class="tabla-env"><table><thead><tr>' + s.encabezados.map((e) => "<th>" + e + "</th>").join("") + "</tr></thead><tbody>";
      s.filas.slice(0, 400).forEach((f) => {
        h += "<tr>" + f.map((v) => "<td" + (esNum(v) ? ' class="n"' : "") + ">" + fmt(v) + "</td>").join("") + "</tr>";
      });
      h += "</tbody></table></div>";
      if (s.filas.length > 400) h += '<div class="nota">Se muestran las primeras 400 de ' + s.filas.length + ".</div>";
    }
    h += "</div></div>";
  });
  $("#salida").innerHTML = h;
  engancharGrupos("#salida");
}

function engancharGrupos(sel) {
  $$(sel + " .cabeza").forEach((c) =>
    c.addEventListener("click", () => c.parentNode.classList.toggle("abierto")));
  $$(sel + " [data-baja]").forEach((b) =>
    b.addEventListener("click", () => descargar(DIR_GENERADAS + "/" + b.dataset.baja, b.dataset.baja)));
}

/* ---------------- unificación CCSS ---------------- */
function cargarGeneradas() {
  const cont = $("#lista-generadas");
  try {
    generadas = listarGeneradasPy();
  } catch (e) {
    cont.innerHTML = '<div class="alerta mal">No se pudo leer la lista: ' + (e.message || e) + "</div>";
    return;
  }
  if (!generadas.length) {
    cont.innerHTML = '<div class="alerta ojo">Todavía no hay planillas generadas. Genere primero las del mes en la pestaña «Generar planilla del período».</div>';
    $("#unir").disabled = true; return;
  }
  cont.innerHTML = generadas.map((a) =>
    '<label class="fila-arch"><input type="checkbox" class="ck-gen" value="' + a.nombre + '" data-mes="' + a.mes + '">' +
    '<span style="flex:1"><span class="nm">Pago del ' + a.pago + "</span>" +
    '<span class="dt"> &nbsp;·&nbsp; período ' + a.periodo + "</span>" +
    '<br><span class="dt">' + a.nombre + "</span></span></label>").join("");
  const mes = generadas[0].mes;
  $$(".ck-gen").forEach((c) => {
    c.addEventListener("change", revisarSeleccion);
    if (c.dataset.mes === mes) c.checked = true;
  });
  revisarSeleccion();
}

function seleccionadas() {
  return $$(".ck-gen").filter((c) => c.checked).map((c) => c.value);
}

function revisarSeleccion() {
  const sel = seleccionadas();
  $("#unir").disabled = sel.length === 0;
  if (sel.length && !$("#etiqueta").value) {
    const m = generadas.filter((a) => sel.indexOf(a.nombre) >= 0);
    if (m.length) {
      const MES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
        "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre"];
      const p = m[m.length - 1].mes.split("/");
      $("#etiqueta").value = MES[parseInt(p[0], 10)] + " " + p[1];
    }
  }
}

async function unificar() {
  const sel = seleccionadas();
  if (!sel.length) return;
  $("#cargando-ccss").classList.remove("oculto");
  $("#salida-ccss").innerHTML = "";
  $("#unir").disabled = true;
  await new Promise((r) => setTimeout(r, 50));
  try {
    const d = unificarPy(sel, $("#etiqueta").value);
    await persistirGeneradas();
    $("#cargando-ccss").classList.add("oculto");
    $("#unir").disabled = false;
    pintarCCSS(d);
    $("#salida-ccss").scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (e) {
    $("#cargando-ccss").classList.add("oculto");
    $("#unir").disabled = false;
    $("#salida-ccss").innerHTML = '<div class="tarjeta"><div class="alerta mal"><strong>No se pudo unificar.</strong><br>' +
      (e.message || e) + "</div></div>";
  }
}

function pintarCCSS(d) {
  const ORD = ["SAL", "IC", "EX", "INS"];
  const NOM = { SAL: "Salario", IC: "Inclusión", EX: "Exclusión", INS: "Incapacidad" };
  let h = '<div class="tarjeta">' +
    '<div class="alerta bien" style="margin-top:0"><strong>Listo.</strong> Se generó <strong>' + d.archivo + "</strong>.</div>" +
    '<div class="descargas" style="margin:16px 0 4px">' +
    '<button class="descargar" data-baja="' + d.archivo + '">Descargar el reporte</button></div></div>';
  h += '<div class="tarjeta"><h2 style="margin-top:0">Resumen</h2><div class="cifras">' +
    '<div class="cifra"><div class="et">Personas</div><div class="vl">' + d.personas + "</div></div>" +
    '<div class="cifra"><div class="et">Días laborados</div><div class="vl">' + fmt(d.dias) + "</div></div>" +
    '<div class="cifra"><div class="et">Total bruto</div><div class="vl">' + fmt(d.bruto) + "</div></div>";
  ORD.forEach((e) => {
    if (d.estados[e]) h += '<div class="cifra"><div class="et">' + e + " — " + NOM[e] + '</div><div class="vl">' + d.estados[e] + "</div></div>";
  });
  h += '</div><div style="margin-top:16px;color:var(--suave);font-size:13.5px">Planillas unidas: ' +
    d.periodos.map((p) => "pago del " + p.pago).join(" &nbsp;·&nbsp; ") + "</div></div>";

  h += "<h2>Revisión antes de enviar</h2>";
  d.secciones.forEach((s, i) => {
    const ok = s.cantidad === 0;
    h += '<div class="grupo' + (!ok && i < 3 ? " abierto" : "") + '"><div class="cabeza">' +
      '<span class="flecha">&#9654;</span><h3>' + s.titulo + "</h3>" +
      '<span class="pastilla' + (ok ? "" : " ojo") + '">' + (ok ? "sin novedad" : s.cantidad + (s.cantidad === 1 ? " caso" : " casos")) + "</span></div><div class=\"cuerpo\">";
    if (ok) h += '<div style="color:var(--verde);padding-top:10px">Nada que revisar aquí.</div>';
    else {
      h += '<div class="tabla-env" style="margin-top:12px"><table><thead><tr>' +
        s.encabezados.map((e) => "<th>" + e + "</th>").join("") + "</tr></thead><tbody>";
      s.filas.slice(0, 300).forEach((f) => {
        h += "<tr>" + f.map((v) => "<td" + (esNum(v) ? ' class="n"' : "") + ">" + fmt(v) + "</td>").join("") + "</tr>";
      });
      h += "</tbody></table></div>";
    }
    h += "</div></div>";
  });

  h += '<h2>Union Final</h2><div class="tarjeta"><div class="tabla-env"><table><thead><tr>' +
    ["Carné", "Días", "Puesto", "Nombre", "Identificación", "Bruto", "Estado", "Fecha IC", "Fecha EX"]
      .map((x) => "<th>" + x + "</th>").join("") + "</tr></thead><tbody>";
  d.muestra.slice(0, 300).forEach((f) => {
    h += "<tr>" + f.map((v, i) =>
      i === 6 ? '<td><span class="chip ' + String(v).toLowerCase() + '">' + v + "</span></td>"
        : "<td" + (esNum(v) ? ' class="n"' : "") + ">" + fmt(v) + "</td>").join("") + "</tr>";
  });
  h += "</tbody></table></div>";
  if (d.muestra.length > 300) h += '<div class="nota">Se muestran las primeras 300 de ' + d.muestra.length + ". El archivo trae todas.</div>";
  h += "</div>";
  $("#salida-ccss").innerHTML = h;
  engancharGrupos("#salida-ccss");
}
