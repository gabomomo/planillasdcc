/* ============================================================
   Pantalla de acceso

   ADVERTENCIA HONESTA
   -------------------
   Esto es un sitio estatico: no hay servidor que valide nada. La
   comprobacion ocurre en el navegador de quien entra, asi que
   cualquiera con conocimientos tecnicos la puede saltar.

   Sirve para que nadie entre por casualidad. NO es una cerradura.

   Lo unico que protege de verdad es que este sitio no contiene
   datos: los archivos de planilla viven en el navegador de cada
   persona (IndexedDB, que es por origen y por computadora), no aqui.

   La contrasena no esta escrita en ningun lado: solo queda su
   derivacion PBKDF2-SHA256 con 310.000 iteraciones y sal aleatoria,
   que no permite recuperarla.
   ============================================================ */

const ACCESO = {
  correo: "kmoran@dcccr.com",
  salt: "z2cVWsXE/L90dHCIQld9Pg==",
  hash: "MhOrhtTbxJluz8JHW4XGvXIXQWIYqPQ7G4WJjkf0hRk=",
  iter: 310000,
  clave: "planilla-acceso",
  horas: 12,           // cuanto dura la sesion antes de volver a pedir la clave
};

function b64aBytes(s) {
  const bin = atob(s);
  const u = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
  return u;
}

async function derivar(pass) {
  const base = await crypto.subtle.importKey(
    "raw", new TextEncoder().encode(pass), "PBKDF2", false, ["deriveBits"]);
  const bits = await crypto.subtle.deriveBits(
    { name: "PBKDF2", salt: b64aBytes(ACCESO.salt), iterations: ACCESO.iter, hash: "SHA-256" },
    base, 256);
  const u = new Uint8Array(bits);
  let s = "";
  for (let i = 0; i < u.length; i++) s += String.fromCharCode(u[i]);
  return btoa(s);
}

function sesionValida() {
  try {
    const g = JSON.parse(localStorage.getItem(ACCESO.clave) || "null");
    return !!(g && g.hasta && Date.now() < g.hasta);
  } catch (e) { return false; }
}

function abrirSesion() {
  localStorage.setItem(ACCESO.clave,
    JSON.stringify({ hasta: Date.now() + ACCESO.horas * 3600 * 1000 }));
}

function cerrarSesion() {
  localStorage.removeItem(ACCESO.clave);
  location.reload();
}

function pintarLogin() {
  const d = document.createElement("div");
  d.id = "login";
  d.style.cssText = "position:fixed;inset:0;background:var(--fondo);z-index:300;" +
    "display:flex;align-items:center;justify-content:center;padding:20px";
  d.innerHTML =
    '<div class="tarjeta" style="max-width:420px;width:100%;margin:0">' +
    '<h1 style="font-size:21px;margin:0 0 4px">Generador de Planilla Bisemanal</h1>' +
    '<div class="sub" style="margin-bottom:22px">Ingeniería Estrella S.A.</div>' +
    '<label class="campo" for="lg-correo">Correo</label>' +
    '<input type="email" id="lg-correo" autocomplete="username" ' +
    'style="width:100%;padding:10px 12px;border:1px solid var(--linea);border-radius:8px;font-size:15px;font-family:inherit">' +
    '<label class="campo" for="lg-clave" style="margin-top:14px">Contraseña</label>' +
    '<input type="password" id="lg-clave" autocomplete="current-password" ' +
    'style="width:100%;padding:10px 12px;border:1px solid var(--linea);border-radius:8px;font-size:15px;font-family:inherit">' +
    '<div id="lg-error" class="alerta mal oculto" style="margin-top:14px"></div>' +
    '<button class="principal" id="lg-entrar" style="margin-top:18px">Entrar</button>' +
    '<div style="color:var(--suave);font-size:12.5px;margin-top:16px;line-height:1.5">' +
    'Sus archivos de planilla no se guardan en este sitio: quedan en este ' +
    'navegador y no salen de esta computadora.</div></div>';
  document.body.appendChild(d);

  const err = (m) => {
    const e = document.getElementById("lg-error");
    e.classList.remove("oculto"); e.innerHTML = m;
  };

  const entrar = async () => {
    const c = document.getElementById("lg-correo").value.trim().toLowerCase();
    const p = document.getElementById("lg-clave").value;
    const btn = document.getElementById("lg-entrar");
    document.getElementById("lg-error").classList.add("oculto");
    if (!c || !p) { err("Escriba el correo y la contraseña."); return; }
    btn.disabled = true; btn.textContent = "Verificando…";
    try {
      const h = await derivar(p);
      if (c === ACCESO.correo && h === ACCESO.hash) {
        abrirSesion();
        d.remove();
        arrancar();
      } else {
        err("<strong>Correo o contraseña incorrectos.</strong>");
        btn.disabled = false; btn.textContent = "Entrar";
      }
    } catch (e) {
      err("No se pudo verificar: " + (e.message || e));
      btn.disabled = false; btn.textContent = "Entrar";
    }
  };

  document.getElementById("lg-entrar").addEventListener("click", entrar);
  ["lg-correo", "lg-clave"].forEach((id) =>
    document.getElementById(id).addEventListener("keydown", (e) => {
      if (e.key === "Enter") entrar();
    }));
  setTimeout(() => document.getElementById("lg-correo").focus(), 100);
}

/* Punto de entrada: decide si pide la clave o arranca directo. */
function iniciar() {
  if (sesionValida()) { arrancar(); return; }
  document.getElementById("arranque").style.display = "none";
  pintarLogin();
}
