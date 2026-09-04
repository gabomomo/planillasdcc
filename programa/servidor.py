#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interfaz web LOCAL del generador de planilla bisemanal.

Levanta un servidor en la propia computadora (127.0.0.1) y abre el navegador.
Nada sale a internet: el archivo se procesa aqui mismo y los resultados se
guardan en la carpeta de la planilla.
"""

import os
import re
import io
import sys
import json
import time
import socket
import shutil
import datetime
import threading
import webbrowser
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
_LIB = os.path.join(AQUI, "lib")
if os.path.isdir(_LIB) and _LIB not in sys.path:
    sys.path.insert(0, _LIB)

import planilla as P  # noqa: E402

CARPETA = os.path.dirname(AQUI)


def version_programa():
    """Cambia cuando se actualiza el programa: una pantalla vieja lo detecta."""
    t = 0
    for n in ("interfaz.html", "servidor.py", "planilla.py"):
        r = os.path.join(AQUI, n)
        if os.path.isfile(r):
            t += int(os.path.getmtime(r))
    return str(t)
CARPETA_PLANTILLAS = os.path.join(CARPETA, "Plantillas")
CARPETA_HORAS = os.path.join(CARPETA, "Horas recibidas")
PUERTO_BASE = 8765
MAX_SUBIDA = 60 * 1024 * 1024


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def parse_multipart(cuerpo, boundary):
    """Parser minimo de multipart/form-data (sin dependencias externas)."""
    campos = {}
    sep = b"--" + boundary
    for parte in cuerpo.split(sep)[1:-1]:
        if parte.startswith(b"\r\n"):
            parte = parte[2:]
        cabeza, _, datos = parte.partition(b"\r\n\r\n")
        if datos.endswith(b"\r\n"):
            datos = datos[:-2]
        cab = cabeza.decode("utf-8", "replace")
        m = re.search(r'name="([^"]*)"', cab)
        if not m:
            continue
        nombre = m.group(1)
        fn = re.search(r'filename="([^"]*)"', cab)
        if fn and fn.group(1):
            campos[nombre] = (fn.group(1), datos)
        else:
            campos[nombre] = datos.decode("utf-8", "replace")
    return campos


def _listar(carpeta):
    out = []
    if not os.path.isdir(carpeta):
        return out
    for n in sorted(os.listdir(carpeta)):
        if n.lower().endswith((".xlsx", ".xlsm")) and not n.startswith("~$"):
            ruta = os.path.join(carpeta, n)
            mt = os.path.getmtime(ruta)
            out.append({
                "nombre": n,
                "modificado": time.strftime("%d/%m/%Y %H:%M", time.localtime(mt)),
                "mtime": mt,
                "tam": round(os.path.getsize(ruta) / 1024.0),
            })
    return out


def horas_de_la_carpeta():
    """Archivos de horas guardados en 'Horas recibidas', del mas nuevo al mas viejo.

    Se ordena por la fecha real (mtime). Ordenar por el texto "dd/mm/aaaa" no da
    orden cronologico: comparaba primero el dia y ponia arriba archivos viejos.
    """
    out = _listar(CARPETA_HORAS)
    out.sort(key=lambda x: x["mtime"], reverse=True)
    return out


def xlsx_de_la_carpeta():
    salida = []
    for n in sorted(os.listdir(CARPETA_PLANTILLAS) if os.path.isdir(CARPETA_PLANTILLAS) else []):
        if n.lower().endswith((".xlsx", ".xlsm")) and not n.startswith("~$"):
            ruta = os.path.join(CARPETA_PLANTILLAS, n)
            salida.append({
                "nombre": n,
                "modificado": time.strftime("%d/%m/%Y %H:%M",
                                            time.localtime(os.path.getmtime(ruta))),
                "mtime": os.path.getmtime(ruta),
                "tam": round(os.path.getsize(ruta) / 1024.0),
                "es_madre": "madre" in P.sin_tildes(n) and "lista" not in P.sin_tildes(n),
                "es_lista": False,
            })
    salida.sort(key=lambda x: x["nombre"])
    return salida


SUBIDAS = {}          # id -> ruta del archivo ya subido
_candado = threading.Lock()


def guardar_temporal(nombre, datos):
    """Guarda un archivo subido y devuelve (id, ruta).

    Cada subida va en su propia subcarpeta: si se suben dos archivos distintos
    que se llaman igual (tipico al unificar planillas de varias computadoras),
    antes el segundo pisaba al primero y se unificaba dos veces el mismo.
    """
    with _candado:
        idx = "a%d" % (len(SUBIDAS) + 1)
        tmp = os.path.join(CARPETA, ".subidas", idx)
    if not os.path.isdir(tmp):
        os.makedirs(tmp)
    limpio = re.sub(r"[^\w \-.()]", "_", nombre) or "archivo.xlsx"
    ruta = os.path.join(tmp, limpio)
    with open(ruta, "wb") as f:
        f.write(datos)
    with _candado:
        SUBIDAS[idx] = ruta
    return idx, ruta


def ruta_de_horas(campos):
    """Acepta el archivo subido, su id, o el nombre de uno de 'Horas recibidas'."""
    idx = campos.get("id")
    if idx and idx in SUBIDAS:
        return SUBIDAS[idx], idx
    horas = campos.get("horas")
    if isinstance(horas, tuple):
        return guardar_temporal(horas[0], horas[1])[1], None
    if isinstance(horas, str) and horas.strip():
        cand = os.path.join(CARPETA_HORAS, os.path.basename(horas))
        if os.path.isfile(cand):
            return cand, None
    return None, None


def inspeccionar(ruta_horas, ruta_madre, pago=None):
    """Lo que se puede saber ANTES de procesar, para mostrarlo en pantalla."""
    info = {"ok": True}
    dias = P.detectar_periodo(ruta_horas)
    if not dias:
        return {"ok": False,
                "error": "Ese archivo no tiene la hoja 'Horas de trabajo' con las "
                         "fechas del periodo en la fila 2. Revise que sea el archivo "
                         "de horas correcto."}
    info["dias"] = len(dias)
    info["inicio"] = dias[0].strftime("%d/%m/%Y")
    info["fin"] = dias[-1].strftime("%d/%m/%Y")
    info["ciclo_ok"] = (len(dias) == P.DIAS_CICLO)

    ruta_estado, nombre_estado, primera = P.elegir_estado(CARPETA, ruta_madre, pago)
    est = P.leer_estado(ruta_estado)
    prev_fin, prev_pago = est["fin"], est["pago"]
    info["plantilla"] = os.path.basename(ruta_madre)
    info["estado"] = nombre_estado
    info["primera"] = primera
    info["salida"] = None

    info["periodo_anterior_fin"] = prev_fin.strftime("%d/%m/%Y") if prev_fin else None
    if prev_pago:
        info["pago_sugerido"] = (prev_pago + datetime.timedelta(days=P.DIAS_CICLO)).isoformat()
    else:
        info["pago_sugerido"] = (dias[-1] + datetime.timedelta(days=12)).isoformat()

    info["continuidad"] = None
    if prev_fin:
        esperado = prev_fin + datetime.timedelta(days=1)
        d = (dias[0] - esperado).days
        if d > 0:
            info["continuidad"] = "Quedarian %d dias sin pagar entre el periodo anterior (cerro el %s) y este." % (d, prev_fin.strftime("%d/%m/%Y"))
        elif d < 0:
            info["continuidad"] = "Se estarian pagando %d dias dos veces: el periodo anterior cerro el %s." % (abs(d), prev_fin.strftime("%d/%m/%Y"))

    info["ventana"] = est["posicion"]
    fp = pago or (P.a_fecha(datetime.datetime.strptime(info["pago_sugerido"],
                                                       "%Y-%m-%d")) )
    if fp:
        info["salida"] = "%s%s.xlsx" % (P.PREFIJO_SALIDA, fp.strftime("%d-%m-%Y"))
        yaesta = os.path.join(CARPETA, P.CARPETA_GENERADAS, info["salida"])
        info["ya_existe"] = os.path.isfile(yaesta)
    return info


def generadas_para_unir():
    salida = []
    for g in P.listar_generadas(CARPETA):
        salida.append({
            "nombre": g["nombre"],
            "pago": g["pago"].strftime("%d/%m/%Y"),
            "periodo": "%s al %s" % (g["ini"].strftime("%d/%m/%Y") if g["ini"] else "?",
                                     g["fin"].strftime("%d/%m/%Y") if g["fin"] else "?"),
            "mes": "%02d/%d" % (g["pago"].month, g["pago"].year),
        })
    salida.reverse()
    return salida


def a_json_secciones(avisos, secciones):
    out = []
    for clave, titulo, encab, nota in secciones:
        filas = avisos.items.get(clave, [])
        out.append({"clave": clave, "titulo": titulo, "encabezados": encab,
                    "nota": nota, "cantidad": len(filas),
                    "filas": [[("" if v is None else v) for v in f] for f in filas]})
    return out


def a_json_avisos(avisos):
    salida = []
    for clave, titulo, encabezados, nota in P.SECCIONES:
        filas = avisos.items.get(clave, [])
        salida.append({
            "clave": clave,
            "titulo": re.sub(r"^\d+[ab]? ", "", titulo),
            "encabezados": encabezados,
            "nota": nota,
            "filas": [[("" if v is None else v) for v in f] for f in filas],
            "cantidad": len(filas),
        })
    return salida


# ---------------------------------------------------------------------------
# Servidor
# ---------------------------------------------------------------------------

class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, *a):
        pass

    def _enviar(self, codigo, cuerpo, tipo="application/json; charset=utf-8",
                extra=None):
        if isinstance(cuerpo, str):
            cuerpo = cuerpo.encode("utf-8")
        self.send_response(codigo)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(cuerpo)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Connection", "close")
        self.close_connection = True
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(cuerpo)

    def _json(self, obj, codigo=200):
        self._enviar(codigo, json.dumps(obj, ensure_ascii=False, default=str))

    # -- GET ----------------------------------------------------------------
    def do_GET(self):
        ruta = urllib.parse.urlparse(self.path)
        if ruta.path in ("/", "/index.html"):
            return self._enviar(200, PAGINA, "text/html; charset=utf-8")
        if ruta.path == "/guia":
            g = os.path.join(AQUI, "LEEME.txt")
            if not os.path.isfile(g):
                return self._enviar(404, "No encontre la guia.",
                                    "text/plain; charset=utf-8")
            with open(g, encoding="utf-8") as fh:
                return self._enviar(200, fh.read(), "text/plain; charset=utf-8")
        if ruta.path == "/salud":
            return self._json({"app": "planilla", "version": version_programa()})
        if ruta.path == "/generadas":
            return self._json({"archivos": generadas_para_unir()})
        if ruta.path == "/archivos":
            return self._json({"archivos": xlsx_de_la_carpeta(),
                               "horas": horas_de_la_carpeta(),
                               "carpeta": CARPETA_PLANTILLAS,
                               "carpeta_horas": CARPETA_HORAS})
        if ruta.path == "/descargar":
            q = urllib.parse.parse_qs(ruta.query)
            nombre = (q.get("f") or [""])[0]
            base = os.path.basename(nombre)
            destino = None
            for d in (os.path.join(CARPETA, P.CARPETA_GENERADAS),
                      CARPETA_PLANTILLAS, CARPETA_HORAS, CARPETA):
                cand = os.path.join(d, base)
                if os.path.isfile(cand):
                    destino = cand
                    break
            if not destino:
                return self._json({"error": "No existe"}, 404)
            with open(destino, "rb") as f:
                datos = f.read()
            cab = {"Content-Disposition":
                   "attachment; filename*=UTF-8''%s"
                   % urllib.parse.quote(os.path.basename(destino))}
            return self._enviar(
                200, datos,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                cab)
        return self._json({"error": "No encontrado"}, 404)

    # -- POST ---------------------------------------------------------------
    def do_POST(self):
        largo = int(self.headers.get("Content-Length") or 0)
        if largo > MAX_SUBIDA:
            return self._json({"ok": False, "error": "El archivo es demasiado grande."}, 413)
        ctype = self.headers.get("Content-Type") or ""
        m = re.search(r"boundary=([^;]+)", ctype)
        if not m:
            return self._json({"ok": False, "error": "Envio invalido."}, 400)
        cuerpo = b""
        while len(cuerpo) < largo:
            trozo = self.rfile.read(min(65536, largo - len(cuerpo)))
            if not trozo:
                break
            cuerpo += trozo
        campos = parse_multipart(cuerpo, m.group(1).strip('"').encode())

        try:
            if self.path == "/subir":
                horas = campos.get("horas")
                if not isinstance(horas, tuple):
                    return self._json({"ok": False, "error": "No llego el archivo."})
                idx, ruta = guardar_temporal(horas[0], horas[1])
                return self._json({"ok": True, "id": idx,
                                   "nombre": os.path.basename(ruta)})

            if self.path == "/subir-planilla":
                # una planilla traida de otra computadora o de un mes anterior
                arch = campos.get("planilla") or campos.get("horas")
                if not isinstance(arch, tuple):
                    return self._json({"ok": False, "error": "No llego el archivo."})
                idx, ruta = guardar_temporal(arch[0], arch[1])
                try:
                    ficha = P.resumen_planilla_ccss(ruta)
                except Exception as e:
                    with _candado:
                        SUBIDAS.pop(idx, None)
                    shutil.rmtree(os.path.dirname(ruta), ignore_errors=True)
                    return self._json({"ok": False, "error": str(e)})
                ficha["ok"] = True
                ficha["id"] = idx
                ficha["subido"] = True
                return self._json(ficha)

            if self.path == "/unificar":
                nombres = [n for n in (campos.get("archivos") or "").split("|") if n]
                rutas = [os.path.join(CARPETA, P.CARPETA_GENERADAS,
                                      os.path.basename(n)) for n in nombres]
                faltan = [os.path.basename(r) for r in rutas if not os.path.isfile(r)]
                if faltan:
                    return self._json({"ok": False,
                                       "error": "No encontre: " + ", ".join(faltan)})
                # planillas subidas desde otro lado
                for idx in (campos.get("subidos") or "").split("|"):
                    if not idx:
                        continue
                    ruta = SUBIDAS.get(idx)
                    if not ruta or not os.path.isfile(ruta):
                        return self._json({
                            "ok": False,
                            "error": "Se perdio uno de los archivos subidos. "
                                     "Vuelva a subirlo."})
                    rutas.append(ruta)
                if not rutas:
                    return self._json({"ok": False,
                                       "error": "No se escogio ninguna planilla."})
                salida, filas, periodos, av = P.unificar(
                    rutas, CARPETA, campos.get("etiqueta") or None,
                    log=lambda *a: None)
                SEC = [
                    ("movimiento", "Movimientos detectados",
                     ["Carné", "Nombre", "Estado", "Fecha IC", "Fecha EX",
                      "De dónde salió"], ""),
                    ("sin_fecha", "Salidas sin fecha",
                     ["Carné", "Nombre", "Estado", "Detalle", "Nota"], ""),
                    ("ingreso_dudoso", "Fechas de ingreso a confirmar",
                     ["Carné", "Nombre", "Fecha usada", "De dónde salió",
                      "Qué hacer"], ""),
                    ("multiple", "Casos con más de una condición",
                     ["Carné", "Nombre", "Estado asignado", "Detalle", "Nota"], ""),
                    ("parcial", "Aparecen en algunos períodos",
                     ["Carné", "Nombre", "En cuántos", "De cuántos",
                      "Faltó en"], ""),
                ]
                estados = {}
                for f in filas:
                    estados[f["estado"]] = estados.get(f["estado"], 0) + 1
                return self._json({
                    "ok": True,
                    "archivo": os.path.basename(salida),
                    "carpeta": P.CARPETA_GENERADAS,
                    "personas": len(filas),
                    "estados": estados,
                    "bruto": round(sum(f["bruto"] for f in filas), 2),
                    "dias": round(sum(f["dias"] for f in filas), 2),
                    "periodos": [{"archivo": p["archivo"],
                                  "pago": p["pago"].strftime("%d/%m/%Y"),
                                  "origen": p.get("origen", "")}
                                 for p in periodos],
                    "secciones": a_json_secciones(av, SEC),
                    "muestra": [[f["carne"], f["dias"], f["puesto"], f["nombre"],
                                 f["cedula"], f["bruto"], f["estado"],
                                 f["f_ic"].strftime("%d/%m/%Y") if f["f_ic"] else "",
                                 f["f_ex"].strftime("%d/%m/%Y") if f["f_ex"] else ""]
                                for f in filas],
                })

            ruta_horas, _ = ruta_de_horas(campos)
            madre = campos.get("madre") or ""
            ruta_madre = os.path.join(CARPETA_PLANTILLAS, os.path.basename(madre))

            if not ruta_horas or not os.path.isfile(ruta_horas):
                return self._json({"ok": False,
                                   "error": "Se perdio el archivo de horas. Vuelva a "
                                            "seleccionarlo con el boton Cambiar."})
            if not os.path.isfile(ruta_madre):
                return self._json({"ok": False, "error": "No encontre la Planilla Madre."})

            if self.path == "/inspeccionar":
                pg = campos.get("pago")
                fp = None
                if pg:
                    try:
                        fp = datetime.datetime.strptime(pg, "%Y-%m-%d").date()
                    except ValueError:
                        fp = None
                return self._json(inspeccionar(ruta_horas, ruta_madre, fp))

            if self.path == "/generar":
                pago = datetime.datetime.strptime(campos["pago"], "%Y-%m-%d").date()
                salida, reporte, resumen, avisos = P.procesar(
                    ruta_horas, ruta_madre, CARPETA, pago,
                    dias_blanco_como_ausencia=(campos.get("blancos") == "1"),
                    prorratear_ingreso=(campos.get("ingresos") == "1"),
                    log=lambda *a: None)
                return self._json({
                    "ok": True,
                    "carpeta": P.CARPETA_GENERADAS,
                    "resumen": [[k, v] for k, v in resumen],
                    "sesiones": a_json_avisos(avisos),
                    "planilla": os.path.basename(salida),
                    "reporte": os.path.basename(reporte),
                })
        except Exception as e:
            import traceback
            traceback.print_exc()
            return self._json({"ok": False, "error": str(e)})
        return self._json({"ok": False, "error": "Ruta desconocida"}, 404)


def ya_esta_abierto(puerto):
    """True si en ese puerto ya responde ESTE mismo programa."""
    import urllib.request
    try:
        with urllib.request.urlopen("http://127.0.0.1:%d/salud" % puerto,
                                    timeout=1.5) as r:
            return json.loads(r.read().decode("utf-8")).get("app") == "planilla"
    except Exception:
        return False


def elegir_puerto(base):
    """Devuelve (puerto, ya_abierto)."""
    for p in range(base, base + 25):
        s = socket.socket()
        # mismo SO_REUSEADDR que usa el servidor real: sin esto, un puerto que
        # acaba de quedar libre (TIME_WAIT) se ve ocupado y el programa se
        # mudaba de puerto dejando pantallas viejas apuntando al anterior.
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            s.bind(("127.0.0.1", p))
            s.close()
            return p, False
        except OSError:
            s.close()
            if ya_esta_abierto(p):
                return p, True
    return base, False


def main():
    with open(os.path.join(AQUI, "interfaz.html"), encoding="utf-8") as f:
        global PAGINA
        PAGINA = f.read()
    puerto, abierto = elegir_puerto(PUERTO_BASE)
    url = "http://127.0.0.1:%d/" % puerto
    print("")
    print("  =========================================================")
    print("     GENERADOR DE PLANILLA BISEMANAL")
    print("  =========================================================")
    print("")
    if abierto:
        print("  El programa YA ESTABA ABIERTO. Se reutiliza esa ventana.")
        print("  Direccion:  %s" % url)
        print("")
        print("  Puede cerrar esta ventana negra: la otra es la que manda.")
        print("")
        sys.stdout.flush()
        webbrowser.open(url)
        time.sleep(2)
        return
    srv = ThreadingHTTPServer(("127.0.0.1", puerto), Handler)
    print("  Se abrio en el navegador:  %s" % url)
    print("")
    print("  Todo se procesa en esta computadora. Nada sale a internet.")
    print("")
    print("  Para cerrar: cierre esta ventana negra, o presione Ctrl+C")
    print("")
    sys.stdout.flush()
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n  Cerrado.\n")
    finally:
        tmp = os.path.join(CARPETA, ".subidas")
        if os.path.isdir(tmp):
            shutil.rmtree(tmp, ignore_errors=True)


PAGINA = ""

if __name__ == "__main__":
    main()
