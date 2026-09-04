#!/bin/bash
# ============================================================
#  GENERADOR DE PLANILLA BISEMANAL - Ingenieria Estrella S.A.
#  Doble clic para abrir. Se abre solo en el navegador.
# ============================================================

cd "$(dirname "$0")" || exit 1
printf '\e[8;26;88t'
clear 2>/dev/null

TITULO="Generador de Planilla Bisemanal"

# Muestra un dialogo del sistema. Devuelve 0 si eligieron el boton de accion.
dialogo() {
    local msg="$1" ok="$2" no="$3"
    local r
    r=$(osascript -e "display dialog \"$msg\" with title \"$TITULO\" \
        buttons {\"$no\", \"$ok\"} default button \"$ok\" with icon caution" 2>/dev/null)
    case "$r" in *"$ok"*) return 0 ;; esac
    return 1
}

aviso() {
    osascript -e "display dialog \"$1\" with title \"$TITULO\" \
        buttons {\"Entendido\"} default button 1 with icon stop" >/dev/null 2>&1
}

buscar_python() {
    PY=""
    for c in /usr/bin/python3 /usr/local/bin/python3 /opt/homebrew/bin/python3 python3; do
        if command -v "$c" >/dev/null 2>&1 && "$c" -c "" >/dev/null 2>&1; then
            PY="$c"; return 0
        fi
    done
    return 1
}

# ---------- falta la libreria incluida ----------
if [ ! -d "programa/lib/openpyxl" ]; then
    aviso "Falta la carpeta programa/lib.\n\nEsa carpeta trae la libreria que lee los archivos de Excel.\n\nCopie de nuevo la carpeta completa del programa, con todo lo que tiene adentro."
    echo ""
    echo "  FALTA LA CARPETA  programa/lib"
    echo "  Copie de nuevo la carpeta completa del programa."
    echo ""
    read -r -p "  Presione ENTER para cerrar."
    exit 1
fi

# ---------- falta Python ----------
if ! buscar_python; then
    MSG="Esta computadora todavia no tiene el motor que necesita el programa.\n\nSe instala UNA sola vez con el instalador oficial de Apple. Tarda unos minutos y necesita internet.\n\nNo se instala solo: es un componente del sistema y lo tiene que autorizar usted.\n\nQuiere abrir el instalador de Apple ahora?"
    if dialogo "$MSG" "Abrir el instalador" "Ahora no"; then
        xcode-select --install 2>/dev/null
        aviso "Siga las instrucciones de la ventana de Apple que acaba de aparecer.\n\nCuando termine la instalacion, vuelva a dar doble clic en Generar Planilla y el programa abrira normalmente.\n\n(Si Apple dice que ya estan instaladas, simplemente vuelva a dar doble clic.)"
    else
        aviso "Cuando quiera hacerlo:\n\n1. Abra la aplicacion Terminal\n2. Escriba:  xcode-select --install\n3. Presione ENTER\n\nAl terminar, vuelva a dar doble clic en Generar Planilla."
    fi
    echo ""
    echo "  Falta instalar las Herramientas de linea de comandos de Apple."
    echo "  Cuando termine, vuelva a dar doble clic en Generar Planilla."
    echo ""
    read -r -p "  Presione ENTER para cerrar."
    exit 1
fi

# ---------- todo listo ----------
echo ""
echo "  Guia de uso:  programa/LEEME.txt"
echo ""
"$PY" "programa/servidor.py"
