# Generador de Planilla Bisemanal

Herramienta interna de Ingeniería Estrella S.A. para armar la planilla
bisemanal del personal operativo y el reporte mensual para la CCSS.

El ciclo es de **14 días exactos**: 26 pagos al año, no 24.

## Dos maneras de usarlo, un solo motor

| | Dónde vive | Cómo corre |
|---|---|---|
| **Local** | [`programa/`](programa/) | Doble clic al lanzador. Levanta un servidor en `127.0.0.1` y se maneja desde el navegador. Necesita Python 3. |
| **Navegador** | [`docs/`](docs/) → <https://gabomomo.github.io/planillasdcc/> | Todo dentro del navegador con [Pyodide](https://pyodide.org/). No necesita instalar nada. |

Las dos usan el **mismo** `planilla.py`. Está duplicado a propósito
(`programa/planilla.py` y `docs/planilla.py`) porque GitHub Pages sirve
archivos estáticos y no sigue enlaces simbólicos. Para que no se separen:

- el gancho de `pre-commit` no deja commitear si difieren
- la revisión automática de GitHub falla si difieren

Si cambia el motor, copie uno sobre el otro antes de commitear:

```sh
cp programa/planilla.py docs/planilla.py
```

**En ningún caso los archivos salen a internet.** Todo se calcula en la
computadora de quien lo usa.

## Instalación para trabajar en el código

```sh
git clone https://github.com/gabomomo/planillasdcc.git
cd planillasdcc
git config core.hooksPath .githooks    # una sola vez: protege los datos
```

## Sobre la pantalla de acceso del sitio

El sitio pide correo y contraseña, pero **no es seguridad real**: al ser un
sitio estático, la comprobación ocurre en el navegador de quien entra y
cualquiera con conocimientos técnicos la puede saltar. Sirve para que nadie
entre por casualidad. La contraseña no está en el código: solo queda su
derivación PBKDF2-SHA256 con 310.000 iteraciones y sal aleatoria.

Lo que sí protege de verdad es que **este sitio no contiene datos**. Los
archivos de planilla se guardan en el navegador de cada persona (IndexedDB,
que es por origen y por computadora), nunca aquí.

Si en algún momento hace falta protección de verdad, el camino es un dominio
propio detrás de Cloudflare Access o equivalente, que valida contra el correo
antes de servir la página. GitHub Pages por sí solo no puede hacerlo.

## Advertencia: nunca suba datos de planilla

Este repositorio es **público** y contiene solo el programa. La Planilla
Madre, los archivos de horas y las planillas generadas llevan nombres,
cédulas, teléfonos y salarios de personas reales. El `.gitignore` los
bloquea y el gancho de `pre-commit` detiene el commit, pero la
responsabilidad final es de quien lo hace. Borrar un archivo después **no**
lo quita del historial de Git.

## Cómo se usa

La guía completa para quien opera la planilla está en
[`programa/LEEME.txt`](programa/LEEME.txt): qué archivos hacen falta, los dos
criterios del paso 3, cómo unificar el mes para la CCSS y qué hacer si algo
falla.

## La jornada mixta y nocturna

La Planilla Madre **calcula** el diferencial de la hora mixta (columna Y) y
de la nocturna (columna AE), pero su fórmula de «Total Salario Bruto»
(columna AU) **no los suma**: solo toma las columnas de horas *extra* mixtas
y nocturnas. O sea que ese diferencial hoy no se paga.

El programa respeta eso, porque su trabajo es llenar la Madre, no cambiar lo
que paga. El monto exacto por persona sale siempre en la hoja «4 Jornada
mixta y nocturna» del reporte de revisión, para que contabilidad decida. Si
corresponde pagarlo, hay que cambiar la fórmula de la columna AU en la
Planilla Madre.

## Estructura

```
programa/          version local
  servidor.py        el servidor en 127.0.0.1
  interfaz.html      la pantalla
  planilla.py        el motor de calculo
  lib/               openpyxl incluido, para no instalar nada
  LEEME.txt          guia para quien opera la planilla
docs/              version de navegador (esto es lo que sirve Pages)
  index.html         la pantalla
  app.js             arranque de Pyodide y guardado en IndexedDB
  ui.js              la logica de la pantalla
  acceso.js          la pantalla de acceso
  planilla.py        copia identica del motor
```
