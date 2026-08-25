# Generador de Planilla Bisemanal

Herramienta interna de Ingeniería Estrella S.A. para armar la planilla
bisemanal del personal operativo y el reporte mensual para la CCSS.

## Cómo funciona

Todo corre **dentro del navegador**. No hay servidor ni base de datos:

- [Pyodide](https://pyodide.org/) ejecuta Python compilado a WebAssembly
- `openpyxl` lee y escribe los archivos de Excel, incluidos los **colores
  de celda** de los que dependen las ausencias e incapacidades
- Los archivos se guardan en el almacenamiento del navegador (IndexedDB)

**Ningún archivo se sube a ningún servidor.** Los datos de planilla nunca
salen de la computadora de quien la usa.

## Sobre la pantalla de acceso

El sitio pide correo y contraseña, pero **no es seguridad real**: al ser un
sitio estático, la comprobación ocurre en el navegador de quien entra y
cualquiera con conocimientos técnicos la puede saltar. Sirve para que
nadie entre por casualidad.

Lo que sí protege de verdad es que **este sitio no contiene datos**. Los
archivos de planilla se guardan en el navegador de cada persona
(IndexedDB, que es por origen y por computadora), nunca aquí.

La contraseña no está escrita en el código: solo queda su derivación
PBKDF2-SHA256 con 310.000 iteraciones y sal aleatoria.

Si en algún momento hace falta protección de verdad, el camino es poner
un dominio propio detrás de Cloudflare Access (o equivalente), que valida
contra el correo antes de servir la página. GitHub Pages por sí solo no
puede hacerlo.

## Advertencia importante

Este repositorio contiene **solo el programa**. No suba nunca la Planilla
Madre ni ninguna planilla generada: llevan nombres, cédulas, cuentas
bancarias y salarios. El `.gitignore` bloquea los archivos de Excel, pero
la responsabilidad final es de quien hace el commit.

## Archivos

| | |
|---|---|
| `index.html` | la pantalla |
| `app.js` | arranque de Pyodide y guardado en el navegador |
| `ui.js` | la lógica de la pantalla |
| `acceso.js` | la pantalla de acceso |
| `planilla.py` | el motor de cálculo |
