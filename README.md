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
| `planilla.py` | el motor de cálculo |
