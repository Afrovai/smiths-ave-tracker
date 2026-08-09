# Pruebas locales — sin tocar tu planilla ni tu Drive

Simulador de las APIs de Google Apps Script (`SpreadsheetApp`, `DriveApp`, `Utilities`, `ContentService`) corriendo en Node, cargado con una foto de tus datos reales (`real-snapshot.md`, texto plano, sin fórmulas ni nada sensible). Corre `Code.gs` tal cual, sin modificarlo.

## Por qué existe

Antes, cada cambio de código requería que lo pegaras en Apps Script y lo desplegaras para poder probarlo — varias vueltas de ida y vuelta. Ahora los cambios se prueban acá primero, tantas veces como haga falta, y solo se te pide pegar/desplegar cuando ya está confirmado que funciona.

## Dos suites — cuál usa qué datos

- **`run.js`** — datos **inventados** (Ana Test, Beto Test, etc.). Este es el que vive en git, seguro para un repo público.
- **`test-real-data.js`** — datos **reales** tuyos (`real-snapshot.md`, una foto de tu planilla en un momento dado). Ninguno de los dos archivos se sube a git (ver `.gitignore` en la raíz del repo) — el repo es público y esos datos no deben quedar ahí.

## Cómo correrlas

```
cd apps-script/test
node run.js              # datos inventados, siempre disponible
node test-real-data.js   # datos reales, solo si existe real-snapshot.md
```

Si todo pasa, termina con `N OK, 0 FAIL`. Si algo falla, se corrige `Code.gs` y se corre de nuevo — sin necesidad de estar frente a la planilla.

## Qué cubre

- El resumen (`summary`) da los totales correctos por inquilino, por pieza, y los "totales de tu Excel" son números reales, no texto.
- Crear una ficha nueva en "Arrendatarios" y editarla (mismo nombre = actualiza, no duplica).
- Subida de foto de ID (sin tocar Drive real — se verifica que se llama correctamente).
- División de cuenta compartida, con y sin gente cargada en el registro.
- "To Landlord" completa la fila prellenada de la fecha en vez de duplicar.

## Actualizar la foto de datos

`real-snapshot.md` es una copia de tus datos en un momento dado — no se actualiza sola. Si quieres que las pruebas reflejen datos más recientes, hay que volver a descargarla (se le pide a Claude).

## Límite de esto

Corre 100% en tu computadora — nunca escribe en tu Drive ni en tu planilla real. Antes de dar por buena una función nueva importante (sobre todo algo que toque Drive o la planilla), igual conviene una prueba controlada real, una sola vez, no en cada cambio chico.
