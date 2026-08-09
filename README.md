# 10 Smiths Avenue, Redcliffe — app de registro rápido

App web simple (para el celular) para anotar al momento: gastos de la casa, pagos de arriendo de los inquilinos, y pagos al arrendador — sin abrir la planilla de Google Sheets. Cada registro se agrega solo como fila nueva en la planilla real ("10 Smiths Avenue, Redcliffe").

## Cómo está armado

- **`docs/`** — la app (HTML/CSS/JS, sin dependencias externas). Se publica en GitHub Pages (GitHub solo permite servir Pages desde `/` o `/docs`, por eso el nombre).
- **`apps-script/`** — el código que corre en Google Apps Script, conectado a la planilla, que recibe lo que manda la app y escribe las filas. Ver `apps-script/README.md` para desplegarlo (paso manual, una sola vez, ~5 min).

## Inteligencia incluida

Si registras un gasto de tipo **Electricity, Water, Gas o Internet**, la app calcula automáticamente el porcentaje que le corresponde pagar al inquilino del Room 2 (50% por defecto) y crea un **cobro pendiente** en la pestaña "Tenants" — no tienes que hacer el cálculo ni la segunda anotación a mano.

## Qué NO hace (a propósito)

No toca ni recalcula las columnas de totales acumulados que ya existen en la planilla (Rent/Bond Held/Internet/Water/Electricity/Gas por inquilino) — solo agrega filas nuevas usando las columnas de cada pestaña. Si en algún momento quieres que esos totales se actualicen solos, lo mejor es convertirlos a fórmulas (`SUMIF`/`SUMIFS`) dentro de la misma planilla — pídemelo y te lo dejo armado.

## Privacidad

La app pide un **PIN** (6 dígitos) antes de mostrar nada — evita que alguien que encuentre el link por casualidad la abra y la use. No es seguridad fuerte (el PIN vive en el código de la página, visible para quien sepa mirarlo), es solo una barrera contra el acceso casual. Para que esto funcione, la implementación de Apps Script debe seguir con acceso **"Cualquier usuario"** (no "Solo yo") — la privacidad la da el PIN del frontend, no el login de Google.

## Puesta en marcha (resumen)

1. Despliega el backend — ver [`apps-script/README.md`](./apps-script/README.md).
2. Abre la app publicada, toca el ⚙ y pega la URL del Web App que te dio Google (el secreto ya viene precargado).
3. Listo — queda guardado en el celular, no hay que repetir la configuración cada vez.

## Cómo se prueban los cambios (sin que tengas que estar en la computadora a cada rato)

Los cambios de `Code.gs` se prueban primero contra una copia de tus datos reales, en `apps-script/test/` — sin tocar tu planilla ni tu Drive, y sin pedirte nada. Solo cuando esas pruebas pasan se te manda el `Code.gs` final para que lo pegues y despliegues **una vez** — no en cada ajuste chico. Ver [`apps-script/test/README.md`](./apps-script/test/README.md).
