# Desplegar el backend (Google Apps Script)

Esto conecta la app web con tu planilla real. Se hace una sola vez, toma ~5 minutos, y **tiene que hacerlo tu cuenta de Google** (no se puede automatizar desde afuera — Google exige que tú autorices el acceso a tu propia planilla).

## Pasos

1. Abre la planilla **"10 Smiths Avenue, Redcliffe"** en Google Sheets.
2. Menú **Extensiones → Apps Script**.
3. Se abre un editor con un archivo `Code.gs` vacío (o con contenido de ejemplo). Borra todo y pega el contenido completo del archivo [`Code.gs`](./Code.gs) de esta carpeta.
4. Guarda (ícono de disco o Ctrl+S). Ponle un nombre al proyecto, ej. "Smiths Ave API".
5. Arriba a la derecha, botón **Implementar → Nueva implementación**.
6. Click en el ícono de engranaje junto a "Seleccionar tipo" → elige **Aplicación web**.
7. Configura:
   - **Ejecutar como:** Yo (tu cuenta)
   - **Quién tiene acceso:** Cualquier usuario ("Anyone")
8. Click **Implementar**. Google te va a pedir autorizar permisos la primera vez — acepta (es tu propia planilla, es seguro).
9. Copia la **URL de la aplicación web** que te muestra al final (termina en `/exec`). Esa es la URL que vas a pegar en la app web, en la pantalla de configuración inicial.

## El secreto compartido

El archivo `Code.gs` ya trae un secreto único generado para ti:

```
5EGVxQhUJ2RsQ10dFUEkHAuM
```

La app web ya viene con este mismo valor precargado — no tienes que hacer nada, a menos que quieras cambiarlo. Si lo cambias en un lado, cámbialo también en el otro (en `Code.gs` la constante `SECRET`, y en la app web en Configuración).

**Nota de seguridad honesta:** este secreto viaja dentro del código de la app web, que es pública (GitHub Pages). No es una protección fuerte — es solo una barrera contra que alguien encuentre la URL por accidente y escriba datos random. Para este uso (una planilla de arriendo personal, no información sensible de terceros) es una relación costo/beneficio razonable. Si en algún momento quieres algo más serio, la alternativa es exigir login de Google en el Web App ("Cualquier usuario con cuenta de Google") y validar el email del que llama dentro del script.

## Si necesitas volver a desplegar después de editar `Code.gs`

Cada vez que cambies el código: **Implementar → Administrar implementaciones → ícono de lápiz → Nueva versión → Implementar**. Si creas una implementación totalmente nueva en vez de editar la existente, la URL cambia y hay que actualizarla en la app web.

## Cómo probar que quedó bien

Con la URL ya copiada, puedes probar desde el navegador (reemplaza `TU_URL` y pega el secreto):

```
TU_URL?secret=5EGVxQhUJ2RsQ10dFUEkHAuM&recent=3
```

Si responde con JSON (`{"ok":true,"recent":{...}}`) en vez de un error de Google, quedó bien desplegado.
