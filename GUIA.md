# App de incidencias · Escuela 26-27 — Guía de puesta en marcha

La app escribe **directamente en tu hoja de Google**. Las hojas de los meses siguen igual: los responsables apuntan desde el móvil, y tú sigues viendo y exportando el Excel como siempre.

```
incidencias/
├── apps-script/
│   ├── Code.gs        ← el "motor" (va dentro de la hoja de Google)
│   └── Index.html     ← la app (copia de app/index.html)
└── app/               ← la misma app, preparada para instalarse en el móvil
    ├── index.html, manifest.webmanifest, sw.js, icon-192.png, icon-512.png
```

---

## 1. Pasar el Excel a Hoja de cálculo de Google (una sola vez)

Ahora mismo el archivo es un **.xlsx** subido a Drive. Para que funcione la app tiene que ser una hoja de Google nativa:

1. Abre el archivo en Drive → **Archivo → Guardar como Hojas de cálculo de Google**.
2. Se crea una copia nueva. **Esa** es la que se queda a partir de ahora: compártela con quien la necesite y quita o renombra el .xlsx viejo (por ejemplo `NO USAR – …`) para que nadie siga apuntando ahí.

## 2. Pegar el código

1. En la hoja nueva: **Extensiones → Apps Script**.
2. Borra lo que haya en `Código.gs` y pega todo el contenido de `apps-script/Code.gs`.
3. Pulsa **+ → HTML**, llámalo exactamente `Index` y pega el contenido de `apps-script/Index.html`.
4. Guarda (icono del disquete).

## 3. Crear las hojas de configuración

1. En el editor de Apps Script, elige la función **`configurarInicial`** en el desplegable y pulsa **Ejecutar**. La primera vez Google te pedirá permisos: acéptalos (es tu propio script).
2. Vuelve a la hoja. Verás hojas nuevas que empiezan por `_`:

| Hoja | Qué tienes que hacer |
|---|---|
| **_Accesos** | Un **código** por responsable (ya generados). La fila *Administración* (`*`) eres tú: ves todos los grupos. Alaitz tiene sus dos grupos (Media y Ribera) en un solo código. Si un responsable tiene datos propios en la hoja *Responsables*, la columna D indica su nombre allí; si está vacía, no verá la pestaña «Mis datos» (por ejemplo, Alaitz: si quieres que la vea, añade su fila en *Responsables* y pon el nombre aquí). |
| **_Monitores** | Lista de monitores por grupo con sus **días fijos L M X J V** (casillas). Los he rellenado a partir de lo que había en octubre y noviembre, **pero muchos monitores no tenían días apuntados**: revisa y marca los días que va cada uno. Para dar de baja a alguien desmarca **Activo**; para un monitor nuevo, añade una fila. |
| **_Festivos** | Días sin actividad (no salen premarcados). Hay algunos festivos de Navarra puestos; **añade las vacaciones de Navidad y Semana Santa** de la escuela. |
| **_Ajustes** | `MESES_CERRADOS`: cuando hayas enviado un mes a la asesoría, escribe aquí `Octubre` (o `Octubre; Noviembre`…) y nadie podrá modificarlo desde la app. |
| _Dias, _Registro | Las rellena la app sola (días concretos marcados e historial de quién guardó qué). No tocar. |

Puedes volver a ejecutar `configurarInicial` cuando quieras (también desde el menú **App incidencias** de la hoja): solo añade lo que falte, nunca borra.

## 4. Publicar

En Apps Script: **Implementar → Nueva implementación → tipo «Aplicación web»**
- Ejecutar como: **Yo**
- Quién tiene acceso: **Cualquier usuario**

Copia la URL que termina en **`/exec`**.

> Los responsables **no necesitan cuenta de Google**. La seguridad la da su código: el servidor comprueba en cada petición qué grupos puede ver ese código y rechaza cualquier otro.

> Si más adelante cambias el código: **Implementar → Gestionar implementaciones → editar (lápiz) → Versión: Nueva**. Así la URL no cambia.

## 5. Cómo se instalan los responsables la app

### Opción A — rápida (sin nada más)
Manda la URL `/exec`. En el móvil la abren y hacen **Compartir → Añadir a pantalla de inicio** (iPhone) o **⋮ → Añadir a pantalla de inicio** (Android). Funciona, pero se abre dentro del navegador y Google muestra un aviso arriba.

### Opción B — app instalada de verdad (recomendada)
1. En `app/index.html` busca `const API_URL = '';` y pega dentro tu URL `/exec`.
2. Sube la carpeta `app/` a un alojamiento gratuito: **GitHub Pages** o **Netlify Drop** (arrastrar la carpeta en app.netlify.com/drop).
3. Manda esa dirección. Al abrirla, el móvil ofrece **Instalar**: queda con su icono, a pantalla completa y sin barras del navegador.

**Enlace directo con código:** si añades `?c=CODIGO` al final de la dirección (p. ej. `https://…/?c=T2L8A4`), el responsable entra directamente sin escribir el código. Después queda recordado en su móvil.

## 6. Cómo funciona para el responsable

1. Elige el **mes** arriba (octubre a junio).
2. Ve **solo sus monitores**, con sus días fijos y el estado: *Pendiente*, *En hoja* (había datos de antes) o *✓ Hecho*.
3. Al tocar un monitor sale un **calendario de lunes a viernes** con sus días fijos ya marcados (menos festivos). Desmarca los días que no fue y marca los extra. La app cuenta sola los viajes de cada día de la semana y los escribe en las columnas LUNES…VIERNES.
4. Si quedan días fijos sin marcar, propone ponerlos como **ausencias** con un toque.
5. Partidos de liga, sustituciones, km especiales, peajes y observaciones (con atajos «Sustituye a…», «Falta el…»).
6. Pestaña **Mis datos**: sus km totales, partidos, sustituciones, peajes y observaciones del mes (hoja *Responsables*).

Si un mes todavía no tiene hoja (Diciembre, Enero…), la app la crea sola con la misma cabecera y todos los monitores activos.
