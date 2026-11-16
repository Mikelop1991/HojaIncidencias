# App de incidencias · Escuela 26-27 — Guía

**App:** https://mikelop1991.github.io/HojaIncidencias/

## Cómo está montado

| Dónde | Qué hay | Quién lo ve |
|---|---|---|
| **Hoja compartida del registro** | Solo tu plantilla de siempre: `Responsables`, `Octubre`, `Noviembre`… | Responsables y monitores |
| **Hoja privada «Configuración app incidencias»** | El programa (Apps Script) y las pestañas `_Accesos` (códigos), `_Monitores` (días fijos), `_Festivos`, `_Ajustes`, `_Dias`, `_Registro` | Solo tú |

- La app escribe **solo en las columnas E–O** (viajes L–V, liga, ausencias, sustituciones, km, peajes, observaciones) de la fila de cada monitor, y en las columnas C–G de `Responsables`. El resto de la plantilla no se toca.
- Si alguien rellena el Excel a mano, la app lee lo que haya en la hoja; lo de la hoja manda.
- La lista de monitores sale de **la hoja de cada mes**: si añades o quitas a alguien en el Excel, la app lo refleja.
- En la app solo salen los meses que ya tienen hoja + el mes en curso. Cuando alguien guarda en un mes que aún no existe, la hoja se crea **duplicando la del mes anterior** (mismo formato y monitores, columnas de datos vacías). También puedes crearla tú: menú **App incidencias → Crear hoja del mes siguiente** (en la hoja privada).

## Cambio a la hoja privada (una vez)

1. Crea una hoja de cálculo nueva en tu Drive: **«Configuración app incidencias»**. No la compartas.
2. En ella: **Extensiones → Apps Script**. Pega `apps-script/Code.gs` en `Código.gs` y crea un archivo HTML `Index` con `apps-script/Index.html`.
3. Arriba del todo de `Code.gs`, pega la URL de la hoja compartida del registro en `URL_HOJA_REGISTRO = '...'`. Guarda.
4. Ejecuta **`configurarInicial`** (acepta los permisos). Mueve tus pestañas `_Accesos`, `_Monitores`… de la hoja compartida a la privada **con los códigos que ya tengas** y las borra de la compartida.
5. Ejecuta **`cargarFestivosCalendario`**: carga los días sin actividad del calendario 26-27.
6. **Implementar → Nueva implementación → Aplicación web** (Ejecutar como: Yo · Acceso: Cualquier usuario). Copia la URL `/exec` nueva y pásasela a Claude (o pégala en `config.js` del repositorio).
7. En la **hoja compartida**: Extensiones → Apps Script → **Implementar → Gestionar implementaciones → Archivar**, y después borra el código de `Código.gs` e `Index`. Así nadie con acceso al Excel puede ver el programa.

## Día a día

- **Códigos** (`_Accesos`): cambia el de cualquiera escribiendo el que quieras en la columna A (sin distinguir mayúsculas ni tildes; que no se repitan). Se aplica en 1 minuto, o al momento con **App incidencias → Aplicar cambios de configuración ya**.
- **Días fijos** (`_Monitores`): marca L M X J V de cada monitor.
- **Festivos** (`_Festivos`): ya cargados del calendario; puedes añadir o quitar filas.
- **Cerrar un mes** (`_Ajustes` → `MESES_CERRADOS`): p. ej. `Octubre; Noviembre`. Ya no se podrá cambiar desde la app.
- **Enlace directo para cada responsable:** `https://mikelop1991.github.io/HojaIncidencias/?c=CODIGO`

## En la app

- Arriba, junto al nombre: **● Conectado al Excel** (verde) = todo guardado. Azul = enviando. Rojo = sin conexión (los cambios se guardan en el móvil y se envían solos al volver la cobertura). Al tocarlo, se fuerza la sincronización.
- La primera vez sale un aviso para **instalar la app** (botón directo en Android; en iPhone explica Compartir → Añadir a pantalla de inicio). También desde el menú ☰ → Instalar.
- En ordenador los monitores salen en cuadrícula y la ficha se abre como panel lateral (Esc para cerrar).
