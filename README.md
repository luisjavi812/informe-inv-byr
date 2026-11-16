# BYR · Depuración semanal de inventario

Aplicación web para comparar el inventario del **sistema** (PDF) contra el **conteo físico** (Excel), decidir qué placas **eliminar**, **incluir** o dejar como **OK**, detectar anomalías, excluir los ingresos posteriores a la fecha de corte y guardar el histórico de cada corte.

Hecha para Representaciones BYR SAS. Funciona sin build step y sin servidor propio: HTML, CSS y JavaScript puro.

---

## Qué hace

Al cargar los archivos la app pide la **fecha de corte** del conteo físico. Las placas de la hoja de ingresos con fecha posterior al corte **no se les hace nada** (ya están en el sistema) y se listan en la vista **Excluidas**.

| Estado | Significado | Acción |
| --- | --- | --- |
| **ELIMINAR** | Está en el sistema y no aparece en el conteo físico. | Quitar la placa del sistema. Fila roja con texto blanco. |
| **INCLUIR** | Está en el físico y no aparece en el PDF del sistema. | Agregar la placa al sistema. Fila verde con texto oscuro. |
| **OK** | Está en ambos lados. Las categorías del sistema y los talleres físicos **no se comparan entre sí**. | Sin acción. |
| **EXCLUIDA** | Ingresada después de la fecha de corte. | Sin acción; solo se muestra en «Excluidas». |

Además:

- Una sección por categoría/taller/parqueadero, con contadores por estado y tabla desplegable.
- Alternar agrupación: **por categoría del sistema** o **por taller físico** (las placas del PARQUEO se agrupan por la categoría con la que aparecen en el sistema, o «PARQUEO (sin categoría)»; los títulos que son solo «TALLER» se agrupan como «TALLER (sin nombre)»).
- Filtros por taller, tipo de vehículo, estado y búsqueda de placa.
- Copiar placas (una por línea) por sección o de todo el corte, para incluir y para eliminar.
- **Diálogo de confirmación** antes de copiar o exportar la lista de ELIMINAR, con la cantidad de placas.
- Exporta el resultado filtrado a **CSV y Excel** (`.xlsx`) y las anomalías a CSV (`;` + BOM UTF-8, compatible con Excel en español).
- Anomalías: duplicados en cada lado, placas **casi idénticas** (distancia de edición 1, ej. QZS756/QSZ756), códigos R repetidos, R que apunta a otra placa, R basura, filas sin placa válida y **placas de prueba con lista editable** en Ajustes.
- Guarda un histórico de cortes y compara contra el corte anterior: pendientes que persisten, resueltos y nuevos. No permite dos cortes con la misma fecha sin **reemplazar** el anterior.
- El histórico funciona en **modo local** (localStorage) o en **Supabase** con autenticación y RLS.

---

## Archivos

```
byr-inventario/
├── index.html          # estructura de la interfaz
├── styles.css          # diseño responsive (lavanda + azul)
├── app.js              # parseo, comparación, render, historial
├── supabase/
│   └── schema.sql      # tablas, índices, RLS y permisos
└── README.md
```

Los archivos de entrada no se suben a ningún servidor: el PDF y el Excel se procesan en el navegador con `pdf.js` y `SheetJS`. Solo los cortes que decidas guardar viajan a Supabase (si está configurado).

---

## Uso rápido

1. Abre la app (o ejecútala localmente, ver abajo).
2. Pulsa **Cargar archivos** o arrastra:
   - `inventariosistema.pdf`
   - `INVENTARIO FISICO ... .xlsx`
3. Cuando el Excel trae hoja de ingresos, la app pide la **fecha de corte** (sugerencia: un día antes del ingreso más antiguo del log). Las placas ingresadas después se excluyen y se listan en **Excluidas**.
4. Revisa los contadores en **Resumen** y las placas en **Comparación**.
5. Usa **Copiar eliminar** / **Copiar incluir** / **Exportar CSV** / **Exportar Excel**.
6. Pulsa **Guardar corte** para conservar el histórico de la semana.

### Atajos de la pantalla de comparación

- **Copiar eliminar (N):** todas las placas a eliminar que pasan los filtros actuales (pide confirmación).
- **Copiar incluir (N):** todas las placas a incluir que pasan los filtros actuales.
- **Exportar CSV / Exportar Excel:** comparativo completo filtrado con estado, categorías, tipo, R y observación (pide confirmación si incluye eliminar).
- El botón de copiar dentro de cada sección copia solo ese taller; el icono con signo menos copia solo sus placas a eliminar (también con confirmación).

---

## Formatos aceptados

### PDF del sistema

Se espera el reporte agrupado por taller/parqueadero, con títulos como:

```
TALLER SENCILLO B&R ( 89 )
PLACA        No. R        VERIFICADO
SZY653       N/A          _____
...
```

El parser:

- Reconoce títulos con su conteo declarado `( n )` y las placas que vienen debajo.
- Ignora encabezado de empresa, NIT, TEL, IVA, dirección, fecha, separadores, «VERIFICADO POR», «TOTAL VEHICULOS» y las líneas `NIT :` / `TEL :` que se cuelan en medio de las listas (aunque vengan pegadas a una placa).
- Acepta filas sin código R (ej. `SPO334 _____`).
- Valida el conteo de cada bloque contra el `( n )` del título y el total contra `TOTAL VEHICULOS : 760`; si no cuadra, muestra advertencia.
- Conserva placas con formatos irregulares (`TEK`, `SCHRA`, `CHASISSINPLATA`, etc.) para que el total no se pierda; esas placas aparecen en la anomalía **Filas sin placa válida**.
- Normaliza a mayúsculas y quita espacios, guiones y puntos antes de comparar.

> Un PDF escaneado como imagen no se puede leer sin OCR. Usa el PDF exportado por el sistema.

### Excel del conteo físico

Se aceptan `.xlsx`, `.xls` y `.csv`. Las hojas se clasifican **por contenido**:

1. **Hoja de ingresos:** se detecta porque sus filas traen fecha/hora (el nombre de la hoja cambia cada semana y no se usa). De ella salen las placas que se **excluyen** si son posteriores a la fecha de corte.
2. **Bloques con título** (hoja `TALLERES`): el nombre del taller está en una celda y las placas debajo, en una o varias columnas. Los títulos que son solo «TALLER» se agrupan como **«TALLER (sin nombre)»**.
3. **Hoja sin títulos** (`PARQUEO`): cuadrícula de placas. Cada placa toma como parqueadero la categoría con la que aparece en el sistema, o «PARQUEO (sin categoría)» si no existe allí.
4. **Tabla con encabezados:** una fila con `PLACA`, `TALLER`, `R`, `TIPO`. El orden y las mayúsculas no importan.

---

## Estados y seguridad

- Cada fila muestra dos columnas: **Categoría / Parqueadero (sistema)** y **Taller físico**.
- Antes de **Copiar** o **Exportar** la lista de ELIMINAR aparece un diálogo de confirmación con la cantidad de placas.
- Accesibilidad: el rojo lleva texto blanco con contraste AA y cada estado va indicado con etiqueta e icono, no solo con color.

---

## Anomalías

En **Ajustes → Placas de prueba** puedes editar la lista separada por comas (se guarda en el navegador).

---

## Historial

### Modo local

Sin credenciales de Supabase, el botón **Guardar corte** guarda la cabecera y el detalle en `localStorage`. El histórico vive en ese navegador y ese equipo. Sirve para probar la app o para uso individual.

### Supabase (recomendado para el equipo)

1. Crea un proyecto en [supabase.com](https://supabase.com).
2. Abre **SQL Editor**, pega el contenido de [`supabase/schema.sql`](supabase/schema.sql) y ejecútalo.
3. En **Authentication → Providers**, deja habilitado **Email**. Para un equipo interno puedes desactivar **Confirm email** (Authentication → Settings) y así el alta es inmediata.
4. En **Authentication → URL Configuration**:
   - **Site URL:** la URL de producción en Vercel (por ejemplo `https://byr-inventario.vercel.app`).
   - **Redirect URLs:** agrega también `http://127.0.0.1:8765/**` si vas a probar en local. Sin esto, los enlaces mágicos y la confirmación por correo no vuelven a la app.
5. En **Project Settings → API**, copia:
   - **Project URL**
   - **anon public key**
6. En la app, entra a **Ajustes → Conexión Supabase**, pega ambos valores y pulsa **Guardar y conectar**.
7. Ve a **Historial → Crear cuenta** o **Iniciar sesión**.
8. Guarda un corte. Al seleccionarlo verás el comparativo contra el corte anterior.

La llave `anon` es pública: la seguridad real la dan las políticas RLS del esquema. Cada usuario solo puede leer, insertar y borrar sus propios cortes (`auth.uid() = user_id`). El detalle se borra en cascada al borrar la cabecera.

#### Credenciales al desplegar

Puedes dejar los campos vacíos y configurarlos desde Ajustes, o inyectarlos antes de cargar `app.js`:

```html
<script>
  window.BYR_SUPABASE = {
    url: 'https://tu-proyecto.supabase.co',
    anonKey: 'eyJhbGciOi...'
  };
</script>
<script type="module" src="app.js"></script>
```

No uses la `service_role` key en el navegador.

---

## Ejecutar en local

Los módulos ES y los CDN no funcionan abriendo `index.html` con doble clic (`file://`). Levanta un servidor estático:

```powershell
cd byr-inventario
python -m http.server 8765
```

Abre `http://127.0.0.1:8765`.

Alternativas: `npx serve .`, la extensión Live Server de VS Code, etc.

---

## Publicar en GitHub + Vercel

### URLs en producción

| URL | Qué es |
| --- | --- |
| **https://problemacero.com/informe-inv-byr** | La ruta bonita (proxy del sitio principal a Vercel). |
| https://informe-inv-byr.vercel.app | El proyecto directo de Vercel (mismo contenido). |
| https://github.com/luisjavi812/informe-inv-byr | El código fuente. |
| https://informe-inv-byr.vercel.app/guia | Guía GuíaFlow del despliegue. |

### Cómo publicar un cambio

```powershell
cd byr-inventario
git add -A
git commit -m "Cambio: descripción"
git push
npx vercel --prod
```

El rewrite de `problemacero.com` apunta al proyecto de Vercel, así que con eso la ruta bonita se actualiza sola (un único repositorio fuente).

Si prefieres despliegue automático al hacer `git push`: en Vercel → Project `informe-inv-byr` → Settings → Git → conecta el repo y desactiva en Settings → Deployment Protection cualquier protección de SSO para producción.

> La ruta bonita depende de dos piezas en el repo `Portal-apps`: el rewrite en `next.config.ts` (PR #1) y la ruta pública en `src/lib/supabase/proxy.ts` (PR #2). Si algún día deja de responder, revisa que esas dos líneas sigan ahí.

---

## Detalle de las anomalías

| Anomalía | Qué busca |
| --- | --- |
| Duplicados en el físico | La misma placa aparece más de una vez en el Excel. |
| Duplicados en el sistema | La misma placa aparece más de una vez en el PDF. |
| Placas casi idénticas | Pares con distancia de edición 1: sustitución, inserción, borrado o transposición adyacente (ej. `QZS756` / `QSZ756`). |
| Placas de prueba | Lista editable en Ajustes: `ELO12345`, `OOO1234`, `AND1234`, `KAR1234`, `MENSUALIDADTALLER1` y similares. |
| Basura en columna R | Vacío, `Xxxx`, `N/V`, `N/A`, comas, valores sin código real. Se muestran tal cual. |
| Códigos R repetidos | Un mismo R asignado a varias placas. |
| R que apunta a otra placa | El R coincide con la placa de otro vehículo del inventario. |
| Filas sin placa válida | No cumplen `AAA123` ni el formato de remolque (letra + 4–6 dígitos). |

La lista de anomalías es informativa: no cambia el estado de la placa en el comparativo.

---

## Notas y límites

- **Todo se procesa en el navegador.** No hay backend propio ni carga de archivos a un servidor.
- El histórico local no se sincroniza entre equipos; para eso usa Supabase.
- El PDF debe ser texto seleccionable, no una imagen escaneada.
- Con los datos de noviembre: 760 en sistema, 273 únicas en el conteo (161 en TALLERES + 112 en PARQUEO), 15 ingresos del 16-nov excluidos, resultado **613 a eliminar y 141 a incluir**.
- Si el Excel no trae hoja de ingresos no se pide fecha de corte y no hay exclusiones.
- El botón de copiar usa el portapapeles del navegador. Si el navegador lo bloquea, usa **Exportar CSV** o **Exportar Excel**.
- El CSV se genera con separador `;` y BOM UTF-8 para que Excel en español abra las columnas correctamente.
- `pdf.js` y `SheetJS` se cargan desde CDN con versiones fijas; si la red corporativa bloquea los CDN, hay que alojar esas librerías localmente y cambiar las rutas en `app.js`.

---

## Mantenimiento

- Para cambiar colores, tipografía o espaciados: `styles.css` (variables en `:root`).
- Para cambiar reglas de normalización, detección de anomalías, exclusiones o columnas del CSV: `app.js`, secciones 2, 4 y 5.
- Para cambiar el esquema del histórico: `supabase/schema.sql` y las funciones `persistCorte`, `findCorteByFecha` y `fetchDetalles` en `app.js`.
- Si cambia el formato del PDF o del Excel, ajusta `parseSystemLines` / `parsePhysicalSheets` (la hoja de ingresos se detecta por contenido, no por nombre).

---

Representaciones BYR SAS · Depuración semanal de inventario
