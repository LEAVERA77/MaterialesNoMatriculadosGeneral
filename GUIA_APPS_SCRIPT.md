> **Guía para armar el Apps Script propio de `MaterialesNoMatriculadosGeneral`.**
> Tiempo estimado: **15 minutos**. No hace falta saber programar: son copiar, pegar y tocar botones.
> Regla de oro: **el proyecto de Apps Script de producción no se abre ni se toca.**

---

# Guía: backend propio con Google Apps Script

## Antes de empezar — 3 cosas para chequear

| # | Chequeo | Por qué |
|---|---|---|
| 1 | Entrá a <https://script.google.com> con la **misma cuenta de Google** que usás siempre para esto. | El script queda en esa cuenta; si usás otra, después no lo encontrás. |
| 2 | **No abras** el proyecto del script de producción. | Un proyecto = un código = una URL. Editarlo cambia producción al instante. |
| 3 | Tené a mano el repo <https://github.com/LEAVERA77/MaterialesNoMatriculadosGeneral> | Ahí está el archivo que vas a copiar. |

---

## Paso 1 · Crear el token de GitHub (5 min)

Es la “llave” que el script usa para pedirle a GitHub que actualice el Excel.

1. GitHub → tu foto (arriba a la derecha) → **Settings**.
2. En la barra izquierda, bien abajo: **Developer settings**.
3. **Personal access tokens → Fine-grained tokens → Generate new token**.
4. Completá:

   | Campo | Qué poner |
   |---|---|
   | **Token name** | `enersa-general-appsscript` |
   | **Expiration** | 90 días (o lo que prefieras; anotá la fecha) |
   | **Resource owner** | `LEAVERA77` |
   | **Repository access** | **Only select repositories** → elegí **solo** `MaterialesNoMatriculadosGeneral` |
   | **Permissions → Repository permissions → Contents** | **Read and write** |

5. **Generate token** y **copiá el token ahora** (empieza con `github_pat_...`). No se vuelve a mostrar.

> Por qué “solo este repositorio”: si algún día el script quedara apuntando a producción por error,
> GitHub igual le negaría la escritura. Es una segunda cerradura.

> ¿Preferís no crear otro? Podés reusar el que ya me pasaste (vence el **23/12/2026**),
> pero ese tiene acceso a los dos repos, o sea que la protección es solo el “modo seguro” del código.

---

## Paso 2 · Crear el proyecto de Apps Script (2 min)

1. Entrá a <https://script.google.com> → **Nuevo proyecto** (arriba a la izquierda).
2. Arriba dice *Proyecto sin título*. Hacé clic y ponele: **`ENERSA Materiales · General`**.
3. Borrá **todo** lo que aparece en el editor: clic dentro del código → `Ctrl+A` → `Supr`.
   (Queda vacío, sin `function myFunction() {}`.)

---

## Paso 3 · Pegar el código (2 min)

1. Abrí el archivo del repo:
   <https://github.com/LEAVERA77/MaterialesNoMatriculadosGeneral/blob/main/apps-script/Code.gs>
2. Botón **Raw** (arriba a la derecha del cuadro del código) o el ícono de copiar.
3. Volvé a Apps Script, pegá con `Ctrl+V` en el editor.
4. Guardá: **`Ctrl+S`**. Si el nombre del proyecto muestra un asterisco, quedó sin guardar.

---

## Paso 4 · Completar la configuración (2 min)

Arriba del archivo están las 4 líneas a completar. Solo tocás **una**:

```js
const GH_OWNER = 'LEAVERA77';                          // ya está bien
const GH_REPO  = 'MaterialesNoMatriculadosGeneral';    // ya está bien — NO cambiar
const GH_TOKEN = 'PEGAR_AQUI_EL_TOKEN';                // ← pegar acá el token del Paso 1
const SPREADSHEET_ID = '';                             // ← DEJAR VACÍO
```

* **`GH_TOKEN`**: borrá `PEGAR_AQUI_EL_TOKEN` (con las comillas simples) y pegá el `github_pat_...`
  entre las comillas. Tiene que quedar así: `const GH_TOKEN = 'github_pat_11AW...';`
* **`SPREADSHEET_ID`**: dejalo **vacío** (`''`).
  ⚠️ **No pongas el ID de la planilla de producción**: eso haría que las pruebas escriban
  filas en su hoja `Historial`. Si querés respaldo en Google Sheets, creá una planilla nueva.

Guardá de nuevo con **`Ctrl+S`**.

---

## Paso 5 · Publicarlo como aplicación web (3 min)

1. Arriba a la derecha: botón azul **Implementar** → **Nueva implementación**.
2. A la izquierda del diálogo, engranaje ⚙️ **Seleccionar tipo** → **Aplicación web**.
3. Completá:

   | Campo | Qué elegir |
   |---|---|
   | **Descripción** | `backend propio General v1` |
   | **Ejecutar como** | **Yo (tu@gmail.com)** |
   | **Quién tiene acceso** | **Cualquier persona** |

   > “Cualquier persona” suena raro, pero es necesario: el navegador de los empleados no tiene
   > sesión de Google. Sin esto, el sitio no puede hablar con el script.
   > El script no queda “abierto”: solo acepta los POST del formulario, y valida legajo y stock.

4. **Implementar**.
5. La primera vez pide permisos → **Autorizar acceso** → elegí tu cuenta →
   aparece “Google no verificó esta aplicación” → **Configuración avanzada** →
   **Ir a ENERSA Materiales · General (no seguro)** → **Permitir**.
   (Ese cartel es normal en scripts propios, no es un error.)
6. Copiá la **URL de la aplicación web**: es larga y termina en **`/exec`**.
   Guardala, la usás en el Paso 7.

---

## Paso 6 · Verificar que apunta al repo correcto (1 min)

Pegá esa URL `/exec` **en una pestaña nueva del navegador** y apretá Enter.
Tenés que ver un texto JSON como este:

```json
{
  "ok": true,
  "build": "general-2026-11-23-r4",
  "servicio": "ENERSA materiales",
  "entorno": "backend propio",
  "repo_destino": "LEAVERA77/MaterialesNoMatriculadosGeneral",
  "modo_seguro": "ok",
  "token_cargado": true,
  "token_detalle": { "largo": 93, "formato": "valido", "pista": "github_pat_1…axPC", "fue_saneado": false },
  "planilla_configurada": false,
  "motivo": "",
  "detalle": ""
}
```

Checklist:

| Campo | Valor esperado | Si no coincide |
|---|---|---|
| `repo_destino` | `LEAVERA77/MaterialesNoMatriculadosGeneral` | Te quedó mal `GH_REPO` → volvé al Paso 4 |
| `build` | `general-2026-11-23-r4` | Otro valor o ausente → el `/exec` corre **código viejo** → *Nueva versión* (nota de abajo) |
| `token_cargado` | `true` | Falta pegar el token → Paso 4 y **Nueva versión** |
| `token_detalle.largo` | ~93 (`github_pat_…`) o 40 (`ghp_…`) | Si da mucho menos, el token se pegó **incompleto** |
| `token_detalle.formato` | `"valido"` | `placeholder` = no lo pegaste; `formato_inesperado` = cortado o con basura |
| `modo_seguro` | `"ok"` | Dice `BLOQUEADO (es producción)` → `GH_REPO` está en producción |

### Atajo: verificar sin salir del editor

En el editor de Apps Script, arriba hay un desplegable con el nombre de la función. Elegí **`verificar`**
y apretá **▶ Ejecutar**. Abajo, en *Registro de ejecuciones*, aparece el mismo JSON.
Sirve para revisar la configuración antes de siquiera publicar.

> **Muy importante:** si cambiás el código después de implementar, el cambio **no llega solo** a la URL.
> Hay que ir a **Implementar → Administrar implementaciones → ✏️ (lápiz) → Versión: Nueva versión → Implementar**.
> La URL `/exec` **no cambia**. Sin este paso seguís viendo el código viejo (y el JSON te lo delata).

---

## Paso 7 · Conectar el sitio con este script (1 min)

En el repo, abrí `index.html` y buscá esta línea (está en la sección *1. CONFIGURACIÓN*, cerca de la línea 769):

```js
const urlGoogle = "PEGAR_AQUI_URL_DEL_APPS_SCRIPT_PROPIO";
```

Reemplazala por tu URL `/exec`:

```js
const urlGoogle = "https://script.google.com/macros/s/AKfycb.../exec";
```

Después, una de estas dos:

* **Opción rápida (web de GitHub):** <https://github.com/LEAVERA77/MaterialesNoMatriculadosGeneral/edit/main/index.html>
  → pegás el cambio → **Commit changes** → *Commit directly to the `main` branch*.
* **Opción cómoda:** pegame la URL `/exec` acá en el chat y lo commiteo yo, con la prueba de humo incluida.

Esperá 1–2 minutos (GitHub Pages republica solo) y recargá con `Ctrl+F5`.

---

## Paso 8 · Prueba de humo (5 min)

En <https://leavera77.github.io/MaterialesNoMatriculadosGeneral/>:

1. **Login** con un legajo de prueba. Si es un legajo sin PIN, te va a pedir crearlo
   (mínimo 4 dígitos) — eso ya es una escritura real: va a la hoja `Empleados` de **este** repo.
2. **Retiro** de **1 unidad** de un material que tenga stock.
3. Esperá ~1 minuto y mirá:
   * Repo → **Actions** → tiene que aparecer **Actualizar Materiales** en verde.
   * El Excel del repo → hoja `Historial` con una fila nueva: legajo, `DV-...` y huella.
4. **Control de aislamiento (el más importante):**
   <https://github.com/LEAVERA77/MaterialesNoMatriculados/commits/main>
   **No** debe aparecer ningún commit nuevo de hoy. Si aparece, pará y avisame.

Después de probar, si querés dejar el stock como estaba, poné a mano la cantidad en la hoja
`Materiales` (editar el `.xlsx` en GitHub no es cómodo: mejor pedirme que prepare un commit de reversión).

---

## Si algo falla

| Síntoma en pantalla | Qué significa | Solución |
|---|---|---|
| “Backend propio sin configurar…” | El HTML todavía tiene `PEGAR_AQUI...` | Paso 7 |
| “No se pudo guardar el PIN (payload_invalido)” | El script desplegado es la versión vieja | *Nueva versión* (nota del Paso 6) |
| “No se pudo guardar el PIN (modo_seguro)” | `GH_REPO` quedó en producción | Paso 4 + *Nueva versión* |
| “No se pudo guardar el PIN (falta_token)” | Token vacío o mal pegado | Paso 4 + *Nueva versión* |
| “No se pudo conectar con el servidor” | URL mal copiada, o el acceso no es “Cualquier persona” | Revisar Paso 5 y la URL |
| El JSON del Paso 6 no aparece / da error | La implementación no está publicada como *Aplicación web* | Paso 5 otra vez |
| `token_cargado: true` **y** `detalle: "Falta cargar GH_TOKEN"` a la vez | Es imposible con el código del repo: el `/exec` está corriendo **otra versión** (editada a mano, o un archivo de más en el proyecto) | Ver recuadro de abajo |
| `build` distinto al esperado | El deploy quedó viejo | *Nueva versión* |
| El sitio dice “registrado” pero el Excel no cambia | Falla la Action de GitHub | Repo → **Actions** → abrir el run en rojo → **Settings → Actions → General → Workflow permissions: Read and write** |
| El stock tarda en reflejarse | Caché de `raw.githubusercontent` | Normal: 1–3 min |

---

## Si el diagnóstico se contradice o da raro

Síntoma típico: `"token_cargado": true` y `"detalle": "Falta cargar GH_TOKEN."` en el mismo JSON.
No puede pasar con el código del repo (las dos cosas salen de la misma variable), así que el `/exec`
**no está ejecutando este archivo**. Causas, en orden de probabilidad:

1. **El deploy quedó viejo.** Editaste el código pero no creaste la versión:
   *Implementar → Administrar implementaciones → ✏️ → Versión: **Nueva versión** → Implementar*.
2. **Hay otro archivo en el proyecto.** En el panel izquierdo *Archivos* tiene que haber **uno solo**
   (`Código.gs` o `Code.gs`). Si hay dos, el segundo puede redefinir funciones sin avisar:
   borralo y volvé a pegar el contenido del repo.
3. **El pegado quedó incompleto** (se cortó a la mitad, o quedó el `function myFunction() {}`).
   Solución: clic en el editor → `Ctrl+A` → `Supr` → pegar de nuevo → `Ctrl+S`.

El sello `build` del JSON y el `token_detalle` existen justamente para distinguir estos casos de un vistazo.

---

## Cómo queda la separación

| | Producción | Este entorno |
|---|---|---|
| Proyecto Apps Script | el de siempre (no se toca) | **nuevo**, `ENERSA Materiales · General` |
| URL `/exec` | la de siempre | **distinta**, la que acabás de crear |
| `GH_REPO` en el script | `MaterialesNoMatriculados` | `MaterialesNoMatriculadosGeneral` |
| Repo que se modifica | producción | pruebas |
| Excel | el de producción | el de este repo |
| Token de GitHub | el suyo | uno **nuevo**, con acceso solo a este repo |
| Planilla de respaldo | la de producción | vacía (`''`) |

Si querés volver todo atrás: en `index.html` volvés a poner `PEGAR_AQUI_URL_DEL_APPS_SCRIPT_PROPIO`,
y en Apps Script borrás la implementación. Producción nunca se enteró.
