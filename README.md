# Materiales No Matriculados · ENERSA María Grande

Aplicación web (HTML estático, GitHub Pages) para registrar **retiros** y **devoluciones** de material y herramientas, con stock actualizado automáticamente en `listado_materiales_ENERSA.xlsx`.

> **Novedad (v2):** los empleados se identifican **por legajo + PIN** antes de operar, y **cada movimiento queda asociado a la huella del dispositivo** desde el que se hizo.

> **Novedad (v3 · este repositorio):** `MaterialesNoMatriculadosGeneral` tiene **backend propio**.
> Ya no usa el backend de `LEAVERA77/MaterialesNoMatriculados` (producción), que **no se toca**.
> Ver **§0**.

---

## 0. Backend propio: este repositorio es independiente de producción

Este repositorio (`LEAVERA77/MaterialesNoMatriculadosGeneral`) tiene **su propio backend** y
**no comparte nada** con `LEAVERA77/MaterialesNoMatriculados` (producción):

| Pieza | Producción (`MaterialesNoMatriculados`) | Este repo (`MaterialesNoMatriculadosGeneral`) |
|---|---|---|
| Excel de stock | su `listado_materiales_ENERSA.xlsx` | **su** `listado_materiales_ENERSA.xlsx`, independiente |
| GitHub Action `actualizar.yml` | la suya | una copia propia (`concurrency` independiente) |
| Repo destino del `repository_dispatch` (`GH_REPO` en `Code.gs`) | `MaterialesNoMatriculados` | **`MaterialesNoMatriculadosGeneral`** |
| Apps Script (Web App `/exec`) | proyecto de producción | **proyecto NUEVO y aparte** (URL distinta) |
| Sitio publicado | `leavera77.github.io/MaterialesNoMatriculados/` | `leavera77.github.io/MaterialesNoMatriculadosGeneral/` |

Lo único compartido es la **copia inicial de datos** (legajos, materiales, historial) y las
reglas de negocio del código. De acá en adelante evolucionan por separado: un movimiento
hecho en este sitio **solo** modifica el Excel de este repositorio.

### Qué falta para que el backend propio quede operativo (una sola vez)

1. **Workflow permissions** — *Settings → Actions → General → Workflow permissions*:
   **Read and write** (sin esto la Action no puede commitear el Excel de este repo).
2. **Pages** — *Settings → Pages → Source*: **GitHub Actions**.
3. **Apps Script nuevo** (no reutilizar el de producción):
   1. Creá un proyecto **nuevo** en <https://script.google.com>.
   2. Pegá `apps-script/Code.gs` y completá:
      ```js
      const GH_OWNER = 'LEAVERA77';
      const GH_REPO  = 'MaterialesNoMatriculadosGeneral';   // ya viene así
      const GH_TOKEN = '...';        // PAT fine-grained: Contents Read/Write SOLO en este repo
      const SPREADSHEET_ID = '...';  // opcional (validación de login / respaldo)
      ```
   3. **Implementar → Nueva implementación → Aplicación web** → *Ejecutar como* **Yo**,
      *Quién tiene acceso* **Cualquier persona**. Copiá la URL terminada en `/exec`.
4. **Pegar la URL en el HTML**: en `index.html`, reemplazá
   `"PEGAR_AQUI_URL_DEL_APPS_SCRIPT_PROPIO"` por la URL `/exec` del paso 3.

> Mientras el HTML siga con `PEGAR_AQUI...`, la app **no envía nada**: muestra el aviso
> *"Backend propio sin configurar"*. Es a propósito, para que nunca escriba en producción.
> La lectura de stock e historial sí funciona, porque lee el Excel de **este** repositorio.

> El token del Apps Script conviene que tenga acceso **solo** a `MaterialesNoMatriculadosGeneral`.
> Si también se lo habilita para producción, el aislamiento queda dependiendo de una única
> línea (`GH_REPO` en `Code.gs`).

---

## 1. Cómo funciona ahora

```
Empleado abre la página
        │
        ▼
[ Identificación ]  legajo + PIN  ──►  válido? ──no──► vuelve a pedir (bloqueo 1 min tras 5 intentos)
        │ sí
        ▼
Formulario habilitado (Retirar / Devolver)  +  barra de sesión con nombre, legajo y dispositivo
        │
        ▼
POST al Apps Script  ──►  GitHub Action  ──►  Excel actualizado (stock + historial + hoja Dispositivos)
```

* **Sin identificarse no se puede operar**: los botones `Retirar` y `Devolver` arrancan bloqueados.
* La sesión dura **12 horas** y se guarda por pestaña (`sessionStorage`). El botón **Cambiar** cierra la sesión.
* Consultar **Stock** sigue pidiendo la contraseña de siempre. Ver **Historial** es libre.

### Legajo + PIN

* El empleado ingresa su **legajo** (lista incluida, 25 empleados) y su **PIN personal**.
* **PIN inicial = los últimos 4 dígitos del legajo** (ej.: legajo `10246` → PIN `0246`).
* Tras 5 intentos fallidos el acceso se bloquea 60 segundos.
* Los PIN se guardan como **hash SHA-256** en el código (nunca el número en texto).

### Huella / ID del dispositivo

Con cada operación se envía y se guarda:

| Dato | Ejemplo | Para qué sirve |
|---|---|---|
| `dispositivo_id` | `DV-184466A654CB` | Identificador **persistente** del navegador (se genera la primera vez y queda en `localStorage`). Sirve para reconocer el mismo equipo en operaciones futuras. |
| `dispositivo_huella` | `0ba26e9e…a51a5c` | **Hash SHA-256** de las características del equipo (SO, navegador, pantalla, zona horaria, idioma, núcleos, memoria, user-agent). No depende del tamaño de la ventana ni de la sesión. |
| `dispositivo_resumen` | `Android 13 · Chrome 120 · 390x844` | Lectura humana del equipo. |
| `timestamp_cliente` / `zona_horaria` | `2026-11-19T12:15:00Z` / `America/Argentina/Cordoba` | Momento exacto del lado del empleado. |

En pantalla, el chip **`DV-…·huella`** de la barra de sesión (y el login) abre el detalle completo. El comprobante impreso también muestra el ID y la huella.

> Aclaración honesta: la huella identifica **el navegador/equipo**, no a la persona. Es una huella de dispositivo (no biometría) y en celulares suele cambiar si se reinstala el navegador o se borran los datos del sitio; el `dispositivo_id` se pierde en ese caso, pero la huella (hash de hardware/SO) normalmente se mantiene.

---

## 2. Dónde queda registrado

### Hoja `Historial` (columnas nuevas al final, G·H·I)

| A Fecha | B Usuario | C Operacion | D Material | E Cantidad | F Uso | **G Legajo** | **H Dispositivo** | **I Huella** |
|---|---|---|---|---|---|---|---|---|

Los registros viejos siguen ahí; sus últimas columnas quedan vacías. Nada se rompe.

### Hoja `Dispositivos` (nueva, se crea sola)

`ID · Huella · Detalle · Último usuario · Último legajo · Última operación · Operaciones`

Sirve para auditar: *qué equipos operaron, quién los usó por última vez y cuántas operaciones hicieron*.

### Hoja `Empleados` (nueva)

`LEGAJO · APELLIDO Y NOMBRE · PIN · ACTIVO` + notas en la columna F.

* Podés **cambiar un PIN** escribiéndolo en la columna C (texto o hash SHA-256).
* `ACTIVO = NO` bloquea ese legajo.
* Podés **agregar empleados**: la app lee esta hoja al iniciar y suma/actualiza legajos sin tocar el código.
* La GitHub Action también valida el legajo contra esta hoja: un legajo inexistente o inactivo **no modifica el stock**.

---

## 3. Instalación (3 pasos)

### Paso 1 — Archivos en el repositorio (este repo: `MaterialesNoMatriculadosGeneral`)

Subir a la rama `main`:

```
index.html                                  (reemplaza el actual)
listado_materiales_ENERSA.xlsx              (agrega las hojas Empleados y Dispositivos)
scripts/actualizar_stock.py                 (nuevo)
apps-script/Code.gs                         (nuevo, referencia)
.github/workflows/actualizar.yml            (actualizado)
.github/workflows/deploy.yml                (corregido: publica el sitio estático)
README.md                                   (este archivo)
```

> **Importante:** `index.html` y el `.xlsx` deben subirse **juntos**. Si subís solo el HTML, los legajos igual funcionan (la lista está embebida), pero la hoja `Empleados` no se podrá administrar desde el Excel.

### Paso 2 — Google Apps Script

Abrir el proyecto de Apps Script del Web App, reemplazar el código por `apps-script/Code.gs` y completar:

```js
const GH_OWNER = 'LEAVERA77';
const GH_REPO  = 'MaterialesNoMatriculadosGeneral';  // backend propio (producción NO se toca)
const GH_TOKEN = '...';        // PAT con "Contents: Read and write" SOLO en este repo
const SPREADSHEET_ID = '...';  // opcional: planilla con la hoja Empleados
```

Luego **Implementar → Administrar implementaciones → ✏️ → Nueva versión**. La URL `/exec` no cambia, así que no hay que tocar el HTML.

> Si preferís no tocar el Apps Script ahora, **todo sigue funcionando igual**: la app envía los campos nuevos y el script viejo los ignora (el Excel se actualiza sin legajo ni dispositivo). Pero para que la trazabilidad quede guardada en el Excel, hay que hacer este paso.

### ⚠️ Límite de GitHub: 10 propiedades en el payload

`repository_dispatch` **rechaza con error 422** cualquier `client_payload` con más de 10 propiedades.
Por eso `apps-script/Code.gs` envía los datos del equipo **agrupados** en un objeto `dispositivo`
y `scripts/actualizar_stock.py` acepta los dos formatos (plano `dispositivo_id` o anidado `dispositivo.id`).
Si agregás campos nuevos, no pases de 10 claves de primer nivel.

### Paso 3 — Verificar

1. GitHub → **Settings → Pages** → Source: *GitHub Actions* (el workflow de deploy corregido publica `index.html`).
2. Abrir la URL del sitio: debe aparecer la pantalla **“Identificación de empleado”**.
3. Probar con un legajo real y su PIN (últimos 4 dígitos).
4. Hacer un retiro de prueba y revisar el Excel: fila en `Historial` con legajo, ID de dispositivo y huella.

### Ojo con la caché de `raw.githubusercontent`

Después de un movimiento, el Excel tarda **1 a 3 minutos** en reflejarse en el sitio (caché de GitHub).
El historial del día puede mostrar el último movimiento con algo de retraso; no es un error de la app.

### Validación del lado del servidor (opcional, recomendada)

En `index.html`:

```js
const USAR_VALIDACION_SERVIDOR = false;  // → true después de actualizar el Apps Script
```

Hoy el PIN se valida en el navegador (rápido y funciona sin conexión). Con el Apps Script actualizado y esta opción en `true`, además se valida en Google, y la hoja `Empleados` pasa a ser la autoridad.

---

## 4. Seguridad: qué protege y qué no

* ✅ Impide que un empleado cargue movimientos **sin identificarse con un legajo válido**.
* ✅ Todo movimiento queda con **legajo + dispositivo + huella + fecha/hora**, y el legajo se revalida en la GitHub Action.
* ✅ Los PIN no viajan ni se guardan en texto en el código (hash SHA-256).
* ⚠️ La validación del PIN en el navegador es una **barrera de uso**, no una autenticación fuerte: alguien con conocimientos técnicos y el archivo en mano podría saltearla. La defensa real es la del punto 2 del Paso 3 (validación en servidor) + la revisión del `Historial`.
* ⚠️ El repositorio es público (lo necesita GitHub Pages), así que el código y la lista de nombres son visibles. Los PIN en la hoja `Empleados` también: por eso conviene mantenerlos como hash o cambiar los iniciales.
* ⚠️ La huella de dispositivo no es infalible: dos equipos iguales pueden parecerse y un mismo equipo puede cambiar si se actualiza el navegador. Es evidencia de auditoría, no una prueba criptográfica de identidad.

---

## 5. Accesos a autorizar

| # | Acceso | Para qué | Estado |
|---|---|---|---|
| 1 | **Token de GitHub (fine-grained PAT)** con `Contents: Read and write` **solo** en `LEAVERA77/MaterialesNoMatriculadosGeneral` | Para que pueda subir los cambios al repo (hoy no tengo credenciales) | ⬜ pendiente |
| 2 | **GitHub Settings → Actions → General → Workflow permissions**: *Read and write* | Para que la Action pueda commitear el Excel actualizado | ⚠️ verificar |
| 3 | **GitHub Settings → Pages** → Source: *GitHub Actions* | Para publicar el sitio | ⚠️ verificar |
| 4 | **Google Apps Script**: `GH_TOKEN` (PAT con `Contents: Read and write`) y permiso de implementación | Para que el script dispare la Action y (opcional) valide el login | ⬜ a completar |
| 5 | (Opcional) ID de la planilla de Google | Solo si querés respaldo en Google Sheets y validación en servidor | ⬜ opcional |

El token del punto 1 lo generás en **GitHub → Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token** (expiración corta, 7 días) y lo revocás cuando terminemos. No hace falta que lo pegues en el chat: alcanza con que lo cargues donde te indique.

---

## 6. Estructura

```
index.html                          App completa (login por legajo + huella de dispositivo)
listado_materiales_ENERSA.xlsx      Materiales · Historial · Empleados
scripts/actualizar_stock.py         Lógica de la GitHub Action (stock + historial + dispositivos)
apps-script/Code.gs                 Web App de Google (recibe el POST y dispara la Action)
.github/workflows/actualizar.yml    Action que actualiza el Excel
.github/workflows/deploy.yml        Action que publica el sitio
```

## 7. Cambios respecto de la versión anterior

* Se eliminó el `<select>` de usuarios: ahora el empleado se identifica por **legajo** y el nombre se completa solo.
* Se agregó el bloqueo de operaciones hasta validar identidad y el manejo de sesión (12 h, botón *Cambiar*).
* Los payloads ahora incluyen `legajo` y los campos `dispositivo_*` (retrocompatible con el Apps Script viejo).
* `Historial` muestra las columnas **Legajo** y **Dispositivo**; el comprobante también.
* El Excel suma las hojas **Empleados** y **Dispositivos**.
* La GitHub Action valida legajo y stock, y escribe las columnas nuevas.
* Se corrigió `deploy.yml` (fallaba con `npm ci`: el proyecto Next ya no existe).
* Se eliminó un script de Cloudflare que había quedado pegado al final del HTML y no cumplía ninguna función.
