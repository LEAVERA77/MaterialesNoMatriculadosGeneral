# Puesta en marcha del backend propio (LEAVERA77/MaterialesNoMatriculadosGeneral)

Este repositorio es **independiente** de `LEAVERA77/MaterialesNoMatriculados` (producción).
Nada de lo que hagas acá modifica producción, y nada de producción hace falta tocar.

Los archivos ya están commiteados en `main`. Lo que sigue son los pasos que se hacen
**una sola vez** en GitHub y en Google, y que no se pueden automatizar desde el repositorio.

---

## 1. GitHub — permisos de las Actions (2 minutos)

En <https://github.com/LEAVERA77/MaterialesNoMatriculadosGeneral>:

| Dónde | Qué | Por qué |
|---|---|---|
| **Settings → Actions → General → Workflow permissions** | Marcar **Read and write permissions** | Sin esto, `actualizar.yml` no puede commitear el Excel |
| **Settings → Actions → General** | Que las Actions estén habilitadas | Para que el `repository_dispatch` dispare el workflow |
| **Settings → Pages → Source** | **GitHub Actions** | Publica `index.html` en `https://leavera77.github.io/MaterialesNoMatriculadosGeneral/` |

Verificación rápida: pestaña **Actions** → tiene que aparecer el workflow
**“Actualizar Materiales”** y correr **“Deploy to GitHub Pages”** después del último push.

---

## 2. Google — Apps Script NUEVO (5 minutos)

> **No reutilices** el proyecto de Apps Script de producción: si lo editás, producción cambia.

1. Entrá a <https://script.google.com> → **Nuevo proyecto**. Ponele un nombre claro,
   por ejemplo `ENERSA Materiales · General`.
2. Reemplazá todo el contenido de `Código.gs` por el de **`apps-script/Code.gs`** de este repo
   y completá las constantes:

   ```js
   const GH_OWNER = 'LEAVERA77';
   const GH_REPO  = 'MaterialesNoMatriculadosGeneral';   // ya viene así, no cambiar
   const GH_TOKEN = '...';           // PAT fine-grained: Contents Read and write
   const SPREADSHEET_ID = '...';     // opcional (login en servidor / respaldo en Sheets)
   ```

3. **Implementar → Nueva implementación → Aplicación web**:
   * *Ejecutar como:* **Yo**
   * *Quién tiene acceso:* **Cualquier persona**
   * **Implementar** y autorizar los permisos que pida.
4. Copiá la URL que termina en **`/exec`**.

### Token para el Apps Script

GitHub → *Settings → Developer settings → Personal access tokens → Fine-grained tokens →
Generate new token*:

* **Resource owner:** `LEAVERA77`
* **Repository access:** *Only select repositories* → **solo** `MaterialesNoMatriculadosGeneral`
* **Permissions → Repository permissions → Contents:** *Read and write*
* Expiración corta (7 días) mientras se prueba; después renovalo por uno más largo o
  revocalo cuando corresponda.

> Si el token tiene acceso también a producción, el aislamiento queda dependiendo solo de la
> línea `GH_REPO`. Mejor: dale acceso únicamente a este repositorio.

---

## 3. Pegar la URL del backend en el sitio (1 minuto)

En `index.html`, sección **1. CONFIGURACIÓN**, reemplazá:

```js
const urlGoogle = "PEGAR_AQUI_URL_DEL_APPS_SCRIPT_PROPIO";
```

por la URL `/exec` del paso 2.4:

```js
const urlGoogle = "https://script.google.com/macros/s/AKfycb.../exec";
```

Commit y push a `main`. GitHub Pages republica solo.

> Mientras diga `PEGAR_AQUI...`, el sitio **no envía nada** y muestra el aviso
> “Backend propio sin configurar”. Es intencional: así es imposible que escriba en producción.

---

## 4. Prueba de humo (5 minutos)

1. Abrir <https://leavera77.github.io/MaterialesNoMatriculadosGeneral/>.
2. Debe aparecer **“Identificación de empleado”**. Entrar con un legajo de prueba.
3. Hacer un **retiro** de 1 unidad de un material con stock.
4. En el repo: **Actions → Actualizar Materiales** en verde y un commit
   `Movimiento: retiro · Legajo ... · ...`.
5. En el Excel de **este** repo: fila nueva en `Historial` con legajo, dispositivo y huella.
6. **Control de aislamiento:** el Excel y los commits de `MaterialesNoMatriculados`
   (producción) tienen que quedar **sin cambios**.

---

## 5. Si algo falla

| Síntoma | Causa probable | Solución |
|---|---|---|
| Aviso “Backend propio sin configurar” | Falta la URL del paso 3 | Pegar la URL `/exec` en `index.html` |
| `ok:false` / 401 · 403 en los logs del Apps Script | `GH_TOKEN` vacío o sin permisos | Regenerar el PAT con *Contents: Read and write* en este repo |
| La app dice “registrado” pero el Excel no cambia | Action deshabilitada o sin *Read and write* | Revisar el paso 1 y la pestaña Actions |
| El sitio muestra la versión vieja | Caché de Pages/CDN | Esperar 1–2 min y recargar con `Ctrl+F5` |
| El stock no se refresca al instante | Caché de `raw.githubusercontent` | Normal: tarda 1–3 min (igual que en producción) |

---

## Nota sobre pruebas locales

`index.html` es autocontenido: se puede abrir directamente en el navegador. Sin la URL del
backend propio, permite navegar y ver stock/historial (lectura), pero **no** registrar
movimientos ni crear PIN. Eso es lo esperado.
