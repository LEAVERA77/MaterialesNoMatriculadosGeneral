# Cómo subir estos cambios al repositorio

Todos los archivos ya están modificados, probados y commiteados localmente en la carpeta `repo/`.
Falta publicarlos en GitHub. Hay dos caminos.

---

## Opción A — Me autorizás un token y lo subo yo (2 minutos)

1. Entrá a **GitHub → tu foto → Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token**.
2. Configurá:
   * **Token name:** `arena-materiales` (o cualquiera)
   * **Expiration:** 7 días (lo revocás cuando terminemos)
   * **Resource owner:** `LEAVERA77`
   * **Repository access:** *Only select repositories* → `MaterialesNoMatriculados`
   * **Permissions → Repository permissions → Contents:** `Read and write`
   * Nada más. No necesita workflows, ni admin, ni pages.
3. Generá el token y **pegalo acá en el chat** (o mejor: cargalo en la variable que te indique).
4. Yo hago el push, verifico que la Action corra y te aviso cuando el sitio esté publicado.
5. **Revocá el token** al terminar: mismo menú → el token → *Delete*.

> Si tu organización tiene SSO, puede pedirte “authorize” después de crear el token.

## Opción B — Los subís vos (sin darme ningún acceso)

Ya te dejé todo empaquetado:

* **`MaterialesNoMatriculados_v2.zip`** → los 7 archivos con la estructura de carpetas lista para arrastrar.
* **`actualizacion_legajo_dispositivo.patch`** → para aplicar por consola:
  ```bash
  git clone https://github.com/LEAVERA77/MaterialesNoMatriculados
  cd MaterialesNoMatriculados
  git am /ruta/actualizacion_legajo_dispositivo.patch
  git push
  ```

Con el ZIP, en GitHub: **Add file → Upload files**, arrastrás `index.html`, `listado_materiales_ENERSA.xlsx` y `README.md`, y además creás las carpetas `scripts/`, `apps-script/` y `.github/workflows/` subiendo sus archivos respectivos (GitHub respeta las rutas del ZIP si lo arrastrás completo). Commit a `main`.

---

## Verificaciones que hay que hacer en GitHub (no requieren darme acceso)

| Dónde | Qué | Por qué |
|---|---|---|
| **Settings → Actions → General → Workflow permissions** | Marcar **Read and write permissions** | Sin esto la Action no puede commitear el Excel actualizado |
| **Settings → Pages → Source** | **GitHub Actions** | Publica `index.html`; el `deploy.yml` viejo intentaba un build de Node que ya no existe |
| **Settings → Actions → General** | Que las Actions estén habilitadas | Para que corra la actualización de stock |

## Después de subir

1. **Actions → Deploy to GitHub Pages** en verde y **Actualizar Materiales** disponible.
2. Abrir el sitio: tiene que aparecer la pantalla **“Identificación de empleado”**.
3. Probar con un legajo real y PIN = últimos 4 dígitos del legajo.
4. Actualizar el **Google Apps Script** con `apps-script/Code.gs` para que el legajo y la huella del dispositivo empiecen a guardarse en el Excel (sin ese paso, todo funciona igual, pero las columnas G·H·I quedan vacías).
