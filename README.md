# Fondos de Inversión Hencorp — actualización automática (gratis, sin base de datos)

Robot que cada día lee los PDF de **Comportamiento Histórico** de los 6 fondos en
hencorpgestora.com, calcula los rendimientos y deja **siempre actualizados**:

- `data.json` — datos que consume la WebApp.
- `fondos-hencorp-rendimientos.xlsx` — el Excel (7 hojas) listo para descargar.
- `index.html` — WebApp de una página (tablas + gráficas + botón Descargar Excel).

Todo corre **gratis** en GitHub (Actions = programador; el repo = "base de datos"; Pages = hosting).
**No usa Supabase ni ningún servicio de pago.**

## Cómo dejarlo andando (una sola vez, ~10 min)

1. **Cuenta GitHub** (gratis) en https://github.com si no tenés.
2. **Creá un repositorio** nuevo (ver nota de privacidad abajo).
3. **Subí estos archivos** tal cual, respetando la carpeta `.github/workflows/`.
   (Web: *Add file → Upload files*, arrastrá todo. O por `git push`.)
4. **Activá Actions:** pestaña **Actions** → *I understand my workflows, enable them*.
5. **Corré una vez a mano:** Actions → **Actualizar fondos Hencorp** → **Run workflow**.
   En 1–2 min genera y commitea `data.json` y el `.xlsx`.
6. **Activá la WebApp (Pages):** Settings → **Pages** → Source: *Deploy from a branch* →
   rama `main`, carpeta `/ (root)` → Save. En ~1 min queda publicada en:
   `https://TU-USUARIO.github.io/TU-REPO/`
7. **Listo.** Cada día a las **05:00 (hora de El Salvador)** se actualiza solo.
   También podés ejecutarlo cuando quieras con **Run workflow**.

## Frecuencia
Los datos del sitio son **diarios** (hay valor cuota por día). El robot corre **a diario**
y es **idempotente**: si un día no hay cambios, no commitea nada. Para cambiar la hora,
editá el `cron` en `.github/workflows/update.yml` (está en UTC; El Salvador = UTC−6).

## Nota de privacidad (importante)
GitHub **Pages gratis requiere repositorio público**. Los datos publicados ya son
**públicos** (están en el sitio de Hencorp), así que no se expone nada confidencial.
Si preferís que NO sea público:
- Usá un **repo privado** (Actions sigue siendo gratis) y **omití Pages**: descargás el
  `.xlsx` directo del repo, o se puede adaptar el robot para que lo suba a tu **OneDrive**.
- (Pages desde repo privado requeriría plan pago; por eso la WebApp pública usa repo público.)

## Mantenimiento
- Si Hencorp cambia el **formato del PDF**, el parseo podría romperse. El robot está hecho
  para **no sobrescribir** si falla algún fondo (conserva el último dato bueno) y marca error
  en la pestaña Actions.
- GitHub desactiva los cron tras 60 días sin actividad; como el robot commitea, se mantiene activo.

## Correrlo en tu compu (opcional)
```bash
pip install -r requirements.txt
python scraper.py
```

## Archivos
- `scraper.py` — descubre el PDF del mes, lo parsea y calcula (Python + pdfplumber).
- `build_excel.py` — genera el Excel de 7 hojas.
- `requirements.txt` — dependencias.
- `.github/workflows/update.yml` — programación diaria.
- `index.html` — WebApp.
