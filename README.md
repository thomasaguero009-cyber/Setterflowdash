# SetterFlow · Factor Studios

Dashboard de performance para el equipo de setters, con datos en vivo desde Google Sheets.

## Archivos

- **`index.html`** — el dashboard completo (frontend). Es un solo archivo, no necesita build ni instalación.
- **`apps_script_setterflow.gs`** — el backend, va pegado dentro de Google Apps Script (en tu Google Sheet real), no en GitHub.
- **`preview.jpg`** — imagen de vista previa para cuando se comparte el link (WhatsApp, redes, etc.).

## Cómo subirlo a GitHub

### 1. Crear el repositorio
1. Andá a [github.com/new](https://github.com/new)
2. Ponele un nombre (ej: `setterflow-dashboard`)
3. Dejalo en **Público** (necesario para GitHub Pages gratis)
4. Creá el repositorio

### 2. Subir los archivos
**Opción fácil (sin usar la terminal):**
1. Entrá al repositorio recién creado
2. Click en **"Add file" → "Upload files"**
3. Arrastrá `index.html`, `preview.jpg` (y `apps_script_setterflow.gs` si querés tenerlo ahí de respaldo, aunque ese va en Apps Script, no en la web)
4. Click en **"Commit changes"**

### 3. Activar GitHub Pages
1. Andá a **Settings** (del repositorio) → **Pages** (en el menú de la izquierda)
2. En "Branch", elegí **main** y la carpeta **/ (root)**
3. Guardá
4. Esperá 1-2 minutos — tu link va a quedar algo como:
   `https://TU-USUARIO.github.io/setterflow-dashboard/`

### 4. Actualizar la imagen de vista previa (opcional pero recomendado)
El `index.html` tiene el link de la imagen de portada apuntando a Netlify. Si tu sitio principal ahora es GitHub Pages, buscá esta línea (aparece 2 veces, cerca del principio del archivo) y cambiá la URL por la tuya de GitHub Pages:

```html
<meta property="og:image" content="https://setterflow-factorstudios.netlify.app/preview.jpg">
```

Cambiala a algo como:
```html
<meta property="og:image" content="https://TU-USUARIO.github.io/setterflow-dashboard/preview.jpg">
```

## Apps Script (el backend — no va en GitHub)

Esto se instala aparte, dentro de tu Google Sheet:
1. Abrí tu Google Sheet real (donde están las hojas de THOMI, FLOR, VALERIA, FRANCO)
2. **Extensiones → Apps Script**
3. Borrá lo que haya y pegá todo el contenido de `apps_script_setterflow.gs`
4. **Guardar** (ícono de disquete)
5. **Implementar → Nueva implementación** (o "Administrar implementaciones" si ya existe una, y ahí "Editar" → "Nueva versión")
   - Tipo: Aplicación web
   - Ejecutar como: Yo (tu cuenta)
   - Quién tiene acceso: **Cualquier usuario** (para que funcione para todo el equipo, no solo vos)
6. Copiá la URL que te da (termina en `/exec`)
7. Confirmá que esa URL sea la misma que está en `index.html` en la constante `WRITE_URL` (buscala con Ctrl+F en el archivo)

## Actualizar el sitio en el futuro

Cada vez que yo te dé una versión nueva de `index.html`:
1. Repetí el paso 2 de arriba (Upload files → arrastrar el nuevo `index.html` → Commit)
2. GitHub Pages se actualiza solo en 1-2 minutos
3. Recordale al equipo hacer **Ctrl+Shift+R** (o Cmd+Shift+R en Mac) la primera vez que entren después de una actualización, para que el navegador no les muestre una versión vieja guardada en caché
