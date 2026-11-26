# Nube BI — Guía rápida

App tipo Power BI que convierte tus Excel en dashboards automáticos y los guarda ordenados en tu Google Drive.
Es una **app web instalable (PWA)**: la misma app se instala en **Windows** y en **Android**.

## Qué hace sola

| Función | Cómo |
|---|---|
| Detecta columnas | Fechas, montos, cantidades, categorías, códigos y texto |
| Limpia | Filas vacías, encabezados repetidos, columnas vacías; avisa de duplicados |
| Recomienda gráficos | Tendencia por día/mes, comparación por categoría, participación (dona), KPIs |
| Ordena en Drive | `Nube BI / <Ventas, Finanzas, Inventario, Personal, Marketing, General> / <año>` |
| Recomienda archivos | Detecta archivos con las mismas columnas y ofrece **combinarlos** |
| Se actualiza sola | Revisa Drive cada 20 s; los Excel "vigilados" en la PC se releen al guardarlos |
| Funciona sin internet | Usa la última copia guardada en el dispositivo |

## 1. Probarla ahora (modo local, sin Drive)

Doble clic en **`iniciar.bat`**. Se abre `http://localhost:8080`.
Sube los archivos de la carpeta `ejemplos/` para verla funcionando.

> Requiere Python (ya lo tienes instalado). No cierres la ventana negra mientras uses la app.

## 2. Conectar Google Drive (una sola vez, unos 15 min)

> ✅ **Ya hecho (26-11-2026):** proyecto `nube-bi` en Google Cloud, Drive API habilitada,
> tu correo como usuario de prueba y el ID de cliente puesto en `config.js`.
> Si alguien más va a usar la app, agrégalo en **Google Auth Platform → Público → Usuarios de prueba**.
> Los pasos de abajo quedan como referencia.

1. Entra a <https://console.cloud.google.com/> con tu cuenta de Google → **Crear proyecto** → nombre: `Nube BI`.
2. Menú **APIs y servicios → Biblioteca** → busca **Google Drive API** → **Habilitar**.
3. **APIs y servicios → Pantalla de consentimiento de OAuth** (o "Google Auth Platform"):
   - Tipo de usuario: **Externo**.
   - Nombre de la app: `Nube BI`; correo de asistencia: el tuyo.
   - En **Usuarios de prueba** agrega tu propio correo (y los de quien vaya a usarla).
4. **APIs y servicios → Credenciales → Crear credenciales → ID de cliente de OAuth**:
   - Tipo: **Aplicación web**.
   - **Orígenes de JavaScript autorizados**: agrega `http://localhost:8080`
     (y más tarde la dirección donde la publiques, ver paso 3).
   - Crear → copia el **ID de cliente** (termina en `.apps.googleusercontent.com`).
5. Abre `config.js` y pégalo:
   ```js
   GOOGLE_CLIENT_ID: "123456-abc.apps.googleusercontent.com",
   ```
6. Recarga la app → **Conectar Google Drive** → acepta.

**Permiso que pide:** `drive.file`, la app **solo ve los archivos que ella misma sube**. No puede leer ni borrar el resto de tus 5 TB.

## 3. Usarla en Android (y en cualquier PC)

El teléfono no puede abrir `localhost` de tu PC, así que la app debe estar publicada en una dirección `https://`. Opciones gratis:

- **Netlify Drop**: entra a <https://app.netlify.com/drop> y arrastra esta carpeta completa. Te da una dirección tipo `https://nube-bi-123.netlify.app`.
- **GitHub Pages**: sube la carpeta a un repositorio y activa Pages.

Después agrega esa dirección en **Orígenes de JavaScript autorizados** (paso 2.4).

**Instalar:**
- **Android**: abre la dirección en Chrome → menú ⋮ → **Instalar app** (o "Agregar a pantalla principal").
- **Windows**: abre la dirección en Edge o Chrome → icono de instalar en la barra de direcciones (o el botón **Instalar app** dentro de Nube BI).

## Cómo funciona la actualización automática

- **Entre dispositivos**: lo que subes en Windows aparece en Android (y viceversa) en máximo 20 segundos, mientras la app esté abierta.
- **Excel editado en tu PC** (Windows, Chrome/Edge): usa **Abrir y vigilar Excel** o arrastra el archivo a la ventana. Cada vez que guardes en Excel, el dashboard se actualiza y la nueva versión se sube a Drive. Tras reiniciar el navegador aparece **Reactivar vigilancia** (el navegador exige un clic para volver a dar permiso).
- **Mismo nombre = misma fuente**: si subes otra vez `ventas.xlsx`, se reemplaza en vez de duplicarse.
- **Combinados**: se recalculan solos cuando cambia cualquiera de sus archivos.
- **Los cambios se ven sin recargar la página**: los números se animan y los gráficos se actualizan en su lugar.

## Límites a tener en cuenta

- **Sesión de Google de 1 hora**: Google no permite sesiones largas sin un servidor. La app la renueva sola en tu siguiente toque; si estuvo inactiva, el indicador de arriba dirá *"Toca para reconectar Drive"*.
- **Tamaño**: el análisis ocurre en el dispositivo. Bien hasta ~200–300 mil filas en PC y ~100 mil en teléfono. Para millones de filas hace falta un servidor (BigQuery).
- **Solo con la app abierta**: los navegadores no permiten revisar Drive en segundo plano con la app cerrada; al abrirla se pone al día sola.
- Los **combinados** se guardan solo en cada dispositivo (no en Drive); en el otro dispositivo se recrean con un clic en "Combinar y comparar".
- Los archivos que subas **directamente** a Drive (fuera de la app) no se ven, por el permiso `drive.file`; súbelos desde la app.

## Archivos

| Archivo | Qué es |
|---|---|
| `index.html`, `styles.css` | Interfaz |
| `app.js` | Lógica: Drive, sincronización, vigilancia, dashboard |
| `worker.js` | Motor de análisis (corre en segundo plano para que la app no se trabe) |
| `config.js` | Tu ID de cliente de Google y tiempos de actualización |
| `sw.js`, `manifest.webmanifest`, `icons/` | Hacen que se instale como app y funcione sin internet |
| `ejemplos/` | Datos de prueba |
| `iniciar.bat` | Abre la app en este PC |
