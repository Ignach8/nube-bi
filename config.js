// Configuración de Nube BI.
// Pega aquí el "ID de cliente" de Google (ver GUIA.md, paso 2).
// Mientras esté vacío, la app funciona en "modo local" (sin Drive).
window.NUBE_CONFIG = {
  GOOGLE_CLIENT_ID: "446790735784-41hh220m6s0t889kriguu9l2tgjn5jo6.apps.googleusercontent.com",

  // Cada cuántos segundos revisa Drive en busca de cambios.
  POLL_SECONDS: 20,

  // Cada cuántos segundos revisa los Excel locales vigilados (solo Windows/Chrome/Edge).
  LOCAL_WATCH_SECONDS: 4,

  // Carpeta principal que la app crea en tu Drive.
  ROOT_FOLDER: "Nube BI"
};
