# Proyecto web institucional
Antiguo y Primitivo Rito de Memphis Misraim
Portal público + Panel de control administrado

## Estado
Repositorio preparado para despliegue provisional en GitHub Pages / Firebase Hosting.
Frontend estático con conexión opcional a Firebase Firestore.

## Para alzar ahora
1) Ajustar `firebaseConfig` en cada HTML bajo `Web/` por el proyecto real si no querés exponer la config, o dejarlo si es intencional para dev.
2) Activá Firebase Hosting contra la carpeta `Web/` usando `firebase.json` incluido.
3) Para GitHub Pages, usá `Web/` como root de publicación o mové el portal a `/docs`.

## Privilegios y datos sensibles
- El panel posee bloqueo por documento `usuarios/{uid}`.
- Existe función oculta de consola en `dashboard.html`: `window.hacermeAdmin()`.
- Quitá esa función antes del deploy público real.

## Estructura relevante
- Portal público: `Web/index.html`, `Web/auth.html`
- Panel: `Web/dashboard*.html`, `Web/ficha_obrero.html`, `Web/registro_obreros_pasaporte.html`
- Datos modelo: Firestore collections `obreros`, `logias`, `decretos`, `actas`, `tesoreria`, `usuarios`
