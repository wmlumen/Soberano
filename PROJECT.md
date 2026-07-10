# Arquitectura del Proyecto — Nuevo Oriente Memphis Misraim

## Identidad
- Nombre: Antiguo y Primitivo Rito de Memphis Misraim
- Portal institucional orientado a la gestión masónica y publicación pública.
- Sede/Oriente de referencia: Asunción, Paraguay.

## Hosting objetivo
- Firebase Hosting apuntando a carpeta `Web/`
- GitHub Pages como publicación provisional alternativa
- Backend datos: Firebase Firestore (`mmpy-146a6`)

## Modelo funcional confirmado
- Portal público: `index.html`
- Acceso/Registro: `auth.html`
- Panel de Control: `dashboard*.html`
- Identidad visual: paleta `#050b14`, `#121c2d`, `#d4af37`, tipografía Cinzel + Montserrat

## Modelo de datos Firestore
- `obreros`
  - `pasaporte_numero`, `nombres`, `apellidos`, `grado`, `logia`, `oriente`, `nacionalidad`
  - `fechaIniciacion`, `fechaAumento`, `fechaExaltacion`
  - `esAfiliado`, `fechaAfiliacion`, `decretoAfiliacion`, `orienteOrigenAfiliacion`, `logiaAumentoAfiliacion`
  - `cargos[]`, `timestamp`
  - Subcolección: `obreros/{id}/planchas` → `titulo`, `fecha`, `grado`, `contenido`, `timestamp`
- `logias`
  - `numero`, `nombre`, `oriente`, `venerable`
- `decretos`
  - `numero`, `fecha`, `asunto`, `emisor`
- `actas`
  - `fecha`, `numero`, `estamento`, `tipo`, `archivo_url`
- `tesoreria`
  - `fecha`, `concepto`, `estamento`, `taller`, `banco`, `tipo`, `monto`
- `usuarios`
  - `is_admin_sitio`, `puede_ver_buzones`, `rol`, `displayName`, `email`

## Mapa estamental y talleres
- Soberano Santuario
  - Consolidado Soberano Santuario
- Supremo Consejo
  - Consolidado Supremo Consejo
  - Capítulo Rosacruz No. 1
  - Areópago No. 1
- Gran Logia
  - Consolidado Gran Logia
  - Logia Luxor No. 1
  - Logia Osiris No. 2
  - Logia Horus No. 3

## Catálogo financiero provisional
Ingresos: Capitación, Saco de la Viuda, Donaciones
Egresos: Pago de Templo, Impresiones, Regalos, Alquileres
Cuentas: Banco ITAÚ, Banco Familiar, Caja Chica

## Hacer / No hacer
- Hacer: preservar la lógica Firestore actual para no perder datos en el puente.
- Hacer: aislar `firebaseConfig` por entorno cuando se publique.
- Hacer: separar estilos globales reutilizables.
- No hacer: borrar ni renombrar colecciones mientras haya usuarios activos.
- No hacer: exponer en GitHub la función `hacermeAdmin` en deploy productivo.
