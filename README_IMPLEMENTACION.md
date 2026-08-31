# Dashboard de Donaciones FIAS — Enterprise V3

## Qué corrige esta versión
- Reforma visual completa: nueva navegación, panel ejecutivo, tarjetas por donación, vistas Ejecutiva / Análisis / Tabla y filtros compactos.
- La donación Nro. 7 se encuentra incluida en `data/control-donaciones.json` y además en un respaldo local incorporado dentro del HTML.
- Si el HTML se abre directamente con `file://`, se usa el respaldo local de 7 registros.
- En GitHub Pages, el dashboard intenta primero leer `./data/control-donaciones.json`; si falla, prueba rutas alternativas y finalmente usa el respaldo local.
- La donación más reciente se ordena primero por fecha de intención y luego por Nro.

## Estructura de publicación
Subir el contenido respetando exactamente:
- `Donaciones.html`
- `data/control-donaciones.json`
- `data/control-donaciones-meta.json`
- `scripts/build-data.mjs`
- `package.json`
- `.github/workflows/update-control-donaciones.yml`

## Verificación inmediata
Al abrir la nueva versión debe observarse en la parte superior:
- Registros: 7
- Último Nro.: Nro. 7
- Tarjeta destacada: Donación Nro. 7 · Rendimiento REM
- Valor: USD 174.435,86

## Fuente
En producción, la fuente oficial continúa siendo el JSON generado por GitHub Actions desde el Excel. El respaldo incorporado solo evita una pantalla vacía cuando el HTML se abre localmente o la ruta JSON no está disponible.
