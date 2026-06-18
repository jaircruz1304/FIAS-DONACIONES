# Implementación sin CORS para Dashboard de Donaciones FIAS

Esta versión evita leer OneDrive desde el navegador. GitHub Actions descarga el Excel en el servidor de GitHub, genera `data/control-donaciones.json` y el HTML lee ese JSON desde el mismo origen de GitHub Pages.

## Archivos

- `Donaciones.html`: dashboard público sin datos hardcodeados.
- `data/control-donaciones.json`: datos publicados para el dashboard.
- `scripts/build-data.mjs`: script que descarga el Excel público y genera el JSON.
- `.github/workflows/update-control-donaciones.yml`: automatización para actualizar el JSON cada 6 horas o manualmente.
- `package.json`: dependencias Node.js.

## Pasos

1. Sube todos estos archivos a la raíz del repositorio de GitHub Pages.
2. Verifica que GitHub Pages esté activo.
3. En GitHub, entra a Actions > Actualizar control de donaciones > Run workflow.
4. Abre `Donaciones.html` desde GitHub Pages.

## Resultado

El navegador solo lee `./data/control-donaciones.json`, por lo que no hay CORS con OneDrive.
