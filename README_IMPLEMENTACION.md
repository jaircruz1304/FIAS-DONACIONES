# Dashboard de Donaciones FIAS — versión Enterprise

## Diagnóstico de la incidencia

El Excel fuente, el JSON generado y la metadata contienen el nuevo registro. La causa principal de la demora de visualización era la **frecuencia de sincronización de 6 horas** del workflow. La versión Enterprise reduce el ciclo a **30 minutos** y mantiene ejecución manual.

## Arquitectura

1. OneDrive / SharePoint: `CONTROL DE DONACIONES 2025-2026 vf.xlsx`.
2. GitHub Actions descarga el XLSX.
3. `scripts/build-data.mjs` transforma la hoja `Control` a `data/control-donaciones.json`.
4. Se genera `data/control-donaciones-meta.json` con fecha, cantidad de registros y fuente.
5. `Donaciones.html` consulta ambos archivos desde GitHub Pages.
6. La página vuelve a consultar automáticamente cada 5 minutos y permite **Actualizar vista** manualmente.

## Mejoras incorporadas

- Sincronización programada cada 30 minutos.
- Validación automática de cantidad de registros, Nro. vacíos y duplicados antes de publicar.
- Indicador visible de última sincronización e integridad.
- Consulta `cache: no-store` con parámetro anti-caché.
- Autoactualización de la vista cada 5 minutos sin perder filtros activos.
- Tabla ordenada por Nro. descendente para mostrar primero las donaciones más recientes.
- Acciones prioritarias excluyen registros finalizados.
- Nueva interfaz corporativa Enterprise, sin imagen decorativa externa, con jerarquía visual más sobria.
- KPI adicional de procesos finalizados.

## Estructura en el repositorio

```
/Donaciones.html
/package.json
/scripts/build-data.mjs
/data/control-donaciones.json
/data/control-donaciones-meta.json
/.github/workflows/update-control-donaciones.yml
/README_IMPLEMENTACION.md
```

## Puesta en producción

1. Reemplazar `Donaciones.html` por la versión Enterprise.
2. Reemplazar `.github/workflows/update-control-donaciones.yml`.
3. Mantener `scripts/build-data.mjs`, `package.json` y la carpeta `data`.
4. Ejecutar manualmente el workflow una vez desde **Actions**.
5. Confirmar que la metadata reporta el mismo número de registros que el JSON.
6. Abrir GitHub Pages y pulsar **Actualizar vista** si la página estaba abierta previamente.

> Nota: “Actualizar vista” vuelve a consultar el JSON ya publicado. No ejecuta GitHub Actions ni descarga directamente desde OneDrive. La actualización del origen la realiza el workflow.
