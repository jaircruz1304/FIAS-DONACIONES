# Dashboard de Donaciones FIAS — lectura directa de Excel

## Arquitectura actualizada
La aplicación ya no depende del workflow YAML para mostrar los cambios del Excel. El orden de lectura es:

1. **Excel en línea (fuente principal)**: en cada apertura, al presionar **Actualizar** y automáticamente cada 60 segundos, el navegador intenta descargar y leer el archivo Excel de SharePoint con SheetJS.
2. **JSON de contingencia**: si el Excel no puede ser leído (por ejemplo, bloqueo CORS de SharePoint), se utiliza `data/control-donaciones.json`.
3. **Respaldo incorporado en el HTML**: si tampoco está disponible el JSON, se muestra el snapshot validado incluido en `Donaciones.html`.

El Excel y el snapshot de contingencia de esta entrega contienen **8 registros**.

## Archivos principales
- `Donaciones.html`: dashboard y lógica de lectura directa.
- `config.js`: URLs de la fuente Excel y configuración opcional del proxy.
- `data/control-donaciones.json`: respaldo JSON, no fuente primaria.
- `data/control-donaciones-meta.json`: metadatos del respaldo.
- `.github/workflows/update-control-donaciones.yml`: mantiene actualizado el JSON de contingencia; la interfaz ya no depende de su ejecución.
- `worker/excel-proxy.js`: proxy CORS opcional para garantizar la lectura directa desde GitHub Pages cuando SharePoint bloquea solicitudes cross-origin.

## Publicación inmediata
Publique todos los archivos en el repositorio conservando la estructura. Si SharePoint permite la descarga directa desde el navegador, no se requiere ninguna configuración adicional.

### Si el navegador muestra “Respaldo JSON activo”
Eso significa que SharePoint bloqueó la lectura directa por CORS. Para mantener la lectura en tiempo real:

1. Cree un Cloudflare Worker y pegue el contenido de `worker/excel-proxy.js`.
2. Publique el Worker y copie su URL pública, por ejemplo `https://donaciones-excel.<cuenta>.workers.dev`.
3. Abra `config.js` y coloque esa URL en `excelProxyUrl`.
4. Vuelva a publicar `config.js`.

Desde ese momento el dashboard consultará el Excel a través del Worker y el JSON quedará únicamente como contingencia.

## Comportamiento del botón Actualizar
El botón **Actualizar** no ejecuta GitHub Actions. Vuelve a consultar el Excel fuente en ese mismo momento y reconstruye las vistas, filtros, KPIs y gráficos con la información recibida.
