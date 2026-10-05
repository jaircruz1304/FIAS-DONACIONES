// Configuración de fuentes del Dashboard de Donaciones FIAS.
// La aplicación intenta primero leer el Excel. Si el navegador bloquea el acceso
// directo de SharePoint por CORS, configure excelProxyUrl con el Worker incluido
// en worker/excel-proxy.js. El JSON queda únicamente como contingencia.
window.DONACIONES_CONFIG = {
  excelProxyUrl: '',
  directExcelUrls: [
    'https://fiasec-my.sharepoint.com/personal/jcruzg_fias_org_ec/_layouts/15/download.aspx?share=IQAUjBvkE6TPT7eb5vhBTZxcAaAltunzyoWrEzM5q5mVINo',
    'https://fiasec-my.sharepoint.com/:x:/g/personal/jcruzg_fias_org_ec/IQAUjBvkE6TPT7eb5vhBTZxcAaAltunzyoWrEzM5q5mVINo?download=1'
  ]
};
