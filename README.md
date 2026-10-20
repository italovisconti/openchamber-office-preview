# Office Preview para OpenChamber

Visor local y de solo lectura de archivos **XLSX, DOCX y PPTX**, integrado en las pestañas de archivos de OpenChamber. Sin cuentas ni servicios de conversión externos.

## Instalación

1. Descarga el **[ZIP precompilado v0.1.0](openchamber-office-preview/dist/openchamber-office-preview-0.1.0.zip)** usando «Download raw file» en GitHub.
2. Abre **Settings → Extensions** en OpenChamber e instala el ZIP.
3. Selecciona un `.xlsx`, `.docx` o `.pptx` en el árbol de archivos.

No hace falta compilar ni instalar dependencias para usarlo. El paquete pesa aproximadamente **1 MB**.

## Qué ofrece

- **XLSX:** hojas visibles, valores y formato básico, navegación paginada y resultados de fórmulas guardados en el archivo.
- **DOCX:** texto, tablas e imágenes, con desplazamiento para páginas anchas.
- **PPTX:** diapositivas visibles, formas, imágenes y gráficos compatibles con el motor incluido.
- Recursos locales, temas claro/oscuro y sin botones de edición o guardado.

## Compatibilidad y límites

Probado en **OpenChamber web 2.5.0 con Chromium**. Desktop todavía no está probado. Máximo 20 MB por archivo, con límites adicionales de seguridad y tamaño de hojas. Se rechazan documentos con contenido activo, recursos externos y otras características no compatibles.

La representación es aproximada: no sustituye Microsoft Office, no recalcula fórmulas ni reproduce animaciones o vídeo. Las pruebas observaron los originales sin cambios; no constituyen una garantía universal sobre todas las versiones o rutas de guardado del host.

## Desarrollo

Consulta el **[README de la extensión](openchamber-office-preview/README.md)**, las **[pruebas y limitaciones verificadas](openchamber-office-preview/docs/verification.md)** y los **[avisos de terceros](openchamber-office-preview/THIRD_PARTY_NOTICES.md)**.

Código propio bajo licencia [MIT](openchamber-office-preview/LICENSE); los motores conservan sus respectivas licencias, incluidas en el ZIP.
