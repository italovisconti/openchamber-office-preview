# Selección de motores

Fecha: 2026-10-20. Versiones fijadas en `package.json` y `bun.lock`.

## Decisión

Usar motores independientes en el navegador, sin framework de interfaz ni editor ofimático completo. Cada formato tiene un bundle IIFE distinto. Los recursos y el worker se incluyen localmente; no se solicitan permisos de red ni de archivos.

| Formato | Motor | Motivo y límites |
| --- | --- | --- |
| XLSX | ExcelJS 4.4.0 + SSF 0.11.2 | Lectura de hojas/estilos y formato numérico existente. Cuadrícula propia de solo lectura, paginada. No recalcula fórmulas. |
| DOCX | docx-preview 0.4.1 | Motor ligero existente para texto, tablas e imágenes. Paginación aproximada; altChunks deshabilitados, fuentes embebidas ignoradas. |
| PPTX | @file-viewer/pptx 3.2.2 + capacidad oficial de gráficos 3.2.2 | Distribución Apache-2.0 con componentes permisivos; worker autocontenido local. Copia efímera filtrada para no revelar diapositivas ocultas/huérfanas. Gráficos básicos mediante Billboard incluido localmente. |

JSZip se usa para la copia efímera PPTX; fflate aplica presupuestos de expansión y genera un ZIP canónico con las partes inspeccionadas antes de los motores. Los bytes originales quedan fuera de los motores y del worker.

## Alternativas consideradas

- `pptx-preview`: no adoptado por sus condiciones de redistribución de fuentes.
- PPTXjs histórico: no adoptado automáticamente por su cadena antigua de dependencias y mayor coste de mantenimiento.
- LibreOffice → PDF.js: alternativa de conversión nativa, no implementada. Añadiría instalación, proceso local, permisos y aislamiento. Requeriría aprobación explícita; no se instala como fallback silencioso.
- Suite/editor completo: no necesario para un visor de solo lectura.

## Evidencia y limitaciones

El ZIP se ha instalado ya compilado en OpenChamber web 2.5.0 con configuración de pruebas independiente. Chromium abrió un documento ficticio de cada formato desde el árbol real, en el iframe del host. Se verificaron valores XLSX, cambio de hoja, texto/tablas DOCX y texto/formas/orden de diapositivas PPTX. La diapositiva y hoja ocultas del ensayo no se mostraron.

La prueba inicial detectó un gráfico PPTX omitido; se resolvió habilitando la capacidad oficial de gráficos y se comprobó visualmente el gráfico de barras y sus etiquetas. No se debe interpretar `ready` como una garantía de fidelidad completa. No se ha establecido una versión histórica mínima ni probado la aplicación desktop.

La distribución mide aproximadamente 1,1 MB comprimidos y 3 MB descomprimidos, incluidos avisos y licencias. Son medidas del paquete construido, no estimaciones de tamaños npm. Las cifras exactas se imprimen al empaquetar.

## Fuentes

- SDK y contrato instalado: `@openchamber/sdk@2.5.0`; host `@openchamber/web@2.5.0`.
- Referencia: <https://github.com/openchamber/openchamber-excalidraw>.
- ExcelJS: <https://github.com/exceljs/exceljs>.
- DOCX: <https://github.com/VolodymyrBaydalka/docxjs>.
- PPTX: <https://github.com/flyfish-dev/file-viewer>, artefacto npm `@file-viewer/pptx@3.2.2`.
- Términos aplicables y textos distribuidos: `THIRD_PARTY_NOTICES.md` y `licenses/` del ZIP.

La referencia Excalidraw se consultó sin fijar un commit histórico; no se usa esa consulta para afirmar compatibilidad con versiones anteriores.
