# Office Preview para OpenChamber

## Steps
- [ ] Confirmar compatibilidad de OpenChamber y probar los tres motores con documentos representativos.
- [ ] Crear la extensión binaria e integrar el contrato de solo lectura sin modificar OpenChamber.
- [ ] Implementar XLSX con navegación entre hojas y representación segura de valores y formato básico.
- [ ] Implementar DOCX y PPTX con el motor validado para cada formato.
- [ ] Integrar la experiencia visual con OpenChamber y resolver estados de carga, límites y errores.
- [ ] Verificar seguridad, integridad de originales y funcionamiento offline dentro de OpenChamber.
- [ ] Empaquetar la extensión instalable y documentar instalación, compatibilidad y limitaciones.

## Objetivo y alcance
Crear `openchamber-office-preview`: al seleccionar un `.xlsx`, `.docx` o `.pptx` en el árbol de archivos de OpenChamber, el documento se muestra dentro de su pestaña existente, como ocurre con la extensión de Excalidraw. No es una aplicación separada ni un editor de Office.

Los tres formatos forman parte de la entrega. El orden de implementación no implica diferir ninguno. El visor será local, de solo lectura, sin cuentas, servicios de conversión externos, telemetría ni recursos CDN en tiempo de ejecución. Instalación y descarga de dependencias durante el desarrollo sí pueden necesitar Internet.

Fuera de alcance: editar o guardar documentos, recalcular fórmulas, ejecutar macros, animaciones de PowerPoint, colaboración y formatos antiguos `.doc/.xls/.ppt`. Tampoco construir parsers OOXML propios o un reemplazo de Microsoft Office. Se conservará siempre el original.

## Hechos verificados y límites de esta planificación
Revisado el 20 de octubre de 2026:
- La extensión oficial Excalidraw registra `contributes.fileEditors`; su manifiesto indica OpenChamber `>=2.0.4`. Esa cifra NO prueba que todas las capacidades binarias actuales existan en esa versión.
- El SDK en `main` admite `content: "binary"` y entrega `Uint8Array` mediante `onFileOpen`. El host sigue siendo dueño del archivo, de las pestañas, del refresco y del fullscreen.
- El contrato actual limita cada documento a 20.000.000 bytes. Un archivo mayor vuelve al visor integrado; no prometer apertura ilimitada.
- Las extensiones se cargan en un iframe aislado. Se requieren bundles clásicos IIFE; fuentes, estilos, workers y otros recursos deben acompañar el paquete. No asumir que un import dinámico o una URL relativa generada por una biblioteca funcionará por relay.
- OpenChamber instala extensiones ya compiladas; no construye TypeScript al instalarlas. Hay que distribuir los archivos generados.
- El soporte documentado de extensiones es web y desktop; móvil y VS Code quedan fuera de la compatibilidad inicial.
- La API no expone en el contrato revisado un manifiesto de «visor sin guardado». `readOnly` lo comunica el host, no lo impone nuestra extensión. La política de solo lectura requiere prueba de integración, no solo ocultar botones.
- `docx-preview` reconoce límites de paginación. `pptx-preview` permite usar su paquete, pero su README restringe redistribución/modificación de fuentes; no tratarlo como una dependencia permisiva por defecto.
- Los especialistas externos no pudieron completar sus revisiones por falta de saldo. Este plan se apoya en la lectura directa de las fuentes enlazadas; la revisión visual queda asignada al diseñador durante la ejecución.

## Arquitectura propuesta
Usar Excalidraw como referencia del puente con OpenChamber y del empaquetado, no copiar su editor ni arrastrar sus dependencias.

Flujo: árbol de archivos → OpenChamber lee el binario → iframe de Office Preview recibe los bytes → adaptador del formato los representa → navegación local. No pedir acceso general al proyecto si los bytes ya llegan por el host.

Una extensión, una entrada HTML y tres registros de archivo binario —uno por formato— permiten atribuir errores y cargar únicamente el motor necesario. No declarar un panel independiente, comandos de chat, herramientas de IA ni permisos de red que no sean necesarios.

### Elección de motores: decisión antes de construir el producto
La recomendación es intentar primero motores en el navegador, porque evita procesos locales, permisos amplios y dependencias del sistema. La primera etapa debe cerrar la decisión para LOS TRES formatos; no empezar con dos y descubrir al final que PowerPoint no encaja.

- XLSX: evaluar ExcelJS para lectura y una cuadrícula de solo lectura. No usar un editor completo solo para mostrar celdas. El formato numérico debe verificarse explícitamente: parsear un valor no equivale a mostrarlo como Excel.
- DOCX: probar `docx-preview`, con recursos locales y contenido activo deshabilitado. Comparar tablas, imágenes, encabezados, pies y saltos; no prometer coincidencia exacta de páginas.
- PPTX: evaluar licencia y redistribución de `pptx-preview` antes de adoptarlo. PPTXjs es una referencia MIT, pero sus dependencias antiguas —jQuery 1.x y JSZip 2.x en el README— lo descartan como elección automática. No construir un renderer de diapositivas propio.
- Si ningún motor de navegador supera las pruebas de PPTX o de fidelidad DOCX, la alternativa concreta es LibreOffice headless → PDF temporal → PDF.js dentro del mismo iframe. XLSX conservaría su cuadrícula, no se convertiría obligatoriamente en páginas impresas.

No implementar simultáneamente dos motores por formato ni crear un sistema genérico de proveedores. Elegir uno por formato después de las pruebas. Un cambio a conversión nativa requiere presentar al usuario sus dependencias y permisos antes de instalarlo: los servicios de extensión tienen acceso completo del usuario y no son una sandbox de seguridad. Si no hay una ruta viable, informar el bloqueo y pedir decisión; no reducir silenciosamente la entrega.

## Etapa 1 — Prueba técnica y decisión
**Responsables:** investigador técnico para fuentes/licencias; implementador para prototipo; coordinador valida evidencia.

1. Identificar la versión instalada de OpenChamber y fijar un commit/tag de referencia del SDK. Probar recepción de bytes con una extensión mínima. Determinar la versión mínima REAL para `content: "binary"` y declararla en `engines.openchamber`.
2. Revisar `src/main.tsx`, `scripts/build.ts` y el manifiesto de Excalidraw; conservar los avisos MIT si se reutiliza código.
3. Crear fixtures ficticios: un XLSX con varias hojas, fechas, porcentajes, moneda, combinaciones y fórmulas con/sin valor guardado; un DOCX con tablas, imágenes y varias secciones; un PPTX con varias diapositivas, imágenes, formas y un gráfico. Añadir los tres archivos ya existentes en `output/` al ensayo, sin alterarlos.
4. Medir dentro del iframe real: apertura, legibilidad, número de hojas/diapositivas, recursos bloqueados por CSP, peso de bundle y consumo de memoria. Un demo independiente en navegador no basta.
5. Guardar `docs/engine-decision.md` con versiones exactas, licencias, dependencias transitivas, resultados, capturas y elección por formato. Para browser-only: cero accesos externos y ninguna dependencia incompatible con el empaquetado del host.

**Salida exigida:** una ruta probada para cada formato y una lista explícita de limitaciones. Si se necesita conversión nativa, aprobar el diseño del servicio antes de continuar.

## Etapa 2 — Extensión y garantía de solo lectura
**Archivos previstos**, dentro de una carpeta nueva `openchamber-office-preview/` del checkout actual; no convertir ni sobrescribir el proyecto de trabajo de oficina:
- `package.json`, lockfile, `tsconfig.json`: identidad, registros de los tres formatos, SDK y versiones fijadas.
- `editor/index.html`, `scripts/build.ts`: entrada y construcción reproducible IIFE, recursos locales.
- `src/main.ts`: conexión al host y ciclo de vida del documento.
- `src/read-only-session.ts`: conservar bytes originales y política de snapshots.
- `src/viewers/xlsx.ts`, `docx.ts`, `pptx.ts`: un adaptador concreto por formato.
- `src/ui/`: componentes visuales y estilos propiedad del diseñador.
- `tests/fixtures/`, pruebas de contrato y `README.md`, `THIRD_PARTY_NOTICES.md`.

Contrato mínimo propuesto para los adaptadores: `mount(container, bytes, { signal }): Promise<{ dispose(): void }>`; el puente conserva el original y el adaptador recibe una copia si la biblioteca puede mutarlo. Evitar abstracciones extra hasta que una necesidad real lo justifique.

- Nunca emitir `dirty: true` ni `edited: true` por zoom, selección, scroll o cambio de hoja. Nunca llamar a `writeFile`, upload ni `requestFileSave`.
- Para `onFileSnapshot('save')`, rechazar expresamente con «Este visor es de solo lectura». No serializar ni devolver bytes transformados para guardar.
- Para `handoff`, devolver los bytes originales con versión estable si el host los necesita para mover la vista; demostrar que no causa escrituras. Si el host no permite esta separación, resolverlo antes de declarar la garantía de solo lectura.
- Verificar Ctrl/Cmd+S, guardado desde la barra y autosave: el SDK intercepta el atajo, por lo que no basta omitir nuestro botón.
- Al cambiar de archivo o recargar: cancelar el trabajo anterior y descartar resultados tardíos. Liberar DOM, workers y object URLs al desmontar.

**Prueba decisiva:** hashes, mtime y registro de llamadas de escritura sin cambios tras abrir, navegar, fullscreen y cerrar. Un hash igual por sí solo no demuestra que no hubo reescritura.

## Etapa 3 — Contenido de los tres formatos
**XLSX:** mostrar hojas visibles en orden; nombres y Unicode correctos; filas/columnas, texto, números, fechas, moneda y porcentajes; formato básico y celdas combinadas según lo validado en etapa 1. Los resultados de fórmulas proceden del archivo: no recalcular ni mostrar cero cuando no existe valor cacheado. Indicar «Sin resultado guardado» y explicar que los valores almacenados pueden estar desactualizados. No activar enlaces externos ni macros. No revelar hojas ocultas por defecto. Para grandes rangos, ventana virtualizada o paginación explícita, sin construir millones de celdas DOM.

**DOCX:** texto legible, estilos básicos, tablas e imágenes, encabezados y pies donde el motor los soporte. Mantener colores del documento. Identificar la vista como aproximada cuando la paginación no sea fiable; no inventar correspondencia con páginas de Word.

**PPTX:** mostrar todas las diapositivas visibles en orden, con composición, texto, formas e imágenes verificadas. Vista estática: no animaciones, reproducción de vídeo ni scripts. Advertir sobre elementos no representados o diferencias con Office; no entregar una extracción de texto como si fuera una previsualización visual.

**Si se aprueba LibreOffice:** agregar solo entonces `service/` y su permiso declarado. Usar copias de los bytes, nombres internos generados y un perfil temporal aislado; nunca abrir por una ruta arbitraria enviada por el iframe. Ejecutar sin shell, con macros/actualización de vínculos desactivadas, límites de tiempo/memoria y bloqueo de salida de red verificable. El permiso del SDK no proporciona ese bloqueo. No publicar un endpoint sin autenticación: usar el loopback/token y proxy del host. Diseñar conversión asíncrona y transporte por bloques para respetar el timeout SDK de 20 s y las respuestas máximas de 8.000.000 bytes. Gestionar TTL y limpiar exclusivamente temporales propios. No dejar PDFs derivados en la carpeta del usuario. La ubicación «local» es el host de OpenChamber; en un servidor remoto no significa la laptop del navegador.

## Etapa 4 — Experiencia integrada
**Propietario:** diseñador; coordinador revisa textos; implementador solo realiza ajustes mecánicos.

El diseño debe resolver la navegación de hojas/documentos/diapositivas y zoom usando el espacio de la pestaña existente, sin duplicar árbol de archivos, pestañas ni fullscreen del host. Debe respetar el tema recibido del SDK sin recolorear el contenido del documento. Textos en español como mínimo y accesibilidad de teclado, etiquetas y foco visibles.

Estados obligatorios verificables: cargando, documento listo, vacío, corrupto, cifrado/no soportado, demasiado grande y fallo del motor. Ninguno puede dejar una pantalla blanca o un spinner eterno. No ofrecer botones de edición ni guardado. Los controles exactos y su distribución se definen tras la revisión del diseñador; no se implementa un rediseño de OpenChamber.

## Etapa 5 — Seguridad y evidencia de aceptación
**Presupuesto de verificación:** pruebas pequeñas en los límites propios y una pasada de integración real; no probar internamente bibliotecas de terceros ni añadir suites duplicadas por cobertura.

- Implementador: pruebas de contrato del host —binario, snapshots, dirty siempre falso, cambio rápido A→B, abort y limpieza— y de valores XLSX sin cache, fechas y porcentajes.
- Implementador: inspeccionar ZIP/OOXML con límites de bytes descomprimidos y entradas, más timeout recuperable. El límite comprimido de 20 MB del host NO previene ZIP bombs. Rechazar archivos truncados, cifrados y tipos internos distintos de la extensión; nunca expandir contenido arbitrario a disco.
- Implementador: no insertar texto XLSX como HTML; no ejecutar HTML embebido/altChunks de DOCX, scripts SVG, vínculos activos peligrosos ni contenido OLE. No cargar relaciones, imágenes, fuentes o vídeos externos.
- Coordinador: abrir un fixture por formato desde el árbol real de OpenChamber, cambiar pestañas, usar fullscreen, refrescar después de modificar una copia externamente y repetir sin Internet. Capturar prueba visual de cada formato.
- Coordinador: registrar accesos de red al abrir y navegar; permitir solo recursos del host/paquete, nunca conversiones externas. Si hay servicio nativo, comprobar por separado sus conexiones: el registro del navegador no lo cubre.
- Coordinador: verificar integridad/no escrituras con autosave activado y desactivado, guardado manual y Ctrl/Cmd+S. Reabrir los originales con la aplicación de oficina como control adicional.
- Diseñador: comprobar legibilidad, controles, carga y errores en panel estrecho y fullscreen, tema oscuro/claro y navegación por teclado.

La extensión no se considera terminada si solo funciona en su demo, si falta uno de los tres formatos, si requiere Internet para abrir documentos o si el modo visor reescribe originales.

## Etapa 6 — Instalación y entrega
Generar una carpeta/ZIP instalable y probar instalación desde Settings → Extensions sin ejecutar un build en el equipo de destino. Incluir bundles, CSS, fuentes con licencia y workers locales. Comprobar los límites del instalador del host: la documentación actual indica ZIP de 20 MB, 500 archivos y 40 MB descomprimidos; no asumir que un build enorme será instalable. Usar Git si es necesario y validar su ruta de instalación por separado.

README con versión mínima comprobada, matriz real de compatibilidad, instrucciones offline, limitaciones de fórmulas/fidelidad, permisos y dependencias nativas si se eligieron. No instalar LibreOffice, habilitar permisos, publicar repositorios, hacer push ni crear PR sin la aprobación que corresponda. Bump de versión para futuras actualizaciones Git. El plan no se incorpora a commits del proyecto.

## Fuentes primarias
- Referencia y distribución: https://github.com/openchamber/openchamber-excalidraw
- Puente de ejemplo: https://github.com/openchamber/openchamber-excalidraw/blob/main/src/main.tsx
- Contrato binario: https://github.com/openchamber/openchamber/blob/main/packages/sdk/src/file-editor.ts
- SDK, CSP y límites: https://github.com/openchamber/openchamber/blob/main/packages/sdk/DOCUMENTATION.md
- Integración Files: https://github.com/openchamber/openchamber/blob/main/packages/ui/src/components/views/files/DOCUMENTATION.md
- Servicio local y permisos: https://github.com/openchamber/openchamber/blob/main/packages/sdk/GUEST_SERVICES.md
- DOCX: https://github.com/VolodymyrBaydalka/docxjs
- XLSX: https://github.com/exceljs/exceljs
- PPTX y condiciones: https://github.com/501351981/pptx-preview
- Alternativa PPTX MIT, no recomendación automática: https://github.com/meshesha/PPTXjs
- Conversión local: https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html