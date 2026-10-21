# Evidencia de aceptación

Este archivo define las pruebas; no afirma que ya hayan pasado. Una compilación o un demo aislado no prueban compatibilidad con OpenChamber.

| Afirmación | Responsable | Evidencia mínima |
| --- | --- | --- |
| Cada formato se abre dentro de la pestaña del host | Coordinador | Un XLSX, DOCX y PPTX representativos abiertos desde el árbol real; captura de cada uno y versión exacta del host |
| Los originales no se escriben | Ejecutor: contrato; coordinador: integración | Save rechazado, handoff idéntico, dirty falso; observar llamadas de escritura, hash y mtime antes/después con guardado manual, atajo y autosave |
| A→B no muestra resultados tardíos de A | Ejecutor | Prueba determinista con finalización de A retrasada, cancelación y limpieza observables |
| Los documentos no requieren servicios externos | Coordinador | Capturar solicitudes en apertura y navegación; repetir sin Internet, permitiendo recursos locales del host/paquete |
| Archivos hostiles no producen ejecución ni expansión ilimitada | Ejecutor | Límites de entradas y bytes realmente descomprimidos, tipos OOXML, rechazo de cifrado/truncado, enlaces/altChunks/medios activos inertes y timeout |
| El contenido básico es útil | Ejecutor y diseñador | Varias hojas/diapositivas, Unicode, imágenes/tablas; fechas/porcentajes/moneda y fórmulas sin caché; panel estrecho y temas claro/oscuro |
| La distribución se instala ya compilada | Coordinador | Instalar artefacto local en host sin build; medir archivos, bytes ZIP y descomprimidos; avisos y licencias incluidos |

## Presupuesto

Pruebas enfocadas por lane y una ejecución completa tras reconciliar cada pasada. Reutilizar evidencia mientras no cambien código, entradas o entorno relevantes. Repetir solo lo invalidado por cambios. La revisión visual requiere abrir capturas, no solo generarlas.

## Límites de las conclusiones

Si no se dispone de una API, fixture o interacción necesaria, registrar la afirmación como pendiente junto con el bloqueo concreto. No sustituir la observación de escrituras por un hash idéntico, ni la prueba en iframe real por un demo independiente. No declarar terminada la extensión si falta cualquiera de los tres formatos.

## Resultado de v0.1.0 — 2026-10-20

Artefacto probado: `openchamber-office-preview-0.1.0.zip`.

- SHA-256: `80f1704abdeddf73d7fd5c5276a6894bfdcb04efa7d3c8b3f3c3b0c09e9666f5`.
- 1.067.738 bytes ZIP, 3.000.731 bytes descomprimidos, 283 archivos.
- `bun test tests/security tests/viewers tests/bridge`: **32 pasan, 0 fallan**.
- `bun test` completo, incluida validación del manifiesto: **33 pasan, 0 fallan**.
- `bun run typecheck`, `bun run package` y `unzip -t`: correctos.
- Revisión de seguridad final: sin bloqueantes en los riesgos revisados, tras corregir alias ZIP, recursos CSS DOCX, montajes pendientes y presupuestos XLSX evadibles.

### Integración observada

Se instaló ese ZIP, sin compilar en destino, en una instancia independiente de **OpenChamber web 2.5.0**, accesible solo por loopback y con HOME/configuración/datos nuevos. No se alteró la instancia habitual del usuario. Navegador: Chromium 156 automatizado con Playwright.

Desde el árbol real se abrieron documentos ficticios XLSX, DOCX y PPTX. Se observaron: dos hojas visibles, Unicode, porcentajes/moneda, resultados de fórmulas presentes y ausentes; texto y tablas DOCX; dos diapositivas PPTX con texto, formas y gráfico de barras. Hoja y diapositiva ocultas no aparecieron. También cargó una imagen PNG embebida en un DOCX independiente. Se abrieron capturas para revisar la legibilidad, incluido el tema oscuro.

Se repitieron apertura, cambio de hoja, expansión del panel y Ctrl+S/guardado manual con `autoSaveEnabled` explícitamente desactivado y activado. Para los tres originales del recorrido se conservaron SHA-256, inode, mtime y ctime; el observador inotify del directorio registró **cero eventos de escritura**. Las pruebas del puente establecen rechazo de save, handoff con bytes originales, aislamiento A→B, timeout y limpieza tardía.

Los recorridos del navegador bloquearon destinos ajenos al host local. Las solicitudes externas observadas procedían de logos de proveedores en la interfaz de OpenChamber, no de los documentos. La prueba adversarial DOCX confirmó mediante violaciones CSP que intentos CSS externos y del mismo origen se bloquearon **antes de llegar al routing de red**, mientras que la imagen PNG `data:` legítima cargó.

### Alcance pendiente y límites

- Desktop, móvil, relay remoto y otras versiones/navegadores no se han validado. La versión mínima histórica no está establecida.
- La evidencia de no escritura corresponde a esos archivos, configuración y recorridos; no demuestra todas las rutas internas del host. El handoff se comprueba en contrato, no como auditoría exhaustiva del host.
- No se ha establecido fidelidad exacta de paginación, SmartArt, todos los gráficos o todas las imágenes. Se rechazan características activas y relaciones externas, incluidos enlaces ordinarios.
- El watchdog permite recuperar esperas asíncronas; no interrumpe JavaScript síncrono. Los presupuestos previos limitan trabajo, no equivalen a aislamiento de proceso.
- `bun audit` conserva un aviso transitivo de `uuid` usado por ExcelJS. El uso instalado es `v4`; el aviso descrito afecta a v3/v5/v6 con buffer proporcionado. Véase `THIRD_PARTY_NOTICES.md`; no se afirma una auditoría limpia.

### Distribución en GitHub Releases — 2026-10-21

El ZIP se distribuye como asset de la release `v0.1.0`, no como archivo versionado en `dist`. Se reempaquetó únicamente para actualizar el enlace de instalación del README incluido; los bundles compilados no cambiaron respecto al artefacto probado arriba.

- Asset: `openchamber-office-preview-0.1.0.zip`, 1.067.801 bytes; 3.000.912 bytes descomprimidos y 283 archivos.
- SHA-256: `c6d9ad1ddc5233a7fbe6157d7133ee5162b1c6a045e5c50c2dfa9d12a6681480`.

## Actualización v0.1.1 — 2026-10-21

- `bun test`: **42 pasan, 0 fallan**; typecheck, build y empaquetado correctos.
- ZIP: 1.069.472 bytes, 3.005.844 bytes descomprimidos, 283 archivos. SHA-256: `c029b604edeb2dca8d2638ef1707dc12dc6abc471a127eba47cebad7837411ea`.
- Chromium confirmó fondo blanco y controles claros en los tres formatos, tanto con el host claro como oscuro. Se revisó una captura con documentos ficticios; los colores definidos por el documento no se invierten.
- Se reprodujo un rechazo PPTX causado por una relación interna de configuración de impresora. La copia canónica descarta únicamente las partes `printerSettings` válidas, sus relaciones y overrides; no se habilitan binarios/OLE arbitrarios. La muestra privada abrió en el iframe real y su hash y mtime originales no cambiaron. Ni la muestra ni su contenido se incorporan al repositorio.
- Las regresiones sintéticas cubren destinos externos, MIME incorrecto, ubicaciones no admitidas y retirada de workbooks de gráficos con diferencias de mayúsculas/minúsculas.
- Se repitió la prueba CSP DOCX: solicitudes CSS externas y del mismo origen bloqueadas antes del routing, con PNG embebido legítimo cargado.
- La compilación sincroniza el CSS inline de DOCX con la hoja generada para impedir que conserve una versión anterior del tema.
- No se añadió un botón nativo «Abrir en app»: el SDK 2.5.0 no expone esa acción a extensiones. Debe utilizarse el control propio del host; no se solicitan permisos nuevos ni se sustituye por una descarga.
