# Third-party notices

This extension bundles and redistributes third-party code. The build pins direct dependency versions in `package.json` and `bun.lock`. The package process copies all installed dependency `LICENSE*`, `NOTICE*`, and worker notice documents to `licenses/` in the installable archive; retain that directory with the compiled bundles.

Direct runtime dependencies:

| Package | Pinned version | License | Full text in install ZIP |
| --- | --- | --- | --- |
| `@openchamber/sdk` | 2.5.0 | MIT | `licenses/@openchamber/sdk/LICENSE` |
| `exceljs` | 4.4.0 | MIT | `licenses/exceljs/LICENSE` |
| `docx-preview` | 0.4.1 | Apache-2.0 | `licenses/docx-preview/LICENSE` |
| `@file-viewer/capability-pptx-charts` | 3.2.2 | Apache-2.0 | `licenses/@file-viewer/capability-pptx-charts/LICENSE` |
| `@file-viewer/pptx` | 3.2.2 | Apache-2.0 | `licenses/@file-viewer/pptx/LICENSE` |
| `fflate` | 0.8.3 | MIT | `licenses/fflate/LICENSE` |
| `jszip` | 3.10.2 | MIT OR GPL-3.0-or-later | `licenses/jszip/LICENSE.markdown` (both full license texts) |
| `ssf` | 0.11.2 | Apache-2.0 | `licenses/ssf/LICENSE` |

The install archive includes the full license documents themselves under `licenses/` (including transitive package copyright notices), not merely external links. It also includes `@file-viewer/pptx`'s complete worker notice at `licenses/file-viewer-pptx-worker-THIRD_PARTY_LICENSES.txt`. Preserve those files and this notice when redistributing the bundles. The product's own license is in the root `LICENSE` file.

## Audit note

`bun audit` currently reports `uuid@8.3.2`, brought in transitively by ExcelJS, for GHSA-w5hq-g745-h8pq. The advisory affects versions below 11.1.1 when the v3/v5/v6 APIs receive a caller-provided output buffer. ExcelJS's only use of this dependency found in the installed source is `uuid.v4()` (for generated XLSX conditional-formatting identifiers); the affected APIs and buffer argument are not used. We retain the transitive version rather than force an unverified major upgrade. Reassess when ExcelJS updates its dependency or the product begins using other UUID APIs.
