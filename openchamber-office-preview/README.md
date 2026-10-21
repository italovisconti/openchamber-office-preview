# OpenChamber Office Preview

Read-only previews for `.xlsx`, `.docx`, and `.pptx` files in OpenChamber's existing file tab. The extension uses the bytes supplied by OpenChamber and renders them in the browser; it is not an Office editor.

## Install

In OpenChamber, open **Settings → Extensions** and paste `https://github.com/italovisconti/openchamber-office-preview.git` into the Git URL installer. For SSH, use `git@github.com:italovisconti/openchamber-office-preview.git`; the OpenChamber server must have a configured Git/SSH identity with repository access. This identity is configured on the server instance, not in the browser. The repository root manifest installs the prebuilt package; no host build is needed. Alternatively, download the [prebuilt v0.1.0 ZIP](https://github.com/italovisconti/openchamber-office-preview/releases/download/v0.1.0/openchamber-office-preview-0.1.0.zip) from [GitHub Releases](https://github.com/italovisconti/openchamber-office-preview/releases/tag/v0.1.0), not the source-code archives, and install that ZIP. For local development, run the commands below from this directory; `bun run package` creates `dist/openchamber-office-preview-0.1.0.zip` locally. Built ZIPs are distributed as release assets, not tracked in Git.

## Compatibility status

The manifest declares OpenChamber `>=2.5.0`; 2.5.0 is the first version checked, not a verified historical minimum. Installation and XLSX, DOCX, and PPTX opening were verified in an isolated OpenChamber 2.5.0 web host using Chromium. In a separate DOCX security check, Chromium confirmed external and same-origin CSS image requests were blocked by `img-src`, with zero requests reaching network routing; a legitimate embedded PNG loaded. Desktop-host behavior remains untested.

## Privacy and read-only behavior

Documents are parsed from the bytes delivered by the host. The extension declares no network origins, conversion service, telemetry, or filesystem capability, and bundles its rendering assets and PPTX worker locally. Chromium verified DOCX CSS image requests to both external and same-origin URLs are blocked and embedded images still load; this is not a general guarantee that all host operations are offline.

The UI does not offer editing. Save snapshots are rejected; handoff returns a copy of the original bytes for the host's source view. This extension never reports edits or calls a file-write API. In the observed web-host test, fixture hashes, modification metadata and directory write events stayed unchanged after opening, navigation, panel expansion and save attempts. This is evidence for those tested paths, not a universal guarantee about every host version or operation.

## Development

Requires Bun 1.3+.

```sh
bun install --frozen-lockfile
bun run typecheck
bun run test:bridge
bun run test:viewers
bun run test:security
bun run build
bun run package
```

`bun run package` produces precompiled classic IIFE bundles, local styles, and a PPTX worker embedded in its bundle, then creates a ZIP ready for the Extensions installer. The archive is checked against OpenChamber's documented limits of 20 MB compressed, 40 MB expanded, and 500 files. The three renderer bundles are built separately so each file type loads only its own renderer.

## Engines and licenses

ExcelJS, `docx-preview`, `@file-viewer/pptx`, and `@file-viewer/capability-pptx-charts` provide the format renderers. Exact direct dependency versions and license terms are listed in `package.json` and `THIRD_PARTY_NOTICES.md`. The PPTX worker and third-party worker notices are included locally; full license texts for direct and transitive dependencies are included in `licenses/` in the install archive. Keep those notices when redistributing.

## Limitations

- OpenChamber hands binary editors at most 20,000,000 bytes; larger files remain in the host's viewer.
- XLSX safety limits include 100,000 rows, 500 columns, 250,000 populated cells, 200 worksheets, and a combined 10,000 merged-cell area. Oversized sheets are rejected before ExcelJS loads them.
- Opening has a 30-second asynchronous deadline and cancels cooperative work; JavaScript timers cannot interrupt a synchronous parser or renderer already running.
- Formula values are not recalculated; cached formula results may be absent or stale.
- PPTX chart rendering uses the bundled renderer capability and cached chart data in the presentation; embedded chart workbooks are removed before rendering and are never opened or recalculated. Unsupported chart types or data may not render.
- The read-only preview parses documents in the browser. Complex or malformed Office features may be omitted or rejected; it is not a substitute for Office rendering.
- DOCX pagination and Office rendering fidelity are not guaranteed.
- PPTX animations, video playback, and editing are unsupported.
- Documents containing external relationships (including ordinary web hyperlinks), embedded active content or unsupported features are conservatively rejected rather than partially trusted.
- Web Chromium checks cover installation, opening all three formats, chart rendering, DOCX external/same-origin CSS image blocking and embedded PNG loading. Desktop behavior is untested. See [verification](docs/verification.md) for the scope of the observed no-write evidence.
