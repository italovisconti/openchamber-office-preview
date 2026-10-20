import { prepareOfficeArchive } from './security/preflight';
import type { MountedViewer, ViewerAdapter, ViewerNavigation } from './viewers/types';
import type { OfficeFormat, ShellState } from './ui/shell';

export interface SessionShell {
  viewport: HTMLElement;
  setState(state: ShellState): void;
  setNavigation(items: { id: string; label: string }[], activeId: string, onSelect: (id: string) => void): void;
}

/** One active file generation; original bytes never enter a renderer/worker. */
export class ReadOnlySession {
  private generation = 0;
  private original: Uint8Array<ArrayBuffer> | null = null;
  private version = '';
  private controller: AbortController | null = null;
  private mounted: MountedViewer | null = null;
  private activeContainer: HTMLElement | null = null;

  constructor(
    private readonly shell: SessionShell,
    private readonly format: OfficeFormat,
    private readonly adapter: ViewerAdapter,
    private readonly onNavigation?: (navigation: ViewerNavigation) => void,
    private readonly preflight: typeof prepareOfficeArchive = prepareOfficeArchive,
    private readonly timeoutMs = 30_000,
  ) {}

  /** Retires the previous viewer synchronously, then performs guarded async work. */
  open(input: Uint8Array<ArrayBuffer>): Promise<void> {
    this.closeCurrent();
    const generation = ++this.generation;
    this.version = `open-${generation}`;
    const original = new Uint8Array(input);
    this.original = original;
    const controller = new AbortController();
    this.controller = controller;
    const container = document.createElement('div');
    container.className = 'office-generation';
    this.activeContainer = container;
    this.shell.viewport.replaceChildren(container);
    this.shell.setNavigation([], '', () => undefined);
    this.shell.setState({ kind: 'loading' });

    // Begin only after ownership and DOM have been switched synchronously.
    const operation = Promise.resolve().then(async () => {
      const canonical = await this.preflight(original, this.format, { signal: controller.signal });
      this.assertCurrent(generation, controller.signal);
      // Only inspected, canonicalized parts reach the renderer. Original bytes remain for handoff.
      const rendererBytes = canonical.slice();
      const mounted = await this.adapter(container, rendererBytes, {
        signal: controller.signal,
        onNavigation: (navigation) => {
          if (!this.isCurrent(generation, controller.signal)) return;
          this.onNavigation?.(navigation);
        },
      });
      if (!this.isCurrent(generation, controller.signal)) {
        mounted.dispose();
        return;
      }
      this.mounted = mounted;
      this.shell.setState({ kind: 'ready' });
    });
    let timedOut = false;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const deadline = new Promise<never>((_, reject) => {
      if (this.timeoutMs <= 0) return;
      timer = setTimeout(() => {
        if (!this.isCurrent(generation, controller.signal)) return;
        timedOut = true;
        controller.abort();
        this.closeCurrent();
        this.shell.setState({ kind: 'error', message: 'La vista previa excedió el tiempo límite. Puedes volver a intentarlo.' });
        reject(new Error('La vista previa excedió el tiempo límite. Puedes volver a intentarlo.'));
      }, this.timeoutMs);
    });
    return Promise.race([operation, deadline]).catch((error: unknown) => {
      if (timedOut) throw error;
      if (!this.isCurrent(generation, controller.signal)) return;
      this.closeCurrent();
      this.shell.setState({ kind: 'error', message: errorMessage(error) });
      throw error;
    }).finally(() => { if (timer) clearTimeout(timer); });
  }

  snapshot(purpose: 'save' | 'handoff') {
    if (purpose === 'save') throw new Error('Este visor es de solo lectura');
    if (!this.original) throw new Error('No hay bytes originales disponibles');
    return { bytes: new Uint8Array(this.original), version: this.version };
  }

  dispose(): void {
    this.closeCurrent();
    this.generation += 1;
  }

  private isCurrent(generation: number, signal: AbortSignal): boolean {
    return generation === this.generation && !signal.aborted;
  }

  private assertCurrent(generation: number, signal: AbortSignal): void {
    if (!this.isCurrent(generation, signal)) throw new DOMException('Apertura cancelada', 'AbortError');
  }

  private closeCurrent(): void {
    // Abort and cleanup happen synchronously, before any await in open().
    this.controller?.abort();
    this.controller = null;
    this.mounted?.dispose();
    this.mounted = null;
    this.original = null;
    this.version = '';
    this.activeContainer?.replaceChildren();
    this.activeContainer = null;
    this.shell.viewport.replaceChildren();
  }
}

function errorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim()) return error.message.slice(0, 500);
  return 'Error desconocido al abrir el documento.';
}
