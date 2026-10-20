/** Navigation requested by an adapter (for example, selecting a worksheet). */
export interface ViewerNavigation { index: number; count: number; label: string }
export interface ViewerMountOptions {
  signal: AbortSignal;
  onNavigation?: (navigation: ViewerNavigation) => void;
}
export interface MountedViewer { dispose(): void }
export type ViewerAdapter = (container: HTMLElement, bytes: Uint8Array, options: ViewerMountOptions) => Promise<MountedViewer>;

export function throwIfAborted(signal: AbortSignal): void {
  if (signal.aborted) throw new DOMException('Viewer mount aborted', 'AbortError');
}
