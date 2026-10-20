import { PptxViewer } from '@file-viewer/pptx';
import '@file-viewer/capability-pptx-charts';
import { pptxWorkerSource } from '../pptx-worker-source';
import { throwIfAborted, type ViewerAdapter } from './types';
import { visibleSlidesCopy } from './pptx-package';

export const mount: ViewerAdapter = async (container, bytes, { signal }) => {
  throwIfAborted(signal);
  const copy = await visibleSlidesCopy(bytes);
  throwIfAborted(signal);
  const url = URL.createObjectURL(new Blob([pptxWorkerSource], { type: 'text/javascript' }));
  const root = document.createElement('div');
  root.className = 'office-pptx-viewer';
  let disposed = false;
  let worker: Worker | undefined;
  let viewer: PptxViewer | undefined;
  let resolveRender!: () => void;
  let rejectRender!: (error: unknown) => void;
  const rendered = new Promise<void>((resolve, reject) => { resolveRender = resolve; rejectRender = reject; });
  const abort = () => { disposed = true; root.remove(); worker?.terminate(); viewer?.destroy(); URL.revokeObjectURL(url); rejectRender(new DOMException('Viewer mount aborted', 'AbortError')); };
  signal.addEventListener('abort', abort, { once: true });
  try {
    container.replaceChildren(root);
    const opened = PptxViewer.open(copy, root, {
      workerFactory: () => { worker = new Worker(url, { type: 'classic' }); return worker; },
      engineOptions: { mediaProcess: false, slideMode: false, keyBoardShortCut: false }, lazySlides: true,
      presentationFullscreen: false, zipLimits: { maxFileBytes: 20_000_000 },
      onRenderComplete: resolveRender, onError: rejectRender,
    });
    // Capture cleanup immediately; abort must not wait on open/render.
    void opened.then((instance) => { viewer = instance; if (disposed) instance.destroy(); }, rejectRender);
    await rendered;
    if (signal.aborted || disposed) throw new DOMException('Viewer mount aborted', 'AbortError');
  } catch (error) { root.remove(); worker?.terminate(); viewer?.destroy(); throw error; }
  return { dispose() { disposed = true; signal.removeEventListener('abort', abort); root.remove(); worker?.terminate(); viewer?.destroy(); URL.revokeObjectURL(url); } };
};
