import { renderAsync } from 'docx-preview';
import { throwIfAborted, type ViewerAdapter } from './types';

/** Render into a generation-owned node; remove it before awaiting on disposal. */
export const mount: ViewerAdapter = async (container, bytes, { signal }) => {
  throwIfAborted(signal);
  const generation = document.createElement('div');
  generation.className = 'office-docx-generation';
  let disposed = false;
  let observer: MutationObserver | null = null;
  const links = new Set<HTMLAnchorElement>();
  const blockLinks = (event: Event) => { if ((event.target as Element | null)?.closest('a')) { event.preventDefault(); event.stopImmediatePropagation(); } };
  generation.addEventListener('click', blockLinks, true);
  const onAbort = () => { disposed = true; observer?.disconnect(); observer = null; generation.remove(); };
  signal.addEventListener('abort', onAbort, { once: true });
  try {
    const rendered = renderAsync(bytes.slice().buffer, generation, generation, {
      renderAltChunks: false, renderComments: false, renderChanges: false,
      experimental: false, ignoreFonts: true, useBase64URL: true,
    });
    // Render while detached; if link elements appear during render, strip them as soon as possible.
    observer = new MutationObserver(() => generation.querySelectorAll('a').forEach((a) => { links.add(a); a.removeAttribute('href'); a.removeAttribute('target'); }));
    observer.observe(generation, { childList: true, subtree: true });
    await rendered;
    observer.disconnect(); observer = null;
    if (signal.aborted || disposed) { generation.remove(); throw new DOMException('Viewer mount aborted', 'AbortError'); }
    generation.querySelectorAll('a').forEach((link) => links.add(link));
    links.forEach((link) => {
      link.removeAttribute('href'); link.removeAttribute('target'); link.setAttribute('aria-disabled', 'true');
    });
    container.replaceChildren(generation);
  } catch (error) {
    observer?.disconnect(); observer = null;
    signal.removeEventListener('abort', onAbort);
    generation.removeEventListener('click', blockLinks, true);
    generation.remove();
    throw error;
  }
  return { dispose() { disposed = true; signal.removeEventListener('abort', onAbort); observer?.disconnect(); observer = null; generation.removeEventListener('click', blockLinks, true); generation.remove(); } };
};
