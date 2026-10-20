import './styles.css';

export type OfficeFormat = 'xlsx' | 'docx' | 'pptx';
export type ShellState = { kind: 'loading' | 'ready' | 'empty' | 'error'; message?: string };
export type NavigationItem = { id: string; label: string };
const stateMessages: Record<ShellState['kind'], string> = {
  loading: 'Abriendo documento…',
  ready: '',
  empty: 'Este documento no contiene contenido para mostrar.',
  error: 'No se pudo abrir este documento.',
};

/** Creates the preview chrome. The host owns tabs, file tree and fullscreen; renderer owns viewport content. */
export function createShell(root: HTMLElement) {
  const shell = document.createElement('section');
  shell.className = 'office-shell';
  shell.setAttribute('aria-label', 'Vista previa del documento');
  const toolbar = document.createElement('header'); toolbar.className = 'office-toolbar';
  const identity = document.createElement('div'); identity.className = 'office-identity';
  const badge = document.createElement('span'); badge.className = 'office-badge';
  const name = document.createElement('span'); name.className = 'office-name';
  name.textContent = 'Documento';
  identity.append(badge, name);
  const nav = document.createElement('nav'); nav.className = 'office-navigation';
  nav.setAttribute('aria-label', 'Secciones del documento');
  toolbar.append(identity, nav);
  const viewport = document.createElement('main'); viewport.className = 'office-viewport';
  viewport.tabIndex = 0; viewport.setAttribute('aria-label', 'Contenido del documento');
  const status = document.createElement('div'); status.className = 'office-status';
  status.setAttribute('role', 'status'); status.setAttribute('aria-live', 'polite');
  const statusText = document.createElement('span'); statusText.className = 'office-status-text';
  status.append(statusText); shell.append(toolbar, viewport, status); root.append(shell);
  let state: ShellState['kind'] = 'empty';

  function setState(next: ShellState) {
    state = next.kind; shell.dataset.state = next.kind;
    // Keep renderer DOM untouched whenever it is ready.
    status.hidden = next.kind === 'ready';
    status.classList.toggle('is-loading', next.kind === 'loading');
    statusText.textContent = next.message ?? stateMessages[next.kind];
    if (next.kind !== 'ready') viewport.setAttribute('aria-busy', String(next.kind === 'loading'));
    else viewport.removeAttribute('aria-busy');
  }
  function setDocument(documentName: string, format: OfficeFormat) {
    name.textContent = documentName || 'Documento';
    badge.textContent = format.toUpperCase(); badge.dataset.format = format;
    badge.setAttribute('aria-label', `Formato ${format.toUpperCase()}`);
  }
  function setNavigation(items: NavigationItem[], activeId: string, onSelect: (id: string) => void) {
    nav.replaceChildren(); nav.hidden = items.length === 0;
    items.forEach((item, index) => {
      const button = document.createElement('button'); button.type = 'button';
      button.className = 'office-nav-item'; button.textContent = item.label;
      button.setAttribute('aria-current', item.id === activeId ? 'true' : 'false');
      button.tabIndex = item.id === activeId || (!items.some(x => x.id === activeId) && index === 0) ? 0 : -1;
      button.addEventListener('click', () => onSelect(item.id));
      button.addEventListener('keydown', event => {
        if (event.key !== 'ArrowRight' && event.key !== 'ArrowLeft') return;
        event.preventDefault(); const delta = event.key === 'ArrowRight' ? 1 : -1;
        const next = (index + delta + items.length) % items.length;
        const target = nav.querySelectorAll<HTMLButtonElement>('button')[next]; target.focus(); target.click();
      });
      nav.append(button);
    });
  }
  function dispose() { shell.remove(); }
  setState({ kind: 'empty' });
  return { viewport, setState, setDocument, setNavigation, dispose };
}
