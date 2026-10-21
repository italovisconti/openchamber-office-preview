import { connectHost, type FileEditorDocument } from '@openchamber/sdk';
import { applyHostReady } from '@openchamber/sdk/ui';
import { createShell, type OfficeFormat } from './ui/shell';
import type { ViewerAdapter } from './viewers/types';
import { ReadOnlySession } from './read-only-session';

export function startOfficePreview(format: OfficeFormat, adapter: ViewerAdapter): void {
  const host = connectHost();
  const root = document.getElementById('root');
  if (!root) throw new Error('Falta el contenedor principal del visor.');
  const shell = createShell(root);
  const session = new ReadOnlySession(shell, format, adapter);
  let generation = 0;

  host.onReady((context) => {
    applyHostReady(context, document.documentElement);
    // Documents and native form controls keep a light canvas even in a dark host.
    document.documentElement.style.colorScheme = 'light';
  });
  host.onFileOpen((file: FileEditorDocument) => {
    const current = ++generation;
    if (file.encoding !== 'binary') {
      shell.setState({ kind: 'error', message: 'El host no entregó el archivo como binario.' });
      session.dispose();
      return;
    }
    shell.setDocument(file.name, format);
    void session.open(file.bytes).catch((error: unknown) => {
      // Session renders the real error message. Log diagnostic details without
      // turning a failed opening into an unhandled rejection.
      if (current === generation) console.error('[office-preview] opening failed', error);
    });
  });
  host.onFileSnapshot((purpose) => session.snapshot(purpose));
  window.addEventListener('pagehide', () => {
    generation += 1;
    session.dispose();
    shell.dispose();
    host.dispose();
  }, { once: true });
}
