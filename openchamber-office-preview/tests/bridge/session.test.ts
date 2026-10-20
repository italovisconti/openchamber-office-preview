import { afterEach, describe, expect, test } from 'bun:test';
import { readFile, access } from 'node:fs/promises';
import { ReadOnlySession, type SessionShell } from '../../src/read-only-session';
import type { ViewerAdapter } from '../../src/viewers/types';

const oldDocument = globalThis.document;
const states: Array<{ kind: string; message?: string }> = [];
const makeElement = () => ({
  children: [] as unknown[],
  replaceChildren(...children: unknown[]) { this.children = children; },
});

function setup(adapter: ViewerAdapter, preflight: typeof import('../../src/security/preflight').prepareOfficeArchive = async (bytes) => bytes, timeoutMs = 30_000) {
  Object.defineProperty(globalThis, 'document', { configurable: true, value: { createElement: makeElement } });
  states.length = 0;
  const viewport = makeElement();
  const shell: SessionShell = {
    viewport: viewport as unknown as HTMLElement,
    setState: (state) => states.push(state),
    setNavigation: () => {},
  };
  return { session: new ReadOnlySession(shell, 'pptx', adapter, undefined, preflight, timeoutMs), viewport };
}

afterEach(() => {
  Object.defineProperty(globalThis, 'document', { configurable: true, value: oldDocument });
});

describe('read-only file-editor bridge', () => {
  test('manifest ships one local IIFE entry per binary format', async () => {
    const root = new URL('../../', import.meta.url);
    const pkg = JSON.parse(await readFile(new URL('package.json', root), 'utf8'));
    const editors = pkg.openchamber.contributes.fileEditors;
    expect(pkg.openchamber.engines.openchamber).toBe('>=2.5.0');
    expect(editors.map((editor: { match: string[] }) => editor.match[0])).toEqual(['*.xlsx', '*.docx', '*.pptx']);
    for (const editor of editors) {
      await access(new URL(editor.entry, root));
      const name = editor.id as string;
      const html = await readFile(new URL(editor.entry, root), 'utf8');
      const script = await readFile(new URL(`editor/${name}.js`, root), 'utf8');
      expect(html).toContain(`${name}.js`);
      expect(script).not.toMatch(/^\s*import\s/m);
    }
  });

  test('handoff is a per-open copy; save is rejected', async () => {
    let adapterBytes: Uint8Array | undefined;
    const adapter: ViewerAdapter = async (_container, bytes) => {
      adapterBytes = bytes;
      return { dispose() {} };
    };
    const { session } = setup(adapter);
    const input = new Uint8Array([1, 2, 3]);
    await session.open(input);
    input[0] = 9;
    adapterBytes![1] = 8;
    const handoff = session.snapshot('handoff');
    expect([...handoff.bytes!]).toEqual([1, 2, 3]);
    expect(handoff.version).toBe('open-1');
    handoff.bytes![0] = 7;
    expect([...session.snapshot('handoff').bytes!]).toEqual([1, 2, 3]);
    expect(() => session.snapshot('save')).toThrow('solo lectura');
  });

  test('renderer receives the canonical archive while handoff remains the untouched original', async () => {
    let rendererBytes: Uint8Array | undefined;
    const { session } = setup(async (_container, bytes) => { rendererBytes = bytes; return { dispose() {} }; }, async () => new Uint8Array([8, 9]));
    const input = new Uint8Array([1, 2, 3]);
    await session.open(input);
    expect([...rendererBytes!]).toEqual([8, 9]);
    expect([...session.snapshot('handoff').bytes!]).toEqual([1, 2, 3]);
  });

  test('A is aborted and removed before B mount; late A disposal cannot own B', async () => {
    const pending: Array<{ resolve: (value: { dispose(): void }) => void; signal: AbortSignal; container: HTMLElement }> = [];
    let disposedA = 0;
    const adapter: ViewerAdapter = (container, _bytes, { signal }) => new Promise((resolve) => {
      pending.push({ resolve, signal, container });
    });
    const { session, viewport } = setup(adapter);
    const a = session.open(new Uint8Array([1]));
    await Promise.resolve(); await Promise.resolve();
    const containerA = pending[0].container;
    const b = session.open(new Uint8Array([2]));
    expect(pending[0].signal.aborted).toBe(true);
    expect((viewport as unknown as { children: unknown[] }).children[0]).not.toBe(containerA);
    await Promise.resolve(); await Promise.resolve();
    pending[0].resolve({ dispose() { disposedA += 1; } });
    pending[1].resolve({ dispose() {} });
    await Promise.all([a, b]);
    expect(disposedA).toBe(1);
    expect(states.at(-1)?.kind).toBe('ready');
    expect(session.snapshot('handoff').version).toBe('open-2');
  });

  test('preflight failures reach UI and reject with the original error', async () => {
    const failure = new Error('OOXML package does not match the requested format');
    const { session } = setup(async () => ({ dispose() {} }), async () => { throw failure; });
    await expect(session.open(new Uint8Array([1]))).rejects.toBe(failure);
    expect(states.at(-1)).toEqual({ kind: 'error', message: failure.message });
    expect(() => session.snapshot('handoff')).toThrow('No hay bytes originales');
  });

  test('mount watchdog reports recoverable timeout, disposes a late mount, and permits generation B', async () => {
    const pending: Array<{ resolve: (value: { dispose(): void }) => void }> = [];
    let disposedLate = 0;
    const adapter: ViewerAdapter = () => new Promise((resolve) => pending.push({ resolve }));
    const { session } = setup(adapter, undefined, 15);
    const first = session.open(new Uint8Array([1]));
    await Promise.resolve(); await Promise.resolve();
    await expect(first).rejects.toThrow(/tiempo límite/);
    expect(states.at(-1)).toEqual({ kind: 'error', message: expect.stringMatching(/tiempo límite/) });
    pending[0].resolve({ dispose() { disposedLate++; } });
    await Promise.resolve(); await Promise.resolve();
    const second = session.open(new Uint8Array([2]));
    await Promise.resolve(); await Promise.resolve();
    pending[1].resolve({ dispose() {} });
    await second;
    expect(disposedLate).toBe(1);
    expect(states.at(-1)?.kind).toBe('ready');
  });
});
