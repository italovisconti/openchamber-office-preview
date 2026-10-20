import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

test('DOCX iframe has restrictive resource CSP and embeds package CSS inline', async () => {
  const html = await readFile(new URL('../../editor/docx.html', import.meta.url), 'utf8');
  const css = await readFile(new URL('../../editor/docx.css', import.meta.url), 'utf8');
  assert.match(html, /http-equiv="Content-Security-Policy"/);
  for (const directive of ["img-src data: blob:", "font-src 'none'", "media-src 'none'", "connect-src 'none'", "frame-src 'none'", "object-src 'none'", "base-uri 'none'", "style-src 'unsafe-inline'"]) assert.ok(html.includes(directive), `missing CSP directive: ${directive}`);
  assert.ok(html.includes(`<style>${css}</style>`));
  assert.doesNotMatch(html, /<link[^>]+stylesheet/);
});
