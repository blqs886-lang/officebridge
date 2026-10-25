import test from 'node:test';
import assert from 'node:assert/strict';
import { makePortable, OFFLINE_POLICY } from '../scripts/build-portable.mjs';

test('portable packaging preserves JS dollar syntax and neutralizes HTML closing tags', async () => {
  const js = 'const text = "$& $` $\' </script>";\n//# sourceMappingURL=private.map';
  const css = 'a:after {content:"$& </style>"}';
  const html = await makePortable({
    html: '<!DOCTYPE html><html><head><meta charset="UTF-8"><link rel="icon" href="./favicon.svg"><script type="module" src="./app.js"></script><link rel="stylesheet" href="./app.css"></head><body><main></main></body></html>',
    readAsset: async name => name.endsWith('.js') ? js : css,
    favicon: Buffer.from('<svg/>'), notices: 'MIT <script>not executable</script> & contributors',
  });
  assert.ok(html.includes('const text = "$& $` $\' <\\/script>";'));
  assert.equal((html.match(/<script\b/g) || []).length, 1);
  assert.equal((html.match(/<\/script>/g) || []).length, 1);
  assert.ok(html.includes('a:after {content:"$& <\\/style>"}'));
  assert.ok(html.includes(OFFLINE_POLICY));
  assert.ok(html.includes('MIT &lt;script&gt;not executable&lt;/script&gt; &amp; contributors'));
  assert.ok(!html.includes('private.map'));
  assert.ok(html.includes('data:image/svg+xml;base64,'));
});

test('portable packager refuses an unresolved external dependency', async () => {
  await assert.rejects(makePortable({ html: '<link rel="preload" href="https://example.invalid/lib.js">', readAsset: async () => '', favicon: Buffer.from(''), notices: '' }), /external resource/);
});
