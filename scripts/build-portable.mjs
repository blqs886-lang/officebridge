import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export const OFFLINE_POLICY = "default-src 'none'; script-src 'unsafe-inline' blob: data:; worker-src blob: data:; style-src 'unsafe-inline'; img-src data: blob:; connect-src 'none'; base-uri 'none'; form-action 'none'";

export async function makePortable({ html, readAsset, favicon, notices }) {
  for (const match of [...html.matchAll(/<script\b[^>]*\bsrc="([^"]+)"[^>]*><\/script>/g)]) {
    const source = await readAsset(match[1]);
    const safe = source.replace(/<\/script/gi, '<\\/script').replace(/^\/\/# sourceMappingURL=.*$/gm, '');
    // A callback is essential: replacement strings interpret dollar syntax in
    // JavaScript bundles as replacement metacharacters and corrupt the app.
    html = html.replace(match[0], () => `<script type="module">${safe}</script>`);
  }
  for (const match of [...html.matchAll(/<link\b[^>]*rel="stylesheet"[^>]*href="([^"]+)"[^>]*>/g)]) {
    const css = await readAsset(match[1]);
    html = html.replace(match[0], () => `<style>${css.replace(/<\/style/gi, '<\\/style')}</style>`);
  }
  html = html.replace(/href="(?:\.\/|\/)favicon\.svg"/, () => `href="data:image/svg+xml;base64,${favicon.toString('base64')}"`);
  html = html.replace('<meta charset="UTF-8"', () => `<meta http-equiv="Content-Security-Policy" content="${OFFLINE_POLICY}">\n    <meta charset="UTF-8"`);
  const structure = html.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '').replace(/<style>[\s\S]*?<\/style>/gi, '');
  if (/<(?:script|link)[^>]*(?:src|href)="(?!data:|blob:|#)/i.test(structure)) throw new Error('Portable build still contains an external resource.');
  // Keep licensing inside the single file as well as in the release ZIP.
  const escaped = notices.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
  html = html.replace('</body>', () => `<template id="officebridge-license-notices"><pre>${escaped}</pre></template>\n</body>`);
  return html;
}

export async function buildPortable(root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')) {
  const html = await makePortable({
    html: await fs.readFile(path.join(root, 'dist/index.html'), 'utf8'),
    readAsset: resource => fs.readFile(path.join(root, 'dist', resource), 'utf8'),
    favicon: await fs.readFile(path.join(root, 'public/favicon.svg')),
    notices: `${await fs.readFile(path.join(root, 'LICENSE'), 'utf8')}\n\n${await fs.readFile(path.join(root, 'THIRD_PARTY_NOTICES.md'), 'utf8')}`,
  });
  await fs.mkdir(path.join(root, 'release'), { recursive: true });
  const output = path.join(root, 'release/OfficeBridge.html');
  await fs.writeFile(output, html);
  console.log(`Portable offline app: ${output} (${Buffer.byteLength(html)} bytes)`);
  return output;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) await buildPortable();
