import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const dependencies = [
  ['fflate', 'LICENSE', 'Archive compression and decompression, included in the browser app.'],
  ['@xmldom/xmldom', 'LICENSE', 'XML parsing, included in the browser app.'],
  ['pdfjs-dist', 'LICENSE', 'Mozilla PDF.js renders locally selected PDF proofs in the browser. Distributed under the Apache License 2.0.'],
  ['vite', 'LICENSE.md', 'Build tooling and generated browser helpers. The full upstream notice is retained.'],
];
let output = '# Third-party notices\n\nOfficeBridge is MIT licensed. The following notices cover the runtime libraries and build helpers distributed with the portable app. Development dependencies retain their licenses in their respective packages.\n\nNo Microsoft, Apple or other proprietary font binaries are distributed. Mentioning a font name does not grant a font license or install it.\n';
for (const [name, licenseFile, usage] of dependencies) {
  const dir = path.join(root, 'node_modules', name);
  const pkg = JSON.parse(await fs.readFile(path.join(dir, 'package.json'), 'utf8'));
  const license = await fs.readFile(path.join(dir, licenseFile), 'utf8');
  // Normalize trailing whitespace only; preserve every upstream notice word.
  output += `\n## ${name} ${pkg.version}\n\n${usage}\n\n${license.trim().replace(/[ \t]+$/gm, '')}\n`;
}
await fs.writeFile(path.join(root, 'THIRD_PARTY_NOTICES.md'), output);
console.log('Updated THIRD_PARTY_NOTICES.md from installed, locked dependencies.');
