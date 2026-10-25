import test from 'node:test';
import assert from 'node:assert/strict';
import { deflateSync, zipSync, Inflate } from 'fflate';
import { readArchive, crc32 } from '../src/engine/archive.js';

const encode = text => new TextEncoder().encode(text);

function locateCentral(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocd = bytes.length - 22;
  assert.equal(view.getUint32(eocd, true), 0x06054b50);
  return view.getUint32(eocd + 16, true);
}

function mixedPayload(size = 3 * 1024 * 1024 + 17) {
  const bytes = new Uint8Array(size);
  let state = 12345;
  for (let i = 0; i < bytes.length; i++) {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    bytes[i] = i % 8 === 0 ? state >>> 24 : 65 + i % 3;
  }
  return bytes;
}

// A small valid ZIP assembled independently of the production ZIP reader. It
// exercises streaming-writer descriptors with and without their optional marker.
function descriptorZip({ signature = true, compressed = true } = {}) {
  const name = encode('ppt/media/synthetic.bin');
  const original = encode('An entirely synthetic document part. '.repeat(4));
  const data = compressed ? deflateSync(original) : original;
  const descriptorLength = signature ? 16 : 12;
  const central = 30 + name.length + data.length + descriptorLength;
  const end = central + 46 + name.length;
  const bytes = new Uint8Array(end + 22);
  const view = new DataView(bytes.buffer);
  const set16 = (offset, value) => view.setUint16(offset, value, true);
  const set32 = (offset, value) => view.setUint32(offset, value, true);
  set32(0, 0x04034b50); set16(4, 20); set16(6, 0x0808); set16(8, compressed ? 8 : 0);
  set16(26, name.length); bytes.set(name, 30); bytes.set(data, 30 + name.length);
  let descriptor = 30 + name.length + data.length;
  if (signature) { set32(descriptor, 0x08074b50); descriptor += 4; }
  set32(descriptor, crc32(original)); set32(descriptor + 4, data.length); set32(descriptor + 8, original.length);
  set32(central, 0x02014b50); set16(central + 4, 20); set16(central + 6, 20);
  set16(central + 8, 0x0808); set16(central + 10, compressed ? 8 : 0);
  set32(central + 16, crc32(original)); set32(central + 20, data.length); set32(central + 24, original.length);
  set16(central + 28, name.length); bytes.set(name, central + 46);
  set32(end, 0x06054b50); set16(end + 8, 1); set16(end + 10, 1);
  set32(end + 12, 46 + name.length); set32(end + 16, central);
  return { bytes, original };
}

test('validated EOCD remains authoritative when the ZIP comment contains a false EOCD signature', () => {
  const original = encode('safe payload');
  const base = zipSync({ 'data.bin': original });
  const comment = new Uint8Array(96);
  const commentView = new DataView(comment.buffer);
  // The false record does not reach the real end of the comment and must not be
  // chosen by a decompressor that rescans without checking the comment length.
  commentView.setUint32(12, 0x06054b50, true);
  commentView.setUint16(20, 40, true);
  commentView.setUint16(22, 40, true);
  commentView.setUint32(28, 0xfffffff0, true);
  const bytes = new Uint8Array(base.length + comment.length);
  bytes.set(base); bytes.set(comment, base.length);
  new DataView(bytes.buffer).setUint16(base.length - 2, comment.length, true);
  const archive = readArchive(bytes);
  assert.deepEqual(archive.entries.get('data.bin'), original);
  assert.equal(archive.entries.size, 1);
});

test('false uncompressed size cannot turn a longer DEFLATE stream into an accepted truncated part', () => {
  const original = encode('ABCD'.repeat(1000));
  const bytes = zipSync({ 'data.bin': original }, { level: 6 });
  const view = new DataView(bytes.buffer);
  const central = locateCentral(bytes);
  const prefix = original.subarray(0, 5);
  // Both headers and CRC agree on the prefix; only the actual compressed stream
  // reveals that the declared size is false.
  view.setUint32(14, crc32(prefix), true);
  view.setUint32(22, prefix.length, true);
  view.setUint32(central + 16, crc32(prefix), true);
  view.setUint32(central + 24, prefix.length, true);
  assert.throws(() => readArchive(bytes), /校验失败|压缩数据损坏/);
});

test('valid multi-chunk DEFLATE preserves a several-MiB part below the expansion-ratio limit', () => {
  const original = mixedPayload();
  const bytes = zipSync({ 'ppt/media/synthetic.bin': original }, { level: 6 });
  const view = new DataView(bytes.buffer);
  const central = locateCentral(bytes);
  const compressedSize = view.getUint32(central + 20, true);
  assert.ok(compressedSize > 8 * 1024);
  assert.ok(original.length / compressedSize < 300);
  const archive = readArchive(bytes);
  assert.deepEqual(archive.entries.get('ppt/media/synthetic.bin'), original);
  assert.equal(archive.expandedBytes, original.length);
});

test('false-size long compressed streams stop before subsequent chunks are fed to the decoder', () => {
  const original = mixedPayload();
  const bytes = zipSync({ 'data.bin': original }, { level: 6 });
  const view = new DataView(bytes.buffer);
  const central = locateCentral(bytes);
  const compressedSize = view.getUint32(central + 20, true);
  const prefix = original.subarray(0, 16);
  view.setUint32(14, crc32(prefix), true);
  view.setUint32(22, prefix.length, true);
  view.setUint32(central + 16, crc32(prefix), true);
  view.setUint32(central + 24, prefix.length, true);

  // Observe decoder input rather than assert a machine-dependent time limit.
  // The same stream is a valid 3 MiB payload in the preceding test.
  const originalPush = Inflate.prototype.push;
  let fedBytes = 0;
  let pushes = 0;
  Inflate.prototype.push = function (chunk, final) {
    fedBytes += chunk.length;
    pushes += 1;
    return originalPush.call(this, chunk, final);
  };
  try {
    assert.throws(() => readArchive(bytes), /压缩数据损坏/);
    assert.equal(pushes, 1);
    assert.ok(fedBytes <= 8 * 1024);
    assert.ok(fedBytes < compressedSize / 10, `Consumed ${fedBytes} of ${compressedSize} compressed bytes`);
  } finally { Inflate.prototype.push = originalPush; }
});

test('large stored blocks inside a valid DEFLATE stream are not mistaken for stalled decompression', () => {
  const original = new Uint8Array(256 * 1024 + 7);
  let state = 12345;
  for (let i = 0; i < original.length; i++) {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    original[i] = state >>> 24;
  }
  const archive = readArchive(zipSync({ 'noise.bin': original }, { level: 6 }));
  assert.deepEqual(archive.entries.get('noise.bin'), original);
});

test('long trailing garbage cannot force the streaming decoder to repeatedly retain the entire tail', () => {
  const original = encode('small valid content');
  const base = zipSync({ 'data.bin': original }, { level: 6 });
  const central = locateCentral(base);
  const trailer = new Uint8Array(512 * 1024);
  const bytes = new Uint8Array(base.length + trailer.length);
  bytes.set(base.subarray(0, central));
  bytes.set(trailer, central);
  bytes.set(base.subarray(central), central + trailer.length);
  const view = new DataView(bytes.buffer);
  const compressedSize = new DataView(base.buffer).getUint32(central + 20, true) + trailer.length;
  view.setUint32(18, compressedSize, true);
  view.setUint32(central + trailer.length + 20, compressedSize, true);
  view.setUint32(bytes.length - 6, central + trailer.length, true);
  const originalPush = Inflate.prototype.push;
  let fedBytes = 0;
  Inflate.prototype.push = function (chunk, final) {
    fedBytes += chunk.length;
    return originalPush.call(this, chunk, final);
  };
  try {
    assert.throws(() => readArchive(bytes), /解压超限/);
    assert.ok(fedBytes < compressedSize / 2);
  } finally { Inflate.prototype.push = originalPush; }
});

test('empty content with a valid DEFLATE stream remains valid', () => {
  const archive = readArchive(zipSync({ 'empty.bin': new Uint8Array(0) }, { level: 6 }));
  assert.equal(archive.entries.get('empty.bin').length, 0);
});

test('a method-8 entry cannot omit the DEFLATE stream even when its declared content is empty', () => {
  const bytes = zipSync({ 'empty.bin': new Uint8Array(0) }, { level: 0 });
  const view = new DataView(bytes.buffer);
  const central = locateCentral(bytes);
  view.setUint16(8, 8, true);
  view.setUint16(central + 10, 8, true);
  assert.throws(() => readArchive(bytes), /压缩数据损坏/);
});

test('stream completion also rejects declared content longer than the actual output', () => {
  const bytes = zipSync({ 'data.bin': encode('bounded stream') }, { level: 6 });
  const view = new DataView(bytes.buffer);
  const central = locateCentral(bytes);
  const expanded = view.getUint32(central + 24, true);
  view.setUint32(22, expanded + 17, true);
  view.setUint32(central + 24, expanded + 17, true);
  assert.throws(() => readArchive(bytes), /压缩数据损坏|校验失败/);
});

test('case-insensitive part-name collisions are rejected before decoding entries', () => {
  const bytes = zipSync({ 'ppt/slides/slide1.xml': encode('one'), 'PPT/slides/SLIDE1.xml': encode('two') });
  assert.throws(() => readArchive(bytes), /重复|大小写冲突/);
});

test('percent-decoded part-name aliases cannot silently select a different resource', () => {
  const bytes = zipSync({ 'ppt/media/icon.svg': encode('one'), 'ppt/media/%69con.svg': encode('two') });
  assert.throws(() => readArchive(bytes), /重复|大小写冲突/);
});

test('URI-encoded traversal and separators are rejected as unsafe archive paths', () => {
  for (const name of ['ppt/%2e%2e/evil.xml', 'ppt/%5cserver.xml', '%2fppt/evil.xml', 'ppt/%2e/evil.xml']) {
    const bytes = zipSync({ [name]: encode('synthetic') });
    assert.throws(() => readArchive(bytes), /不安全/);
  }
});

test('valid compressed ZIP data descriptors with and without markers are accepted', () => {
  for (const signature of [false, true]) {
    const { bytes, original } = descriptorZip({ signature });
    const archive = readArchive(bytes);
    assert.deepEqual(archive.entries.get('ppt/media/synthetic.bin'), original);
  }
});

test('valid stored data descriptors remain readable without treating descriptor bytes as content', () => {
  const { bytes, original } = descriptorZip({ compressed: false });
  const archive = readArchive(bytes);
  assert.deepEqual(archive.entries.get('ppt/media/synthetic.bin'), original);
  assert.equal(archive.expandedBytes, original.length);
});

test('stored entries with contradictory compressed and expanded lengths are rejected', () => {
  const bytes = zipSync({ 'data.bin': encode('payload') }, { level: 0 });
  const view = new DataView(bytes.buffer);
  const central = locateCentral(bytes);
  view.setUint32(22, 6, true);
  view.setUint32(central + 24, 6, true);
  assert.throws(() => readArchive(bytes), /大小信息不一致/);
});

test('reader respects a Uint8Array view without reading unrelated prefix and suffix bytes', () => {
  const original = encode('payload');
  const zip = zipSync({ 'data.bin': original });
  const buffer = new Uint8Array(zip.length + 80);
  buffer.fill(0xa5); buffer.set(zip, 37);
  const archive = readArchive(buffer.subarray(37, 37 + zip.length));
  assert.deepEqual(archive.entries.get('data.bin'), original);
});
