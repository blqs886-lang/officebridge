import assert from 'node:assert/strict';
import test from 'node:test';
import { fontPairHint, unsupportedFileMessage } from '../src/ui/input-help.js';

test('unsupported Apple and legacy files name the source app and the export format', () => {
  for (const [name, source, format] of [
    ['答辩.最终.KEY', 'Keynote', '.pptx'], ['季度.numbers', 'Numbers', '.xlsx'],
    ['旧报告.PPT', 'PowerPoint', '.pptx'], ['旧报表.xls', 'Excel', '.xlsx'],
  ]) {
    const message = unsupportedFileMessage(name);
    for (const expected of [name, source, format]) assert.ok(message.includes(expected));
  }
});

test('macro documents keep their macros and unknown extensions get a plain supported-format hint', () => {
  for (const name of ['模型.xlsm', '动画.PPTM']) {
    const message = unsupportedFileMessage(name);
    assert.match(message, /保留原件/);
    assert.doesNotMatch(message, /另存为|改名/);
  }
  assert.match(unsupportedFileMessage('讲义.pdf'), /原稿/);
  assert.match(unsupportedFileMessage('预算.csv'), /编码/);
  for (const name of ['工作簿', '测试.constructor', '文稿.key.zip']) {
    assert.match(unsupportedFileMessage(name), /\.pptx.*\.xlsx/);
  }
});

test('font guidance explains known family pairs and leaves unknown custom mappings unlabelled', () => {
  assert.match(fontPairHint('PingFang SC', 'Microsoft YaHei'), /苹方.*微软雅黑/);
  assert.match(fontPairHint('Menlo', 'Courier New'), /错列/);
  assert.match(fontPairHint('Helvetica Neue', 'Arial'), /换行/);
  assert.equal(fontPairHint('Menlo', 'Arial'), '');
  assert.equal(fontPairHint('My font', 'My other font'), '');
});
