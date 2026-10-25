import { analyzeOffice, repairOffice } from '../engine/index.js';
import { createDemo } from '../engine/demo.js';

self.addEventListener('message', async ({ data }) => {
  const { id, type, bytes, fileName, options = {} } = data;
  try {
    if (type === 'analyze') {
      const report = await analyzeOffice(bytes, fileName, options);
      self.postMessage({ id, result: report });
    } else if (type === 'repair') {
      const result = await repairOffice(bytes, fileName, options);
      self.postMessage({ id, result }, [result.bytes.buffer]);
    } else if (type === 'demo') {
      const result = await createDemo(data.format);
      self.postMessage({ id, result }, [result.buffer]);
    } else {
      throw new Error('不支持的处理请求。');
    }
  } catch (error) {
    self.postMessage({
      id,
      error: { message: error?.message || '无法处理此文件。', code: error?.code || 'PROCESSING_ERROR' },
    });
  }
});
