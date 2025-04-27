const JSZip = require('jszip');
const { processRels, processStyles } = require('./lib/processor');
const { combineDocuments, combineRelationships, combineAppInfo, mergeFiles, updateModifiedDate } = require('./lib/combine');

/**
 * 合并多个 Word 文档的类。
 * 该类用于处理多个 Word 文件，将它们合并为一个单一的 Word 文件。
 * 
 * @class CombineWord
 * @param {object} [options={}] - 配置选项。
 * @param {boolean} [options.pageBreak=true] - 是否在合并的文档中插入分页符，默认为 `true`。
 * @param {Array} [files=[]] - 要合并的 Word 文件数组，每个文件应为二进制格式。
 */
class CombineWord {
  constructor(options = {}, files = []) {
    /**
     * 是否在文档之间插入分页符
     * @type {boolean}
     */
    this._pageBreak = typeof options.pageBreak !== 'undefined' ? !!options.pageBreak : true;

    /**
     * 包含所有输入文件的 JSZip 实例
     * @type {Array<JSZip>}
     */
    this._files = files.map(file => new JSZip(file));

    /**
     * 合并后的 Word 文档
     * @type {JSZip|null}
     */
    this.docx = null;

    // 如果传入文件，则加载和处理文档
    if (this._files.length > 0) {
      this.preloadDocuments(this._files);
    }
  }

  /**
   * 预处理文件中的所有内容，包括关系（rels）、样式和文档的合并。
   * 
   * @param {Array<JSZip>} files - 包含要合并的 Word 文件的 JSZip 实例数组。
   * @private
   */
  preloadDocuments(files) {
    let relIdOffset = -1;
    let styleIdOffset = -1;
    let styleNumIdOffset = 0;
    let relationIndex = { headerIndex: 1, footerIndex: 1, mediaIndex: 1 };

    // 遍历每个文件，处理其关系（rels）和样式
    for (const zip of files) {
      const { nextRelId, fileIndexMap } = processRels(zip, relIdOffset, relationIndex);
      const { nextStyleId, styleNumStart } = processStyles(zip, styleIdOffset, styleNumIdOffset);

      // 更新偏移量和索引
      relationIndex = fileIndexMap;
      relIdOffset = nextRelId;
      styleIdOffset = nextStyleId;
      styleNumIdOffset = styleNumStart;
    }

    // 合并文档内容、关系和媒体文件
    combineDocuments(files, this._pageBreak);
    combineRelationships(files);
    combineAppInfo(files);

    updateModifiedDate(files);
    this.docx = mergeFiles(files); // 合并所有文件并生成最终的 ZIP 文件
  }

  /**
   * 将合并后的 Word 文档保存为指定类型的文件。
   * 
   * @param {string} type - 要保存的文件类型，通常为 'blob' 或 'base64'。
   * @param {function} callback - 回调函数，接收生成的文件数据作为参数。
   */
  save(type, callback) {
    const zip = this.docx;

    // 生成并回调文件数据
    callback(zip.generate({
      type: type,
      compression: "DEFLATE", // 设置压缩算法
      compressionOptions: { level: 4 } // 设置压缩级别
    }));
  }
}

module.exports = CombineWord;
