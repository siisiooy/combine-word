const JSZip = require('jszip');
const { processRels, processStyles } = require('./lib/processor');
const { combineDocuments, combineRelationships, combineAppInfo, mergeFiles, updateCore } = require('./lib/combine');

/**
 * 合并多个 Word 文档的类。
 * 该类用于处理多个 Word 文件，将它们合并为一个单一的 Word 文件。
 * 
 * @param {object} [options={}] - 配置选项。
 * @param {boolean} [options.pageBreak=true] - 是否在合并的文档中插入分页符，默认为 `true`。设置为 `false` 时，文档内容会直接连接，没有分页符。
 * @param {string} [options.title] - 合并后文档的标题。
 * @param {string} [options.subject] - 合并后文档的主题。
 * @param {string} [options.author] - 合并后文档的作者。
 * @param {string} [options.keywords] - 合并后文档的关键词。
 * @param {string} [options.description] - 合并后文档的描述。
 * @param {string} [options.lastModifiedBy] - 最后修改该文档的用户。
 * @param {string} [options.vision] - 合并后文档的版本。
 * @param {Array} [files=[]] - 要合并的 Word 文件数组，每个文件应为二进制格式。可以是通过读取文件的方式获得的 Blob 或 ArrayBuffer。
 * 
 * @example
 * const combineWord = new CombineWord({
 *   pageBreak: true,
 *   title: "合并文档",
 *   author: "张三"
 * }, [file1, file2]);
 * combineWord.save('nodebuffer', (fileData) => {
 *   // 在此处理保存的文件数据
 * });
 */

// TODO: 三个文件合并的话，rel关系会出问题，ID重复且文件复制忽略了 customXml
class CombineWord {
  constructor(options = {}, files = []) {
    /**
     * 是否在文档之间插入分页符
     * @type {boolean}
     */
    this._pageBreak = typeof options.pageBreak !== 'undefined' ? !!options.pageBreak : true;

    /**
      * 校验并设置文档信息
      * @type {object}
      */
    this._docInfo = {
      title: this._validateString(options.title),
      subject: this._validateString(options.subject),
      author: this._validateString(options.author),
      keywords: this._validateString(options.keywords),
      description: this._validateString(options.description),
      lastModifiedBy: this._validateString(options.lastModifiedBy),
      vision: this._validateString(options.vision),
    };

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
   * 校验字段是否为字符串，若不是则抛出错误
   * @param {any} value - 需要验证的字段值
   * @param {string} fieldName - 字段名称，用于报错提示
   * @returns {string} - 字符串类型的字段值
   * @throws {Error} - 如果字段不是字符串，则抛出错误
   */
  _validateString(value, fieldName) {
    if (value !== undefined && value !== null && typeof value !== 'string') {
      throw new Error(`${fieldName} must be a string. Received: ${typeof value}`);
    }
    return value;  // 如果值为空，返回空字符串
  }

  /**
   * 预处理文件中的所有内容，包括关系（rels）、样式和文档的合并。
   * 
   * @param {Array<JSZip>} files - 包含要合并的 Word 文件的 JSZip 实例数组。
   * @private
   */
  preloadDocuments(files) {
    // 遍历每个文件，处理其关系（rels）和样式
    processRels(files);
    processStyles(files);

    // 合并文档内容和媒体文件
    combineDocuments(files, this._pageBreak);
    combineRelationships(files);
    combineAppInfo(files);

    updateCore(files, this._docInfo);
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
