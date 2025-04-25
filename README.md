# CombineWord - 合并多个 Word 文档

`CombineWord` 是一个用于将多个 Word 文档合并为一个单一 Word 文件的 JavaScript 类。它支持在合并文档之间插入分页符，并且能够处理文档的关系（rels）、样式和其他元数据的合并。该类基于 `JSZip` 库实现，支持以二进制格式合并多个 Word 文件。

## 安装

通过 npm 安装 `combine-word` 包：

```bash
npm install combine-word
```


## 使用示例

```javascript
const CombineWord = require('combine-word');
const fs = require('fs');

// 读取多个 Word 文件（以二进制格式）
const file1 = fs.readFileSync('file1.docx');
const file2 = fs.readFileSync('file2.docx');

// 创建一个 CombineWord 实例并合并文档
const combine = new CombineWord({}, [file1, file2]);

// 保存合并后的文档为 Blob 类型的文件
combine.save('nodebuffer', (fileData) => {
  fs.writeFileSync('combined.docx', fileData);
});
```

## 参数说明
- `files`: 要合并的 Word 文件数组。

- `options`: 配置选项。
  `options.pageBreak`: 多个文件拼合是否在合并时插入分页符。



