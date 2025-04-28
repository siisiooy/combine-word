# CombineWord - 合并多个 Word 文档

`CombineWord` 是一个简单的 JavaScript 库，旨在将多个 Word 文档（.docx 文件）合并成一个单一的文档。它支持合并文档的内容、样式、页眉、页脚和文档基础信息，并且可以灵活地处理分页符。



## 安装

通过 npm 安装 `combine-word` 包：

```bash
npm install combine-word
```



## 使用示例

#### 在 Node.js 中使用

```javascript
const CombineWord = require("combine-word");
const fs = require("fs");

// 读取多个 Word 文件
const file1 = fs.readFileSync("file1.docx");
const file2 = fs.readFileSync("file2.docx");

// 创建一个 CombineWord 实例并合并文档
const combine = new CombineWord({ pageBreak: true, title: "Doc Title" }, [
  file1,
  file2,
]);

// 保存合并后的文档为 nodebuffer 类型的文件
combine.save("nodebuffer", (fileData) => {
  fs.writeFileSync("combined.docx", fileData);
});
```

#### 在浏览器中使用

```html
<script src="combine-word.js"></script>
<script>
  // 创建 CombineWord 实例，传入需要合并的文件数组（以 ArrayBuffer 格式）
  const files = [/* ArrayBuffer 格式的文件 */];
  const combineWord = new CombineWord({ pageBreak: true }, files);

  // 保存合并后的文件
  combineWord.save('blob', (blob) => {
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'merged.docx';
    link.click();
  });
</script>

```



## API说明

#### `CombineWord(options, files) ` :

##### 参数

- `files` (Array) : 要合并的 Word 文件数组。
- `options` (object) : 配置选项。
  - `options.pageBreak` (boolean) : *多个文件拼合是否在合并时插入分页符*。
  - `options.title` (string) : *合并后文档的标题。*
  - `options.subject` (string) : *合并后文档的主题。*
  - `options.author` (string) : *合并后文档的作者。*
  - `options.keywords` (string) : *合并后文档的标记关键词。*
  - `options.description` (string) : *合并后文档的备注。*
  - `options.lastModifiedBy` (string) : *合并后文档的最后一次保存者。*
  - `options.vision` (string) : *合并后文档的修订号。*



#### `save(type, callback)` :

##### 参数

- `type` (string): 保存文件的类型 ( 如 `'nodebuffer'` )。
- `callback` (function): 合并后文件的回调函数，返回文件数据。





## 兼容性

- **Node.js**：支持在 Node.js 环境中运行，适用于文件读取、合并等操作。
- **浏览器**：支持在浏览器环境中使用，通过打包后的脚本文件，提供文件下载功能。



## 贡献

欢迎对 **Combine Word** 进行贡献！你可以通过提交 issues 或 Pull Requests 来改进这个项目。



## 许可证

该项目遵循 [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0) 。

