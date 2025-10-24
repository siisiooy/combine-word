# CombineWord - 合并多个 Word 文档

Welcome to the **CombineWord** project! This project is a JavaScript library designed to merge multiple Word documents into one. You can choose to read the documentation in **Chinese** or **English** by clicking the links below.

## 选择语言 / Choose Language

- [中文文档](README.zh.md)
- [English Documentation](README.en.md)

---


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



## 许可证

该项目遵循 [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0) 许可证。

This project is licensed under the [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0).



## TODO

- [ ]. 三个及以上的文件合并可能存在问题，需要进一步测试和修复。
- [ ]. 页眉页脚的插入性内容过多时，可能会导致合并后提示需要修复文档，需要进行优化。
- [ ]. 更丰富的插入内容适配。
- [ ]. 选择不分节后的页数计算存在问题。

## VERSION 版本更新日志

### v1.0.6
- 修复合并后分节文档页眉页脚默认为连续上一节的问题

### v1.0.5
- 修复合并后关系文件引用异常问题

### v1.0.4
- 增加文档基础信息
- 增加默认合并文件进行分节

### v1.0.3
- 修复表格样式合并失效问题

### v1.0.2
- 修复合并图片丢失问题

### v1.0.1
- 添加了页脚页眉的合并功能
- 修复表格合并异常问题

### v1.0.0
- 初始版本，基础文件合并功能