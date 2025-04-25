const fs = require('fs-extra');
const path = require('path');
const CombineWord = require('../index'); // 确保你的 CombineWord 文件路径正确
const { file } = require('jszip/lib/object');

// 读取 example 目录下的 Word 文件
const inputDir = path.resolve(__dirname, "./example"); // 需要合并的 Word 文件目录
const outputDir = path.resolve(__dirname, "./output"); // 结果存储在 output 目录
const outputFilePath = path.join(outputDir, "merged.docx");

// 确保 output 目录存在
fs.ensureDirSync(outputDir);

// 读取 example 目录下的 .docx 文件
const files = fs.readdirSync(inputDir)
  .filter(file => file.endsWith(".docx"))
  .map(file => path.join(inputDir, file));

if (files.length < 2) {
  console.error("请提供至少两个 Word 文件进行合并！");
  process.exit(1);
}

// 读取文件为 ArrayBuffer
const fileBuffers = files.map(file => fs.readFileSync(file).buffer);

// 合并文件
const docx = new CombineWord({ pageBreak: true }, fileBuffers);
docx.save("nodebuffer", buffer => {
  fs.writeFileSync(outputFilePath, buffer);
  console.log(`合并完成！文件保存至: ${outputFilePath}`);
});
