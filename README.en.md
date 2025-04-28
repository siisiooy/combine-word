# CombineWord - Merge Multiple Word Documents

`CombineWord` is a simple JavaScript library designed to merge multiple Word (.docx) documents into a single document. It supports merging document content, styles, headers, footers, and basic document information, and it can handle page breaks flexibly.



## Installation

Install the `combine-word` package via npm:

```bash
npm install combine-word
```



## Usage Examples

#### Using in Node.js

```javascript
const CombineWord = require("combine-word");
const fs = require("fs");

// Read multiple Word files (in binary format)
const file1 = fs.readFileSync("file1.docx");
const file2 = fs.readFileSync("file2.docx");

// Create a CombineWord instance and merge the documents
const combine = new CombineWord({ pageBreak: true, title: "Doc Title" }, [
  file1,
  file2,
]);

// Save the merged document as a node buffer
combine.save("nodebuffer", (fileData) => {
  fs.writeFileSync("combined.docx", fileData);
});
```

#### Using in the Browser

```html
<script src="combine-word.js"></script>
<script>
  // Create a CombineWord instance, passing in the files array (in ArrayBuffer format)
  const files = [/* ArrayBuffer format files */];
  const combineWord = new CombineWord({ pageBreak: true }, files);

  // Save the merged file
  combineWord.save('blob', (blob) => {
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'merged.docx';
    link.click();
  });
</script>

```



## API Documentation

#### `CombineWord(options, files) ` :

##### Parameters

- `files` (Array) : An array of Word files to merge.
- `options` (object) : Configuration options.
  - `options.pageBreak` (boolean) : *Whether to insert a page break when merging multiple files*.
  - `options.title` (string) : *The title of the merged document*.
  - `options.subject` (string) : *The subject of the merged document*.
  - `options.author` (string) : *The author of the merged document*.
  - `options.keywords` (string) : *The keywords of the merged document*.
  - `options.description` (string) : *The description of the merged document*.
  - `options.lastModifiedBy` (string) : *The last person who modified the merged document*.
  - `options.vision` (string) : *The version of the merged document*.



#### `save(type, callback)` :

##### Parameters

- `type` (string): The file type to save ( e.g., `'nodebuffer'` ).
- `callback` (function): The callback function that receives the merged file data.





## Compatibility

- **Node.js**：Supports running in a Node.js environment, suitable for file reading, merging, etc.
- **浏览器**：Supports usage in the browser environment via a bundled script file, offering a file download feature.



## Contributing

We welcome contributions to **Combine Word** ！You can improve this project by submitting issues or Pull Requests.



## License

This project is licensed under the [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0) 。

