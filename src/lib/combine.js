const { DOMParser, XMLSerializer } = require('xmldom');



/**
 * 修改 _files[0] 中 coreProperties.xml 的 dcterms:modified 为当前时间。
 * 
 * @param {Array} _files - 包含多个 Word 文件的 ZIP 文件数组。
 */
function updateCore(_files, docInfo = {
  title: '',
  subject: '',
  author: '',
  vision: '',
  keywords: '',
  description: '',
  lastModifiedBy: '',
}) {
  const corePropertiesPath = 'docProps/core.xml';  // coreProperties 文件路径
  const xml = _files[0].file(corePropertiesPath).asText();  // 获取文件内容

  if (!xml) {
    console.warn('coreProperties.xml not found in the first file.');
    return;  // 如果文件不存在，返回
  }
  // 获取当前时间并格式化为 W3CDTF 格式 (例如：2025-04-25T08:24:00Z)
  const currentDate = new Date();
  currentDate.setMilliseconds(0);  // 去除毫秒部分

  // 获取 UTC 时间并手动格式化为 "YYYY-MM-DDTHH:MM:SSZ"
  const formattedDate = currentDate.toISOString().slice(0, 17) + '00Z';  // 获取到秒并加上 Z 表示 UTC 时间

  const updateCoreProperties = {
    'dc:title': docInfo.title,
    'dc:subject': docInfo.subject,
    'dc:creator': docInfo.author,
    'cp:keywords': docInfo.keywords,
    'dc:description': docInfo.description,
    'cp:lastModifiedBy': docInfo.lastModifiedBy,
    'cp:revision': docInfo.vision,
    'dcterms:modified': formattedDate,
  }

  // 使用 DOMParser 解析 XML
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');
  for (const [key, value] of Object.entries(updateCoreProperties)) {
    updateOrCreateElement(doc, key, value);  // 更新或创建元素
  }

  // 将修改后的 XML 转为字符串
  const builder = new XMLSerializer();
  const updatedXml = builder.serializeToString(doc);

  // 将更新后的 coreProperties.xml 写回文件
  _files[0].file(corePropertiesPath, updatedXml);
}

/**
 * 检查文档中是否存在指定的元素，如果存在则更新其文本内容，否则创建新元素并添加到文档中。
 * 
 * @param {Document} doc - 需要更新的 XML 或 HTML 文档对象。
 * @param {string} tagName - 要查找或创建的元素的标签名。
 * @param {string} textContent - 要设置到元素中的文本内容。
 * 
 * @returns {void}
 */
function updateOrCreateElement(doc, tagName, textContent) {
  if (textContent === undefined || textContent === null) {
    return;  // 如果文本内容为空，直接返回
  }
  let element = doc.getElementsByTagName(tagName)[0];
  if (element) {
    element.textContent = textContent;
    return;  // 如果元素存在，更新其文本内容并返回
  }
  element = doc.createElement(tagName);
  element.textContent = textContent;
  doc.documentElement.appendChild(element);  // 将新元素添加到文档
}

/**
 * 处理 Word 文档中的自定义 XML 文件（app.xml），并提取应用程序信息。
 * 
 * @param {object} zip - 包含 Word 文档的 ZIP 文件对象。
 * @param {object} [appInfo={ Pages:0, Words:0, Characters:0, Lines:0, CharactersWithSpaces:0 }] - 默认值，包含要提取的应用信息。
 * @returns {object} - 返回更新后的 appInfo 对象，包含文档中的相关统计信息。
 */
function combineAppInfo(_files) {
  // TODO: 处理 app.xml 中的自定义 XML 文件时，空标签会变为单标签，待修复
  const appPath = 'docProps/app.xml';  // 自定义 XML 文件路径
  const appInfo = { Pages: 0, Words: 0, Characters: 0, Lines: 0, CharactersWithSpaces: 0 };
  // 遍历每个文件
  for (let i = 0; i < _files.length; i++) {
    const zip = _files[i];
    const xml = zip.file(appPath).asText();

    // 使用 DOMParser 解析 XML
    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'application/xml');

    // 使用 DOM 查询获取相应的节点值
    const pages = doc.getElementsByTagName('Pages')[0]?.textContent || 0;
    const words = doc.getElementsByTagName('Words')[0]?.textContent || 0;
    const characters = doc.getElementsByTagName('Characters')[0]?.textContent || 0;
    const lines = doc.getElementsByTagName('Lines')[0]?.textContent || 0;
    const charactersWithSpaces = doc.getElementsByTagName('CharactersWithSpaces')[0]?.textContent || 0;

    // 将提取到的值赋给 appInfo 对象
    appInfo.Pages += parseInt(pages, 10);
    appInfo.Words += parseInt(words, 10);
    appInfo.Characters += parseInt(characters, 10);
    appInfo.Lines += parseInt(lines, 10);
    appInfo.CharactersWithSpaces += parseInt(charactersWithSpaces, 10);
  }
  const xml = _files[0].file(appPath).asText();
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');

  // 获取根元素 <Properties>，并更新各个统计信息
  const properties = doc.getElementsByTagName('Properties')[0];

  // 更新 appInfo 内容
  properties.getElementsByTagName('Pages')[0].textContent = appInfo.Pages.toString();
  properties.getElementsByTagName('Words')[0].textContent = appInfo.Words.toString();
  properties.getElementsByTagName('Characters')[0].textContent = appInfo.Characters.toString();
  properties.getElementsByTagName('Lines')[0].textContent = appInfo.Lines.toString();
  properties.getElementsByTagName('CharactersWithSpaces')[0].textContent = appInfo.CharactersWithSpaces.toString();

  // 将更新后的 XML 内容序列化为字符串
  const builder = new XMLSerializer();
  const updatedXml = builder.serializeToString(doc);

  // 将更新后的 app.xml 文件写回到第一个文件
  _files[0].file(appPath, updatedXml);
}


/**
 * 合并多个 Word 文档的内容（document.xml）。
 * 遍历每个文档，将每个文档的内容合并到第一个文档中，保持结构一致。
 * 
 * @param {Array} _files - 包含多个 Word 文档的 ZIP 文件数组。
 */
function combineDocuments(_files, _pageBreak) {
  const docPath = 'word/document.xml';
  const documents = [];
  let docXml;

  // 遍历每个文件
  for (let i = 0; i < _files.length; i++) {
    const zip = _files[i];
    const isLast = !!(i === _files.length - 1);
    const isBaseFile = !!(i === 0);

    const xml = zip.file(docPath).asText();
    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'application/xml');

    // 获取<w:body>标签
    const body = doc.getElementsByTagName('w:body')[0];

    if (isBaseFile) {
      // 只处理第一个文件，保存清除 <w:body> 标签的内容
      docXml = doc.cloneNode(true); // 克隆 <w:body>，以便修改
      const bodyDoc = docXml.getElementsByTagName('w:body')[0];
      // 清空 w:body 标签的内容
      while (bodyDoc.firstChild) {
        bodyDoc.removeChild(bodyDoc.firstChild);
      }
    }

    // 获取每个文件 body 内所有子元素并存储到 documents 数组中
    const bodyClone = body.cloneNode(true); // 克隆 body
    const sectPrElements = bodyClone.getElementsByTagName('w:sectPr');

    // 对于每个文件的 <w:sectPr> 元素
    Array.from(sectPrElements).forEach(sectPr => {

      // <w:lastRenderedPageBreak />
      if (sectPr.parentNode === bodyClone) { // 确保 sectPr 的父级是 <w:body>
        if (!isLast) {
          sectPr.parentNode.removeChild(sectPr);
          const doc = bodyClone.ownerDocument; // 通过 bodyClone 获取对应文档对象

          // 替换 <w:sectPr> 为 <w:p><w:rPr>  <w:sectPr>  </w:rPr></w:p>
          const newSectPr = doc.createElement('w:p');
          // w:pPr
          const pPr = doc.createElement('w:pPr');
          const rPr = doc.createElement('w:rPr');
          pPr.appendChild(rPr);
          pPr.appendChild(sectPr);

          newSectPr.appendChild(pPr);

          if (!!_pageBreak) {
            const r = doc.createElement('w:r');
            const br = doc.createElement('w:br');
            br.setAttribute('w:type', 'page');
            r.appendChild(br);
            newSectPr.appendChild(r);
            bodyClone.appendChild(newSectPr);
          }
        }
      }
    });
    documents.push(bodyClone); // 将克隆的 body 存储在 documents 数组中
  }

  // 获取 docXml 中的 <w:body> 标签，并将 documents 数组的内容添加到其中
  const bodyNode = docXml.getElementsByTagName('w:body')[0];

  // 将 documents 中的所有元素的子元素添加到 docXml 的 body 中
  documents.forEach(docElement => {
    // 遍历 docElement 的子节点，将每个子节点添加到 bodyNode 中
    Array.from(docElement.childNodes).forEach(childNode => {
      bodyNode.appendChild(childNode);
    });
  });

  // 结合 docXml 和 documents 构成新的 xml
  const serializer = new XMLSerializer();
  const updatedXml = serializer.serializeToString(docXml); // 将更新后的 docXml 转换为字符串
  _files[0].file(docPath, updatedXml);
}

/**
 * 合并多个 Word 文档的关系文件（document.xml.rels）。
 * 遍历每个文件的关系文件，将每个文档的关系合并到第一个文件中。
 * 
 * @param {Array} _files - 包含多个 Word 文档的 ZIP 文件数组。
 */
function combineRelationships(_files) {
  let docRelsXml;
  let docTypeXml;
  let docFontTableXml;
  const relsPath = "word/_rels/document.xml.rels";
  const typePath = '[Content_Types].xml';
  const fontTablePath = "word/fontTable.xml";
  // 遍历每个文件
  for (const zip of _files) {
    const relsXml = zip.file(relsPath).asText();
    const parser = new DOMParser();
    const relsDoc = parser.parseFromString(relsXml, 'application/xml');

    // 获取所有的 <Relationship> 元素
    const relationships = relsDoc.getElementsByTagName('Relationship');

    if (!docRelsXml) {
      // 第一次遇到文件，保存其 document.xml.rels 的内容作为基
      docRelsXml = relsDoc;
    } else {
      // 获取第一个文件的 <Relationships> 元素
      const relationshipsElement = docRelsXml.getElementsByTagName('Relationships')[0];
      // 获取现有的 Target 属性集合，以便后续进行检查
      const existingTargets = Array.from(relationshipsElement.getElementsByTagName('Relationship'))
        .map(r => r.getAttribute('Target'));

      // 遍历当前文件的所有 <Relationship> 元素并将其添加到第一个文件的 <Relationships> 中
      for (let i = 0; i < relationships.length; i++) {
        const relationship = relationships[i];
        const target = relationship.getAttribute('Target');

        // 检查是否已有相同的 Target 属性
        if (!existingTargets.includes(target)) {
          // 如果没有重复的 Target，克隆并添加
          relationshipsElement.appendChild(relationship.cloneNode(true));
          existingTargets.push(target); // 更新现有的 Target 列表
        }
      }
    }

    const typeXml = zip.file(typePath).asText();
    const typeDoc = parser.parseFromString(typeXml, 'application/xml');

    // 获取所有的 <Override> 元素
    const overrides = typeDoc.getElementsByTagName('Override');
    const extensions = typeDoc.getElementsByTagName('Default');

    if (!docTypeXml) {
      docTypeXml = typeDoc;
    } else {
      // 获取第一个文件的 <Types> 元素
      const typesElement = docTypeXml.getElementsByTagName('Types')[0];
      const typesTargets = Array.from(typesElement.getElementsByTagName('Override')).map(r => r.getAttribute('PartName'));

      for (let i = 0; i < overrides.length; i++) {
        const overridesEle = overrides[i];
        const partName = overridesEle.getAttribute('PartName');

        if (!typesTargets.includes(partName)) {
          typesElement.appendChild(overridesEle.cloneNode(true));
          typesTargets.push(partName);
        }
      }


      const defaultExtensions = Array.from(typesElement.getElementsByTagName('Default')).map(r => r.getAttribute('Extension'));
      for (let i = 0; i < extensions.length; i++) {
        const extensionEle = extensions[i];
        const Extension = extensionEle.getAttribute('Extension');

        if (!defaultExtensions.includes(Extension)) {
          typesElement.appendChild(extensionEle.cloneNode(true));
          defaultExtensions.push(Extension);
        }
      }
    }

    // 处理 fontTable.xml 合并
    const fontTableXml = zip.file(fontTablePath).asText();
    const fontTableDoc = parser.parseFromString(fontTableXml, 'application/xml');

    if (!docFontTableXml) {
      // 第一次遇到 fontTable.xml，保存其内容作为基
      docFontTableXml = fontTableDoc;
    } else {
      const fontsElement = docFontTableXml.getElementsByTagName('w:fonts')[0];
      const existingFonts = Array.from(fontsElement.getElementsByTagName('w:font'))
        .map(font => font.getAttribute('w:name'));

      const newFonts = fontTableDoc.getElementsByTagName('w:font');
      for (let i = 0; i < newFonts.length; i++) {
        const newFont = newFonts[i];
        const fontName = newFont.getAttribute('w:name');

        if (!existingFonts.includes(fontName)) {
          // 如果没有重复的字体，克隆并添加
          fontsElement.appendChild(newFont.cloneNode(true));
          existingFonts.push(fontName); // 更新现有字体列表
        }
      }
    }
  }

  // 生成更新后的 document.xml.rels 字符串
  const serializer = new XMLSerializer();
  const updatedRelsXml = serializer.serializeToString(docRelsXml);
  const updatedTypeXml = serializer.serializeToString(docTypeXml);
  const updatedFontTableXml = serializer.serializeToString(docFontTableXml);

  // 将合并后的文件保存到第一个文件中
  _files[0].file(relsPath, updatedRelsXml);
  _files[0].file(typePath, updatedTypeXml);
  _files[0].file(fontTablePath, updatedFontTableXml);
}

/**
 * 合并多个 Word 文件，将它们的内容合并到一个 ZIP 文件中。
 * 
 * @param {Array} _files - 包含多个 Word 文档的 ZIP 文件数组。
 * @returns {object} - 返回合并后的 ZIP 文件。
 */
function mergeFiles(_files) {
  let baseZip;
  let baseStyleDoms;
  const stylePath = 'word/styles.xml';

  // 遍历每个文件
  for (const zip of _files) {
    if (!baseZip) {
      // 第一个文件作为基，保存其内容
      baseZip = zip;
      const baseXmlText = zip.file(stylePath).asText();
      const parser = new DOMParser();
      const baseDoc = parser.parseFromString(baseXmlText, 'application/xml');
      // 获取 base styles.xml 中的 w:styles 标签
      baseStyleDoms = baseDoc.getElementsByTagName('w:styles')[0];
    } else {
      // 获取当前文件的媒体文件和 header/footer 文件
      const mediaFolder = zip.folder("word");

      // 复制 word/media 中的文件到第一个文件对应目录
      if (mediaFolder) {
        Object.keys(mediaFolder.files).forEach(fileName => {
          // TODO: 'word/endnotes', 'word/footnotes','word/numbering' 这类文件没做对应的合并处理，
          if (['word/media/', 'word/header', 'word/footer', 'word/endnotes', 'word/footnotes', 'word/numbering'].some(item => fileName.startsWith(item))) {
            const file = mediaFolder.files[fileName];
            baseZip.file(fileName, file.asUint8Array());  // 将文件复制到 baseZip 中
          }
          if (fileName === stylePath) {
            const xmlText = mediaFolder.files[fileName].asText();
            const parser = new DOMParser();
            const doc = parser.parseFromString(xmlText, 'application/xml');
            const styleDoms = doc.getElementsByTagName('w:styles')[0];
            // 将 styleDoms 中的 w:style 添加到第一个文件的 styles.xml 中
            if (styleDoms && baseStyleDoms) {
              Array.from(styleDoms.childNodes).forEach(childNode => {
                if (childNode.tagName === 'w:style') {
                  baseStyleDoms.appendChild(childNode);
                }
              });
            }
          }
        });
      }
    }
  }

  const updatedXml = new XMLSerializer().serializeToString(baseStyleDoms);
  baseZip.file(stylePath, updatedXml);

  return baseZip; // 返回更新后的第一个文件
}

module.exports = { combineDocuments, combineRelationships, combineAppInfo, mergeFiles, updateCore };
