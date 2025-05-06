const { DOMParser, XMLSerializer } = require('xmldom');

/**
 * 处理 Word 文档中的关系（rels）XML。
 * 此函数会重命名一些元素，如页眉、页脚和媒体文件，更新关系 ID，并处理文件的删除和添加。
 * 
 * @param {object} _files - 包含 Word 文档内容的 ZIP 文件对象数组。
 * @returns {void} 函数没有返回值，直接修改 ZIP 文件中的内容。
 */
function processRels(_files) {
  const relsPath = 'word/_rels/document.xml.rels';

  const fileIndexMap = {
    headerIndex: 1,
    footerIndex: 1,
    mediaIndex: 1,
  };
  let relId = 1;
  const targetList = [];

  // 遍历每个文件
  for (let fileIndex = 0; fileIndex < _files.length; fileIndex++) {
    const zip = _files[fileIndex];
    const parser = new DOMParser();
    const xml = zip.file(relsPath).asText();
    const doc = parser.parseFromString(xml, 'application/xml');
    const relationships = doc.getElementsByTagName('Relationship');
    const renameTasks = [];
    const mapping = {};

    // 遍历所有关系，处理每一个
    for (let i = 0; i < relationships.length; i++) {
      const rel = relationships[i];
      let oldName = '';
      let newName = '';
      const target = rel.getAttribute('Target');
      if (!['header', 'footer', 'media/'].some(item => target.startsWith(item)) && targetList.includes(target)) {
        continue; // 如果目标已经存在，跳过处理
      }
      if (!['header', 'footer', 'media/'].some(item => target.startsWith(item))) {
        targetList.push(target);
      }

      if (target.includes('header')) {
        oldName = target.split('/').pop(); // 获取文件名，比如 "header1.xml"
        const ext = oldName.slice(oldName.lastIndexOf('.')); // 获取文件扩展名，比如 ".xml"
        newName = `header${fileIndexMap.headerIndex++}${ext}`; // 例如 "header2.xml"
      }
      if (target.includes('footer')) {
        oldName = target.split('/').pop();
        const ext = oldName.slice(oldName.lastIndexOf('.'));
        newName = `footer${fileIndexMap.footerIndex++}${ext}`;
      }
      if (target.includes('media/')) {
        oldName = target;
        newName = oldName.replace(/(\d+)(?=\.[^.]+$)/, `${fileIndexMap.mediaIndex++}`);
      }
      // 将需要重命名的文件任务存储到数组中
      if (oldName) {
        const folder = zip.file(`word/${oldName}`).asUint8Array();
        if (folder) {
          renameTasks.push({
            oldName: `word/${oldName}`,
            newName: `word/${newName}`,
            folder
          });
        }
        rel.setAttribute('Target', newName);
      }

      // 如果是基本的关系文件，更新目标路径
      if (fileIndex === 0) {
        const oldId = rel.getAttribute('Id');
        relId = oldId && (Number(oldId.replace('rId', '')) > relId) ? Number(oldId.replace('rId', '')) : relId;
      } else {
        const oldId = rel.getAttribute('Id');
        const newId = 'rId' + (relId++);
        rel.setAttribute('Id', newId);
        mapping[oldId] = newId;
      }
    }
    const builder = new XMLSerializer();
    const updated = builder.serializeToString(doc);
    zip.file(relsPath, updated);

    if (fileIndex === 0) relId++;

    // 执行文件重命名任务
    renameTasks.forEach(task => {
      zip.remove(task.oldName); // 删除旧文件
    });
    renameTasks.forEach(task => {
      zip.file(task.newName, task.folder); // 添加重命名后的文件
    });

    // 处理相关文件引用
    processDocument(zip, mapping);
    processTypeRels(zip, renameTasks);
  }
}

/**
 * 处理 Word 文档中的 [Content_Types].xml 文件，更新与文件类型相关的引用。
 * 
 * @param {object} zip - 包含 Word 文档内容的 ZIP 文件对象。
 * @param {Array} replaceMapping - 一个包含需要替换的文件名和路径的映射数组。
 *    每个元素是一个对象，包含 oldName 和 newName 属性。
 */
function processTypeRels(zip, replaceMapping = []) {
  const relsPath = '[Content_Types].xml';
  const xml = zip.file(relsPath).asText();
  if (!xml || replaceMapping.length === 0) {
    return;
  }

  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');
  const overrides = doc.getElementsByTagName('Override');

  for (let i = 0; i < overrides.length; i++) {
    const rel = overrides[i];
    const target = rel.getAttribute('PartName');

    for (const pathObj of replaceMapping) {
      const oldPath = `/${pathObj.oldName}`;
      const newPath = `/${pathObj.newName}`;
      if (target === oldPath) {
        rel.setAttribute('PartName', newPath);
      }
    }
  }

  const builder = new XMLSerializer();
  const updated = builder.serializeToString(doc);
  zip.file(relsPath, updated);
}

/**
 * 处理 Word 文档中的 styles.xml 文件，重命名样式 ID，并处理样式的继承、下一样式、链接和编号。
 *
 * @param {object} _files - 包含 Word 文档内容的 ZIP 文件对象数组。
 * @returns {void} 函数没有返回值，直接修改 ZIP 文件中的内容。
 */
function processStyles(_files) {
  const relsPath = 'word/styles.xml';
  const parser = new DOMParser();
  const builder = new XMLSerializer();
  let styleStart = -1;
  let styleNumStart = 0;

  // 遍历每个文件
  for (let i = 0; i < _files.length; i++) {
    const zip = _files[i];
    const xml = zip.file(relsPath).asText();
    const doc = parser.parseFromString(xml, 'application/xml');
    const styles = doc.getElementsByTagName('w:style');
    const stylesArray = Array.from(styles);

    // 根据 w:styleId 从小到大排序
    stylesArray.sort((a, b) => {
      const idA = a.getAttribute('w:styleId');
      const idB = b.getAttribute('w:styleId');

      // 如果是纯数字 ID，可以将它们作为数字进行比较
      if (/^\d+$/.test(idA) && /^\d+$/.test(idB)) {
        return Number(idA) - Number(idB);
      }

      // 否则按字母顺序比较
      return idA.localeCompare(idB);
    });

    const mapping = {};
    const updateIds = ['w:basedOn', 'w:next', 'w:link', 'w:numId'];
    let cur = styleStart;
    let numCur = styleNumStart;

    // 遍历样式，处理每一个样式 ID
    for (let i = 0; i < stylesArray.length; i++) {
      const style = stylesArray[i];
      let oldId = style.getAttribute('w:styleId');
      let newId;
      if (!mapping[oldId]) {
        // 如果 ID 为纯数字，直接使用数字 ID
        if (/^\d+$/.test(oldId)) {
          newId = Number(cur) + styleNumStart; // 使用数字 ID
          numCur = Number(cur) > numCur ? Number(cur) : numCur;
        } else {
          newId = cur === -1 ? 'a' : 'a' + cur;
          cur++;
        }
        mapping[oldId] = newId;
      }
      newId = mapping[oldId];
      style.setAttribute('w:styleId', newId); // 更新样式 ID
      for (const tagName of updateIds) {
        processStyleAttribute(style, tagName, 'w:val', mapping); // 更新样式属性
      }
    }

    const updated = builder.serializeToString(doc);
    zip.file(relsPath, updated);
    processDocument(zip, mapping);
    processFileRels(zip, mapping);
  }
}

/**
 * 更新指定元素中指定属性的 ID 值，根据传入的映射表将旧的 ID 替换为新的 ID。
 * 
 * @param {Element} style - 当前样式元素，用于获取目标标签和属性。
 * @param {string} tagName - 目标标签的名称，通常是样式元素的子标签。
 * @param {string} attributeName - 需要更新的属性名称（如 'w:val' 或其他属性）。
 * @param {Object} mapping - 一个映射对象，键是旧的 ID，值是新的 ID，用于替换。
 * 
 * @returns {void} 函数没有返回值，直接修改元素的属性。
 */
function processStyleAttribute(style, tagName, attributeName, mapping) {
  const element = style.getElementsByTagName(tagName)[0];
  if (element) {
    const oldId = element.getAttribute(attributeName);
    if (oldId) {
      element.setAttribute(attributeName, mapping[oldId] || oldId); // 更新 ID
    }
  }
}

/**
 * 更新 XML 节点中的特定属性值。
 *
 * @param {NodeList} nodes - 需要更新的 XML 节点集合。
 * @param {string} attr - 需要修改的属性名。
 * @param {string} newValue - 新的属性值。
 */
function updateNodeAttributes(docXml) {
  const nodesKey = ['w:p', 'w:tr'];
  const attributeObject = {
    'w14:textId': '77777777',
    'w:rsidR': '00000000',
    'w:rsidRDefault': '00000000'
  };
  for (const key of nodesKey) {
    const nodes = docXml.getElementsByTagName(key)
    Array.from(nodes).forEach(node => {
      for (const [attr, newValue] of Object.entries(attributeObject)) {
        if (node.hasAttribute(attr)) {
          node.setAttribute(attr, newValue);
        }
      }
    });
  }
}

/**
 * 替换 XML 内容中的旧 ID 为新 ID。
 *
 * @param {string} xml - 原始 XML 内容。
 * @param {Object} mapping - 包含旧 ID 和新 ID 映射的对象。
 * @returns {string} - 更新后的 XML 内容。
 */
function replaceIdsInXml(xml, mapping) {
  const timestamp = 'combine_id_' + new Date().getTime();
  // 替换文档中的旧 ID 为新 ID
  for (const [oldId, newId] of Object.entries(mapping)) {
    const re = new RegExp(`"${oldId}"`, 'g');
    xml = xml.replace(re, `"${timestamp}${newId}"`);
  }
  // 移除临时时间戳
  return xml.replace(new RegExp(`${timestamp}`, 'g'), '');
}

/**
 * 处理 Word 文档中的 document.xml 文件，并更新内容中的 ID。
 *
 * @param {object} zip - 包含 Word 文档内容的 ZIP 文件对象。
 * @param {object} mapping - 一个包含旧 ID 和新 ID 映射的对象。
 */
function processDocument(zip, mapping) {
  const relsPath = 'word/document.xml';
  let xml = zip.file(relsPath).asText();


  // 替换 ID
  xml = replaceIdsInXml(xml, mapping);

  const docXml = new DOMParser().parseFromString(xml, 'application/xml');

  // 更新 <w:p> 和 <w:tr> 标签的属性
  updateNodeAttributes(docXml);

  // 重新序列化 XML 文档为字符串
  const serializer = new XMLSerializer();
  const updatedXml = serializer.serializeToString(docXml);
  // 将更新后的 XML 内容写回到 ZIP 文件中
  zip.file(relsPath, updatedXml);
}

/**
 * 处理 ZIP 文件中的 Word 文档关系文件（如 header 和 footer），根据提供的 ID 映射替换文件内容中的 ID，并修改特定 XML 属性。
 * 
 * @param {JSZip} zip - 需要处理的 ZIP 文件，包含一个 `word` 文件夹，其中包含待处理的 Word 文档文件。
 * @param {Object} mapping - 一个映射对象，其中键是旧的 ID，值是新的 ID。用于在文件内容中替换对应的 ID。
 * 
 * @returns {void} 函数没有返回值，直接修改 ZIP 文件中的内容。
 * 
 * @description
 * 1. 遍历 ZIP 文件中的 `word/media` 文件夹，查找以 `word/header` 或 `word/footer` 开头的文件。
 * 2. 对找到的文件内容进行文本替换，替换所有旧 ID 为新的 ID，使用映射对象 `mapping`。
 * 3. 修改 XML 中的 `<w:p>` 和 `<w:tr>` 标签的属性：将 `w14:textId` 设置为 '77777777'，`w:rsidR` 和 `w:rsidRDefault` 设置为 '00000000'。
 * 4. 最后，更新修改后的 XML 内容并将其写回 ZIP 文件。
 */
function processFileRels(zip, mapping) {
  const mediaFolder = zip.folder("word");

  if (mediaFolder) {
    Object.keys(mediaFolder.files).forEach(fileName => {
      if (['word/header', 'word/footer'].some(item => fileName.startsWith(item))) {
        const file = mediaFolder.files[fileName];
        let xml = file.asText();

        // 替换 ID
        xml = replaceIdsInXml(xml, mapping);

        const docXml = new DOMParser().parseFromString(xml, 'application/xml');

        // 更新 <w:p> 和 <w:tr> 标签的属性
        updateNodeAttributes(docXml);

        // 重新序列化 XML 文档为字符串
        const serializer = new XMLSerializer();
        const updatedXml = serializer.serializeToString(docXml);
        // 将更新后的 XML 内容写回到 ZIP 文件中
        zip.file(fileName, updatedXml);
      }
    });
  }
}

module.exports = { processRels, processStyles };
