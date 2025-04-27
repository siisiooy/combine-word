const path = require('path');
const { DOMParser, XMLSerializer } = require('xmldom');

/**
 * 处理 Word 文档中的关系（rels）XML。
 * 此函数会重命名一些元素，如页眉、页脚和媒体文件，更新关系 ID，并处理文件的删除和添加。
 * 
 * @param {object} zip - 包含 Word 文档内容的 ZIP 文件对象。该对象通常来自于一个 ZIP 库，如 JSZip。
 * @param {number} [relIdStart=-1] - 起始关系 ID。如果未提供，则默认为 -1。用于控制生成的关系 ID。
 * @param {object} [startIndex={ headerIndex: 1, footerIndex: 1, mediaIndex: 1 }] - 页眉、页脚和媒体文件的起始索引，默认为 { headerIndex: 1, footerIndex: 1, mediaIndex: 1 }。
 * @returns {object} - 返回一个包含更新后的关系 ID 和文件索引映射的对象。
 *    - nextRelId: 下一个可用的关系 ID。
 *    - fileIndexMap: 包含页眉、页脚和媒体文件的索引映射。
 */
function processRels(zip, relIdStart = -1, startIndex = { headerIndex: 1, footerIndex: 1, mediaIndex: 1 }) {
  const relsPath = 'word/_rels/document.xml.rels';
  const xml = zip.file(relsPath).asText();
  if (!xml) {
    return { mapping: {}, nextRelId: relIdStart };
  }

  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');
  const relationships = doc.getElementsByTagName('Relationship');

  // 用来存储需要重命名的文件任务
  const renameTasks = [];
  const mapping = {};
  const fileIndexMap = startIndex;
  const baseRelFile = relIdStart === -1;
  let cur = relIdStart === -1 ? 1 : relIdStart;

  // 遍历所有关系，处理每一个
  for (let i = 0; i < relationships.length; i++) {
    const rel = relationships[i];
    let oldName = '';
    let newName = '';
    const target = rel.getAttribute('Target');

    if (target.includes('header')) {
      oldName = path.basename(target); // 例如 "header1.xml"
      const ext = path.extname(oldName); // 例如 ".xml"
      newName = `header${fileIndexMap.headerIndex++}${ext}`; // 例如 "header2.xml"
    }
    if (target.includes('footer')) {
      oldName = path.basename(target);
      const ext = path.extname(oldName);
      newName = `footer${fileIndexMap.footerIndex++}${ext}`;
    }
    if (target.includes('media/')) {
      oldName = target;
      newName = oldName.replace(/(\d+)(?=\.[^.]+$)/, `${fileIndexMap.mediaIndex++}`);
    }

    // 如果是基本的关系文件，更新目标路径
    if (baseRelFile) {
      const oldId = rel.getAttribute('Id');
      cur = oldId && Number(oldId.replace('rId', '')) > cur ? Number(oldId.replace('rId', '')) : cur;
      if (["header", "footer", "media/"].some(item => target.includes(item))) {
        rel.setAttribute('Target', newName);
      }
    } else {
      if (["header", "footer", "media/"].some(item => target.includes(item))) {
        const oldId = rel.getAttribute('Id');
        const newId = 'rId' + (cur++);
        mapping[oldId] = newId;
        rel.setAttribute('Id', newId);
        rel.setAttribute('Target', newName);
      }
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
    }
  }

  // 执行文件重命名任务
  renameTasks.forEach(task => {
    zip.remove(task.oldName); // 删除旧文件
  });
  renameTasks.forEach(task => {
    zip.file(task.newName, task.folder); // 添加重命名后的文件
  });

  if (baseRelFile) cur++;
  const builder = new XMLSerializer();
  const updated = builder.serializeToString(doc);
  zip.file(relsPath, updated);

  // 处理相关文件引用
  processDocument(zip, mapping);
  processTypeRels(zip, renameTasks);
  return { nextRelId: cur, fileIndexMap };
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
 * @param {object} zip - 包含 Word 文档内容的 ZIP 文件对象。
 * @param {number} [styleStart=-1] - 起始样式 ID。如果未提供，则默认为 -1。用于控制样式 ID 的生成。
 * @returns {object} - 返回一个包含下一个样式 ID 的对象。
 *    - nextStyleId: 下一个可用的样式 ID。
 */
function processStyles(zip, styleStart = -1) {
  const relsPath = 'word/styles.xml';
  const xml = zip.file(relsPath).asText();
  if (!xml) {
    return { mapping: {}, nextStyleId: styleStart };
  }

  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');
  const styles = doc.getElementsByTagName('w:style');

  const mapping = {};
  let cur = styleStart;

  // 遍历样式，处理每一个样式 ID
  for (let i = 0; i < styles.length; i++) {
    const style = styles[i];
    const oldId = style.getAttribute('w:styleId');
    const newId = cur === -1 ? 'a' : 'a' + cur;
    mapping[oldId] = newId;
    style.setAttribute('w:styleId', newId);
    cur++;
  }

  // 处理每个样式的继承关系、下一样式、链接和编号
  for (let i = 0; i < styles.length; i++) {
    const style = styles[i];

    // 处理 'basedOn' 关系（继承关系）
    const basedOn = style.getElementsByTagName('w:basedOn')[0];
    if (basedOn) {
      const oldBasedOnId = basedOn.getAttribute('w:val');
      if (oldBasedOnId) {
        basedOn.setAttribute('w:val', mapping[oldBasedOnId] || oldBasedOnId); // 更新继承关系的 ID
      }
    }

    // 处理 'next' 样式（下一样式）
    const next = style.getElementsByTagName('w:next')[0];
    if (next) {
      const oldNextId = next.getAttribute('w:val');
      if (oldNextId) {
        next.setAttribute('w:val', mapping[oldNextId] || oldNextId); // 更新下一样式的 ID
      }
    }

    // 处理 'link' 样式
    const link = style.getElementsByTagName('w:link')[0];
    if (link) {
      const oldLinkId = link.getAttribute('w:val');
      if (oldLinkId) {
        link.setAttribute('w:val', mapping[oldLinkId] || oldLinkId); // 更新链接样式的 ID
      }
    }

    // 处理 'numId'（列表编号）引用
    const numId = style.getElementsByTagName('w:numId')[0];
    if (numId) {
      const oldNumId = numId.getAttribute('w:val');
      if (oldNumId) {
        numId.setAttribute('w:val', mapping[oldNumId] || oldNumId); // 更新列表编号的 ID
      }
    }
  }
  console.log("========= mapping =========\n", mapping);
  const builder = new XMLSerializer();
  const updated = builder.serializeToString(doc);
  zip.file(relsPath, updated);

  processDocument(zip, mapping);
  processFileRels(zip, mapping);

  return { nextStyleId: cur };
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

  const timestamp = 'combine_id_' + new Date().getTime();
  // 替换文档中的旧 ID 为新 ID
  for (const [oldId, newId] of Object.entries(mapping)) {
    const re = new RegExp(`"${oldId}"`, 'g');
    xml = xml.replace(re, `"${timestamp}${newId}"`);
  }
  xml = xml.replace(new RegExp(`${timestamp}`, 'g'), '');
  zip.file(relsPath, xml);
}


function processFileRels(zip, mapping) {
  const mediaFolder = zip.folder("word");
  // 复制 word/media 中的文件到第一个文件对应目录
  if (mediaFolder) {
    Object.keys(mediaFolder.files).forEach(fileName => {
      if (['word/header', 'word/footer'].some(item => fileName.startsWith(item))) {
        const file = mediaFolder.files[fileName];
        let xml = file.asText();
        const timestamp = 'combine_id_' + new Date().getTime();

        for (const [oldId, newId] of Object.entries(mapping)) {
          const re = new RegExp(`"${oldId}"`, 'g');
          xml = xml.replace(re, `"${timestamp}${newId}"`);
        }

        xml = xml.replace(new RegExp(`${timestamp}`, 'g'), '');
        mediaFolder.file(fileName, xml);
      }
    });
  }
}

module.exports = { processRels, processStyles };
