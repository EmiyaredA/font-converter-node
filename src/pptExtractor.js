/**
 * PPT字体提取工具
 * 用于从PPTX文件中提取嵌入的字体
 */

import fs from 'fs-extra';
import path from 'path';
import AdmZip from 'adm-zip';
import { convertEOTtoTTF } from './eotConverter.js';
import { XMLParser } from 'fast-xml-parser';

/**
 * 从PPTX文件中提取所有嵌入的字体
 * @param {string} pptxPath PPTX文件路径
 * @param {string} outputDir 输出目录
 * @returns {Promise<string[]>} 提取的字体文件路径列表
 */
export async function extractFontsFromPPTX(pptxPath, outputDir) {
  const extractedFonts = [];
  
  try {
    console.log(`开始从PPTX文件提取字体: ${pptxPath}`);
    
    // 确保输出目录存在
    await fs.ensureDir(outputDir);
    
    // 使用AdmZip打开PPTX文件 (PPTX实际上是ZIP文件)
    const zip = new AdmZip(pptxPath);
    const zipEntries = zip.getEntries();
    
    // 同时提取字体样式信息
    const embeddedFontMap = await extractTypefaceInfo(zip, outputDir);
    // 遍历所有条目，查找字体文件
    for (const entry of zipEntries) {
      // 查找ppt/fonts/目录下的所有字体文件
      if (entry.entryName.startsWith('ppt/fonts/') && !entry.isDirectory) {
        const fontFileName = path.basename(entry.entryName);
        const fileName = embeddedFontMap[fontFileName]
        const outputPath = path.join(outputDir, fileName.replace(' ', '___') + '.fntData');
        const outputPath2 = path.join(outputDir, fontFileName);
        
        // 提取字体文件
        zip.extractEntryTo(entry.entryName, outputDir, false, true);

        fs.copyFileSync(outputPath2, outputPath);
        
        console.log(`提取字体文件: ${fileName}`);
        extractedFonts.push(outputPath);
      }
    }

    return extractedFonts;
  } catch (error) {
    console.error(`提取字体时发生错误: ${error.message}`);
    return extractedFonts;
  }
}

/**
 * 从PPTX文件中提取字体样式信息
 * @param {AdmZip} zip 已打开的PPTX文件(ZIP)
 * @param {string} outputDir 输出目录
 * @returns {Promise<Object>} 提取的字体样式信息
 */
export async function extractTypefaceInfo(zip, outputDir) {
  try {
    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@_'
    });

    const embeddedFontMap = {};
    let embeddedFontLst = [];
    
    // 尝试查找并解析presentation.xml文件，它包含主题和字体信息
    const presentationEntry = zip.getEntry('ppt/presentation.xml');
    if (presentationEntry) {
      const presentationXml = presentationEntry.getData().toString('utf8');
      const presentationData = parser.parse(presentationXml);
      embeddedFontLst = presentationData?.['p:presentation']?.['p:embeddedFontLst']?.['p:embeddedFont'] || [];
    }

    const presentationEntryRels = zip.getEntry('ppt/_rels/presentation.xml.rels');
    if (presentationEntryRels) {
      const presentationXml = presentationEntryRels.getData().toString('utf8');
      const presentationData = parser.parse(presentationXml);

      const themeRels = presentationData.Relationships.Relationship.filter(rel => rel['@_Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font');
      
      themeRels.forEach(rel => {
        const rid = rel['@_Id']
        const font = embeddedFontLst.find(font => font['p:regular']['@_r:id'] === rid);
        const path = (rel['@_Target'] || '').split('/').pop();
        embeddedFontMap[path] = font['p:font']['@_typeface'];
      });
    }

    return embeddedFontMap;
  } catch (error) {
    console.error(`提取字体样式信息时发生错误: ${error.message}`);
    return {};
  }
}

/**
 * 从主题文件中提取字体样式信息
 * @param {AdmZip} zip 已打开的PPTX文件(ZIP)
 * @param {string} masterId 幻灯片母版ID
 * @param {string} outputDir 输出目录
 * @param {Object} typefaceInfo 字体样式信息对象（将被修改）
 */
async function extractThemeTypefaces(zip, masterId, outputDir, typefaceInfo) {
  try {
    // 查找并解析关系文件
    const relsEntry = zip.getEntry('ppt/_rels/presentation.xml.rels');
    if (!relsEntry) return;
    
    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@_'
    });
    
    const relsXml = relsEntry.getData().toString('utf8');
    const relsData = parser.parse(relsXml);
    
    // 保存关系文件以便检查
    await fs.writeFile(path.join(outputDir, 'presentation.xml.rels'), relsXml);
    
    // 查找主题文件路径
    if (relsData?.Relationships?.Relationship) {
      const relationships = Array.isArray(relsData.Relationships.Relationship)
        ? relsData.Relationships.Relationship
        : [relsData.Relationships.Relationship];
      
      const masterRel = relationships.find(rel => rel['@_Id'] === masterId);
      if (masterRel && masterRel['@_Target']) {
        const masterPath = 'ppt/' + masterRel['@_Target'];
        const masterEntry = zip.getEntry(masterPath);
        
        if (masterEntry) {
          const masterXml = masterEntry.getData().toString('utf8');
          const masterData = parser.parse(masterXml);
          
          // 保存母版文件以便检查
          await fs.writeFile(path.join(outputDir, path.basename(masterPath)), masterXml);
          
          // 查找主题引用
          if (masterData?.p?.sldMaster?.['@_theme']) {
            const themeId = masterData.p.sldMaster['@_theme'];
            
            // 查找主题关系文件
            const masterRelsPath = masterPath.replace('.xml', '.xml.rels');
            const masterRelsEntry = zip.getEntry('ppt/_rels/' + path.basename(masterRelsPath));
            
            if (masterRelsEntry) {
              const masterRelsXml = masterRelsEntry.getData().toString('utf8');
              const masterRelsData = parser.parse(masterRelsXml);
              
              // 保存主题关系文件以便检查
              await fs.writeFile(path.join(outputDir, path.basename(masterRelsPath)), masterRelsXml);
              
              // 查找主题文件路径
              if (masterRelsData?.Relationships?.Relationship) {
                const masterRelationships = Array.isArray(masterRelsData.Relationships.Relationship)
                  ? masterRelsData.Relationships.Relationship
                  : [masterRelsData.Relationships.Relationship];
                
                const themeRel = masterRelationships.find(rel => rel['@_Id'] === themeId);
                if (themeRel && themeRel['@_Target']) {
                  const themePath = path.dirname(masterPath) + '/' + themeRel['@_Target'];
                  const themeEntry = zip.getEntry(themePath);
                  
                  if (themeEntry) {
                    const themeXml = themeEntry.getData().toString('utf8');
                    const themeData = parser.parse(themeXml);
                    
                    // 保存主题文件以便检查
                    await fs.writeFile(path.join(outputDir, path.basename(themePath)), themeXml);
                    
                    // 提取主题中的字体信息
                    if (themeData?.a?.theme?.a?.fontScheme) {
                      const fontScheme = themeData.a.theme.a.fontScheme;
                      
                      // 提取标题字体
                      if (fontScheme.majorFont?.latin) {
                        typefaceInfo.titleFont = fontScheme.majorFont.latin['@_typeface'];
                      }
                      
                      // 提取正文字体
                      if (fontScheme.minorFont?.latin) {
                        typefaceInfo.bodyFont = fontScheme.minorFont.latin['@_typeface'];
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  } catch (error) {
    console.error(`提取主题字体样式信息时发生错误: ${error.message}`);
  }
}

/**
 * 处理提取的字体文件，将EOT字体转换为TTF
 * @param {string[]} fontFilePaths 字体文件路径列表
 * @param {string} outputDir TTF输出目录
 * @returns {Promise<string[]>} 转换后的TTF字体文件路径列表
 */
export async function convertExtractedFonts(fontFilePaths, outputDir) {
  const convertedFonts = [];
  
  try {
    // 确保输出目录存在
    await fs.ensureDir(outputDir);
    
    for (const fontPath of fontFilePaths) {
      const fontName = path.basename(fontPath);
      const baseName = path.basename(fontName, path.extname(fontName));
      const ttfPath = path.join(outputDir, `${baseName}.ttf`);
      
      if (fontName.toLowerCase().endsWith('.fntdata')) {
        console.log(`转换字体文件: ${fontName}`);
        const success = await convertEOTtoTTF(fontPath, ttfPath);
        if (success) {
          convertedFonts.push(ttfPath);
        }
      }
      // 可以扩展支持其他字体格式的转换
    }
    
    console.log(`完成字体转换，共转换${convertedFonts.length}个字体文件`);
    return convertedFonts;
  } catch (error) {
    console.error(`转换字体时发生错误: ${error.message}`);
    return convertedFonts;
  }
} 