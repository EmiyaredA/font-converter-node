/**
 * EOT字体转换工具
 * 将EOT (Embedded OpenType)格式字体文件转换为TTF格式
 */

import fs from 'fs-extra';
import path from 'path';
import ttf2woff2 from 'ttf2woff2';

// EOT文件头的大小通常是固定的，但实际值可能因文件而异
const EOT_HEADER_SIZE = 512;

/**
 * 将EOT格式字体转换为TTF格式
 * @param {string} eotFilePath EOT文件路径
 * @param {string} ttfOutputPath 输出的TTF文件路径
 * @returns {Promise<boolean>} 是否转换成功
 */
export async function convertEOTtoTTF(eotFilePath, ttfOutputPath) {
  try {
    console.log(`开始转换EOT字体文件: ${eotFilePath}`);
    
    // 读取EOT文件
    const eotData = await fs.readFile(eotFilePath);
    
    // 提取TTF数据
    const ttfData = extractTTFFromEOT(eotData);
    console.log(`提取TTF数据成功，长度为: ${ttfData.length}字节`, ttfData);
    if (!ttfData || ttfData.length === 0) {
      console.error('无法从EOT文件中提取TTF数据');
      return false;
    }

    // 转换为WOFF2格式
    // const woff2Data = ttf2woff2(ttfData);
    // if (!woff2Data || woff2Data.length === 0) {
    //   console.error('无法将TTF数据转换为WOFF2格式');
    //   return false;
    // }
    
    // 确保输出目录存在
    await fs.ensureDir(path.dirname(ttfOutputPath));
    console.log(`成功将EOT字体转换为WOFF2格式，保存到: ${ttfOutputPath}`);
    
    // 保存TTF文件
    await fs.writeFile(ttfOutputPath, ttfData);
    
    console.log(`成功将EOT字体转换为TTF格式，保存到: ${ttfOutputPath}`);
    return true;
  } catch (error) {
    console.error(`转换字体文件时发生错误: ${error.message}`);
    return false;
  }
}

/**
 * 使用FontFace API加载转换后的WOFF2字体文件
 * @param {string} fontName 字体名称
 * @param {string} woff2FilePath WOFF2文件路径
 * @param {Object} options 加载选项
 * @param {string} options.fontStyle 字体样式 (默认: 'normal')
 * @param {string} options.fontWeight 字体粗细 (默认: 'normal')
 * @returns {Promise<FontFace|null>} 加载的FontFace对象
 */
export async function loadWoff2Font(fontName, woff2FilePath, options = {}) {
  try {
    // 检查是否在浏览器环境中
    if (!isBrowser()) {
      throw new Error('此方法仅适用于浏览器环境');
    }
    
    // 默认选项
    const { 
      fontStyle = 'normal', 
      fontWeight = 'normal' 
    } = options;
    
    // 读取WOFF2文件
    const fontData = await fs.readFile(woff2FilePath);
    
    // 创建Blob对象
    const fontBlob = new Blob([fontData], { type: 'font/woff2' });
    
    // 创建Object URL
    const fontUrl = URL.createObjectURL(fontBlob);
    
    // 创建FontFace对象
    const fontFace = new FontFace(fontName, `url(${fontUrl})`, {
      style: fontStyle,
      weight: fontWeight
    });
    
    // 加载字体
    await fontFace.load();
    
    // 将字体添加到字体列表
    document.fonts.add(fontFace);
    
    console.log(`成功加载字体: ${fontName}`);
    return fontFace;
  } catch (error) {
    console.error(`加载字体时发生错误: ${error.message}`);
    return null;
  }
}

/**
 * 直接从WOFF2字节数据加载字体（适用于浏览器环境）
 * @param {string} fontName 字体名称
 * @param {Buffer|Uint8Array} woff2Data WOFF2字体的字节数据
 * @param {Object} options 加载选项
 * @param {string} options.fontStyle 字体样式 (默认: 'normal')
 * @param {string} options.fontWeight 字体粗细 (默认: 'normal')
 * @param {boolean} options.autoRelease 是否在加载后自动释放资源 (默认: false)
 * @returns {Promise<{fontFace: FontFace, cleanup: Function}|null>} 加载的FontFace对象和清理函数
 */
export async function loadWoff2FontFromData(fontName, woff2Data, options = {}) {
  try {
    // 检查是否在浏览器环境中
    if (!isBrowser()) {
      throw new Error('此方法仅适用于浏览器环境');
    }
    
    // 默认选项
    const { 
      fontStyle = 'normal', 
      fontWeight = 'normal',
      autoRelease = false
    } = options;
    
    // 创建Blob对象
    const fontBlob = new Blob([woff2Data], { type: 'font/woff2' });
    
    // 创建Object URL
    const fontUrl = URL.createObjectURL(fontBlob);
    
    // 创建FontFace对象
    const fontFace = new FontFace(fontName, `url(${fontUrl})`, {
      style: fontStyle,
      weight: fontWeight
    });
    
    // 加载字体
    await fontFace.load();
    
    // 将字体添加到字体列表
    document.fonts.add(fontFace);
    
    // 创建清理函数
    const cleanup = () => {
      try {
        document.fonts.delete(fontFace);
        URL.revokeObjectURL(fontUrl);
        console.log(`已释放字体资源: ${fontName}`);
      } catch (error) {
        console.error(`释放字体资源时发生错误: ${error.message}`);
      }
    };
    
    // 如果设置了自动释放，则在字体加载完成后释放资源
    if (autoRelease) {
      // 延迟释放，确保字体已被应用
      setTimeout(cleanup, 1000);
    }
    
    console.log(`成功加载字体: ${fontName}`);
    return { fontFace, cleanup };
  } catch (error) {
    console.error(`加载字体时发生错误: ${error.message}`);
    return null;
  }
}

/**
 * 检测当前是否在浏览器环境中运行
 * @returns {boolean} 是否在浏览器环境中
 */
function isBrowser() {
  return typeof window !== 'undefined' && typeof document !== 'undefined';
}

/**
 * 从EOT文件数据中提取TTF数据
 * EOT格式是TTF/OTF格式的包装器，包含了一个头部和元数据，后面跟着实际的TTF/OTF字体数据
 * @param {Buffer} eotData EOT文件的字节数据
 * @returns {Buffer|null} TTF字体数据
 */
function extractTTFFromEOT(eotData) {
  try {
    // EOT文件格式相对复杂，以下是一个简化的实现
    
    // 尝试寻找TTF签名 (0x00010000 或 'OTTO' 或 'true' 或 'typ1')
    let ttfStartOffset = -1;
    for (let i = 0; i < eotData.length - 4; i++) {
      if ((eotData[i] === 0x00 && eotData[i + 1] === 0x01 && eotData[i + 2] === 0x00 && eotData[i + 3] === 0x00) ||
          (eotData[i] === 0x4F && eotData[i + 1] === 0x54 && eotData[i + 2] === 0x54 && eotData[i + 3] === 0x4F) || // 'OTTO'
          (eotData[i] === 0x74 && eotData[i + 1] === 0x72 && eotData[i + 2] === 0x75 && eotData[i + 3] === 0x65) || // 'true'
          (eotData[i] === 0x74 && eotData[i + 1] === 0x79 && eotData[i + 2] === 0x70 && eotData[i + 3] === 0x31)) { // 'typ1'
        ttfStartOffset = i;
        break;
      }
    }
    
    // 如果找不到TTF签名，则尝试使用固定偏移
    if (ttfStartOffset === -1) {
      console.log('无法找到TTF签名，尝试使用固定偏移量');
      ttfStartOffset = EOT_HEADER_SIZE;
    }
    
    // 提取TTF数据
    const ttfData = eotData.slice(ttfStartOffset);
    
    console.log(`从EOT文件中提取了${ttfData.length}字节的TTF数据`);
    return ttfData;
  } catch (error) {
    console.error(`提取TTF数据时发生错误: ${error.message}`);
    return null;
  }
}

/**
 * 检查是否是有效的TTF文件头
 * @param {Buffer} ttfData TTF文件数据
 * @returns {boolean} 是否有有效的TTF头
 */
export function hasValidTTFHeader(ttfData) {
  if (!ttfData || ttfData.length < 4) {
    return false;
  }
  
  // 检查TTF魔数
  return (ttfData[0] === 0x00 && ttfData[1] === 0x01 && ttfData[2] === 0x00 && ttfData[3] === 0x00) || // TrueType
         (ttfData[0] === 0x4F && ttfData[1] === 0x54 && ttfData[2] === 0x54 && ttfData[3] === 0x4F) || // OpenType/CFF ('OTTO')
         (ttfData[0] === 0x74 && ttfData[1] === 0x72 && ttfData[2] === 0x75 && ttfData[3] === 0x65) || // TrueType (Apple) ('true')
         (ttfData[0] === 0x74 && ttfData[1] === 0x79 && ttfData[2] === 0x70 && ttfData[3] === 0x31);   // Type 1 ('typ1')
} 