/**
 * 测试PPTX字体样式提取功能
 */

import path from 'path';
import { fileURLToPath } from 'url';
import { extractFontsFromPPTX, extractTypefaceInfo } from './src/pptExtractor.js';
import AdmZip from 'adm-zip';
import fs from 'fs-extra';

// 获取当前文件的目录
const __dirname = path.dirname(fileURLToPath(import.meta.url));

// 配置
const testPptxPath = process.argv[2] || path.join(__dirname, 'test-files', 'sample.pptx');
const outputDir = path.join(__dirname, 'output', path.basename(testPptxPath, '.pptx'));

async function main() {
  try {
    console.log('开始测试typeface提取功能...');
    console.log(`输入PPTX文件: ${testPptxPath}`);
    console.log(`输出目录: ${outputDir}`);
    
    // 确保PPTX文件存在
    if (!await fs.pathExists(testPptxPath)) {
      console.error(`错误: PPTX文件不存在: ${testPptxPath}`);
      console.log('用法: node test-typeface.js [pptx文件路径]');
      process.exit(1);
    }
    
    // 确保输出目录存在
    await fs.ensureDir(outputDir);
    
    // 提取字体文件
    const extractedFonts = await extractFontsFromPPTX(testPptxPath, outputDir);
    console.log(`提取了${extractedFonts.length}个字体文件`);
    
    // 如果需要单独测试typeface提取
    console.log('\n独立测试typeface提取功能:');
    const zip = new AdmZip(testPptxPath);
    const typefaceInfo = await extractTypefaceInfo(zip, outputDir);
    
    console.log('\n提取的字体样式信息:');
    console.log(JSON.stringify(typefaceInfo, null, 2));
    
    console.log('\n测试完成!');
  } catch (error) {
    console.error(`测试过程中发生错误: ${error.message}`);
    console.error(error.stack);
  }
}

// 运行主函数
main().catch(err => {
  console.error('程序执行失败:', err);
  process.exit(1);
}); 