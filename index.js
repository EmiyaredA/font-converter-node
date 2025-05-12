#!/usr/bin/env node

/**
 * PPT字体转换工具
 * 用于提取PPT中的字体文件并转换为标准TTF格式
 */

import fs from 'fs';
import path from 'path';
import { program } from 'commander';
import { convertEOTtoTTF } from './src/eotConverter.js';
import { extractFontsFromPPTX, convertExtractedFonts } from './src/pptExtractor.js';
import FormData from 'form-data';
import fetch from 'node-fetch';
import { exec } from 'child_process';

// 版本和描述信息
program
  .name('ppt-font-converter')
  .description('将PPT中的EOT(fntdata)字体文件转换为TTF格式')
  .version('1.0.0');

// 命令行参数定义
program
  .option('-p, --pptx <file>', 'PPTX文件路径')
  .option('-f, --font <file>', '单个字体文件路径')
  .option('-o, --output <dir>', '输出目录', 'output');

// 解析命令行参数
program.parse();
const options = program.opts();

/**
 * 主函数
 */
async function main() {
  let { pptx, font, output } = options;
  console.log(pptx, font, output);
  // 如果没有指定参数，使用默认示例
  if (!pptx && !font) {
    
    const defaultFontPath = './output/extracted/font1.fntdata';
    if (fs.existsSync(defaultFontPath)) {
      font = defaultFontPath;
      console.log(`使用默认样例文件: ${font}`);
    } else {
      console.error(`未指定输入文件且未找到默认字体文件，请指定--pptx或--font参数`);
      program.help();
      return;
    }
  }
  
  try {
    // 处理PPTX文件
    if (pptx) {
      console.log(`开始处理PPTX文件: ${pptx}`);
      
      // 提取字体
      const extractDir = path.join(output, 'extracted');
      const extractedFonts = await extractFontsFromPPTX(pptx, extractDir);
      
      if (extractedFonts.length === 0) {
        console.warn('未从PPTX文件中提取到任何字体');
      } else {
        // 转换提取的字体
        const convertDir = path.join(output, 'converted');
        const convertedFonts = await convertExtractedFonts(extractedFonts, convertDir);
        
        console.log(`成功转换${convertedFonts.length}个字体文件`);
        
        // 输出转换后的字体文件路径
          for (const fontPath of convertedFonts) {
            console.log(`转换后的字体文件: ${fontPath}`);
            const fontName = path.basename(fontPath);
            const baseName = path.basename(fontName, path.extname(fontName)).replace('___', ' ');
            
            try {
              // 获取文件状态以检查文件大小
              const stats = fs.statSync(fontPath);
              console.log(`准备上传文件: ${fontPath}, 大小: ${stats.size} 字节`);
              
              // 使用child_process执行curl命令
              const curlCommand = `curl -vL 'http://sensenote-test.sensetime.com/graphql' \
                -H "Content-Type: multipart/form-data" \
                -F 'operations={"query": "mutation ($font: Upload!, $name: String!) { putFont(name: $name, font: $font) }", "variables": {"font": null, "name": "${baseName}"}}' \
                -F 'map={"file":["variables.font"]}' \
                -F 'file=@${fontPath}'`;
              
              exec(curlCommand, (error, stdout, stderr) => {
                if (error) {
                  console.error(`执行curl命令失败: ${error}`);
                  return;
                }
                console.log(`上传成功，响应: ${stdout}`);
                if (stderr) console.error(`curl stderr: ${stderr}`);
              });
            } catch (error) {
              console.error(`上传文件失败: ${error.message}`);
            }
          }
      }
    }
    
    // 处理单个字体文件
    if (font) {
      console.log(`开始处理单个字体文件: ${font}`);
      
      const fontName = path.basename(font);
      const baseName = path.basename(fontName, path.extname(fontName));
      const ttfPath = path.join(output, `${baseName}.ttf`);
      
      // 转换字体
      const success = await convertEOTtoTTF(font, ttfPath);
      
      if (success) {
        console.log(`成功转换字体文件: ${ttfPath}`);
      } else {
        console.error(`转换字体文件失败`);
      }
    }
  } catch (error) {
    console.error(`处理过程中发生错误: ${error.message}`);
    process.exit(1);
  }
}

// 执行主函数
main().catch(error => {
  console.error(`程序运行出错: ${error.message}`);
  process.exit(1);
}); 