# PPT字体转换工具 (Node.js版)

这个工具用于解析PowerPoint (PPTX) 文件中的嵌入字体文件（特别是.fntdata格式），并将其转换为标准TTF字体格式。

## 功能特点

- 从PPTX文件中提取嵌入的字体文件
- 将EOT (Embedded OpenType) 格式的.fntdata文件转换为TTF格式
- 支持单个字体文件的直接转换
- 支持批量处理文件夹中的多个PPTX或字体文件
- 提供命令行界面和Node.js API

## 环境要求

- Node.js 14.0或更高版本
- npm或yarn包管理器

## 安装

```bash
# 本地安装
npm install

# 全局安装（可选）
npm install -g .
```

## 使用方法

### 命令行方式

转换单个字体文件：

```bash
# 本地运行
node index.js -f ../dsadsadsa/ppt/fonts/font16.fntdata -o output

# 全局安装后运行
ppt-font-converter -f path/to/font.fntdata -o output
```

处理整个PPTX文件：

```bash
# 本地运行
node index.js -p path/to/presentation.pptx -o output

# 全局安装后运行
ppt-font-converter -p path/to/presentation.pptx -o output
```

处理包含多个PPTX或字体文件的文件夹：

```bash
# 本地运行
node index.js -d path/to/folder -o output

# 全局安装后运行
ppt-font-converter -d path/to/folder -o output
```

### 命令行参数

```
选项:
  -V, --version            显示版本号
  -p, --pptx <file>        PPTX文件路径
  -f, --font <file>        单个字体文件路径
  -d, --dir <directory>    包含PPTX或字体文件的文件夹路径
  -o, --output <dir>       输出目录 (默认: "output")
  -h, --help               显示帮助信息
```

### 在Node.js项目中使用

```javascript
import { convertEOTtoTTF } from './src/eotConverter.js';
import { extractFontsFromPPTX, convertExtractedFonts } from './src/pptExtractor.js';

// 转换单个字体文件
const success = await convertEOTtoTTF('path/to/font.fntdata', 'path/to/output.ttf');

// 从PPTX中提取并转换所有字体
const extractedFonts = await extractFontsFromPPTX('presentation.pptx', 'extracted_fonts');
const convertedFonts = await convertExtractedFonts(extractedFonts, 'converted_fonts');
```

## 技术说明

PPTX文件实际上是一个ZIP存档，其中包含了各种XML文件和资源。嵌入的字体文件通常存储在`ppt/fonts/`目录下，格式为`.fntdata`（实际上是EOT格式）。

转换过程：
1. 解析EOT文件头
2. 提取出嵌入的TTF数据
3. 保存为新的TTF文件

### EOT转换细节

EOT (Embedded OpenType) 是Microsoft开发的字体格式，专门为网页设计。它基于TrueType/OpenType字体，但添加了额外的头部和数据结构。转换过程中，我们需要：

1. 跳过EOT头部（通常约512字节）
2. 查找TTF签名（如0x00010000或"OTTO"）
3. 提取TTF数据并保存

## 项目结构

```
font-converter-node/
├── index.js          # 主程序入口
├── package.json      # 项目配置
├── README.md         # 使用说明
└── src/
    ├── eotConverter.js  # EOT转TTF转换器
    └── pptExtractor.js  # PPT字体提取器
```

## 已知限制

- 当前实现使用简化的EOT解析，可能无法处理所有EOT文件变体
- 某些高度压缩或加密的字体可能无法正确转换

## 许可证

MIT许可证 