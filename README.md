# Markdown到Word公文格式转换工具

将Markdown文件转换为符合GB/T 9704-2012《党政机关公文格式》标准的Word文档。

## 功能特点

- 符合GB/T 9704-2012国家标准
- 支持LaTeX数学公式、表格和列表
- 智能处理Obsidian格式图片
- 自动设置 A4 纸张（210mm × 297mm）和公文格式（页边距、字体、字号、页码）

## 安装

### 系统要求
- Python 3.6+
- Pandoc（必须）：[安装说明](https://pandoc.org/installing.html)

### 安装依赖
```bash
pip3 install -r requirements.txt
```

## 项目结构

```
md-to-word/
├── src/                         # 源代码目录
│   ├── __init__.py
│   ├── core/                    # 核心处理模块
│   │   ├── __init__.py
│   │   ├── markdown_preprocessor.py
│   │   ├── pandoc_processor.py
│   │   └── word_postprocessor.py
│   ├── formatters/              # 格式化器模块
│   │   ├── __init__.py
│   │   ├── base_formatter.py
│   │   ├── document_title_formatter.py
│   │   ├── image_formatter.py
│   │   ├── list_formatter.py
│   │   ├── page_formatter.py
│   │   ├── paragraph_formatter.py
│   │   └── table_formatter.py
│   ├── utils/                   # 工具模块
│   │   ├── __init__.py
│   │   ├── config_validator.py
│   │   ├── constants.py
│   │   ├── exceptions.py
│   │   ├── path_validator.py
│   │   └── xpath_cache.py
│   └── config/                  # 配置模块
│       ├── __init__.py
│       └── config.py
├── docs/                        # 文档目录
│   ├── architecture.md          # 架构文档
│   ├── configuration.md         # 配置指南
│   └── security.md              # 安全策略
├── examples/                    # 示例文件
│   ├── example.md
│   └── example.docx
├── filters/                     # Pandoc过滤器
│   └── chinese_filter.lua
├── md_to_word.py               # 主程序入口
├── requirements.txt            # Python依赖
├── README.md                   # 项目说明
├── CLAUDE.md                   # 项目状态记录
└── LICENSE                     # MIT许可证
```

## 配置说明

### 环境变量配置

本工具支持通过环境变量配置 Obsidian 相关路径，方便在不同环境下使用。支持的环境变量包括：

- `OBSIDIAN_VAULT_NAME` - Obsidian Vault 名称
- `OBSIDIAN_ATTACHMENTS_FOLDER` - 附件文件夹名称  
- `OBSIDIAN_VAULT_PATH` - Vault 完整路径

详细配置说明请查看 [配置指南](docs/configuration.md)。

### 代码配置

可在 `src/config/config.py` 中调整以下设置：
- 字体和字号配置
- A4 纸张和页边距设置
- 图片处理选项
- Pandoc转换参数

### 安全性

- 命令注入防护：使用 `subprocess.run([arg1, arg2, ...])` 的列表参数方式，避免 shell 解析
- 路径遍历防护：验证所有文件路径，防止目录遍历攻击
- XML注入防护：使用安全的XML API构建元素

## 使用方法

```bash
# 基本用法
python3 md_to_word.py document.md

# 指定输出文件
python3 md_to_word.py input.md -o output.docx

# 非交互覆盖已存在输出
python3 md_to_word.py input.md -o output.docx --force

# 使用自定义 Obsidian Vault
OBSIDIAN_VAULT_NAME="我的笔记" python3 md_to_word.py document.md
```

## 输入验证和支持说明

### 支持的输入格式

#### 文件类型要求
- **支持的文件扩展名**：`.md`、`.markdown`
- **文件编码**：UTF-8（推荐），工具会自动处理UTF-8编码的文件
- **文件位置**：支持任意路径的Markdown文件（绝对路径或相对路径）

#### 文件大小限制
- **无硬性文件大小限制**：工具设计用于处理各种大小的文档
- **性能考虑**：大型文档（>10MB）可能需要更长的处理时间
- **内存使用**：处理超大文档时可能需要更多系统内存

### 路径要求

#### 输入文件路径
- **支持绝对路径**：`/Users/username/Documents/document.md`
- **支持相对路径**：`./document.md`、`../docs/document.md`
- **自动创建输出目录**：如果输出路径的目录不存在，工具会自动创建

#### 图片路径支持
- **相对路径**：相对于Markdown文件的路径
- **绝对路径**：完整的文件系统路径
- **Obsidian格式**：`![[filename]]`（无需路径，自动搜索）
- **自动搜索路径**（按优先级）：
  1. Markdown文件所在目录
  2. Obsidian附件目录（可配置）
  3. `./images`目录
  4. `./assets`目录
  5. 当前工作目录

### 支持的Markdown功能

#### 完全支持的功能
- **标题**：`##`（一级）、`###`（二级），其他级别作为正文处理
- **段落**：标准段落文本，自动应用首行缩进
- **列表**：
  - 无序列表：`-`、`*`（自动转换为`-`）
  - 有序列表：自动转换为正文格式，保持编号
  - 多级列表：支持缩进的嵌套列表
- **表格**：标准Markdown表格语法，自动格式化
- **图片**：
  - 标准格式：`![alt text](image.png)`
  - Obsidian格式：`![[image.png]]`
  - 支持格式：PNG、JPG、JPEG、GIF、BMP、SVG、WEBP
- **数学公式**：
  - 行内公式：`$E=mc^2$`
  - 块级公式：`$$E=mc^2$$`
  - 使用MathML渲染，支持完整LaTeX语法
- **加粗文本**：`**text**`或`__text__`（自动去除加粗标记）
- **代码块**：保留但不应用特殊格式
- **引用块**：`>`开头的引用文本

#### 自动处理的元素
- **YAML Front Matter**：自动过滤，不出现在输出中
- **结尾元数据**：Date、标签等元数据自动过滤
- **文件名标题**：使用文件名作为文档标题
- **一级标题（#）**：自动跳过，不影响文档结构
- **图片标题**：智能处理，移除文件名，保留有意义的描述

#### 限制和注意事项
- **不支持的Markdown扩展**：
  - 脚注
  - 任务列表（`- [ ]`）
  - 定义列表
  - HTML内嵌代码
- **标题层级限制**：只识别二级（##）和三级（###）标题，更深层级作为正文
- **表格限制**：不支持跨行或跨列的复杂表格
- **代码高亮**：代码块保留但不应用语法高亮

### 图片格式要求

#### 支持的图片格式
- **位图格式**：PNG、JPG、JPEG、GIF、BMP、WEBP
- **矢量格式**：SVG（转换时可能栅格化）
- **文件名**：支持中文、英文、数字、下划线、连字符
- **无扩展名文件**：自动尝试匹配支持的格式

#### 图片处理特性
- **自动全宽显示**：图片自动调整为页面全宽（156mm）
- **保持纵横比**：自动计算高度，确保图片不变形
- **居中对齐**：所有图片自动水平居中
- **文字环绕**：默认Top and Bottom模式（可配置）
- **智能标题**：自动清理图片文件名，保留有意义的描述

## 格式说明

### 标题层级
- `##` → 一级标题（黑体）
- `###` → 二级标题（楷体）
- 其他 → 正文（仿宋）

### 支持的Markdown元素
- 数学公式：`$E=mc^2$`、`$$E=mc^2$$`
- 表格：标准Markdown表格语法
- 列表：有序列表自动转为正文格式
- 图片：支持`![]()`和`![[]]`格式

## 版本信息

**v2.4.0**（2025-07-24）
- 修复有序列表和多级编号处理问题
- 确保字体格式统一

## 许可

MIT License
