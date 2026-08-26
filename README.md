# Markdown 到 Word 公文格式转换工具

将正文型 Markdown 转换为 DOCX，并按 GB/T 9704-2012 的核心版式要求统一页面、字体、字号、文档网格、段落、标题和页码。项目面向个人稳定使用，不试图覆盖版记、密级、发文字号等全部可选公文要素。仓库保留一份 [GB/T 9704-2012 参考文件](<docs/GB T 9704 2012.pdf>)，用于离线核对版式条文。

## 主要能力

- A4 纵向页面，28 字 × 22 行文档网格，奇偶页页码分别外侧对齐。
- 中文按内容层级选择小标宋、仿宋、黑体或楷体；所有拉丁字母和数字统一使用 Times New Roman。
- 支持正文、两级标题、有序与无序列表、表格、LaTeX 数学公式、标准 Markdown 图片和 Obsidian 图片。
- 自动处理附件说明、正文分页符、图片说明，以及单一机关盖章公文的署名和成文日期。
- 通过临时文件生成、DOCX 重新打开验证和原子替换，避免失败时覆盖已有输出。

## 安装

需要 Python 3.11 或更高版本，以及可从 `PATH` 调用的 [Pandoc](https://pandoc.org/installing.html)。推荐使用 [uv](https://docs.astral.sh/uv/) 管理环境，仓库已提交 `uv.lock` 锁定全部依赖版本：

```bash
uv sync
```

也可以使用标准工具链创建虚拟环境并安装运行依赖：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

先检查环境，再转换文件：

```bash
python md_to_word.py --check-config
python md_to_word.py input.md -o output.docx
```

输出文件存在时，CLI 默认询问是否覆盖；自动化环境可增加 `--force`。使用 `-v` 可输出调试日志。仓库中的自包含示例可用 `python3 md_to_word.py examples/公文格式示例.md -o /tmp/md-to-word-example.docx --force` 直接转换。

## Markdown 编写约定

### 文档标题与层级标题

如果 Markdown 中恰好有一个一级标题（`# Title`），转换器会将其文本作为文档内部标题，并从正文中移除该标题；没有一级标题时才使用 Markdown 文件名。文档内部标题使用二号方正小标宋并居中排版，不依赖 YAML frontmatter。存在多个一级标题时，继续使用文件名作为文档标题，并将已编写的标题层级下移。常规输入建议使用：

- `##` 表示一级标题，使用三号黑体。
- `###` 表示二级标题，使用三号楷体。
- 更深层级按三号仿宋正文处理。
- 正文和层级标题的段前、段后均为 0，并对齐正文网格；文档标题使用二号小标宋和独立的精确行距。

YAML frontmatter 是可选的。存在时会被过滤，不会出现在 Word 中，也不负责驱动落款识别。

### 列表

有序列表会转换为带可见编号的普通正文段落，从而避免 Word 或 WPS 自动编号产生不稳定的缩进与字体。无序列表的 `-` 和 `*` 会转换为 Word 原生项目符号列表：标记统一为 `•`，换行使用悬挂对齐，嵌套层级每级右移两个正文字符。

### 附件说明

附件按自然 Markdown 编写即可：`附件：` 独占一行，空一行后依次写 `1.`、`2.`、`3.`。转换器会将各项生成独立的附件说明段落，自动对齐第二条及后续序号；附件名称较长时，续行与上一行名称首字对齐。不要手工添加全角空格、制表符或行尾空格。

例如：

> 附件：
>
> 1. Example implementation plan
> 2. Project detail table
> 3. Acceptance checklist

转换器还会确保附件说明前恰好有一个对齐文档网格的空行。

### 落款与成文日期

单一机关盖章公文可在 Markdown 最后写署名和日期，不需要 frontmatter：

> The final body paragraph.
>
> Example Municipal Government Office
> 2026-07-14

识别条件是：署名是日期之前最近的非空行，署名前有一个用于分隔正文的空行；署名与日期之间可以有空行，转换时会统一规范。日期可以是最后一个正文行，也可以紧邻一个位于文档末尾的附件说明之前。末尾的 `---`、单词标签（如 `#work`）和 `Date:` 笔记元数据会先被过滤，不影响识别。日期接受 ISO 形式或不补零的中文形式。

转换器保留落款与附件说明在 Markdown 中的先后顺序，在落款前生成两个正文网格空行，将日期规范为中文形式并右空四字，再让署名以日期为准居中编排。这里只处理文字位置，不插入印章图片。

### 图片与说明

支持标准 Markdown 图片和 Obsidian 图片。相对路径优先基于 Markdown 源文件目录解析，再查找配置的 Obsidian 路径、`images`、`assets` 和当前目录。图片按版心全宽缩放并保持纵横比；图片锚点段落会压缩，因此图片和 caption 之间不产生额外空行。找不到本地图片时转换失败，不发布不完整文档。

### 分页与其他元素

正文中的独立 `---`、`***` 或 `___` 会转换为分页符。文件开头成对的 `---` 仍按 YAML frontmatter 边界处理。表格和公式由 Pandoc 转换后再统一格式化；复杂跨行、跨列表格、脚注、任务列表、定义列表和内嵌 HTML 不属于稳定支持范围。代码块和引用可由 Pandoc 保留，但没有专门的公文样式。

## 配置

集中配置位于 `src/config/config.py`，包括页面、页边距、字体、字号、文档网格、列表、表格、图片和 Pandoc 参数。默认关键值如下：

- 页面：A4 纵向；上 34.58 mm、下 32.58 mm、左 28 mm、右 26 mm。
- 网格：每行 28 字、每页 22 行；正文行距 10.39 mm，文档标题行距 12.51 mm。
- 字号：文档标题二号，正文与层级标题三号，表格及图片说明四号，页码四号。
- 段落：正文首行缩进 2 字；段前、段后均为 0。

Obsidian 路径可通过以下环境变量覆盖：

- `OBSIDIAN_VAULT_PATH`：Vault 的绝对路径。
- `OBSIDIAN_VAULT_NAME`：供自动发现使用的 Vault 名称。
- `OBSIDIAN_ATTACHMENTS_FOLDER`：Vault 内的附件目录名。

完整说明见[配置指南](docs/configuration.md)。Word 和 WPS 会替换本机未安装的字体；要获得可重复的输出，应在转换机器上安装配置字体，或将字体名改为该环境中可用的等价字体。

## 输入、输出与失败行为

- 输入扩展名支持 `.md`、`.markdown`、`.mdown`、`.mkd` 和 `.mdwn`，编码应为 UTF-8，文件上限为 100 MB。
- 输入和输出可使用绝对或相对路径；输出目录不存在时会自动创建。
- 输出固定为 `.docx`；指定其他后缀时会自动改为 `.docx`。
- 生成与后处理在目标目录的临时文件中完成。只有文件能被 `python-docx` 重新打开时才原子替换目标文件。
- Pandoc、图片、格式化或验证失败时，已有输出保持不变，临时文件会被清理。

安全边界和威胁模型见[安全说明](docs/security.md)，模块职责见[架构说明](docs/architecture.md)。

## 开发与 CI

安装固定版本的开发工具。使用 uv 时，`uv sync` 默认同时安装 `dev` 依赖组（ruff、mypy、pip-audit）；使用标准工具链时单独安装：

```bash
# uv（推荐）：
uv sync
# 标准工具链：
python -m pip install -r requirements.txt -r requirements-dev.txt
```

提交前运行与 CI 相同的检查：

```bash
ruff check .
ruff format --check .
mypy
pip-audit -r requirements.txt
python md_to_word.py --check-config
python -m unittest discover -v
```

GitHub Actions 会在 push、pull request 和手工触发时运行质量检查，并在 Python 3.11 与 3.14 上执行测试。具体任务以 [CI 工作流](.github/workflows/ci.yml) 为准；依赖 Pandoc 的本地测试在 Pandoc 不可用时自动跳过。

## 版本与许可

当前版本为 **v2.7.0**，采用 [MIT License](LICENSE)。
