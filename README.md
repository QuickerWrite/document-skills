# Documents Skills

独立的 Word、PowerPoint、Excel 技能与 Python 工具包。无需 QuickerWrite 服务端、数据库、模型密钥或网络服务。提供可编辑 Office 文件生成、结构检查、局部修改、CSV 交换及可选的 PDF/PNG 转换。

## 安装与快速开始

需要 Python 3.12+。从源码安装（不假定 PyPI 已发布）：

```sh
git clone https://github.com/QuickerWrite/documents-skills.git
cd documents-skills
python -m venv .venv
. .venv/bin/activate
pip install .
document-skills doctor
document-skills create examples/budget.json budget.xlsx
document-skills inspect budget.xlsx --sheet Budget --range A1:B4
document-skills create examples/report.json report.docx
document-skills create examples/deck.json deck.pptx
```

仅生成/编辑 Office 文件不需要 LibreOffice。PDF 转换与 Excel 重算需要自行安装 LibreOffice（Linux/macOS；Windows 可通过 WSL）及所需字体。可用 `LIBREOFFICE_BIN` 指定程序路径。

```sh
document-skills recalculate budget.xlsx calculated.xlsx
document-skills render calculated.xlsx budget.pdf
document-skills preview budget.pdf page-1.png --page 1
```

默认不覆盖现有文件；明确需要时使用 `--overwrite`。读取失败或转换失败返回非零退出码，不交付部分输出。旧入口 `qw-documents-worker --format docx < input.docx > output.pdf` 保持兼容。

## 能力与边界

| 格式 | 支持 | 明确边界 |
|---|---|---|
| Excel | 多工作表、显式公式、样式、列宽、冻结行列、筛选、柱状/折线图、范围检查、单元格编辑、CSV | openpyxl 不计算公式；重算需 LibreOffice；复杂 Excel 特性可能发生变化 |
| Word | 原生标题、段落、列表、表格、图片、分页等生成，正文/表格检查，正文 run 编辑 | 不提供修订、批注及复杂域编辑；检查不是完整正文 XML 展开 |
| PowerPoint | 原生可编辑幻灯片，模板主题，元素检查和局部 XML 修改 | 局部修改不支持组内对象、继承坐标或空文本对象 |
| PDF/PNG | Office→PDF、指定页 PNG 预览 | 依赖字体与 LibreOffice，无法保证与 Microsoft Office 像素一致 |

`validate` 检查安全边界与解析可读性，不是完整 OOXML schema、公式正确性或视觉质量认证。请保留源文件，重算后检查公式与缓存，并实际检查预览。

## 安装技能

将 `skills/word`、`skills/ppt`、`skills/excel` 复制到支持 SKILL.md 的客户端技能目录，例如 `~/.codex/skills/`。每个目录应完整保留 `references`。同时在该客户端可用的 Python 环境安装本工具包。wheel 也在 `<环境前缀>/share/document-skills/skills/` 内附带这三套技能。

技能使用本地真实可执行命令，不依赖私有平台工具名称。应用平台的 file_id、鉴权、队列和用户确认交互属于适配层，见 [集成指南](docs/integration.md)。

## 文档

- [API 和命令](docs/api.md)
- [安全边界](SECURITY.md)
- [研究来源与许可证决策](docs/research.md)
- [开发与验证](CONTRIBUTING.md)
- 格式说明：[Excel](skills/excel/references/workflows.md)、[Word](skills/word/references/workflows.md)、[PPT](skills/ppt/references/workflows.md)

本仓库使用 MIT 许可证。没有复制限制再分发的 Office skill 源码。测试使用合成数据，不附带客户文件或密钥。
