# API / CLI

Python API 输入 Office 内容为 `bytes`，输出生成/修改结果为 `bytes`，不自动写文件。调用方负责输出存储、授权和资源隔离。

| 模块 | 函数 | 结果 |
|---|---|---|
| `spreadsheet` | `create_xlsx(spec)` | XLSX bytes；schema 见 Excel workflows |
| `spreadsheet` | `inspect_xlsx(data, sheet=None, cell_range='A1:T50')` | sheets/cells，value/type/cached/error；范围外不检查 |
| `spreadsheet` | `edit_xlsx(data, changes)` | XLSX bytes；cell/value 或 style |
| `spreadsheet` | `csv_to_xlsx(text, sheet='Data')` | 全部按文本导入 |
| `spreadsheet` | `xlsx_to_csv(data, sheet, cell_range='A1:T50')` | CSV 字符串；使用缓存并防文本公式注入 |
| `render_word` | `render_word(title, content, styles, options=None)` | DOCX bytes |
| `word` | `inspect_docx(data)` / `edit_docx(data, changes)` | 正文段落和表格 / 保留 run 格式的修改 |
| `render_ppt` | `render_ppt(title, content, styles, ppt_options=None, *, template_bytes=None)` | PPTX bytes |
| `pptx_patch` | `inspect_pptx(data, slide_number=None)` / `patch_pptx(data, patches)` | 元素列表 / 局部修改 |
| `render` | `recalculate_xlsx(data, cancelled=lambda:False)` | LibreOffice 重算副本 |

历史内部入口 `_libreoffice_pdf`, `_guard_ooxml`, `_pdf_document`, `_page_png` 仍保留，供旧 worker 适配器使用；新调用方优先使用 CLI。它们以下划线标记，未来 API 稳定化时可能新增公共名称。

CLI：`document-skills --help`；`create INPUT.json OUTPUT.{docx,pptx,xlsx}`、`inspect INPUT`、`edit INPUT OUTPUT --changes PATCH.json`、`validate INPUT`、`render INPUT OUTPUT.pdf`、`preview INPUT OUTPUT.png --page N --width 1000`、`recalculate INPUT.xlsx OUTPUT.xlsx`、`csv-import INPUT.csv OUTPUT.xlsx`、`csv-export INPUT.xlsx OUTPUT.csv --sheet NAME --range A1:D20`。

读取上限 80 MiB，Office 解压总量 250 MiB、单部件 50 MiB、最多 12,000 部件；PDF 最多 200 页；转换最长 180 秒。PNG 页码从 1 开始，宽度 64–4096，实际渲染还受高度 1800 限制。输出默认拒绝覆盖，`--overwrite` 才替换；同目录临时文件原子发布。命令错误输出到 stderr，退出码 1；参数错误退出码 2。

CLI JSON 用 UTF-8，路径参数需按 shell 规则引用。不要把动态公式拼入 shell：写到 JSON 文件。返回的缓存值不保证新鲜；编辑保存后缓存可能为空。CSV 仅导出指定范围，不会自动覆盖整张表。Excel 样式/图表回写依赖 openpyxl，不提供原 ZIP 字节保真保证。
