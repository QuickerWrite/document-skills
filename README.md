# QuickerWrite 文档技能

维护 PPT、Word 技能提示词，以及不依赖数据库的 Office 渲染与 PPT 元素修改能力。仓库为私有仓库，不存放应用凭据或用户文档。

## 文档处理流程

- PPT：HTML 布局 → Chromium 提取 → 可编辑 PPTX → LibreOffice 转 PDF → PDFium 页面图 → AI 检查 → 带版本校验的元素精修。
- Word：保持原生 DOCX 生成，使用同一条实际文件渲染与检查链路，不经过 HTML 转换。
- 选中元素通过独立的结构化上下文传递；用户输入区只显示修改要求，文件编号、哈希和内部提示词不进入可编辑正文。

## 安装与运行

Python 3.12 及以上，安装 `pip install .`。运行环境需要 LibreOffice Writer/Impress/Calc、Chromium、Noto CJK 和 Liberation 字体。

```sh
qw-documents-worker --format docx < input.docx > output.pdf
```

工作器从标准输入读取文件，向标准输出写入 PDF，无需数据库权限。PDFium 渲染与 PPT 元素读取、精修函数也可供任务适配层调用。

## 发布与部署

服务端独立的 `worker-office` 队列负责权限、取消、存储和版本校验。发布包固定本仓库的 Git 提交，并使用 SHA-256 清单验证随包源码；qwctl 统一安装依赖、检查实际渲染并管理独立服务或容器。生产机器不需要 GitHub 凭据。更新时修改固定版本并重新构建服务端发布包，再走常规 qwctl 升级流程。

Dashi、Guizang 保持独立的 AGPL 引擎服务，不把其实现复制到本私有仓库。当前拆分覆盖渲染、PPT 元素操作和提示词；HTML 转换与生成适配层仍位于服务端工作器。

## 维护约定

所有 QuickerWrite 仓库的 GitHub 简介统一使用中文，README 的项目说明也使用中文；技术名称、命令和接口标识保留原文。
