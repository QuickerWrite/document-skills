# 上游研究与来源

研究日期：2026-11-05。通过 GitHub 搜索 Office / agent skills 并阅读以下项目。以下为工程思路比较，本仓库没有复制第三方 skill 正文或实现。

| 来源 | 查阅内容 | 采用的工程决策 |
|---|---|---|
| [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI/tree/main/skills) | officecli-xlsx、officecli-docx；仓库 Apache-2.0 | 先检查再修改、显式元素定位、结构与视觉分开验证、公式缓存与原式分开检查、宽表按页宽适配 |
| [OpenAI skills PDF](https://github.com/openai/skills/tree/main/skills/.curated/pdf) | PDF 技能的布局检查流程 | 增加独立 PDF/PNG 预览入口；文本提取不能替代页面检查 |
| [Anthropic XLSX](https://github.com/anthropics/skills/tree/main/skills/xlsx) | SKILL.md 与专用 LICENSE.txt | 发现该目录限制复制/再分发，明确排除代码和正文复用；公式重算能力自行实现 |

不采用上游未经验证的特定客户端行为或强制性设计规则。例如不会要求所有文档达到固定封面填充比例，也不会把验证通过当成数据正确。CLI 帮助、实际运行和目标 Office 客户端表现才是能力依据。

第一方原生 Word/PPT 生成器来自 QuickerWrite 服务端，已移除应用导入依赖；本库不包含鉴权、支付、模型配置、连接凭据或客户数据。依赖项保留其各自许可证；MIT 只适用于本仓库第一方代码与文档。
