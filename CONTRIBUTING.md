# 开发与验证

```sh
python -m venv .venv
. .venv/bin/activate
pip install -e '.[dev]'
pytest
ruff check .
python -m build
```

安装 LibreOffice 后 `pytest` 自动运行真实 PDF 转换及公式重算测试；缺少时标记 skip，不得报告为通过。不要提交生成产物、私人文档、.env 或访问凭据。测试使用合成数据，新增修复应验证可观察结果与不变项，而非只匹配实现文本。

修改 API 时同步相应 skills/references 与 docs/api.md；修改技能后执行 skill-creator 的 quick_validate.py（若本地有该工具）。发布 wheel 必须验证附带技能资源。当前没有自动发布到 PyPI 的工作流。
