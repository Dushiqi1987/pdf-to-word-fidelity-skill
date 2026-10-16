# PDF 转可编辑 Word Skill

把 PDF 内容重建为可编辑 Word，并按你提供的 DOCX 样例或文字要求排版。重点保留正文、原生公式、真正的表格，以及目录、编号和页眉之间的一致性。

## 怎么用

把本仓库链接、PDF 和格式样例交给 Codex，例如：

> 请读取并使用 https://github.com/Dushiqi1987/pdf-to-word-fidelity-skill 的 SKILL.md，将完整仓库下载到工作区（不只读取入口文件），按其中流程把附带 PDF 转成 Word。格式以附带 DOCX 为准。公式必须原生可编辑，表格保持三线表。先检查环境，完成后逐页核对，结果保存到工作区。

只给链接不会在所有客户端自动安装 skill。Codex 需要能访问仓库并读取文件；也可以下载整个目录到 Codex 的 skills 目录。无需先安装 documents、pdf 或其他 skill，但需要 Codex 可运行本地 Python，并允许准备缺失的运行依赖。

没有 DOCX 样例时，可以直接写格式要求，例如纸张、页边距、正文字体字号、行距、标题层级、页眉和目录。未指定时，优先参考 PDF 并说明采用的规则。

## 可以期待什么

- 正文保持文字，不把整页截图塞进 Word。
- 公式转换为 Word 原生 OMML，支持检查求和上下限、分式、矩阵和公式编号。
- 表格为可编辑单元格，三线表不等于普通网格表。
- 可按要求设置目录、阿拉伯数字章节编号和奇偶页页眉。
- 后续修改以你最新编辑的 Word 为基准，只改指定内容。
- 结构审计与渲染检查分开记录，不用一个“转换成功”掩盖未验证的问题。

## 重要限制

这不是对所有 PDF 保证“一键、完全无损”的转换器。PDF 经常没有可恢复的语义结构，公式尤其需要 AI 对照原图重建。扫描件还取决于 OCR、清晰度和符号辨认。样例 Word 与 PDF 的字体、行距、版心不同，也会导致页数变化；版式一致和原分页一致可能无法同时满足。

Python 依赖可由 bootstrap 脚本检查或准备。渲染需要本地 Microsoft Word 或 LibreOffice 等受支持工具；Windows Word 原生渲染通常更适合检查目标 Word 效果。没有渲染器时仍可做结构检查，但不能声称已逐页视觉验证。字体、OCR、Office 软件和许可证需由使用环境合法提供。

## 内容结构

- `SKILL.md`：Codex 的执行入口。
- `scripts/`：环境检查、输入分析、文档构建、原生公式、渲染与审计工具。
- `references/workflow.md`：完整操作流程与验收规则。
- `references/lessons-and-provenance.md`：本次方法积累、常见失误与来源说明。
- `evals/evals.json`：可复用的场景验收提示。
- `examples/demo-plan.json`：不含原稿内容的公开合成示例。
- `tests/`：可运行的结构回归测试。

## 运行验证

```text
python scripts/bootstrap.py --install
python -m unittest discover -s tests -v
python scripts/build_docx.py --plan examples/demo-plan.json --output outputs/draft.docx
python scripts/render_verify.py outputs/draft.docx --outdir qa --finalized-docx outputs/final.docx
python scripts/audit_docx.py outputs/final.docx --expected examples/demo-plan.json
```

使用 Codex 依赖工具返回的 Python 路径，不要混用解释器。`--finalized-docx` 会创建新文件并更新目录页码；只有 Word 渲染路径支持此完成步骤。原稿不变。默认渲染为只读，适合后续局部修改。仅有 LibreOffice 时可渲染，但目录域缓存仍需本地编辑器处理并复核。

## 隐私与来源

发布包不包含用户的 PDF、Word、私人路径、账号信息、转换中间件或原文片段。默认在本地处理文档，不上传到第三方转换网站。调用在线 OCR 或其他外部服务前，应先向用户说明并取得相应授权。

本项目复用的是工作方法，不捆绑其他厂商 skill 的全文、Microsoft Office 转换样式表、字体或商业软件。第三方依赖遵循各自许可证。详见 [来源与经验](references/lessons-and-provenance.md)。
