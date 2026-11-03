---
name: pdf-translation
description: 当需要修改、调试、排查本项目的文档翻译程序（PDF/Word/Excel/PPT，含 pdf2docx/OCR/写回）时使用
---

# mail_translation 开发手册（工作流与验证）

> **项目现状、代码约束与全部经验教训**以根目录 **AGENTS.md** 为唯一权威（每次会话自动加载）。
> 动手前先读 AGENTS.md；本手册只提供"怎么做"的工作流与验证工具箱，避免两处重复漂移。

## 标准工作流

1. **读 AGENTS.md**：确认现状（`index13.py` 为工作文件、API 配置、已知坑），并明确本次改动波及范围：
   - 改 `qwen_translate` / `process_*` → 影响全部格式
   - 改 `create_docx_pdf2docx` / `_split_tab_cells` / `_clean_fake_borders` / `_replace_para_text` → 只影响 pdf2docx Word 版
   - 改 `apply_translations_to_pdf` / `_plan_group_layout` → 只影响 PDF 版
2. **编码**：在 `index13.py` 上改，中文注释；运行/测试一律 `.venv\Scripts\python.exe`
3. **回归验证（mock 翻译、不调 API）**——最小固定基线：
   - `证据清单.pdf` → 表格实竖线 = **134**
   - `社外秘(関係会社開示可)_20260829圧作労働災害報告.pdf` → 实竖线 = **84**；"場所名|静電" 为拆开的 2 格，"静電"格底纹 = **FFCCFF**
   - `【静電群馬】…pdf` → body 段[2] 含 `w:br`；表1 R4 "8月3日"/"14:30" 为独立两格
   - 更新检查：`_check_update` 用本地临时 `version.json`（monkeypatch `UPDATE_VERSION_FILE`）测"新版本→提示 / 同版本·无文件·坏数据→None"；GUI 提示位置需模拟 `_check_update_async` 截图确认（**不访问共享目录**）
4. **渲染级验证**：`docx2pdf.py`（Word COM 转 PDF）→ PyMuPDF 渲染 → 与原版/基线目视比对
5. **清理**：删除 tmp_* 与调试脚本；对比图按需保留给用户

## 验证工具箱

- **布局回归法**（排版类修复的黄金标准）：docx2pdf 转 PDF 后，对比两版的"表格横线 y 列表"+"`search_for` 相同文本 bbox"——应逐条一致
- **边框计数**：遍历 `w:tcBorders` 的 `w:start`/`w:end`，`val` 非 nil/none 计入
- **docx 合法性**：zipfile 读 `word/document.xml` + lxml 严格解析 + python-docx 打开（**不能代替真实 Word 验证**）
- **import index13 测试**：必须先 `index13.log_box = _MockLog()`（insert/see 两个空方法即可）；
  纯清洗验证：`index13.process_pdf_translations = lambda texts: {}`

## 样例文件速查（各自测什么）

| 文件 | 测试点 |
|---|---|
| `证据清单.pdf` | 纯表格竖线清洗（真线保留 134） |
| `【静電群馬】…pdf` | 混合排版/照片标注、段落软换行(br)、相邻格；与 2.0 基线对比 |
| `社外秘…圧作….pdf` | tab 误合并格（場所名\|静電）、拆格+采样底色、vMerge 同步拆 |
| `经改部规则集.pdf` | 扫描版 OCR（rotation=270）；"源==目标"时模型复读英文语料的现象样本 |
| `version.json`（项目内） | 更新检查模板（发布时复制到共享目录并改版本号） |

## 环境坑速查

- **PowerShell 不能编辑文本文件**（乱码毁文件）——改代码只用编辑工具或 UTF-8 脚本
- 中文输出脚本加 `$env:PYTHONIOENCODING='utf-8'`；复杂逻辑写 .py 文件（`python -c` 引号/反斜杠坑多）
- 中文路径在 `python -c` 里用 `\uXXXX` 转义传参最稳
