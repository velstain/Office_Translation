# AGENTS.md — Office AI 翻译工具

## 项目定位
Tkinter GUI 的办公文档翻译工具，调用本地 LLM API 将 Excel/Word/PPT/PDF 翻译为目标语言。

## 开发入口
- **最新版本**：`index13.py` — 始终在此文件上工作；`index12.py` 是上一版兜底，仍保留
- **虚拟环境**：`.venv`（Python 3.12.6）。**必须用 `.venv\Scripts\python.exe` 运行**——系统 Python 没有 xlwings/pdf2docx
- **没有** `requirements.txt` 或 `pyproject.toml`，依赖需手动安装（如 `pdf2docx` 用户已手动装过）

## 核心依赖及用途
| 库 | 用途 |
|---|---|
| `PyMuPDF` (import fitz) | PDF 解析/写入/渲染 |
| `python-docx` | Word 读写、Word 生成 |
| `python-pptx` | PPT 读写 |
| `xlwings` | Excel 读写（可保留图片/图表，非 openpyxl） |
| `pdf2docx` | PDF→DOCX 转换（仅 Word 翻译版用） |
| `Pillow` / `numpy` | 图片处理（已随项目安装） |
| `tkinter` | GUI（内置） |

## API 配置
- **当前 API**：`http://10.237.192.48:8000/v1/chat/completions`；模型 `qwen36-35b-a3b`
- 支持 `response_format: json_object` 批量翻译；超时 `REQUEST_TIMEOUT = 1200` 秒
- 默认目标语言是"日语"（`lang_box.current(2)`），不是中文；脚本测试时 `TARGET_LANG=""` 会让模型自选语言（如英文），需手动设 `index13.TARGET_LANG='中文'`

## 编程规范
- 全程使用中文注释
- Word/Excel/PPT 翻译逻辑稳定，**不要改动**；GUI（`get_user_config`）不要改动（例外：经用户批准加入的"有新版本可用"提示，仅在翻译按钮下方显示）

## 版本与更新检查（仅提示，不自动下载/替换）
- 版本号：`index13.py` 顶部 `APP_VERSION`（当前 `"5.0"`）
- 新版本信息：共享目录 `UPDATE_SHARE_DIR`（默认 `\\10.237.105.210\install\AI\Office AI翻译工具V5`）下的 `version.json`：`{"version": "5.1", "notes": "..."}`
- 启动 1.5s 后**后台线程**读取并比对（SMB 不可达时 `open` 会阻塞数十秒，**绝不能放主线程**）；有新版本时在翻译按钮正下方显示图标+"有新版本可用"，点击打开共享目录（失败弹窗："无法访问路径…可能是新版本存放路径发生了变化或是网络问题"）
- **发版流程**：①改 `APP_VERSION` ②打包 `pyinstaller index13.spec --noconfirm` ③覆盖共享 V5 的 `翻译工具.exe`+`_internal` ④改共享 `version.json` 的 `version`（**①④必须同步**）
- **权限边界**：用户明确要求 agent **不操作共享文件夹**（打包、替换由用户自己做）；项目内 `version.json` 仅作模板
- 图标 `icon/download.png`（16×16，由 receipt_filler 项目的 `icon/下载.svg` 光栅化而来）随 spec 的 datas 打包

## PDF 翻译（index13 当前实现）的关键发现

### 文本提取与分组
- span 可能字符级拆分（"序号"→"序"+"号"）；以 block 内 line 合并 span，再按**单元格**分组
- 分组优先用 `page.find_tables()` 的单元格边界（行中心点归属单元格）；无表格页 fallback 启发式（间隙<字号×1.5 且上行不以句号结尾则合并换行行）
- 含图片 PDF 的 find_tables 可能把跨多行文字合并成一个大单元格（如静電群馬右侧 10 行合并）→ 涂白/写字时按行处理避图片

### 写回（`apply_translations_to_pdf`）
- **页面 /Rotate（90/270）**：`get_text`/`find_tables`/`get_image_rects` 返回**未旋转**坐标，而涂白/写字（`_vis_*` + `rotate=page.rotation`）按**视觉坐标**。`extract_texts_from_pdf` 必须把所有坐标 `* page.rotation_matrix` 转成视觉坐标（含 span bbox/origin、表格 cell、图片 rect）——否则旋转页文本组碎片化、写回文字被旋转 90°（261009 案例：整页文字纵向排列“切开”文件）
- **涂白必须用 `Shape.commit(overlay=True)`**：`page.draw_rect(overlay=True)` 在本 PyMuPDF 版本对"扫描底图+文字层"的 PDF 实际落在背景层，盖不住图片上的原文（静态白块测试验证）
- **"扫描底图 + 原生文字层"页面**（如静電群馬：电子文字层 + 整页扫描底）：文字层与底图存在偏移 → 逐行白块盖不严会露出原文叠加 → 对 `page_fullimage` 页且组与照片重叠 <10% 时改用**组矩形大块涂白**；与照片重叠大的组仍逐行（避免盖照片）
- **文档固有限制**：OCR 文字层不全（扫描件部分标题无文字层）、竖排文字（縦書き）不支持翻译 → 这类文本保留原文，属文档固有而非 bug
- **`insert_textbox` 在此 PyMuPDF 版本对 CJK 排版有 bug**（9 字宽空需 493pt 高度、溢出时不写入任何 span）——**不要使用**，用 `insert_text` + `_wrap_text` 手动拆行
- `insert_textbox` 返回值：**正数=文字放下（未使用空间），负数=溢出（不写入）**，别搞反
- 字体统一 `china-s`（CJK+拉丁）；粗体 `flags & 16` → 偏移 0.4pt 双重渲染；颜色从 span `color` 整数转浮点 RGB
- 列聚类**必须 15pt 阈值**（40pt 会把序号列/证据名称列合并导致 box 错乱写越界）；文字右边界用 `min(列右界-2, 页宽)` 限制
- 书名「数字范围保护」：`2-5 页` 翻译前剥离前缀 `2-5` 只译"页"再拼回（LLM 会丢"-"）
- 字号：**单行优先**（缩字号保单行，≥7pt 可读；不可行才拆行）；总行高 ≤ `下一同列组顶部 - 组顶 - 3pt`（max_y 硬边界）
- 涂白**按行**（不整组涂白）、图片在行间不被盖；写字遇图片跳到图片下方

### Word 翻译版（生成两份）
- `create_docx_pdf2docx()` → `名字_翻译版.docx`：pdf2docx 转 docx 后替换文本（按单元格宽度缩字号，不改行高）
- `create_docx_from_pdf()` → `名字_翻译版_背景版.docx`：每页 PyMuPDF 渲染 PNG 为整页浮动背景图（behindDoc=1）+ 翻译文本绝对定位浮动文本框（behindDoc=0）

### pdf2docx 版替换/清洗的坑（9/14 修复）
- **段落软换行必须保行数**：pdf2docx 的多行段落含 `w:br`（多 run + 软换行）；LLM 译文常把 `\n` 丢成空格 → 替换后段落少行 → **整页后续内容整体上移**（静電群馬段[2] 少一行 = 后面全部上移 11.6pt，即"所有文字往上偏移"）。`_replace_para_text` 已实现按原文行长度比例切分重建 `w:br`（`_split_like_lines`），不依赖 LLM 是否保留 `\n`
- **边框清洗绝不能改表格结构**：曾做"无线边界→合并两格"，替换循环仍用清洗前的 (ri,ci) 旧索引 → 写入合并格被覆盖（"08月3日"被"14:30"覆盖丢失）。现改为只清线（`_nil_border`）不动结构；`_merge_cell_pair` 已弃用（碎标题合并留待专门机制）
- 布局回归验证法：docx2pdf 转 PDF 后，对比**表格横线 y 列表** + `search_for` 相同文本 bbox（应逐条一致）
- **tab 合并格要拆**：pdf2docx 会把原 PDF 中"无分界线的相邻格"并成一个 gs=N 的跨列格（文本 `A\tB`，如 '場所名\t静電'）→ `_split_tab_cells`（在提取 positions 前调用）按 tab 拆回 gs=1 的格：列宽取 tblGrid、拆缝按原 PDF 线条定边框（无线→nil）、底色用文本 bbox 采样回填（如粉底 #ffccff）、vMerge 延续格要同步拆（否则合并对齐错）。触发条件：gs==tab 段数
- **PDF 版超长译文兜底**：`_plan_group_layout` 在缩到 5pt 仍放不下时，按可容纳行数截断加"…"（防文本溢出与下方内容重叠）
- **模型输出防护**（`qwen_translate`）：输出含提示词特征（`_PROMPT_MARKERS`）或空/暴长（>原文×3+40）→ 回退原文
- **"源==目标语言"时的模型异常**：扫描版场景下"中→中/日→日"（用户选错语言）时，模型会复读英文语料/模板（如 "(Phase 1) System Requirements Research…"，old.py 可复现，产物见 `经改部规则集_异常英文.pdf`）——不正常用法，用户接受无视。**不要加"文本已是目标语言→直接保留原文"的短路**：日译中时，纯汉字日语（"労働災害""静電"）、同形异义词（如'担当'=负责人）"会被误判为中文而不翻译，危及日译中主场景（增加后已移除）

### Word 背景版的坑（血泪教训）
- **`w:drawing` 必须放 `w:r` 内**：`p.add_run()._r.append(parse_xml(tb_xml))`，直接 `p._p.append` 是非法层级，Word 报错打不开
- **文本必须 XML 转义**（`from xml.sax.saxutils import escape`）：译文含 `& < >` 会损坏 docx
- 浮动文本框 docPr id 从大数（100）开始，避免与 python-docx 自动 id 冲突
- 背景会残留原文日文：渲染前需涂白文字区，但**涂白是背景版硬伤**——会遮表格边框、紫色表头等背景颜色，且 TextBox 定位复杂表格会歪。背景版暂不作为主攻方向

### pdf2docx 方案限制（已验证）
- 把 PDF 矢量层（标注/箭头/表格线/紫表头）压扁成**不透明整页 PNG**（如静電群馬 image7 2208×2752 白底整页），白色占 82% 且被图形边界分隔——**flood-fill 去白方案无效**（实测只清 2.1% 外部白），ChatGPT 式 alpha 处理不通用
- 已验证可行方案：**从原 PDF 对"照片+标注区并集"局部渲染合成图**（PDF 渲染引擎保证层叠正确），替换 Word 对应图片——这是通解

## 运行 / 调试
- 运行：`.venv\Scripts\python.exe index13.py`（或激活 .venv 后 `python index13.py`）
- **import index13 做测试必须先 mock**：`index13.log_box = _MockLog()`（log_box 为 None 会 AttributeError）；翻译测试 mock `process_pdf_translations` 避免真实调 API 超时
- PowerShell 下 `python -c` 反斜杠/中文转义坑多，复杂逻辑写 .py 脚本文件；脚本含中文输出建议 `$env:PYTHONIOENCODING='utf-8'`
- 如果在编写代码过程中发生了严重的失误，无论怎么修复，运行结果也无法回到正轨，请直接停手，与其修不了硬修，让用户使用撤销指令撤销本轮做的更改往往更方便。

## 测试文件与命名
- 样例 PDF：`证据清单.pdf`（纯表格）、`【静電群馬】労働災害発生報告書...pdf`（含 6 照片+矢量标注层，测图片问题）
- 输出命名：`名字_{目标语言}.pdf/docx`（如 `证据清单_日语.docx`）、背景版 `名字_{目标语言}_背景版.docx`
- 验证 Word 文件合法性：zipfile 读 `word/document.xml` + lxml 严格解析 + python-docx 打开（但**不能代替真实 Word 验证**——schema 问题只有 Word 能发现）

## 交互要求
- 全程中文思考和回答；只有用户让"先检查/确认再做"时，先给结论让用户确认再动代码

## powershell的坑
powershell作为工具链可以用，但是不能拿来写或者改文本，会导致严重的乱码

