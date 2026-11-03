# PDF 翻译程序改进建议

## 一、总体判断

目前不建议直接替换 PyMuPDF，也不建议立刻把 RapidOCR 整套换掉。

你的程序已经不是“简单抽取 PDF 文本再覆盖写回”，而是已经围绕 PyMuPDF 构建了一套较完整的 PDF 版面处理逻辑，包括：

- `get_text("dict")` 提取文本、字号、颜色、位置等信息；
- `find_tables()` 尝试识别表格单元格；
- 行级、单元格级文本合并；
- 样式突变拆组；
- 图片区域拆组；
- 背景色采样与原文覆盖；
- 动态字号与换行；
- 页面旋转坐标修正；
- PDF 译文版；
- `pdf2docx` 可编辑 Word 版；
- “整页背景 + 浮动译文文本框”的高保真 Word 版；
- 扫描 PDF 的 RapidOCR 识别流程。

因此，目前系统的主要瓶颈已经不再是“PDF 解析库选错了”，而是：

1. 翻译服务耗时；
2. PDF 与 pdf2docx 之间的文本映射不统一；
3. OCR 的文档级分流过于粗糙；
4. 扫描件缺少版面结构识别；
5. 文本排版越来越接近自行实现一套 layout engine；
6. 缺少系统化的 PDF 回归测试集。

---

## 二、先看当前耗时：真正慢的不是 PDF 解析

这次日志总耗时约 **217.54 秒**：

| 环节 | 耗时 | 判断 |
|---|---:|---|
| PDF 分析 | 0.11 s | 已经非常快，几乎没有优化价值 |
| 45 条文本翻译 | 141.91 s | **第一大瓶颈** |
| PDF 写回 | 1.33 s | 很快 |
| “pdf2docx 转 Word” | 72.38 s | **当前计时有较大误导性** |
| 背景版 DOCX | 1.23 s | 非常快 |

需要特别注意的是：

```text
[pdf2docx转Word] 72.38s
```

这个时间并不等于 `pdf2docx` 本身用了 72.38 秒。

当前 `create_docx_pdf2docx()` 的实际流程是：

```text
PDF
→ pdf2docx 转换
→ 打开 DOCX
→ 再次扫描 DOCX 文本
→ 检查 cache
→ 发现 3 条未命中
→ 再调用一次 LLM 翻译
→ 替换 Word 文本
→ 清理表格边框
→ 保存 DOCX
```

所以 72.38 秒实际上包含：

```text
pdf2docx_convert_time
+ docx_parse_time
+ cache_match_time
+ extra_translation_time
+ replace_time
+ save_time
```

由于前一个 45 条翻译批次已经用了 141.91 秒，后面 3 条文本重新调用模型，很可能占据了这 72 秒中的相当部分。

### 建议

首先把计时拆开：

```python
t0 = time.time()
cv.convert(tmp_path)
t1 = time.time()
log(f"[pdf2docx纯转换] {t1 - t0:.2f}s")

doc = Document(tmp_path)
t2 = time.time()
log(f"[DOCX解析] {t2 - t1:.2f}s")

# cache matching
t3 = time.time()
log(f"[DOCX缓存匹配] {t3 - t2:.2f}s")

# extra translation
t4 = time.time()
log(f"[DOCX补翻译] {t4 - t3:.2f}s")

doc.save(output)
t5 = time.time()
log(f"[DOCX保存] {t5 - t4:.2f}s")
```

否则现在很难判断到底应该优化 `pdf2docx`，还是应该优化第二次 LLM 请求。

---

## 三、PyMuPDF 不建议换成 pdfminer.six 或 pdfplumber

### 1. PyMuPDF 对当前程序的价值不仅是“简单好用”

对你的任务而言，PyMuPDF 最大的优势是：

```text
文本解析
+ 坐标系统
+ 字体属性
+ 图片定位
+ 页面渲染
+ 表格检测
+ PDF 修改
+ PDF 保存
```

这些功能全部位于同一个 API 和坐标体系内。

你的现有代码已经深度使用：

```python
page.get_text("dict")
page.find_tables()
page.get_images()
page.get_image_rects()
page.get_pixmap()
page.insert_text()
page.draw_rect()
```

如果把文本解析部分换成其他库，最终写回 PDF 时仍然很可能需要回到 PyMuPDF。

---

### 2. pdfminer.six 更适合什么

`pdfminer.six` 的优势主要是底层文本解析。

它可以提供：

```text
LTPage
→ LTTextBox
→ LTTextLine
→ LTChar
```

因此对于以下异常 PDF，可能比 PyMuPDF 更有帮助：

- 字符顺序异常；
- 特殊字体编码；
- 竖排文字；
- 字符级字体属性分析；
- 极端复杂的文本对象。

但它主要是解析工具，不适合替代你当前整套 PDF 编辑与渲染流程。

因此更适合做：

> 疑难页面 fallback parser

而不是主引擎。

---

### 3. pdfplumber 更适合什么

`pdfplumber` 的优势主要是：

- 字符级坐标；
- 线条与矩形；
- 表格分析；
- 可视化调试；
- 版面问题排查。

但它同样不是 PDF 写回引擎。

而且 PyMuPDF 当前的表格识别能力已经借鉴/移植了 pdfplumber 的相关思路，因此单纯为了“表格更准”而把整个程序迁移到 pdfplumber，收益很可能有限。

### 推荐架构

```text
主引擎：
PyMuPDF

异常 PDF：
├─ pdfplumber：表格 / 版面 debug
└─ pdfminer.six：字符级异常 fallback
```

不要默认把三个库全部跑一遍。

---

## 四、当前一个明确问题：PDF 的“唯一文本”实际上没有真正去重

你当前类似这样的逻辑：

```python
for group in groups:
    group_text = group["text"]
    idx = len(unique_texts)
    unique_texts.append(group_text)
    text_to_id[group_text] = idx
    group["id"] = idx
```

即使两个 group 的文本完全一样，也仍然都会 append 到 `unique_texts`。

例如：

```text
审核
审核
审核
```

理论上应该只翻译一次，但当前可能形成 3 个待翻译 item。

### 建议修改

```python
for group in groups:
    group_text = group["text"]

    if group_text not in text_to_id:
        idx = len(unique_texts)
        text_to_id[group_text] = idx
        unique_texts.append(group_text)
    else:
        idx = text_to_id[group_text]

    group["id"] = idx
```

扫描 PDF 的 OCR 路径也应采用相同处理。

### 收益

- 减少 LLM 输入字符数；
- 减少翻译时间；
- 降低 API 请求量；
- 提高缓存命中率；
- 同一术语翻译更一致。

这是一个低风险、高收益改动。

---

## 五、真正的架构问题：目前存在两个“文本世界”

现在实际上有两套独立文本解析流程。

### 路径 A

```text
PDF
→ PyMuPDF
→ group
→ source_text
→ 翻译
```

### 路径 B

```text
PDF
→ pdf2docx
→ DOCX
→ paragraph / cell
→ source_text
→ cache lookup
```

问题是，同一句话经过两套解析器以后，不一定完全一样。

例如 PDF 中：

```text
安全管理
实施情况
```

pdf2docx 可能变成：

```text
安全管理实施情况
```

或者出现：

```text
安全管理 实施情况
```

还可能有：

- `\n` / `\r` 差异；
- 全角 / 半角差异；
- Unicode 兼容字符差异；
- 多余空格；
- span 合并差异；
- 连字符差异；
- 表格单元格拆分差异。

因此就会出现：

```text
PDF已经翻译过
↓
DOCX再次解析
↓
cache miss
↓
第二次调用LLM
```

---

## 六、建议建立统一的中间层 Layout IR

建议逐步把 PDF 解析结果统一成一个中间数据结构：

```text
LayoutNode
├─ node_id
├─ page
├─ bbox
├─ source_text
├─ normalized_text
├─ translated_text
├─ font_size
├─ font_family
├─ bold
├─ italic
├─ color
├─ object_type
├─ confidence
└─ source_engine
```

之后整个流程变成：

```text
PDF
 ↓
Layout IR
 ↓
一次统一翻译
 ↓
├─ PDF Renderer
├─ Background DOCX Renderer
└─ pdf2docx Matcher
```

这样 `pdf2docx` 不再重新决定“这句话有没有翻译过”，而是负责：

> 将 DOCX 中的文本匹配回已经存在的 LayoutNode。

---

## 七、优先加入文本归一化匹配

第一步不需要复杂 NLP。

先做：

```python
import re
import unicodedata

def normalize_text(text):
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\r", "\n")
    text = re.sub(r"\s+", " ", text)
    return text.strip()
```

匹配顺序可以设计为：

```text
1. 原始文本 exact match
2. normalized exact match
3. 去空格 match
4. 连续 PDF node 拼接 match
5. fuzzy match
6. 最后才重新调用 LLM
```

例如：

```text
PDF node 1 = 安全管理
PDF node 2 = 实施情况

DOCX paragraph = 安全管理实施情况
```

可以组合两个连续 node 去匹配。

这样比“匹配失败就重新翻译”更合理。

---

## 八、真正最大的速度瓶颈是 LLM 翻译

当前翻译逻辑基本是：

```text
构造 batch
→ HTTP 请求
→ 等待完整结果
→ 下一个 batch
```

全部串行阻塞。

当前配置：

```python
MAX_BATCH_CHARS = 4000
SINGLE_TEXT_THRESHOLD = 1000
```

但 `4000` 是否是最优 batch size，目前没有实测依据。

### 建议做 benchmark

测试：

| 单批字符数 | 并发 1 | 并发 2 | 并发 4 |
|---:|---:|---:|---:|
| 500 | 测试 | 测试 | 测试 |
| 1000 | 测试 | 测试 | 测试 |
| 2000 | 测试 | 测试 | 测试 |
| 4000 | 测试 | 测试 | 测试 |

记录：

```text
总 wall time
平均单批耗时
失败率
JSON解析失败率
GPU显存
模型吞吐
```

如果你的 Qwen 服务支持 continuous batching，多批并行很可能优于单个超大 batch。

但这必须实测，不能直接假设。

---

## 九、让 pdf2docx 与 LLM 翻译并行

现在逻辑大致是：

```text
PDF分析
↓
等待LLM 141秒
↓
PDF写回
↓
再开始pdf2docx
```

但 `pdf2docx` 并不依赖译文。

因此完全可以：

```text
               ┌─ LLM翻译
PDF分析完成 ────┤
               └─ pdf2docx转换
```

两个任务同时进行。

待翻译结果返回后：

```text
LLM译文
+
pdf2docx生成的DOCX
↓
统一匹配与写回
```

如果 pdf2docx 本身需要 20～30 秒，这部分耗时就可以被隐藏在 LLM 等待期间。

---

## 十、翻译缓存应该持久化

目前 cache 主要存在内存中。

程序关闭以后：

```text
cache → 消失
```

建议改为 SQLite。

### 推荐 key

```text
normalized_source_text
target_language
model_name
prompt_version
```

数据库例如：

```sql
CREATE TABLE translations (
    source_hash TEXT,
    normalized_source TEXT,
    target_lang TEXT,
    model_name TEXT,
    prompt_version TEXT,
    translated_text TEXT,
    PRIMARY KEY (
        source_hash,
        target_lang,
        model_name,
        prompt_version
    )
);
```

这样：

- 同一份文件再次运行；
- 不同文件出现相同术语；
- pdf2docx 二次匹配；
- 批量处理同类企业文档；

都可以复用旧翻译。

---

## 十一、PDF 排版建议继续深挖 PyMuPDF，而不是换解析器

当前你已经开始自行实现：

```text
_wrap_text
→ get_text_length
→ 算字体
→ 算行数
→ 算baseline
→ insert_text
```

这实际上已经开始接近自己写一个文字排版引擎。

尤其当前粗体是通过再次偏移写一遍文字来模拟，这种方案会越来越难维护。

### 建议测试 `Page.insert_htmlbox()`

复杂文本可以考虑：

```text
短标签 / 简单单行
→ insert_text

长文本 / 多行 / 复杂样式
→ insert_htmlbox
```

`insert_htmlbox()` 更适合：

- 自动换行；
- 字体；
- bold；
- italic；
- 颜色；
- 对齐；
- HTML/CSS；
- 复杂 Unicode shaping；
- 在目标 Rect 内缩放。

这可能比继续增加 `_wrap_text()` 的特殊规则更稳定。

---

## 十二、扫描 PDF 的字体大小不要过度依赖 OCR 框高度

目前类似：

```python
fs = max(r.height * 0.78, 6)
```

问题在于 OCR bbox 自身就是有噪声的。

同一字号文字可能因为：

- 字符不同；
- 边缘检测；
- 图像模糊；
- 日文假名与汉字高度不同；
- bbox上下留白不同；

得到不同高度。

所以：

```text
OCR框高度 ≠ 原始字体大小
```

### 更合理的策略

OCR 主要回答：

> 文字在哪里？

字体大小由排版阶段决定：

```text
bbox
+ 文本长度
+ 当前区域类型
+ 基准字号
↓
自动fit
```

例如：

```text
标题 → 基准 14pt
正文 → 基准 10.5pt
表格 → 基准 9pt
小标注 → 基准 7.5pt
```

然后动态缩小到能放进 bbox 为止。

这比不断调整 `0.78` 更稳。

---

## 十三、扫描 PDF 当前最应该改的是“什么时候 OCR”

目前判断方式大致是：

```python
total_text = 0

for page in pdf:
    total_text += len(page.get_text().strip())

return total_text > 20
```

这意味着：

> 整份 PDF 只要总共超过 20 个文本字符，就判定为文本型 PDF。

这是比较危险的。

例如：

```text
第1页：真正文本
第2页：扫描图片
第3页：扫描图片
第4页：扫描图片
```

第1页只要超过20个字符，整个文件都不会进入 OCR 模式。

还有一种常见情况：

```text
扫描页
+
扫描软件自动加入的页码/隐藏OCR层/水印
```

也可能误判成普通文本 PDF。

---

## 十四、建议改为逐页、最终逐区域的混合模式

理想流程：

```text
Page 1
native text 足够
→ PyMuPDF

Page 2
没有文本 + 大图覆盖95%
→ OCR

Page 3
部分native text + 一个扫描表格
→ native text + OCR局部区域
```

页级判断指标可以包括：

```text
native_text_chars
image_area_ratio
number_of_images
text_bbox_area
page_area
vector_object_count
```

例如：

```python
if text_chars < 20 and image_area_ratio > 0.7:
    mode = "ocr"
elif text_chars > threshold:
    mode = "native"
else:
    mode = "hybrid"
```

这对真实办公 PDF 的鲁棒性提升，会明显大于单纯更换 OCR 引擎。

---

## 十五、RapidOCR 不必急着替换，但配置值得检查

目前你使用的是 RapidOCR 日文模型。

整体方向没有问题。

但应该确认当前版本的 RapidOCR：

- detection language 配置；
- recognition language 配置；
- OCRVersion；
- model type；
- inference engine；
- 日文模型版本。

尤其要确认 `Det.lang_type` 与 `Rec.lang_type` 是否使用了正确的枚举类型。

建议显式区分：

```text
LangDet
LangRec
```

而不是都直接使用 `LangRec.JAPAN`。

具体参数需要依据你当前安装的 RapidOCR 版本核对。

---

## 十六、是否应该直接换成 PaddleOCR

不建议仅因为“RapidOCR识别不完美”就整个换成 PaddleOCR。

因为 RapidOCR 本身就大量使用 PaddleOCR 系列模型。

如果只是：

```text
RapidOCR + PP-OCR模型
```

换成：

```text
PaddleOCR + 类似PP-OCR模型
```

识别精度未必发生质变。

### 更合理的方法

建立自己的 OCR benchmark。

取真实困难样本：

```text
30～50 页
```

覆盖：

- 小字号日文；
- 模糊扫描；
- 倾斜扫描；
- 表格；
- 图片标注；
- 低对比度文字；
- 彩色背景；
- 竖排文字。

测试：

```text
RapidOCR 当前模型
RapidOCR PP-OCRv5
RapidOCR PP-OCRv6 small
RapidOCR PP-OCRv6 medium
```

比较：

| 指标 | 含义 |
|---|---|
| Character Error Rate | 字符识别错误率 |
| Missing Box Rate | 漏文字框比例 |
| False Positive Rate | 误识别比例 |
| Bounding Box Stability | 位置是否稳定 |
| Page Time | 单页耗时 |
| Memory | 内存 / 显存消耗 |

然后再决定是否换模型或换框架。

---

## 十七、真正值得考虑 PaddleOCR 的地方：结构分析

当前扫描 PDF 基本是：

```text
OCR
→ bbox
→ 按 y 排序
→ 按 gap 合并
→ 当作一行
```

OCR 本身再准，也不知道：

- 哪个是标题；
- 哪个是正文；
- 哪些文字属于同一个表格单元格；
- 哪些文字属于照片标注；
- 哪些内容属于两栏排版；
- 阅读顺序是什么。

因此扫描件真正的天花板其实不是 OCR recognition，而是：

> Layout Analysis

这里才是 PP-StructureV3 之类系统的价值。

它可以帮助处理：

```text
layout detection
table recognition
reading order
formula
chart
```

### 推荐方式

不要所有页面都调用。

采用：

```text
普通页
→ PyMuPDF / RapidOCR

复杂低置信页面
→ PP-StructureV3 fallback
```

例如只有检测到：

```text
OCR框很多
+
多列
+
大量线条
+
疑似表格
```

才触发结构模型。

---

## 十八、OCR 可以增加“两阶段识别”

当前扫描版大概是：

```text
整页200 DPI
→ RapidOCR
→ score < 0.45直接丢弃
```

问题是小字很容易：

```text
低置信度
→ 被丢弃
→ 最终文件完全没有这段文字
```

### 推荐流程

```text
第一阶段
整页 180～200 DPI
快速 OCR
↓
筛出可疑区域
├─ confidence低
├─ 字符串异常短
├─ bbox异常
├─ 疑似小字号
└─ 疑似漏识别
↓
第二阶段
只对该区域 clip
300～400 DPI重新渲染
↓
二次OCR
```

这样：

- 大部分页面维持高速；
- 小字区域获得更高精度；
- 不需要整页400 DPI OCR。

---

## 十九、扫描件最好尽量直接使用原始嵌入图像

如果 PDF 页面本质上是：

```text
一张300 DPI / 600 DPI扫描JPEG
```

现在你可能在做：

```text
原始扫描图
→ PDF
→ PyMuPDF重新渲染200 DPI
→ OCR
```

等于主动降采样。

如果可以识别到：

```text
单张大图覆盖页面绝大多数区域
```

可以尝试：

```text
提取原始嵌入图片
→ 直接OCR
```

这样可能保留更多小字细节。

需要注意：

- 页面 rotation；
- cropbox；
- image transform；
- 多图拼版；

因此适合作为扫描页专用优化，而不是所有PDF默认使用。

---

## 二十、背景版 Word 的方案建议保留

背景版 Word 的核心思路其实非常合理：

> 原 PDF 页面作为视觉真值，只替换文字层。

你当前方案：

```text
PDF页面
→ 高清PNG背景
→ 清除原文字区域
→ Word页面背景
→ 在对应位置添加译文浮动文本框
```

对以下复杂内容非常有价值：

- 图片；
- 箭头；
- 标注；
- CAD式线条；
- 特殊矢量图；
- 复杂表格；
- 图文叠加。

任何“真正恢复成Word语义结构”的方案都存在不确定性。

因此两种 Word 输出应该继续保留：

```text
可编辑结构版
= best effort

背景保真版
= fidelity fallback
```

这是合理的产品设计，不是临时补丁。

---

## 二十一、可以加入“页面复杂度评分”

未来可以自动判断：

> 这个 PDF 是否值得跑 pdf2docx。

例如统计：

```text
image_area_ratio
vector_count
annotation_count
text_image_overlap
table_count
rotation
column_count
```

构造一个复杂度：

```text
complexity_score
```

如果很低：

```text
生成：
PDF + pdf2docx版 + 背景版
```

如果很高：

```text
提示：
该页面存在复杂图文叠加，背景版预计更可靠
```

甚至可以选择性跳过 pdf2docx，节省时间。

---

## 二十二、文本型 PDF 可以考虑真正删除原文字

目前主要使用：

```text
draw_rect()
→ 用背景色盖住原文
→ 再写译文
```

缺点是：

1. 原文字仍然存在于 PDF 文本层；
2. 搜索可能搜到原文；
3. 复制可能复制出原文；
4. 背景复杂时可能看到覆盖痕迹。

对文本型 PDF，可以测试 PyMuPDF redaction：

```text
标记原文区域
→ apply_redactions()
→ 删除原文字
→ 保留图片/矢量对象
→ 插入译文
```

然后：

```text
文本PDF
→ redaction + insert_htmlbox

扫描PDF
→ 图像背景处理 + OCR + 译文覆盖
```

两类 PDF 使用不同策略，会比全部依赖“涂白”更加干净。

---

## 二十三、推荐的优化优先级

| 优先级 | 改动 | 主要收益 |
|---|---|---|
| **P0** | 修复 PDF / OCR 真正去重 | 减少翻译量和模型调用 |
| **P0** | 拆分 `pdf2docx转Word` 的计时 | 找出真正性能瓶颈 |
| **P0** | PDF 与 pdf2docx 使用 normalized text 匹配 | 尽量消灭第二轮 LLM |
| **P0** | SQLite 持久化翻译缓存 | 重复文档和术语几乎零成本 |
| **P0** | benchmark LLM batch size 与并发 | 当前最大速度优化来源 |
| **P1** | pdf2docx 与 LLM 翻译并行 | 隐藏转换耗时 |
| **P1** | 文档级 OCR 判断改成逐页 / 混合模式 | 大幅提高混合 PDF 鲁棒性 |
| **P1** | 检查 RapidOCR 模型及 `LangDet/LangRec` 配置 | 提高 OCR 稳定性 |
| **P1** | 200 DPI 快速 OCR + 低置信区域高 DPI 复识 | 提高小字识别率 |
| **P1** | 复杂文本改用 `insert_htmlbox()` | 提高排版和长译文稳定性 |
| **P1** | Native PDF 使用 redaction 删除原文字 | 避免隐藏原文和覆盖伪影 |
| **P2** | PP-StructureV3 作为复杂扫描页 fallback | 提升表格和版面理解 |
| **P2** | pdfplumber / pdfminer.six 作为疑难页 fallback | 提高特殊 PDF 兼容性 |

---

## 二十四、最推荐首先修改的三项

如果不希望大幅重构，建议首先只改三个地方。

### 1. 真正去重 PDF / OCR 文本

风险非常低，并且直接减少模型输入量。

### 2. 消灭 pdf2docx 的第二轮 LLM 翻译

核心做法：

```text
原文
→ Unicode NFKC
→ whitespace normalize
→ exact normalized match
→ node组合匹配
→ fuzzy match
→ 最后才重新翻译
```

同时拆开 pdf2docx 内部计时。

### 3. OCR 改成逐页混合模式

从：

```text
整个PDF：
文本版 OR 扫描版
```

改成：

```text
每一页：
native / OCR / hybrid
```

这对真实业务 PDF 的稳定性提升很大。

---

# 二十五、下一阶段最重要的工程改进：建立 PDF 回归测试集

目前开发模式更接近：

```text
出现一个坏样本
→ 找问题
→ 增加一个规则
→ 等待下一个坏样本
```

程序进入稳定阶段以后，这种方法的收益会逐渐降低，并容易出现：

```text
修好了PDF A
↓
却把PDF B改坏
```

建议建立固定测试集，例如 30～50 份 PDF：

```text
01_normal_text.pdf
02_border_table.pdf
03_borderless_table.pdf
04_photo_annotation.pdf
05_rotated_page.pdf
06_vertical_japanese.pdf
07_low_res_scan.pdf
08_skew_scan.pdf
09_text_scan_mixed.pdf
10_multi_column.pdf
11_vector_heavy.pdf
12_colored_cells.pdf
13_small_font.pdf
14_nested_table.pdf
...
```

每次修改代码后全部自动跑一次。

---

## 二十六、建议自动记录的回归指标

### 性能指标

```text
PDF parse time
OCR time
LLM time
PDF write time
pdf2docx conversion time
DOCX processing time
background DOCX time
total time
```

### 文本指标

```text
source text count
unique text count
cache hit rate
OCR box count
low confidence OCR count
second translation count
```

### 排版指标

```text
minimum font size
overflow count
out-of-page count
text-image overlap count
untranslated text count
```

最终日志可以类似：

```text
PDF: example.pdf

Pages: 12
Native pages: 8
OCR pages: 3
Hybrid pages: 1

Raw groups: 183
Unique groups: 126
Cache hit: 81
LLM translated: 45

OCR low confidence: 7
Second-pass OCR: 5
Recovered: 4

Overflow: 2
Min font size: 6.5pt

PDF parse: 0.31s
OCR: 3.82s
LLM: 41.71s
PDF write: 1.92s
pdf2docx: 8.31s
DOCX replace: 0.73s
Background DOCX: 2.18s

Total: 48.67s
```

这样下一次收到“不佳样本”时，就不再只是依赖肉眼观察，而是能够快速判断失败发生在哪一层。

---

# 二十七、最终建议架构

长期可以逐步演化成：

```text
                    PDF
                     │
             ┌───────▼────────┐
             │ Page Classifier │
             └───────┬────────┘
                     │
         ┌───────────┼───────────┐
         │           │           │
      Native       Scanned      Hybrid
         │           │           │
     PyMuPDF      RapidOCR    PyMuPDF + OCR
         │           │           │
         └───────────┼───────────┘
                     ▼
                  Layout IR
                     │
             Normalized Text
                     │
            Persistent Cache
                     │
               LLM Translation
                     │
         ┌───────────┼──────────────┐
         │           │              │
         ▼           ▼              ▼
   PDF Renderer   pdf2docx      Background DOCX
   redaction +       │          bitmap + textbox
   htmlbox           │
                     ▼
               Layout Matcher
```

这样整个系统的职责会变得更清晰：

```text
OCR负责：
看见文字

Layout负责：
判断文字在哪里、属于什么结构

Translation负责：
翻译一次

Renderer负责：
决定如何输出

Cache负责：
避免重复工作
```

---

# 二十八、结论

目前没有充分理由把 PyMuPDF 替换为 pdfminer.six 或 pdfplumber。

更合理的技术路线是：

> **继续以 PyMuPDF 为 PDF 主引擎，将 pdfminer.six / pdfplumber 作为疑难 PDF 的辅助解析器；保留 RapidOCR，但升级 OCR 分流、模型配置和二次识别机制；把主要开发精力转向统一 Layout IR、文本匹配、持久缓存、并行处理和回归测试。**

当前日志已经说明：

> **PDF 解析和写回本身并不是性能瓶颈，LLM 翻译才是。**

因此，现阶段最值得优先解决的不是：

> “有没有比 PyMuPDF 更强的库？”

而是：

> “如何让同一段文本只解析一次、只翻译一次、只布局一次，同时让多个输出格式共享同一份结构信息。”

如果这一步完成，你的程序会从“多个翻译脚本拼接在一起”逐步变成一个真正的：

> **文档解析 → 统一翻译 → 多格式渲染系统。**

---

## 参考资料

- PyMuPDF Documentation  
  https://pymupdf.readthedocs.io/

- PyMuPDF `Page` API / Redaction  
  https://pymupdf.readthedocs.io/en/latest/page.html

- pdfminer.six Layout Analysis  
  https://pdfminersix.readthedocs.io/en/latest/tutorial/extract_pages.html

- pdfplumber  
  https://github.com/jsvine/pdfplumber

- RapidOCR Documentation  
  https://rapidai.github.io/RapidOCRDocs/

- RapidOCR Model List  
  https://rapidai.github.io/RapidOCRDocs/main/model_list/

- PaddleOCR PP-StructureV3  
  https://www.paddleocr.ai/main/en/version3.x/algorithm/PP-StructureV3/PP-StructureV3.html
