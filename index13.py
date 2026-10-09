# PDF翻译方案改进，用pdf2docx处理，并加入背景板功能。

import os
import json
import re
import shutil
from xml.sax.saxutils import escape
import urllib.request
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import time
import threading
import queue
import stat
import sys
import fitz
fitz.no_recommend_layout()  # 关闭 find_tables 的控制台"pymupdf_layout"推荐提示

import xlwings as xw
from docx import Document
from docx.shared import Pt
from pptx import Presentation

# ==========================
# 配置
# ==========================

# API_URL = "http://10.237.192.42:8000/v1/chat/completions"
# MODEL_NAME = "qwen36-35b-a3b"

# 备用1
# API_URL = "http://10.237.192.46:8000/v1/chat/completions"
# MODEL_NAME = "qwen3.6-35b-a3b"
# 备用2
# API_URL = "http://10.237.192.41:11434/v1/chat/completions"
# MODEL_NAME = "qwen3.6:35b"
# 备用3
API_URL = "http://10.237.192.48:8000/v1/chat/completions"
MODEL_NAME = "qwen3.6-35b-a3b"

REQUEST_TIMEOUT = 1200

# ★ 模型思考模式开关（重要）：
#   False = 关闭模型"思考"——翻译速度快约 10 倍以上（实测 39.6s → 2.7s/批）；
#   True  = 恢复默认思考模式（如果发现关闭后翻译质量明显变差，改回 True 即可）；
#   本参数通过请求体 chat_template_kwargs.enable_thinking 传给 vLLM 服务端。
ENABLE_THINKING = False

# 批量翻译的并发线程数（并发逐条翻译，控制总耗时；如服务端压力大可调小）
BATCH_CONCURRENCY = 4

MAX_BATCH_CHARS = 4000
SINGLE_TEXT_THRESHOLD = 1000

# 全局变量储存
root = None
log_box = None
SOURCE_PATH = ""
SOURCE_TYPE = ""
TARGET_LANG = ""
cache = {}
last_lang = ""
istranslating = False
translation_button = None

# ==========================
# GUI
# ==========================
LANGUAGES = ["英文", "中文", "日语", "泰语", "韩语", "法语", "德语", "西班牙语", "俄语", "意大利语", "葡萄牙语"]

def get_user_config():

    global root, log_box
    root = tk.Tk()
    root.title("Office AI 翻译工具")
    root.geometry("700x560")
    root.resizable(False, False)

    tk.Label(root,text="文件/文件夹:").place(x=15,y=25)

    path_var = tk.StringVar()
    path_entry = tk.Entry(root,textvariable=path_var,width=70)
    path_entry.place(x=100,y=25)

    mode_var = tk.StringVar()
    mode_var.set("file")

    def choose_file():

        filename = filedialog.askopenfilename(
            title="请选择文件",
            filetypes=[
                ("Office文件","*.xlsx *.xlsm *.xls *.docx *.pptx *.pdf"),
                ("所有文件","*.*")
            ]
        )

        if filename:
            path_var.set(filename)
            mode_var.set("file")


    def choose_folder():

        folder = filedialog.askdirectory(title="请选择文件夹")

        if folder:
            path_var.set(folder)
            mode_var.set("folder")

    tk.Button(root,text="选择文件",command=choose_file).place(x=100,y=60)

    tk.Button(root,text="选择文件夹",command=choose_folder).place(x=200,y=60)

    tk.Label(root,text="目标语言:").place(x=15,y=105)

    lang_box = ttk.Combobox(root,values=LANGUAGES,width=20)

    lang_box.place(x=100,y=105)

    lang_box.current(2)

    # 模型思考模式：默认关闭（提速约 10 倍）；若发现翻译质量差可勾选恢复
    think_var = tk.BooleanVar(master=root, value=False)
    tk.Checkbutton(root, text="深度思考（开启后会显著降速）",
                   variable=think_var).place(x=320, y=103)

    tk.Label(root,text="API地址:").place(x=15,y=135)

    api_var = tk.StringVar(value=API_URL)
    api_entry = tk.Entry(root,textvariable=api_var,width=70)
    api_entry.place(x=100,y=135)

    tk.Label(root,text="模型名称:").place(x=15,y=165)

    model_var = tk.StringVar(value=MODEL_NAME)
    model_entry = tk.Entry(root,textvariable=model_var,width=70)
    model_entry.place(x=100,y=165)

    tk.Label(root,text="运行日志:").place(x=15,y=200)

    log_box = tk.Text(root,width=82,height=16)

    log_box.place(x=15,y=230)

    def start():

        global SOURCE_PATH
        global SOURCE_TYPE
        global TARGET_LANG
        global istranslating
        global API_URL
        global MODEL_NAME
        global ENABLE_THINKING

        SOURCE_PATH = path_var.get()
        SOURCE_TYPE = mode_var.get()
        TARGET_LANG = lang_box.get()
        ENABLE_THINKING = bool(think_var.get())

        # 允许用户修改API地址与模型名称（留空则保持原有配置）
        if api_var.get().strip():
            API_URL = api_var.get().strip()
        if model_var.get().strip():
            MODEL_NAME = model_var.get().strip()

        if not istranslating:
            istranslating = True
            threading.Thread(target=translate_worker,daemon=True).start()
            translation_button.config(state=tk.DISABLED, text="翻译中...", bg="gray")

    def on_close():
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)

    global translation_button
    translation_button = tk.Button(root,text="开始翻译",width=15,command=start)
    translation_button.place(x=280,y=450)

    root.after(120, _poll_log)   # 启动日志队列轮询（必须在 mainloop 之前）
    root.mainloop()

# ==========================
# 日志函数（线程安全：工作线程只写队列，主线程轮询刷新 Tk 控件）
# ==========================
_log_queue = queue.Queue()


def log(*args):

    msg = " ".join(
        str(x)
        for x in args
    )

    _log_queue.put(msg)


def _poll_log():
    """主线程日志轮询：从队列取消息刷新到日志框（避免工作线程直接操作 Tk）"""
    try:
        while True:
            msg = _log_queue.get_nowait()
            log_box.insert(tk.END, msg + "\n")
            log_box.see(tk.END)
    except queue.Empty:
        pass
    except Exception:
        pass
    try:
        if root is not None:
            root.after(120, _poll_log)
    except Exception:
        pass

# ==========================
# 翻译函数
# ==========================
def qwen_translate(text, target_lang="中文"):

    if not text or not str(text).strip():return ""

    text = str(text).strip()
    messages = [
        {"role": "system", "content": f"你是一个专业翻译引擎，请将用户输入翻译成{target_lang}。保持原意。只输出翻译结果，不要解释，不要添加引号。"},
        {"role": "user", "content": text}
    ]
    try:
        req = urllib.request.Request(
            API_URL,
            data=json.dumps({"model": MODEL_NAME, "messages": messages, "temperature": 0.1,
                             "stream": False,
                             "chat_template_kwargs": {"enable_thinking": ENABLE_THINKING}}).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as resp:
            result = json.loads(resp.read().decode("utf-8"))
        return result["choices"][0]["message"]["content"].strip()
    except Exception as e:
        log(f"翻译失败：{e}")
        return text

def qwen_translate_batch(text_list, target_lang="中文"):
    """批量翻译接口（线程池并发逐条）。
    说明：实测模型对"长批量 JSON"的项-内容对应不稳定（会漏项/整体错位，
    如忽略"(1/2)"导致后续译文前移），故改为并发单条翻译——逐条独立、绝无错位；
    靠并发控制总耗时（单条约 0.2~2 秒）。
    """
    if not text_list:
        return {}
    from concurrent.futures import ThreadPoolExecutor

    def _one(item):
        idx, text = item
        return idx, qwen_translate(text, target_lang)

    out = {}
    try:
        with ThreadPoolExecutor(max_workers=BATCH_CONCURRENCY) as ex:
            for idx, trans in ex.map(_one, text_list):
                out[idx] = trans
    except Exception as e:
        log("并发翻译异常，退回串行：", e)
        out = {idx: qwen_translate(text, target_lang) for idx, text in text_list}
    return out

# ==========================
# 通用翻译流程
# ==========================


def process_translations(unique_texts,text_to_id):
    """统一的翻译流程，返回 text -> translation 的映射（入参已过滤：text 均在缓存未命中）"""
    log(f"\n发现 {len(unique_texts)} 条唯一文本需要翻译（缓存未命中）")

    id_to_translation = {}
    # 页码/纯符号/纯数字类文本（如 "(1/2)"、"2009-1-1"）不送 LLM：
    # 模型会忽略或省略这类项，导致同批后续译文整体前移错位（已实测复现），直接原样保留
    symbol_re = re.compile(r'^[\s\d\W_]+$')
    # 两阶段：先收集批次，再统一发送（便于打印 [i/N] 进度）
    batch = []
    current_chars = 0
    batches = []
    singles = []
    for idx, text in enumerate(unique_texts):
        if symbol_re.match(text):
            id_to_translation[idx] = text
            continue
        if len(text) > SINGLE_TEXT_THRESHOLD:
            singles.append(idx)
            continue
        if current_chars + len(text) > MAX_BATCH_CHARS:
            batches.append(batch)
            batch = []
            current_chars = 0
        batch.append((idx, text))
        current_chars += len(text)
    if batch:
        batches.append(batch)

    total = len(batches) + len(singles)
    for i, b in enumerate(batches, 1):
        log(f"  [{i}/{total}] 发送批次: {len(b)} 条")
        result = qwen_translate_batch(b, TARGET_LANG)
        id_to_translation.update(result)
    for i, idx in enumerate(singles, len(batches) + 1):
        log(f"  [{i}/{total}] 长文本单独翻译 [{idx}]")
        id_to_translation[idx] = qwen_translate(unique_texts[idx], TARGET_LANG)

    return {text: id_to_translation.get(idx, text) for text, idx in text_to_id.items() if text in text_to_id}

# ==========================
# Excel 处理
# ==========================
def _iter_xls_shapes_com(shapes_com):
    """递归遍历 Excel 形状（含组合形状 GroupItems 嵌套，深度不限）"""
    for shp in shapes_com:
        try:
            if shp.Type == 6:  # msoGroup
                yield from _iter_xls_shapes_com(shp.GroupItems)
                continue
        except Exception:
            pass
        yield shp


def _shape_paragraphs(shp):
    """读取形状（文本框/艺术字）的段落文本，返回 [(段号, 原文)]。
    注：TextRange.Paragraphs 集合在 win32com 动态派发下不可用（"不支持集合"），
    故用 Text 整串按 \\r/\\v 拆分。"""
    try:
        tf2 = shp.TextFrame2
        if tf2.HasText:
            raw = tf2.TextRange.Text
            return [(i, p) for i, p in enumerate(raw.replace('\v', '\r').split('\r'), 1)]
    except Exception:
        pass
    try:
        tf = shp.TextFrame
        if tf.HasText:
            raw = tf.Characters().Text
            return [(i, p) for i, p in enumerate(raw.replace('\v', '\r').split('\r'), 1)]
    except Exception:
        pass
    return []


def _set_shape_paragraph(shp, idx, new_text):
    """写回形状文本的第 idx 段（整串拆分-替换-回写，保持段落结构）"""
    try:
        tf2 = shp.TextFrame2
        if tf2.HasText:
            tr = tf2.TextRange
            parts = tr.Text.replace('\v', '\r').split('\r')
            if 1 <= idx <= len(parts):
                parts[idx - 1] = new_text
                tr.Text = '\r'.join(parts)
            return True
    except Exception:
        pass
    try:
        tf = shp.TextFrame
        if tf.HasText:
            chars = tf.Characters()
            parts = chars.Text.replace('\v', '\r').split('\r')
            if 1 <= idx <= len(parts):
                parts[idx - 1] = new_text
                chars.Text = '\r'.join(parts)
            return True
    except Exception:
        pass
    return False


def process_excel(wb):

    all_cells = []
    shape_items = []
    unique_texts = []
    text_to_id = {}
    total_unique = 0

    for sheet in wb.sheets:

        log(f"扫描工作表:{sheet.name}")
        try:
            if sheet.api.ProtectContents:
                log(f"发现保护工作表:{sheet.name}")

                try:
                    sheet.api.Unprotect()
                    log(f"已解除保护:{sheet.name}")

                except Exception:
                    log(f"无法解除保护:{sheet.name}，可能需要密码")
                    continue

        except Exception:
            pass

        # 形状文本框（含组合嵌套）：与单元格一起收集、统一翻译
        try:
            for shp in _iter_xls_shapes_com(sheet.api.Shapes):
                for pi, raw in _shape_paragraphs(shp):
                    text = raw.replace('\r', '').replace('\v', '').strip()
                    if not text or text.startswith('='):
                        continue
                    shape_items.append((shp, pi, text))
                    if text not in text_to_id:
                        total_unique += 1
                        if text not in cache:
                            idx = len(unique_texts)
                            text_to_id[text] = idx
                            unique_texts.append(text)
        except Exception:
            pass

        used_range = sheet.used_range
        values = used_range.value

        if values is None or len(values)==0:
            log(f"{sheet.name}: 无有效数据")
            continue

        if not isinstance(values, list):
            values = [[values]]

        elif values and not isinstance(values[0], list):

            values = [
                [v]
                for v in values
            ]


        rows = len(values)
        cols = len(values[0]) if rows else 0

        log(f"{sheet.name}翻译范围: {rows}行 {cols}列")

        for r in range(rows):
            for c in range(cols):

                value = values[r][c]
                if not isinstance(value,str):
                    continue

                text=value.strip()

                if (not text or text.startswith("=")):
                    continue

                address = used_range[r,c].address

                all_cells.append((sheet,address,text))

                if (
                    text not in text_to_id
                ):
                    total_unique += 1
                    if text not in cache:
                        idx=len(unique_texts)
                        text_to_id[text]=idx
                        unique_texts.append(text)

    log(f"发现{len(unique_texts)}条Excel文本（缓存未命中）")
    log(f"  Excel文本总数：{total_unique}, Cache命中：{total_unique - len(unique_texts)}")

    translation_map = process_translations(unique_texts, text_to_id)

    for sheet,address,orig in all_cells:

        # 先查本次翻译结果；未命中则查全局缓存（已缓存文本也要写回）
        if orig in translation_map:
            new_text = translation_map[orig]
        else:
            new_text = cache.get(orig, orig)

        if new_text != orig:
            sheet.range(address).value = new_text

    # 写回形状文本框译文（按段落替换，保持段落结构）
    for shp, pi, orig in shape_items:
        if orig in translation_map:
            new_text = translation_map[orig]
        else:
            new_text = cache.get(orig, orig)
        if new_text != orig:
            try:
                _set_shape_paragraph(shp, pi, new_text)
            except Exception:
                pass

    return translation_map

# ==========================
# Docx 处理
# ==========================
def _convert_legacy_office(path):
    """旧格式(.doc/.ppt) → 临时新格式(.docx/.pptx)，使用 Office COM 转换。
    失败抛异常（调用方捕获并提示原因）；成功返回临时文件路径（调用方负责删除）。
    纯 Python 无可靠库可转换旧二进制 Office 格式，故采用 COM（本机需装 Office）。"""
    import tempfile
    import win32com.client

    ext = os.path.splitext(path)[1].lower()
    fd, tmp = tempfile.mkstemp(suffix='.docx' if ext == '.doc' else '.pptx')
    os.close(fd)
    os.remove(tmp)  # COM SaveAs 要求目标文件不存在

    if ext == '.doc':
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            d = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
            d.SaveAs2(os.path.abspath(tmp), FileFormat=16)  # 16 = wdFormatXMLDocument(.docx)
            d.Close(False)
        finally:
            word.Quit()
    elif ext == '.ppt':
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            try:
                ppt.DisplayAlerts = 1  # ppAlertsNone：抑制"是否替换"弹窗
            except Exception:
                pass
            p = ppt.Presentations.Open(os.path.abspath(path), ReadOnly=True, WithWindow=False)
            p.SaveAs(os.path.abspath(tmp), 24)  # 24 = ppSaveAsOpenXMLPresentation(.pptx)
            p.Close()
        finally:
            ppt.Quit()
    else:
        raise ValueError("不支持的旧格式: %s" % ext)

    if not os.path.exists(tmp):
        raise RuntimeError("Office 转换未生成输出文件")
    return tmp


def extract_texts_from_docx(doc):

    paragraphs_to_translate = []
    unique_texts = []
    text_to_id = {}
    total_unique = 0

    def add_paragraph(para):

        nonlocal total_unique

        text = para.text.strip()

        if not text: return
        if text.startswith("="): return

        paragraphs_to_translate.append((para,text))

        if text not in text_to_id:
            total_unique += 1
            if text not in cache:
                idx=len(unique_texts)
                text_to_id[text]=idx
                unique_texts.append(text)

    for para in doc.paragraphs:
        add_paragraph(para)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    add_paragraph(para)

    # 文本框（含文本框套文本框、嵌套表格，深度不限）：
    # 收集所有 w:txbxContent，段落归属"最近的"txbxContent（嵌套时不重复处理）
    from docx.oxml.ns import qn as _qn
    from docx.text.paragraph import Paragraph as _DxParagraph
    try:
        for txbx in doc.element.body.iter(_qn('w:txbxContent')):
            for p_elm in txbx.iter(_qn('w:p')):
                anc = p_elm.getparent()
                nearest = None
                while anc is not None:
                    if anc.tag == _qn('w:txbxContent'):
                        nearest = anc
                        break
                    anc = anc.getparent()
                if nearest is not txbx:
                    continue
                try:
                    add_paragraph(_DxParagraph(p_elm, doc))
                except Exception:
                    continue
    except Exception:
        pass

    log(f"  Word文本总数：{total_unique}, Cache命中：{total_unique - len(unique_texts)}, 需要翻译：{len(unique_texts)}")

    return (
        paragraphs_to_translate,
        unique_texts,
        text_to_id
    )

def apply_translations_to_docx(paragraphs_to_translate, translation_map):
    for para, orig_text in paragraphs_to_translate:
        # 先查本次翻译结果；未命中则查全局缓存（extract 阶段跳过的"已缓存文本"也要写回）
        if orig_text in translation_map:
            translated = translation_map[orig_text]
        else:
            translated = cache.get(orig_text, orig_text)
        if translated == orig_text:
            continue
        if para.runs:
            para.runs[0].text = translated
            for run in para.runs[1:]:
                run.text = ""
        else:
            para.text = translated

# ==========================
# Pptx 处理
# ==========================
def extract_texts_from_pptx(prs):
    paragraphs_to_translate = []
    unique_texts = []
    text_to_id = {}
    total_unique = 0

    def iter_shapes(shapes):
        for shape in shapes:
            if shape.shape_type == 6:
                yield from iter_shapes(shape.shapes)
            else:
                yield shape

    for slide in prs.slides:
        for shape in iter_shapes(slide.shapes):

            if shape.shape_type == 19:

                for row in shape.table.rows:
                    for cell in row.cells:
                        for para in cell.text_frame.paragraphs:

                            text = para.text.strip()

                            if text and not text.startswith('='):

                                paragraphs_to_translate.append((para, text))

                                if text not in text_to_id:
                                    total_unique += 1
                                    if text not in cache:
                                        idx = len(unique_texts)
                                        text_to_id[text] = idx
                                        unique_texts.append(text)

            elif shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if text and not text.startswith('='):
                        paragraphs_to_translate.append((para, text))
                        if text not in text_to_id:
                            total_unique += 1
                            if text not in cache:
                                idx = len(unique_texts)
                                text_to_id[text] = idx
                                unique_texts.append(text)

    log(f"  PPT文本总数：{total_unique}, Cache命中：{total_unique - len(unique_texts)}, 需要翻译：{len(unique_texts)}")
    return paragraphs_to_translate, unique_texts, text_to_id

def apply_translations_to_pptx(paragraphs_to_translate, translation_map):

    def replace_paragraph_text_preserve_style(paragraph, text):

        runs = paragraph.runs

        if not runs:
            paragraph.text = text
            return

        runs[0].text = text

        for run in runs[1:]:
            run.text = ""

    for para, orig_text in paragraphs_to_translate:

        # 先查本次翻译结果；未命中则查全局缓存（已缓存文本也要写回）
        if orig_text in translation_map:
            translated = translation_map[orig_text]
        else:
            translated = cache.get(orig_text, orig_text)

        if translated == orig_text:
            continue

        replace_paragraph_text_preserve_style(para,translated)

# ==========================
# PDF处理（v13：单元格级合并 —— 同一单元格的自动换行行合并为完整文本翻译）
# ==========================

def _is_sentence_end(text):
    """判断行文本是否以句末标点结尾（用于识别单元格的最后一行）"""
    return text.endswith(('。', '！', '？', '；', '…', '”', '"'))


def extract_texts_from_pdf(pdf):
    """
    提取PDF文本：
    1. 以block/line为单位合并span为"行"
    2. 将同一单元格内垂直相邻的自动换行行合并为"组"（完整文本送LLM）
    返回 (groups, unique_texts, text_to_id)
    """
    lines_to_translate = []
    unique_texts = []
    text_to_id = {}

    for page_index, page in enumerate(pdf):
        data = page.get_text("dict")
        block_order = 0

        for block in data["blocks"]:
            if "lines" not in block:
                continue

            for line in block["lines"]:
                span_texts = []
                span_bboxes = []
                max_fontsize = 0
                max_flags = 0
                font = ""
                origin = None
                text_color = None

                for span in line["spans"]:
                    text = span["text"]
                    stripped = text.strip()
                    if not stripped:
                        continue
                    span_texts.append(text)
                    span_bboxes.append(fitz.Rect(span["bbox"]))
                    max_fontsize = max(max_fontsize, span["size"])
                    max_flags |= span.get("flags", 0)
                    if not font:
                        font = span.get("font", "")
                    if origin is None and "origin" in span:
                        origin = fitz.Point(span["origin"])
                    c = span.get("color")
                    if c is not None and text_color is None:
                        text_color = c

                # 如果span没有origin字段，从bbox估算基线位置
                if origin is None and span_bboxes:
                    first_bbox = span_bboxes[0]
                    origin = fitz.Point(first_bbox.x0, first_bbox.y1 - 2)

                if not span_texts:
                    continue

                full_text = span_texts[0]
                for i in range(1, len(span_texts)):
                    gap = span_bboxes[i].x0 - span_bboxes[i-1].x1
                    if gap > max_fontsize * 0.3:
                        full_text += " "
                    full_text += span_texts[i]

                full_text = full_text.strip()
                if not full_text or full_text.startswith("="):
                    continue

                min_x = min(b.x0 for b in span_bboxes)
                min_y = min(b.y0 for b in span_bboxes)
                max_x = max(b.x1 for b in span_bboxes)
                max_y = max(b.y1 for b in span_bboxes)

                lines_to_translate.append({
                    "page": page_index,
                    "rect": fitz.Rect(min_x, min_y, max_x, max_y),
                    "text": full_text,
                    "fontsize": max_fontsize if max_fontsize > 0 else 10,
                    "font": font,
                    "origin": origin,
                    "flags": max_flags,
                    "color": text_color,
                    "block_order": block_order,
                })

            block_order += 1

    # 将行合并为"单元格组"（优先用表格检测，失败时用启发式）
    groups = _group_lines_into_cells(lines_to_translate, pdf)

    # 照片区拆组：含"行与图片相交"的组（照片上的标注文字，常被find_tables合并成大组）
    # 逐行拆成独立组——每行独立翻译/写回原位，避免"组译文行数<组行数"导致后半行原文残留
    # 样式分块拆组：组内行样式突变（字号/颜色/粗体不同——如 小标题+正文 混在一格）
    # 按连续样式段拆开，标题与正文各自成组（更精确翻译+样式继承）
    page_images = {}
    for p in pdf:
        img_rects = []
        for im in p.get_images(full=True):
            img_rects.extend(p.get_image_rects(im[0]))
        page_images[p.number] = img_rects

    def _row_on_image(pg, r):
        for img in page_images.get(pg, []):
            if (r.x0 < img.x1 and r.x1 > img.x0 and
                    r.y0 < img.y1 and r.y1 > img.y0):
                return True
        return False

    def _row_style(row):
        """行的样式签名：字号/颜色/粗体"""
        fs = row.get("fontsize") or 0
        c = row.get("color")
        bold = (row.get("flags", 0) or 0) & 16
        return (round(fs, 1), c, bold)

    def _make_ng(g, row):
        """按行复制组（行级样式已继承）"""
        ng = dict(g)
        ng["rows"] = [row]
        ng["rect"] = row["rect"]
        ng["text"] = row["text"]
        ng["fontsize"] = row.get("fontsize", g.get("fontsize", 10))
        ng["origin"] = row.get("origin")
        ng["color"] = row.get("color", g.get("color"))
        ng["flags"] = row.get("flags", g.get("flags", 0))
        return ng

    def make_style_ng(g, block):
        """按连续样式块复制组：组rect=块合并区、样式=块首行样式"""
        if not block:
            return None
        ng = dict(g)
        ng["rows"] = block
        min_x = min(r["rect"].x0 for r in block)
        min_y = min(r["rect"].y0 for r in block)
        max_x = max(r["rect"].x1 for r in block)
        max_y = max(r["rect"].y1 for r in block)
        ng["rect"] = fitz.Rect(min_x, min_y, max_x, max_y)
        ng["text"] = "\n".join(r["text"] for r in block)
        head = block[0]
        ng["fontsize"] = head.get("fontsize", g.get("fontsize", 10))
        ng["origin"] = head.get("origin")
        ng["color"] = head.get("color", g.get("color"))
        ng["flags"] = head.get("flags", g.get("flags", 0))
        return ng

    split_groups = []
    for g in groups:
        has_image_row = any(_row_on_image(g["page"], row["rect"]) for row in g["rows"])
        if not has_image_row:
            # 样式分块：相邻行样式签名不同则切开（标题与正文分离）
            blocks = []
            cur = []
            for row in g["rows"]:
                if cur and _row_style(row) != _row_style(cur[-1]):
                    blocks.append(cur)
                    cur = []
                cur.append(row)
            if cur:
                blocks.append(cur)
            if len(blocks) <= 1:
                split_groups.append(g)
            else:
                for block in blocks:
                    ng = make_style_ng(g, block)
                    if ng is not None:
                        split_groups.append(ng)
            continue
        # 逐行独立成组（继承组的字号/颜色/粗体等元数据）
        for row in g["rows"]:
            split_groups.append(_make_ng(g, row))
    groups = split_groups

    for group in groups:
        group_text = group["text"]
        if group_text not in text_to_id:
            idx = len(unique_texts)
            text_to_id[group_text] = idx
            unique_texts.append(group_text)
        else:
            idx = text_to_id[group_text]
        group["id"] = idx

    log(f"PDF发现{len(unique_texts)}条唯一文本（{len(groups)}个文本组，{len(lines_to_translate)}行）")
    return groups, unique_texts, text_to_id


def _group_lines_into_cells(lines, pdf):
    """
    将行分组为"单元格组"：
    1. 优先使用 page.find_tables() 检测表格单元格边界，行按中心点归属单元格
    2. 无表格或检测失败时，fallback 启发式（跨列行独立 + 列内gap/句号判断）
    """
    CROSS_RATIO = 0.4
    GAP_COEF = 1.5
    COL_TOL = 40
    page_width = max(p.rect.width for p in pdf) if pdf else 595.3

    pages = {}
    for it in lines:
        pages.setdefault(it["page"], []).append(it)

    groups = []
    for page_id in sorted(pages.keys()):
        page_lines = pages[page_id]
        cell_rects = []
        table_ok = False

        # 尝试表格检测
        try:
            tables = pdf[page_id].find_tables()
            if tables and tables.tables:
                for t in tables.tables:
                    for c in t.cells:
                        cell_rects.append(fitz.Rect(c))
                table_ok = len(cell_rects) > 0
        except Exception:
            table_ok = False

        if table_ok:
            # 行 → 单元格归属（按行中心点）
            for it in page_lines:
                cx = (it["rect"].x0 + it["rect"].x1) / 2
                cy = (it["rect"].y0 + it["rect"].y1) / 2
                it["cell"] = None
                it["cell_rect"] = None
                for ri, cr in enumerate(cell_rects):
                    if cr.x0 <= cx <= cr.x1 and cr.y0 <= cy <= cr.y1:
                        it["cell"] = ri
                        it["cell_rect"] = cr
                        break

            # 同一单元格的行合并；表格外的行独立
            cell_groups = {}
            for it in page_lines:
                key = it["cell"]
                if key is None:
                    cell_groups.setdefault(("free", id(it)), [it])
                else:
                    cell_groups.setdefault(("cell", key), []).append(it)
            for key, cl in cell_groups.items():
                cl.sort(key=lambda it: (it["rect"].y0, it["rect"].x0))
                if key[0] == "cell":
                    groups.append(cl)
                else:
                    # 表格外的独立行，逐行成组
                    for it in cl:
                        groups.append([it])
        else:
            # fallback：启发式分组
            normal = []
            for it in page_lines:
                if it["rect"].width > page_width * CROSS_RATIO:
                    groups.append([it])
                else:
                    normal.append(it)
            if not normal:
                continue
            x0_vals = sorted(set(it["rect"].x0 for it in normal))
            col_map = {}
            col_id = 0
            last_x0 = None
            for x0 in x0_vals:
                if last_x0 is None or x0 - last_x0 > COL_TOL:
                    col_id += 1
                col_map[x0] = col_id
                last_x0 = x0
            col_lines = {}
            for it in normal:
                cid = col_map[it["rect"].x0]
                col_lines.setdefault(cid, []).append(it)
            for cid in sorted(col_lines.keys()):
                cl = sorted(col_lines[cid], key=lambda it: (it["rect"].y0, it["rect"].x0))
                current = [cl[0]]
                for it in cl[1:]:
                    prev = current[-1]
                    gap = it["rect"].y0 - prev["rect"].y1
                    size = max(it["fontsize"], prev["fontsize"])
                    near_size = abs(it["fontsize"] - prev["fontsize"]) < 2
                    small_gap = 0 <= gap <= size * GAP_COEF
                    prev_ends_sentence = _is_sentence_end(prev["text"])
                    if near_size and small_gap and not prev_ends_sentence:
                        current.append(it)
                    else:
                        groups.append(current)
                        current = [it]
                groups.append(current)

    # 构建组对象（合并"竖排短行"：同x0、相邻y、每行<=2字的行合并为一行，
    # 避免"序号"这类竖排词被行级拆翻时拆成"序"/"号"碎片）
    result = []
    for rows in groups:
        if not rows:
            continue
        rows = _merge_short_rows(rows)
        min_x = min(r["rect"].x0 for r in rows)
        min_y = min(r["rect"].y0 for r in rows)
        max_x = max(r["rect"].x1 for r in rows)
        max_y = max(r["rect"].y1 for r in rows)
        merged_text = "".join(r["text"] for r in rows)
        merged_text = merged_text.strip()
        if not merged_text:
            continue
        first = rows[0]
        result.append({
            "page": first["page"],
            "rect": fitz.Rect(min_x, min_y, max_x, max_y),
            "text": merged_text,
            "rows": rows,
            "fontsize": max(r["fontsize"] for r in rows),
            "font": first["font"],
            "origin": first["origin"],
            "flags": max(r.get("flags", 0) for r in rows),
            "color": first["color"],
            "block_order": first["block_order"],
            # 容器边界（表格单元格整格 rect）：供排版用整格宽度；涂白仍用文字 rect
            "container": first.get("cell_rect"),
        })

    return result


def _group_lines_to_paragraphs(lines):
    """按block_order分组：同一block内、字号接近的连续行属于同一段落。"""
    paragraphs = []
    current_para = []

    for item in lines:
        if not current_para:
            current_para = [item]
            continue

        prev = current_para[-1]
        same_block = item.get("block_order") == prev.get("block_order")
        same_x0 = abs(item["rect"].x0 - prev["rect"].x0) < 8
        near_size = abs(item["fontsize"] - prev["fontsize"]) < 2
        near_vert = item["rect"].y0 - prev["rect"].y1 < max(item["fontsize"], prev["fontsize"]) * 2.5

        if same_block and same_x0 and near_size and near_vert:
            current_para.append(item)
        else:
            paragraphs.append(current_para)
            current_para = [item]

    if current_para:
        paragraphs.append(current_para)

    return paragraphs


def _wrap_text(text, fontname, fontsize, max_width):
    """按最大宽度将文本拆成多行（逐字符拆分，适用于CJK）"""
    lines = []
    current = ""
    for ch in text:
        test = current + ch
        if fitz.get_text_length(test, fontname=fontname, fontsize=fontsize) <= max_width or not current:
            current = test
        else:
            lines.append(current)
            current = ch
    if current:
        lines.append(current)
    return lines


def _cover_rect_pad(r):
    """统一涂白矩形参数：横扩0.5pt、纵内缩0.25pt——
    贴近文字笔画、不碰表格线（PDF版本 draw_rect 与 背景版本 像素替换共用同一套）。"""
    return (r.x0 - 0.5, r.y0 + 0.25, r.x1 + 0.5, r.y1 - 0.25)


def _merge_short_rows(rows):
    """合并"竖排短行"：同x0（差<3pt）、相邻y（间距<2.5倍字号）、每行<=2字的行合并为一行，
    避免"序号"这类竖排词被拆成"序"/"号"片段（词碎片导致翻译抖动+段映射错位）。"""
    merged = []
    i = 0
    while i < len(rows):
        r = rows[i]
        if len(r["text"].strip()) <= 2 and i + 1 < len(rows):
            nxt = rows[i + 1]
            if (len(nxt["text"].strip()) <= 2
                    and abs(nxt["rect"].x0 - r["rect"].x0) < 3
                    and 0 < nxt["rect"].y0 - r["rect"].y1 < max(r["fontsize"] or 10, nxt["fontsize"] or 10) * 2.5):
                nr = dict(r)
                nr["text"] = r["text"].strip() + nxt["text"].strip()
                nr["rect"] = fitz.Rect(r["rect"].x0, r["rect"].y0, max(r["rect"].x1, nxt["rect"].x1), nxt["rect"].y1)
                merged.append(nr)
                i += 2
                continue
        merged.append(r)
        i += 1
    return merged


def _build_col_map(groups, page_width, cross_ratio=0.4):
    """列聚类：按组 rect.x0 聚类（15pt 阈值），返回 (col_map, col_max_right)。
    阈值必须够小以区分相邻列（序号列/证据名称列），否则 box 错乱写越界。"""
    normal = [g for g in groups if g["rect"].width <= page_width * cross_ratio]
    x0_vals = sorted(set(g["rect"].x0 for g in normal))
    col_map = {}
    col_id = 0
    last_x0 = None
    for x0 in x0_vals:
        if last_x0 is None or x0 - last_x0 > 15:
            col_id += 1
        col_map[x0] = col_id
        last_x0 = x0
    col_max_right = {}
    for g in normal:
        cid = col_map[g["rect"].x0]
        col_max_right[cid] = max(col_max_right.get(cid, 0), g["rect"].x1)
    return col_map, col_max_right


def _plan_group_layout(group, translated, page_width, col_map, col_max_right, fontname="china-s", cross_ratio=0.4, right_limit=None):
    """计算一组翻译文本的布局计划：字号、拆行、每行写入坐标(x,y)与线宽。
    - x：按对应原文行 rect.x0 左对齐；超列盒右界则右对齐回截
    - y：每行用自己原文行的 origin.y（基线），合并大单元格(行间相距远)也不粘连；
      origin 缺失时用 rect.y0+fontsize 估算
    - 跳图：仅对"原文行不与图片相交"的行生效（照片上的标注行译文照写原位）
    返回 dict：{fontsize, lines, line_xy, line_w, box_x0, box_x1, rect, rows}
    """
    rect = group["rect"]
    fontsize = group["fontsize"]

    # box范围：普通组从组起点延伸到列右边界（留2pt边距），跨列行用自身范围
    if rect.width > page_width * cross_ratio:
        box_x0 = rect.x0
        box_x1 = rect.x1
    else:
        cid = col_map.get(rect.x0, 1)
        box_x0 = rect.x0
        # box 右界：
        # - 容器边界可用时用整格宽（表格单元格）
        # - 否则：若同 y 带右侧紧邻有其它组（right_limit）→ 严格止于其左侧 2pt（防重叠，如 TBM 案例）
        #         若右侧空旷 → 放宽到列右界（标题/独立行可充分利用空间，不再被压窄）
        container = group.get("container")
        if container is not None and container.width > rect.width:
            box_x1 = min(col_max_right.get(cid, rect.x1) - 2, page_width, container.x1 - 1)
        elif right_limit is not None:
            box_x1 = min(col_max_right.get(cid, rect.x1) - 2, page_width, right_limit - 2)
        else:
            # 右侧空旷：放宽到列右界（至少保留本格右缘+8pt 的余量）
            box_x1 = min(max(col_max_right.get(cid, rect.x1), rect.x1 + 8) - 2, page_width)
        if box_x1 <= box_x0:
            box_x1 = rect.x1
    box_w = box_x1 - box_x0
    if box_w <= 0:
        box_w = rect.width

    # 可用高度：从"首行基线 base_y" 到 max_y，留2pt边距（预算与真实行链一致：
    # 行链=首个基线 + (n-1)*fs*1.25 行距 + 最后一行字高 fs*1.3）
    first = group["rows"][0] if group["rows"] else None
    o0 = first.get("origin") if first else None
    if o0 is not None and hasattr(o0, "y"):
        base_y = o0.y
    else:
        base_y = rect.y0 + fontsize
    max_y0 = group.get("max_y", rect.y1 + fontsize * 3)
    max_h = max_y0 - base_y - 2
    if max_h < rect.height:
        max_h = rect.height

    # 字号（单行优先 ≥7pt；否则多行）
    fs_single = None
    fs = fontsize
    while fs >= 5:
        tw = fitz.get_text_length(translated, fontname=fontname, fontsize=fs)
        if tw <= box_w and fs * 1.3 <= max_h:
            fs_single = fs
            break
        fs -= 0.5

    fs_multi = None
    lines_multi = []
    fs = fontsize
    while fs >= 5:
        lines_multi_t = _wrap_text(translated, fontname, fs, box_w)
        # 真实行链总高：(n-1)*fs*1.25（行距）+ fs*1.3（末行字高）
        n_lines = len(lines_multi_t)
        if (n_lines - 1) * fs * 1.25 + fs * 1.3 <= max_h:
            fs_multi = fs
            lines_multi = lines_multi_t
            break
        fs -= 0.5

    if fs_single is not None and fs_single >= 7:
        fontsize = fs_single
        lines = [translated]
    elif fs_multi is not None:
        fontsize = fs_multi
        lines = lines_multi
    else:
        fontsize = 5
        tw = fitz.get_text_length(translated, fontname=fontname, fontsize=fontsize)
        if tw <= box_w:
            lines = [translated]
        else:
            lines = _wrap_text(translated, fontname, fontsize, box_w)

    rows = group["rows"]
    line_xy = []
    line_w = []
    line_colors = []
    line_bold = []
    n = len(rows)
    # 连续行距模式（2.4 布局）：首行用原文基线（或组顶估算），
    # 之后逐行 fs*1.25 均匀步进——避免"吸附碎片行 origin/x0"导致行距跳格、
    # 后半句偏右、超行溢出到下一格（照片区已按行拆组为单行组，不受影响）
    first = rows[0]
    base_x = first["rect"].x0
    for i, line in enumerate(lines):
        wid = fitz.get_text_length(line, fontname=fontname, fontsize=fontsize)
        x = base_x
        if x + wid > box_x1:
            x = max(box_x1 - wid, base_x)
        y = base_y + i * fontsize * 1.25
        line_xy.append((x, y))
        line_w.append(wid)
        # 逐行样式：优先本行 span 的颜色/粗体，无则用组级
        c = first.get("color")
        if c is None:
            c = group.get("color")
        line_colors.append(c)
        fl = first.get("flags", 0)
        if fl is None:
            fl = group.get("flags", 0) or 0
        line_bold.append((fl & 16) != 0)

    return {"fontsize": fontsize, "lines": lines, "line_xy": line_xy,
            "line_w": line_w, "box_x0": box_x0, "box_x1": box_x1,
            "line_colors": line_colors, "line_bold": line_bold, "rows": rows}


def apply_translations_to_pdf(pdf, groups, translation_map, scanned=False, scanned_pages=None, skip_paint=False):
    """
    写回PDF（按"单元格组"处理，两阶段）：
    - 阶段1：全部涂白
    - 阶段2：每组翻译文本手动拆行后用 insert_text 写入（自动换行模拟原PDF）
    - 字号：渐进缩小直到行数×行高不超过组高+余量
    - scanned=True / scanned_pages=set：扫描底页（整页图）不当作"照片"，照常涂白写字
    - skip_paint=True：跳过涂白（配合 redaction 预处理：原文已被真删除，底色由原矢量保留）
    """

    def _is_scanned_page(pg):
        if scanned_pages is not None:
            return pg in scanned_pages
        return scanned

    fontname = "china-s"
    page_width = max(p.rect.width for p in pdf) if pdf else 595.3
    CROSS_RATIO = 0.4

    # 列聚类（排除跨列行）用于限制box宽度。
    # 注意：阈值必须够小（15pt）以区分"序号列"和"证据名称列"等相邻列，
    # 否则列被合并后 box 起点/终点会错乱，导致文字越界到邻列
    col_map, col_max_right = _build_col_map(groups, page_width, CROSS_RATIO)

    # 每页图片矩形（用于涂白和写字时避开图片，防止图片被遮挡；
    # 扫描版时整页图不当作"照片"，行标记为可涂白）
    page_fullimage = {}
    page_images = {}
    for p in pdf:
        area = p.rect.width * p.rect.height
        full = False
        imgs = []
        for im in p.get_images(full=True):
            for r in p.get_image_rects(im[0]):
                imgs.append(r)
                if area > 0 and (r.width * r.height) > area * 0.6:
                    full = True
        page_images[p.number] = imgs
        page_fullimage[p.number] = full

    import io
    from PIL import Image
    import numpy as np

    # 每页渲染2倍像素图（"底色克隆"采样：行内下缘条带众数，主导底色判定）
    page_pixs = {}
    for p in pdf:
        pm = p.get_pixmap(matrix=fitz.Matrix(2, 2), colorspace=fitz.csRGB)
        page_pixs[p.number] = pm

    def sample_bg_color(pm, rect, zoom=2.0):
        """行主体带采样：rect 内部条带（y0+1.5pt~y1-1.5pt）像素众数（排除近黑文字/线），
        避免旧"下缘条带"扫入下一格底纹（如 7年6个月/静电5个月 紫格文字贴下格黄底）；
        主体带无效时退回上缘带（y0-0.3~y0+1.5pt）；颜色量化为24档聚众数，倾向白时返回白。"""
        try:
            img = Image.frombytes("RGB", (pm.width, pm.height), pm.samples)

            def _band(y0pt, y1pt):
                x0 = max(int(rect.x0 * zoom) - 6, 0)
                x1 = min(int(rect.x1 * zoom) + 6, pm.width)
                y0 = int(y0pt * zoom)
                y1 = int(y1pt * zoom)
                if y0 < 0 or y1 > pm.height or y1 <= y0 or x1 <= x0:
                    return None
                arr = np.asarray(img.crop((x0, y0, x1, y1))).astype(np.int32)
                quant = arr // 24 * 24
                colors, counts = np.unique(quant.reshape(-1, 3), axis=0, return_counts=True)
                order = counts.argsort()[::-1]
                for idx in order:
                    r_, g_, b_ = colors[idx]
                    if r_ + g_ + b_ < 150:
                        continue
                    top = colors[idx]
                    n_top = counts[idx]
                    if n_top > counts.sum() * 0.25:
                        mask = (quant == top).all(axis=2)
                        avg = arr[mask].mean(axis=0)
                        return tuple(float(v) / 255.0 for v in avg)
                return None

            main = _band(rect.y0 + 1.5, rect.y1 - 1.5)
            if main is not None:
                return main
            fallback = _band(rect.y0 - 0.3, rect.y0 + 1.5)
            if fallback is not None:
                return fallback
            return (1, 1, 1)
        except Exception:
            return (1, 1, 1)

    # ========== 计算每组可用高度（max_y）——供布局计划用 ==========
    # - 普通组：同页同列下一组的顶部
    # - 跨列行组：同页任意列的下一个组的顶部（跨列行占满宽度，下方任何组都算相邻）
    col_groups = {}
    free_groups = []
    for group in groups:
        if group["rect"].width > page_width * CROSS_RATIO:
            free_groups.append(group)
        else:
            cid = col_map.get(group["rect"].x0, 1)
            col_groups.setdefault((group["page"], cid), []).append(group)

    for key, gl in col_groups.items():
        gl.sort(key=lambda g: g["rect"].y0)
        page_h = pdf[gl[0]["page"]].rect.height if pdf else 842
        for i, g in enumerate(gl):
            if i + 1 < len(gl):
                g["max_y"] = gl[i + 1]["rect"].y0  # 同列下一组顶部
            else:
                g["max_y"] = page_h - 10

    for page_id in sorted(set(g["page"] for g in free_groups)):
        fg = [g for g in free_groups if g["page"] == page_id]
        fg.sort(key=lambda g: g["rect"].y0)
        page_h = pdf[page_id].rect.height if pdf else 842
        all_y = sorted(g["rect"].y0 for g in groups if g["page"] == page_id)
        for i, g in enumerate(fg):
            nxt = [y for y in all_y if y > g["rect"].y0 + 1]
            g["max_y"] = min(nxt) if nxt else page_h - 10

    # ========== 阶段1：全部涂白 ==========
    # 思路：先计算每组布局计划（字号/拆行/每行写入坐标），按"行原文区±1pt ∪ 行译文区"
    # 精确涂白——不再整列盒子全宽（避免盖邻格/照片上的超长色块）。
    # - 与图片相交的行（照片上的标注）：白色涂白（标注框通常白底）
    # - 其余行：底色 = 行内下缘条带像素众数（紫/粉/黄/白自适配，白格内嵌紫带也能取白）
    # 注：页面带 /Rotate（如扫描件 rotation=270）时，所有提取/渲染坐标均为"视觉坐标"，
    # 而 draw_rect/insert_text 用"未旋转页面坐标"——统一经 derotation_matrix 变换。

    def _vis_pt(page, x, y):
        if page.rotation:
            return fitz.Point(x, y) * page.derotation_matrix
        return fitz.Point(x, y)

    def _vis_rect(page, rect):
        if page.rotation:
            return rect * page.derotation_matrix
        return rect

    # 每组的"右邻组 x0"（同页、y 带重叠、x0 在右侧最近者）：
    # 有右邻 → box 严格止于其左侧（防重叠）；无右邻 → 可放宽到列右界（标题/独立行）
    right_neighbor = {}
    by_page = {}
    for g in groups:
        by_page.setdefault(g["page"], []).append(g)
    for _pg, gs in by_page.items():
        gs_sorted = sorted(gs, key=lambda x: x["rect"].x0)
        for i, g in enumerate(gs_sorted):
            r = g["rect"]
            nb = None
            for j in range(i + 1, len(gs_sorted)):
                r2 = gs_sorted[j]["rect"]
                if r2.y0 < r.y1 and r2.y1 > r.y0:
                    nb = r2.x0
                    break
            right_neighbor[id(g)] = nb

    plans = {}
    for gi, group in enumerate(groups):
        translated = translation_map.get(group["text"], group["text"])
        if translated == group["text"]:
            continue

        pg = group["page"]
        page = pdf[pg]
        pm = page_pixs.get(pg)
        plan = _plan_group_layout(group, translated, page_width, col_map, col_max_right, fontname,
                                  right_limit=right_neighbor.get(id(group)))
        plans[gi] = plan

        fs = plan["fontsize"]
        lines = plan["lines"]
        rows = group["rows"]

        if skip_paint:
            continue  # redaction 预处理版：原文已被真删除，无需涂白

        # 涂白所有原文本行（原文区±1pt 全覆盖——即使该行没有对应译文行，
        # 避免中文译文比原文短时露出后半原文（日文残留））
        for row in rows:
            r = row["rect"]
            on_image = False
            for img in page_images.get(pg, []):
                if (r.x0 < img.x1 and r.x1 > img.x0 and
                        r.y0 < img.y1 and r.y1 > img.y0):
                    on_image = True
                    break
            if _is_scanned_page(pg) and page_fullimage.get(pg):
                on_image = False  # 整页图为扫描底（非照片），允许涂白
            bg = (1, 1, 1)
            if not on_image and pm is not None:
                bg = sample_bg_color(pm, r)
            cx0, cy0, cx1, cy1 = _cover_rect_pad(r)
            page.draw_rect(_vis_rect(page, fitz.Rect(cx0, cy0, cx1, cy1)),
                           color=bg, fill=bg, overlay=True)

        # 再涂白每条译文行实际矩形（尺寸由译文文本决定，确保译文底色与其区域一致）
        for i, line in enumerate(lines):
            row_ref = rows[min(i, len(rows) - 1)]
            r = row_ref["rect"]
            on_image = False
            for img in page_images.get(pg, []):
                if (r.x0 < img.x1 and r.x1 > img.x0 and
                        r.y0 < img.y1 and r.y1 > img.y0):
                    on_image = True
                    break
            if _is_scanned_page(pg) and page_fullimage.get(pg):
                on_image = False  # 扫描底整页图，译文带可涂白
            bg = (1, 1, 1)
            if not on_image and pm is not None:
                bg = sample_bg_color(pm, r)
            x, y = plan["line_xy"][i]
            wid = plan["line_w"][i]
            # 译文区竖直带与原文区统一（横扩按译文宽度）；基线在带内，
            # 避免译文区向下越入表格线
            cx0, cy0, cx1, cy1 = _cover_rect_pad(r)
            page.draw_rect(_vis_rect(page, fitz.Rect(x - 1.2, cy0, x + wid + 1.2, cy1)),
                           color=bg, fill=bg, overlay=True)

    # ========== 阶段2：按组写翻译文本（用阶段1的计划，坐标预计算好了） ==========
    small_count = 0          # 译文被缩到 7pt 以下的组数（结束时汇总提醒）
    min_small_fs = 99.0
    for gi, group in enumerate(groups):
        translated = translation_map.get(group["text"], group["text"])
        if translated == group["text"]:
            continue

        page = pdf[group["page"]]
        plan = plans[gi]
        fontsize = plan["fontsize"]
        lines = plan["lines"]
        if fontsize < 7:
            small_count += 1
            min_small_fs = min(min_small_fs, fontsize)

        for i, line in enumerate(lines):
            # 逐行样式继承：颜色/粗体取该行 span（无则组级）
            c = plan["line_colors"][i]
            if c is not None:
                text_color = ((c >> 16) & 0xFF) / 255.0, ((c >> 8) & 0xFF) / 255.0, (c & 0xFF) / 255.0
            else:
                text_color = (0, 0, 0)
            is_bold = plan["line_bold"][i]

            x, y = plan["line_xy"][i]
            page.insert_text(
                _vis_pt(page, x, y), line,
                fontsize=fontsize,
                fontname=fontname,
                color=text_color,
                overlay=True,
                rotate=page.rotation,
            )
            if is_bold:
                page.insert_text(
                    _vis_pt(page, x + 0.4, y), line,
                    fontsize=fontsize,
                    fontname=fontname,
                    color=text_color,
                    overlay=True,
                    rotate=page.rotation,
                )

    if small_count:
        log(f"  ⚠️ {small_count} 处译文因空间不足缩小到 7pt 以下（最小 {min_small_fs:.1f}pt），"
            f"建议人工检查显示效果（可加大容器或精简译文改善）")

def is_text_pdf(pdf):
    total_text = 0
    for page in pdf:
        text = page.get_text().strip()
        total_text += len(text)
    return total_text > 20


# ---------- 扫描版 PDF：OCR 提取 ----------
_OCR_ENGINE = None


def _get_ocr_engine():
    """懒加载 RapidOCR（日文模型，modelscope 源，已缓存到 .venv 包内）"""
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        from rapidocr import RapidOCR, LangRec
        log("  [OCR] 加载 RapidOCR 引擎（首次加载稍慢）...")
        _OCR_ENGINE = RapidOCR(params={
            "Det.lang_type": LangRec.JAPAN,
            "Rec.lang_type": LangRec.JAPAN,
        })
    return _OCR_ENGINE


def extract_texts_from_scanned_pdf(pdf, dpi=300, pages=None):
    """扫描版 PDF 提取：每页渲染图像 → RapidOCR → 行级框转组。
    pages：只处理指定页号集合（None=全部页；供混合 PDF 逐页分流用）。
    返回 (groups, unique_texts, text_to_id)，结构与 extract_texts_from_pdf 一致，
    写回/翻译流程可直接复用。"""
    from PIL import Image
    import numpy as np

    groups = []
    unique_texts = []
    text_to_id = {}
    scale = dpi / 72.0
    engine = _get_ocr_engine()

    for pi, page in enumerate(pdf):
        if pages is not None and pi not in pages:
            continue
        pm = page.get_pixmap(dpi=dpi, colorspace=fitz.csRGB)
        img = Image.frombytes("RGB", (pm.width, pm.height), pm.samples)
        arr = np.asarray(img)
        out = engine(arr)
        rows = []
        for bx, txt, sc in zip(out.boxes, out.txts, out.scores):
            try:
                if not txt or float(sc) < 0.45:
                    continue
                x0 = min(p[0] for p in bx)
                y0 = min(p[1] for p in bx)
                x1 = max(p[0] for p in bx)
                y1 = max(p[1] for p in bx)
            except Exception:
                continue
            r = fitz.Rect(x0 / scale, y0 / scale, x1 / scale, y1 / scale)
            fs = max(r.height * 0.78, 6)
            rows.append({"text": txt.strip(), "rect": r, "fontsize": fs,
                         "color": 0, "flags": 0})
        # 同一 y 带内、x 紧邻的框合并为一行（避免半字/缺字框拆成多个组）
        rows.sort(key=lambda r: (round(r["rect"].y1, 1), r["rect"].x0))
        merged = []
        for row in rows:
            if merged:
                last = merged[-1]
                dy = abs((row["rect"].y0 + row["rect"].y1) / 2 - (last["rect"].y0 + last["rect"].y1) / 2)
                gap = row["rect"].x0 - last["rect"].x1
                if dy < max(row["rect"].height, last["rect"].height) * 0.6 and -4 < gap < last["fontsize"] * 2.2:
                    nb = fitz.Rect(min(last["rect"].x0, row["rect"].x0),
                                   min(last["rect"].y0, row["rect"].y0),
                                   max(last["rect"].x1, row["rect"].x1),
                                   max(last["rect"].y1, row["rect"].y1))
                    last["text"] = last["text"] + row["text"]
                    last["rect"] = nb
                    last["fontsize"] = max(last["fontsize"], row["fontsize"])
                    continue
            merged.append(dict(row))
        # 垂直合并成"段落组"：相邻行（同列缩进、行距近、上一行不以句末标点结尾）合并为一组。
        # 与 native 版的分组行为对齐——否则每行一组时，译文（多行）没有下探空间，
        # 被迫兜底放行导致多组译文叠在一起。
        paragraphs = []
        cur = []
        for row in merged:
            if cur:
                last = cur[-1]
                dy = row["rect"].y0 - last["rect"].y1
                fs_ref = max(last["fontsize"], row["fontsize"])
                x_ok = abs(row["rect"].x0 - last["rect"].x0) < fs_ref * 2.2
                y_ok = -2.5 < dy < fs_ref * 1.0
                prev_end = last["text"].rstrip()
                cont = not prev_end.endswith(("。", "．", ".", "！", "？", "!", "?"))
                if x_ok and y_ok and cont:
                    cur.append(row)
                    continue
            if cur:
                paragraphs.append(cur)
            cur = [row]
        if cur:
            paragraphs.append(cur)

        for para in paragraphs:
            para_text = "".join(r["text"] for r in para).strip()
            if not para_text:
                continue
            p_rect = fitz.Rect(
                min(r["rect"].x0 for r in para),
                min(r["rect"].y0 for r in para),
                max(r["rect"].x1 for r in para),
                max(r["rect"].y1 for r in para),
            )
            if para_text in text_to_id:
                idx = text_to_id[para_text]
            else:
                idx = len(unique_texts)
                text_to_id[para_text] = idx
                unique_texts.append(para_text)
            groups.append({
                "id": idx, "text": para_text, "rect": p_rect,
                "page": pi, "cell": None, "rows": para,
                "fontsize": max(r["fontsize"] for r in para), "origin": None,
                "color": 0, "flags": 0,
            })
    return groups, unique_texts, text_to_id


def _cluster_values(values, tolerance):
    """将数值聚类，返回每个值的簇ID"""
    sorted_vals = sorted(set(values))
    clusters = {}
    cluster_id = 0
    last_val = None
    for v in sorted_vals:
        if last_val is None or v - last_val > tolerance:
            cluster_id += 1
        clusters[v] = cluster_id
        last_val = v
    return clusters


def _detect_column_boundaries(x0_list, gap_threshold=25):
    """按列间隙检测列边界：x0值的大跳跃（>gap_threshold）视为列分界。
    返回 [(col_min, col_max), ...] 每列的x范围。"""
    sorted_x0 = sorted(set(x0_list))
    columns = []
    col_start = sorted_x0[0]
    col_end = sorted_x0[0]

    for x0 in sorted_x0[1:]:
        if x0 - col_end > gap_threshold:
            columns.append((col_start, col_end))
            col_start = x0
        col_end = x0
    columns.append((col_start, col_end))
    return columns


def _assign_to_column(x0, columns):
    """将x0分配到最近的列"""
    for ci, (col_min, col_max) in enumerate(columns):
        if x0 >= col_min - 5 and x0 <= col_max + 5:
            return ci
    # 找最近的列
    best_ci = 0
    best_dist = float("inf")
    for ci, (col_min, col_max) in enumerate(columns):
        mid = (col_min + col_max) / 2
        dist = abs(x0 - mid)
        if dist < best_dist:
            best_dist = dist
            best_ci = ci
    return best_ci


def _clean_fake_borders(doc):
    """清洗 pdf2docx 推断出的"假竖向缝"：
    某列缝在一行被 gridSpan 合并（说明原文件该处连续、无竖线），
    却在另一行被拆成两个独立 cell（pdf2docx 按 span 间隙误切）→
    判定为假缝，移除该缝两侧 cell 的 start/end 边框；真实缝（所有行一致）不动。
    """
    from docx.oxml.ns import qn
    from lxml import etree

    def _nil_border(tc, tag):
        tcPr = tc.find(qn('w:tcPr'))
        borders = tcPr.find(qn('w:tcBorders')) if tcPr is not None else None
        if borders is None:
            return
        b = borders.find(qn(tag))
        if b is not None:
            b.set(qn('w:val'), 'nil')

    def walk_table(tbl):
        rows_info = []
        for tr in tbl.iter(qn('w:tr')):
            info = []
            col = 0
            for tc in tr.findall(qn('w:tc')):
                span = 1
                tcPr = tc.find(qn('w:tcPr'))
                if tcPr is not None:
                    gs = tcPr.find(qn('w:gridSpan'))
                    if gs is not None:
                        try:
                            span = int(gs.get(qn('w:val')) or 1)
                        except Exception:
                            span = 1
                info.append((tc, col, col + span))
                col += span
            if info:
                rows_info.append(info)
        max_col = max((r[2] for row in rows_info for r in row), default=0)
        for k in range(max_col - 1):
            has_merge = False
            has_split = False
            for info in rows_info:
                for tc, s, e in info:
                    if e - s > 1 and s <= k < e:
                        has_merge = True
                if any(a[2] == k + 1 and b[1] == k + 1 for a, b in zip(info, info[1:])):
                    has_split = True
            if has_merge and has_split:
                for info in rows_info:
                    for tc, s, e in info:
                        if s == k + 1:
                            _nil_border(tc, 'w:start')
                        if e == k + 1:
                            _nil_border(tc, 'w:end')
        # 嵌套表递归
        for tr in tbl.iter(qn('w:tr')):
            for tc in tr.findall(qn('w:tc')):
                ntbl = tc.find(qn('w:tbl'))
                if ntbl is not None:
                    walk_table(ntbl)

    for tbl in doc.tables:
        walk_table(tbl._tbl)


def create_docx_pdf2docx(pdf_path, output):
    """
    Word翻译版（pdf2docx方案）：保留可编辑的表格/段落结构。
    局限：PDF中的矢量层（标注/箭头等）会被pdf2docx压扁为不透明图片，
    含标注层的PDF（如现场照片标注）可能互相遮挡；此时请用 _背景版.docx。
    处理：按单元格宽度缩小字号（参考PDF的字号算法，保持布局不变）。
    """
    import tempfile
    from pdf2docx import Converter

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = tmp.name

    t0 = time.time()
    log("  使用 pdf2docx 转换 PDF → DOCX ...")
    cv = Converter(pdf_path)
    cv.convert(tmp_path)
    cv.close()
    t1 = time.time()
    log(f"  [pdf2docx纯转换] {t1 - t0:.2f}s")

    doc = Document(tmp_path)

    positions = []
    unique_texts = []
    text_to_id = {}

    for pi, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text or text.startswith("="):
            continue
        if text not in text_to_id:
            idx = len(unique_texts)
            unique_texts.append(text)
            text_to_id[text] = idx
        positions.append(("para", pi, text))

    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                for pi, para in enumerate(cell.paragraphs):
                    text = para.text.strip()
                    if not text or text.startswith("="):
                        continue
                    if text not in text_to_id:
                        idx = len(unique_texts)
                        unique_texts.append(text)
                        text_to_id[text] = idx
                    positions.append(("cell", ti, ri, ci, pi, text))

    t2 = time.time()
    log(f"  [DOCX解析+提取] {t2 - t1:.2f}s")
    log(f"  DOCX发现{len(unique_texts)}条唯一文本（{len(positions)}个位置）")

    new_texts = [t for t in unique_texts if t not in cache]
    if new_texts:
        log(f"  DOCX有{len(new_texts)}条文本缓存未命中，需要另外翻译")
        new_translations = process_pdf_translations(new_texts)
        for text, trans in new_translations.items():
            if trans != text:
                cache[text] = trans
    t3 = time.time()
    log(f"  [DOCX补翻译] {t3 - t2:.2f}s")

    for pos in positions:
        pos_type = pos[0]
        orig = pos[-1]
        translated = cache.get(orig, orig)
        if translated == orig:
            continue
        if pos_type == "para":
            pi = pos[1]
            _replace_para_text(doc.paragraphs[pi], translated)
        elif pos_type == "cell":
            ti, ri, ci, pi = pos[1:5]
            cell = doc.tables[ti].rows[ri].cells[ci]
            para = cell.paragraphs[pi]
            _replace_para_text_fit(para, translated, cell)

    _clean_fake_borders(doc)
    doc.save(output)
    os.unlink(tmp_path)
    t4 = time.time()
    log(f"  [DOCX替换+保存] {t4 - t3:.2f}s")


def create_docx_from_pdf(pdf_path, output, groups=None):
    """
    生成Word翻译版（整页保真背景 + 逐行底色克隆浮动文本框）：
    - 每页用 PyMuPDF 渲染原页为高清PNG作为整页背景（不涂白——底纹/表格线/照片/标注全保留）
    - 每个翻译文本行用独立浮动文本框（behindDoc=0 浮于背景上），
      文本框填充色=该行底色克隆（逐行采样，紫/黄/粉/白自适配，盖住原文字且不突元）
    - 文本继承字号/颜色/粗体（逐行样式）
    背景版不再涂白：视觉与PDF版一致，同时文本可编辑（精修后转PDF无重大样式问题）。
    """
    import io
    from docx import Document
    from docx.shared import Pt, Mm, Inches
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml
    from lxml import etree

    pdf = fitz.open(pdf_path)
    page_w_pt = pdf[0].rect.width
    page_h_pt = pdf[0].rect.height

    # 提取文字组（与PDF流程相同的分组逻辑；从主流程传入 groups 时直接复用，避免重新解析）
    if groups is None:
        groups, unique_texts, text_to_id = extract_texts_from_pdf(pdf)
    col_map, col_max_right = _build_col_map(groups, page_w_pt)

    # 创建文档并设置页面尺寸 = PDF页面
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Mm(page_w_pt / 72 * 25.4)
    sec.page_height = Mm(page_h_pt / 72 * 25.4)
    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Mm(0)

    # 每页分组
    pages = {}
    for g in groups:
        pages.setdefault(g["page"], []).append(g)

    drawing_id = 100  # 从100开始，避免与python-docx自动生成的docPr id冲突
    for page_idx in range(len(pdf)):
        # 逐页尺寸：页面尺寸变化时新建 section（混合纸张 PDF）；否则沿用分页符
        cur_w_pt = pdf[page_idx].rect.width
        cur_h_pt = pdf[page_idx].rect.height
        if page_idx > 0:
            prev_w_pt = pdf[page_idx - 1].rect.width
            prev_h_pt = pdf[page_idx - 1].rect.height
            if abs(cur_w_pt - prev_w_pt) > 1 or abs(cur_h_pt - prev_h_pt) > 1:
                from docx.enum.section import WD_SECTION
                s = doc.add_section(WD_SECTION.NEW_PAGE)
                s.page_width = Mm(cur_w_pt / 72 * 25.4)
                s.page_height = Mm(cur_h_pt / 72 * 25.4)
                s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Mm(0)
            else:
                doc.add_page_break()
        # 每页一个锚定段落：本页所有浮动对象（背景图+文本框）锚定在同段落，
        # 段落流只有一行空段，避免多段排布撑出多余页面、浮动对象跑错页
        ap = doc.add_paragraph()
        ar = ap.add_run()
        # 背景图：像素级"底色克隆涂白"（替换文本区像素为采样底色；紫/黄/粉/白自适配，
        # 不产生 draw_rect 多边形遮盖的伪影；照片上标注框涂白）
        pm2 = pdf[page_idx].get_pixmap(matrix=fitz.Matrix(2, 2), colorspace=fitz.csRGB)
        bg_png = _make_cleaned_bg(pdf[page_idx], page_idx, pages.get(page_idx, []), pm2)
        _add_bg_image(doc, bg_png, cur_w_pt, cur_h_pt, drawing_id, ar)
        drawing_id += 1
        # 该页翻译文字文本框（每组一框：框=组rect原文区域；透明填充，文字只作为译文层）
        for g in pages.get(page_idx, []):
            trans = cache.get(g["text"], g["text"])
            if not trans or trans == g["text"]:
                continue
            _add_group_textbox(doc, trans, g, cur_w_pt, cur_h_pt, pdf[page_idx], pm2, drawing_id, ar)
            drawing_id += 1

    pdf.close()
    doc.save(output)


def _add_group_textbox(doc, text, group, page_w_pt, page_h_pt, page, pm, did, run):
    """每组一个浮动文本框：框=组rect（原文文本区±1pt，遮盖等宽完整的原日文区域）。
    译文在框内换行排版、字号按框宽/框高内适配（行数*1.32*fs ≤ 框高），
    内边距0、文字顶对齐——框恰好盖住原文区：原日文不露、译文不削字。"""
    import math
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml

    rect = group["rect"]
    fontsize = group["fontsize"] or 10
    group_w = rect.width
    group_h = rect.height
    box_w = min(group_w + 2, page_w_pt - rect.x0 + 1)  # 框宽（界内）

    # 字号适配：width_limit（单行需要缩行宽）与 height_limit（行数×行高 ≤ 框高）
    def est_lines(fs):
        tw = fitz.get_text_length(text, fontname="china-s", fontsize=fs)
        return max(1, math.ceil(tw / max(box_w - 2, 1))), tw

    fs = fontsize
    while fs >= 5:
        n, tw = est_lines(fs)
        # 单行优先：能单行且行高≤框高即用
        if n == 1 and fs * 1.32 <= group_h + 2:
            break
        if fs * 1.32 <= group_h + 2 and tw <= box_w - 2:
            break
        fs -= 0.5
        if fs < 8:
            break
    # 兜底：8pt 以下继续适配高度为主
    while fs >= 5:
        n, tw = est_lines(fs)
        if n * fs * 1.32 <= group_h + 2:
            break
        fs -= 0.5
    fs = max(fs, 5)
    n_final, tw_final = est_lines(fs)

    # 框几何：左上角=组rect左上-1pt（覆盖原文区），高=组高+2（至少单行高）
    EMU = 12700
    x_emu = int((rect.x0 - 1) * EMU)
    y_emu = int((rect.y0 - 1) * EMU)
    w_emu = int(box_w * EMU)
    h_emu = int(max(group_h + 2, n_final * fs * 1.32 + 1) * EMU)

    # 底色：框=原文区，底色克隆（照片上标注（与图相交）→ 白）
    on_image = False
    try:
        for im in page.get_images(full=True):
            for ir in page.get_image_rects(im[0]):
                if (rect.x0 < ir.x1 and rect.x1 > ir.x0 and
                        rect.y0 < ir.y1 and rect.y1 > ir.y0):
                    on_image = True
                    break
    except Exception:
        on_image = False
    bg = (1, 1, 1)
    if not on_image and pm is not None:
        bg = _sample_bg_color_pm(pm, rect)
    bg_hex = "{:02X}{:02X}{:02X}".format(
        int(round(bg[0] * 255)), int(round(bg[1] * 255)), int(round(bg[2] * 255)))

    # 样式（组级颜色/粗体）
    c = group.get("color")
    color_hex = "{:06X}".format(c & 0xFFFFFF) if c is not None else "000000"
    bold = '<w:b/>' if (group.get("flags", 0) or 0) & 16 != 0 else ''

    tb_xml = (
        '<w:drawing {ns}>'
        '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        'relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page"><wp:posOffset>{x}</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="page"><wp:posOffset>{y}</wp:posOffset></wp:positionV>'
        '<wp:extent cx="{w}" cy="{h}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapNone/>'
        '<wp:docPr id="{did}" name="TextBox {did}"/>'
        '<wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/>'
        '<a:ln w="0"><a:noFill/></a:ln>'
        '</wps:spPr>'
        '<wps:txbx><w:txbxContent>'
        '<w:p><w:pPr>'
        '<w:spacing w:before="0" w:after="0" w:line="264" w:lineRule="auto"/>'
        '<w:jc w:val="left"/></w:pPr>'
        '<w:r><w:rPr><w:sz w:val="{sz}"/>'
        '<w:color w:val="{color}"/>{bold}</w:rPr>'
        '{text}'
        '</w:r>'
        '</w:p>'
        '</w:txbxContent></wps:txbx>'
        '<wps:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"/>'
        '</wps:wsp></a:graphicData></a:graphic>'
        '</wp:anchor>'
        '</w:drawing>'
    ).format(
        ns=nsdecls("wp", "a", "w") + ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
        x=x_emu, y=y_emu, w=w_emu, h=h_emu, did=did,
        sz=int(fs * 2), bold=bold, color=color_hex, bg=bg_hex,
        text=_text_br(escape(text)),
    )

    # 关键：w:drawing 必须放在 w:r 内（w:p > w:r > w:drawing）
    run._r.append(parse_xml(tb_xml))


def _text_br(escaped_text):
    """转义后的文本转 XML 片段：\n → <w:br/>（Word 文本框内换行）"""
    if "\n" not in escaped_text:
        return f'<w:t xml:space="preserve">{escaped_text}</w:t>'
    return "<w:br/>".join(
        f'<w:t xml:space="preserve">{p}</w:t>' for p in escaped_text.split("\n"))


def _sample_bg_color_pm(pm, rect, zoom=2.0):
    """与 apply_translations_to_pdf 统一的"行主体带"采样（避免循环内重复实现）——
    取 rect 内部带（y0+1.5pt ~ y1-1.5pt）非近黑像素众数，
    避免下缘条带扫入下一格底纹（如 7年6个月/静电5个月 紫格文字贴下格黄底）；
    主体带无效时退回上缘带（y0-0.3 ~ y0+1.5pt）；量化24档聚众数，倾向白时返回白。"""
    try:
        from PIL import Image
        import numpy as np
        img = Image.frombytes("RGB", (pm.width, pm.height), pm.samples)

        def _band(y0pt, y1pt):
            x0 = max(int(rect.x0 * zoom) - 6, 0)
            x1 = min(int(rect.x1 * zoom) + 6, pm.width)
            y0 = int(y0pt * zoom)
            y1 = int(y1pt * zoom)
            if y0 < 0 or y1 > pm.height or y1 <= y0 or x1 <= x0:
                return None
            arr = np.asarray(img.crop((x0, y0, x1, y1))).astype(np.int32)
            quant = arr // 24 * 24
            colors, counts = np.unique(quant.reshape(-1, 3), axis=0, return_counts=True)
            order = counts.argsort()[::-1]
            for idx in order:
                r_, g_, b_ = colors[idx]
                if r_ + g_ + b_ < 150:
                    continue
                top = colors[idx]
                n_top = counts[idx]
                if n_top > counts.sum() * 0.25:
                    mask = (quant == top).all(axis=2)
                    avg = arr[mask].mean(axis=0)
                    return tuple(float(v) / 255.0 for v in avg)
            return None

        main = _band(rect.y0 + 1.5, rect.y1 - 1.5)
        if main is not None:
            return main
        fallback = _band(rect.y0 - 0.3, rect.y0 + 1.5)
        if fallback is not None:
            return fallback
        return (1, 1, 1)
    except Exception:
        return (1, 1, 1)


def _make_cleaned_bg(pdf_page, page_idx, groups, pm2, zoom=150 / 72):
    """像素级"底色克隆涂白"的背景PNG：
    渲染原页150dpi → 对每个文本行区域，用采样底色（紫/黄/粉/白）替换其像素；
    照片上的标注行（与图片相交）→ 涂白（标注框通常白底）。
    不产生 draw_rect 多边形遮盖的伪影（跨线/跨底色），表格线不受影响。
    返回 PNG bytes。"""
    import io
    from PIL import Image
    import numpy as np

    pix = pdf_page.get_pixmap(dpi=150)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    arr = np.asarray(img).copy()

    for g in groups:
        for row in g["rows"]:
            r = row["rect"]
            # 未翻译行（译文==原文，如"直接"/"有"/"TBM-KY"）保留原文，不涂白
            if cache.get(g["text"], g["text"]) == g["text"]:
                continue
            # 判定是否在照片上
            on_image = False
            try:
                for im in pdf_page.get_images(full=True):
                    for ir in pdf_page.get_image_rects(im[0]):
                        if (r.x0 < ir.x1 and r.x1 > ir.x0 and
                                r.y0 < ir.y1 and r.y1 > ir.y0):
                            on_image = True
                            break
            except Exception:
                on_image = False
            if on_image:
                bg = (255, 255, 255)
            else:
                bg = _sample_bg_color_pm(pm2, r)
                bg = tuple(int(round(v * 255)) for v in bg)
            # 像素区域：与 PDF 版统一（横扩0.5pt、纵内缩0.25pt）
            cx0, cy0, cx1, cy1 = _cover_rect_pad(r)
            x0 = max(int(cx0 * zoom), 0)
            x1 = min(int(cx1 * zoom), arr.shape[1])
            y0 = max(int(cy0 * zoom), 0)
            y1 = min(int(cy1 * zoom), arr.shape[0])
            if x0 < x1 and y0 < y1:
                arr[y0:y1, x0:x1, :] = bg

    out_img = Image.fromarray(arr)
    buf = io.BytesIO()
    out_img.save(buf, format="PNG")
    return buf.getvalue()


def _add_bg_image(doc, pixmap, page_w_pt, page_h_pt, did, run):
    """把渲染的页面PNG作为整页浮动背景图（behindDoc=1，置于文字下方），锚定到传入run
    （与译文文本框共用同一锚定段落，避免背景/文字锚点不同导致跨页漂移）"""
    import io
    from docx.shared import Inches
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml
    from lxml import etree

    run.add_picture(io.BytesIO(pixmap),
                    width=Inches(page_w_pt / 72), height=Inches(page_h_pt / 72))
    inline = run._r.find(qn("w:drawing") + "/" + qn("wp:inline"))
    if inline is None:
        return
    graphic = inline.find(qn("a:graphic"))
    cx = inline.find(qn("wp:extent")).get("cx")
    cy = inline.find(qn("wp:extent")).get("cy")
    docPr = inline.find(qn("wp:docPr"))
    did_attr = docPr.get("id")
    dname = docPr.get("name")

    anchor_xml = (
        '<wp:anchor {ns} distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        'relativeHeight="0" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        '<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'   # ← 补上
        '<wp:wrapNone/>'
        '<wp:docPr id="{did}" name="{name}"/>'
        '<wp:cNvGraphicFramePr/>{graphic}</wp:anchor>'
    ).format(ns=nsdecls("wp", "a"), cx=cx, cy=cy, did=did_attr, name=dname,
             graphic=etree.tostring(graphic, encoding="unicode"))
    anchor = parse_xml(anchor_xml)
    inline.getparent().replace(inline, anchor)


def _add_translation_textbox(doc, text, group, page_w_pt, did):
    """把翻译文本放入绝对定位的浮动文本框（无边框无填充，继承字号/颜色/粗体）"""
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml

    rect = group["rect"]
    fontsize = group["fontsize"]
    if fontsize <= 0:
        fontsize = 10

    # 文本框宽度：组宽 + 余量（不超页面）
    box_w = min(rect.width * 2 + 20, page_w_pt - rect.x0)
    # 字号：单行优先缩字号，使翻译文本能放入框宽（参考PDF算法）
    fs = fontsize
    while fs >= 5:
        tw = fitz.get_text_length(text, fontname="china-s", fontsize=fs)
        if tw <= box_w:
            break
        fs -= 0.5
    fs = max(fs, 5)
    # 估算所需高度
    est_lines = max(1, int(fitz.get_text_length(text, fontname="china-s", fontsize=fs) / box_w) + 1)
    box_h = max(rect.height, est_lines * fs * 1.3 + 4)

    # 颜色
    color_hex = "000000"
    if group.get("color") is not None:
        c = group["color"]
        color_hex = "{:06X}".format(c & 0xFFFFFF)
    bold = '<w:b/>' if (group.get("flags", 0) & 16) != 0 else ''

    EMU = 12700
    x_emu = int(rect.x0 * EMU)
    y_emu = int(rect.y0 * EMU)
    w_emu = int(box_w * EMU)
    h_emu = int(box_h * EMU)

    tb_xml = (
        '<w:drawing {ns}>'
        '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        'relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page"><wp:posOffset>{x}</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="page"><wp:posOffset>{y}</wp:posOffset></wp:positionV>'
        '<wp:extent cx="{w}" cy="{h}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'          # ← 补上
        '<wp:wrapNone/>'
        '<wp:docPr id="{did}" name="TextBox {did}"/>'
        '<wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'  # ← off 改为 0,0
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/>'                                            # ← 补上形状填充
        '<a:ln w="0"><a:noFill/></a:ln>'
        '</wps:spPr>'
        '<wps:txbx><w:txbxContent>'
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="MS Mincho" w:eastAsia="MS Mincho"/>'
        '<w:color w:val="{color}"/><w:sz w:val="{sz}"/>{bold}</w:rPr>'
        '<w:t xml:space="preserve">{text}</w:t></w:r>'
        '</w:p>'
        '</w:txbxContent></wps:txbx>'
        '<wps:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"/>'
        '</wps:wsp></a:graphicData></a:graphic>'
        '</wp:anchor>'
        '</w:drawing>'
    ).format(
        ns=nsdecls("wp", "a", "w") + ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
        x=x_emu, y=y_emu, w=w_emu, h=h_emu, did=did,
        sz=int(fs * 2), bold=bold, color=color_hex,
        text=escape(text)
    )

    p = doc.add_paragraph()
    run = p.add_run()
    # 关键：w:drawing 必须放在 w:r 内（w:p > w:r > w:drawing），
    # 直接 append 到 w:p 是非法层级，Word 打开会报错
    run._r.append(parse_xml(tb_xml))


def _unset_fixed_row_height(doc):
    """表格固定行高(exact)改为自动(atLeast)，避免翻译文本被截断"""
    from docx.oxml.ns import qn
    for table in doc.tables:
        for tr in table._tbl.findall(qn('w:tr')):
            trPr = tr.find(qn('w:trPr'))
            if trPr is not None:
                trH = trPr.find(qn('w:trHeight'))
                if trH is not None and trH.get(qn('w:hRule')) == 'exact':
                    trH.set(qn('w:hRule'), 'atLeast')


def _replace_para_text_fit(para, new_text, cell):
    """替换段落文本，并按单元格宽度缩小字号（参考PDF的字号算法）"""
    width_pt = cell.width / 12700 if cell.width else 100
    _replace_para_text(para, new_text)

    if not para.runs:
        return
    orig_size = para.runs[0].font.size.pt if para.runs[0].font.size else 11
    if orig_size <= 0:
        orig_size = 11

    fs = orig_size
    while fs >= 5:
        tw = fitz.get_text_length(new_text, fontname="china-s", fontsize=fs)
        if tw <= width_pt * 1.2 or width_pt <= 0:
            break
        fs -= 0.5
    fs = max(fs, 5)

    for run in para.runs:
        if run.text:
            run.font.size = Pt(fs)


def _replace_para_text(para, new_text):
    """替换段落文本，保留第一个run的样式"""
    if para.runs:
        para.runs[0].text = new_text
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.text = new_text


def process_pdf_translations(unique_texts):

    log(f"\n发现 {len(unique_texts)} 条PDF文本需要翻译（缓存未命中）")

    # 保护数字范围模式（如 "2-5 页" → 先翻译"页"，翻译后拼回"2-5"）
    # 避免LLM把 "2-5" 翻译成 "25"（丢失连字符）
    range_restore = {}
    protected_texts = list(unique_texts)
    range_pat = re.compile(r'^(\d+)[-－](\d+)\s+(.+)$')

    for idx, text in enumerate(unique_texts):
        m = range_pat.match(text)
        if m:
            prefix = m.group(1) + "-" + m.group(2)
            suffix = m.group(3)
            if suffix.strip():
                range_restore[idx] = prefix
                protected_texts[idx] = suffix  # 只把"页"送给LLM翻译

    id_to_translation = {}

    # 页码/纯符号/纯数字类文本（如 "(1/2)"、"2009-1-1"）不送 LLM：
    # 模型会忽略或省略这类项，导致同批后续译文整体前移错位（已实测复现），直接原样保留
    symbol_re = re.compile(r'^[\s\d\W_]+$')

    # 两阶段：先收集批次，再统一发送（便于打印 [i/N] 进度）
    batch = []
    current_chars = 0
    batches = []
    singles = []
    for idx, text in enumerate(protected_texts):
        if not text.strip():
            continue
        if symbol_re.match(text):
            id_to_translation[idx] = text
            continue
        if len(text) > SINGLE_TEXT_THRESHOLD:
            singles.append(idx)
            continue
        if current_chars + len(text) > MAX_BATCH_CHARS:
            batches.append(batch)
            batch = []
            current_chars = 0
        batch.append((idx, text))
        current_chars += len(text)
    if batch:
        batches.append(batch)

    total = len(batches) + len(singles)
    for i, b in enumerate(batches, 1):
        log(f"  [{i}/{total}] 发送批次: {len(b)}条")
        result = qwen_translate_batch(b, TARGET_LANG)
        id_to_translation.update(result)
    for i, idx in enumerate(singles, len(batches) + 1):
        log(f"  [{i}/{total}] 长文本单独翻译 [{idx}]")
        id_to_translation[idx] = qwen_translate(protected_texts[idx], TARGET_LANG)

    # 恢复被保护的数字范围
    for idx, prefix in range_restore.items():
        if idx in id_to_translation:
            id_to_translation[idx] = prefix + " " + id_to_translation[idx]

    # 统一缓存 key：原文 → 译文（与 Excel/Word/PPT 流程一致，便于 pdf2docx 版复用）
    return {text: id_to_translation.get(idx, text) for idx, text in enumerate(unique_texts)}

# ==========================
# 主流程
# ==========================
def get_files(folder):

    supported = [".xlsx",".xlsm",".xls",".docx",".pptx",".doc",".ppt",".pdf"]
    files = []

    for filename in os.listdir(folder):

        filepath = os.path.join(folder,filename)

        if os.path.isfile(filepath):

            ext = os.path.splitext(filename)[1].lower()

            if ext in supported:
                files.append(filepath)

    return files

def translate_worker():

    global last_lang
    global istranslating

    if TARGET_LANG != last_lang:
        cache.clear()
        last_lang = TARGET_LANG

    try:
        if SOURCE_TYPE=="file":
            process_single_file(SOURCE_PATH)

        elif SOURCE_TYPE=="folder":

            files = get_files(SOURCE_PATH)
            total=len(files)
            log(f"发现{total}个文件")

            for index,file in enumerate(files):

                log(f"[{index+1}/{total}]")
                process_single_file(file)

        log("全部任务完成")

    except Exception as e: log("程序异常:"+str(e))

    finally: 
        istranslating = False
        root.after(0, lambda: translation_button.config(state=tk.NORMAL, text="开始翻译", bg="SystemButtonFace"))

def process_single_file(file_path):

    cells_to_translate = []
    unique_texts = []
    text_to_id = {}

    start_time = time.time()

    log("="*40)

    log(f"开始处理:{file_path}")

    folder = os.path.dirname(file_path)

    filename = os.path.basename(file_path)

    name, ext = os.path.splitext(filename)

    ext = ext.lower()

    try:
        # ===========================
        # Excel
        # ===========================
        if ext in [".xlsx",".xlsm",".xls"]:

            log("检测到Excel文件")
            final_output = os.path.join(folder, f"{name}_{TARGET_LANG}{ext}")
            app = xw.App(visible=False, add_book=False)
            is_repaired = False
            temp_repaired = None

            try:
                try:
                    wb = app.books.open(file_path, update_links=False)
                    log("✅ 文件正常打开，开始翻译...")
                    is_repaired = False
                except Exception:
                    log("⚠️ 正常打开失败，尝试修复...")
                    temp_repaired = os.path.join(folder, f"{name}_temp_repaired{ext}")
                    if os.path.exists(temp_repaired):
                        os.remove(temp_repaired)
                    try:
                        wb = app.api.Workbooks.Open(
                            os.path.normpath(file_path),
                            UpdateLinks=False,
                            CorruptLoad=1
                        )
                        wb.SaveAs(os.path.normpath(temp_repaired))
                        wb.Close(SaveChanges=False)
                        app.quit()
                        time.sleep(0.5)
                        app = xw.App(visible=False, add_book=False)
                        wb = app.books.open(temp_repaired, update_links=False)
                        is_repaired = True
                        log("✅ 修复成功，使用修复后的临时文件")
                    except Exception as repair_err:
                        log(f"❌ 修复过程中出错: {repair_err}")
                        raise

                translation_map = process_excel(wb)
                # 只缓存成功翻译（译文为空/与原文相同不缓存——防止失败结果被永久缓存）
                cache.update({k: v for k, v in translation_map.items() if v and v != k})

                if is_repaired:
                    wb.api.SaveAs(os.path.normpath(final_output))
                    wb.close()
                    app.quit()
                    time.sleep(0.5)
                    if os.path.exists(temp_repaired):
                        os.remove(temp_repaired)
                        log("🗑️ 已删除临时修复文件")
                else:
                    wb.api.SaveAs(os.path.normpath(final_output))
                    wb.close()
                    app.quit()

                log(f"✅ 翻译完成，输出文件：{final_output}")

            except Exception as e:
                log(f"❌ 处理过程中出错: {e}")
                try:
                    app.quit()
                except:
                    pass
                exit(1)

            output_file = os.path.join(
                folder,
                name+"_"+TARGET_LANG+ext
            )

        # ===========================
        # Word&PPT（含旧格式 .doc/.ppt 先转新格式）
        # ===========================
        elif ext in ['.docx', '.pptx', '.doc', '.ppt']:

            work_file = file_path
            legacy_tmp = None
            if ext in ['.doc', '.ppt']:
                log(f"检测到旧格式{ext}文件，尝试用 Office 转换为新格式...")
                try:
                    legacy_tmp = _convert_legacy_office(file_path)
                    work_file = legacy_tmp
                    log("✅ 旧格式转换成功，继续翻译")
                except Exception as conv_err:
                    log(f"❌ 旧格式转换失败: {conv_err}")
                    log("   可能原因：本机未安装 Office / 文件损坏 / 文件被占用；"
                        "请手动另存为 .docx 或 .pptx 后重试")
                    return

            new_ext = os.path.splitext(work_file)[1].lower()
            NEW_FILE = os.path.join(folder, f"{name}_{TARGET_LANG}{new_ext}")
            shutil.copy2(work_file, NEW_FILE)
            
            os.chmod(
                NEW_FILE,
                stat.S_IWRITE
            )
            try:

                if new_ext == '.docx':
                    log("检测到Word文件")
                    doc = Document(NEW_FILE)

                    texts_to_translate, unique_texts, text_to_id = extract_texts_from_docx(doc)

                    apply_func = apply_translations_to_docx
                    save_func = lambda: doc.save(NEW_FILE)

                else:
                    log("检测到PPT文件")
                    prs = Presentation(NEW_FILE)

                    texts_to_translate, unique_texts, text_to_id = extract_texts_from_pptx(prs)

                    apply_func = apply_translations_to_pptx
                    save_func = lambda: prs.save(NEW_FILE)

                translation_map = process_translations(unique_texts, text_to_id)
                # 只缓存成功翻译（译文为空/与原文相同不缓存——防止失败结果被永久缓存）
                cache.update({k: v for k, v in translation_map.items() if v and v != k})

                apply_func(texts_to_translate, translation_map)

                save_func()

                log(f"✅ 翻译完成，输出文件：{NEW_FILE}")

            except Exception as e:
                log(f"❌ 处理过程中出错: {e}")
                exit(1)
            finally:
                if legacy_tmp and os.path.exists(legacy_tmp):
                    try:
                        os.remove(legacy_tmp)
                    except Exception:
                        pass

        # ===========================
        # PDF
        # ===========================
        elif ext == '.pdf':

            log("检测到PDF文件")
            pdf = fitz.open(file_path)

            # 逐页分流：native（有文本层）/ scanned（无文本+大图覆盖）
            scan_pages = set()
            for p in pdf:
                chars = len(p.get_text().strip())
                area = p.rect.width * p.rect.height
                img_area = 0.0
                for im in p.get_images(full=True):
                    for r in p.get_image_rects(im[0]):
                        img_area = max(img_area, r.width * r.height)
                if chars < 20 and area > 0 and img_area > area * 0.5:
                    scan_pages.add(p.number)
            all_scanned = bool(scan_pages) and len(scan_pages) == len(pdf)

            t0 = time.time()
            if all_scanned:
                log("⚠️ 全页为扫描图，需图片转文字，启用 OCR 引擎...")
                log("⚠️ 识别精度取决于图片质量，可能发生各种识别错误。建议使用原始PDF或Word、excel等类型的源文件进行翻译")
                groups, unique_texts, text_to_id = extract_texts_from_scanned_pdf(pdf, dpi=200, pages=scan_pages)
                t1 = time.time()
                log(f"  [扫描版OCR分析] {t1 - t0:.2f}s，识别 {len(groups)} 组")
            elif scan_pages:
                log(f"⚠️ 混合PDF（{len(scan_pages)}/{len(pdf)} 页为扫描图），逐页分流处理...")
                groups, unique_texts, text_to_id = extract_texts_from_pdf(pdf)
                g_ocr, u_ocr, t_ocr = extract_texts_from_scanned_pdf(pdf, dpi=200, pages=scan_pages)
                groups = groups + g_ocr
                # 合并后重建唯一映射（去重）
                unique_texts = []
                text_to_id = {}
                for g in groups:
                    t = g["text"]
                    if t not in text_to_id:
                        text_to_id[t] = len(unique_texts)
                        unique_texts.append(t)
                    g["id"] = text_to_id[t]
                t1 = time.time()
                log(f"  [混合分析] OCR识别 {len(g_ocr)} 组，合计 {len(unique_texts)} 条唯一文本（{t1 - t0:.2f}s）")
            else:
                groups, unique_texts, text_to_id = extract_texts_from_pdf(pdf)
                t1 = time.time()
                log(f"  [PDF分析] {t1 - t0:.2f}s")

            if not groups:
                log("⚠️ 未识别到可翻译文本，跳过该文件")
                pdf.close()
                return

            miss = [t for t in unique_texts if t not in cache]
            hit_count = len(unique_texts) - len(miss)
            log(f"  PDF文本组：{len(unique_texts)}, Cache命中：{hit_count}，未命中：{len(miss)}")
            if miss:
                new_trans = process_pdf_translations(miss)
                # 只缓存成功翻译（译文为空/与原文相同不缓存——防止失败结果被永久缓存）
                cache.update({k: v for k, v in new_trans.items() if v and v != k})
            translation_map = {t: cache.get(t, t) for t in unique_texts}

            t2 = time.time()
            log(f"  [翻译/命中合并] {t2 - t1:.2f}s")
            apply_translations_to_pdf(pdf, groups, translation_map,
                                      scanned_pages=scan_pages if scan_pages else None)

            pdf_output = os.path.join(folder, name+"_"+TARGET_LANG+".pdf")
            pdf.save(pdf_output)
            pdf.close()
            t3 = time.time()
            log(f"  [PDF写回] {t3 - t2:.2f}s")

            docx_output = os.path.join(folder, name+"_"+TARGET_LANG+".docx")
            if all_scanned:
                log("确认无文本层/可编辑内容，跳过word转换，直接生成背景版DOCX")
            else:
                create_docx_pdf2docx(file_path, docx_output)
            t4 = time.time()
            log(f"  [pdf2docx转Word] {t4 - t3:.2f}s")
            docx_bg_output = os.path.join(folder, name+"_"+TARGET_LANG+"_背景版.docx")
            create_docx_from_pdf(file_path, docx_bg_output, groups)
            t5 = time.time()
            log(f"  [背景版DOCX] {t5 - t4:.2f}s")
            log("✅ 翻译完成，输出文件：")
            log(f"    {pdf_output}")
            log(f"    {docx_output}")
            log(f"    {docx_bg_output}")

        else:
            log(f"不支持的格式:{ext}")
            return

        cost = time.time()-start_time
        log(f"耗时:{cost:.2f}秒")

    except Exception as e:
        log(f"处理失败:{file_path}")
        log(str(e))

if __name__ == "__main__":

    config = get_user_config()

    if config is None:
        sys.exit(0)

    SOURCE_PATH = config["path"]
    TARGET_LANG = config["lang"]
    SOURCE_TYPE = config["type"]
    log_box = config["log_box"]
