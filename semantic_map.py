# -*- coding: utf-8 -*-
"""
字段名语义对齐（嵌入模型）
=====================================================================
干什么：把 PDF 数据表里的字段名（"Oper. Temper. 操作温度"）匹配到
        报价表模板的表头（"阀前温度℃"）。

为什么不用大模型：表格坐标定位靠 pdfplumber 的 cells 就够准，且快 100 倍；
真正需要"语义"的只有字段名对齐这一步。用 bge-small-zh（int8 量化，24MB，
CPU 上毫秒级）就够，7B 模型在 12 核 CPU 上一份要几十秒，还会把
"三偏心蝶阀""A350 LF2" 这类术语幻觉掉。

模型目录：models/bge-small-zh-v1.5-onnx/
    onnx/model_quantized.onnx       + .onnx_data（外置权重）
    tokenizer.json

没模型也能跑：自动退回"中文词干 + 字符 2-gram"的词法匹配，
准确率低一些，但零依赖。
"""
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
MODEL_DIR = os.path.join(HERE, 'models', 'bge-small-zh-v1.5-onnx')

# 报价表里这些列不该被字段自动占掉（它们由数量/公式/人工决定）
SKIP_HEADER = re.compile(r'单价|合计|总价|金额|小计|序号|备注|选型建议')

_SESS = None
_TOK = None
_MODEL_READY = None


# ---------------------------- 模型加载 ----------------------------

def _model_paths():
    d = MODEL_DIR
    onnx = os.path.join(d, 'onnx', 'model_quantized.onnx')
    tok = os.path.join(d, 'tokenizer.json')
    if not os.path.isfile(onnx):
        alt = os.path.join(d, 'onnx', 'model.onnx')
        onnx = alt if os.path.isfile(alt) else onnx
    return onnx, tok


def available():
    """模型文件在不在。"""
    global _MODEL_READY
    if _MODEL_READY is None:
        onnx, tok = _model_paths()
        _MODEL_READY = os.path.isfile(onnx) and os.path.isfile(tok)
    return _MODEL_READY


def _load():
    global _SESS, _TOK
    if _SESS is not None:
        return _SESS, _TOK
    import onnxruntime as ort
    from tokenizers import Tokenizer
    onnx, tok = _model_paths()
    so = ort.SessionOptions()
    so.intra_op_num_threads = min(4, os.cpu_count() or 1)
    so.log_severity_level = 3
    _SESS = ort.InferenceSession(onnx, so, providers=['CPUExecutionProvider'])
    _TOK = Tokenizer.from_file(tok)
    _TOK.enable_truncation(max_length=64)
    _TOK.enable_padding(length=64)
    return _SESS, _TOK


# ---------------------------- 文本清洗 ----------------------------

def clean_label(s):
    """PDF 字段名 → 干净的中文标签。
    'Oper. Temper. 操作温度' → '操作温度'
    'Tag No. 位号'           → '位号'
    纯英文的（如 'Size'）保留原样。"""
    if not s:
        return ''
    t = str(s).replace('\n', ' ')
    # 去掉常见的竖排乱序尾巴和多余符号
    t = re.sub(r'[A-Za-z][A-Za-z\.\s/&\-\(\)]*', ' ', t)     # 英文部分
    t = re.sub(r'[^\u4e00-\u9fff0-9]+', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    if t:
        return t
    # 全英文的保留
    t2 = re.sub(r'[^\w\s\.]', ' ', str(s))
    return re.sub(r'\s+', ' ', t2).strip()


def _ngrams(s, n=2):
    s = re.sub(r'\s+', '', s or '')
    if len(s) <= n:
        return {s} if s else set()
    return {s[i:i + n] for i in range(len(s) - n + 1)}


def _lexical_sim(a, b):
    """零依赖兜底：字符 2-gram 余弦 + 子串加权。"""
    A, B = _ngrams(a), _ngrams(b)
    if not A or not B:
        return 0.0
    inter = len(A & B)
    cos = inter / ((len(A) * len(B)) ** 0.5)
    if a and b and (a in b or b in a):
        cos = max(cos, 0.75)
    return cos


# ---------------------------- 编码 ----------------------------

def encode(texts):
    """返回 L2 归一化后的向量矩阵 (n, 512)。模型缺失时返回 None。"""
    if not texts:
        return None
    if not available():
        return None
    try:
        import numpy as np
        sess, tok = _load()
        enc = tok.encode_batch([str(t) for t in texts])
        ids = np.array([e.ids for e in enc], dtype=np.int64)
        mask = np.array([e.attention_mask for e in enc], dtype=np.int64)
        out = sess.run(None, {'input_ids': ids,
                              'attention_mask': mask,
                              'token_type_ids': np.zeros_like(ids)})
        names = [o.name for o in sess.get_outputs()]
        vec = out[names.index('sentence_embedding')] if 'sentence_embedding' in names else out[0][:, 0]
        norm = np.linalg.norm(vec, axis=1, keepdims=True)
        norm[norm == 0] = 1.0
        return vec / norm
    except Exception:
        return None


# ---------------------------- 对齐 ----------------------------

def read_template_headers(template_path, header_row=None):
    """读报价表模板表头 → ([(列字母, 表头文字)], 表头行号)。

    这种报价表常常是**两行表头**：第一行是主体字段（序号/位号/…/阀杆），
    第二行是右侧附件区的子表头（作用型式/型号/接口尺寸…）。
    所以先定位主体行，再用下一行把该行空着的列补上。
    """
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    # 老 .xls 也能读：openpyxl 按扩展名拦，用文件句柄绕过去
    with open(template_path, 'rb') as fh:
        wb = load_workbook(fh, read_only=False, data_only=True)
    ws = wb.worksheets[0]

    def row_cells(r):
        return {c: re.sub(r'\s+', ' ', str(ws.cell(r, c).value)).strip()
                for c in range(1, ws.max_column + 1)
                if ws.cell(r, c).value not in (None, '')}

    best_r = header_row
    if not best_r:
        # 主表头行 = A 列是"序号"或 B 列是"位号"的那一行。
        # 不能用"列数最多"来挑：这类模板右半部分还有一段配件区表头
        # （作用型式/型号/接口尺寸…），列数反而更多，会选错。
        for r in range(1, min(10, ws.max_row) + 1):
            a = re.sub(r'\s+', '', str(ws.cell(r, 1).value or ''))
            b = re.sub(r'\s+', '', str(ws.cell(r, 2).value or ''))
            if a in ('序号', 'No.', 'No') or b in ('位号', 'TagNo.', 'TagNo'):
                best_r = r
                break
        if not best_r:
            cands = []
            for r in range(1, min(10, ws.max_row) + 1):
                cells = row_cells(r)
                txt = ' '.join(cells.values())
                if re.search(r'位号|Tag\s*No|序号', txt, re.I):
                    cands.append((len(cells), r))
            best_r = max(cands)[1] if cands else 1

    main = row_cells(best_r)
    sub = row_cells(best_r + 1) if best_r + 1 <= ws.max_row else {}

    out = []
    for c in range(1, ws.max_column + 1):
        t = main.get(c) or sub.get(c)
        if not t:
            continue
        out.append((get_column_letter(c), t))
    wb.close()
    return out, best_r


def load_memory(mapping_dir=None):
    """从所有已保存的版式配置里，把"字段名 → 报价表列"沉淀成词表。

    这是第一层，也是最准的一层：同一批客户的数据表，字段名高度重复，
    历史配置里已经有人工确认过的答案，直接复用。
    返回 {干净字段名: Counter({列字母: 次数})}
    """
    import glob
    import json
    from collections import Counter
    d = mapping_dir or os.path.join(HERE, 'mappings')
    mem = {}
    for f in sorted(glob.glob(os.path.join(d, '*.json'))):
        try:
            with open(f, 'r', encoding='utf-8') as fh:
                cfg = json.load(fh)
        except Exception:
            continue
        for x in (cfg.get('fields') or []):
            if not x.get('to') or not x.get('label'):
                continue
            k = clean_label(x['label'])
            if k:
                mem.setdefault(k, Counter())[x['to']] += 1
    return mem


def _memory_hit(mem, label):
    """先精确、再互相包含，都没有返回 None。"""
    if not mem:
        return None
    k = clean_label(label)
    if not k:
        return None
    if k in mem:
        return mem[k].most_common(1)[0][0], 1.0
    best = None
    for mk, cnt in mem.items():
        if not mk:
            continue
        if mk in k or k in mk:
            # 越长的匹配越可信，短词（如"温度"）容易乱套
            s = 0.9 - 0.1 * (min(len(mk), len(k)) / max(len(mk), len(k)))
            if best is None or s > best[1]:
                best = (cnt.most_common(1)[0][0], s)
    return best


def suggest(labels, headers, min_score=0.62, one_to_one=True, memory=None):
    """三层对齐：①历史记忆 ②嵌入模型 ③（对不上的留空，交给人工框选）

    返回与 labels 等长的列表：
      {'label','to','header','score','src'}  src = '记忆' / '模型' / ''
    """
    if not labels or not headers:
        return []
    cand = [(c, h) for c, h in headers if not SKIP_HEADER.search(h)]
    if not cand:
        cand = list(headers)
    hdr_map = {c: h for c, h in cand}

    res = [{'label': labels[i], 'to': '', 'header': '', 'score': 0.0, 'src': ''}
           for i in range(len(labels))]
    used_h = set()

    # —— 第一层：历史记忆 ——
    for i, l in enumerate(labels):
        hit = _memory_hit(memory, l)
        if not hit:
            continue
        to, sc = hit
        if to not in hdr_map or (one_to_one and to in used_h):
            continue
        res[i] = {'label': l, 'to': to, 'header': hdr_map[to],
                  'score': round(sc, 3), 'src': '记忆'}
        used_h.add(to)

    # —— 第二层：嵌入模型（只补剩下的）——
    todo = [i for i, r in enumerate(res) if not r['to']]
    if not todo:
        return res

    cl = [clean_label(labels[i]) for i in todo]
    ch = [clean_label(h) or h for _, h in cand]
    L = encode(cl)
    H = encode(ch)
    if L is not None and H is not None:
        import numpy as np
        S = L @ H.T
        backend = 'bge-small-zh'
    else:
        import numpy as np
        S = np.array([[_lexical_sim(a, b) for b in ch] for a in cl], dtype='float32')
        backend = '词法兜底'

    pairs = [(float(S[a, j]), a, j) for a in range(S.shape[0])
             for j in range(S.shape[1])]
    pairs.sort(key=lambda t: (-t[0], t[2], t[1]))

    done = set()
    for sc, a, j in pairs:
        if sc < min_score:
            break
        if a in done:
            continue
        if one_to_one and cand[j][0] in used_h:
            continue
        done.add(a)
        used_h.add(cand[j][0])
        i = todo[a]
        res[i] = {'label': labels[i], 'to': cand[j][0], 'header': cand[j][1],
                  'score': round(sc, 3), 'src': '模型'}
    for r in res:
        r.setdefault('backend', backend)
        r['backend'] = backend if r['src'] == '模型' else r.get('backend', '记忆')
    return res

