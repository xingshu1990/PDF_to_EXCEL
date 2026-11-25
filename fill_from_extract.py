"""
通用填表引擎：气动阀门数据表 PDF / 提取表 → 选型报价表
=====================================================================
设计原则
--------
1. 代码只有一份，版式差异全部放进 mappings/*.json。
2. 加一种新版式 = 加一个 JSON，不改本文件。
3. 遇到 mappings 里没有的版式，自动按"字段名关键词"做一次尽力填表，
   并把这次识别到的结构存成新 JSON，下次直接复用。
4. 填表优先直接读 PDF 文字层（所有数据表类 PDF 都是文字版）；
   只有给了"提取表 xlsx"的时候才走提表路线。

三种解析方式（配置里 parse.parse_mode）
------------------------------------
  pdf_inline  直接读原始 PDF 表格，按 col_map 指定的列位取值（推荐）
  col         读"忠实转录"提取表 xlsx，列位自动探测（双阀并排版式用）
  auto        先用 pdf_inline，失败再退 col

对外接口
--------
  fill_pdf(pdf, 模板, 输出)               ← PDF 直接出报价表（主入口）
  fill_extract(提取表, 模板, 输出)         ← 提取表出报价表（兼容旧流程）
  detect_layout(路径)                      ← 识别版式，返回 (配置路径, 打分表)
  probe_pdf(pdf)                           ← 结构指纹（列数/编号列/台数）
  list_layouts()                           ← 列出 mappings 里的版式
"""
import os
import re
import json
import glob
import time
import hashlib
import shutil
import unicodedata
from collections import Counter

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter, column_index_from_string

HERE = os.path.dirname(os.path.abspath(__file__))
MAP_DIR = os.path.join(HERE, 'mappings')

# 报价表模板默认名（用户目录里的正式模板）
DEFAULT_TEMPLATE_NAMES = [
    '气动球阀（蝶阀）选型报价表.xls',
]


# ============================== 基础工具 ==============================

def _clean(v):
    if v is None:
        return ''
    return ' '.join(str(v).split())


def _load_config(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def list_layouts():
    """列出所有可用版式配置（含自动学习的）

    mappings/ 里还放着"备注翻译词典.json"这种数据文件，它不是版式配置
    （没有 fields/parse），必须跳过，否则会被当成候选版式参与打分。
    """
    out = []
    for p in sorted(glob.glob(os.path.join(MAP_DIR, '*.json'))):
        try:
            cfg = _load_config(p)
        except Exception:
            continue
        if not isinstance(cfg, dict) or not (cfg.get('fields') or cfg.get('parse')):
            continue
        if cfg.get('templates') or cfg.get('phrases'):
            continue
        out.append({
            'file': p,
            'id': cfg.get('id', ''),
            'name': cfg.get('name', os.path.basename(p)),
            'auto': bool(cfg.get('auto')),
        })
    return out


def _looks_label(s):
    """判断某格是"字段名"还是"值"。
    规则：含中文 且 含≥2个连续英文字母 且 **不含数字** → 字段名。
      字段名：'位号 TAG NO.' '数量 QUANTITIES' '管道材质：PIPE MATERIAL'
      值：'全通径球阀'（无英文）、'DN100'（无中文）、
          '1/2〞NPT(内)'（有数字）、'2区，组别IIC，T4'（有数字）
    带数字这一条是关键：像 '1/2"NPT(内)'、'机械式×2，触点容量：250VAC 3A'
    这种真正的值，光看"中文+英文"会被误杀。"""
    if not s:
        return False
    if re.search(r'\d', s):
        return False
    return bool(re.search(r'[\u4e00-\u9fff]', s) and re.search(r'[A-Za-z]{2,}', s))


# ============================== 字段名（标签）匹配 ==============================
# 为什么要有这一层
# ----------------
# 数据表里的"第几项"完全不可靠：同一家设计院改版会插一行（seq 整体 +1）、
# 多页文件后面几页还可能多一格（列位整体右移）、左右并排的版式甚至有两套 seq。
# 写死的 seq / 列号在这些情况下会静默取到隔壁那一项，比报错更糟。
#
# 所以取值统一走三步：
#   ① pat（版式里显式写的字段名正则）—— 最准，优先
#   ② 字段名相似度（中文按 2-gram、英文按整词）—— 覆盖没写 pat 的配置
#   ③ 写死的 seq —— 只在名字完全对不上时兜底
# 值本身也按"一定在字段名右边"来找，不迷信绝对列号。

LABEL_MIN_SCORE = 0.62      # 字段名相似的"最低"可信线（最后兜底用）
LABEL_HIGH_SCORE = 0.85     # 达到这条线才允许盖过版式里写死的 seq

# 明显不是"值"的占位词：数据表里当子表头/单位用
_JUNK_TOKENS = {
    'sch', 'w.t', 'w.t.', 'min', 'max', 'nor', 'normal', 'units', 'unit',
    'no.', 'no', 'rev', 'date', 'item', 'remark', 'remarks',
    '最大', '正常', '最小', '单位', '设计', '数量单位',
}
_JUNK_TOKENS_LC = {str(x).strip().lower() for x in _JUNK_TOKENS}

# 纯数值，或"0~40 / -20～80"这种量程写法：只有这种才拼单位
_NUM_OR_RANGE = re.compile(r'^[≤≥<>]?\s*-?\d+(\.\d+)?'
                           r'(\s*[~～\-—]\s*[≤≥<>]?\s*-?\d+(\.\d+)?)?$')

# 字段名里去掉后不影响识别的词（问路式停用词）
_STOP_TOKENS = {'no', 'the', 'of', 'and', 'to', 'per', 'type', 'size',
                'number', 'nominal', '号', '值', '项'}


def _is_junk(s):
    return str(s or '').strip().lower() in _JUNK_TOKENS_LC


# pdfplumber 偶尔把 "406.4 mm" 拆成 "406. mm 4"（小数点后的数字被排到单位后面），
# 这类文本顺序错乱在原表里也看得出来，取值时顺手纠正回来。
_TXT_ORDER = re.compile(r'^(\d+)\.\s*(mm|in|cm|m)\s*(\d+)$', re.I)


def tidy_value(v):
    s = str(v if v is not None else '').strip()
    m = _TXT_ORDER.match(s)
    if m:
        return f'{m.group(1)}.{m.group(3)} {m.group(2)}'
    return s


def norm_label(s):
    """字段名归一化：全角转半角、去掉空格/标点/序号，只留中文+字母+数字。"""
    if not s:
        return ''
    s = unicodedata.normalize('NFKC', str(s)).lower()
    s = re.sub(r'[\s\u3000]+', '', s)
    s = re.sub(r'[*#:：;；,，.。()（）\[\]【】\-_/\\|]+', '', s)
    return s


def label_tokens(s):
    """拆词：英文/数字按整词，中文按 2-gram（'设计温度' → 设计/计温/温度）。"""
    s = norm_label(s)
    toks = set(re.findall(r'[a-z][a-z0-9]*|\d+', s))
    cn = ''.join(re.findall(r'[\u4e00-\u9fff]', s))
    toks |= {cn[i:i + 2] for i in range(len(cn) - 1)}
    if len(cn) == 1:
        toks.add(cn)
    return {t for t in toks if t not in _STOP_TOKENS}


def label_similarity(a, b):
    """字段名相似度 0~1（a 是我们配置里的标签，b 是 PDF 那一行的字段名）。

    分词重合按 **a 的词被覆盖了多少** 算，不按两边的较小值算：
    '介质状态' 三个词对上 'Fluid State 介质状态' = 1.0；
    对上 'Fluid Name 介质' 只覆盖 1/3 = 0.28 —— 以前按较小值算会给 0.85，
    于是"介质状态"抢走"介质名称"那一行，S 列被填成介质名。
    """
    na, nb = norm_label(a), norm_label(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    short, long_ = (na, nb) if len(na) <= len(nb) else (nb, na)
    # 包含关系只认"够长"的那个方向：'执行机构' 这种区块名是会撞上
    # '执行机构型号' 的，4 个汉字太短，一律当噪声放过去给分词去判。
    if len(short) >= 5 and short in long_:
        return 0.95
    ta, tb = label_tokens(a), label_tokens(b)
    if not ta or not tb:
        return 0.0
    inter = len(ta & tb)
    if not inter:
        return 0.0
    return min(1.0, inter / len(ta)) * 0.9


def _label_owns_row(label, name, name2):
    """这个人工修正的前缀是不是这一行自己的字段名（而不是同行并排的另一个字段）。"""
    lab = str(label or '').strip()
    if not lab:
        return False
    for cand in (name or '', name2 or ''):
        if cand and label_similarity(lab, cand) >= LABEL_MIN_SCORE:
            return True
    return False


def _find_row_by_label(valve, label, min_score=None):
    """在整台阀里按字段名找一行（seq 漂了 / 一份配置合并了多套版式时用）。"""
    min_score = LABEL_HIGH_SCORE if min_score is None else min_score
    best, best_s = None, 0.0
    for f in (valve.get('fields') or {}).values():
        s = max(label_similarity(label, f.get('name') or ''),
                label_similarity(label, f.get('name2') or ''))
        if s >= min_score and s > best_s:
            best, best_s = f, s
    return best


def resolve_field(valve, item, fields=None):
    """定位配置项 item（比如"设计温度→I"）在这台阀上对应哪一行。

    返回 (seq, 该行的字段 dict)；找不到返回 (None, None)。
    fields 可传入本台阀的字段表（默认取 valve['fields']）。
    """
    fields = fields if fields is not None else (valve.get('fields') or {})
    if not fields:
        return None, None
    pairs = []
    for k, f in fields.items():
        try:
            pairs.append((int(k), f))
        except (TypeError, ValueError):
            continue
    if not pairs:
        return None, None
    by_seq = dict(pairs)

    want = item.get('seq')
    try:
        want = int(want) if want is not None else None
    except (TypeError, ValueError):
        want = None

    def _near(keys):
        """同名项出现多次（'防护等级'常见三处）时取离原 seq 最近那个。"""
        if want is None:
            return min(keys)
        return min(keys, key=lambda k: (abs(k - want), k))

    # ① pat：版式里显式写的字段名正则
    pat_s = item.get('pat')
    if pat_s:
        try:
            p = re.compile(pat_s, re.I)
        except re.error:
            p = None
        if p:
            hits = [k for k, f in pairs
                    if p.search(('%s %s' % (f.get('name') or '',
                                            f.get('name2') or '')).strip())]
            if hits:
                k = _near(hits)
                return k, by_seq[k]

    def _label_best(min_score):
        scored = []
        for k, f in pairs:
            s = max(label_similarity(label, f.get('name') or ''),
                    label_similarity(label, f.get('name2') or ''))
            if s >= min_score:
                scored.append((s, k))
        if not scored:
            return None
        top_score = max(s for s, _ in scored)
        return _near([k for s, k in scored if s == top_score])

    # ② 字段名高度吻合（几乎就是同一项）→ 采用，用来容忍改版插行
    label = item.get('label') or ''
    if label:
        k = _label_best(LABEL_HIGH_SCORE)
        if k is not None:
            return k, by_seq[k]

    # ③ 版式里写死的 seq：老配置的标签是中文简称、和原文名对不上时，
    #    以人工校过的 seq 为准（比模糊匹配可靠）
    if want is not None and want in by_seq:
        return want, by_seq[want]

    # ④ 兜底：只有中等相似度也认，总比什么都不填强
    if label:
        k = _label_best(LABEL_MIN_SCORE)
        if k is not None:
            return k, by_seq[k]
    return None, None


def _scan_value(row, start, end, skip=(), taken=()):
    """从 start 往右找第一个"像值"的格子，返回 (值, 列号)。

    end 是右栏序号列的起点：左右并排版式里左侧的值绝不能越过它。
    skip 是序号列（序号本身不是值）；taken 是已经被别处用掉的列。
    """
    end = max(0, min(end, len(row)))
    skip = set(skip) | set(taken)
    seq_texts = {str(_cell(row, i)).strip() for i in skip if _cell(row, i)}
    for i in range(max(0, start), end):
        if i in skip:
            continue
        v = _cell(row, i)
        if not v or _looks_label(v) or _is_unit_only(v) or _is_junk(v):
            continue
        if str(v).strip() in seq_texts:
            continue          # 和序号格同值 → 是隔壁栏的序号，不是值
        return v, i
    return '', None


def _find_unit(row, start, end, main_idx=None):
    """主值右边紧邻的纯单位格（'℃' / 'MPaG' / 'mm'），返回 (单位, 列号)。

    单位必须是不含数字的短串，并且一遇到下一个"像值"的格子就停——
    单位永远紧贴在值右边，越过别的值再往后找只会捡到隔壁项的。
    """
    end = max(0, min(end, len(row)))
    for i in range(max(0, start), end):
        if main_idx is not None and i == main_idx:
            continue
        u = _cell(row, i)
        if not u:
            continue
        if _is_junk(u):
            continue
        if _unit_ok(u) and not re.search(r'\d', str(u)):
            return u, i
        if not _looks_label(u) and not _is_unit_only(u):
            return '', None
    return '', None


# ============================== 结构缓存 ==============================
# 同一份 PDF 反复填表时（换模板、调映射、开核对窗口），不必每次重解析。
# 键 = 绝对路径 + 大小 + 修改时间 + 配置指纹 + 引擎版本；任一变化自动失效。

ENGINE_VER = '2026-11-25.03'
CACHE_DIR = os.path.join(HERE, '_cache')


def _cache_key(pdf_path, tag):
    try:
        st = os.stat(pdf_path)
        raw = f'{os.path.abspath(pdf_path)}|{st.st_size}|{int(st.st_mtime)}|{tag}|{ENGINE_VER}'
    except OSError:
        raw = f'{pdf_path}|{tag}|{ENGINE_VER}'
    return hashlib.md5(raw.encode('utf-8')).hexdigest()


def _cache_load(pdf_path, tag):
    p = os.path.join(CACHE_DIR, _cache_key(pdf_path, tag) + '.json')
    if not os.path.isfile(p):
        return None
    try:
        with open(p, 'r', encoding='utf-8') as f:
            obj = json.load(f)
    except Exception:
        return None
    # JSON 会把字典的整数键写成字符串，读回来必须还原成 int，
    # 否则 fields.get(seq) 全部落空（表现为"填了 0 个字段"）。
    for v in obj.get('valves') or []:
        fld = v.get('fields') or {}
        if fld and not all(isinstance(k, int) for k in fld):
            v['fields'] = {int(k): val for k, val in fld.items()}
    return obj


def _cache_save(pdf_path, tag, obj):
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        p = os.path.join(CACHE_DIR, _cache_key(pdf_path, tag) + '.json')
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(obj, f, ensure_ascii=False)
    except Exception:
        pass


def clear_cache():
    """清空结构缓存（改了 mappings 或怀疑结果不对时用）"""
    shutil.rmtree(CACHE_DIR, ignore_errors=True)


# ============================== PDF 结构指纹 ==============================

def _table_cells(page):
    """取页面最大的一张表，同时返回单元格坐标。

    返回 (data, bbox)：
      data[i][j] -> 清洗后的单元格文本
      bbox[i][j] -> [x0, top, x1, bottom]（PDF 点，左上为原点）或 None
    bbox 与 data 的行列索引严格对齐（都来自同一个 Table 对象），
    这是红线框能精确落到"取值的那一格"上的前提。
    """
    try:
        tabs = page.find_tables()
    except Exception:
        return None, None
    if not tabs:
        return None, None
    t = max(tabs, key=lambda x: (x.bbox[2] - x.bbox[0]) * (x.bbox[3] - x.bbox[1]))
    try:
        data = t.extract()
    except Exception:
        return None, None
    rows = t.rows
    out, bx = [], []
    for i, row in enumerate(data):
        cells = rows[i].cells if i < len(rows) else []
        out.append([_clean(c) for c in row])
        bx.append([[round(v, 1) for v in c] if c else None for c in cells])
    return out, bx


def _table_of(page):
    """只要文本：list[list[str]]"""
    data, _ = _table_cells(page)
    return data


def _find_seq_col(rows, ncols):
    """找出"编号列"：该列有大量纯整数、且整体递增。"""
    best, best_hit = None, 0
    for c in range(ncols):
        vals = []
        for row in rows:
            if c < len(row):
                v = row[c]
                if re.fullmatch(r'\d{1,3}', v):
                    vals.append(int(v))
        if len(vals) < 5:
            continue
        inc = sum(1 for a, b in zip(vals, vals[1:]) if b > a)
        if inc >= len(vals) * 0.6 and len(vals) > best_hit:
            best, best_hit = c, len(vals)
    return best


def probe_pdf(pdf_path, max_pages=3):
    """结构指纹：列数、编号列、每页台数、首页文字样本。
    用于在 mappings 里挑选匹配的版式配置。结果带缓存。"""
    tag = f'probe{max_pages}'
    cached = _cache_load(pdf_path, tag)
    if cached:
        cached['ncols'] = set(cached.get('ncols') or [])
        return cached

    import pdfplumber
    info = {'pages': 0, 'ncols': set(), 'ncols_all': [], 'seq_col': None,
            'seq_cols': [], 'labels': [], 'valves_per_page': 1, 'text': ''}
    with pdfplumber.open(pdf_path) as pdf:
        info['pages'] = len(pdf.pages)
        texts, rows_all, labels = [], [], []
        for page in pdf.pages[:max_pages]:
            t = _table_of(page)
            if t:
                nc = max(len(r) for r in t)
                info['ncols'].add(nc)
                info['ncols_all'].append(nc)
                rows_all.extend(t)
                # 首页/前几页的字段名清单：版式识别靠它算"这份配置能认出这篇表多少项"
                for e in _generic_rows(t, None):
                    if e['name']:
                        labels.append(e['name'])
            try:
                texts.append(page.extract_text() or '')
            except Exception:
                pass
        info['text'] = '\n'.join(texts)[:6000]
        info['labels'] = labels[:400]
        if rows_all:
            nc = max(info['ncols']) if info['ncols'] else 0
            info['seq_col'] = _find_seq_col(rows_all, nc)
            info['seq_cols'] = _seq_cols_in(rows_all, nc)
            for row in rows_all[:12]:
                ints = [c for c in row if re.fullmatch(r'\d{1,2}', c)]
                if len(ints) >= 2 and '序号' in ' '.join(row):
                    info['valves_per_page'] = 2
                    break

    _cache_save(pdf_path, tag, dict(info, ncols=sorted(info['ncols'])))
    return info


# ============================== 阀门类型识别 ==============================
# 新手拿到一份没见过的数据表时，最先能判断的就是"这是什么阀"。
# 先认类型 → 只在同类型的版式配置里挑 → 挑不出来才让人核对。
# 这样不用每来一份新 PDF 都去翻 mappings 手选。

VALVE_TYPES = [
    ('蝶阀',   r'蝶阀|蝴蝶阀|butterfly', 'struct'),
    ('球阀',   r'球阀|ball\s*valve',    'struct'),
    ('闸阀',   r'闸阀|gate\s*valve',    'struct'),
    ('截止阀', r'截止阀|globe\s*valve', 'struct'),
    ('旋塞阀', r'旋塞阀|plug\s*valve',  'struct'),
    ('隔膜阀', r'隔膜阀|diaphragm\s*valve', 'struct'),
    ('调节阀', r'调节阀|控制阀|control\s*valve', 'func'),
    ('切断阀', r'切断阀|shut\s*-?\s*off|shutoff', 'func'),
    ('开关阀', r'开关阀|on\s*-?\s*off\s*valve',  'func'),
]


def detect_valve_type(text):
    """从 PDF 文字里认阀门类型。

    返回 {'primary': '球阀', 'all': [('球阀',5),...],
          'structural': [...], 'functional': [...], 'by': 'content'}

    一份表里"球阀"和"切断阀"经常同时出现：前者是阀体结构，后者是功能。
    primary 优先取**结构类型**（它决定报价表怎么填），功能词另放 functional，
    但两者都进 all —— 筛候选配置时一个都不能漏。
    有的规格书文件名写"切断阀"、正文标题写"气动球阀规格书"，以正文为准。
    """
    low = (text or '').lower()
    if not low:
        return {'primary': '', 'all': [], 'structural': [],
                'functional': [], 'by': 'none'}
    struct_hits, func_hits = [], []
    for name, pat, group in VALVE_TYPES:
        n = len(re.findall(pat, low, re.I))
        if n:
            (struct_hits if group == 'struct' else func_hits).append((name, n))
    order = {name: i for i, (name, _, _) in enumerate(VALVE_TYPES)}
    struct_hits.sort(key=lambda x: (-x[1], order[x[0]]))
    func_hits.sort(key=lambda x: (-x[1], order[x[0]]))
    all_hits = struct_hits + func_hits
    if not all_hits:
        return {'primary': '', 'all': [], 'structural': [],
                'functional': [], 'by': 'none'}
    return {'primary': (struct_hits or func_hits)[0][0],
            'all': all_hits, 'structural': struct_hits,
            'functional': func_hits, 'by': 'content'}


def _cfg_valve_types(cfg):
    """配置声明它适用的阀门类型。没写就是通用配置。"""
    vt = cfg.get('valve_type')
    if not vt:
        return []
    if isinstance(vt, str):
        vt = [vt]
    return [str(x).strip() for x in vt if str(x).strip()]


def _candidate_layouts(vtypes, cfg_list=None):
    """按阀门类型先筛一遍候选配置。

    规则：配置声明了 valve_type 且与识别结果有交集 → 入选；
    配置没声明 valve_type → 视为通用，也入选（兼容老配置）。
    筛完一个都不剩 → 退回全部配置（宁可多打分，也不要直接判"不认识"）。
    """
    if cfg_list is None:
        cfg_list = [c['file'] for c in list_layouts()]
    if not cfg_list:
        return []
    want = set(vtypes or [])
    if not want:
        return list(cfg_list)
    picked, generic = [], []
    for p in cfg_list:
        try:
            cfg = _load_config(p)
        except Exception:
            continue
        vt = _cfg_valve_types(cfg)
        if not vt:
            generic.append(p)
        elif want & set(vt):
            picked.append(p)
    if not picked:
        return list(cfg_list)
    # 同类型配置在前，通用配置兜底在后；顺序影响不大，打分才决定胜负
    return picked + generic


# ============================== 版式识别 ==============================

# 打分门槛：字段名命中（≤10）+ 结构吻合（5）+ 类型吻合（3）+ 特征词。
# 真匹配实测都在 15 分以上，误判（同列数的另一种阀、只撞了两三个通用字段名）
# 在 8 分以下。门槛取 10，宁可多走一次人工核对，也不要悄悄填错表。
SURE_SCORE = 10


def _label_coverage(cfg, labels):
    """这份配置的字段名，在这份 PDF 的字段名里能对上多少。

    这是"自动认版式"的关键证据：同一族版式改版、插行、多一列，score_words
    可能一个都不命中，但字段名（位号/设计温度/泄露等级…）几乎不变。
    返回 {'matched': 命中项数, 'ratio': 命中率, 'matched_labels': [...]}。
    """
    labels = [l for l in (labels or []) if l]
    items = [f for f in (cfg.get('fields') or [])
             if (f.get('label') or f.get('pat')) and not f.get('meta')]
    if not labels or not items:
        return {'matched': 0, 'ratio': 0.0, 'matched_labels': []}
    hit_labels = []
    for it in items:
        pat_s = it.get('pat')
        if pat_s:
            try:
                p = re.compile(pat_s, re.I)
            except re.error:
                p = None
            if p and any(p.search(l) for l in labels):
                hit_labels.append(it.get('label') or pat_s)
                continue
        lab = it.get('label') or ''
        if lab and any(label_similarity(lab, l) >= LABEL_MIN_SCORE for l in labels):
            hit_labels.append(lab)
    # 分母取"配置项数"和"PDF 字段名数"里小的那个：配置写得比表细或比表粗都不该扣分
    denom = max(1, min(len(items), len(labels)))
    return {'matched': len(hit_labels),
            'ratio': min(1.0, len(hit_labels) / denom),
            'matched_labels': hit_labels}


def _score_cfg(cfg, text, struct, vtypes=None):
    """给单个配置打分。struct 为 None 表示输入不是 PDF（只按文字判断）。

    打分 = 字段名命中（最多 +10）+ 结构吻合（+5）+ 版式特征词（每个 +1）
         + 阀门类型吻合（+3）− 自动配置折旧（−1）
    不再因为"列数对不上"或"特征词没命中"一票否决：同一族版式经常多一列、
    少一行，那样会被误判成全新版式，白白丢掉已经配好的 mapping。
    """
    det = cfg.get('detect', {})
    low = text.lower()

    cov = {'matched': 0, 'ratio': 0.0, 'matched_labels': []}
    if struct:
        cov = _label_coverage(cfg, struct.get('labels'))
    label_score = 0
    if cov['matched'] >= 3:                    # 只命中一两项不算证据
        label_score = min(10, int(round(cov['ratio'] * 10)))

    struct_ok = True
    penalty = 0
    if struct:
        st = cfg.get('struct') or {}
        want_nc = st.get('ncols')
        if want_nc:
            want = set(want_nc if isinstance(want_nc, list) else [want_nc])
            if not (want & set(struct['ncols'])):
                struct_ok = False              # 列数不同不再否决，只是拿不到结构分
                penalty += 2
        want_v = st.get('valves_per_page')
        if want_v and int(want_v) != int(struct['valves_per_page']):
            penalty += 2

    # must_any：有字段名证据时降级为软条件，没有证据时才否决
    must = det.get('must_any', [])
    if must and not any(w.lower() in low for w in must):
        if cov['ratio'] < 0.45:
            return -1

    score = label_score
    score += sum(1 for w in det.get('score_words', []) if w.lower() in low)
    if struct and struct_ok:
        score += 5
    score -= penalty
    if cfg.get('auto'):
        # 自动生成的配置降一档：分数够（≥8）照样能自动采用，但同分时
        # 优先用人工核过的版式。以前扣 3 分 + 结构分只有 5，自动配置
        # 永远到不了门槛，等于"学了一次再也不复用"。
        score -= 3
    if (cfg.get('parse') or {}).get('parse_mode') == 'pdf_inline':
        score += 2          # 有明确列位的正式配置，比关键词猜的可靠
    # 阀门类型：对上 +3，明确对不上 −4（防止"同列数的另一种阀"抢走匹配）
    vt = _cfg_valve_types(cfg)
    if vtypes and vt:
        score += 3 if (set(vt) & set(vtypes)) else -4
    return score


def detect_layout(source_path, cfg_list=None, with_info=False):
    """识别版式。source_path 可以是 PDF 或提取表 xlsx。

    返回 (最佳配置路径 | None, [(分数, 路径, 名称), ...])
    with_info=True 时返回第三项 {'valve_type': {...}, 'sure': bool, 'threshold': 8}
    """
    if cfg_list is None:
        cfg_list = [c['file'] for c in list_layouts()]
    if not cfg_list:
        return (None, [], {'valve_type': {'primary': '', 'all': []},
                           'sure': False, 'threshold': SURE_SCORE}) if with_info \
            else (None, [])

    is_pdf = source_path.lower().endswith('.pdf')
    struct, text = None, ''
    if is_pdf:
        try:
            struct = probe_pdf(source_path)
            text = struct['text']
        except Exception:
            struct, text = None, ''
    else:
        try:
            wb = load_workbook(source_path, read_only=True, data_only=True)
            parts = []
            for ws in wb.worksheets[:2]:
                for row in ws.iter_rows(values_only=True):
                    parts.extend(_clean(v) for v in row if v is not None)
            wb.close()
            text = ' '.join(parts)[:6000]
        except Exception:
            text = ''

    # 第一步：认阀门类型；第二步：只在同类型的配置里挑
    vinfo = detect_valve_type(text)
    vtypes = [t for t, _ in vinfo['all']]
    pool = _candidate_layouts(vtypes, cfg_list)

    scored = []
    for p in pool:
        try:
            cfg = _load_config(p)
        except Exception:
            continue
        s = _score_cfg(cfg, text, struct, vtypes)
        if s < 0:
            continue
        scored.append((s, p, cfg.get('name', '')))
    scored.sort(key=lambda x: -x[0])

    info = {'valve_type': vinfo, 'sure': False, 'threshold': SURE_SCORE,
            'pool': len(pool), 'total': len(cfg_list), 'coverage': {}}
    if not scored:
        return (None, [], info) if with_info else (None, [])
    try:
        info['coverage'] = _label_coverage(_load_config(scored[0][1]),
                                           (struct or {}).get('labels'))
    except Exception:
        pass

    # 门槛只对 PDF 生效：提取表 xlsx 那条路拿不到结构分（+5），
    # 分数天生偏低，套 8 分会把本来能用的配置全判死。
    need = _load_config(scored[0][1]).get('detect', {}).get('min_score', 1)
    if struct:
        need = max(int(need), SURE_SCORE)
    info['threshold'] = need
    if scored[0][0] < need:
        return (None, scored, info) if with_info else (None, scored)
    info['sure'] = True
    return (scored[0][1], scored, info) if with_info else (scored[0][1], scored)


# ============================== PDF 解析（带坐标） ==============================

# 纯单位/占位符：这些不是真正的值，遇到就用副值顶上
_UNIT_ONLY = {'℃', '°C', '%', '％', 'dBA', 'MPa', 'MPaG', 'Mpa', 'MPa-g', 'Mpa-g',
              'MPa.A', 'MPaA', 'kg/h', 'm3/h', 'kg/m3', 'mPa.S', 'Cp', 'cp',
              'mm', 'in', '——', '---', '-', '/', '—',
              # 各家数据表写法不统一：'Cp/mpa.S'、'MpaA'、'mPa.s' 都是单位，
              # 落在这个集合外面就会被当成值写进备注（"粘度：Cp/mpa.S"）。
              'Cp/mpa.S', 'Cp/mPa.S', 'mPa.s', 'mpa.s', 'MpaA', 'MpaG',
              'm3/h', 'Nm3/h', 'kg/cm2', 'KPa', 'kPa', 'Pa', 'mbar'}
# 大小写不敏感比较：'MpaA' 和 'MPaA' 是同一个单位
_UNIT_ONLY_LC = {str(x).strip().lower() for x in _UNIT_ONLY}


def _is_unit_only(s):
    return str(s or '').strip().lower() in _UNIT_ONLY_LC


def _unit_ok(u):
    """单位格能不能直接拼到值后面。

    要挡掉的不是"空"，而是各种占位符：数据表里没单位时爱写 '——' / '---' / '-'，
    直接拼上去会变成 '28——'。另外中文也不算单位（那是备注串到单位列了）。"""
    if not u:
        return False
    u = str(u).strip()
    if not u or len(u) > 5:
        return False
    if re.fullmatch(r'[\-—–－_/\\\.\s]+', u):
        return False
    if re.search(r'[\u4e00-\u9fff]', u):
        return False
    return True


# 值本身已经带单位的样子：'1.685MPa.A'。这种不要再拼一次。
_HAS_UNIT = re.compile(r'\d\s*[A-Za-z%℃]')
# 纯数值，允许 '≤85' '>10' '<0.5' 这种带比较符的写法（噪声等级就长这样）
_NUM_LIKE = re.compile(r'^[≤≥<>]?\s*-?\d+(\.\d+)?$')


def _cell(row, i):
    if i is None or i < 0 or i >= len(row):
        return ''
    return row[i]


def _row_bboxes(row, bxs):
    """这一行每个格子的坐标，供核对窗口把红框画到"真正取值的那一格"。

    只存非空格：一份表 20 多列 × 每台阀一行，存全了缓存会翻倍。
    """
    out = {}
    for i in range(len(row or [])):
        b = _bbox_at(bxs, i)
        if b:
            out[i] = b
    return out


# 单位写法统一一下：'deg C' 在报价表里写 ℃ 更符合工程习惯
_UNIT_FIX = {'degc': '℃', 'degf': '℉',
             'mpa-g': 'MPaG', 'mpag': 'MPaG', 'mpa.g': 'MPaG', 'mpa': 'MPa',
             'kg/m3': 'kg/m³', 'kg/cm2': 'kg/cm²', 'm3/h': 'm³/h', 'nm3/h': 'Nm³/h'}


def _unit_text(u):
    s = str(u or '').strip()
    if not s:
        return ''
    return _UNIT_FIX.get(s.lower().replace(' ', ''), s)


def _bbox_at(bxs, i):
    if not bxs or i is None or i < 0 or i >= len(bxs):
        return None
    return bxs[i]


# 划掉的值 = 否定。工程设计图里常用"文字上划一道线"表示"这一项不要/不配"，
# 但文字本身还是照常写（如"手轮装置 → 配带"上划线，实际是"不配带"）。
# 这种线在 pdfplumber 的 lines/rects 里常常解析不出来，要从 PyMuPDF 的
# drawings 里找一条细长的斜线/横线，看它是否穿过这个单元格的文字。
_STRIKE_CACHE = {}


def _strike_lines(pdf_path, page_idx):
    """取这一页上"像删除线"的线段，带缓存。"""
    key = (os.path.abspath(pdf_path), int(page_idx))
    if key in _STRIKE_CACHE:
        return _STRIKE_CACHE[key]
    segs = []
    try:
        import pymupdf as fitz
        doc = fitz.open(pdf_path)
        try:
            page = doc[page_idx]
            for dr in page.get_drawings():
                if dr.get('type') not in ('s', 'fs'):
                    continue
                box = dr.get('rect')
                if not box:
                    continue
                w, h = box.x1 - box.x0, box.y1 - box.y0
                # 删除线：横向细长（允许轻微倾斜），不能是表格边框那种整行宽
                if w < 8 or w > 200 or h > 8 or h >= w * 0.4:
                    continue
                for it in dr.get('items') or []:
                    if it and it[0] == 'l' and len(it) >= 3:
                        a, b = it[1], it[2]
                        segs.append((float(a.x), float(a.y),
                                     float(b.x), float(b.y)))
        finally:
            doc.close()
    except Exception:
        pass
    _STRIKE_CACHE[key] = segs
    return segs


def _seg_crosses_box(seg, box, ratio=0.35):
    """线段是否穿过 bbox 的中部（不是擦边）。"""
    x0, y0, x1, y1 = box
    ax, ay, bx, by = seg
    w, h = x1 - x0, y1 - y0
    if w <= 0 or h <= 0:
        return False
    # 采样线段，看有多少点落在 bbox 中间那一带
    hits = 0
    n = 12
    for i in range(n + 1):
        t = i / n
        px, py = ax + (bx - ax) * t, ay + (by - ay) * t
        if (x0 + w * 0.02 <= px <= x1 - w * 0.02 and
                y0 + h * ratio <= py <= y1 - h * ratio):
            hits += 1
    return hits >= max(2, n // 4)


def _struck_value(v, bbox, pdf_path, page_idx):
    """值被划线划掉 → 返回否定写法；没划线返回原值。"""
    if not v or not bbox or len(bbox) < 4:
        return v
    for seg in _strike_lines(pdf_path, page_idx):
        if _seg_crosses_box(seg, bbox):
            # 已经是"不带/无/不"就不重复加
            if re.match(r'^(不|无|未|non)', str(v).strip(), re.I):
                return v
            return f'不{v}'
    return v


def _parse_one_valve(row, cm, bxs=None):
    """从一行里取出该阀门的 {编号, 主值, 副值}，并记下取值单元格的坐标。

    取值优先级（实测覆盖四种版式）：
      1) 第二字段名列是"值"而不是字段名 → 取它（阀体/执行机构段，值就写在这一列）
      2) 否则取主值列（不对主值列做字段名过滤——'2区，组别IIC，T4'、'法兰连接 RF'
         这类真正的值，用"含中文+英文"去判会被误杀）
      3) 主值是纯单位（℃/MPa/kg/h…）→ 用副值顶上
    """
    seq = _cell(row, cm.get('seq'))
    if not re.fullmatch(r'\d{1,3}', seq):
        return None

    i_name2, i_val, i_val2 = cm.get('name2'), cm.get('val'), cm.get('val2')
    name2 = _cell(row, i_name2)
    val = _cell(row, i_val)
    val2 = _cell(row, i_val2)
    name = _cell(row, cm.get('name'))

    # 少数版式里字段名太长，尾巴被排到了"值"那一列：
    # 例如 'Power Failure Pos 电源故障阀' + '位'（12 列版第 49 行）。
    # 不能靠"值很短"去猜——'带'、'无'、'FC'、'侧装'、'2"' 都是合法值，
    # 所以这种续名由版式配置显式声明（parse.col_map.name_cont）。
    nc = cm.get('name_cont') or []
    if name2 and any(str(x) == str(seq) for x in nc):
        name = f'{name}{name2}'
        name2 = ''

    # 两列字段名都空，且值列里没有一个数字 → 不是数据行。
    # 典型是 @Max./@Nor./@Min./Unit 那种子表头（会把 '@ Max.' 读成值）。
    # 但不能一刀切：'250 / Sch30 / 7.8 / 257.4 / mm' 这种"续行"也没有字段名，
    # 却是管道尺寸那几项真正的值，所以要求"值列无数字"才跳过。
    # 注意只看值列：序号格里本身就带数字，算进去就永远跳不掉。
    if not name and not name2:
        vals = [_cell(row, i) for i in (i_val, cm.get('val_nor'), i_val2)
                if i is not None]
        if not any(re.search(r'\d', v) for v in vals if v):
            return None

    main, mi = '', None
    if name2 and not _looks_label(name2):
        main, mi = name2, i_name2
    elif val:
        main, mi = val, i_val

    # 有的表把工艺数据拆成三列 @Max. / @Nor. / @Min.（部分数据表如此），
    # 温度这类参数经常只填了 Nor. 一列，主值列是空的。
    # 只在主值确实为空时兜底取 Nor.，不影响已经取到 Max. 的行。
    if not main:
        i_nor = cm.get('val_nor')
        if i_nor is not None:
            v_nor = _cell(row, i_nor)
            if v_nor and not _looks_label(v_nor):
                main, mi = v_nor, i_nor

    sub, si = '', None
    if val2 and not _looks_label(val2):
        sub, si = val2, i_val2

    # 人工框选写回的补丁：用户在核对窗口里指过"这一项的值在这一列"，
    # 直接以它为准（既能补空，也能纠正取错的）。
    ovs = cm.get('overrides')
    if ovs and not cm.get('_skip_ov'):
        for o in ovs:
            try:
                if int(o.get('seq', -1)) != int(seq):
                    continue
            except (TypeError, ValueError):
                continue
            c = o.get('col')
            v = _cell(row, c)
            if not v or _looks_label(v):
                continue
            if o.get('label'):
                # 按前缀绑定：只记下"这个前缀用哪一列"，不整体覆盖主/副值。
                # 真正取值时由 _value_for_label 按 label 精确取，避免一行的
                # 另一个字段被顺带改掉（防火结构 / 防杆吹出 都是"是"的坑）。
                continue
            if o.get('sub'):
                sub, si = v, c
            else:
                main, mi = v, c

    if _is_unit_only(main):
        if sub:
            main, mi, sub, si = sub, si, '', None
        else:
            main, mi = '', None

    # —— 行内相对定位（抗列漂移 / 抗左右并排）——
    # 两种情况都会让写死的 val 列取偏：
    #   ① 多页表的后面几页某一项多插了一格，整行右移；
    #   ② 左右并排版式（本页 21 列）里，左栏 val 列可能正好落在右栏的序号上。
    # 统一处理：值是"字段名右边第一个像值的格子"，右边界是右栏序号列。
    seq_col = cm.get('seq')
    name_col = cm.get('name')
    i_seq_r = cm.get('seq_r')
    seq_idx = {c for c in (seq_col, i_seq_r) if isinstance(c, int)}
    ncols = len(row)
    if isinstance(i_seq_r, int) and isinstance(name_col, int) and i_seq_r > name_col:
        right_edge = i_seq_r
    else:
        right_edge = ncols
    # 主值等于本行某个序号格 → 那是隔壁栏的序号被当成值了（实测 "用途=63"）
    if main and re.fullmatch(r'\d{1,3}', str(main)) and str(main) in {
            str(_cell(row, i)).strip() for i in seq_idx if _cell(row, i)} \
            and int(main) != int(seq):
        main, mi = '', None
    if not main and not sub:
        v, i = _scan_value(row, (name_col if isinstance(name_col, int) else 0) + 1,
                           right_edge, skip=seq_idx)
        if v:
            main, mi = v, i

    # 单位：配置没写 unit_col 时自动取主值右边紧邻的纯单位格
    unit, unit_idx = '', None
    if mi is not None and cm.get('unit_auto') and _NUM_OR_RANGE.match(str(main or '')):
        u, ui = _find_unit(row, mi + 1, right_edge, mi)
        if u:
            unit, unit_idx = u, ui

    # 按前缀的人工修正：{label: (值, 列号, bbox)}，供备注/字段精确取值
    label_vals, label_bbox, label_cols = {}, {}, {}
    for o in (ovs or []) if not cm.get('_skip_ov') else []:
        lab = o.get('label')
        if not lab:
            continue
        try:
            if int(o.get('seq', -1)) != int(seq):
                continue
        except (TypeError, ValueError):
            continue
        c = o.get('col')
        v = _cell(row, c)
        if not v or _looks_label(v):
            continue
        label_vals[str(lab)] = v
        label_bbox[str(lab)] = _bbox_at(bxs, c)
        label_cols[str(lab)] = c

    # 人工修正如果指的就是这一行自己的字段名，主值也跟着改——
    # 否则核对窗口里 "已保存" 了、红框那格显示的还是旧值，看着像没生效。
    for lab, lv in label_vals.items():
        if not _label_owns_row(lab, name, name2):
            continue
        main, mi = lv, label_cols.get(lab)
        break

    return {'seq': int(seq), 'name': name, 'name2': name2,
            'main': tidy_value(main), 'sub': tidy_value(sub), 'raw': row,
            'bbox': _bbox_at(bxs, mi), 'sub_bbox': _bbox_at(bxs, si),
            'bboxes': _row_bboxes(row, bxs),
            'col': mi, 'sub_col': si, 'unit': unit, 'unit_col': unit_idx,
            'label_vals': label_vals, 'label_bbox': label_bbox}


def _parse_right_section(row, cm, bxs=None):
    """单阀版式右侧附加栏（定位器/电磁阀/限位开关/气源装置）。
    值列位置不稳定，从 val_r 起扫，跳过字段名，第一个是主值、第二个是副值。"""
    seq = _cell(row, cm.get('seq_r'))
    if not re.fullmatch(r'\d{1,3}', seq):
        return None
    start = cm.get('val_r', 10)
    skip = {cm.get('seq_r'), cm.get('name_r')}
    cands = []
    for i in range(start, len(row)):
        if i in skip:
            continue
        v = row[i]
        # 单位/占位符/子表头（℃、Max、Units）不是值：不挡掉的话
        # "最小工作温度"会取到单位格 '℃'，写进报价表就是一行垃圾。
        if not v or _looks_label(v) or _is_unit_only(v) or _is_junk(v):
            continue
        cands.append((v, i))
    unit, unit_idx = '', None
    if cands and _NUM_OR_RANGE.match(str(cands[0][0] or '')):
        unit, unit_idx = _find_unit(row, cands[0][1] + 1, len(row), cands[0][1])
    return {'seq': int(seq), 'name': _cell(row, cm.get('name_r')), 'name2': '',
            'main': tidy_value(cands[0][0]) if cands else '',
            'sub': tidy_value(cands[1][0]) if len(cands) > 1 else '',
            'raw': row, 'sec': 1,
            'unit': unit, 'unit_col': unit_idx,
            'bboxes': _row_bboxes(row, bxs),
            'bbox': _bbox_at(bxs, cands[0][1]) if cands else None,
            'sub_bbox': _bbox_at(bxs, cands[1][1]) if len(cands) > 1 else None}


def label_sources_for_row(cfg, seq, name, name2):
    """列出这一行所有"字段名前缀 → 取值"的候选对应关系。

    同一个序号里经常并排两个字段（如 seq10 = 防火结构 col7 / 防杆吹出 col9），
    核对窗口要把它们都显示出来，并允许人工框选时改前缀归属。
    三项来源：fields[]（按 seq 绑定，写报价表列）、remark.include_labels
    （按字段名正则匹配，写备注列）、以及原始 name/name2 兜底。
    """
    out, seen = [], set()

    def add(label, which, to, src, src_col=None):
        key = str(label or '').strip()
        if not key or key in seen:
            return
        seen.add(key)
        out.append({'label': key, 'which': which, 'to': to or '', 'src': src,
                    'src_col': src_col})

    for item in cfg.get('fields', []):
        if item.get('meta') or item.get('seq') is None:
            continue
        if str(item['seq']) != str(seq):
            # seq 漂移（改版插行）时按字段名也能认出这一行归哪一项
            lab = item.get('label') or ''
            if not lab or max(label_similarity(lab, name or ''),
                              label_similarity(lab, name2 or '')) < LABEL_MIN_SCORE:
                continue
        add(item.get('label') or name,
            'sub' if int(item.get('sub', 0)) == 1 else 'main',
            item.get('to'), '报价表列', item.get('src_col'))

    rcfg = cfg.get('remark') or {}
    for spec in rcfg.get('include_labels') or []:
        pat = spec.get('pat')
        if not pat:
            continue
        try:
            p = re.compile(pat, re.I)
        except re.error:
            continue
        which = 'main' if (name and p.search(name)) else (
            'sub' if (name2 and p.search(name2)) else None)
        if not which:
            continue
        add(spec.get('label') or pat, which, rcfg.get('to', 'Q'), '备注',
            spec.get('src_col'))

    if name:
        add(name, 'main', '', '字段名')
    if name2:
        add(name2, 'sub', '', '第二字段名')
    return out


def _pick_tag(fields):
    """挑位号：优先字段名里出现 Tag/位号 的那一行，其次 seq2、seq1。
    不能直接取 seq1——有的版式 seq1 是"序号"，位号在 seq2。"""
    best = ''
    for seq, f in fields.items():
        name = (f.get('name') or '') + ' ' + (f.get('name2') or '')
        if f.get('main') and re.search(r'tag|位号', name, re.I):
            return f['main']
    for seq in (2, 1):
        v = fields.get(seq, {}).get('main')
        if v:
            best = v
            break
    return best


# 汇总小表的表头 → 内部键。各家写法不同，所以用正则而不是写死文字。
_SUMMARY_KEYS = [
    ('tag',      r'位号|tag'),
    ('service',  r'用途|service|介质名称'),
    ('pipe',     r'管道号|pipe\s*no'),
    ('pid',      r'流程图|p\s*&\s*i\s*d|p&id'),
    ('fluid',    r'介质|fluid'),
    ('temp',     r'操作温度|温度|temp'),
    ('pressure', r'操作压力|压力|press'),
    ('dp',       r'关闭压差|差压|dp'),
    ('remark',   r'备注|remark'),
]


def _summary_meta(data, sumcfg):
    """从表格下方的"汇总小表"里取 位号/用途/介质 等。

    有些规格书主规格区根本没有位号，位号只写在下方的汇总表里。
    结构固定为：一行表头（含"位号/Tag"）→ 紧跟若干数据行。
    返回 [{'tag': '0204A-XV-2001A', 'service': '富液过滤器切断球阀', ...}, ...]
    """
    if not sumcfg or not data:
        return []
    head_any = sumcfg.get('header_any') or ['位号', 'Tag']
    key_pats = sumcfg.get('keys') or {}
    hdr_i, colmap = -1, {}
    for i, row in enumerate(data):
        hit = sum(1 for c in row
                  if c and any(w.lower() in c.lower() for w in head_any))
        if hit < 1:
            continue
        cm_ = {}
        for j, c in enumerate(row):
            if not c:
                continue
            for key, pat in _SUMMARY_KEYS:
                pat2 = key_pats.get(key) or pat
                if re.search(pat2, c, re.I) and key not in cm_:
                    cm_[key] = j
        # 表头至少要认出"位号"这一列，否则不是汇总表
        if 'tag' in cm_ and len(cm_) >= 2:
            hdr_i, colmap = i, cm_
            break
    if hdr_i < 0:
        return []

    tag_re = sumcfg.get('tag_pat') or r'^[A-Za-z0-9][A-Za-z0-9\-_/\.]{3,}$'
    out = []
    for row in data[hdr_i + 1:]:
        tag = _cell(row, colmap.get('tag'))
        if not tag or not re.search(tag_re, tag):
            continue
        meta = {'tag': tag}
        for key, j in colmap.items():
            if key == 'tag':
                continue
            v = _cell(row, j)
            if v:
                meta[key] = v
        out.append(meta)
    return out


def _parse_page(page, cm, vpp, pageno, cm_by_ncols=None):
    """解析一页，返回 (该页阀门列表, 页面元信息)。阀门里带 bbox。
    cm_by_ncols：同一版式不同页列数不同时（如首页 9 列、其余 8 列），按列数切换列位。"""
    data, bxs = _table_cells(page)
    if data and cm_by_ncols:
        nc = max(len(r) for r in data)
        alt = cm_by_ncols.get(str(nc))
        if alt:
            cm = alt          # 已带 overrides（parse_pdf_detail 注入过）
    meta_info = {'width_pt': round(page.width, 1), 'height_pt': round(page.height, 1),
                 'table_bbox': None}
    if not data:
        return [], meta_info

    # 表格外框 = 所有单元格的并集
    xs0 = [c[0] for row in (bxs or []) for c in row if c]
    ys0 = [c[1] for row in (bxs or []) for c in row if c]
    xs1 = [c[2] for row in (bxs or []) for c in row if c]
    ys1 = [c[3] for row in (bxs or []) for c in row if c]
    if xs0:
        meta_info['table_bbox'] = [round(min(xs0), 1), round(min(ys0), 1),
                                   round(max(xs1), 1), round(max(ys1), 1)]

    buckets = [dict() for _ in range(vpp)]
    for ri, row in enumerate(data):
        rbx = bxs[ri] if bxs and ri < len(bxs) else None
        if vpp == 1:
            d = _parse_one_valve(row, cm, rbx)
            if d:
                cur = buckets[0].get(d['seq'])
                if not cur or (not cur['main'] and d['main']):
                    buckets[0][d['seq']] = d
                elif not cur['sub'] and d['sub']:
                    cur['sub'] = d['sub']
                    cur['sub_bbox'] = d['sub_bbox']
            r = _parse_right_section(row, cm, rbx)
            if r:
                cur = buckets[0].get(r['seq'])
                if not cur:
                    buckets[0][r['seq']] = r
                elif not cur['main'] and r['main']:
                    cur['main'], cur['bbox'] = r['main'], r['bbox']
                    cur['sub'] = cur['sub'] or r['sub']
                    cur['sub_bbox'] = cur['sub_bbox'] or r['sub_bbox']
        else:
            for vi in range(2):
                cmx = dict(cm)
                if vi == 1:
                    cmx['val'] = cm.get('val_r')
                    cmx['val2'] = cm.get('val2_r')
                    cmx['_skip_ov'] = True   # 右侧那台阀的列位不同，人工补丁不套用
                d = _parse_one_valve(row, cmx, rbx)
                if d:
                    cur = buckets[vi].get(d['seq'])
                    if not cur or (not cur['main'] and d['main']):
                        buckets[vi][d['seq']] = d

    out = []
    # 位号写在下方汇总表里的版式：先把汇总表读出来
    sumcfg = cm.get('_summary')
    sum_rows = _summary_meta(data, sumcfg) if sumcfg else []
    # 汇总表也没有时，退一步在整页文字里按正则捞一个位号
    tag_rx = cm.get('_tag_regex')
    text_tag = ''
    if tag_rx and not sum_rows:
        try:
            ptxt = page.extract_text() or ''
        except Exception:
            ptxt = ''
        m = re.search(tag_rx, ptxt)
        if m:
            text_tag = m.group(0).strip()

    for vi, fields in enumerate(buckets):
        tag = _pick_tag(fields)
        meta_ = None
        if sum_rows:
            # 一台阀一页时直接取第一行；多台时按顺序对号入座
            if len(sum_rows) == len([b for b in buckets if b]):
                idxs = [i for i, b in enumerate(buckets) if b]
                meta_ = sum_rows[idxs.index(vi)]
            elif vi == 0:
                meta_ = sum_rows[0]
        if meta_ and meta_.get('tag'):
            tag = meta_['tag']
        elif text_tag:
            tag = text_tag
        if not tag:
            continue
        v = {'page': f'第{pageno}页', 'page_idx': pageno - 1, 'index': vi,
             'tag': tag, 'fields': fields}
        if meta_:
            v['meta'] = meta_
        out.append(v)
    return out, meta_info


def _parse_worker(args):
    """多进程工作单元：解析指定页（可 pickle）。"""
    pdf_path, cm, vpp, page_idxs, cm_by_ncols = args
    import pdfplumber
    valves, metas = [], {}
    with pdfplumber.open(pdf_path) as pdf:
        for pi in page_idxs:
            vs, meta = _parse_page(pdf.pages[pi], cm, vpp, pi + 1, cm_by_ncols)
            # 划线=否定：值上有删除线时改成"不XX"（如"配带"→"不配带"）。
            # 放在这一层统一做，主值/副值/备注三条路都会用到改后的值。
            if vs and _strike_lines(pdf_path, pi):
                for v in vs:
                    for f in (v.get('fields') or {}).values():
                        if f.get('main') and f.get('bbox'):
                            f['main'] = _struck_value(f['main'], f['bbox'],
                                                      pdf_path, pi)
                        if f.get('sub') and f.get('sub_bbox'):
                            f['sub'] = _struck_value(f['sub'], f['sub_bbox'],
                                                     pdf_path, pi)
            valves.extend(vs)
            metas[str(pi + 1)] = meta
    return valves, metas


def _cfg_tag(cfg):
    parse = cfg.get('parse') or {}
    return 'parse:%s:%s' % (cfg.get('id', ''),
                            json.dumps(parse, sort_keys=True, ensure_ascii=False))


def _with_overrides(cm, ovs, parse=None):
    """把 parse.overrides（人工框选写回的修正）挂到列位表上。
    overrides 放在 parse 下便于人看，但解析函数拿到的是 col_map，
    所以这里做一次注入；按列数切换列位时也要重新注入。

    顺带把 parse 层的两个"版式级"设置也挂上去，因为 _parse_page 只拿得到 col_map：
      tag_regex —— 位号用正则从整页文字里捞（主规格区没有位号字段的版式）
      summary   —— 位号/用途/介质写在下方汇总小表的版式
    """
    c = dict(cm or {})
    if ovs:
        c['overrides'] = ovs
    if parse:
        if parse.get('tag_regex'):
            c['_tag_regex'] = parse['tag_regex']
        if parse.get('summary'):
            c['_summary'] = parse['summary']
        # 单位自动拼接是版式级开关，但解析函数只拿得到 col_map，这里带下去
        if parse.get('unit_auto'):
            c['unit_auto'] = True
    return c


def parse_pdf_detail(pdf_path, cfg, workers=None, use_cache=True):
    """解析 PDF，返回 {'valves': [...], 'pages': {页码: 元信息}}。
    阀门字段里带 bbox（PDF 点坐标），供红线框窗口使用。
    策略：结构缓存 → 多进程分页 → 单进程兜底。"""
    parse = cfg.get('parse') or {}
    # 未知版式自动存的配置是 generic 模式：不按 col_map 取列，而是按
    # “标签 vs 值”分类。它的 col_map 只是从样本反推的一个参考，真实的
    # 值列散在各处；若照旧套 col_map，会一个字段都取不到，核对窗口和
    # 填表都会误报“没解析到阀门”。这里必须按配置声明的模式分流。
    if parse.get('parse_mode') == 'generic':
        valves, _ = parse_pdf_generic(pdf_path, cfg=cfg)
        _attach_notes(pdf_path, cfg, valves)
        return {'valves': valves, 'pages': {}, 'cached': False}
    ovs = parse.get('overrides') or []
    cm = _with_overrides(parse.get('col_map') or {}, ovs, parse)
    byn = parse.get('col_map_by_ncols') or None
    cm_by_ncols = ({k: _with_overrides(v, ovs, parse) for k, v in byn.items()}
                   if byn else None)
    vpp = int(parse.get('valves_per_page', 1))
    tag = _cfg_tag(cfg)

    if use_cache:
        cached = _cache_load(pdf_path, tag)
        if cached:
            cached['cached'] = True
            # 备注来自页面文字、不进解析缓存的结构里，读缓存时补一次
            _attach_notes(pdf_path, cfg, cached.get('valves') or [])
            return cached

    import pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        n = len(pdf.pages)

    idxs = list(range(n))
    valves, metas = [], {}
    cpu = os.cpu_count() or 1
    if workers is None:
        workers = min(6, cpu, max(1, n // 4)) if n >= 8 else 1

    if workers > 1:
        from concurrent.futures import ProcessPoolExecutor
        chunks = [idxs[i::workers] for i in range(workers)]
        chunks = [c for c in chunks if c]
        try:
            with ProcessPoolExecutor(max_workers=len(chunks)) as ex:
                for vs, ms in ex.map(_parse_worker,
                                     [(pdf_path, cm, vpp, c, cm_by_ncols) for c in chunks]):
                    valves.extend(vs)
                    metas.update(ms)
        except Exception:
            valves, metas = [], {}
            workers = 1

    if not valves and not metas:
        valves, metas = _parse_worker((pdf_path, cm, vpp, idxs, cm_by_ncols))

    valves.sort(key=lambda v: (v.get('page_idx', 0), v.get('index', 0)))
    _attach_notes(pdf_path, cfg, valves)
    detail = {'valves': valves, 'pages': metas, 'cached': False}

    if use_cache:
        _cache_save(pdf_path, tag, detail)
    return detail


def parse_pdf(pdf_path, cfg):
    """兼容接口：只要阀门列表"""
    return parse_pdf_detail(pdf_path, cfg)['valves']


def audit_pdf(pdf_path, cfg=None):
    """空值体检：把报价表里的空格分成两类，回答"为什么有这么多空格"。

      ① PDF 原始就是空的 —— 客户没填，程序填不出来是正常的
      ② PDF 里有值却没取到 —— 真正的漏读，需要改配置或改引擎

    返回 {'total','filled','blank_pdf','suspect':[...],'dead_fields':[...],'part_fields':[...]}
    """
    if cfg is None:
        cp, scores = detect_layout(pdf_path)
        if not cp and scores:
            # 分数没到门槛时 cp 是 None，但候选里往往有上次存下的自动配置；
            # 直接把 None 传给 _load_config 会抛 TypeError，菜单11 看起来
            # 就是"没有任何提示"。这里退回候选的第一名。
            cp = scores[0][1]
        cfg = (_load_config(cp) if cp else
               {'id': 'generic', 'name': '通用兜底',
                'parse': {'parse_mode': 'generic'}, 'fields': []})
    detail = parse_pdf_detail(pdf_path, cfg)
    valves = detail['valves']
    parse = cfg.get('parse') or {}
    cm = parse.get('col_map') or {}
    seq_l, seq_r = cm.get('seq'), cm.get('seq_r')
    name_l, name_r = cm.get('name'), cm.get('name_r')
    seq_cols = {c for c in (seq_l, seq_r) if isinstance(c, int)}

    # meta 项（从下方汇总表取位号/用途）没有 seq，也不参与"漏读"体检
    flds = [f for f in (cfg.get('fields') or [])
            if f.get('to') and not f.get('meta') and f.get('seq') is not None]
    per, suspect = {}, []
    for v in valves:
        for f in flds:
            # 体检和填表必须用同一套定位逻辑，否则"体检说没漏、表里却是空的"
            seq, rec = resolve_field(v, f)
            if seq is None:
                seq = int(f['seq'])
            sub = bool(f.get('sub'))
            val = ''
            if rec:
                val = (rec.get('sub') if sub else rec.get('main')) or ''
            key = (f.get('to'), f.get('label'), seq, sub)
            per.setdefault(key, 0)
            if val:
                per[key] += 1
                continue
            raw = (rec or {}).get('raw') or []
            used = {rec.get(k) for k in ('col', 'sub_col') if rec and rec.get(k) is not None}
            # 这一项属于左栏还是右栏：左右并排版式两栏的取值窗口不一样，
            # 用同一个区间扫会把隔壁栏的值当成"漏读"（以前报一堆假漏读）。
            in_right = (isinstance(seq_r, int) and isinstance(seq, int)
                        and isinstance(name_r, int)
                        and _cell(raw, seq_r) == str(seq)
                        and (not isinstance(seq_l, int) or _cell(raw, seq_l) != str(seq)))
            if in_right:
                lo, hi = name_r + 1, len(raw)
            else:
                lo = (name_l + 1) if isinstance(name_l, int) else 0
                hi = (seq_r if isinstance(seq_r, int) and seq_r > lo else len(raw))
            cands = []
            for i in range(max(0, lo), min(hi, len(raw))):
                if i in seq_cols or i in used:
                    continue
                cell = raw[i]
                if not cell or _looks_label(cell) or _is_unit_only(cell) or _is_junk(cell):
                    continue
                if re.fullmatch(r'\d{1,3}', cell):
                    continue
                cands.append((i, cell))
            if cands:
                suspect.append({'tag': v['tag'], 'to': f.get('to'), 'label': f.get('label'),
                                'seq': seq, 'sub': sub, 'cands': cands})

    n = len(valves)
    total = n * len(flds)
    filled = sum(per.values())
    return {
        'config': cfg.get('name'),
        'valves': n,
        'fields': len(flds),
        'total': total,
        'filled': filled,
        'blank_pdf': total - filled - len(suspect),
        'suspect': suspect,
        'dead_fields': [k for k, c in per.items() if c == 0],
        'part_fields': [(k, c) for k, c in per.items() if 0 < c < n],
    }



def _valve_groups(rows, seq_cols=None):
    """一页里"几台阀并排"？返回每台阀各自的值列组。

    判据：表头行里 Unit / 最大 / 正常 / 最小 这套"值列表头"重复出现。
    气动球阀303机组 这类表一页左右并排两台阀，表头就是
      Unit | 最大 | 正常 | 最小 | Unit | 最大 | 正常 | 最小
    重复两遍 = 两台阀。只有一遍 = 一台阀（值列 = 全部值列）。
    两组之间若还夹着序号列，那是"左栏/右栏"两段（同一台阀），不是两台阀。
    """
    import re as _re
    tok = _re.compile(r'(@?\s*(?:max|nor|normal|min)\.?|最大|正常|最小|units?|单位)', _re.I)
    unit_tok = _re.compile(r'(units?|单位)', _re.I)
    best, hdr_row = None, None
    for row in rows:
        cols = [i for i, v in enumerate(row) if v and tok.search(str(v).strip())]
        if len(cols) >= 2 and (best is None or len(cols) > len(best)):
            best, hdr_row = cols, row
    if not best:
        return []
    starts = [c for c in best if unit_tok.search(str(hdr_row[c] or ''))]
    if len(starts) >= 2:
        groups = []
        for k, st in enumerate(starts):
            end_c = starts[k + 1] - 1 if k + 1 < len(starts) else best[-1]
            groups.append([c for c in best if st <= c <= end_c])
    else:
        groups, cur = [], [best[0]]
        for c in best[1:]:
            if c - cur[-1] <= 3:
                cur.append(c)
            else:
                groups.append(cur)
                cur = [c]
        groups.append(cur)
    if len(groups) < 2:
        return []
    if any(groups[0][-1] < sc < groups[1][0] for sc in (seq_cols or [])):
        return []
    out = []
    for g in groups:
        lo, hi = max(0, g[0]), g[-1]
        if lo > 0 and _is_unit_only(_cell(rows[0], lo - 1) or ''):
            lo -= 1
        out.append(list(range(lo, hi + 1)))
    return out


def _group_pick(raw, cols):
    """从这台阀自己的值列组里取 (主值, 副值, max, nor, min, 列号, 单位)。"""
    unit = ''
    if cols and _is_unit_only(_cell(raw, cols[0]) or ''):
        unit = _cell(raw, cols[0])
        cols = cols[1:]
    vals = [(i, _cell(raw, i)) for i in cols]
    vals = [(i, v) for i, v in vals if v and not _is_unit_only(v)
            and not _is_junk(v) and not _looks_label(v)]
    # cols 现在按 [max, nor, min] 排（或 [主值] / [主值, 副值]）
    max_i = nor_i = min_i = None
    if len(cols) >= 3:
        max_i, nor_i, min_i = cols[0], cols[1], cols[2]
    if nor_i is not None:
        nor = _cell(raw, nor_i)
        main = nor if nor and not _is_unit_only(nor) else (
            vals[0][1] if vals else '')
    else:
        main = vals[0][1] if vals else ''
    sub = vals[1][1] if len(vals) > 1 and vals[1][1] != main else ''
    mx = _cell(raw, max_i) if max_i is not None else ''
    mn = _cell(raw, min_i) if min_i is not None else ''
    return (tidy_value(main), tidy_value(sub), tidy_value(mx),
            tidy_value(mn), unit, vals[0][0] if vals else None)


def _seq_cols_in(rows, ncols):
    """找出所有"编号列"。多数版式只有一列；左右并排的版式有两列。

    判据：该列有 ≥5 个纯整数、且整体递增。左右并排的数据表（如
    气动截断阀数据表：左栏 1~60、右栏 61~120）两列都满足，返回 [0, 12]。
    """
    found = []
    for c in range(ncols):
        vals = []
        for row in rows:
            if c < len(row):
                v = row[c]
                if re.fullmatch(r'\d{1,3}', v or ''):
                    vals.append(int(v))
        if len(vals) < 5:
            continue
        inc = sum(1 for a, b in zip(vals, vals[1:]) if b > a)
        if inc >= len(vals) * 0.6:
            found.append((c, len(vals)))
    found.sort(key=lambda x: -x[1])
    return [c for c, _ in found[:2]]


def _generic_rows(data, bxs):
    """把一张表拆成"字段行"：seq / 字段名 / 主值 / 副值 / 单位 / 原始行 / 坐标。

    一页里如果有两套序号（左右并排），两栏都会解析出来——这正是以前
    附件区（执行机构/电磁阀/限位开关/气源装置）永远填不出来的原因：
    旧实现只认第一套序号，右栏那些项根本没进 fields。
    """
    ncols = max((len(r) for r in data), default=0)
    if not ncols:
        return []
    seq_cols = _seq_cols_in(data, ncols)
    if not seq_cols:
        return []
    # 先找"值列"，**按栏分别算**：每栏里出现最多的那几列是值列；
    # 序号和值列之间那些"纯中文短词"列（旁路手阀/壁厚/管线防护类型）是字段名，
    # 绝不能当成值——否则"管线压力等级"会取到"旁路手阀"。
    seq_texts = {str(_cell(row, c)).strip() for row in data for c in seq_cols
                 if _cell(row, c)}
    nrow = max(1, len(data))
    cnt, chinese = {}, {}
    for c in range(ncols):
        n = cn = 0
        for row in data:
            v = _cell(row, c)
            if not v or _is_unit_only(v) or _is_junk(v) or _looks_label(v):
                continue
            if re.fullmatch(r'\d{1,3}', str(v)) or str(v).strip() in seq_texts:
                continue
            n += 1
            if re.fullmatch(r'[\u4e00-\u9fff（）()、/·\s]{1,10}', str(v).strip()):
                cn += 1
        cnt[c], chinese[c] = n, cn

    val_cols, key_cols = set(), set()
    for si, sc in enumerate(seq_cols):
        end = seq_cols[si + 1] if si + 1 < len(seq_cols) else ncols
        rng = [c for c in range(sc + 1, end)]
        rng2 = [c for c in rng if c != sc + 1]
        dom = max(rng2, key=lambda c: cnt.get(c, 0)) if rng2 else None
        for c in rng:
            if c == sc + 1:
                key_cols.add(c)          # 序号右边第一列永远是字段名
                continue
            if dom is not None and c < dom and cnt[c] and chinese[c] / cnt[c] >= 0.6:
                key_cols.add(c)          # 值列左边、几乎全是中文短词 → 字段名
                continue
            if cnt[c] >= max(2, int(nrow * 0.3)) or \
                    (dom is not None and c > dom and cnt[c] >= max(2, int(nrow * 0.08))):
                val_cols.add(c)
    val_cols -= key_cols

    out = []
    for ri, row in enumerate(data):
        rbx = bxs[ri] if bxs and ri < len(bxs) else None
        for si, sc in enumerate(seq_cols):
            s = _cell(row, sc)
            if not re.fullmatch(r'\d{1,3}', s or ''):
                continue
            # 本栏的右边界 = 下一套序号列的起点（右栏则是行尾）
            end = seq_cols[si + 1] if si + 1 < len(seq_cols) else len(row)
            # 序号和第一个值列之间是"字段名区"：第一格是字段名，后面的是
            # 同行的第二个字段名（如 '入口管线尺寸 | 壁厚'）
            first_val = min([c for c in val_cols if sc < c < end], default=end)
            keys = []
            for i in range(sc + 1, first_val):
                v = _cell(row, i)
                if v and not _is_unit_only(v) and not _is_junk(v):
                    keys.append(v)
            name = keys[0] if keys else ''
            name2 = keys[1] if len(keys) > 1 else ''
            vals = []
            for i in range(max(first_val, sc + 1), end):
                if i not in val_cols:
                    continue
                v = _cell(row, i)
                if not v or _looks_label(v) or _is_unit_only(v) or _is_junk(v):
                    continue
                if str(v).strip() in seq_texts:
                    continue          # 隔壁栏的序号，不是值
                vals.append((v, i))
            main, mi = vals[0] if vals else ('', None)
            sub, si2 = vals[1] if len(vals) > 1 else ('', None)
            unit, ui = ('', None)
            if mi is not None and _NUM_OR_RANGE.match(str(main or '')):
                unit, ui = _find_unit(row, mi + 1, end, mi)
            out.append({'seq': int(s), 'name': name, 'name2': name2, 'raw': row,
                        'main': tidy_value(main), 'sub': tidy_value(sub),
                        'col': mi, 'sub_col': si2,
                        'unit': unit, 'unit_col': ui,
                        'bboxes': _row_bboxes(row, rbx),
                        'bbox': _bbox_at(rbx, mi) if mi is not None else None,
                        'sub_bbox': _bbox_at(rbx, si2) if si2 is not None else None,
                        'sec': si})
    return out


def _apply_overrides_to_fields(fields, ovs):
    """把人工修正套到 generic 解析出来的字段上。

    旧实现里 generic 模式直接 return，overrides 一条都不生效——用户在核对
    窗口框完、配置里也写进去了，但下次填表还是老样子，看起来就是"没保存"。
    """
    for o in ovs or []:
        try:
            seq = int(o.get('seq', -1))
        except (TypeError, ValueError):
            continue
        f = fields.get(seq)
        if not f:
            continue
        c = o.get('col')
        row = f.get('raw') or []
        if c is None or not (0 <= int(c) < len(row)):
            continue
        v = _cell(row, c)
        if not v or _looks_label(v):
            continue
        lab = o.get('label')
        if lab:
            # 按前缀绑定：真取值时由 resolve_field 读 label_vals，不整体覆盖
            f.setdefault('label_vals', {})[str(lab)] = v
            # 前缀就是这一行自己的字段名时，主值同步改，红框/取值才马上跟着变
            if _label_owns_row(lab, f.get('name'), f.get('name2')):
                f['main'], f['col'] = v, int(c)
        elif o.get('sub'):
            f['sub'], f['sub_col'] = v, int(c)
        else:
            f['main'], f['col'] = v, int(c)
    return fields


def parse_pdf_generic(pdf_path, max_pages=0, cfg=None):
    """未知版式的兜底解析：不依赖 col_map，用"字段名 vs 值"分类。

    max_pages>0 时只取前 N 页做样本（人工核对阶段用，省得等整本解析完）。
    cfg 给了就把 parse.overrides 套上去（人工修正必须在这个模式生效）。
    """
    import pdfplumber
    st = probe_pdf(pdf_path, max_pages=max_pages or 3)
    ovs = (((cfg or {}).get('parse') or {}).get('overrides') or [])
    out = []
    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages[:max_pages] if max_pages else pdf.pages
        for pi, page in enumerate(pages, 1):
            data, bxs = _table_cells(page)
            if not data:
                continue
            entries = _generic_rows(data, bxs)
            if not entries:
                continue
            # 一页两台阀并排（值列表头 Max/Nor/Min 出现两遍）→ 拆成两台
            groups = _valve_groups(data)
            if len(groups) >= 2:
                got = False
                for gi, cols in enumerate(groups[:2]):
                    f2 = {}
                    for e in entries:
                        main_, sub_, mx, mn, unit, ci = _group_pick(
                            e.get('raw') or [], cols)
                        if not main_ and not sub_:
                            continue
                        if e['seq'] in f2:
                            continue
                        item = dict(e)
                        item.update({'main': main_, 'sub': sub_,
                                     'max': mx, 'nor': main_, 'min': mn,
                                     'unit': unit, 'col': ci, 'sec': gi})
                        f2[e['seq']] = item
                    if not f2:
                        continue
                    tag2 = _pick_tag(f2)
                    if tag2:
                        out.append({'page': f'第{pi}页', 'page_idx': pi - 1,
                                    'index': gi, 'tag': tag2, 'fields': f2,
                                    'two_valves': True})
                        got = True
                if got:
                    continue
            fields = {}
            for e in entries:
                seq = e['seq']
                if seq not in fields:
                    fields[seq] = e
                else:
                    cur = fields[seq]
                    if not cur.get('main') and e.get('main'):
                        fields[seq] = e
                    elif not cur.get('sub') and e.get('sub'):
                        cur['sub'], cur['sub_col'] = e['sub'], e['sub_col']
                        cur['sub_bbox'] = e.get('sub_bbox')
            if ovs:
                _apply_overrides_to_fields(fields, ovs)
            tag = _pick_tag(fields)
            if tag:
                out.append({'page': f'第{pi}页', 'page_idx': pi - 1, 'index': 0,
                            'tag': tag, 'fields': fields})
    return out, st


# ============================== 映射核对用的 middle 结构 ==============================
# 参考 MinerU 的 middle.json：每页若干带 bbox 的语义块。
# 左边红线框、右边内容清单都从这里渲染，前端不需要再碰 PDF 解析逻辑。

def _page_text_blocks(pdf_path, page_idx):
    """用 PyMuPDF 取页面文字块（标题/正文/页脚），比 pdfplumber 快很多。"""
    try:
        import pymupdf as fitz
    except Exception:
        try:
            import fitz
        except Exception:
            return []
    try:
        doc = fitz.open(pdf_path)
        page = doc[page_idx]
        d = page.get_text('dict')
        doc.close()
    except Exception:
        return []
    blocks = []
    for b in d.get('blocks', []):
        if b.get('type') != 0:
            continue
        txt = ' '.join(''.join(sp.get('text', '') for sp in ln.get('spans', []))
                       for ln in b.get('lines', [])).strip()
        txt = ' '.join(txt.split())
        if not txt:
            continue
        sizes = [sp.get('size', 0) for ln in b.get('lines', []) for sp in ln.get('spans', [])]
        blocks.append({'text': txt, 'bbox': [round(v, 1) for v in b['bbox']],
                       'size': round(max(sizes) if sizes else 0, 1)})
    return blocks


def _overlap(a, b):
    if not a or not b:
        return False
    ix = min(a[2], b[2]) - max(a[0], b[0])
    iy = min(a[3], b[3]) - max(a[1], b[1])
    return ix > 0 and iy > 0


def _row_bands(bboxes):
    """把一组 bbox 按纵向合并成"行带"，用于判断某段文字是不是表格里的字段行。
    整页表格的 table_bbox 往往盖住页眉，用它排除会把标题一起排掉；
    改用"有字段的那些行"来排除，标题/页脚/备注就能留下来。"""
    bands = []
    for b in sorted(bboxes, key=lambda x: x[1]):
        if bands and b[1] <= bands[-1][1] + 3:
            bands[-1][1] = max(bands[-1][1], b[3])
        else:
            bands.append([b[1], b[3]])
    return bands


def _in_bands(bbox, bands, ratio=0.4):
    """文字块纵向有 ratio 以上落在字段行带里 → 视为表格内容"""
    if not bbox or not bands:
        return False
    h = max(1.0, bbox[3] - bbox[1])
    for y0, y1 in bands:
        ov = min(bbox[3], y1) - max(bbox[1], y0)
        if ov > 0 and ov / h >= ratio:
            return True
    return False


# ============================== 页面备注（Notes / 表格下方说明） ==============================
# 数据表底部经常有"备注"：有的写成表格最后一行（Notes: … / 备注：…），
# 有的在表格下面单独一段；**同一份 PDF 不同页的备注还可能不一样**。
# 报价表里给它两列：备注原文 + 翻译备注（英文备注翻译成中文后写在第二列）。

_NOTE_LABEL = re.compile(r'^\s*(?:notes?|note|备注)\s*[:：]?\s*', re.I)
# 页眉页脚/版权/页码这类"表格下方的文字"不是备注，别当备注抓
_NOTE_SKIP = re.compile(
    r'未经.{0,8}许可|不得扩散|版权所有|版权|copyright|all rights reserved|'
    r'^\s*第\s*\d+\s*页|^\s*共\s*\d+\s*页|^\s*\d+\s*/\s*\d+\s*$', re.I)
# 纯文档编号页脚（'JL 02ZC11 IN5-72'）：全大写字母+数字，没有小写字母
_NOTE_CODE = re.compile(r'^[A-Z][A-Z0-9 \-./_#]{3,}$')


def _split_note_labels(s):
    """去掉 'Notes:'/'备注：' 标记；同一段备注被重复贴两遍时只留一份。"""
    s = _clean(s)
    if not s:
        return ''
    # 冒号可有可无：'备注 …' 后面常常直接跟正文，没有冒号
    parts = [p.strip() for p in re.split(r'\s*(?:notes?|备注)\s*[:：]?\s*', s, flags=re.I)]
    parts = [p for p in parts if p]
    if not parts:
        return ''
    if len(parts) >= 2 and (parts[0] == parts[1] or parts[0] in parts[1]):
        return parts[1]
    return ' '.join(parts)


def _page_note(pdf_path, page_idx, fn_cfg):
    """取某一页的备注：① 表内 Notes/备注 行 ② 表格下方的文字。"""
    min_len = int(fn_cfg.get('min_len', 8))
    note = ''
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_idx]
            data, bxs = _table_cells(page)
            # ① 表内备注行
            for row in (data or []):
                txt = ' '.join(c for c in row if c)
                if _NOTE_LABEL.match(txt) and len(txt) >= min_len:
                    note = _split_note_labels(txt)
                    break
            # ② 表格下方文字
            if not note and fn_cfg.get('below_table', True):
                ys = [c[3] for r in bxs or [] for c in r if c]
                ymax = max(ys) if ys else page.height
                lines = []
                for b in _page_text_blocks(pdf_path, page_idx):
                    if b['bbox'][1] < ymax - 2:
                        continue
                    t = _clean(b['text'])
                    if not t or len(t) < min_len or _NOTE_SKIP.search(t):
                        continue
                    if _NOTE_CODE.match(t) and re.search(r'\d', t):
                        continue          # 文档编号，不是备注
                    if t not in lines:
                        lines.append(t)
                note = ' '.join(lines)
    except Exception:
        return ''
    return note


def _attach_notes(pdf_path, cfg, valves):
    """把每页备注挂到该页的每台阀上（fill 时按行写进两列）。"""
    fn_cfg = (cfg.get('parse') or {}).get('footnote') or {}
    if not fn_cfg or not valves:
        return valves
    by_page = {}
    for v in valves:
        pi = int(v.get('page_idx', 0) or 0)
        if pi not in by_page:
            by_page[pi] = _page_note(pdf_path, pi, fn_cfg)
        v['footnote'] = by_page[pi]
    return valves


# ---------------- 备注翻译（离线词典，可自己往里加） ----------------
NOTE_DICT_PATH = os.path.join(HERE, 'mappings', '备注翻译词典.json')
_NOTE_DICT = None


def note_dict():
    """读备注翻译词典（第一次用时加载，改了文件重启程序生效）。"""
    global _NOTE_DICT
    if _NOTE_DICT is not None:
        return _NOTE_DICT
    try:
        with open(NOTE_DICT_PATH, 'r', encoding='utf-8') as f:
            _NOTE_DICT = json.load(f) or {}
    except Exception:
        _NOTE_DICT = {}
    return _NOTE_DICT


def _tm_templates(d):
    out = []
    for it in d.get('templates') or []:
        en, zh = it.get('en'), it.get('zh')
        if not en or not zh:
            continue
        pat = re.escape(en).replace(r'\{n\}', r'([0-9][0-9.]*)')
        try:
            out.append((re.compile(pat, re.I), zh))
        except re.error:
            continue
    return out


def translate_note(text):
    """备注翻译：先用「整句/模板」翻译记忆，再按长短语替换。

    词典里没有的英文会原样留在"翻译备注"里 —— 看到没翻的，
    把它加进 mappings/备注翻译词典.json 的 phrases 即可（下次生效）。
    """
    s = _clean(text)
    if not s:
        return ''
    # 已经是中文（没有成词的英文）→ 原样返回
    if len(re.findall(r'[A-Za-z]{3,}', s)) < 2:
        return s
    d = note_dict()
    for rx, zh in _tm_templates(d):
        def _sub(m, zh=zh):
            out = zh
            for i, g in enumerate(m.groups(), 1):
                out = out.replace('{n}', g, 1) if '{n}' in out else out
            return out
        s = rx.sub(_sub, s)
    phrases = sorted((p for p in (d.get('phrases') or [])
                      if p.get('en') and p.get('zh')),
                     key=lambda p: -len(p['en']))
    for p in phrases:
        s = re.sub(re.escape(p['en']), p['zh'], s, flags=re.I)
    s = re.sub(r'\s+([，。；：、])', r'\1', s)
    s = re.sub(r'\s{2,}', ' ', s)
    # 中文后面的英文标点换成中文标点：'并配管.' → '并配管。'
    s = re.sub(r'(?<=[\u4e00-\u9fff])\s*\.\s*', '。', s)
    s = re.sub(r'(?<=[\u4e00-\u9fff])\s*,\s*', '，', s)
    return s.strip()


def build_middle(pdf_path, cfg, detail=None):
    """组装 middle 结构：每页 = 若干 block（field / title / text）。
    field 块带 to 列、标签、值、bbox；title/text 是表格外的页面文字。"""
    if detail is None:
        detail = parse_pdf_detail(pdf_path, cfg)
    valves = detail.get('valves') or []
    pages_meta = detail.get('pages') or {}

    # seq -> 映射项（可能一个 seq 对应多列，如 29→AF + C）
    # 注意：fields 里可能有 {"meta":"tag","to":"B"} 这种从表格下方汇总表取值的项，
    # 它没有 seq 也没有 bbox，画不了红框，核对窗口里直接跳过（填表时仍然生效）。
    by_seq = {}
    for item in cfg.get('fields', []):
        if item.get('meta') or item.get('seq') is None:
            continue
        by_seq.setdefault(int(item['seq']), []).append(item)

    # 备注里的字段（include_labels）也要画框：以前只画 fields[] 的，
    # 一行里两个同值（如 seq10 防火结构/防杆吹出 都是"是"）时，
    # 第二个前缀在核对窗口里根本看不到，也就没法核对/框选改它。
    # 具体来源由 label_sources_for_row 在下面逐行算。

    pages = {}
    for v in valves:
        pi = v.get('page_idx', 0)
        pno = pi + 1
        meta = pages_meta.get(str(pno), {})
        pg = pages.setdefault(pno, {
            'page_idx': pi, 'page_no': pno,
            'width_pt': meta.get('width_pt', 0), 'height_pt': meta.get('height_pt', 0),
            'table_bbox': meta.get('table_bbox'), 'blocks': [],
        })
        for seq, f in sorted(v['fields'].items(), key=lambda kv: int(kv[0])):
            srcs = label_sources_for_row(cfg, seq, f.get('name'), f.get('name2'))
            # 一个格子只画一个框：同一格可能被多个来源指认
            # （fields 用原文名、remark 用中文短名），只留 mapped 的那条。
            slots = {}
            for item in by_seq.get(int(seq)) or []:
                is_sub = bool(int(item.get('sub', 0)) == 1)
                which = 'sub' if is_sub else 'main'
                lab = item.get('label') or f.get('name') or str(seq)
                prev = slots.get(which)
                if prev and not prev[0]:      # 已有 mapped 项，别覆盖
                    continue
                slots[which] = (bool(item.get('to')), lab, item.get('to'),
                                item.get('src_col'))
            for src in srcs:
                which = src['which']
                prev = slots.get(which)
                # 原文名（src='字段名'/'第二字段名'）优先级最低
                if prev and (prev[0] or src['src'] in ('字段名', '第二字段名')):
                    continue
                slots[which] = (bool(src.get('to')), src['label'],
                                src.get('to') or None, src.get('src_col'))

            for which, (mapped, lab, to, src_col) in slots.items():
                val = f.get('sub') if which == 'sub' else f.get('main')
                bbox = f.get('sub_bbox') if which == 'sub' else f.get('bbox')
                # 配了 src_col 的项，红框要画到"真正取走的那一格"上
                # （@Nor 那一列、法兰 Standard/Size/Rating/Face 各列）
                if src_col is not None:
                    cols = src_col if isinstance(src_col, (list, tuple)) else [src_col]
                    bxs_map = f.get('bboxes') or {}
                    got = [bxs_map.get(int(c)) for c in cols if bxs_map.get(int(c))]
                    if got:
                        val = ' / '.join(str(_cell(f.get('raw') or [], int(c)))
                                         for c in cols
                                         if _cell(f.get('raw') or [], int(c))) or val
                        # 只画第一格的框（多列合并映射就画最左边那格）
                        bbox = got[0]
                if not val:
                    continue
                pg['blocks'].append({
                    'type': 'field', 'seq': seq, 'label': lab, 'to': to,
                    'value': val, 'bbox': bbox,
                    'tag': v['tag'], 'mapped': mapped,
                })

        # 写进"新增列"（extra_columns）的格子也要画红框：值确实进了报价表，
        # 不画的话核对窗口里这些格子是青绿的，看着像"没映射"。
        for g in cfg.get('extra_columns') or []:
            for s in (g.get('subs') or [g]):
                _, ff = resolve_field(v, s)
                if not ff:
                    continue
                sc = s.get('src_col')
                raw = ff.get('raw') or []
                bbox = None
                if sc is not None:
                    if isinstance(sc, (list, tuple)):
                        boxes = [ff.get('bboxes', {}).get(int(i)) for i in sc]
                        boxes = [b for b in boxes if b]
                        val = ' '.join(str(_cell(raw, int(i))) for i in sc
                                       if _cell(raw, int(i)))
                        bbox = boxes[0] if boxes else None
                    else:
                        val = _cell(raw, int(sc))
                        bbox = (ff.get('bboxes') or {}).get(int(sc))
                else:
                    val = ff.get('main') or ff.get('sub')
                    bbox = ff.get('bbox') or ff.get('sub_bbox')
                if not val or not bbox:
                    continue
                pg['blocks'].append({
                    'type': 'field', 'seq': int(s.get('seq') or 0),
                    'label': str(g.get('title') or s.get('header') or ''),
                    'to': '新增列', 'value': val, 'bbox': bbox,
                    'tag': v['tag'], 'mapped': True,
                })

    # 补上表格外的页面文字（标题 / 正文 / 备注）
    for pno, pg in pages.items():
        field_boxes = [b['bbox'] for b in pg['blocks']
                       if b['type'] == 'field' and b.get('bbox')]
        bands = _row_bands(field_boxes)
        tb = pg.get('table_bbox')
        for b in _page_text_blocks(pdf_path, pg['page_idx']):
            if _in_bands(b['bbox'], bands):
                continue          # 落在字段行上，已由 field 块表达
            # 表格框内的短词基本是表头/分区标签（Unit、Max.Flow、BODY…），
            # 留长句（备注、说明），避免右侧面板被这些词淹没。
            if tb and _overlap(b['bbox'], tb) and len(b['text']) < 24:
                continue
            pg['blocks'].append({'type': 'title' if b['size'] >= 14 else 'text',
                                 'text': b['text'], 'bbox': b['bbox'], 'mapped': False})
        # 页面文字排在前面，字段在后
        pg['blocks'].sort(key=lambda x: (0 if x['type'] in ('title', 'text') else 1,
                                         (x.get('bbox') or [0, 0, 0, 0])[1]))

    # 每一项在全文的取值情况：有几台填在主值列、几台填在副值列。
    # 有的数据表某页这一行是空的、另一页又有值，只看当前页会判断错，
    # 所以统计放在 middle 里，核对窗口的弹窗直接引用。
    seq_stats = {}
    for v in valves:
        for seq, f in v['fields'].items():
            s = seq_stats.setdefault(int(seq), {'main': 0, 'sub': 0, 'total': 0})
            s['total'] += 1
            if f.get('main'):
                s['main'] += 1
            if f.get('sub'):
                s['sub'] += 1

    return {
        'schema': 'tablekit.middle', 'schema_version': '1.0',
        'pdf': os.path.abspath(pdf_path),
        'config': cfg.get('name', ''), 'config_id': cfg.get('id', ''),
        'valves': len(valves), 'cached': detail.get('cached', False),
        'pages': [pages[k] for k in sorted(pages)],
        'seq_stats': {str(k): v for k, v in seq_stats.items()},
    }


# ============================== 关键词兜底映射 ==============================

# (正则, 目标列, 变换, 是否合并)  —— 顺序敏感：越具体越靠前
# 只在"没见过版式"时用来兜底生成映射（fields_from_keywords），以及人工框选
# 后猜测"这一项该写哪一列"。手写版式配置不读这里。
_KEYWORD_RULES = [
    # —— 附件区：执行机构 / 电磁阀 / 阀位开关 / 气源装置 ——
    # 这些项以前全靠人工补，现在按字段名也能自动落到附件区那些列上
    (r'solenoid\s*model|电磁阀型号', 'BY', None, False),
    (r'limit\s*switch\s*model|限位开关型号|阀位开关型号', 'CR', None, False),
    (r'actuator\s*type|acutator\s*type|执行机构型式|执行机构形式|执行机构类型|执行器类型',
     'BK', None, False),
    (r'actuator\s*(size|model)|acutator\s*(size|model)|执行机构型号|执行器型号',
     'BM', None, False),
    (r'rated\s*travel|额定行程', 'BU', None, False),
    (r'air\s*failure|气源故障|故障时阀位|故障位置|fail\s*position', 'BO', 'failure_pos', False),
    (r'hand\s*wheel|handle\s*wheel|手轮', 'BP', None, False),
    (r'travel\s*time|行程时间|动作时间|全行程', 'O', None, False),
    (r'action\s*mode|动作方式|作用形式', 'BQ', None, False),
    (r'全压差开关|full\s*diff', 'BX', None, False),
    (r'travel\s*indicator|阀位指示', 'BV', None, False),
    (r'防沙处理|sediment', 'BW', None, False),
    (r'失电|断电阀位置|de-?energize|electricity\s*failure', 'CE', None, False),
    (r'供电|power\s*supply|电源(?!故障)', 'CG', None, False),
    (r'类型/数量|type\s*/\s*quantity', 'CQ', None, False),
    (r'solenoid\s*type|电磁阀形式', 'CB', None, False),
    (r'limit\s*switch\s*type|限位开关形式|阀位开关形式', 'CY', None, False),
    (r'field\s*indicator|现场指示', 'CU', None, False),
    (r'接点容量|contacts?\s*/?\s*rating|触电容量', 'DD', None, False),
    (r'动作点|actuation\s*point|触电位置|触点位置', 'CX', None, False),
    (r'气连接口|气源接口', 'DI', None, False),
    (r'压力表|pressure\s*gauge', 'DJ', None, False),
    (r'储气罐|air\s*storer', 'DK', None, False),
    (r'过滤减压|air\s*regulat|filter\s*-?\s*regulator', 'DE', None, False),
    (r'供气压力|气源压力|air\s*(press|supp)', 'BR', None, False),
    (r'铭牌|custom\s*tag', 'BI', None, False),
    (r'特殊需求|special\s*preparation', 'BJ', None, False),
    (r'设计温度下最大压力|max\s*press', 'BG', None, False),
    (r'最小工作温度|min\s*working', 'BH', None, False),
    (r'防爆标志|explosion\s*proof|防爆等级', 'CA', None, False),
    (r'防护等级|ip\s*code|enclosure\s*protec', 'BZ', None, False),
    (r'电气连接|conduit|elec\.?\s*conn', 'CH', None, False),
    (r'pneu|气接口|气动接口', 'DG', None, False),

    # —— 主规格区 ——
    (r'tag|位号', 'B', None, False),
    (r'quantit|数量', 'F', None, False),
    (r'service|用途', 'G', None, False),
    (r'p&id|pid\s*no', 'R', None, False),
    (r'valve\s*code|阀门编码', 'AD', None, False),
    (r'upstr.*rating|上游管线压力等级', 'X', None, False),
    (r'pressure\s*rating|公称压力|压力等级', 'C', 'pressure_class', False),
    (r'upstr.*(nominal|size)|上游管线尺寸', 'Y', None, False),
    (r'upstr.*schedule|上游管线规格', 'Z', None, False),
    (r'upstr.*material|上游管线材质', 'AA', None, False),
    (r'line\s*no|pipe\s*no|管线号|管道号|工艺管号', 'W', None, False),
    (r'公称尺寸|normal\s*size|line\s*size|管道尺寸|管线尺寸|公称通径|body\s*size|'
     r'port\s*size|阀口径|口径', 'D', None, False),
    (r'pipe\s*class|管道等级', 'AE', None, False),
    (r'pipe\s*material|line\s*material|管道材质', 'AA', None, False),
    (r'area\s*class|危险区域', 'U', None, False),
    (r'min\s*ign|最小燃点', 'V', None, False),
    (r'fluid\s*state|介质状态', 'S', None, False),
    (r'fluid\s*phase|介质相态', 'T', None, False),
    (r'fluid|介质', 'H', None, False),
    (r'design\s*press|设计压力', 'M', None, False),
    (r'design\s*temp|设计温度', 'I', None, False),
    (r'oper\s*press|操作压力|inlet\s*press|入口压力|阀前压力', 'L', None, False),
    (r'oper\s*temp|操作温度|temperature|入口温度', 'J', None, False),
    (r'shut[\s\-]*off|最大关闭差压', 'K', None, False),
    (r'valve\s*type|body\s*type|阀门类型|阀形式|阀门形式|形式', 'N', None, False),
    (r'end\s*conn|连接尺寸|conn\.?\s*type|连接形式|process\s*connection', 'AF', None, False),
    (r'construction\s*type|结构形式', 'AG', None, False),
    (r'special\s*construction|特殊结构', 'AH', None, False),
    (r'operator\s*type|操作方式', 'AI', None, False),
    (r'line\s*design\s*code|管线设计标准', 'AB', None, False),
    (r'design\s*(standard|code)|设计标准', 'AJ', None, False),
    (r'dimensions|结构长度', 'AK', None, False),
    (r'bonnet\s*mat|阀盖材质', 'AL', None, True),
    (r'body\s*mat|本体材质|阀体材质|阀体\s*&\s*阀盖', 'AL', None, True),
    (r'seat\s*mat|阀座材质', 'AM', None, True),
    (r'trim\s*mat|阀芯材质|内件\s*&\s*密封面', 'AM', None, True),
    (r'stem\s*mat|阀杆材质|阀杆', 'AN', None, False),
    (r'spring|弹簧', 'AO', None, False),
    (r'packing|填料', 'AP', None, False),
    (r'bolt.*material|螺栓&?螺母材质', 'BE', None, False),
    (r'gasket\s*(material|mat)|垫片材质', 'BF', None, False),
    (r'bolt|螺栓', 'AQ', None, False),
    (r'gasket|垫片', 'AR', None, False),
    (r'sour\s*service|酸性工况', 'AS', None, False),
    (r'inspection|检验级别', 'AT', None, False),
    (r'fugitive|逸散', 'AU', None, False),
    (r'leak|泄漏|泄露', 'AV', None, False),
    (r'hydrostatic\s*shell|壳体试验\(液压', 'AW', None, False),
    (r'pneumatic\s*shell|壳体试验\(气压', 'AX', None, False),
    (r'hydrostatic\s*seat|密封试验\(液压', 'AY', None, False),
    (r'lp\s*gas|低压密封', 'AZ', None, False),
    (r'hp\s*gas|高压密封', 'BA', None, False),
    (r'fire\s*test|耐火试验', 'BB', None, False),
    (r'cryogenic|低温测试', 'BC', None, False),
    (r'companion\s*flange|配对法兰', 'BD', None, False),
]


def fields_from_keywords(valve):
    """按字段名关键词生成映射项，用于未知版式兜底。"""
    items, used = [], set()
    for seq, f in sorted(valve['fields'].items()):
        name = (f.get('name') or '') + ' ' + (f.get('name2') or '')
        low = name.lower()
        for pat, to, tf, merge in _KEYWORD_RULES:
            if not re.search(pat, low):
                continue
            if to in used and not merge:
                break
            items.append({'seq': seq, 'to': to, 'label': f.get('name', '')[:16],
                          'transform': tf, 'merge': merge})
            used.add(to)
            break
    return items


def _ov_key(o):
    """人工修正的身份：序号 + 主/副 + 前缀名。
    带 label 的是"按前缀取值"，和旧的"按主/副取值"互不覆盖。"""
    try:
        return (int(o.get('seq', -1)), bool(o.get('sub')),
                str(o.get('label') or ''))
    except (TypeError, ValueError):
        return (-1, False, '')


def _guess_to_for_label(label, used):
    """按字段名猜这一项该写报价表哪一列（只用于人工框选后的补映射）。"""
    for pat, to, tf, merge in _KEYWORD_RULES:
        if to in used and not merge:
            continue
        try:
            if re.search(pat, str(label or ''), re.I):
                return to, tf
        except re.error:
            continue
    return '', None


def add_override(cfg_path, seq, col, sub=False, label=''):
    """把"第 seq 项、前缀 label 的值在第 col 列"写回版式配置。

    label 为空 → 老写法，按主值/副值整体覆盖（兼容旧配置）；
    label 非空 → 精确绑定到某个字段名前缀，如"防杆吹出"取第 9 列，
                 这样一行两个同值时也能各自纠正，不会张冠李戴。

    另外：如果这个前缀压根还没有"写到哪一列"的映射（未知版式常见），
    顺手按字段名猜一列补上——否则框了半天，值仍然没地方去，看起来就像
    "修正没生效"。

    返回 (是否新增修正, 当前修正条数, 自动补上的目标列或 '')。
    """
    with open(cfg_path, 'r', encoding='utf-8') as f:
        cfg = json.load(f)
    parse = cfg.setdefault('parse', {})
    ovs = parse.get('overrides') or []
    new = {'seq': int(seq), 'col': int(col)}
    if sub:
        new['sub'] = True
    if label:
        new['label'] = str(label)
    key = _ov_key(new)
    hit = False
    for i, o in enumerate(ovs):
        if _ov_key(o) == key:
            ovs[i] = new
            hit = True
            break
    if not hit:
        ovs.append(new)
    parse['overrides'] = ovs

    added = ''
    if label:
        flds = cfg.get('fields') or []
        # 标签可能是"截断过的配置名"，也可能是原文字段名，用相似度认，别死比字符串
        mapped = [f for f in flds if f.get('to') and
                  (str(f.get('label') or '') == str(label) or
                   label_similarity(f.get('label'), label) >= 0.9)]
        if not mapped:
            used = {f.get('to') for f in flds if f.get('to')}
            to, tf = _guess_to_for_label(label, used)
            if to:
                item = {'seq': int(seq), 'to': to, 'label': str(label)}
                pat = _field_pat_from_name(label)
                if pat:
                    item['pat'] = pat
                if tf:
                    item['transform'] = tf
                if sub:
                    item['sub'] = 1
                flds.append(item)
                cfg['fields'] = flds
                added = to
    with open(cfg_path, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    return (not hit), len(ovs), added


def remove_override(cfg_path, seq, sub=False, label=''):
    """删掉一条人工修正（label 参与身份，能精确删"按前缀"那条）。"""
    with open(cfg_path, 'r', encoding='utf-8') as f:
        cfg = json.load(f)
    parse = cfg.get('parse') or {}
    ovs = parse.get('overrides') or []
    key = (int(seq), bool(sub), str(label or ''))
    keep = [o for o in ovs if _ov_key(o) != key]
    parse['overrides'] = keep
    cfg['parse'] = parse
    with open(cfg_path, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    return len(keep)


def override_for_label(ovs, seq, label):
    """取"第 seq 项、前缀 label"的人工指定列；没有返回 None。"""
    for o in ovs or []:
        try:
            if (int(o.get('seq', -1)) == int(seq) and
                    str(o.get('label') or '') == str(label)):
                return o.get('col')
        except (TypeError, ValueError):
            continue
    return None


def _field_pat_from_name(name):
    """把一个字段名变成"抗漂移"的正则。

    取中文字串（'设计温度 Design Temperature' → 设计温度），没有中文就取
    前几个英文词。生成配置时给每项都写上 pat，下次同一族版式插行/多一列，
    靠字段名也能找回来，不再依赖写死的 seq。
    """
    s = _clean(name)
    if not s:
        return ''
    cn = re.findall(r'[\u4e00-\u9fff]{2,}', s)
    if cn:
        return r'.*?'.join(re.escape(x) for x in cn)
    parts = re.findall(r'[A-Za-z][A-Za-z0-9.\-]*', s)
    if parts:
        return r'\s*'.join(re.escape(x) for x in parts[:4])
    return re.escape(s)


def _save_auto_config(pdf_path, struct, valves):
    """把自动识别到的结构存成新配置，下次直接复用。

    和旧版的三点区别：
      ① 每个 fields 项都带 pat（字段名正则）：同族版式插了一行、换了一列，
         下次靠名字也能认出来；
      ② col_map 补 seq_r/name_r/val_r：左右并排版式（一套表两套序号）才算
         真的"认下来了"，否则附件区永远解析不到；
      ③ 文件已存在时不再直接 return，而是刷新 fields，**保留**人工 overrides
         和手工补过的目标列——以前一刷新就把人工修正冲掉/或者干脆不更新。
    """
    try:
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        slug = re.sub(r'[^\w\u4e00-\u9fff]+', '_', base)[:40].strip('_')
        path = os.path.join(MAP_DIR, f'自动_{slug}.json')
        sample = valves[0]
        ncols = max(struct['ncols']) if struct['ncols'] else 0
        # 注意：seq_col 合法值可以是 0（第一列）。不能用 `or 1`，
        # 否则 0 会被当成假值写成 1，整份配置的 seq/name 列全部错位，
        # 核对窗口框选时取不到数字序号，弹窗永远不出现。
        seq_cols = struct.get('seq_cols') or []
        seq_col = struct.get('seq_col')
        if seq_col is None:
            seq_col = seq_cols[0] if seq_cols else 1

        def _mode_col(vals):
            c = Counter(int(v) for v in vals if v is not None)
            return c.most_common(1)[0][0] if c else None

        left_main, left_sub, right_main, right_sub = [], [], [], []
        for v in valves:
            for f in (v.get('fields') or {}).values():
                sec = int(f.get('sec') or 0)
                if f.get('col') is not None:
                    (right_main if sec else left_main).append(f['col'])
                if f.get('sub_col') is not None:
                    (right_sub if sec else left_sub).append(f['sub_col'])
        col_map = {'seq': seq_col, 'name': seq_col + 1}
        if len(seq_cols) > 1:
            col_map['seq_r'] = seq_cols[1]
            col_map['name_r'] = seq_cols[1] + 1
        cm_left, cm_left_sub = _mode_col(left_main), _mode_col(left_sub)
        cm_right, cm_right_sub = _mode_col(right_main), _mode_col(right_sub)
        if cm_left is not None:
            col_map['val'] = cm_left
        if cm_left_sub is not None:
            col_map['val2'] = cm_left_sub
        if cm_right is not None:
            col_map['val_r'] = cm_right
        if cm_right_sub is not None:
            col_map['val2_r'] = cm_right_sub
        # 一页两台阀（左右各一套值列）：让配置记住"这台页有两台"
        two_valves = any(v.get('two_valves') for v in valves)
        vpp = 2 if two_valves else struct['valves_per_page']

        # 旧配置里人工调过的东西要留住：overrides 是核对窗口一条条框出来的
        old = {}
        if os.path.isfile(path):
            try:
                old = _load_config(path)
            except Exception:
                old = {}

        items = fields_from_keywords(sample)
        for it in items:
            f = sample['fields'].get(it['seq']) or {}
            pat = _field_pat_from_name(f.get('name'))
            if pat:
                it['pat'] = pat
        # 关键词生成不出来的目标列（人工补过 fields 的）追加保留，别丢
        have = {it.get('to') for it in items if it.get('to')}
        for it in (old.get('fields') or []):
            if it.get('to') and it.get('to') not in have:
                items.append(it)
                have.add(it['to'])

        vinfo = detect_valve_type(struct.get('text', ''))
        vts = [t for t, _ in vinfo.get('all') or []]
        # 特征词：取这份表里出现过的、够长的字段名，供老式打分参考
        text = struct.get('text') or ''
        words = []
        for it in items:
            lab = _clean(it.get('label') or '')
            if len(lab) >= 6 and lab in text and lab not in words:
                words.append(lab)
            if len(words) >= 8:
                break
        cfg = {
            'id': f'auto_{slug}',
            'name': f'自动识别：{base[:40]}',
            'auto': True,
            'valve_type': vts,
            'note': '由引擎按字段名关键词自动生成。已给每项写 pat，'
                    '同族版式插行/换列也能认出来；核对窗口的人工修正会写在这里。',
            'detect': {'must_any': [], 'score_words': words, 'min_score': 0},
            'struct': {'ncols': sorted(struct['ncols']),
                       'valves_per_page': vpp,
                       'seq_col': seq_col},
            'parse': {'parse_mode': 'generic',
                      'valves_per_page': vpp,
                      'unit_auto': True,
                      'col_map': col_map},
            'fields': items,
            'defaults': (old.get('defaults') or {'F': 1}),
        }
        for k in ('remark', 'value_transform', 'extra_columns'):
            if old.get(k):
                cfg[k] = old[k]
        ovs = ((old.get('parse') or {}).get('overrides') or [])
        if ovs:
            cfg['parse']['overrides'] = ovs
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return path
    except Exception:
        return None


def ensure_config_file(pdf_path, cfg=None):
    """保证这份 PDF 有一份**可写的**版式配置，返回它的路径。

    核对窗口要靠它保存人工修正。以前窗口在"没配置"时直接拒绝保存，
    用户框了半天等于白框；现在现场解析一份存进 mappings/ 再返回。
    """
    if cfg is None:
        cfg = {}
    try:
        valves, struct = parse_pdf_generic(pdf_path, cfg=cfg or None)
    except Exception:
        return None
    if not valves:
        return None
    return _save_auto_config(pdf_path, struct, valves)


# ============================== 提取表解析（双阀并排） ==============================

def _find_main_cols(ws, ncol):
    """从"序号"行探测左右两阀的主值列"""
    for r in range(1, ws.max_row + 1):
        b = _clean(ws.cell(r, 2).value)
        c = _clean(ws.cell(r, 3).value)
        if b == '1' and ('序号' in c or 'No.' in c):
            nums = []
            for cc in range(4, ncol + 1):
                v = _clean(ws.cell(r, cc).value)
                if re.fullmatch(r'\d+', v):
                    nums.append(cc)
            if len(nums) >= 2:
                return nums[0], nums[1]
    return None, None


def _find_sub_cols(ws, ncol, left_c, right_c):
    """从"双字段行"探测左右两阀的副值列"""
    for r in range(1, ws.max_row + 1):
        if not _clean(ws.cell(r, 4).value):
            continue
        valcols = [cc for cc in range(5, ncol + 1) if _clean(ws.cell(r, cc).value)]
        if len(valcols) >= 3 and left_c in valcols and right_c in valcols:
            lc = [c for c in valcols if c < right_c]
            rc = [c for c in valcols if c > right_c]
            lsub = max(lc) if lc else None
            rsub = min(rc) if rc else None
            if lsub and rsub:
                return lsub, rsub
    return None, None


def parse_extract_sheet(ws):
    """解析单个 sheet，返回 [{'seq','values':[左,右],'sub_values':[左,右]}]"""
    ncol = ws.max_column
    left_c, right_c = _find_main_cols(ws, ncol)
    if not left_c or not right_c:
        return []
    lsub, rsub = _find_sub_cols(ws, ncol, left_c, right_c)
    rows = []
    for r in range(1, ws.max_row + 1):
        b = _clean(ws.cell(r, 2).value)
        if not re.fullmatch(r'\d+', b):
            continue
        rows.append({
            'seq': int(b),
            'values': [_clean(ws.cell(r, left_c).value), _clean(ws.cell(r, right_c).value)],
            'sub_values': [_clean(ws.cell(r, lsub).value) if lsub else '',
                           _clean(ws.cell(r, rsub).value) if rsub else ''],
        })
    return rows


def _locate_pdf(extract_path):
    """根据提取表路径找同名原始 PDF（去掉 _提取 后缀）"""
    base = os.path.splitext(os.path.basename(extract_path))[0].replace('_提取', '').rstrip()
    d = os.path.dirname(os.path.abspath(extract_path))
    if not os.path.isdir(d):
        return None
    for f in os.listdir(d):
        if f.lower().endswith('.pdf') and f.rstrip() == base + '.pdf':
            return os.path.join(d, f)
    norm = re.sub(r'\s+', '', base)
    for f in os.listdir(d):
        if f.lower().endswith('.pdf') and re.sub(r'\s+', '', f[:-4]) == norm:
            return os.path.join(d, f)
    return None


def _parse_extract_xlsx(extract_path):
    wb = load_workbook(extract_path, read_only=True, data_only=True)
    valves = []
    for ws in wb.worksheets:
        rows = parse_extract_sheet(ws)
        if not rows:
            continue
        for vi in range(2):
            fields = {}
            for row in rows:
                v = row['values'][vi]
                sv = row['sub_values'][vi]
                if not v and sv:
                    v, sv = sv, ''
                fields[row['seq']] = {'seq': row['seq'], 'name': '', 'name2': '',
                                      'main': v, 'sub': sv, 'raw': []}
            tag = fields.get(2, {}).get('main', '')
            if not tag:
                continue
            valves.append({'page': ws.title, 'index': vi, 'tag': tag, 'fields': fields})
    wb.close()
    return valves


def _pdf_to_temp_extract(pdf_path):
    """走"提取表"兜底路线时：先用提取引擎导出一份临时提取表，返回路径"""
    import tempfile
    from pdf_table_master import extract_pdf, write_excel
    tmpdir = tempfile.mkdtemp(prefix='pdf2quote_')
    out = os.path.join(tmpdir, os.path.splitext(os.path.basename(pdf_path))[0] + '_提取.xlsx')
    res = extract_pdf(pdf_path, layout='auto', quiet=True)
    write_excel(res, out)
    return out


# ============================== 取阀门列表 ==============================

def extract_valves(source_path, cfg, source_pdf=None):
    """按配置取出所有阀门。source 可以是 PDF 或提取表 xlsx。"""
    mode = (cfg.get('parse') or {}).get('parse_mode', 'pdf_inline')
    is_pdf = source_path.lower().endswith('.pdf')

    if mode == 'generic':
        pdf = source_path if is_pdf else (source_pdf or _locate_pdf(source_path))
        if not pdf or not os.path.isfile(pdf):
            raise RuntimeError('该版式需要原始 PDF，但没找到')
        valves, _ = parse_pdf_generic(pdf, cfg=cfg)
        if not valves:
            raise RuntimeError('按通用方式没解析到阀门，PDF 可能没有文字层。')
        _attach_notes(pdf, cfg, valves)
        return valves

    if mode in ('pdf_inline', 'auto'):
        pdf = source_path if is_pdf else (source_pdf or _locate_pdf(source_path))
        if pdf and os.path.isfile(pdf):
            valves = parse_pdf(pdf, cfg)
            if valves:
                return valves
        if mode == 'pdf_inline' and not pdf:
            raise RuntimeError('该版式需要原始 PDF，但没找到（提取表旁边应有同名 PDF）')

    # col 模式：读提取表
    if is_pdf:
        extract_path = _pdf_to_temp_extract(source_path)
    else:
        extract_path = source_path
    valves = _parse_extract_xlsx(extract_path)
    if not valves:
        raise RuntimeError('提取表里没有解析到阀门（位号为空），请检查版式配置。')
    return valves


# ============================== 值转换 ==============================

def _transform(value, rule):
    if not value or not rule:
        return value
    if rule == 'failure_pos':
        v = value.strip().upper()
        if v in ('CLOSE', 'FC', 'FAIL CLOSE') or v.startswith('FC'):
            return '关'
        if v in ('OPEN', 'FO', 'FAIL OPEN') or v.startswith('FO'):
            return '开'
        if v in ('LAST', 'FL', 'FAIL LAST') or v.startswith('FL'):
            return '保位'
        return value
    if rule == 'pressure_class':
        v = value.strip()
        m = re.search(r'(?:CL|CLASS|ANSI|ASME)\s*-?\s*(\d+)', v, re.I) or \
            re.search(r'(\d+)\s*#', v)
        if m:
            return 'CL' + m.group(1)
        m = re.search(r'PN\s*(\d+)', v, re.I)
        if m:
            return 'PN' + m.group(1)
        return v
    if rule == 'min_temp':
        # 设计温度那一行常写成 '@Max 90 / @Min -45'，最低耐温取 -45；
        # 有的写成 '/'（不要求最低温度），那就没有值。
        v = str(value or '').strip()
        m = re.search(r'-\s*\d+(?:\.\d+)?', v)
        return m.group(0).replace(' ', '') if m else ''
    return value


# ============================== 写报价表 ==============================

def _copy_row_style(ws, src_row, dst_row, max_col):
    from copy import copy
    for c in range(1, max_col + 1):
        s = ws.cell(src_row, c)
        d = ws.cell(dst_row, c)
        if s.has_style:
            d._style = copy(s._style)
        if s.alignment:
            d.alignment = copy(s.alignment)


def _copy_cell_style(src, dst):
    from copy import copy
    if src.has_style:
        dst._style = copy(src._style)
    if src.number_format:
        dst.number_format = src.number_format
    if src.alignment:
        dst.alignment = copy(src.alignment)


def _insert_cols_with_merges(ws, at_col, n):
    """在 at_col 后面插入 n 列，并把合并区域、列宽一起搬过去。

    openpyxl 的 insert_cols 只搬单元格，不会动 merged_cells 和 column_dimensions，
    模板里 A1:A2、EA1:ED1、A15:EI15 这些合并全都要手动重建，否则表头会散架。
    """
    old_max = ws.max_column
    merges = [(m.min_row, m.min_col, m.max_row, m.max_col)
              for m in ws.merged_cells.ranges]
    widths = {}
    for c in range(1, old_max + 1):
        L = get_column_letter(c)
        d = ws.column_dimensions.get(L)
        if d is not None and d.width:
            widths[c] = d.width

    for m in list(ws.merged_cells.ranges):
        try:
            ws.unmerge_cells(str(m))
        except Exception:
            pass

    ws.insert_cols(at_col + 1, amount=n)

    for r1, c1, r2, c2 in merges:
        c1n = c1 + n if c1 > at_col else c1
        c2n = c2 + n if c2 > at_col else c2
        try:
            ws.merge_cells(start_row=r1, start_column=c1n,
                           end_row=r2, end_column=c2n)
        except Exception:
            pass
    for c, w in widths.items():
        cn = c + n if c > at_col else c
        try:
            ws.column_dimensions[get_column_letter(cn)].width = w
        except Exception:
            pass


def _add_extra_columns(ws, valves, cfg, start_row, total_row):
    """在模板最右侧追加"分组表头 + 子列"的列。

    对应数据表里 @Max./@Nor. 这种一行两组值的情况：报价表模板往往只有一列，
    塞不下两个值，于是在右边新开一组，标题横跨、下面拆 Max./Nor. 两列。

    配置（版式 JSON 里）：
      "extra_columns": [
        {"title": "入口压力",
         "subs": [
           {"header": "Max.", "seq": 12, "src_col": 9,  "unit_col": 18},
           {"header": "Nor.", "seq": 12, "src_col": 12, "unit_col": 18}
         ]},
        {"title": "阀盖类型", "seq": 31}
      ]

    - 有 subs：第 1 行分组标题横跨，第 2 行写各子列名
    - 没有 subs：单列，标题纵向合并占第 1-2 行（模板里 A1:A2 那种）
    - src_col：直接从这个字段的 raw 第几列取值（不走 main/sub 那套优先级）；
      不写就用 main/sub
    """
    groups = cfg.get('extra_columns') or []
    if not groups:
        return 0

    added = 0
    for g in groups:
        # after 指定的位置在 _prepare_extra_columns 里已经插好；
        # 没指定的就接在最后一列后面（不动任何已有列字母）
        at = g.get('_at_col')
        c0 = at + 1 if at else ws.max_column + 1
        subs = g.get('subs') or []
        single = not subs
        if single:
            subs = [{'seq': g.get('seq'), 'src_col': g.get('src_col'),
                     'unit_col': g.get('unit_col')}]
        n = len(subs)
        c_end = c0 + n - 1
        L0, L1 = get_column_letter(c0), get_column_letter(c_end)

        # 第 1 行：分组标题（插在已有分组表头里面的，第 1 行是合并区不能写，
        # 只写第 2 行；否则 openpyxl 会报 MergedCell 只读）
        sub_only = bool(g.get('_sub_only'))
        if sub_only:
            for i, s in enumerate(subs):
                ws.cell(2, c0 + i,
                        (g.get('title') if i == 0 else (s.get('header') or '')))
        else:
            ws.cell(1, c0, g.get('title') or '')
            try:
                if single:
                    ws.merge_cells(f'{L0}1:{L0}2')       # 单列：纵向占两行
                elif n > 1:
                    ws.merge_cells(f'{L0}1:{L1}1')       # 分组：横向合并
            except Exception:
                pass
            # 第 2 行：子列名（单列时已被合并占掉，不写）
            if not single:
                for i, s in enumerate(subs):
                    ws.cell(2, c0 + i, s.get('header') or '')

        # 数据行
        placeholder = bool(g.get('placeholder'))
        for vi, valve in enumerate(valves):
            r = start_row + vi
            if r >= total_row:
                break
            if placeholder:
                continue          # 占位列：只要表头，不写值
            for i, s in enumerate(subs):
                # 附加列也支持按字段名定位：配了 label/pat 就先按名字找这一行的
                # 第几项，写死 seq 只在名字完全对不上时兜底（和 fields 同一套规则）。
                _, f = resolve_field(valve, s)
                if not f:
                    continue
                raw = f.get('raw') or []
                sc = s.get('src_col')
                if sc is not None:
                    # 指定了列就直取（@Max./@Nor. 这种一行两组值的场景）
                    v = _cell(raw, sc)
                else:
                    # 没指定就用引擎已经判断好的主值/副值
                    v = str(f.get('main') or f.get('sub') or '')
                if not v or _is_unit_only(v) or _looks_label(v):
                    continue
                ucol = s.get('unit_col')
                if ucol is not None and _NUM_LIKE.match(str(v)):
                    u = _cell(raw, ucol)
                    if _unit_ok(u):
                        v = f'{v} {u}'
                ws.cell(r, c0 + i, v)

        # 样式：表头抄左边一列（模板最后一列），数据行抄本行已有的单元格
        for i in range(n):
            col = c0 + i
            if c0 > 1:
                _copy_cell_style(ws.cell(1, c0 - 1), ws.cell(1, col))
                _copy_cell_style(ws.cell(2, c0 - 1), ws.cell(2, col))
            for vi in range(min(len(valves), max(0, total_row - start_row))):
                r = start_row + vi
                _copy_cell_style(ws.cell(r, c0 - 1), ws.cell(r, col))
            try:
                ws.column_dimensions[get_column_letter(col)].width = 11
            except Exception:
                pass

        added += n
    return added


def _shrink_empty_columns(ws, start_row, total_row, min_w=8):
    """数据行整列为空的，把列宽收窄，让报价表不至于被一堆空列撑得很宽。

    不动数据行区域的横向合并——改宽度会把合并区挤变形。
    表头（第 1 行）的分组合并不管：那只是分类标题，缩窄数据列不会破坏数据，
    顶多标题显示被截断，而模板里这类分组几乎横跨了全部列，一律跳过等于不缩。
    """
    merged_cols = set()
    for m in ws.merged_cells.ranges:
        # 只看数据行范围内的横向合并。总计行下方的备注行是 'A16:EI16' 这种
        # 横跨全表的大合并，算进来等于所有列都跳过，缩窄就完全不生效了。
        if m.max_col > m.min_col and start_row <= m.min_row and m.max_row < total_row:
            merged_cols.update(range(m.min_col, m.max_col + 1))

    n = 0
    for c in range(1, ws.max_column + 1):
        if c in merged_cols:
            continue
        has = False
        for r in range(start_row, total_row):
            v = ws.cell(r, c).value
            if v not in (None, ''):
                has = True
                break
        if has:
            continue
        L = get_column_letter(c)
        try:
            # 整列没值：表头字号改成 10 号宋体，列宽尽量收到最窄
            for r in (1, 2):
                cell = ws.cell(r, c)
                cell.font = Font(name='宋体', size=10)
            ws.column_dimensions[L].width = min(min_w, 6) if min_w else 6
            n += 1
        except Exception:
            pass
    return n


def _remark_parts(valve, remark_cfg):
    """拼备注列（默认 Q 列）的文字，返回 ['流量：13380 kg/h', ...]。

    两条路并行，同一项只出现一次：
      ① include_labels —— 按**字段名正则**匹配，不怕 seq 漂移。
         各家数据表的序号编号不统一（这份 30 项、那份 29 项、球阀版 44 项），
         但"Flow Rate 流量"这几个字不会变，所以按名字找最稳。
      ② include_seqs —— 按序号取，保留给已知版式的固化清单，补 ① 没覆盖到的项。

    remark_cfg 支持：
      to            目标列，默认 'Q'
      unit_col      单位列（整段共用一个时写这里）
      triplet       {"nor":12,"max":9,"min":16}
                    配了 nor 就用它当主值（正常工况），Max/Min 放括号里；
                    没配 nor 就用 col_map 取到的值当主值
      include_labels [{"pat":"flow\\s*rate|流量","label":"流量","unit_col":18}]
      include_seqs   [11, 14, 29, ...]
      labels         {"11":"流量"} 给 include_seqs 用
    """
    ucol = remark_cfg.get('unit_col')
    trip = remark_cfg.get('triplet') or {}
    labels = remark_cfg.get('labels') or {}

    def _fmt(f, label, ucol_i, which='main', any_side=False, src_col=None):
        raw = f.get('raw') or []
        # 人工框选按前缀改过的话，以它为准（如"防杆吹出"明确取第 9 列）
        lv = (f.get('label_vals') or {}).get(label)
        if lv:
            return f'{label}：{lv}'
        # src_col：点名取这一行的第几列（@Max/@Nor/@Min 里的 @Nor 就走这条，
        # 否则 col_map 取到的永远是 @Max，备注里就变成"流量：110187"）
        if src_col is not None:
            if isinstance(src_col, (list, tuple)):
                sep = ' '
                v0 = sep.join(x for x in (_cell(raw, int(i)) for i in src_col)
                              if x and not _is_unit_only(x))
            else:
                v0 = _cell(raw, int(src_col))
            if _is_unit_only(v0) or _looks_label(v0):
                v0 = ''
        # which 决定取哪一格：字段名匹配到 name 就取主值，匹配到 name2 就取副值。
        # 这一条很关键——同一行里 name2 常常是另一项（'电磁阀数量 / 配置方式'），
        # 一律取主值会把数量 1 当成"配置方式"写进备注。
        elif which == 'sub':
            v0 = str(f.get('sub') or '')
        else:
            v0 = str(f.get('main') or '')
            # 有的标签本身就是"两项合起来"的意思（'选择Cv/特性'、'输出扭矩/气源管线'），
            # 左值空、右值有值时要拿右值——这种在配置里写 any_side:true。
            # 否则只有 name2 本身是空的时候才让副值顶上。
            if not v0:
                n2 = str(f.get('name2') or '')
                if any_side or not n2 or not _looks_label(n2):
                    v0 = str(f.get('sub') or '')
        # 三列版式里 col_map 的 val 往往是 @Max.（最大值），而报价/选型要看
        # 正常工况 @Nor.。配了 triplet.nor 就以它为准，Max/Min 退到括号里。
        if trip and trip.get('nor') is not None and src_col is None:
            vn = _cell(raw, trip['nor'])
            if vn and not _is_unit_only(vn) and not _looks_label(vn):
                v0 = vn
        if not v0:
            return None
        v = v0
        # 单位单独占一列（12 列版 col6=Unit），拼上才算完整：
        # '13380' → '13380 kg/h'，'≤85' → '≤85 dBA'
        if ucol_i is not None and _NUM_LIKE.match(v):
            u = _unit_text(_cell(raw, ucol_i))
            if _unit_ok(u) and not re.search(r'\d', str(u)):
                v = f'{v} {u}'
        elif f.get('unit') and _NUM_OR_RANGE.match(str(v0)):
            # 列位不固定的版式：解析层已经从"值右边紧邻"抓到单位
            v = f'{v} {_unit_text(f["unit"])}'
        # 有的表把工艺参数拆成 Max./Norm./Min. 三列，主值取的是 Norm.
        # 三者不一致时把另外两个也带上，否则只看到中间值会漏掉上限。
        # 只对数值型生效：'Yes（带过滤减压）' 这种配置型字段的 Max/Min 列
        # 往往是别的含义（供气压力 0.4MPa-g），拼进去就成噪音了。
        if trip and _NUM_LIKE.match(v0):
            # Max/Min 各自独立判断：很多表只填了 @Max. 和 @Nor.，@Min. 空着，
            # 用"两个都非空才追加"会连 Max 一起丢掉。
            extra = []
            vmax = _cell(raw, trip.get('max'))
            vmin = _cell(raw, trip.get('min'))
            if vmax and vmax != v0 and not _is_unit_only(vmax):
                extra.append(f'Max {vmax}')
            if vmin and vmin != v0 and not _is_unit_only(vmin):
                extra.append(f'Min {vmin}')
            if extra:
                v = f'{v}（{"；".join(extra)}）'
        return f'{label}：{v}'

    parts, seen = [], set()

    specs = remark_cfg.get('include_labels') or []
    # 一旦配了字段名规则，就**完全不用** include_seqs。
    # 原因：seq 会随 PDF 漂移（这份 30 项、那份 29 项、球阀版 44 项），
    # 按 seq 取会取到隔壁那一行，出现"流量：1.7 MPa-g"这种张冠李戴。
    # 字段名（Flow Rate 流量）不会漂，所以有它就以它为准，不留 seq 兜底。
    seq_list = [] if specs else (remark_cfg.get('include_seqs') or [])

    # ① 字段名驱动
    for spec in specs:
        pat_s = spec.get('pat') or ''
        if not pat_s:
            continue
        try:
            pat = re.compile(pat_s, re.I)
        except re.error:
            continue
        label = spec.get('label') or pat_s
        if label in seen:
            continue
        for seq, f in sorted(valve['fields'].items(), key=lambda kv: int(kv[0])):
            seq = int(seq)
            nm = f.get('name') or ''
            nm2 = f.get('name2') or ''
            if nm and pat.search(nm):
                which = 'main'
            elif nm2 and pat.search(nm2):
                which = 'sub'
            else:
                continue
            s = _fmt(f, label, spec.get('unit_col', ucol), which,
                     any_side=bool(spec.get('any_side')),
                     src_col=spec.get('src_col'))
            if not s:
                continue          # 匹配到了但这项这台阀是空的，继续找下一行同名项
            parts.append(s)
            seen.add(label)
            break

    # ② 序号驱动（仅当这个版式没配字段名规则时用）
    for seq in seq_list:
        seq = int(seq)
        label = labels.get(str(seq), str(seq))
        if label in seen:
            continue
        f = valve['fields'].get(seq)
        # seq 只对"当初写下这份配置的那套版式"成立：一份配置合并了多套版式、
        # 或设计院改版插了一行时，第 seq 项的字段名可能已经对不上标签。
        # 对不上就按标签另找一行；找不到宁可跳过，也不要把隔壁项的值当成本项。
        if f is not None and label and not _label_owns_row(
                label, f.get('name'), f.get('name2')):
            f = _find_row_by_label(valve, label)
        if not f:
            continue
        s = _fmt(f, label, ucol)
        if s:
            parts.append(s)
            seen.add(label)

    return parts


def _prepare_columns(ws, cfg):
    """准备"新加的列"：备注原文/翻译备注（插在备注列后面）+ extra_columns。

    返回 (列偏移函数, 备注两列的列号[或空])。

    多组插入必须**从右往左**一次规划完：先在左边插一列，右边所有列都要右移，
    先前记下的列号就全废了（实测会把"壁厚/内径"的值写到隔壁列去）。
    所以先按模板列号排好、右边先插，再统一按"某模板列左边插了多少列"算偏移。
    """
    fn = (cfg.get('parse') or {}).get('footnote') or {}
    plan, order = [], 0

    def _inside_group_header(at0):
        """插在"分组表头"（第 1 行横向合并）中间时，表头格会被合并吃掉。
        BK1:BX1 这种分组里再插列写表头会撞 MergedCell，这种就改为追加到最右边。"""
        for m in ws.merged_cells.ranges:
            if m.min_row == 1 and m.max_row == 1 and m.min_col < at0 + 1 <= m.max_col:
                return True
        return False

    if fn:
        try:
            at0 = column_index_from_string(fn.get('after', 'Q'))
        except Exception:
            at0 = None
        if at0:
            plan.append({'at0': at0, 'n': 2, 'kind': 'fn', 'payload': fn,
                         'order': order})
            order += 1
    for g in cfg.get('extra_columns') or []:
        after = g.get('after')
        if not after:
            continue
        try:
            at0 = column_index_from_string(after)
        except Exception:
            continue
        if _inside_group_header(at0):
            # 插在分组表头（BK1:BX1 这种）里面：列要插，但表头只能写第 2 行，
            # 第 1 行那一格是合并区（写进去会撞 MergedCell）。
            plan.append({'at0': at0, 'n': max(1, len(g.get('subs') or [])),
                         'kind': 'extra', 'payload': g, 'order': order,
                         'sub_only': True})
            order += 1
            continue
        plan.append({'at0': at0, 'n': max(1, len(g.get('subs') or [])),
                     'kind': 'extra', 'payload': g, 'order': order})
        order += 1

    # 从右往左插：右边的先插，左边再插时不会挪动已经插好的右边列。
    # 因此每个插入点就是它自己的模板列号（右边的插入不会改变左边列的列号）。
    # 同一列的多个分组按倒序插，最后正好排成配置里的先后顺序。
    for p in sorted(plan, key=lambda x: (-x['at0'], -x['order'])):
        p['at'] = p['at0']
        _insert_cols_with_merges(ws, p['at'], p['n'])

    def shift(c):
        """模板里的列号 → 插完新列后的实际列号。"""
        return c + sum(q['n'] for q in plan if q['at0'] < c)

    fn_cols = []
    for p in plan:
        if p['kind'] == 'fn':
            fn_cols = [shift(p['at0']) + 1, shift(p['at0']) + 2]
            titles = (p['payload'].get('original_title') or '备注原文',
                      p['payload'].get('translated_title') or '翻译备注')
            for c, title in zip(fn_cols, titles):
                ws.cell(1, c, title)
                try:
                    ws.merge_cells(start_row=1, start_column=c, end_row=2, end_column=c)
                except Exception:
                    pass
                _copy_cell_style(ws.cell(1, shift(p['at0'])), ws.cell(1, c))
                try:
                    ws.column_dimensions[get_column_letter(c)].width = 34
                except Exception:
                    pass
        else:
            # 交给 _add_extra_columns：记下"插入点"的最终列号（它会在 +1 处开始写）。
            # 同一个 after 有多组时要依次往右排，否则两组会抢同一列。
            left_same = sum(q['n'] for q in plan
                            if q['kind'] == 'extra' and q['at0'] == p['at0']
                            and q['order'] < p['order'])
            p['payload']['_at_col'] = shift(p['at0']) + left_same
            if p.get('sub_only'):
                p['payload']['_sub_only'] = True
    return shift, fn_cols


def _fill_sheet(ws, valves, cfg, start_row):
    """把阀门列表写进工作表，返回统计"""
    # 带 after 的附加列要先插进去，并拿到"原列字母 -> 实际列号"的偏移函数。
    # 必须在取 max_col 之前做，否则后续复制样式会漏掉新插入的列。
    col_shift, fn_cols = _prepare_columns(ws, cfg)
    max_col = ws.max_column

    total_row = None
    for r in range(1, ws.max_row + 1):
        if _clean(ws.cell(r, 1).value) == '总计':
            total_row = r
            break
    if total_row is None:
        total_row = start_row + 7

    need_rows = len(valves)
    have_rows = total_row - start_row

    # —— 快照"总计行及其下方"，避免 insert_rows 破坏合并 ——
    tail_start = total_row
    tail_snapshot = {}
    for r in range(tail_start, ws.max_row + 1):
        for c in range(1, max_col + 1):
            v = ws.cell(r, c).value
            if v is not None:
                tail_snapshot[(r, c)] = v
    tail_merges = [(m.min_row, m.min_col, m.max_row, m.max_col)
                   for m in ws.merged_cells.ranges if m.min_row >= tail_start]
    for mr, mc1, Mr, mc2 in tail_merges:
        try:
            ws.unmerge_cells(f'{get_column_letter(mc1)}{mr}:{get_column_letter(mc2)}{Mr}')
        except Exception:
            pass

    if need_rows > have_rows:
        ins = need_rows - have_rows
        ws.insert_rows(total_row, amount=ins)
        for r in range(start_row + have_rows, start_row + need_rows):
            _copy_row_style(ws, start_row, r, max_col)
        total_row += ins
    elif need_rows < have_rows:
        # 模板自带若干行示例。阀少的时候，多出来的空行要真删掉——
        # 只清空不删的话，报价表尾部会挂一串没有序号的空行，
        # 看上去就像"这些数据没提取出来"。
        cut = have_rows - need_rows
        cut_at = start_row + need_rows
        for mr, mc1, Mr, mc2 in [(m.min_row, m.min_col, m.max_row, m.max_col)
                                 for m in ws.merged_cells.ranges
                                 if m.min_row >= cut_at and m.max_row < total_row]:
            try:
                ws.unmerge_cells(f'{get_column_letter(mc1)}{mr}:'
                                 f'{get_column_letter(mc2)}{Mr}')
            except Exception:
                pass
        ws.delete_rows(cut_at, amount=cut)
        total_row -= cut

    shift = total_row - tail_start
    for (r, c), v in tail_snapshot.items():
        ws.cell(r + shift, c).value = v
    for mr, mc1, Mr, mc2 in tail_merges:
        try:
            ws.merge_cells(f'{get_column_letter(mc1)}{mr+shift}:{get_column_letter(mc2)}{Mr+shift}')
        except Exception:
            pass

    for r in range(start_row, total_row):
        for c in range(1, max_col + 1):
            ws.cell(r, c).value = None

    ws.cell(total_row, 1, '总计')
    ws.cell(total_row, 6, f'=SUM(F{start_row}:F{total_row-1})')

    transforms = cfg.get('value_transform', {})
    defaults = cfg.get('defaults', {})
    remark_cfg = cfg.get('remark', {})
    # 单位是否自动拼：列位不固定时靠解析层"值右边紧邻的单位格"拿到
    unit_auto = bool((cfg.get('parse') or {}).get('unit_auto'))
    detail = []

    for vi, valve in enumerate(valves):
        r = start_row + vi
        for col, val in defaults.items():
            ws.cell(r, col_shift(column_index_from_string(col)), val)
        ws.cell(r, 1, vi + 1)

        for item in cfg.get('fields', []):
            # 有的版式把 位号/用途/介质 写在表格下方的汇总小表里，
            # 这些值不在 fields[seq] 里，用 "meta": "service" 直接点名取。
            if item.get('meta'):
                v = (valve.get('meta') or {}).get(item['meta'], '')
                if not v:
                    continue
                try:
                    ws.cell(r, col_shift(column_index_from_string(item['to'])), v)
                except Exception:
                    pass
                continue
            # 抗漂移：先按字段名（pat → 字段名相似度）找这一台的第几项，
            # 名字完全对不上才退回写死的 seq。改版插行、多页表后几页多一格、
            # 同一份表左右两栏各有一套序号，都靠这一步纠正。
            seq, f = resolve_field(valve, item)
            if not f:
                continue
            # 人工框选按前缀改过的值优先（fields 项的 label 就是前缀名）
            lab = item.get('label') or ''
            lv = (f.get('label_vals') or {}).get(lab) if lab else None
            sc = item.get('src_col')
            if sc is not None:
                # src_col：直接点名"取这一行的第几列"。一行里并排好几组数据
                # （@Max/@Nor/@Min、法兰 Standard/Size/Rating/Face）时，靠
                # 字段名+位置猜不出来，必须让配置能精确指定是哪一格。
                raw = f.get('raw') or []
                if isinstance(sc, (list, tuple)):
                    sep = item.get('src_sep', ' ')
                    vals = [_cell(raw, int(i)) for i in sc]
                    v = sep.join(x for x in vals
                                 if x and not _is_unit_only(x) and not _is_junk(x))
                else:
                    v = _cell(raw, int(sc))
                if _is_unit_only(v) or _looks_label(v):
                    v = ''
            else:
                v = lv or (f.get('sub', '') if int(item.get('sub', 0)) == 1
                           else f.get('main', ''))
            rule = item.get('transform') or transforms.get(str(seq))
            if rule:
                v = _transform(v, rule)
            if v and item.get('unit_col') is not None:
                raw = f.get('raw') or []
                u = _unit_text(_cell(raw, item['unit_col']))
                if _unit_ok(u) and _NUM_OR_RANGE.match(str(v)):
                    v = f'{v} {u}'
            elif v and (item.get('unit_auto') or unit_auto):
                # 单位由解析层在"值右边紧邻"抓到（各行列位不一样也能拿到）
                u = _unit_text(f.get('unit'))
                if u and _NUM_OR_RANGE.match(str(v)):
                    v = f'{v} {u}'
            if not v:
                continue
            try:
                to = item['to']
                # 有些版式同一编号既可能是"类型"也可能是"数量"，纯数字时改填到数量列
                if item.get('if_numeric') and re.fullmatch(r'\d+', str(v)):
                    to = item['if_numeric']
                col = col_shift(column_index_from_string(to))
                if item.get('merge'):
                    old = ws.cell(r, col).value
                    if old not in (None, ''):
                        olds = str(old)
                        # 用 " / " 分段比较，避免 '316SS/STL' 被拆成 316SS 和 STL
                        parts = [p.strip() for p in olds.split(' / ')]
                        v = olds if str(v) in parts else f'{olds} / {v}'
                ws.cell(r, col, v)
            except Exception:
                pass

        if remark_cfg:
            parts = _remark_parts(valve, remark_cfg)
            if parts:
                ws.cell(r, col_shift(column_index_from_string(remark_cfg.get('to', 'Q'))),
                        '；'.join(parts))

        # 每页表格下方的备注：原文 + 翻译（不同页的备注写在各自那台阀那一行）
        fn_cfg = (cfg.get('parse') or {}).get('footnote') or {}
        if fn_cols and fn_cfg:
            note = _clean(valve.get('footnote') or '')
            if note:
                ws.cell(r, fn_cols[0], note)
                ws.cell(r, fn_cols[1],
                        translate_note(note) if fn_cfg.get('translate', True) else note)

        detail.append({'row': r, 'tag': valve['tag'], 'page': valve['page'],
                       'fields': len(valve['fields'])})

    # 数据表里一行两组值（@Max./@Nor.）而模板只有一列的项，
    # 在报价表最右侧新开一组列，标题横跨、下面拆成两列。
    n_extra = _add_extra_columns(ws, valves, cfg, start_row, total_row)

    # 整列没数据的收窄列宽（配了 shrink_empty:false 可关，宽度用 shrink_width 调）
    if cfg.get('shrink_empty', True):
        _shrink_empty_columns(ws, start_row, total_row,
                              min_w=cfg.get('shrink_width', 8))

    return {'data_rows': f'{start_row}-{total_row-1}', 'total_row': total_row,
            'detail': detail, 'extra_columns': n_extra}


def _fill(source_path, template_path, out_path, config_path=None,
          start_row=3, source_pdf=None, allow_generic=True):
    """主流程：识别版式 → 解析 → 写表。"""
    if not os.path.isfile(source_path):
        raise FileNotFoundError(f'源文件不存在: {source_path}')
    if not os.path.isfile(template_path):
        raise FileNotFoundError(f'模板不存在: {template_path}')

    is_pdf = source_path.lower().endswith('.pdf')
    learned = None

    cfg_path = config_path
    scores = []
    if not cfg_path:
        cfg_path, scores = detect_layout(source_path)

    if not cfg_path:
        if not allow_generic or not is_pdf:
            raise RuntimeError('无法识别版式，且没有可用的兜底方案。候选打分: %s' % scores)
        # 未知版式：按关键词尽力填，并把结构存成新配置
        valves, struct = parse_pdf_generic(source_path)
        if not valves:
            raise RuntimeError('这个 PDF 没解析出任何阀门，可能没有文字层或表格结构特殊。')
        learned = _save_auto_config(source_path, struct, valves)
        cfg = {
            'id': 'generic', 'name': '通用关键词映射（未知版式兜底）',
            'fields': fields_from_keywords(valves[0]),
            'defaults': {'F': 1},
        }
        cfg_path = learned or '(内置兜底)'
    else:
        cfg = _load_config(cfg_path)
        valves = extract_valves(source_path, cfg, source_pdf=source_pdf)

    if not valves:
        raise RuntimeError('没有解析到任何阀门（位号为空）。')

    with open(template_path, 'rb') as fh:
        wb = load_workbook(fh, read_only=False, data_only=False)
    ws = wb.worksheets[0]
    stat = _fill_sheet(ws, valves, cfg, start_row)
    wb.save(out_path)
    wb.close()

    return {
        'config': os.path.basename(cfg_path) if os.path.sep in str(cfg_path) else str(cfg_path),
        'layout': cfg.get('name', ''),
        'output': out_path,
        'filled_rows': len(valves),
        'mapped_fields': len(cfg.get('fields', [])),
        'data_rows': stat['data_rows'],
        'total_row': stat['total_row'],
        'extra_columns': stat.get('extra_columns', 0),
        'learned_config': learned,
        'detail': stat['detail'],
    }


def fill_pdf(pdf_path, template_path, out_path, config_path=None, start_row=3):
    """PDF 直接生成报价表（主入口）"""
    return _fill(pdf_path, template_path, out_path,
                 config_path=config_path, start_row=start_row)


def fill_extract(extract_path, template_path, out_path, config_path=None,
                 start_row=3, source_pdf=None):
    """提取表 xlsx 生成报价表（兼容旧流程；能自动找到同名 PDF 时优先读 PDF）"""
    return _fill(extract_path, template_path, out_path, config_path=config_path,
                 start_row=start_row, source_pdf=source_pdf)


def template_columns(tpl=None):
    """模板列清单：[(列字母, 表头文字)]，表头取第 1、2 行拼起来。"""
    from openpyxl import load_workbook
    tpl = tpl or default_template()
    out = []
    if not tpl or not os.path.isfile(tpl):
        return out
    try:
        with open(tpl, 'rb') as fh:
            wb = load_workbook(fh, data_only=True)
        ws = wb.worksheets[0]
        for c in range(1, ws.max_column + 1):
            parts = []
            for r in (1, 2):
                v = ws.cell(r, c).value
                if v not in (None, ''):
                    parts.append(str(v).strip().replace('\n', ''))
            if parts:
                out.append((get_column_letter(c), ' '.join(parts)))
        wb.close()
    except Exception:
        pass
    return out


def upsert_field(cfg_path, seq, label, to=None, sub=0, src_col=None,
                 insert_before=None, pat=None):
    """把手动框选出来的"字段名 → 写入列"写进版式配置。

    to            目标列字母；None 表示只在配置里登记取值列（不指定写到哪）
    src_col       这个值在 PDF 里的列号（框选那一格）
    insert_before 给了就"插到这一列前面"：不动已有列，新增一列（Excel 手插列的
                  逻辑），老列往右挪；不覆盖原有内容。
    返回 (动作, 说明)：动作 ∈ {'update','add','insert','skip'}
    """
    import json
    with open(cfg_path, 'r', encoding='utf-8') as f:
        cfg = json.load(f)
    label = str(label or '').strip()
    try:
        seq = int(seq)
    except (TypeError, ValueError):
        return 'skip', 'seq 无效'
    sub = 1 if sub else 0

    # ① 插入新列
    if insert_before:
        try:
            idx = column_index_from_string(str(insert_before).strip().upper())
        except Exception:
            return 'skip', f'列字母无效: {insert_before}'
        after = get_column_letter(idx - 1) if idx > 1 else None
        extras = cfg.get('extra_columns') or []
        item = {'title': label[:24] or f'seq{seq}', 'label': label, 'seq': seq}
        if src_col is not None:
            item['src_col'] = int(src_col)
        if after:
            item['after'] = after
        if sub:
            item['sub'] = 1
        for e in extras:
            if e.get('title') == item['title'] and e.get('seq') == seq:
                e.update(item)
                break
        else:
            extras.append(item)
        cfg['extra_columns'] = extras
        with open(cfg_path, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        clear_cache()
        return 'insert', f'在 {after or "首列"} 后面插入新列「{item["title"]}」'

    # ② 更新/新增 fields 项
    fields = cfg.get('fields') or []
    hit = None
    for f in fields:
        if f.get('seq') is None or f.get('meta'):
            continue
        if int(f['seq']) != seq:
            continue
        if int(f.get('sub', 0)) != sub:
            continue
        if f.get('label') and label and label_similarity(f['label'], label) >= 0.85:
            hit = f
            break
    if hit is None and label:
        hit = next((f for f in fields if f.get('label') == label
                    and int(f.get('sub', 0)) == sub), None)
    action = 'update' if hit else 'add'
    item = hit if hit else {}
    if hit is None:
        fields.append(item)
    item['seq'] = seq
    if label:
        item['label'] = label
        item.setdefault('pat', _field_pat_from_name(label))
    if sub:
        item['sub'] = 1
    if to:
        item['to'] = str(to).strip().upper()
    if src_col is not None:
        item['src_col'] = int(src_col)
    cfg['fields'] = fields
    with open(cfg_path, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    clear_cache()
    return action, (f'第 {seq} 项「{label}」→ {item.get("to") or "(未指定列)"}'
                    f'（取 PDF 第 {src_col} 列）')


def remove_field(cfg_path, seq, sub=0, label=None):
    """删掉一条 fields 映射（人工修正列表里删"字段映射"用）。"""
    import json
    with open(cfg_path, 'r', encoding='utf-8') as f:
        cfg = json.load(f)
    keep, n = [], len(cfg.get('fields') or [])
    for f in cfg.get('fields') or []:
        same = (f.get('seq') is not None and int(f['seq']) == int(seq)
                and int(f.get('sub', 0)) == int(1 if sub else 0))
        if same and (not label or str(f.get('label') or '') == str(label)):
            continue
        keep.append(f)
    cfg['fields'] = keep
    with open(cfg_path, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    clear_cache()
    return n - len(keep)


def remove_extra_column(cfg_path, title):
    """删掉一条 extra_columns（新增列）。"""
    import json
    with open(cfg_path, 'r', encoding='utf-8') as f:
        cfg = json.load(f)
    extras = cfg.get('extra_columns') or []
    keep = [e for e in extras if str(e.get('title') or '') != str(title)]
    cfg['extra_columns'] = keep
    with open(cfg_path, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    clear_cache()
    return len(extras) - len(keep)


def default_template():
    for n in DEFAULT_TEMPLATE_NAMES:
        p = os.path.join(HERE, n)
        if os.path.isfile(p):
            return p
    return ''


if __name__ == '__main__':
    import sys
    if len(sys.argv) < 4:
        print('用法: python fill_from_extract.py <PDF或提取表> <模板> <输出> [配置.json]')
        sys.exit(1)
    cfg = sys.argv[4] if len(sys.argv) > 4 else None
    t0 = time.time()
    rep = fill_pdf(sys.argv[1], sys.argv[2], sys.argv[3], config_path=cfg)
    rep['elapsed'] = round(time.time() - t0, 2)
    print(json.dumps(rep, ensure_ascii=False, indent=2))
