"""
PDF 表格识别 · 提取 · 准确性测试 —— 一体化交互菜单（单文件自包含版）
=====================================================================
所有功能都在这一个文件里，不依赖任何兄弟模块，避免文件丢失导致崩溃。
运行后按数字选功能、按提示输路径即可，无需记命令行参数。

运行（依赖装在当前 python 里即可）：
    python pdf_table_master.py --menu
或直接双击项目根目录的「完整菜单.bat」。

菜单功能（按用途分组，新手先看 H 说明）：
    【最常用】
    1. PDF → 报价表（自动识别版式，直接出填好的报价表）
    2. 批量：整个目录的 PDF → 报价表
    【导出与核对】
    3. 映射核对窗口（左边 PDF + 红线框，右边提取内容）
    4. 把 PDF 表格导出成 Excel（提取表，供人工查看）
    5. 测试识别准不准（与人工核对版逐格比对）
    6. 生成人工核对副本
    【高级】
    7. 批量提取并检查表格结构
    8. 测试能否识别这种表格（结构识别报告）
    9. 查看已保存的版式清单（mappings）
    H. 新手说明 / 我该选哪个
    0. 退出

版式（mappings/*.json）：一种 PDF 版式对应一个 JSON，记录"第几项填哪一列"。
  新 PDF 进来先按正文认阀门类型（球阀/闸阀/截止阀/蝶阀/切断阀/调节阀），
  再按字段名命中率在同类型版式里挑最像的一份；挑不出来只取前 3 页弹红线框
  窗口给人核对，确认正确后自动出完整 Excel，不用新手手改 JSON。

作用范围：**只处理有文字层的 PDF**（设计院发来的数据表都属于这种）。
  项目不再集成 OCR：文字层缺失时直接报错，不会拿错误结果去填表。
依赖：pdfplumber pymupdf openpyxl xlrd
"""
import os
import re
import sys
import json
import time
import glob
import shutil
import tempfile
import unicodedata
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))                   # 本脚本所在目录
BASE = os.path.dirname(HERE)                                        # 上级目录
# 优先从自身目录加载同目录模块（fill_from_extract / pdf_review / semantic_map），
# 找不到时回退到上级目录——这样工具箱既自包含、又兼容原位置。
for _p in (HERE, BASE):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================ 通用工具

def clean(cell):
    if cell is None:
        return ''
    return ' '.join(str(cell).split())


def parse_page_range(spec, total):
    if not spec:
        return list(range(total))
    out = []
    for part in spec.split(','):
        part = part.strip()
        if '-' in part:
            a, b = part.split('-', 1)
            out.extend(range(int(a) - 1, min(int(b), total)))
        elif part.isdigit():
            out.append(int(part) - 1)
    return [p for p in dict.fromkeys(out) if 0 <= p < total]


def unquote(path):
    p = path.strip()
    if len(p) >= 2 and p[0] == p[-1] and p[0] in ('"', "'"):
        p = p[1:-1]
    return p.strip()


def _pick_local(patterns):
    """在本目录里找最新的一个候选文件，给"识别结果/标准答案"当默认值用。"""
    hits = []
    for pat in patterns:
        hits.extend(glob.glob(os.path.join(HERE, pat)))
    hits = [h for h in hits if os.path.isfile(h)]
    if not hits:
        return ''
    hits.sort(key=os.path.getmtime, reverse=True)
    return hits[0]


def ask_path(prompt, must_exist=True, default=None, must_file=False, exts=()):
    """输入路径。

    must_file=True → 必须是文件（不能是目录）；exts 再限定后缀。
    以前这里只判断"路径存在"，而菜单 5/6 的默认值恰好是个目录，
    于是直接回车就把目录当 xlsx 交给 openpyxl，报
    "openpyxl does not support  file format"（格式名还是空的），用户根本看不懂。
    """
    if default is None:
        default = HERE
    # 要文件时，目录型的默认值不显示、不接受——否则回车就掉坑里
    show_default = default if (default and not (must_file and os.path.isdir(default))) else ''
    while True:
        raw = input(prompt + (f' [回车用: {show_default}]' if show_default else '') + ': ')
        p = unquote(raw) if raw.strip() else (show_default or '')
        if not p:
            if must_file:
                # 没有可用默认值、用户又直接回车：不要在这里死循环，
                # 返回空串让上层退回主菜单（上层会提示先按哪几步生成文件）。
                print('  ! 这里要选一个**文件**（不是目录）。'
                      '可以右键文件→复制路径粘贴进来；直接回车 = 返回主菜单。')
                return ''
            print('  ! 路径不能为空')
            continue
        p = os.path.abspath(p) if must_file else p
        if must_exist and not os.path.exists(p):
            print(f'  ! 路径不存在: {p}  请重新输入'); continue
        if must_file and os.path.isdir(p):
            print(f'  ! 这是目录，不是文件: {p}')
            continue
        if exts and not p.lower().endswith(tuple(exts)):
            print(f'  ! 这里只认 {" / ".join(exts)} 文件，'
                  f'你给的是: {os.path.basename(p)}')
            continue
        return p


def ask_choice(prompt, options, default=None):
    print(f'\n{prompt}')
    for i, (val, desc) in enumerate(options, 1):
        print(f'   {i}. {desc}{" (默认)" if val == default else ""}')
    while True:
        raw = input(f'请输入序号(1-{len(options)})'
                    + (f'，回车默认 {default}' if default is not None else '') + ': ').strip()
        if not raw and default is not None:
            return default
        if raw.isdigit() and 1 <= int(raw) <= len(options):
            return options[int(raw) - 1][0]
        print('  ! 输入无效')


def ask_yesno(prompt, default=False):
    raw = input(f'{prompt} [{"Y/n" if default else "y/N"}]: ').strip().lower()
    if not raw:
        return default
    return raw in ('y', 'yes', '是')


def pause():
    # 没有控制台（比如被别的程序重定向了输入）时不要让整个程序崩在 input() 上
    try:
        input('\n[按回车返回主菜单]')
    except EOFError:
        pass


# ============================================================ 提取引擎

def _find_seq_columns(raw):
    col_count = max((len(r) for r in raw), default=0)
    found = []
    for ci in range(col_count):
        nums = []
        for row in raw:
            if ci < len(row):
                v = clean(row[ci])
                if re.fullmatch(r'\d+', v):
                    nums.append(int(v))
        if len(nums) >= 5:
            inc = sum(1 for a, b in zip(nums, nums[1:]) if b >= a) / max(1, len(nums) - 1)
            if inc >= 0.7:
                found.append(ci)
    return found[:1] + found[-1:] if len(found) > 2 else found


def _compact_row(row, col_slice):
    parts = []
    for ci in col_slice:
        if ci >= len(row):
            continue
        v = clean(row[ci])
        if v and (not parts or v != parts[-1]):
            parts.append(v)
    return parts


def extract_text_tables(page, layout='auto'):
    tables = page.extract_tables()
    result = []
    for raw in tables:
        if not raw:
            continue
        seq = _find_seq_columns(raw)
        ncol = max((len(r) for r in raw), default=0)
        dual = (layout == 'dual') or (layout == 'auto' and len(seq) >= 2)
        if dual and len(seq) >= 2:
            lid, rid = seq[0], seq[-1]
            rows = []
            for row in raw:
                left = _compact_row(row, range(lid, rid))
                right = _compact_row(row, range(rid, ncol))
                if left or right:
                    rows.append([' | '.join(left), ' | '.join(right)])
            result.append({'style': 'dual-kv', 'rows': rows,
                           'meta': {'seq_cols': seq}})
        else:
            rows = [[clean(c) for c in row] for row in raw]
            rows = [r for r in rows if any(r)]
            result.append({'style': 'grid', 'rows': rows, 'meta': {}})
    return result


def _fmt_sec(s):
    s = max(0, int(s))
    if s < 60:
        return f'{s}s'
    return f'{s//60}m{s%60:02d}s'


def _print_progress(idx, n, t_start, pageno, source='text', cached=False):
    """单行刷新的进度条：百分比 + 已用 + 预计剩余"""
    elapsed = time.time() - t_start
    pct = idx / n * 100 if n else 100
    eta = (elapsed / idx * (n - idx)) if idx else 0
    width = 22
    filled = int(width * idx / n) if n else width
    bar = '#' * filled + '-' * (width - filled)
    tag = '已缓存' if cached else source
    sys.stdout.write(f'\r  [{bar}] {idx}/{n} {pct:5.1f}% | 第{pageno}页 {tag} | '
                     f'已用 {_fmt_sec(elapsed)} 剩 {_fmt_sec(eta)}   ')
    sys.stdout.flush()
    if idx >= n:
        sys.stdout.write('\n')
        sys.stdout.flush()


def extract_pdf(pdf_path, layout='auto', pages=None, partial_path=None, quiet=False):
    """逐页提取（只读文字层）。

    支持进度条与断点续跑。partial_path 给出时每完成一页追加一行 JSONL。
    项目不做 OCR：遇到没有文字层的页会明确报告，而不是给出错误结果。
    """
    import pdfplumber
    t_start = time.time()
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        selected = parse_page_range(pages, total)

        # 断点：读取已完成的页
        done = {}
        if partial_path and os.path.isfile(partial_path):
            try:
                with open(partial_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if not line:
                            continue
                        try:
                            r = json.loads(line)
                            done[r['page']] = r
                        except Exception:
                            continue
            except Exception:
                done = {}

        out = []
        n = len(selected)
        for idx, i in enumerate(selected, 1):
            pageno = i + 1
            if pageno in done:                      # 断点命中，跳过重算
                out.append(done[pageno])
                if not quiet:
                    _print_progress(idx, n, t_start, pageno,
                                    done[pageno].get('source', 'text'), cached=True)
                continue
            t0 = time.time()
            tables = extract_text_tables(pdf.pages[i], layout=layout)
            r = {'page': pageno, 'source': 'text', 'tables': tables,
                 'elapsed': round(time.time() - t0, 2)}
            out.append(r)
            if partial_path:                        # 每页落盘，崩溃不丢
                try:
                    with open(partial_path, 'a', encoding='utf-8') as f:
                        f.write(json.dumps(r, ensure_ascii=False) + '\n')
                except Exception:
                    pass
            if not quiet:
                _print_progress(idx, n, t_start, pageno)

    empty = [p['page'] for p in out if not p['tables']]
    if empty and not quiet:
        print(f'  ! 第 {",".join(map(str, empty[:8]))} 页没识别到表格：'
              f'可能是扫描件/图片，本工具只处理有文字层的 PDF。')

    c = Counter(p['source'] for p in out)
    return {'file': os.path.basename(pdf_path), 'pages': out,
            'summary': {'pages': len(out), 'text_pages': c.get('text', 0),
                        'tables': sum(len(p['tables']) for p in out)}}


def write_excel(result, out_path):
    wb = Workbook(); wb.remove(wb.active)
    thin = Side(style='thin', color='C8C8C8')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hf = PatternFill('solid', fgColor='DCE6F1')
    sf = PatternFill('solid', fgColor='FFF2CC')
    for page in result['pages']:
        ws = wb.create_sheet(title=f"第{page['page']}页")
        ws.cell(1, 1, f"第{page['page']}页 · 来源={page['source']} · "
                      f"{len(page['tables'])}个表（忠实转录）")
        ws.cell(1, 1).font = Font(bold=True, size=12); ws.cell(1, 1).fill = hf
        cur = 3
        for ti, t in enumerate(page['tables']):
            ws.cell(cur, 1, f"表{ti+1} [{t['style']}] {len(t['rows'])}行")
            ws.cell(cur, 1).font = Font(bold=True, color='555555'); cur += 1
            for ri, row in enumerate(t['rows']):
                for ci, val in enumerate(row):
                    cell = ws.cell(cur+ri, ci+1, val if val else None)
                    cell.border = border
                    cell.alignment = Alignment(vertical='top', wrap_text=True)
                    if ci == 0:
                        cell.fill = sf
            cur += len(t['rows']) + 1
        ncol = max((len(r) for t in page['tables'] for r in t['rows']), default=2)
        for ci in range(1, min(ncol, 30)+1):
            ws.column_dimensions[get_column_letter(ci)].width = 60 if ncol <= 2 else 22
        ws.freeze_panes = 'A3'
    wb.save(out_path)
    return out_path


def write_json(result, out_path):
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    return out_path


# ============================================================ 测试框架

def _norm(v, ignore_space=True, ignore_width=True):
    if v is None:
        return ''
    s = str(v)
    if ignore_width:
        s = unicodedata.normalize('NFKC', s)
    if ignore_space:
        s = ''.join(s.split())
    return s


def test_structure(pdf_inputs, out_xlsx):
    paths = []
    for inp in pdf_inputs:
        if os.path.isfile(inp):
            paths.append(inp)
        elif os.path.isdir(inp):
            paths.extend(sorted(glob.glob(os.path.join(inp, '*.pdf'))))
    paths = list(dict.fromkeys(paths))

    wb = Workbook(); wb.remove(wb.active)
    thin = Side(style='thin', color='C8C8C8')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hf = PatternFill('solid', fgColor='DCE6F1')
    okf = PatternFill('solid', fgColor='E2EFDA')
    badf = PatternFill('solid', fgColor='FCE4D6')
    ws = wb.create_sheet('结构识别报告')
    heads = ['PDF', '页', '判定来源', '表格数', '表格样式', '行数', '列数', '耗时(s)', '备注']
    for ci, h in enumerate(heads, 1):
        c = ws.cell(1, ci, h); c.font = Font(bold=True); c.fill = hf; c.border = border

    r = 2
    stats = {'pages': 0, 'styles': {}}
    for p in paths:
        print(f'  结构测试: {os.path.basename(p)}', flush=True)
        try:
            result = extract_pdf(p, layout='auto')
        except Exception as e:
            ws.cell(r, 1, os.path.basename(p))
            ws.cell(r, 9, f'失败: {type(e).__name__}: {e}'); ws.cell(r, 9).fill = badf
            r += 1; continue
        for page in result['pages']:
            stats['pages'] += 1
            styles = [t['style'] for t in page['tables']]
            for s in styles:
                stats['styles'][s] = stats['styles'].get(s, 0) + 1
            nrows = sum(len(t['rows']) for t in page['tables'])
            ncols = max((max((len(row) for row in t['rows']), default=0)
                         for t in page['tables']), default=0)
            vals = [os.path.basename(p), page['page'], '文字层',
                    len(page['tables']), '; '.join(styles) or '(无)', nrows, ncols,
                    page['elapsed'], 'OK' if page['tables'] else '未识别到表格']
            for ci, v in enumerate(vals, 1):
                c = ws.cell(r, ci, v); c.border = border
                c.alignment = Alignment(vertical='center', wrap_text=True)
            ws.cell(r, 9).fill = okf if page['tables'] else badf
            r += 1

    ws2 = wb.create_sheet('汇总')
    for ci, h in enumerate(['指标', '值'], 1):
        c = ws2.cell(1, ci, h); c.font = Font(bold=True); c.fill = hf
    rows = [('测试PDF数', len(paths)), ('总页数', stats['pages'])] + \
           [(f'样式-{k}', v) for k, v in stats['styles'].items()]
    for i, (k, v) in enumerate(rows, 2):
        ws2.cell(i, 1, k); ws2.cell(i, 2, v)
    ws2.column_dimensions['A'].width = 22; ws2.column_dimensions['B'].width = 14
    for ci, w in enumerate([34, 6, 10, 8, 20, 8, 8, 9, 30], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = 'A2'
    wb.save(out_xlsx)
    return {'pdfs': len(paths), 'pages': stats['pages'],
            'styles': stats['styles'], 'report': out_xlsx}


def _load_cells(path):
    """读一个工作簿里所有非空单元格。

    用文件句柄打开：模板那种"xlsx 改名为 .xls"也能读。
    目录/非 Excel 文件给出能看懂的中文错误——openpyxl 对目录报的是
    "does not support  file format"（格式名是空的），等于没说。
    """
    if os.path.isdir(path):
        raise ValueError(f'这是目录，不是 Excel 文件：{path}')
    if not os.path.isfile(path):
        raise ValueError(f'文件不存在：{path}')
    try:
        with open(path, 'rb') as fh:
            wb = load_workbook(fh, data_only=True)
    except Exception as e:
        raise ValueError(
            f'打不开这个 Excel：{os.path.basename(path)}（{type(e).__name__}）。'
            f'确认它是 .xlsx 或本工具导出的表，且没有被 Excel 独占打开') from e
    data = {}
    for ws in wb.worksheets:
        cells = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None and str(cell.value).strip():
                    cells[(cell.row, cell.column)] = cell.value
        data[ws.title] = cells
    wb.close()
    return data


def _rows_by_key(cells):
    rows = {}
    for (r, c), v in cells.items():
        rows.setdefault(r, {})[c] = v
    out = {}
    for r, rc in rows.items():
        cols = sorted(rc); key = str(rc[cols[0]])
        for c in cols:
            out.setdefault(key, {})[c - cols[0]] = rc[c]
    return out


def test_content(result_x, golden_x, out_x, align='position',
                 ignore_space=True, ignore_width=True):
    res = _load_cells(result_x); gold = _load_cells(golden_x)
    names = sorted(set(res) | set(gold))
    common = [s for s in names if s in res and s in gold]
    only_r = [s for s in names if s in res and s not in gold]
    only_g = [s for s in names if s in gold and s not in res]
    if not common:
        raise ValueError(
            '两份表没有同名 sheet，没法逐格比对。'
            f'识别结果的 sheet：{list(res)}；标准答案的 sheet：{list(gold)}')

    wb = Workbook(); wb.remove(wb.active)
    thin = Side(style='thin', color='C8C8C8')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hf = PatternFill('solid', fgColor='DCE6F1')
    good = PatternFill('solid', fgColor='E2EFDA'); bad = PatternFill('solid', fgColor='FCE4D6')
    wsum = wb.create_sheet('准确率汇总')
    for ci, h in enumerate(['指标', '值'], 1):
        c = wsum.cell(1, ci, h); c.font = Font(bold=True); c.fill = hf
    wd = wb.create_sheet('逐格差异')
    for ci, h in enumerate(['Sheet', '位置', '标准答案', '识别结果', '严格一致', '归一化一致'], 1):
        c = wd.cell(1, ci, h); c.font = Font(bold=True); c.fill = hf

    total = sok = nok = 0; dr = 2
    for s in common:
        rc, gc = res[s], gold[s]
        if align == 'position':
            pairs = []
            for k in sorted(set(rc) | set(gc)):
                pairs.append((f'{get_column_letter(k[1])}{k[0]}', gc.get(k), rc.get(k)))
        else:
            rm, gm = _rows_by_key(rc), _rows_by_key(gc)
            pairs = []
            for key in sorted(set(rm) | set(gm)):
                for off in sorted(set(rm.get(key, {})) | set(gm.get(key, {}))):
                    pairs.append((f'键[{key}] 偏移{off}',
                                  gm.get(key, {}).get(off), rm.get(key, {}).get(off)))
        for loc, gv, rv in pairs:
            if gv is None and rv is None:
                continue
            total += 1
            strict = gv is not None and rv is not None and str(gv) == str(rv)
            norm = _norm(gv, ignore_space, ignore_width) == _norm(rv, ignore_space, ignore_width)
            sok += int(strict); nok += int(norm)
            if not norm:
                for ci, v in enumerate([s, loc, str(gv or ''), str(rv or ''),
                                        '是' if strict else '否', '否'], 1):
                    c = wd.cell(dr, ci, v); c.border = border
                wd.cell(dr, 6).fill = bad; dr += 1

    pct = lambda x: f'{x/total*100:.2f}%' if total else 'N/A'
    for i, (k, v) in enumerate([
        ('标准答案', os.path.basename(golden_x)), ('识别结果', os.path.basename(result_x)),
        ('对齐方式', align), ('共同Sheet', ', '.join(common)),
        ('仅结果有', ', '.join(only_r) or '(无)'), ('仅标准有', ', '.join(only_g) or '(无)'),
        ('比对单元格总数', total), ('严格一致数', sok), ('严格准确率', pct(sok)),
        ('归一化一致数', nok), ('归一化准确率', pct(nok)), ('差异数', total-nok),
    ], 2):
        wsum.cell(i, 1, k); wsum.cell(i, 2, v)
        if k in ('严格准确率', '归一化准确率'):
            wsum.cell(i, 2).font = Font(bold=True, size=12)
    wsum.column_dimensions['A'].width = 20; wsum.column_dimensions['B'].width = 58
    for ci, w in enumerate([22, 24, 40, 40, 10, 12], 1):
        wd.column_dimensions[get_column_letter(ci)].width = w
    wd.freeze_panes = 'A2'
    wb.save(out_x)
    return {'total': total, 'strict_acc': sok/total if total else 0,
            'norm_acc': nok/total if total else 0, 'report': out_x}


def run_all(inputs, outdir):
    os.makedirs(outdir, exist_ok=True)
    paths = []
    for inp in inputs:
        if os.path.isfile(inp): paths.append(inp)
        elif os.path.isdir(inp): paths.extend(sorted(glob.glob(os.path.join(inp, '*.pdf'))))
    paths = list(dict.fromkeys(paths))
    for p in paths:
        base = os.path.splitext(os.path.basename(p))[0]
        print(f'\n=== {base} ===', flush=True)
        part = os.path.join(outdir, base + '_提取.partial.jsonl')
        result = extract_pdf(p, layout='auto', partial_path=part)
        out = os.path.join(outdir, f'{base}_提取.xlsx')
        write_excel(result, out)
        print(f'  已保存: {out}', flush=True)
        if os.path.isfile(part):
            try:
                os.remove(part)
            except Exception:
                pass
    s = test_structure(paths, os.path.join(outdir, '结构识别报告.xlsx'))
    return s


# ============================================================ 菜单功能

def _outdir_for(p):
    """输出目录：给的是 PDF 文件 → 它所在目录；给的是目录 → 目录本身（绝不往上跳一级）"""
    ap = os.path.abspath(p)
    return ap if os.path.isdir(ap) else os.path.dirname(ap)


def _partial_for(pdf_path, outdir):
    """断点文件：<输出目录>/<PDF名>_提取.partial.jsonl"""
    stem = os.path.splitext(os.path.basename(pdf_path))[0]
    return os.path.join(outdir, stem + '_提取.partial.jsonl')


def _partial_count(path):
    """断点文件里已完成的页数"""
    if not path or not os.path.isfile(path):
        return 0
    n = 0
    try:
        with open(path, 'r', encoding='utf-8') as f:
            for line in f:
                if line.strip():
                    n += 1
    except Exception:
        return 0
    return n


def func_extract():
    print('\n--- 4. 把 PDF 表格导出成 Excel ---')
    p = ask_path('PDF 文件或目录')
    p = os.path.abspath(p)
    layout = ask_choice('文本表布局', [
        ('auto', '自动判别双栏/网格（推荐）'),
        ('dual', '双栏键值表'), ('grid', '网格表')], 'auto')
    pages = input('页码范围（如 1-4 / 1,3；回车=全部）: ').strip() or None
    want_json = ask_yesno('同时导出 JSON？', False)
    outdir = ask_path('输出目录', must_exist=False, default=_outdir_for(p))
    outdir = os.path.abspath(outdir)
    os.makedirs(outdir, exist_ok=True)
    print(f'  输出目录: {outdir}')
    paths = ([p] if os.path.isfile(p) else
             sorted([os.path.join(p, f) for f in os.listdir(p) if f.lower().endswith('.pdf')]))
    if not paths:
        print('  ! 目录里没有 PDF'); pause(); return
    for f in paths:
        print(f'\n提取: {f}', flush=True)
        part = _partial_for(f, outdir)
        n_done = _partial_count(part)
        resume = False
        if n_done:
            print(f'  发现断点文件，已完成 {n_done} 页：{os.path.basename(part)}')
            resume = ask_yesno('  是否从断点续跑？（否=重新提取全部）', True)
            if not resume:
                try:
                    os.remove(part)
                except Exception:
                    pass
                n_done = 0                     # 重头跑，本轮要重新记断点
        try:
            result = extract_pdf(f, layout=layout, pages=pages,
                                 partial_path=(part if (n_done == 0 or resume) else None))
        except KeyboardInterrupt:
            print(f'\n  ! 已中断。已完成的页保存在断点文件里：')
            print(f'    {part}')
            print('    重新运行选 2 并输入同一个 PDF，会问你是否续跑。')
            pause(); return
        except Exception as e:
            print(f'  ! 失败: {type(e).__name__}: {e}'); continue
        out = os.path.join(outdir, os.path.splitext(os.path.basename(f))[0] + '_提取.xlsx')
        write_excel(result, out); print(f'  已保存: {out}')
        if want_json:
            write_json(result, os.path.splitext(out)[0] + '.json')
        print(f'  汇总: {result["summary"]}')
        # 成功落盘后才删断点，避免半途失败丢数据
        if os.path.isfile(part):
            try:
                os.remove(part)
                print('  断点文件已清理')
            except Exception:
                pass
    pause()


def func_structure():
    print('\n--- 8. 测试能否识别这种表格（结构识别）---')
    p = ask_path('PDF 文件或目录')
    out = ask_path('报告输出（xlsx）', must_exist=False,
                   default=os.path.join(HERE, '结构识别报告.xlsx'))
    print('\n测试中...', flush=True)
    try:
        s = test_structure([p], out)
    except Exception as e:
        print(f'  ! 失败: {type(e).__name__}: {e}'); pause(); return
    print(f'\n完成: {s["pdfs"]}个PDF / {s["pages"]}页')
    print(f'  样式: {json.dumps(s["styles"], ensure_ascii=False)}')
    print(f'  报告: {s["report"]}')
    pause()


def func_content():
    print('\n--- 5. 测试识别准不准（内容准确性测试）---')
    print('比对两份**同结构**的 xlsx：一份是程序导出的识别结果，')
    print('一份是你人工改到完全正确的标准答案。标准流程：')
    print('  ① 选 4「把 PDF 表格导出成 Excel」 → 得到 <PDF名>_提取.xlsx')
    print('  ② 选 6「生成人工核对副本」      → 得到 <...>_人工核对.xlsx，手工改对')
    print('  ③ 回来选 5，分别选这两个文件，出逐格准确率报告。')
    print('  （想测报价表准不准也行：把生成的报价表复制一份、人工改对，两份一起选。）\n')
    exts = ('.xlsx', '.xlsm', '.xls')
    result = ask_path('识别结果 xlsx（result）', must_file=True, exts=exts,
                      default=_pick_local(['*_提取.xlsx']) or None)
    golden = ask_path('人工核对标准答案 xlsx（golden）', must_file=True, exts=exts,
                      default=_pick_local(['*人工核对*.xlsx', '*_核对*.xlsx']) or None)
    if not result or not golden:
        print('  ! 两个文件都要选。先按上面 ① ② 步生成，再回来比对。')
        pause(); return
    if os.path.abspath(result) == os.path.abspath(golden):
        print('  ! 两份选的是同一个文件，比对结果恒为 100%，没有意义。')
        pause(); return
    align = ask_choice('对齐方式', [
        ('position', '按行列位置（结构完全一致时，推荐）'),
        ('value', '按行首键（容忍少量列偏移）')], 'position')
    ignore_space = not ask_yesno('把空白差异算作不同？（默认忽略）', False)
    ignore_width = not ask_yesno('把全/半角差异算作不同？（默认忽略）', False)
    out = ask_path('报告输出（xlsx）', must_exist=False,
                   default=os.path.join(HERE, '准确率报告.xlsx'))
    if os.path.isdir(out):
        print(f'  ! 报告输出要给一个文件路径，不是目录: {out}')
        pause(); return
    print('\n比对中...', flush=True)
    try:
        r = test_content(result, golden, out, align, ignore_space, ignore_width)
    except Exception as e:
        print(f'  ! 失败: {e}'); pause(); return
    print(f'\n结果: {r["total"]}格 | 严格 {r["strict_acc"]*100:.2f}% | '
          f'归一化 {r["norm_acc"]*100:.2f}%')
    print(f'  报告: {r["report"]}')
    pause()


def func_runall():
    print('\n--- 7. 批量提取并检查表格结构 ---')
    p = ask_path('PDF 文件或目录')
    p = os.path.abspath(p)
    out = ask_path('输出目录', must_exist=False, default=_outdir_for(p))
    out = os.path.abspath(out)
    print('\n处理中...', flush=True)
    try:
        s = run_all([p], out)
        print(f'\n完成: {s["pages"]}页, 样式 {s["styles"]}')
    except Exception as e:
        print(f'  ! 失败: {type(e).__name__}: {e}')
    print(f'  输出: {out}')
    pause()


def func_golden():
    print('\n--- 6. 生成人工核对副本（golden 模板）---')
    exts = ('.xlsx', '.xlsm', '.xls')
    src = ask_path('提取结果 xlsx', must_file=True, exts=exts,
                   default=_pick_local(['*_提取.xlsx']) or None)
    if not src:
        print('  ! 没有选提取表。先选 4「把 PDF 表格导出成 Excel」生成一份。')
        pause(); return
    dst = ask_path('保存为（xlsx）', must_exist=False,
                   default=os.path.splitext(src)[0] + '_人工核对.xlsx')
    if os.path.isdir(dst):
        print(f'  ! 保存位置要给一个文件路径，不是目录: {dst}')
        pause(); return
    if os.path.abspath(src) == os.path.abspath(dst):
        print('  ! 目标不能等于源文件'); pause(); return
    shutil.copy2(src, dst)
    print(f'已复制: {dst}')
    print(f'手工改到完全正确后，回主菜单选 5，把「{os.path.basename(src)}」'
          f'和「{os.path.basename(dst)}」比对，就知道识别准不准了。')
    pause()


def ask_pdf_path(prompt, default=None):
    """输入 PDF 路径或目录。回车默认用 bat 所在目录，列出里面的 PDF 让选。

    支持三种输入：
      ① 直接回车        → 用默认目录（bat 所在路径），列出 PDF 让你选序号
      ② 输入一个目录    → 列出该目录的 PDF 让你选序号
      ③ 输入 PDF 路径   → 直接用（可以把 PDF 拖到 bat 图标上，或右键粘贴路径）
    """
    if default is None:
        default = HERE
    while True:
        raw = input(prompt + (f' [回车用当前目录: {default}]' if default else '') + ': ')
        p = unquote(raw) if raw.strip() else default
        if not p:
            print('  ! 路径不能为空'); continue
        if os.path.isdir(p):
            pdfs = sorted(f for f in os.listdir(p) if f.lower().endswith('.pdf'))
            if not pdfs:
                print(f'  ! 该目录下没有 PDF: {p}')
                print('    可以输入别的目录，或把 PDF 拖到 bat 图标上。')
                continue
            print(f'\n  目录 {p} 里有 {len(pdfs)} 个 PDF：')
            for i, f in enumerate(pdfs, 1):
                print(f'     {i}. {f}')
            sel = input('  输入序号选一个（回车=1，输入 d 换目录）: ').strip().lower()
            if sel == 'd':
                continue
            idx = int(sel) - 1 if sel.isdigit() and 1 <= int(sel) <= len(pdfs) else 0
            return os.path.join(p, pdfs[idx])
        if not os.path.exists(p):
            print(f'  ! 路径不存在: {p}'); continue
        if not p.lower().endswith('.pdf'):
            print('  ! 这不是 PDF 文件，请重新输入'); continue
        return p


def ask_template_path(prompt, default=None):
    """输入模板路径；校验它确实是报价表（表头含"位号"），防止选成提取结果"""
    if default is None:
        default = _find_default_template()
    while True:
        raw = input(prompt + (f' [回车用: {default}]' if default else '') + ': ')
        p = unquote(raw) if raw.strip() else default
        if not p:
            print('  ! 路径不能为空'); continue
        if not os.path.exists(p) or os.path.isdir(p):
            print(f'  ! 文件不存在: {p}'); continue
        # 校验表头：报价表 B1 应为"位号"
        # 直接把文件对象交给 openpyxl，不复制临时 .xlsx，避免 Windows 临时文件句柄占用
        try:
            with open(p, 'rb') as fh:
                wb = load_workbook(fh, read_only=True, data_only=True)
                try:
                    b1 = str(wb.worksheets[0]['B1'].value or '')
                finally:
                    wb.close()
            if '位号' not in b1:
                print(f'  ! 这不像报价表模板（B1="{b1}"），请选择"选型报价表"文件'); continue
        except Exception as e:
            print(f'  ! 无法读取模板: {type(e).__name__}: {e}'); continue
        return p


def _find_default_template():
    """默认报价表模板（候选名统一放在引擎里维护）"""
    try:
        from fill_from_extract import default_template
        p = default_template()
        if p:
            return p
    except Exception:
        pass
    for n in ('气动球阀（蝶阀）选型报价表.xls',
              '气动球阀（蝶阀）选型报价表2_已填充.xls'):
        p = os.path.join(HERE, n)
        if os.path.isfile(p):
            return p
    return ''


def _first_pdf():
    """tablekit 目录里第一个 PDF，用作输入默认值"""
    pdfs = sorted(glob.glob(os.path.join(HERE, '*.pdf')))
    return pdfs[0] if pdfs else ''


def func_fill():
    print('\n--- 1. PDF → 报价表（自动识别版式）---')
    print('选一个气动阀门数据表 PDF，程序自动判断版式，直接生成填好的报价表。\n')
    pdf = ask_pdf_path('数据表 PDF 路径')
    tpl = ask_template_path('报价表模板路径', default=_find_default_template())
    out = ask_path('输出文件路径（xlsx）', must_exist=False,
                   default=os.path.join(os.path.dirname(pdf),
                                        os.path.splitext(os.path.basename(pdf))[0] + '_报价表.xlsx'))
    _do_fill(pdf, tpl, out)
    pause()


def func_fill_batch():
    print('\n--- 2. 批量：整个目录的 PDF → 报价表 ---')
    d = ask_path('PDF 所在目录', must_exist=True, default=HERE)
    pdfs = sorted(p for p in glob.glob(os.path.join(d, '*.pdf')) if os.path.isfile(p))
    if not pdfs:
        print(f'  ! 该目录下没有 PDF: {d}'); pause(); return
    print(f'  找到 {len(pdfs)} 个 PDF：')
    for i, p in enumerate(pdfs, 1):
        print(f'     {i}. {os.path.basename(p)}')
    tpl = ask_template_path('报价表模板路径', default=_find_default_template())
    outdir = ask_path('输出目录', must_exist=False, default=d)
    os.makedirs(outdir, exist_ok=True)

    ok, fail = [], []
    for i, p in enumerate(pdfs, 1):
        base = os.path.splitext(os.path.basename(p))[0]
        out = os.path.join(outdir, base + '_报价表.xlsx')
        print(f'\n[{i}/{len(pdfs)}] {os.path.basename(p)}')
        try:
            rep = _do_fill(p, tpl, out, auto_yes=True)
            (ok if rep else fail).append(os.path.basename(p))
        except Exception as e:
            print(f'  ! 失败: {type(e).__name__}: {e}')
            fail.append(os.path.basename(p))

    print('\n' + '=' * 62)
    print(f'批量完成：成功 {len(ok)} 个，失败 {len(fail)} 个')
    for f in fail:
        print(f'  ! {f}')
    pause()


def _pick_layout(layouts):
    """手动选一个版式配置"""
    print('\n可用版式配置:')
    for i, l in enumerate(layouts, 1):
        tag = '  [自动学习]' if l.get('auto') else ''
        print(f'   {i}. {l["name"]}{tag}')
        print(f'      文件: {os.path.basename(l["file"])}')
    sel = input('  请选择序号 [回车=1]: ').strip() or '1'
    idx = int(sel) - 1 if sel.isdigit() and 1 <= int(sel) <= len(layouts) else 0
    return layouts[idx]['file']


def _review_unknown(pdf, auto_yes=False):
    """识别不出版式时的新手流程：

    ① 打印认到的阀门类型和候选分数（让人知道为什么没自动采用）
    ② 只解析前 3 页，存成一份"自动配置"
    ③ 弹【映射核对窗口（红线框）】给人看框得对不对
    ④ 关窗后问一句：对 → 用这份配置出完整 Excel；不对 → 停下，提示补 mappings

    返回可用的配置路径；用户否决或失败返回 None。
    """
    import fill_from_extract as E

    try:
        _, scores, info = E.detect_layout(pdf, with_info=True)
    except Exception:
        scores, info = [], {}

    vt = info.get('valve_type') or {}
    hits = vt.get('all') or []
    if hits:
        print('  阀门类型线索: ' + '、'.join(f'{t}×{n}' for t, n in hits[:5]))
    else:
        print('  阀门类型: 没认出来（球阀/闸阀/截止阀/蝶阀/切断阀/调节阀 都没命中）')
    if scores:
        print('  最接近的候选: ' + ', '.join(
            f'{os.path.basename(p)}={s}' for s, p, _ in scores[:3]))
    if scores and not info.get('sure', False):
        print(f'  最高分没到门槛 {info.get("threshold", 8)}，不能自动采用——先请你核对。')
        cov = info.get('coverage') or {}
        if cov.get('matched'):
            print(f'  （最像的那份配置能对上 {cov["matched"]} 个字段名，'
                  f'命中率 {cov.get("ratio", 0):.0%}）')
    elif not scores:
        print('  没有任何版式配置匹配得上——先请你核对。')

    if auto_yes:
        print('\n  【新版式】批量模式无人值守，跳过核对窗口，先用自动配置继续。')
    else:
        print('\n  【新版式核对】只取前 3 页做样本，马上打开红线框窗口。')
        print('    红线框 = 已映射并会写进报价表；青绿框 = 没映射。')
        print('    框得都对：关窗后选"是"，自动接着出完整 Excel。')
        print('    框错位：在窗口里拖框修正，保存会立刻写进 mappings/*.json。')

    try:
        valves, struct = E.parse_pdf_generic(pdf, max_pages=3)
    except Exception as e:
        print(f'  ! 前 3 页解析失败: {type(e).__name__}: {e}')
        return None
    if not valves:
        print('  ! 前 3 页没解析出阀门（没有文字层，或表格结构特殊）。')
        return None
    cfg_path = E._save_auto_config(pdf, struct, valves)
    if not cfg_path:
        print('  ! 自动配置保存失败。')
        return None
    print(f'  样本: {len(valves)} 台阀门 / 前 {min(3, info.get("pages", 3))} 页')

    # 核对窗口只看前 3 页：多页 PDF 临时裁一份，避免为了看样本等整本解析完。
    review_pdf, tmp = pdf, None
    try:
        import pymupdf as fitz
        src = fitz.open(pdf)
        if src.page_count > 3:
            dst = fitz.open()
            dst.insert_pdf(src, from_page=0, to_page=2)
            tmp = os.path.join(tempfile.gettempdir(), 'tablekit_review_3p.pdf')
            dst.save(tmp)
            dst.close()
            review_pdf = tmp
        src.close()
    except Exception:
        pass

    # 批量模式没人看着屏幕，不要弹窗口把整批卡住
    if not auto_yes:
        try:
            from pdf_review import open_review
            print('  正在打开核对窗口...', flush=True)
            open_review(review_pdf, cfg_path=cfg_path)
        except Exception as e:
            print(f'  ! 核对窗口打开失败: {type(e).__name__}: {e}')
    if tmp:
        try:
            os.remove(tmp)
        except OSError:
            pass

    if auto_yes:
        return cfg_path
    if ask_yesno('  映射正确吗？选"是"立即生成完整报价表', False):
        return cfg_path
    print('  已取消自动填表。请补充 mappings 配置（或修正框选）后再试。')
    return None


def _do_fill(pdf, tpl, out, auto_yes=False):
    """单个 PDF 填表：识别版式 → 填 → 汇报。失败返回 None。"""
    try:
        from fill_from_extract import fill_pdf, detect_layout, list_layouts, MAP_DIR
    except Exception as e:
        print(f'  ! 无法加载填表引擎: {type(e).__name__}: {e}')
        return None

    layouts = list_layouts()
    if not layouts:
        print('\n  ! mappings 目录里没有任何版式配置，无法填表。')
        return None

    cfg = None
    try:
        guessed, scores, info = detect_layout(pdf, with_info=True)
    except Exception as e:
        print(f'  ! 版式识别出错: {type(e).__name__}: {e}')
        guessed, scores, info = None, [], {}

    if guessed:
        gname = next((l['name'] for l in layouts if l['file'] == guessed),
                     os.path.basename(guessed))
        vt = (info.get('valve_type') or {}).get('all') or []
        if vt:
            print('  阀门类型: ' + '、'.join(f'{t}×{n}' for t, n in vt[:4]))
        print(f'  识别版式: {gname}')
        if scores and len(scores) > 1:
            print('  候选打分: ' + ', '.join(
                f'{os.path.basename(p)}={s}' for s, p, _ in scores[:4]))
        if auto_yes or ask_yesno('  用这个版式？（选否手动指定）', True):
            cfg = guessed
        else:
            cfg = _pick_layout(layouts)
    else:
        cfg = _review_unknown(pdf, auto_yes=auto_yes)
        if not cfg:
            return None

    try:
        rep = fill_pdf(pdf, tpl, out, config_path=cfg)
    except Exception as e:
        print(f'  ! 失败: {type(e).__name__}: {e}')
        print('    若是新版式，可打开 mappings 里的配置补全字段映射。')
        return None

    if rep['filled_rows'] == 0:
        print('  ! 填入 0 台：该 PDF 与所选配置不匹配。')
        return None
    print(f'  完成: 版式【{rep["layout"]}】')
    print(f'    填入 {rep["filled_rows"]} 台阀门，映射 {rep["mapped_fields"]} 个字段')
    print(f'    数据区第 {rep["data_rows"]} 行，总计在第 {rep["total_row"]} 行')
    if rep.get('learned_config'):
        print(f'    已保存新版式配置: {os.path.basename(rep["learned_config"])}（下次自动复用）')
    print(f'    输出: {rep["output"]}')

    # 填完直接弹核对窗口，确认"哪一格进了表"
    if not auto_yes and ask_yesno('  打开映射核对窗口看红线框？', True):
        try:
            from pdf_review import open_review
            print('  正在解析并打开窗口...', flush=True)
            open_review(pdf, cfg)
        except Exception as e:
            print(f'  ! 核对窗口打开失败: {type(e).__name__}: {e}')
    return rep


def func_review():
    print('\n--- 3. 映射核对窗口（红线框）---')
    print('左边显示 PDF 页面，被取走的值用红线框圈出（青绿框=没进表）；')
    print('右边列出提取到的字段与页面文字。\n')
    print('【怎么读这个窗口】')
    print('  红线框 = 这一格的值已经写进报价表；青绿框 = 没映射，没进表。')
    print('  左侧那些 Tag No./Service 是"字段名"，不是值，本来就不画框。')
    print('\n【可以在窗口里直接改】')
    print('  在左边页面上用鼠标拖一个框 → 告诉程序这格属于哪一项 → 保存。')
    print('  修正会写进该版式的配置，下次同版式的 PDF 自动生效。')
    print('\n【主值和副值怎么选】')
    print('  一张数据表经常在同一个序号行里并排放两个值，比如：')
    print('      序号 9 ： 设计压力 0.54  ｜  设计温度 680')
    print('      序号 5 ： 管道等级 T2A1  ｜  管道材质 A333 Gr.6,SMLS')
    print('  左边那一格叫【主值】，右边那一格叫【副值】。')
    print('  选错了，值就会进错报价表的列，所以框完要在弹窗里确认一下。')
    print('  只有一路值的行（没有左右两格）→ 一律选主值，弹窗会自动锁定。\n')
    pdf = ask_pdf_path('数据表 PDF 路径')
    try:
        from pdf_review import open_review
    except Exception as e:
        print(f'  ! 无法加载核对窗口: {type(e).__name__}: {e}')
        pause(); return
    print('\n正在解析并打开窗口...', flush=True)
    try:
        m = open_review(pdf)
        print(f'  已关闭窗口。共 {len(m["pages"])} 页，{m["valves"]} 台阀门。')
    except Exception as e:
        print(f'  ! 打开失败: {type(e).__name__}: {e}')
    pause()


def func_align():
    import fill_from_extract as E
    from fill_from_extract import default_template

    print('\n--- 11. 智能字段对齐（记忆 + 嵌入模型）---')
    print('遇到没见过的版式时用：自动猜"PDF 的哪一项该填报价表的哪一列"。')
    print('分三层：①历史配置里的同名字段（最准）②bge-small-zh 语义匹配')
    print('        ③猜不准的留空，之后用第 3 项的框选人工补。\n')
    pdf = ask_pdf_path('数据表 PDF 路径')
    if not pdf or not os.path.exists(pdf):
        print('  ! 文件不存在'); pause(); return
    tpl = ask_template_path('报价表模板路径', default=default_template())
    if not tpl or not os.path.exists(tpl):
        print('  ! 模板不存在'); pause(); return

    import fill_from_extract as E
    try:
        import semantic_map as S
    except Exception as e:
        print(f'  ! 无法加载语义模块: {type(e).__name__}: {e}'); pause(); return

    print('\n正在解析...', flush=True)
    cfg_path, _ = E.detect_layout(pdf)
    if cfg_path:
        cfg = E._load_config(cfg_path)
        valves = E.parse_pdf_detail(pdf, cfg)['valves']
    else:
        valves, struct = E.parse_pdf_generic(pdf)
        cfg_path = E._save_auto_config(pdf, struct, valves) if valves else None
        cfg = E._load_config(cfg_path) if cfg_path else {'fields': [], 'name': '通用'}
        print('  （新版式，已存成新配置）')
    if not valves:
        print('  ! 没解析到阀门'); pause(); return

    v = valves[0]
    labels, seqs = [], []
    for seq, f in sorted(v['fields'].items()):
        # 只取"字段名"那一列。name2 里经常放的是值（F347H、全通径球阀），
        # 拼进去会把语义匹配带偏；实测 name 单独用比 name+name2 准 6 个百分点。
        n1 = (f.get('name') or '').strip()
        n2 = (f.get('name2') or '').strip()
        nm = n1 or n2
        if nm:
            labels.append(nm)
            seqs.append(int(seq))
    if not labels:
        print('  ! 没读到字段名'); pause(); return

    headers, hrow = S.read_template_headers(tpl)
    mem = S.load_memory()
    print(f'  模板表头第 {hrow} 行，共 {len(headers)} 列；记忆词表 {len(mem)} 条')
    print('  正在对齐...', flush=True)
    res = S.suggest(labels, headers, memory=mem)

    manual = {int(f['seq']): f.get('to') for f in (cfg.get('fields') or [])
              if f.get('to') and not f.get('sub')}
    print(f"\n  {'序号':<5}{'PDF 字段':<26}{'→ 报价表列':<12}{'依据':<8}")
    print('  ' + '-' * 66)
    n_hit = n_new = 0
    for s, x in zip(seqs, res):
        old = manual.get(s)
        tag = ''
        if old:
            tag = '（已有）'
            n_hit += 1
        elif x['to']:
            n_new += 1
        print(f"  {s:<5}{x['label'][:24]:<26}"
              f"{(x['to'] + ' ' + (x['header'] or '')[:10]) if x['to'] else '—':<14}"
              f"{x['src'] or '—':<8}{tag}")
    blank = sum(1 for x in res if not x['to'])
    print(f"\n  已有人工映射 {n_hit} 项｜本次新增建议 {n_new} 项｜猜不出 {blank} 项")

    if not cfg_path:
        print('\n  没有可写的配置文件，无法保存。'); pause(); return
    if n_new == 0:
        print('\n  没有需要新增的映射。'); pause(); return
    if not ask_yesno(f'\n  把这 {n_new} 条建议写进版式配置？', True):
        pause(); return

    # 只补空缺，且同一列不重复占用
    used = {f.get('to') for f in (cfg.get('fields') or []) if f.get('to')}
    added = 0
    for s, x in zip(seqs, res):
        if not x['to'] or s in manual or x['to'] in used:
            continue
        cfg['fields'].append({'seq': s, 'to': x['to'],
                              'label': (x['header'] or x['label'])[:16]})
        used.add(x['to'])
        added += 1
    with open(cfg_path, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    E.clear_cache()
    print(f'  已写入 {added} 条 → {os.path.basename(cfg_path)}')

    if ask_yesno('  现在就用新配置重新生成报价表？', True):
        out = os.path.splitext(pdf)[0] + '_报价表.xlsx'
        try:
            rep = E.fill_pdf(pdf, tpl, out)
            print(f"  已生成：{os.path.basename(out)}")
            print(f"  填入 {rep['filled_rows']} 台，映射 {rep['mapped_fields']} 个字段，"
                  f"数据区 {rep['data_rows']} 行")
        except Exception as e:
            print(f'  ! 生成失败: {type(e).__name__}: {e}')
    print('\n  提示：跑一遍第 10 项「空值体检」，剩下的空格再回第 3 项框选补。')
    pause()


def func_audit():
    print('\n--- 10. 空值体检（为什么表里有很多空格）---')
    print('把报价表里的空格分成两类：')
    print('  · PDF 原始就是空的 —— 客户没填，填不出来是正常的')
    print('  · PDF 里有值却没取到 —— 真正的漏读，需要修')
    print()
    pdf = ask_pdf_path('数据表 PDF 路径')
    if not pdf or not os.path.exists(pdf):
        print('  ! 文件不存在'); pause(); return
    try:
        from fill_from_extract import audit_pdf
    except Exception as e:
        print(f'  ! 无法加载引擎: {e}'); pause(); return

    print('\n正在解析并体检...', flush=True)
    r = audit_pdf(pdf)
    total = max(r['total'], 1)
    print(f"\n  版式：{r['config']}")
    print(f"  {r['valves']} 台 × {r['fields']} 字段 = {r['total']} 格")
    print(f"  已填 {r['filled']} ({r['filled']*100//total}%)｜"
          f"PDF 本就空白 {r['blank_pdf']}｜疑似漏读 {len(r['suspect'])}")

    if r['suspect']:
        print('\n  【疑似漏读】PDF 这一格有值，但没进报价表：')
        for s in r['suspect'][:40]:
            txt = ' / '.join(f'[{i}]{c[:20]}' for i, c in s['cands'][:3])
            pos = '副值' if s['sub'] else '主值'
            print(f"    {s['tag']:<16} {s['to']:<4} 第{s['seq']}项{pos} {s['label']:<14} → {txt}")
        if len(r['suspect']) > 40:
            print(f"    ... 还有 {len(r['suspect'])-40} 条")
        print('  → 这些要在 mappings 里补字段，或调整 parse 的列位。')
    else:
        print('\n  【没有漏读】所有空格都是 PDF 原始就空着的。')

    if r['dead_fields']:
        print(f"\n  【整列全空 {len(r['dead_fields'])} 项】这份数据表里没有该项，或客户整份未填：")
        print('    ' + '、'.join(f'{k[1]}({k[0]})' for k in r['dead_fields']))
    if r['part_fields']:
        print(f"\n  【部分空 {len(r['part_fields'])} 项】有的阀填了、有的没填：")
        for k, c in r['part_fields']:
            print(f"    {k[1]}({k[0]}) {c}/{r['valves']}")

    # 存一份报告，方便发给客户确认
    try:
        rep = os.path.splitext(pdf)[0] + '_空值体检.txt'
        with open(rep, 'w', encoding='utf-8') as f:
            f.write(f"版式：{r['config']}\n{r['valves']} 台 × {r['fields']} 字段 = {r['total']} 格\n")
            f.write(f"已填 {r['filled']}｜PDF 本就空白 {r['blank_pdf']}｜疑似漏读 {len(r['suspect'])}\n\n")
            for s in r['suspect']:
                f.write(f"{s['tag']}\t{s['to']}\t第{s['seq']}项\t{s['label']}\t"
                        + '\t'.join(f'[{i}]{c}' for i, c in s['cands']) + '\n')
            f.write('\n整列全空：' + '、'.join(f'{k[1]}({k[0]})' for k in r['dead_fields']) + '\n')
        print(f'\n  报告已保存：{os.path.basename(rep)}')
    except Exception:
        pass
    pause()


def func_layouts():
    print('\n--- 9. 已保存的版式清单 ---')
    try:
        from fill_from_extract import list_layouts, MAP_DIR
    except Exception as e:
        print(f'  ! 无法加载引擎: {e}'); pause(); return
    layouts = list_layouts()
    if not layouts:
        print('  mappings 目录里还没有任何版式配置。'); pause(); return
    print(f'  共 {len(layouts)} 个版式（每个对应一种 PDF 版式，新增只需加一个 JSON）\n')
    for i, l in enumerate(layouts, 1):
        tag = '  [自动学习]' if l.get('auto') else ''
        print(f'   {i}. {l["name"]}{tag}')
        print(f'      {os.path.basename(l["file"])}')
    print('\n  目录:', MAP_DIR)
    print('  同一个版式的 PDF 以后会直接命中这里，不用再配。')
    pause()


def func_map():
    print('\n--- 4. 批量映射（工作台）---')
    print('把这份 PDF 的所有字段列成一张表，下拉选「写入报价表哪一列」，')
    print('可以多选整批指定；保存即写进版式配置，同版式以后自动生效。\n')
    pdf = ask_pdf_path('数据表 PDF 路径')
    try:
        from map_workbench import open_workbench
    except Exception as e:
        print(f'  ! 无法加载映射工作台: {type(e).__name__}: {e}')
        pause(); return
    open_workbench(pdf)
    pause()


MENU_GROUPS = [
    ('【常用】', [
        ('1', 'PDF → 报价表', func_fill, '自动识别版式，直接出填好的报价表'),
        ('2', '批量：整个目录的 PDF → 报价表', func_fill_batch, '一次出多份报价表'),
        ('3', '映射核对窗口（红线框）', func_review, '看哪一格进了报价表'),
        ('4', '批量映射（工作台）', func_map, '一屏把"字段名→哪一列"定完'),
    ]),
    ('【进阶】', [
        ('5', '把 PDF 表格导出成 Excel', func_extract, '要人工看表格原文时用'),
        ('6', '空值体检（为什么有很多空格）', func_audit, '分清"PDF 没填"和"漏读"'),
        ('7', '已保存的版式清单', func_layouts, 'mappings 里有哪些版式'),
        ('8', '智能字段对齐（记忆 + 模型）', func_align, '新版式自动猜该填哪一列'),
    ]),
]

ACTIONS = {k: fn for _, items in MENU_GROUPS for k, _, fn, _ in items}


def func_help():
    print('\n--- 新手说明：我该选哪个？---')
    print('【做报价表】标准流程（现在只要一步）：')
    print('  选 1，输入气动阀门数据表 PDF，再选报价表模板，')
    print('  程序自动判断这是哪种版式，直接生成填好数据的报价表。')
    print('  手上有好几个 PDF：选 2，输一个目录，一次全出。')
    print('\n【版式是怎么回事】')
    print('  每种不同的数据表版式，在 mappings 里存一个 JSON（记的是"哪一项填哪一列"）。')
    print('  新 PDF 进来：先从正文认阀门类型（球阀/闸阀/截止阀/蝶阀/切断阀/调节阀），')
    print('  再拿它的字段名和每份配置的字段名对一遍，命中率最高的那份自动采用。')
    print('  同一家设计院改版、多页文件后几页多一格、左右两栏各一套序号，')
    print('  都不影响：取值是按字段名+行内相对位置找的，不靠写死的序号/列号。')
    print('  完全没见过版式：只取前 3 页弹红线框窗口让你核对。')
    print('  确认框得对 → 关窗后选"是"，马上出完整 Excel；')
    print('  框错位 → 在窗口里拖框修正，**保存即写入 mappings/*.json**，下次直接生效。')
    print('  选 9 可以看现在一共存了哪些版式。')
    print('\n【想确认"哪一格进了报价表"】')
    print('  选 3：左边是 PDF 页面，被取走的值用红线框圈出来；')
    print('  右边是提取到的字段（序号 / 字段 / 值 / 写入列）和页面标题正文。')
    print('  点右边任意一行，左边会跳到对应格子；左右方向键翻页，Ctrl+滚轮缩放。')
    print('  功能 1 填完表也会问你要不要顺手打开这个窗口。')
    print('\n【其他功能】')
    print('  - 想把 PDF 表格原文导成 Excel 人工看：选 4。')
    print('  - 想知道识别准不准：先选 4 导出，选 6 生成核对副本，')
    print('    人工改好后回来选 5 比对。')
    print('  - 一次提取多个 PDF 的表格：选 7。')
    print('  - 本工具只处理有文字层的 PDF（设计院发来的数据表都属于这种）。')
    print('    纯图片/扫描件没有文字层，读不出内容——这种需要先去 OCR，本工具不做。')
    print('\n小提醒：')
    print('  - 模板要选「选型报价表」，不要把提取结果当模板。')
    print('  - 输入路径时，可以直接把文件拖进窗口，不用手打。')
    print('  - 生成的报价表默认和 PDF 放在同一个目录，文件名是「原名_报价表.xlsx」。')
    print('  - 每个功能结束后都会回到这个菜单。')
    print('\n【页数多的时候（200页以上）】')
    print('  - 会显示进度条：[####----] 页码/总页 百分比 已用时间 预计剩余。')
    print('  - 处理期间每完成一页就存一次断点，文件名叫「xxx_提取.partial.jsonl」。')
    print('  - 如果卡死、断电、或者你按了 Ctrl+C，重新运行选 3 并输入同一个 PDF，')
    print('    程序会问「是否从断点续跑」，选 Y 就跳过已完成的页，不会从头再来。')
    print('  - 全部成功后断点文件会自动删除，不用手动清理。')
    print('  - 解析结果缓存会加快二次填表，怀疑结果不对时用菜单里清缓存。')
    print('  - 填报价表只读 PDF 文字层，几十页通常几秒完成。')
    pause()


def quick(pdf=None):
    """一键模式：给一个 PDF，直接在它旁边建同名文件夹，把报价表放进去。

    大多数人只想要这一步，不需要看菜单。bat 支持把 PDF 直接拖上来。
    """
    import fill_from_extract as E
    from fill_from_extract import default_template

    print('\n' + '=' * 58)
    print('  PDF → 报价表（一键）')
    print('=' * 58)
    print('  把 PDF 拖到 bat 图标上，或直接粘贴路径；')
    print('  也可以直接按回车，列出本文件夹里的 PDF 让你选。')
    print('  （注意：拖到已经打开的窗口里没用，Windows 窗口不接收文件拖放，')
    print('    要拖到 bat 文件的图标上。）')
    print('  生成的报价表会放进「PDF 同名文件夹」里。')
    print('  （要看更多功能，双击「完整菜单.bat」）\n')

    if not pdf:
        try:
            pdf = ask_pdf_path('PDF 路径')
        except EOFError:
            print('  ! 没有输入 PDF 路径。'
                  '把 PDF 拖到「启动工具.bat」图标上，或双击后按回车从列表里选。')
            return 1
    if not pdf:
        print('  ! 没有输入路径'); return 1
    if not os.path.isfile(pdf):
        print(f'  ! 文件不存在: {pdf}'); return 1
    if not pdf.lower().endswith('.pdf'):
        print('  ! 这不是 PDF 文件'); return 1

    tpl = default_template()
    if not tpl or not os.path.isfile(tpl):
        print('  ! 找不到报价表模板（气动球阀（蝶阀）选型报价表.xls）')
        print('    请把它和本程序放在同一目录。'); return 1

    stem = os.path.splitext(os.path.basename(pdf))[0]
    out_dir = os.path.join(os.path.dirname(os.path.abspath(pdf)), stem)
    out_path = os.path.join(out_dir, f'{stem}_报价表.xlsx')

    print(f'\n  正在处理：{os.path.basename(pdf)}', flush=True)
    try:
        os.makedirs(out_dir, exist_ok=True)
    except Exception as e:
        print(f'  ! 无法创建文件夹: {type(e).__name__}: {e}'); return 1

    t0 = time.time()
    try:
        # 先认阀门类型 → 匹配版式；认不出来就走"前 3 页 + 红线框"核对流程
        guessed, scores, info = E.detect_layout(pdf, with_info=True)
        if guessed:
            cfg = guessed
            vt = (info.get('valve_type') or {}).get('all') or []
            if vt:
                print('  阀门类型: ' + '、'.join(f'{t}×{n}' for t, n in vt[:4]))
            print(f'  自动匹配版式: {os.path.basename(cfg)}')
        else:
            cfg = _review_unknown(pdf)
            if not cfg:
                return 1
        rep = E.fill_pdf(pdf, tpl, out_path, config_path=cfg)
    except Exception as e:
        print(f'\n  ! 处理失败: {type(e).__name__}: {e}')
        print('    如果是全新版式，请用「完整菜单.bat」的第 3 项核对后重试。')
        return 1
    dt = time.time() - t0

    print('\n  完成！')
    print(f'    版式      {rep["layout"]}')
    print(f'    阀门台数  {rep["filled_rows"]} 台')
    print(f'    映射字段  {rep["mapped_fields"]} 个')
    print(f'    耗时      {dt:.1f} 秒')
    print(f'    输出文件  {out_path}')
    if rep.get('learned_config'):
        print('    （这是新版式，已自动记下来，下次同样的表直接出）')
    print('\n  文件夹已打开，可以直接看报价表。')
    try:
        os.startfile(out_dir)
    except Exception:
        pass
    return 0


def main():
    while True:
        print('\n' + '=' * 62)
        print('  PDF 表格处理工具（完整菜单）')
        print('=' * 62)
        print('  只要出报价表 → 用「启动工具.bat」更省事，不用看菜单。')
        print('  问路径时直接按回车 = 用本文件夹；问 PDF 时会列出 PDF 让你选序号。')
        print('  不确定怎么选时，按 H 看说明。')
        for group, items in MENU_GROUPS:
            print('\n' + group)
            for k, desc, _, tip in items:
                print(f'   {k}. {desc}   —— {tip}')
        print('\n   0. 退出')
        print('   H. 新手说明 / 我该选哪个')
        try:
            choice = input('\n请选择: ').strip().upper()
        except EOFError:
            print('\n(输入已结束) 已退出。')
            break
        if choice == '0':
            print('已退出。'); break
        if choice == 'H':
            func_help(); continue
        hit = ACTIONS.get(choice)
        if hit is None:
            print(f'  ! 无效选择，请输入 0-{len(ACTIONS)} 或 H。'); continue
        try:
            hit()
        except KeyboardInterrupt:
            print('\n  (已取消)')
        except Exception as e:
            print(f'\n  ! 错误: {type(e).__name__}: {e}')


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if a and not a.startswith('-')]
    flags = [a for a in sys.argv[1:] if a.startswith('-')]
    if '--menu' in flags:
        main()
    elif '--map' in flags:
        # 批量映射工作台：拖 PDF 上来直接进；没给就列出本目录的 PDF 让选
        pdf = args[0] if args else None
        if not pdf:
            try:
                pdf = ask_pdf_path('要批量映射的 PDF')
            except EOFError:
                pdf = ''
        if pdf and os.path.isfile(pdf):
            from map_workbench import open_workbench
            open_workbench(pdf)
            sys.exit(0)
        print('  ! 没有可用的 PDF 路径。')
        sys.exit(1)
    else:
        # 默认一键模式；命令行给了 PDF 路径就直接用（支持拖拽到 bat 上）
        # 结尾不自己等回车——bat 里统一 pause，避免要按两次。
        sys.exit(quick(args[0] if args else None))
