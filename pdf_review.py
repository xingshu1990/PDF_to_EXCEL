"""
映射核对窗口（红线框预览）
=====================================================================
打开一个 PDF，左边显示页面、把"被映射取走的那一格"用红线框圈出来，
右边列出提取到的内容：字段清单（序号 / 字段 / 值 / 写入列）＋ 页面标题正文。

数据来自 fill_from_extract.build_middle() 生成的 middle 结构
（参考 MinerU 的 middle.json：每页若干带 bbox 的语义块）。
本文件只负责画，不参与解析，也不改任何表。

单独运行：
    python pdf_review.py <PDF> [配置.json]
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import tkinter as tk
from tkinter import ttk

# 引擎模块：多个回调（含 _pick_cell）都要用，统一在模块级导入一次，
# 避免某个函数里忘了局部 import 导致 UnboundLocalError。
import fill_from_extract as engine

# 配色
C_BG = '#f5f6f8'
C_PAPER = '#ffffff'
C_RED = '#e02020'          # 已映射（艳红）
C_RED_FILL = '#ffe8e8'
C_TEAL = '#00897b'         # 未映射（青绿，和艳红反差最大）
C_BLUE = '#0050c8'         # 选中高亮 / 拖框
C_ORANGE_TXT = '#b26a00'   # 警示文字（"已有修正"）
C_TEXT = '#222222'
C_MUTED = '#777777'

FONT_UI = ('Microsoft YaHei UI', 10)
FONT_SM = ('Microsoft YaHei UI', 9)
FONT_MONO = ('Consolas', 9)


# ============================== 渲染工具 ==============================

def render_page_image(pdf_path, page_idx, zoom=1.4):
    """把某页渲染成 PIL Image（1 点 = zoom 像素）"""
    try:
        import pymupdf as fitz
    except Exception:
        import fitz
    from PIL import Image
    doc = fitz.open(pdf_path)
    page = doc[page_idx]
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    img = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
    doc.close()
    return img


def render_page_png(pdf_path, middle, page_no, out_png, zoom=1.6):
    """离线导出一页带红线框的 PNG（不开窗口，用于核对/存档）"""
    from PIL import ImageDraw
    img = render_page_image(pdf_path, page_no - 1, zoom)
    draw = ImageDraw.Draw(img)
    page = next((p for p in middle['pages'] if p['page_no'] == page_no), None)
    if page:
        for b in page['blocks']:
            if b['type'] != 'field' or not b.get('bbox'):
                continue
            x0, y0, x1, y1 = b['bbox']
            draw.rectangle([x0 * zoom, y0 * zoom, x1 * zoom, y1 * zoom],
                           outline=C_RED if b.get('mapped') else C_TEAL, width=2)
    img.save(out_png)
    return out_png


# ============================== 人工框选对话框 ==============================

class PickDialog(tk.Toplevel):
    """框选后确认：这一格属于第几项、算主值还是副值。"""

    def __init__(self, master, info, has_override, on_ok, on_del=None):
        super().__init__(master)
        self.title('人工框选 — 确认这一格')
        self.configure(bg=C_BG)
        self.resizable(False, False)
        self.transient(master)
        self.on_ok = on_ok
        self.on_del = on_del
        self.info = info
        self.result = None

        pad = dict(padx=14, pady=6)
        tk.Label(self, text='你框中的单元格', bg=C_BG, fg=C_MUTED,
                 font=FONT_SM).grid(row=0, column=0, columnspan=2, sticky='w', **pad)

        box = tk.Frame(self, bg=C_PAPER, relief='solid', bd=1)
        box.grid(row=1, column=0, columnspan=2, sticky='we', padx=14, pady=(0, 8))
        tk.Label(box, text=f'第 {info["seq"]} 项  {info["label"]}', bg=C_PAPER, fg=C_TEXT,
                 font=FONT_UI, anchor='w').pack(fill='x', padx=8, pady=(6, 0))
        txt = info.get('text') or ''
        tk.Label(box, text=(f'内容：{txt}' if txt else '（本页这一格是空的）'),
                 bg=C_PAPER, fg=C_BLUE if txt else C_MUTED,
                 font=FONT_UI, anchor='w', wraplength=470,
                 justify='left').pack(fill='x', padx=8, pady=(2, 0))
        tk.Label(box, text=f'第 {info["col"]} 列', bg=C_PAPER, fg=C_MUTED,
                 font=FONT_SM, anchor='w').pack(fill='x', padx=8, pady=(0, 6))

        # —— 前缀：这一格属于哪个字段 ——
        # 一行常并排两个字段名（防火结构 / 防杆吹出），以前前缀写死不能改，
        # 框错了只能改主/副值，改不了归属。这里列出该行全部候选让人选。
        cands = info.get('candidates') or []
        self.var_label = tk.StringVar()
        if cands:
            tk.Label(self, text='这一格属于哪个字段？（可改，改完按这个前缀取值）',
                     bg=C_BG, fg=C_TEXT, font=FONT_UI).grid(
                row=9, column=0, columnspan=2, sticky='w', padx=14, pady=(2, 2))
            from tkinter import ttk
            cb = ttk.Combobox(self, textvariable=self.var_label,
                              values=[c['display'] for c in cands],
                              state='normal', width=64)
            cb.grid(row=10, column=0, columnspan=2, sticky='w', padx=20, pady=(0, 4))
            cur = info.get('label') or ''
            pick = 0
            for i, c in enumerate(cands):
                if c['label'] == cur or cur in c['display']:
                    pick = i
                    break
            cb.current(pick)
            self._cands = cands
        else:
            self._cands = []

        # —— 这一格是什么（字段名 / 主值 / 副值）——
        self.var_mode = tk.StringVar(value='value')
        tk.Label(self, text='这一格是：', bg=C_BG, fg=C_TEXT,
                 font=FONT_UI).grid(row=11, column=0, sticky='w', padx=14, pady=(8, 0))
        mf = tk.Frame(self, bg=C_BG)
        mf.grid(row=11, column=1, sticky='w', pady=(8, 0))
        if info.get('force_mode') == 'key':
            self.var_mode.set('key')
        for val, title in (('key', '字段名'), ('value', '主值'), ('sub', '副值')):
            tk.Radiobutton(mf, text=title, variable=self.var_mode, value=val, bg=C_BG,
                           fg=C_TEXT, font=FONT_UI, activebackground=C_BG).pack(side='left')

        # —— 写入报价表哪一列 ——
        cols = getattr(info, 'get', lambda k, d=None: d)('columns', None) or []
        self.var_to = tk.StringVar()
        self.var_ins = tk.IntVar(value=1)
        tk.Label(self, text='写入列：', bg=C_BG, fg=C_TEXT, font=FONT_UI).grid(
            row=12, column=0, sticky='w', padx=14, pady=(6, 0))
        self.cbo_to = ttk.Combobox(
            self, textvariable=self.var_to, width=46, state='readonly',
            values=['(不指定，沿用原配置)'] + [f'{L}  {n}' for L, n in cols])
        self.cbo_to.grid(row=12, column=1, sticky='w', pady=(6, 0))
        tk.Checkbutton(self, text='插到该列前面（不覆盖原列）', variable=self.var_ins,
                       bg=C_BG, fg=C_TEXT, font=FONT_SM,
                       activebackground=C_BG).grid(row=13, column=1, sticky='w')

        # —— 主值 / 副值的说明 ——
        tk.Label(self, text='这一格算作哪一路值？（两个都能选）', bg=C_BG, fg=C_TEXT,
                 font=FONT_UI).grid(row=2, column=0, columnspan=2, sticky='w',
                                    padx=14, pady=(2, 2))

        self.var_sub = tk.IntVar(value=1 if info.get('guess_sub') else 0)
        rec = info.get('recommend') or ''       # '主值' / '副值' / ''

        def opt(row, val, title, desc):
            label = title + ('　← 建议' if rec == title else '')
            tk.Radiobutton(self, text=label, variable=self.var_sub, value=val,
                           bg=C_BG, fg=(C_BLUE if rec == title else C_TEXT),
                           font=FONT_UI, activebackground=C_BG,
                           anchor='w').grid(row=row, column=0, sticky='nw',
                                            padx=(20, 6), pady=(4, 0))
            tk.Label(self, text=desc, bg=C_BG, fg=C_MUTED, font=FONT_SM,
                     wraplength=420, justify='left', anchor='w').grid(
                row=row, column=1, sticky='w', pady=(6, 0))

        opt(3, 0, '主值',
            '这一项的第一个值（表格里靠左的那格）。'
            '报价表该字段默认取它。')
        opt(4, 1, '副值',
            '同一行并排的第二个值（靠右那格）。'
            '一行里塞了两项时才用，例如「设计压力 0.54 ｜ 设计温度 680」，'
            '右边那个就是副值。')

        # 全文统计：这一项在别的页/别的台上到底有没有值，是主值多还是副值多
        hint = info.get('hint') or ''
        if hint:
            tk.Label(self, text=hint, bg=C_BG, fg=C_BLUE, font=FONT_SM,
                     wraplength=520, justify='left', anchor='w').grid(
                row=5, column=0, columnspan=2, sticky='w', padx=20, pady=(8, 0))

        # 复用情况：这一项是不是已经有修正了
        old = info.get('existing')
        if old:
            tk.Label(self, text=f'⚠ 这一项已有修正：按「{old}」取第 {info.get("old_col")} 列，'
                                f'保存会覆盖它。',
                     bg=C_BG, fg=C_ORANGE_TXT, font=FONT_SM,
                     wraplength=520, justify='left', anchor='w').grid(
                row=6, column=0, columnspan=2, sticky='w', padx=20, pady=(6, 0))

        tk.Label(self, text='保存后写进该版式的配置，同版式的 PDF 全部按这一列取值，'
                            '不用再框。工具条「人工修正」里能看到并删除已保存的条目。',
                 bg=C_BG, fg=C_MUTED, font=FONT_SM,
                 wraplength=520, justify='left').grid(
            row=7, column=0, columnspan=2, sticky='w', padx=14, pady=(10, 2))

        btns = tk.Frame(self, bg=C_BG)
        btns.grid(row=8, column=0, columnspan=2, sticky='e', padx=14, pady=(6, 12))
        if has_override and on_del:
            tk.Button(btns, text='删除这项的人工修正', font=FONT_SM,
                      command=self._del).pack(side='left', padx=(0, 8))
        tk.Button(btns, text='键不对，去框正确的键', font=FONT_SM,
                  command=self._rekey).pack(side='left', padx=(0, 8))
        tk.Button(btns, text='取消', font=FONT_SM, command=self.destroy).pack(side='left')
        tk.Button(btns, text='保存到版式配置', font=FONT_SM, default='active',
                  command=self._ok).pack(side='left', padx=(8, 0))

        self.bind('<Return>', lambda e: self._ok())
        self.bind('<Escape>', lambda e: self.destroy())
        self.grab_set()
        self.update_idletasks()
        # 居中到主窗口
        mx, my = master.winfo_rootx(), master.winfo_rooty()
        mw, mh = master.winfo_width(), master.winfo_height()
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f'+{mx + (mw - w) // 2}+{my + (mh - h) // 3}')

    def _chosen_label(self):
        val = self.var_label.get().strip()
        for c in self._cands:
            if c['display'] == val:
                return c['label']
        return val          # 允许手打一个新的字段名

    def _ok(self):
        to = self.var_to.get().strip()
        to = to.split()[0] if to and not to.startswith('(') else ''
        self.result = ('ok', int(self.info['seq']), int(self.info['col']),
                       bool(self.var_sub.get()), self._chosen_label(),
                       self.var_mode.get(), to, bool(self.var_ins.get()))
        self.destroy()

    def _del(self):
        self.result = ('del', int(self.info['seq']), 0,
                       bool(self.var_sub.get()), self._chosen_label())
        self.destroy()

    def _rekey(self):
        to = self.var_to.get().strip()
        to = to.split()[0] if to and not to.startswith('(') else ''
        self.result = ('rekey', int(self.info['seq']), int(self.info['col']),
                       bool(self.var_sub.get()), '', 'key', to,
                       bool(self.var_ins.get()))
        self.destroy()


# ============================== 人工修正清单 ==============================

class OverrideListDialog(tk.Toplevel):
    """列出这个版式所有的"人工改动"，能逐条删。

    三类都列出来（以前只列 overrides，所以在工作台/核对窗口里改的东西看不到）：
      o = 取值修正（overrides：第几项取 PDF 第几列）
      f = 字段映射（fields：字段名 → 报价表哪一列）
      x = 新增列（extra_columns：插在模板哪一列后面）
    """

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.title('人工修正 / 人工审核')
        self.configure(bg=C_BG)
        self.transient(master)
        self.geometry('820x460')

        tk.Label(self, text='这些改动都写在版式配置里，同版式的 PDF 每次自动套用。'
                            '选中后可以删除。',
                 bg=C_BG, fg=C_MUTED, font=FONT_SM, anchor='w').pack(
            fill='x', padx=12, pady=(10, 6))

        box = tk.Frame(self, bg=C_BG)
        box.pack(fill='both', expand=True, padx=12)
        cols = ('seq', 'label', 'kind', 'col', 'where')
        self.tree = ttk.Treeview(box, columns=cols, show='headings', height=14)
        for c, t, w, al in (('seq', '序号', 52, 'center'),
                            ('label', '字段/项', 230, 'w'),
                            ('kind', '改动类型', 150, 'w'),
                            ('col', 'PDF列', 56, 'center'),
                            ('where', '写入/位置', 190, 'w')):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor=al, stretch=(c in ('label', 'where')))
        vs = ttk.Scrollbar(box, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vs.pack(side='right', fill='y')

        bar = tk.Frame(self, bg=C_BG)
        bar.pack(fill='x', padx=12, pady=10)
        tk.Button(bar, text='删除选中', font=FONT_SM,
                  command=self._del_sel).pack(side='left')
        tk.Button(bar, text='关闭', font=FONT_SM,
                  command=self.destroy).pack(side='right')
        self.lbl = tk.Label(bar, text='', bg=C_BG, fg=C_MUTED, font=FONT_SM)
        self.lbl.pack(side='left', padx=12)

        self._rows = []
        self.refresh()
        self.grab_set()

    def refresh(self):
        import fill_from_extract as engine
        self.tree.delete(*self.tree.get_children())
        self._rows = []
        cfg = engine._load_config(self.app.cfg_path) if self.app.cfg_path else {}
        stats = self.app.middle.get('seq_stats') or {}

        for i, o in enumerate((cfg.get('parse') or {}).get('overrides') or []):
            seq, sub = int(o.get('seq', 0)), bool(o.get('sub'))
            st = stats.get(str(seq)) or {}
            n = st.get('sub' if sub else 'main', 0)
            tot = st.get('total', 0)
            iid = f'o{i}'
            self.tree.insert('', 'end', iid=iid, values=(
                seq, (o.get('label') or '') + ('（副值）' if sub else ''),
                ('按前缀' if o.get('label') else '按主/副值') + '取值',
                o.get('col'), f'{n}/{tot} 台有值'))
            self._rows.append({'kind': 'o', 'obj': o, 'iid': iid})

        for i, f in enumerate(cfg.get('fields') or []):
            if f.get('meta') or f.get('seq') is None:
                continue
            if f.get('src_col') is None and not f.get('fixed'):
                continue          # 原始配置里的老字段不列，只看手动改过的
            iid = f'f{i}'
            self.tree.insert('', 'end', iid=iid, values=(
                int(f['seq']), (f.get('label') or '') + ('（副值）' if f.get('sub') else ''),
                '字段映射', f.get('src_col') or '', f'→ {f.get("to") or "(未指定)"}'))
            self._rows.append({'kind': 'f', 'obj': f, 'iid': iid})

        for i, e in enumerate(cfg.get('extra_columns') or []):
            if e.get('seq') is None:
                continue
            iid = f'x{i}'
            self.tree.insert('', 'end', iid=iid, values=(
                int(e['seq']), e.get('title') or e.get('label') or '',
                '新增列', e.get('src_col') or '',
                f'插在 {e.get("after") or "末尾"} 之后'))
            self._rows.append({'kind': 'x', 'obj': e, 'iid': iid})

        self.lbl.config(text=f'共 {len(self._rows)} 条改动'
                             f'（o=取值修正 {sum(1 for r in self._rows if r["kind"] == "o")}，'
                             f'f=字段映射 {sum(1 for r in self._rows if r["kind"] == "f")}，'
                             f'x=新增列 {sum(1 for r in self._rows if r["kind"] == "x")}）')

    def _del_sel(self):
        sel = self.tree.selection()
        if not sel:
            self.lbl.config(text='先在列表里选中要删的条目')
            return
        import fill_from_extract as engine
        n = 0
        for r in self._rows:
            if r['iid'] not in sel:
                continue
            o = r['obj']
            if r['kind'] == 'o':
                engine.remove_override(self.app.cfg_path, int(o['seq']),
                                       bool(o.get('sub')), o.get('label') or '')
            elif r['kind'] == 'f':
                engine.remove_field(self.app.cfg_path, int(o['seq']),
                                    bool(o.get('sub')), o.get('label'))
            else:
                engine.remove_extra_column(self.app.cfg_path, o.get('title'))
            n += 1
        self.app._reload(f'已删除 {n} 条改动')
        self.refresh()


class ReviewApp(tk.Tk):
    def __init__(self, pdf_path, middle, cfg=None, cfg_path=None):
        super().__init__()
        self.pdf_path = pdf_path
        self.middle = middle
        self.cfg = cfg or {}
        self.columns = engine.template_columns()
        self.pending_key = None
        # 防御：调用方要是把配置 dict 当路径传进来（老代码有过这种写法），
        # 这里直接丢掉，下面保存时会自动建一份可写的配置。
        self.cfg_path = cfg_path if isinstance(cfg_path, str) else None
        self.pages = middle.get('pages') or []
        self.page_i = 0
        self.zoom = 1.4
        self._fit_done = False
        self._img = None
        self._rect_of = {}          # tree iid -> canvas rect id
        self._block_of = {}         # tree iid -> block dict
        self._grid_cache = {}       # page_idx -> (data, bboxes)
        self._band = None           # 拖拽中的框 id
        self._band_xy = None
        self.show_unmapped = tk.BooleanVar(value=False)

        self.title(f'映射核对 — {os.path.basename(pdf_path)}')
        self.geometry('1520x940')
        self.minsize(1100, 700)
        self.configure(bg=C_BG)
        self._build()
        self._bind_keys()
        if self.pages:
            self.show_page(0)
        # 窗口还没完成布局时 canvas 宽度是无效值，取不到就退回默认 140%，
        # 结果是页面比画布宽、红框全跑到右侧看不见。所以等布局稳定后再适配一次。
        self.after(120, self._fit_once)

    def _fit_once(self, force=False):
        """按画布宽度适配缩放，保证整页（连同红框）都看得见。"""
        if self._fit_done and not force:
            return
        self.update_idletasks()
        page = self.cur_page()
        if not page:
            return
        if not self._fit_done:
            try:
                self.pane.sash_place(0, max(520, int(self.winfo_width() * 0.56)), 0)
            except Exception:
                pass
            self.update_idletasks()
        cw = self.canvas.winfo_width()
        if cw < 60:                      # 还没布局好，稍后再试
            self.after(150, self._fit_once)
            return
        pw = page.get('width_pt') or 595
        z = max(0.5, min(2.0, round((cw - 30) / pw, 2)))
        self._fit_done = True
        if abs(z - self.zoom) > 0.01:
            self.zoom = z
            self.show_page(self.page_i)

    # ---------------- 界面 ----------------
    def _build(self):
        # 顶部工具条
        bar = tk.Frame(self, bg=C_BG, padx=10, pady=6)
        bar.pack(side='top', fill='x')

        tk.Label(bar, text='页码', bg=C_BG, fg=C_TEXT, font=FONT_UI).pack(side='left')
        self.btn_prev = tk.Button(bar, text='◀ 上一页', font=FONT_SM, command=self.prev_page)
        self.btn_prev.pack(side='left', padx=(6, 2))
        self.lbl_page = tk.Label(bar, text='-/-', bg=C_BG, fg=C_TEXT, font=FONT_UI, width=9)
        self.lbl_page.pack(side='left')
        self.btn_next = tk.Button(bar, text='下一页 ▶', font=FONT_SM, command=self.next_page)
        self.btn_next.pack(side='left', padx=(2, 14))

        tk.Label(bar, text='缩放', bg=C_BG, fg=C_TEXT, font=FONT_UI).pack(side='left')
        tk.Button(bar, text='－', font=FONT_SM, width=3,
                  command=lambda: self.set_zoom(self.zoom - 0.2)).pack(side='left', padx=2)
        self.lbl_zoom = tk.Label(bar, text='140%', bg=C_BG, fg=C_TEXT, font=FONT_UI, width=6)
        self.lbl_zoom.pack(side='left')
        tk.Button(bar, text='＋', font=FONT_SM, width=3,
                  command=lambda: self.set_zoom(self.zoom + 0.2)).pack(side='left', padx=2)
        tk.Button(bar, text='适应宽度', font=FONT_SM,
                  command=lambda: self._fit_once(force=True)).pack(side='left', padx=(6, 0))

        tk.Checkbutton(bar, text='同时显示未映射字段（青绿框）', variable=self.show_unmapped,
                       bg=C_BG, fg=C_TEXT, font=FONT_SM, activebackground=C_BG,
                       command=self.redraw_boxes).pack(side='left', padx=16)

        tk.Button(bar, text='导出本页 PNG', font=FONT_SM,
                  command=self.export_png).pack(side='left', padx=(0, 8))
        self.btn_ov = tk.Button(bar, text='人工修正 (0)', font=FONT_SM,
                                command=self.show_overrides)
        self.btn_ov.pack(side='left', padx=(0, 16))

        tk.Label(bar, text='（没进表）', bg=C_BG, fg=C_TEAL,
                 font=FONT_SM).pack(side='right')
        tk.Label(bar, text='青绿框=未映射', bg=C_BG, fg=C_TEAL,
                 font=FONT_SM).pack(side='right')
        tk.Label(bar, text='　｜　', bg=C_BG, fg=C_MUTED,
                 font=FONT_SM).pack(side='right')
        tk.Label(bar, text='红框=已写入报价表', bg=C_BG, fg=C_RED,
                 font=FONT_SM).pack(side='right')

        # 主体：左右分栏
        pane = tk.PanedWindow(self, orient='horizontal', bg=C_BG, sashwidth=6, bd=0)
        pane.pack(side='top', fill='both', expand=True, padx=8, pady=(0, 8))
        self.pane = pane

        # 左：PDF 画布
        left = tk.Frame(pane, bg=C_BG)
        self.canvas = tk.Canvas(left, bg='#d9dde3', highlightthickness=0)
        vs = ttk.Scrollbar(left, orient='vertical', command=self.canvas.yview)
        hs = ttk.Scrollbar(left, orient='horizontal', command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.canvas.grid(row=0, column=0, sticky='nsew')
        vs.grid(row=0, column=1, sticky='ns')
        hs.grid(row=1, column=0, sticky='ew')
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)
        pane.add(left, minsize=520)

        # 右：内容面板
        right = tk.Frame(pane, bg=C_BG)
        pane.add(right, minsize=420)

        self.lbl_info = tk.Label(right, text='', bg=C_BG, fg=C_MUTED, font=FONT_SM,
                                 anchor='w', justify='left')
        self.lbl_info.pack(side='top', fill='x', pady=(0, 4))

        # 字段表
        box1 = tk.LabelFrame(right, text=' 提取到的映射内容 ', bg=C_BG, fg=C_TEXT,
                             font=FONT_UI, padx=4, pady=4)
        box1.pack(side='top', fill='both', expand=True)
        cols = ('seq', 'label', 'value', 'to')
        self.tree = ttk.Treeview(box1, columns=cols, show='headings', height=14)
        for c, t, w, a in (('seq', '序号', 52, 'center'),
                           ('label', '字段', 150, 'w'),
                           ('value', '提取值', 260, 'w'),
                           ('to', '写入列', 64, 'center')):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor=a, stretch=(c in ('label', 'value')))
        ts = ttk.Scrollbar(box1, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=ts.set)
        self.tree.pack(side='left', fill='both', expand=True)
        ts.pack(side='right', fill='y')
        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.tree.bind('<Double-1>', self.edit_block)

        # 页面文字（标题 / 正文）
        box2 = tk.LabelFrame(right, text=' 页面文字（标题 / 正文 / 备注） ', bg=C_BG,
                             fg=C_TEXT, font=FONT_UI, padx=4, pady=4)
        box2.pack(side='top', fill='both', expand=True, pady=(8, 0))
        self.txt = tk.Text(box2, height=10, wrap='word', font=FONT_SM,
                           bg=C_PAPER, fg=C_TEXT, relief='flat', padx=6, pady=4)
        tsc = ttk.Scrollbar(box2, orient='vertical', command=self.txt.yview)
        self.txt.configure(yscrollcommand=tsc.set)
        self.txt.pack(side='left', fill='both', expand=True)
        tsc.pack(side='right', fill='y')
        self.txt.tag_configure('title', foreground='#0b5cad', font=('Microsoft YaHei UI', 10, 'bold'))
        self.txt.tag_configure('muted', foreground=C_MUTED, font=FONT_SM)
        self.txt.configure(state='disabled')

        self.lbl_status = tk.Label(self, text='', bg='#e8ebef', fg=C_TEXT, font=FONT_SM,
                                   anchor='w', padx=10, pady=3)
        self.lbl_status.pack(side='bottom', fill='x')

    def _bind_keys(self):
        self.bind('<Left>', lambda e: self.prev_page())
        self.bind('<Right>', lambda e: self.next_page())
        self.bind('<Prior>', lambda e: self.prev_page())
        self.bind('<Next>', lambda e: self.next_page())
        self.canvas.bind('<Control-MouseWheel>',
                         lambda e: self.set_zoom(self.zoom + (0.2 if e.delta > 0 else -0.2)))
        self.canvas.bind('<MouseWheel>',
                         lambda e: self.canvas.yview_scroll(-1 if e.delta > 0 else 1, 'units'))
        # 人工框选：左键拖一个矩形
        self.canvas.bind('<ButtonPress-1>', self._band_start)
        self.canvas.bind('<B1-Motion>', self._band_move)
        self.canvas.bind('<ButtonRelease-1>', self._band_end)

    # ---------------- 人工框选 ----------------
    def _band_start(self, ev):
        self.canvas.delete('band')
        self._band_xy = (self.canvas.canvasx(ev.x), self.canvas.canvasy(ev.y))
        self._band = self.canvas.create_rectangle(
            *self._band_xy, *self._band_xy, outline=C_BLUE, width=2,
            dash=(4, 3), tags='band')

    def _band_move(self, ev):
        if not self._band or not self._band_xy:
            return
        x0, y0 = self._band_xy
        x1, y1 = self.canvas.canvasx(ev.x), self.canvas.canvasy(ev.y)
        self.canvas.coords(self._band, x0, y0, x1, y1)

    def _band_end(self, ev):
        if not self._band or not self._band_xy:
            return
        x0, y0 = self._band_xy
        x1, y1 = self.canvas.canvasx(ev.x), self.canvas.canvasy(ev.y)
        self.canvas.delete('band')
        self._band = self._band_xy = None
        if abs(x1 - x0) < 5 or abs(y1 - y0) < 5:
            return                      # 当成普通点击，不算框选
        self._pick_cell(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))

    def _grid(self, page_idx):
        """取这一页的单元格文本与坐标（与引擎同一套算法，保证框到哪就是哪）。"""
        if page_idx in self._grid_cache:
            return self._grid_cache[page_idx]
        import pdfplumber
        import fill_from_extract as engine
        with pdfplumber.open(self.pdf_path) as pdf:
            data, bxs = engine._table_cells(pdf.pages[page_idx])
        self._grid_cache[page_idx] = (data, bxs)
        return data, bxs

    def _seq_col(self, ncols):
        """这一页该用哪个列位表（有的版式首页与内页列数不同）。"""
        parse = self.cfg.get('parse') or {}
        cm = parse.get('col_map') or {}
        byn = parse.get('col_map_by_ncols') or {}
        if byn:
            cm = byn.get(str(ncols)) or cm
        return cm

    def _pick_cell(self, x0, y0, x1, y1):
        page = self.cur_page()
        if not page:
            return
        z = self.zoom
        # 画布像素 → PDF 点
        px0, py0, px1, py1 = x0 / z, y0 / z, x1 / z, y1 / z
        cx, cy = (px0 + px1) / 2, (py0 + py1) / 2
        try:
            data, bxs = self._grid(page['page_idx'])
        except Exception as e:
            self.lbl_status.config(text=f'读取表格结构失败: {type(e).__name__}: {e}')
            return
        if not data:
            return
        ncols = max(len(r) for r in data)
        cm = self._seq_col(ncols)
        seq_col = cm.get('seq')

        hit = None
        for ri, row_bx in enumerate(bxs or []):
            for ci, cell in enumerate(row_bx or []):
                if not cell:
                    continue
                if cell[0] <= cx <= cell[2] and cell[1] <= cy <= cell[3]:
                    hit = (ri, ci)
                    break
            if hit:
                break
        if not hit:
            self.lbl_status.config(text='框的位置不在表格格线上，请对着一个单元格拖。')
            return

        ri, ci = hit
        row = data[ri]
        text = (row[ci] or '').strip() if ci < len(row) else ''

        # 左右两栏并排的版式（本页 21 列）：同一行里左栏有自己的序号/字段名，
        # 右栏(执行机构/附件)另有一套 seq_r/name_r。框到右栏时必须用右栏列位，
        # 否则会把「类型Type=拨叉式」误认成左栏的「用途Service」。
        seq_r = cm.get('seq_r')
        is_right = (seq_r is not None and ci >= seq_r)
        use_seq = seq_r if is_right else seq_col
        use_name = cm.get('name_r') if is_right else cm.get('name')

        seq = (row[use_seq] or '').strip() if use_seq is not None and use_seq < len(row) else ''
        if not seq.isdigit():
            # 右栏没取到序号时，退回左栏再试一次（有的行只有左栏序号）
            if is_right and seq_col is not None:
                seq = (row[seq_col] or '').strip() if seq_col < len(row) else ''
                use_name = cm.get('name')
            if not seq.isdigit():
                self.lbl_status.config(
                    text=f'这一格所在行取不到序号（第 {ri + 1} 行），可能是表头或说明文字。')
                return
        # 空格子也允许：有的数据表某页这一行是空的、别的页有值，
        # 这时候用户框空格是在表达"这一项的值在这一列"，属于版式级规则。

        name = (row[use_name] if use_name is not None and use_name < len(row) else '') or ''
        name2 = (row[cm.get('name2')] if cm.get('name2') is not None else '') or ''
        if is_right:
            name2 = ''   # 右栏用 name_r 单一字段名，别混进左栏的 name2
        label = name2 if name2 and not name else name

        main_c, nor_c, sub_c = cm.get('val'), cm.get('val_nor'), cm.get('val2')

        # 看全文，不只看这一页：这一项在 24 台里主值列有几台有值、副值列有几台
        stat = (self.middle.get('seq_stats') or {}).get(str(seq)) or {}
        n_main, n_sub, n_all = stat.get('main', 0), stat.get('sub', 0), stat.get('total', 0)

        # 建议：优先信全文统计，其次看这一格落在哪一列
        if n_sub > n_main:
            recommend = '副值'
        elif n_main > n_sub:
            recommend = '主值'
        else:
            recommend = '副值' if (sub_c is not None and ci == sub_c) else '主值'
        guess_sub = (recommend == '副值')

        here = '（本页这格是空的）' if not text else f'「{text[:18]}」'
        if n_all:
            base = (f'全文 {n_all} 台里：主值列有值 {n_main} 台，副值列有值 {n_sub} 台'
                    f' → 建议「{recommend}」。你框的是第 {ci} 列{here}。')
        else:
            base = f'你框的是第 {ci} 列{here} → 建议「{recommend}」。'
        hint = base

        # 这一行所有可选前缀（fields + 备注 + 原始字段名），让人能改归属
        srcs = engine.label_sources_for_row(self.cfg, seq, name, name2)
        cands = []
        for s in srcs:
            tail = []
            if s.get('src'):
                tail.append(s['src'])
            if s.get('to'):
                tail.append('→' + s['to'])
            cands.append({'label': s['label'],
                          'display': s['label'] + ('（' + ' '.join(tail) + '）' if tail else '')})
        if not cands:
            cands = [{'label': label, 'display': label}]

        # 这一项是不是已经有修正了（按前缀精确匹配）
        ovs = (self.cfg.get('parse') or {}).get('overrides') or []
        chosen_label = None
        for c in cands:
            if c['label'] == label or label in c['label']:
                chosen_label = c['label']
                break
        chosen_label = chosen_label or cands[0]['label']
        existing, old_col = '', ''
        for o in ovs:
            if str(o.get('seq')) != str(seq):
                continue
            if o.get('label') and o.get('label') != chosen_label:
                continue
            existing = '副值' if o.get('sub') else '主值'
            old_col = o.get('col')
            break

        info = {'seq': int(seq), 'col': ci, 'text': text, 'label': chosen_label,
                'force_mode': 'key' if getattr(self, 'pending_value', None) else '',
                'columns': getattr(self, 'columns', []),
                'guess_sub': guess_sub, 'recommend': recommend, 'hint': hint,
                'existing': existing, 'old_col': old_col, 'candidates': cands}

        has = bool(existing)

        # 没有可写的配置文件时现场建一份：以前这里直接拒绝保存，
        # 用户框了半天等于白框。未知版式走的正是这条路。
        if not self.cfg_path:
            try:
                self.cfg_path = engine.ensure_config_file(self.pdf_path, self.cfg)
            except Exception as e:
                self.cfg_path = None
                self.lbl_status.config(text=f'自动建配置失败: {type(e).__name__}: {e}')
                return
            if not self.cfg_path:
                self.lbl_status.config(text='这份 PDF 解析不出字段，无法自动建配置。')
                return
            self.cfg = engine._load_config(self.cfg_path)
            self.lbl_status.config(
                text=f'已自动建配置：{os.path.basename(self.cfg_path)}，修正会写在这里。')

        dlg = PickDialog(self, info, has, on_ok=None, on_del=None)
        self.wait_window(dlg)
        if not dlg.result:
            return
        res = tuple(dlg.result) + (None,) * 8
        kind, s, c, sub, lab, mode, to, ins = res[:8]
        if kind == 'rekey':
            self.pending_value = {'col': c, 'sub': sub, 'to': to, 'ins': ins}
            self.lbl_status.config(
                text='已记住这一格的值。现在请框住正确的字段名（弹窗里已锁定“字段名”），'
                     '保存后就把这一格接到它下面。')
            return
        if kind == 'ok' and mode == 'key' and lab and getattr(self, 'pending_value', None):
            pv = self.pending_value
            try:
                act, desc = engine.upsert_field(
                    self.cfg_path, s, lab, to=pv.get('to') or None,
                    sub=pv.get('sub'), src_col=pv.get('col'),
                    insert_before=(pv.get('to') if (pv.get('to') and pv.get('ins'))
                                   else None))
            except Exception as e:
                self.lbl_status.config(text=f'写配置失败: {type(e).__name__}: {e}')
                return
            self.pending_value = None
            self.pending_key = None
            self.cfg = engine._load_config(self.cfg_path)
            self._reload(f'{act}：{desc}', seq=s, sub=pv.get('sub'))
            return
        if kind == 'ok' and mode == 'key' and lab:
            self.pending_key = {'seq': s, 'label': lab}
            self.lbl_status.config(
                text=f'已记住字段名「{lab}」（第 {s} 项）。请再框住它对应的值，'
                     f'在弹窗里选「主值/副值」。')
            return
        if kind == 'ok' and mode in ('value', 'sub') and lab:
            pk = self.pending_key or {'seq': s, 'label': lab}
            try:
                act, desc = engine.upsert_field(
                    self.cfg_path, pk['seq'], pk['label'],
                    to=to or None, sub=(mode == 'sub'), src_col=c,
                    insert_before=(to if (to and ins) else None))
            except Exception as e:
                self.lbl_status.config(text=f'写配置失败: {type(e).__name__}: {e}')
                return
            self.pending_key = None
            self.cfg = engine._load_config(self.cfg_path)
            self._reload(f'{act}：{desc}', seq=pk['seq'], sub=(mode == 'sub'))
            return
        # 不要在这里局部 import engine：Python 会把整个函数里的 engine
        # 都当成局部变量，导致前面第 612 行 label_sources_for_row 报
        # UnboundLocalError。模块顶部已导入，直接用即可。
        try:
            if kind == 'del':
                n = engine.remove_override(self.cfg_path, s, sub, lab)
                msg = f'已删除第 {s} 项「{lab}」的修正（剩 {n} 条）'
            else:
                _, n, added = engine.add_override(self.cfg_path, s, c, sub, lab)
                msg = (f'已写入 {os.path.basename(self.cfg_path)}：第 {s} 项「{lab}」'
                       f'按「{"副值" if sub else "主值"}」取第 {c} 列（共 {n} 条修正）')
                if added:
                    msg += f'｜顺手补了映射 → {added} 列'
        except Exception as e:
            self.lbl_status.config(text=f'写配置失败: {type(e).__name__}: {e}')
            return
        self._reload(msg, seq=s, sub=sub)

    def _reload(self, msg='', seq=None, sub=False):
        """改完配置后重解析并刷新界面。
        顺便报告这条修正对全文的影响（多少台有值了），让用户看到"是否复用"。"""
        import fill_from_extract as engine
        try:
            engine.clear_cache()
            cfg = engine._load_config(self.cfg_path)
            detail = engine.parse_pdf_detail(self.pdf_path, cfg)
            self.middle = engine.build_middle(self.pdf_path, cfg, detail)
            self.cfg = cfg
            self.pages = self.middle.get('pages') or []
            self._grid_cache.clear()
            self.show_page(min(self.page_i, len(self.pages) - 1))

            tail = '｜已重新解析并生效'
            if seq is not None:
                st = (self.middle.get('seq_stats') or {}).get(str(seq)) or {}
                n = st.get('sub' if sub else 'main', 0)
                tot = st.get('total', 0)
                tail = f'｜第 {seq} 项现在 {n}/{tot} 台有值'
            if self.cfg_path:
                tail += f'｜配置文件 {os.path.basename(self.cfg_path)}'
            self.lbl_status.config(text=msg + tail)
        except Exception as e:
            self.lbl_status.config(text=f'重解析失败: {type(e).__name__}: {e}')

    def show_overrides(self):
        if not self.cfg_path:
            self.lbl_status.config(text='当前没有可写的版式配置。')
            return
        dlg = OverrideListDialog(self, self)
        self.wait_window(dlg)
        self._sync_override_btn()

    def _sync_override_btn(self):
        n = len((self.cfg.get('parse') or {}).get('overrides') or [])
        try:
            self.btn_ov.config(text=f'人工修正 ({n})')
        except Exception:
            pass

    def export_png(self):
        page = self.cur_page()
        if not page:
            return
        out = os.path.splitext(self.pdf_path)[0] + f'_第{page["page_no"]}页_红线框.png'
        try:
            render_page_png(self.pdf_path, self.middle, page['page_no'], out, zoom=1.8)
            self.lbl_status.config(text=f'已导出：{os.path.basename(out)}')
        except Exception as e:
            self.lbl_status.config(text=f'导出失败: {type(e).__name__}: {e}')

    # ---------------- 页面切换 ----------------
    def cur_page(self):
        return self.pages[self.page_i] if 0 <= self.page_i < len(self.pages) else None

    def prev_page(self):
        if self.page_i > 0:
            self.show_page(self.page_i - 1)

    def next_page(self):
        if self.page_i < len(self.pages) - 1:
            self.show_page(self.page_i + 1)

    def set_zoom(self, z):
        z = max(0.6, min(3.0, round(z, 2)))
        if abs(z - self.zoom) < 1e-6:
            return
        self.zoom = z
        self.show_page(self.page_i)

    def show_page(self, i):
        if not self.pages:
            return
        self.page_i = max(0, min(i, len(self.pages) - 1))
        page = self.cur_page()

        self.lbl_page.config(text=f'{page["page_no"]}/{len(self.pages)}')
        self.lbl_zoom.config(text=f'{int(self.zoom * 100)}%')
        self.btn_prev.config(state='normal' if self.page_i > 0 else 'disabled')
        self.btn_next.config(text='下一页 ▶',
                             state='normal' if self.page_i < len(self.pages) - 1 else 'disabled')

        try:
            img = render_page_image(self.pdf_path, page['page_idx'], self.zoom)
        except Exception as e:
            self.lbl_status.config(text=f'页面渲染失败: {type(e).__name__}: {e}')
            return
        from PIL import ImageTk
        self._img = ImageTk.PhotoImage(img)
        self.canvas.delete('all')
        self.canvas.create_image(0, 0, anchor='nw', image=self._img, tags='page')
        self.canvas.configure(scrollregion=(0, 0, img.width, img.height))
        self.redraw_boxes()
        self.fill_tree(page)
        self.fill_text(page)
        self.canvas.yview_moveto(0)
        self.canvas.xview_moveto(0)

    # ---------------- 画框 ----------------
    def redraw_boxes(self):
        page = self.cur_page()
        if not page:
            return
        self.canvas.delete('box')
        z = self.zoom
        n_map = n_un = 0
        for b in page['blocks']:
            if b['type'] != 'field' or not b.get('bbox'):
                continue
            mapped = bool(b.get('mapped'))
            if mapped:
                n_map += 1
            else:
                n_un += 1
            if not mapped and not self.show_unmapped.get():
                continue
            x0, y0, x1, y1 = [v * z for v in b['bbox']]
            self.canvas.create_rectangle(x0, y0, x1, y1,
                                         outline=C_RED if mapped else C_TEAL,
                                         width=2, tags=('box', f'f{id(b)}'))
        n_ov = len((self.cfg.get('parse') or {}).get('overrides') or [])
        self._sync_override_btn()
        self.lbl_status.config(
            text=f'本页 {len(page["blocks"])} 个块｜红线框 {n_map} 处（已进报价表）'
                 f'｜未映射 {n_un} 处{"（青绿框）" if self.show_unmapped.get() else "（勾选上方选项可看）"}'
                 f'｜全文 {self.middle.get("valves", 0)} 台'
                 f'｜版式：{self.middle.get("config", "")}'
                 f'{"｜人工修正 " + str(n_ov) + " 条" if n_ov else ""}'
                 f'{"｜缓存" if self.middle.get("cached") else ""}')

    # ---------------- 右侧：字段清单 ----------------
    def fill_tree(self, page):
        self.tree.delete(*self.tree.get_children())
        self._rect_of.clear()
        self._block_of.clear()
        fields = [b for b in page['blocks'] if b['type'] == 'field']
        fields.sort(key=lambda b: (b.get('seq') or 0,
                                   0 if b.get('mapped') else 1))
        for i, b in enumerate(fields):
            if not b.get('mapped') and not self.show_unmapped.get():
                continue
            iid = f'f{i}'
            self.tree.insert('', 'end', iid=iid,
                             values=(b.get('seq', ''), b.get('label', ''),
                                     b.get('value', ''), b.get('to') or '—'))
            self._block_of[iid] = b
            if b.get('bbox'):
                self._rect_of[iid] = f'f{id(b)}'
        self.lbl_info.config(
            text=f'第 {page["page_no"]} 页｜字段 {len(fields)} 个'
                 f'（已映射 {sum(1 for b in fields if b.get("mapped"))}）')

    def edit_block(self, _evt=None):
        """双击右侧字段 → 直接改字段名/写入列（不用重新框选）。"""
        sel = self.tree.selection()
        if not sel:
            return
        b = self._block_of.get(sel[0])
        if not b:
            return
        top = tk.Toplevel(self)
        top.title('改这一项的映射')
        top.configure(bg=C_BG)
        top.transient(self)
        tk.Label(top, text=f'序号 {b.get("seq")}｜值：{str(b.get("value"))[:40]}',
                 bg=C_BG, fg=C_MUTED, font=FONT_SM).pack(anchor='w', padx=12, pady=(10, 4))

        tk.Label(top, text='字段名（可以直接改成你想要的名字）', bg=C_BG,
                 fg=C_TEXT, font=FONT_UI).pack(anchor='w', padx=12)
        var_lab = tk.StringVar(value=str(b.get('label') or ''))
        tk.Entry(top, textvariable=var_lab, width=48).pack(padx=12, pady=(2, 8))

        tk.Label(top, text='写入报价表哪一列', bg=C_BG, fg=C_TEXT,
                 font=FONT_UI).pack(anchor='w', padx=12)
        cols = getattr(self, 'columns', []) or []
        var_to = tk.StringVar(value=str(b.get('to') or ''))
        ttk.Combobox(top, textvariable=var_to, width=44, state='readonly',
                     values=['(不指定)'] + [f'{L}  {n}' for L, n in cols]).pack(
            padx=12, pady=(2, 8))

        def save():
            lab = var_lab.get().strip()
            to = var_to.get().strip()
            to = to.split()[0] if to and not to.startswith('(') else ''
            if not lab:
                return
            try:
                act, desc = engine.upsert_field(self.cfg_path, int(b['seq']), lab,
                                                to=to or None)
            except Exception as e:
                self.lbl_status.config(text=f'写配置失败: {type(e).__name__}: {e}')
                return
            top.destroy()
            self.cfg = engine._load_config(self.cfg_path)
            self._reload(f'{act}：{desc}', seq=int(b['seq']), sub=False)

        def delete():
            lab = str(b.get('label') or '').strip()
            seq = int(b['seq'])
            n = 0
            try:
                n += engine.remove_field(self.cfg_path, seq, False, lab or None)
                n += engine.remove_field(self.cfg_path, seq, True, lab or None)
                n += engine.remove_override(self.cfg_path, seq, False, lab) or 0
                if lab:
                    n += engine.remove_extra_column(self.cfg_path, lab[:24])
            except Exception as e:
                self.lbl_status.config(text=f'删除失败: {type(e).__name__}: {e}')
                return
            top.destroy()
            self.cfg = engine._load_config(self.cfg_path)
            self._reload(f'已删除第 {seq} 项「{lab}」的映射（{n} 处）', seq=seq, sub=False)

        bar = tk.Frame(top, bg=C_BG)
        bar.pack(fill='x', pady=(0, 12))
        tk.Button(bar, text='删除这条映射', fg=C_RED, command=delete).pack(
            side='left', padx=(12, 0))
        tk.Button(bar, text='保存', command=save).pack(side='right', padx=(6, 12))
        tk.Button(bar, text='取消', command=top.destroy).pack(side='right')
        top.grab_set()

    def on_select(self, _evt=None):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        b = self._block_of.get(iid)
        if not b or not b.get('bbox'):
            return
        z = self.zoom
        x0, y0, x1, y1 = [v * z for v in b['bbox']]
        self.canvas.delete('hl')
        self.canvas.create_rectangle(x0 - 2, y0 - 2, x1 + 2, y1 + 2,
                                     outline='#0050c8', width=3, tags='hl')
        h = max(1, self.canvas.winfo_height())
        total = max(1, int(self.canvas.cget('scrollregion').split()[3].split('.')[0])
                    if self.canvas.cget('scrollregion') else h)
        self.canvas.yview_moveto(max(0.0, (y0 - h * 0.3) / total))
        self.canvas.xview_moveto(0)

    # ---------------- 右侧：页面文字 ----------------
    def fill_text(self, page):
        self.txt.configure(state='normal')
        self.txt.delete('1.0', 'end')
        blocks = [b for b in page['blocks'] if b['type'] in ('title', 'text')]
        if not blocks:
            self.txt.insert('end', '（本页表格外没有可提取的文字）', 'muted')
        for b in blocks:
            tag = 'title' if b['type'] == 'title' else None
            label = '【标题】' if b['type'] == 'title' else '· '
            self.txt.insert('end', label + b['text'] + '\n', tag)
        self.txt.configure(state='disabled')


# ============================== 对外入口 ==============================

def open_review(pdf_path, cfg_path=None, cfg=None):
    """打开核对窗口。cfg_path/cfg 都不给时自动识别版式。

    窗口里可以直接拖框人工修正，修正会写回该版式的配置 JSON，
    下次遇到同版式的 PDF 自动生效。
    """
    import fill_from_extract as engine
    if not isinstance(cfg_path, str):        # 防御：被传进来的是配置 dict 而不是路径
        cfg_path = None
    if cfg is None:
        scores = []
        if not cfg_path:
            cfg_path, scores = engine.detect_layout(pdf_path)
        if cfg_path:
            cfg = engine._load_config(cfg_path)
        else:
            # 分数没到门槛 = 新版式。优先复用上次自动生成的配置：它带
            # fields，红框才画得出来；更重要的是它有真实路径，人工框选
            # 的修正才能写回 JSON（cfg_path=None 时窗口无法保存）。
            # 候选里没有就现场解析一份存盘。
            for s, p, name in scores or []:
                if os.path.basename(p).startswith('自动_'):
                    cfg_path = p
                    break
            if not cfg_path:
                # 连候选都没有：现场解析一份存进 mappings/，保证窗口能保存修正
                try:
                    cfg_path = engine.ensure_config_file(pdf_path)
                except Exception:
                    cfg_path = None
            if cfg_path:
                cfg = engine._load_config(cfg_path)
            else:
                cfg = {'id': 'generic', 'name': '通用兜底',
                       'parse': {'parse_mode': 'generic'}, 'fields': []}
    detail = engine.parse_pdf_detail(pdf_path, cfg)
    if not detail['valves']:
        raise RuntimeError('这份 PDF 没解析到阀门，无法核对。')
    middle = engine.build_middle(pdf_path, cfg, detail)
    app = ReviewApp(pdf_path, middle, cfg=cfg, cfg_path=cfg_path)
    app.mainloop()
    return middle


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python pdf_review.py <PDF> [配置.json]')
        sys.exit(1)
    open_review(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
