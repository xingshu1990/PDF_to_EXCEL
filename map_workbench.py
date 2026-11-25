# -*- coding: utf-8 -*-
"""
映射工作台（批量映射）
=====================================================================
左边 PDF（红框=已映射、青绿=未映射），右边"字段名 → 写入列"表格，两边联动：
  · 点表格某一行 → 左边高亮对应那一格；
  · 点左边某一格 → 右边选中对应行；
  · 多选几行 → 上方下拉选模板列 → 「应用到选中行」，一屏定几十项；
  · 「保存到版式配置」写 mappings/*.json；「保存并生成报价表」顺带出表。

用法（从「批量映射.bat」进来）：
    python map_workbench.py <PDF> [配置.json]
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import tkinter as tk
from tkinter import ttk, messagebox

import fill_from_extract as engine

C_BG = '#f5f6f8'
C_TEXT = '#222222'
C_MUTED = '#777777'
C_RED = '#e02020'
C_TEAL = '#00897b'
C_BLUE = '#0050c8'
C_OK = '#1a7f37'
C_TODO = '#b26a00'
FONT_UI = ('Microsoft YaHei UI', 10)
FONT_SM = ('Microsoft YaHei UI', 9)


def template_columns(tpl):
    """模板列清单：[(列字母, 表头文字)]。"""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    out = []
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


class Row:
    __slots__ = ('seq', 'label', 'value', 'to', 'sub', 'src')

    def __init__(self, seq, label, value, to='', sub=0, src=None):
        self.seq, self.label, self.value = seq, label, value
        self.to, self.sub, self.src = to, sub, src


class Workbench(tk.Tk):
    def __init__(self, pdf_path, cfg_path=None):
        super().__init__()
        self.pdf = pdf_path
        self.cfg_path = cfg_path
        self.mid = {'pages': []}
        self.pages = []
        self.blocks = []          # 当前页的字段块
        self.rows = []
        self.zoom = 1.3
        self._img = None
        self._rect_of = {}        # row iid -> canvas rect id
        self.dirty = set()        # 本次手动改过的行（只有它们会被写入）
        self.tpl = engine.default_template()
        self.columns = template_columns(self.tpl) if self.tpl else []
        self.title(f'映射工作台 — {os.path.basename(pdf_path)}')
        self.geometry('1560x900')
        self.configure(bg=C_BG)
        self._build()
        self.reload()

    # ---------------- 界面 ----------------
    def _build(self):
        pane = tk.PanedWindow(self, orient='horizontal', bg=C_BG, sashwidth=6, bd=0)
        pane.pack(fill='both', expand=True, padx=6, pady=(6, 0))
        left = tk.Frame(pane, bg=C_BG)
        right = tk.Frame(pane, bg=C_BG)
        pane.add(left, minsize=620)
        pane.add(right, minsize=560)

        # —— 左：PDF ————
        bar = tk.Frame(left, bg=C_BG)
        bar.pack(side='top', fill='x')
        tk.Button(bar, text='◀', font=FONT_SM, width=3, command=self.prev_page).pack(side='left')
        self.lbl_page = tk.Label(bar, text='-/-', bg=C_BG, fg=C_TEXT, font=FONT_UI, width=10)
        self.lbl_page.pack(side='left')
        tk.Button(bar, text='▶', font=FONT_SM, width=3, command=self.next_page).pack(side='left')
        self.var_valve = tk.StringVar()
        self.cbo_valve = ttk.Combobox(bar, textvariable=self.var_valve, width=16,
                                      state='readonly')
        self.cbo_valve.pack(side='left', padx=(10, 0))
        self.cbo_valve.bind('<<ComboboxSelected>>', lambda e: self.switch_valve())
        tk.Label(bar, text='红框=已映射  青绿=未映射', bg=C_BG, fg=C_MUTED,
                 font=FONT_SM).pack(side='right')

        cbox = tk.Frame(left, bg='#d9dde3')
        cbox.pack(side='top', fill='both', expand=True, pady=(4, 0))
        self.canvas = tk.Canvas(cbox, bg='#d9dde3', highlightthickness=0)
        vs = ttk.Scrollbar(cbox, orient='vertical', command=self.canvas.yview)
        hs = ttk.Scrollbar(cbox, orient='horizontal', command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.canvas.grid(row=0, column=0, sticky='nsew')
        vs.grid(row=0, column=1, sticky='ns')
        hs.grid(row=1, column=0, sticky='ew')
        cbox.rowconfigure(0, weight=1)
        cbox.columnconfigure(0, weight=1)
        self.canvas.bind('<Button-1>', self.on_canvas_click)

        # —— 右：映射表 ————
        top = tk.Frame(right, bg=C_BG, padx=6, pady=4)
        top.pack(side='top', fill='x')
        tk.Label(top, text='模板列：', bg=C_BG, fg=C_TEXT, font=FONT_UI).pack(side='left')
        self.var_col = tk.StringVar()
        self.cbo = ttk.Combobox(top, textvariable=self.var_col, width=34, state='readonly',
                                values=[f'{L}  {n}' for L, n in self.columns])
        self.cbo.pack(side='left', padx=(0, 4))
        tk.Button(top, text='应用到选中行', font=FONT_SM,
                  command=self.apply_to_selected).pack(side='left')
        tk.Button(top, text='清空', font=FONT_SM,
                  command=self.clear_selected).pack(side='left', padx=(4, 0))
        self.var_insert = tk.BooleanVar(value=True)
        tk.Checkbutton(top, text='插到该列前面（不覆盖）', variable=self.var_insert,
                       bg=C_BG, fg=C_TEXT, font=FONT_SM,
                       activebackground=C_BG).pack(side='left', padx=(10, 0))

        top2 = tk.Frame(right, bg=C_BG, padx=6)
        top2.pack(side='top', fill='x')
        tk.Button(top2, text='只显示未映射', font=FONT_SM,
                  command=self.toggle_filter).pack(side='left')
        tk.Button(top2, text='按字段名自动对齐', font=FONT_SM,
                  command=self.auto_align).pack(side='left', padx=(4, 0))
        tk.Button(top2, text='打开核对窗口', font=FONT_SM,
                  command=self.open_review).pack(side='left', padx=(4, 0))
        tk.Button(top2, text='保存并生成报价表', font=FONT_SM, default='active',
                  command=lambda: self.save(True)).pack(side='right')
        tk.Button(top2, text='保存到版式配置', font=FONT_SM,
                  command=lambda: self.save(False)).pack(side='right', padx=(0, 6))

        box = tk.Frame(right, bg=C_BG)
        box.pack(side='top', fill='both', expand=True, padx=6, pady=(6, 0))
        cols = ('seq', 'label', 'value', 'to')
        self.tree = ttk.Treeview(box, columns=cols, show='headings', selectmode='extended')
        for c, t, w, a in (('seq', '序号', 50, 'center'), ('label', '字段名', 250, 'w'),
                           ('value', '值', 210, 'w'), ('to', '写入列', 110, 'center')):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor=a, stretch=(c in ('label', 'value')))
        vs2 = ttk.Scrollbar(box, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs2.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vs2.pack(side='right', fill='y')
        self.tree.tag_configure('done', foreground=C_OK)
        self.tree.tag_configure('todo', foreground=C_TODO)
        self.tree.bind('<<TreeviewSelect>>', lambda e: self.on_row_select())
        self.tree.bind('<Double-1>', lambda e: self.apply_to_selected())

        self.lbl = tk.Label(self, text='', bg='#e8ebef', fg=C_TEXT, font=FONT_SM,
                            anchor='w', padx=10, pady=4)
        self.lbl.pack(side='bottom', fill='x', pady=(6, 0))

    # ---------------- 数据 ----------------
    def reload(self):
        if self.cfg_path and os.path.isfile(self.cfg_path):
            self.cfg = engine._load_config(self.cfg_path)
        else:
            chosen, _ = engine.detect_layout(self.pdf)
            self.cfg_path = chosen or engine.ensure_config_file(self.pdf)
            self.cfg = engine._load_config(self.cfg_path) if self.cfg_path else {}
        self.valves = engine.parse_pdf_detail(self.pdf, self.cfg)['valves']
        self.mid = engine.build_middle(self.pdf, self.cfg)
        self.pages = self.mid.get('pages') or []
        self.vi = 0
        self.cbo_valve.config(values=[f'阀门{i+1}  {v["tag"]}'
                                      for i, v in enumerate(self.valves)])
        if self.valves:
            self.var_valve.set(f'阀门1  {self.valves[0]["tag"]}')
        self.build_rows()
        self.show_page(0)

    def build_rows(self):
        if not self.valves:
            self.rows = []
            return
        v = self.valves[min(self.vi, len(self.valves) - 1)]
        cur = {}
        for item in self.cfg.get('fields') or []:
            if item.get('to') and item.get('seq') is not None:
                cur[(int(item['seq']), int(item.get('sub', 0)))] = item['to']
        rows = []
        for seq, f in sorted(v['fields'].items(), key=lambda kv: int(kv[0])):
            if f.get('main') or f.get('name'):
                rows.append(Row(int(seq), str(f.get('name') or ''), str(f.get('main') or ''),
                                cur.get((int(seq), 0), ''), 0, f.get('col')))
            if f.get('sub') or f.get('name2'):
                rows.append(Row(int(seq), str(f.get('name2') or f.get('name') or ''),
                                str(f.get('sub') or ''), cur.get((int(seq), 1), ''), 1,
                                f.get('sub_col')))
        self.rows = rows
        self.filter_empty = getattr(self, 'filter_empty', False)
        self.refresh()

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(self.rows):
            if getattr(self, 'filter_empty', False) and r.to:
                continue
            self.tree.insert('', 'end', iid=str(i),
                             values=(r.seq, r.label[:52], r.value[:52], r.to or ''),
                             tags=('done' if r.to else 'todo',))
        self.n_map = sum(1 for r in self.rows if r.to)
        self._status(f'{len(self.rows)} 行｜已映射 {self.n_map}｜配置：'
                     f'{os.path.basename(self.cfg_path or "(无)")}')

    # ---------------- PDF 侧 ----------------
    def switch_valve(self):
        s = self.var_valve.get()
        try:
            self.vi = int(s.split('阀门')[1].split()[0]) - 1
        except Exception:
            self.vi = 0
        self.build_rows()
        v = self.valves[self.vi] if self.valves else None
        if v:
            self.show_page(v.get('page_idx', 0))

    def show_page(self, idx):
        if not self.pages:
            return
        idx = max(0, min(idx, len(self.pages) - 1))
        page = self.pages[idx]
        self.page_i = idx
        self.lbl_page.config(text=f'{page["page_no"]}/{len(self.pages)}')
        try:
            import pdf_review
            img = pdf_review.render_page_image(self.pdf, page['page_idx'], self.zoom)
        except Exception as e:
            self._status(f'页面渲染失败: {e}'); return
        from PIL import ImageTk
        self._img = ImageTk.PhotoImage(img)
        self.canvas.delete('all')
        self.canvas.create_image(0, 0, anchor='nw', image=self._img, tags='page')
        self.canvas.configure(scrollregion=(0, 0, img.width, img.height))
        self.blocks = [b for b in page['blocks'] if b['type'] == 'field' and b.get('bbox')]
        self.draw_boxes()
        self.fill_left(page)

    def draw_boxes(self):
        self.canvas.delete('box')
        z = self.zoom
        for b in self.blocks:
            x0, y0, x1, y1 = [v * z for v in b['bbox']]
            color = C_RED if b.get('mapped') else C_TEAL
            self.canvas.create_rectangle(x0, y0, x1, y1, outline=color, width=2, tags='box')

    def fill_left(self, page):
        pass

    def on_canvas_click(self, ev):
        x, y = self.canvas.canvasx(ev.x), self.canvas.canvasy(ev.y)
        z = self.zoom
        hit = None
        for b in self.blocks:
            x0, y0, x1, y1 = [v * z for v in b['bbox']]
            if x0 <= x <= x1 and y0 <= y <= y1:
                hit = b
                break
        if not hit:
            return
        for i, r in enumerate(self.rows):
            if r.seq == hit.get('seq') and (r.value == hit.get('value')
                                            or r.label[:10] == str(hit.get('label'))[:10]):
                self.tree.selection_set(str(i))
                self.tree.see(str(i))
                return

    def on_row_select(self):
        sel = self.tree.selection()
        if not sel:
            return
        r = self.rows[int(sel[0])]
        self.canvas.delete('hl')
        z = self.zoom
        for b in self.blocks:
            if b.get('seq') != r.seq:
                continue
            if r.value and b.get('value') != r.value and str(b.get('label'))[:8] != r.label[:8]:
                continue
            x0, y0, x1, y1 = [v * z for v in b['bbox']]
            self.canvas.create_rectangle(x0 - 2, y0 - 2, x1 + 2, y1 + 2,
                                         outline=C_BLUE, width=3, tags='hl')
            h = max(1, self.canvas.winfo_height())
            total = max(1, int(float(str(self.canvas.cget('scrollregion')).split()[3])))
            self.canvas.yview_moveto(max(0.0, (y0 - h * 0.3) / total))
            break

    def prev_page(self):
        self.show_page(getattr(self, 'page_i', 0) - 1)

    def next_page(self):
        self.show_page(getattr(self, 'page_i', 0) + 1)

    # ---------------- 操作 ----------------
    def _picked_col(self):
        s = self.var_col.get().strip()
        return s.split()[0] if s else ''

    def _selected(self):
        return [int(i) for i in self.tree.selection()]

    def apply_to_selected(self):
        col = self._picked_col()
        rows = self._selected()
        if not col:
            self._status('先在上面选一个模板列'); return
        if not rows:
            self._status('先在表格里选行（可 Ctrl/Shift 多选）'); return
        for i in rows:
            self.rows[i].to = col
            self.dirty.add(i)
        self.refresh()
        self.tree.selection_set([str(i) for i in rows])
        self._status(f'已给 {len(rows)} 行指定 → {col}（别忘了保存）')

    def clear_selected(self):
        for i in self._selected():
            self.rows[i].to = ''
            self.dirty.add(i)
        self.refresh()

    def toggle_filter(self):
        self.filter_empty = not getattr(self, 'filter_empty', False)
        self.refresh()

    def auto_align(self):
        try:
            import semantic_map as S
        except Exception as e:
            self._status(f'语义模块不可用: {e}'); return
        headers, _ = S.read_template_headers(self.tpl)
        mem = S.load_memory()
        res = S.suggest([r.label for r in self.rows], headers, memory=mem)
        n = 0
        for r, x in zip(self.rows, res):
            if not r.to and x.get('to'):
                r.to = x['to']; n += 1
        self.refresh()
        self._status(f'自动对齐补了 {n} 行；没把握的留空，请手动指定')

    def open_review(self):
        # 用子进程打开核对窗口：同一个进程里开两个 Tk 根窗口会撞
        # TclError: image "pyimageN" doesn't exist。
        import subprocess
        script = os.path.join(HERE, 'pdf_review.py')
        args = [sys.executable, script, self.pdf]
        if self.cfg_path:
            args.append(self.cfg_path)
        try:
            # 必须用 Popen（非阻塞）：run() 会卡住 Tk 的事件循环，工作台就假死
            self._proc = subprocess.Popen(args)
        except Exception as e:
            messagebox.showerror('打开失败', f'{type(e).__name__}: {e}')
            return
        self._status('核对窗口已打开（工作台可以继续用）；关窗后会自动刷新配置')
        self._poll_review()

    def _poll_review(self):
        proc = getattr(self, '_proc', None)
        if proc is None:
            return
        if proc.poll() is None:
            self.after(400, self._poll_review)
            return
        self._proc = None
        self.reload()
        self._status('核对窗口已关闭；配置已重新加载')

    def save(self, make_xlsx=False):
        if not self.cfg_path:
            self._status('没有可写的版式配置'); return
        import json
        cfg = engine._load_config(self.cfg_path)
        fields = cfg.get('fields') or []
        index = {}
        for i, f in enumerate(fields):
            if f.get('to') and f.get('seq') is not None:
                index[(int(f['seq']), int(f.get('sub', 0)))] = i
        added = updated = inserts = 0
        extras = cfg.get('extra_columns') or []
        by_title = {e.get('title'): e for e in extras}
        from openpyxl.utils import get_column_letter, column_index_from_string
        for ri, r in enumerate(self.rows):
            if ri not in self.dirty or not r.to:
                continue
            label = (r.label or '').strip()
            if self.var_insert.get():
                # 插到所选列前面：Excel 里就是在 D 前插一列，老 D 往后挪
                idx = column_index_from_string(r.to)
                after = get_column_letter(idx - 1) if idx > 1 else None
                item = {'title': label[:24] or f'seq{r.seq}', 'label': label, 'seq': r.seq}
                if r.src is not None:
                    item['src_col'] = r.src
                if after:
                    item['after'] = after
                if not (r.value or r.src):
                    item['placeholder'] = True
                key_t = item['title']
                if key_t in by_title:
                    by_title[key_t].update(item)
                else:
                    extras.append(item); by_title[key_t] = item
                inserts += 1
                continue
            key = (r.seq, r.sub)
            item = {'seq': r.seq, 'to': r.to, 'label': label}
            pat = engine._field_pat_from_name(label)
            if pat:
                item['pat'] = pat
            if r.sub:
                item['sub'] = 1
            if key in index:
                fields[index[key]].update(item); updated += 1
            else:
                fields.append(item); index[key] = len(fields) - 1; added += 1
        cfg['fields'] = fields
        if self.var_insert.get():
            cfg['extra_columns'] = extras
        self.dirty.clear()
        with open(self.cfg_path, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        engine.clear_cache()
        msg = (f'已保存到 {os.path.abspath(self.cfg_path)}：新增 {added} 项、更新 {updated} 项、插入新列 {inserts} 项')
        if make_xlsx:
            out = os.path.splitext(self.pdf)[0] + '_报价表.xlsx'
            try:
                rep = engine.fill_pdf(self.pdf, self.tpl, out, config_path=self.cfg_path)
                self._status(f'已生成 {os.path.basename(out)}：{rep["filled_rows"]} 台、'
                             f'映射 {rep["mapped_fields"]} 项')
            except Exception as e:
                messagebox.showerror('生成失败', f'{type(e).__name__}: {e}')
            self._status(msg + '｜' + self.lbl.cget('text'))
        else:
            self.reload()
            self._status(msg + '  ← 已写入版式配置，同版式以后自动生效')
            messagebox.showinfo('已保存', msg)

    def _status(self, s):
        self.lbl.config(text=s)


def open_workbench(pdf_path, cfg_path=None):
    app = Workbench(pdf_path, cfg_path)
    app.mainloop()


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python map_workbench.py <PDF> [配置.json]')
        sys.exit(1)
    open_workbench(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
