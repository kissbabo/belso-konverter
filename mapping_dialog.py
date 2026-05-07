# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk
from openpyxl.utils import get_column_letter, column_index_from_string
import re

BELSO_ROLES = [
    ('mennyiseg',  'Mennyiseg',      '#DCE6F1'),
    ('egyseg',     'Egyseg',         '#DCE6F1'),
    ('anyag_egy',  'Anyag egysar',   '#FFF3CD'),
    ('dij_egy',    'Dij egysar',     '#FFF3CD'),
    ('anyag_ossz', 'Anyag ossz',     '#E2EFDA'),
    ('dij_ossz',   'Dij ossz',       '#E2EFDA'),
    ('mindossz',   'Mindosszesen',   '#F3E5F5'),
]

BELSO_ROLE_LABELS = {
    'mennyiseg':  'Mennyiség',
    'egyseg':     'Egység',
    'anyag_egy':  'Anyag egységár',
    'dij_egy':    'Díj egységár',
    'anyag_ossz': 'Anyag összár',
    'dij_ossz':   'Díj összár',
    'mindossz':   'Mindösszesen  (A+D)',
}

def col_to_letter(col_idx):
    if col_idx is None:
        return ''
    return get_column_letter(col_idx)

def cell_ref_to_col(ref):
    if not ref or str(ref).strip() == '':
        return None
    ref = str(ref).strip().upper()
    m = re.match(r'([A-Z]+)', ref)
    if m:
        try:
            return column_index_from_string(m.group(1))
        except Exception:
            return None
    return None


class MappingDialog(tk.Toplevel):
    def __init__(self, parent, wb_src, auto_mappings):
        super().__init__(parent)
        self.title("Oszlop parositas ellenorzese")
        self.resizable(True, True)
        self.grab_set()
        self._wb      = wb_src
        self._auto    = auto_mappings
        self._result  = None
        self._widgets = {}
        self._build(parent)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _build(self, parent):
        pw = parent.winfo_width() or 700
        ph = parent.winfo_height() or 560
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        w = max(860, pw + 260)
        h = max(560, ph + 80)
        self.geometry(f"{w}x{h}+{max(0,px-130)}+{max(0,py-40)}")
        self.configure(bg="#F0F2F5")

        hdr = tk.Frame(self, bg="#1A3A5C", height=50)
        hdr.pack(fill='x')
        hdr.pack_propagate(False)
        tk.Label(hdr, text="Oszlop parositas ellenorzese",
                 bg="#1A3A5C", fg="white",
                 font=("Segoe UI", 11, "bold")).pack(side='left', padx=16, pady=12)
        tk.Label(hdr, text="Ellenorizd es modositsd, majd kattints az Elfogadas gombra",
                 bg="#1A3A5C", fg="#A0C4E8",
                 font=("Segoe UI", 8)).pack(side='right', padx=16)

        outer = tk.Frame(self, bg="#F0F2F5")
        outer.pack(fill='both', expand=True)

        canvas = tk.Canvas(outer, bg="#F0F2F5", highlightthickness=0)
        vsb = tk.Scrollbar(outer, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        inner = tk.Frame(canvas, bg="#F0F2F5")
        win_id = canvas.create_window((0, 0), window=inner, anchor='nw')

        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(win_id, width=e.width))
        canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))

        self._build_header(inner)

        SKIP = {'belso', 'foosszesito', 'osszesito', '00_elolap'}
        sheets = [s for s in self._wb.sheetnames
                  if not any(k in s.lower().replace('ő','o').replace('ö','o').replace('é','e').replace('í','i').replace('á','a').replace('ü','u') for k in SKIP)]
        for i, sname in enumerate(sheets):
            bg = "#FFFFFF" if i % 2 == 0 else "#F7F8FA"
            self._build_row(inner, sname, bg)

        btn_bar = tk.Frame(self, bg="#E8EDF2", height=54)
        btn_bar.pack(fill='x', side='bottom')
        btn_bar.pack_propagate(False)

        tk.Label(btn_bar, text="Sarga mezok szerkeszthetok. Ures = automatikus.",
                 bg="#E8EDF2", fg="#595959",
                 font=("Segoe UI", 8)).pack(side='left', padx=14)

        tk.Button(btn_bar, text="Megse", command=self._on_cancel,
                  font=("Segoe UI", 9), bg="#E8EDF2", relief='flat',
                  padx=12, pady=7, cursor='hand2').pack(side='right', padx=8, pady=8)

        tk.Button(btn_bar, text="  Elfogadas es konvertals  ",
                  command=self._on_ok,
                  font=("Segoe UI", 10, "bold"),
                  bg="#1A6B3C", fg="white", relief='flat',
                  padx=16, pady=7, cursor='hand2').pack(side='right', pady=8)

    def _build_header(self, parent):
        hf = tk.Frame(parent, bg="#1F497D")
        hf.pack(fill='x', padx=6, pady=(6, 0))
        tk.Label(hf, text="Munkalap neve",
                 bg="#1F497D", fg="white", font=("Segoe UI", 9, "bold"),
                 anchor='w', padx=8, width=24).grid(
            row=0, column=0, sticky='nsew', padx=1, pady=2)
        tk.Label(hf, text="Javasolt oszlop (auto)",
                 bg="#2C6FAC", fg="white", font=("Segoe UI", 8, "bold"),
                 anchor='center', width=18).grid(
            row=0, column=1, sticky='nsew', padx=1, pady=2)
        tk.Label(hf, text="Elfogadott oszlop  (modosithato  v)",
                 bg="#1E6B3C", fg="white", font=("Segoe UI", 8, "bold"),
                 anchor='center', width=24).grid(
            row=0, column=2, sticky='nsew', padx=1, pady=2)
        hf.columnconfigure(0, weight=3)
        hf.columnconfigure(1, weight=1)
        hf.columnconfigure(2, weight=2)

    def _build_row(self, parent, sname, bg):
        """
        Vertikalis elrendezes: a munkalap neve + minden szerepkor egy sorban.
        Oszlopok: [Munkalap / szerepkor neve] | [Javasolt] | [Elfogadott combo]
        """
        auto = self._auto.get(sname, {})
        ws   = self._wb[sname]
        hr   = auto.get('header_row', 1)

        col_options = ['']
        for c in range(1, ws.max_column + 1):
            hval = ws.cell(hr, c).value
            lbl  = get_column_letter(c)
            if hval and str(hval).strip():
                lbl += f"  ({str(hval).strip()[:20]})"
            col_options.append(lbl)

        # Keretlap: lapnev fejlec + szerepkori sorok
        outer = tk.Frame(parent, bg=bg, highlightthickness=1,
                         highlightbackground="#D0D8E4")
        outer.pack(fill='x', padx=6, pady=3)

        # Lapnev fejlec sor
        name_row = tk.Frame(outer, bg="#E8EDF7")
        name_row.pack(fill='x')
        tk.Label(name_row, text=sname,
                 bg="#E8EDF7", fg="#1A3A5C",
                 font=("Segoe UI", 9, "bold"),
                 anchor='w', padx=10, pady=4).pack(side='left')

        self._widgets[sname] = {}

        for i, (role, _, row_bg) in enumerate(BELSO_ROLES):
            label    = BELSO_ROLE_LABELS[role]
            col_idx  = auto.get(role)
            javasolt = col_to_letter(col_idx) if col_idx else '-'
            rb       = "#FAFCFF" if i % 2 == 0 else "#F3F7FF"

            role_row = tk.Frame(outer, bg=rb)
            role_row.pack(fill='x')
            role_row.columnconfigure(0, weight=3)
            role_row.columnconfigure(1, weight=1)
            role_row.columnconfigure(2, weight=2)

            # Szerepkor neve
            tk.Label(role_row, text=f"   {label}",
                     bg=rb, fg="#333333",
                     font=("Segoe UI", 8),
                     anchor='w', width=28).grid(
                row=0, column=0, sticky='w', padx=(16,4), pady=3)

            # Javasolt (olvasható, kek)
            tk.Label(role_row, text=javasolt,
                     bg="#DCE6F1", fg="#1A3A5C",
                     font=("Consolas", 9, "bold"),
                     width=8, anchor='center',
                     relief='flat', padx=4).grid(
                row=0, column=1, sticky='ew', padx=4, pady=3)

            # Elfogadott combo (sarga, szerkesztheto)
            var = tk.StringVar(value=javasolt if javasolt != '-' else '')
            combo = ttk.Combobox(role_row, textvariable=var,
                                  values=col_options,
                                  width=22,
                                  font=("Segoe UI", 9))
            combo.grid(row=0, column=2, sticky='ew', padx=(4,8), pady=3)

            self._widgets[sname][role] = var

    def _on_ok(self):
        result = {}
        for sname, roles in self._widgets.items():
            sheet_map = {}
            for role, var in roles.items():
                val = var.get().strip()
                sheet_map[role] = cell_ref_to_col(val)
            result[sname] = sheet_map
        self._result = result
        self.destroy()

    def _on_cancel(self):
        self._result = None
        self.destroy()

    def get_result(self):
        return self._result


def show_mapping_dialog(parent, wb_src, auto_mappings):
    style = ttk.Style(parent)
    style.configure('Sarga.TCombobox', fieldbackground='#FFFF99', background='#FFFF99')
    dlg = MappingDialog(parent, wb_src, auto_mappings)
    parent.wait_window(dlg)
    return dlg.get_result()
