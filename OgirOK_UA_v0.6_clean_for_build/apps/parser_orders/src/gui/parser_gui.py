import os, sys, json, time, zipfile, re, hashlib
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import tkinter.font as tkfont

# Ensure project root is on sys.path (so `import apps...` works even when running as a script)
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../../.."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from apps.parser_orders.src.core.excel_autofill import autofill_template
from apps.parser_orders.src.core.diagnostics import load_unknowns
from apps.parser_orders.src.core.docx_parser import build_combined_v2
from apps.parser_orders.src.core.category_engine import CategoryEngine, make_regex_from_phrase
from apps.parser_orders.src.core.os_utils import open_path

BASE_DIR = PROJECT_ROOT
# In PyInstaller build (frozen), use directory of the executable as base.
if getattr(sys, 'frozen', False):
    try:
        BASE_DIR = os.path.dirname(sys.executable)
    except Exception:
        BASE_DIR = PROJECT_ROOT
TEMPLATES = os.path.join(BASE_DIR, "templates")
CONFIG = os.path.join(BASE_DIR, "config")
OUTPUT = os.path.join(BASE_DIR, "output")
DOCS_DIR = os.path.join(BASE_DIR, "input", "documents")


# Special categories (driven by headings) -> UI tags/colors
SPECIAL_CATEGORY_TAGS = {
    "Прикомандировано до медичної роти (в/ч А0989)": "SPECIAL_MED",
    "З щорічної основної відпустки": "SPECIAL_VACATION",
    "У відрядження": "SPECIAL_TRIP",
}

def _ui_tag_for_event_and_category(event_type: str, category: str | None) -> tuple:
    """Return a Treeview/Text tag based on category (preferred) and event type."""
    if category:
        t = SPECIAL_CATEGORY_TAGS.get(str(category).strip())
        if t:
            return (t,)
    et = (event_type or "").strip().lower()
    if et == "arrival":
        return ("ARRIVAL",)
    if et == "departure":
        return ("DEPARTURE",)
    return ()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ОгрОК_UA — Парсер наказів (DOCX → Excel)")
        self.geometry("1100x760")

        # App icon (assets/icon.png). Keep a reference in self to avoid GC.
        try:
            _ip = os.path.join(BASE_DIR, "assets", "icon.png")
            if os.path.exists(_ip):
                self._app_icon = tk.PhotoImage(file=_ip)
                self.iconphoto(True, self._app_icon)
        except Exception:
            pass


        # UI settings (font + highlight colors)
        self.ui_settings_path = os.path.join(CONFIG, "ui_settings.json")
        self.ui_settings = self._load_ui_settings()
        self._apply_global_style()

        self.template_path = os.path.join(TEMPLATES, "N_fixed.xlsx")
        self.rules_path = os.path.join(CONFIG, "autofill", "passport_rules.json")
        self.mapping_path = os.path.join(CONFIG, "autofill", "mapping.json")
        self.template_sources_path = os.path.join(CONFIG, "autofill", "template_sources.json")
        try:
            from apps.parser_orders.src.core.excel_autofill import ensure_mapping_file
            ensure_mapping_file(self.mapping_path, self.template_sources_path)
        except Exception:
            # mapping may be user-managed; do not block GUI startup
            pass
        self.learning_dir = os.path.join(CONFIG, "learning")
        self.categories_dir = os.path.join(CONFIG, "categories")
        os.makedirs(self.categories_dir, exist_ok=True)
        self.category_rules_path = os.path.join(self.categories_dir, "event_categories_rules_v1.json")
        self.user_category_rules_path = os.path.join(self.categories_dir, "user_category_rules.json")
        self._cat_engine = None
        # Hidden people (remove from list)
        self.ignored_people_path = os.path.join(CONFIG, "ignored_people.json")
        self._ignored_people = self._load_ignored_people()

        self.diag_path = os.path.join(OUTPUT, "diagnostics", "unknowns_latest.json")
        self.output_dir = os.path.join(OUTPUT, "exports", "excel")
        self.parsed_dir = os.path.join(OUTPUT, "parsed")
        os.makedirs(self.parsed_dir, exist_ok=True)
        self.combined_path = os.path.join(self.parsed_dir, "combined_v2.json")
        self.exports_txt_dir = os.path.join(OUTPUT, "exports", "txt")
        self.exports_json_dir = os.path.join(OUTPUT, "exports", "json")
        os.makedirs(self.exports_txt_dir, exist_ok=True)
        os.makedirs(self.exports_json_dir, exist_ok=True)
        os.makedirs(self.output_dir, exist_ok=True)

        self.selected_path = tk.StringVar(value="")
        self.last_excel = None
        self.last_json = None

        # People tab caches
        self._docx_by_name_cache = {}  # doc_name -> full path
        self._current_person_pib = None
        self._current_doc_filter = None  # doc_name or None
        self._current_event_filter = None  # 'arrival'/'departure'/None
        self._current_category_filter = None  # category label or None
        self._mention_pos = {}  # (doc, block_id) -> dict(text indices)

        self._ensure_category_engine()


        nb = ttk.Notebook(self)
        self.nb = nb
        nb.pack(fill="both", expand=True)

        self.tab_input = ttk.Frame(nb)
        self.tab_run = ttk.Frame(nb)
        self.tab_results = ttk.Frame(nb)
        self.tab_quality = ttk.Frame(nb)
        self.tab_settings = ttk.Frame(nb)
        self.tab_people = ttk.Frame(nb)

        nb.add(self.tab_input, text="1) Вхід")
        nb.add(self.tab_run, text="2) Обробка")
        nb.add(self.tab_results, text="3) Результати")
        nb.add(self.tab_people, text="👤 Люди (A–Я)")
        nb.add(self.tab_quality, text="🧩 Правила парсера")
        nb.add(self.tab_settings, text="⚙ Налаштування")

        self._build_input()
        self._build_run()
        self._build_results()
        self.build_people()
        self._build_quality_rules()
        self._build_settings()

        self._scan_default()
        self.after(300, self._auto_run_if_docs)

    # ---------- Input
    def _build_input(self):
        frm = ttk.Frame(self.tab_input, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Джерело DOCX (файл або папка). Якщо не вибрати — береться input/documents/").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.selected_path, width=90).grid(row=1, column=0, sticky="we", padx=(0,8))
        btns = ttk.Frame(frm)
        btns.grid(row=1, column=1, sticky="e")
        ttk.Button(btns, text="📄 Файл", command=self._pick_file).pack(side="left", padx=4)
        ttk.Button(btns, text="📁 Папка", command=self._pick_folder).pack(side="left", padx=4)
        ttk.Button(btns, text="🔍 Сканувати", command=self._scan).pack(side="left", padx=4)

        tree_frame = ttk.Frame(frm)
        tree_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=10)

        self.tree = ttk.Treeview(tree_frame, columns=("path",), show="headings", height=22)
        self.tree.heading("path", text="Знайдені DOCX")
        self.tree.column("path", width=900)
        ysb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        frm.rowconfigure(2, weight=1)
        frm.columnconfigure(0, weight=1)

        actions = ttk.Frame(frm)
        actions.grid(row=3, column=0, columnspan=2, sticky="we")
        ttk.Button(actions, text="⚡ Обробити все (input/documents)", command=self._run_default).pack(side="left")
        ttk.Button(actions, text="▶ Обробити вибране", command=self._run_selected).pack(side="left", padx=8)

    def _pick_file(self):
        p = filedialog.askopenfilename(filetypes=[("DOCX","*.docx")])
        if p:
            self.selected_path.set(p)
            self._scan()

    def _pick_folder(self):
        p = filedialog.askdirectory()
        if p:
            self.selected_path.set(p)
            self._scan()

    def _collect_docx(self, path):
        if not path:
            path = DOCS_DIR
        files = []
        if os.path.isfile(path) and path.lower().endswith(".docx"):
            return [path]
        if os.path.isdir(path):
            for root, dirs, fs in os.walk(path):
                # Ignore common temporary/system directories.
                dirs[:] = [d for d in dirs if d not in {"__pycache__", ".git", ".idea", "node_modules"}]
                for f in fs:
                    # Word temporary lock files start with ~$ and are not valid input documents.
                    if f.startswith("~$"):
                        continue
                    if f.lower().endswith(".docx"):
                        files.append(os.path.normpath(os.path.join(root, f)))
        # Avoid duplicate paths when a caller supplies overlapping sources.
        return sorted(set(files), key=lambda x: x.lower())

    def _scan_default(self):
        self.selected_path.set("")
        self._scan()

    def _scan(self):
        path = self.selected_path.get().strip()
        files = self._collect_docx(path)
        for i in self.tree.get_children():
            self.tree.delete(i)
        for f in files:
            self.tree.insert("", "end", values=(f,))
        self.status_var.set(f"Знайдено DOCX: {len(files)}")

    # ---------- Run
    def _build_run(self):
        frm = ttk.Frame(self.tab_run, padding=12)
        frm.pack(fill="both", expand=True)

        self.status_var = tk.StringVar(value="Готово")
        ttk.Label(frm, textvariable=self.status_var).pack(anchor="w")

        self.pb = ttk.Progressbar(frm, orient="horizontal", mode="determinate", maximum=100)
        self.pb.pack(fill="x", pady=6)

        log_frame = ttk.Frame(frm)
        log_frame.pack(fill="both", expand=True, pady=8)
        self.log = tk.Text(log_frame, height=28, wrap="none")
        self._apply_text_widget_style(self.log)
        ysb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        xsb = ttk.Scrollbar(log_frame, orient="horizontal", command=self.log.xview)
        self.log.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        bottom = ttk.Frame(frm)
        bottom.pack(fill="x")
        ttk.Button(bottom, text="🧹 Очистити лог", command=lambda: self.log.delete("1.0","end")).pack(side="left")
        ttk.Button(bottom, text="📂 Відкрити output", command=lambda: open_path(OUTPUT)).pack(side="left", padx=6)

    def _log(self, s):
        self.log.insert("end", s + "\n")
        self.log.see("end")
        self.update_idletasks()

    # ---------- UI settings (font + highlight colors)
    def _load_ui_settings(self) -> dict:
        """Load UI settings from config/ui_settings.json.

        Structure:
          {
            "font": {"family": "Consolas", "size": 12},
            "colors": {"BLOCK_ACTIVE": "#...", ...},
            "wrap_person_text": true
          }
        """
        # defaults
        try:
            base_font = tkfont.nametofont("TkTextFont")
            fam = base_font.actual("family")
            size = int(base_font.actual("size"))
        except Exception:
            fam, size = "Consolas", 12

        defaults = {
            "font": {"family": fam, "size": size},
            "colors": {
                "BLOCK_ACTIVE": "#fff3b0",
                "MATCH_ACTIVE": "#ffe066",
                "CAT_MATCH": "#ffd8a8",
                "PIB_MATCH": "#d3f9d8",
                "FIND": "#cfe8ff",
                "EVT_ARRIVAL": "#d3f9d8",
                "EVT_DEPARTURE": "#ffd6d6",
                "EVT_SPECIAL_MED": "#d0ebff",
                "EVT_SPECIAL_VACATION": "#fff3bf",
                "EVT_SPECIAL_TRIP": "#e5dbff",
            },
            "wrap_person_text": True,
        }

        path = self.ui_settings_path
        if not path or not os.path.exists(path):
            return defaults
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                return defaults
        except Exception:
            return defaults

        # merge with defaults
        out = defaults
        try:
            if isinstance(data.get("font"), dict):
                out["font"]["family"] = str(data["font"].get("family") or out["font"]["family"])
                out["font"]["size"] = int(data["font"].get("size") or out["font"]["size"])
            if isinstance(data.get("colors"), dict):
                for k, v in data["colors"].items():
                    if isinstance(v, str) and v.startswith("#") and len(v) in (4, 7):
                        out["colors"][k] = v
            if "wrap_person_text" in data:
                out["wrap_person_text"] = bool(data.get("wrap_person_text"))
        except Exception:
            return defaults
        return out

    def _save_ui_settings(self) -> None:
        os.makedirs(os.path.dirname(self.ui_settings_path), exist_ok=True)
        with open(self.ui_settings_path, "w", encoding="utf-8") as f:
            json.dump(self.ui_settings, f, ensure_ascii=False, indent=2)


    def _load_ignored_people(self) -> set:
        """Load hidden people list from config/ignored_people.json."""
        path = getattr(self, "ignored_people_path", None)
        if not path or not os.path.exists(path):
            return set()
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return set(str(x).strip() for x in data if str(x).strip())
        except Exception:
            return set()
        return set()

    def _save_ignored_people(self):
        """Persist hidden people list."""
        path = getattr(self, "ignored_people_path", None)
        if not path:
            return
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(sorted(self._ignored_people), f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def hide_selected_person(self):
        """Hide selected person from the People list without deleting parsed data."""
        pib = self._get_selected_pib()
        if not pib:
            messagebox.showwarning("Нема вибору", "Оберіть людину зі списку.")
            return
        self._ignored_people.add(pib)
        self._save_ignored_people()
        self.apply_people_filter(select_first=True)

    def restore_person_to_list(self):
        """Restore one or more hidden people to the visible People list."""
        hidden = sorted(self._ignored_people, key=lambda x: x.upper())
        if not hidden:
            messagebox.showinfo("Список порожній", "Немає прихованих людей для повернення.")
            return

        win = tk.Toplevel(self)
        win.title("Повернути до списку")
        win.geometry("620x430")
        win.transient(self)
        win.grab_set()

        outer = ttk.Frame(win, padding=12)
        outer.pack(fill="both", expand=True)
        ttk.Label(
            outer,
            text="Виберіть людину або декілька людей, яких потрібно повернути до списку:"
        ).pack(anchor="w", pady=(0, 8))

        box = ttk.Frame(outer)
        box.pack(fill="both", expand=True)
        tree = ttk.Treeview(box, columns=("pib",), show="headings", selectmode="extended")
        tree.heading("pib", text="ПІБ")
        tree.column("pib", width=540, anchor="w")
        ysb = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=ysb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)

        for pib in hidden:
            tree.insert("", "end", values=(pib,))

        btns = ttk.Frame(outer)
        btns.pack(fill="x", pady=(10, 0))

        def restore_selected():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("Нема вибору", "Оберіть хоча б одну людину.", parent=win)
                return
            restored = []
            for iid in selected:
                vals = tree.item(iid, "values")
                if vals:
                    pib = str(vals[0]).strip()
                    if pib in self._ignored_people:
                        self._ignored_people.remove(pib)
                        restored.append(pib)
            self._save_ignored_people()
            self.apply_people_filter(select_first=True)
            win.destroy()
            messagebox.showinfo("Повернення", f"Повернуто до списку: {len(restored)}", parent=self)

        ttk.Button(btns, text="↩ Повернути вибране", command=restore_selected).pack(side="left")
        ttk.Button(btns, text="Скасувати", command=win.destroy).pack(side="right")
    def _apply_global_style(self) -> None:
        """Apply font+colors to ttk widgets and text widgets when they exist."""
        font_family = (self.ui_settings.get("font") or {}).get("family")
        font_size = int((self.ui_settings.get("font") or {}).get("size") or 12)
        try:
            style = ttk.Style(self)
            style.configure("TLabel", font=(font_family, font_size))
            style.configure("TButton", font=(font_family, font_size))
            style.configure("TEntry", font=(font_family, font_size))
            style.configure("Treeview", font=(font_family, font_size))
            style.configure("Treeview.Heading", font=(font_family, font_size, "bold"))
        except Exception:
            pass

    def _apply_text_widget_style(self, widget: tk.Widget) -> None:
        font_family = (self.ui_settings.get("font") or {}).get("family")
        font_size = int((self.ui_settings.get("font") or {}).get("size") or 12)
        try:
            widget.configure(font=(font_family, font_size))
        except Exception:
            return

    def _auto_run_if_docs(self):
        files = self._collect_docx(DOCS_DIR)
        if files:
            # auto-run only when nothing selected
            if not self.selected_path.get().strip():
                self._log("[AUTO] Знайдено DOCX у input/documents — запускаю обробку.")
                self._run_pipeline(files)

    def _run_default(self):
        files = self._collect_docx(DOCS_DIR)
        if not files:
            messagebox.showwarning("Немає DOCX", "Покладіть DOCX у input/documents/")
            return
        self._run_pipeline(files)

    def _run_selected(self):
        path = self.selected_path.get().strip()
        files = self._collect_docx(path)
        if not files:
            messagebox.showwarning("Немає DOCX", "Не знайдено DOCX у вибраному шляху.")
            return
        self._run_pipeline(files)

    def _run_pipeline(self, docx_files):
        # simple progress stages: parse->json->excel->diagnostics
        self.pb["value"] = 0
        self.status_var.set(f"Парсинг DOCX: {len(docx_files)}")
        self._log(f"Парсинг DOCX ({len(docx_files)} файл(ів))...")
        self.update()

        combined = build_combined_v2(docx_files, template_path=self.template_path, learning_dir=self.learning_dir)
        processed = combined.get("processed_files") or []
        errors = combined.get("errors") or []
        self._log(f"Успішно оброблено: {len(processed)}; помилок: {len(errors)}")
        for err in errors:
            self._log(f"[DOCX ERROR] {err.get('file')}: {err.get('error_type')}: {err.get('error')}")
        if not processed and errors:
            messagebox.showerror("Помилка обробки", "Жоден DOCX не вдалося обробити. Деталі дивіться у журналі вкладки «Обробка».")
            self.status_var.set("Помилка")
            return
        self.pb["value"] = 35
        self.update()

        combined_path = os.path.join(self.parsed_dir, "combined_v2.json")
        with open(combined_path, "w", encoding="utf-8") as f:
            json.dump(combined, f, ensure_ascii=False, indent=2)
        self.last_json = combined_path
        self._log(f"JSON збережено: {combined_path}")
        self.pb["value"] = 55
        self.update()

        # flatten records
        records = []
        for person in combined.get("people", []):
            pib = person.get("pib","")
            for r in person.get("records", []):
                if not r.get("pib"):
                    r["pib"] = pib
                records.append(r)

        self.status_var.set("Автозаповнення Excel...")
        out = autofill_template(
            records,
            template_path=self.template_path,
            output_dir=self.output_dir,
            rules_path=self.rules_path,
            mapping_path=self.mapping_path,
            diagnostics_path=self.diag_path,
            learning_dir=self.learning_dir
        )
        self.last_excel = out
        self._log(f"Excel готовий: {out}")
        self.pb["value"] = 90
        self.update()

        self._refresh_results()
        self.refresh_people_view()
        self._refresh_quality_tree()
        self.pb["value"] = 100
        self.status_var.set("Готово ✅")

    # ---------- Results
    def _build_results(self):
        frm = ttk.Frame(self.tab_results, padding=12)
        frm.pack(fill="both", expand=True)

        self.res_excel = tk.StringVar(value="")
        self.res_json = tk.StringVar(value="")
        self.res_diag = tk.StringVar(value=self.diag_path)

        ttk.Label(frm, text="Останній Excel:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.res_excel, width=90).grid(row=1, column=0, sticky="we")
        ttk.Button(frm, text="Відкрити файл", command=self._open_last_excel).grid(row=1, column=1, padx=6)

        ttk.Label(frm, text="Останній JSON:").grid(row=2, column=0, sticky="w", pady=(10,0))
        ttk.Entry(frm, textvariable=self.res_json, width=90).grid(row=3, column=0, sticky="we")
        ttk.Button(frm, text="Відкрити файл", command=self._open_last_json).grid(row=3, column=1, padx=6)

        ttk.Label(frm, text="Діагностика (unknowns):").grid(row=4, column=0, sticky="w", pady=(10,0))
        ttk.Entry(frm, textvariable=self.res_diag, width=90).grid(row=5, column=0, sticky="we")
        ttk.Button(frm, text="Відкрити файл", command=self._open_diag).grid(row=5, column=1, padx=6)

        ttk.Separator(frm).grid(row=6, column=0, columnspan=2, sticky="we", pady=12)

        ttk.Button(frm, text="📂 Відкрити папку Excel", command=lambda: open_path(self.output_dir)).grid(row=7, column=0, sticky="w")
        ttk.Button(frm, text="📂 Відкрити output/parsed", command=lambda: open_path(self.parsed_dir)).grid(row=7, column=1, sticky="e")

        frm.columnconfigure(0, weight=1)

    def _refresh_results(self):
        self.res_excel.set(self.last_excel or "")
        self.res_json.set(self.last_json or "")
        self.res_diag.set(self.diag_path)

    def _open_last_excel(self):
        if self.last_excel and os.path.exists(self.last_excel):
            open_path(self.last_excel)
        else:
            messagebox.showinfo("Немає", "Ще не створено Excel.")

    def _open_last_json(self):
        if self.last_json and os.path.exists(self.last_json):
            open_path(self.last_json)
        else:
            messagebox.showinfo("Немає", "Ще не створено JSON.")

    def _open_diag(self):
        if os.path.exists(self.diag_path):
            open_path(self.diag_path)
        else:
            messagebox.showinfo("Немає", "Ще не створено діагностику.")

    # ---------- Parser quality / rules files
    def _build_quality_rules(self):
        """Manage JSON files that directly influence parser quality/accuracy."""
        frm = ttk.Frame(self.tab_quality, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(
            frm,
            text=(
                "JSON-файли правил та конфігурацій, які використовує парсер. "
                "Завантаження замінює відповідний файл у config/. Перед заміною "
                "автоматично створюється резервна копія."
            ),
            wraplength=1000,
            justify="left"
        ).grid(row=0, column=0, columnspan=4, sticky="we", pady=(0, 10))

        cols = ("file", "destination", "status")
        self.quality_tree = ttk.Treeview(frm, columns=cols, show="headings", height=18)
        self.quality_tree.heading("file", text="JSON-файл")
        self.quality_tree.heading("destination", text="Призначення")
        self.quality_tree.heading("status", text="Стан")
        self.quality_tree.column("file", width=260, anchor="w")
        self.quality_tree.column("destination", width=460, anchor="w")
        self.quality_tree.column("status", width=180, anchor="w")
        self.quality_tree.grid(row=1, column=0, columnspan=3, sticky="nsew")

        ysb = ttk.Scrollbar(frm, orient="vertical", command=self.quality_tree.yview)
        ysb.grid(row=1, column=3, sticky="ns")
        self.quality_tree.configure(yscrollcommand=ysb.set)

        actions = ttk.Frame(frm)
        actions.grid(row=2, column=0, columnspan=4, sticky="we", pady=(10, 6))
        ttk.Button(actions, text="📥 Завантажити JSON", command=self._import_quality_json).pack(side="left", padx=(0, 6))
        ttk.Button(actions, text="🔍 Перевірити всі", command=self._validate_quality_jsons).pack(side="left", padx=6)
        ttk.Button(actions, text="📤 Експортувати всі", command=self._export_quality_jsons).pack(side="left", padx=6)
        ttk.Button(actions, text="📂 Відкрити config", command=lambda: open_path(CONFIG)).pack(side="left", padx=6)
        ttk.Button(actions, text="🔄 Оновити", command=self._refresh_quality_tree).pack(side="right")

        info = ttk.LabelFrame(frm, text="Файли, що впливають на парсинг / класифікацію / автозаповнення", padding=8)
        info.grid(row=3, column=0, columnspan=4, sticky="we", pady=(6, 0))
        self.quality_info = tk.Text(info, height=7, wrap="word")
        self._apply_text_widget_style(self.quality_info)
        self.quality_info.pack(fill="both", expand=True)
        self.quality_info.insert("1.0", self._quality_files_help_text())
        self.quality_info.configure(state="disabled")

        frm.rowconfigure(1, weight=1)
        frm.columnconfigure(1, weight=1)
        self._refresh_quality_tree()

    def _quality_file_map(self):
        return {
            "event_categories_rules_v1.json": os.path.join("categories", "event_categories_rules_v1.json"),
            "user_category_rules.json": os.path.join("categories", "user_category_rules.json"),
            "units.json": os.path.join("learning", "units.json"),
            "reasons.json": os.path.join("learning", "reasons.json"),
            "arrival_sources.json": os.path.join("learning", "arrival_sources.json"),
            "events.json": os.path.join("learning", "events.json"),
            "mapping.json": os.path.join("autofill", "mapping.json"),
            "passport_rules.json": os.path.join("autofill", "passport_rules.json"),
            "template_sources.json": os.path.join("autofill", "template_sources.json"),
            "military_order.schema.json": os.path.join("schema", "military_order.schema.json")
        }

    def _quality_files_help_text(self):
        return (
            "Активні файли:\n"
            "• categories/event_categories_rules_v1.json — базові правила категоризації;\n"
            "• categories/user_category_rules.json — користувацькі правила категоризації;\n"
            "• learning/*.json — словники навчання/розпізнавання;\n"
            "• autofill/mapping.json — мапінг для автозаповнення;\n"
            "• autofill/passport_rules.json — правила автозаповнення;\n"
            "• autofill/template_sources.json — джерела даних шаблону;\n"
            "• schema/military_order.schema.json — контроль структури результату парсингу.\n\n"
            "Файли config/ui_settings.json та config/ignored_people.json не замінюються цією вкладкою, "
            "оскільки вони є налаштуваннями інтерфейсу/списку прихованих людей, а не правилами точності парсера."
        )

    def _refresh_quality_tree(self):
        if not hasattr(self, "quality_tree"):
            return
        for iid in self.quality_tree.get_children():
            self.quality_tree.delete(iid)
        fmap = self._quality_file_map()
        for name, rel in fmap.items():
            path = os.path.join(CONFIG, rel)
            status = "OK" if os.path.isfile(path) else "ВІДСУТНІЙ"
            if status == "OK":
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        json.load(f)
                except Exception as exc:
                    status = f"ПОМИЛКА JSON: {type(exc).__name__}"
            self.quality_tree.insert("", "end", values=(name, os.path.join("config", rel), status))

    def _validate_quality_file(self, path):
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if os.path.basename(path) == "military_order.schema.json":
            try:
                import jsonschema
                jsonschema.Draft7Validator.check_schema(data)
            except ImportError:
                return "JSON OK; jsonschema не встановлено"
        return "OK"

    def _validate_quality_jsons(self):
        errors = []
        ok = 0
        for name, rel in self._quality_file_map().items():
            path = os.path.join(CONFIG, rel)
            if not os.path.isfile(path):
                errors.append(f"{name}: файл відсутній")
                continue
            try:
                self._validate_quality_file(path)
                ok += 1
            except Exception as exc:
                errors.append(f"{name}: {exc}")
        self._refresh_quality_tree()
        if errors:
            messagebox.showerror("Перевірка JSON", f"Коректних: {ok}\nПомилок: {len(errors)}\n\n" + "\n".join(errors[:20]))
        else:
            messagebox.showinfo("Перевірка JSON", f"Усі {ok} активних JSON-файлів коректні.")

    def _import_quality_json(self):
        paths = filedialog.askopenfilenames(
            title="Виберіть JSON-файли правил парсера",
            filetypes=[("JSON", "*.json"), ("Усі файли", "*.*")]
        )
        if not paths:
            return

        fmap = self._quality_file_map()
        unknown = []
        planned = []
        for src in paths:
            name = os.path.basename(src)
            rel = fmap.get(name)
            if rel:
                planned.append((src, os.path.join(CONFIG, rel)))
            else:
                unknown.append(src)

        if unknown:
            unknown_dir = os.path.join(CONFIG, "quality_imports")
            os.makedirs(unknown_dir, exist_ok=True)
            for src in unknown:
                planned.append((src, os.path.join(unknown_dir, os.path.basename(src))))

        if not planned:
            return

        overwrite = [dst for _, dst in planned if os.path.exists(dst)]
        if overwrite:
            answer = messagebox.askyesno(
                "Підтвердження заміни",
                "Деякі активні JSON-файли вже існують. Перед заміною буде створено backup.\n\n"
                "Продовжити?"
            )
            if not answer:
                return

        import shutil
        backup_root = os.path.join(CONFIG, "backups", time.strftime("%Y%m%d_%H%M%S"))
        imported = []
        errors = []
        for src, dst in planned:
            try:
                self._validate_quality_file(src)
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                if os.path.exists(dst):
                    rel = os.path.relpath(dst, CONFIG)
                    bdst = os.path.join(backup_root, rel)
                    os.makedirs(os.path.dirname(bdst), exist_ok=True)
                    shutil.copy2(dst, bdst)
                shutil.copy2(src, dst)
                imported.append(os.path.relpath(dst, CONFIG))
            except Exception as exc:
                errors.append(f"{os.path.basename(src)}: {exc}")

        self._refresh_quality_tree()
        try:
            self._cat_engine = None
            self._ensure_category_engine()
        except Exception:
            pass

        msg = f"Завантажено: {len(imported)}"
        if backup_root and overwrite:
            msg += f"\nBackup: {backup_root}"
        if errors:
            msg += "\n\nПомилки:\n" + "\n".join(errors)
            messagebox.showerror("Імпорт JSON", msg)
        else:
            messagebox.showinfo("Імпорт JSON", msg)

    def _export_quality_jsons(self):
        dst_dir = filedialog.askdirectory(title="Папка для експорту JSON-файлів правил")
        if not dst_dir:
            return
        import shutil
        count = 0
        for name, rel in self._quality_file_map().items():
            src = os.path.join(CONFIG, rel)
            if os.path.isfile(src):
                out = os.path.join(dst_dir, name)
                shutil.copy2(src, out)
                count += 1
        messagebox.showinfo("Експорт JSON", f"Експортовано {count} файлів.")

    def build_people(self):
        frm = ttk.Frame(self.tab_people, padding=12)
        frm.pack(fill="both", expand=True)

        # --- toolbar
        top = ttk.Frame(frm)
        top.pack(fill="x", pady=(0, 8))

        ttk.Button(top, text="⮜/⮞", width=5, command=self._toggle_people_sidebar).pack(side="left")
        ttk.Button(top, text="🔄 Оновити список (A–Я)", command=self.refresh_people_view).pack(side="left")
        ttk.Button(top, text="🗑 Видалити зі списку", command=self.hide_selected_person).pack(side="left", padx=6)
        ttk.Button(top, text="↩ Повернути до списку", command=self.restore_person_to_list).pack(side="left", padx=6)
        ttk.Button(top, text="💾 TXT (людина)", command=self.export_person_txt).pack(side="left", padx=6)
        ttk.Button(top, text="💾 JSON (людина)", command=self.export_person_json).pack(side="left")
        ttk.Button(top, text="📦 Експорт ВСІХ (ZIP)", command=self.export_all_people_zip).pack(side="left", padx=6)
        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Button(top, text="📄 Відкрити DOCX джерело", command=self.open_selected_docx_source).pack(side="left")
        ttk.Button(top, text="🗎 Оригінал (людина)", command=self.open_person_original_view).pack(side="left", padx=6)
        ttk.Button(top, text="🏷 Перегляд по категоріях", command=self.open_category_browser).pack(side="left")
        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Button(top, text="➕ Правило з виділення", command=self._add_rule_from_selection).pack(side="left")
        ttk.Button(top, text="🧾 Правила", command=self.open_user_category_rules_manager).pack(side="left", padx=6)

        ttk.Label(top, text="Пошук ПІБ:").pack(side="left", padx=(18, 4))
        self.people_filter = tk.StringVar(value="")
        ent = ttk.Entry(top, textvariable=self.people_filter, width=28)
        ent.pack(side="left")
        ent.bind("<KeyRelease>", lambda e: self.apply_people_filter())

        # --- layout: left sidebar + right content
        body = ttk.Panedwindow(frm, orient="horizontal")
        body.pack(fill="both", expand=True)

        left = ttk.Frame(body)
        right = ttk.Frame(body)
        body.add(left, weight=1)
        body.add(right, weight=3)

        # store for collapse/expand
        self._people_paned = body
        self._people_left = left
        self._people_right = right
        self._people_sidebar_collapsed = False

        # helper: create treeview with both scrollbars
        def _mk_tree(parent, height: int, show: str, columns: tuple[str, ...]):
            box = ttk.Frame(parent)
            tree = ttk.Treeview(box, columns=columns, show=show, height=height)
            ysb = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
            xsb = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
            tree.grid(row=0, column=0, sticky="nsew")
            ysb.grid(row=0, column=1, sticky="ns")
            xsb.grid(row=1, column=0, sticky="ew")
            box.rowconfigure(0, weight=1)
            box.columnconfigure(0, weight=1)
            return box, tree

        # LEFT: people list (tree) + per-person doc list
        ttk.Label(left, text="ПІБ (алфавіт):").pack(anchor="w")
        people_box, self.people_tree = _mk_tree(left, height=18, show="headings", columns=("pib", "blocks", "records"))
        people_box.pack(fill="both", expand=True)
        self.people_tree.heading("pib", text="ПІБ")
        self.people_tree.heading("blocks", text="Блоки")
        self.people_tree.heading("records", text="Записи")
        self.people_tree.column("pib", width=240)
        self.people_tree.column("blocks", width=60, anchor="center")
        self.people_tree.column("records", width=60, anchor="center")
        self.people_tree.bind("<<TreeviewSelect>>", lambda e: self.show_selected_person())
        self.people_tree.bind("<Delete>", lambda e: self.hide_selected_person())

        docs_box = ttk.Frame(left)
        docs_box.pack(fill="both", expand=False, pady=(10, 0))
        ttk.Label(docs_box, text="Документи цієї людини:").pack(anchor="w")
        docs_tree_box, self.docs_tree = _mk_tree(docs_box, height=9, show="tree headings", columns=("count",))
        docs_tree_box.pack(fill="both", expand=True)
        self.docs_tree.heading("#0", text="DOCX / Блок")
        self.docs_tree.heading("count", text="К-сть")
        self.docs_tree.column("#0", width=260)
        self.docs_tree.column("count", width=60, anchor="center")
        self.docs_tree.bind("<<TreeviewSelect>>", lambda e: self._on_docs_tree_select())
        self.docs_tree.bind("<Double-1>", lambda e: self._on_docs_tree_open_or_jump())

        cats_box = ttk.Frame(left)
        cats_box.pack(fill="both", expand=False, pady=(10, 0))
        ttk.Label(cats_box, text="Категорії (Надійшло/Вибуло):").pack(anchor="w")
        cats_tree_box, self.cats_tree = _mk_tree(cats_box, height=9, show="tree headings", columns=("count",))
        cats_tree_box.pack(fill="both", expand=True)
        self.cats_tree.heading("#0", text="Категорія")
        self.cats_tree.heading("count", text="К-сть")
        self.cats_tree.column("#0", width=260)
        self.cats_tree.column("count", width=60, anchor="center")
        self.cats_tree.bind("<<TreeviewSelect>>", lambda e: self._on_cats_tree_select())

        cat_btns = ttk.Frame(cats_box)
        cat_btns.pack(fill="x", pady=(6, 0))
        ttk.Button(cat_btns, text="✖ Скинути категорію", command=self._clear_category_filter).pack(side="left")


        # RIGHT: text + fragments
        right_top = ttk.Frame(right)
        right_top.pack(fill="x")
        ttk.Label(right_top, text="Повʼязаний текст (повні блоки наказів):").pack(side="left")
        ttk.Label(right_top, text="Пошук у тексті:").pack(side="left", padx=(18, 4))
        self.find_var = tk.StringVar(value="")
        find_ent = ttk.Entry(right_top, textvariable=self.find_var, width=28)
        find_ent.pack(side="left")
        ttk.Button(right_top, text="⬇", width=3, command=lambda: self.find_in_person_text(next_=True)).pack(side="left", padx=2)
        ttk.Button(right_top, text="⬆", width=3, command=lambda: self.find_in_person_text(next_=False)).pack(side="left", padx=2)
        ttk.Button(right_top, text="🧹", width=3, command=self.clear_find_highlight).pack(side="left", padx=2)

        ttk.Separator(right_top, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Button(right_top, text="A-", width=3, command=lambda: self._adjust_font_size(-1)).pack(side="left")
        ttk.Button(right_top, text="A+", width=3, command=lambda: self._adjust_font_size(1)).pack(side="left", padx=2)
        ttk.Button(right_top, text="↔", width=3, command=self._toggle_person_wrap).pack(side="left", padx=(6, 0))

        # Vertical split
        vpan = ttk.Panedwindow(right, orient="vertical")
        vpan.pack(fill="both", expand=True, pady=(6, 0))

        text_frame = ttk.Frame(vpan)
        bottom_frame = ttk.Frame(vpan)
        vpan.add(text_frame, weight=3)
        vpan.add(bottom_frame, weight=2)

        wrap_mode = "word" if bool(self.ui_settings.get("wrap_person_text", True)) else "none"
        self.person_text = tk.Text(text_frame, wrap=wrap_mode)
        self._apply_text_widget_style(self.person_text)
        txt_ysb = ttk.Scrollbar(text_frame, orient="vertical", command=self.person_text.yview)
        txt_xsb = ttk.Scrollbar(text_frame, orient="horizontal", command=self.person_text.xview)
        self.person_text.configure(yscrollcommand=txt_ysb.set, xscrollcommand=txt_xsb.set)
        self.person_text.grid(row=0, column=0, sticky="nsew")
        txt_ysb.grid(row=0, column=1, sticky="ns")
        txt_xsb.grid(row=1, column=0, sticky="ew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        # Tags for interaction
        # Tags for interaction (colors from ui_settings)
        try:
            base_font = tkfont.Font(self.person_text, self.person_text.cget("font"))
            self._font_bold = tkfont.Font(self.person_text, base_font)
            self._font_bold.configure(weight="bold")
        except Exception:
            self._font_bold = None
        if self._font_bold:
            self.person_text.tag_configure("HDR", font=self._font_bold)
        colors = (self.ui_settings.get("colors") or {})
        self.person_text.tag_configure("BLOCK_ACTIVE", background=colors.get("BLOCK_ACTIVE", "#fff3b0"))
        self.person_text.tag_configure("MATCH_ACTIVE", background=colors.get("MATCH_ACTIVE", "#ffe066"))
        self.person_text.tag_configure("CAT_MATCH", background=colors.get("CAT_MATCH", "#ffd8a8"))
        self.person_text.tag_configure("PIB_MATCH", background=colors.get("PIB_MATCH", "#d3f9d8"))
        self.person_text.tag_configure("FIND", background=colors.get("FIND", "#cfe8ff"))
        self.person_text.tag_configure("ARRIVAL", background=colors.get("EVT_ARRIVAL", "#d3f9d8"))
        self.person_text.tag_configure("DEPARTURE", background=colors.get("EVT_DEPARTURE", "#ffd6d6"))
        self.person_text.tag_configure("SPECIAL_MED", background=colors.get("EVT_SPECIAL_MED", "#d0ebff"))
        self.person_text.tag_configure("SPECIAL_VACATION", background=colors.get("EVT_SPECIAL_VACATION", "#fff3bf"))
        self.person_text.tag_configure("SPECIAL_TRIP", background=colors.get("EVT_SPECIAL_TRIP", "#e5dbff"))

        self.person_text.bind("<Button-3>", self._person_text_context_menu)

        # Bottom: fragments table
        ttk.Label(bottom_frame, text="Фрагменти / блоки (double-click → перейти в текст):").pack(anchor="w")
        frag_box = ttk.Frame(bottom_frame)
        frag_box.pack(fill="both", expand=True)
        self.frag_tree = ttk.Treeview(
            frag_box,
            columns=("idx", "date", "event", "category", "doc", "block", "snippet"),
            show="headings",
            height=10,
        )
        for col, title, w in [
            ("idx", "#", 40),
            ("date", "Дата", 90),
            ("event", "Подія", 80),
            ("category", "Категорія", 200),
            ("doc", "DOCX", 180),
            ("block", "Блок", 60),
            ("snippet", "Snippet", 520),
        ]:
            self.frag_tree.heading(col, text=title)
            self.frag_tree.column(col, width=w, anchor=("center" if col in ("idx", "block") else "w"))
        frag_ysb = ttk.Scrollbar(frag_box, orient="vertical", command=self.frag_tree.yview)
        frag_xsb = ttk.Scrollbar(frag_box, orient="horizontal", command=self.frag_tree.xview)
        self.frag_tree.configure(yscrollcommand=frag_ysb.set, xscrollcommand=frag_xsb.set)
        self.frag_tree.grid(row=0, column=0, sticky="nsew")
        frag_ysb.grid(row=0, column=1, sticky="ns")
        frag_xsb.grid(row=1, column=0, sticky="ew")
        frag_box.rowconfigure(0, weight=1)
        frag_box.columnconfigure(0, weight=1)
        self.frag_tree.bind("<Double-1>", lambda e: self.jump_to_selected_fragment())
        self.frag_tree.bind("<<TreeviewSelect>>", lambda e: self.jump_to_selected_fragment())

        # Row coloring by event_type (arrival vs departure)
        try:
            c = (self.ui_settings.get("colors") or {})
            self.frag_tree.tag_configure("ARRIVAL", background=c.get("EVT_ARRIVAL", "#d3f9d8"))
            self.frag_tree.tag_configure("DEPARTURE", background=c.get("EVT_DEPARTURE", "#ffd6d6"))
            self.frag_tree.tag_configure("SPECIAL_MED", background=c.get("EVT_SPECIAL_MED", "#d0ebff"))
            self.frag_tree.tag_configure("SPECIAL_VACATION", background=c.get("EVT_SPECIAL_VACATION", "#fff3bf"))
            self.frag_tree.tag_configure("SPECIAL_TRIP", background=c.get("EVT_SPECIAL_TRIP", "#e5dbff"))
            # Optional: docs tree child rows
            if hasattr(self, "docs_tree"):
                self.docs_tree.tag_configure("ARRIVAL", background=c.get("EVT_ARRIVAL", "#d3f9d8"))
                self.docs_tree.tag_configure("DEPARTURE", background=c.get("EVT_DEPARTURE", "#ffd6d6"))
                self.docs_tree.tag_configure("SPECIAL_MED", background=c.get("EVT_SPECIAL_MED", "#d0ebff"))
                self.docs_tree.tag_configure("SPECIAL_VACATION", background=c.get("EVT_SPECIAL_VACATION", "#fff3bf"))
                self.docs_tree.tag_configure("SPECIAL_TRIP", background=c.get("EVT_SPECIAL_TRIP", "#e5dbff"))
            # People tree rows (ПІБ list)
            if hasattr(self, "people_tree"):
                self.people_tree.tag_configure("ARRIVAL", background=c.get("EVT_ARRIVAL", "#d3f9d8"))
                self.people_tree.tag_configure("DEPARTURE", background=c.get("EVT_DEPARTURE", "#ffd6d6"))
                self.people_tree.tag_configure("SPECIAL_MED", background=c.get("EVT_SPECIAL_MED", "#d0ebff"))
                self.people_tree.tag_configure("SPECIAL_VACATION", background=c.get("EVT_SPECIAL_VACATION", "#fff3bf"))
                self.people_tree.tag_configure("SPECIAL_TRIP", background=c.get("EVT_SPECIAL_TRIP", "#e5dbff"))
                self.people_tree.tag_configure("BOTH", background=c.get("EVT_BOTH", "#fff3bf"))
        except Exception:
            pass

        # Internal caches
        self._people_cache = {}  # pib -> person dict
        self._people_all_pibs = []
        self._people_filtered_pibs = []
        self._frag_meta = {}  # tree item -> meta
        self._docs_meta = {}  # docs_tree item -> meta (doc or block)
        self._doc_tree_meta = {}  # docs_tree item -> meta

    def _toggle_people_sidebar(self):
        """Collapse/expand the left sidebar in People tab."""
        try:
            pan = getattr(self, "_people_paned", None)
            left = getattr(self, "_people_left", None)
            if not pan or not left:
                return
            if not getattr(self, "_people_sidebar_collapsed", False):
                try:
                    pan.forget(left)
                except Exception:
                    return
                self._people_sidebar_collapsed = True
            else:
                try:
                    pan.insert(0, left, weight=1)
                except Exception:
                    return
                self._people_sidebar_collapsed = False
        except Exception:
            return

    def _toggle_person_wrap(self):
        if not hasattr(self, "person_text"):
            return
        try:
            cur = str(self.person_text.cget("wrap"))
            new_wrap = "none" if cur != "none" else "word"
            self.person_text.configure(wrap=new_wrap)
            self.ui_settings["wrap_person_text"] = (new_wrap != "none")
            self._save_ui_settings()
        except Exception:
            pass

    def _adjust_font_size(self, delta: int):
        try:
            cur = int((self.ui_settings.get("font") or {}).get("size") or 12)
            new_size = max(8, min(28, cur + int(delta)))
            self.ui_settings.setdefault("font", {})["size"] = new_size
            self._save_ui_settings()
            self._apply_global_style()
            # re-apply to existing text widgets
            for wname in ("person_text", "log", "unknown_units", "unknown_reasons", "unknown_arrivals"):
                w = getattr(self, wname, None)
                if w is not None:
                    self._apply_text_widget_style(w)
            self._reconfigure_people_text_tags()
        except Exception:
            pass

    def _reconfigure_people_text_tags(self):
        if not hasattr(self, "person_text"):
            return
        colors = (self.ui_settings.get("colors") or {})
        try:
            # rebuild bold font for headers
            base_font = tkfont.Font(self.person_text, self.person_text.cget("font"))
            self._font_bold = tkfont.Font(self.person_text, base_font)
            self._font_bold.configure(weight="bold")
            self.person_text.tag_configure("HDR", font=self._font_bold)
        except Exception:
            pass
        try:
            self.person_text.tag_configure("BLOCK_ACTIVE", background=colors.get("BLOCK_ACTIVE", "#fff3b0"))
            self.person_text.tag_configure("MATCH_ACTIVE", background=colors.get("MATCH_ACTIVE", "#ffe066"))
            self.person_text.tag_configure("CAT_MATCH", background=colors.get("CAT_MATCH", "#ffd8a8"))
            self.person_text.tag_configure("PIB_MATCH", background=colors.get("PIB_MATCH", "#d3f9d8"))
            self.person_text.tag_configure("FIND", background=colors.get("FIND", "#cfe8ff"))
        except Exception:
            pass

    def open_person_original_view(self):
        """Open a separate window with the full raw blocks for the selected person."""
        pib = self._current_person_pib or self._get_selected_pib()
        if not pib:
            messagebox.showwarning("Немає людини", "Оберіть людину в списку.")
            return
        person = self._people_cache.get(pib) or {}
        mentions = person.get("mentions", []) or []

        def mention_key(m):
            dt = (m.get("date") or "").strip()
            dkey = self._parse_dt(dt) if dt else (9999, 12, 31)
            doc = (m.get("doc") or "").strip()
            bid = int(m.get("block_id") or 0)
            return (doc, dkey, bid)

        all_mentions = sorted(mentions, key=mention_key)

        win = tk.Toplevel(self)
        win.title(f"Оригінал — {pib}")
        win.geometry("980x700")

        outer = ttk.Frame(win, padding=10)
        outer.pack(fill="both", expand=True)

        topbar = ttk.Frame(outer)
        topbar.pack(fill="x")
        ttk.Label(topbar, text=f"ПІБ: {pib}").pack(side="left")
        ttk.Button(topbar, text="📄 Відкрити DOCX (якщо вибрано фрагмент)", command=self.open_selected_docx_source).pack(side="right")

        text_box = ttk.Frame(outer)
        text_box.pack(fill="both", expand=True, pady=(8, 0))
        wrap_mode = "word" if bool(self.ui_settings.get("wrap_person_text", True)) else "none"
        txt = tk.Text(text_box, wrap=wrap_mode)
        self._apply_text_widget_style(txt)
        ysb = ttk.Scrollbar(text_box, orient="vertical", command=txt.yview)
        xsb = ttk.Scrollbar(text_box, orient="horizontal", command=txt.xview)
        txt.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        txt.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        text_box.rowconfigure(0, weight=1)
        text_box.columnconfigure(0, weight=1)

        # fill content
        txt.insert("end", f"ПІБ: {pib}\n")
        txt.insert("end", "-" * 90 + "\n\n")
        for i, m in enumerate(all_mentions, start=1):
            doc = (m.get("doc") or "").strip() or "(невідомо)"
            bid = int(m.get("block_id") or 0)
            dt = (m.get("date") or "").strip() or "—"
            et = (m.get("event_type") or "").strip() or "—"
            text = (m.get("text") or "").strip()
            cat = m.get("_cat")
            if cat is None:
                # Prefer parser-provided category (if the parser already ran CategoryEngine).
                parsed_cat = m.get("category")
                if parsed_cat is not None:
                    cat = parsed_cat
                else:
                    c0, _, _ = self._classify_block_category(text, (m.get("event_type") or ""))
                    cat = c0
            head = f"[{i}] {dt} | {et} | {cat or '—'} | {doc} | блок {bid}\n"
            txt.insert("end", head)
            txt.insert("end", text + "\n\n")
        txt.configure(state="disabled")

    def _build_global_category_index(self) -> dict:
        """Build a global index: event_type -> category_label -> list[entry]."""
        idx: dict[str, dict[str, list[dict]]] = {"arrival": {}, "departure": {}, "unknown": {}}
        for pib, person in (self._people_cache or {}).items():
            for m in (person.get("mentions") or []):
                text = (m.get("text") or "").strip()
                et = (m.get("event_type") or "").strip()
                if et not in ("arrival", "departure"):
                    et = "unknown"
                # ensure category cached
                if m.get("_cat") is None and "_cat" not in m:
                    has_parsed = any(k in m for k in ("category", "rule_id", "cat_span"))
                    if has_parsed and (m.get("category") is not None or m.get("rule_id") is not None or m.get("cat_span") is not None):
                        m["_cat"], m["_cat_rule"], m["_cat_span"] = m.get("category"), m.get("rule_id"), m.get("cat_span")
                    else:
                        cat0, rule0, span0 = self._classify_block_category(text, m.get("event_type") or "")
                        m["_cat"], m["_cat_rule"], m["_cat_span"] = cat0, rule0, span0
                cat = m.get("_cat") or "(без категорії)"
                entry = {
                    "pib": pib,
                    "date": (m.get("date") or "").strip(),
                    "doc": (m.get("doc") or "").strip(),
                    "block_id": int(m.get("block_id") or 0),
                    "event_type": (m.get("event_type") or "").strip(),
                    "category": cat,
                    "snippet": " ".join(text.split())[:240],
                }
                idx.setdefault(et, {}).setdefault(cat, []).append(entry)
        return idx

    def open_category_browser(self):
        """Browse blocks across people by the same category (arrival/departure)."""
        if not getattr(self, "_people_cache", None):
            self.refresh_people_view()
        idx = self._build_global_category_index()

        win = tk.Toplevel(self)
        win.title("Перегляд по категоріях")
        win.geometry("1100x700")

        outer = ttk.Frame(win, padding=10)
        outer.pack(fill="both", expand=True)

        pan = ttk.Panedwindow(outer, orient="horizontal")
        pan.pack(fill="both", expand=True)

        left = ttk.Frame(pan)
        right = ttk.Frame(pan)
        pan.add(left, weight=1)
        pan.add(right, weight=3)

        ttk.Label(left, text="Категорії:").pack(anchor="w")
        cat_tree_box = ttk.Frame(left)
        cat_tree_box.pack(fill="both", expand=True)
        cat_tree = ttk.Treeview(cat_tree_box, columns=("count",), show="tree headings", height=18)
        cat_tree.heading("#0", text="Категорія")
        cat_tree.heading("count", text="К-сть")
        cat_tree.column("#0", width=260)
        cat_tree.column("count", width=70, anchor="center")
        ysb = ttk.Scrollbar(cat_tree_box, orient="vertical", command=cat_tree.yview)
        xsb = ttk.Scrollbar(cat_tree_box, orient="horizontal", command=cat_tree.xview)
        cat_tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        cat_tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        cat_tree_box.rowconfigure(0, weight=1)
        cat_tree_box.columnconfigure(0, weight=1)

        # Right: fragments list
        topbar = ttk.Frame(right)
        topbar.pack(fill="x")
        ttk.Label(topbar, text="Фільтр:").pack(side="left")
        qv = tk.StringVar(value="")
        ent = ttk.Entry(topbar, textvariable=qv, width=40)
        ent.pack(side="left", padx=6)
        ttk.Label(topbar, text="(ПІБ/DOCX/текст)").pack(side="left")

        frag_box = ttk.Frame(right)
        frag_box.pack(fill="both", expand=True, pady=(8, 0))
        cols = ("pib", "date", "event", "category", "doc", "block", "snippet")
        frag_tree = ttk.Treeview(frag_box, columns=cols, show="headings", height=22)
        for col, title, w in [
            ("pib", "ПІБ", 220),
            ("date", "Дата", 90),
            ("event", "Подія", 80),
            ("category", "Категорія", 200),
            ("doc", "DOCX", 170),
            ("block", "Блок", 60),
            ("snippet", "Snippet", 520),
        ]:
            frag_tree.heading(col, text=title)
            frag_tree.column(col, width=w, anchor=("center" if col in ("block",) else "w"))
        ysb2 = ttk.Scrollbar(frag_box, orient="vertical", command=frag_tree.yview)
        xsb2 = ttk.Scrollbar(frag_box, orient="horizontal", command=frag_tree.xview)
        frag_tree.configure(yscrollcommand=ysb2.set, xscrollcommand=xsb2.set)
        frag_tree.grid(row=0, column=0, sticky="nsew")
        ysb2.grid(row=0, column=1, sticky="ns")
        xsb2.grid(row=1, column=0, sticky="ew")
        frag_box.rowconfigure(0, weight=1)
        frag_box.columnconfigure(0, weight=1)

        # meta map
        meta: dict[str, dict] = {}

        def _fill_cat_tree():
            cat_tree.delete(*cat_tree.get_children(""))
            for ev in ("arrival", "departure", "unknown"):
                total = sum(len(v) for v in idx.get(ev, {}).values())
                ev_iid = cat_tree.insert("", "end", text=ev, values=(total,))
                for cat in sorted(idx.get(ev, {}).keys(), key=lambda x: x.lower()):
                    cat_tree.insert(ev_iid, "end", text=cat, values=(len(idx[ev][cat]),))
                cat_tree.item(ev_iid, open=True)

        def _current_entries():
            sel = cat_tree.selection()
            if not sel:
                return []
            iid = sel[0]
            parent = cat_tree.parent(iid)
            if not parent:
                # event selected
                ev = cat_tree.item(iid, "text")
                out = []
                for cat in idx.get(ev, {}):
                    out.extend(idx[ev][cat])
                return out
            ev = cat_tree.item(parent, "text")
            cat = cat_tree.item(iid, "text")
            return list(idx.get(ev, {}).get(cat, []))

        def _apply_filter(entries: list[dict]):
            q = (qv.get() or "").strip().lower()
            if not q:
                return entries
            out = []
            for e in entries:
                blob = f"{e.get('pib','')} {e.get('doc','')} {e.get('snippet','')} {e.get('category','')}".lower()
                if q in blob:
                    out.append(e)
            return out

        def _refresh_frag_list(*_):
            frag_tree.delete(*frag_tree.get_children(""))
            meta.clear()
            entries = _apply_filter(_current_entries())
            for e in entries:
                iid = frag_tree.insert("", "end", values=(
                    e.get("pib"),
                    e.get("date") or "—",
                    e.get("event_type") or "—",
                    e.get("category") or "—",
                    e.get("doc") or "—",
                    e.get("block_id"),
                    e.get("snippet") or "",
                ))
                meta[iid] = e

        def _jump_to_entry(event=None):
            sel = frag_tree.selection()
            if not sel:
                return
            e = meta.get(sel[0]) or {}
            pib = e.get("pib")
            if not pib or pib not in self._people_cache:
                return
            # switch to people tab and show
            try:
                if hasattr(self, "nb"):
                    self.nb.select(self.tab_people)
            except Exception:
                pass
            self._current_person_pib = pib
            self._current_doc_filter = None
            self._current_event_filter = (e.get("event_type") or "").strip() if (e.get("event_type") or "").strip() in ("arrival", "departure") else None
            self._current_category_filter = e.get("category") if e.get("category") not in ("(без категорії)", None) else None
            self._render_person_into_ui(pib, self._people_cache[pib], doc_filter=self._current_doc_filter)
            # select matching fragment row
            target_doc = e.get("doc")
            target_bid = int(e.get("block_id") or 0)
            for iid2, m in (self._frag_meta or {}).items():
                if (m.get("doc") == target_doc) and int(m.get("block_id") or 0) == target_bid:
                    try:
                        self.frag_tree.selection_set(iid2)
                        self.frag_tree.focus(iid2)
                        self.jump_to_selected_fragment()
                    except Exception:
                        pass
                    break

        _fill_cat_tree()
        cat_tree.bind("<<TreeviewSelect>>", _refresh_frag_list)
        ent.bind("<KeyRelease>", _refresh_frag_list)
        frag_tree.bind("<Double-1>", _jump_to_entry)
        # select first event
        roots = cat_tree.get_children("")
        if roots:
            cat_tree.selection_set(roots[0])
            _refresh_frag_list()

    def open_user_category_rules_manager(self):
        """Allow enabling/disabling/deleting user rules created from selection."""
        os.makedirs(os.path.dirname(self.user_category_rules_path), exist_ok=True)
        if not os.path.exists(self.user_category_rules_path):
            with open(self.user_category_rules_path, "w", encoding="utf-8") as f:
                json.dump({"rules": []}, f, ensure_ascii=False, indent=2)

        def load_rules():
            try:
                with open(self.user_category_rules_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                rules = data.get("rules") or []
                return [r for r in rules if isinstance(r, dict)]
            except Exception:
                return []

        def save_rules(rules: list[dict]):
            with open(self.user_category_rules_path, "w", encoding="utf-8") as f:
                json.dump({"rules": rules}, f, ensure_ascii=False, indent=2)

        win = tk.Toplevel(self)
        win.title("User правила категорій")
        win.geometry("1050x550")

        outer = ttk.Frame(win, padding=10)
        outer.pack(fill="both", expand=True)

        topbar = ttk.Frame(outer)
        topbar.pack(fill="x")
        ttk.Label(topbar, text="Фільтр:").pack(side="left")
        qv = tk.StringVar(value="")
        qent = ttk.Entry(topbar, textvariable=qv, width=40)
        qent.pack(side="left", padx=6)

        box = ttk.Frame(outer)
        box.pack(fill="both", expand=True, pady=(8, 0))

        cols = ("enabled", "event", "label", "regex", "priority", "id")
        tree = ttk.Treeview(box, columns=cols, show="headings", height=18)
        for c, t, w in [
            ("enabled", "ON", 55),
            ("event", "Подія", 80),
            ("label", "Категорія", 200),
            ("regex", "Regex", 420),
            ("priority", "Priority", 70),
            ("id", "ID", 180),
        ]:
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor=("center" if c in ("enabled", "priority") else "w"))
        ysb = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        xsb = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)

        btns = ttk.Frame(outer)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="✅/🚫 Увімкнути/вимкнути", command=lambda: toggle()).pack(side="left")
        ttk.Button(btns, text="🗑 Видалити", command=lambda: delete()).pack(side="left", padx=6)
        ttk.Button(btns, text="📂 Відкрити JSON", command=lambda: open_path(self.user_category_rules_path)).pack(side="right")

        rules = load_rules()
        meta: dict[str, dict] = {}

        def norm_rx(r):
            m = r.get("match") or {}
            any_ = m.get("any") or []
            return str(any_[0]) if any_ else ""

        def refresh():
            tree.delete(*tree.get_children(""))
            meta.clear()
            q = (qv.get() or "").strip().lower()
            for r in rules:
                rid = r.get("id") or ""
                ev = r.get("event_type") or ""
                lbl = r.get("label") or ""
                rx = norm_rx(r)
                pr = int(r.get("priority") or 0)
                en = bool(r.get("enabled", True))
                blob = f"{ev} {lbl} {rx} {rid}".lower()
                if q and q not in blob:
                    continue
                iid = tree.insert("", "end", values=("ON" if en else "OFF", ev, lbl, rx, pr, rid))
                meta[iid] = r

        def toggle():
            sel = tree.selection()
            if not sel:
                return
            r = meta.get(sel[0])
            if not r:
                return
            r["enabled"] = not bool(r.get("enabled", True))
            save_rules(rules)
            self._ensure_category_engine()
            refresh()
            # keep UI consistent
            if self._current_person_pib and self._current_person_pib in self._people_cache:
                self._render_person_into_ui(self._current_person_pib, self._people_cache[self._current_person_pib], doc_filter=self._current_doc_filter)

        def delete():
            sel = tree.selection()
            if not sel:
                return
            r = meta.get(sel[0])
            if not r:
                return
            rid = r.get("id")
            if not rid:
                return
            # remove
            rules[:] = [x for x in rules if x.get("id") != rid]
            save_rules(rules)
            self._ensure_category_engine()
            refresh()
            if self._current_person_pib and self._current_person_pib in self._people_cache:
                self._render_person_into_ui(self._current_person_pib, self._people_cache[self._current_person_pib], doc_filter=self._current_doc_filter)

        qent.bind("<KeyRelease>", lambda e: refresh())
        refresh()



    def _ensure_category_engine(self):
        """Load category rules engine if possible. Never crash UI if rules are missing."""
        try:
            if os.path.exists(self.category_rules_path):
                self._cat_engine = CategoryEngine(self.category_rules_path, self.user_category_rules_path)
            else:
                self._cat_engine = None
        except Exception:
            self._cat_engine = None

    def _add_rule_from_selection(self):
        """Create a new category rule from the currently selected phrase in person_text."""
        if not hasattr(self, "person_text"):
            return
        try:
            phrase = self.person_text.get(tk.SEL_FIRST, tk.SEL_LAST)
        except Exception:
            messagebox.showwarning("Нема виділення", "Виділіть у тексті слово/фразу, а потім натисніть 'Правило з виділення'.")
            return
        phrase = (phrase or "").strip()
        if not phrase:
            messagebox.showwarning("Нема виділення", "Виділіть у тексті слово/фразу (не порожньо).")
            return

        # Default event from selected fragment if any
        default_event = "arrival"
        sel = self.frag_tree.selection() if hasattr(self, "frag_tree") else ()
        if sel:
            meta = self._frag_meta.get(sel[0]) or {}
            et = (meta.get("event_type") or "").strip()
            if et in ("arrival", "departure"):
                default_event = et

        self._open_rule_dialog(phrase, default_event)

    def _open_rule_dialog(self, phrase: str, default_event: str):
        """Modal dialog to create rule from phrase."""
        if self._cat_engine is None:
            self._ensure_category_engine()

        top = tk.Toplevel(self)
        top.title("Створити правило з виділення")
        top.geometry("560x260")
        top.transient(self)
        top.grab_set()

        frm = ttk.Frame(top, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Виділена фраза:").grid(row=0, column=0, sticky="w")
        phrase_var = tk.StringVar(value=phrase)
        ent_phrase = ttk.Entry(frm, textvariable=phrase_var, state="readonly")
        ent_phrase.grid(row=0, column=1, sticky="ew", padx=(8,0))

        ttk.Label(frm, text="Подія:").grid(row=1, column=0, sticky="w", pady=(8,0))
        event_var = tk.StringVar(value=default_event if default_event in ("arrival","departure") else "arrival")
        ev_box = ttk.Frame(frm)
        ev_box.grid(row=1, column=1, sticky="w", pady=(8,0))
        ttk.Radiobutton(ev_box, text="arrival (Надійшло)", variable=event_var, value="arrival").pack(side="left")
        ttk.Radiobutton(ev_box, text="departure (Вибуло)", variable=event_var, value="departure").pack(side="left", padx=(10,0))

        ttk.Label(frm, text="Категорія (label):").grid(row=2, column=0, sticky="w", pady=(8,0))
        cat_var = tk.StringVar(value="")
        # combobox with existing labels + allow typing
        values = []
        if self._cat_engine:
            values = self._cat_engine.arrival_labels if event_var.get()=="arrival" else self._cat_engine.departure_labels
        cb = ttk.Combobox(frm, textvariable=cat_var, values=values, state="normal")
        cb.grid(row=2, column=1, sticky="ew", padx=(8,0), pady=(8,0))

        ttk.Label(frm, text="Regex (match.any):").grid(row=3, column=0, sticky="w", pady=(8,0))
        regex_var = tk.StringVar(value=make_regex_from_phrase(phrase))
        ent_rx = ttk.Entry(frm, textvariable=regex_var)
        ent_rx.grid(row=3, column=1, sticky="ew", padx=(8,0), pady=(8,0))

        prio_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text="Додати вгорі (вищий пріоритет)", variable=prio_var).grid(row=4, column=1, sticky="w", padx=(8,0), pady=(8,0))

        # update combobox options when event changes
        def _refresh_cb(*_):
            if not self._cat_engine:
                cb.configure(values=[])
                return
            cb.configure(values=self._cat_engine.arrival_labels if event_var.get()=="arrival" else self._cat_engine.departure_labels)
        event_var.trace_add("write", _refresh_cb)

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(14,0))
        ttk.Button(btns, text="Скасувати", command=top.destroy).pack(side="right")
        def _save():
            label = (cat_var.get() or "").strip()
            rx = (regex_var.get() or "").strip()
            ev = event_var.get()
            if not label:
                messagebox.showwarning("Нема категорії", "Оберіть або введіть назву категорії (label).")
                return
            if not rx:
                messagebox.showwarning("Нема regex", "Regex порожній. Вкажіть regex для match.any.")
                return
            try:
                re.compile(rx, re.IGNORECASE)
            except re.error as e:
                messagebox.showerror("Regex помилка", f"Regex невалідний:\n{e}")
                return
            try:
                self._persist_category_rule(event_type=ev, label=label, regex=rx, priority_high=bool(prio_var.get()))
            except Exception as e:
                messagebox.showerror("Помилка", f"Не можу зберегти правило:\n{e}")
                return
            top.destroy()
        ttk.Button(btns, text="Зберегти правило", command=_save).pack(side="right", padx=(0,8))

        frm.columnconfigure(1, weight=1)

    def _persist_category_rule(self, event_type: str, label: str, regex: str, priority_high: bool = True):
        """Write a new rule into user_category_rules.json (and reload)."""
        # ensure user rules file exists
        if not os.path.exists(self.user_category_rules_path):
            with open(self.user_category_rules_path, "w", encoding="utf-8") as f:
                json.dump({"version": 1, "rules": []}, f, ensure_ascii=False, indent=2)

        with open(self.user_category_rules_path, "r", encoding="utf-8") as f:
            udata = json.load(f) if f.readable() else {"version": 1, "rules": []}
        if not isinstance(udata, dict):
            udata = {"version": 1, "rules": []}
        udata.setdefault("version", 1)
        udata.setdefault("rules", [])
        if not isinstance(udata["rules"], list):
            udata["rules"] = []

        # stable-ish id
        base = f"{event_type}:{label}:{regex}".encode("utf-8")
        hid = hashlib.sha256(base).hexdigest()[:8]
        slug = re.sub(r"[^a-z0-9_]+", "_", label.lower())
        rid = f"{'arr' if event_type=='arrival' else 'dep'}_{slug}_{hid}"

        rule = {
            "id": rid,
            "event_type": event_type,
            "label": label,
            "match": {"any": [regex]},
            "priority": "high" if priority_high else "low",
            "enabled": True,
        }

        # de-dup by id (or same label+regex+event)
        def _same(r: dict) -> bool:
            try:
                return (r.get("id") == rid) or (
                    (r.get("event_type") == event_type)
                    and (r.get("label") == label)
                    and (r.get("match") or {}).get("any") == [regex]
                )
            except Exception:
                return False
        udata["rules"] = [r for r in udata["rules"] if not _same(r)]

        # prepend high priority rules
        if priority_high:
            udata["rules"].insert(0, rule)
        else:
            udata["rules"].append(rule)

        with open(self.user_category_rules_path, "w", encoding="utf-8") as f:
            json.dump(udata, f, ensure_ascii=False, indent=2)

        self._ensure_category_engine()

        # refresh current person view
        pib = self._current_person_pib
        if pib and pib in self._people_cache:
            self._render_person_into_ui(pib, self._people_cache[pib], doc_filter=self._current_doc_filter)

    def _person_text_context_menu(self, event):
        menu = tk.Menu(self, tearoff=False)
        menu.add_command(label="Копіювати", command=lambda: self.person_text.event_generate("<<Copy>>"))
        menu.add_separator()
        menu.add_command(label="➕ Правило з виділення", command=self._add_rule_from_selection)
        menu.tk_popup(event.x_root, event.y_root)


    def refresh_people_view(self):
        # Load combined_v2.json and build alphabetical index by PIB
        self._people_cache = {}
        self._people_all_pibs = []
        self._people_filtered_pibs = []
        # clear UI lists
        if hasattr(self, "people_tree"):
            for iid in self.people_tree.get_children():
                self.people_tree.delete(iid)
        if hasattr(self, "docs_tree"):
            for iid in self.docs_tree.get_children():
                self.docs_tree.delete(iid)
        if hasattr(self, "frag_tree"):
            for iid in self.frag_tree.get_children():
                self.frag_tree.delete(iid)
        self._frag_meta = {}
        self._docs_meta = {}
        self._mention_pos = {}
        if not os.path.exists(self.combined_path):
            self.person_text.delete("1.0", "end")
            self.person_text.insert("end", "Немає combined_v2.json. Спочатку обробіть DOCX.\n")

            return
        try:
            with open(self.combined_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            self.person_text.delete("1.0", "end")
            self.person_text.insert("end", f"Не можу прочитати combined_v2.json: {e}\n")
            return

        people = data.get("people", []) if isinstance(data, dict) else []
        people_sorted = sorted(people, key=lambda p: (p.get("pib") or "").upper())

        for p in people_sorted:
            pib = (p.get("pib") or "").strip()
            if not pib:
                continue
            self._people_cache[pib] = p
            self._people_all_pibs.append(pib)

        self.apply_people_filter(select_first=True)

    def apply_people_filter(self, select_first=False):
        # Filter people by substring (case-insensitive) and keep alphabetical order
        q = (self.people_filter.get() if hasattr(self, "people_filter") else "").strip().upper()
        if not q:
            self._people_filtered_pibs = list(self._people_all_pibs)
        else:
            self._people_filtered_pibs = [p for p in self._people_all_pibs if q in p.upper()]
        # ignore hidden people
        try:
            hidden = getattr(self, "_ignored_people", set()) or set()
            if hidden:
                self._people_filtered_pibs = [p for p in self._people_filtered_pibs if p not in hidden]
        except Exception:
            pass

        # rebuild tree
        if hasattr(self, "people_tree"):
            for iid in self.people_tree.get_children():
                self.people_tree.delete(iid)
            for pib in self._people_filtered_pibs:
                p = self._people_cache.get(pib) or {}
                blocks = len(p.get("mentions") or [])
                records = len(p.get("records") or [])
                # Tag PІБ row by event types present for this person
                has_arrival = False
                has_departure = False
                for mm in (p.get("mentions") or []):
                    et0 = (mm.get("event_type") or "").strip().lower()
                    if et0 == "arrival":
                        has_arrival = True
                    elif et0 == "departure":
                        has_departure = True
                    if has_arrival and has_departure:
                        break
                if has_arrival and has_departure:
                    tags = ("BOTH",)
                elif has_arrival:
                    tags = ("ARRIVAL",)
                elif has_departure:
                    tags = ("DEPARTURE",)
                else:
                    tags = ()
                self.people_tree.insert("", "end", values=(pib, blocks, records), tags=tags)

        if select_first and self._people_filtered_pibs and hasattr(self, "people_tree"):
            first = self.people_tree.get_children()[0]
            self.people_tree.selection_set(first)
            self.people_tree.focus(first)
            self.show_selected_person()
        elif not self._people_filtered_pibs:
            self.person_text.delete("1.0", "end")
            self.person_text.insert("end", "Немає збігів.\n")


    def _parse_dt(self, s):
        try:
            d, m, y = s.split(".")
            return (int(y), int(m), int(d))
        except Exception:
            return (9999, 12, 31)

    def _get_selected_pib(self):
        if not hasattr(self, "people_tree"):
            return None
        sel = self.people_tree.selection()
        if not sel:
            return None
        vals = self.people_tree.item(sel[0], "values")
        if not vals:
            return None
        return vals[0]

    def _resolve_docx_path(self, doc_name: str):
        """Try to locate the DOCX by name under current input selection and default docs dir."""
        if not doc_name:
            return None
        if doc_name in self._docx_by_name_cache and os.path.exists(self._docx_by_name_cache[doc_name]):
            return self._docx_by_name_cache[doc_name]

        # candidate roots
        roots = []
        sel = (self.selected_path.get() or "").strip()
        if sel:
            if os.path.isdir(sel):
                roots.append(sel)
            elif os.path.isfile(sel):
                roots.append(os.path.dirname(sel))
        roots.append(DOCS_DIR)

        target = doc_name.lower()
        for root in roots:
            if not os.path.isdir(root):
                continue
            for r, _, fs in os.walk(root):
                for f in fs:
                    if f.lower() == target:
                        full = os.path.join(r, f)
                        self._docx_by_name_cache[doc_name] = full
                        return full
        return None

    def open_selected_docx_source(self):
        """Open DOCX source for selected fragment or selected doc in the left panel."""
        doc_name = None
        # prefer fragment selection
        if hasattr(self, "frag_tree"):
            sel = self.frag_tree.selection()
            if sel and sel[0] in self._frag_meta:
                doc_name = self._frag_meta[sel[0]].get("doc")
        # else doc panel selection
        if not doc_name and hasattr(self, "docs_tree"):
            dsel = self.docs_tree.selection()
            if dsel:
                meta = (self._docs_meta.get(dsel[0]) or {}) if hasattr(self, "_docs_meta") else {}
                if meta.get("type") in ("doc", "block"):
                    doc_name = meta.get("doc")
        # else current filter
        if not doc_name:
            doc_name = self._current_doc_filter

        if not doc_name:
            messagebox.showwarning("Немає документа", "Оберіть фрагмент або документ.")
            return
        path = self._resolve_docx_path(doc_name)
        if not path:
            messagebox.showwarning("Не знайдено", f"Не можу знайти файл: {doc_name}\n\nСпробуйте покласти його у input/documents або оберіть папку у вкладці 'Вхід'.")
            return
        open_path(path)

    def _event_match_regex(self, event_type: str):
        et = (event_type or "").strip().lower()
        if et == "arrival":
            return re.compile(r"(прибув\w*|прибут\w*|зарахув\w*|призван\w*|направлен\w*)", re.IGNORECASE)
        if et == "departure":
            return re.compile(
                r"(звільн\w*|виведен\w*|переведен\w*|виключ\w*|загиб\w*|помер\w*|призупинен\w*|полон\w*|сзч)",
                re.IGNORECASE,
            )
        return re.compile(r"(звільн\w*|прибув\w*|виведен\w*|зарахув\w*)", re.IGNORECASE)

    def _select_fragment_row(self, doc: str, block_id: int):
        if not hasattr(self, "frag_tree"):
            return
        for iid, meta in (self._frag_meta or {}).items():
            if meta.get("doc") == doc and int(meta.get("block_id") or 0) == int(block_id):
                self.frag_tree.selection_set(iid)
                self.frag_tree.focus(iid)
                self.frag_tree.see(iid)
                return

    def _on_docs_tree_select(self):
        """Docs panel behavior:

        - select '(усі)'  -> clear doc filter
        - select doc node -> set filter to this doc
        - select block node -> ensure filter doc + jump to block
        """
        if not hasattr(self, "docs_tree"):
            return
        sel = self.docs_tree.selection()
        if not sel:
            return
        meta = (self._docs_meta or {}).get(sel[0]) or {}
        t = meta.get("type")
        pib = self._current_person_pib
        person = self._people_cache.get(pib) if pib else None
        if not person:
            return

        if t == "all":
            if self._current_doc_filter is None:
                return
            self._current_doc_filter = None
            self._render_person_into_ui(pib, person, doc_filter=None)
            return

        if t == "doc":
            doc = meta.get("doc")
            if not doc:
                return
            if self._current_doc_filter == doc:
                return
            self._current_doc_filter = doc
            self._render_person_into_ui(pib, person, doc_filter=doc)
            return

        if t == "block":
            doc = meta.get("doc")
            bid = meta.get("block_id")
            if not doc or bid is None:
                return
            bid = int(bid)
            if self._current_doc_filter != doc:
                self._current_doc_filter = doc
                self._render_person_into_ui(pib, person, doc_filter=doc)
            self._select_fragment_row(doc, bid)
            self.jump_to_selected_fragment()

    def _on_docs_tree_open_or_jump(self):
        """Double click: open the underlying DOCX (and keep selection logic)."""
        if not hasattr(self, "docs_tree"):
            return
        sel = self.docs_tree.selection()
        if not sel:
            return
        meta = (self._docs_meta or {}).get(sel[0]) or {}
        if meta.get("type") in ("doc", "block"):
            doc_name = meta.get("doc")
            if doc_name:
                path = self._resolve_docx_path(doc_name)
                if path:
                    open_path(path)

    def find_in_person_text(self, next_: bool = True):
        q = (self.find_var.get() if hasattr(self, "find_var") else "").strip()
        if not q:
            return
        self.person_text.tag_remove("FIND", "1.0", "end")

        # start search from insert mark
        start = self.person_text.index(tk.INSERT)
        if next_:
            pos = self.person_text.search(q, start, nocase=1, stopindex="end")
            if not pos:
                pos = self.person_text.search(q, "1.0", nocase=1, stopindex="end")
        else:
            pos = self.person_text.search(q, start, nocase=1, stopindex="1.0", backwards=True)
            if not pos:
                pos = self.person_text.search(q, "end", nocase=1, stopindex="1.0", backwards=True)

        if not pos:
            return
        end = f"{pos}+{len(q)}c"
        self.person_text.tag_add("FIND", pos, end)
        self.person_text.mark_set(tk.INSERT, end)
        self.person_text.see(pos)

    def clear_find_highlight(self):
        self.person_text.tag_remove("FIND", "1.0", "end")

    def jump_to_selected_fragment(self):
        if not hasattr(self, "frag_tree"):
            return
        sel = self.frag_tree.selection()
        if not sel:
            return
        meta = self._frag_meta.get(sel[0]) or {}
        doc = meta.get("doc")
        bid = meta.get("block_id")
        if not doc or bid is None:
            return
        key = (doc, int(bid))
        pos = self._mention_pos.get(key)
        if not pos:
            return

        # clear previous highlights
        self.person_text.tag_remove("BLOCK_ACTIVE", "1.0", "end")
        self.person_text.tag_remove("MATCH_ACTIVE", "1.0", "end")
        self.person_text.tag_remove("PIB_MATCH", "1.0", "end")
        self.person_text.tag_remove("CAT_MATCH", "1.0", "end")

        self.person_text.tag_add("BLOCK_ACTIVE", pos["block_start"], pos["block_end"])
        self.person_text.see(pos["block_start"])
        self.person_text.mark_set(tk.INSERT, pos["block_start"])

        raw = meta.get("text") or ""
        rx = self._event_match_regex(meta.get("event_type") or "")
        m = rx.search(raw)
        if m:
            s_off, e_off = m.span()
            s_idx = f"{pos['content_start']}+{s_off}c"
            e_idx = f"{pos['content_start']}+{e_off}c"
            self.person_text.tag_add("MATCH_ACTIVE", s_idx, e_idx)
            self.person_text.see(s_idx)

        # Highlight the person's PIB itself (first match in the block)
        pib = self._current_person_pib or ""
        pib = pib.strip()
        if pib:
            try:
                pm = re.search(re.escape(pib), raw, flags=re.IGNORECASE)
            except re.error:
                pm = None
            if pm:
                ps, pe = pm.span()
                ps_idx = f"{pos['content_start']}+{ps}c"
                pe_idx = f"{pos['content_start']}+{pe}c"
                self.person_text.tag_add("PIB_MATCH", ps_idx, pe_idx)

        # Highlight category match span (rule-based)
        cat_span = meta.get("cat_span")
        if cat_span and isinstance(cat_span, (list, tuple)) and len(cat_span) == 2:
            try:
                cs, ce = int(cat_span[0]), int(cat_span[1])
            except Exception:
                cs, ce = None, None
            if cs is not None and ce is not None and ce > cs:
                cs_idx = f"{pos['content_start']}+{cs}c"
                ce_idx = f"{pos['content_start']}+{ce}c"
                self.person_text.tag_add("CAT_MATCH", cs_idx, ce_idx)
                self.person_text.see(cs_idx)


    def _classify_block_category(self, block_text: str, event_type: str):
        """Return (category_label, rule_id, span_tuple) for a block."""
        if not self._cat_engine:
            return (None, None, None)
        try:
            res = self._cat_engine.classify(block_text or "", event_hint=event_type)
            return (res.get("category"), res.get("rule_id"), res.get("match_span"))
        except Exception:
            return (None, None, None)
    def _render_person_into_ui(self, pib: str, person: dict, doc_filter: str | None = None):
        """Render per-person view: docs list, fragments list and full block text.

        - docs_tree: shows DOCX -> count of blocks
        - frag_tree: shows one row per block (variant A)
        - person_text: prints full blocks, keeps index mapping for jump/highlight
        """

        mentions = person.get("mentions", []) or []
        ranks = person.get("ranks", []) or []
        variants = person.get("variants", []) or []

        def mention_key(m):
            dt = (m.get("date") or "").strip()
            dkey = self._parse_dt(dt) if dt else (9999, 12, 31)
            doc = (m.get("doc") or "").strip()
            bid = int(m.get("block_id") or 0)
            return (doc, dkey, bid)

        all_mentions = sorted(mentions, key=mention_key)

        # Precompute categories for mentions (cache into dict)
        # Prefer parser-provided category fields if present (stable spans for highlight).
        for _m in all_mentions:
            if _m.get("_cat") is None and "_cat" in _m:
                continue

            has_parsed = any(k in _m for k in ("category", "rule_id", "cat_span"))
            parsed_cat = _m.get("category")
            parsed_span = _m.get("cat_span")
            parsed_rule = _m.get("rule_id")

            if has_parsed and (parsed_cat is not None or parsed_span is not None or parsed_rule is not None):
                _m["_cat"] = parsed_cat
                _m["_cat_rule"] = parsed_rule
                _m["_cat_span"] = parsed_span
                continue

            txt = (_m.get("text") or "").strip()
            et0 = (_m.get("event_type") or "").strip()
            cat0, rule0, span0 = self._classify_block_category(txt, et0)
            _m["_cat"] = cat0
            _m["_cat_rule"] = rule0
            _m["_cat_span"] = span0

        base_mentions = all_mentions
        # Apply event/category filters from category tree
        if self._current_event_filter in ("arrival", "departure"):
            base_mentions = [m for m in base_mentions if (m.get("event_type") or "").strip() == self._current_event_filter]
        if self._current_category_filter:
            base_mentions = [m for m in base_mentions if (m.get("_cat") or "") == self._current_category_filter]
        filtered_mentions = [m for m in base_mentions if (m.get("doc") or "") == doc_filter] if doc_filter else base_mentions

        # ---- docs panel (DOCX -> blocks)
        self._docs_meta = {}
        if hasattr(self, "docs_tree"):
            for iid in self.docs_tree.get_children():
                self.docs_tree.delete(iid)

            # Count blocks per doc
            counts: dict[str, int] = {}
            per_doc: dict[str, list[dict]] = {}
            for m in all_mentions:
                d = (m.get("doc") or "").strip() or "(невідомо)"
                counts[d] = counts.get(d, 0) + 1
                per_doc.setdefault(d, []).append(m)

            total = sum(counts.values())
            all_iid = self.docs_tree.insert("", "end", text="(усі)", values=(total,))
            self._docs_meta[all_iid] = {"type": "all"}

            for d in sorted(per_doc.keys(), key=lambda x: x.lower()):
                doc_iid = self.docs_tree.insert("", "end", text=d, values=(counts.get(d, 0),))
                self._docs_meta[doc_iid] = {"type": "doc", "doc": d}
                # Add blocks as children
                for m in sorted(per_doc[d], key=mention_key):
                    bid = int(m.get("block_id") or 0)
                    dt = (m.get("date") or "").strip()
                    et = (m.get("event_type") or "").strip()
                    cat, cat_rule, cat_span = m.get("_cat"), m.get("_cat_rule"), m.get("_cat_span")
                    lbl = f"{dt or '—'} | {et or '—'} | {cat or '—'} | блок {bid}"
                    child = self.docs_tree.insert(doc_iid, "end", text=lbl, values=("",), tags=_ui_tag_for_event_and_category(et, cat))
                    self._docs_meta[child] = {
                        "type": "block",
                        "doc": d,
                        "block_id": bid,
                        "date": dt,
                        "event_type": et,
                        "text": (m.get("text") or "").strip(),
                    }

            # Select current filter doc if set
            if doc_filter:
                for iid in self.docs_tree.get_children(""):
                    meta = self._docs_meta.get(iid) or {}
                    if meta.get("type") == "doc" and meta.get("doc") == doc_filter:
                        self.docs_tree.selection_set(iid)
                        self.docs_tree.focus(iid)
                        self.docs_tree.item(iid, open=True)
                        break

        # ---- fragments table
        if hasattr(self, "frag_tree"):
            for iid in self.frag_tree.get_children():
                self.frag_tree.delete(iid)
        self._frag_meta = {}

        for idx, m in enumerate(filtered_mentions, start=1):
            doc = (m.get("doc") or "").strip()
            bid = int(m.get("block_id") or 0)
            dt = (m.get("date") or "").strip()
            et = (m.get("event_type") or "").strip()
            text = (m.get("text") or "").strip()
            cat, cat_rule, cat_span = m.get("_cat"), m.get("_cat_rule"), m.get("_cat_span")

            snippet = " ".join(text.split())
            if len(snippet) > 220:
                snippet = snippet[:217] + "..."
            iid = self.frag_tree.insert("", "end", values=(idx, dt, et, cat, doc, bid, snippet), tags=_ui_tag_for_event_and_category(et, cat))
            self._frag_meta[iid] = {
                "doc": doc,
                "block_id": bid,
                "date": dt,
                "event_type": et,
                "text": text,
                "category": cat,
                "cat_rule_id": cat_rule,
                "cat_span": cat_span,
            }

        # ---- text area
        self.person_text.tag_remove("BLOCK_ACTIVE", "1.0", "end")
        self.person_text.tag_remove("MATCH_ACTIVE", "1.0", "end")
        self.clear_find_highlight()

        self.person_text.configure(state="normal")
        self.person_text.delete("1.0", "end")
        self._mention_pos = {}

        # Header
        self.person_text.insert("end", f"ПІБ: {pib}\n", "HDR")
        if ranks:
            self.person_text.insert("end", f"Звання: {', '.join(ranks)}\n")
        if variants:
            self.person_text.insert("end", "Варіанти: " + "; ".join([str(v) for v in variants if v]) + "\n")
        if doc_filter:
            self.person_text.insert("end", f"Фільтр документа: {doc_filter}\n")
        self.person_text.insert("end", "-" * 88 + "\n")

        if not filtered_mentions:
            self.person_text.insert("end", "Немає блоків для показу (перевірте фільтр).\n")
            return

        for idx, m in enumerate(filtered_mentions, start=1):
            doc = (m.get("doc") or "").strip()
            bid = int(m.get("block_id") or 0)
            dt = (m.get("date") or "").strip()
            et = (m.get("event_type") or "").strip()
            sec = m.get("section_path") or []
            sec_s = " → ".join([s for s in sec if s])
            text = (m.get("text") or "").strip()

            block_start = self.person_text.index("end-1c")
            cat, cat_rule, cat_span = m.get("_cat"), m.get("_cat_rule"), m.get("_cat_span")
            header = f"[{idx}] {dt or '—'} | {et or '—'} | {cat or '—'} | {doc} | блок {bid}\n"
            self.person_text.insert("end", header, ("HDR",) + _ui_tag_for_event_and_category(et, cat))
            if sec_s:
                self.person_text.insert("end", f"  Контекст: {sec_s}\n")

            content_start = self.person_text.index("end-1c")
            self.person_text.insert("end", text + "\n\n")
            block_end = self.person_text.index("end-1c")
            content_end = f"{content_start}+{len(text)}c"
            self._mention_pos[(doc, bid)] = {
                "block_start": block_start,
                "block_end": block_end,
                "content_start": content_start,
                "content_end": content_end,
            }


    def _clear_category_filter(self):
        if self._current_person_pib and self._current_person_pib in self._people_cache:
            self._current_event_filter = None
            self._current_category_filter = None
            self._render_person_into_ui(self._current_person_pib, self._people_cache[self._current_person_pib], doc_filter=self._current_doc_filter)

    def _on_cats_tree_select(self):
        if not hasattr(self, "cats_tree"):
            return
        sel = self.cats_tree.selection()
        if not sel:
            return
        meta = (self._cats_meta or {}).get(sel[0]) or {}
        pib = self._current_person_pib
        person = self._people_cache.get(pib) if pib else None
        if not person:
            return

        if meta.get("type") == "all":
            self._current_event_filter = None
            self._current_category_filter = None
        elif meta.get("type") == "event":
            self._current_event_filter = meta.get("event")
            self._current_category_filter = None
        elif meta.get("type") == "cat":
            self._current_event_filter = meta.get("event")
            self._current_category_filter = meta.get("label")

        self._render_person_into_ui(pib, person, doc_filter=self._current_doc_filter)

    def show_selected_person(self):
        pib = self._get_selected_pib()
        if not pib:
            return
        person = self._people_cache.get(pib)
        if not person:
            return

        self._current_person_pib = pib
        self._current_doc_filter = None
        self._current_event_filter = None
        self._current_category_filter = None
        self._render_person_into_ui(pib, person, doc_filter=None)

    def export_person_txt(self):
        pib = self._get_selected_pib()
        if not pib:
            messagebox.showwarning("Нема вибору", "Оберіть людину зі списку.")
            return
        safe = "".join([c for c in pib if c.isalnum() or c in " _-"]).strip().replace(" ", "_")
        path = os.path.join(self.exports_txt_dir, f"{safe}.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write(self.person_text.get("1.0", "end").strip() + "\n")
        messagebox.showinfo("Збережено", f"TXT: {path}")

    def export_person_json(self):
        pib = self._get_selected_pib()
        if not pib:
            messagebox.showwarning("Нема вибору", "Оберіть людину зі списку.")
            return
        person = self._people_cache.get(pib)
        if not person:
            return
        safe = "".join([c for c in pib if c.isalnum() or c in " _-"]).strip().replace(" ", "_")
        path = os.path.join(self.exports_json_dir, f"{safe}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(person, f, ensure_ascii=False, indent=2)
        messagebox.showinfo("Збережено", f"JSON: {path}")


    def export_all_people_zip(self):
        # Export all people to TXT+JSON and pack into one ZIP
        if not os.path.exists(self.combined_path):
            messagebox.showwarning("Немає даних", "Спочатку обробіть DOCX, щоб зʼявився combined_v2.json.")
            return
        if not self._people_cache:
            self.refresh_people_view()
        if not self._people_cache:
            messagebox.showwarning("Немає даних", "Список людей порожній.")
            return

        ts = time.strftime("%Y%m%d_%H%M%S")
        out_dir = os.path.join(OUTPUT, "exports", "people_all")
        os.makedirs(out_dir, exist_ok=True)
        zip_path = os.path.join(out_dir, f"people_all_{ts}.zip")

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for pib, person in sorted(self._people_cache.items(), key=lambda kv: kv[0].upper()):
                safe = "".join([c for c in pib if c.isalnum() or c in " _-"]).strip().replace(" ", "_")
                z.writestr(f"json/{safe}.json", json.dumps(person, ensure_ascii=False, indent=2))
                txt = self._render_person_text(pib, person)
                z.writestr(f"txt/{safe}.txt", txt + "\n")

        messagebox.showinfo("Готово", f"ZIP експортовано:\n{zip_path}")
        open_path(out_dir)

    def _render_person_text(self, pib, person):
        variants = person.get("variants", [])
        ranks = person.get("ranks", [])
        records = person.get("records", [])
        mentions = person.get("mentions", [])

        def rec_key(r):
            dt = r.get("arrival_date") or r.get("departure_date") or ""
            dkey = self._parse_dt(dt) if dt else (9999, 12, 31)
            order = (r.get("order_number") or "")
            et = (r.get("event_type") or "")
            return (dkey, order, et)

        records_sorted = sorted(records, key=rec_key)

        # Mentions: keep full block text and show it as "пункти наказу"
        def mention_key(m):
            dt = (m.get("date") or "")
            dkey = self._parse_dt(dt) if dt else (9999, 12, 31)
            doc = (m.get("doc") or "")
            bid = int(m.get("block_id") or 0)
            return (doc, dkey, bid)

        mentions_sorted = sorted(mentions, key=mention_key)

        out = []
        out.append(f"ПІБ: {pib}")
        if ranks:
            out.append(f"Звання: {', '.join(ranks)}")
        if variants:
            out.append("Варіанти написання: " + "; ".join([str(v) for v in variants if v]))
        if mentions_sorted:
            out.append(f"Знайдено пунктів/блоків у наказах: {len(mentions_sorted)}")
        out.append("-" * 70)

        if mentions_sorted:
            # Full structured view (preferred)
            for idx, m in enumerate(mentions_sorted, start=1):
                doc = (m.get("doc") or "").strip()
                bid = m.get("block_id")
                et = (m.get("event_type") or "").strip()
                dt = (m.get("date") or "").strip()
                sec = m.get("section_path") or []
                sec_s = " → ".join([s for s in sec if s])
                text = (m.get("text") or "").strip()

                header = [f"Пункт {idx}"]
                if doc:
                    header.append(f"Документ: {doc}")
                if bid:
                    header.append(f"Блок: {bid}")
                if dt:
                    header.append(f"Дата: {dt}")
                if et:
                    header.append(f"Подія: {et}")
                out.append(" | ".join(header))
                if sec_s:
                    out.append(f"  Контекст: {sec_s}")
                if text:
                    out.append("  Текст (повністю):")
                    # indent block lines
                    for ln in text.splitlines():
                        out.append("    " + ln)
                out.append("")

        else:
            # Fallback: show records (legacy)
            for r in records_sorted:
                et = (r.get("event_type") or "").strip()
                date = r.get("arrival_date") or r.get("departure_date") or ""
                pos = r.get("position") or ""
                frm = r.get("from_unit") or ""
                to = r.get("to_unit") or ""
                reason = r.get("departure_reason") or ""
                src = r.get("arrival_source") or ""
                order = r.get("order_number") or ""
                note = r.get("note") or ""

                line = []
                if date: line.append(date)
                if et: line.append(et)
                if order: line.append(order)
                header = " | ".join(line) if line else "Запис"
                out.append(header)
                if pos: out.append(f"  Посада: {pos}")
                if frm: out.append(f"  Звідки: {frm}")
                if to: out.append(f"  Куди: {to}")
                if src: out.append(f"  Джерело прибуття: {src}")
                if reason: out.append(f"  Причина вибуття: {reason}")
                if note:
                    out.append("  Текст (повністю):")
                    for ln in str(note).splitlines():
                        out.append("    " + ln)
                out.append("")

        return "\n".join(out)

    def _build_settings(self):
        frm = ttk.Frame(self.tab_settings, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Шаблон Excel (N_fixed.xlsx):").grid(row=0, column=0, sticky="w")
        self.tpl_var = tk.StringVar(value=self.template_path)
        ttk.Entry(frm, textvariable=self.tpl_var, width=90).grid(row=1, column=0, sticky="we")
        ttk.Button(frm, text="Обрати", command=self._pick_template).grid(row=1, column=1, padx=6)

        ttk.Label(frm, text="Папка з DOCX за замовчуванням:").grid(row=2, column=0, sticky="w", pady=(10,0))
        self.docs_var = tk.StringVar(value=DOCS_DIR)
        ttk.Entry(frm, textvariable=self.docs_var, width=90).grid(row=3, column=0, sticky="we")
        ttk.Button(frm, text="Відкрити", command=lambda: open_path(DOCS_DIR)).grid(row=3, column=1, padx=6)

        ttk.Label(frm, text="Вихід Excel:").grid(row=4, column=0, sticky="w", pady=(10,0))
        self.out_var = tk.StringVar(value=self.output_dir)
        ttk.Entry(frm, textvariable=self.out_var, width=90).grid(row=5, column=0, sticky="we")
        ttk.Button(frm, text="Відкрити", command=lambda: open_path(self.output_dir)).grid(row=5, column=1, padx=6)

        ttk.Separator(frm).grid(row=6, column=0, columnspan=2, sticky="we", pady=12)
        ttk.Button(frm, text="Застосувати (шаблон/вихід)", command=self._apply_settings).grid(row=7, column=0, sticky="w")

        # --- UI settings (font + highlight colors)
        ui = ttk.LabelFrame(frm, text="UI — шрифт + підсвітка", padding=10)
        ui.grid(row=8, column=0, columnspan=2, sticky="nsew", pady=(12, 0))

        cur_font = self.ui_settings.get("font") or {}
        self.ui_font_family_var = tk.StringVar(value=cur_font.get("family") or "")
        self.ui_font_size_var = tk.StringVar(value=str(cur_font.get("size") or 12))
        self.ui_wrap_var = tk.BooleanVar(value=bool(self.ui_settings.get("wrap_person_text", True)))

        ttk.Label(ui, text="Font family:").grid(row=0, column=0, sticky="w")
        fam_values = sorted({f for f in tkfont.families() if f})
        fam_box = ttk.Combobox(ui, textvariable=self.ui_font_family_var, values=fam_values, width=34)
        fam_box.grid(row=0, column=1, sticky="w", padx=(8, 0))

        ttk.Label(ui, text="Size:").grid(row=0, column=2, sticky="w", padx=(18, 0))
        size_sp = ttk.Spinbox(ui, from_=8, to=28, textvariable=self.ui_font_size_var, width=5)
        size_sp.grid(row=0, column=3, sticky="w", padx=(8, 0))

        ttk.Checkbutton(ui, text="Wrap (person text)", variable=self.ui_wrap_var).grid(row=0, column=4, sticky="w", padx=(18, 0))

        ttk.Separator(ui).grid(row=1, column=0, columnspan=5, sticky="we", pady=10)

        ttk.Label(ui, text="Кольори підсвітки:").grid(row=2, column=0, sticky="w")
        colors = self.ui_settings.get("colors") or {}
        self._ui_color_buttons = {}
        tags = [
            ("BLOCK_ACTIVE", "Block"),
            ("MATCH_ACTIVE", "Match"),
            ("CAT_MATCH", "Category"),
            ("PIB_MATCH", "PIB"),
            ("FIND", "Find"),
        ]
        row = 3
        col = 0
        for tag, label in tags:
            if col >= 5:
                row += 1
                col = 0
            b = tk.Button(ui, text=label, width=10, relief="raised",
                          bg=colors.get(tag, "#dddddd"),
                          command=lambda t=tag: self._pick_ui_color(t))
            b.grid(row=row, column=col, sticky="w", padx=(0, 8), pady=(4, 4))
            self._ui_color_buttons[tag] = b
            col += 1

        btn_row = row + 1
        ttk.Button(ui, text="Застосувати UI", command=self._apply_ui_settings_from_controls).grid(row=btn_row, column=0, sticky="w", pady=(10, 0))
        ttk.Button(ui, text="Скинути кольори", command=self._reset_ui_colors).grid(row=btn_row, column=1, sticky="w", pady=(10, 0))

        ui.columnconfigure(1, weight=1)
        frm.columnconfigure(0, weight=1)

    def _pick_template(self):
        p = filedialog.askopenfilename(filetypes=[("XLSX","*.xlsx")])
        if p:
            self.tpl_var.set(p)

    def _apply_settings(self):
        # only template and output dir are mutable in this simple build
        tpl = self.tpl_var.get().strip()
        out = self.out_var.get().strip()
        if not os.path.exists(tpl):
            messagebox.showerror("Помилка", "Шаблон не знайдено.")
            return
        self.template_path = tpl
        self.output_dir = out
        os.makedirs(self.output_dir, exist_ok=True)
        messagebox.showinfo("OK", "Налаштування застосовано.")
        self._scan()

    def _pick_ui_color(self, tag: str):
        """Choose a highlight color for a given tag (settings tab)."""
        try:
            initial = (self.ui_settings.get("colors") or {}).get(tag) or "#ffffff"
            color = colorchooser.askcolor(initialcolor=initial, title=f"Колір — {tag}")[1]
            if not color:
                return
            self.ui_settings.setdefault("colors", {})[tag] = color
            # update button preview if exists
            btn = getattr(self, "_ui_color_buttons", {}).get(tag) if hasattr(self, "_ui_color_buttons") else None
            if btn is not None:
                try:
                    btn.configure(bg=color)
                except Exception:
                    pass
            self._save_ui_settings()
            self._reconfigure_people_text_tags()
        except Exception:
            return

    def _apply_ui_settings_from_controls(self):
        """Apply font + wrap + highlight colors from Settings tab."""
        fam = (getattr(self, "ui_font_family_var", tk.StringVar(value="")).get() or "").strip()
        try:
            size = int((getattr(self, "ui_font_size_var", tk.StringVar(value="12")).get() or "12").strip())
        except Exception:
            size = 12
        wrap_on = bool(getattr(self, "ui_wrap_var", tk.BooleanVar(value=True)).get())

        if fam:
            self.ui_settings.setdefault("font", {})["family"] = fam
        self.ui_settings.setdefault("font", {})["size"] = max(8, min(28, int(size)))
        self.ui_settings["wrap_person_text"] = wrap_on
        self._save_ui_settings()

        # apply globally
        self._apply_global_style()
        for wname in ("person_text", "log", "unknown_units", "unknown_reasons", "unknown_arrivals"):
            w = getattr(self, wname, None)
            if w is not None:
                self._apply_text_widget_style(w)

        if hasattr(self, "person_text"):
            try:
                self.person_text.configure(wrap=("word" if wrap_on else "none"))
            except Exception:
                pass
        self._reconfigure_people_text_tags()
        messagebox.showinfo("OK", "UI налаштування застосовано.")

    def _reset_ui_colors(self):
        defaults = {
            "BLOCK_ACTIVE": "#fff3b0",
            "MATCH_ACTIVE": "#ffe066",
            "CAT_MATCH": "#ffd8a8",
            "PIB_MATCH": "#d3f9d8",
            "FIND": "#cfe8ff",
        }
        self.ui_settings.setdefault("colors", {}).update(defaults)
        # refresh preview buttons
        if hasattr(self, "_ui_color_buttons"):
            for k, b in self._ui_color_buttons.items():
                if k in defaults:
                    try:
                        b.configure(bg=defaults[k])
                    except Exception:
                        pass
        self._save_ui_settings()
        self._reconfigure_people_text_tags()

if __name__ == "__main__":
    App().mainloop()
