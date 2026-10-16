import os
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import openpyxl
import pandas as pd

# ضبط المظهر كـ Light كاختيار أولي
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")


# ==================== نافذة التعديل وإضافة موظف ====================
class EditTeacherWindow(ctk.CTkToplevel):

  def __init__(
      self, parent, ppr, teacher_data, on_save_callback, is_new=False
  ):
    super().__init__(parent)
    self.parent = parent
    self.ppr = str(ppr) if ppr else ""
    self.teacher_data = teacher_data
    self.on_save_callback = on_save_callback
    self.is_new = is_new

    title_text = (
        "إضافة موظف جديد"
        if is_new
        else f"تعديل بيانات الموظف : {self.ppr}"
    )
    self.title(title_text)
    self.geometry("860x720")
    self.grab_set()

    self.grid_columnconfigure(0, weight=1)
    self.grid_rowconfigure(1, weight=1)

    nom_ar = self.teacher_data.get("NOM_AR", "")
    nom_fr = self.teacher_data.get("NOM_FR", "")
    header_title = (
        "إدراج سجل إداري جديد"
        if is_new
        else f"بطاقة الموظف : {nom_ar} ({nom_fr})"
    )

    self.header_frame = ctk.CTkFrame(self, fg_color="#1F538D", corner_radius=10)
    self.header_frame.grid(row=0, column=0, padx=15, pady=12, sticky="ew")

    self.lbl_head = ctk.CTkLabel(
        self.header_frame,
        text=header_title,
        font=ctk.CTkFont(family="Arial", size=16, weight="bold"),
        text_color="white",
    )
    self.lbl_head.pack(pady=10)

    self.scroll_frame = ctk.CTkScrollableFrame(
        self, label_text="البيانات الشخصية والإدارية"
    )
    self.scroll_frame.grid(row=1, column=0, padx=15, pady=5, sticky="nsew")
    self.scroll_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

    self.entries = {}

    fields_definition = [
        ("PPR", "رقم التأجير :"),
        ("CIN", "رقم البطاقة الوطنية :"),
        ("NOM_AR", "الاسم الكامل بالعربية :"),
        ("PRENOM_AR", "الاسم الشخصي :"),
        ("NOM_FR", "الاسم الكامل بالفرنسية :"),
        ("SEX", "الجنس (M/F) :"),
        ("SIT_FAM", "الحالة العائلية :"),
        ("DATE_NAISSANCE", "تاريخ الازدياد :"),
        ("المهام", "المهمة المسندة :"),
        ("FONCTION", "المهمة بالفرنسية :"),
        ("الاطار", "الإطار الحالي :"),
        ("الدرجة", "الدرجة :"),
        ("C_ETAB", "كود المؤسسة أو الفرعية :"),
        ("CD_ETAB", "كود التعيين الإداري :"),
        ("الوحدة المدرسية.1", "اسم الفرعية :"),
        ("المؤسسة", "المجموعة المدرسية المركزية :"),
        ("DATE_DEBUT_AFF", "تاريخ التعيين بالمؤسسة :"),
        ("DATE_ANC_ADM", "تاريخ التوظيف الإداري :"),
        ("DATE_EFFET_GD", "تاريخ مفعول الدرجة :"),
        ("DATE_ANC_GD", "أقدمية الدرجة :"),
        ("DATE_EFFET_ELO", "تاريخ مفعول الرتبة :"),
        ("DATE_ANC_ELO", "أقدمية الرتبة :"),
    ]

    row_idx = 0
    for i, (col_key, label_text) in enumerate(fields_definition):
      grid_col_lbl = 0 if i % 2 == 0 else 2
      grid_col_ent = 1 if i % 2 == 0 else 3

      lbl = ctk.CTkLabel(
          self.scroll_frame,
          text=label_text,
          anchor="w",
          font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
      )
      lbl.grid(row=row_idx, column=grid_col_lbl, padx=12, pady=7, sticky="w")

      val = str(self.teacher_data.get(col_key, "") or "")
      entry = ctk.CTkEntry(
          self.scroll_frame, width=220, font=ctk.CTkFont(family="Arial", size=12)
      )
      entry.insert(0, val)
      entry.grid(
          row=row_idx, column=grid_col_ent, padx=10, pady=7, sticky="ew"
      )

      self.entries[col_key] = entry

      if i % 2 == 1:
        row_idx += 1

    self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
    self.btn_frame.grid(row=2, column=0, padx=15, pady=12, sticky="ew")

    self.btn_save = ctk.CTkButton(
        self.btn_frame,
        text="حفظ التعديلات",
        fg_color="#27AE60",
        hover_color="#1E8449",
        font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
        width=150,
        command=self.save_data,
    )
    self.btn_save.pack(side="right", padx=10)

    self.btn_cancel = ctk.CTkButton(
        self.btn_frame,
        text="إلغاء الأمر",
        fg_color="#7F8C8D",
        hover_color="#95A5A6",
        font=ctk.CTkFont(family="Arial", size=13),
        width=100,
        command=self.destroy,
    )
    self.btn_cancel.pack(side="right", padx=5)

  def save_data(self):
    updated_dict = {}
    for col_key, entry in self.entries.items():
      updated_dict[col_key] = entry.get().strip()

    ppr_val = updated_dict.get("PPR", "").strip()
    if not ppr_val:
      messagebox.showwarning("تنبيه", "يرجى إدخال رقم التأجير أولاً")
      return

    self.on_save_callback(ppr_val, updated_dict, self.is_new)
    messagebox.showinfo("تأكيد", "تم حفظ وتحديث المعطيات بنجاح")
    self.destroy()


# ==================== التطبيق الرئيسي ====================
class RH2026App(ctk.CTk):

  def __init__(self):
    super().__init__()

    self.title("منظومة تدبير الموارد البشرية 2026")
    self.geometry("1400x840")

    self.df = pd.DataFrame()
    self.current_filtered_df = pd.DataFrame()
    self.active_file_path = "RH2026.xlsx"

    # نظام التصفح بالصفحات
    self.page_size = 50
    self.current_page = 1
    self.total_pages = 1

    self.grid_columnconfigure(1, weight=1)
    self.grid_rowconfigure(0, weight=1)

    # ---------------- القائمة الجانبية ----------------
    self.sidebar_frame = ctk.CTkFrame(self, width=250, corner_radius=0)
    self.sidebar_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

    self.logo_label = ctk.CTkLabel(
        self.sidebar_frame,
        text="إدارة الموارد البشرية\nالموسم الدراسي 2026",
        font=ctk.CTkFont(family="Arial", size=17, weight="bold"),
    )
    self.logo_label.pack(pady=15)

    self.line1 = ctk.CTkFrame(self.sidebar_frame, height=2, fg_color="#BDC3C7")
    self.line1.pack(fill="x", padx=15, pady=5)

    self.btn_load_file = ctk.CTkButton(
        self.sidebar_frame,
        text="فتح ملف المعطيات",
        font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
        command=self.load_excel_manually,
    )
    self.btn_load_file.pack(pady=8, padx=15, fill="x")

    self.btn_add_emp = ctk.CTkButton(
        self.sidebar_frame,
        text="إضافة موظف جديد",
        fg_color="#2980B9",
        hover_color="#1F618D",
        font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
        command=self.open_add_dialog,
    )
    self.btn_add_emp.pack(pady=8, padx=15, fill="x")

    self.btn_save_source = ctk.CTkButton(
        self.sidebar_frame,
        text="حفظ التعديلات في الملف",
        fg_color="#D35400",
        hover_color="#A04000",
        font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
        command=self.save_to_source_excel,
    )
    self.btn_save_source.pack(pady=8, padx=15, fill="x")

    self.btn_export = ctk.CTkButton(
        self.sidebar_frame,
        text="تصدير النتائج إلى إكسيل",
        fg_color="#27AE60",
        hover_color="#1E8449",
        font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
        command=self.export_excel,
    )
    self.btn_export.pack(pady=8, padx=15, fill="x")

    self.lbl_status = ctk.CTkLabel(
        self.sidebar_frame,
        text="جاري فحص الملف...",
        text_color="#7F8C8D",
        wraplength=220,
        font=ctk.CTkFont(family="Arial", size=12),
    )
    self.lbl_status.pack(pady=10)

    # نمط المظهر
    self.lbl_theme = ctk.CTkLabel(
        self.sidebar_frame,
        text="مظهر الشاشة :",
        font=ctk.CTkFont(family="Arial", size=12),
    )
    self.lbl_theme.pack(side="bottom", pady=(0, 5))
    self.theme_switch = ctk.CTkOptionMenu(
        self.sidebar_frame,
        values=["فاتح", "داكن"],
        command=self.change_appearance,
        font=ctk.CTkFont(family="Arial", size=12),
    )
    self.theme_switch.pack(side="bottom", pady=(0, 15), padx=20, fill="x")

    # ---------------- المنطقة الرئيسية ----------------
    self.main_frame = ctk.CTkFrame(self)
    self.main_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
    self.main_frame.grid_columnconfigure(0, weight=1)
    self.main_frame.grid_rowconfigure(4, weight=1)

    # ---------------- شريط البحث والفلاتر ----------------
    self.filter_frame = ctk.CTkFrame(self.main_frame, corner_radius=10)
    self.filter_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=8)

    # 1. كود المؤسسة
    self.lbl_etab = ctk.CTkLabel(
        self.filter_frame,
        text="كود المؤسسة :",
        font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
    )
    self.lbl_etab.grid(row=0, column=0, padx=6, pady=8)
    self.entry_etab = ctk.CTkEntry(
        self.filter_frame,
        placeholder_text="03242Y مثلاً",
        width=130,
        font=ctk.CTkFont(family="Arial", size=13),
    )
    self.entry_etab.grid(row=0, column=1, padx=6, pady=8)
    self.entry_etab.bind("<Return>", lambda e: self.filter_by_etab())

    self.btn_filter_etab = ctk.CTkButton(
        self.filter_frame,
        text="بحث بالمؤسسة",
        font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
        width=120,
        command=self.filter_by_etab,
    )
    self.btn_filter_etab.grid(row=0, column=2, padx=6, pady=8)

    self.lbl_sep = ctk.CTkLabel(
        self.filter_frame, text="|", font=ctk.CTkFont(size=16), text_color="gray"
    )
    self.lbl_sep.grid(row=0, column=3, padx=6, pady=8)

    # 2. بحث سريع
    self.lbl_quick = ctk.CTkLabel(
        self.filter_frame,
        text="بحث سريع :",
        font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
    )
    self.lbl_quick.grid(row=0, column=4, padx=6, pady=8)
    self.entry_quick = ctk.CTkEntry(
        self.filter_frame,
        placeholder_text="رقم التأجير أو الاسم أو البطاقة...",
        width=230,
        font=ctk.CTkFont(family="Arial", size=12),
    )
    self.entry_quick.grid(row=0, column=5, padx=6, pady=8)
    self.entry_quick.bind("<Return>", lambda e: self.search_action())

    self.btn_search = ctk.CTkButton(
        self.filter_frame,
        text="بحث فوري",
        font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
        width=90,
        command=self.search_action,
    )
    self.btn_search.grid(row=0, column=6, padx=6, pady=8)

    self.btn_reset = ctk.CTkButton(
        self.filter_frame,
        text="عرض الكل",
        fg_color="#7F8C8D",
        hover_color="#95A5A6",
        font=ctk.CTkFont(family="Arial", size=12),
        width=80,
        command=self.show_all,
    )
    self.btn_reset.grid(row=0, column=7, padx=8, pady=8)

    # ---------------- بطاقات المؤشرات الرقمية (KPIs) ----------------
    self.kpi_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
    self.kpi_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=4)
    self.kpi_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

    self.card_total = self.create_kpi_card(
        self.kpi_frame, 0, "مجموع الموظفين", "0", "#2980B9"
    )
    self.card_females = self.create_kpi_card(
        self.kpi_frame, 1, "عدد الإناث", "0 (0%)", "#8E44AD"
    )
    self.card_males = self.create_kpi_card(
        self.kpi_frame, 2, "عدد الذكور", "0 (0%)", "#16A085"
    )
    self.card_etab = self.create_kpi_card(
        self.kpi_frame, 3, "المؤسسة الحالية", "جميع المؤسسات", "#D35400"
    )

    # ---------------- جدول إحصاء المهام ----------------
    self.stats_frame = ctk.CTkFrame(self.main_frame, corner_radius=8)
    self.stats_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=5)

    self.lbl_stats_title = ctk.CTkLabel(
        self.stats_frame,
        text="توزيع الموظفين حسب المهمة والإطار :",
        font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
    )
    self.lbl_stats_title.pack(anchor="w", padx=12, pady=4)

    self.stats_tree = ttk.Treeview(
        self.stats_frame,
        columns=("FONCTION", "COUNT", "PERCENT"),
        show="headings",
        height=3,
    )
    self.stats_tree.heading("FONCTION", text="المهمة أو الإطار")
    self.stats_tree.heading("COUNT", text="العدد الإجمالي")
    self.stats_tree.heading("PERCENT", text="النسبة")
    self.stats_tree.column("FONCTION", anchor="center", width=420)
    self.stats_tree.column("COUNT", anchor="center", width=140)
    self.stats_tree.column("PERCENT", anchor="center", width=140)
    self.stats_tree.pack(fill="x", padx=10, pady=5)

    # ---------------- شريط التحكم في الصفحات ----------------
    self.pagination_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
    self.pagination_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=2)

    self.btn_prev = ctk.CTkButton(
        self.pagination_frame,
        text="السابق",
        font=ctk.CTkFont(family="Arial", size=12),
        width=70,
        command=self.prev_page,
    )
    self.btn_prev.pack(side="left", padx=5)

    self.lbl_page_info = ctk.CTkLabel(
        self.pagination_frame,
        text="صفحة 1 من 1",
        font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
    )
    self.lbl_page_info.pack(side="left", padx=10)

    self.btn_next = ctk.CTkButton(
        self.pagination_frame,
        text="التالي",
        font=ctk.CTkFont(family="Arial", size=12),
        width=70,
        command=self.next_page,
    )
    self.btn_next.pack(side="left", padx=5)

    self.lbl_page_size = ctk.CTkLabel(
        self.pagination_frame,
        text="عدد السجلات بالصفحة :",
        font=ctk.CTkFont(family="Arial", size=12),
    )
    self.lbl_page_size.pack(side="right", padx=(10, 5))

    self.combo_page_size = ctk.CTkOptionMenu(
        self.pagination_frame,
        values=["50", "100", "200", "الكل"],
        width=80,
        command=self.change_page_size,
        font=ctk.CTkFont(family="Arial", size=12),
    )
    self.combo_page_size.pack(side="right", padx=5)

    # ---------------- جدول اللائحة الاسمية ----------------
    self.table_frame = ctk.CTkFrame(self.main_frame, corner_radius=8)
    self.table_frame.grid(row=4, column=0, sticky="nsew", padx=10, pady=5)
    self.table_frame.grid_columnconfigure(0, weight=1)
    self.table_frame.grid_rowconfigure(0, weight=1)

    self.display_columns = [
        "PPR",
        "CIN",
        "NOM_AR",
        "NOM_FR",
        "المهام",
        "الاطار",
        "الدرجة",
        "C_ETAB",
        "الوحدة المدرسية.1",
        "المؤسسة",
        "DATE_DEBUT_AFF",
        "DATE_ANC_ADM",
    ]

    self.tree = ttk.Treeview(
        self.table_frame,
        columns=self.display_columns,
        show="headings",
        selectmode="browse",
    )

    headers = {
        "PPR": "رقم التأجير",
        "CIN": "رقم البطاقة",
        "NOM_AR": "الاسم بالعربية",
        "NOM_FR": "الاسم بالفرنسية",
        "المهام": "المهمة",
        "الاطار": "الإطار",
        "الدرجة": "الدرجة",
        "C_ETAB": "كود المؤسسة",
        "الوحدة المدرسية.1": "اسم الفرعية",
        "المؤسسة": "المجموعة المدرسية",
        "DATE_DEBUT_AFF": "تاريخ التعيين بالمؤسسة",
        "DATE_ANC_ADM": "تاريخ التوظيف",
    }

    for col, text in headers.items():
      self.tree.heading(col, text=text)
      self.tree.column(col, anchor="center", width=115)

    self.tree.column("NOM_AR", width=130)
    self.tree.column("NOM_FR", width=140)
    self.tree.column("الوحدة المدرسية.1", width=130)
    self.tree.column("المؤسسة", width=150)

    self.scrollbar_y = ttk.Scrollbar(
        self.table_frame, orient="vertical", command=self.tree.yview
    )
    self.scrollbar_x = ttk.Scrollbar(
        self.table_frame, orient="horizontal", command=self.tree.xview
    )
    self.tree.configure(
        yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set
    )

    self.tree.grid(row=0, column=0, sticky="nsew")
    self.scrollbar_y.grid(row=0, column=1, sticky="ns")
    self.scrollbar_x.grid(row=1, column=0, sticky="ew")

    self.setup_table_style()
    self.tree.bind("<Double-1>", self.on_teacher_double_click)

    # شريط تلميح سفلي
    self.bottom_bar = ctk.CTkFrame(
        self.main_frame, height=28, fg_color="transparent"
    )
    self.bottom_bar.grid(row=5, column=0, sticky="ew", padx=10, pady=2)
    self.lbl_hint = ctk.CTkLabel(
        self.bottom_bar,
        text="تلميح : انقر مرتين على أي سطر لتعديل معطيات الموظف أو طباعتها",
        text_color="#7F8C8D",
        font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
    )
    self.lbl_hint.pack(side="left")

    self.auto_load_file()

  def setup_table_style(self):
    style = ttk.Style()
    style.theme_use("clam")

    style.configure(
        "Treeview.Heading",
        font=("Arial", 10, "bold"),
        background="#2980B9",
        foreground="white",
        relief="flat",
        padding=5,
    )
    style.map("Treeview.Heading", background=[("active", "#1F538D")])

    style.configure(
        "Treeview",
        font=("Arial", 10),
        rowheight=26,
        background="#FFFFFF",
        foreground="#1C2833",
        fieldbackground="#FFFFFF",
        relief="flat",
    )
    style.map(
        "Treeview",
        background=[("selected", "#5DADE2")],
        foreground=[("selected", "white")],
    )

    self.tree.tag_configure(
        "evenrow", background="#FFFFFF", foreground="#1A252C"
    )
    self.tree.tag_configure("oddrow", background="#EBF5FB", foreground="#1A252C")

    self.stats_tree.tag_configure(
        "evenrow", background="#FFFFFF", foreground="#1A252C"
    )
    self.stats_tree.tag_configure(
        "oddrow", background="#EBF5FB", foreground="#1A252C"
    )

  def change_appearance(self, mode):
    eng_mode = "Dark" if mode == "داكن" else "Light"
    ctk.set_appearance_mode(eng_mode)
    if eng_mode == "Dark":
      self.tree.tag_configure(
          "evenrow", background="#2C3E50", foreground="#ECF0F1"
      )
      self.tree.tag_configure(
          "oddrow", background="#34495E", foreground="#ECF0F1"
      )
    else:
      self.tree.tag_configure(
          "evenrow", background="#FFFFFF", foreground="#1A252C"
      )
      self.tree.tag_configure(
          "oddrow", background="#EBF5FB", foreground="#1A252C"
      )

  def create_kpi_card(self, parent, col, title, initial_val, color):
    frame = ctk.CTkFrame(parent, fg_color=color, corner_radius=10, height=75)
    frame.grid(row=0, column=col, padx=6, pady=4, sticky="nsew")
    frame.grid_propagate(False)

    lbl_t = ctk.CTkLabel(
        frame, text=title, font=ctk.CTkFont(family="Arial", size=12), text_color="#EAECEE"
    )
    lbl_t.pack(pady=(8, 2))
    lbl_v = ctk.CTkLabel(
        frame,
        text=initial_val,
        font=ctk.CTkFont(family="Arial", size=16, weight="bold"),
        text_color="white",
    )
    lbl_v.pack()
    return lbl_v

  def auto_load_file(self):
    default_name = "RH2026.xlsx"
    if os.path.exists(default_name):
      self.active_file_path = default_name
      self.read_data(default_name)
    else:
      self.lbl_status.configure(
          text="لم يتم العثور على الملف تلقائياً", text_color="orange"
      )

  def load_excel_manually(self):
    path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    if path:
      self.active_file_path = path
      self.read_data(path)

  def read_data(self, path):
    try:
      df = pd.read_excel(path)
      df.columns = [str(c).strip() for c in df.columns]

      for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()

      self.df = df
      self.show_all()
      self.lbl_status.configure(
          text=(
              f"تم الاتصال بقاعدة البيانات\nعدد الموظفين: {len(self.df):,}"
          ),
          text_color="#27AE60",
      )
    except Exception as e:
      messagebox.showerror("خطأ", f"تعذر قراءة ملف الإكسيل :\n{str(e)}")

  def update_views(self, dataframe, title_info="جميع المؤسسات"):
    self.current_filtered_df = dataframe
    total = len(dataframe)

    if self.page_size == -1:
      self.total_pages = 1
    else:
      self.total_pages = max(1, (total + self.page_size - 1) // self.page_size)

    self.current_page = 1
    self.render_current_page(title_info)

  def render_current_page(self, title_info="جميع المؤسسات"):
    self.tree.delete(*self.tree.get_children())
    total = len(self.current_filtered_df)

    if total == 0:
      self.lbl_page_info.configure(text="لا توجد نتائج مطابقة")
      self.stats_tree.delete(*self.stats_tree.get_children())
      self.card_total.configure(text="0")
      self.card_females.configure(text="0 (0%)")
      self.card_males.configure(text="0 (0%)")
      self.card_etab.configure(text=title_info[:24])
      return

    if self.page_size == -1:
      page_df = self.current_filtered_df
      start_idx = 0
      end_idx = total
    else:
      start_idx = (self.current_page - 1) * self.page_size
      end_idx = min(start_idx + self.page_size, total)
      page_df = self.current_filtered_df.iloc[start_idx:end_idx]

    for idx, (_, row) in enumerate(page_df.iterrows()):
      values = [row.get(col, "") for col in self.display_columns]
      tag = "evenrow" if idx % 2 == 0 else "oddrow"
      self.tree.insert("", "end", values=values, tags=(tag,))

    if self.page_size == -1:
      self.lbl_page_info.configure(text=f"عرض الكل : {total:,} موظف")
    else:
      self.lbl_page_info.configure(
          text=(
              f"صفحة {self.current_page} من {self.total_pages} (من {start_idx+1}"
              f" إلى {end_idx} من أصل {total:,})"
          )
      )

    self.card_total.configure(text=f"{total:,}")

    if "SEX" in self.current_filtered_df.columns and total > 0:
      females = (self.current_filtered_df["SEX"].str.upper() == "F").sum()
      males = (self.current_filtered_df["SEX"].str.upper() == "M").sum()
      self.card_females.configure(
          text=f"{females:,} ({females/total*100:.1f}%)"
      )
      self.card_males.configure(text=f"{males:,} ({males/total*100:.1f}%)")
    else:
      self.card_females.configure(text="0")
      self.card_males.configure(text="0")

    self.card_etab.configure(text=title_info[:24])

    self.stats_tree.delete(*self.stats_tree.get_children())
    if (
        not self.current_filtered_df.empty
        and "المهام" in self.current_filtered_df.columns
    ):
      counts = self.current_filtered_df["المهام"].value_counts()
      for idx, (fonc, count) in enumerate(counts.items()):
        if fonc:
          pct = f"{(count / total * 100):.1f}%"
          tag = "evenrow" if idx % 2 == 0 else "oddrow"
          self.stats_tree.insert(
              "", "end", values=(fonc, f"{count:,}", pct), tags=(tag,)
          )
      self.lbl_stats_title.configure(
          text=f"توزيع الموظفين حسب المهمة والإطار ({title_info}) :"
      )

  def next_page(self):
    if self.current_page < self.total_pages:
      self.current_page += 1
      self.render_current_page(self.card_etab.cget("text"))

  def prev_page(self):
    if self.current_page > 1:
      self.current_page -= 1
      self.render_current_page(self.card_etab.cget("text"))

  def change_page_size(self, val):
    if val == "الكل":
      self.page_size = -1
    else:
      self.page_size = int(val)
    self.update_views(self.current_filtered_df, self.card_etab.cget("text"))

  def show_all(self):
    if not self.df.empty:
      self.entry_etab.delete(0, "end")
      self.entry_quick.delete(0, "end")
      self.update_views(self.df, "جميع المؤسسات المسجلة")

  def filter_by_etab(self):
    code = self.entry_etab.get().strip().upper()
    if not code:
      messagebox.showwarning("تنبيه", "يرجى كتابة كود المؤسسة أولاً")
      return

    if self.df.empty:
      return

    mask = (self.df["C_ETAB"].astype(str).str.upper() == code) | (
        self.df["CD_ETAB"].astype(str).str.upper() == code
    )
    res = self.df[mask]

    if not res.empty:
      r0 = res.iloc[0]
      nom_sub = r0.get("الوحدة المدرسية.1", code)
      self.update_views(res, f"{nom_sub} ({code})")
    else:
      messagebox.showinfo(
          "تنبيه", f"لا يوجد أي موظف مسجل بهذا الكود : {code}"
      )

  # حل مشكلة البطء: البحث يتم فوراً كعملية واحدة سريعة جداً عند الضغط على Enter أو الزر فقط
  def search_action(self):
    query = self.entry_quick.get().strip()
    if not query:
      self.show_all()
      return

    if self.df.empty:
      return

    query_lower = query.lower()
    mask = (
        self.df["PPR"].astype(str).str.contains(query, regex=False)
        | self.df["CIN"]
            .astype(str)
            .str.lower()
            .str.contains(query_lower, regex=False)
        | self.df["NOM_AR"].astype(str).str.contains(query, regex=False)
        | self.df["NOM_FR"]
            .astype(str)
            .str.lower()
            .str.contains(query_lower, regex=False)
        | self.df["الوحدة المدرسية.1"]
            .astype(str)
            .str.contains(query, regex=False)
        | self.df["المؤسسة"].astype(str).str.contains(query, regex=False)
    )
    res = self.df[mask]
    self.update_views(res, f"نتيجة البحث عن : {query}")

  def on_teacher_double_click(self, event):
    selected_item = self.tree.selection()
    if not selected_item:
      return

    item_values = self.tree.item(selected_item[0], "values")
    if not item_values:
      return

    ppr = str(item_values[0]).strip()
    teacher_rows = self.df[self.df["PPR"].astype(str) == ppr]

    if not teacher_rows.empty:
      teacher_dict = teacher_rows.iloc[0].to_dict()
      EditTeacherWindow(
          self, ppr, teacher_dict, self.apply_teacher_update, is_new=False
      )

  def open_add_dialog(self):
    if self.df.empty:
      messagebox.showwarning("تنبيه", "يرجى فتح قاعدة البيانات أولاً")
      return
    empty_data = {col: "" for col in self.df.columns}
    EditTeacherWindow(
        self, "", empty_data, self.apply_teacher_update, is_new=True
    )

  def apply_teacher_update(self, ppr, updated_dict, is_new=False):
    if is_new:
      new_row_df = pd.DataFrame([updated_dict])
      self.df = pd.concat([self.df, new_row_df], ignore_index=True)
    else:
      mask = self.df["PPR"].astype(str) == str(ppr)
      for col_name, new_val in updated_dict.items():
        if col_name in self.df.columns:
          self.df.loc[mask, col_name] = str(new_val)

    cur_code = self.entry_etab.get().strip().upper()
    if cur_code:
      self.filter_by_etab()
    else:
      self.search_action()

  def save_to_source_excel(self):
    if self.df.empty:
      messagebox.showwarning("تنبيه", "لا توجد معطيات للحفظ")
      return

    confirm = messagebox.askyesno(
        "تأكيد الحفظ",
        "هل تريد حفظ التعديلات مباشرة في ملف العمل الأصلي ؟",
    )
    if confirm:
      try:
        with pd.ExcelWriter(self.active_file_path, engine="openpyxl") as writer:
          self.df.to_excel(writer, index=False, sheet_name="PERSONNEL")
        messagebox.showinfo(
            "تأكيد", "تم حفظ كافة التعديلات في ملف Excel بنجاح"
        )
      except Exception as e:
        messagebox.showerror(
            "خطأ",
            "تعذر حفظ الملف، يرجى إغلاقه أولاً إن كان مفتوحاً في Excel"
            f" :\n{str(e)}",
        )

  def export_excel(self):
    if self.current_filtered_df.empty:
      messagebox.showwarning("تنبيه", "لا توجد معطيات للتصدير")
      return
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
    )
    if save_path:
      try:
        with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
          self.current_filtered_df.to_excel(
              writer, index=False, sheet_name="اللائحة الاسمية"
          )
          if "المهام" in self.current_filtered_df.columns:
            stats = (
                self.current_filtered_df["المهام"]
                .value_counts()
                .reset_index()
            )
            stats.columns = ["المهمة أو الإطار", "العدد الإجمالي"]
            stats.to_excel(writer, index=False, sheet_name="إحصاء الأطر")
        messagebox.showinfo("تأكيد", "تم تصدير اللائحة والإحصاء بنجاح")
      except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ أثناء التصدير :\n{str(e)}")


if __name__ == "__main__":
  app = RH2026App()
  app.mainloop()
