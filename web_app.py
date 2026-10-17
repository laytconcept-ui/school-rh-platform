import glob
import io
import os
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="بوابة تدبير الموارد البشرية | RH Portal",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# تخصيص المظهر وتنسيق الواجهة وإخفاء إشعار "Press Enter to apply"
st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&family=Tajawal:wght@400;500;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Cairo', 'Tajawal', sans-serif !important;
        direction: rtl;
        text-align: right;
    }

    div[data-testid="InputInstructions"] {
        display: none !important;
    }
    small[data-testid="stFormHelp"] {
        display: none !important;
    }

    .stApp {
        background-color: #F8FAFC !important;
    }

    .portal-header {
        background: linear-gradient(135deg, #1B365D 0%, #2A5298 100%);
        padding: 22px 28px;
        border-radius: 14px;
        color: white;
        margin-bottom: 22px;
        box-shadow: 0 4px 15px rgba(27, 54, 93, 0.15);
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .portal-title {
        font-size: 23px;
        font-weight: 800;
        margin: 0;
        color: #FFFFFF;
    }
    .portal-sub {
        font-size: 13px;
        color: #D6E4FF;
        margin-top: 4px;
    }
    .portal-badge {
        background: rgba(255, 255, 255, 0.18);
        border: 1px solid rgba(255, 255, 255, 0.3);
        padding: 6px 16px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: 600;
    }

    .kpi-container {
        display: flex;
        gap: 12px;
        margin-bottom: 18px;
        flex-wrap: wrap;
    }
    .kpi-box {
        flex: 1;
        min-width: 180px;
        background: #FFFFFF;
        padding: 16px 18px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        border-right: 5px solid #2563EB;
        border-top: 1px solid #E2E8F0;
        border-left: 1px solid #E2E8F0;
        border-bottom: 1px solid #E2E8F0;
    }
    .kpi-label {
        color: #64748B;
        font-size: 12px;
        font-weight: 600;
        margin-bottom: 5px;
    }
    .kpi-num {
        color: #0F172A;
        font-size: 20px;
        font-weight: 800;
    }

    .stTextInput input, .stSelectbox select {
        border-radius: 8px !important;
        border: 1.5px solid #CBD5E1 !important;
        padding: 9px 12px !important;
        font-size: 14px !important;
        background-color: #FFFFFF !important;
        color: #1E293B !important;
    }
    .stTextInput input:focus, .stSelectbox select:focus {
        border-color: #2563EB !important;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12) !important;
    }

    .stButton button {
        border-radius: 8px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        padding: 8px 22px !important;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 2px solid #E2E8F0;
        padding-bottom: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0 !important;
        padding: 10px 22px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        color: #64748B !important;
        background-color: transparent !important;
    }
    .stTabs [aria-selected="true"] {
        color: #2563EB !important;
        border-bottom: 3px solid #2563EB !important;
        background-color: #EFF6FF !important;
    }
</style>
""",
    unsafe_allow_html=True,
)


def clean_value(val):
  if pd.isna(val) or val is None:
    return ""
  s = str(val).strip()
  if s.endswith(".0"):
    s = s[:-2]
  return s


def find_excel_file():
  possible_names = [
      "RH2026.xlsx",
      "rh2026.xlsx",
      "Rh2026.xlsx",
      "RH2026.XLSX",
      "RH.xlsx",
      "rh.xlsx",
  ]
  for name in possible_names:
    if os.path.exists(name):
      return name
  rh_files = glob.glob("*RH*.xlsx") + glob.glob("*rh*.xlsx")
  if rh_files:
    return rh_files[0]
  return None


@st.cache_data
def load_data(source):
  try:
    df = pd.read_excel(source)
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.columns:
      df[col] = df[col].apply(clean_value)
    return df
  except Exception as e:
    st.error(f"حدث خطأ أثناء قراءة ملف الإكسيل: {e}")
    return pd.DataFrame()


detected_file = find_excel_file()
if detected_file:
  df_source = load_data(detected_file)
else:
  st.sidebar.warning("لم يتم العثور على ملف RH2026.xlsx في المستودع.")
  uploaded = st.sidebar.file_uploader(
      "يرجى رفع ملف RH2026.xlsx يدوياً:", type=["xlsx", "xls"]
  )
  if uploaded:
    df_source = load_data(uploaded)
  else:
    df_source = pd.DataFrame()

if df_source.empty:
  st.info("👈 يرجى رفع ملف قاعدة البيانات لبدء التشغيل.")
  st.stop()

cols_all = list(df_source.columns)


def detect_col(candidates, fallback):
  for cand in candidates:
    for c in cols_all:
      if c.strip().lower() == cand.strip().lower():
        return c
  return fallback


COL_C_ETAB = detect_col(["C_ETAB", "كود المؤسسة"], cols_all[0])
COL_UNIT_CODE = detect_col(
    ["c_unité", "c_unite", "c-unite", "الوحدة المدرسية", "كود الوحدة"],
    COL_C_ETAB,
)
COL_UNIT_NAME = detect_col(
    ["الوحدة المدرسية.1", "اسم الفرعية", "اسم الوحدة"], COL_UNIT_CODE
)
COL_INSTITUTION = detect_col(
    ["المؤسسة", "اسم المجموعة", "المجموعة المدرسية"], COL_UNIT_NAME
)
COL_TASK = detect_col(
    ["المهام", "FONCTION", "fonct", "المهمة"],
    cols_all[1] if len(cols_all) > 1 else cols_all[0],
)
COL_PPR = detect_col(["PPR", "رقم التأجير", "تأجير"], cols_all[0])

if "df" not in st.session_state or len(st.session_state.df) != len(df_source):
  st.session_state.df = df_source.copy()

df = st.session_state.df


# ==================== نافذة التعديل المنبثقة الشاملة (Modal Dialog) ====================
if hasattr(st, "dialog"):
  dialog_decorator = st.dialog("✏️ بطاقة تعديل المعطيات الشاملة للأستاذ")
else:

  def dialog_decorator(func):

    def wrapper(*args, **kwargs):
      with st.expander("✏️ بطاقة تعديل المعطيات الشاملة للأستاذ", expanded=True):
        return func(*args, **kwargs)

    return wrapper


@dialog_decorator
def show_teacher_edit_dialog(ppr_selected):
  match_indices = df.index[df[COL_PPR] == str(ppr_selected)].tolist()
  if not match_indices:
    st.error("تعذر العثور على سجل هذا الأستاذ!")
    return

  idx = match_indices[0]
  t = df.loc[idx]

  st.markdown(
      f"### 👤 {t.get('NOM_AR', '')} | {t.get('NOM_FR', '')} (PPR:"
      f" `{ppr_selected}`)"
  )
  st.caption(
      "يمكنك تعديل أي معلومة ثم الضغط على زر الحفظ لتحديث قاعدة البيانات"
      " فوراً."
  )

  with st.form("modal_edit_teacher_form"):
    st.markdown("##### 📌 المعطيات الشخصية والمهنية")
    c1, c2, c3 = st.columns(3)
    v_nom_ar = c1.text_input("الاسم بالعربية:", t.get("NOM_AR", ""))
    v_nom_fr = c2.text_input("الاسم بالفرنسية:", t.get("NOM_FR", ""))
    v_cin = c3.text_input("رقم البطاقة (CIN):", t.get("CIN", ""))

    v_sex = c1.text_input("الجنس (M/F):", t.get("SEX", ""))
    v_sit = c2.text_input("الحالة العائلية:", t.get("SIT_FAM", ""))
    v_nais = c3.text_input("تاريخ الازدياد:", t.get("DATE_NAISSANCE", ""))

    st.markdown("##### 🏫 معطيات التعيين والوحدة المدرسية")
    c4, c5, c6 = st.columns(3)
    v_c_etab = c4.text_input("كود المؤسسة (C_ETAB):", t.get(COL_C_ETAB, ""))
    v_u_code = c5.text_input("كود الوحدة المدرسية:", t.get(COL_UNIT_CODE, ""))
    v_u_name = c6.text_input(
        "اسم الوحدة المدرسية (الفرعية):", t.get(COL_UNIT_NAME, "")
    )

    v_inst = c4.text_input(
        "المجموعة المدرسية المركزية:", t.get(COL_INSTITUTION, "")
    )
    v_task = c5.text_input("المهمة:", t.get(COL_TASK, ""))
    v_cadre = c6.text_input("الإطار:", t.get("الاطار", ""))

    st.markdown("##### 📜 الترقية والتواريخ الإدارية")
    c7, c8, c9 = st.columns(3)
    v_grade = c7.text_input("الدرجة:", t.get("الدرجة", ""))
    v_d_aff = c8.text_input(
        "تاريخ التعيين بالمؤسسة:", t.get("DATE_DEBUT_AFF", "")
    )
    v_d_rec = c9.text_input("تاريخ التوظيف:", t.get("DATE_ANC_ADM", ""))

    v_d_gd = c7.text_input("تاريخ مفعول الدرجة:", t.get("DATE_EFFET_GD", ""))
    v_d_elo = c8.text_input("تاريخ مفعول الرتبة:", t.get("DATE_EFFET_ELO", ""))
    v_cd_etab = c9.text_input("كود التعيين (CD_ETAB):", t.get("CD_ETAB", ""))

    submitted = st.form_submit_button(
        "💾 حفظ كافة التعديلات في قاعدة البيانات",
        type="primary",
        use_container_width=True,
    )
    if submitted:
      updates = {
          "NOM_AR": v_nom_ar,
          "NOM_FR": v_nom_fr,
          "CIN": v_cin,
          "SEX": v_sex,
          "SIT_FAM": v_sit,
          "DATE_NAISSANCE": v_nais,
          COL_C_ETAB: v_c_etab,
          COL_UNIT_CODE: v_u_code,
          COL_UNIT_NAME: v_u_name,
          COL_INSTITUTION: v_inst,
          COL_TASK: v_task,
          "الاطار": v_cadre,
          "الدرجة": v_grade,
          "DATE_DEBUT_AFF": v_d_aff,
          "DATE_ANC_ADM": v_d_rec,
          "DATE_EFFET_GD": v_d_gd,
          "DATE_EFFET_ELO": v_d_elo,
          "CD_ETAB": v_cd_etab,
      }
      for col_name, val in updates.items():
        if col_name in df.columns:
          df.loc[idx, col_name] = clean_value(val)

      st.session_state.df = df
      st.success("✓ تم حفظ كافة التعديلات وتحديث السجل بنجاح!")
      st.rerun()


# ==================== القائمة الجانبية ====================
with st.sidebar:
  st.markdown(
      """
        <div style="text-align: center; padding: 10px 0 15px 0;">
            <div style="font-size: 38px; margin-bottom: 5px;">🏛️</div>
            <h3 style="margin:0; color:#1E293B; font-weight:800;">بوابة الموارد البشرية</h3>
            <span style="color:#64748B; font-size:12px;">مصلحة الموارد البشرية والتخطيط</span>
        </div>
    """,
      unsafe_allow_html=True,
  )

  st.markdown("---")
  st.metric("إجمالي الموظفين المسجلين", f"{len(df):,} فرد")

  buf_all = io.BytesIO()
  with pd.ExcelWriter(buf_all, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="PERSONNEL")
  st.download_button(
      label="📥 تحميل قاعدة المعطيات كاملة (Excel)",
      data=buf_all.getvalue(),
      file_name="RH2026_Updated.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      use_container_width=True,
  )

  st.markdown("<br>", unsafe_allow_html=True)
  if st.button("🔄 إعادة ضبط وتحديث القاعدة", use_container_width=True):
    if detected_file:
      st.session_state.df = load_data(detected_file)
      st.success("تم تحديث البيانات بنجاح!")
      st.rerun()

# ==================== الترويسة ====================
st.markdown(
    """
    <div class="portal-header">
        <div>
            <h1 class="portal-title">منظومة تدبير الموارد البشرية وتتبع المؤسسات والفرعيات</h1>
            <div class="portal-sub">فلترة مباشرة بكود المؤسسة (C_ETAB) واستعلام تفاعلي لبطاقة الأستاذ - الموسم الدراسي 2026</div>
        </div>
        <div class="portal-badge">نسخة الويب التفاعلية v5.0</div>
    </div>
""",
    unsafe_allow_html=True,
)

# ==================== التبويبات ====================
tab_etab, tab_search, tab_stats, tab_edit = st.tabs([
    "🏫 استعلام المؤسسة والوحدات التابعة",
    "🔍 التنقيب والبحث الفردي",
    "📊 لوحة المؤشرات والإحصاء",
    "✏️ تدبير وتعديل سجل موظف",
])

# -------------------------------------------------------------
# 1. تبويب استعلام المؤسسة بكود C_ETAB
# -------------------------------------------------------------
with tab_etab:
  with st.form("etab_filter_form"):
    st.markdown("##### 📍 تحديد المؤسسة التعليمية بكود المؤسسة (C_ETAB)")
    col_c1, col_c2, col_btn = st.columns([1.5, 2.5, 1])

    etab_input = (
        col_c1.text_input(
            "أدخل كود المؤسسة (C_ETAB):", placeholder="مثال: 03254L"
        )
        .strip()
        .upper()
    )

    all_etabs = [""]
    if COL_C_ETAB in df.columns:
      all_etabs += sorted(
          list(df[COL_C_ETAB].replace("", pd.NA).dropna().unique())
      )

    etab_dropdown = col_c2.selectbox("أو اختر كود C_ETAB من اللائحة:", all_etabs)

    col_btn.write("")
    col_btn.write("")
    btn_apply = col_btn.form_submit_button(
        "🔎 تأكيد الاستعلام", use_container_width=True, type="primary"
    )

  target_code = etab_input if etab_input else etab_dropdown

  if target_code:
    res_etab = df[
        df[COL_C_ETAB].astype(str).str.strip().str.upper() == target_code
    ]

    if not res_etab.empty:
      r0 = res_etab.iloc[0]
      institution_name = r0.get(COL_INSTITUTION, "")

      total_staff = len(res_etab)
      females_cnt = (
          (res_etab["SEX"].str.upper() == "F").sum()
          if "SEX" in res_etab.columns
          else 0
      )
      males_cnt = (
          (res_etab["SEX"].str.upper() == "M").sum()
          if "SEX" in res_etab.columns
          else 0
      )

      st.markdown(
          f"""
                <div class="kpi-container">
                    <div class="kpi-box" style="border-right-color:#2563EB;">
                        <div class="kpi-label">المجموعة المدرسية / المؤسسة</div>
                        <div class="kpi-num" style="font-size:18px;">{institution_name if institution_name else 'مؤسسة مستقلة'}</div>
                    </div>
                    <div class="kpi-box" style="border-right-color:#0D9488;">
                        <div class="kpi-label">كود المؤسسة المطابق (C_ETAB)</div>
                        <div class="kpi-num" style="font-size:18px;">{target_code}</div>
                    </div>
                    <div class="kpi-box" style="border-right-color:#8B5CF6;">
                        <div class="kpi-label">مجموع العاملين بالمؤسسة</div>
                        <div class="kpi-num">{total_staff} أطر</div>
                    </div>
                    <div class="kpi-box" style="border-right-color:#F59E0B;">
                        <div class="kpi-label">توزيع الجنسين</div>
                        <div class="kpi-num" style="font-size:16px;">{females_cnt} إناث ({females_cnt/total_staff*100:.0f}%) | {males_cnt} ذكور</div>
                    </div>
                </div>
            """,
          unsafe_allow_html=True,
      )

      # جدول الوحدات المدرسية التابعة
      st.markdown(
          f"#### 🏫 الوحدات المدرسية التابعة للمؤسسة `{target_code}`:"
      )

      raw_grp_cols = [
          c for c in [COL_UNIT_CODE, COL_UNIT_NAME] if c in res_etab.columns
      ]
      clean_grp = list(dict.fromkeys(raw_grp_cols))

      if clean_grp:
        sub_units_summary = (
            res_etab.groupby(clean_grp).size().reset_index(name="عدد العاملين")
        )
        rename_map = {}
        if COL_UNIT_CODE in sub_units_summary.columns:
          rename_map[COL_UNIT_CODE] = "كود الوحدة المدرسية (c_unité)"
        if (
            COL_UNIT_NAME in sub_units_summary.columns
            and COL_UNIT_NAME != COL_UNIT_CODE
        ):
          rename_map[COL_UNIT_NAME] = "اسم الوحدة المدرسية"
        sub_units_summary.rename(columns=rename_map, inplace=True)

        if "اسم الوحدة المدرسية" not in sub_units_summary.columns:
          sub_units_summary["اسم الوحدة المدرسية"] = sub_units_summary.get(
              "كود الوحدة المدرسية (c_unité)", ""
          )
        sub_units_summary["النسبة"] = (
            sub_units_summary["عدد العاملين"] / total_staff * 100
        ).map("{:.1f}%".format)
      else:
        sub_units_summary = pd.DataFrame()

      col_units, col_filter_unit = st.columns([2.5, 1.5])
      with col_units:
        if not sub_units_summary.empty:
          st.dataframe(
              sub_units_summary, use_container_width=True, hide_index=True
          )

      with col_filter_unit:
        st.markdown("###### تصفية العرض حسب وحدة مدرسية معينة:")
        unit_opts = ["عرض جميع الوحدات المدرسية التابعة"]
        if (
            not sub_units_summary.empty
            and "كود الوحدة المدرسية (c_unité)" in sub_units_summary.columns
        ):
          for _, r in sub_units_summary.iterrows():
            u_code = r.get("كود الوحدة المدرسية (c_unité)", "")
            u_name = r.get("اسم الوحدة المدرسية", "")
            unit_opts.append(f"{u_code} - {u_name}")

        selected_unit_choice = st.selectbox(
            "اختر الوحدة لعرض أطرها فقط:",
            unit_opts,
            key="unit_select_choice_etab",
        )

      if selected_unit_choice != "عرض جميع الوحدات المدرسية التابعة":
        unit_code_selected = (
            selected_unit_choice.split(" - ")[0].strip().upper()
        )
        final_display_df = res_etab[
            res_etab[COL_UNIT_CODE].astype(str).str.strip().str.upper()
            == unit_code_selected
        ]
        st.info(
            f"عرض أطر الوحدة المدرسية: **{selected_unit_choice}** (العدد:"
            f" {len(final_display_df)} أطر)"
        )
      else:
        final_display_df = res_etab

      st.markdown("---")
      stats_tasks = (
          final_display_df[COL_TASK].value_counts().reset_index()
      )
      stats_tasks.columns = ["المهمة", "العدد"]
      if len(final_display_df) > 0:
        stats_tasks["النسبة"] = (
            stats_tasks["العدد"] / len(final_display_df) * 100
        ).map("{:.1f}%".format)

      # ==================== اللائحة الاسمية للعاملين بالمؤسسة مع ميزة النقر والتعديل ====================
      st.markdown("#### 📋 اللائحة الاسمية للعاملين بالمؤسسة:")
      st.caption(
          "💡 **انقر على أي سطر في الجدول** لفتح واجهة منفردة (نافذة تفاعلية)"
          " تمكنك من معاينة وتعديل معطيات الأستاذ وحفظها مباشرة."
      )

      # إظهار عمود الوحدة المدرسية وكودها بوضوح في الواجهة
      disp_cols = [
          COL_PPR,
          "CIN",
          "NOM_AR",
          "NOM_FR",
          COL_TASK,
          COL_UNIT_NAME,
          COL_UNIT_CODE,
          COL_C_ETAB,
          "الدرجة",
          "DATE_DEBUT_AFF",
      ]
      clean_disp = list(dict.fromkeys(disp_cols))
      avail_cols = [c for c in clean_disp if c in final_display_df.columns]

      # إعداد جدول تفاعلي يدعم تحديد السطر فوراً
      selection = st.dataframe(
          final_display_df[avail_cols],
          use_container_width=True,
          height=300,
          on_select="rerun",
          selection_mode="single-row",
      )

      # فتح واجهة التعديل المستقلة عند اختيار أي سطر
      if (
          selection
          and selection.get("selection")
          and selection["selection"].get("rows")
      ):
        selected_row_idx = selection["selection"]["rows"][0]
        clicked_ppr = str(
            final_display_df.iloc[selected_row_idx][COL_PPR]
        ).strip()
        show_teacher_edit_dialog(clicked_ppr)

      # زر تصدير التقرير
      buf_e = io.BytesIO()
      with pd.ExcelWriter(buf_e, engine="openpyxl") as writer:
        final_display_df.to_excel(
            writer, index=False, sheet_name="اللائحة الاسمية"
        )
        if not sub_units_summary.empty:
          sub_units_summary.to_excel(
              writer, index=False, sheet_name="الوحدات المدرسية"
          )
        stats_tasks.to_excel(writer, index=False, sheet_name="إحصاء الأطر")
      st.download_button(
          label=f"📥 تصدير تقرير المؤسسة {target_code} (Excel)",
          data=buf_e.getvalue(),
          file_name=f"RH_{target_code}.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
    else:
      st.warning(
          "لم يتم العثور على أي موظف مسجل بكود المؤسسة C_ETAB:"
          f" `{target_code}`"
      )
  else:
    st.info(
        "👈 المرجو إدخال كود المؤسسة (C_ETAB) ثم الضغط على **تأكيد الاستعلام**."
    )

# -------------------------------------------------------------
# 2. تبويب التنقيب الفردي
# -------------------------------------------------------------
with tab_search:
  with st.form("search_form"):
    st.markdown("##### 🔎 استعلام دقيق بالرقم المالي، الاسم أو البطاقة الوطنية")
    q_col1, q_col2 = st.columns([3, 1])
    query_val = q_col1.text_input(
        "أدخل رقم التأجير (PPR)، الاسم، أو رقم البطاقة (CIN):",
        placeholder="مثال: 1114650 أو محمد",
    ).strip()
    q_col2.write("")
    q_col2.write("")
    btn_query = q_col2.form_submit_button(
        "🔍 تنفيذ الاستعلام", use_container_width=True, type="primary"
    )

  if query_val:
    q_clean = clean_value(query_val).lower()
    mask_ppr = df[COL_PPR].astype(str).str.strip() == q_clean
    mask_cin = (
        df["CIN"].astype(str).str.strip().str.lower() == q_clean
        if "CIN" in df.columns
        else False
    )
    mask_names = df["NOM_AR"].astype(str).str.contains(
        query_val, regex=False
    ) | df["NOM_FR"].astype(str).str.lower().str.contains(
        q_clean, regex=False
    )
    res_q = df[mask_ppr | mask_cin | mask_names]

    if not res_q.empty:
      st.success(
          f"تم العثور على **{len(res_q)}** نتيجة مطابقة (انقر على السطر"
          " لتعديله):"
      )
      raw_disp_cols = [
          COL_PPR,
          "CIN",
          "NOM_AR",
          "NOM_FR",
          COL_TASK,
          "الاطار",
          "الدرجة",
          COL_UNIT_NAME,
          COL_UNIT_CODE,
          COL_C_ETAB,
          "DATE_DEBUT_AFF",
      ]
      disp_cols = list(dict.fromkeys(raw_disp_cols))
      cols_to_render = [c for c in disp_cols if c in res_q.columns]

      sel_q = st.dataframe(
          res_q[cols_to_render],
          use_container_width=True,
          on_select="rerun",
          selection_mode="single-row",
      )

      if sel_q and sel_q.get("selection") and sel_q["selection"].get("rows"):
        r_idx = sel_q["selection"]["rows"][0]
        clicked_q_ppr = str(res_q.iloc[r_idx][COL_PPR]).strip()
        show_teacher_edit_dialog(clicked_q_ppr)

      if len(res_q) == 1:
        indiv = res_q.iloc[0]
        with st.container():
          st.markdown("---")
          st.markdown(
              "#### 👤 بطاقة المسار الإداري:"
              f" {indiv.get('NOM_AR', '')} ({indiv.get('NOM_FR', '')})"
          )
          c_b1, c_b2, c_b3 = st.columns(3)
          c_b1.markdown(f"**رقم التأجير (PPR):** `{indiv.get(COL_PPR, '')}`")
          c_b1.markdown(f"**رقم البطاقة (CIN):** `{indiv.get('CIN', '')}`")
          c_b1.markdown(f"**تاريخ الازدياد:** {indiv.get('DATE_NAISSANCE', '')}")

          c_b2.markdown(f"**المهمة:** {indiv.get(COL_TASK, '')}")
          c_b2.markdown(
              f"**الإطار والدرجة:** {indiv.get('الاطار', '')} -"
              f" {indiv.get('الدرجة', '')}"
          )
          c_b2.markdown(
              f"**تاريخ التوظيف:** {indiv.get('DATE_ANC_ADM', '')}"
          )

          c_b3.markdown(
              "**المؤسسة (المجموعة):**"
              f" {indiv.get(COL_INSTITUTION, '')} (`{indiv.get(COL_C_ETAB, '')}`)"
          )
          c_b3.markdown(
              "**الوحدة المدرسية (الفرعية):**"
              f" {indiv.get(COL_UNIT_NAME, '')} (`{indiv.get(COL_UNIT_CODE, '')}`)"
          )
          c_b3.markdown(
              "**تاريخ التعيين بالمؤسسة:**"
              f" {indiv.get('DATE_DEBUT_AFF', '')}"
          )
    else:
      st.warning("لم يتم العثور على أي موظف بهذه المعطيات.")

# -------------------------------------------------------------
# 3. تبويب لوحة الإحصاء
# -------------------------------------------------------------
with tab_stats:
  st.markdown("##### 📈 مؤشرات التوزيع الإجمالي للموارد البشرية بالإقليم")

  total_all = len(df)
  females_all = (
      (df["SEX"].str.upper() == "F").sum() if "SEX" in df.columns else 0
  )
  males_all = (df["SEX"].str.upper() == "M").sum() if "SEX" in df.columns else 0
  main_etab_count = df[COL_C_ETAB].nunique() if COL_C_ETAB in df.columns else 0
  sub_unit_count = (
      df[COL_UNIT_CODE].nunique() if COL_UNIT_CODE in df.columns else 0
  )

  st.markdown(
      f"""
        <div class="kpi-container">
            <div class="kpi-box" style="border-right-color:#2563EB;">
                <div class="kpi-label">إجمالي الموظفين</div>
                <div class="kpi-num">{total_all:,} موظف</div>
            </div>
            <div class="kpi-box" style="border-right-color:#8B5CF6;">
                <div class="kpi-label">عدد الإناث</div>
                <div class="kpi-num">{females_all:,} ({females_all/total_all*100:.1f}%)</div>
            </div>
            <div class="kpi-box" style="border-right-color:#0D9488;">
                <div class="kpi-label">عدد الذكور</div>
                <div class="kpi-num">{males_all:,} ({males_all/total_all*100:.1f}%)</div>
            </div>
            <div class="kpi-box" style="border-right-color:#EA580C;">
                <div class="kpi-label">المؤسسات / الوحدات المدرسية</div>
                <div class="kpi-num" style="font-size:16px;">{main_etab_count} مؤسسة | {sub_unit_count} وحدة مدرسية</div>
            </div>
        </div>
    """,
      unsafe_allow_html=True,
  )

  col_g1, col_g2 = st.columns(2)
  with col_g1:
    st.markdown("###### 📊 أعلى 10 مهام من حيث عدد الموظفين:")
    if COL_TASK in df.columns:
      top_tasks_counts = df[COL_TASK].value_counts().head(10)
      top_tasks = pd.DataFrame({
          "المهمة / الإطار": top_tasks_counts.index,
          "العدد الإجمالي": top_tasks_counts.values,
      })
      st.dataframe(top_tasks, use_container_width=True, hide_index=True)

  with col_g2:
    st.markdown("###### 🏫 أكبر 10 مجموعات مدرسية من حيث عدد الأطر:")
    if COL_INSTITUTION in df.columns:
      top_etabs_counts = df[COL_INSTITUTION].value_counts().head(10)
      top_etabs = pd.DataFrame({
          "المؤسسة التعليمية": top_etabs_counts.index,
          "عدد الموظفين": top_etabs_counts.values,
      })
      st.dataframe(top_etabs, use_container_width=True, hide_index=True)

# -------------------------------------------------------------
# 4. تبويب تعديل سجل موظف بالبحث المباشر
# -------------------------------------------------------------
with tab_edit:
  st.markdown("##### ✏️ نموذج تعديل أو تحيين المعطيات الإدارية لموظف")
  col_p1, col_p2 = st.columns([2, 1])
  target_ppr = clean_value(
      col_p1.text_input(
          "أدخل رقم التأجير (PPR) لتحميل استمارة التعديل:",
          placeholder="مثال: 1114650",
      ).strip()
  )

  if target_ppr and COL_PPR in df.columns:
    match_idx = df.index[
        df[COL_PPR].astype(str).str.strip() == target_ppr
    ].tolist()
    if match_idx:
      t_row = df.loc[match_idx[0]]
      st.success(
          f"تم تحميل سجل: **{t_row.get('NOM_AR', '')}**"
          f" ({t_row.get('NOM_FR', '')})"
      )

      with st.form("edit_employee_form_direct"):
        f1, f2, f3 = st.columns(3)
        u_nom_ar = f1.text_input(
            "الاسم الكامل بالعربية:", t_row.get("NOM_AR", "")
        )
        u_nom_fr = f2.text_input("الاسم بالفرنسية:", t_row.get("NOM_FR", ""))
        u_cin = f3.text_input(
            "رقم البطاقة الوطنية (CIN):", t_row.get("CIN", "")
        )

        u_task = f1.text_input("المهمة الحالية:", t_row.get(COL_TASK, ""))
        u_cadre = f2.text_input("الإطار الإداري:", t_row.get("الاطار", ""))
        u_grade = f3.text_input("الدرجة:", t_row.get("الدرجة", ""))

        u_c_etab = f1.text_input(
            "كود المؤسسة (C_ETAB):", t_row.get(COL_C_ETAB, "")
        )
        u_unit_code = f2.text_input(
            "كود الوحدة المدرسية:", t_row.get(COL_UNIT_CODE, "")
        )
        u_unit_name = f3.text_input(
            "اسم الوحدة المدرسية:", t_row.get(COL_UNIT_NAME, "")
        )

        btn_commit = st.form_submit_button(
            "💾 حفظ وتثبيت التعديلات في الجلسة", type="primary"
        )
        if btn_commit:
          if "NOM_AR" in df.columns:
            df.loc[match_idx[0], "NOM_AR"] = u_nom_ar
          if "NOM_FR" in df.columns:
            df.loc[match_idx[0], "NOM_FR"] = u_nom_fr
          if "CIN" in df.columns:
            df.loc[match_idx[0], "CIN"] = u_cin
          if COL_TASK in df.columns:
            df.loc[match_idx[0], COL_TASK] = u_task
          if "الاطار" in df.columns:
            df.loc[match_idx[0], "الاطار"] = u_cadre
          if "الدرجة" in df.columns:
            df.loc[match_idx[0], "الدرجة"] = u_grade
          if COL_C_ETAB in df.columns:
            df.loc[match_idx[0], COL_C_ETAB] = u_c_etab
          if COL_UNIT_CODE in df.columns:
            df.loc[match_idx[0], COL_UNIT_CODE] = u_unit_code
          if COL_UNIT_NAME in df.columns:
            df.loc[match_idx[0], COL_UNIT_NAME] = u_unit_name
          st.session_state.df = df
          st.success(
              "✓ تم حفظ التعديلات بنجاح! يمكنك الآن تحميل الملف المحين من"
              " القائمة الجانبية."
          )
          st.rerun()
    else:
      st.warning("رقم التأجير غير موجود بقاعدة البيانات.")
