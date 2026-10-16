import glob
import io
import os
import openpyxl
import pandas as pd
import streamlit as st

# ضبط إعدادات الصفحة
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

    /* إخفاء جملة Press Enter to apply */
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


# ==================== دوال قراءة وتنظيف البيانات ====================
def find_excel_file():
  possible_names = ["RH2026.xlsx", "rh2026.xlsx", "Rh2026.xlsx", "RH2026.XLSX"]
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
      df[col] = df[col].fillna("").astype(str).str.strip()
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

if "df" not in st.session_state or len(st.session_state.df) != len(df_source):
  st.session_state.df = df_source.copy()

df = st.session_state.df

# ==================== الشريط الجانبي ====================
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
            <div class="portal-sub">استعلام دقيق للوحدات المدرسية والمجموعات المركزية - الموسم الدراسي 2026</div>
        </div>
        <div class="portal-badge">نسخة الويب السحابية v4.0</div>
    </div>
""",
    unsafe_allow_html=True,
)

# ==================== التبويبات ====================
tab_etab, tab_search, tab_stats, tab_edit = st.tabs([
    "🏫 استعلام المؤسسة والوحدات المدرسية التابعة",
    "🔍 التنقيب والبحث الفردي",
    "📊 لوحة المؤشرات والإحصاء",
    "✏️ تدبير وتعديل سجل موظف",
])

# -------------------------------------------------------------
# 1. تبويب استعلام المؤسسة والوحدات التابعة لها
# -------------------------------------------------------------
with tab_etab:
  with st.form("etab_filter_form"):
    st.markdown("##### 📍 تحديد المؤسسة التعليمية أو كود الفرعية (C_ETAB)")
    col_c1, col_c2, col_btn = st.columns([1.5, 2.5, 1])

    etab_input = (
        col_c1.text_input(
            "أدخل كود المؤسسة أو الفرعية (C_ETAB):",
            placeholder="مثال: 03242Y أو 03246C",
        )
        .strip()
        .upper()
    )

    all_etabs = [""] + sorted(
        list(df["C_ETAB"].replace("", pd.NA).dropna().unique())
    )
    etab_dropdown = col_c2.selectbox("أو اختر كود المؤسسة من اللائحة:", all_etabs)

    col_btn.write("")
    col_btn.write("")
    btn_apply = col_btn.form_submit_button(
        "🔎 تأكيد الاستعلام", use_container_width=True, type="primary"
    )

  target_code = etab_input if etab_input else etab_dropdown

  if target_code:
    # البحث عن الكود سواء كان كود الفرعية C_ETAB أو كود المركزية في عمود 'الوحدة المدرسية'
    found_rows = df[
        (df["C_ETAB"].astype(str).str.upper() == target_code)
        | (df["الوحدة المدرسية"].astype(str).str.upper() == target_code)
    ]

    if not found_rows.empty:
      # استخراج كود المجموعة الأم واسم المؤسسة الجامعة
      parent_code = found_rows.iloc[0].get("الوحدة المدرسية", target_code)
      institution_name = found_rows.iloc[0].get("المؤسسة", "")

      # جلب كل العاملين في هذه المؤسسة المركزية وجميع الفرعيات التابعة لها
      all_institution_staff = df[
          df["الوحدة المدرسية"].astype(str).str.upper() == parent_code
      ]
      if all_institution_staff.empty:
        all_institution_staff = found_rows

      total_staff = len(all_institution_staff)
      females_cnt = (
          (all_institution_staff["SEX"].str.upper() == "F").sum()
          if "SEX" in all_institution_staff.columns
          else 0
      )
      males_cnt = (
          (all_institution_staff["SEX"].str.upper() == "M").sum()
          if "SEX" in all_institution_staff.columns
          else 0
      )

      # بطاقات معلومات المؤسسة
      st.markdown(
          f"""
                <div class="kpi-container">
                    <div class="kpi-box" style="border-right-color:#2563EB;">
                        <div class="kpi-label">المؤسسة / المجموعة المدرسية</div>
                        <div class="kpi-num" style="font-size:18px;">{institution_name}</div>
                    </div>
                    <div class="kpi-box" style="border-right-color:#0D9488;">
                        <div class="kpi-label">كود المركزية (المؤسسة الأم)</div>
                        <div class="kpi-num" style="font-size:18px;">{parent_code}</div>
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

      # ==================== جدول الوحدات المدرسية التابعة (الفرعيات) ====================
      st.markdown(
          "#### 🏫 الوحدات المدرسية والفرعيات التابعة لهذه المؤسسة:"
      )

      sub_units_summary = (
          all_institution_staff.groupby(["C_ETAB", "الوحدة المدرسية.1"])
          .size()
          .reset_index(name="عدد العاملين")
      )
      sub_units_summary.columns = [
          "كود الفرعية (C_ETAB)",
          "اسم الوحدة المدرسية / الفرعية",
          "عدد العاملين بها",
      ]
      sub_units_summary["النسبة"] = (
          sub_units_summary["عدد العاملين بها"] / total_staff * 100
      ).map("{:.1f}%".format)

      col_units, col_filter_unit = st.columns([2.5, 1.5])
      with col_units:
        st.dataframe(
            sub_units_summary, use_container_width=True, hide_index=True
        )

      with col_filter_unit:
        st.markdown("###### تصفية العرض حسب وحدة مدرسية معينة:")
        unit_options = ["عرض جميع الوحدات التابعة"] + [
            f"{r['كود الفرعية (C_ETAB)']} - {r['اسم الوحدة المدرسية / الفرعية']}"
            for _, r in sub_units_summary.iterrows()
        ]
        selected_unit_choice = st.selectbox(
            "اختر الفرعية لعرض أطرها فقط:", unit_options
        )

      # تحديد العرض النهائي حسب اختيار الوحدة
      if selected_unit_choice != "عرض جميع الوحدات التابعة":
        unit_code_selected = selected_unit_choice.split(" - ")[0].strip()
        final_display_df = all_institution_staff[
            all_institution_staff["C_ETAB"].astype(str).str.upper()
            == unit_code_selected
        ]
        st.info(
            f"عرض أطر الفرعية: **{selected_unit_choice}** (العدد:"
            f" {len(final_display_df)} أطر)"
        )
      else:
        final_display_df = all_institution_staff

      # ==================== جدول توزيع الأطر واللائحة الاسمية ====================
      st.markdown("---")
      col_t1, col_t2 = st.columns([1.2, 2.8])
      with col_t1:
        st.markdown("###### 📊 توزيع العاملين حسب المهمة:")
        stats_tasks = final_display_df["المهام"].value_counts().reset_index()
        stats_tasks.columns = ["المهمة", "العدد"]
        stats_tasks["النسبة"] = (
            stats_tasks["العدد"] / len(final_display_df) * 100
        ).map("{:.1f}%".format)
        st.dataframe(stats_tasks, use_container_width=True, hide_index=True)

      with col_t2:
        st.markdown(
            "###### 📋 اللائحة الاسمية للعاملين (مع كود الفرعية واسمها):"
        )
        disp_cols = [
            "PPR",
            "CIN",
            "NOM_AR",
            "NOM_FR",
            "المهام",
            "C_ETAB",
            "الوحدة المدرسية.1",
            "المؤسسة",
            "الدرجة",
            "DATE_DEBUT_AFF",
        ]
        avail_cols = [c for c in disp_cols if c in final_display_df.columns]
        st.dataframe(
            final_display_df[avail_cols], use_container_width=True, height=280
        )

      # زر تصدير تقرير المؤسسة ووحداتها
      buf_e = io.BytesIO()
      with pd.ExcelWriter(buf_e, engine="openpyxl") as writer:
        final_display_df.to_excel(
            writer, index=False, sheet_name="اللائحة الاسمية"
        )
        sub_units_summary.to_excel(
            writer, index=False, sheet_name="الوحدات المدرسية"
        )
        stats_tasks.to_excel(writer, index=False, sheet_name="إحصاء الأطر")
      st.download_button(
          label="📥 تصدير تقرير المؤسسة ووحداتها المدرسية (Excel)",
          data=buf_e.getvalue(),
          file_name=f"RH_{target_code}.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
    else:
      st.warning(f"لم يتم العثور على أي موظف مسجل بالكود: `{target_code}`")
  else:
    st.info(
        "👈 المرجو إدخال كود المؤسسة أو اختياره من القائمة ثم الضغط على **تأكيد"
        " الاستعلام**."
    )

# -------------------------------------------------------------
# 2. تبويب التنقيب الفردي
# -------------------------------------------------------------
with tab_search:
  with st.form("search_form"):
    st.markdown("##### 🔎 استعلام سريع بالرقم المالي، الاسم أو البطاقة الوطنية")
    q_col1, q_col2 = st.columns([3, 1])
    query_val = q_col1.text_input(
        "أدخل رقم التأجير (PPR)، الاسم، أو رقم البطاقة (CIN):",
        placeholder="مثال: 379274 أو محمد",
    ).strip()
    q_col2.write("")
    q_col2.write("")
    btn_query = q_col2.form_submit_button(
        "🔍 تنفيذ الاستعلام", use_container_width=True, type="primary"
    )

  if query_val:
    q_lower = query_val.lower()
    search_cols = [
        c
        for c in [
            "PPR",
            "CIN",
            "NOM_AR",
            "NOM_FR",
            "المؤسسة",
            "الوحدة المدرسية.1",
            "C_ETAB",
        ]
        if c in df.columns
    ]
    mask_q = df[search_cols].apply(
        lambda row: row.astype(str)
        .str.lower()
        .str.contains(q_lower, regex=False)
        .any(),
        axis=1,
    )
    res_q = df[mask_q]

    if not res_q.empty:
      st.success(f"تم العثور على **{len(res_q)}** نتيجة مطابقة:")
      disp_cols = [
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
      ]
      cols_to_render = [c for c in disp_cols if c in res_q.columns]
      st.dataframe(res_q[cols_to_render], use_container_width=True)

      if len(res_q) == 1:
        indiv = res_q.iloc[0]
        with st.container():
          st.markdown("---")
          st.markdown(
              "#### 👤 بطاقة المسار الإداري:"
              f" {indiv.get('NOM_AR', '')} ({indiv.get('NOM_FR', '')})"
          )
          c_b1, c_b2, c_b3 = st.columns(3)
          c_b1.markdown(f"**رقم التأجير (PPR):** `{indiv.get('PPR', '')}`")
          c_b1.markdown(f"**رقم البطاقة (CIN):** `{indiv.get('CIN', '')}`")
          c_b1.markdown(f"**تاريخ الازدياد:** {indiv.get('DATE_NAISSANCE', '')}")

          c_b2.markdown(f"**المهمة:** {indiv.get('المهام', '')}")
          c_b2.markdown(
              f"**الإطار والدرجة:** {indiv.get('الاطار', '')} -"
              f" {indiv.get('الدرجة', '')}"
          )
          c_b2.markdown(
              f"**تاريخ التوظيف:** {indiv.get('DATE_ANC_ADM', '')}"
          )

          c_b3.markdown(f"**المجموعة المدرسية:** {indiv.get('المؤسسة', '')}")
          c_b3.markdown(
              "**الوحدة المدرسية (الفرعية):**"
              f" {indiv.get('الوحدة المدرسية.1', '')} (`{indiv.get('C_ETAB', '')}`)"
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
  sub_count = df["C_ETAB"].nunique() if "C_ETAB" in df.columns else 0
  main_count = (
      df["الوحدة المدرسية"].nunique() if "الوحدة المدرسية" in df.columns else 0
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
                <div class="kpi-num" style="font-size:16px;">{main_count} مركزية | {sub_count} فرعية</div>
            </div>
        </div>
    """,
      unsafe_allow_html=True,
  )

  col_g1, col_g2 = st.columns(2)
  with col_g1:
    st.markdown("###### 📊 أعلى 10 مهام من حيث عدد الموظفين:")
    if "المهام" in df.columns:
      top_tasks_counts = df["المهام"].value_counts().head(10)
      top_tasks = pd.DataFrame({
          "المهمة / الإطار": top_tasks_counts.index,
          "العدد الإجمالي": top_tasks_counts.values,
      })
      st.dataframe(top_tasks, use_container_width=True, hide_index=True)

  with col_g2:
    st.markdown("###### 🏫 أكبر 10 مجموعات مدرسية من حيث عدد الأطر:")
    if "المؤسسة" in df.columns:
      top_etabs_counts = df["المؤسسة"].value_counts().head(10)
      top_etabs = pd.DataFrame({
          "المؤسسة التعليمية": top_etabs_counts.index,
          "عدد الموظفين": top_etabs_counts.values,
      })
      st.dataframe(top_etabs, use_container_width=True, hide_index=True)

# -------------------------------------------------------------
# 4. تبويب تعديل سجل موظف
# -------------------------------------------------------------
with tab_edit:
  st.markdown("##### ✏️ نموذج تعديل أو تحيين المعطيات الإدارية لموظف")

  col_p1, col_p2 = st.columns([2, 1])
  target_ppr = col_p1.text_input(
      "أدخل رقم التأجير (PPR) لتحميل استمارة التعديل:",
      placeholder="مثال: 379274",
  ).strip()

  if target_ppr:
    match_idx = df.index[df["PPR"] == target_ppr].tolist()
    if match_idx:
      t_row = df.loc[match_idx[0]]
      st.success(
          f"تم تحميل سجل: **{t_row.get('NOM_AR', '')}**"
          f" ({t_row.get('NOM_FR', '')})"
      )

      with st.form("edit_employee_form"):
        f1, f2, f3 = st.columns(3)
        u_nom_ar = f1.text_input(
            "الاسم الكامل بالعربية:", t_row.get("NOM_AR", "")
        )
        u_nom_fr = f2.text_input("الاسم بالفرنسية:", t_row.get("NOM_FR", ""))
        u_cin = f3.text_input(
            "رقم البطاقة الوطنية (CIN):", t_row.get("CIN", "")
        )

        u_task = f1.text_input("المهمة الحالية:", t_row.get("المهام", ""))
        u_cadre = f2.text_input("الإطار الإداري:", t_row.get("الاطار", ""))
        u_grade = f3.text_input("الدرجة:", t_row.get("الدرجة", ""))

        u_c_etab = f1.text_input(
            "كود الفرعية (C_ETAB):", t_row.get("C_ETAB", "")
        )
        u_sub_name = f2.text_input(
            "اسم الوحدة المدرسية (الفرعية):", t_row.get("الوحدة المدرسية.1", "")
        )
        u_grp_name = f3.text_input(
            "المجموعة المدرسية المركزية:", t_row.get("المؤسسة", "")
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
          if "المهام" in df.columns:
            df.loc[match_idx[0], "المهام"] = u_task
          if "الاطار" in df.columns:
            df.loc[match_idx[0], "الاطار"] = u_cadre
          if "الدرجة" in df.columns:
            df.loc[match_idx[0], "الدرجة"] = u_grade
          if "C_ETAB" in df.columns:
            df.loc[match_idx[0], "C_ETAB"] = u_c_etab
          if "الوحدة المدرسية.1" in df.columns:
            df.loc[match_idx[0], "الوحدة المدرسية.1"] = u_sub_name
          if "المؤسسة" in df.columns:
            df.loc[match_idx[0], "المؤسسة"] = u_grp_name
          st.session_state.df = df
          st.success(
              "✓ تم حفظ التعديلات بنجاح! يمكنك الآن تحميل الملف المحين من"
              " القائمة الجانبية."
          )
          st.rerun()
    else:
      st.warning("رقم التأجير غير موجود بقاعدة البيانات.")
