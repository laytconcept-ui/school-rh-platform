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

# ==================== هندسة التصميم والـ CSS الاحترافي ====================
# تم إخفاء تلميح "Press Enter to apply" وإعطاء مظهر متناسق ومريح للعين شبيه بالمنصات الحكومية الرسمية
st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&family=Tajawal:wght@400;500;700&display=swap');
    
    /* توجيه الصفحة بالكامل لليمين واستخدام خط عربي راقي */
    html, body, [class*="css"] {
        font-family: 'Cairo', 'Tajawal', sans-serif !important;
        direction: rtl;
        text-align: right;
    }

    /* إخفاء جملة Press Enter to apply نهائياً */
    div[data-testid="InputInstructions"] {
        display: none !important;
    }
    small[data-testid="stFormHelp"] {
        display: none !important;
    }

    /* خلفية الصفحة العامة ناعمة وواضحة */
    .stApp {
        background-color: #F8FAFC !important;
    }

    /* ترويسة المنصة الرسمية */
    .portal-header {
        background: linear-gradient(135deg, #1B365D 0%, #2A5298 100%);
        padding: 24px 30px;
        border-radius: 14px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(27, 54, 93, 0.15);
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .portal-title {
        font-size: 24px;
        font-weight: 800;
        margin: 0;
        color: #FFFFFF;
    }
    .portal-sub {
        font-size: 14px;
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

    /* صناديق البحث والبطاقات التفاعلية */
    .search-card {
        background: #FFFFFF;
        padding: 20px 25px;
        border-radius: 12px;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.04);
        border: 1px solid #E2E8F0;
        margin-bottom: 22px;
    }

    /* بطاقات المؤشرات الرقمية KPIs */
    .kpi-container {
        display: flex;
        gap: 15px;
        margin-bottom: 22px;
    }
    .kpi-box {
        flex: 1;
        background: #FFFFFF;
        padding: 18px 20px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        border-right: 5px solid #2563EB;
        border-top: 1px solid #E2E8F0;
        border-left: 1px solid #E2E8F0;
        border-bottom: 1px solid #E2E8F0;
        transition: transform 0.2s;
    }
    .kpi-box:hover {
        transform: translateY(-2px);
    }
    .kpi-label {
        color: #64748B;
        font-size: 13px;
        font-weight: 600;
        margin-bottom: 6px;
    }
    .kpi-num {
        color: #0F172A;
        font-size: 22px;
        font-weight: 800;
        letter-spacing: -0.5px;
    }

    /* تحسين حقول الإدخال */
    .stTextInput input, .stSelectbox select {
        border-radius: 8px !important;
        border: 1.5px solid #CBD5E1 !important;
        padding: 10px 14px !important;
        font-size: 14px !important;
        background-color: #FFFFFF !important;
        color: #1E293B !important;
    }
    .stTextInput input:focus, .stSelectbox select:focus {
        border-color: #2563EB !important;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12) !important;
    }

    /* تحسين أزرار المنصة */
    .stButton button {
        border-radius: 8px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        padding: 8px 22px !important;
        transition: all 0.2s ease !important;
    }
    
    /* تحسين شريط التبويبات Tabs */
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


# ==================== تحميل البيانات ====================
@st.cache_data
def load_data():
  file_path = "RH2026.xlsx"
  if not os.path.exists(file_path):
    return pd.DataFrame()
  df = pd.read_excel(file_path)
  df.columns = [str(c).strip() for c in df.columns]
  for col in df.columns:
    df[col] = df[col].fillna("").astype(str).str.strip()
  return df


df_source = load_data()

# فحص توفر الملف
if df_source.empty:
  st.error(
      "⚠️ لم يتم العثور على ملف `RH2026.xlsx` داخل المستودع. يرجى التأكد من رفعه"
      " في نفس المجلد."
  )
  uploaded = st.file_uploader(
      "أو يمكنك رفع الملف مؤقتاً لتشغيله الآن:", type=["xlsx", "xls"]
  )
  if uploaded:
    df_source = pd.read_excel(uploaded)
    df_source.columns = [str(c).strip() for c in df_source.columns]
    for col in df_source.columns:
      df_source[col] = df_source[col].fillna("").astype(str).str.strip()
  else:
    st.stop()

# حفظ البيانات داخل session_state
if "df" not in st.session_state or len(st.session_state.df) != len(df_source):
  st.session_state.df = df_source.copy()

df = st.session_state.df

# ==================== الشريط الجانبي ====================
with st.sidebar:
  st.markdown(
      """
        <div style="text-align: center; padding: 10px 0 20px 0;">
            <div style="font-size: 40px; margin-bottom: 5px;">🏛️</div>
            <h3 style="margin:0; color:#1E293B; font-weight:800;">بوابة الموارد البشرية</h3>
            <span style="color:#64748B; font-size:12px;">مديرية التربية الوطنية والتخطيط</span>
        </div>
    """,
      unsafe_allow_html=True,
  )

  st.markdown("---")
  st.metric("إجمالي الموظفين المسجلين", f"{len(df):,} فرد")

  # تنزيل النسخة المحدثة
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
    st.session_state.df = load_data()
    st.success("تم تحديث البيانات من الملف الرئيسي!")
    st.rerun()

# ==================== الترويسة الرئيسية ====================
st.markdown(
    """
    <div class="portal-header">
        <div>
            <h1 class="portal-title">منظومة تدبير الموارد البشرية وتتبع المؤسسات</h1>
            <div class="portal-sub">لوحة قيادة تفاعلية مخصصة للمصلحة - الموسم الدراسي 2026</div>
        </div>
        <div class="portal-badge">نسخة الويب السحابية v3.5</div>
    </div>
""",
    unsafe_allow_html=True,
)

# ==================== التبويبات التفاعلية للمنصة ====================
tab_etab, tab_search, tab_stats, tab_edit = st.tabs([
    "🏫 استعلام موظفي المؤسسة",
    "🔍 التنقيب والبحث الفردي",
    "📊 لوحة المؤشرات والإحصاء",
    "✏️ تدبير وتعديل سجل موظف",
])

# -------------------------------------------------------------
# 1. تبويب استعلام موظفي المؤسسة
# -------------------------------------------------------------
with tab_etab:
  with st.form("etab_form"):
    st.markdown("##### 📍 تحديد المؤسسة التعليمية أو الوحدة المدرسية")
    col_code, col_select, col_btn = st.columns([1.5, 2.5, 1])

    input_code = (
        col_code.text_input(
            "كود المؤسسة أو الفرعية:", placeholder="مثال: 03242Y"
        )
        .strip()
        .upper()
    )

    # قائمة فريدة من الأكواد والأسماء للاختيار السريع
    etab_choices = [""] + sorted(
        list(df["C_ETAB"].replace("", pd.NA).dropna().unique())
    )
    selected_dropdown = col_select.selectbox(
        "أو اختر من الأكواد المسجلة:", etab_choices
    )

    col_btn.write("")
    col_btn.write("")
    btn_search_etab = col_btn.form_submit_button(
        "🔎 تأكيد البحث", use_container_width=True, type="primary"
    )

  chosen_etab = input_code if input_code else selected_dropdown

  if chosen_etab:
    mask = (df["C_ETAB"].astype(str).str.upper() == chosen_etab) | (
        df["CD_ETAB"].astype(str).str.upper() == chosen_etab
    )
    res_etab = df[mask]

    if not res_etab.empty:
      r0 = res_etab.iloc[0]
      nom_sub = r0.get("الوحدة المدرسية.1", chosen_etab)
      nom_grp = r0.get("المؤسسة", "")

      # بطاقات المؤشرات الخاصة بالمؤسسة
      total_etab = len(res_etab)
      f_count = (
          (res_etab["SEX"].str.upper() == "F").sum()
          if "SEX" in res_etab.columns
          else 0
      )
      m_count = (
          (res_etab["SEX"].str.upper() == "M").sum()
          if "SEX" in res_etab.columns
          else 0
      )

      st.markdown(
          f"""
                <div class="kpi-container">
                    <div class="kpi-box" style="border-right-color:#2563EB;">
                        <div class="kpi-label">المؤسسة / الفرعية</div>
                        <div class="kpi-num" style="font-size:18px;">{nom_sub}</div>
                    </div>
                    <div class="kpi-box" style="border-right-color:#0D9488;">
                        <div class="kpi-label">المجموعة المدرسية</div>
                        <div class="kpi-num" style="font-size:18px;">{nom_grp if nom_grp else 'مركزية مستقلة'}</div>
                    </div>
                    <div class="kpi-box" style="border-right-color:#8B5CF6;">
                        <div class="kpi-label">مجموع العاملين</div>
                        <div class="kpi-num">{total_etab} أطر</div>
                    </div>
                    <div class="kpi-box" style="border-right-color:#F59E0B;">
                        <div class="kpi-label">توزيع الجنس (إناث / ذكور)</div>
                        <div class="kpi-num" style="font-size:17px;">{f_count} إناث ({f_count/total_etab*100:.0f}%) | {m_count} ذكور</div>
                    </div>
                </div>
            """,
          unsafe_allow_html=True,
      )

      # توزيع المهام بالمؤسسة
      col_t1, col_t2 = st.columns([1.2, 2.8])
      with col_t1:
        st.markdown("###### 📊 توزيع العاملين حسب المهمة:")
        stats_etab = res_etab["المهام"].value_counts().reset_index()
        stats_etab.columns = ["المهمة", "العدد"]
        stats_etab["النسبة"] = (
            stats_etab["العدد"] / total_etab * 100
        ).map("{:.1f}%".format)
        st.dataframe(stats_etab, use_container_width=True, hide_index=True)

      with col_t2:
        st.markdown("###### 📋 اللائحة الاسمية لأطر المؤسسة:")
        show_cols = [
            "PPR",
            "CIN",
            "NOM_AR",
            "NOM_FR",
            "المهام",
            "الدرجة",
            "DATE_DEBUT_AFF",
            "DATE_ANC_ADM",
        ]
        avail_cols = [c for c in show_cols if c in res_etab.columns]
        st.dataframe(
            res_etab[avail_cols], use_container_width=True, height=260
        )

      # زر تصدير اللائحة
      buf_e = io.BytesIO()
      with pd.ExcelWriter(buf_e, engine="openpyxl") as writer:
        res_etab.to_excel(writer, index=False, sheet_name=chosen_etab[:25])
        stats_etab.to_excel(writer, index=False, sheet_name="إحصاء الأطر")
      st.download_button(
          label=f"📥 تصدير تقرير المؤسسة {chosen_etab} (Excel)",
          data=buf_e.getvalue(),
          file_name=f"RH_{chosen_etab}.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
    else:
      st.warning(f"لم يتم العثور على أي موظف مسجل بالكود: `{chosen_etab}`")
  else:
    st.info(
        "👈 المرجو إدخال كود المؤسسة أو اختياره من القائمة ثم الضغط على **تأكيد"
        " البحث**."
    )

# -------------------------------------------------------------
# 2. تبويب التنقيب والبحث الفردي
# -------------------------------------------------------------
with tab_search:
  with st.form("search_form"):
    st.markdown("##### 🔎 استعلام سريع بالرقم المالي، الاسم أو البطاقة الوطنية")
    q_col1, q_col2 = st.columns([3, 1])
    query_val = q_col1.text_input(
        "أدخل رقم التأجير (PPR)، الاسم بالعربية، أو رقم البطاقة (CIN):",
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
          "الوحدة المدرسية.1",
          "المؤسسة",
          "DATE_DEBUT_AFF",
      ]
      cols_to_render = [c for c in disp_cols if c in res_q.columns]
      st.dataframe(res_q[cols_to_render], use_container_width=True)

      # عرض بطاقة تفصيلية عند اختيار موظف واحد
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

          c_b3.markdown(f"**المؤسسة:** {indiv.get('المؤسسة', '')}")
          c_b3.markdown(
              f"**الفرعية:** {indiv.get('الوحدة المدرسية.1', '')}"
              f" (`{indiv.get('C_ETAB', '')}`)"
          )
          c_b3.markdown(
              "**تاريخ التعيين بالمؤسسة:**"
              f" {indiv.get('DATE_DEBUT_AFF', '')}"
          )
    else:
      st.warning("لم يتم العثور على أي موظف بهذه المعطيات.")

# -------------------------------------------------------------
# 3. تبويب لوحة المؤشرات والإحصاء العام
# -------------------------------------------------------------
with tab_stats:
  st.markdown("##### 📈 مؤشرات التوزيع الإجمالي للموارد البشرية بالإقليم")

  total_all = len(df)
  females_all = (
      (df["SEX"].str.upper() == "F").sum() if "SEX" in df.columns else 0
  )
  males_all = (df["SEX"].str.upper() == "M").sum() if "SEX" in df.columns else 0
  etab_count = df["C_ETAB"].nunique() if "C_ETAB" in df.columns else 0

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
                <div class="kpi-label">المؤسسات والفرعيات المسجلة</div>
                <div class="kpi-num">{etab_count} مؤسسة</div>
            </div>
        </div>
    """,
      unsafe_allow_html=True,
  )

  col_g1, col_g2 = st.columns(2)
  with col_g1:
    st.markdown("###### 📊 أعلى 10 مهام من حيث عدد الموظفين:")
    top_tasks = df["المهام"].value_counts().head(10).reset_index()
    top_tasks.columns = ["المهمة / الإطار", "العدد الإجمالي"]
    st.dataframe(top_tasks, use_container_width=True, hide_index=True)

  with col_g2:
    st.markdown("###### 🏫 أكبر 10 مؤسسات من حيث عدد الأطر:")
    top_etabs = df["المؤسسة"].value_counts().head(10).reset_index()
    top_etabs.columns = ["المؤسسة التعليمية", "عدد الموظفين"]
    st.dataframe(top_etabs, use_container_width=True, hide_index=True)

# -------------------------------------------------------------
# 4. تبويب تدبير وتعديل سجل موظف
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

        u_etab_code = f1.text_input(
            "كود المؤسسة (C_ETAB):", t_row.get("C_ETAB", "")
        )
        u_sub_name = f2.text_input(
            "اسم الفرعية:", t_row.get("الوحدة المدرسية.1", "")
        )
        u_grp_name = f3.text_input(
            "المجموعة المدرسية:", t_row.get("المؤسسة", "")
        )

        btn_commit = st.form_submit_button(
            "💾 حفظ وتثبيت التعديلات في الجلسة", type="primary"
        )
        if btn_commit:
          df.loc[match_idx[0], "NOM_AR"] = u_nom_ar
          df.loc[match_idx[0], "NOM_FR"] = u_nom_fr
          df.loc[match_idx[0], "CIN"] = u_cin
          df.loc[match_idx[0], "المهام"] = u_task
          df.loc[match_idx[0], "الاطار"] = u_cadre
          df.loc[match_idx[0], "الدرجة"] = u_grade
          df.loc[match_idx[0], "C_ETAB"] = u_etab_code
          df.loc[match_idx[0], "الوحدة المدرسية.1"] = u_sub_name
          df.loc[match_idx[0], "المؤسسة"] = u_grp_name
          st.session_state.df = df
          st.success(
              "✓ تم حفظ التعديلات بنجاح! يمكنك الآن تحميل الملف المحين من"
              " القائمة الجانبية."
          )
          st.rerun()
    else:
      st.warning("رقم التأجير غير موجود بقاعدة البيانات.")
