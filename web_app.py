import io
import os
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="منظومة تدبير الموارد البشرية 2026",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# تخصيص المظهر وتنسيق الواجهة لتبدو أنيقة ومريحة للعين بخلفية فاتحة ولمسات زرقاء
st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Tajawal', sans-serif;
        direction: rtl;
        text-align: right;
    }
    .main {
        background-color: #F8FAFC;
    }
    .metric-card {
        background: white;
        padding: 15px 20px;
        border-radius: 10px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
        border-top: 4px solid #2980B9;
        text-align: center;
    }
    .metric-title {
        color: #7F8C8D;
        font-size: 13px;
        font-weight: 500;
        margin-bottom: 5px;
    }
    .metric-value {
        color: #2C3E50;
        font-size: 22px;
        font-weight: bold;
    }
</style>
""",
    unsafe_allow_html=True,
)


# ==================== تحميل البيانات ====================
@st.cache_data
def load_data(file_source="RH2026.xlsx"):
  if isinstance(file_source, str) and not os.path.exists(file_source):
    return pd.DataFrame()
  df = pd.read_excel(file_source)
  df.columns = [str(c).strip() for c in df.columns]
  for col in df.columns:
    df[col] = df[col].fillna("").astype(str).str.strip()
  return df


# فحص تواجد الملف الافتراضي أو إتاحة رفعه
if os.path.exists("RH2026.xlsx"):
  df_raw = load_data("RH2026.xlsx")
else:
  st.sidebar.warning(
      "لم يتم العثور على RH2026.xlsx في المجلد، يرجى رفعه أدناه."
  )
  uploaded_file = st.sidebar.file_uploader(
      "رفع ملف RH2026.xlsx", type=["xlsx", "xls"]
  )
  if uploaded_file:
    df_raw = load_data(uploaded_file)
  else:
    df_raw = pd.DataFrame()

if df_raw.empty:
  st.info(
      "يرجى التأكد من وجود ملف `RH2026.xlsx` داخل مستودع GitHub بجانب ملف الكود."
  )
  st.stop()

# حفظ البيانات في حالة الجلسة (Session State) لتمكين التعديل الحي
if "df" not in st.session_state or len(st.session_state.df) != len(df_raw):
  st.session_state.df = df_raw.copy()

df = st.session_state.df

# ==================== القائمة الجانبية ====================
st.sidebar.title("🏛️ تدبير الموارد البشرية")
st.sidebar.caption("الموسم الدراسي 2026")
st.sidebar.markdown("---")

st.sidebar.success(f"✓ متصل بالقاعدة: ({len(df):,} موظفاً)")

# زر إعادة تحميل البيانات الأصلية
if st.sidebar.button("🔄 إعادة تحميل المعطيات الأصلية"):
  st.session_state.df = load_data("RH2026.xlsx")
  st.rerun()

# تصدير البيانات إلى Excel
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
  df.to_excel(writer, index=False, sheet_name="PERSONNEL")
st.sidebar.download_button(
    label="📥 تصدير قاعدة البيانات كاملة (Excel)",
    data=buffer.getvalue(),
    file_name="RH2026_Export.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ==================== محرك التصفية والبحث ====================
st.markdown("### 🔍 الاستعلام والبحث السريع")

col1, col2, col3 = st.columns([1.5, 2, 1])

# 1. كود المؤسسة
etab_input = (
    col1.text_input(
        "كود المؤسسة أو الفرعية:", placeholder="مثال: 03242Y أو 03252J"
    )
    .strip()
    .upper()
)

# 2. بحث فوري شامل
search_query = col2.text_input(
    "بحث بالاسم، رقم التأجير (PPR) أو البطاقة (CIN):",
    placeholder="اكتب للبحث...",
).strip()

# تطبيق الفلترة
filtered_df = df.copy()

if etab_input:
  mask = (filtered_df["C_ETAB"].astype(str).str.upper() == etab_input) | (
      filtered_df["CD_ETAB"].astype(str).str.upper() == etab_input
  )
  filtered_df = filtered_df[mask]

if search_query:
  q_lower = search_query.lower()
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
      if c in filtered_df.columns
  ]
  mask = filtered_df[search_cols].apply(
      lambda row: row.astype(str).str.lower().str.contains(q_lower).any(),
      axis=1,
  )
  filtered_df = filtered_df[mask]

# ==================== بطاقات المؤشرات (KPIs) ====================
total_res = len(filtered_df)
females = (
    (filtered_df["SEX"].str.upper() == "F").sum()
    if "SEX" in filtered_df.columns and total_res > 0
    else 0
)
males = (
    (filtered_df["SEX"].str.upper() == "M").sum()
    if "SEX" in filtered_df.columns and total_res > 0
    else 0
)

etab_label = "جميع المؤسسات"
if etab_input and not filtered_df.empty:
  r0 = filtered_df.iloc[0]
  sub = r0.get("الوحدة المدرسية.1", etab_input)
  grp = r0.get("المؤسسة", "")
  etab_label = f"{sub} ({grp})" if grp and grp != sub else sub

k1, k2, k3, k4 = st.columns(4)
k1.markdown(
    '<div class="metric-card"><div class="metric-title">مجموع'
    f' الموظفين</div><div class="metric-value">{total_res:,} فرد</div></div>',
    unsafe_allow_html=True,
)
k2.markdown(
    '<div class="metric-card"><div class="metric-title">عدد الإناث</div><div'
    f' class="metric-value">{females:,}'
    f" ({(females/total_res*100 if total_res else 0):.1f}%)</div></div>",
    unsafe_allow_html=True,
)
k3.markdown(
    '<div class="metric-card"><div class="metric-title">عدد الذكور</div><div'
    f' class="metric-value">{males:,}'
    f" ({(males/total_res*100 if total_res else 0):.1f}%)</div></div>",
    unsafe_allow_html=True,
)
k4.markdown(
    '<div class="metric-card"><div class="metric-title">المؤسسة'
    ' المحددة</div><div class="metric-value"'
    f' style="font-size:16px;">{etab_label[:28]}</div></div>',
    unsafe_allow_html=True,
)

st.write("")

# ==================== إحصاء الأطر والمهام ====================
if not filtered_df.empty and "المهام" in filtered_df.columns:
  with st.expander(
      "📊 إحصاء وتوزيع العاملين حسب المهمة والإطار", expanded=True
  ):
    stats = filtered_df["المهام"].value_counts().reset_index()
    stats.columns = ["المهمة / الإطار", "العدد الإجمالي"]
    stats["النسبة المئوية"] = (
        stats["العدد الإجمالي"] / total_res * 100
    ).map("{:.1f}%".format)

    st.dataframe(stats, use_container_width=True, hide_index=True)

# ==================== جدول اللائحة الاسمية التفصيلية ====================
st.markdown("### 👥 اللائحة الاسمية للعاملين")

display_cols = [
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
available_cols = [c for c in display_cols if c in filtered_df.columns]

# تصدير اللائحة المصفاة
buf_filtered = io.BytesIO()
with pd.ExcelWriter(buf_filtered, engine="openpyxl") as writer:
  filtered_df[available_cols].to_excel(
      writer, index=False, sheet_name="اللائحة الاسمية"
  )
st.download_button(
    label="📥 تصدير اللائحة الحالية المعروضة (Excel)",
    data=buf_filtered.getvalue(),
    file_name="RH_Export_Selection.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# عرض جدول البيانات مع إمكانية التحديد
st.dataframe(filtered_df[available_cols], use_container_width=True, height=400)

# ==================== تعديل معطيات موظف ====================
st.markdown("---")
with st.expander("✏️ تعديل بيانات موظف محدد", expanded=False):
  c_sel, _ = st.columns([2, 2])
  selected_ppr = c_sel.selectbox(
      "اختر رقم التأجير (PPR) لتعديل معطياته:",
      [""] + list(filtered_df["PPR"].unique()),
  )

  if selected_ppr:
    teacher_idx = df.index[df["PPR"] == selected_ppr].tolist()
    if teacher_idx:
      t_row = df.loc[teacher_idx[0]]
      with st.form("edit_form"):
        st.subheader(
            f"تعديل: {t_row.get('NOM_AR', '')} ({t_row.get('NOM_FR', '')})"
        )
        f1, f2 = st.columns(2)
        new_nom_ar = f1.text_input("الاسم بالعربية:", t_row.get("NOM_AR", ""))
        new_nom_fr = f2.text_input("الاسم بالفرنسية:", t_row.get("NOM_FR", ""))
        new_cin = f1.text_input("رقم البطاقة الوطنية:", t_row.get("CIN", ""))
        new_task = f2.text_input("المهمة:", t_row.get("المهام", ""))
        new_cadre = f1.text_input("الإطار:", t_row.get("الاطار", ""))
        new_grade = f2.text_input("الدرجة:", t_row.get("الدرجة", ""))
        new_etab_code = f1.text_input(
            "كود المؤسسة:", t_row.get("C_ETAB", "")
        )
        new_sub_name = f2.text_input(
            "اسم الفرعية:", t_row.get("الوحدة المدرسية.1", "")
        )

        submitted = st.form_submit_button("💾 حفظ التعديلات")
        if submitted:
          df.loc[teacher_idx[0], "NOM_AR"] = new_nom_ar
          df.loc[teacher_idx[0], "NOM_FR"] = new_nom_fr
          df.loc[teacher_idx[0], "CIN"] = new_cin
          df.loc[teacher_idx[0], "المهام"] = new_task
          df.loc[teacher_idx[0], "الاطار"] = new_cadre
          df.loc[teacher_idx[0], "الدرجة"] = new_grade
          df.loc[teacher_idx[0], "C_ETAB"] = new_etab_code
          df.loc[teacher_idx[0], "الوحدة المدرسية.1"] = new_sub_name
          st.session_state.df = df
          st.success("تم تحديث البيانات بنجاح!")
          st.rerun()
