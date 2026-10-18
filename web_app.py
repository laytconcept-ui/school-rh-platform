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
    initial_sidebar_state="collapsed",
)

# تخصيص التصميم ليتوافق 100% مع شاشات الهواتف والحواسيب معاً
st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&family=Tajawal:wght@400;500;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Cairo', 'Tajawal', sans-serif !important;
        direction: rtl;
        text-align: right;
        line-height: 1.6;
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

    /* ترويسة المنصة */
    .portal-header {
        background: linear-gradient(135deg, #1B365D 0%, #2A5298 100%);
        padding: 16px 20px;
        border-radius: 12px;
        color: white;
        margin-bottom: 16px;
        box-shadow: 0 4px 15px rgba(27, 54, 93, 0.15);
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        justify-content: space-between;
        align-items: center;
    }
    .portal-title {
        font-size: 19px;
        font-weight: 800;
        margin: 0;
        color: #FFFFFF;
        line-height: 1.4;
    }
    .portal-sub {
        font-size: 12px;
        color: #D6E4FF;
        margin-top: 3px;
    }
    .portal-badge {
        background: rgba(255, 255, 255, 0.18);
        border: 1px solid rgba(255, 255, 255, 0.3);
        padding: 4px 12px;
        border-radius: 16px;
        font-size: 11px;
        font-weight: 600;
    }

    /* بطاقات المؤشرات الرقمية - متجاوبة ومرنة */
    .kpi-container {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        margin-bottom: 16px;
    }
    .kpi-box {
        flex: 1 1 calc(50% - 10px);
        min-width: 140px;
        background: #FFFFFF;
        padding: 12px 14px;
        border-radius: 10px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
        border-right: 4px solid #2563EB;
        border-top: 1px solid #E2E8F0;
        border-left: 1px solid #E2E8F0;
        border-bottom: 1px solid #E2E8F0;
    }
    .kpi-label {
        color: #64748B;
        font-size: 11px;
        font-weight: 600;
        margin-bottom: 4px;
    }
    .kpi-num {
        color: #0F172A;
        font-size: 16px;
        font-weight: 800;
        word-break: break-word;
    }

    /* الحقول والأزرار */
    .stTextInput input, .stSelectbox select {
        border-radius: 8px !important;
        border: 1.5px solid #CBD5E1 !important;
        padding: 8px 12px !important;
        font-size: 13px !important;
        background-color: #FFFFFF !important;
        color: #1E293B !important;
    }
    .stButton button {
        border-radius: 8px !important;
        font-weight: 700 !important;
        font-size: 13px !important;
        padding: 8px 18px !important;
    }

    /* شريط التبويبات المتجاوب */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        border-bottom: 2px solid #E2E8F0;
        overflow-x: auto;
        white-space: nowrap;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 6px 6px 0 0 !important;
        padding: 8px 14px !important;
        font-weight: 700 !important;
        font-size: 12px !important;
    }

    /* مواءمة شاشات الهاتف المحمول */
    @media (max-width: 768px) {
        .portal-header {
            flex-direction: column;
            align-items: flex-start;
        }
        .portal-title {
            font-size: 17px !important;
        }
        .kpi-box {
            flex: 1 1 100% !important;
        }
        [data-testid="column"] {
            width: 100% !important;
            flex: 1 1 100% !important;
            min-width: 100% !important;
        }
        div[data-testid="stDialog"] div[role="dialog"] {
            width: 95vw !important;
            max-width: 95vw !important;
            margin: auto !important;
            padding: 10px !important;
        }
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
    df = pd.read_excel(source, dtype=str).fillna("")
    df.columns = [str(c).strip() for c in df.columns]

    for c in [
        "PPR",
        "CIN",
        "C_ETAB",
        "CD_ETAB",
        "الوحدة المدرسية",
        "c_unité",
        "c_unite",
    ]:
      if c in df.columns:
        df[c] = df[c].apply(clean_value)

    # تحديد صريح لاسم الوحدة المدرسية لمنع ظهور الكود بدلاً من الاسم
    if "الوحدة المدرسية.1" in df.columns:
      df["اسم_الوحدة_المدرسية"] = df["الوحدة المدرسية.1"].astype(str).str.strip()
    elif "اسم الفرعية" in df.columns:
      df["اسم_الوحدة_المدرسية"] = df["اسم الفرعية"].astype(str).str.strip()
    else:
      df["اسم_الوحدة_المدرسية"] = (
          df.get("الوحدة المدرسية", "").astype(str).str.strip()
      )

    if "c_unité" in df.columns:
      df["كود_الوحدة_المدرسية"] = df["c_unité"].astype(str).str.strip()
    elif "c_unite" in df.columns:
      df["كود_الوحدة_المدرسية"] = df["c_unite"].astype(str).str.strip()
    elif "الوحدة المدرسية" in df.columns:
      df["كود_الوحدة_المدرسية"] = df["الوحدة المدرسية"].astype(str).str.strip()
    else:
      df["كود_الوحدة_المدرسية"] = df.get("C_ETAB", "").astype(str).str.strip()

    return df
  except Exception as e:
    st.error(f"حدث خطأ أثناء قراءة ملف الإكسيل: {e}")
    return pd.DataFrame()


def export_to_excel(df_to_export, sheet_name="Data"):
  out = io.BytesIO()
  with pd.ExcelWriter(out, engine="openpyxl") as writer:
    df_to_export.to_excel(writer, index=False, sheet_name=sheet_name)
  return out.getvalue()


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
COL_UNIT_CODE = "كود_الوحدة_المدرسية"
COL_UNIT_NAME = "اسم_الوحدة_المدرسية"
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

# ==================== نافذة التعديل المنبثقة ====================
if hasattr(st, "dialog"):
  dialog_decorator = st.dialog(
      "✏️ بطاقة تعديل المعطيات الشاملة للأستاذ", width="large"
  )
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
      f"### 👤 {t.get('NOM_AR', '')} | {t.get('NOM_FR', '')} (رقم التأجير:"
      f" `{ppr_selected}`)"
  )

  with st.form("modal_edit_teacher_form"):
    st.markdown("##### 📌 المعطيات الشخصية والمهنية")
    c1, c2 = st.columns(2)
    v_nom_ar = c1.text_input("الاسم الكامل بالعربية:", t.get("NOM_AR", ""))
    v_nom_fr = c2.text_input("الاسم بالفرنسية:", t.get("NOM_FR", ""))
    v_cin = c1.text_input("رقم البطاقة (CIN):", t.get("CIN", ""))
    v_sex = c2.text_input("الجنس (M/F):", t.get("SEX", ""))
    v_sit = c1.text_input("الحالة العائلية:", t.get("SIT_FAM", ""))
    v_nais = c2.text_input("تاريخ الازدياد:", t.get("DATE_NAISSANCE", ""))

    st.markdown("##### 🏫 معطيات التعيين والوحدة المدرسية")
    c3, c4 = st.columns(2)
    v_c_etab = c3.text_input("كود المؤسسة (C_ETAB):", t.get(COL_C_ETAB, ""))
    v_u_code = c4.text_input(
        "كود الوحدة المدرسية (c_unité):", t.get(COL_UNIT_CODE, "")
    )
    v_u_name = c3.text_input(
        "اسم الوحدة المدرسية (الفرعية):", t.get(COL_UNIT_NAME, "")
    )
    v_inst = c4.text_input(
        "المجموعة المدرسية المركزية:", t.get(COL_INSTITUTION, "")
    )
    v_task = c3.text_input("المهمة:", t.get(COL_TASK, ""))
    v_cadre = c4.text_input("الإطار الإداري:", t.get("الاطار", ""))

    st.markdown("##### 📜 الوضعية الإدارية والترقية")
    c5, c6 = st.columns(2)
    v_grade = c5.text_input("الدرجة:", t.get("الدرجة", ""))
    v_d_aff = c6.text_input(
        "تاريخ التعيين بالمؤسسة:", t.get("DATE_DEBUT_AFF", "")
    )
    v_d_rec = c5.text_input("تاريخ التوظيف:", t.get("DATE_ANC_ADM", ""))
    v_d_gd = c6.text_input("تاريخ مفعول الدرجة:", t.get("DATE_EFFET_GD", ""))
    v_d_elo = c5.text_input("تاريخ مفعول الرتبة:", t.get("DATE_EFFET_ELO", ""))
    v_cd_etab = c6.text_input("كود التعيين (CD_ETAB):", t.get("CD_ETAB", ""))

    st.write("")
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
          "اسم_الوحدة_المدرسية": v_u_name,
          "كود_الوحدة_المدرسية": v_u_code,
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
      st.success("✓ تم حفظ كافة التعديلات بنجاح!")
      st.rerun()


# ==================== الشريط الجانبي ====================
with st.sidebar:
  st.markdown(
      """
        <div style="text-align: center; padding: 10px 0 15px 0;">
            <div style="font-size: 34px; margin-bottom: 5px;">🏛️</div>
            <h3 style="margin:0; color:#1E293B; font-weight:800; font-size:18px;">بوابة الموارد البشرية</h3>
            <span style="color:#64748B; font-size:12px;">مصلحة الموارد البشرية والتخطيط</span>
        </div>
    """,
      unsafe_allow_html=True,
  )

  st.markdown("---")
  st.metric("إجمالي الموظفين المسجلين", f"{len(df):,} فرد")

  if st.sidebar.button(
      "📥 تجهيز ملف Excel للتحميل الكامل", use_container_width=True
  ):
    st.session_state["ready_export_all"] = export_to_excel(df, "PERSONNEL")

  if "ready_export_all" in st.session_state:
    st.sidebar.download_button(
        label="💾 تنزيل ملف Excel الكامل",
        data=st.session_state["ready_export_all"],
        file_name="RH2026_Updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

  st.markdown("<br>", unsafe_allow_html=True)
  if st.button("🔄 إعادة ضبط وتحديث القاعدة", use_container_width=True):
    if detected_file:
      st.session_state.df = load_data(detected_file)
      st.session_state.pop("ready_export_all", None)
      st.success("تم تحديث البيانات بنجاح!")
      st.rerun()

# ==================== الترويسة ====================
st.markdown(
    """
    <div class="portal-header">
        <div>
            <h1 class="portal-title">منظومة تدبير الموارد البشرية وتتبع المؤسسات</h1>
            <div class="portal-sub">فلترة مباشرة بكود المؤسسة (C_ETAB) واستعلام تفاعلي - 2026</div>
        </div>
        <div class="portal-badge">نسخة الهواتف والحواسيب v6.0</div>
    </div>
""",
    unsafe_allow_html=True,
)

# ==================== التبويبات ====================
tab_etab, tab_search, tab_stats, tab_edit = st.tabs([
    "🏫 استعلام المؤسسة",
    "🔍 البحث الفردي",
    "📊 المؤشرات والإحصاء",
    "✏️ تعديل سجل موظف",
])

# -------------------------------------------------------------
# 1. تبويب استعلام المؤسسة بكود C_ETAB
# -------------------------------------------------------------
with tab_etab:
  with st.form("etab_filter_form"):
    st.markdown("##### 📍 تحديد المؤسسة بكود C_ETAB")
    col_c1, col_c2 = st.columns(2)

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
          [x for x in df[COL_C_ETAB].unique() if x and x != "nan"]
      )

    etab_dropdown = col_c2.selectbox("أو اختر كود C_ETAB من اللائحة:", all_etabs)

    btn_apply = st.form_submit_button(
        "🔎 تأكيد الاستعلام", use_container_width=True, type="primary"
    )

  target_code = etab_input if etab_input else etab_dropdown

  if target_code:
    # مطابقة تامة وصريحة لكود C_ETAB
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
                        <div class="kpi-num" style="font-size:16px;">{institution_name if institution_name else 'مؤسسة مستقلة'}</div>
                    </div>
                    <div class="kpi-box" style="border-right-color:#0D9488;">
                        <div class="kpi-label">كود المؤسسة (C_ETAB)</div>
                        <div class="kpi-num" style="font-size:16px;">{target_code}</div>
                    </div>
                    <div class="kpi-box" style="border-right-color:#8B5CF6;">
                        <div class="kpi-label">مجموع العاملين</div>
                        <div class="kpi-num">{total_staff} أطر</div>
                    </div>
                    <div class="kpi-box" style="border-right-color:#F59E0B;">
                        <div class="kpi-label">توزيع الجنسين</div>
                        <div class="kpi-num" style="font-size:15px;">{females_cnt} إناث | {males_cnt} ذكور</div>
                    </div>
                </div>
            """,
          unsafe_allow_html=True,
      )

      # جدول الوحدات المدرسية التابعة
      st.markdown(
          f"#### 🏫 الوحدات المدرسية التابعة للمؤسسة `{target_code}`:"
      )

      sub_units_summary = (
          res_etab.groupby([COL_UNIT_CODE, COL_UNIT_NAME])
          .size()
          .reset_index(name="عدد العاملين")
      )
      sub_units_summary.columns = [
          "كود الوحدة المدرسية",
          "اسم الوحدة المدرسية (الفرعية)",
          "عدد العاملين بها",
      ]
      sub_units_summary["النسبة"] = (
          sub_units_summary["عدد العاملين بها"] / total_staff * 100
      ).map("{:.1f}%".format)

      col_units, col_filter_unit = st.columns([2, 1])
      with col_units:
        st.dataframe(
            sub_units_summary, use_container_width=True, hide_index=True
        )

      with col_filter_unit:
        st.markdown("###### تصفية العرض حسب فرعية:")
        unit_opts = ["عرض جميع الوحدات المدرسية التابعة"]
        for _, r in sub_units_summary.iterrows():
          u_code = r.get("كود الوحدة المدرسية", "")
          u_name = r.get("اسم الوحدة المدرسية (الفرعية)", "")
          unit_opts.append(f"{u_code} - {u_name}")

        selected_unit_choice = st.selectbox(
            "اختر الفرعية لعرض أطرها فقط:",
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
            f"عرض أطر الفرعية: **{selected_unit_choice}** (العدد:"
            f" {len(final_display_df)} أطر)"
        )
      else:
        final_display_df = res_etab

      st.markdown("---")
      # إحصاء المهام
      stats_tasks = (
          final_display_df[COL_TASK].value_counts().reset_index()
      )
      stats_tasks.columns = ["المهمة", "العدد"]
      if len(final_display_df) > 0:
        stats_tasks["النسبة"] = (
            stats_tasks["العدد"] / len(final_display_df) * 100
        ).map("{:.1f}%".format)

      # اللائحة الاسمية للعاملين بالمؤسسة
      st.markdown("#### 📋 اللائحة الاسمية للعاملين بالمؤسسة:")
      st.caption(
          "💡 **انقر على سطر أي أستاذ في الجدول** لفتح النافذة العريضة لتعديل"
          " معطياته الإدارية والشخصية فورياً."
      )

      # ترتيب الأعمدة: اسم الوحدة المدرسية يظهر بوضوح بجانب كودها
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

      rename_disp = {
          COL_PPR: "رقم التأجير",
          "CIN": "رقم البطاقة",
          "NOM_AR": "الاسم بالعربية",
          "NOM_FR": "الاسم بالفرنسية",
          COL_TASK: "المهمة",
          COL_UNIT_NAME: "اسم الوحدة المدرسية (الفرعية)",
          COL_UNIT_CODE: "كود الوحدة المدرسية",
          COL_C_ETAB: "كود المؤسسة",
          "الدرجة": "الدرجة",
          "DATE_DEBUT_AFF": "تاريخ التعيين",
      }

      rendered_df = final_display_df[avail_cols].rename(columns=rename_disp)

      selection = st.dataframe(
          rendered_df,
          use_container_width=True,
          height=320,
          on_select="rerun",
          selection_mode="single-row",
      )

      # فتح النافذة المنبثقة عند اختيار السطر
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

      # تصدير التقرير
      if st.button("📥 تجهيز ملف تقرير المؤسسة (Excel)"):
        st.session_state[f"export_{target_code}"] = export_to_excel(
            final_display_df, "اللائحة الاسمية"
        )

      if f"export_{target_code}" in st.session_state:
        st.download_button(
            label=f"💾 تنزيل تقرير {target_code} (Excel)",
            data=st.session_state[f"export_{target_code}"],
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
# 2. تبويب التنقيب الفردي السريع
# -------------------------------------------------------------
with tab_search:
  with st.form("search_form"):
    st.markdown(
        "##### 🔎 استعلام فوري سريع بالرقم المالي، الاسم أو البطاقة الوطنية"
    )
    col_q1, col_q2 = st.columns([3, 1])
    query_val = col_q1.text_input(
        "أدخل رقم التأجير (PPR)، الاسم، أو رقم البطاقة (CIN):",
        placeholder="مثال: 1114650 أو محمد",
    ).strip()
    btn_query = col_q2.form_submit_button(
        "🔍 تنفيذ الاستعلام", use_container_width=True, type="primary"
    )

  if query_val:
    q_clean = clean_value(query_val).lower()

    # بحث متّجهي سريع جداً
    mask_ppr = df[COL_PPR].astype(str).str.strip() == q_clean
    mask_cin = (
        df["CIN"].astype(str).str.strip().str.lower() == q_clean
        if "CIN" in df.columns
        else False
    )
    mask_names = df["NOM_AR"].astype(str).str.contains(
        query_val, regex=False, na=False
    ) | df["NOM_FR"].astype(str).str.lower().str.contains(
        q_clean, regex=False, na=False
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

      rename_disp_q = {
          COL_PPR: "رقم التأجير",
          "CIN": "رقم البطاقة",
          "NOM_AR": "الاسم بالعربية",
          "NOM_FR": "الاسم بالفرنسية",
          COL_TASK: "المهمة",
          COL_UNIT_NAME: "اسم الوحدة المدرسية (الفرعية)",
          COL_UNIT_CODE: "كود الوحدة",
          COL_C_ETAB: "كود المؤسسة",
          "الدرجة": "الدرجة",
          "DATE_DEBUT_AFF": "تاريخ التعيين",
      }

      rendered_q = res_q[cols_to_render].rename(columns=rename_disp_q)

      sel_q = st.dataframe(
          rendered_q,
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
          c_b1, c_b2 = st.columns(2)
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

          c_b1.markdown(
              "**المؤسسة (المجموعة):**"
              f" {indiv.get(COL_INSTITUTION, '')} (`{indiv.get(COL_C_ETAB, '')}`)"
          )
          c_b2.markdown(
              "**الوحدة المدرسية (الفرعية):**"
              f" {indiv.get(COL_UNIT_NAME, '')} (`{indiv.get(COL_UNIT_CODE, '')}`)"
          )
          c_b1.markdown(
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
                <div class="kpi-num" style="font-size:15px;">{main_etab_count} مؤسسة | {sub_unit_count} وحدة مدرسية</div>
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
        f1, f2 = st.columns(2)
        u_nom_ar = f1.text_input(
            "الاسم الكامل بالعربية:", t_row.get("NOM_AR", "")
        )
        u_nom_fr = f2.text_input("الاسم بالفرنسية:", t_row.get("NOM_FR", ""))
        u_cin = f1.text_input(
            "رقم البطاقة الوطنية (CIN):", t_row.get("CIN", "")
        )
        u_task = f2.text_input("المهمة الحالية:", t_row.get(COL_TASK, ""))
        u_cadre = f1.text_input("الإطار الإداري:", t_row.get("الاطار", ""))
        u_grade = f2.text_input("الدرجة:", t_row.get("الدرجة", ""))

        u_c_etab = f1.text_input(
            "كود المؤسسة (C_ETAB):", t_row.get(COL_C_ETAB, "")
        )
        u_unit_code = f2.text_input(
            "كود الوحدة المدرسية:", t_row.get(COL_UNIT_CODE, "")
        )
        u_unit_name = f1.text_input(
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
