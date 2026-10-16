import io
import os
import glob
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="بوابة تدبير الموارد البشرية | RH Portal",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# تخصيص المظهر وتنسيق الواجهة وإخفاء إشعار "Press Enter to apply"
st.markdown("""
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

    .stButton button {
        border-radius: 8px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        padding: 8px 22px !important;
        transition: all 0.2s ease !important;
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
""", unsafe_allow_html=True)

# ==================== دوال قراءة وتنظيف البيانات ====================
def find_excel_file():
    """البحث الذكي عن ملف RH2026 بأي صيغة حروف ممكنة"""
    possible_names = ["RH2026.xlsx", "rh2026.xlsx", "Rh2026.xlsx", "RH2026.XLSX"]
    for name in possible_names:
        if os.path.exists(name):
            return name
    # إن لم يجد بالاسم، يبحث عن أي ملف إكسيل يبدأ بـ RH
    rh_files = glob.glob("*RH*.xlsx") + glob.glob("*rh*.xlsx")
    if rh_files:
        return rh_files[0]
    return None

@st.cache_data
def load_data(source):
    try:
        df = pd.read_excel(source)
        # تنظيف فوري للأعمدة لمنع أي KeyError
        df.columns = [str(c).strip() for c in df.columns]
        for col in df.columns:
            df[col] = df[col].fillna('').astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"حدث خطأ أثناء قراءة ملف الإكسيل: {e}")
        return pd.DataFrame()

# تحديد مصدر الملف
detected_file = find_excel_file()
if detected_file:
    df_source = load_data(detected_file)
else:
    st.sidebar.warning("لم يتم العثور على ملف RH2026.xlsx في المستودع.")
    uploaded = st.sidebar.file_uploader("يرجى رفع ملف RH2026.xlsx يدوياً:", type=["xlsx", "xls"])
    if uploaded:
        df_source = load_data(uploaded)
    else:
        df_source = pd.DataFrame()

if df_source.empty:
    st.info("👈 يرجى رفع ملف قاعدة البيانات لبدء التشغيل.")
    st.stop()

# توحيد أسماء الأعمدة البديلة تلقائياً لتفادي KeyError
col_etab = "C_ETAB" if "C_ETAB" in df_source.columns else ("CD_ETAB" if "CD_ETAB" in df_source.columns else df_source.columns[0])
col_sub = "الوحدة المدرسية.1" if "الوحدة المدرسية.1" in df_source.columns else ("الوحدة المدرسية" if "الوحدة المدرسية" in df_source.columns else col_etab)
col_grp = "المؤسسة" if "المؤسسة" in df_source.columns else ("ETAB" if "ETAB" in df_source.columns else col_sub)
col_task = "المهام" if "المهام" in df_source.columns else ("FONCTION" if "FONCTION" in df_source.columns else ("fonct" if "fonct" in df_source.columns else df_source.columns[1]))

if "df" not in st.session_state or len(st.session_state.df) != len(df_source):
    st.session_state.df = df_source.copy()

df = st.session_state.df

# ==================== الشريط الجانبي ====================
with st.sidebar:
    st.markdown("""
        <div style="text-align: center; padding: 10px 0 20px 0;">
            <div style="font-size: 40px; margin-bottom: 5px;">🏛️</div>
            <h3 style="margin:0; color:#1E293B; font-weight:800;">بوابة الموارد البشرية</h3>
            <span style="color:#64748B; font-size:12px;">مديرية التربية الوطنية والتخطيط</span>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.metric("إجمالي الموظفين المسجلين", f"{len(df):,} فرد")
    
    buf_all = io.BytesIO()
    with pd.ExcelWriter(buf_all, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='PERSONNEL')
    st.download_button(
        label="📥 تحميل قاعدة المعطيات كاملة (Excel)",
        data=buf_all.getvalue(),
        file_name="RH2026_Updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🔄 إعادة ضبط وتحديث القاعدة", use_container_width=True):
        if detected_file:
            st.session_state.df = load_data(detected_file)
            st.success("تم تحديث البيانات بنجاح!")
            st.rerun()

# ==================== الترويسة ====================
st.markdown("""
    <div class="portal-header">
        <div>
            <h1 class="portal-title">منظومة تدبير الموارد البشرية وتتبع المؤسسات</h1>
            <div class="portal-sub">لوحة قيادة تفاعلية مخصصة للمصلحة - الموسم الدراسي 2026</div>
        </div>
        <div class="portal-badge">نسخة الويب السحابية v3.6</div>
    </div>
""", unsafe_allow_html=True)

# ==================== التبويبات ====================
tab_etab, tab_search, tab_stats, tab_edit = st.tabs([
    "🏫 استعلام موظفي المؤسسة",
    "🔍 التنقيب والبحث الفردي",
    "📊 لوحة المؤشرات والإحصاء",
    "✏️ تدبير وتعديل سجل موظف"
])

# ---------------- 1. تبويب استعلام المؤسسة ----------------
with tab_etab:
    with st.form("etab_form"):
        st.markdown("##### 📍 تحديد المؤسسة التعليمية أو الوحدة المدرسية")
        col_code, col_select, col_btn = st.columns([1.5, 2.5, 1])
        
        input_code = col_code.text_input("كود المؤسسة أو الفرعية:", placeholder="مثال: 03242Y").strip().upper()
        
        etab_choices = [""] + sorted(list(df[col_etab].replace('', pd.NA).dropna().unique())) if col_etab in df.columns else [""]
        selected_dropdown = col_select.selectbox("أو اختر من الأكواد المسجلة:", etab_choices)
        
        col_btn.write("")
        col_btn.write("")
        btn_search_etab = col_btn.form_submit_button("🔎 تأكيد البحث", use_container_width=True, type="primary")

    chosen_etab = input_code if input_code else selected_dropdown

    if chosen_etab:
        cond1 = df[col_etab].astype(str).str.upper() == chosen_etab
        cond2 = (df["CD_ETAB"].astype(str).str.upper() == chosen_etab) if "CD_ETAB" in df.columns else cond1
        res_etab = df[cond1 | cond2]

        if not res_etab.empty:
            r0 = res_etab.iloc[0]
            nom_sub = r0.get(col_sub, chosen_etab)
            nom_grp = r0.get(col_grp, "")
            
            total_etab = len(res_etab)
            f_count = (res_etab["SEX"].str.upper() == 'F').sum() if "SEX" in res_etab.columns else 0
            m_count = (res_etab["SEX"].str.upper() == 'M').sum() if "SEX" in res_etab.columns else 0
            
            st.markdown(f"""
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
            """, unsafe_allow_html=True)

            col_t1, col_t2 = st.columns([1.2, 2.8])
            with col_t1:
                st.markdown("###### 📊 توزيع العاملين حسب المهمة:")
                if col_task in res_etab.columns:
                    stats_counts = res_etab[col_task].value_counts()
                    stats_etab = pd.DataFrame({
                        "المهمة": stats_counts.index,
                        "العدد": stats_counts.values,
                        "النسبة": (stats_counts.values / total_etab * 100)
                    })
                    stats_etab["النسبة"] = stats_etab["النسبة"].map("{:.1f}%".format)
                    st.dataframe(stats_etab, use_container_width=True, hide_index=True)

            with col_t2:
                st.markdown("###### 📋 اللائحة الاسمية لأطر المؤسسة:")
                show_cols = ["PPR", "CIN", "NOM_AR", "NOM_FR", col_task, "الدرجة", "DATE_DEBUT_AFF", "DATE_ANC_ADM"]
                avail_cols = [c for c in show_cols if c in res_etab.columns]
                st.dataframe(res_etab[avail_cols], use_container_width=True, height=260)

            buf_e = io.BytesIO()
            with pd.ExcelWriter(buf_e, engine='openpyxl') as writer:
                res_etab.to_excel(writer, index=False, sheet_name=chosen_etab[:25])
            st.download_button(
                label=f"📥 تصدير تقرير المؤسسة {chosen_etab} (Excel)",
                data=buf_e.getvalue(),
                file_name=f"RH_{chosen_etab}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning(f"لم يتم العثور على أي موظف مسجل بالكود: `{chosen_etab}`")
    else:
        st.info("👈 المرجو إدخال كود المؤسسة ثم الضغط على **تأكيد البحث**.")

# ---------------- 2. تبويب التنقيب الفردي ----------------
with tab_search:
    with st.form("search_form"):
        st.markdown("##### 🔎 استعلام سريع بالرقم المالي، الاسم أو البطاقة الوطنية")
        q_col1, q_col2 = st.columns([3, 1])
        query_val = q_col1.text_input("أدخل رقم التأجير (PPR)، الاسم، أو رقم البطاقة (CIN):", placeholder="مثال: 379274 أو محمد").strip()
        q_col2.write("")
        q_col2.write("")
        btn_query = q_col2.form_submit_button("🔍 تنفيذ الاستعلام", use_container_width=True, type="primary")

    if query_val:
        q_lower = query_val.lower()
        search_cols = [c for c in ["PPR", "CIN", "NOM_AR", "NOM_FR", col_grp, col_sub] if c in df.columns]
        mask_q = df[search_cols].apply(lambda row: row.astype(str).str.lower().str.contains(q_lower, regex=False).any(), axis=1)
        res_q = df[mask_q]

        if not res_q.empty:
            st.success(f"تم العثور على **{len(res_q)}** نتيجة مطابقة:")
            disp_cols = ["PPR", "CIN", "NOM_AR", "NOM_FR", col_task, "الاطار", "الدرجة", col_sub, col_grp, "DATE_DEBUT_AFF"]
            cols_to_render = [c for c in disp_cols if c in res_q.columns]
            st.dataframe(res_q[cols_to_render], use_container_width=True)
            
            if len(res_q) == 1:
                indiv = res_q.iloc[0]
                with st.container():
                    st.markdown("---")
                    st.markdown(f"#### 👤 بطاقة المسار الإداري: {indiv.get('NOM_AR', '')} ({indiv.get('NOM_FR', '')})")
                    c_b1, c_b2, c_b3 = st.columns(3)
                    c_b1.markdown(f"**رقم التأجير (PPR):** `{indiv.get('PPR', '')}`")
                    c_b1.markdown(f"**رقم البطاقة (CIN):** `{indiv.get('CIN', '')}`")
                    c_b1.markdown(f"**تاريخ الازدياد:** {indiv.get('DATE_NAISSANCE', '')}")
                    
                    c_b2.markdown(f"**المهمة:** {indiv.get(col_task, '')}")
                    c_b2.markdown(f"**الإطار والدرجة:** {indiv.get('الاطار', '')} - {indiv.get('الدرجة', '')}")
                    c_b2.markdown(f"**تاريخ التوظيف:** {indiv.get('DATE_ANC_ADM', '')}")

                    c_b3.markdown(f"**المؤسسة:** {indiv.get(col_grp, '')}")
                    c_b3.markdown(f"**الفرعية:** {indiv.get(col_sub, '')} (`{indiv.get(col_etab, '')}`)")
                    c_b3.markdown(f"**تاريخ التعيين بالمؤسسة:** {indiv.get('DATE_DEBUT_AFF', '')}")
        else:
            st.warning("لم يتم العثور على أي موظف بهذه المعطيات.")

# ---------------- 3. تبويب لوحة الإحصاء ----------------
with tab_stats:
    st.markdown("##### 📈 مؤشرات التوزيع الإجمالي للموارد البشرية بالإقليم")
    
    total_all = len(df)
    females_all = (df["SEX"].str.upper() == 'F').sum() if "SEX" in df.columns else 0
    males_all = (df["SEX"].str.upper() == 'M').sum() if "SEX" in df.columns else 0
    etab_count = df[col_etab].nunique() if col_etab in df.columns else 0

    st.markdown(f"""
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
    """, unsafe_allow_html=True)

    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.markdown("###### 📊 أعلى 10 مهام من حيث عدد الموظفين:")
        if col_task in df.columns:
            top_tasks_counts = df[col_task].value_counts().head(10)
            top_tasks = pd.DataFrame({"المهمة / الإطار": top_tasks_counts.index, "العدد الإجمالي": top_tasks_counts.values})
            st.dataframe(top_tasks, use_container_width=True, hide_index=True)

    with col_g2:
        st.markdown("###### 🏫 أكبر 10 مؤسسات من حيث عدد الأطر:")
        if col_grp in df.columns:
            top_etabs_counts = df[col_grp].value_counts().head(10)
            top_etabs = pd.DataFrame({"المؤسسة التعليمية": top_etabs_counts.index, "عدد الموظفين": top_etabs_counts.values})
            st.dataframe(top_etabs, use_container_width=True, hide_index=True)

# ---------------- 4. تبويب تعديل سجل موظف ----------------
with tab_edit:
    st.markdown("##### ✏️ نموذج تعديل أو تحيين المعطيات الإدارية لموظف")
    
    col_p1, col_p2 = st.columns([2, 1])
    target_ppr = col_p1.text_input("أدخل رقم التأجير (PPR) لتحميل استمارة التعديل:", placeholder="مثال: 379274").strip()
    
    ppr_col = "PPR" if "PPR" in df.columns else df.columns[0]
    if target_ppr:
        match_idx = df.index[df[ppr_col] == target_ppr].tolist()
        if match_idx:
            t_row = df.loc[match_idx[0]]
            st.success(f"تم تحميل سجل: **{t_row.get('NOM_AR', '')}** ({t_row.get('NOM_FR', '')})")
            
            with st.form("edit_employee_form"):
                f1, f2, f3 = st.columns(3)
                u_nom_ar = f1.text_input("الاسم الكامل بالعربية:", t_row.get("NOM_AR", ""))
                u_nom_fr = f2.text_input("الاسم بالفرنسية:", t_row.get("NOM_FR", ""))
                u_cin = f3.text_input("رقم البطاقة الوطنية (CIN):", t_row.get("CIN", ""))
                
                u_task = f1.text_input("المهمة الحالية:", t_row.get(col_task, ""))
                u_cadre = f2.text_input("الإطار الإداري:", t_row.get("الاطار", ""))
                u_grade = f3.text_input("الدرجة:", t_row.get("الدرجة", ""))

                u_etab_code = f1.text_input("كود المؤسسة:", t_row.get(col_etab, ""))
                u_sub_name = f2.text_input("اسم الفرعية:", t_row.get(col_sub, ""))
                u_grp_name = f3.text_input("المجموعة المدرسية:", t_row.get(col_grp, ""))

                btn_commit = st.form_submit_button("💾 حفظ وتثبيت التعديلات في الجلسة", type="primary")
                if btn_commit:
                    if "NOM_AR" in df.columns: df.loc[match_idx[0], "NOM_AR"] = u_nom_ar
                    if "NOM_FR" in df.columns: df.loc[match_idx[0], "NOM_FR"] = u_nom_fr
                    if "CIN" in df.columns: df.loc[match_idx[0], "CIN"] = u_cin
                    if col_task in df.columns: df.loc[match_idx[0], col_task] = u_task
                    if "الاطار" in df.columns: df.loc[match_idx[0], "الاطار"] = u_cadre
                    if "الدرجة" in df.columns: df.loc[match_idx[0], "الدرجة"] = u_grade
                    if col_etab in df.columns: df.loc[match_idx[0], col_etab] = u_etab_code
                    if col_sub in df.columns: df.loc[match_idx[0], col_sub] = u_sub_name
                    if col_grp in df.columns: df.loc[match_idx[0], col_grp] = u_grp_name
                    st.session_state.df = df
                    st.success("✓ تم حفظ التعديلات بنجاح! يمكنك الآن تحميل الملف المحين من القائمة الجانبية.")
                    st.rerun()
        else:
            st.warning("رقم التأجير غير موجود بقاعدة البيانات.")
