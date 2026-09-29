import streamlit as st
import pandas as pd
import io

# 1. إعدادات الصفحة والمظهر
st.set_page_config(
    page_title="منصة تدبير واستعلام الموارد البشرية والمؤسسات",
    page_icon="🏫",
    layout="wide"
)

# 2. نظام مصادقة وتسجيل دخول بسيط وآمن
USERS = {
    "admin": "admin@2026",
    "inspecteur": "insp@2026",
    "directeur": "dir@2026"
}

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False


def login_form():
    st.markdown("<h2 style='text-align: center; color: #1e40af;'>🏫 تسجيل الدخول للمنصة التربوية</h2>",
                unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            username = st.text_input("اسم المستخدم:")
            password = st.text_input("كلمة المرور:", type="password")
            submitted = st.form_submit_button("دخول المنصة 🔐", use_container_width=True)
            if submitted:
                if username in USERS and USERS[username] == password:
                    st.session_state.authenticated = True
                    st.session_state.username = username
                    st.rerun()
                else:
                    st.error("اسم المستخدم أو كلمة المرور غير صحيحة.")


if not st.session_state.authenticated:
    login_form()
    st.stop()

# ----------------- بعد تسجيل الدخول: الواجهة الرئيسية -----------------

# الشريط الجانبي (Sidebar)
with st.sidebar:
    st.markdown(f"**مرحباً بك: `{st.session_state.username}`**")
    if st.button("تسجيل الخروج 🚪"):
        st.session_state.authenticated = False
        st.rerun()

    st.divider()
    st.header("إدارة البيانات")
    uploaded_file = st.file_uploader("رفع ملف Excel جديد (RH):", type=["xlsx", "xls"])


# تحميل البيانات (إما من الملف المرفوع أو ملف افتراضي)
@st.cache_data
def get_data(file_source):
    if file_source is not None:
        df = pd.read_excel(file_source)
    else:
        # مسار الملف الافتراضي إن وجد على السيرفر
        try:
            df = pd.read_excel("RH2026.xlsx")
        except Exception:
            return pd.DataFrame()
    return df.fillna("").astype(str)


df = get_data(uploaded_file)

st.markdown(
    "<h1 style='color: #1e40af; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px;'>🏫 منصة تدبير واستعلام الموارد البشرية والمؤسسات التعليمية</h1>",
    unsafe_allow_html=True)

if df.empty:
    st.info("👋 مرحباً بك! يرجى رفع ملف الإكسيل من القائمة الجانبية (Sidebar) للبدء في الاستعلام.")
    st.stop()

# الكشف عن عمود المؤسسة ETAB
etab_col = None
for col in df.columns:
    if str(col).strip().upper() == "ETAB":
        etab_col = col
        break

# 3. بطاقات المؤشرات السريعة (KPIs)
m1, m2, m3 = st.columns(3)
with m1:
    st.metric(label="👥 إجمالي السجلات", value=f"{len(df):,}")
with m2:
    total_etabs = df[etab_col].replace("", pd.NA).dropna().nunique() if etab_col else "غير متوفر"
    st.metric(label="🏫 عدد المؤسسات التعليمية (ETAB)", value=f"{total_etabs}")
with m3:
    st.metric(label="📂 حالة قاعدة البيانات", value="نشطة ومحدثة ✅")

st.divider()

# 4. التبويبات الرئيسية (البحث والاستعلام / إحصائيات المؤسسات)
tab_search, tab_stats = st.tabs(["🔍 البحث والاستعلام", "📊 إحصائيات المؤسسات (ETAB)"])

# ===== التبويب 1: البحث المتقدم وتصدير النتائج =====
with tab_search:
    col_input, col_etab_filter = st.columns([2, 1])
    with col_input:
        search_query = st.text_input("بحث نصي (الاسم الكامل، رقم التأجير، التخصص...):", placeholder="اكتب للبحث...")

    with col_etab_filter:
        if etab_col:
            etab_list = ["الكل"] + sorted([e for e in df[etab_col].unique() if e])
            selected_etab = st.selectbox("تصفية حسب المؤسسة (ETAB):", etab_list)
        else:
            selected_etab = "الكل"

    # تصفية البيانات
    filtered_df = df.copy()
    if search_query:
        mask = filtered_df.apply(lambda row: row.str.lower().str.contains(search_query.lower(), regex=False).any(),
                                 axis=1)
        filtered_df = filtered_df[mask]

    if selected_etab != "الكل" and etab_col:
        filtered_df = filtered_df[filtered_df[etab_col] == selected_etab]

    # عرض النتائج
    st.markdown(f"**عدد النتائج المطابقة:** `{len(filtered_df):,}`")
    st.dataframe(filtered_df, use_container_width=True, height=450)

    # زر تصدير النتائج إلى Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        filtered_df.to_excel(writer, index=False, sheet_name="Résultats")

    st.download_button(
        label="💾 تصدير النتائج المفلترة إلى Excel",
        data=output.getvalue(),
        file_name="resultats_recherche.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ===== التبويب 2: إحصائيات وتوزيع الموارد البشرية =====
with tab_stats:
    if etab_col:
        st.subheader("📊 توزيع الأساتذة والموظفين حسب المؤسسة")

        stats_df = df[etab_col].value_counts().reset_index()
        stats_df.columns = ["المؤسسة (ETAB)", "عدد العاملين"]

        col_tbl, col_chart = st.columns([1, 1])
        with col_tbl:
            st.dataframe(stats_df, use_container_width=True, height=400)

            # تصدير الإحصائيات
            stats_output = io.BytesIO()
            with pd.ExcelWriter(stats_output, engine="openpyxl") as writer:
                stats_df.to_excel(writer, index=False, sheet_name="Statistiques_ETAB")

            st.download_button(
                label="📑 تصدير تقرير المؤسسات إلى Excel",
                data=stats_output.getvalue(),
                file_name="statistiques_etab.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col_chart:
            st.write("**أكبر 10 مؤسسات من حيث عدد الأطر:**")
            st.bar_chart(stats_df.set_index("المؤسسة (ETAB)").head(10))
    else:
        st.warning("لم يتم العثور على عمود باسم 'ETAB' في الملف لحساب الإحصائيات.")