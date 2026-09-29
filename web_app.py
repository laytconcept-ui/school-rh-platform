import streamlit as st
import pandas as pd
import io
import os

# 1. إعداد الصفحة
st.set_page_config(
    page_title="المنصة الرقمية لتدبير واستعلام الموارد البشرية",
    page_icon="🇲🇦",
    layout="wide"
)

# 2. تخصيص التصميم، الألوان، الخلفية، وإخفاء أشرطة التحكم
custom_css = """
<style>
/* إخفاء شريط أدوات وأزرار التحكم الرسمية لـ Streamlit */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
.stDeployButton {display:none;}
[data-testid="stToolbar"] {visibility: hidden; display: none;}
[data-testid="stDecoration"] {display: none;}
[data-testid="stStatusWidget"] {visibility: hidden;}

/* خلفية الصفحة: تدرج أزرق تعليمي احترافي مع صورة هندسية أنيقة وخفيفة */
.stApp {
    background-image: linear-gradient(rgba(240, 244, 248, 0.94), rgba(240, 244, 248, 0.94)),
                      url("background.jpg");
    background-size: cover;
    background-position: center;
    background-attachment: fixed;
    direction: rtl;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
}

/* تنسيق بطاقة تسجيل الدخول وحاويات العرض */
[data-testid="stForm"], .css-card {
    background: rgba(255, 255, 255, 0.95);
    border-radius: 16px;
    padding: 28px;
    box-shadow: 0 10px 25px rgba(0, 35, 75, 0.08);
    border: 1px solid rgba(226, 232, 240, 0.8);
}

/* بطاقات المؤشرات (Metrics) */
[data-testid="metric-container"] {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    padding: 14px 20px;
    border-radius: 12px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.03);
    border-right: 5px solid #1e40af;
}

/* تحسين شكل الأزرار */
.stButton > button, .stDownloadButton > button {
    border-radius: 8px;
    font-weight: 600;
    transition: all 0.2s ease-in-out;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(30, 64, 175, 0.2);
}

/* محاذاة الجداول وحقول الإدخال للغة العربية */
input, select {
    text-align: right !important;
}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# 3. إدارة المستخدمين والمصادقة
USERS = {
    "admin": "admin@2026",
    "inspecteur": "insp@2026",
    "directeur": "dir@2026"
}

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def login_form():
    col1, col2, col3 = st.columns([1, 1.8, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown(
            """
            <div style='text-align: center; margin-bottom: 20px;'>
                <div style='font-size: 55px;'>🏛️</div>
                <h2 style='color: #1e3a8a; margin-bottom: 5px;'>فضاء تدبير الموارد البشرية</h2>
                <p style='color: #64748b; font-size: 14px;'>المنصة الرقمية للاستعلام وتتبع الوضعيات الإدارية والمؤسسات</p>
            </div>
            """, 
            unsafe_allow_html=True
        )
        with st.form("login_form"):
            username = st.text_input("👤 اسم المستخدم:", placeholder="أدخل اسم الحساب")
            password = st.text_input("🔒 كلمة المرور:", type="password", placeholder="••••••••")
            submitted = st.form_submit_button("تسجيل الدخول إلى المنصة 🔐", use_container_width=True)
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

# ----------------- الواجهة الرئيسية بعد تسجيل الدخول -----------------

# القائمة الجانبية (Sidebar)
with st.sidebar:
    st.markdown(f"### مرحباً بك 👋\n**`{st.session_state.username}`**")
    if st.button("تسجيل الخروج 🚪", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()
    
    st.divider()
    st.subheader("⚙️ إدارة ملف المعطيات")
    uploaded_file = st.file_uploader("رفع نسخة محدثة (Excel):", type=["xlsx", "xls"])

# دالة قراءة البيانات التلقائية
@st.cache_data
def get_data(file_source):
    if file_source is not None:
        return pd.read_excel(file_source).fillna("").astype(str)
    
    for fname in os.listdir("."):
        if fname.lower().endswith((".xlsx", ".xls")) and not fname.startswith("~$"):
            try:
                df = pd.read_excel(fname)
                return df.fillna("").astype(str)
            except Exception as e:
                st.error(f"خطأ في قراءة ملف {fname}: {e}")
    return pd.DataFrame()

df = get_data(uploaded_file)

# الترويسة الرئيسية
st.markdown(
    """
    <div style='background: white; border-radius: 12px; padding: 18px 24px; margin-bottom: 20px; box-shadow: 0 4px 14px rgba(0,0,0,0.04); border-right: 6px solid #1e40af;'>
        <h2 style='color: #1e3a8a; margin: 0; font-size: 24px;'>🏫 منصة استعلام وتدبير الموارد البشرية والمؤسسات التعليمية</h2>
        <p style='color: #64748b; margin: 5px 0 0 0; font-size: 14px;'>قاعدة بيانات الأطر الإدارية والتربوية وتوزيع الهيئات التعليمية</p>
    </div>
    """,
    unsafe_allow_html=True
)

if df.empty:
    st.info("👋 يرجى التأكد من رفع ملف الإكسيل في المستودع ليتم تحميل المعطيات تلقائياً.")
    st.stop()

# التحقق من عمود المؤسسة (ETAB)
etab_col = None
for col in df.columns:
    if str(col).strip().upper() == "ETAB":
        etab_col = col
        break

# عرض المؤشرات السريعة (KPIs)
m1, m2, m3 = st.columns(3)
with m1:
    st.metric(label="👥 إجمالي الموظفين المسجلين", value=f"{len(df):,}")
with m2:
    total_etabs = df[etab_col].replace("", pd.NA).dropna().nunique() if etab_col else "غير محدد"
    st.metric(label="🏫 عدد المؤسسات التعليمية (ETAB)", value=f"{total_etabs}")
with m3:
    st.metric(label="📂 وضعية المنظومة", value="متصلة ومحدثة ✅")

st.markdown("<br>", unsafe_allow_html=True)

# التبويبات
tab_search, tab_stats = st.tabs(["🔍 البحث والاستعلام المتقدم", "📊 إحصائيات وتوزيع المؤسسات"])

# التبويب 1: البحث
with tab_search:
    col_input, col_etab_filter = st.columns([2, 1])
    with col_input:
        search_query = st.text_input("🔎 بحث سريع (الاسم، رقم التأجير، التخصص، المهمة...):", placeholder="اكتب كلمة البحث هنا...")
    
    with col_etab_filter:
        if etab_col:
            etab_list = ["الكل"] + sorted([e for e in df[etab_col].unique() if e])
            selected_etab = st.selectbox("🏫 تصفية حسب المؤسسة:", etab_list)
        else:
            selected_etab = "الكل"

    filtered_df = df.copy()
    if search_query:
        mask = filtered_df.apply(lambda row: row.str.lower().str.contains(search_query.lower(), regex=False).any(), axis=1)
        filtered_df = filtered_df[mask]

    if selected_etab != "الكل" and etab_col:
        filtered_df = filtered_df[filtered_df[etab_col] == selected_etab]

    st.markdown(f"**عدد السجلات المطابقة:** `{len(filtered_df):,}`")
    st.dataframe(filtered_df, use_container_width=True, height=450)

    # زر التحميل بتنسيق أنيق
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        filtered_df.to_excel(writer, index=False, sheet_name="Résultats")
    
    st.download_button(
        label="📥 استخراج النتائج المصفاة (Excel)",
        data=output.getvalue(),
        file_name="resultats_recherche.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# التبويب 2: الإحصائيات
with tab_stats:
    if etab_col:
        st.subheader("📊 توزيع الموارد البشرية حسب المؤسسة التعليمية")
        
        stats_df = df[etab_col].value_counts().reset_index()
        stats_df.columns = ["المؤسسة (ETAB)", "عدد العاملين"]

        c1, c2 = st.columns([1, 1])
        with c1:
            st.dataframe(stats_df, use_container_width=True, height=400)
            
            stats_output = io.BytesIO()
            with pd.ExcelWriter(stats_output, engine="openpyxl") as writer:
                stats_df.to_excel(writer, index=False, sheet_name="Statistiques_ETAB")
            
            st.download_button(
                label="📑 تصدير تقرير المؤسسات (Excel)",
                data=stats_output.getvalue(),
                file_name="statistiques_etab.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with c2:
            st.write("**أكبر 10 مؤسسات من حيث عدد الأطر:**")
            st.bar_chart(stats_df.set_index("المؤسسة (ETAB)").head(10))
    else:
        st.warning("عمود المؤسسة 'ETAB' غير محدد في قاعدة البيانات.")
