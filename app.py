import streamlit as st
import pandas as pd
from docx import Document
import tempfile
import os

st.set_page_config(page_title="كيان أحمد سمير حجر - النظام الضريبي المتكامل", layout="wide")

DATA_DIR = "data_files"
os.makedirs(DATA_DIR, exist_ok=True)

DB_PATH = os.path.join(DATA_DIR, "real_estate_db.xlsx")
PROFIT_PATH = os.path.join(DATA_DIR, "profit_db.xlsx")
NAT_PATH = os.path.join(DATA_DIR, "national_db.xlsx")
TP_PATH = os.path.join(DATA_DIR, "taxpayer_db.xlsx")
FS_PATH = os.path.join(DATA_DIR, "file_status_db.xlsx")
CENTRAL_PATH = os.path.join(DATA_DIR, "central_query_db.xlsx")
MEMO_DB_PATH = os.path.join(DATA_DIR, "memo_taxpayers_db.xlsx")

# تهيئة قواعد البيانات في الجلسة
if "db_sheets" not in st.session_state:
    st.session_state.db_sheets = {}
    if os.path.exists(DB_PATH):
        try:
            df_temp = pd.read_excel(DB_PATH)
            df_temp.columns = df_temp.columns.str.strip()
            st.session_state.db_sheets["شيت_افتراضي.xlsx"] = df_temp.to_dict('records')
        except:
            pass

if "central_db" not in st.session_state:
    if os.path.exists(CENTRAL_PATH):
        try:
            df_temp = pd.read_excel(CENTRAL_PATH)
            df_temp.columns = df_temp.columns.str.strip()
            st.session_state.central_db = df_temp.to_dict('records')
        except:
            st.session_state.central_db = []
    else:
        st.session_state.central_db = []

if "tax_records" not in st.session_state:
    st.session_state.tax_records = []

if "national_db" not in st.session_state:
    if os.path.exists(NAT_PATH):
        df_temp = pd.read_excel(NAT_PATH)
        df_temp.columns = df_temp.columns.str.strip()
        st.session_state.national_db = df_temp.to_dict('records')
    else:
        st.session_state.national_db = []

if "taxpayer_db" not in st.session_state:
    if os.path.exists(TP_PATH):
        df_temp = pd.read_excel(TP_PATH)
        df_temp.columns = df_temp.columns.str.strip()
        st.session_state.taxpayer_db = df_temp.to_dict('records')
    else:
        st.session_state.taxpayer_db = []

if "profit_db" not in st.session_state:
    if os.path.exists(PROFIT_PATH):
        df_temp = pd.read_excel(PROFIT_PATH)
        df_temp.columns = df_temp.columns.str.strip()
        st.session_state.profit_db = df_temp
    else:
        st.session_state.profit_db = None

if "file_status_db" not in st.session_state:
    if os.path.exists(FS_PATH):
        df_temp = pd.read_excel(FS_PATH)
        df_temp.columns = df_temp.columns.str.strip()
        st.session_state.file_status_db = df_temp.to_dict('records')
    else:
        st.session_state.file_status_db = []

if "memo_taxpayer_db" not in st.session_state:
    if os.path.exists(MEMO_DB_PATH):
        try:
            df_temp = pd.read_excel(MEMO_DB_PATH)
            df_temp.columns = df_temp.columns.str.strip()
            st.session_state.memo_taxpayer_db = df_temp.to_dict('records')
        except:
            st.session_state.memo_taxpayer_db = []
    else:
        st.session_state.memo_taxpayer_db = []

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# شاشة تسجيل الدخول (اسم المستخدم وكلمة المرور: 04842040 مع إخفاء الباسورد)
if not st.session_state.logged_in:
    st.title("🔐 تسجيل الدخول - كيان أحمد سمير حجر")
    st.image("https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?q=80&w=1000&auto=format&fit=crop", use_container_width=True)
    u = st.text_input("اسم المستخدم", placeholder="أدخل اسم المستخدم")
    p = st.text_input("الرقم السري", type="password", placeholder="أدخل الرقم السري")
    if st.button("دخول"):
        if u == "04842040" and p == "04842040":
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("اسم المستخدم أو الرقم السري غير صحيح (يرجى إدخال 04842040)")
else:
    st.sidebar.title("📌 قائمة النظام الضريبي")
    st.sidebar.markdown("### كيان أحمد سمير حجر")
    st.sidebar.image("https://images.unsplash.com/photo-1541888946425-d0fbb18f124b?q=80&w=400&auto=format&fit=crop", use_container_width=True)
    
    app_mode = st.sidebar.selectbox("اختر الشاشة المطلوبة:", [
        "🏠 الرئيسية (الاستعلام المركزي)",
        "🔎 الاستعلام عن سعر المتر (حالات المثل)",
        "🧾 حساب الضريبة ٢.٥٪",
        "📚 سجلات الضريبة",
        "🪪 البحث بالرقم القومي",
        "📊 تعاملات الممولين",
        "📝 مذكرة الفحص الاحترافية",
        "📈 نسب صافي الربح",
        "📁 موقف الملف",
        "🚪 تسجيل الخروج"
    ])

    if app_mode == "🚪 تسجيل الخروج":
        st.session_state.logged_in = False
        st.rerun()

    # ==========================================
    # 1. الرئيسية (الاستعلام المركزي)
    # ==========================================
    elif app_mode == "🏠 الرئيسية (الاستعلام المركزي)":
        st.title("🏠 كيان أحمد سمير حجر - النظام الضريبي المتكامل")
        st.image("https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?q=80&w=1200&auto=format&fit=crop", use_container_width=True)
        st.markdown("---")

        st.subheader("📂 1. رفع وحفظ شيت الإكسيل المركزي")
        central_file = st.file_uploader("اختر شيت الإكسيل (الاسم، رقم الملف، الرقم القومي)", type=["xlsx", "xls", "csv"], key="cent_up")
        
        if central_file is not None:
            if st.button("💾 حفظ الشيت في النظام بصفة دائمة", key="save_cent"):
                try:
                    df_c = pd.read_excel(central_file) if not central_file.name.endswith(".csv") else pd.read_csv(central_file)
                    df_c.columns = df_c.columns.str.strip()
                    st.session_state.central_db = df_c.to_dict('records')
                    df_c.to_excel(CENTRAL_PATH, index=False)
                    st.success("تم رفع وحفظ شيت الاستعلام المركزي بنجاح دائم!")
                except Exception as e:
                    st.error(f"خطأ أثناء قراءة الملف: {e}")

        st.markdown("---")
        with st.expander("👁 2. معاينة شيت الاستعلام المركزي المرفوع بالكامل", expanded=False):
            if st.session_state.central_db:
                st.dataframe(pd.DataFrame(st.session_state.central_db), use_container_width=True)
                st.write(f"إجمالي عدد السجلات: {len(st.session_state.central_db)}")
            else:
                st.info("لا توجد بيانات مركزية مرفوعة حتى الآن.")

        st.markdown("---")
        st.subheader("🔍 3. الاستعلام المركزي (بالاسم أو الرقم القومي)")
        
        if st.session_state.central_db:
            df_curr = pd.DataFrame(st.session_state.central_db)
            cols = list(df_curr.columns)
            
            def_name = next((c for c in cols if 'اسم' in str(c) or 'name' in str(c).lower()), cols[0])
            def_file = next((c for c in cols if 'ملف' in str(c) or 'file' in str(c).lower()), cols[1] if len(cols)>1 else cols[0])
            def_nid = next((c for c in cols if 'قومي' in str(c) or 'id' in str(c).lower() or 'national' in str(c).lower()), cols[2] if len(cols)>2 else cols[0])

            sc1, sc2, sc3 = st.columns(3)
            with sc1:
                col_name_sel = st.selectbox("عمود اسم الممول:", cols, index=cols.index(def_name) if def_name in cols else 0)
            with sc2:
                col_file_sel = st.selectbox("عمود رقم الملف:", cols, index=cols.index(def_file) if def_file in cols else 0)
            with sc3:
                col_nid_sel = st.selectbox("عمود الرقم القومي:", cols, index=cols.index(def_nid) if def_nid in cols else 0)

            st.markdown("---")
            c_inp1, c_inp2 = st.columns(2)
            with c_inp1:
                search_name = st.text_input("بحث بالاسم (أو جزء منه):", placeholder="مثال: محمد")
            with c_inp2:
                search_nid = st.text_input("بحث بالرقم القومي:", placeholder="أدخل الرقم القومي كاملاً أو جزئياً")

            if st.button("🔎 تنفيذ الاستعلام المركزي", key="run_cent"):
                results = []
                for item in st.session_state.central_db:
                    val_name = str(item.get(col_name_sel, ""))
                    val_nid = str(item.get(col_nid_sel, ""))
                    
                    match_n = search_name == "" or search_name.lower() in val_name.lower()
                    match_id = search_nid == "" or search_nid in val_nid
                    
                    if (search_name != "" or search_nid != "") and match_n and match_id:
                        results.append(item)
                
                if results:
                    st.success(f"✅ تم العثور على ({len(results)}) نتيجة مطابقة:")
                    for res in results:
                        r_name = res.get(col_name_sel, "غير متوفر")
                        r_file = res.get(col_file_sel, "غير متوفر")
                        r_nid = res.get(col_nid_sel, "غير متوفر")
                        
                        st.markdown(f"""
                        <div style="background-color: #f2f8f8; border: 2px solid #075d68; border-radius: 12px; padding: 20px; margin-bottom: 15px;">
                            <h3 style="color: #075d68; margin-top: 0; border-bottom: 1px solid #075d68; padding-bottom: 8px;">📑 تقرير بيانات الممول</h3>
                            <p style="font-size: 18px; margin: 8px 0;"><b>👤 اسم الممول:</b> {r_name}</p>
                            <p style="font-size: 18px; margin: 8px 0;"><b>🪪 الرقم القومي:</b> {r_nid}</p>
                            <p style="font-size: 18px; margin: 8px 0;"><b>📁 رقم الملف:</b> {r_file}</p>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.warning("⚠️ لا توجد نتائج مطابقة لبيانات البحث المدخلة.")
        else:
            st.info("الرجاء رفع شيت الإكسيل المركزي أولاً.")

    # ==========================================
    # 2. الاستعلام عن سعر المتر (حالات المثل)
    # ==========================================
    elif app_mode == "🔎 الاستعلام عن سعر المتر (حالات المثل)":
        st.title("🔎 الاستعلام عن سعر المتر العقاري (حالات المثل)")
        st.image("https://images.unsplash.com/photo-1545324418-cc1a3fa10c00?q=80&w=1200&auto=format&fit=crop", use_container_width=True)
        st.markdown("---")
        
        uploaded_files = st.file_uploader("اختر شيت أو عدة شيتات إكسيل لحالات المثل", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="db_up_multi")
        
        if uploaded_files:
            if st.button("💾 حفظ وإضافة جميع الشيتات المرفوعة للنظام", key="save_multi_db"):
                try:
                    for up_f in uploaded_files:
                        df_u = pd.read_excel(up_f) if not up_f.name.endswith(".csv") else pd.read_csv(up_f)
                        if 'اسم الشارع' not in df_u.columns and len(df_u) > 1:
                            for idx, row in df_u.iterrows():
                                row_str = str(row.values)
                                if 'اسم الشارع' in row_str or 'ق. الاصلية' in row_str or 'المساحة' in row_str:
                                    df_u.columns = df_u.iloc[idx]
                                    df_u = df_u.iloc[idx+1:].reset_index(drop=True)
                                    break
                        df_u.columns = df_u.columns.str.strip()
                        st.session_state.db_sheets[up_f.name] = df_u.to_dict('records')
                    st.success("✅ تم رفع وحفظ الشيتات الجديدة بنجاح!")
                except Exception as e:
                    st.error(f"خطأ أثناء معالجة الملفات: {e}")

        st.markdown("---")

        if st.session_state.db_sheets:
            sheet_names_list = ["الكل (دمج كل الشيتات)"] + list(st.session_state.db_sheets.keys())
            selected_sheet = st.selectbox("📂 حدد الشيت المراد البحث فيه:", sheet_names_list)
            
            active_records = []
            if selected_sheet == "الكل (دمج كل الشيتات)":
                for s_name, recs in st.session_state.db_sheets.items():
                    active_records.extend(recs)
            else:
                active_records = st.session_state.db_sheets.get(selected_sheet, [])

            if active_records:
                df_view = pd.DataFrame(active_records)
                df_view.columns = df_view.columns.str.strip()
                cols_db = list(df_view.columns)
                
                with st.expander(f"👁️ معاينة بيانات الشيت النشط ({selected_sheet})", expanded=False):
                    st.dataframe(df_view, use_container_width=True)
                    st.write(f"إجمالي عدد الصفوف في العرض الحالي: {len(active_records)}")
                
                st.subheader("🔗 مطابقة أعمدة الشيت الحالي (العنوان، القيمة، المساحة، التاريخ، الحصة، الكود، والاختيار)")
                
                def find_col(keywords):
                    for col in cols_db:
                        for kw in keywords:
                            if kw in str(col).lower():
                                return col
                    return cols_db[0] if cols_db else ""

                def_street = find_col(['اسم الشارع', 'عنوان', 'شارع', 'وصف', 'مكان', 'address', 'street'])
                def_value = find_col(['ق. الاصلية', 'قيمة', 'ثمن', 'مبلغ', 'العقد', 'value', 'price'])
                def_est = find_col(['ق. التقديرية', 'تقديرية', 'estimated'])
                def_area = find_col(['مساحة', 'المساحة', 'area', 'space'])
                def_year = find_col(['التاريخ', 'سنة', 'عام', 'year', 'date'])
                def_share = find_col(['الحصة', 'share'])
                def_choice = find_col(['اختيار', 'choice'])
                def_code = find_col(['كود', 'code'])

                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    col_street = st.selectbox("عمود العنوان (اسم الشارع):", cols_db, index=cols_db.index(def_street) if def_street in cols_db else 0)
                    col_share = st.selectbox("عمود الحصة:", cols_db, index=cols_db.index(def_share) if def_share in cols_db else 0)
                with c2:
                    col_value = st.selectbox("عمود قيمة العقد الأصلية:", cols_db, index=cols_db.index(def_value) if def_value in cols_db else 0)
                    col_choice = st.selectbox("عمود الاختيار:", cols_db, index=cols_db.index(def_choice) if def_choice in cols_db else 0)
                with c3:
                    col_est = st.selectbox("عمود القيمة التقديرية:", cols_db, index=cols_db.index(def_est) if def_est in cols_db else 0)
                    col_code = st.selectbox("عمود الكود:", cols_db, index=cols_db.index(def_code) if def_code in cols_db else 0)
                with c4:
                    col_area = st.selectbox("عمود المساحة:", cols_db, index=cols_db.index(def_area) if def_area in cols_db else 0)
                    col_year = st.selectbox("عمود التاريخ:", cols_db, index=cols_db.index(def_year) if def_year in cols_db else 0)

                st.markdown("---")
                sc1, sc2, sc3 = st.columns([2, 1, 1])
                with sc1:
                    q = st.text_input("جزء من العنوان أو الشارع", placeholder="سموحة، سيدي جابر...")
                with sc2:
                    all_years = sorted(list(set(str(item.get(col_year, "الكل")) for item in active_records if pd.notna(item.get(col_year)))))
                    selected_year = st.selectbox("سنة البيع / التاريخ", ["كل السنوات"] + all_years)
                with sc3:
                    max_results = st.selectbox("عدد النتائج", [10, 20, 50, 100, 500], index=1)

                filtered = []
                for item in active_records:
                    street_val = str(item.get(col_street, ""))
                    year_val = str(item.get(col_year, ""))
                    
                    match_q = q == "" or q.lower() in street_val.lower()
                    match_year = selected_year == "كل السنوات" or selected_year in year_val
                    
                    if match_q and match_year:
                        filtered.append(item)

                filtered = filtered[:max_results]

                if filtered:
                    st.markdown("---")
                    total_contract_value = 0.0
                    total_area_value = 0.0
                    min_m_price = float('inf')
                    max_m_price = 0.0

                    for row in filtered:
                        try:
                            raw_v = str(row.get(col_value, 0))
                            v = float(raw_v.replace(',', '').replace('None', '0').replace('nan', '0').strip() or 0)
                        except:
                            v = 0.0
                        try:
                            raw_a = str(row.get(col_area, 0))
                            a = float(raw_a.replace(',', '').replace('None', '0').replace('nan', '0').strip() or 0)
                        except:
                            a = 0.0

                        if a > 0:
                            total_contract_value += v
                            total_area_value += a
                            m_price = v / a
                            if m_price < min_m_price:
                                min_m_price = m_price
                            if m_price > max_m_price:
                                max_m_price = m_price

                    avg_price = (total_contract_value / total_area_value) if total_area_value > 0 else 0.0
                    if min_m_price == float('inf'):
                        min_m_price = 0.0

                    st.markdown(f"""
                    <div style="background-color: #fffcf2; border: 2px solid #d5a63a; border-radius: 15px; padding: 25px; text-align: center; margin: 20px 0;">
                        <p style="color: #6c757d; margin: 0; font-size: 16px;">متوسط سعر المتر (كيان أحمد سمير حجر)</p>
                        <h1 style="color: #075d68; font-size: 48px; font-weight: bold; margin: 10px 0;">{avg_price:,.2f} جنيه</h1>
                    </div>
                    """, unsafe_allow_html=True)

                    m1, m2, m3, m4 = st.columns(4)
                    with m1:
                        st.metric("التصرفات المطابقة", len(filtered))
                    with m2:
                        st.metric("أقل سعر متر", f"{min_m_price:,.2f}")
                    with m3:
                        st.metric("أعلى سعر متر", f"{max_m_price:,.2f}")
                    with m4:
                        st.metric("إجمالي العقود", f"{total_contract_value:,.2f}")

                    st.markdown("<br>", unsafe_allow_html=True)

                    if st.button("🖨 طباعة تقرير حالات المثل (صلاحيات الطباعة - ملف Word احترافي)", key="btn_word_market"):
                        doc = Document()
                        doc.add_heading("كيان أحمد سمير حجر - تقرير حالات المثل العقارية", 0)
                        doc.add_paragraph(f"متوسط سعر المتر: {avg_price:,.2f} جنيه")
                        doc.add_paragraph(f"إجمالي عدد الحالات المطابقة: {len(filtered)}")
                        doc.add_heading("تفاصيل العقود المطابقة:", level=1)
                        
                        for r_item in filtered:
                            doc.add_paragraph(f"العنوان: {r_item.get(col_street, '')} | القيمة: {r_item.get(col_value, '')} | المساحة: {r_item.get(col_area, '')} | الحصة: {r_item.get(col_share, '')} | الكود: {r_item.get(col_code, '')}")
                        
                        tmp_f = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
                        doc.save(tmp_f.name)
                        with open(tmp_f.name, "rb") as f:
                            st.download_button("💾 اضغط هنا لتحميل تقرير حالات المثل بصيغة Word", f, file_name="Real_Estate_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

                    st.markdown("<br>", unsafe_allow_html=True)
                    table_data = []
                    for item in filtered:
                        try:
                            v = float(str(item.get(col_value, 0)).replace(',', '').replace('None', '0').strip() or 0)
                        except:
                            v = 0.0
                        try:
                            a = float(str(item.get(col_area, 0)).replace(',', '').replace('None', '0').strip() or 0)
                        except:
                            a = 0.0
                        m_p = (v / a) if a > 0 else 0.0
                        table_data.append({
                            "سعر المتر": f"{m_p:,.2f}",
                            "قيمة العقد": f"{v:,.2f}",
                            "المساحة": f"{a:,.2f}",
                            "التاريخ": str(item.get(col_year, "")),
                            "الحصة": str(item.get(col_share, "")),
                            "العنوان": str(item.get(col_street, ""))
                        })
                    st.dataframe(pd.DataFrame(table_data), use_container_width=True, hide_index=True)
                else:
                    st.warning("⚠️ لا توجد نتائج مطابقة لعملية البحث المدخلة.")
            else:
                st.info("الشيت المحدد فارغ حالياً.")
        else:
            st.info("الرجاء رفع شيت واحد أو عدة شيتات لحالات المثل العقارية لبدء العمل.")

    # ==========================================
    # 3. حساب الضريبة ٢.٥٪
    # ==========================================
    elif app_mode == "🧾 حساب الضريبة ٢.٥٪":
        st.title("🧾 حساب الضريبة بنسبة ٢.٥٪")
        st.markdown("---")
        c1, c2 = st.columns(2)
        with c1:
            seller = st.text_input("اسم البائع")
            tax_date = st.date_input("تاريخ البيع")
            property_desc = st.text_area("وصف المبيع")
        with c2:
            buyer = st.text_input("اسم المشتري")
            property_address = st.text_input("عنوان المبيع")
            sale_amount = st.number_input("مبلغ البيع (جنيه)", min_value=0.0, step=0.01)

        tax_val = sale_amount * 0.025
        st.markdown(f"""
        <div style="background-color: #fff8e6; border: 1px solid #d5a63a; border-radius: 10px; padding: 15px; text-align: center; margin-top: 15px;">
            <p style="margin:0; font-weight:bold;">قيمة الضريبة المستحقة بنسبة ٢.٥٪</p>
            <h2 style="color: #075d68; margin: 5px 0;">{tax_val:,.2f} جنيه</h2>
        </div>
        """, unsafe_allow_html=True)

        if st.button("💾 حفظ سجل الضريبة"):
            st.session_state.tax_records.append({
                "seller": seller, "buyer": buyer, "date": str(tax_date),
                "address": property_address, "desc": property_desc,
                "amount": sale_amount, "tax": tax_val
            })
            st.success("تم حفظ سجل الضريبة بنجاح!")

    elif app_mode == "📚 سجلات الضريبة":
        st.title("📚 سجلات حساب الضريبة")
        st.markdown("---")
        if st.session_state.tax_records:
            df_tax = pd.DataFrame(st.session_state.tax_records)
            df_tax.columns = ['البائع', 'المشتري', 'تاريخ البيع', 'العنوان', 'الوصف', 'مبلغ البيع', 'الضريبة ٢.٥٪']
            st.dataframe(df_tax, use_container_width=True, hide_index=True)
            
            if st.button("🖨 طباعة سجلات الضريبة (صلاحيات الطباعة - ملف Word)", key="print_tax_records"):
                doc = Document()
                doc.add_heading("كيان أحمد سمير حجر - سجلات حساب الضريبة ٢.٥٪", 0)
                for rec in st.session_state.tax_records:
                    doc.add_paragraph(f"البائع: {rec['seller']} | المشتري: {rec['buyer']} | المبلغ: {rec['amount']} | الضريبة: {rec['tax']}")
                tmp_f = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
                doc.save(tmp_f.name)
                with open(tmp_f.name, "rb") as f:
                    st.download_button("💾 تحميل ملف السجلات بصيغة Word", f, file_name="Tax_Records.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        else:
            st.info("لا توجد سجلات ضريبية محفوظة حتى الآن.")

    elif app_mode == "🪪 البحث بالرقم القومي":
        st.title("🪪 البحث بالرقم القومي للممولين")
        st.markdown("---")
        nat_file = st.file_uploader("رفع شيت Excel لقاعدة الممولين", type=["xlsx", "xls", "csv"], key="nat_up")
        if nat_file is not None:
            try:
                df_nat = pd.read_excel(nat_file) if not nat_file.name.endswith(".csv") else pd.read_csv(nat_file)
                st.session_state.national_db = df_nat.to_dict('records')
                df_nat.to_excel(NAT_PATH, index=False)
                st.success("تم رفع وحفظ قاعدة الممولين بنجاح!")
            except Exception as e:
                st.error(f"خطأ: {e}")

        nid_query = st.text_input("أدخل الرقم القومي (14 رقماً)", max_chars=14)
        if nid_query and len(nid_query) == 14:
            found = next((item for item in st.session_state.national_db if str(item.get('national_id', item.get('الرقم_القومي', ''))) == nid_query), None)
            if found:
                st.success("تم العثور على بيانات الممول:")
                st.json(found)
                if st.button("🖨 طباعة نتيجة البحث بالرقم القومي (صلاحيات الطباعة - ملف Word)", key="print_nat_res"):
                    doc = Document()
                    doc.add_heading("كيان أحمد سمير حجر - نتيجة الاستعلام بالرقم القومي", 0)
                    for k, v in found.items():
                        doc.add_paragraph(f"{k}: {v}")
                    tmp_f = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
                    doc.save(tmp_f.name)
                    with open(tmp_f.name, "rb") as f:
                        st.download_button("💾 تحميل نتيجة البحث بصيغة Word", f, file_name="National_ID_Result.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            else:
                st.warning("لم يتم العثور على ممول بهذا الرقم القومي.")

    elif app_mode == "📊 تعاملات الممولين":
        st.title("📊 استعلام تعاملات الممولين")
        st.markdown("---")
        tp_file = st.file_uploader("رفع شيت تعاملات الممولين", type=["xlsx", "xls", "csv"], key="tp_up")
        if tp_file is not None:
            try:
                df_tp = pd.read_excel(tp_file) if not tp_file.name.endswith(".csv") else pd.read_csv(tp_file)
                st.session_state.taxpayer_db = df_tp.to_dict('records')
                df_tp.to_excel(TP_PATH, index=False)
                st.success("تم رفع وحفظ شيت تعاملات الممولين بنجاح!")
            except Exception as e:
                st.error(f"خطأ: {e}")

        reg_no = st.text_input("رقم التسجيل الضريبي")
        if reg_no and st.session_state.taxpayer_db:
            results = [item for item in st.session_state.taxpayer_db if str(item.get('reg_no', item.get('رقم_التسجيل', ''))) == reg_no]
            if results:
                st.dataframe(pd.DataFrame(results), use_container_width=True)
                if st.button("🖨 طباعة تعاملات الممول (صلاحيات الطباعة - ملف Word)", key="print_tp_res"):
                    doc = Document()
                    doc.add_heading(f"كيان أحمد سمير حجر - تعاملات الممول برقم تسجيل: {reg_no}", 0)
                    for r in results:
                        doc.add_paragraph(str(r))
                    tmp_f = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
                    doc.save(tmp_f.name)
                    with open(tmp_f.name, "rb") as f:
                        st.download_button("💾 تحميل تقرير التعاملات بصيغة Word", f, file_name=f"Taxpayer_Transactions_{reg_no}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            else:
                st.warning("لا توجد تعاملات لهذا الرقم الضريبي.")

    # ==========================================
    # 7. مذكرة الفحص الاحترافية المطورة بالكامل
    # ==========================================
    elif app_mode == "📝 مذكرة الفحص الاحترافية":
        st.title("📝 مذكرة الفحص الضريبي الاحترافية")
        st.markdown("إدارة ومعاينة شيت ممولي مذكرة الفحص والملء التلقائي بناءً على رقم التسجيل مع تفعيل النسخ واللصق الكامل.")
        st.markdown("---")

        with st.expander("📂 رفع أو تحديث شيت ممولي الفحص (Excel) ومعاينته للنسخ", expanded=False):
            memo_up = st.file_uploader("اختر شيت الإكسيل (رقم التسجيل، اسم الممول، رقم الملف، النشاط، العنوان، الكيان)", type=["xlsx", "xls", "csv"], key="memo_sheet_uploader")
            if memo_up is not None:
                if st.button("💾 حفظ الشيت في النظام", key="save_memo_sheet"):
                    try:
                        df_m_up = pd.read_excel(memo_up) if not memo_up.name.endswith(".csv") else pd.read_csv(memo_up)
                        df_m_up.columns = df_m_up.columns.str.strip()
                        st.session_state.memo_taxpayer_db = df_m_up.to_dict('records')
                        df_m_up.to_excel(MEMO_DB_PATH, index=False)
                        st.success("✅ تم حفظ شيت ممولي الفحص بنجاح!")
                    except Exception as e:
                        st.error(f"خطأ في قراءة الشيت: {e}")

            if st.session_state.memo_taxpayer_db:
                st.markdown("<b>معاينة البيانات المحفوظة (يمكنك النسخ منها مباشرة ولصقها بالأسفل):</b>", unsafe_allow_html=True)
                st.dataframe(pd.DataFrame(st.session_state.memo_taxpayer_db), use_container_width=True)

        st.markdown("---")
        
        auto_reg = st.text_input("🔎 أدخل رقم التسجيل الضريبي للبحث والملء التلقائي (يمكنك اللصق هنا):", placeholder="اكتب أو ألصق رقم التسجيل هنا...")
        
        found_data = {}
        if auto_reg and st.session_state.memo_taxpayer_db:
            for rec in st.session_state.memo_taxpayer_db:
                for k, v in rec.items():
                    if auto_reg.strip() in str(v):
                        found_data = rec
                        break
                if found_data:
                    break

        def get_val(keys):
            if not found_data:
                return ""
            for k in found_data.keys():
                if any(kw in str(k).lower() for kw in keys):
                    return str(found_data[k])
            return ""

        default_reg = auto_reg if auto_reg else get_val(['تسجيل', 'reg', 'رقم'])
        default_name = get_val(['اسم', 'name', 'الممول'])
        default_file = get_val(['ملف', 'file'])
        default_act = get_val(['نشاط', 'activity'])
        default_addr = get_val(['عنوان', 'address', 'شارع'])
        default_legal = get_val(['كيان', 'legal', 'شكل'])

        col1, col2 = st.columns(2)
        with col1:
            memo_reg = st.text_input("رقم التسجيل الضريبي (9 أرقام)", value=default_reg)
            memo_name = st.text_input("اسم الممول", value=default_name)
            memo_file = st.text_input("رقم الملف", value=default_file)
            memo_activity = st.text_input("النشاط", value=default_act)
        with col2:
            memo_address = st.text_input("العنوان", value=default_addr)
            memo_legal = st.text_input("الكيان القانوني", value=default_legal)
            memo_years = st.text_input("سنوات الفحص (مثال: 2020 إلى 2023)")

        st.markdown("---")
        st.subheader("📌 تفاصيل ومحتوى مذكرة الفحص")
        
        # 1. مقدمة المذكرة (فارغة تماماً ليكتبها المستخدم بنفسه)
        memo_intro = st.text_area("1. مقدمة المذكرة:", placeholder="اكتب مقدمة مذكرة الفحص هنا...")

        # 2. إقرارات الممول
        st.markdown("---")
        st.markdown("#### 2. إقرارات الممول")
        taxpayer_return_choice = st.radio("هل قدم الممول الإقرار؟", ["نعم", "لا"], key="ret_choice")
        if taxpayer_return_choice == "نعم":
            return_details = st.text_input("حدد السنوات والقيمة المقدم عنها الإقرار:", placeholder="مثال: سنة 2021 مقدم عنها الإقرار بمبلغ...")
        else:
            return_details = "لم يقدم / لا يوجد"
            st.info("تم اعتماد: لم يقدم / لا يوجد")

        # 3. الاطلاعات الأربعة تحت بعض بالترتيب
        st.markdown("---")
        st.markdown("#### 3. الاطلاعات الرسمية")
        
        # أ. الخصم والتحصيل
        disc_choice = st.radio("أ. الخصم والتحصيل:", ["يوجد", "لا يوجد"], key="disc_c")
        disc_text = st.text_input("بيانات الخصم والتحصيل:", placeholder="أدخل تفاصيل الخصم والتحصيل...") if disc_choice == "يوجد" else "لا يوجد"

        # ب. القيمة المضافة
        vat_choice = st.radio("ب. القيمة المضافة:", ["يوجد", "لا يوجد"], key="vat_c")
        vat_text = st.text_input("بيانات القيمة المضافة:", placeholder="أدخل تفاصيل القيمة المضافة...") if vat_choice == "يوجد" else "لا يوجد"

        # ج. الفاتورة الإلكترونية
        e_invoice_choice = st.radio("ج. الفاتورة الإلكترونية:", ["يوجد", "لا يوجد"], key="inv_c")
        e_invoice_text = st.text_input("بيانات الفاتورة الإلكترونية:", placeholder="أدخل تفاصيل الفاتورة الإلكترونية...") if e_invoice_choice == "يوجد" else "لا يوجد"

        # د. الجمارك
        customs_choice = st.radio("د. الجمارك:", ["يوجد", "لا يوجد"], key="cust_c")
        customs_text = st.text_input("بيانات الجمارك:", placeholder="أدخل تفاصيل الجمارك...") if customs_choice == "يوجد" else "لا يوجد"

        # 4. محاضر الأعمال
        st.markdown("---")
        st.markdown("#### 4. محاضر الأعمال")
        inspection_record = st.text_input("معاينة بتاريخ يوم كذا:", placeholder="أدخل تاريخ ويوم المعاينة وتفاصيلها...")
        discussion_record = st.text_input("مناقشة بتاريخ يوم كذا:", placeholder="أدخل تاريخ ويوم المناقشة وتفاصيلها...")

        # 5. الفحص والمحاسبة (فارغ تماماً ليكتب فيه المستخدم ما يحتاجه)
        st.markdown("---")
        audit_and_accounting = st.text_area("5. الفحص والمحاسبة:", placeholder="اكتب نتائج الفحص والمحاسبة هنا...")

        # 6. التوقيعات (مأمور ومراجع جنباً بجنب وبحقول فارغة)
        st.markdown("---")
        st.markdown("#### التوقيعات والاعتماد")
        sig_col1, sig_col2 = st.columns(2)
        with sig_col1:
            tax_officer = st.text_input("توقيع المأمور:", placeholder="اسم / توقيع المأمور")
        with sig_col2:
            tax_reviewer = st.text_input("توقيع المراجع:", placeholder="اسم / توقيع المراجع")

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("📥 طباعة وتحميل مذكرة الفحص بصيغة Word احترافي (صلاحيات الطباعة)", key="btn_memo_word"):
            doc = Document()
            doc.add_heading("كيان أحمد سمير حجر - مذكرة الفحص الضريبي الاحترافية", 0)
            doc.add_paragraph(f"رقم التسجيل الضريبي: {memo_reg} | اسم الممول: {memo_name}")
            doc.add_paragraph(f"رقم الملف: {memo_file} | النشاط: {memo_activity}")
            doc.add_paragraph(f"العنوان: {memo_address} | الكيان القانوني: {memo_legal} | سنوات الفحص: {memo_years}")
            
            doc.add_heading("1. مقدمة المذكرة:", level=1)
            doc.add_paragraph(memo_intro if memo_intro else "لا توجد مقدمة مضافة.")

            doc.add_heading("2. إقرارات الممول:", level=1)
            doc.add_paragraph(f"الحالة: {taxpayer_return_choice} | التفاصيل: {return_details}")

            doc.add_heading("3. الاطلاعات الرسمية:", level=1)
            doc.add_paragraph(f"- الخصم والتحصيل: {disc_text}")
            doc.add_paragraph(f"- القيمة المضافة: {vat_text}")
            doc.add_paragraph(f"- الفاتورة الإلكترونية: {e_invoice_text}")
            doc.add_paragraph(f"- الجمارك: {customs_text}")

            doc.add_heading("4. محاضر الأعمال:", level=1)
            doc.add_paragraph(f"- معاينة: {inspection_record}")
            doc.add_paragraph(f"- مناقشة: {discussion_record}")

            doc.add_heading("5. الفحص والمحاسبة:", level=1)
            doc.add_paragraph(audit_and_accounting if audit_and_accounting else "لم يتم ادخال بيانات الفحص والمحاسبة.")

            doc.add_heading("التوقيعات والاعتماد:", level=1)
            doc.add_paragraph(f"المأمور: {tax_officer}                         المراجع: {tax_reviewer}")
            
            tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            doc.save(tmp_file.name)
            with open(tmp_file.name, "rb") as f:
                st.download_button("💾 اضغط هنا لتحميل ملف الـ Word نهائياً", f, file_name=f"Audit_Memo_{memo_reg}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.success("تم إعداد تقرير الـ Word بنجاح وجاهز للتحميل والطباعة!")

    # ==========================================
    # 8. نسب صافي الربح
    # ==========================================
    elif app_mode == "📈 نسب صافي الربح":
        st.title("📈 البحث في نسب صافي الربح (تعليمات 12 و 65 الضريبية)")
        st.markdown("---")
        pr_file = st.file_uploader("رفع شيت نسب صافي الربح", type=["xlsx", "xls", "csv"], key="pr_up")
        if pr_file is not None:
            try:
                df_profit = pd.read_excel(pr_file) if not pr_file.name.endswith(".csv") else pd.read_csv(pr_file)
                df_profit.columns = df_profit.columns.str.strip()
                st.session_state.profit_db = df_profit
                df_profit.to_excel(PROFIT_PATH, index=False)
                st.success("تم رفع وحفظ شيت نسب صافي الربح بنجاح!")
            except Exception as e:
                st.error(f"خطأ في رفع شيت نسب الربح: {e}")

        if st.session_state.profit_db is not None:
            with st.expander("👁️ معاينة شيت نسب صافي الربح المحفوظ بالكامل", expanded=False):
                st.dataframe(st.session_state.profit_db, use_container_width=True)
                st.write(f"إجمالي عدد الصفوف: {len(st.session_state.profit_db)}")

            st.markdown("---")
            st.subheader("🔗 مطابقة أعمدة شيت نسب الربح (المسلسل، الكود، البند، البيان، صافي الربح 12، صافي الربح 65)")
            cols_p = list(st.session_state.profit_db.columns)
            
            def find_p_col(kw_list):
                for col in cols_p:
                    for kw in kw_list:
                        if kw in str(col).lower():
                            return col
                return cols_p[0] if cols_p else ""

            p1, p2, p3, p4, p5, p6 = st.columns(6)
            with p1:
                col_serial = st.selectbox("عمود المسلسل:", cols_p, index=cols_p.index(find_p_col(['مسلسل', 'serial'])) if find_p_col(['مسلسل', 'serial']) in cols_p else 0)
            with p2:
                col_code = st.selectbox("عمود كود النشاط:", cols_p, index=cols_p.index(find_p_col(['كود', 'code'])) if find_p_col(['كود', 'code']) in cols_p else 0)
            with p3:
                col_item = st.selectbox("عمود البند:", cols_p, index=cols_p.index(find_p_col(['بند', 'item'])) if find_p_col(['بند', 'item']) in cols_p else 0)
            with p4:
                col_name = st.selectbox("عمود البيان (مسمى النشاط):", cols_p, index=cols_p.index(find_p_col(['بيان', 'نشاط', 'name'])) if find_p_col(['بيان', 'نشاط', 'name']) in cols_p else 0)
            with p5:
                col_t12 = st.selectbox("صافي الربح 12:", cols_p, index=cols_p.index(find_p_col(['12'])) if find_p_col(['12']) in cols_p else 0)
            with p6:
                col_t65 = st.selectbox("صافي الربح 65:", cols_p, index=cols_p.index(find_p_col(['65'])) if find_p_col(['65']) in cols_p else 0)

            st.markdown("---")
            st.subheader("🔍 البحث برقم كود النشاط")
            act_code = st.text_input("أدخل كود النشاط:", max_chars=10)
            
            if act_code:
                df_p = st.session_state.profit_db
                match_row = df_p[df_p[col_code].astype(str).str.contains(act_code, na=False)]
                
                if not match_row.empty:
                    activity_title = match_row.iloc[0][col_name]
                    rate_12 = match_row.iloc[0][col_t12]
                    rate_65 = match_row.iloc[0][col_t65]
                    
                    st.success(f"✅ تم العثور على نشاط: {activity_title}")
                    
                    r_col1, r_col2 = st.columns(2)
                    with r_col1:
                        st.markdown(f"""
                        <div style="background-color: #f2f8f8; border: 2px solid #075d68; border-radius: 12px; padding: 20px; text-align: center;">
                            <h3 style="color: #075d68; margin-top: 0;">📋 نسبة الربح طبقاً لـ (تعليمات 12)</h3>
                            <p style="font-size: 28px; font-weight: bold; color: #212529; margin: 10px 0;">{rate_12}</p>
                        </div>
                        """, unsafe_allow_html=True)
                    with r_col2:
                        st.markdown(f"""
                        <div style="background-color: #fff8e6; border: 2px solid #d5a63a; border-radius: 12px; padding: 20px; text-align: center;">
                            <h3 style="color: #b8860b; margin-top: 0;">📋 نسبة الربح طبقاً لـ (تعليمات 65)</h3>
                            <p style="font-size: 28px; font-weight: bold; color: #212529; margin: 10px 0;">{rate_65}</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("🖨 طباعة نتيجة نسب صافي الربح (صلاحيات الطباعة - ملف Word)", key="print_profit_res"):
                        doc = Document()
                        doc.add_heading("كيان أحمد سمير حجر - تقرير نسب صافي الربح", 0)
                        doc.add_paragraph(f"كود النشاط: {act_code}")
                        doc.add_paragraph(f"بيان النشاط: {activity_title}")
                        doc.add_paragraph(f"نسبة الربح طبقاً لتعليمات 12: {rate_12}")
                        doc.add_paragraph(f"نسبة الربح طبقاً لتعليمات 65: {rate_65}")
                        tmp_f = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
                        doc.save(tmp_f.name)
                        with open(tmp_f.name, "rb") as f:
                            st.download_button("💾 تحميل تقرير نسب الربح بصيغة Word", f, file_name=f"Profit_Rate_{act_code}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                else:
                    st.warning("⚠️ كود النشاط المدخل غير موجود في الشيت المحفوظ.")
        else:
            st.info("الرجاء رفع شيت نسب صافي الربح أولاً.")

    elif app_mode == "📁 موقف الملف":
        st.title("📁 الاستعلام عن موقف الملف")
        st.markdown("---")
        fs_file = st.file_uploader("رفع قاعدة موقف الملفات", type=["xlsx", "xls", "csv"], key="fs_up")
        if fs_file is not None:
            try:
                df_fs = pd.read_excel(fs_file) if not fs_file.name.endswith(".csv") else pd.read_csv(fs_file)
                st.session_state.file_status_db = df_fs.to_dict('records')
                df_fs.to_excel(FS_PATH, index=False)
                st.success("تم رفع وحفظ قاعدة موقف الملفات بنجاح!")
            except Exception as e:
                st.error(f"خطأ: {e}")

        file_no_query = st.text_input("أدخل رقم الملف")
        if file_no_query and st.session_state.file_status_db:
            results = [item for item in st.session_state.file_status_db if str(item.get('file_no', item.get('رقم_الملف', ''))) == file_no_query]
            if results:
                st.dataframe(pd.DataFrame(results), use_container_width=True)
                if st.button("🖨 طباعة موقف الملف (صلاحيات الطباعة - ملف Word)", key="print_fs_res"):
                    doc = Document()
                    doc.add_heading(f"كيان أحمد سمير حجر - موقف الملف برقم: {file_no_query}", 0)
                    for r in results:
                        doc.add_paragraph(str(r))
                    tmp_f = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
                    doc.save(tmp_f.name)
                    with open(tmp_f.name, "rb") as f:
                        st.download_button("💾 تحميل تقرير موقف الملف بصيغة Word", f, file_name=f"File_Status_{file_no_query}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            else:
                st.warning("لم يتم العثور على بيانات لهذا الملف.")