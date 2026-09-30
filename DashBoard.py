import streamlit as st
import pandas as pd
import plotly.express as px
import os

# 1. ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="Company Sales & Deposit Analytics App",
    page_icon="📊",
    layout="wide"
)

# กำหนดสีประจำบริษัท
color_map = {
    'YKT': '#4682B4',  # สีฟ้า
    'POM': '#FFD700',  # สีเหลือง
    'SPG': '#FF7F50'   # สีส้ม
}

# กำหนดชื่อไฟล์กลางในระบบ
CENTRAL_FILE_PATH = "20.09.69 แปลง.xlsx"

# ดึงชื่อไฟล์มาตัดนามสกุลออกเพื่อเอาไปแสดงผลเป็นวันที่ (เช่น "20.09.69 แปลง" หรือดึงเฉพาะวันที่)
if os.path.exists(CENTRAL_FILE_PATH):
    # ดึงชื่อไฟล์แบบไม่มีนามสกุล
    file_display_name = os.path.splitext(CENTRAL_FILE_PATH)[0]
else:
    file_display_name = "-"

# แสดงหัวข้อรายงานพร้อมดึงชื่อวันที่ตามไฟล์
st.markdown(f"### 📊 ระบบรายงาน Dashboard ยอดขายฝากและเงินเก๊ะ ประจำวันที่: {file_display_name}")
st.markdown("---")

# 2. ระบบจัดการสิทธิ์แอดมินใน Sidebar
st.sidebar.markdown("### 📁 จัดการข้อมูล Excel")

# ตั้งรหัสผ่านสำหรับแอดมินที่นี่ (สามารถเปลี่ยนเลขได้ตามต้องการ)
ADMIN_PASSWORD = "1234" 

# สร้างปุ่มให้เลือกโหมด (สำหรับแอดมินใส่รหัส)
mode_options = ["ดูข้อมูลทั่วไป (ผู้ชมทั่วไป)"]
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

if st.session_state.is_admin:
    mode_options.append("อัปโหลดไฟล์ Excel ใหม่ (แอดมิน)")

selected_mode = st.sidebar.selectbox("เลือกโหมดการใช้งาน:", mode_options)

# ถ้ายัังไม่ได้เป็นแอดมิน ให้มีช่องกรอกรหัสผ่านด้านล่าง
if not st.session_state.is_admin:
    st.sidebar.markdown("---")
    with st.sidebar.expander("🔐 สำหรับแอดมิน (อัปเดตข้อมูล)"):
        password_input = st.text_input("กรอกรหัสผ่านแอดมิน", type="password")
        if st.button("เข้าสู่ระบบแอดมิน"):
            if password_input == ADMIN_PASSWORD:
                st.session_state.is_admin = True
                st.success("เข้าสู่ระบบสำเร็จ!")
                st.rerun()
            else:
                st.error("รหัสผ่านไม่ถูกต้อง")

# กำหนดเป้าหมายไฟล์ที่จะดึงมาแสดงผล
target_file = None

if st.session_state.is_admin and selected_mode == "อัปโหลดไฟล์ Excel ใหม่ (แอดมิน)":
    st.sidebar.markdown("---")
    st.sidebar.markdown("#### 📤 อัปโหลดไฟล์ประจำวัน")
    uploaded_file = st.sidebar.file_uploader("เลือกไฟล์ Excel ของคุณ", type=["xlsx", "xls"])
    if uploaded_file is not None:
        # บันทึกไฟล์ด้วยชื่อเดิม (หรือจะใช้ชื่อตามไฟล์ที่อัปโหลดเข้ามาจริงก็ได้)
        CENTRAL_FILE_PATH = uploaded_file.name
        with open(CENTRAL_FILE_PATH, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.sidebar.success(f"✅ อัปโหลดไฟล์ {uploaded_file.name} สำเร็จ!")
        st.rerun()
    
    if st.sidebar.button("ออกจากระบบแอดมิน"):
        st.session_state.is_admin = False
        st.rerun()

# ตรวจสอบไฟล์ที่จะนำมาประมวลผล (อ่านจากไฟล์กลางบนระบบ)
if os.path.exists(CENTRAL_FILE_PATH):
    target_file = CENTRAL_FILE_PATH
else:
    target_file = None

# 3. เริ่มประมวลผลข้อมูลหากมีไฟล์พร้อมใช้งาน
if target_file is not None:
    try:
        df_raw = pd.read_excel(target_file, sheet_name='ส่งเสี่ย', header=5)
        
        # จัดการชื่อคอลัมน์
        col_mapping = {
            df_raw.columns[0]: 'ลำดับ',
            df_raw.columns[1]: 'บริษัท_raw',
            df_raw.columns[2]: 'สาขา',
            df_raw.columns[3]: 'ยอดขายฝาก',
            df_raw.columns[4]: 'ยอดเงินเก๊ะใหญ่',
            df_raw.columns[5]: 'ยอดเงินเก๊ะขายฝาก'
        }
        
        df = df_raw.rename(columns=col_mapping)[['ลำดับ', 'บริษัท_raw', 'สาขา', 'ยอดขายฝาก', 'ยอดเงินเก๊ะใหญ่', 'ยอดเงินเก๊ะขายฝาก']].copy()
        
        # กรองเฉพาะแถวข้อมูลสาขาจริง
        df = df.dropna(subset=['สาขา', 'ลำดับ'])
        df = df[~df['สาขา'].astype(str).str.contains('รวม|หมายเหตุ|\*', na=False)]
        df['ลำดับ'] = pd.to_numeric(df['ลำดับ'], errors='coerce')
        df = df.dropna(subset=['ลำดับ'])
        
        # แปลงข้อมูลตัวเลข
        for col in ['ยอดขายฝาก', 'ยอดเงินเก๊ะใหญ่', 'ยอดเงินเก๊ะขายฝาก']:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            
        df['ยอดรวม'] = df['ยอดขายฝาก'] + df['ยอดเงินเก๊ะใหญ่'] + df['ยอดเงินเก๊ะขายฝาก']
            
        def map_company(val):
            val_str = str(val).strip().upper()
            if val_str == 'POM':
                return 'POM'
            elif val_str == 'SPG':
                return 'SPG'
            else:
                return 'YKT'
                
        df['บริษัท'] = df['บริษัท_raw'].apply(map_company)
        
        # Sidebar ตัวกรองข้อมูลสำหรับผู้ใช้งานทั่วไป
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 🎯 ตัวกรองการแสดงผล")
        
        metric_option = st.sidebar.radio(
            "เลือกมุมมองข้อมูล:",
            ("All (รวมยอดทั้งหมด)", "ยอดขายฝาก", "ยอดเงินเก๊ะใหญ่", "ยอดเงินเก๊ะขายฝาก")
        )
        
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 🏢 เลือกบริษัท (ติ๊กเลือกได้มากกว่า 1)")
        
        chk_ykt = st.sidebar.checkbox("YKT", value=True)
        chk_pom = st.sidebar.checkbox("POM", value=True)
        chk_spg = st.sidebar.checkbox("SPG", value=True)
        
        selected_companies = []
        if chk_ykt: selected_companies.append('YKT')
        if chk_pom: selected_companies.append('POM')
        if chk_spg: selected_companies.append('SPG')
        
        if metric_option == "ยอดขายฝาก":
            selected_col = 'ยอดขายฝาก'
            metric_title = "ยอดขายฝาก"
        elif metric_option == "ยอดเงินเก๊ะใหญ่":
            selected_col = 'ยอดเงินเก๊ะใหญ่'
            metric_title = "ยอดเงินเก๊ะใหญ่"
        elif metric_option == "ยอดเงินเก๊ะขายฝาก":
            selected_col = 'ยอดเงินเก๊ะขายฝาก'
            metric_title = "ยอดเงินเก๊ะขายฝาก"
        else:
            selected_col = 'ยอดรวม'
            metric_title = "รวมยอดทั้ง 3 หัวข้อ"
            
        if selected_companies:
            df_filtered = df[df['บริษัท'].isin(selected_companies)].copy()
        else:
            df_filtered = pd.DataFrame(columns=df.columns)
            
        # สถิติสูงสุด (KPI Cards)
        max_sales_dep = df_filtered.loc[df_filtered['ยอดขายฝาก'].idxmax()] if not df_filtered.empty else None
        max_cash_big = df_filtered.loc[df_filtered['ยอดเงินเก๊ะใหญ่'].idxmax()] if not df_filtered.empty else None
        max_cash_dep = df_filtered.loc[df_filtered['ยอดเงินเก๊ะขายฝาก'].idxmax()] if not df_filtered.empty else None
        max_total_branch = df_filtered.loc[df_filtered['ยอดรวม'].idxmax()] if not df_filtered.empty else None

        st.markdown(f"#### 🏆 สรุปสถิติยอดสูงสุด (Top Performance)")
        
        kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
        
        def render_clean_card(title, value, subtitle):
            return f"""
            <div style="background-color: #F8F9FA; padding: 16px; border-radius: 10px; border: 1px solid #E0E0E0; box-shadow: 0 1px 3px rgba(0,0,0,0.05); height: 120px; display: flex; flex-direction: column; justify-content: space-between;">
                <div style="color: #6C757D; font-size: 13px; font-weight: 600;">{title}</div>
                <div style="color: #212529; font-size: 18px; font-weight: bold; letter-spacing: -0.5px;">{value}</div>
                <div style="color: #0D6EFD; font-size: 12px; font-weight: 500; background-color: #E7F1FF; padding: 2px 8px; border-radius: 4px; width: fit-content;">{subtitle}</div>
            </div>
            """

        with kpi_col1:
            val_str = f"{max_sales_dep['ยอดขายฝาก']:,.2f} บาท" if max_sales_dep is not None and not df_filtered.empty else "0.00 บาท"
            sub_str = f"สาขา: {max_sales_dep['สาขา']} ({max_sales_dep['บริษัท']})" if max_sales_dep is not None and not df_filtered.empty else "สาขา: -"
            st.markdown(render_clean_card("🔥 ยอดขายฝากสูงสุด", val_str, sub_str), unsafe_allow_html=True)
            
        with kpi_col2:
            val_str = f"{max_cash_big['ยอดเงินเก๊ะใหญ่']:,.2f} บาท" if max_cash_big is not None and not df_filtered.empty else "0.00 บาท"
            sub_str = f"สาขา: {max_cash_big['สาขา']} ({max_cash_big['บริษัท']})" if max_cash_big is not None and not df_filtered.empty else "สาขา: -"
            st.markdown(render_clean_card("🔥 เงินเก๊ะใหญ่สูงสุด", val_str, sub_str), unsafe_allow_html=True)
            
        with kpi_col3:
            val_str = f"{max_cash_dep['ยอดเงินเก๊ะขายฝาก']:,.2f} บาท" if max_cash_dep is not None and not df_filtered.empty else "0.00 บาท"
            sub_str = f"สาขา: {max_cash_dep['สาขา']} ({max_cash_dep['บริษัท']})" if max_cash_dep is not None and not df_filtered.empty else "สาขา: -"
            st.markdown(render_clean_card("🔥 เงินเก๊ะขายฝากสูงสุด", val_str, sub_str), unsafe_allow_html=True)
            
        with kpi_col4:
            val_str = f"{max_total_branch['ยอดรวม']:,.2f} บาท" if max_total_branch is not None and not df_filtered.empty else "0.00 บาท"
            sub_str = f"สาขา: {max_total_branch['สาขา']} ({max_total_branch['บริษัท']})" if max_total_branch is not None and not df_filtered.empty else "สาขา: -"
            st.markdown(render_clean_card("⭐ ยอดรวมสูงสุด (ทุกหัวข้อ)", val_str, sub_str), unsafe_allow_html=True)
            
        st.markdown("---")
        
        # แผนภูมิวงกลม
        st.markdown(f"### 🥧 สัดส่วนและรายละเอียดของแต่ละบริษัท ({metric_title})")
        chart_col, info_col = st.columns([2, 1])
        
        df_all_grouped = df.groupby('บริษัท')[selected_col].sum().reset_index()
        df_all_grouped.rename(columns={selected_col: 'Total_Metric'}, inplace=True)
        total_sum_all = df_all_grouped['Total_Metric'].sum()
        
        with chart_col:
            fig_pie = px.pie(
                df_all_grouped, 
                names='บริษัท', 
                values='Total_Metric', 
                hole=0.4,
                color='บริษัท',
                color_discrete_map=color_map
            )
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            fig_pie.update_layout(height=450)
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with info_col:
            st.markdown("#### 📋 รายละเอียดยอดแต่ละบริษัท")
            val_dict = dict(zip(df_all_grouped['บริษัท'], df_all_grouped['Total_Metric']))
            
            ykt_val = val_dict.get('YKT', 0)
            pom_val = val_dict.get('POM', 0)
            spg_val = val_dict.get('SPG', 0)
            
            ykt_pct = (ykt_val / total_sum_all * 100) if total_sum_all > 0 else 0
            pom_pct = (pom_val / total_sum_all * 100) if total_sum_all > 0 else 0
            spg_pct = (spg_val / total_sum_all * 100) if total_sum_all > 0 else 0
            
            st.markdown(f"🔹 **บริษัท YKT ({ykt_pct:.1f}%):**")
            st.info(f"**{ykt_val:,.2f}** บาท")
            
            st.markdown(f"🔸 **บริษัท POM ({pom_pct:.1f}%):**")
            st.warning(f"**{pom_val:,.2f}** บาท")
            
            st.markdown(f"🔸 **บริษัท SPG ({spg_pct:.1f}%):**")
            st.error(f"**{spg_val:,.2f}** บาท")
            
            st.markdown(f"💰 **ยอดรวมทุกบริษัท (100.0%):**")
            st.success(f"**{total_sum_all:,.2f}** บาท")

        st.markdown("---")
        
        # ตารางแสดงผล
        st.markdown(f"🏆 **5 สาขาที่สูงสุด และ 5 สาขาที่ต่ำที่สุด ({metric_title})**")
        if not df_filtered.empty:
            df_sorted = df_filtered.sort_values(by=selected_col, ascending=False).reset_index(drop=True)
            
            if metric_option == "All (รวมยอดทั้งหมด)":
                cols_to_show = ['ลำดับ', 'สาขา', 'บริษัท', 'ยอดขายฝาก', 'ยอดเงินเก๊ะใหญ่', 'ยอดเงินเก๊ะขายฝาก', 'ยอดรวม']
                format_dict = {'ลำดับ': '{:.0f}', 'ยอดขายฝาก': '{:,.2f}', 'ยอดเงินเก๊ะใหญ่': '{:,.2f}', 'ยอดเงินเก๊ะขายฝาก': '{:,.2f}', 'ยอดรวม': '{:,.2f}'}
            else:
                cols_to_show = ['ลำดับ', 'สาขา', 'บริษัท', selected_col]
                format_dict = {'ลำดับ': '{:.0f}', selected_col: '{:,.2f}'}
            
            top_5 = df_sorted.head(5)[cols_to_show]
            bottom_5 = df_sorted.tail(min(5, len(df_sorted))).sort_values(by=selected_col, ascending=True)[cols_to_show]
            
            st.markdown("**🔥 5 สาขายอดสูงสุด**")
            st.dataframe(top_5.style.format(format_dict), use_container_width=True, hide_index=True)
            st.markdown("")
            st.markdown("**❄️ 5 สาขายอดต่ำที่สุด**")
            st.dataframe(bottom_5.style.format(format_dict), use_container_width=True, hide_index=True)
        else:
            st.warning("กรุณาติ๊กเลือกบริษัทอย่างน้อย 1 บริษัท")

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์: {e}")
else:
    st.warning("⚠️ ยังไม่พบไฟล์ข้อมูลกลางในระบบ")
    st.info("💡 คำแนะนำสำหรับแอดมิน: กรุณาขยายเมนู 'สำหรับแอดมิน' ด้านซ้าย ใส่รหัสผ่านเพื่ออัปโหลดไฟล์ตั้งต้นขึ้นระบบครับ")
