# main_app.py
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import re
import numpy as np

# ==============================================================================
# PAGE CONFIGURATION
# ==============================================================================
st.set_page_config(
    layout="wide",
    page_title="KPI Dashboard",
    page_icon="📊"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@300;400;500;600;700&display=swap');
    html, body, [data-testid="stAppViewContainer"], [data-testid="stSidebar"] {
        font-family: 'Kanit', sans-serif;
    }
    .kpi-title { font-size: 2rem; font-weight: 700; color: #1E3A8A; text-align: center; margin-bottom: 0.5rem; }
    .kpi-n-value { font-size: 1.2rem; font-weight: 500; color: #4B5563; text-align: center; margin-bottom: 2rem; }
    .interpretation-box { border: 1px solid #E5E7EB; border-radius: 12px; padding: 1.5rem; background-color: #FFFFFF; font-size: 1rem; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1); }
    .interp-item { margin-bottom: 0.8rem; display: flex; flex-wrap: wrap; line-height: 1.6; }
    .interp-label { font-weight: 600; color: #4B5563; margin-right: 8px; flex-shrink: 0; }
    .interp-value { font-weight: 400; color: #1F2937; }
    .interp-criteria { margin-top: 1.2rem; margin-bottom: 1.2rem; }
    .interp-criteria ul { list-style: none; padding-left: 0; margin-top: 0.5rem; }
    .interp-criteria li { margin-bottom: 0.4rem; display: flex; align-items: center; }
    .color-swatch { width: 14px; height: 14px; border-radius: 4px; margin-right: 10px; display: inline-block; border: 1px solid rgba(0,0,0,0.1); }
    .interp-summary { margin-top: 1.2rem; padding: 1rem; background-color: #EFF6FF; border-left: 5px solid #3B82F6; border-radius: 4px; }
    .interp-summary .interp-label { color: #1E3A8A; }
    .stDataFrame th { white-space: normal !important; overflow-wrap: break-word !important; width: auto !important; }
    .stDataFrame td { text-align: center !important; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# DATA LOADING AND PREPARATION (*** UPDATED FUNCTION ***)
# ==============================================================================
@st.cache_data
def load_kpi_data(source):
    if source is None:
        return pd.DataFrame(), "No source"

    try:
        if hasattr(source, 'name'): # Handles Streamlit UploadedFile
            file_info = f"ไฟล์ที่อัปโหลด: `{source.name}`"
            df = pd.read_excel(source)
        else: # Handles a URL
            file_info = "ไฟล์เริ่มต้นจาก GitHub"
            # Use openpyxl engine which is required for reading xlsx from URLs
            df = pd.read_excel(source, engine='openpyxl')

        # Find the correct header row by looking for 'N'
        header_row_index = -1
        for i, row in df.head(10).iterrows():
            if 'N' in row.values:
                header_row_index = i
                break
        
        if header_row_index == -1:
            st.error("ไม่พบคอลัมน์ 'N' ใน 10 แถวแรกของไฟล์ ไม่สามารถประมวลผลได้")
            return pd.DataFrame(), file_info

        # Re-read the data with the correct header
        if hasattr(source, 'name'):
             source.seek(0) # Reset file pointer for re-reading
             df = pd.read_excel(source, header=header_row_index)
        else:
             df = pd.read_excel(source, header=header_row_index, engine='openpyxl')

        if 'Unnamed: 3' in df.columns:
            df.rename(columns={'Unnamed: 3': 'kpi_name'}, inplace=True)
        else:
            st.error("ไม่พบคอลัมน์ 'Unnamed: 3' ที่คาดว่าจะมีชื่อตัวชี้วัด")
            return pd.DataFrame(), file_info
            
        df_kpi = df[pd.to_numeric(df['N'], errors='coerce').notna()].copy()
        numeric_cols = ['N', 'KPI Value', 'P25', 'Median', 'P75']
        for col in numeric_cols:
            df_kpi[col] = pd.to_numeric(df_kpi[col], errors='coerce')
        df_kpi.dropna(subset=['kpi_name', 'KPI Value', 'P25', 'Median', 'P75'], inplace=True)
        return df_kpi, file_info
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านหรือประมวลผลไฟล์: {e}")
        return pd.DataFrame(), "Error"

# ==============================================================================
# HELPER FUNCTIONS (UNCHANGED)
# ==============================================================================
def plot_kpi_gauge(kpi_data):
    kpi_name = kpi_data['kpi_name']
    value, p25, median, p75 = kpi_data['KPI Value'], kpi_data['P25'], kpi_data['Median'], kpi_data['P75']
    n_value = int(kpi_data['N'])
    lower_is_better = any(re.search(keyword, kpi_name, re.IGNORECASE) for keyword in
                          ['เสียชีวิต', 'ตาย', 'ติดเชื้อ', 'แทรกซ้อน', 'ระยะเวลา'])
    red, orange, yellow, green = '#EF4444', '#F97316', '#FBBF24', '#22C55E'

    max_val = max(value, p75) * 1.2 if max(value, p75) > 0 else 1

    steps = ([{'range': [0, p25], 'color': green}, {'range': [p25, median], 'color': yellow},
              {'range': [median, p75], 'color': orange}, {'range': [p75, max_val], 'color': red}]
             if lower_is_better else
             [{'range': [0, p25], 'color': red}, {'range': [p25, median], 'color': orange},
              {'range': [median, p75], 'color': yellow}, {'range': [p75, max_val], 'color': green}])

    if lower_is_better:
        st.info("สำหรับตัวชี้วัดนี้: **ค่าที่ต่ำกว่า** ถือว่าดีกว่า")
    else:
        st.info("สำหรับตัวชี้วัดนี้: **ค่าที่สูงกว่า** ถือว่าดีกว่า")

    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=value, number={'valueformat': '.2f', 'font': {'size': 50}},
        title={'text': "KPI Value", 'font': {'size': 24}},
        gauge={ 'axis': {'range': [0, max_val]}, 'bar': {'color': "#1E3A8A", 'thickness': 0.3},
            'bgcolor': "white", 'borderwidth': 2, 'bordercolor': "#D1D5DB", 'steps': steps }))

    annotation_text = (f"<b>P25:</b> {p25:.2f}<br>" f"<b>Median:</b> {median:.2f}<br>" f"<b>P75:</b> {p75:.2f}")
    fig.add_annotation( x=0.05, y=0.95, xref="paper", yref="paper", text=annotation_text, showarrow=False,
        font=dict(family="Kanit, sans-serif", size=12, color="#1F2937"), align="left",
        bordercolor="#E5E7EB", borderwidth=1, borderpad=4, bgcolor="#F9FAFB", opacity=0.9)
    fig.update_layout(height=400, margin=dict(t=50, r=50, b=50, l=50))
    st.markdown(f'<p class="kpi-title">{kpi_name}</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="kpi-n-value">N = {n_value:,}</p>', unsafe_allow_html=True)
    st.plotly_chart(fig, use_container_width=True)

def interpret_kpi_data(kpi_data):
    kpi_name, value, p25, median, p75 = kpi_data['kpi_name'], kpi_data['KPI Value'], kpi_data['P25'], kpi_data['Median'], kpi_data['P75']
    unit = " นาที" if "นาที" in kpi_name else ("%" if "ร้อยละ" in kpi_name or "%" in kpi_name else "")
    lower_is_better = any(re.search(keyword, kpi_name, re.IGNORECASE) for keyword in ['เสียชีวิต', 'ตาย', 'ติดเชื้อ', 'แทรกซ้อน', 'ระยะเวลา'])
    levels = [{'range': (0, p25), 'label': 'ระดับที่ 1'}, {'range': (p25, median), 'label': 'ระดับที่ 2'},
              {'range': (median, p75), 'label': 'ระดับที่ 3'}, {'range': (p75, np.inf), 'label': 'ระดับที่ 4'}]
    level_colors = (['เขียว', 'เหลือง', 'ส้ม', 'แดง'] if lower_is_better else ['แดง', 'ส้ม', 'เหลือง', 'เขียว'])
    for i, level in enumerate(levels): level['color'] = level_colors[i]
    interpretation = "ไม่สามารถระบุระดับได้"
    for level in levels:
        if level['range'][0] <= value < level['range'][1]:
            interpretation = f"ค่าที่วัดได้ <strong>{value:,.2f}{unit}</strong> อยู่ในช่วงของ <strong>{level['label']} ({level['color']})</strong>"
            break
    if value >= p75: interpretation = f"ค่าที่วัดได้ <strong>{value:,.2f}{unit}</strong> อยู่ในช่วงของ <strong>{levels[-1]['label']} ({levels[-1]['color']})</strong>"
    color_map_hex = {'แดง': '#EF4444', 'ส้ม': '#F97316', 'เหลือง': '#FBBF24', 'เขียว': '#22C55E'}
    criteria_html = "<ul>"
    for level in levels:
        min_r, max_r = level['range']
        range_str = f"{min_r:,.2f} - {max_r:,.2f}{unit}" if max_r != np.inf else f"มากกว่า {min_r:,.2f}{unit}"
        color_hex = color_map_hex.get(level['color'], '#D1D5DB')
        criteria_html += f'<li><span class="color-swatch" style="background-color: {color_hex};"></span><strong>{level["label"]} ({level["color"]}):</strong>&nbsp;{range_str}</li>'
    criteria_html += "</ul>"
    output = f"""<div class="interpretation-box"><div class="interp-item"><div class="interp-label">หัวข้อหลัก:</div><div class="interp-value">{kpi_name}</div></div><div class="interp-item"><div class="interp-label">ช่วงเวลาของข้อมูล:</div><div class="interp-value">ตุลาคม 2567 ถึง กันยายน 2568 (จากข้อมูลในไฟล์)</div></div><div class="interp-item"><div class="interp-label">ค่าที่วัดได้:</div><div class="interp-value" style="font-size: 1.1rem; font-weight: 700; color: #1E3A8A;">{value:,.2f}{unit}</div></div><div class="interp-criteria"><div class="interp-label">การแบ่งเกณฑ์:</div>{criteria_html}</div><div class="interp-summary"><div class="interp-label">สรุป/การตีความ:</div><div class="interp-value">{interpretation}</div></div></div>"""
    return output

# ==============================================================================
# MAIN APP LAYOUT (*** UPDATED SECTION ***)
# ==============================================================================
st.title("MCH : THIP")

with st.sidebar:
    st.header("อัปโหลดข้อมูล")
    uploaded_file = st.file_uploader(
        "เลือกไฟล์ XLSX (หากไม่อัปโหลดจะใช้ไฟล์เริ่มต้น)",
        type=['xlsx']
    )

# --- Define the default file URL ---
DEFAULT_FILE_URL = "https://raw.githubusercontent.com/HOIARRTool/mthip/main/mthip2.xlsx"

# --- Data Loading Logic ---
data_source = uploaded_file if uploaded_file is not None else DEFAULT_FILE_URL

if uploaded_file is None:
    st.info("กำลังใช้ข้อมูลเริ่มต้นจาก GitHub หากต้องการใช้ข้อมูลอื่น กรุณาอัปโหลดไฟล์ใหม่")

df_kpi, source_info = load_kpi_data(data_source)

st.sidebar.markdown(f"**แหล่งข้อมูล:**\n{source_info}")

# --- Main Content Area ---
if not df_kpi.empty:
    st.success(f"โหลดข้อมูลสำเร็จ! พบตัวชี้วัด {len(df_kpi)} รายการ")
    kpi_list = df_kpi['kpi_name'].unique().tolist()
    selected_kpi_name = st.selectbox("**กรุณาเลือกตัวชี้วัดที่ต้องการดู:**", options=kpi_list)
    st.markdown("---")

    if selected_kpi_name:
        selected_kpi_data = df_kpi[df_kpi['kpi_name'] == selected_kpi_name].iloc[0]

        col1, col2 = st.columns([3, 2])

        with col1:
            plot_kpi_gauge(selected_kpi_data)
        with col2:
            st.subheader("การแปลผลข้อมูล 📝")
            interpretation_result = interpret_kpi_data(selected_kpi_data)
            st.markdown(interpretation_result, unsafe_allow_html=True)
else:
    st.warning("ไม่สามารถโหลดข้อมูลได้ กรุณาตรวจสอบไฟล์หรือ URL อีกครั้ง")
