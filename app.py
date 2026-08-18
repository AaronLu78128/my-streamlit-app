import streamlit as st
import pandas as pd
import plotly.express as px
import os

# 1. 頁面基本設定
st.set_page_config(
    page_title="客訴 List 數據分析平台",
    page_icon="📊",
    layout="wide"
)

# 自訂 CSS：精簡側邊欄與元件樣式
st.markdown("""
    <style>
    [data-testid="stSidebar"] { font-size: 13px; }
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stMultiSelect label,
    [data-testid="stSidebar"] .stRadio label { font-size: 13px !important; }
    [data-testid="stSidebar"] div[role="combobox"] { min-height: 32px !important; font-size: 13px !important; }
    [data-testid="stSidebar"] .stButton button { padding: 2px 8px !important; font-size: 12px !important; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 客訴 List 數據分析與多視角對照平台")

# 2. 建立檔案大廳資料夾
UPLOAD_DIR = "uploaded_files"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

if "current_file" not in st.session_state:
    st.session_state.current_file = None

# 3. 側邊欄：檔案大廳與上傳
st.sidebar.header("🏛️ 1. 檔案大廳 (歷史檔案庫)")

saved_files = [
    f for f in os.listdir(UPLOAD_DIR) 
    if f.endswith(('.xlsx', '.xls', '.xlsm')) and not f.startswith('~$')
]

uploaded_file = st.sidebar.file_uploader("➕ 上傳新檔案至大廳 (.xlsx, .xlsm)", type=["xlsx", "xls", "xlsm"])

if uploaded_file is not None:
    file_path = os.path.join(UPLOAD_DIR, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.sidebar.success(f"✅ 已上傳 `{uploaded_file.name}`！")
    st.session_state.current_file = uploaded_file.name
    saved_files = [
        f for f in os.listdir(UPLOAD_DIR) 
        if f.endswith(('.xlsx', '.xls', '.xlsm')) and not f.startswith('~$')
    ]

st.sidebar.markdown("---")
st.sidebar.subheader("📂 檔案清單")

if saved_files and (st.session_state.current_file not in saved_files):
    st.session_state.current_file = saved_files[0]

if saved_files:
    for filename in saved_files:
        col_del, col_select = st.sidebar.columns([1, 4])
        
        if col_del.button("❌", key=f"del_{filename}", help=f"刪除 {filename}"):
            file_to_delete = os.path.join(UPLOAD_DIR, filename)
            if os.path.exists(file_to_delete):
                os.remove(file_to_delete)
                if st.session_state.current_file == filename:
                    st.session_state.current_file = None
                st.rerun()
        
        is_selected = (filename == st.session_state.current_file)
        btn_label = f"📌 {filename}" if is_selected else f"📄 {filename}"
        btn_type = "primary" if is_selected else "secondary"
        
        if col_select.button(btn_label, key=f"select_{filename}", type=btn_type, use_container_width=True):
            st.session_state.current_file = filename
            st.rerun()
else:
    st.sidebar.info("💡 目前大廳是空的，請先在上方上傳 Excel 檔案！")

# 4. 讀取與解析 Excel 資料
def load_and_clean_excel(file_path, sheet_name=None):
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names
    
    if not sheet_name or sheet_name not in sheet_names:
        sheet_name = sheet_names[0]
        for name in sheet_names:
            if "客" in name or "list" in name.lower():
                sheet_name = name
                break

    df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
    df.columns = [str(col).strip() for col in df.columns]
    valid_cols = [col for col in df.columns if col != "" and not col.startswith("Unnamed") and col.lower() != "nan"]
    
    if len(valid_cols) == 0:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
        df.columns = [str(col).strip() for col in df.columns]
        valid_cols = [col for col in df.columns if col != "" and not col.startswith("Unnamed") and col.lower() != "nan"]

    df = df[valid_cols]
    df = df.dropna(how='all')
    return df, sheet_names, sheet_name

# 5. 核心函式：圖表計算與生成（模組化獨立邏輯）
def process_and_render_chart(df, x_axis, legend_axis, y_axis, chart_type, chart_height=420):
    color_col = None if legend_axis == "無" else legend_axis
    calc_df = df.copy()

    # 專一判定邏輯：僅「不良數量」做加總，其餘算筆數
    if y_axis == "不良數量":
        calc_df["不良數量"] = pd.to_numeric(calc_df["不良數量"], errors='coerce').fillna(0)
        y_metric = "不良數量 (總和)" if color_col == "不良數量" else "不良數量"

        if color_col and x_axis != color_col:
            grouped_df = calc_df.groupby([x_axis, color_col])["不良數量"].sum().reset_index(name=y_metric)
        else:
            grouped_df = calc_df.groupby([x_axis])["不良數量"].sum().reset_index(name=y_metric)
            color_col = None
    else:
        y_metric = "數量"
        if color_col and x_axis != color_col:
            grouped_df = calc_df.groupby([x_axis, color_col]).size().reset_index(name="數量")
        else:
            grouped_df = calc_df.groupby([x_axis]).size().reset_index(name="數量")
            color_col = None

    # 繪製 Plotly 圖表
    if chart_type == "堆疊直條圖 (Stacked)":
        fig = px.bar(grouped_df, x=x_axis, y=y_metric, color=color_col, barmode="stack", text_auto=True)
    elif chart_type == "聚集直條圖 (Grouped)":
        fig = px.bar(grouped_df, x=x_axis, y=y_metric, color=color_col, barmode="group", text_auto=True)
    elif chart_type == "折線圖 (Line)":
        fig = px.line(grouped_df, x=x_axis, y=y_metric, color=color_col, markers=True)
    elif chart_type == "圓餅圖 (Pie)":
        fig = px.pie(grouped_df, names=x_axis, values=y_metric)

    fig.update_layout(
        font=dict(family="Microsoft JhengHei", size=13),
        height=chart_height,
        margin=dict(l=10, r=10, t=25, b=10),
        xaxis_title=x_axis,
        yaxis_title=y_metric,
        legend_title=color_col if color_col else ""
    )
    return fig, grouped_df

# 6. 主頁面顯示設定
selected_filename = st.session_state.current_file

if not selected_filename or not os.path.exists(os.path.join(UPLOAD_DIR, selected_filename)):
    st.info("👈 請先從左側檔案大廳上傳或選擇要分析的 Excel 檔案。")
else:
    file_path = os.path.join(UPLOAD_DIR, selected_filename)

    st.sidebar.markdown("---")
    st.sidebar.header("⚙️ 2. 圖表數量設定")
    num_charts = st.sidebar.selectbox("請選擇同時顯示的圖表數量：", options=[1, 2, 3, 4], index=0)

    st.subheader(f"📌 當前分析檔案：`{selected_filename}` (同時顯示 {num_charts} 個圖表視角)")

    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names

        # 動態建立多欄版面
        cols = st.columns(num_charts)

        for idx in range(num_charts):
            with cols[idx]:
                st.markdown(f"### 📊 圖表視角 {idx+1}")
                
                # 參數設定區塊
                with st.expander(f"⚙️ 設定視角 {idx+1} 參數", expanded=True):
                    default_sheet_idx = min(idx, len(sheet_names) - 1)
                    chosen_sheet = st.selectbox(
                        "工作表 (Sheet)", 
                        options=sheet_names, 
                        index=default_sheet_idx, 
                        key=f"sheet_{idx}"
                    )
                    
                    df, _, _ = load_and_clean_excel(file_path, chosen_sheet)
                    cols_list = df.columns.tolist()

                    if cols_list:
                        default_x_idx = min(idx, len(cols_list) - 1)
                        x_axis = st.selectbox(
                            "X 軸（主要類別）", 
                            options=cols_list, 
                            index=default_x_idx, 
                            key=f"x_{idx}"
                        )
                        
                        legend_opts = ["無"] + cols_list
                        legend_axis = st.selectbox(
                            "圖例（顏色分類）", 
                            options=legend_opts, 
                            index=0, 
                            key=f"leg_{idx}"
                        )
                        
                        y_opts = ["資料筆數 (Count)"] + cols_list
                        y_axis = st.selectbox(
                            "統計欄位（Y 軸）",
                            options=y_opts,
                            index=0,
                            key=f"y_{idx}"
                        )
                        
                        chart_type = st.selectbox(
                            "圖表類型", 
                            options=["聚集直條圖 (Grouped)", "堆疊直條圖 (Stacked)", "折線圖 (Line)", "圓餅圖 (Pie)"],
                            index=0,
                            key=f"chart_{idx}"
                        )

                # 呼叫統一函式繪圖與數據整理
                if cols_list:
                    fig, grouped_df = process_and_render_chart(
                        df, x_axis, legend_axis, y_axis, chart_type, 
                        chart_height=450 if num_charts == 1 else 380
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)

                    with st.expander("🔍 查看統計數據明細"):
                        st.dataframe(grouped_df, use_container_width=True)

    except Exception as e:
        st.error(f"讀取或解析檔案時發生錯誤：{e}")

# 7. 側邊欄留白
for _ in range(5):
    st.sidebar.write("")
