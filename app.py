import streamlit as st
import pandas as pd
import plotly.express as px
import os
import json

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

# 2. 建立檔案大廳與常用資料夾儲存機制
UPLOAD_DIR = "uploaded_files"
FAVORITES_FILE = "favorites.json"
CONFIG_FILE = "saved_configs.json"

if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

def load_favorites():
    if os.path.exists(FAVORITES_FILE):
        try:
            with open(FAVORITES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_favorites(fav_list):
    with open(FAVORITES_FILE, "w", encoding="utf-8") as f:
        json.dump(fav_list, f, ensure_ascii=False, indent=4)

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

# 4. 側邊欄：⭐ 常用圖表庫 (常用資料夾)
st.sidebar.markdown("---")
st.sidebar.header("📁 常用圖表庫 (常用資料夾)")

fav_list = load_favorites()

if not fav_list:
    st.sidebar.info("💡 目前沒有常用圖表。繪製完圖表後，點擊圖表下方的「⭐ 加入常用」即可存入！")
else:
    for idx_fav, fav_item in enumerate(fav_list):
        col_fav_del, col_fav_load = st.sidebar.columns([1, 4])
        
        # 刪除常用功能
        if col_fav_del.button("❌", key=f"del_fav_{idx_fav}", help=f"刪除常用：{fav_item['name']}"):
            fav_list.pop(idx_fav)
            save_favorites(fav_list)
            st.sidebar.success(f"🗑️ 已刪除「{fav_item['name']}」")
            st.rerun()
        
        # 點擊直接載入常用圖表至第一個視角
        if col_fav_load.button(f"📌 {fav_item['name']}", key=f"load_fav_{idx_fav}", use_container_width=True):
            st.session_state.current_file = fav_item["file"]
            st.session_state["sheet_0"] = fav_item["sheet"]
            st.session_state["x_0"] = fav_item["x_axis"]
            st.session_state["x_filter_0"] = fav_item["selected_x_vals"]
            st.session_state["leg_0"] = fav_item["legend_axis"]
            st.session_state["y_0"] = fav_item["y_axis"]
            st.session_state["chart_0"] = fav_item["chart_type"]
            st.rerun()

# 5. 讀取與解析 Excel 資料
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

# 6. 核心函式：圖表計算與生成
def process_and_render_chart(df, x_axis, legend_axis, y_axis, chart_type, chart_height=420):
    color_col = None if legend_axis == "無" else legend_axis
    calc_df = df.copy()

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

# 7. 主頁面顯示設定
selected_filename = st.session_state.current_file

if not selected_filename or not os.path.exists(os.path.join(UPLOAD_DIR, selected_filename)):
    st.info("👈 請先從左側檔案大廳上傳或選擇要分析的 Excel 檔案。")
else:
    file_path = os.path.join(UPLOAD_DIR, selected_filename)

    st.sidebar.markdown("---")
    st.sidebar.header("⚙️ 2. 圖表數量設定")
    
    num_charts = st.sidebar.selectbox(
        "請選擇同時顯示的圖表數量：", 
        options=[1, 2, 3, 4], 
        index=0,
        key="num_charts_select"
    )

    st.subheader(f"📌 當前分析檔案：`{selected_filename}` (同時顯示 {num_charts} 個圖表視角)")

    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names

        cols = st.columns(num_charts)

        for idx in range(num_charts):
            with cols[idx]:
                st.markdown(f"### 📊 圖表視角 {idx+1}")
                
                with st.expander(f"⚙️ 設定視角 {idx+1} 參數", expanded=True):
                    sheet_key = f"sheet_{idx}"
                    if sheet_key in st.session_state and st.session_state[sheet_key] not in sheet_names:
                        st.session_state[sheet_key] = sheet_names[0]

                    chosen_sheet = st.selectbox(
                        "工作表 (Sheet)", 
                        options=sheet_names, 
                        key=sheet_key
                    )
                    
                    df, _, _ = load_and_clean_excel(file_path, chosen_sheet)
                    cols_list = df.columns.tolist()[:15]

                    if cols_list:
                        x_key = f"x_{idx}"
                        if x_key in st.session_state and st.session_state[x_key] not in cols_list:
                            st.session_state[x_key] = cols_list[0]

                        x_axis = st.selectbox(
                            "X 軸（主要類別）", 
                            options=cols_list, 
                            key=x_key
                        )
                        
                        # X 軸範圍選取過濾
                        unique_x_vals = df[x_axis].dropna().unique().tolist()
                        filter_key = f"x_filter_{idx}"
                        
                        if filter_key in st.session_state and isinstance(st.session_state[filter_key], list):
                            valid_vals = [v for v in st.session_state[filter_key] if v in unique_x_vals]
                            st.session_state[filter_key] = valid_vals if valid_vals else unique_x_vals
                        elif filter_key not in st.session_state:
                            st.session_state[filter_key] = unique_x_vals

                        selected_x_vals = st.multiselect(
                            f"🔍 選擇 X 軸 ({x_axis}) 顯示範圍：",
                            options=unique_x_vals,
                            key=filter_key
                        )
                        
                        df_filtered = df[df[x_axis].isin(selected_x_vals)]
                        
                        legend_opts = ["無"] + cols_list
                        leg_key = f"leg_{idx}"
                        if leg_key in st.session_state and st.session_state[leg_key] not in legend_opts:
                            st.session_state[leg_key] = "無"

                        legend_axis = st.selectbox(
                            "圖例（顏色分類）", 
                            options=legend_opts, 
                            key=leg_key
                        )
                        
                        y_opts = ["資料筆數 (Count)"] + cols_list
                        y_key = f"y_{idx}"
                        if y_key in st.session_state and st.session_state[y_key] not in y_opts:
                            st.session_state[y_key] = "資料筆數 (Count)"

                        y_axis = st.selectbox(
                            "統計欄位（Y 軸）",
                            options=y_opts,
                            key=y_key
                        )
                        
                        chart_opts = ["聚集直條圖 (Grouped)", "堆疊直條圖 (Stacked)", "折線圖 (Line)", "圓餅圖 (Pie)"]
                        chart_key = f"chart_{idx}"
                        if chart_key in st.session_state and st.session_state[chart_key] not in chart_opts:
                            st.session_state[chart_key] = chart_opts[0]

                        chart_type = st.selectbox(
                            "圖表類型", 
                            options=chart_opts,
                            key=chart_key
                        )

                # 繪圖 logic
                if cols_list:
                    if not df_filtered.empty:
                        fig, grouped_df = process_and_render_chart(
                            df_filtered, x_axis, legend_axis, y_axis, chart_type, 
                            chart_height=450 if num_charts == 1 else 380
                        )
                        
                        st.plotly_chart(fig, use_container_width=True)

                        # 📌 按鈕：將繪製好的圖表加入常用資料夾
                        with st.popover("⭐ 將此圖表存入常用資料夾"):
                            default_fav_title = f"{chosen_sheet}-{x_axis}"
                            fav_title = st.text_input("常用圖表名稱：", value=default_fav_title, key=f"fav_title_in_{idx}")
                            if st.button("💾 確認儲存", key=f"btn_save_fav_{idx}", use_container_width=True):
                                if fav_title.strip():
                                    current_favs = load_favorites()
                                    new_fav = {
                                        "name": fav_title.strip(),
                                        "file": selected_filename,
                                        "sheet": chosen_sheet,
                                        "x_axis": x_axis,
                                        "selected_x_vals": selected_x_vals,
                                        "legend_axis": legend_axis,
                                        "y_axis": y_axis,
                                        "chart_type": chart_type
                                    }
                                    # 若名稱重複則覆蓋舊的，否則新增
                                    current_favs = [f for f in current_favs if f["name"] != fav_title.strip()]
                                    current_favs.append(new_fav)
                                    save_favorites(current_favs)
                                    st.success(f"🎉 已將「{fav_title.strip()}」存入常用資料夾！")
                                    st.rerun()
                                else:
                                    st.warning("⚠️ 請輸入圖表名稱！")

                        with st.expander("🔍 查看統計數據明細"):
                            st.dataframe(grouped_df, use_container_width=True)
                    else:
                        st.warning("⚠️ 請至少勾選一個 X 軸範圍項目以顯示圖表。")

    except Exception as e:
        st.error(f"讀取或解析檔案時發生錯誤：{e}")

# 8. 側邊欄留白
for _ in range(5):
    st.sidebar.write("")