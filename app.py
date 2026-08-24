import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import json
import gc

# 1. 頁面基本設定
st.set_page_config(
    page_title="客訴 List 數據分析平台",
    page_icon="📊",
    layout="wide"
)

# 2. 檔案與常用資料夾儲存機制
UPLOAD_DIR = "uploaded_files"
FAVORITES_FILE = "favorites.json"

if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR, exist_ok=True)

def load_favorites():
    if os.path.exists(FAVORITES_FILE):
        try:
            with open(FAVORITES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_favorites(fav_list):
    try:
        with open(FAVORITES_FILE, "w", encoding="utf-8") as f:
            json.dump(fav_list, f, ensure_ascii=False, indent=4)
    except Exception as e:
        st.error(f"儲存常用時發生錯誤：{e}")

if "fav_list" not in st.session_state:
    st.session_state.fav_list = load_favorites()

if "current_file" not in st.session_state:
    st.session_state.current_file = None

# 用於清空上傳元件的 key 計數器
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# ---------------------------------------------------------
# 3. 側邊欄 (Sidebar) 結構 — Main Menu & File List 佈局
# ---------------------------------------------------------

# --- MAIN MENU 區塊 ---
st.sidebar.caption("MAIN MENU")
app_mode = st.sidebar.radio(
    "主選單",
    options=["🖥️ Overall View (總覽看板)", "📊 多視角分析儀表板"],
    index=0,
    label_visibility="collapsed"
)

st.sidebar.divider()

# --- FILE REPOSITORY 區塊 ---
st.sidebar.caption("FILE REPOSITORY (檔案庫)")

saved_files = [
    f for f in os.listdir(UPLOAD_DIR) 
    if f.endswith(('.xlsx', '.xls', '.xlsm')) and not f.startswith('~$')
]

uploaded_file = st.sidebar.file_uploader(
    "➕ 上傳新檔案", 
    type=["xlsx", "xls", "xlsm"], 
    key=f"uploader_{st.session_state.uploader_key}"
)

if uploaded_file is not None:
    file_path = os.path.join(UPLOAD_DIR, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    st.session_state.current_file = uploaded_file.name
    st.session_state.uploader_key += 1
    st.rerun()

if saved_files and (st.session_state.current_file not in saved_files):
    st.session_state.current_file = saved_files[0]

# 檔案選擇清單
if saved_files:
    for filename in saved_files:
        col_del, col_select = st.sidebar.columns([1, 5])
        
        if col_del.button("✕", key=f"del_{filename}", help=f"刪除 {filename}"):
            file_to_delete = os.path.join(UPLOAD_DIR, filename)
            if st.session_state.current_file == filename:
                st.session_state.current_file = None
            gc.collect()
            if os.path.exists(file_to_delete):
                try:
                    os.remove(file_to_delete)
                    st.rerun()
                except Exception as e:
                    st.sidebar.error(f"刪除失敗: {e}")

        is_selected = (filename == st.session_state.current_file)
        btn_label = f"🟢 {filename}" if is_selected else f"⚪ {filename}"
        
        if col_select.button(btn_label, key=f"select_{filename}", use_container_width=True):
            st.session_state.current_file = filename
            st.rerun()

# VIEW CONFIG 整合進 FILE REPOSITORY 區塊內部
if app_mode == "📊 多視角分析儀表板":
    st.sidebar.markdown("---")
    st.sidebar.caption("VIEW CONFIG")
    num_charts = st.sidebar.selectbox("顯示視角數量：", options=[1, 2, 3, 4], index=0)
else:
    num_charts = 1

st.sidebar.divider()

# --- FAVORITE CHARTS 區塊 ---
st.sidebar.caption("FAVORITE CHARTS (常用庫)")

if not st.session_state.fav_list:
    st.sidebar.caption("💡 點擊圖表下方「⭐ 加入常用」即可在此快速載入。")
else:
    for idx_fav, fav_item in enumerate(st.session_state.fav_list):
        col_f_del, col_f_btn = st.sidebar.columns([1, 5])
        if col_f_del.button("✕", key=f"del_fav_{idx_fav}"):
            st.session_state.fav_list.pop(idx_fav)
            save_favorites(st.session_state.fav_list)
            st.rerun()
        if col_f_btn.button(f"📌 {fav_item['name']}", key=f"load_fav_{idx_fav}", use_container_width=True):
            st.session_state.current_file = fav_item["file"]
            st.session_state["sheet_0"] = fav_item["sheet"]
            st.session_state["x_0"] = fav_item["x_axis"]
            st.session_state["x_filter_0"] = fav_item["selected_x_vals"]
            st.session_state["leg_0"] = fav_item["legend_axis"]
            st.session_state["y_0"] = fav_item["y_axis"]
            st.session_state["chart_0"] = fav_item["chart_type"]
            st.rerun()

# ---------------------------------------------------------
# 4. 數據讀取與繪圖處理
# ---------------------------------------------------------
@st.cache_data(ttl=3600)
def load_and_clean_excel(file_path, sheet_name=None):
    with pd.ExcelFile(file_path) as xls:
        sheet_names = xls.sheet_names
        if not sheet_name or sheet_name not in sheet_names:
            sheet_name = sheet_names[0]
            for name in sheet_names:
                if "客" in name or "list" in name.lower():
                    sheet_name = name
                    break

        # 以 header=1 (即 Excel 第 2 列為欄位標題) 讀取資料
        df = pd.read_excel(xls, sheet_name=sheet_name, header=1)
        df.columns = [str(col).strip() for col in df.columns]
        valid_cols = [col for col in df.columns if col != "" and not col.startswith("Unnamed") and col.lower() != "nan"]
        
        if len(valid_cols) == 0:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
            df.columns = [str(col).strip() for col in df.columns]
            valid_cols = [col for col in df.columns if col != "" and not col.startswith("Unnamed") and col.lower() != "nan"]

        df = df[valid_cols].dropna(how='all')

        for col in df.columns:
            if pd.api.types.is_float_dtype(df[col]):
                non_nulls = df[col].dropna()
                if (non_nulls % 1 == 0).all():
                    df[col] = df[col].astype("Int64")

        return df, sheet_names, sheet_name

def process_and_render_chart(df, x_axis, legend_axis, y_axis, chart_type, chart_height=380):
    calc_df = df.copy()
    calc_df[x_axis] = calc_df[x_axis].astype(str)
    color_col = None if legend_axis == "無" else legend_axis

    if color_col and color_col in calc_df.columns:
        calc_df[color_col] = calc_df[color_col].astype(str)

    exact_defect_col = "不良數量" if "不良數量" in calc_df.columns else None

    if y_axis == "不良數量" and exact_defect_col:
        calc_df[exact_defect_col] = pd.to_numeric(
            calc_df[exact_defect_col].astype(str).str.extract(r'(\d+)')[0], 
            errors='coerce'
        ).fillna(0)
        target_col = exact_defect_col
        y_metric = "不良數量 (總和)"
    else:
        target_col = None
        y_metric = "資料筆數"

    if target_col:
        if color_col and x_axis != color_col:
            grouped_df = calc_df.groupby([x_axis, color_col], as_index=False)[target_col].sum()
            grouped_df.rename(columns={target_col: y_metric}, inplace=True)
        else:
            grouped_df = calc_df.groupby([x_axis], as_index=False)[target_col].sum()
            grouped_df.rename(columns={target_col: y_metric}, inplace=True)
            color_col = None
    else:
        if color_col and x_axis != color_col:
            grouped_df = calc_df.groupby([x_axis, color_col], as_index=False).size().rename(columns={"size": y_metric})
        else:
            grouped_df = calc_df.groupby([x_axis], as_index=False).size().rename(columns={"size": y_metric})
            color_col = None

    # 🎨 定義圖中對應的專屬顏色對照表
    CUSTOM_COLOR_MAP = {
        "EOS/ESD": "#0066CC",                   # 深藍
        "溫校異常": "#66C2FF",                 # 天藍
        "系統應用": "#FF1A3C",                 # 鮮紅
        "供應商": "#FFA8B6",                   # 粉紅
        "Sample 受損": "#00B386",             # 綠/青綠
        "Design problem": "#66E6A3",          # 淺綠
        "其他": "#FF8033",                     # 橘色
        "測試coverage": "#FFC658",             # 黃色
        "組裝製程_(壓合、燒錄等)": "#6A3AAD"      # 紫色
    }

    # 預設備用通用調色盤
    color_palette = [
        "#0066CC", "#66C2FF", "#FF1A3C", "#FFA8B6", 
        "#00B386", "#66E6A3", "#FF8033", "#FFC658", "#6A3AAD"
    ]

    # 🎯 針對直條圖/折線圖進行數字優先的正向排序
    def extract_num(series):
        extracted = series.astype(str).str.extract(r'(\d+)')[0]
        return pd.to_numeric(extracted, errors='coerce')

    sort_cols = [x_axis]
    if color_col:
        sort_cols.append(color_col)
    
    grouped_df.sort_values(
        by=sort_cols, 
        key=extract_num,
        inplace=True
    )

    # 建立類別順序字典
    cat_orders = {}
    if x_axis in grouped_df.columns:
        cat_orders[x_axis] = sorted(grouped_df[x_axis].unique(), key=lambda v: pd.to_numeric(''.join(filter(str.isdigit, str(v))), errors='coerce') if any(c.isdigit() for c in str(v)) else v)
    if color_col and color_col in grouped_df.columns:
        cat_orders[color_col] = sorted(grouped_df[color_col].unique(), key=lambda v: pd.to_numeric(''.join(filter(str.isdigit, str(v))), errors='coerce') if any(c.isdigit() for c in str(v)) else v)

    if chart_type == "聚集直條圖 (Grouped)" and color_col:
        pivot_df = grouped_df.pivot(index=x_axis, columns=color_col, values=y_metric)
        fig = go.Figure()
        
        num_colors = len(color_palette)
        for i, cat in enumerate(pivot_df.columns):
            y_data = pivot_df[cat]
            color = CUSTOM_COLOR_MAP.get(str(cat), color_palette[i % num_colors])
            fig.add_trace(go.Bar(
                x=pivot_df.index,
                y=y_data,
                name=str(cat),
                marker_color=color,
                text=[f"{int(val):,}" if pd.notnull(val) and val > 0 else "" for val in y_data],
                textposition='auto'
            ))
        fig.update_layout(
            barmode='group', bargap=0.2, bargroupgap=0.03,
            font=dict(family="Microsoft JhengHei", size=12),
            height=chart_height, margin=dict(l=10, r=10, t=30, b=10),
            xaxis_title=x_axis, yaxis_title=y_metric, legend_title=color_col
        )
        fig.update_xaxes(type='category')
        return fig, grouped_df

    elif chart_type == "圓餅圖 (Pie)":
        # 🎯 圓餅圖專屬處理：依數值從大到小排序，確保從 12 點鐘方向順時針依比例排列
        pie_df = grouped_df.sort_values(by=y_metric, ascending=False)
        
        fig = px.pie(
            pie_df, 
            names=x_axis, 
            values=y_metric,
            color=x_axis,
            color_discrete_map=CUSTOM_COLOR_MAP,
            color_discrete_sequence=color_palette,
            category_orders={x_axis: pie_df[x_axis].tolist()}
        )
        
        # 設定順時針排列與起始角度
        fig.update_traces(
            sort=False,            # 停用預設內部自動排序，採用傳入的排序
            direction='clockwise', # 順時針方向排列
            rotation=0             # 0度從正上方 (12 點鐘方向) 開始
        )

        fig.update_layout(
            font=dict(family="Microsoft JhengHei", size=12),
            height=chart_height, margin=dict(l=10, r=10, t=30, b=10)
        )
        return fig, pie_df

    else:
        if chart_type == "堆疊直條圖 (Stacked)":
            fig = px.bar(
                grouped_df, x=x_axis, y=y_metric, color=color_col, barmode="stack", text_auto=True,
                color_discrete_map=CUSTOM_COLOR_MAP,
                color_discrete_sequence=color_palette,
                category_orders=cat_orders
            )
        elif chart_type == "折線圖 (Line)":
            fig = px.line(
                grouped_df, x=x_axis, y=y_metric, color=color_col, markers=True,
                color_discrete_map=CUSTOM_COLOR_MAP,
                color_discrete_sequence=color_palette,
                category_orders=cat_orders
            )
        else:
            fig = px.bar(
                grouped_df, x=x_axis, y=y_metric, color=color_col, text_auto=True,
                color_discrete_map=CUSTOM_COLOR_MAP,
                color_discrete_sequence=color_palette,
                category_orders=cat_orders
            )

        fig.update_layout(
            font=dict(family="Microsoft JhengHei", size=12),
            height=chart_height, margin=dict(l=10, r=10, t=30, b=10),
            xaxis_title=x_axis, yaxis_title=y_metric, legend_title=color_col if color_col else ""
        )
        fig.update_xaxes(type='category')
        return fig, grouped_df
# ---------------------------------------------------------
# 5. 主頁面切換：Overall View VS 多視角分析
# ---------------------------------------------------------

selected_filename = st.session_state.current_file

if not selected_filename or not os.path.exists(os.path.join(UPLOAD_DIR, selected_filename)):
    st.info("👈 請先從左側 File Repository 上傳或選取 Excel 分析檔案。")

# --- 模式 A：🖥️ Overall View (總覽看板) ---
elif app_mode == "🖥️ Overall View (總覽看板)":
    st.title("🖥️ Overall View 數據總覽看板")
    st.caption(f"📌 當前分析檔案：`{selected_filename}`")
    
    if not st.session_state.fav_list:
        st.warning("⚠️ 目前常用庫中沒有儲存的圖表！請在「📊 多視角分析儀表板」設定圖表並點選「⭐ 加入常用」，圖表將會自動排列出現在這裡。")
    else:
        file_path = os.path.join(UPLOAD_DIR, selected_filename)
        
        cols_per_row = 2
        fav_items = st.session_state.fav_list
        
        for i in range(0, len(fav_items), cols_per_row):
            grid_cols = st.columns(cols_per_row)
            for j in range(cols_per_row):
                if i + j < len(fav_items):
                    fav = fav_items[i + j]
                    with grid_cols[j]:
                        st.subheader(f"📌 {fav['name']}")
                        try:
                            df, _, _ = load_and_clean_excel(file_path, fav['sheet'])
                            
                            if fav['x_axis'] in df.columns:
                                df_filtered = df[df[fav['x_axis']].isin(fav['selected_x_vals'])]
                                
                                fig, _ = process_and_render_chart(
                                    df_filtered, 
                                    fav['x_axis'], 
                                    fav['legend_axis'], 
                                    fav['y_axis'], 
                                    fav['chart_type'],
                                    chart_height=350
                                )
                                st.plotly_chart(fig, use_container_width=True)
                            else:
                                st.error(f"欄位 `{fav['x_axis']}` 在目前的工作表中不存在。")
                        except Exception as e:
                            st.error(f"無法載入圖表 `{fav['name']}`: {e}")

# --- 模式 B：📊 多視角分析儀表板 ---
else:
    st.title("📊 客訴 List 多視角分析儀表板")
    file_path = os.path.join(UPLOAD_DIR, selected_filename)

    # 🎯 指定的的 X 軸與圖例目標欄位名稱清單
    TARGET_X_COLS = ["代理商", "客戶", "產品概述", "Month", "Lot no.", "產品別", "責任歸屬", "責任歸屬1", "問題分類"]
    TARGET_LEGEND_COLS = ["代理商", "客戶", "Month", "產品別", "責任歸屬", "責任歸屬1", "問題分類"]

    try:
        with pd.ExcelFile(file_path) as xls:
            sheet_names = xls.sheet_names

        cols = st.columns(num_charts)

        for idx in range(num_charts):
            with cols[idx]:
                st.subheader(f"📊 圖表視角 {idx+1}")
                
                with st.expander(f"⚙️ 設定視角 {idx+1} 參數", expanded=True):
                    sheet_key = f"sheet_{idx}"
                    if sheet_key in st.session_state and st.session_state[sheet_key] not in sheet_names:
                        st.session_state[sheet_key] = sheet_names[0]

                    chosen_sheet = st.selectbox("工作表 (Sheet)", options=sheet_names, key=sheet_key)
                    df, _, _ = load_and_clean_excel(file_path, chosen_sheet)
                    
                    actual_cols = df.columns.tolist()
                    
                    # 🎯 只保留 Excel 中實際存在的指定欄位 (兼具安全性與彈性)
                    x_opts = [col for col in TARGET_X_COLS if col in actual_cols]
                    if not x_opts:  # 若無名稱完全相符，預設顯示指定的目標清單
                        x_opts = TARGET_X_COLS

                    legend_raw_opts = [col for col in TARGET_LEGEND_COLS if col in actual_cols]
                    if not legend_raw_opts:
                        legend_raw_opts = TARGET_LEGEND_COLS
                        
                    legend_opts = ["無"] + list(dict.fromkeys(legend_raw_opts))

                    if x_opts:
                        x_key = f"x_{idx}"
                        if x_key in st.session_state and st.session_state[x_key] not in x_opts:
                            st.session_state[x_key] = x_opts[0]

                        x_axis = st.selectbox("X 軸（主要類別）", options=x_opts, key=x_key)
                        
                        if x_axis in df.columns:
                            unique_x_vals = df[x_axis].dropna().unique().tolist()
                        else:
                            unique_x_vals = []

                        filter_key = f"x_filter_{idx}"
                        
                        if filter_key not in st.session_state:
                            st.session_state[filter_key] = unique_x_vals
                        elif isinstance(st.session_state[filter_key], list):
                            st.session_state[filter_key] = [v for v in st.session_state[filter_key] if v in unique_x_vals]

                        selected_x_vals = st.multiselect(f"🔍 選擇 X 軸顯示範圍：", options=unique_x_vals, key=filter_key)
                        
                        if x_axis in df.columns:
                            df_filtered = df[df[x_axis].isin(selected_x_vals)]
                        else:
                            df_filtered = pd.DataFrame()
                        
                        leg_key = f"leg_{idx}"
                        if leg_key in st.session_state and st.session_state[leg_key] not in legend_opts:
                            st.session_state[leg_key] = "無"

                        legend_axis = st.selectbox("圖例（顏色分類）", options=legend_opts, key=leg_key)
                        
                        y_opts = ["資料筆數", "不良數量"]
                        y_key = f"y_{idx}"
                        if y_key in st.session_state and st.session_state[y_key] not in y_opts:
                            st.session_state[y_key] = "資料筆數"

                        y_axis = st.selectbox("統計欄位（Y 軸）", options=y_opts, key=y_key)
                        
                        chart_opts = ["聚集直條圖 (Grouped)", "堆疊直條圖 (Stacked)", "折線圖 (Line)", "圓餅圖 (Pie)"]
                        chart_key = f"chart_{idx}"
                        if chart_key in st.session_state and st.session_state[chart_key] not in chart_opts:
                            st.session_state[chart_key] = chart_opts[0]

                        chart_type = st.selectbox("圖表類型", options=chart_opts, key=chart_key)

                if x_opts:
                    if not df_filtered.empty:
                        fig, grouped_df = process_and_render_chart(
                            df_filtered, x_axis, legend_axis, y_axis, chart_type, 
                            chart_height=420 if num_charts == 1 else 360
                        )
                        st.plotly_chart(fig, use_container_width=True)

                        with st.popover("⭐ 將此圖表存入常用資料夾"):
                            if legend_axis and legend_axis != "無":
                                default_fav_title = f"{x_axis}-{y_axis}({legend_axis})"
                            else:
                                default_fav_title = f"{x_axis}-{y_axis}"
                                
                            fav_title = st.text_input("常用圖表名稱：", value=default_fav_title, key=f"fav_title_in_{idx}")
                            
                            if st.button("💾 確認儲存", key=f"btn_save_fav_{idx}", use_container_width=True):
                                if fav_title.strip():
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
                                    st.session_state.fav_list = [f for f in st.session_state.fav_list if f["name"] != fav_title.strip()]
                                    st.session_state.fav_list.append(new_fav)
                                    save_favorites(st.session_state.fav_list)
                                    st.success(f"🎉 已將「{fav_title.strip()}」存入常用資料夾！")
                                    st.rerun()

                        with st.expander("🔍 查看統計數據明細"):
                            st.dataframe(grouped_df, use_container_width=True)
                    else:
                        st.warning("⚠️ 請至少勾選一個 X 軸範圍項目。")

    except Exception as e:
        st.error(f"檔案解析失敗：{e}")