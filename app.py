import streamlit as st
import pandas as pd
import colorsys
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os
import json
import gc

# 1. 頁面基本設定
st.set_page_config(
    page_title="客訴數據分析平台",
    page_icon="📊",
    layout="wide"
)

if "entered" not in st.session_state:
    st.session_state.entered = False

# 尚未點擊 Enter 時顯示封面頁
if not st.session_state.entered:
    # ---------------------------------------------------------
    # 1. 注入自訂背景圖片 CSS
    # ---------------------------------------------------------
    # 💡 提示：您可以將背景圖片放在程式碼同目錄下，並將下面的網址替換為您的圖片路徑或 URL
    bg_image_url = "https://www.ist4u.com/uploads/index_banner/tw/banner_2.png"
    
    st.markdown(
        f"""
        <style>
        /* 設定主頁面背景圖片與遮罩 */
        .stApp {{
            background: linear-gradient(rgba(255, 255, 255, 0.6), rgba(255, 255, 255, 0.6)), 
                        url("{bg_image_url}");
            background-size: contain;
            background-position: center;
            background-repeat: no-repeat;
            background-attachment: fixed;
            
        }}
        
        /* 隱藏封面頁的預設頁首/頁尾 (可選，讓畫面更乾淨) */
        header {{visibility: hidden;}}
        footer {{visibility: hidden;}}
        
        </style>
        """,
        unsafe_allow_html=True
    )

    # 2. 最頂部全寬跑馬燈 (連續顯示三個)
    st.markdown("""
        <div style="background-color: transparent; padding: 10px 0; margin-bottom: 20px;">
            <marquee behavior="scroll" direction="left" scrollamount="7" style="color: #1E293B; font-size: 16px; font-weight: bold; font-family: 'Microsoft JhengHei';">
                Data Analytics & Quality Dashboard @ IST QA TEAM &nbsp;&nbsp;&nbsp;&nbsp;•&nbsp;&nbsp;&nbsp;&nbsp; Data Analytics & Quality Dashboard @ IST QA TEAM &nbsp;&nbsp;&nbsp;&nbsp;•&nbsp;&nbsp;&nbsp;&nbsp; Data Analytics & Quality Dashboard @ IST QA TEAM
            </marquee>
        </div>
        """, unsafe_allow_html=True)

    # 垂直間距
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    # 3. 中間標題區塊
    _, col_title, _ = st.columns([1, 2, 1])
    with col_title:
        st.markdown(
            """
            <div style='text-align: center;'>
                <h1 style='font-size: 3.5rem; margin-bottom: 10px; color: #0F172A; text-shadow: 1px 1px 2px rgba(0,0,0,0.1);'>客訴數據分析平台</h1>
                <p style='font-size: 1.8rem; color: #475569; font-weight: 500;'>
                    Data Analytics & Quality Dashboard
                </p>
            </div>
            """, 
            unsafe_allow_html=True
        )
        
   # 4. 推到底部的垂直間距（增加 <br> 數量讓按鈕再往下移動）
    st.markdown("<br><br><br><br><br><br><br><br>", unsafe_allow_html=True)
    
    # 5. 下方小型居中 Enter 按鈕
    _, col_btn, _ = st.columns([3, 1, 3])
    with col_btn:
        if st.button("Enter", use_container_width=True, type="primary"):
            st.session_state.entered = True
            st.rerun()

# 點擊 Enter 後顯示原本的儀表板內容
else:
    # (保持原本 else 裡面的程式碼即可)

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
    st.sidebar.caption("MAIN MENU (主選單)")
    app_mode = st.sidebar.radio(
        "主選單",
        options=["🖥️ Overall View (總覽看板)", "📊 Data Analysis (繪圖)"],
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
        "➕ Upload New File", 
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
    if app_mode == "📊 Data Analysis (繪圖)":
        st.sidebar.markdown("---")
        st.sidebar.caption("VIEW CONFIG (視角設定)")
        num_charts = st.sidebar.selectbox("顯示視角數量：", options=[1, 2, 3, 4], index=0)
    else:
        num_charts = 1

    st.sidebar.divider()

    # --- FAVORITE CHARTS 區塊 ---
    st.sidebar.caption("FAVORITE CHARTS (常用庫)")

    # 💡 注入 CSS 讓按鈕內的文字/符號完美居中，並移除多餘內邊距
    st.sidebar.markdown("""
        <style>
        /* 針對 sidebar 中的小型按鈕強制置中 */
        div[data-testid="stSidebar"] button {
            padding: 0px !important;
            display: flex !important;
            justify-content: center !important;
            align-items: center !important;
            text-align: center !important;
            min-height: 38px !important;
        }
        </style>
    """, unsafe_allow_html=True)

    if not st.session_state.fav_list:
        st.sidebar.caption("💡 點擊圖表下方「⭐ 加入常用」即可在此快速載入。")
    else:
        for idx_fav, fav_item in enumerate(st.session_state.fav_list):
            # 建立 4 個小欄位：[上移, 下移, 刪除, 圖表名稱按鈕]
            col_up, col_down, col_f_del, col_f_btn = st.sidebar.columns([0.8, 0.8, 0.8, 4.6])
            
            # 1. 上移按鈕 (第一個項目時不顯示或停用)
            if idx_fav > 0:
                if col_up.button("▲", key=f"up_fav_{idx_fav}", help="向上移動"):
                    # 交換位置
                    st.session_state.fav_list[idx_fav], st.session_state.fav_list[idx_fav - 1] = \
                        st.session_state.fav_list[idx_fav - 1], st.session_state.fav_list[idx_fav]
                    save_favorites(st.session_state.fav_list)
                    st.rerun()

            # 2. 下移按鈕 (最後一個項目時不顯示或停用)
            if idx_fav < len(st.session_state.fav_list) - 1:
                if col_down.button("▼", key=f"down_fav_{idx_fav}", help="向下移動"):
                    # 交換位置
                    st.session_state.fav_list[idx_fav], st.session_state.fav_list[idx_fav + 1] = \
                        st.session_state.fav_list[idx_fav + 1], st.session_state.fav_list[idx_fav]
                    save_favorites(st.session_state.fav_list)
                    st.rerun()

            # 3. 刪除按鈕
            if col_f_del.button("✕", key=f"del_fav_{idx_fav}", help="刪除"):
                st.session_state.fav_list.pop(idx_fav)
                save_favorites(st.session_state.fav_list)
                st.rerun()

            # 4. 點擊載入圖表設定按鈕
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

    def get_weight_based_colors(categories, values):
        """根據數據比重 (values) 動態生成顏色對照表：

        - 色相 (Hue)：將 360 度均勻切分，確保各項目顏色不重複 - 飽和度 & 亮度：比重越大越鮮豔深濃，比重越小越淡/亮
        """
        num_items = len(categories)
        val_arr = np.array(values, dtype=float)

        val_min = val_arr.min() if len(val_arr) > 0 else 0
        val_max = val_arr.max() if len(val_arr) > 0 else 1
        val_range = val_max - val_min if val_max != val_min else 1

        color_map = {}

        for i, (cat, val) in enumerate(zip(categories, values)):
            # 1. 色相：均勻間隔旋轉，保證不重複
            hue = (i * (360 / max(num_items, 1))) / 360.0

            # 2. 計算比重相對權重 (0.0 ~ 1.0)
            norm_weight = (val - val_min) / val_range

            # 3. 比重大 -> 高飽和度 (0.85)，比重小 -> 低飽和度/淡色 (0.35)
            saturation = 0.35 + (norm_weight * 0.50)

            # 4. 比重大 -> 顏色較濃深 (Lightness 0.45)，比重小 -> 顏色較明亮 (Lightness 0.75)
            lightness = 0.75 - (norm_weight * 0.30)

            # HLS 轉 RGB 並轉 HEX 格式
            r, g, b = colorsys.hls_to_rgb(hue, lightness, saturation)
            hex_color = f"#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}"
            color_map[str(cat)] = hex_color

        return color_map


    def process_and_render_chart(
        df, x_axis, legend_axis, y_axis, chart_type, chart_height=380
    ):
        calc_df = df.copy()
        calc_df[x_axis] = calc_df[x_axis].astype(str)
        color_col = None if legend_axis == "無" else legend_axis

        if color_col and color_col in calc_df.columns:
            calc_df[color_col] = calc_df[color_col].astype(str)

        exact_defect_col = "不良數量" if "不良數量" in calc_df.columns else None

        if y_axis == "不良數量" and exact_defect_col:
            calc_df[exact_defect_col] = (
                pd.to_numeric(
                    calc_df[exact_defect_col].astype(str).str.extract(r"(\d+)")[0],
                    errors="coerce",
                )
                .fillna(0)
            )
            target_col = exact_defect_col
            y_metric = "不良數量 (總和)"
        else:
            target_col = None
            y_metric = "資料筆數"

        # 資料聚合
        if target_col:
            if color_col and x_axis != color_col:
                grouped_df = calc_df.groupby(
                    [x_axis, color_col], as_index=False, observed=False
                )[target_col].sum()
                grouped_df.rename(columns={target_col: y_metric}, inplace=True)
            else:
                grouped_df = calc_df.groupby(
                    [x_axis], as_index=False, observed=False
                )[target_col].sum()
                grouped_df.rename(columns={target_col: y_metric}, inplace=True)
                color_col = x_axis
        else:
            if color_col and x_axis != color_col:
                grouped_df = (
                    calc_df.groupby([x_axis, color_col], as_index=False, observed=False)
                    .size()
                    .rename(columns={"size": y_metric})
                )
            else:
                grouped_df = (
                    calc_df.groupby([x_axis], as_index=False, observed=False)
                    .size()
                    .rename(columns={"size": y_metric})
                )
                color_col = x_axis

        # 🎯 針對直條圖/折線圖進行數字優先的正向排序
        def extract_num(series):
            extracted = series.astype(str).str.extract(r"(\d+)")[0]
            return pd.to_numeric(extracted, errors="coerce")

        sort_cols = [x_axis]
        if color_col and color_col != x_axis:
            sort_cols.append(color_col)

        grouped_df.sort_values(by=sort_cols, key=extract_num, inplace=True)

        # 建立類別順序字典
        cat_orders = {}
        if x_axis in grouped_df.columns:
            cat_orders[x_axis] = sorted(
                grouped_df[x_axis].unique(),
                key=lambda v: pd.to_numeric(
                    "".join(filter(str.isdigit, str(v))), errors="coerce"
                )
                if any(c.isdigit() for c in str(v))
                else v,
            )
        if color_col and color_col in grouped_df.columns:
            cat_orders[color_col] = sorted(
                grouped_df[color_col].unique(),
                key=lambda v: pd.to_numeric(
                    "".join(filter(str.isdigit, str(v))), errors="coerce"
                )
                if any(c.isdigit() for c in str(v))
                else v,
            )

        # 🎨 系統動態自動計算配色（依據數據比重總和調整飽和度與亮度）
        color_target_col = color_col if color_col else x_axis
        weight_summary = (
            grouped_df.groupby(color_target_col)[y_metric]
            .sum()
            .sort_values(ascending=False)
        )

        DYNAMIC_COLOR_MAP = get_weight_based_colors(
            weight_summary.index.tolist(), weight_summary.values.tolist()
        )

        display_legend_title = color_col if legend_axis != "無" else ""

        if chart_type == "聚集直條圖 (Grouped)" and legend_axis != "無":
            pivot_df = grouped_df.pivot(
                index=x_axis, columns=color_col, values=y_metric
            )
            fig = go.Figure()

            for cat in pivot_df.columns:
                y_data = pivot_df[cat]
                color = DYNAMIC_COLOR_MAP.get(str(cat), "#3B82F6")
                fig.add_trace(
                    go.Bar(
                        x=pivot_df.index,
                        y=y_data,
                        name=str(cat),
                        marker_color=color,
                        text=[
                            f"{int(val):,}" if pd.notnull(val) and val > 0 else ""
                            for val in y_data
                        ],
                        textposition="auto",
                    )
                )
            fig.update_layout(
                barmode="group",
                bargap=0.2,
                bargroupgap=0.03,
                font=dict(family="Microsoft JhengHei", size=12),
                height=chart_height,
                margin=dict(l=10, r=10, t=30, b=10),
                xaxis_title=x_axis,
                yaxis_title=y_metric,
                legend_title=display_legend_title,
            )
            fig.update_xaxes(type="category")
            return fig, grouped_df

        elif chart_type == "圓餅圖 (Pie)":
            pie_df = grouped_df.sort_values(by=y_metric, ascending=False)

            fig = px.pie(
                pie_df,
                names=x_axis,
                values=y_metric,
                color=x_axis,
                color_discrete_map=DYNAMIC_COLOR_MAP,
                category_orders={x_axis: pie_df[x_axis].tolist()},
            )

            fig.update_traces(sort=False, direction="clockwise", rotation=0)

            fig.update_layout(
                font=dict(family="Microsoft JhengHei", size=12),
                height=chart_height,
                margin=dict(l=10, r=10, t=30, b=10),
            )
            return fig, pie_df

        else:
            if chart_type == "堆疊直條圖 (Stacked)":
                fig = px.bar(
                    grouped_df,
                    x=x_axis,
                    y=y_metric,
                    color=color_col,
                    barmode="stack",
                    text_auto=True,
                    color_discrete_map=DYNAMIC_COLOR_MAP,
                    category_orders=cat_orders,
                )
            elif chart_type == "折線圖 (Line)":
                fig = px.line(
                    grouped_df,
                    x=x_axis,
                    y=y_metric,
                    color=color_col,
                    markers=True,
                    color_discrete_map=DYNAMIC_COLOR_MAP,
                    category_orders=cat_orders,
                )
            else:
                fig = px.bar(
                    grouped_df,
                    x=x_axis,
                    y=y_metric,
                    color=color_col,
                    text_auto=True,
                    color_discrete_map=DYNAMIC_COLOR_MAP,
                    category_orders=cat_orders,
                )

            fig.update_layout(
                font=dict(family="Microsoft JhengHei", size=12),
                height=chart_height,
                margin=dict(l=10, r=10, t=30, b=10),
                xaxis_title=x_axis,
                yaxis_title=y_metric,
                legend_title=display_legend_title,
            )
            fig.update_xaxes(type="category")
            return fig, grouped_df
    # ---------------------------------------------------------
    # 5. 主頁面切換：Overall View VS 多視角分析
    # ---------------------------------------------------------

    selected_filename = st.session_state.current_file

    if not selected_filename or not os.path.exists(os.path.join(UPLOAD_DIR, selected_filename)):
        st.info("👈 請先從左側 File Repository 上傳或選取 Excel 分析檔案。")

    # --- 模式 A：🖥️ Overall View (總覽看板) ---
    elif app_mode == "🖥️ Overall View (總覽看板)":
        st.title("🖥️ Overall View ")
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
                    idx = i + j  # 當前圖表的唯一索引
                    if idx < len(fav_items):
                        fav = fav_items[idx]
                        with grid_cols[j]:
                            st.subheader(f"📌 {fav['name']}")
                            try:
                                df, _, _ = load_and_clean_excel(file_path, fav['sheet'])
                                
                                x_axis_raw = fav.get("x_axis", "")
                                # 取得組合欄位中的各個子欄位
                                x_cols = [c.strip() for c in x_axis_raw.split(" / ")]
                                valid_x_cols = [c for c in x_cols if c in df.columns]

                                if valid_x_cols:
                                    df_filtered = df.copy()

                                    # 1. 動態建立組合 X 軸欄位 (若為單一欄位也會妥善處理)
                                    df_filtered[x_axis_raw] = (
                                        df_filtered[valid_x_cols]
                                        .fillna("")
                                        .astype(str)
                                        .apply(lambda row: " / ".join([str(item) for item in row]), axis=1)
                                    )

                                    # 2. 🎯 關鍵修復：將 selected_x_vals 與欄位內容皆「強制轉成字串」再比對，避免型態不匹配
                                    selected_x_vals = [str(v) for v in fav.get('selected_x_vals', [])]
                                    if selected_x_vals:
                                        df_filtered = df_filtered[df_filtered[x_axis_raw].astype(str).isin(selected_x_vals)]

                                    # 3. 繪製圖表
                                    if not df_filtered.empty:
                                        fig, _ = process_and_render_chart(
                                            df_filtered, 
                                            x_axis_raw, 
                                            fav['legend_axis'], 
                                            fav['y_axis'], 
                                            fav['chart_type'],
                                            chart_height=350
                                        )
                                        
                                        # 渲染圖表
                                        st.plotly_chart(fig, use_container_width=True)
                                        
                                        # 📝 圖片說明 / 報告備註輸入框
                                        note_key = f"overall_note_{idx}_{fav['name']}"
                                        initial_note = fav.get("note", st.session_state.get(note_key, ""))
                                        
                                        note_text = st.text_area(
                                            label="📝 圖片說明",
                                            value=initial_note,
                                            placeholder="請在此輸入...",
                                            key=note_key,
                                            height=90
                                        )
                                        
                                        # 即時更新回 session_state 與檔案
                                        if st.session_state.fav_list[idx].get("note") != note_text:
                                            st.session_state.fav_list[idx]["note"] = note_text
                                            save_favorites(st.session_state.fav_list)
                                    else:
                                        st.warning(f"⚠️ `{fav['name']}` 篩選後無數據，無法生成圖表。")

                                else:
                                    st.error(f"欄位 `{fav['x_axis']}` 在目前的工作表中不存在。")
                            except Exception as e:
                                st.error(f"無法載入圖表 `{fav['name']}`: {e}")
    # --- 模式 B：📊 多視角分析儀表板 ---
    else:
        st.title("📊 Data Analysis")
        file_path = os.path.join(UPLOAD_DIR, selected_filename)

        FILTER_X_COLS = ["Month", "代理商", "客戶", "產品概述", "責任歸屬1", "Lot no.", "問題分類"]
        TARGET_LEGEND_COLS = ["代理商", "客戶", "Month", "產品別", "責任歸屬", "責任歸屬1", "問題分類"]

        try:
            with pd.ExcelFile(file_path) as xls:
                sheet_names = xls.sheet_names

            for idx in range(num_charts):
                st.subheader(f"圖表視角 {idx+1}")
                
                with st.expander(f"⚙️ 設定視角 {idx+1} 基本數據源", expanded=True):
                    sheet_key = f"sheet_{idx}"
                    if sheet_key in st.session_state and st.session_state[sheet_key] not in sheet_names:
                        st.session_state[sheet_key] = sheet_names[0]

                    chosen_sheet = st.selectbox("工作表 (Sheet)", options=sheet_names, key=sheet_key)
                    df, _, _ = load_and_clean_excel(file_path, chosen_sheet)
                    actual_cols = df.columns.tolist()

                chart_col, filter_col = st.columns([7, 3])

                with filter_col:
                    st.markdown("### X軸(欄位篩選器)")

                    df_filtered = df.copy()
                    active_x_cols = []  # 記錄哪些欄位有被使用者選取

                    # 1. 進行欄位篩選與狀態保持 (避免頁面切換被 Streamlit 自動清空)
                    for col_name in FILTER_X_COLS:
                        if col_name in actual_cols:
                            avail_opts = sorted([str(v) for v in df[col_name].dropna().unique().tolist() if str(v).strip() != ""])
                            
                            # 🔑 用於真正持久儲存的變數名稱
                            saved_state_key = f"saved_filter_{col_name}_{idx}"
                            # 🔑 multiselect 暫存使用名稱
                            widget_key = f"widget_filter_{col_name}_{idx}"

                            # 初始化持久狀態
                            if saved_state_key not in st.session_state:
                                st.session_state[saved_state_key] = []

                            # 校正：剔除不在目前選項內的無效值
                            st.session_state[saved_state_key] = [
                                v for v in st.session_state[saved_state_key] if v in avail_opts
                            ]

                            # 定義 callback：當使用者操作元件時，同步寫回持久狀態
                            def update_filter(s_key=saved_state_key, w_key=widget_key):
                                st.session_state[s_key] = st.session_state[w_key]

                            selected_vals = st.multiselect(
                                f"📌 {col_name}", 
                                options=avail_opts,
                                default=st.session_state[saved_state_key], # 從持久狀態讀取預設值
                                key=widget_key,
                                on_change=update_filter # 異動時即時寫回
                            )
                            
                            if selected_vals:
                                df_filtered = df_filtered[df_filtered[col_name].astype(str).isin(selected_vals)]
                                active_x_cols.append(col_name)

                    st.markdown("---")
                    st.markdown("### 其他圖表設定")

                    # 2. 自動生成組合 X 軸
                    if active_x_cols:
                        x_axis = " / ".join(active_x_cols)
                        
                        df_filtered[x_axis] = (
                            df_filtered[active_x_cols]
                            .fillna("")
                            .astype(str)
                            .apply(lambda row: " / ".join([str(item) for item in row]), axis=1)
                        )
                        current_selected_x_vals = sorted(df_filtered[x_axis].dropna().unique().tolist())
                    else:
                        x_axis = ""
                        current_selected_x_vals = []

                    legend_raw_opts = [col for col in TARGET_LEGEND_COLS if col in actual_cols] or TARGET_LEGEND_COLS
                    legend_opts = ["無"] + list(dict.fromkeys(legend_raw_opts))

                    leg_key = f"leg_{idx}"
                    legend_axis = st.selectbox("🎨 圖例 (分類)", options=legend_opts, key=leg_key)

                    y_opts = ["資料筆數", "不良數量"]
                    y_axis = st.selectbox("📊 Y 軸 (統計項目)", options=y_opts, key=f"y_{idx}")

                    chart_opts = ["聚集直條圖 (Grouped)", "堆疊直條圖 (Stacked)", "折線圖 (Line)", "圓餅圖 (Pie)"]
                    chart_type = st.selectbox("📈 圖表類型", options=chart_opts, key=f"chart_{idx}")

                # -------------------------------------------------------------
                # 📌 左側面板：圖表繪製區
                # -------------------------------------------------------------
                with chart_col:
                    if x_axis and not df_filtered.empty:
                        fig, grouped_df = process_and_render_chart(
                            df_filtered, x_axis, legend_axis, y_axis, chart_type, 
                            chart_height=480
                        )
                        st.plotly_chart(fig, use_container_width=True)

                        with st.popover("⭐ 將此圖表存入常用資料夾"):
                            default_fav_title = f"{x_axis}-{y_axis}({legend_axis})" if legend_axis != "無" else f"{x_axis}-{y_axis}"
                            fav_title = st.text_input("常用圖表名稱：", value=default_fav_title, key=f"fav_title_in_{idx}")
                            
                            if st.button("💾 確認儲存", key=f"btn_save_fav_{idx}", use_container_width=True):
                                if fav_title.strip():
                                    new_fav = {
                                        "name": fav_title.strip(),
                                        "file": selected_filename,
                                        "sheet": chosen_sheet,
                                        "x_axis": x_axis,
                                        "selected_x_vals": current_selected_x_vals,
                                        "legend_axis": legend_axis,
                                        "y_axis": y_axis,
                                        "chart_type": chart_type,
                                        "note": ""
                                    }
                                    st.session_state.fav_list = [f for f in st.session_state.fav_list if f["name"] != fav_title.strip()]
                                    st.session_state.fav_list.append(new_fav)
                                    save_favorites(st.session_state.fav_list)
                                    st.success(f"🎉 已儲存「{fav_title.strip()}」！")
                                    st.rerun()

                        with st.expander("🔍 查看統計數據明細"):
                            st.dataframe(grouped_df, use_container_width=True)
                    else:
                        st.info("請在右側面板的欄位篩選器中**勾選至少一個欄位的內容**來生成圖表。")

                st.markdown("---")

        except Exception as e:
            st.error(f"檔案解析失敗：{e}")