import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io

st.set_page_config(page_title="COGS 9-Box Bubble Chart", layout="wide")

# ==========================================
# HEADER
# ==========================================
st.markdown("""
    <div style="background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); 
                padding: 1.2rem; border-radius: 10px; margin-bottom: 1rem; color: white;
                text-align: center;">
        <h2 style="margin:0; color: #38bdf8; font-weight: 800;">
            🫧 COGS (%) vs GROSS SALES 9-BOX
        </h2>
        <p style="margin: 0.2rem 0 0 0; opacity: 0.9; font-size: 1rem;">
            Advanced SKU Portfolio Evaluation & Price Tracking
        </p>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# FILE UPLOADER & DATA LOADING
# ==========================================
st.markdown("### 📂 Upload Data PL")
uploaded_file = st.file_uploader("Upload file Excel/CSV (e.g., Sample_PL_2026.xlsx)", type=["xlsx", "csv"])

if uploaded_file is None:
    st.info("☝️ Silakan upload file data P&L lu untuk memulai.")
    st.stop()

try:
    file_ext = uploaded_file.name.split('.')[-1].lower()
    has_price_data = False
    pi_produk_keys = set()
    pi_period_dict = {}

    # 1. READ DATA WITH MULTI-SHEET EXCEL SUPPORT
    if file_ext == 'xlsx':
        xls = pd.ExcelFile(uploaded_file)
        if 'PL_2026' in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name='PL_2026')
        else:
            df = pd.read_excel(xls)  # Fallback to first sheet

        # Check and process mapping sheets
        if 'Price Increase' in xls.sheet_names and 'Mapping' in xls.sheet_names:
            df_pi = pd.read_excel(xls, sheet_name='Price Increase')
            df_map = pd.read_excel(xls, sheet_name='Mapping')

            df_pi.columns = df_pi.columns.str.strip()
            df_map.columns = df_map.columns.str.strip()

            # Execute mapping logic via 'Old material number'
            if 'Old material number' in df_pi.columns and 'Old material number' in df_map.columns and 'Produk_Key' in df_map.columns:
                df_pi['Old material number'] = df_pi['Old material number'].astype(str).str.strip().str.replace(r'\.0$',
                                                                                                                '',
                                                                                                                regex=True)
                df_map['Old material number'] = df_map['Old material number'].astype(str).str.strip().str.replace(
                    r'\.0$', '', regex=True)
                df_map['Produk_Key'] = df_map['Produk_Key'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

                # Format Date in Price Increase
                if 'Price Increase Period' in df_pi.columns:
                    df_pi['Price Increase Period'] = pd.to_datetime(df_pi['Price Increase Period'], errors='coerce')

                df_pi_mapped = df_pi.merge(df_map, on='Old material number', how='inner')
                pi_produk_keys = set(df_pi_mapped['Produk_Key'].tolist())
                pi_period_dict = dict(zip(df_pi_mapped['Produk_Key'], df_pi_mapped['Price Increase Period']))
                has_price_data = True
    else:
        # Fallback for old CSV format
        df = pd.read_csv(uploaded_file, sep=';', dtype=str)
        if len(df.columns) == 1:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, sep=',', dtype=str)

    df.columns = df.columns.str.strip()


    # Currency Cleaner Function
    def clean_currency(x):
        if pd.isna(x): return 0.0
        if isinstance(x, (int, float)): return float(x)
        x_str = str(x).strip()
        if x_str in ['-', '']: return 0.0
        x_str = x_str.replace('.', '').replace(',', '.')
        try:
            return float(x_str)
        except:
            return 0.0


    # Ensure required columns exist
    req_cols = ['Date', 'Product Name', 'Gross Sales', 'Return', 'COGS_Regular', 'Total COGS', 'Quantity']
    missing_cols = [c for c in req_cols if c not in df.columns]

    if missing_cols:
        st.error(f"⚠️ Kolom berikut tidak ditemukan di sheet data P&L: {', '.join(missing_cols)}")
        st.stop()

    # Apply cleaner to numbers
    df['Gross Sales'] = df['Gross Sales'].apply(clean_currency)
    df['Return'] = df['Return'].apply(clean_currency)
    df['COGS_Regular'] = df['COGS_Regular'].apply(clean_currency)
    df['Total COGS'] = df['Total COGS'].apply(clean_currency)
    df['Quantity'] = df['Quantity'].apply(clean_currency)

    # Define Gross Sales
    df['Calc_Gross_Sales'] = df['Gross Sales'] + df['Return']

    # 2. PARSE DATE & FILTER JAN 2026 ONWARDS
    if pd.api.types.is_datetime64_any_dtype(df['Date']):
        df['Date'] = pd.to_datetime(df['Date'])
    else:
        df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y', errors='coerce')

    df = df[df['Date'] >= '2026-01-01']

    if df.empty:
        st.warning("⚠️ Tidak ada data transaksi mulai dari Januari 2026 di file ini.")
        st.stop()

    # Map periods
    df['Month_Year'] = df['Date'].dt.strftime('%B %Y')
    df['YYYYMM'] = df['Date'].dt.strftime('%Y%m')

    # Flag SKUs with Price Increase
    if 'Produk_Key' in df.columns:
        df['Produk_Key_Str'] = df['Produk_Key'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        df['Price_Increase_Flag'] = df['Produk_Key_Str'].isin(pi_produk_keys)
        if has_price_data:
            df['Price_Increase_Period'] = df['Produk_Key_Str'].map(pi_period_dict)
        else:
            df['Price_Increase_Period'] = pd.NaT
    else:
        df['Price_Increase_Flag'] = False
        df['Price_Increase_Period'] = pd.NaT

    # ==========================================
    # SETTINGS: PERIOD & METRIC
    # ==========================================
    st.markdown("---")
    st.markdown("### ⚙️ Evaluation Settings")

    col_set1, col_set2, col_set3 = st.columns(3)

    with col_set1:
        cogs_selector = st.selectbox(
            "Y-Axis Metric (COGS Component):",
            ["Total COGS", "COGS_Regular"],
            help="Pilih komponen COGS yang ingin dievaluasi (Sumbu Y)."
        )

    with col_set2:
        period_mode = st.radio(
            "Time Evaluation Mode:",
            ["Monthly", "Year-to-Date (YTD)"],
            horizontal=True
        )

    with col_set3:
        # Sort unique months chronologically
        months_sorted = df[['YYYYMM', 'Month_Year']].drop_duplicates().sort_values('YYYYMM')
        month_list = months_sorted['Month_Year'].tolist()

        if period_mode == "Monthly":
            selected_month = st.selectbox("Select Month:", month_list)
            df_filtered = df[df['Month_Year'] == selected_month].copy()
            chart_period_title = selected_month
        else:
            selected_month = st.selectbox("YTD up to Month:", month_list)
            selected_yyyymm = months_sorted[months_sorted['Month_Year'] == selected_month]['YYYYMM'].values[0]
            df_filtered = df[df['YYYYMM'] <= selected_yyyymm].copy()
            chart_period_title = f"YTD (Jan 2026 - {selected_month.split(' ')[0]})"

    # ==========================================
    # AGGREGATE DATA PER SKU & HARD-FILTER
    # ==========================================
    df_sku = df_filtered.groupby('Product Name').agg(
        Total_Gross_Sales=('Calc_Gross_Sales', 'sum'),
        Total_COGS_Reg=('COGS_Regular', 'sum'),
        Total_COGS_All=('Total COGS', 'sum'),
        Selected_COGS=(cogs_selector, 'sum'),
        Total_Qty=('Quantity', 'sum'),
        Price_Increase=('Price_Increase_Flag', 'max'),
        Price_Increase_Period=('Price_Increase_Period', 'first'),
        Produk_Key_Str=('Produk_Key_Str', 'first')
    ).reset_index()

    # 🧹 HARD-FILTER OUTLIER
    df_sku_valid = df_sku[
        (df_sku['Total_Gross_Sales'] >= 1) &
        (df_sku['Total_COGS_Reg'] >= 1) &
        (df_sku['Total_COGS_All'] >= 1)
        ].copy()

    if df_sku_valid.empty:
        st.warning("⚠️ Tidak ada data valid (Gross Sales, COGS Regular, & Total COGS >= Rp 1) pada periode ini.")
        st.stop()

    # Formatting
    df_sku_valid['COGS (%)'] = (df_sku_valid['Selected_COGS'] / df_sku_valid['Total_Gross_Sales']) * 100
    df_sku_valid['Bubble_Size'] = df_sku_valid['Total_Qty'].abs().replace(0, 1)
    df_sku_valid['Price_Increase_Label'] = np.where(df_sku_valid['Price_Increase'], 'Naik Harga', 'Harga Tetap')
    df_sku_valid['Periode Naik Harga'] = df_sku_valid['Price_Increase_Period'].dt.strftime('%Y-%m-%d').fillna('-')

    df_sku_valid['Format_Sales'] = df_sku_valid['Total_Gross_Sales'].apply(lambda x: f"Rp {x:,.0f}")
    df_sku_valid['Format_COGS'] = df_sku_valid['Selected_COGS'].apply(lambda x: f"Rp {x:,.0f}")
    df_sku_valid['Format_Qty'] = df_sku_valid['Total_Qty'].apply(lambda x: f"{x:,.0f} Pcs")

    # ==========================================
    # SMART-SCALING & OUTLIER CONTROL
    # ==========================================
    st.markdown("#### 🎚️ Smart-Scaling Control")

    col_out1, col_out2 = st.columns(2)
    with col_out1:
        min_outlier_limit_x = st.number_input(
            "MIN Gross Sales Limit for Average Calculation:",
            value=0.0,
            step=1000000.0,
            help="SKUs di bawah limit ini tidak masuk hitungan garis rata-rata (Center Line X), tapi tetap divisualisasikan."
        )

    with col_out2:
        scale_factor_x = st.number_input(
            "Right Box X-Axis Scale (Multiplier):",
            value=0.2,
            step=0.1,
            min_value=0.001,
            max_value=10.0,
            help="Set < 1.0 (misal 0.2) untuk mengkompresi ruang kosong X-Axis bagian kanan."
        )

    # ==========================================
    # MATRIX THRESHOLDS
    # ==========================================
    st.info(
        "💡 **INFO:** The X and Y axes use the dynamic Midpoint value as the absolute center point based on your custom Low and High thresholds. The Center lines are marked in Red.")

    normal_x_df = df_sku_valid[df_sku_valid['Total_Gross_Sales'] >= min_outlier_limit_x]
    robust_avg_x = normal_x_df['Total_Gross_Sales'].mean() if not normal_x_df.empty else df_sku_valid[
        'Total_Gross_Sales'].mean()
    if pd.isna(robust_avg_x) or robust_avg_x <= 0: robust_avg_x = 1000000.0

    actual_avg_y = (df_sku_valid['Selected_COGS'].sum() / df_sku_valid['Total_Gross_Sales'].sum() * 100) if \
    df_sku_valid['Total_Gross_Sales'].sum() > 0 else df_sku_valid['COGS (%)'].mean()
    if pd.isna(actual_avg_y) or actual_avg_y <= 0: actual_avg_y = 50.0

    def_x_low = robust_avg_x * 0.66
    def_x_high = robust_avg_x * 1.33
    def_y_low = actual_avg_y * 0.8
    def_y_high = actual_avg_y * 1.2

    with st.form("cogs_threshold_form"):
        col_t1, col_t2, col_t3, col_t4 = st.columns(4)
        with col_t1:
            x_low_val = st.number_input("X-Axis Low to Med (Gross Sales)", value=float(def_x_low), step=1000000.0)
        with col_t2:
            x_high_val = st.number_input("X-Axis Med to High (Gross Sales)", value=float(def_x_high), step=1000000.0)
        with col_t3:
            y_low_val = st.number_input("Y-Axis Low to Med (COGS %)", value=float(def_y_low), step=1.0)
        with col_t4:
            y_high_val = st.number_input("Y-Axis Med to High (COGS %)", value=float(def_y_high), step=1.0)

        run_chart = st.form_submit_button("▶ RENDER BUBBLE CHART", type="primary")

    # Midpoints
    mid_x = (x_low_val + x_high_val) / 2.0
    mid_y = (y_low_val + y_high_val) / 2.0

    # ==========================================
    # MATRIX ASSIGNMENT & LOGIC
    # ==========================================
    b1_id = 'Box 1 (High COGS, Low Sales)'
    b2_id = 'Box 2 (Med COGS, Low Sales)'
    b3_id = 'Box 3 (Low COGS, Low Sales)'
    b4_id = 'Box 4 (High COGS, Med Sales)'
    b5_id = 'Box 5 (Med COGS, Med Sales)'
    b6_id = 'Box 6 (Low COGS, Med Sales)'
    b7_id = 'Box 7 (High COGS, High Sales)'
    b8_id = 'Box 8 (Med COGS, High Sales)'
    b9_id = 'Box 9 (Low COGS, High Sales)'


    def assign_9box(df_target):
        conditions = [
            (df_target['COGS (%)'] > y_high_val) & (df_target['Total_Gross_Sales'] < x_low_val),
            (df_target['COGS (%)'] >= y_low_val) & (df_target['COGS (%)'] <= y_high_val) & (
                        df_target['Total_Gross_Sales'] < x_low_val),
            (df_target['COGS (%)'] < y_low_val) & (df_target['Total_Gross_Sales'] < x_low_val),
            (df_target['COGS (%)'] > y_high_val) & (df_target['Total_Gross_Sales'] >= x_low_val) & (
                        df_target['Total_Gross_Sales'] <= x_high_val),
            (df_target['COGS (%)'] >= y_low_val) & (df_target['COGS (%)'] <= y_high_val) & (
                        df_target['Total_Gross_Sales'] >= x_low_val) & (df_target['Total_Gross_Sales'] <= x_high_val),
            (df_target['COGS (%)'] < y_low_val) & (df_target['Total_Gross_Sales'] >= x_low_val) & (
                        df_target['Total_Gross_Sales'] <= x_high_val),
            (df_target['COGS (%)'] > y_high_val) & (df_target['Total_Gross_Sales'] > x_high_val),
            (df_target['COGS (%)'] >= y_low_val) & (df_target['COGS (%)'] <= y_high_val) & (
                        df_target['Total_Gross_Sales'] > x_high_val),
            (df_target['COGS (%)'] < y_low_val) & (df_target['Total_Gross_Sales'] > x_high_val)
        ]
        return np.select(conditions, [b1_id, b2_id, b3_id, b4_id, b5_id, b6_id, b7_id, b8_id, b9_id], default=b5_id)


    df_sku_valid['Dynamic 9-Box Category'] = assign_9box(df_sku_valid)
    box_counts = df_sku_valid['Dynamic 9-Box Category'].value_counts().to_dict()
    total_items = len(df_sku_valid)


    # Scaling function
    def apply_custom_x_scale(v):
        if v <= x_high_val:
            return v
        else:
            return x_high_val + (v - x_high_val) * scale_factor_x


    df_sku_valid['X_Plot'] = df_sku_valid['Total_Gross_Sales'].apply(apply_custom_x_scale)

    # Global boundaries setup
    x_min_raw = df_sku_valid['Total_Gross_Sales'].min()
    x_max_raw = df_sku_valid['Total_Gross_Sales'].max()
    x_span_raw = x_max_raw - x_min_raw if x_max_raw != x_min_raw else robust_avg_x
    plot_x_min = x_min_raw - (abs(x_span_raw) * 0.05) if x_min_raw >= 0 else x_min_raw - (abs(x_span_raw) * 0.05)
    plot_x_max = x_max_raw + (abs(x_span_raw) * 0.05)

    y_min_raw = df_sku_valid['COGS (%)'].min()
    y_max_raw = df_sku_valid['COGS (%)'].max()
    y_span_raw = y_max_raw - y_min_raw if y_max_raw != y_min_raw else 100.0
    plot_y_min = y_min_raw - (y_span_raw * 0.05)
    plot_y_max = max(y_max_raw, y_high_val * 1.5) + (y_span_raw * 0.05)


    def fmt_idr(val):
        if val >= 1e9: return f"Rp{val / 1e9:,.1f}Bn"
        if val >= 1e6: return f"Rp{val / 1e6:,.0f}M"
        return f"Rp{val:,.0f}"


    # Helper function to inject common annotations/lines
    def apply_common_layout(fig, custom_box_counts=None):
        b_counts = custom_box_counts if custom_box_counts is not None else box_counts
        tot_items = sum(b_counts.values()) if custom_box_counts is not None else total_items

        fig.add_vline(x=apply_custom_x_scale(x_low_val), line_dash="dash", line_color="black", opacity=0.7,
                      annotation_text=f" Low ({fmt_idr(x_low_val)})", annotation_position="top left")
        fig.add_vline(x=apply_custom_x_scale(x_high_val), line_dash="dash", line_color="black", opacity=0.7,
                      annotation_text=f" High ({fmt_idr(x_high_val)})", annotation_position="top right")
        fig.add_hline(y=y_low_val, line_dash="dash", line_color="black", opacity=0.7,
                      annotation_text=f"Low ({y_low_val:.1f}%)", annotation_position="bottom right")
        fig.add_hline(y=y_high_val, line_dash="dash", line_color="black", opacity=0.7,
                      annotation_text=f"High ({y_high_val:.1f}%)", annotation_position="top right")

        # Center Line Red
        fig.add_vline(x=apply_custom_x_scale(mid_x), line_dash="dot", line_color="red", opacity=0.6,
                      annotation_text=f" ← Center X ({fmt_idr(mid_x)})", annotation_font_color="red",
                      annotation_position="bottom right")
        fig.add_hline(y=mid_y, line_dash="dot", line_color="red", opacity=0.6,
                      annotation_text=f"Center Y ({mid_y:.1f}%)", annotation_position="top right",
                      annotation_font_color="red")

        mapped_x_low, mapped_x_high = apply_custom_x_scale(x_low_val), apply_custom_x_scale(x_high_val)
        mapped_plot_x_min, mapped_plot_x_max = apply_custom_x_scale(plot_x_min), apply_custom_x_scale(plot_x_max)

        b_coords = {
            b1_id: {'x': (mapped_plot_x_min + mapped_x_low) / 2, 'y': (y_high_val + plot_y_max) / 2, 'id': 'B1'},
            b2_id: {'x': (mapped_plot_x_min + mapped_x_low) / 2, 'y': (y_low_val + y_high_val) / 2, 'id': 'B2'},
            b3_id: {'x': (mapped_plot_x_min + mapped_x_low) / 2, 'y': (plot_y_min + y_low_val) / 2, 'id': 'B3'},
            b4_id: {'x': (mapped_x_low + mapped_x_high) / 2, 'y': (y_high_val + plot_y_max) / 2, 'id': 'B4'},
            b5_id: {'x': (mapped_x_low + mapped_x_high) / 2, 'y': (y_low_val + y_high_val) / 2, 'id': 'B5'},
            b6_id: {'x': (mapped_x_low + mapped_x_high) / 2, 'y': (plot_y_min + y_low_val) / 2, 'id': 'B6'},
            b7_id: {'x': (mapped_x_high + mapped_plot_x_max) / 2, 'y': (y_high_val + plot_y_max) / 2, 'id': 'B7'},
            b8_id: {'x': (mapped_x_high + mapped_plot_x_max) / 2, 'y': (y_low_val + y_high_val) / 2, 'id': 'B8'},
            b9_id: {'x': (mapped_x_high + mapped_plot_x_max) / 2, 'y': (plot_y_min + y_low_val) / 2, 'id': 'B9'}
        }

        for box_name, coords in b_coords.items():
            if b_counts.get(box_name, 0) > 0:
                fig.add_annotation(
                    x=coords['x'], y=coords['y'], xref="x", yref="y",
                    text=f"<span style='color:#1f2937;'><b>{coords['id']}</b></span><br><b>{b_counts[box_name]} SKUs</b>",
                    showarrow=False, font=dict(size=13, color="#374151"),
                    bgcolor="rgba(255, 255, 255, 0.85)", bordercolor="rgba(15, 23, 42, 0.3)",
                    borderwidth=1, borderpad=4
                )

        fig.add_annotation(
            x=0.98, y=0.98, xref="paper", yref="paper",
            text=f"<b>TOTAL:<br>{tot_items} SKUs</b>", showarrow=False,
            font=dict(size=13, color="white"), bgcolor="#374151",
            bordercolor="black", borderwidth=1, borderpad=5
        )

        ticks_x_actual = sorted(list(set([plot_x_min, 0, x_low_val, mid_x, x_high_val, plot_x_max])))
        fig.update_xaxes(
            range=[mapped_plot_x_min, mapped_plot_x_max],
            tickvals=[apply_custom_x_scale(v) for v in ticks_x_actual],
            ticktext=[fmt_idr(v) for v in ticks_x_actual],
            title_text="Gross Sales Absolute (IDR)"
        )
        fig.update_yaxes(range=[plot_y_min, plot_y_max], title_text=f"{cogs_selector} (%)")
        fig.update_layout(
            height=650, hovermode="closest",
            legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, title=None),
            margin=dict(t=50, b=50, l=50, r=50)
        )
        return fig


    # Configs
    hover_conf = {
        'Format_Sales': True,
        'Format_COGS': True,
        'COGS (%)': ':.2f',
        'Format_Qty': True,
        'Dynamic 9-Box Category': False,
        'Price_Increase_Label': False,
        'Bubble_Size': False,
        'X_Plot': False
    }

    color_map_main = {
        b1_id: '#c03d32', b2_id: '#d89f0e', b3_id: '#3871b6',
        b4_id: '#c03d32', b5_id: '#d89f0e', b6_id: '#3871b6',
        b7_id: '#d89f0e', b8_id: '#319b5e', b9_id: '#319b5e'
    }

    # ==========================================
    # MULTI-TAB CHARTS RENDER
    # ==========================================
    st.markdown("---")
    tab_main, tab_pi = st.tabs(["📊 Main COGS 9-Box", "📈 Price Increase Tracker"])

    # ----- TAB 1: MAIN MATRIX -----
    with tab_main:
        fig_main = px.scatter(
            df_sku_valid, x='X_Plot', y='COGS (%)', size='Bubble_Size',
            color='Dynamic 9-Box Category', color_discrete_map=color_map_main,
            hover_name='Product Name', hover_data=hover_conf,
            labels={'Format_Sales': 'Gross Sales', 'Format_COGS': cogs_selector, 'Format_Qty': 'Quantity'},
            title=f"9-Box Bubble Chart: {cogs_selector} (%) vs Gross Sales ({chart_period_title})",
            size_max=60, render_mode='svg'
        )
        fig_main = apply_common_layout(fig_main)
        st.plotly_chart(fig_main, use_container_width=True)

    # ----- TAB 2: PRICE INCREASE TRACKER -----
    with tab_pi:
        if not has_price_data:
            st.warning(
                "⚠️ Data identifikasi Price Increase belum tersedia. Harap upload data Excel multi-sheet yang berisi sheet 'Price Increase' & 'Mapping'.")
        else:
            color_map_pi = {
                'Naik Harga': '#319b5e',  # Hijau
                'Harga Tetap': '#3871b6'  # Biru
            }
            fig_pi = px.scatter(
                df_sku_valid, x='X_Plot', y='COGS (%)', size='Bubble_Size',
                color='Price_Increase_Label', color_discrete_map=color_map_pi,
                hover_name='Product Name', hover_data=hover_conf,
                labels={'Format_Sales': 'Gross Sales', 'Format_COGS': cogs_selector, 'Format_Qty': 'Quantity',
                        'Price_Increase_Label': 'Status Harga'},
                title=f"Price Increase Tracker: {cogs_selector} (%) vs Gross Sales ({chart_period_title})",
                size_max=60, render_mode='svg'
            )
            fig_pi = apply_common_layout(fig_pi)

            count_naik = df_sku_valid[df_sku_valid['Price_Increase_Label'] == 'Naik Harga'].shape[0]
            count_tetap = df_sku_valid[df_sku_valid['Price_Increase_Label'] == 'Harga Tetap'].shape[0]

            fig_pi.add_annotation(
                x=0.01, y=0.98, xref="paper", yref="paper",
                text=f"<b>Status Harga:</b><br><span style='color:#319b5e'>Naik Harga: {count_naik} SKUs</span><br><span style='color:#3871b6'>Harga Tetap: {count_tetap} SKUs</span>",
                showarrow=False, font=dict(size=13, color="#1f2937"),
                bgcolor="rgba(255, 255, 255, 0.95)", bordercolor="rgba(15, 23, 42, 0.3)",
                borderwidth=1, borderpad=6, align="left"
            )
            st.plotly_chart(fig_pi, use_container_width=True)

            # MISSING SKU DETECTOR
            master_pi_keys = set(pi_produk_keys)
            plotted_pi_keys = set(df_sku_valid[df_sku_valid['Price_Increase']]['Produk_Key_Str'].dropna())
            missing_pi_keys = master_pi_keys - plotted_pi_keys

            if missing_pi_keys:
                missing_df = df_pi_mapped[df_pi_mapped['Produk_Key'].isin(missing_pi_keys)][
                    ['Old material number', 'Produk_Key', 'Product Name', 'Price Increase Period']].copy()
                missing_df['Price Increase Period'] = missing_df['Price Increase Period'].dt.strftime(
                    '%Y-%m-%d').fillna('-')

                keys_in_raw = set(df_filtered['Produk_Key_Str'].dropna())
                keys_in_sku_pre_filter = set(df_sku['Produk_Key_Str'].dropna())


                def get_reason(k):
                    if k not in keys_in_raw:
                        return "Bolos (Tidak ada transaksi di periode terpilih)"
                    elif k not in keys_in_sku_pre_filter:
                        return "Terfilter (Nilai Sales/COGS < Rp 1)"
                    else:
                        return "Terfilter (Nilai Sales/COGS < Rp 1)"


                missing_df['Alasan Hilang'] = missing_df['Produk_Key'].apply(get_reason)

                st.markdown(
                    f"**⚠️ Terdapat {len(missing_pi_keys)} SKU Naik Harga yang tidak ter-plot ke dalam grafik utama:**")

                # Excel Download for Missing SKUs
                col_mis1, col_mis2 = st.columns([1, 4])
                with col_mis1:
                    buffer_missing = io.BytesIO()
                    with pd.ExcelWriter(buffer_missing, engine='openpyxl') as writer:
                        missing_df.to_excel(writer, index=False, sheet_name='Missing_SKUs')
                        worksheet_mis = writer.sheets['Missing_SKUs']
                        for col in worksheet_mis.columns:
                            max_len = max([len(str(cell.value)) if cell.value is not None else 0 for cell in col])
                            worksheet_mis.column_dimensions[col[0].column_letter].width = max_len + 2
                    st.download_button(
                        label="📥 Download Missing SKUs (Excel)",
                        data=buffer_missing.getvalue(),
                        file_name=f"Missing_SKUs_Price_Increase_{chart_period_title.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.dataframe(missing_df, use_container_width=True, hide_index=True)

            # ==========================================
            # BEFORE & AFTER TOGGLE CHART (APPLES-TO-APPLES)
            # ==========================================
            st.markdown("---")
            st.markdown("#### 🔄 Before & After Price Increase Analysis")
            st.info(
                "Visualisasi khusus membandingkan pergerakan performa **SKU yang sama persis** (Apples-to-Apples). Sistem membandingkan agregat transaksi sebelum vs sesudah tanggal naik harga untuk SKU yang aktif di kedua periode tersebut.")

            toggle_mode = st.radio("Pilih Mode Timeline:", ["Before Price Increase", "After Price Increase"],
                                   horizontal=True)

            df_pi_trx = df_filtered[df_filtered['Price_Increase_Flag'] == True].copy()


            # Fungsi Agregasi & Filter per Periode
            def get_agg_valid(df_sub):
                if df_sub.empty: return pd.DataFrame()
                df_agg = df_sub.groupby('Product Name').agg(
                    Total_Gross_Sales=('Calc_Gross_Sales', 'sum'),
                    Total_COGS_Reg=('COGS_Regular', 'sum'),
                    Total_COGS_All=('Total COGS', 'sum'),
                    Selected_COGS=(cogs_selector, 'sum'),
                    Total_Qty=('Quantity', 'sum')
                ).reset_index()
                return df_agg[
                    (df_agg['Total_Gross_Sales'] >= 1) &
                    (df_agg['Total_COGS_Reg'] >= 1) &
                    (df_agg['Total_COGS_All'] >= 1)
                    ].copy()


            df_raw_before = df_pi_trx[df_pi_trx['Date'] < df_pi_trx['Price_Increase_Period']]
            df_raw_after = df_pi_trx[df_pi_trx['Date'] >= df_pi_trx['Price_Increase_Period']]

            df_before_agg = get_agg_valid(df_raw_before)
            df_after_agg = get_agg_valid(df_raw_after)

            if df_before_agg.empty or df_after_agg.empty:
                st.warning("⚠️ Tidak cukup data untuk melakukan perbandingan Before-After (salah satu periode kosong).")
            else:
                # 🍎 MURNI APPLES-TO-APPLES COMPARISON (INTERSECTION) 🍎
                common_skus = set(df_before_agg['Product Name']).intersection(set(df_after_agg['Product Name']))

                # Proses data BEFORE
                df_before_agg['COGS (%)'] = (df_before_agg['Selected_COGS'] / df_before_agg['Total_Gross_Sales']) * 100
                df_before_agg['Dynamic 9-Box Category'] = assign_9box(df_before_agg)

                # Proses data AFTER
                df_after_agg['COGS (%)'] = (df_after_agg['Selected_COGS'] / df_after_agg['Total_Gross_Sales']) * 100
                df_after_agg['Dynamic 9-Box Category'] = assign_9box(df_after_agg)

                if toggle_mode == "Before Price Increase":
                    df_ba = df_before_agg[df_before_agg['Product Name'].isin(common_skus)].copy()
                    chart_title_ba = "BEFORE Price Increase"
                else:
                    df_ba = df_after_agg[df_after_agg['Product Name'].isin(common_skus)].copy()
                    chart_title_ba = "AFTER Price Increase"

                if df_ba.empty:
                    st.warning(
                        f"Setelah mencari irisan SKU yang aktif di kedua periode, tidak ada data tersisa untuk {chart_title_ba}.")
                else:
                    df_ba['Bubble_Size'] = df_ba['Total_Qty'].abs().replace(0, 1)
                    df_ba['Price_Increase_Label'] = 'Naik Harga'

                    df_ba['Format_Sales'] = df_ba['Total_Gross_Sales'].apply(lambda x: f"Rp {x:,.0f}")
                    df_ba['Format_COGS'] = df_ba['Selected_COGS'].apply(lambda x: f"Rp {x:,.0f}")
                    df_ba['Format_Qty'] = df_ba['Total_Qty'].apply(lambda x: f"{x:,.0f} Pcs")
                    df_ba['X_Plot'] = df_ba['Total_Gross_Sales'].apply(apply_custom_x_scale)

                    b_counts_ba = df_ba['Dynamic 9-Box Category'].value_counts().to_dict()

                    fig_ba = px.scatter(
                        df_ba, x='X_Plot', y='COGS (%)', size='Bubble_Size',
                        color='Dynamic 9-Box Category', color_discrete_map=color_map_main,
                        hover_name='Product Name', hover_data=hover_conf,
                        labels={'Format_Sales': 'Gross Sales', 'Format_COGS': cogs_selector, 'Format_Qty': 'Quantity'},
                        title=f"{chart_title_ba} Matrix: {cogs_selector} (%) vs Gross Sales | Total Terkalkulasi: {len(df_ba)} SKUs Apples-to-Apples",
                        size_max=60, render_mode='svg'
                    )

                    fig_ba = apply_common_layout(fig_ba, custom_box_counts=b_counts_ba)
                    st.plotly_chart(fig_ba, use_container_width=True)

                    # 🧭 SKU MIGRATION TRACKER TABLE
                    st.markdown("#### 🧭 SKU Migration Tracker")
                    st.caption(
                        "Detail pergerakan kotak (Box) masing-masing SKU dari sebelum naik harga menjadi setelah naik harga.")

                    # Merge data Before dan After khusus untuk SKU yang masuk irisan
                    df_mig_before = df_before_agg[df_before_agg['Product Name'].isin(common_skus)][
                        ['Product Name', 'Total_Gross_Sales', 'COGS (%)', 'Dynamic 9-Box Category']].copy()
                    df_mig_after = df_after_agg[df_after_agg['Product Name'].isin(common_skus)][
                        ['Product Name', 'Total_Gross_Sales', 'COGS (%)', 'Dynamic 9-Box Category']].copy()

                    df_migration = pd.merge(df_mig_before, df_mig_after, on='Product Name',
                                            suffixes=(' (Before)', ' (After)'))


                    # Fungsi untuk nentuin status migrasi
                    def get_migration_status(row):
                        box_b = row['Dynamic 9-Box Category (Before)'].split(' ')[1]
                        box_a = row['Dynamic 9-Box Category (After)'].split(' ')[1]
                        if box_b == box_a:
                            return f"Tetap di Box {box_b}"
                        else:
                            return f"Pindah: Box {box_b} ➡️ Box {box_a}"


                    df_migration['Status Pergerakan'] = df_migration.apply(get_migration_status, axis=1)

                    # Rapihin kolom untuk display
                    df_migration = df_migration[[
                        'Product Name', 'Status Pergerakan',
                        'Dynamic 9-Box Category (Before)', 'Dynamic 9-Box Category (After)',
                        'Total_Gross_Sales (Before)', 'Total_Gross_Sales (After)',
                        'COGS (%) (Before)', 'COGS (%) (After)'
                    ]].sort_values('Status Pergerakan')

                    # Excel Download for Migration Tracker
                    col_mig1, col_mig2 = st.columns([1, 4])
                    with col_mig1:
                        buffer_mig = io.BytesIO()
                        with pd.ExcelWriter(buffer_mig, engine='openpyxl') as writer:
                            df_migration.to_excel(writer, index=False, sheet_name='Migration_Tracker')
                            ws_mig = writer.sheets['Migration_Tracker']
                            for col in ws_mig.columns:
                                max_len = max([len(str(cell.value)) if cell.value is not None else 0 for cell in col])
                                ws_mig.column_dimensions[col[0].column_letter].width = max_len + 2
                        st.download_button(
                            label="📥 Download Migration Tracker (Excel)",
                            data=buffer_mig.getvalue(),
                            file_name=f"SKU_Migration_Tracker_{chart_period_title.replace(' ', '_')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )

                    st.dataframe(
                        df_migration.style.format({
                            'Total_Gross_Sales (Before)': 'Rp {:,.0f}',
                            'Total_Gross_Sales (After)': 'Rp {:,.0f}',
                            'COGS (%) (Before)': '{:.2f}%',
                            'COGS (%) (After)': '{:.2f}%'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )

                    # 📉 SKU DROP-OFF (Bolosan) TABLE
                    lost_skus = set(df_before_agg['Product Name']) - set(df_after_agg['Product Name'])
                    if lost_skus:
                        st.markdown("---")
                        st.markdown("#### 📉 SKU Drop-off (Hilang Setelah Naik Harga)")
                        st.caption(
                            "Daftar SKU yang memiliki penjualan sebelum naik harga, namun **TIDAK ADA** penjualan (atau terfilter < Rp 1) setelah harga dinaikkan.")

                        df_lost = df_before_agg[df_before_agg['Product Name'].isin(lost_skus)].copy()
                        df_lost = df_lost[['Product Name', 'Total_Gross_Sales', 'COGS (%)', 'Dynamic 9-Box Category',
                                           'Total_Qty']].sort_values('Total_Gross_Sales', ascending=False)

                        # Excel Download for SKU Drop-off
                        col_lost1, col_lost2 = st.columns([1, 4])
                        with col_lost1:
                            buffer_lost = io.BytesIO()
                            with pd.ExcelWriter(buffer_lost, engine='openpyxl') as writer:
                                df_lost.to_excel(writer, index=False, sheet_name='Lost_SKUs')
                                ws_lost = writer.sheets['Lost_SKUs']
                                for col in ws_lost.columns:
                                    max_len = max(
                                        [len(str(cell.value)) if cell.value is not None else 0 for cell in col])
                                    ws_lost.column_dimensions[col[0].column_letter].width = max_len + 2
                            st.download_button(
                                label="📥 Download SKU Drop-off (Excel)",
                                data=buffer_lost.getvalue(),
                                file_name=f"SKU_DropOff_After_Price_Increase_{chart_period_title.replace(' ', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )

                        st.dataframe(
                            df_lost.style.format({
                                'Total_Gross_Sales': 'Rp {:,.0f}',
                                'COGS (%)': '{:.2f}%',
                                'Total_Qty': '{:,.0f} Pcs'
                            }),
                            use_container_width=True,
                            hide_index=True
                        )

    # ==========================================
    # DATA TABLE & EXPORT
    # ==========================================
    st.markdown("---")
    st.markdown("### 📋 SKU Details")

    # Clean UI DataFrame
    drop_cols = ['Bubble_Size', 'X_Plot', 'Format_Sales', 'Format_COGS', 'Format_Qty', 'Price_Increase',
                 'Produk_Key_Str', 'Price_Increase_Period']
    df_display = df_sku_valid.drop(columns=[c for c in drop_cols if c in df_sku_valid.columns]).sort_values(
        'Total_Gross_Sales', ascending=False)

    col_dl1, col_dl2 = st.columns([1, 4])
    with col_dl1:
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            df_display.to_excel(writer, index=False, sheet_name='COGS_Details')
            worksheet_f = writer.sheets['COGS_Details']
            for col in worksheet_f.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                    except:
                        pass
                worksheet_f.column_dimensions[column].width = (max_length + 2)

        st.download_button(
            label="📥 Download Data (Excel)",
            data=buffer_excel.getvalue(),
            file_name=f"COGS_9Box_{cogs_selector.replace(' ', '_')}_{chart_period_title.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    st.dataframe(
        df_display.style.format({
            'Total_Gross_Sales': 'Rp {:,.0f}',
            'Total_COGS_Reg': 'Rp {:,.0f}',
            'Total_COGS_All': 'Rp {:,.0f}',
            'Selected_COGS': 'Rp {:,.0f}',
            'COGS (%)': '{:.2f}%',
            'Total_Qty': '{:,.0f}'
        }),
        use_container_width=True,
        hide_index=True
    )

except Exception as e:
    st.error(f"Terjadi kesalahan saat memproses data P&L. Pastikan format file sesuai. Error Log: {e}")
