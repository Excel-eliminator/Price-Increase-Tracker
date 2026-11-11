import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io

st.set_page_config(page_title="COGS 9-Box Bubble Chart", layout="wide")


@st.cache_data(show_spinner="Optimizing & caching data into memory...")
def load_and_clean_data(file_bytes, file_name):
    file_ext = file_name.split('.')[-1].lower()
    has_price_data = False
    baseline_prices = {}

    file_buffer = io.BytesIO(file_bytes)

    if file_ext == 'xlsx':
        xls = pd.ExcelFile(file_buffer)
        df = pd.read_excel(xls, sheet_name='PL_2026') if 'PL_2026' in xls.sheet_names else pd.read_excel(xls)

        if 'Price List 2026' in xls.sheet_names:
            df_pl = pd.read_excel(xls, sheet_name='Price List 2026')
            df_pl.columns = df_pl.columns.str.strip()
            if all(c in df_pl.columns for c in ['Produk_Key', 'Date', 'Price']):
                df_pl['Produk_Key'] = df_pl['Produk_Key'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                df_pl['Date'] = pd.to_datetime(df_pl['Date'], dayfirst=True, errors='coerce')

                if not pd.api.types.is_numeric_dtype(df_pl['Price']):
                    s = df_pl['Price'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                    df_pl['Price'] = pd.to_numeric(s, errors='coerce').fillna(0.0)
                else:
                    df_pl['Price'] = df_pl['Price'].fillna(0.0)

                df_pl = df_pl.dropna(subset=['Date', 'Produk_Key']).sort_values(by=['Produk_Key', 'Date'])
                baseline_prices = df_pl.groupby('Produk_Key')['Price'].first().to_dict()
                has_price_data = True
            else:
                df_pl = pd.DataFrame()
        else:
            df_pl = pd.DataFrame()
    else:
        df = pd.read_csv(file_buffer, sep=';', dtype=str)
        if len(df.columns) == 1:
            file_buffer.seek(0)
            df = pd.read_csv(file_buffer, sep=',', dtype=str)
        df_pl = pd.DataFrame()

    df.columns = df.columns.str.strip()
    req_cols = ['Month', 'Year', 'Produk_Key', 'Product Name', 'Gross Sales', 'Return', 'COGS_Regular', 'Total COGS',
                'Quantity', 'Market']
    missing_cols = [c for c in req_cols if c not in df.columns]

    if missing_cols:
        return None, None, False, {}, missing_cols, None

    def vectorize_clean(series):
        if pd.api.types.is_numeric_dtype(series): return series.fillna(0.0)
        s = series.astype(str).str.strip().replace(['', '-'], '0')
        s = s.str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
        return pd.to_numeric(s, errors='coerce').fillna(0.0)

    for col in ['Gross Sales', 'Return', 'COGS_Regular', 'Total COGS', 'Quantity']:
        df[col] = vectorize_clean(df[col])

    df['Market'] = df['Market'].fillna('Unknown').astype(str).str.strip()
    df['Calc_Gross_Sales'] = df['Gross Sales'] + df['Return']
    df['Year_Str'] = df['Year'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df['Month_Str'] = df['Month'].astype(str).str.strip()
    df['Date'] = pd.to_datetime(df['Month_Str'] + ' ' + df['Year_Str'], format='%B %Y', errors='coerce')

    df = df.dropna(subset=['Date'])
    df['Month_Year'] = df['Date'].dt.strftime('%B %Y')
    df['YYYYMM'] = df['Date'].dt.strftime('%Y%m')
    df['Produk_Key_Str'] = df['Produk_Key'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

    # seluruh histori di file -> dipakai HANYA untuk baseline "Before Price Increase"
    df_hist = df.copy()
    # jendela evaluasi (Monthly / YTD / Custom) tetap mulai Jan 2026 seperti sebelumnya
    df = df[df['Date'] >= '2026-01-01'].copy()

    return df, df_pl, has_price_data, baseline_prices, [], df_hist


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

st.markdown("### 📂 Upload Data Files")
col_up1, col_up2 = st.columns(2)
with col_up1:
    uploaded_file = st.file_uploader("1. Upload P&L Data (e.g., Sample_PL_2026.xlsx)", type=["xlsx", "csv"])
with col_up2:
    uploaded_target_file = st.file_uploader("2. Upload Target SKUs (Optional)", type=["xlsx", "csv"])

if uploaded_file is None:
    st.info("☝️ Please upload your P&L data file to begin.")
    st.stop()

df, df_pl, has_price_data, baseline_prices, missing_cols, df_hist = load_and_clean_data(uploaded_file.getvalue(),
                                                                               uploaded_file.name)

if missing_cols:
    st.error(f"⚠️ Missing columns: {', '.join(missing_cols)}")
    st.stop()

if df.empty:
    st.warning("⚠️ No transaction data starting from January 2026.")
    st.stop()

target_skus = set()
target_new_skus = set()
target_mapping = {}
target_name_mapping = {}

if uploaded_target_file is not None:
    try:
        if uploaded_target_file.name.endswith('.xlsx'):
            df_tgt = pd.read_excel(uploaded_target_file)
        else:
            df_tgt = pd.read_csv(uploaded_target_file, sep=None, engine='python')

        cols = df_tgt.columns
        key_col = cols[0]
        new_key_col = cols[1] if len(cols) > 1 else key_col
        new_name_col = cols[2] if len(cols) > 2 else None

        for col in cols:
            cl = str(col).lower()
            if ('old' in cl or 'original' in cl) and ('sku' in cl or 'code' in cl or 'key' in cl): key_col = col
            if 'new' in cl and ('sku' in cl or 'code' in cl or 'key' in cl): new_key_col = col
            if 'new' in cl and ('name' in cl or 'desc' in cl): new_name_col = col

        df_tgt['Clean_Key'] = df_tgt[key_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        df_tgt['Clean_New_Key'] = df_tgt[new_key_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

        target_skus = set(df_tgt['Clean_Key'])
        target_new_skus = set(df_tgt['Clean_New_Key'])
        target_mapping = dict(zip(df_tgt['Clean_Key'], df_tgt['Clean_New_Key']))

        if new_name_col:
            target_name_mapping = dict(zip(df_tgt['Clean_New_Key'], df_tgt[new_name_col].astype(str).str.strip()))

        st.success(
            f"✅ Loaded {len(target_skus)} Target SKUs successfully (Consolidated to {len(target_new_skus)} SKUs).")
    except Exception as e:
        st.error(f"Failed to load Target SKUs: {e}")

st.markdown("---")
st.markdown("### ⚙️ Evaluation Settings")

col_set1, col_set2, col_set3, col_set4 = st.columns(4)

with col_set1:
    cogs_selector = st.selectbox("Y-Axis Metric (COGS Component):", ["Total COGS", "COGS_Regular"])

with col_set2:
    period_mode = st.radio("Time Evaluation Mode:", ["Monthly", "Year-to-Date (YTD)", "Custom Range"], horizontal=True)

with col_set3:
    months_sorted = df[['YYYYMM', 'Month_Year']].drop_duplicates().sort_values('YYYYMM')
    month_list = months_sorted['Month_Year'].tolist()

    if period_mode == "Monthly":
        selected_month = st.selectbox("Select Month:", month_list)
        df_filtered = df[df['Month_Year'] == selected_month].copy()
        chart_period_title = selected_month
    elif period_mode == "Year-to-Date (YTD)":
        selected_month = st.selectbox("YTD up to Month:", month_list)
        selected_yyyymm = months_sorted[months_sorted['Month_Year'] == selected_month]['YYYYMM'].values[0]
        df_filtered = df[df['YYYYMM'] <= selected_yyyymm].copy()
        chart_period_title = f"YTD (Jan 2026 - {selected_month.split(' ')[0]})"
    else:
        start_month, end_month = st.select_slider("Select Period Range (From - To):", options=month_list,
                                                  value=(month_list[0], month_list[-1]))
        start_yyyymm = months_sorted[months_sorted['Month_Year'] == start_month]['YYYYMM'].values[0]
        end_yyyymm = months_sorted[months_sorted['Month_Year'] == end_month]['YYYYMM'].values[0]

        if start_yyyymm > end_yyyymm:
            start_yyyymm, end_yyyymm = end_yyyymm, start_yyyymm
            start_month, end_month = end_month, start_month

        df_filtered = df[(df['YYYYMM'] >= start_yyyymm) & (df['YYYYMM'] <= end_yyyymm)].copy()
        chart_period_title = f"{start_month} to {end_month}"

    end_month_str = selected_month if period_mode in ["Monthly", "Year-to-Date (YTD)"] else end_month
    period_end_date = pd.to_datetime(end_month_str, format='%B %Y') + pd.offsets.MonthEnd(1)

    if period_mode == "Monthly":
        start_month_str = selected_month
        before_end_date = period_end_date
    elif period_mode == "Year-to-Date (YTD)":
        start_month_str = month_list[0]
        before_end_date = pd.to_datetime(start_month_str, format='%B %Y') + pd.offsets.MonthEnd(1)
    else:
        start_month_str = start_month
        before_end_date = pd.to_datetime(start_month_str, format='%B %Y') + pd.offsets.MonthEnd(1)

    pi_produk_keys = set()
    pi_period_dict = {}

    if has_price_data and not df_pl.empty:
        df_pl_period = df_pl[df_pl['Date'] <= period_end_date]
        max_prices_in_period = df_pl_period.groupby('Produk_Key')['Price'].max()

        for sku, max_price in max_prices_in_period.items():
            baseline = baseline_prices.get(sku, max_price)
            if max_price > baseline:
                pi_produk_keys.add(sku)
                sku_data = df_pl_period[(df_pl_period['Produk_Key'] == sku) & (df_pl_period['Price'] > baseline)]
                if not sku_data.empty:
                    pi_period_dict[sku] = sku_data['Date'].min()

    df_filtered['Price_Increase_Flag'] = df_filtered['Produk_Key_Str'].isin(pi_produk_keys)
    df_filtered['Price_Increase_Period'] = df_filtered['Produk_Key_Str'].map(
        pi_period_dict) if has_price_data else pd.NaT

with col_set4:
    market_list = ["All Markets"] + sorted([m for m in df['Market'].unique() if m != 'Unknown'])
    selected_market = st.selectbox("Select Market:", market_list)

    if selected_market != "All Markets":
        df_filtered = df_filtered[df_filtered['Market'] == selected_market].copy()
        chart_period_title += f" | {selected_market}"

# ---- histori penuh untuk baseline Before (semua transaksi di file sebelum tanggal naik SKU) ----
df_hist_m = df_hist.copy() if selected_market == "All Markets" else df_hist[df_hist['Market'] == selected_market].copy()
df_hist_m['Price_Increase_Flag'] = df_hist_m['Produk_Key_Str'].isin(pi_produk_keys)
df_hist_m['Price_Increase_Period'] = df_hist_m['Produk_Key_Str'].map(pi_period_dict) if has_price_data else pd.NaT
hist_range_lbl = (f"{df_hist['Date'].min():%b %Y} - {df_hist['Date'].max():%b %Y}") if not df_hist.empty else "-"

df_sku = df_filtered.groupby('Produk_Key_Str').agg(
    Product_Name=('Product Name', 'first'),
    Total_Gross_Sales=('Calc_Gross_Sales', 'sum'),
    Total_COGS_Reg=('COGS_Regular', 'sum'),
    Total_COGS_All=('Total COGS', 'sum'),
    Selected_COGS=(cogs_selector, 'sum'),
    Total_Qty=('Quantity', 'sum'),
    Price_Increase=('Price_Increase_Flag', 'max'),
    Price_Increase_Period=('Price_Increase_Period', 'first')
).reset_index()

df_sku.rename(columns={'Product_Name': 'Product Name'}, inplace=True)
df_sku_valid = df_sku[(df_sku['Total_Gross_Sales'] >= 1) & (df_sku['Selected_COGS'] >= 1)].copy()

if df_sku_valid.empty:
    st.warning("⚠️ No valid data found for the selected Market & Period.")
    st.stop()

st.markdown("#### 🎚️ Smart-Scaling & Pareto Control")
col_out1, col_out2, col_out3, col_out4 = st.columns(4)
with col_out1: min_outlier_limit_x = st.number_input("MIN Gross Sales Limit:", value=0.0, step=1000000.0)
with col_out2: scale_factor_x = st.number_input("Right Box X-Scale (Multiplier):", value=0.2, step=0.1, min_value=0.001,
                                                max_value=10.0)
with col_out3: scale_factor_y = st.number_input("Top Box Y-Scale (Multiplier):", value=0.2, step=0.1, min_value=0.001,
                                                max_value=10.0)
with col_out4: pareto_limit = st.slider("Pareto Filter (Top % Sales):", min_value=10, max_value=100, value=100, step=5)

pre_sku_count = len(df_sku_valid)
pre_qty = df_sku_valid['Total_Qty'].sum()
pre_sales = df_sku_valid['Total_Gross_Sales'].sum()
pre_gm = (df_sku_valid['Total_Gross_Sales'] - df_sku_valid['Total_COGS_Reg']).sum()
pre_gp = (df_sku_valid['Total_Gross_Sales'] - df_sku_valid['Total_COGS_All']).sum()

if pareto_limit < 100:
    df_sku_valid = df_sku_valid.sort_values('Total_Gross_Sales', ascending=False)
    df_sku_valid['Cum_Sales'] = df_sku_valid['Total_Gross_Sales'].cumsum()
    target_frac = pareto_limit / 100.0
    df_sku_valid['Keep'] = (df_sku_valid['Cum_Sales'] / pre_sales).shift().fillna(0) < target_frac
    df_sku_valid = df_sku_valid[df_sku_valid['Keep']].drop(columns=['Cum_Sales', 'Keep'])
    if df_sku_valid.empty:
        st.warning("⚠️ Pareto filter resulted in empty data.")
        st.stop()

post_sku_count = len(df_sku_valid)
post_qty = df_sku_valid['Total_Qty'].sum()
post_sales = df_sku_valid['Total_Gross_Sales'].sum()
post_gm = (df_sku_valid['Total_Gross_Sales'] - df_sku_valid['Total_COGS_Reg']).sum()
post_gp = (df_sku_valid['Total_Gross_Sales'] - df_sku_valid['Total_COGS_All']).sum()

post_gm_ratio = (post_gm / post_sales * 100) if post_sales != 0 else 0
post_gp_ratio = (post_gp / post_sales * 100) if post_sales != 0 else 0


def perc(num, den): return f"{(num / den * 100):.1f}%" if den != 0 else "0.0%"


def fmt_bn(val): return f"Rp {val / 1e9:,.1f} Bn"


html_metrics = f"""
<div style="margin-top: 15px; margin-bottom: 25px; font-family: sans-serif;">
    <div style="display: flex; gap: 20px; flex-wrap: wrap; margin-bottom: 20px;">
        <div style="flex: 1; min-width: 200px; padding: 20px; background-color: white; border-radius: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border: 1px solid #e2e8f0;">
            <p style="margin: 0 0 5px 0; font-size: 14px; color: #475569; font-weight: 500;">Total SKUs</p>
            <h2 style="margin: 0 0 10px 0; font-size: 32px; color: #1e293b; font-weight: 400;">{post_sku_count:,.0f}</h2>
            <span style="font-size: 12px; color: #10b981; background-color: #d1fae5; padding: 4px 10px; border-radius: 20px; font-weight: 600;">↑ {perc(post_sku_count, pre_sku_count)} from {pre_sku_count:,.0f} Global SKUs</span>
        </div>
        <div style="flex: 1; min-width: 250px; padding: 20px; background-color: white; border-radius: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border: 1px solid #e2e8f0;">
            <p style="margin: 0 0 5px 0; font-size: 14px; color: #475569; font-weight: 500;">Quantity</p>
            <h2 style="margin: 0 0 10px 0; font-size: 32px; color: #1e293b; font-weight: 400;">{post_qty:,.0f} <span style="font-size: 20px; color: #64748b; font-weight: 300;">({perc(post_qty, pre_qty)})</span></h2>
            <span style="font-size: 12px; color: #10b981; background-color: #d1fae5; padding: 4px 10px; border-radius: 20px; font-weight: 600;">↑ {perc(post_qty, pre_qty)} from total slice</span>
        </div>
        <div style="flex: 1; min-width: 250px; padding: 20px; background-color: white; border-radius: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border: 1px solid #e2e8f0;">
            <p style="margin: 0 0 5px 0; font-size: 14px; color: #475569; font-weight: 500;">Gross Sales</p>
            <h2 style="margin: 0 0 10px 0; font-size: 32px; color: #1e293b; font-weight: 400;">{fmt_bn(post_sales)} <span style="font-size: 20px; color: #64748b; font-weight: 300;">({perc(post_sales, pre_sales)})</span></h2>
            <span style="font-size: 12px; color: #10b981; background-color: #d1fae5; padding: 4px 10px; border-radius: 20px; font-weight: 600;">↑ {perc(post_sales, pre_sales)} from total slice</span>
        </div>
    </div>
    <div style="display: flex; gap: 20px; flex-wrap: wrap;">
        <div style="flex: 1; min-width: 250px; padding: 20px; background-color: white; border-radius: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border: 1px solid #e2e8f0;">
            <p style="margin: 0 0 5px 0; font-size: 14px; color: #475569; font-weight: 500;">Gross Profit</p>
            <h2 style="margin: 0 0 10px 0; font-size: 32px; color: #1e293b; font-weight: 400;">{fmt_bn(post_gp)} <span style="font-size: 20px; color: #64748b; font-weight: 300;">({perc(post_gp, pre_gp)})</span></h2>
            <span style="font-size: 12px; color: #10b981; background-color: #d1fae5; padding: 4px 10px; border-radius: 20px; font-weight: 600;">↑ {post_gp_ratio:.1f}% margin ratio</span>
        </div>
        <div style="flex: 1; min-width: 250px; padding: 20px; background-color: white; border-radius: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border: 1px solid #e2e8f0;">
            <p style="margin: 0 0 5px 0; font-size: 14px; color: #475569; font-weight: 500;">Gross Margin</p>
            <h2 style="margin: 0 0 10px 0; font-size: 32px; color: #1e293b; font-weight: 400;">{fmt_bn(post_gm)} <span style="font-size: 20px; color: #64748b; font-weight: 300;">({perc(post_gm, pre_gm)})</span></h2>
            <span style="font-size: 12px; color: #10b981; background-color: #d1fae5; padding: 4px 10px; border-radius: 20px; font-weight: 600;">↑ {post_gm_ratio:.1f}% margin ratio</span>
        </div>
        <div style="flex: 1; min-width: 250px; padding: 20px; background-color: transparent; border: 0px;"></div>
    </div>
</div>
""".replace('\n', '')
st.markdown(html_metrics, unsafe_allow_html=True)

df_sku_valid['COGS (%)'] = (df_sku_valid['Selected_COGS'] / df_sku_valid['Total_Gross_Sales']) * 100
df_sku_valid['Bubble_Size'] = df_sku_valid['Total_Qty'].abs().replace(0, 1)
df_sku_valid['Price_Increase_Label'] = np.where(df_sku_valid['Price_Increase'], 'Price Increased', 'Stable Price')
df_sku_valid['Periode Naik Harga'] = df_sku_valid['Price_Increase_Period'].dt.strftime('%Y-%m-%d').fillna('-')
df_sku_valid['Format_Sales'] = df_sku_valid['Total_Gross_Sales'].apply(lambda x: f"Rp {x:,.0f}")
df_sku_valid['Format_COGS'] = df_sku_valid['Selected_COGS'].apply(lambda x: f"Rp {x:,.0f}")
df_sku_valid['Format_Qty'] = df_sku_valid['Total_Qty'].apply(lambda x: f"{x:,.0f} Pcs")

st.info(
    "💡 **INFO:** The X and Y axes use the dynamic Midpoint value as the absolute center point based on your custom Low and High thresholds. The Center lines are marked in Red.")

normal_x_df = df_sku_valid[df_sku_valid['Total_Gross_Sales'] >= min_outlier_limit_x]
robust_avg_x = normal_x_df['Total_Gross_Sales'].mean() if not normal_x_df.empty else df_sku_valid[
    'Total_Gross_Sales'].mean()
if pd.isna(robust_avg_x) or robust_avg_x <= 0: robust_avg_x = 1000000.0

actual_avg_y = (df_sku_valid['Selected_COGS'].sum() / df_sku_valid['Total_Gross_Sales'].sum() * 100) if df_sku_valid[
                                                                                                            'Total_Gross_Sales'].sum() > 0 else \
df_sku_valid['COGS (%)'].mean()
if pd.isna(actual_avg_y) or actual_avg_y <= 0: actual_avg_y = 50.0

with st.form("cogs_threshold_form"):
    col_t1, col_t2, col_t3, col_t4 = st.columns(4)
    with col_t1: x_low_val = st.number_input("X-Axis Low to Med (Gross Sales)", value=float(robust_avg_x * (2 / 3)),
                                             step=1000000.0)
    with col_t2: x_high_val = st.number_input("X-Axis Med to High (Gross Sales)", value=float(robust_avg_x * (4 / 3)),
                                              step=1000000.0)
    with col_t3: y_low_val = st.number_input("Y-Axis Low to Med (COGS %)", value=float(actual_avg_y * (2 / 3)),
                                             step=1.0)
    with col_t4: y_high_val = st.number_input("Y-Axis Med to High (COGS %)", value=float(actual_avg_y * (4 / 3)),
                                              step=1.0)
    st.form_submit_button("▶ RENDER BUBBLE CHART", type="primary")

mid_x, mid_y = (x_low_val + x_high_val) / 2.0, (y_low_val + y_high_val) / 2.0


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
    return np.select(conditions, ['Box 1', 'Box 2', 'Box 3', 'Box 4', 'Box 5', 'Box 6', 'Box 7', 'Box 8', 'Box 9'],
                     default='Box 5')


df_sku_valid['Dynamic 9-Box Category'] = assign_9box(df_sku_valid)


def apply_custom_x_scale(v): return v if v <= x_high_val else x_high_val + (v - x_high_val) * scale_factor_x


def apply_custom_y_scale(v): return v if v <= y_high_val else y_high_val + (v - y_high_val) * scale_factor_y


df_sku_valid['X_Plot'] = df_sku_valid['Total_Gross_Sales'].apply(apply_custom_x_scale)
df_sku_valid['Y_Plot'] = df_sku_valid['COGS (%)'].apply(apply_custom_y_scale)

x_min_raw, x_max_raw = df_sku_valid['Total_Gross_Sales'].min(), df_sku_valid['Total_Gross_Sales'].max()
x_span_raw = x_max_raw - x_min_raw if x_max_raw != x_min_raw else robust_avg_x
plot_x_min = x_min_raw - (abs(x_span_raw) * 0.05)
plot_x_max = x_max_raw + (abs(x_span_raw) * 0.05)

y_min_raw, y_max_raw = df_sku_valid['COGS (%)'].min(), df_sku_valid['COGS (%)'].max()
y_span_raw = y_max_raw - y_min_raw if y_max_raw != y_min_raw else 100.0
plot_y_min, plot_y_max = y_min_raw - (y_span_raw * 0.05), max(y_max_raw, y_high_val * 1.5) + (y_span_raw * 0.05)


def fmt_idr_axis(val):
    if val >= 1e9: return f"Rp{val / 1e9:,.1f}Bn"
    if val >= 1e6: return f"Rp{val / 1e6:,.0f}M"
    return f"Rp{val:,.0f}"


def display_9box_matrix(df_target):
    stats = {}
    for b in [f'Box {i}' for i in range(1, 10)]:
        sub = df_target[df_target['Dynamic 9-Box Category'] == b]
        sales = sub['Total_Gross_Sales'].sum()
        cogs = sub['Selected_COGS'].sum()
        margin = sales - cogs
        stats[b] = {
            'count': len(sub),
            'sales': sales,
            'margin': margin,
            'gp': (margin / sales * 100) if sales > 0 else 0,
            'qty': sub['Total_Qty'].sum()
        }

    def fmt_v(v):
        sign = "-" if v < 0 else ""
        av = abs(v)
        if av >= 1e9: return f"{sign}{av / 1e9:,.0f} Bn"
        if av >= 1e6: return f"{sign}{av / 1e6:,.0f} M"
        return f"{sign}{av:,.0f}"

    def make_box(box_id, bg_color):
        s = stats[box_id]
        return f"""
        <div style="background-color: {bg_color}; color: white; padding: 15px; border-radius: 4px; display: flex; flex-direction: column; justify-content: space-between; text-align: left; min-height: 120px; box-shadow: 1px 1px 4px rgba(0,0,0,0.1);">
            <div style="font-size: 12px; font-weight: 600; opacity: 0.9;">{box_id}</div>
            <div style="margin: 5px 0;">
                <span style="font-size: 28px; font-weight: 700;">{s['count']}</span> 
                <span style="font-size: 12px; opacity: 0.9;">SKUs</span>
            </div>
            <div style="font-size: 11px; font-weight: 600; margin-bottom: 2px;">IDR {fmt_v(s['margin'])} margin &bull; {s['gp']:.0f}% GP</div>
            <div style="font-size: 11px; opacity: 0.9;">Sales: IDR {fmt_v(s['sales'])} &bull; Qty: {s['qty']:,.0f}</div>
        </div>
        """

    html = f"""
    <div style="display: grid; grid-template-columns: 30px 70px 1fr 1fr 1fr; gap: 8px; font-family: sans-serif; margin-bottom: 10px;">
        <div></div><div></div>
        <div style="text-align: center; font-size: 12px; color: #64748b; font-weight: 600;">Sales &lt; Low</div>
        <div style="text-align: center; font-size: 12px; color: #64748b; font-weight: 600;">Med Sales</div>
        <div style="text-align: center; font-size: 12px; color: #64748b; font-weight: 600;">Sales &gt; High</div>

        <div style="grid-row: 2 / span 3; display: flex; align-items: center; justify-content: center;">
            <div style="writing-mode: vertical-rl; transform: rotate(180deg); font-size: 11px; font-weight: 700; color: #64748b; letter-spacing: 2px;">COGS (%)</div>
        </div>
        <div style="display: flex; flex-direction: column; justify-content: center; align-items: flex-end; padding-right: 10px; font-size: 12px; color: #475569;">
            <span style="font-weight: 600; color:#1e293b;">High</span>
            <span>&gt;{y_high_val:.1f}%</span>
        </div>
        {make_box('Box 1', '#c03d32')}
        {make_box('Box 4', '#d89f0e')}
        {make_box('Box 7', '#d89f0e')}

        <div style="display: flex; flex-direction: column; justify-content: center; align-items: flex-end; padding-right: 10px; font-size: 12px; color: #475569;">
            <span style="font-weight: 600; color:#1e293b;">Med</span>
            <span>{y_low_val:.1f}-{y_high_val:.1f}%</span>
        </div>
        {make_box('Box 2', '#d89f0e')}
        {make_box('Box 5', '#d89f0e')}
        {make_box('Box 8', '#319b5e')}

        <div style="display: flex; flex-direction: column; justify-content: center; align-items: flex-end; padding-right: 10px; font-size: 12px; color: #475569;">
            <span style="font-weight: 600; color:#1e293b;">Low</span>
            <span>&lt;{y_low_val:.1f}%</span>
        </div>
        {make_box('Box 3', '#3871b6')}
        {make_box('Box 6', '#3871b6')}
        {make_box('Box 9', '#319b5e')}
    </div>

    <div style="display: flex; justify-content: center; gap: 20px; font-family: sans-serif; font-size: 12px; font-weight: 600; color: #333; margin-bottom: 20px;">
        <div style="display: flex; align-items: center; gap: 5px;"><div style="width: 16px; height: 16px; background-color: #319b5e; border-radius: 2px;"></div> Grow</div>
        <div style="display: flex; align-items: center; gap: 5px;"><div style="width: 16px; height: 16px; background-color: #3871b6; border-radius: 2px;"></div> Keep</div>
        <div style="display: flex; align-items: center; gap: 5px;"><div style="width: 16px; height: 16px; background-color: #d89f0e; border-radius: 2px;"></div> Fix</div>
        <div style="display: flex; align-items: center; gap: 5px;"><div style="width: 16px; height: 16px; background-color: #c03d32; border-radius: 2px;"></div> Exit candidate</div>
    </div>
    """.replace('\n', '')
    st.markdown(html, unsafe_allow_html=True)


def apply_common_layout(fig, custom_box_counts=None, df_len=None):
    b_counts = custom_box_counts if custom_box_counts is not None else df_sku_valid[
        'Dynamic 9-Box Category'].value_counts().to_dict()
    tot_items = sum(b_counts.values()) if df_len is None else df_len

    fig.add_vline(x=apply_custom_x_scale(x_low_val), line_dash="dash", line_color="black", opacity=0.7,
                  annotation_text=f" Low ({fmt_idr_axis(x_low_val)})", annotation_position="top left")
    fig.add_vline(x=apply_custom_x_scale(x_high_val), line_dash="dash", line_color="black", opacity=0.7,
                  annotation_text=f" High ({fmt_idr_axis(x_high_val)})", annotation_position="top right")
    fig.add_hline(y=apply_custom_y_scale(y_low_val), line_dash="dash", line_color="black", opacity=0.7,
                  annotation_text=f"Low ({y_low_val:.1f}%)", annotation_position="bottom right")
    fig.add_hline(y=apply_custom_y_scale(y_high_val), line_dash="dash", line_color="black", opacity=0.7,
                  annotation_text=f"High ({y_high_val:.1f}%)", annotation_position="top right")
    fig.add_vline(x=apply_custom_x_scale(mid_x), line_dash="dot", line_color="red", opacity=0.6,
                  annotation_text=f" ← Center X ({fmt_idr_axis(mid_x)})", annotation_font_color="red",
                  annotation_position="bottom right")
    fig.add_hline(y=apply_custom_y_scale(mid_y), line_dash="dot", line_color="red", opacity=0.6,
                  annotation_text=f"Center Y ({mid_y:.1f}%)", annotation_position="top right",
                  annotation_font_color="red")

    mapped_x_low, mapped_x_high = apply_custom_x_scale(x_low_val), apply_custom_x_scale(x_high_val)
    mapped_plot_x_min, mapped_plot_x_max = apply_custom_x_scale(plot_x_min), apply_custom_x_scale(plot_x_max)
    mapped_y_low, mapped_y_high = apply_custom_y_scale(y_low_val), apply_custom_y_scale(y_high_val)
    mapped_plot_y_min, mapped_plot_y_max = apply_custom_y_scale(plot_y_min), apply_custom_y_scale(plot_y_max)

    b_coords = {
        'Box 1': {'x': (mapped_plot_x_min + mapped_x_low) / 2, 'y': (mapped_y_high + mapped_plot_y_max) / 2,
                  'id': 'B1'},
        'Box 2': {'x': (mapped_plot_x_min + mapped_x_low) / 2, 'y': (mapped_y_low + mapped_y_high) / 2, 'id': 'B2'},
        'Box 3': {'x': (mapped_plot_x_min + mapped_x_low) / 2, 'y': (mapped_plot_y_min + mapped_y_low) / 2, 'id': 'B3'},
        'Box 4': {'x': (mapped_x_low + mapped_x_high) / 2, 'y': (mapped_y_high + mapped_plot_y_max) / 2, 'id': 'B4'},
        'Box 5': {'x': (mapped_x_low + mapped_x_high) / 2, 'y': (mapped_y_low + mapped_y_high) / 2, 'id': 'B5'},
        'Box 6': {'x': (mapped_x_low + mapped_x_high) / 2, 'y': (mapped_plot_y_min + mapped_y_low) / 2, 'id': 'B6'},
        'Box 7': {'x': (mapped_x_high + mapped_plot_x_max) / 2, 'y': (mapped_y_high + mapped_plot_y_max) / 2,
                  'id': 'B7'},
        'Box 8': {'x': (mapped_x_high + mapped_plot_x_max) / 2, 'y': (mapped_y_low + mapped_y_high) / 2, 'id': 'B8'},
        'Box 9': {'x': (mapped_x_high + mapped_plot_x_max) / 2, 'y': (mapped_plot_y_min + mapped_y_low) / 2, 'id': 'B9'}
    }

    for b_name, c in b_coords.items():
        if b_counts.get(b_name, 0) > 0:
            fig.add_annotation(x=c['x'], y=c['y'], xref="x", yref="y",
                               text=f"<span style='color:#1f2937;'><b>{c['id']}</b></span><br><b>{b_counts[b_name]} SKUs</b>",
                               showarrow=False, font=dict(size=13, color="#374151"),
                               bgcolor="rgba(255, 255, 255, 0.85)", bordercolor="rgba(15, 23, 42, 0.3)", borderwidth=1,
                               borderpad=4)

    pareto_text = f"<br><span style='color:#fbbf24; font-size:11px;'><i>Pareto: {pareto_limit}%</i></span>" if pareto_limit < 100 else ""
    fig.add_annotation(x=0.98, y=0.98, xref="paper", yref="paper",
                       text=f"<b>TOTAL:<br>{tot_items} SKUs</b>{pareto_text}", showarrow=False,
                       font=dict(size=13, color="white"), bgcolor="#374151", bordercolor="black", borderwidth=1,
                       borderpad=5)

    ticks_x = sorted(list(set([plot_x_min, 0, x_low_val, mid_x, x_high_val, plot_x_max])))
    ticks_y = sorted(list(set([plot_y_min, 0, y_low_val, mid_y, y_high_val, plot_y_max])))

    fig.update_xaxes(range=[mapped_plot_x_min, mapped_plot_x_max], tickvals=[apply_custom_x_scale(v) for v in ticks_x],
                     ticktext=[fmt_idr_axis(v) for v in ticks_x], title_text="Gross Sales Absolute (IDR)")
    fig.update_yaxes(range=[mapped_plot_y_min, mapped_plot_y_max], tickvals=[apply_custom_y_scale(v) for v in ticks_y],
                     ticktext=[f"{v:.1f}%" for v in ticks_y], title_text=f"{cogs_selector} (%)")
    fig.update_layout(height=650, hovermode="closest",
                      legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, title=None),
                      margin=dict(t=50, b=50, l=50, r=50))
    return fig


hover_conf = {'Format_Sales': True, 'Format_COGS': True, 'COGS (%)': ':.2f', 'Format_Qty': True,
              'Dynamic 9-Box Category': False, 'Price_Increase_Label': False, 'Bubble_Size': False, 'X_Plot': False,
              'Y_Plot': False, 'Produk_Key_Str': True}
color_map_main = {'Box 1': '#c03d32', 'Box 2': '#d89f0e', 'Box 3': '#3871b6', 'Box 4': '#d89f0e', 'Box 5': '#d89f0e',
                  'Box 6': '#3871b6', 'Box 7': '#d89f0e', 'Box 8': '#319b5e', 'Box 9': '#319b5e'}

st.markdown("---")
tab_main, tab_pi, tab_tgt = st.tabs(["📊 Main COGS 9-Box", "📈 Price Increase Tracker", "🎯 Target SKU Intersect"])

with tab_main:
    display_9box_matrix(df_sku_valid)
    fig_main = px.scatter(df_sku_valid, x='X_Plot', y='Y_Plot', size='Bubble_Size', color='Dynamic 9-Box Category',
                          color_discrete_map=color_map_main, hover_name='Product Name', hover_data=hover_conf,
                          labels={'Format_Sales': 'Gross Sales', 'Format_COGS': cogs_selector,
                                  'Format_Qty': 'Quantity'},
                          title=f"9-Box Bubble Chart: {cogs_selector} (%) vs Gross Sales ({chart_period_title})",
                          size_max=60, render_mode='svg')
    st.plotly_chart(apply_common_layout(fig_main, df_len=len(df_sku_valid)), use_container_width=True)

with tab_pi:
    if not has_price_data:
        st.warning(
            "⚠️ Price Increase identification data is not available. Please upload a multi-sheet Excel file containing the 'Price List 2026' sheet.")
    else:
        fig_pi = px.scatter(df_sku_valid, x='X_Plot', y='Y_Plot', size='Bubble_Size', color='Price_Increase_Label',
                            color_discrete_map={'Price Increased': '#319b5e', 'Stable Price': '#3871b6'},
                            hover_name='Product Name', hover_data=hover_conf,
                            labels={'Format_Sales': 'Gross Sales', 'Format_COGS': cogs_selector,
                                    'Format_Qty': 'Quantity', 'Price_Increase_Label': 'Price Status'},
                            title=f"Price Increase Tracker: {cogs_selector} (%) vs Gross Sales ({chart_period_title})",
                            size_max=60, render_mode='svg')
        fig_pi = apply_common_layout(fig_pi, df_len=len(df_sku_valid))

        count_naik, count_tetap = (df_sku_valid['Price_Increase_Label'] == 'Price Increased').sum(), (
                    df_sku_valid['Price_Increase_Label'] == 'Stable Price').sum()
        fig_pi.add_annotation(x=0.01, y=0.98, xref="paper", yref="paper",
                              text=f"<b>Price Status:</b><br><span style='color:#319b5e'>Price Increased: {count_naik} SKUs</span><br><span style='color:#3871b6'>Stable Price: {count_tetap} SKUs</span>",
                              showarrow=False, font=dict(size=13, color="#1f2937"), bgcolor="rgba(255, 255, 255, 0.95)",
                              bordercolor="rgba(15, 23, 42, 0.3)", borderwidth=1, borderpad=6, align="left")
        st.plotly_chart(fig_pi, use_container_width=True)

        missing_pi_keys = set(pi_produk_keys) - set(
            df_sku_valid[df_sku_valid['Price_Increase']]['Produk_Key_Str'].dropna())
        if missing_pi_keys:
            missing_df = df_pl[df_pl['Produk_Key'].isin(missing_pi_keys)][
                ['Produk_Key', 'Product Name', 'Currency']].drop_duplicates(subset=['Produk_Key']).copy()
            missing_df['Price Increase Period'] = missing_df['Produk_Key'].map(pi_period_dict).dt.strftime(
                '%Y-%m-%d').fillna('-')

            keys_in_raw, keys_in_sku_pre = set(df_filtered['Produk_Key_Str'].dropna()), set(
                df_sku['Produk_Key_Str'].dropna())
            missing_df['Missing Reason'] = missing_df['Produk_Key'].apply(
                lambda k: "Missing (No transaction)" if k not in keys_in_raw else (
                    "Filtered (Sales/COGS < Rp 1)" if k not in keys_in_sku_pre else "Removed by Pareto Filter"))

            st.markdown(f"**⚠️ There are {len(missing_pi_keys)} Price Increased SKUs not plotted on the main chart:**")
            col_mis1, _ = st.columns([1, 4])
            with col_mis1:
                buffer_missing = io.BytesIO()
                with pd.ExcelWriter(buffer_missing, engine='openpyxl') as writer: missing_df.to_excel(writer,
                                                                                                      index=False)
                st.download_button("📥 Download Missing SKUs (Excel)", data=buffer_missing.getvalue(),
                                   file_name=f"Missing_SKUs_Price_Increase.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.dataframe(missing_df, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("#### 🔄 Before & After Price Increase Analysis")
        st.caption(f"ℹ️ **Before** = histori dari awal file sampai {start_month_str} sebelum tanggal naik harga tiap SKU. "
                   f"**After** = {chart_period_title}, mulai tanggal naik harga.")
        toggle_mode = st.radio("Select Timeline Mode:", ["Before Price Increase", "After Price Increase"],
                               horizontal=True)

        df_pi_trx = df_filtered[df_filtered['Price_Increase_Flag'] == True].copy()


        def get_agg_valid(df_sub):
            if df_sub.empty: return pd.DataFrame()
            d_agg = df_sub.groupby('Produk_Key_Str').agg(Product_Name=('Product Name', 'first'),
                                                         Total_Gross_Sales=('Calc_Gross_Sales', 'sum'),
                                                         Selected_COGS=(cogs_selector, 'sum'),
                                                         Total_Qty=('Quantity', 'sum')).reset_index()
            return d_agg[(d_agg['Total_Gross_Sales'] >= 1) & (d_agg['Selected_COGS'] >= 1)].rename(
                columns={'Product_Name': 'Product Name'})


        df_pi_hist = df_hist_m[df_hist_m['Price_Increase_Flag'] == True]
        df_before_agg = get_agg_valid(df_pi_hist[(df_pi_hist['Date'] <= before_end_date) & (df_pi_hist['Date'] < df_pi_hist['Price_Increase_Period'])])
        df_after_agg = get_agg_valid(df_pi_trx[df_pi_trx['Date'] >= df_pi_trx['Price_Increase_Period']])

        if df_before_agg.empty or df_after_agg.empty:
            st.warning("⚠️ Not enough data to perform Before-After comparison.")
        else:
            common_skus = set(df_before_agg['Produk_Key_Str']).intersection(set(df_after_agg['Produk_Key_Str']))
            df_before_agg['COGS (%)'] = (df_before_agg['Selected_COGS'] / df_before_agg['Total_Gross_Sales']) * 100
            df_before_agg['Dynamic 9-Box Category'] = assign_9box(df_before_agg)
            df_after_agg['COGS (%)'] = (df_after_agg['Selected_COGS'] / df_after_agg['Total_Gross_Sales']) * 100
            df_after_agg['Dynamic 9-Box Category'] = assign_9box(df_after_agg)

            df_ba = df_before_agg[
                df_before_agg['Produk_Key_Str'].isin(common_skus)].copy() if "Before" in toggle_mode else df_after_agg[
                df_after_agg['Produk_Key_Str'].isin(common_skus)].copy()
            chart_title_ba = "BEFORE Price Increase" if "Before" in toggle_mode else "AFTER Price Increase"

            if df_ba.empty:
                st.warning("No intersection data remains.")
            else:
                df_ba['Bubble_Size'] = df_ba['Total_Qty'].abs().replace(0, 1)
                df_ba['Price_Increase_Label'] = 'Price Increased'
                df_ba['X_Plot'] = df_ba['Total_Gross_Sales'].apply(apply_custom_x_scale)
                df_ba['Y_Plot'] = df_ba['COGS (%)'].apply(apply_custom_y_scale)
                df_ba['Format_Sales'] = df_ba['Total_Gross_Sales'].apply(lambda x: f"Rp {x:,.0f}")
                df_ba['Format_COGS'] = df_ba['Selected_COGS'].apply(lambda x: f"Rp {x:,.0f}")
                df_ba['Format_Qty'] = df_ba['Total_Qty'].apply(lambda x: f"{x:,.0f} Pcs")

                display_9box_matrix(df_ba)
                b_counts_ba = df_ba['Dynamic 9-Box Category'].value_counts().to_dict()
                fig_ba = px.scatter(df_ba, x='X_Plot', y='Y_Plot', size='Bubble_Size', color='Dynamic 9-Box Category',
                                    color_discrete_map=color_map_main, hover_name='Product Name', hover_data=hover_conf,
                                    labels={'Format_Sales': 'Gross Sales', 'Format_COGS': cogs_selector},
                                    title=f"{chart_title_ba} Matrix | Calculated Total: {len(df_ba)} Apples-to-Apples SKUs",
                                    size_max=60, render_mode='svg')
                st.plotly_chart(apply_common_layout(fig_ba, custom_box_counts=b_counts_ba, df_len=len(df_ba)),
                                use_container_width=True)

                st.markdown("#### 🧭 SKU Migration Tracker")
                df_migration = pd.merge(df_before_agg[df_before_agg['Produk_Key_Str'].isin(common_skus)][
                                            ['Produk_Key_Str', 'Product Name', 'Total_Gross_Sales', 'COGS (%)',
                                             'Dynamic 9-Box Category']],
                                        df_after_agg[df_after_agg['Produk_Key_Str'].isin(common_skus)][
                                            ['Produk_Key_Str', 'Total_Gross_Sales', 'COGS (%)',
                                             'Dynamic 9-Box Category']], on='Produk_Key_Str',
                                        suffixes=(' (Before)', ' (After)'))
                df_migration['Movement Status'] = df_migration.apply(
                    lambda r: f"Stays in Box {r['Dynamic 9-Box Category (Before)'].split(' ')[1]}" if r[
                                                                                                          'Dynamic 9-Box Category (Before)'] ==
                                                                                                      r[
                                                                                                          'Dynamic 9-Box Category (After)'] else f"Moved: Box {r['Dynamic 9-Box Category (Before)'].split(' ')[1]} ➡️ Box {r['Dynamic 9-Box Category (After)'].split(' ')[1]}",
                    axis=1)

                col_mig1, _ = st.columns([1, 4])
                with col_mig1:
                    buffer_mig = io.BytesIO()
                    with pd.ExcelWriter(buffer_mig, engine='openpyxl') as writer: df_migration.to_excel(writer,
                                                                                                        index=False)
                    st.download_button("📥 Download Migration Tracker", data=buffer_mig.getvalue(),
                                       file_name=f"SKU_Migration_Tracker.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       key="btn_mig_main")
                st.dataframe(df_migration[['Produk_Key_Str', 'Product Name', 'Movement Status',
                                           'Dynamic 9-Box Category (Before)', 'Dynamic 9-Box Category (After)',
                                           'Total_Gross_Sales (Before)', 'Total_Gross_Sales (After)',
                                           'COGS (%) (Before)', 'COGS (%) (After)']].style.format(
                    {'Total_Gross_Sales (Before)': 'Rp {:,.0f}', 'Total_Gross_Sales (After)': 'Rp {:,.0f}',
                     'COGS (%) (Before)': '{:.2f}%', 'COGS (%) (After)': '{:.2f}%'}), use_container_width=True,
                             hide_index=True)

with tab_tgt:
    if not target_skus:
        st.warning("⚠️ Target SKU data not loaded. Please upload the Target SKU file in the uploader section above.")
    else:
        st.markdown("### 🔀 1602 SKU View Version")
        v_mode = st.radio("Select 1602 Portfolio Version: ❔", [
            f"Versi 1: 1602 Original SKUs (Granular - {len(target_skus):,} Target SKUs)",
            f"Versi 2: Commonized SKUs by New Code (Consolidated - {len(target_new_skus):,} Target SKUs)"
        ], horizontal=True)
        st.markdown("---")

        if "Versi 1" in v_mode:
            df_target_intersect = df_sku_valid[df_sku_valid['Produk_Key_Str'].isin(target_skus)].copy()
        else:
            df_tgt_raw = df_sku_valid[df_sku_valid['Produk_Key_Str'].isin(target_skus)].copy()
            df_tgt_raw['Commonized_Key'] = df_tgt_raw['Produk_Key_Str'].map(target_mapping).fillna(
                df_tgt_raw['Produk_Key_Str'])

            df_target_intersect = df_tgt_raw.groupby('Commonized_Key').agg(
                Total_Gross_Sales=('Total_Gross_Sales', 'sum'),
                Selected_COGS=('Selected_COGS', 'sum'),
                Total_Qty=('Total_Qty', 'sum')
            ).reset_index()

            if target_name_mapping:
                df_target_intersect['Product Name'] = df_target_intersect['Commonized_Key'].map(
                    target_name_mapping).fillna("Commonized SKU " + df_target_intersect['Commonized_Key'].astype(str))
            else:
                df_target_intersect['Product Name'] = "Commonized SKU " + df_target_intersect['Commonized_Key'].astype(
                    str)

            df_target_intersect['COGS (%)'] = (df_target_intersect['Selected_COGS'] / df_target_intersect[
                'Total_Gross_Sales']) * 100
            df_target_intersect['Bubble_Size'] = df_target_intersect['Total_Qty'].abs().replace(0, 1)
            df_target_intersect['Dynamic 9-Box Category'] = assign_9box(df_target_intersect)
            df_target_intersect['X_Plot'] = df_target_intersect['Total_Gross_Sales'].apply(apply_custom_x_scale)
            df_target_intersect['Y_Plot'] = df_target_intersect['COGS (%)'].apply(apply_custom_y_scale)

            df_target_intersect['Format_Sales'] = df_target_intersect['Total_Gross_Sales'].apply(
                lambda x: f"Rp {x:,.0f}")
            df_target_intersect['Format_COGS'] = df_target_intersect['Selected_COGS'].apply(lambda x: f"Rp {x:,.0f}")
            df_target_intersect['Format_Qty'] = df_target_intersect['Total_Qty'].apply(lambda x: f"{x:,.0f} Pcs")
            df_target_intersect['Price_Increase_Label'] = 'Consolidated'
            df_target_intersect.rename(columns={'Commonized_Key': 'Produk_Key_Str'}, inplace=True)

        if df_target_intersect.empty:
            st.warning(
                "⚠️ No intersecting data found between the uploaded Target SKUs and the currently filtered valid SKUs.")
        else:
            display_9box_matrix(df_target_intersect)
            b_counts_tgt = df_target_intersect['Dynamic 9-Box Category'].value_counts().to_dict()
            fig_tgt = px.scatter(df_target_intersect, x='X_Plot', y='Y_Plot', size='Bubble_Size',
                                 color='Dynamic 9-Box Category', color_discrete_map=color_map_main,
                                 hover_name='Product Name', hover_data=hover_conf,
                                 labels={'Format_Sales': 'Gross Sales', 'Format_COGS': cogs_selector,
                                         'Format_Qty': 'Quantity'},
                                 title=f"Target SKU Intersection: {len(df_target_intersect)} SKUs matched", size_max=60,
                                 render_mode='svg')
            st.plotly_chart(
                apply_common_layout(fig_tgt, custom_box_counts=b_counts_tgt, df_len=len(df_target_intersect)),
                use_container_width=True)

            st.markdown("---")
            st.markdown("#### 🔄 Before & After Price Increase Analysis (Target SKUs)")
            st.caption(f"ℹ️ **Before** = histori dari awal file sampai {start_month_str} sebelum tanggal naik harga tiap SKU. "
                       f"**After** = {chart_period_title}, mulai tanggal naik harga.")
            toggle_mode_tgt = st.radio("Select Timeline Mode:", ["Before Price Increase", "After Price Increase"],
                                       horizontal=True, key="toggle_tgt_ba")

            df_pi_trx_tgt = df_filtered[
                (df_filtered['Price_Increase_Flag'] == True) & (df_filtered['Produk_Key_Str'].isin(target_skus))].copy()


            def get_tgt_agg_valid(df_sub):
                if df_sub.empty: return pd.DataFrame()
                df_r = df_sub.copy()
                if "Versi 1" in v_mode:
                    df_r['Agg_Key'] = df_r['Produk_Key_Str']
                else:
                    df_r['Agg_Key'] = df_r['Produk_Key_Str'].map(target_mapping).fillna(df_r['Produk_Key_Str'])

                d_agg = df_r.groupby('Agg_Key').agg(
                    Total_Gross_Sales=('Calc_Gross_Sales', 'sum'),
                    Selected_COGS=(cogs_selector, 'sum'),
                    Total_Qty=('Quantity', 'sum')
                ).reset_index()
                d_agg = d_agg[(d_agg['Total_Gross_Sales'] >= 1) & (d_agg['Selected_COGS'] >= 1)].copy()

                if "Versi 1" in v_mode:
                    name_dict = df_sku.set_index('Produk_Key_Str')['Product Name'].to_dict()
                    d_agg['Product Name'] = d_agg['Agg_Key'].map(name_dict).fillna("Unknown SKU")
                else:
                    if target_name_mapping:
                        d_agg['Product Name'] = d_agg['Agg_Key'].map(target_name_mapping).fillna(
                            "Commonized SKU " + d_agg['Agg_Key'].astype(str))
                    else:
                        d_agg['Product Name'] = "Commonized SKU " + d_agg['Agg_Key'].astype(str)

                d_agg.rename(columns={'Agg_Key': 'Produk_Key_Str'}, inplace=True)
                return d_agg


            df_pi_hist_tgt = df_hist_m[(df_hist_m['Price_Increase_Flag'] == True) & (df_hist_m['Produk_Key_Str'].isin(target_skus))]
            df_tgt_before_agg = get_tgt_agg_valid(
                df_pi_hist_tgt[(df_pi_hist_tgt['Date'] <= before_end_date) & (df_pi_hist_tgt['Date'] < df_pi_hist_tgt['Price_Increase_Period'])])
            df_tgt_after_agg = get_tgt_agg_valid(
                df_pi_trx_tgt[df_pi_trx_tgt['Date'] >= df_pi_trx_tgt['Price_Increase_Period']])

            if df_tgt_before_agg.empty or df_tgt_after_agg.empty:
                st.warning("⚠️ Not enough data to perform Before-After comparison for Target SKUs.")
            else:
                common_tgt_skus = set(df_tgt_before_agg['Produk_Key_Str']).intersection(
                    set(df_tgt_after_agg['Produk_Key_Str']))

                df_tgt_before_agg['COGS (%)'] = (df_tgt_before_agg['Selected_COGS'] / df_tgt_before_agg[
                    'Total_Gross_Sales']) * 100
                df_tgt_before_agg['Dynamic 9-Box Category'] = assign_9box(df_tgt_before_agg)
                df_tgt_after_agg['COGS (%)'] = (df_tgt_after_agg['Selected_COGS'] / df_tgt_after_agg[
                    'Total_Gross_Sales']) * 100
                df_tgt_after_agg['Dynamic 9-Box Category'] = assign_9box(df_tgt_after_agg)

                df_ba_tgt = df_tgt_before_agg[df_tgt_before_agg['Produk_Key_Str'].isin(
                    common_tgt_skus)].copy() if "Before" in toggle_mode_tgt else df_tgt_after_agg[
                    df_tgt_after_agg['Produk_Key_Str'].isin(common_tgt_skus)].copy()
                chart_title_ba_tgt = "BEFORE Price Increase" if "Before" in toggle_mode_tgt else "AFTER Price Increase"

                if df_ba_tgt.empty:
                    st.warning("No intersection data remains for Target SKUs.")
                else:
                    df_ba_tgt['Bubble_Size'] = df_ba_tgt['Total_Qty'].abs().replace(0, 1)
                    df_ba_tgt['Price_Increase_Label'] = 'Price Increased'
                    df_ba_tgt['X_Plot'] = df_ba_tgt['Total_Gross_Sales'].apply(apply_custom_x_scale)
                    df_ba_tgt['Y_Plot'] = df_ba_tgt['COGS (%)'].apply(apply_custom_y_scale)
                    df_ba_tgt['Format_Sales'] = df_ba_tgt['Total_Gross_Sales'].apply(lambda x: f"Rp {x:,.0f}")
                    df_ba_tgt['Format_COGS'] = df_ba_tgt['Selected_COGS'].apply(lambda x: f"Rp {x:,.0f}")
                    df_ba_tgt['Format_Qty'] = df_ba_tgt['Total_Qty'].apply(lambda x: f"{x:,.0f} Pcs")

                    display_9box_matrix(df_ba_tgt)
                    b_counts_ba_tgt = df_ba_tgt['Dynamic 9-Box Category'].value_counts().to_dict()
                    fig_ba_tgt = px.scatter(df_ba_tgt, x='X_Plot', y='Y_Plot', size='Bubble_Size',
                                            color='Dynamic 9-Box Category', color_discrete_map=color_map_main,
                                            hover_name='Product Name', hover_data=hover_conf,
                                            labels={'Format_Sales': 'Gross Sales', 'Format_COGS': cogs_selector},
                                            title=f"{chart_title_ba_tgt} Matrix | Target SKUs: {len(df_ba_tgt)} Apples-to-Apples",
                                            size_max=60, render_mode='svg')
                    st.plotly_chart(
                        apply_common_layout(fig_ba_tgt, custom_box_counts=b_counts_ba_tgt, df_len=len(df_ba_tgt)),
                        use_container_width=True)

                    st.markdown("#### 🧭 Target SKU Migration Tracker")
                    df_migration_tgt = pd.merge(
                        df_tgt_before_agg[df_tgt_before_agg['Produk_Key_Str'].isin(common_tgt_skus)][
                            ['Produk_Key_Str', 'Product Name', 'Total_Gross_Sales', 'COGS (%)',
                             'Dynamic 9-Box Category']],
                        df_tgt_after_agg[df_tgt_after_agg['Produk_Key_Str'].isin(common_tgt_skus)][
                            ['Produk_Key_Str', 'Total_Gross_Sales', 'COGS (%)', 'Dynamic 9-Box Category']],
                        on='Produk_Key_Str', suffixes=(' (Before)', ' (After)'))
                    df_migration_tgt['Movement Status'] = df_migration_tgt.apply(
                        lambda r: f"Stays in Box {r['Dynamic 9-Box Category (Before)'].split(' ')[1]}" if r[
                                                                                                              'Dynamic 9-Box Category (Before)'] ==
                                                                                                          r[
                                                                                                              'Dynamic 9-Box Category (After)'] else f"Moved: Box {r['Dynamic 9-Box Category (Before)'].split(' ')[1]} ➡️ Box {r['Dynamic 9-Box Category (After)'].split(' ')[1]}",
                        axis=1)

                    col_mig2, _ = st.columns([1, 4])
                    with col_mig2:
                        buffer_mig_tgt = io.BytesIO()
                        with pd.ExcelWriter(buffer_mig_tgt, engine='openpyxl') as writer: df_migration_tgt.to_excel(
                            writer, index=False)
                        st.download_button("📥 Download Migration Tracker", data=buffer_mig_tgt.getvalue(),
                                           file_name=f"Target_SKU_Migration_Tracker.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                           key="btn_mig_tgt")
                    st.dataframe(df_migration_tgt[['Produk_Key_Str', 'Product Name', 'Movement Status',
                                                   'Dynamic 9-Box Category (Before)', 'Dynamic 9-Box Category (After)',
                                                   'Total_Gross_Sales (Before)', 'Total_Gross_Sales (After)',
                                                   'COGS (%) (Before)', 'COGS (%) (After)']].style.format(
                        {'Total_Gross_Sales (Before)': 'Rp {:,.0f}', 'Total_Gross_Sales (After)': 'Rp {:,.0f}',
                         'COGS (%) (Before)': '{:.2f}%', 'COGS (%) (After)': '{:.2f}%'}), use_container_width=True,
                                 hide_index=True)

st.markdown("---")
st.markdown("### 📋 SKU Details")
drop_cols = ['Bubble_Size', 'X_Plot', 'Y_Plot', 'Format_Sales', 'Format_COGS', 'Format_Qty', 'Price_Increase',
             'Price_Increase_Period', 'Cum_Sales']
df_display = df_sku_valid.drop(columns=[c for c in drop_cols if c in df_sku_valid.columns]).sort_values(
    'Total_Gross_Sales', ascending=False)
col_dl1, _ = st.columns([1, 4])
with col_dl1:
    buffer_excel = io.BytesIO()
    with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer: df_display.to_excel(writer, index=False,
                                                                                        sheet_name='COGS_Details')
    st.download_button("📥 Download Data (Excel)", data=buffer_excel.getvalue(),
                       file_name=f"COGS_9Box_{cogs_selector.replace(' ', '_')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
st.dataframe(df_display.style.format(
    {'Total_Gross_Sales': 'Rp {:,.0f}', 'Total_COGS_Reg': 'Rp {:,.0f}', 'Total_COGS_All': 'Rp {:,.0f}',
     'Selected_COGS': 'Rp {:,.0f}', 'COGS (%)': '{:.2f}%', 'Total_Qty': '{:,.0f}'}), use_container_width=True,
             hide_index=True)
