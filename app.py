import streamlit as st
import pandas as pd
import yfinance as yf
import numpy as np
import io
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, ScatterChart, Reference, Series
from openpyxl.chart.trendline import Trendline
from openpyxl.formatting.rule import ColorScaleRule

# ==========================================
# 1. HELPER FUNCTIONS
# ==========================================

def format_indian_symbol(raw_input: str) -> str:
    symbol = raw_input.strip().upper()
    index_map = {
        "NIFTY": "^NSEI", "NIFTY 50": "^NSEI", "NIFTY50": "^NSEI",
        "SENSEX": "^BSESN", "BSE SENSEX": "^BSESN", "BANKNIFTY": "^NSEBANK",
        "NIFTY BANK": "^NSEBANK", "NIFTY IT": "^CNXIT",
        "NIFTY MIDCAP": "^NSEMDCP50", "NIFTY SMALLCAP": "^NSI200",
    }
    if symbol in index_map: return index_map[symbol]
    if symbol.startswith("^") or symbol.endswith(".NS") or symbol.endswith(".BO"):
        return symbol
    return f"{symbol}.NS"

def clean_sheet_name(name: str) -> str:
    return name.replace("^", "").replace(":", "_").replace("/", "_").strip()[:30].upper()

@st.cache_data(show_spinner=False)
def get_financial_metrics(ticker_symbol: str) -> tuple[float, float]:
    de_ratio, tax_rate = 0.0, 0.25 
    try:
        t = yf.Ticker(ticker_symbol)
        info = t.info or {}
        
        d_e_raw = info.get("debtToEquity")
        if d_e_raw is not None and isinstance(d_e_raw, (int, float)):
            de_ratio = (d_e_raw / 100.0 if d_e_raw > 5 else float(d_e_raw))
        else:
            try:
                bs = t.balance_sheet
                if not bs.empty and "Total Debt" in bs.index and "Stockholders Equity" in bs.index:
                    total_debt = bs.loc["Total Debt"].iloc[0]
                    total_equity = bs.loc["Stockholders Equity"].iloc[0]
                    if total_equity and total_equity != 0:
                        de_ratio = float(total_debt / total_equity)
            except: pass

        try:
            inc = t.income_stmt
            if not inc.empty and "Tax Provision" in inc.index and "Pretax Income" in inc.index:
                tax_provision = inc.loc["Tax Provision"].iloc[0]
                pretax_income = inc.loc["Pretax Income"].iloc[0]
                if pretax_income > 0 and tax_provision > 0:
                    calc_tax = float(tax_provision / pretax_income)
                    if 0.05 <= calc_tax <= 0.45: tax_rate = calc_tax
        except: pass
    except: pass
    return round(de_ratio, 4), round(tax_rate, 4)

# ==========================================
# 2. EXCEL GENERATION LOGIC (WITH CHARTS)
# ==========================================

def generate_market_excel(
    data_dict: dict,
    financials_dict: dict,
    main_assets: list,
    comp_map: dict,
    benchmark_symbol: str,
    start_date: str,
    end_date: str,
    dash_df: pd.DataFrame
) -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    NAVY_FILL = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    HEADER_FILL = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    ZEBRA_FILL = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    FONT_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    FONT_BOLD = Font(name="Calibri", size=10, bold=True)
    FONT_METRIC = Font(name="Calibri", size=11, bold=True, color="1B365D")

    BORDER_THIN = Border(
        left=Side(style="thin", color="D5D8DC"), right=Side(style="thin", color="D5D8DC"),
        top=Side(style="thin", color="D5D8DC"), bottom=Side(style="thin", color="D5D8DC"),
    )
    ALIGN_CENTER = Alignment(horizontal="center", vertical="center")

    all_symbols = list(data_dict.keys())
    ordered_sheets = [s for s in main_assets if s in all_symbols]
    ordered_sheets += [s for s in all_symbols if s not in main_assets and s != benchmark_symbol]
    if benchmark_symbol in all_symbols:
        ordered_sheets.append(benchmark_symbol)

    bench_clean = clean_sheet_name(benchmark_symbol)

    for user_symbol in ordered_sheets:
        df = data_dict[user_symbol]
        sheet_title = clean_sheet_name(user_symbol)
        ws = wb.create_sheet(title=sheet_title)
        de_ratio, tax_rate = financials_dict.get(user_symbol, (0.0, 0.25))

        ws.merge_cells("A1:G1")
        ws["A1"] = f"MARKET DATA: {sheet_title}"
        ws["A1"].font = FONT_TITLE
        ws["A1"].fill = NAVY_FILL

        headers = ["Date", "Open", "High", "Low", "Close", "Volume", "Daily Return"]
        for c_i, h_text in enumerate(headers, start=1):
            c = ws.cell(row=3, column=c_i, value=h_text)
            c.fill, c.font, c.alignment, c.border = HEADER_FILL, FONT_HEADER, ALIGN_CENTER, BORDER_THIN

        start_row = 4
        for i, (idx_dt, row_data) in enumerate(df.iterrows()):
            r_idx = start_row + i
            ws.cell(row=r_idx, column=1, value=idx_dt.strftime("%Y-%m-%d")).alignment = ALIGN_CENTER
            for col, val in enumerate(["Open", "High", "Low", "Close"], start=2):
                ws.cell(row=r_idx, column=col, value=float(row_data[val])).number_format = "₹#,##0.00"
            ws.cell(row=r_idx, column=6, value=int(row_data["Volume"])).number_format = "#,##0"
            
            if i == 0: ws.cell(row=r_idx, column=7, value=0.0).number_format = "0.00%"
            else: ws.cell(row=r_idx, column=7, value=f"=(E{r_idx}-E{r_idx-1})/E{r_idx-1}").number_format = "0.00%"
            
            for c in range(1, 8):
                ws.cell(row=r_idx, column=c).border = BORDER_THIN
                if i % 2 == 1: ws.cell(row=r_idx, column=c).fill = ZEBRA_FILL

        end_row = start_row + len(df) - 1

        # NOTE: Statistical functions start at G5 to exclude row 4 (0.0% return on Day 1)
        metrics = [
            ("Avg Daily Return", f"=AVERAGE(G5:G{end_row})", "0.000%"),
            ("Annual Return", "=J4*252", "0.00%"),
            ("Daily Vol", f"=_xlfn.STDEV.S(G5:G{end_row})", "0.000%"),
            ("Annual Vol", "=J6*SQRT(252)", "0.00%"),
            ("Daily Var", f"=_xlfn.VAR.S(G5:G{end_row})", "0.00000"),
            (f"Covariance ({bench_clean})", f"=_xlfn.COVARIANCE.S(G5:G{end_row}, '{bench_clean}'!G5:G{end_row})", "0.00000"),
            (f"Bench Var", f"=_xlfn.VAR.S('{bench_clean}'!G5:G{end_row})", "0.00000"),
            ("Beta (Ratio)", "=IF(J10=0, 1, J9/J10)", "0.00"),
            ("Levered Beta", f"=SLOPE(G5:G{end_row}, '{bench_clean}'!G5:G{end_row})", "0.00"),
            ("D/E Ratio", de_ratio, "0.00%"),
            ("Tax Rate", tax_rate, "0.00%"),
            ("Unlevered Beta", "=J12/(1+(1-J14)*J13)", "0.00"),
            ("Correlation", f"=CORREL(G5:G{end_row}, '{bench_clean}'!G5:G{end_row})", "0.000")
        ]
        
        for m_idx, (label, form, fmt) in enumerate(metrics, start=4):
            ws.cell(row=m_idx, column=9, value=label).font = FONT_BOLD
            c_val = ws.cell(row=m_idx, column=10, value=form)
            c_val.font = FONT_METRIC
            c_val.number_format = fmt

        ws.column_dimensions["G"].width = 15
        ws.column_dimensions["I"].width = 25
        ws.column_dimensions["J"].width = 15

        # Regression Scatter Chart inside Target Asset worksheets
        if user_symbol in main_assets and benchmark_symbol in data_dict:
            sc = ScatterChart()
            sc.title = f"{sheet_title} vs {bench_clean} Regression"
            sc.style = 13
            sc.x_axis.title = f"{bench_clean} Returns"
            sc.y_axis.title = f"{sheet_title} Returns"
            sc.width, sc.height = 15, 10
            
            bench_sheet = wb[bench_clean]
            x_values = Reference(bench_sheet, min_col=7, min_row=5, max_row=end_row)
            y_values = Reference(ws, min_col=7, min_row=5, max_row=end_row)
            
            series = Series(y_values, x_values, title_from_data=False)
            series.marker.symbol = "circle"
            series.graphicalProperties.line.noFill = True
            series.trendline = Trendline(trendlineType="linear")
            
            sc.series.append(series)
            ws.add_chart(sc, "M15")

    # ==========================================
    # Executive Summary Sheet
    # ==========================================
    ws_summary = wb.create_sheet(title="Executive Summary", index=0)
    ws_summary.views.sheetView[0].showGridLines = True
    ws_summary.merge_cells("A1:K1")
    ws_summary["A1"] = f"QUANTITATIVE ANALYSIS ({start_date} to {end_date})"
    ws_summary["A1"].font, ws_summary["A1"].fill = FONT_TITLE, NAVY_FILL

    headers_summary = ["Asset", "Avg Daily Ret", "Ann Return", "Ann Volatility", "Levered Beta", "D/E Ratio", "Tax Rate", "Unlevered Beta", "Peer Unlevered Beta", "Peer Relevered Beta", "Correlation"]
    
    for col_idx, text in enumerate(headers_summary, start=1):
        c = ws_summary.cell(row=3, column=col_idx, value=text)
        c.fill, c.font, c.alignment, c.border = HEADER_FILL, FONT_HEADER, ALIGN_CENTER, BORDER_THIN
        ws_summary.column_dimensions[get_column_letter(col_idx)].width = 20

    for idx, row in dash_df.iterrows():
        s_row = 4 + idx
        ws_summary.cell(row=s_row, column=1, value=row["Asset"]).font = FONT_BOLD
        ws_summary.cell(row=s_row, column=2, value=row["Ann Return"]/252).number_format = "0.000%"
        ws_summary.cell(row=s_row, column=3, value=row["Ann Return"]).number_format = "0.00%"
        ws_summary.cell(row=s_row, column=4, value=row["Ann Volatility"]).number_format = "0.00%"
        ws_summary.cell(row=s_row, column=5, value=row["Raw Levered Beta"]).number_format = "0.00"
        ws_summary.cell(row=s_row, column=6, value=row["D/E Ratio"]).number_format = "0.00%"
        ws_summary.cell(row=s_row, column=7, value=financials_dict.get(row["Asset"], (0.0, 0.25))[1]).number_format = "0.00%"
        ws_summary.cell(row=s_row, column=8, value=row["Unlevered Beta"]).number_format = "0.00"
        ws_summary.cell(row=s_row, column=9, value=row["Peer Avg Unlevered Beta"]).number_format = "0.00"
        ws_summary.cell(row=s_row, column=10, value=row["Peer Relevered Beta"]).number_format = "0.00"
        ws_summary.cell(row=s_row, column=11, value=row["Correlation"]).number_format = "0.000"

        for c_i in range(1, 12):
            ws_summary.cell(row=s_row, column=c_i).border = BORDER_THIN

    # Native Excel Bar Chart
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Beta Comparison by Asset"
    chart.y_axis.title = "Beta Value"
    chart.x_axis.title = "Assets"

    cats_ref = Reference(ws_summary, min_col=1, min_row=4, max_row=3 + len(dash_df))
    s1 = Series(Reference(ws_summary, min_col=5, min_row=3, max_row=3 + len(dash_df)), title_from_data=True)
    s2 = Series(Reference(ws_summary, min_col=8, min_row=3, max_row=3 + len(dash_df)), title_from_data=True)
    s3 = Series(Reference(ws_summary, min_col=10, min_row=3, max_row=3 + len(dash_df)), title_from_data=True)
    
    chart.series.extend([s1, s2, s3])
    chart.set_categories(cats_ref)
    chart.width = 18
    chart.height = 10
    ws_summary.add_chart(chart, "M3")

    # ==========================================
    # Correlation Heatmap Sheet
    # ==========================================
    ws_corr = wb.create_sheet(title="Correlation Heatmap", index=1)
    returns_df = pd.DataFrame()
    for sym in main_assets:
        if sym in data_dict:
            returns_df[sym] = data_dict[sym]['Close'].pct_change()
    corr_matrix = returns_df.corr()
    
    ws_corr.cell(row=1, column=1, value="Asset").font = FONT_BOLD
    for c_idx, col_name in enumerate(corr_matrix.columns, 2):
        ws_corr.cell(row=1, column=c_idx, value=col_name).font = FONT_BOLD
        
    for r_idx, row_name in enumerate(corr_matrix.index, 2):
        ws_corr.cell(row=r_idx, column=1, value=row_name).font = FONT_BOLD
        for c_idx, col_name in enumerate(corr_matrix.columns, 2):
            val = corr_matrix.loc[row_name, col_name]
            c = ws_corr.cell(row=r_idx, column=c_idx, value=val)
            c.number_format = "0.00"
            c.border = BORDER_THIN
            
    rule = ColorScaleRule(start_type='min', start_color='F8696B', 
                          mid_type='num', mid_value=0, mid_color='FFEB84', 
                          end_type='max', end_color='63BE7B')
    ws_corr.conditional_formatting.add(f"B2:{get_column_letter(len(corr_matrix.columns)+1)}{len(corr_matrix.index)+1}", rule)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ==========================================
# 3. STREAMLIT APP UI & LOGIC
# ==========================================

st.set_page_config(page_title="Beta & Risk Analytics", layout="wide", initial_sidebar_state="expanded")

st.title("📈 Advanced Cost of Capital & Beta Dashboard")
st.markdown("Analyze market risk, compute unlevered betas, and handle custom competitor peer groups dynamically.")

# --- SIDEBAR CONFIGURATION ---
st.sidebar.header("Configuration")

mode = st.sidebar.radio(
    "Select Output Mode",
    ["Interactive Dashboard Only", "Excel Workbook with Charts Only", "Generate Both"]
)

start_date = st.sidebar.date_input("Start Date", datetime(2021, 4, 1))
end_date = st.sidebar.date_input("End Date", datetime(2026, 3, 31))
benchmark_input = st.sidebar.text_input("Benchmark Index", "NIFTY 50")

# Move optional competitor mapping to Sidebar
with st.sidebar.expander("⚙️ Optional Competitor Mapping (Advanced)", expanded=False):
    st.markdown("""
    *Specify custom peers per company to override default peer averaging.*  
    **Format:** `Target = Peer1, Peer2`  
    **Example:**  
    `TCS = INFY, WIPRO`  
    `RELIANCE = ONGC, BPCL`  
    *(If left blank, defaults to averaging other Target Companies).*
    """)
    custom_peers_str = st.text_area("Competitor Rules", height=120, placeholder="TCS = INFY, WIPRO")

# --- MAIN INPUTS ---
main_tickers_str = st.text_input("Target Companies (comma-separated)", "TCS, INFY, HCLTECH, RELIANCE")

if st.button("Run Financial Analysis", type="primary"):
    main_assets = [s.strip().upper() for s in main_tickers_str.split(",") if s.strip()]
    if not main_assets:
        st.error("Please enter at least one target company.")
        st.stop()

    comp_map = {}
    additional_peers = []
    if custom_peers_str.strip():
        for line in custom_peers_str.strip().split("\n"):
            if "=" in line:
                target, peers = line.split("=")
                target = target.strip().upper()
                peer_list = [p.strip().upper() for p in peers.split(",") if p.strip()]
                comp_map[target] = peer_list
                additional_peers.extend(peer_list)

    all_symbols_to_fetch = list(set(main_assets + additional_peers))
    bench_symbol = benchmark_input.strip().upper()
    if bench_symbol not in all_symbols_to_fetch:
        all_symbols_to_fetch.append(bench_symbol)

    with st.spinner("Downloading market data and computing metrics..."):
        fetched_data = {}
        financials_data = {}
        
        yf_start = start_date.strftime("%Y-%m-%d")
        yf_end = end_date.strftime("%Y-%m-%d")

        for sym in all_symbols_to_fetch:
            ticker = format_indian_symbol(sym)
            try:
                stock = yf.Ticker(ticker)
                df = stock.history(start=yf_start, end=yf_end)
                if not df.empty:
                    df.index = df.index.tz_localize(None)
                    fetched_data[sym] = df
                    financials_data[sym] = get_financial_metrics(ticker)
            except Exception as e:
                st.warning(f"Failed to fetch {sym}: {e}")

        if bench_symbol not in fetched_data:
            st.error(f"Could not fetch benchmark '{bench_symbol}'. Please verify the ticker.")
            st.stop()

        # STRICT ALIGNMENT: Ensure all fetched dataframes share exact trading dates
        common_dates = set(fetched_data[bench_symbol].index)
        for sym in fetched_data:
            common_dates = common_dates.intersection(set(fetched_data[sym].index))
        common_dates = sorted(list(common_dates))

        for sym in list(fetched_data.keys()):
            fetched_data[sym] = fetched_data[sym].loc[common_dates].copy()

    bench_df = fetched_data[bench_symbol]['Close']
    bench_ret = bench_df.pct_change().dropna()

    metrics_cache = {}
    for sym in all_symbols_to_fetch:
        if sym == bench_symbol: continue
        if sym not in fetched_data: continue
        target_ret = fetched_data[sym]['Close'].pct_change().dropna()
        
        aligned = pd.concat([target_ret, bench_ret], axis=1).dropna()
        aligned.columns = ['Target', 'Bench']
        
        mean_ret = aligned['Target'].mean()
        ann_ret = mean_ret * 252
        ann_vol = aligned['Target'].std(ddof=1) * np.sqrt(252)
        
        bench_var = aligned['Bench'].var(ddof=1)
        cov = aligned['Target'].cov(aligned['Bench'])
        
        beta_l = cov / bench_var if bench_var != 0 else 1.0
        de, tax = financials_data.get(sym, (0.0, 0.25))
        beta_u = beta_l / (1 + (1 - tax) * de)
        corr = aligned['Target'].corr(aligned['Bench'])
        
        metrics_cache[sym] = {
            "mean_ret": mean_ret, "ann_ret": ann_ret, "ann_vol": ann_vol,
            "beta_l": beta_l, "de": de, "tax": tax, "beta_u": beta_u, "corr": corr
        }

    dashboard_data = []
    for sym in main_assets:
        if sym not in metrics_cache: continue
        m = metrics_cache[sym]
        
        # Dynamic Custom Competitor Logic Evaluation
        peer_u_beta = 0
        if sym in comp_map and comp_map[sym]:
            peers = [p for p in comp_map[sym] if p in metrics_cache]
            if peers:
                peer_u_beta = np.mean([metrics_cache[p]["beta_u"] for p in peers])
            else: peer_u_beta = np.nan
        else:
            peers = [p for p in main_assets if p != sym and p in metrics_cache]
            if peers:
                peer_u_beta = np.mean([metrics_cache[p]["beta_u"] for p in peers])
            else: peer_u_beta = np.nan
            
        peer_rel_beta = peer_u_beta * (1 + (1 - m["tax"]) * m["de"]) if not np.isnan(peer_u_beta) else np.nan

        dashboard_data.append({
            "Asset": sym,
            "Ann Return": m["ann_ret"],
            "Ann Volatility": m["ann_vol"],
            "Raw Levered Beta": m["beta_l"],
            "D/E Ratio": m["de"],
            "Unlevered Beta": m["beta_u"],
            "Peer Avg Unlevered Beta": peer_u_beta,
            "Peer Relevered Beta": peer_rel_beta,
            "Correlation": m["corr"]
        })

    dash_df = pd.DataFrame(dashboard_data)
    st.session_state['dash_df'] = dash_df
    st.session_state['fetched_data'] = fetched_data
    st.session_state['financials_data'] = financials_data
    st.session_state['main_assets'] = main_assets
    st.session_state['comp_map'] = comp_map
    st.session_state['bench_symbol'] = bench_symbol
    st.session_state['yf_start'] = yf_start
    st.session_state['yf_end'] = yf_end
    st.session_state['ran_analysis'] = True

if 'ran_analysis' in st.session_state and st.session_state['ran_analysis']:
    dash_df = st.session_state['dash_df']
    fetched_data = st.session_state['fetched_data']
    financials_data = st.session_state['financials_data']
    main_assets = st.session_state['main_assets']
    comp_map = st.session_state['comp_map']
    bench_symbol = st.session_state['bench_symbol']
    yf_start = st.session_state['yf_start']
    yf_end = st.session_state['yf_end']

    show_dash = mode in ["Interactive Dashboard Only", "Generate Both"]
    show_excel = mode in ["Excel Workbook with Charts Only", "Generate Both"]

    if show_dash:
        st.divider()
        st.subheader("📊 Executive Summary Dashboard")
        
        formatted_df = dash_df.copy()
        format_dict = {
            "Ann Return": "{:.2%}", "Ann Volatility": "{:.2%}", "D/E Ratio": "{:.2%}",
            "Raw Levered Beta": "{:.2f}", "Unlevered Beta": "{:.2f}",
            "Peer Avg Unlevered Beta": "{:.2f}", "Peer Relevered Beta": "{:.2f}", "Correlation": "{:.3f}"
        }
        st.dataframe(formatted_df.style.format(format_dict).background_gradient(cmap='Blues', subset=['Raw Levered Beta', 'Peer Relevered Beta']), use_container_width=True)

        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Cost of Capital Beta Comparison")
            bar_df = dash_df.melt(id_vars='Asset', value_vars=['Raw Levered Beta', 'Unlevered Beta', 'Peer Relevered Beta'], var_name='Metric', value_name='Beta')
            fig1 = px.bar(bar_df, x='Asset', y='Beta', color='Metric', barmode='group', text_auto='.2f',
                          color_discrete_sequence=['#1B365D', '#6CA0DC', '#E67E22'])
            fig1.update_layout(xaxis_title="", yaxis_title="Beta Value", legend_title="")
            st.plotly_chart(fig1, use_container_width=True)

        with c2:
            st.subheader("Asset Correlation Heatmap")
            returns_df = pd.DataFrame()
            for sym in main_assets:
                if sym in fetched_data:
                    returns_df[sym] = fetched_data[sym]['Close'].pct_change()
            corr_matrix = returns_df.corr()
            
            fig2 = px.imshow(corr_matrix, text_auto='.2f', color_continuous_scale='RdYlGn', aspect="auto")
            st.plotly_chart(fig2, use_container_width=True)

        st.divider()
        st.subheader(f"📈 Target Companies Regression Scatter Plots (vs {bench_symbol})")
        scatter_cols = st.columns(2)
        idx = 0
        
        for sym in main_assets:
            if sym in fetched_data and bench_symbol in fetched_data:
                target_ret = fetched_data[sym]['Close'].pct_change().dropna()
                b_ret = fetched_data[bench_symbol]['Close'].pct_change().dropna()
                
                aligned = pd.concat([target_ret, b_ret], axis=1).dropna()
                aligned.columns = [sym, bench_symbol]
                
                m, c_y = np.polyfit(aligned[bench_symbol], aligned[sym], 1)
                
                fig_scat = px.scatter(aligned, x=bench_symbol, y=sym, opacity=0.6, title=f"{sym} Regression")
                fig_scat.add_trace(go.Scatter(
                    x=aligned[bench_symbol], y=m*aligned[bench_symbol] + c_y, 
                    mode='lines', name='OLS Trendline', line=dict(color='red')
                ))
                fig_scat.update_layout(xaxis_title=f"{bench_symbol} Daily Return", yaxis_title=f"{sym} Daily Return")
                scatter_cols[idx % 2].plotly_chart(fig_scat, use_container_width=True)
                idx += 1

    if show_excel:
        st.divider()
        st.subheader("📥 Excel Export with Native Charts")
        excel_bytes = generate_market_excel(
            fetched_data, financials_data, main_assets, comp_map, bench_symbol,
            yf_start, yf_end, dash_df
        )
        st.download_button(
            label="Download Complete Analytical Workbook (.xlsx)",
            data=excel_bytes,
            file_name=f"Financial_Model_{yf_start}_to_{yf_end}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
