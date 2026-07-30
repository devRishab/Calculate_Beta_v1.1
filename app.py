import streamlit as st
import yfinance as yf
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# ==========================================
# 1. PAGE CONFIGURATION & STYLING
# ==========================================
st.set_page_config(
    page_title="Beta & Valuation Dashboard",
    page_icon="📈",
    layout="wide"
)

st.title("📈 Beta Calculation & Valuation Dashboard")
st.markdown("Accurate Financial Beta, Unlevered Beta, Peer Relevered Beta, and Regression Analysis.")

# ==========================================
# 2. SIDEBAR - INPUTS & CONFIGURATION
# ==========================================
st.sidebar.header("1. Portfolio & Date Settings")

# Date range selection (Defaulted to requested dates)
start_date = st.sidebar.date_input("Start Date", datetime.strptime("2021-04-01", "%Y-%m-%d"))
end_date = st.sidebar.date_input("End Date", datetime.strptime("2026-03-31", "%Y-%m-%d"))

st.sidebar.subheader("Tickers")
target_ticker = st.sidebar.text_input("Target Company Ticker", value="TCS.NS").strip().upper()
peer_tickers_input = st.sidebar.text_input("Peer Tickers (comma-separated)", value="INFY.NS, WIT, TECHM.NS, HCLTECH.NS")
benchmark_ticker = st.sidebar.text_input("Benchmark Market Index", value="^NSEI").strip().upper()

# Parse peer list
peer_list = [p.strip().upper() for p in peer_tickers_input.split(",") if p.strip()]
all_tickers = list(dict.fromkeys([target_ticker] + peer_list + [benchmark_ticker]))

st.sidebar.markdown("---")
st.sidebar.header("2. Capital Structure Inputs")

# Interactive financial inputs for target and peers
financial_data = {}

st.sidebar.subheader(f"Target: {target_ticker}")
t_debt = st.sidebar.number_input(f"{target_ticker} Total Debt", value=1000.0, step=100.0)
t_equity = st.sidebar.number_input(f"{target_ticker} Market Equity", value=10000.0, step=500.0)
t_tax = st.sidebar.number_input(f"{target_ticker} Tax Rate (%)", value=25.0, step=1.0) / 100.0

financial_data[target_ticker] = {"Debt": t_debt, "Equity": t_equity, "Tax_Rate": t_tax}

with st.sidebar.expander("Peer Financial Inputs", expanded=False):
    for peer in peer_list:
        st.markdown(f"**{peer}**")
        p_debt = st.number_input(f"{peer} Debt", value=500.0, step=50.0, key=f"d_{peer}")
        p_equity = st.number_input(f"{peer} Equity", value=5000.0, step=250.0, key=f"e_{peer}")
        p_tax = st.number_input(f"{peer} Tax Rate (%)", value=25.0, step=1.0, key=f"t_{peer}") / 100.0
        financial_data[peer] = {"Debt": p_debt, "Equity": p_equity, "Tax_Rate": p_tax}

# ==========================================
# 3. DATA FETCHING & COMPUTATION ENGINE
# ==========================================
@st.cache_data(ttl=3600)
def fetch_financial_data(tickers, start, end):
    data = yf.download(tickers, start=start, end=end)['Adj Close']
    if isinstance(data, pd.Series):
        data = data.to_frame()
    return data

try:
    with st.spinner("Fetching market data..."):
        price_df = fetch_financial_data(all_tickers, start_date, end_date)

    # Drop missing values across all tickers for date alignment
    price_df = price_df.dropna()
    
    # Calculate simple daily percentage returns (Matches Excel PERCENTCHANGE / SLOPE)
    returns_df = price_df.pct_change().dropna()

    if benchmark_ticker not in returns_df.columns:
        st.error(f"Benchmark ticker '{benchmark_ticker}' not found in fetched data.")
        st.stop()

    benchmark_returns = returns_df[benchmark_ticker]

    # Calculate Raw Levered Betas using OLS Covariance / Variance ratio (Excel SLOPE equivalent)
    results = []
    market_var = np.var(benchmark_returns, ddof=1)

    for ticker in [target_ticker] + peer_list:
        if ticker in returns_df.columns:
            stock_returns = returns_df[ticker]
            
            # Covariance stock vs market
            cov_matrix = np.cov(stock_returns, benchmark_returns)
            raw_beta = cov_matrix[0, 1] / market_var
            
            # Correlation with benchmark (Excel CORREL equivalent)
            correlation = np.corrcoef(stock_returns, benchmark_returns)[0, 1]
            
            # R-Squared
            r_squared = correlation ** 2
            
            # Capital Structure calculations
            debt = financial_data[ticker]["Debt"]
            equity = financial_data[ticker]["Equity"]
            tax_rate = financial_data[ticker]["Tax_Rate"]
            de_ratio = debt / equity if equity > 0 else 0
            
            # Unlevered Beta (Hamada equation)
            unlevered_beta = raw_beta / (1 + (1 - tax_rate) * de_ratio)

            results.append({
                "Ticker": ticker,
                "Type": "Target" if ticker == target_ticker else "Peer",
                "Market Cap / Equity": equity,
                "Total Debt": debt,
                "D/E Ratio": de_ratio,
                "Tax Rate (%)": tax_rate * 100,
                "Raw Levered Beta": raw_beta,
                "Unlevered Beta": unlevered_beta,
                "Correlation": correlation,
                "R-Squared": r_squared
            })

    results_df = pd.DataFrame(results)

    # Peer Average & Median Unlevered Beta
    peer_unlevered_betas = results_df[results_df["Type"] == "Peer"]["Unlevered Beta"]
    avg_peer_unlevered_beta = peer_unlevered_betas.mean() if not peer_unlevered_betas.empty else results_df[results_df["Ticker"] == target_ticker]["Unlevered Beta"].values[0]

    # Target Relevered Beta (using Peer Average Unlevered Beta)
    target_de = financial_data[target_ticker]["Debt"] / financial_data[target_ticker]["Equity"]
    target_tax = financial_data[target_ticker]["Tax_Rate"]
    target_relevered_beta = avg_peer_unlevered_beta * (1 + (1 - target_tax) * target_de)

    # ==========================================
    # 4. DASHBOARD DISPLAY
    # ==========================================
    
    # Key Summary Metrics
    target_row = results_df[results_df["Ticker"] == target_ticker].iloc[0]
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Raw Levered Beta", f"{target_row['Raw Levered Beta']:.3f}")
    col2.metric("Target Unlevered Beta", f"{target_row['Unlevered Beta']:.3f}")
    col3.metric("Peer Avg Unlevered Beta", f"{avg_peer_unlevered_beta:.3f}")
    col4.metric("Target Relevered Beta", f"{target_relevered_beta:.3f}", delta=f"{target_relevered_beta - target_row['Raw Levered Beta']:.3f} vs Raw")

    st.markdown("---")

    # Layout: Visualizations
    tab1, tab2, tab3, tab4 = st.tabs(["📉 Regression Scatter Plot", "📊 Beta Comparison Bar Chart", "🔥 Correlation Heatmap", "📋 Detailed Data Table"])

    # -------------------------------------------------------------
    # TAB 1: Target Regression Scatter Plot (OLS Line)
    # -------------------------------------------------------------
    with tab1:
        st.subheader(f"Target ({target_ticker}) vs Benchmark ({benchmark_ticker}) Regression Scatter Plot")
        
        target_ret = returns_df[target_ticker]
        bench_ret = returns_df[benchmark_ticker]

        # Calculate linear regression line parameters (y = mx + c)
        slope, intercept = np.polyfit(bench_ret, target_ret, 1)
        line_x = np.linspace(bench_ret.min(), bench_ret.max(), 100)
        line_y = slope * line_x + intercept

        fig_scatter = go.Figure()

        # Scatter points
        fig_scatter.add_trace(go.Scatter(
            x=bench_ret,
            y=target_ret,
            mode='markers',
            name='Daily Returns',
            marker=dict(color='#1f77b4', opacity=0.6, size=5)
        ))

        # Regression Line
        fig_scatter.add_trace(go.Scatter(
            x=line_x,
            y=line_y,
            mode='lines',
            name=f'OLS Trendline (Beta = {slope:.3f})',
            line=dict(color='#ff7f0e', width=2)
        ))

        fig_scatter.update_layout(
            title=f"Linear Regression: {target_ticker} vs {benchmark_ticker}<br><sup>Equation: Returns({target_ticker}) = {slope:.3f} × Returns({benchmark_ticker}) + ({intercept:.5f}) | R² = {target_row['R-Squared']:.3f}</sup>",
            xaxis_title=f"Benchmark Return ({benchmark_ticker})",
            yaxis_title=f"Target Return ({target_ticker})",
            template="plotly_white",
            height=500
        )

        st.plotly_chart(fig_scatter, use_container_width=True)

    # -------------------------------------------------------------
    # TAB 2: Bar Chart - Raw, Unlevered, & Relevered Betas
    # -------------------------------------------------------------
    with tab2:
        st.subheader("Beta Comparison Breakdown")
        
        # Prepare data for Beta comparison
        target_raw = target_row['Raw Levered Beta']
        target_unlevered = target_row['Unlevered Beta']
        
        beta_comp_df = pd.DataFrame({
            "Beta Metric": ["Raw Levered Beta", "Unlevered Beta", "Peer Relevered Beta"],
            "Value": [target_raw, target_unlevered, target_relevered_beta]
        })

        fig_bar = px.bar(
            beta_comp_df,
            x="Beta Metric",
            y="Value",
            text_auto='.3f',
            color="Beta Metric",
            color_discrete_sequence=['#2b5c8f', '#4682b4', '#d9534f'],
            title=f"Target ({target_ticker}) Beta Transition Analysis"
        )

        fig_bar.update_layout(
            yaxis_title="Beta Coefficient",
            showlegend=False,
            template="plotly_white",
            height=450
        )

        st.plotly_chart(fig_bar, use_container_width=True)

    # -------------------------------------------------------------
    # TAB 3: Correlation Heatmap
    # -------------------------------------------------------------
    with tab3:
        st.subheader("Returns Correlation Heatmap")
        
        # Filter return columns for target, peers, and benchmark
        corr_tickers = [t for t in all_tickers if t in returns_df.columns]
        corr_matrix = returns_df[corr_tickers].corr()

        fig_heatmap = px.imshow(
            corr_matrix,
            text_auto='.2f',
            color_continuous_scale='Blues',
            title="Correlation Matrix across Target, Peers & Benchmark",
            aspect="auto"
        )

        fig_heatmap.update_layout(height=500)
        st.plotly_chart(fig_heatmap, use_container_width=True)

    # -------------------------------------------------------------
    # TAB 4: Detailed Summary Data Table
    # -------------------------------------------------------------
    with tab4:
        st.subheader("Summary Table")
        
        display_df = results_df.copy()
        display_df.loc[len(display_df)] = {
            "Ticker": f"RELEVERED ({target_ticker})",
            "Type": "Target Relevered",
            "Market Cap / Equity": target_row["Market Cap / Equity"],
            "Total Debt": target_row["Total Debt"],
            "D/E Ratio": target_de,
            "Tax Rate (%)": target_tax * 100,
            "Raw Levered Beta": np.nan,
            "Unlevered Beta": avg_peer_unlevered_beta,
            "Correlation": np.nan,
            "R-Squared": np.nan
        }
        
        st.dataframe(
            display_df.style.format({
                "Market Cap / Equity": "{:,.1f}",
                "Total Debt": "{:,.1f}",
                "D/E Ratio": "{:.2%}",
                "Tax Rate (%)": "{:.1f}%",
                "Raw Levered Beta": "{:.3f}",
                "Unlevered Beta": "{:.3f}",
                "Correlation": "{:.3f}",
                "R-Squared": "{:.3f}"
            }, na_rep="-"),
            use_container_width=True
        )

except Exception as e:
    st.error(f"An error occurred while processing data: {str(e)}")
