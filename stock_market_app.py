"""
Stock Market Analysis System - Enhanced UI
Built with Streamlit | yfinance | Pandas | Plotly
Team Project — Dr. Khalaf
"""

import streamlit as st
import yfinance as yf
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Pro Stock Analyzer",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CUSTOM CSS (The "Masterpiece" UI)
# ─────────────────────────────────────────────
st.markdown("""
<style>
    /* Main Background & Font */
    .stApp {
        background-color: #fcfcfd;
    }
    
    /* Elegant Header */
    .main-header {
        background: linear-gradient(90deg, #0f2027, #203a43, #2c5364);
        padding: 25px;
        border-radius: 12px;
        color: white;
        font-family: 'Inter', sans-serif;
        font-weight: 800;
        font-size: 2.2rem;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }

    /* Professional Info Boxes */
    .info-card {
        background: white;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #eef0f2;
        box-shadow: 2px 2px 10px rgba(0,0,0,0.02);
        margin-bottom: 10px;
    }
    .info-label {
        color: #64748b;
        font-size: 0.8rem;
        text-transform: uppercase;
        font-weight: 600;
        letter-spacing: 0.5px;
    }
    .info-value {
        color: #1e293b;
        font-size: 1.1rem;
        font-weight: 700;
    }

    /* Sidebar Styling */
    section[data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 1px solid #e2e8f0;
    }
    
    /* Customizing Buttons */
    div.stButton > button {
        border-radius: 8px;
        background-color: white;
        border: 1px solid #e2e8f0;
        color: #475569;
        font-weight: 500;
        transition: all 0.3s ease;
    }
    div.stButton > button:hover {
        border-color: #3b82f6;
        color: #3b82f6;
        background-color: #eff6ff;
    }
    
    /* Metrics Styling */
    [data-testid="stMetricValue"] {
        font-weight: 800;
        color: #0f172a;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown('<div class="main-header">📈 Market Intelligence Suite</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🛠️ Configuration")
    
    stock_symbol = st.text_input(
        "Search Ticker",
        value="AAPL",
        help="Type symbol (e.g., TSLA, NVDA)"
    ).upper().strip()

    period_options = {
        "7 Days": "7d", "1 Month": "1mo", "3 Months": "3mo", 
        "6 Months": "6mo", "1 Year": "1y", "5 Years": "5y"
    }
    selected_period_label = st.selectbox("Analysis Horizon", list(period_options.keys()), index=1)
    selected_period = period_options[selected_period_label]

    chart_type = st.radio("Visual Style", ["Line Chart", "Candlestick", "Area Chart"], horizontal=True)

    show_ma = st.toggle("Moving Average Overlay (20d)", value=True)

    st.markdown("---")
    st.markdown("### ⚡ Quick Access")
    popular = ["AAPL", "NVDA", "MSFT", "TSLA", "BTC-USD"]
    
    # Grid of buttons
    cols = st.columns(2)
    for i, sym in enumerate(popular):
        if cols[i % 2].button(sym, use_container_width=True):
            stock_symbol = sym

# ─────────────────────────────────────────────
# BACKEND LOGIC (Member 1 & 2)
# ─────────────────────────────────────────────
@st.cache_data(ttl=300)
def fetch_stock_data(symbol: str, period: str):
    ticker = yf.Ticker(symbol)
    hist = ticker.history(period=period)
    info = ticker.info
    return hist, info

def validate_symbol(symbol: str) -> bool:
    return bool(symbol) and len(symbol) <= 10 # Allow crypto and long tickers

def process_data(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    df = df.copy()
    df["MA20"] = df["Close"].rolling(window=20).mean()
    return df

# ─────────────────────────────────────────────
# VISUALIZATION (Member 3)
# ─────────────────────────────────────────────
def get_chart(df, symbol, type, show_ma):
    fig = go.Figure()
    
    if type == "Line Chart":
        fig.add_trace(go.Scatter(x=df.index, y=df["Close"], name="Price", line=dict(color="#3b82f6", width=3)))
    elif type == "Area Chart":
        fig.add_trace(go.Scatter(x=df.index, y=df["Close"], fill='tozeroy', name="Price", line=dict(color="#3b82f6")))
    else:
        fig.add_trace(go.Candlestick(x=df.index, open=df['Open'], high=df['High'], low=df['Low'], close=df['Close'], name="Market"))

    if show_ma and "MA20" in df.columns:
        fig.add_trace(go.Scatter(x=df.index, y=df["MA20"], name="MA(20)", line=dict(color="#f59e0b", width=1.5, dash='dot')))

    fig.update_layout(
        template="plotly_white",
        hovermode="x unified",
        margin=dict(l=0, r=0, t=30, b=0),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    return fig

# ─────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────
if not validate_symbol(stock_symbol):
    st.warning("Please enter a valid ticker to begin.")
    st.stop()

with st.spinner("Analyzing Market Data..."):
    try:
        hist_df, info = fetch_stock_data(stock_symbol, selected_period)
        if hist_df.empty:
            st.error("No data found for this symbol.")
            st.stop()
        hist_df = process_data(hist_df)
    except:
        st.error("Unable to connect to market servers.")
        st.stop()

# Company Profile Header
c1, c2 = st.columns([3, 1])
with c1:
    st.title(f"{info.get('longName', stock_symbol)}")
with c2:
    st.markdown(f"""
    <div class="info-card">
        <div class="info-label">Exchange</div>
        <div class="info-value">{info.get('exchange', 'N/A')}</div>
    </div>
    """, unsafe_allow_html=True)

# Sector & Industry Cards
col_s1, col_s2, col_s3 = st.columns(3)
with col_s1:
    st.markdown(f'<div class="info-card"><div class="info-label">Sector</div><div class="info-value">{info.get("sector", "N/A")}</div></div>', unsafe_allow_html=True)
with col_s2:
    st.markdown(f'<div class="info-card"><div class="info-label">Industry</div><div class="info-value">{info.get("industry", "N/A")}</div></div>', unsafe_allow_html=True)
with col_s3:
    st.markdown(f'<div class="info-card"><div class="info-label">Currency</div><div class="info-value">{info.get("currency", "USD")}</div></div>', unsafe_allow_html=True)

# Key Performance Metrics
curr = hist_df["Close"].iloc[-1]
prev = hist_df["Close"].iloc[-2] if len(hist_df) > 1 else curr
change = ((curr - prev) / prev) * 100

st.markdown("---")
m_col1, m_col2, m_col3, m_col4 = st.columns(4)
m_col1.metric("Current Price", f"${curr:,.2f}", f"{change:+.2f}%")
m_col2.metric("Period High", f"${hist_df['High'].max():,.2f}")
m_col3.metric("Period Low", f"${hist_df['Low'].min():,.2f}")
m_col4.metric("Avg Volume", f"{hist_df['Volume'].mean():,.0f}")

# Main Visualization
st.plotly_chart(get_chart(hist_df, stock_symbol, chart_type, show_ma), use_container_width=True)

# Detailed Data Table
with st.expander("📂 View Raw Historical Data"):
    st.dataframe(hist_df.sort_index(ascending=False), use_container_width=True)
    
st.markdown("---")
st.caption(f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Data provided by Yahoo Finance")