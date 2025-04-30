import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from prophet import Prophet
import base64
from io import StringIO, BytesIO

# Splash screen logic
if "splash_shown" not in st.session_state:
    st.session_state.splash_shown = False

if not st.session_state.splash_shown:
    st.markdown(
        "<h1 style='text-align: center; font-weight: bold;'>👋 Welcome to Enterprise Sales Analytics!</h1>",
        unsafe_allow_html=True
    )
    st.markdown(
        """
        <div style='text-align: center; font-size: 1.2em; margin-bottom: 30px;'>
            Analyze your sales data, visualize trends, and forecast the future.<br>
            <b>Upload your file and get started!</b>
        </div>
        """,
        unsafe_allow_html=True
    )
    col1, col2, col3 = st.columns([3, 2, 3])
    with col2:
        if st.button("Continue to Dashboard"):
            st.session_state.splash_shown = True
    st.stop()

# Utility functions
def generate_forecast(df, period=365):
    df_prophet = df.rename(columns={'Order Date': 'ds', 'Sales': 'y'})
    model = Prophet(interval_width=0.95)
    model.fit(df_prophet)
    future = model.make_future_dataframe(periods=period)
    forecast = model.predict(future)
    return model, forecast

# Page config
st.set_page_config(
    page_title="Enterprise Sales Analytics",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# Custom CSS for download buttons (optional)
st.markdown("""
<style>
.download-btn {
    background: #4f8bf9;
    color: white !important;
    padding: 10px 20px;
    border-radius: 5px;
    text-decoration: none;
    display: inline-block;
    margin: 10px 0;
    transition: all 0.3s ease;
}
.download-btn:hover {
    background: #3a6db0;
    box-shadow: 0 2px 5px rgba(0,0,0,0.2);
}
</style>
""", unsafe_allow_html=True)

st.title("📊 Enterprise Sales Analytics Platform")
st.markdown("---")

uploaded_file = st.file_uploader(
    "📤 Upload Your Sales Data (Excel/CSV)",
    type=["xlsx", "csv"],
    help="File must include columns: Order Date, Sales, Profit"
)

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, parse_dates=['Order Date'])
        else:
            df = pd.read_excel(uploaded_file, parse_dates=['Order Date'])

        st.sidebar.header("🔍 Filters")

        date_col = 'Order Date' if 'Order Date' in df.columns else df.columns[0]
        min_date = df[date_col].min().date()
        max_date = df[date_col].max().date()

        date_range = st.sidebar.date_input(
            "📅 Date Range",
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )

        sales_range = st.sidebar.slider(
            "💰 Sales Range ($)",
            float(df['Sales'].min()),
            float(df['Sales'].max()),
            (float(df['Sales'].min()), float(df['Sales'].max()))
        )

        category_filter = st.sidebar.multiselect(
            "🗂️ Categories",
            options=df['Category'].unique() if 'Category' in df.columns else [],
            default=df['Category'].unique() if 'Category' in df.columns else []
        )

        filtered_df = df[
            (df[date_col].dt.date >= date_range[0]) &
            (df[date_col].dt.date <= date_range[1]) &
            (df['Sales'].between(sales_range[0], sales_range[1]))
        ]

        if category_filter and 'Category' in df.columns:
            filtered_df = filtered_df[filtered_df['Category'].isin(category_filter)]

        # <-- Added expander showing full uploaded data -->
        with st.expander("📁 Show full uploaded data"):
            st.dataframe(df, use_container_width=True)

        st.header("📊 Real-Time Sales Dashboard")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Sales", f"${filtered_df['Sales'].sum():,.0f}")
        col2.metric("Total Profit", f"${filtered_df['Profit'].sum():,.0f}")
        col3.metric("Avg. Profit Margin", f"{(filtered_df['Profit'].sum() / filtered_df['Sales'].sum())*100:.1f}%")

        # Visualization example
        st.subheader("📈 Sales by Category")
        if 'Category' in filtered_df.columns:
            fig = px.bar(filtered_df.groupby('Category')['Sales'].sum().reset_index(),
                         x='Category', y='Sales', title='Sales by Category')
            st.plotly_chart(fig, use_container_width=True)

        # Forecasting example
        st.subheader("🔮 Sales Forecasting")
        try:
            model, forecast = generate_forecast(filtered_df[['Order Date', 'Sales']], period=90)
            fig_forecast = go.Figure()
            fig_forecast.add_trace(go.Scatter(x=forecast['ds'], y=forecast['yhat'], name='Forecast'))
            st.plotly_chart(fig_forecast, use_container_width=True)
        except Exception as e:
            st.warning(f"Forecasting failed: {e}")

    except Exception as e:
        st.error(f"Failed to process file: {e}")
