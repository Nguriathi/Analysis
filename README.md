📦 Product Invoice Analyzer
[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svgrn Streamlit app for visualizing, editing, and forecasting invoice product data.
Upload your Excel invoice(s) and get instant analytics, summaries, and export-ready files!

🌟 Features
✅ Interactive Data Editing

✅ Adaptive Visualizations (bar, treemap, line charts)

✅ Batch & Single File Processing

✅ Automatic Vessel & Agent Extraction

✅ Forecasting with Prophet

✅ Bullet Point Summaries

✅ Export to Excel Template

✅ Modern UI with Splash Screen and Lottie Animation

🚀 Getting Started
Prerequisites
Python 3.8 or higher

Installation
Clone the repository

bash
git clone https://github.com/your-username/invoice-analyzer.git
cd invoice-analyzer
Install dependencies

bash
pip install -r requirements.txt
or, if you don’t have a requirements file:

bash
pip install streamlit streamlit-lottie pandas plotly openpyxl prophet
Add your files

Place your template.xlsx and assets/animation.json in the project root.

Run the app

bash
streamlit run app.py
📂 File Structure
text
invoice-analyzer/
├── app.py               # Main Streamlit application
├── requirements.txt     # Dependency list
├── README.md            # This file
├── assets/
│   └── animation.json   # Lottie animation for sidebar
├── template.xlsx        # Excel template for export
└── Data.xlsx            # Example invoice data
🧮 Core Functionality
Extracts product tables from your invoice Excel file (analysis sheet).

Displays vessel and agent (if present) above the table in single file mode.

Interactive editing of product data.

Visualizes:

Product quantities (bar/treemap)

Trends and recurring items (line/bar)

Price and quantity outliers (scatter)

Forecasts future quantities using Facebook Prophet.

Exports:

Edited data as Excel

Single file mode: fills your template.xlsx with all info

Sample function for forecasting:

python
def prophet_forecast(df, periods=3):
    ts = df.groupby('INVOICE_DATE')['QTY'].sum().reset_index()
    ts = ts.rename(columns={'INVOICE_DATE': 'ds', 'QTY': 'y'})
    if len(ts) < 2:
        return None
    model = Prophet(yearly_seasonality=False, weekly_seasonality=False, daily_seasonality=False)
    model.fit(ts)
    future = model.make_future_dataframe(periods=periods, freq='MS')
    forecast = model.predict(future)
    return forecast
📊 Supported Data Formats
Format	Features
Excel	Full support (analysis sheet as shown below)
CSV	Not supported (convert to Excel first)
JSON	Not supported
📝 Invoice Format Example
Your Excel file should have an analysis sheet like:

AGENT	...	...
VESSEL	...	...
DOD	...	...
...	...	...
NO	PRODUCT DESCRIPTION	UNIT/PRC
1	Cabbage White	1.2
...	...	...
💡 Usage Tips
Data Requirements

analysis sheet must include columns: NO, PRODUCT DESCRIPTION, UNIT/PRC, UNIT, QTY, TOTAL USD

AGENT and VESSEL fields are auto-extracted if present

Date format: YYYY-MM-DD for DOD

Performance

For large datasets, enable Streamlit caching (@st.cache_data)

🧪 Testing
Test your data extraction and forecasting logic with:

python
python -m pytest tests/
🤝 Contributing
Fork the repository

Create your feature branch (git checkout -b feature/AmazingFeature)

Commit your changes (git commit -m 'Add some AmazingFeature')

Push to the branch (git push origin feature/AmazingFeature)

Open a Pull Request

📜 License
Distributed under the MIT License.

📞 Contact
Project Maintainer: Your Name
Live App: https://chandler.streamlit.app/

🏆 Deployment Options
1. Streamlit Community Cloud
[![Deploy to Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svgx Server

bash
sudo apt install python3-pip
pip install -r requirements.txt
streamlit run app.py --server.port 80
3. Docker
text
FROM python:3.9-slim
COPY . /app
WORKDIR /app
RUN pip install -r requirements.txt
EXPOSE 8501
CMD ["streamlit", "run", "app.py"]
Built with ❤️ using Streamlit for actionable invoice analytics and business intelligence.
