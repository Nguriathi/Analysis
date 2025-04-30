# 📈 Enterprise Sales Analytics Platform

    ```
        https://salesanalytic.streamlit.app/

    ```
A powerful sales analytics dashboard built with [Streamlit](https://streamlit.io/) that provides real-time insights, trend visualization, and sales forecasting capabilities.

---

## 🌟 Features

- ✅ **Interactive Dashboard**
- ✅ **Sales Forecasting** (Facebook Prophet)
- ✅ **Real-Time Filtering**
- ✅ **Data Export Capabilities**
- ✅ **Custom Visualization Themes**
- ✅ **Splash Screen**
- ✅ **Responsive Design**

---

## 🚀 Getting Started

### Prerequisites

Make sure you have Python 3.8+ installed.

Install dependencies:

pip install -r requirements.txt



### Installation

1. **Clone the repository**
    ```
    git clone https://github.com/Nguriathi/Analysis.git
    
    cd enterprise-sales-analytics
    ```

2. **Install dependencies**
    ```
    pip install streamlit pandas plotly prophet openpyxl
    ```

3. **Run the application**
    ```
    streamlit run app.py
    ```



## 📂 File Structure

enterprise-sales-analytics/
├── app.py # Main application code
├── requirements.txt # Dependency list
├── README.md # Documentation
├── config.yaml # (Optional) Authentication config
└── Data.xlsx # Data used as an example for the project


---

## 🧮 Core Functionality

def generate_forecast(df, period=365):
"""Generate sales forecast using Facebook Prophet"""
model = Prophet()
model.fit(df)
future = model.make_future_dataframe(periods=period)
return model.predict(future)


---

## 📊 Supported Data Formats

| Format | Features      |
|--------|--------------|
| CSV    | Full support |
| Excel  | Full support |
| JSON   | Not supported |

---

## 🛠️ Configuration

To enable authentication, edit `config.yaml` file:

credentials:
usernames:
admin:
email: admin@company.com
name: Admin User
password: "hashed_password"
cookie:
name: "sales_analytics"
key: "your_secret_key"
expiry_days: 1


---

## 💡 Usage Tips

1. **Data Requirements**
    - Must contain `Order Date`, `Sales`, and `Profit` columns
    - Date format: `YYYY-MM-DD`
    - Numeric columns should be clean

2. **Performance**
    - Optimal dataset size: <100,000 rows
    - For large datasets, enable caching with `@st.cache`

---

## 🧪 Testing

Run basic tests with:

python -m pytest tests/


---

## 🤝 Contributing

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## 📜 License

Distributed under the MIT License.

---

## 📞 Contact

**Project Maintainer:** Nguriathi

Project Link: https://salesanalytic.streamlit.app/

---

## 🏆 Deployment Options

### 1. Streamlit Community Cloud

[![Deploy to Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://streamlit.io/cloud)

### 2. Linux Server

sudo apt install python3-pip
pip install -r requirements.txt
streamlit run app.py --server.port 80



### 3. Docker

FROM python:3.9-slim
COPY . /app
WORKDIR /app
RUN pip install -r requirements.txt
EXPOSE 8501
CMD ["streamlit", "run", "app.py"]
