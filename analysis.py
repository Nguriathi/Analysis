from logging import PlaceHolder
import streamlit as st  
import pandas as pd  
import plotly.express as px  
import base64  
from io import StringIO, BytesIO 
from streamlit_lottie import st_lottie
import json


def generate_excel_download_link(df):
    towrite = BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_download.xlsx">Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_html_download_link(fig):
    towrite = StringIO()
    fig.write_html(towrite, include_plotlyjs="cdn")
    towrite = BytesIO(towrite.getvalue().encode())
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:text/html;charset=utf-8;base64, {b64}" download="plot.html">Plot</a>'
    return st.markdown(href, unsafe_allow_html=True)

st.set_page_config(page_title='Analysis', layout="wide")

st.markdown("<h1 style='text-align: center; color: white;'>PY SALES ANALYSIS</h1>", unsafe_allow_html=True)


def load_lottiefile(filepath: str):
        with open(filepath, "r") as f:
            return json.load(f)

lottie_news = load_lottiefile("analysis.json")
st_lottie(
        lottie_news,
        speed=0.5,
        reverse=False,
        loop=True,
        height=350,
        width=None,
        key=None,
     )

uploaded_file = st.file_uploader('Choose a XLSX file', type='xlsx')

st.sidebar.markdown("<h1 style='text-align: center; color: white;'>ANALYSIS OPTIONS</h1>", unsafe_allow_html=True)
groupby_column = st.sidebar.selectbox( 'Select Option',( 'Ship Mode', 'Segment', 'City', 'Category', 'Sub-Category'))

if uploaded_file:
    
    st.markdown('---')
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    st.dataframe(df)

        
    # GROUP DATAFRAME
    output_columns = ['Sales', 'Profit']
    df_grouped = df.groupby(by=[groupby_column], as_index=False)[output_columns].sum()

    # PLOT DATAFRAME
    fig = px.bar(
        df_grouped,
        x=groupby_column,
        y='Sales',
        color='Profit',
        color_continuous_scale=['red', 'green', 'blue'],
        template='plotly_white',
        title=f'<b>Sales & Profit by {groupby_column}</b>'
    )
    st.plotly_chart(fig)

    # DOWNLOAD FILES
    st.subheader('Download')
    generate_excel_download_link(df_grouped)
    generate_html_download_link(fig)


hide_st_style = """
            <style>
            #mainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            
            """    
st.markdown(hide_st_style, unsafe_allow_html=True)        
