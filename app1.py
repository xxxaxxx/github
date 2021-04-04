#0# open Anaconda Powershell Prompt (anaconda3) 
#1# cd desktop
#2# cd ai-Excel-Webapp
#3# pip install plotly streamlit pandas openpyxl Pillow
#4# streamlit run app.py


import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image

st.set_page_config(page_title='Strategy Results')
st.header('Strategy Results 2021')
st.subheader('Strategy')

### --- LOAD DATAFRAME
excel_file = 'Survey_Results.xlsx'
sheet_name = 'DATA1'

df = pd.read_excel(excel_file,
                   sheet_name=sheet_name,
                   usecols='B:D',
                   header=3)

df_participants = pd.read_excel(excel_file,
                                sheet_name= sheet_name,
                                usecols='F:G',
                                header=3)
df_participants.dropna(inplace=True)

# --- STREAMLIT SELECTION
department = df['Strategy'].unique().tolist()
ages = df['Total Order'].unique().tolist()

age_selection = st.slider('Total Order:',
                        min_value= min(ages),
                        max_value= max(ages),
                        value=(min(ages),max(ages)))

department_selection = st.multiselect('Strategy:',
                                    department,
                                    default=department)

# --- FILTER DATAFRAME BASED ON SELECTION
mask = (df['Total Order'].between(*age_selection)) & (df['Strategy'].isin(department_selection))
number_of_result = df[mask].shape[0]
st.markdown(f'*Available Results: {number_of_result}*')

# --- GROUP DATAFRAME AFTER SELECTION
df_grouped = df[mask].groupby(by=['Total Order']).count()[['Long']]
df_grouped = df_grouped.rename(columns={'Long': 'Long'})
df_grouped = df_grouped.reset_index()

# --- PLOT BAR CHART
bar_chart = px.bar(df_grouped,
                   x='Total Order',
                   y='Long',
                   text='Long',
                   color_discrete_sequence = ['#F63366']*len(df_grouped),
                   template= 'plotly_white')
st.plotly_chart(bar_chart)

# --- DISPLAY IMAGE & DATAFRAME
col1, col2 = st.beta_columns(2)
image = Image.open('images/survey.jpg')
print(image)
col1.image(image,
        caption='BOT APP1',
        use_column_width=True)
col2.dataframe(df[mask])

