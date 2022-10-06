import streamlit as st
import numpy as np
from utils import *

from google.oauth2 import service_account
from gsheetsdb import connect

# Create a connection object.
credentials = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
    ],
)
conn = connect(credentials=credentials)

st.set_page_config(page_title='Resrouce time utilisation')
st.title('Time utilisation per resource')

# @st.experimental_memo
# def load_data():
#     df = preprocess_data()
#     names = list(df['Resource'].unique())
#     names.sort()
#     cols = ['Project Hours','Admin Hours', 'Nonworking Hours', 'Holiday', 'Total (All Entries)','Utilization (All Entries)',
#     'Comment','time_category_percentage','Type','Team','Assignment Code','Audit Plan Title','Assignment Name','Phase',
#     'Time Category','activity_percentage']

#     return names, df.drop(columns = cols).sort_values('Date',ascending=False)

@st.cache(ttl=600)
def run_query(query):
    rows = conn.execute(query, headers=1)
    rows = rows.fetchall()
    df = pd.DataFrame(rows)
    df['Date'] = pd.to_datetime(df['Date'])
    df['year_month'] = pd.to_datetime(df['year_month']).dt.strftime('%B, %Y')
    names = list(df['Resource'].unique())
    names.sort()

    return names, df

sheet_url = st.secrets["private_gsheets_url"]
names, df = run_query(f'SELECT * FROM "{sheet_url}"')
# col1, col2 = st.columns(2)


with st.sidebar:
    st.header("Dates between")
    start = pd.Timestamp(st.date_input("Start",min(df['Date']),min(df['Date']),max(df['Date'])))
    end = pd.Timestamp(st.date_input("End",max(df['Date']),min(df['Date']),max(df['Date'])))

resources = st.multiselect("Resource",names,names)

fig = time_utilisation(df[(df['Resource'].isin(resources)) & (df['Date'].between(start,end))])

# # source = df
st.plotly_chart(fig)
