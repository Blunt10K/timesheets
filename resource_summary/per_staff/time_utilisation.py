import streamlit as st
import numpy as np
from utils import *

from google.oauth2 import service_account
from gsheetsdb import connect


st.set_page_config(page_title='Resource time utilisation',layout='wide')
st.title('Time utilisation per resource')


@st.experimental_singleton()
def connect_sheets():
    # Create a connection object.
    credentials = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
        ],
    )
    return connect(credentials=credentials)

@st.experimental_memo(ttl=600)
def run_query(query):
    conn = connect_sheets()
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


with st.sidebar:
    st.header("Dates between")
    start = pd.Timestamp(st.date_input("Start",min(df['Date']),min(df['Date']),max(df['Date'])))
    end = pd.Timestamp(st.date_input("End",max(df['Date']),min(df['Date']),max(df['Date'])))

resources = st.multiselect("Resource",names,names)

fig = time_utilisation(df[(df['Resource'].isin(resources)) & (df['Date'].between(start,end))])

# # source = df
st.plotly_chart(fig)
