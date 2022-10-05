import streamlit as st
import numpy as np
from utils import *

# @st.cache
def load_data(wb):
    # exclude nonworking records
    df = wb.loc[wb['time']!='Nonworking']

    # merge Time Category and Phase information in one -- activity
    non_audit = df[df['Time Category'].isna()]['Assignment Name']
    audit = df[~df['Time Category'].isna()]['Time Category']
    df['activity'] = pd.concat((non_audit,audit))

    # extract year from date, calculate working hours
    df['Date'] = pd.to_datetime(df['Date'])
    df = df.assign(year_month=df['Date'].dt.strftime('%B, %Y'))
    df = df.assign(working_hours=df[['Project Hours', 'Admin Hours']].sum(axis=1))

    df = df.groupby(['Resource','year_month','time','activity']).sum().drop(columns=['Project Hours', 'Admin Hours',
        'Nonworking Hours', 'Holiday', 'Total (All Entries)',
        'Utilization (All Entries)'])

    df['monthly_hours'] = df.groupby(level = [0,1]).sum()
    df['time_category_hours'] = df.groupby(level = [0,1,2]).sum()[['working_hours']]

    df['time_category_percentage']  = (100*df['time_category_hours']/df['monthly_hours']).round(1)
    df['activity_percentage'] = (100*df['working_hours']/df['time_category_hours']).round(1)

    df = df.reset_index()

    df.replace(np.nan,None,inplace=True)


    df = df.astype({'time_category_percentage':'string','activity_percentage':'string'})
    df['time_category_percentage']+='%'
    df['activity_percentage']+='%'

    df['time_stat'] = df['time_category_hours'].astype('string') + ' hours (' + df["time_category_percentage"] + ')'
    df['activity_hours']  =  df['working_hours'].astype('string') + ' hours (' + df["activity_percentage"] + ')'

    return df


wb = preprocess_data()
df = load_data(wb)

fig = time_utilisation(df)


st.set_page_config(layout='centered',page_title='Charged hours breakdwon per staff')

# data_load_state = st.text('Loading data...')

# print(df.head())
# data_load_state.text('Done! (using st.cache)')

# st.subheader('Raw data')
# st.write(df)

# # source = df
# st.plotly_chart(fig)
