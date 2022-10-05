import plotly.express as px
import datetime as dt
import pandas as pd

def preprocess_data(path = '../timesheet_dash_data/RESOURCE SUMMARY.xlsx'):
    wb = pd.read_excel(path,'Sheet1',header=1)

    correspondences = {}

    nonworking = ['10th Day','9th Day','8th day','All Saints Day','Annual Leave','Ascension Day','Easter Monday','National Day',
    'Other Leave','Sick Leave','Special Leave granted by the DG','Whit Monday']

    non_audit = ['Administrative Matters & Support','Covid-19 Sanitary Situation','GC and EXB (incl. Annual report)',
        'HR Management & Recruitment', 'IOS Management - Ad-hoc Requests','IOS Team Meetings', 'JIU Coordination',
        'OAC support, preparation and meetings','Participation to UNESCO Working Groups or Task Forces',
        'Policy or Administrative Manual Item Review','Support to Investigation Unit',
        'Trainings or Workshops']

    audit = ['Audit-Ad-hoc request / Advisory','Audit-Annual Planning', 'Audit-QAIP (incl. TeamMate+ Migration)']

    reco = ['Audit-Recommendation Follow-up']

    correspondences['Nonworking'] = nonworking
    correspondences['Non audit time'] = non_audit
    correspondences['Audit Time'] = audit
    correspondences['Reco Follow-up Time'] = reco

    # Insert correspondences for the different Time Categories
    wb['time'] = 'Audit Time'
    for i in correspondences: 
        wb['time'].mask(wb['Time Category'].isin(correspondences[i]),i,inplace=True)

    # Make Phase 5 -- Recommendation followups be repored as Reco Follow up time
    wb['time'].mask(wb['Phase'].astype('string').str.startswith('5'),'Reco Follow-up Time',inplace=True)

    # Remove IOS/... prefix from Assignment name
    prefix_removed = wb['Assignment Name'].str.extract(r'(^IOS/[a-zA-Z0-9_\./]+)-([a-zA-Z0-9_\' -]+)')[1]
    idxs = ~wb['Assignment Name'].str.startswith('IOS/').fillna(False)
    prefix_removed[idxs] = wb[idxs]['Assignment Name']
    wb['Assignment Name'] = prefix_removed


    return wb
        

def build_treemap(df):
    colour_map = dict(zip(df['time'].sort_values().unique(),px.colors.carto.Safe))
    colour_map['(?)'] = 'lightgrey'

    # hovertemplate='label = %{label}<br>%{color:.data}<br>Hours worked = %{value:.2f}'
    hovertemplate = 'label = %{label}<br>Activity = %{customdata[0]}<br>Working hours = %{value}<extra></extra>'
    fig = px.treemap(df,path=[px.Constant('Month'),'year_month','Resource','time','time_stat','activity','activity_hours'],
    hover_data=['time','activity_hours'],width = 1300,height=800,
    values='working_hours',color = 'time',color_discrete_map=colour_map,maxdepth=5)

    fig.update_layout(title='Time Utilization per Month and Auditor',title_x=0.5)
    fig.update_traces(hovertemplate=hovertemplate)

    return fig

def time_utilisation(df):

    return build_treemap(df)