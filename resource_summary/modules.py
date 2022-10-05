import plotly.express as px
import pandas as pd
import numpy as np
from docx import Document
from os import listdir
from os.path import join as osjoin, expanduser

def preprocess_data(path = 'RESOURCE UTILIZATION_20220902.xlsx'):
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
    wb['Reporting categories'] = 'Audit Time'
    for i in correspondences: 
        wb['Reporting categories'].mask(wb['Time Category'].isin(correspondences[i]),i,inplace=True)

    # Make Phase 5 -- Recommendation followups be repored as Reco Follow up time
    wb['Reporting categories'].mask(wb['Phase'].astype('string').str.startswith('5'),'Reco Follow-up Time',inplace=True)


    return wb


def time_utilisation(wb):
   # exclude nonworking records
   df = wb.loc[wb['Reporting categories']!='Nonworking']
   df = df.rename(columns={'Reporting categories':'time'})

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

   colour_map = dict(zip(df['time'].sort_values().unique(),px.colors.carto.Safe))
   colour_map['(?)'] = 'lightgrey'

   # hovertemplate='label = %{label}<br>%{color:.data}<br>Hours worked = %{value:.2f}'
   hovertemplate = 'label = %{label}<br>Activity = %{customdata[0]}<br>Working hours = %{value}<extra></extra>'
   fig = px.treemap(df,path=[px.Constant('Month'),'year_month','Resource','time','time_stat','activity','activity_hours'],
   hover_data=['time','activity_hours'],
   values='working_hours',color = 'time',color_discrete_map=colour_map,maxdepth=5)

   fig.update_layout(title='Time Utilization per Month and Auditor',title_x=0.5)
   fig.update_traces(hovertemplate=hovertemplate)

   return fig


def hours_per_assignment_year_phase(wb):
    df = wb[~wb['Phase'].isna()]

    df = df.assign(Working_hours = df[['Project Hours', 'Admin Hours']].sum(axis = 1))
    df.rename(columns = {'Working_hours':'Working Hours'},inplace=True)

    df['month'] = df['Date'].dt.strftime('%B')
    df['year'] = df['Date'].dt.strftime('%Y')

    # group by assignment, phase and resource
    work_share = (df.groupby(['year','Assignment Name','Phase','Resource'])
            .agg(working_hours = pd.NamedAgg(column="Total (All Entries)", aggfunc="sum")))

    work_share['assignment duration (hours)'] = work_share.groupby(level=[0,1]).sum()
    work_share['phase duration (hours)']=work_share.groupby(level=[0,1,2]).sum()[['working_hours']]

    work_share['phase_share'] = (100* work_share['working_hours']/work_share['phase duration (hours)'])
    work_share['display_phase_share'] = work_share['phase_share'].round(1)
    
    work_share = work_share.reset_index()

    colour_map = dict(zip(work_share['Phase'].sort_values().unique(),px.colors.sequential.Blugrn))
    colour_map['(?)'] = 'lightgrey'
    
    work_share.replace(np.nan,None,inplace=True)

    work_share['staff_share'] = work_share['Resource'] + ', ' + work_share['display_phase_share'].astype('string') + '%'
    work_share['(Average) Working hours']=work_share['working_hours']

    phase_count = work_share.groupby(['year','Assignment Name','Phase']).count().reset_index(level=2)[['Phase']]
    work_share['Phase count'] = work_share['phase_share']/100

    work_share['staff_share'] = work_share['Resource'] + ', ' + work_share['working_hours'].astype('string') + ' hours '\
    + '(' + work_share['display_phase_share'].astype('string') + '%)'


    hovertemplate='label = %{label}<br>Working hours = %{value}<br>Auditor = %{customdata[1]}<br>Phase share = %{customdata[0]}%<extra></extra>'
    fig = px.treemap(work_share,path = [px.Constant('Assignment'),'Assignment Name','year','Phase','staff_share'],
    color='Phase',branchvalues='total',maxdepth=4,hover_data=['display_phase_share','Resource'],
    values= 'working_hours',color_discrete_sequence=px.colors.sequential.Blugrn,color_discrete_map=colour_map,
    title='Profile of Hours Charged per Assignment, broken down by year then phase')

    fig.update_traces(root_color = 'lightgray')
    fig.update_layout(title_x= .5)
    fig.update_traces(hovertemplate=hovertemplate)

    return fig


def hours_per_year(wb):
    df = wb[~wb['Phase'].isna()]

    df = df.assign(Working_hours = df[['Project Hours', 'Admin Hours']].sum(axis = 1))
    df.rename(columns = {'Working_hours':'Working Hours'},inplace=True)

    df['month'] = df['Date'].dt.strftime('%B')
    df['year'] = df['Date'].dt.strftime('%Y')

    # group by assignment, phase and resource
    work_share = (df.groupby(['year','Assignment Name','Phase','Resource'])
            .agg(working_hours = pd.NamedAgg(column="Total (All Entries)", aggfunc="sum")))

    work_share['assignment duration (hours)'] = work_share.groupby(level=[0,1]).sum()
    work_share['phase duration (hours)']=work_share.groupby(level=[0,1,2]).sum()[['working_hours']]

    work_share['phase_share'] = (100* work_share['working_hours']/work_share['phase duration (hours)'])
    work_share['display_phase_share'] = work_share['phase_share'].round(1)
    
    work_share = work_share.reset_index()

    colour_map = dict(zip(work_share['Phase'].sort_values().unique(),px.colors.sequential.Blugrn))
    colour_map['(?)'] = 'lightgrey'
    
    work_share.replace(np.nan,None,inplace=True)

    work_share['staff_share'] = work_share['Resource'] + ', ' + work_share['display_phase_share'].astype('string') + '%'
    work_share['(Average) Working hours']=work_share['working_hours']

    phase_count = work_share.groupby(['year','Assignment Name','Phase']).count().reset_index(level=2)[['Phase']]
    work_share['Phase count'] = work_share['phase_share']/100

    work_share['staff_share'] = work_share['Resource'] + ', ' + work_share['working_hours'].astype('string') + ' hours '\
    + '(' + work_share['display_phase_share'].astype('string') + '%)'


    hovertemplate='label = %{label}<br>Working hours = %{value}<br>Auditor = %{customdata[1]}<br>Phase share = %{customdata[0]}%<extra></extra>'
    fig = px.treemap(work_share,path = [px.Constant('Year'),'year','Assignment Name','Phase','staff_share'],
    color='Phase',branchvalues='total',maxdepth=4,hover_data=['display_phase_share','Resource'],
    values= 'working_hours',color_discrete_sequence=px.colors.sequential.Blugrn,color_discrete_map=colour_map,
    title='Profile of Hours Charged per Year')

    fig.update_traces(root_color = 'lightgray')
    fig.update_layout(title_x= .5)
    fig.update_traces(hovertemplate=hovertemplate)

    return fig


def hours_per_assignment_phase_year(wb):
    df = wb[~wb['Phase'].isna()]

    df = df.assign(Working_hours = df[['Project Hours', 'Admin Hours']].sum(axis = 1))
    df.rename(columns = {'Working_hours':'Working Hours'},inplace=True)

    df['month'] = df['Date'].dt.strftime('%B')
    df['year'] = df['Date'].dt.strftime('%Y')

    # group by assignment, phase and resource
    work_share = (df.groupby(['year','Assignment Name','Phase','Resource'])
            .agg(working_hours = pd.NamedAgg(column="Total (All Entries)", aggfunc="sum")))

    work_share['assignment duration (hours)'] = work_share.groupby(level=[0,1]).sum()
    work_share['phase duration (hours)']=work_share.groupby(level=[0,1,2]).sum()[['working_hours']]

    work_share['phase_share'] = (100* work_share['working_hours']/work_share['phase duration (hours)'])
    work_share['display_phase_share'] = work_share['phase_share'].round(1)
    
    work_share = work_share.reset_index()

    colour_map = dict(zip(work_share['Phase'].sort_values().unique(),px.colors.sequential.Blugrn))
    colour_map['(?)'] = 'lightgrey'
    
    work_share.replace(np.nan,None,inplace=True)

    work_share['staff_share'] = work_share['Resource'] + ', ' + work_share['display_phase_share'].astype('string') + '%'
    work_share['(Average) Working hours']=work_share['working_hours']

    phase_count = work_share.groupby(['year','Assignment Name','Phase']).count().reset_index(level=2)[['Phase']]
    work_share['Phase count'] = work_share['phase_share']/100

    work_share['staff_share'] = work_share['Resource'] + ', ' + work_share['working_hours'].astype('string') + ' hours '\
    + '(' + work_share['display_phase_share'].astype('string') + '%)'


    hovertemplate='label = %{label}<br>Working hours = %{value}<br>Auditor = %{customdata[1]}<br>Phase share = %{customdata[0]}%<extra></extra>'
    fig = px.treemap(work_share,path = [px.Constant('Assignment'),'Assignment Name','Phase','year','staff_share'],
    color='Phase',branchvalues='total',maxdepth=4,hover_data=['display_phase_share','Resource'],
    values= 'working_hours',color_discrete_sequence=px.colors.sequential.Blugrn,color_discrete_map=colour_map,
    title='Profile of Hours Charged per Assignment, broken down by phase then year')

    fig.update_traces(root_color = 'lightgray')
    fig.update_layout(title_x= .5)
    fig.update_traces(hovertemplate=hovertemplate)

    return fig


def extract_load(path = '~/Downloads/OneDrive_1_05-09-2022'):
    header_of_interest = ['Activity/ Persons-days','Planning','Fieldwork','Draft report','Finalization','Total']
    int_cols = ['Planning', 'Fieldwork', 'Draft report', 'Finalization', 'Total']
    dtypes = ['int64']*len(int_cols)

    doc_files = [i for i in listdir(expanduser(path)) if '.docx' in i]

    for f in doc_files:

        sheet_name = f[:-5]
   
        file = expanduser(osjoin(path,f))
        print(sheet_name)
        doc = Document(file)
        tables = doc.tables

        
        for t in tables:
            header = [cell.text for cell in t.rows[0].cells]
            if header == header_of_interest:
                data = [[cell.text for cell in row.cells] for row in t.rows[1:]]
                break

        df = pd.DataFrame(data,columns = header_of_interest)

        for a in int_cols:
            df[a] = pd.to_numeric(df[a],errors = 'coerce',downcast='integer')    
        
        with pd.ExcelWriter('Audit budget days.xlsx',mode = 'a',engine='openpyxl',if_sheet_exists='replace') as writer:
            df.to_excel(writer,sheet_name=sheet_name,index=False)

