import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import datetime

def get_data_from_excel_SBI(uploaded_file):
    wb = load_workbook(uploaded_file)
    ws = wb["Scheme CG"]

    for row in ws.rows:
        for cell in row:
            if cell.value == "Equity Instruments":
                Equity_Instruments_data_start = cell.row + 1
            if cell.value == "Alternate Investments":
                Alternate_Investments_data_start = cell.row + 1

    df = pd.read_excel(uploaded_file,
                                sheet_name = 'Scheme CG',
                                skiprows=Equity_Instruments_data_start - 1,
                                nrows=Alternate_Investments_data_start -Equity_Instruments_data_start -4)
    df = df.drop(['Name of Instruments'], axis=1)
    df = df.rename(columns={'Isin No.': 'ISIN CODE','Mkt_Value':'MARKET VALUE'})
    df.columns = df.columns.str.upper()
    
    # Merge with lookup table to get company names
    lookup_name_with_isin = get_company_name_from_ISIN()
    df = pd.merge(df, lookup_name_with_isin, on='ISIN CODE', how='left')
    
    df = df.applymap(lambda x: x.upper() if type(x) == str else x)
    df['MONTH'] = month
    df['YEAR'] = str(year)
    df['FUND NAME'] = 'SBI'

    return df

def get_data_from_excel_Kotak(uploaded_file):
    df = pd.DataFrame({})
    return df

def get_data_from_excel_ICICI(uploaded_file):
    wb = load_workbook(uploaded_file)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    # Find the row number of the first cell that contains "Subtotal"
    for row in ws.rows:
        for cell in row:
            if cell.value == "Subtotal":
                rows_till = cell.row
                break # break out of inner loop
        else:
            continue # continue if the inner loop did not break
        break # break out of outer loop

    # Load the data from the Excel file into a DataFrame
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=5, nrows=rows_till - 7)

    # Clean up the DataFrame
    df = df.drop(df[(df['Particulars'] == 'Equity Instruments') | (df['Particulars'] == 'Shares')].index)
    df = df.drop(['Particulars'], axis=1)
    df = df.rename(columns={'ISIN No.': 'ISIN CODE','Industry':'INDUSTRY'})
    df.columns = df.columns.str.upper()
    df = df.applymap(lambda x: x.upper() if type(x) == str else x)
    

    # Merge with lookup table to get company names
    lookup_name_with_isin = get_company_name_from_ISIN()
    df = pd.merge(df, lookup_name_with_isin, on='ISIN CODE', how='left')

    # Add additional columns to the DataFrame
    df['MONTH'] = month
    df['YEAR'] = str(year)
    df['FUND NAME'] = 'ICICI'

    return df

def get_data_from_excel_HDFC(uploaded_file):
    df = pd.DataFrame({})
    return df

def get_data_from_excel_AB(uploaded_file):
    wb = load_workbook(uploaded_file)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    # Find the row number of the first cell that contains "Money Market Instruments:-"
    for row in ws.rows:
        for cell in row:
            if cell.value == "Money Market Instruments:-":
                rows_till = cell.row
                break

    # Load the data from the Excel file into a DataFrame
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, skiprows=5, nrows=rows_till - 10)

    # Clean up the DataFrame
    df = df.drop(['Unnamed: 0', 'Ratings', 'Name of the Instrument'], axis=1)
    df = df.rename(columns={'ISIN No.': 'ISIN CODE','Mkt_Value':'MARKET VALUE','Industry ':'INDUSTRY'})
    df.columns = df.columns.str.upper()
    df = df.applymap(lambda x: x.upper() if type(x) == str else x)

    # Merge with lookup table to get company names
    lookup_name_with_isin = get_company_name_from_ISIN()
    df = pd.merge(df, lookup_name_with_isin, on='ISIN CODE', how='left')

    # Add additional columns to the DataFrame
    df['MONTH'] = month
    df['YEAR'] = str(year)
    df['FUND NAME'] = 'ADITYA BIRLA'

    return df

def get_data_from_excel_UTI(uploaded_file):
    df = pd.DataFrame({})
    return df

def get_data_from_excel_LIC(uploaded_file):
    df = pd.DataFrame({})
    return df

def get_data_from_excel_Max(uploaded_file):
    df = pd.DataFrame({})
    return df

def get_data_from_excel_TATA(uploaded_file):
    df = pd.DataFrame({})
    return df

@st.cache_data()
def get_company_name_from_ISIN():
    df = pd.read_csv("C:\\Users\\raksh\\Documents\\Python Workspace\\Streamlit_examples\\market_analysis\\file_dropbox\\NPS trust data\\ind_nifty500list.csv",
                    usecols = ['ISIN Code','Company Name'])
    df.columns = df.columns.str.upper()
    df = df.applymap(lambda x: x.upper() if isinstance(x, str) else x)
    return df

if 'filter_shares' not in st.session_state:
    st.session_state['filter_shares'] = 'value'


st.set_page_config(page_title = 'Market Analysis',  
                    page_icon = ":bar_chart:",
                    layout = 'wide')


st.header('Market Analysis')
st.subheader('NPS Trust Portfolio')

if os.path.exists('scheme_cg.csv'):
    nps_trust_scheme_cg_df = pd.read_csv('scheme_cg.csv')
    nps_trust_scheme_cg_df.sort_values(['COMPANY NAME', 'YEAR','MONTH'], inplace=True)
else:
    nps_trust_scheme_cg_df = None
    st.info('No NPS Scheme CG Data Available')

add_more_data = st.sidebar.checkbox('Add more Data to Analyse')


if add_more_data:
    uploaded_file = st.file_uploader('Upload your file here')
    nps_fund_names = {'SBI':get_data_from_excel_SBI,
                      'Kotak':get_data_from_excel_Kotak,
                      'ICICI':get_data_from_excel_ICICI,
                      'HDFC':get_data_from_excel_HDFC,
                      'Aditya Birla':get_data_from_excel_AB,
                      'UTI':get_data_from_excel_UTI,
                      'LIC':get_data_from_excel_LIC,
                      'Max':get_data_from_excel_Max,
                      'TATA':get_data_from_excel_TATA}
    nps_upload_form = st.form('Upload NPS Trust Files', clear_on_submit=True)
    nps_fund_name = nps_upload_form.selectbox('Select a Fund to Upload', options = nps_fund_names.keys())
    current_year = datetime.datetime.now().year
    year = nps_upload_form.selectbox('Year',options = [year for year in range(current_year,current_year -10, -1)])
    month = nps_upload_form.selectbox('Month',options = [month for month in range (1,13)])
    pre_check = nps_upload_form.checkbox('Are you Sure?')
    submit = nps_upload_form.form_submit_button('Submit', disabled=not uploaded_file, use_container_width = True)
    if pre_check and submit:
        if nps_fund_name in nps_fund_names:
            uploaded_data = nps_fund_names[nps_fund_name](uploaded_file)
            if isinstance(nps_trust_scheme_cg_df, pd.DataFrame):
                nps_trust_scheme_cg_df = uploaded_data.append(nps_trust_scheme_cg_df)
            else:
                nps_trust_scheme_cg_df = uploaded_data
            nps_trust_scheme_cg_df.to_csv('scheme_cg.csv',index = False)
    
group_by_cols = st.sidebar.multiselect('Select Group by Columns', options = ['YEAR', 'MONTH'], default=['YEAR','MONTH']  )
selected_value = st.sidebar.selectbox('Select a Value',options = ['QUANTITY','MARKET VALUE','% OF PORTFOLIO'])
if isinstance(nps_trust_scheme_cg_df, pd.DataFrame):
    company_name = st.sidebar.multiselect('Filter Shares',options = nps_trust_scheme_cg_df['COMPANY NAME'].unique(),default=None)
    if company_name:
        st.session_state['filter_shares'] = company_name
        nps_trust_scheme_cg_df = nps_trust_scheme_cg_df.query('`COMPANY NAME` == @company_name')
        st.dataframe(nps_trust_scheme_cg_df,use_container_width = True)
    nps_trust_scheme_cg_df = nps_trust_scheme_cg_df.pivot_table(values=selected_value, index='COMPANY NAME', columns=group_by_cols, aggfunc='sum')
    st.dataframe(nps_trust_scheme_cg_df,use_container_width = True)