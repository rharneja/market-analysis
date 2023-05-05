import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import os

def get_data_from_excel_SBI(uploaded_file,month,year):
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
    #df = pd.DataFrame({})
    print('Do Nothing')
    return True

def get_data_from_excel_ICICI(uploaded_file,month,year):
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

def get_data_from_excel_AB(uploaded_file,month,year):
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
    df = pd.read_csv("ind_nifty500list.csv",
                    usecols = ['ISIN Code','Company Name'])
    df.columns = df.columns.str.upper()
    df = df.applymap(lambda x: x.upper() if isinstance(x, str) else x)
    return df

def highlight_less_than_previous_and_next(s):
    # create an empty style object
    style = pd.Series(np.nan, index=s.index.astype(float))
    # loop over columns in the dataframe
    for i in range(1, len(s)):
        if isinstance(s, pd.Series) and s[i] < s[i-1]:
            # set the style for the cell to red if the value is less than the previous column's value
            style[i] = 'color: red'
        elif isinstance(s, pd.Series) and s[i] > s[i-1]:
            # set the style for the cell to green if the value is greater than the previous column's value
            style[i] = 'color: green'
        elif isinstance(s, pd.DataFrame) and s.iloc[:, i].lt(s.iloc[:, i-1]).any():
            # set the style for the entire column to red if any value in the column is less than the previous column's value
            style.iloc[:, i] = 'color: red'
        elif isinstance(s, pd.DataFrame) and s.iloc[:, i].gt(s.iloc[:, i-1]).any():
            # set the style for the entire column to green if any value in the column is greater than the previous column's value
            style.iloc[:, i] = 'color: green'
    return style

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

nps_upload_form = st.form('Upload NPS Trust Files', clear_on_submit=True)
uploaded_files = nps_upload_form.file_uploader('Upload your files here',accept_multiple_files=True)
submit = nps_upload_form.form_submit_button('Submit', use_container_width = True)
nps_fund_names = {'SBI':get_data_from_excel_SBI,
                    'KOTAK':get_data_from_excel_Kotak,
                    'ICICI':get_data_from_excel_ICICI,
                    'HDFC':get_data_from_excel_HDFC,
                    'AB':get_data_from_excel_AB,
                    'UTI':get_data_from_excel_UTI,
                    'LIC':get_data_from_excel_LIC,
                    'MAX':get_data_from_excel_Max,
                    'TATA':get_data_from_excel_TATA}
months = {'JAN' : 1, 'FEB' : 2, 'MAR' : 3, 'APR' : 4, 'MAY' : 5, 'JUN' : 6,
         'JUL' : 7, 'AUG' : 8, 'SEP' : 9, 'OCT': 10, 'NOV' : 11, 'DEC' : 12}

if submit:
    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        fund_manager, month, year = file_name.split('_')
        year = year.split('.')[0]
        month = months[month.upper()]

        uploaded_data = nps_fund_names[fund_manager](uploaded_file,month,year)
        if isinstance(nps_trust_scheme_cg_df, pd.DataFrame):
            nps_trust_scheme_cg_df = pd.concat([nps_trust_scheme_cg_df, uploaded_data])
        else:
            nps_trust_scheme_cg_df = uploaded_data
        nps_trust_scheme_cg_df.to_csv('scheme_cg.csv',index = False)
    st.commands.execution_control.rerun()

if st.sidebar.button('Reset'):
    if os.path.exists("scheme_cg.csv"):
        os.remove("scheme_cg.csv")
    st.commands.execution_control.rerun()
group_by_cols = st.sidebar.multiselect('Select Group by Columns', options = ['YEAR', 'MONTH'], default=['YEAR','MONTH']  )
selected_value = st.sidebar.selectbox('Select a Value',options = ['QUANTITY','MARKET VALUE','% OF PORTFOLIO'])
if isinstance(nps_trust_scheme_cg_df, pd.DataFrame):
    company_name = st.sidebar.multiselect('Filter Shares',options = nps_trust_scheme_cg_df['COMPANY NAME'].unique(),default=None)
    if company_name:
        st.session_state['filter_shares'] = company_name
        nps_trust_scheme_cg_df = nps_trust_scheme_cg_df.query('`COMPANY NAME` == @company_name')
        st.dataframe(nps_trust_scheme_cg_df,use_container_width = True)
    
    
    nps_pivot = nps_trust_scheme_cg_df.pivot_table(values=selected_value, index='COMPANY NAME', columns=group_by_cols, aggfunc='sum')
    #nps_pivot = nps_pivot.style.applymap(highlight_less_than_previous_and_next)
    st.dataframe(nps_pivot,use_container_width = True)