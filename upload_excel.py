from snowflake.snowpark.session import Session
from snowflake.snowpark.types import *

import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title='Excel uploader for Snowflake',  initial_sidebar_state="auto", menu_items=None)
s = st.session_state
SF_ACCOUNT = ''#Account locator
SF_USR = ''#Service account 
SF_PWD = ''#password

conn = {'ACCOUNT': SF_ACCOUNT,'USER': SF_USR,'PASSWORD': SF_PWD}
session = Session.builder.configs(conn).create()

st.header("Excel uploader for Snowflake")
data_file = st.sidebar.file_uploader("Upload Excel file",type=['xlsx'])  

if data_file:

    wb = openpyxl.load_workbook(data_file)
    
    ## Select sheet
    sheet_selector = st.sidebar.selectbox("Select sheet:",wb.sheetnames)     
    df = pd.read_excel(data_file,sheet_selector)
    st.markdown(f"### Currently Selected: `{sheet_selector}`")
    database = st.sidebar.text_input('Database: ')
    schema = st.sidebar.text_input('Schema: (By default this is PUBLIC) ')
    table = st.sidebar.text_input('Table: (If empty then filename+sheet name)')
    st.write(df)
    if st.button('Upload'):
        session.use_database(database)
        if schema == '':
            session.use_schema('PUBLIC')
        else:
            session.use_schema(schema)
        if table == '':
             tbl_name = data_file.name.split('.')[0]
             table_name = f"{tbl_name}_{sheet_selector}"
        else:
            table_name = table
        session_df = session.create_dataframe(df)
        st.markdown("Uploading ")
        try:
            session_df.write.mode('Overwrite').save_as_table(table_name)
            
                    
        except ValueError:
            st.error('Upload failed')
        else:
                st.write('File uploaded')
        
