import streamlit as st
from loop_test2 import main as npv_calculate
import pandas as pd
import platform

st.write('# NPV Calculator')

# let user upload a excel
uploaded_file = st.file_uploader(
    "Choose a file",
    accept_multiple_files=False,
    type=['xlsx', 'xls'])

sys = platform.system()
if sys == "Linux":
    st.warning('This app is not supported on Linux system.')


if uploaded_file is not None:
    # get the file name
    file_name = uploaded_file.name

    # save the file
    with open(file_name, 'wb') as f:
        f.write(uploaded_file.getbuffer())

    # use function to calculate NPV
    npv_calculate(file_name)

    # show the result from circular_table.xlsx
    st.write('## Result')
    # read circular_table.xlsx
    df = pd.read_excel('circular_table.xlsx', header=None)
    # show the result
    st.dataframe(df)

