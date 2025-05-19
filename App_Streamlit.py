import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime


conn = st.connection("gsheets", type=GSheetsConnection)
data = conn.read(worksheet="Tumble_cup")
st.subheader("Current Orders")
st.dataframe(data)