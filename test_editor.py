import streamlit as st
import pandas as pd
if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
ed = st.data_editor(st.session_state.df, num_rows="dynamic", key="ed")
st.session_state.df = ed.copy()
st.write(ed)
