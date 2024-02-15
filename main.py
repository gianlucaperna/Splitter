import pandas as pd
import streamlit as st
import io
import xlsxwriter
import numpy as np
# buffer to use for excel writer
buffer = io.BytesIO()

#n_persone = 30
np.float = float
np.int = int

if __name__ == '__main__':

    st.set_page_config(page_title='XLSX Grouper', layout='wide')
    #IMAGE-UPLOAD FILE
    #b.title('XLSX Grouper')
    st.write('Dividere persone in diversi gruppi')
    #st.write(np.__version__)
    #st.write(pd.__version__)
    uploaded_file = st.file_uploader("Choose a file", type="xlsx")
    #DASHBOARD
    if uploaded_file is not None:
    # To read file as bytes:
        st.write(f'You selected {uploaded_file.name}')
        df = pd.read_excel(uploaded_file).reset_index(drop=True)
        columns = df.columns
        col_to_group = st.selectbox("Colonna da usare per fare ordinamento:", columns )
        n_sessione = st.number_input("numero di gruppi", min_value=1, format="%d")
        if st.button("RUN"):

            group = ['Utente - Nome utente', 'Utente - Cognome utente', 'Utente - Luogo ID']
            df[group] = df[group].apply(lambda x: x.str.upper())
            gb = df.groupby(group).agg(list).reset_index()
            gb = gb.sort_values(by=col_to_group)
            gb["session"] =  list(range(1, len(gb) + 1))
            gb["session"] =  gb["session"].apply(lambda x: x % n_sessione)
            gb["session"] =  gb["session"]+1
            df=gb.explode("Utente - ID utente")
            df = df.sort_values(by=["session"]+group)
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Sheet1', index=False)
               # Close the Pandas Excel writer and output the Excel file to the buffer
               # writer.save()

            download2 = st.download_button(
            label="Download data as Excel",
            data=buffer,
            file_name='Executed.xlsx',
            mime='application/vnd.ms-excel'
            )
