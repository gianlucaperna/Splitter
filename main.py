import pandas as pd
import streamlit as st
import io

# buffer to use for excel writer
buffer = io.BytesIO()

#n_persone = 30

def new_assignement(size):
    return np.random.randint(1, 7, size)


if __name__ == '__main__':

    st.set_page_config(page_title='XLSX Grouper', layout='wide')
    #IMAGE-UPLOAD FILE
    #b.title('XLSX Grouper')
    st.write('Dividere persone in diversi gruppi')
    uploaded_file = st.file_uploader("Choose a file", type="xlsx")

    #DASHBOARD
    if uploaded_file is not None:
    # To read file as bytes:
        st.write(f'You selected {uploaded_file.name}')
        df = pd.read_excel(uploaded_file).reset_index(drop=True)
        n_sessione = st.number_input("numero di gruppi", min_value=1, format="%d")
        df["Utente - Luogo"] = df["Utente - Luogo"].str.upper()
        df["Utente - ID utente"] = df["Utente - ID utente"].str.upper()
        df = df.sort_values(by="Utente - Luogo")
        df["session"] =  list(range(1, len(df) + 1))
        df["session"] =  df["session"].apply(lambda x: x % n_sessione)
        df["session"] =  df["session"]+1
        df = df.sort_values(by="session")
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
           # Close the Pandas Excel writer and output the Excel file to the buffer
            writer.save()

            download2 = st.download_button(
            label="Download data as Excel",
            data=buffer,
            file_name='Executed.xlsx',
            mime='application/vnd.ms-excel'
            )
