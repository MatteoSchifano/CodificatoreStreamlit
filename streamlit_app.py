import streamlit as st
from cod import Codificatore, DfExcel
import pandas as pd
from io import BytesIO

pd.options.display.max_columns = 30

st.set_page_config(page_title='Codificatore')

st.title('Codificatore')

# Carica file da codificare
path_file = st.file_uploader('Carica file da codificare', type=['xlsx', 'xls'])

# Carica codice
path_code = st.file_uploader('Carica codice', type=['csv'])

# Separatore
sep = st.text_input('Separatore', value=',')

# Colonna d'inizio codifica
start = st.text_input("Colonna d'inizio codifica", value='A')

# Colonna di fine codifica
end = st.text_input(
    "Colonna di fine codifica (-1 per ultima colonna)", value='-1')

# Soglia di somiglianza (treshold)
treshold = st.number_input(
    "Inserire la soglia di somiglianza (treshold) [0.0, 1.0]", min_value=0.0, max_value=1.0, value=0.7)

# Valore Other
altro = st.number_input("Valore 'Other'", value=95)

# Bottone per controllare i file
if st.button('Controlla file'):
    if path_file is not None and path_code is not None:
        cod = Codificatore(
            file_aperte=path_file,
            file_codice=path_code,
            sep=sep,
            start=start,
            end=end,
            altro=altro,
            treshold=treshold
        )
        st.write('Anteprima del file da codificare:')
        st.write(cod.aperte.df)
        st.write('Anteprima del codice:')
        st.write(cod.codice)
        st.session_state['cod'] = cod
    else:
        st.warning('Per favore, carica entrambi i file.')

# Bottone per eseguire la codifica
if st.button('Codifica'):
    if 'cod' in st.session_state:
        cod = st.session_state['cod']
        st.write('Sto codificando, attendi un attimo...')
        cod.aperte.generate_c()
        cod.aperte.get_preprocessed_data()
        cod.codifica()
        output = cod.aperte.to_excel()
        st.write(
            'Codifica effettuata! Anteprima del file codificato:')
        st.write(cod.aperte.df)
        st.download_button(
            label='Scarica il file codificato',
            data=output,
            file_name='codificato.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.warning('Per favore, controlla i file prima di codificare.')
