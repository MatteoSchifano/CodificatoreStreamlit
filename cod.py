import pandas as pd
from openpyxl.utils import column_index_from_string
import re
import string
from typing import Union
from io import BytesIO
from fuzzywuzzy import fuzz


class DfExcel:
    c_cols = False

    def __init__(self,
                 file,
                 start: Union[int, str] = 1,
                 end: Union[int, str] = -1):
        self.file = file

        try:
            self.start = int(start)
        except ValueError:
            self.start = column_index_from_string(str(start))

        try:
            self.end = int(end)
            if self.end == -1:
                self.end = None
        except ValueError:
            self.end = column_index_from_string(str(end))
            if self.end is None or self.end == -1:
                self.end = None

        self.df = self.load_df()
        self.index = self._check_set_index(self.df)

        self.c_cols = False

    def load_df(self):
        df = pd.read_excel(self.file)
        if self.end is not None:
            df = df.iloc[:, self.start-1:self.end]
        else:
            df = df.iloc[:, self.start-1:]
        return df

    def _check_set_index(self, df: pd.DataFrame):
        # Non impostare l'indice automaticamente
        return False

    def generate_c(self):
        cols = self.df.columns
        new_cols = []

        for i, col in enumerate(cols):
            # Controlla se la colonna è la prima e contiene solo numeri
            if i == 0 and pd.to_numeric(self.df[col], errors='coerce').notnull().all():
                new_cols.append(col)  # Aggiungi solo la colonna originale
            else:
                new_cols.append(col)
                new_cols.append(col + '_c')
                self.df[col + '_c'] = ''

        self.df = self.df[new_cols]
        self.c_cols = True
        return self.df

    def delete_c(self):
        cols_c = [col for col in self.df.columns if col.endswith('_c')]
        self.df.drop(columns=cols_c, inplace=True)
        self.c_cols = False
        return self.df

    @staticmethod
    def clear_str(x: str):
        x = str(x).lower()
        x = str(x).replace("'", " ").replace("’", " ").translate(
            str.maketrans('', '', string.punctuation))
        x = re.sub(' +', ' ', str(x))
        x = x.strip()
        if x == 'nan':
            return ''
        return x

    def get_preprocessed_data(self):
        df = self.df.apply(lambda x: x.apply(
            lambda y: self.clear_str(y)) if x.dtype == 'object' else x)
        return df

    def to_excel(self):
        # Resetta l'indice se è stato impostato
        self.df.reset_index(inplace=True, drop=True)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            self.df.to_excel(writer, index=False)
        output.seek(0)
        return output

    def __repr__(self):
        return self.df.__repr__()


class Codificatore:
    def __init__(self,
                 file_aperte,
                 file_codice,
                 sep: str = ';|,|\t',
                 treshold: float = 0.7,
                 start: Union[int, str] = 1,
                 end: Union[int, str] = -1,
                 altro: int = 95):
        self.file_codice = file_codice
        self.file_aperte = file_aperte
        self.treshold = treshold
        self.sep = sep
        self.start = start
        self.end = end
        self.altro = altro

        self.aperte = DfExcel(self.file_aperte, self.start, self.end)
        self.codice = self.carica_codice()

    def carica_codice(self):
        codice = pd.read_csv(self.file_codice, sep=self.sep, engine='python')
        codice['nome'] = codice['nome'].str.lower()
        codice = codice.assign(
            nome=codice['nome'].str.split('/')).explode('nome', ignore_index=True)
        codice['nome'] = codice['nome'].str.strip()
        return codice

    def confronta(self, x: str, codice: pd.DataFrame, treshold=0.7):
        # Se il valore è None o è una stringa vuota, restituisci None o un valore specifico
        if pd.isnull(x) or x == '':
            return None  # O restituisci un valore diverso se necessario, ad esempio self.altro

        ratei = [fuzz.ratio(x, c) for c in codice.nome]
        max_val = max(ratei)

        if 100 in ratei:
            return codice.loc[ratei.index(100), 'codice']
        elif max_val >= treshold * 100:
            return codice.loc[ratei.index(max_val), 'codice']
        else:
            return self.altro

    def _codifica_set(self, sub_df: pd.DataFrame):
        assert sub_df.shape[1] == 2, f'set non valido {sub_df.columns=}'
        sub_df[sub_df.columns[1]] = sub_df[sub_df.columns[0]].apply(
            lambda x: self.confronta(str(x), self.codice, self.treshold))
        return sub_df


    def codifica(self):
        # Sostituisci i valori None/NaN con stringhe vuote
        self.aperte.df = self.aperte.df.fillna('')

        ls = []  # Inizializza ls una volta qui fuori dal ciclo

        # Cicla attraverso le colonne per applicare la codifica
        for i, col in enumerate(self.aperte.df.columns):
            # Salta la prima colonna se è composta solo da numeri
            if i == 0 and pd.to_numeric(self.aperte.df[col], errors='coerce').notnull().all():
                continue

            ls.append(col)

            # Controlla se hai due colonne da codificare (colonna e la sua colonna "_c")
            if len(ls) == 2:
                # Applica la codifica al set di colonne
                self.aperte.df[ls] = self._codifica_set(self.aperte.df[ls].copy())
                ls = []  # Resetta ls dopo aver codificato il set
