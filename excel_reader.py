import pandas as pd

def excel_to_df(path):
    return pd.read_excel(path)

def get_columns(path):
    df = excel_to_df(path)
    return df.columns.tolist()


