import pandas as pd

pd.set_option('display.max_rows', None)

def excel_to_df(path):
    return pd.read_excel(path)

def get_columns(path):
    df = excel_to_df(path)
    return df.columns.tolist()

def get_height(path):
    df = excel_to_df(path)
    return df.shape[0]

def get_rows(path) :
    df = excel_to_df(path)
    return df.map(lambda x: x.strftime('%Y-%m-%d') if isinstance(x, pd.Timestamp) else x).values.tolist()
