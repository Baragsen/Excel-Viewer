import pandas as pd

# Function to read Excel file into a DataFrame
def excel_to_df(path):
    return pd.read_excel(path)

# Function to get column names from DataFrame
def get_columns(path):
    df = excel_to_df(path)
    return df.columns.tolist()

# Function to get the number of rows in DataFrame
def get_height(path):
    df = excel_to_df(path)
    return df.shape[0]

# Function to get rows of DataFrame as a list
def get_rows(path) :
    # Read Excel file into DataFrame
    df = excel_to_df(path)
    # Convert DataFrame rows to list, formatting timestamp objects if present
    return df.apply(lambda x: x.strftime('%Y-%m-%d') if isinstance(x, pd.Timestamp) else x).values.tolist()
