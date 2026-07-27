# utils.py
import pandas as pd

def extract_mawb_columns(filepath, columns_needed):
    """
    Read MAWB Excel file and return a DataFrame with only the columns
    where the header matches (case‑insensitive) the keys in columns_needed.
    columns_needed is a dict: { 'exact_header_name': 'new_column_name' }
    """
    df = pd.read_excel(filepath)
    df.columns = [str(col).strip() for col in df.columns]
    mapping = {}
    for original, new in columns_needed.items():
        found = None
        for col in df.columns:
            if col.upper() == original.upper():
                found = col
                break
        if found is None:
            raise ValueError(f"Column '{original}' not found in the MAWB file.")
        mapping[found] = new
    result = df[list(mapping.keys())].rename(columns=mapping).copy()
    return result