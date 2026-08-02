# utils.py
import os
import pandas as pd

def ensure_dir(path):
    """Create `path` and return the directory that's actually usable.

    On some serverless hosts (Vercel included), the filesystem available
    while a module is imported/built can be a DIFFERENT snapshot than the
    one available while an actual request is being served - so a directory
    check done once at import/cold-start time can report "writable" and
    still fail later on a real request. The only reliable check is to try
    right at the point of use, every time, and self-heal to /tmp if that
    fails. Cheap when it succeeds (a no-op if the directory already exists),
    so it's safe to call on every request rather than only once at startup.
    """
    try:
        os.makedirs(path, exist_ok=True)
        return path
    except OSError:
        fallback = os.path.join('/tmp', path)
        os.makedirs(fallback, exist_ok=True)
        return fallback

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