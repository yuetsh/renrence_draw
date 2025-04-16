import os

import pandas as pd


import pandas as pd

def get_school_names(filepath):
    first_sheet = pd.read_excel(filepath, sheet_name=0, skiprows=1)
    school1 = first_sheet["学校"].value_counts().idxmax()
    school2 = first_sheet["学校"].value_counts().idxmin()
    return school1, school2

def download(df: pd.DataFrame, target_dir, filepath):
    name = os.path.splitext(os.path.basename(filepath))[0]
    ext = os.path.splitext(filepath)[1]

    new_filepath = os.path.join(target_dir, name + ext)

    if os.path.exists(new_filepath):
        os.remove(new_filepath)

    df.to_excel(new_filepath, index=False)
