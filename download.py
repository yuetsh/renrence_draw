import os

import pandas as pd


def download(df: pd.DataFrame, target_dir, filepath):
    name = os.path.splitext(os.path.basename(filepath))[0]
    ext = os.path.splitext(filepath)[1]

    new_filepath = os.path.join(target_dir, name + ext)

    if os.path.exists(new_filepath):
        os.remove(new_filepath)

    df.to_excel(new_filepath, index=False)
