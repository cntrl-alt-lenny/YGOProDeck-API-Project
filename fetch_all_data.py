"""Download all card data from YGOPRODeck and save to an Excel file.

This script fetches the entire card database using the public API with the
`misc=yes` parameter so additional fields are included. The resulting data is
normalized into a pandas DataFrame, expanded so nested objects are flattened and
exported to an Excel file. The file name includes the current timestamp.

This is similar to the example notebook code used in Google Colab but can be
executed from a normal Python environment.
"""

import requests
import pandas as pd
from datetime import datetime


def expand_column(df: pd.DataFrame, column_name: str, prefix: str) -> pd.DataFrame:
    """Explode and flatten list columns from the API.

    Parameters
    ----------
    df : pandas.DataFrame
        The main DataFrame.
    column_name : str
        The column that contains a list of nested records.
    prefix : str
        Prefix to prepend to the expanded column names.
    """
    if column_name in df.columns:
        df = df.explode(column_name).reset_index(drop=True)
        expanded = pd.json_normalize(df[column_name])
        if not expanded.empty:
            expanded.columns = [f"{prefix}_{c}" for c in expanded.columns]
            df = df.drop(columns=[column_name]).reset_index(drop=True)
            df = pd.concat([df, expanded.reset_index(drop=True)], axis=1)
        else:
            df = df.drop(columns=[column_name])
    return df


def main() -> None:
    """Fetch all card data and export to Excel."""
    url = "https://db.ygoprodeck.com/api/v7/cardinfo.php?misc=yes"
    resp = requests.get(url)
    resp.raise_for_status()
    data = resp.json()["data"]
    df = pd.json_normalize(data, sep="_")

    df = expand_column(df, "card_sets", "set")
    df = expand_column(df, "card_images", "image")
    df = expand_column(df, "card_prices", "price")

    now = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    filename = f"all_yugioh_cards_{now}.xlsx"
    with pd.ExcelWriter(filename, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        df.to_excel(writer, index=False, sheet_name="RAW Data")

    print(f"Saved {len(df)} cards to {filename}")


if __name__ == "__main__":
    main()
