import os
import pandas as pd


def clean_online(
    file_name,
    output_folder="outputs"
):
    """
    Clean Online End-of-Event Evaluation data.

    Parameters
    ----------
    file_name : str
        Raw CSV file.

    output_folder : str

    Returns
    -------
    str
        Path to cleaned Excel file.
    """

    os.makedirs(output_folder, exist_ok=True)

    # =========================================================
    # LOAD CSV
    # =========================================================
    df = pd.read_csv(file_name)

    print("Loaded file successfully.")

    # =========================================================
    # REMOVE EMPTY ROWS/COLUMNS
    # =========================================================
    df.dropna(how="all", inplace=True)
    df.dropna(axis=1, how="all", inplace=True)

    # =========================================================
    # REMOVE DUPLICATE HEADER ROWS
    # =========================================================
    first_col = df.columns[0]

    df = df[
        df[first_col].astype(str).str.strip()
        != first_col
    ]

    # =========================================================
    # CLEAN COLUMN NAMES
    # =========================================================
    df.columns = [

        str(col).strip()

        for col in df.columns

    ]

    # =========================================================
    # REMOVE SYSTEM COLUMNS
    # =========================================================
    drop_cols = [

        "Timestamp",

        "Email Address",

        "Name",

        "Response number"

    ]

    existing = [

        col

        for col in drop_cols

        if col in df.columns

    ]

    df.drop(
        columns=existing,
        inplace=True
    )

    # =========================================================
    # CLEAN TEXT
    # =========================================================
    for col in df.columns:

        if df[col].dtype == object:

            df[col] = (

                df[col]

                .astype(str)

                .str.strip()

            )

    # =========================================================
    # LIKERT MAP
    # =========================================================
    likert_map = {

        "Excellent": 5,

        "Very Good": 4,

        "Good": 3,

        "Neutral": 3,

        "Satisfactory": 3,

        "Fair": 2,

        "Poor": 2,

        "Very Poor": 1,

        "Great Extent": 5,

        "To Some Extent": 4,

        "Moderate Extent": 3,

        "Not Sure": 2,

        "Not At All": 1,

        "Not at All": 1

    }

    # =========================================================
    # DETECT QUALITATIVE COLUMNS
    # =========================================================
    qualitative_keywords = [

        "suggest",

        "comment",

        "additional",

        "topic",

        "interest",

        "improve",

        "experience"

    ]

    qualitative_cols = []

    for col in df.columns:

        name = str(col).lower()

        if any(

            keyword in name

            for keyword in qualitative_keywords

        ):

            qualitative_cols.append(col)

    # =========================================================
    # CONVERT RATING COLUMNS
    # =========================================================
    rating_cols = [

        col

        for col in df.columns

        if col not in qualitative_cols

    ]

    for col in rating_cols:

        df[col] = (

            df[col]

            .replace(likert_map)

        )

        df[col] = pd.to_numeric(

            df[col],

            errors="coerce"

        )

    # =========================================================
    # SAVE
    # =========================================================
    base_name = os.path.splitext(

        os.path.basename(file_name)

    )[0]

    output_file = os.path.join(

        output_folder,

        f"{base_name}_cleaned.xlsx"

    )

    df.to_excel(

        output_file,

        index=False

    )

    print("=" * 40)
    print("ONLINE CLEANING COMPLETED")
    print("=" * 40)

    print(f"Rows    : {df.shape[0]}")
    print(f"Columns : {df.shape[1]}")

    print("\nQualitative Columns:")

    for col in qualitative_cols:

        print(f"- {col}")

    print(f"\nSaved to: {output_file}")

    return output_file