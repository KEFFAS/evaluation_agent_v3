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

    os.makedirs(
        output_folder,
        exist_ok=True
    )

    # =========================================================
    # LOAD CSV
    # =========================================================
    df = pd.read_csv(file_name)

    print("Loaded file successfully.")

    # =========================================================
    # REMOVE EMPTY ROWS/COLUMNS
    # =========================================================
    df.dropna(
        how="all",
        inplace=True
    )

    df.dropna(
        axis=1,
        how="all",
        inplace=True
    )

    # =========================================================
    # CLEAN COLUMN NAMES
    # =========================================================
    df.columns = [

        str(col).strip()

        for col in df.columns

    ]

    # =========================================================
    # REMOVE DUPLICATE HEADER ROWS
    # =========================================================
    first_col = df.columns[0]

    df = df[

        df[first_col]
        .astype(str)
        .str.strip()

        != first_col

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
    # CLEAN TEXT VALUES
    # =========================================================
    for col in df.columns:

        if df[col].dtype == object:

            df[col] = (

                df[col]

                .fillna("")

                .astype(str)

                .str.strip()

            )

    # =========================================================
    # LIKERT MAP
    # =========================================================
    likert_map = {

        # GENERAL RATINGS

        "Excellent": 5,

        "Very Good": 4,

        "Good": 3,

        "Neutral": 3,

        "Satisfactory": 3,

        "Fair": 2,

        "Poor": 2,

        "Very Poor": 1,

        # EXTENT RATINGS

        "Great Extent": 5,

        "To Some Extent": 4,

        "Some Extent": 4,

        "Moderate Extent": 3,

        "Not Sure": 2,

        "Not At All": 1,

        "Not at All": 1

    }

    # =========================================================
    # NORMALIZE LIKERT RESPONSES
    # =========================================================
    normalized_map = {

        str(key).strip().lower(): value

        for key, value
        in likert_map.items()

    }

    # =========================================================
    # DETECT RATING COLUMNS
    # =========================================================
    rating_cols = []

    qualitative_cols = []

    for col in df.columns:

        values = (

            df[col]

            .dropna()

            .astype(str)

            .str.strip()

        )

        # Ignore empty responses
        values = values[
            values != ""
        ]

        if len(values) == 0:

            qualitative_cols.append(col)

            continue

        # Normalize responses
        normalized_values = (

            values

            .str.lower()

        )

        # Count Likert responses
        matching = normalized_values.isin(

            normalized_map.keys()

        ).sum()

        # If majority are Likert responses,
        # treat column as a rating column
        if matching >= len(values) * 0.5:

            rating_cols.append(col)

        else:

            qualitative_cols.append(col)

    # =========================================================
    # CONVERT RATING COLUMNS
    # =========================================================
    for col in rating_cols:

        df[col] = (

            df[col]

            .astype(str)

            .str.strip()

            .str.lower()

            .map(normalized_map)

        )

        df[col] = pd.to_numeric(

            df[col],

            errors="coerce"

        )

    # =========================================================
    # SAVE CLEANED FILE
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

    # =========================================================
    # OUTPUT SUMMARY
    # =========================================================
    print("=" * 50)

    print(
        "ONLINE CLEANING COMPLETED"
    )

    print("=" * 50)

    print(
        f"Rows    : {df.shape[0]}"
    )

    print(
        f"Columns : {df.shape[1]}"
    )

    print("\nRating Columns:")

    for col in rating_cols:

        print(
            f"- {col}"
        )

    print("\nQualitative Columns:")

    for col in qualitative_cols:

        print(
            f"- {col}"
        )

    print(
        f"\nSaved to: {output_file}"
    )

    return output_file