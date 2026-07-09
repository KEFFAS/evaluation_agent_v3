import os
import pandas as pd
from openpyxl.styles import Font


def clean_fe(
    file_name,
    output_folder="outputs"
):
    """
    Cleans Facilitator Evaluation raw data.

    Parameters
    ----------
    file_name : str
        Path to the raw Excel file.

    output_folder : str
        Folder where the cleaned file will be saved.

    Returns
    -------
    str
        Path to the cleaned Excel file.
    """

    os.makedirs(output_folder, exist_ok=True)

    # =========================================================
    # LOAD FILE
    # =========================================================
    df = pd.read_excel(file_name, header=None)

    print("Original shape:", df.shape)

    # =========================================================
    # REMOVE FIRST ROW
    # =========================================================
    df = df.iloc[1:].reset_index(drop=True)

    # =========================================================
    # SET HEADER
    # =========================================================
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)

    # =========================================================
    # REMOVE EMPTY ROWS
    # =========================================================
    df = df.dropna(how="all")

    # =========================================================
    # REMOVE UNNAMED COLUMNS
    # =========================================================
    df = df.loc[
        :,
        ~df.columns.astype(str).str.contains(
            "unnamed",
            case=False
        )
    ]

    # =========================================================
    # CLEAN COLUMN NAMES
    # =========================================================
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.title()
    )

    # =========================================================
    # CLEAN TEXT VALUES
    # =========================================================
    for col in df.columns:

        if df[col].dtype == "object":

            df[col] = (
                df[col]
                .astype(str)
                .str.strip()
            )

    # =========================================================
    # CONVERT NUMERIC COLUMNS
    # =========================================================
    for col in df.columns:

        converted = pd.to_numeric(
            df[col],
            errors="coerce"
        )

        if converted.notna().sum() > len(df) * 0.5:

            df[col] = converted

    # =========================================================
    # CONVERT WHOLE NUMBERS TO INTEGER
    # =========================================================
    for col in df.select_dtypes(include="number").columns:

        if (df[col].dropna() % 1 == 0).all():

            df[col] = df[col].astype("Int64")

    # =========================================================
    # STANDARDIZE COLUMN NAMES
    # =========================================================
    rename_map = {

        "Programme Title": "Program Title",

        "Facilitator": "Lecturer Name",

        "Lecturer": "Lecturer Name",

        "Topic": "Topic Description"

    }

    df = df.rename(columns=rename_map)

    # =========================================================
    # REORDER COLUMNS
    # =========================================================
    desired_order = [

        "Date",

        "Program Title",

        "Topic Description",

        "Lecturer Name",

        "Punctuality",

        "Presentation Flow",

        "Handling Questions",

        "Active Participation Of Learners",

        "Use Of Visual Aids",

        "Relevance Of Subject To Workplace",

        "Use Of Relevant Examples",

        "Knowledge Of Subject",

        "Treats Participants With Dignity And Respect",

        "Variety And Appropriateness Of Training Methods",

        "Like",

        "Suggestions",

        "Status",

        "Session Code",

        "Timetable No",

        "Campus"

    ]

    existing_cols = [

        col

        for col in desired_order

        if col in df.columns

    ]

    remaining_cols = [

        col

        for col in df.columns

        if col not in existing_cols

    ]

    df = df[
        existing_cols +
        remaining_cols
    ]

    # =========================================================
    # SORT BY TOPIC
    # =========================================================
    if "Topic Description" in df.columns:

        df = (

            df

            .sort_values("Topic Description")

            .reset_index(drop=True)

        )

    print("Cleaned shape:", df.shape)

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

    with pd.ExcelWriter(

        output_file,

        engine="openpyxl"

    ) as writer:

        df.to_excel(

            writer,

            index=False,

            sheet_name="Cleaned Data"

        )

        worksheet = writer.sheets["Cleaned Data"]

        for cell in worksheet[1]:

            cell.font = Font(bold=True)

    print(f"✅ Cleaned file saved as: {output_file}")

    return output_file