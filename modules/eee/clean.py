import os
import pandas as pd
from openpyxl.styles import Font


def clean_eee(file_path, output_folder="outputs"):
    """
    Cleans a KSG End of Event Evaluation (EEE) Excel file.

    Parameters
    ----------
    file_path : str
        Path to uploaded Excel file.

    output_folder : str
        Folder where cleaned file will be saved.

    Returns
    -------
    str
        Path to cleaned Excel file.
    """

    os.makedirs(output_folder, exist_ok=True)

    # =========================================================
    # LOAD FILE
    # =========================================================
    df = pd.read_excel(file_path, header=None)

    print(f"Loaded: {file_path}")
    print(f"Original shape: {df.shape}")

    # =========================================================
    # REMOVE FIRST ROW
    # =========================================================
    df = df.iloc[1:].reset_index(drop=True)

    # =========================================================
    # SET TRUE HEADER
    # =========================================================
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)

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
    # SMART NUMERIC CONVERSION
    # =========================================================
    for col in df.columns:

        converted = pd.to_numeric(
            df[col],
            errors="coerce"
        )

        if converted.notna().sum() > len(df) * 0.5:
            df[col] = converted

    # =========================================================
    # STANDARDIZE IDENTIFIERS
    # =========================================================
    rename_map = {

        "Program Name": "Program Title",

        "Coordinator": "Coordinator Name"

    }

    df = df.rename(columns=rename_map)

    # =========================================================
    # DROP SYSTEM COLUMNS
    # =========================================================
    drop_cols = [

        "Average Rating",

        "Status",

        "Course Duration (Days)"

    ]

    existing = [

        col for col in drop_cols

        if col in df.columns

    ]

    df = df.drop(
        columns=existing,
        errors="ignore"
    )

    # =========================================================
    # DETECT KEY COLUMNS
    # =========================================================
    objective_cols = [

        col for col in df.columns

        if "objective" in str(col).lower()

    ]

    expectation_cols = [

        col for col in df.columns

        if "expectation" in str(col).lower()

    ]

    comparison_cols = [

        col for col in df.columns

        if "similar institution" in str(col).lower()

    ]

    qualitative_cols = [

        col for col in df.columns

        if any(

            keyword in str(col).lower()

            for keyword in [

                "suggest",

                "comment",

                "interest",

                "area",

                "training programs"

            ]

        )

    ]

    print("\nDetected Sections")

    print("--------------------------")

    print("Objectives:", len(objective_cols))

    print("Expectations:", len(expectation_cols))

    print("Institution Comparison:", len(comparison_cols))

    print("Qualitative:", len(qualitative_cols))

    # =========================================================
    # SAVE
    # =========================================================
    base_name = os.path.splitext(
        os.path.basename(file_path)
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

        ws = writer.sheets["Cleaned Data"]

        for cell in ws[1]:
            cell.font = Font(bold=True)

    print(f"\nCleaning complete: {output_file}")

    return output_file