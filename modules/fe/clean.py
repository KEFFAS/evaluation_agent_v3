import os
import pandas as pd
from openpyxl.styles import Font


# =========================================================
# CLEAN FACILITATOR EVALUATION DATA
# =========================================================

def clean_fe(
    file_name,
    output_folder="outputs"
):
    """
    Cleans raw Facilitator Evaluation data.

    The function:
    - Removes introductory rows
    - Detects and sets the header row
    - Removes empty rows and unnamed columns
    - Standardizes column names
    - Cleans text values
    - Preserves missing values
    - Converts rating columns to numeric values
    - Reorders columns
    - Saves a cleaned Excel workbook

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

    # =====================================================
    # CREATE OUTPUT FOLDER
    # =====================================================

    os.makedirs(
        output_folder,
        exist_ok=True
    )

    # =====================================================
    # LOAD RAW FILE
    # =====================================================

    df = pd.read_excel(
        file_name,
        header=None
    )

    print(
        "Original shape:",
        df.shape
    )

    # =====================================================
    # REMOVE FIRST ROW
    # =====================================================

    df = (
        df
        .iloc[1:]
        .reset_index(drop=True)
    )

    # =====================================================
    # SET HEADER
    # =====================================================

    df.columns = df.iloc[0]

    df = (
        df
        .iloc[1:]
        .reset_index(drop=True)
    )

    # =====================================================
    # REMOVE FULLY EMPTY ROWS
    # =====================================================

    df = df.dropna(
        how="all"
    )

    # =====================================================
    # REMOVE UNNAMED COLUMNS
    # =====================================================

    df = df.loc[
        :,
        ~df.columns
        .astype(str)
        .str.contains(
            "unnamed",
            case=False,
            na=False
        )
    ]

    # =====================================================
    # CLEAN COLUMN NAMES
    # =====================================================

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.title()
    )

    # =====================================================
    # STANDARDIZE COLUMN NAMES
    # =====================================================

    rename_map = {

        "Programme Title":
            "Program Title",

        "Program":
            "Program Title",

        "Facilitator":
            "Lecturer Name",

        "Lecturer":
            "Lecturer Name",

        "Topic":
            "Topic Description",

        "Session Topic":
            "Topic Description"
    }

    df = df.rename(
        columns=rename_map
    )

    # =====================================================
    # REMOVE DUPLICATE COLUMN NAMES
    # =====================================================

    df = df.loc[
        :,
        ~df.columns.duplicated()
    ]

    # =====================================================
    # CLEAN TEXT VALUES
    # =====================================================

    for col in df.columns:

        if df[col].dtype == "object":

            df[col] = (

                df[col]

                .apply(
                    lambda x:
                    str(x).strip()
                    if pd.notna(x)
                    else pd.NA
                )
            )

            # Replace empty strings with missing values

            df[col] = df[col].replace(

                "",

                pd.NA
            )

    # =====================================================
    # RATING COLUMNS
    # =====================================================

    rating_columns = [

        "Punctuality",

        "Presentation Flow",

        "Handling Questions",

        "Active Participation Of Learners",

        "Use Of Visual Aids",

        "Relevance Of Subject To Workplace",

        "Use Of Relevant Examples",

        "Knowledge Of Subject",

        "Treats Participants With Dignity And Respect",

        "Variety And Appropriateness Of Training Methods"
    ]

    # =====================================================
    # CONVERT RATINGS TO NUMERIC
    # =====================================================

    for col in rating_columns:

        if col in df.columns:

            df[col] = pd.to_numeric(

                df[col],

                errors="coerce"
            )

            # Keep only valid FE ratings

            df[col] = df[col].where(

                df[col].isin(

                    [1, 2, 3, 4, 5]
                )
            )

            # Convert to nullable integer

            df[col] = df[col].astype(
                "Int64"
            )

    # =====================================================
    # CONVERT OTHER MOSTLY NUMERIC COLUMNS
    # =====================================================

    for col in df.columns:

        if col in rating_columns:

            continue

        converted = pd.to_numeric(

            df[col],

            errors="coerce"
        )

        # Convert only if majority of values are numeric

        if converted.notna().sum() > (

            len(df) * 0.5
        ):

            df[col] = converted

    # =====================================================
    # REORDER COLUMNS
    # =====================================================

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

    # =====================================================
    # REMOVE INVALID RECORDS
    # =====================================================

    if "Lecturer Name" in df.columns:

        df = df.dropna(

            subset=[
                "Lecturer Name"
            ]
        )

    if "Topic Description" in df.columns:

        df = df.dropna(

            subset=[
                "Topic Description"
            ]
        )

    # =====================================================
    # SORT DATA
    # =====================================================

    sort_columns = [

        col

        for col in [

            "Lecturer Name",

            "Topic Description"

        ]

        if col in df.columns
    ]

    if sort_columns:

        df = (

            df

            .sort_values(

                sort_columns
            )

            .reset_index(
                drop=True
            )
        )

    print(
        "Cleaned shape:",
        df.shape
    )

    # =====================================================
    # OUTPUT FILE NAME
    # =====================================================

    base_name = os.path.splitext(

        os.path.basename(
            file_name
        )

    )[0]

    output_file = os.path.join(

        output_folder,

        f"{base_name}_cleaned.xlsx"
    )

    # =====================================================
    # SAVE CLEANED FILE
    # =====================================================

    with pd.ExcelWriter(

        output_file,

        engine="openpyxl"
    ) as writer:

        df.to_excel(

            writer,

            index=False,

            sheet_name="Cleaned Data"
        )

        worksheet = writer.sheets[
            "Cleaned Data"
        ]

        # Bold headers

        for cell in worksheet[1]:

            cell.font = Font(
                bold=True
            )

        # Adjust column widths

        for column_cells in worksheet.columns:

            max_length = 0

            column_letter = (

                column_cells[0]
                .column_letter
            )

            for cell in column_cells:

                try:

                    if cell.value:

                        max_length = max(

                            max_length,

                            len(
                                str(
                                    cell.value
                                )
                            )
                        )

                except Exception:

                    pass

            adjusted_width = min(

                max_length + 2,

                50
            )

            worksheet.column_dimensions[
                column_letter
            ].width = adjusted_width

    print(
        f"Cleaned file saved as: {output_file}"
    )

    return output_file