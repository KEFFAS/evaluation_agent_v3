import os
import pandas as pd


def analyze_ce(
    cleaned_file,
    output_folder="outputs"
):
    """
    Analyze Coordinator Evaluation data.

    Parameters
    ----------
    cleaned_file : str
        Path to cleaned Excel file.

    output_folder : str

    Returns
    -------
    str
        Path to analysis workbook.
    """

    os.makedirs(output_folder, exist_ok=True)

    # =========================================================
    # LOAD CLEANED DATA
    # =========================================================
    df = pd.read_excel(cleaned_file)

    print("Loaded file:", cleaned_file)
    print("Shape:", df.shape)

    # =========================================================
    # PROGRAM DETAILS
    # =========================================================
    programme_title = (
        df["Program Title"].iloc[0]
        if "Program Title" in df.columns
        else "N/A"
    )

    coordinator = (
        df["Coordinator Name"].iloc[0]
        if "Coordinator Name" in df.columns
        else "N/A"
    )

    # =========================================================
    # IDENTIFY RATING COLUMNS
    # =========================================================
    rating_cols = [

        col

        for col in df.columns

        if df[col].dtype in [

            "int64",

            "float64",

            "Int64"

        ]

    ]

    exclude_cols = [

        "Timetable No"

    ]

    rating_cols = [

        col

        for col in rating_cols

        if col not in exclude_cols

    ]

    # =========================================================
    # ANALYSIS
    # =========================================================
    results = []

    for col in rating_cols:

        counts = df[col].value_counts().to_dict()

        count5 = counts.get(5, 0)
        count4 = counts.get(4, 0)
        count3 = counts.get(3, 0)
        count2 = counts.get(2, 0)
        count1 = counts.get(1, 0)

        total = (

            count5 +

            count4 +

            count3 +

            count2 +

            count1

        )

        if total > 0:

            p5 = round(count5 / total * 100, 1)
            p4 = round(count4 / total * 100, 1)
            p3 = round(count3 / total * 100, 1)
            p2 = round(count2 / total * 100, 1)
            p1 = round(count1 / total * 100, 1)

        else:

            p5 = p4 = p3 = p2 = p1 = 0

        results.append([

            col,

            p5,

            p4,

            p3,

            p2,

            p1

        ])

    # =========================================================
    # RESULTS TABLE
    # =========================================================
    df_out = pd.DataFrame(

        results,

        columns=[

            "Specific Aspects",

            "Excellent % : 5",

            "Very Good % : 4",

            "Good % : 3",

            "Fair % : 2",

            "Poor % : 1"

        ]

    )

    # =========================================================
    # QUALITATIVE
    # =========================================================
    likes = (

        "; ".join(

            df["Like"]

            .dropna()

            .astype(str)

        )

        if "Like" in df.columns

        else "No responses"

    )

    suggestions = (

        "; ".join(

            df["Suggestions"]

            .dropna()

            .astype(str)

        )

        if "Suggestions" in df.columns

        else "No responses"

    )

    # =========================================================
    # SAVE
    # =========================================================
    base_name = os.path.splitext(

        os.path.basename(cleaned_file)

    )[0]

    output_file = os.path.join(

        output_folder,

        f"{base_name}_analysis.xlsx"

    )

    with pd.ExcelWriter(

        output_file,

        engine="openpyxl"

    ) as writer:

        start_row = 0

        header_df = pd.DataFrame({

            "A": [

                f"Program Title: {programme_title}",

                f"Coordinator Name: {coordinator}"

            ]

        })

        header_df.to_excel(

            writer,

            sheet_name="Analysis",

            index=False,

            header=False,

            startrow=start_row

        )

        table_start = start_row + 4

        df_out.to_excel(

            writer,

            sheet_name="Analysis",

            index=False,

            startrow=table_start

        )

        worksheet = writer.sheets["Analysis"]

        qual_start = table_start + len(df_out) + 3

        worksheet.cell(

            row=qual_start,

            column=1

        ).value = "Most Liked:"

        worksheet.cell(

            row=qual_start + 1,

            column=1

        ).value = likes

        worksheet.cell(

            row=qual_start + 3,

            column=1

        ).value = "Suggestions:"

        worksheet.cell(

            row=qual_start + 4,

            column=1

        ).value = suggestions

    print(f"✅ Analysis saved: {output_file}")

    return output_file