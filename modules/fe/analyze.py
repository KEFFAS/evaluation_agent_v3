import os
import pandas as pd
from openpyxl.styles import Font


def analyze_fe(
    cleaned_file,
    total_participants,
    output_folder="outputs"
):
    """
    Analyze Facilitator Evaluation data.

    Parameters
    ----------
    cleaned_file : str
        Path to cleaned Excel file.

    total_participants : int
        Total participants in the programme.

    output_folder : str

    Returns
    -------
    str
        Path to analyzed workbook.
    """

    os.makedirs(output_folder, exist_ok=True)

    # =========================================================
    # LOAD CLEANED DATA
    # =========================================================
    df = pd.read_excel(cleaned_file)

    # =========================================================
    # STANDARDIZE COLUMN NAMES
    # =========================================================
    df.columns = (

        df.columns

        .astype(str)

        .str.strip()

        .str.title()

    )

    rename_map = {

        "Topic": "Topic Description",

        "Session Topic": "Topic Description",

        "Facilitator": "Lecturer Name",

        "Lecturer": "Lecturer Name"

    }

    df = df.rename(columns=rename_map)

    # =========================================================
    # VALIDATION
    # =========================================================
    required = [

        "Topic Description",

        "Lecturer Name"

    ]

    missing = [

        col

        for col in required

        if col not in df.columns

    ]

    if missing:

        raise ValueError(

            f"Missing required columns: {missing}"

        )

    # =========================================================
    # RATING COLUMNS
    # =========================================================
    rating_cols = [

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

    rating_cols = [

        col

        for col in rating_cols

        if col in df.columns

    ]

    # =========================================================
    # GROUP BY SESSION
    # =========================================================
    grouped = df.groupby("Topic Description")

    # =========================================================
    # OUTPUT FILE
    # =========================================================
    base_name = os.path.splitext(

        os.path.basename(cleaned_file)

    )[0]

    output_file = os.path.join(

        output_folder,

        f"{base_name}_analysis.xlsx"

    )

    # =========================================================
    # WRITE ANALYSIS
    # =========================================================
    with pd.ExcelWriter(

        output_file,

        engine="openpyxl"

    ) as writer:

        for session, group in grouped:

            facilitator = group["Lecturer Name"].iloc[0]

            results = []

            for col in rating_cols:

                counts = group[col].value_counts().to_dict()

                count_5 = counts.get(5, 0)
                count_4 = counts.get(4, 0)
                count_3 = counts.get(3, 0)
                count_2 = counts.get(2, 0)
                count_1 = counts.get(1, 0)

                total_valid = (
                    count_5 +
                    count_4 +
                    count_3 +
                    count_2 +
                    count_1
                )

                non_response = (
                    total_participants -
                    total_valid
                )

                pct45 = (

                    ((count_5 + count_4) / total_valid * 100)

                    if total_valid > 0

                    else 0

                )

                results.append([

                    col,

                    total_participants,

                    non_response,

                    count_5,

                    count_4,

                    count_3,

                    count_2,

                    count_1,

                    total_valid,

                    round(pct45, 1)

                ])

            result_df = pd.DataFrame(

                results,

                columns=[

                    "Indicator",

                    "Total Participants",

                    "Non Response",

                    "5",

                    "4",

                    "3",

                    "2",

                    "1",

                    "Total Valid Responses",

                    "% Scores 4 & 5"

                ]

            )

            likes = "; ".join(

                group["Like"]

                .dropna()

                .astype(str)

            )

            suggestions = "; ".join(

                group["Suggestions"]

                .dropna()

                .astype(str)

            )

            sheet_name = str(session)[:31]

            result_df.to_excel(

                writer,

                sheet_name=sheet_name,

                startrow=3,

                index=False

            )

            ws = writer.sheets[sheet_name]

            ws["A1"] = f"SESSION: {session}"
            ws["A2"] = f"FACILITATOR: {facilitator}"

            ws["A1"].font = Font(bold=True)
            ws["A2"].font = Font(bold=True)

            last = len(result_df) + 5

            ws[f"A{last}"] = "MOST LIKED:"
            ws[f"A{last+1}"] = likes

            ws[f"A{last+3}"] = "SUGGESTIONS:"
            ws[f"A{last+4}"] = suggestions

            ws[f"A{last}"].font = Font(bold=True)
            ws[f"A{last+3}"].font = Font(bold=True)

    print(f"✅ Analysis saved: {output_file}")

    return output_file