import os
import pandas as pd
from openpyxl.styles import Font


def calculate_rating_stats(series):
    """
    Calculate rating counts, total respondents,
    mean score and composite score.
    """

    numeric = pd.to_numeric(series, errors="coerce")

    counts = {
        score: int((numeric == score).sum())
        for score in [5, 4, 3, 2, 1]
    }

    total_respondents = sum(counts.values())

    if total_respondents > 0:

        mean_score = (
            (5 * counts[5]) +
            (4 * counts[4]) +
            (3 * counts[3]) +
            (2 * counts[2]) +
            (1 * counts[1])
        ) / total_respondents

        composite_score = (
            mean_score / 5
        ) * 100

    else:

        mean_score = 0
        composite_score = 0

    return {
        "count_5": counts[5],
        "count_4": counts[4],
        "count_3": counts[3],
        "count_2": counts[2],
        "count_1": counts[1],
        "total_respondents": total_respondents,
        "mean_score": round(mean_score, 2),
        "composite_score": round(composite_score, 1)
    }


def analyze_eee(cleaned_file, output_folder="outputs"):
    """
    Analyze a cleaned End of Event Evaluation dataset.

    Produces:
    1. Course Objectives Achievement analysis
    2. Personal Expectations analysis
    3. Specific Programme Aspects analysis
    4. KSG Comparison analysis
    5. Overall Evaluation Summary
    6. Qualitative Responses

    Parameters
    ----------
    cleaned_file : str
        Path to cleaned Excel file.

    output_folder : str
        Folder to save analysis workbook.

    Returns
    -------
    str
        Path to analysis workbook.
    """

    os.makedirs(output_folder, exist_ok=True)

    # =====================================================
    # LOAD FILE
    # =====================================================
    df = pd.read_excel(cleaned_file)

    print(f"Loaded: {cleaned_file}")
    print(f"Shape: {df.shape}")

    # =====================================================
    # BASIC DETAILS
    # =====================================================
    program_title = (
        df["Program Title"].iloc[0]
        if "Program Title" in df.columns
        else "N/A"
    )

    coordinator = (
        df["Coordinator Name"].iloc[0]
        if "Coordinator Name" in df.columns
        else "N/A"
    )

    program_code = (
        df["Program Code"].iloc[0]
        if "Program Code" in df.columns
        else "N/A"
    )

    venue = (
        df["Venue / Campus"].iloc[0]
        if "Venue / Campus" in df.columns
        else "N/A"
    )

    assistant = (
        df["Program Assistant Name"].iloc[0]
        if "Program Assistant Name" in df.columns
        else "N/A"
    )

    # =====================================================
    # DETECT NUMERIC RATING COLUMNS
    # =====================================================
    rating_cols = []

    for col in df.columns:

        numeric = pd.to_numeric(
            df[col],
            errors="coerce"
        )

        valid_ratings = numeric.dropna()

        if len(valid_ratings) > 0:

            if valid_ratings.isin(
                [1, 2, 3, 4, 5]
            ).all():

                rating_cols.append(col)

    # =====================================================
    # EXCLUDE NON-RATING NUMERIC COLUMNS
    # =====================================================
    exclude = [
        "Timetable No"
    ]

    rating_cols = [
        col
        for col in rating_cols
        if col not in exclude
    ]

    print("\nDetected Rating Columns:")

    for col in rating_cols:

        print(f" - {col}")

    # =====================================================
    # DETECT SPECIAL SECTIONS
    # =====================================================
    objective_col = next(
        (
            col
            for col in rating_cols
            if "objective" in str(col).lower()
        ),
        None
    )

    expectation_col = next(
        (
            col
            for col in rating_cols
            if "expectation" in str(col).lower()
        ),
        None
    )

    comparison_col = next(
        (
            col
            for col in rating_cols
            if (
                "similar institution" in str(col).lower()
                or "similar institutions" in str(col).lower()
            )
        ),
        None
    )

    # =====================================================
    # SECTION 1
    # COURSE OBJECTIVES ACHIEVEMENT
    # =====================================================
    section1_df = pd.DataFrame()

    section1_stats = None

    if objective_col:

        section1_stats = calculate_rating_stats(
            df[objective_col]
        )

        section1_df = pd.DataFrame(
            [[
                "Achievement of Course Objectives",
                section1_stats["count_5"],
                section1_stats["count_4"],
                section1_stats["count_3"],
                section1_stats["count_2"],
                section1_stats["count_1"],
                section1_stats["mean_score"]
            ]],
            columns=[
                "COURSE OBJECTIVES ACHIEVEMENT",
                "5",
                "4",
                "3",
                "2",
                "1",
                "MEAN"
            ]
        )

        section1_df.loc[
            len(section1_df)
        ] = [
            "COMPOSITE SCORE (%)",
            "",
            "",
            "",
            "",
            "",
            section1_stats["composite_score"]
        ]

    # =====================================================
    # SECTION 2
    # FULFILMENT OF PERSONAL EXPECTATIONS
    # =====================================================
    section2_df = pd.DataFrame()

    section2_stats = None

    if expectation_col:

        section2_stats = calculate_rating_stats(
            df[expectation_col]
        )

        section2_df = pd.DataFrame(
            [[
                "Fulfilment of Personal Expectations",
                section2_stats["count_5"],
                section2_stats["count_4"],
                section2_stats["count_3"],
                section2_stats["count_2"],
                section2_stats["count_1"],
                section2_stats["mean_score"]
            ]],
            columns=[
                "FULFILMENT OF EXPECTATIONS",
                "5",
                "4",
                "3",
                "2",
                "1",
                "MEAN"
            ]
        )

        section2_df.loc[
            len(section2_df)
        ] = [
            "COMPOSITE SCORE (%)",
            "",
            "",
            "",
            "",
            "",
            section2_stats["composite_score"]
        ]

    # =====================================================
    # SECTION 3
    # SPECIFIC ASPECTS OF THE TRAINING PROGRAMME
    # =====================================================
    specific_aspects = []

    specific_aspect_stats = []

    special_cols = [
        objective_col,
        expectation_col,
        comparison_col
    ]

    for col in rating_cols:

        if col not in special_cols:

            stats = calculate_rating_stats(
                df[col]
            )

            specific_aspects.append([
                col,
                stats["count_5"],
                stats["count_4"],
                stats["count_3"],
                stats["count_2"],
                stats["count_1"],
                stats["mean_score"]
            ])

            specific_aspect_stats.append(
                stats
            )

    section3_columns = [
        "SPECIFIC ASPECTS",
        "5",
        "4",
        "3",
        "2",
        "1",
        "MEAN"
    ]

    section3_df = pd.DataFrame(
        specific_aspects,
        columns=section3_columns
    )

    section3_overall_mean = 0
    section3_composite = 0

    if specific_aspect_stats:

        section3_overall_mean = round(
            sum(
                item["mean_score"]
                for item in specific_aspect_stats
            )
            / len(specific_aspect_stats),
            2
        )

        section3_composite = round(
            (
                section3_overall_mean
                / 5
            ) * 100,
            1
        )

        section3_df.loc[
            len(section3_df)
        ] = [
            "OVERALL MEAN",
            "",
            "",
            "",
            "",
            "",
            section3_overall_mean
        ]

        section3_df.loc[
            len(section3_df)
        ] = [
            "COMPOSITE SCORE (%)",
            "",
            "",
            "",
            "",
            "",
            section3_composite
        ]

    # =====================================================
    # SECTION 4
    # KSG COMPARED TO SIMILAR INSTITUTIONS
    # =====================================================
    section4_df = pd.DataFrame()

    section4_stats = None

    if comparison_col:

        section4_stats = calculate_rating_stats(
            df[comparison_col]
        )

        section4_df = pd.DataFrame(
            [[
                "Rating of KSG Training",
                section4_stats["count_5"],
                section4_stats["count_4"],
                section4_stats["count_3"],
                section4_stats["count_2"],
                section4_stats["count_1"],
                section4_stats["mean_score"]
            ]],
            columns=[
                "KSG COMPARED TO SIMILAR INSTITUTIONS",
                "5",
                "4",
                "3",
                "2",
                "1",
                "MEAN"
            ]
        )

        section4_df.loc[
            len(section4_df)
        ] = [
            "COMPOSITE SCORE (%)",
            "",
            "",
            "",
            "",
            "",
            section4_stats["composite_score"]
        ]

    # =====================================================
    # OVERALL EVALUATION SUMMARY
    # =====================================================
    summary_data = []

    if section1_stats:

        summary_data.append([
            "Course Objectives Achievement",
            section1_stats["mean_score"],
            section1_stats["composite_score"]
        ])

    if section2_stats:

        summary_data.append([
            "Fulfilment of Personal Expectations",
            section2_stats["mean_score"],
            section2_stats["composite_score"]
        ])

    if specific_aspect_stats:

        summary_data.append([
            "Specific Programme Aspects",
            section3_overall_mean,
            section3_composite
        ])

    if section4_stats:

        summary_data.append([
            "KSG Compared to Similar Institutions",
            section4_stats["mean_score"],
            section4_stats["composite_score"]
        ])

    summary_df = pd.DataFrame(
        summary_data,
        columns=[
            "EVALUATION AREA",
            "MEAN SCORE",
            "COMPOSITE SCORE (%)"
        ]
    )

    if not summary_df.empty:

        overall_mean = round(
            summary_df["MEAN SCORE"].mean(),
            2
        )

        overall_composite = round(
            (
                overall_mean / 5
            ) * 100,
            1
        )

        summary_df.loc[
            len(summary_df)
        ] = [
            "OVERALL EVALUATION SCORE",
            overall_mean,
            overall_composite
        ]

    # =====================================================
    # QUALITATIVE MAPPING
    # =====================================================
    qualitative_mapping = {

        "Suggestions": None,

        "Areas to Add": None,

        "Interest in Other KSG Programmes": None,

        "Additional Training Areas": None,

        "General Comments": None

    }

    for col in df.columns:

        col_lower = str(col).lower()

        if "suggestions on aspects" in col_lower:

            qualitative_mapping[
                "Suggestions"
            ] = col

        elif "other areas you would like added" in col_lower:

            qualitative_mapping[
                "Areas to Add"
            ] = col

        elif "other ksg training programs" in col_lower:

            qualitative_mapping[
                "Interest in Other KSG Programmes"
            ] = col

        elif (
            "other training programs not currently offered"
            in col_lower
        ):

            qualitative_mapping[
                "Additional Training Areas"
            ] = col

        elif "other comments" in col_lower:

            qualitative_mapping[
                "General Comments"
            ] = col

    # =====================================================
    # QUALITATIVE RESPONSES
    # =====================================================
    qualitative_outputs = {}

    for section, column in qualitative_mapping.items():

        if column and column in df.columns:

            responses = (
                df[column]
                .dropna()
                .astype(str)
                .str.strip()
            )

            responses = responses[
                responses != ""
            ]

            qualitative_outputs[
                section
            ] = " ".join(responses)

        else:

            qualitative_outputs[
                section
            ] = ""

    # =====================================================
    # OUTPUT FILE
    # =====================================================
    base_name = os.path.splitext(
        os.path.basename(cleaned_file)
    )[0]

    output_file = os.path.join(
        output_folder,
        f"{base_name}_analysis.xlsx"
    )

    # =====================================================
    # SAVE ANALYSIS
    # =====================================================
    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        # =================================================
        # PROGRAM DETAILS
        # =================================================
        details_df = pd.DataFrame({

            "Field": [

                "Program Title",

                "Coordinator",

                "Program Code",

                "Venue",

                "Program Assistant"

            ],

            "Value": [

                program_title,

                coordinator,

                program_code,

                venue,

                assistant

            ]

        })

        details_df.to_excel(
            writer,
            sheet_name="Program Details",
            index=False
        )

        # =================================================
        # OVERALL SUMMARY
        # =================================================
        summary_df.to_excel(
            writer,
            sheet_name="Overall Summary",
            index=False
        )

        # =================================================
        # SECTION 1
        # =================================================
        section1_df.to_excel(
            writer,
            sheet_name="Section 1 Objectives",
            index=False
        )

        # =================================================
        # SECTION 2
        # =================================================
        section2_df.to_excel(
            writer,
            sheet_name="Section 2 Expectations",
            index=False
        )

        # =================================================
        # SECTION 3
        # =================================================
        section3_df.to_excel(
            writer,
            sheet_name="Section 3 Specific Aspects",
            index=False
        )

        # =================================================
        # SECTION 4
        # =================================================
        section4_df.to_excel(
            writer,
            sheet_name="Section 4 Comparison",
            index=False
        )

        # =================================================
        # QUALITATIVE SHEET
        # =================================================
        qualitative_sheet = []

        for section, text in qualitative_outputs.items():

            qualitative_sheet.append([
                section,
                text
            ])

        qualitative_df = pd.DataFrame(
            qualitative_sheet,
            columns=[
                "Section",
                "Responses"
            ]
        )

        qualitative_df.to_excel(
            writer,
            sheet_name="Qualitative Responses",
            index=False
        )

        # =================================================
        # FORMAT HEADERS
        # =================================================
        for sheet in writer.sheets.values():

            for cell in sheet[1]:

                cell.font = Font(bold=True)

    print(
        f"\nAnalysis completed: {output_file}"
    )

    return output_file