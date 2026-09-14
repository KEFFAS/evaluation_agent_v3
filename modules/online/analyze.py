import os
import pandas as pd
from openpyxl.styles import Font, Alignment


def calculate_rating_stats(series):
    """
    Calculate rating counts, mean score and composite score.
    """

    numeric = pd.to_numeric(
        series,
        errors="coerce"
    )

    numeric = numeric[
        numeric.isin([1, 2, 3, 4, 5])
    ]

    counts = {

        5: int((numeric == 5).sum()),
        4: int((numeric == 4).sum()),
        3: int((numeric == 3).sum()),
        2: int((numeric == 2).sum()),
        1: int((numeric == 1).sum())

    }

    total = len(numeric)

    if total > 0:

        mean_score = round(
            numeric.mean(),
            2
        )

        composite_score = round(
            (mean_score / 5) * 100,
            1
        )

    else:

        mean_score = 0
        composite_score = 0

    return {

        "5": counts[5],
        "4": counts[4],
        "3": counts[3],
        "2": counts[2],
        "1": counts[1],
        "Total": total,
        "Mean": mean_score,
        "Composite Score (%)": composite_score

    }


def analyze_online(
    cleaned_file,
    output_folder="outputs"
):
    """
    Analyze Online End-of-Event Evaluation data.

    Creates an Excel workbook containing:

    1. Overall Evaluation Summary
    2. Course Objectives Analysis
    3. Personal Expectations Analysis
    4. Specific Aspects Analysis
    5. Future Attendance Analysis
    6. Recommendation Analysis
    7. Qualitative Feedback
    """

    os.makedirs(
        output_folder,
        exist_ok=True
    )

    # =========================================================
    # LOAD CLEANED DATA
    # =========================================================

    df = pd.read_excel(
        cleaned_file
    )

    print("=" * 60)
    print("ONLINE EVALUATION ANALYSIS")
    print("=" * 60)

    print(
        f"Rows: {df.shape[0]}"
    )

    print(
        f"Columns: {df.shape[1]}"
    )

    columns = df.columns.tolist()

    # =========================================================
    # IDENTIFY MAIN RATING COLUMNS
    # =========================================================

    objective_col = next(

        (

            col

            for col in columns

            if "objective"
            in str(col).lower()

        ),

        None

    )

    expectation_col = next(

        (

            col

            for col in columns

            if "expectation"
            in str(col).lower()

        ),

        None

    )

    future_col = next(

        (

            col

            for col in columns

            if (

                "future online training"
                in str(col).lower()

                or

                "attend another"
                in str(col).lower()

                or

                "future training"
                in str(col).lower()

            )

        ),

        None

    )

    recommend_col = next(

        (

            col

            for col in columns

            if (

                "recommend"
                in str(col).lower()

            )

        ),

        None

    )

    print("\nDetected Main Columns:")

    print(
        "Objectives:",
        objective_col
    )

    print(
        "Expectations:",
        expectation_col
    )

    print(
        "Future Attendance:",
        future_col
    )

    print(
        "Recommendation:",
        recommend_col
    )

    # =========================================================
    # IDENTIFY QUALITATIVE COLUMNS
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

    qualitative_cols = [

        col

        for col in columns

        if any(

            keyword in str(col).lower()

            for keyword
            in qualitative_keywords

        )

    ]

    # =========================================================
    # IDENTIFY RATING COLUMNS
    # =========================================================

    main_cols = [

        objective_col,
        expectation_col,
        future_col,
        recommend_col

    ]

    rating_cols = []

    for col in columns:

        if col in main_cols:

            continue

        if col in qualitative_cols:

            continue

        numeric = pd.to_numeric(
            df[col],
            errors="coerce"
        )

        valid = numeric.dropna()

        if len(valid) > 0:

            if valid.isin(
                [1, 2, 3, 4, 5]
            ).all():

                rating_cols.append(
                    col
                )

    print("\nSpecific Rating Columns:")

    for col in rating_cols:

        print(
            "-",
            col
        )

    # =========================================================
    # ANALYZE MAIN SECTIONS
    # =========================================================

    section_results = []

    detailed_results = {}

    # ---------------------------------------------------------
    # COURSE OBJECTIVES
    # ---------------------------------------------------------

    if objective_col:

        stats = calculate_rating_stats(
            df[objective_col]
        )

        detailed_results[
            "Course Objectives Achievement"
        ] = pd.DataFrame([

            {

                "Specific Aspect":
                "Course Objectives Achievement",

                **stats

            }

        ])

        section_results.append({

            "Evaluation Area":
            "Course Objectives Achievement",

            "Mean Score":
            stats["Mean"],

            "Composite Score (%)":
            stats["Composite Score (%)"]

        })

    # ---------------------------------------------------------
    # PERSONAL EXPECTATIONS
    # ---------------------------------------------------------

    if expectation_col:

        stats = calculate_rating_stats(
            df[expectation_col]
        )

        detailed_results[
            "Fulfilment of Personal Expectations"
        ] = pd.DataFrame([

            {

                "Specific Aspect":
                "Fulfilment of Personal Expectations",

                **stats

            }

        ])

        section_results.append({

            "Evaluation Area":
            "Fulfilment of Personal Expectations",

            "Mean Score":
            stats["Mean"],

            "Composite Score (%)":
            stats["Composite Score (%)"]

        })

    # =========================================================
    # SPECIFIC ASPECTS
    # =========================================================

    aspect_results = []

    for col in rating_cols:

        stats = calculate_rating_stats(
            df[col]
        )

        aspect_results.append({

            "Specific Aspect":
            col,

            **stats

        })

    aspects_df = pd.DataFrame(
        aspect_results
    )

    if not aspects_df.empty:

        specific_mean = round(

            aspects_df["Mean"].mean(),

            2

        )

        specific_composite = round(

            (specific_mean / 5) * 100,

            1

        )

        section_results.append({

            "Evaluation Area":
            "Specific Aspects of Online Training",

            "Mean Score":
            specific_mean,

            "Composite Score (%)":
            specific_composite

        })

        detailed_results[
            "Specific Aspects"
        ] = aspects_df

    # =========================================================
    # FUTURE ATTENDANCE
    # =========================================================

    if future_col:

        stats = calculate_rating_stats(
            df[future_col]
        )

        detailed_results[
            "Future Attendance"
        ] = pd.DataFrame([

            {

                "Specific Aspect":
                "Likelihood of Attending Future Online Training",

                **stats

            }

        ])

        section_results.append({

            "Evaluation Area":
            "Likelihood of Attending Future Online Training",

            "Mean Score":
            stats["Mean"],

            "Composite Score (%)":
            stats["Composite Score (%)"]

        })

    # =========================================================
    # RECOMMEND KSG
    # =========================================================

    if recommend_col:

        stats = calculate_rating_stats(
            df[recommend_col]
        )

        detailed_results[
            "Recommend KSG"
        ] = pd.DataFrame([

            {

                "Specific Aspect":
                "Willingness to Recommend KSG Online Learning",

                **stats

            }

        ])

        section_results.append({

            "Evaluation Area":
            "Willingness to Recommend KSG Online Learning",

            "Mean Score":
            stats["Mean"],

            "Composite Score (%)":
            stats["Composite Score (%)"]

        })

    # =========================================================
    # OVERALL SUMMARY
    # =========================================================

    summary_df = pd.DataFrame(
        section_results
    )

    if not summary_df.empty:

        overall_mean = round(

            summary_df[
                "Mean Score"
            ].mean(),

            2

        )

        overall_composite = round(

            (overall_mean / 5) * 100,

            1

        )

        overall_row = pd.DataFrame([

            {

                "Evaluation Area":
                "OVERALL EVALUATION SCORE",

                "Mean Score":
                overall_mean,

                "Composite Score (%)":
                overall_composite

            }

        ])

        summary_df = pd.concat(

            [

                summary_df,

                overall_row

            ],

            ignore_index=True

        )

    else:

        overall_mean = 0
        overall_composite = 0

    # =========================================================
    # QUALITATIVE DATA
    # =========================================================

    qualitative_results = []

    for col in qualitative_cols:

        responses = (

            df[col]

            .dropna()

            .astype(str)

            .str.strip()

        )

        responses = responses[

            responses != ""

        ]

        for response in responses:

            qualitative_results.append({

                "Question":
                col,

                "Participant Response":
                response

            })

    qualitative_df = pd.DataFrame(
        qualitative_results
    )

    # =========================================================
    # OUTPUT FILE
    # =========================================================

    base_name = os.path.splitext(

        os.path.basename(
            cleaned_file
        )

    )[0]

    output_file = os.path.join(

        output_folder,

        f"{base_name}_analysis.xlsx"

    )

    # =========================================================
    # SAVE ANALYSIS WORKBOOK
    # =========================================================

    with pd.ExcelWriter(

        output_file,

        engine="openpyxl"

    ) as writer:

        # -----------------------------------------------------
        # OVERALL SUMMARY
        # -----------------------------------------------------

        summary_df.to_excel(

            writer,

            sheet_name="Overall Summary",

            index=False

        )

        # -----------------------------------------------------
        # DETAILED SECTIONS
        # -----------------------------------------------------

        for sheet_name, result_df in detailed_results.items():

            safe_sheet_name = (

                sheet_name[:31]

            )

            result_df.to_excel(

                writer,

                sheet_name=safe_sheet_name,

                index=False

            )

        # -----------------------------------------------------
        # QUALITATIVE FEEDBACK
        # -----------------------------------------------------

        if not qualitative_df.empty:

            qualitative_df.to_excel(

                writer,

                sheet_name="Qualitative Feedback",

                index=False

            )

        # =====================================================
        # FORMAT WORKSHEETS
        # =====================================================

        for worksheet in writer.book.worksheets:

            # Bold header
            for cell in worksheet[1]:

                cell.font = Font(
                    bold=True
                )

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            # Adjust width
            for column_cells in worksheet.columns:

                max_length = 0

                column_letter = (
                    column_cells[0]
                    .column_letter
                )

                for cell in column_cells:

                    try:

                        value_length = len(
                            str(cell.value)
                        )

                        if value_length > max_length:

                            max_length = value_length

                    except Exception:

                        pass

                adjusted_width = min(

                    max_length + 2,

                    50

                )

                worksheet.column_dimensions[
                    column_letter
                ].width = adjusted_width

    # =========================================================
    # PRINT SUMMARY
    # =========================================================

    print("\n" + "=" * 60)

    print(
        "ONLINE ANALYSIS COMPLETED"
    )

    print("=" * 60)

    print(
        f"Overall Mean Score: "
        f"{overall_mean}"
    )

    print(
        f"Overall Composite Score: "
        f"{overall_composite}%"
    )

    print(
        f"Analysis saved to: "
        f"{output_file}"
    )

    return output_file