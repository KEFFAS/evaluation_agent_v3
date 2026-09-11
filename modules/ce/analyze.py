import os
import pandas as pd


def analyze_ce(
    cleaned_file,
    output_folder="outputs"
):
    """
    Analyze Coordinator Evaluation data.

    Generates:
    - Rating distribution percentages
    - Mean score for each indicator
    - Overall Mean Score
    - Composite Score (%)
    - Highest-rated indicator
    - Lowest-rated indicator
    - Qualitative feedback
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
    # CE RATING COLUMNS
    # =========================================================

    expected_rating_cols = [

        "Organization Of Program Opening And Closing",

        "Briefing Participants And Orientation",

        "Leveling Of Participant Expectations",

        "Communication And Feedback",

        "Management Of Timetable And Facilitators",

        "Monitoring Participants Attendance",

        "Program Evaluation",

        "Action Planning",

        "General Administration Of Program"

    ]

    rating_cols = [

        col

        for col in expected_rating_cols

        if col in df.columns

    ]

    if not rating_cols:

        raise ValueError(
            "No Coordinator Evaluation rating columns were found."
        )

    # =========================================================
    # ANALYSIS
    # =========================================================

    results = []

    indicator_means = {}

    for col in rating_cols:

        # Convert values safely to numeric
        ratings = pd.to_numeric(
            df[col],
            errors="coerce"
        )

        # Keep only valid ratings 1–5
        valid_ratings = ratings[
            ratings.isin([1, 2, 3, 4, 5])
        ]

        count5 = (valid_ratings == 5).sum()
        count4 = (valid_ratings == 4).sum()
        count3 = (valid_ratings == 3).sum()
        count2 = (valid_ratings == 2).sum()
        count1 = (valid_ratings == 1).sum()

        total_valid = len(valid_ratings)

        # =====================================================
        # PERCENTAGES
        # =====================================================

        if total_valid > 0:

            p5 = round(count5 / total_valid * 100, 1)
            p4 = round(count4 / total_valid * 100, 1)
            p3 = round(count3 / total_valid * 100, 1)
            p2 = round(count2 / total_valid * 100, 1)
            p1 = round(count1 / total_valid * 100, 1)

            mean_score = round(
                valid_ratings.mean(),
                2
            )

        else:

            p5 = p4 = p3 = p2 = p1 = 0
            mean_score = 0

        # Store mean for composite analysis
        indicator_means[col] = mean_score

        results.append([

            col,

            p5,

            p4,

            p3,

            p2,

            p1,

            mean_score

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

            "Poor % : 1",

            "Mean"

        ]

    )

    # =========================================================
    # OVERALL MEAN
    # =========================================================

    valid_means = [

        mean

        for mean in indicator_means.values()

        if mean > 0

    ]

    overall_mean = round(

        sum(valid_means) / len(valid_means),

        2

    ) if valid_means else 0

    # =========================================================
    # COMPOSITE SCORE
    # =========================================================

    composite_score = round(

        (overall_mean / 5) * 100,

        1

    ) if overall_mean > 0 else 0

    # =========================================================
    # HIGHEST AND LOWEST INDICATORS
    # =========================================================

    if valid_means:

        highest_indicator = max(
            indicator_means,
            key=indicator_means.get
        )

        lowest_indicator = min(
            indicator_means,
            key=indicator_means.get
        )

        highest_score = indicator_means[
            highest_indicator
        ]

        lowest_score = indicator_means[
            lowest_indicator
        ]

    else:

        highest_indicator = "N/A"
        lowest_indicator = "N/A"

        highest_score = 0
        lowest_score = 0

    # =========================================================
    # QUALITATIVE FEEDBACK
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

        # =====================================================
        # HEADER INFORMATION
        # =====================================================

        header_df = pd.DataFrame({

            "A": [

                f"Program Title: {programme_title}",

                f"Coordinator Name: {coordinator}",

                f"Overall Mean Score: {overall_mean}",

                f"Composite Score: {composite_score}%",

                f"Highest Rated Aspect: {highest_indicator} ({highest_score})",

                f"Lowest Rated Aspect: {lowest_indicator} ({lowest_score})"

            ]

        })

        header_df.to_excel(

            writer,

            sheet_name="Analysis",

            index=False,

            header=False,

            startrow=0

        )

        # =====================================================
        # ANALYSIS TABLE
        # =====================================================

        table_start = 8

        df_out.to_excel(

            writer,

            sheet_name="Analysis",

            index=False,

            startrow=table_start

        )

        worksheet = writer.sheets["Analysis"]

        # =====================================================
        # ADD OVERALL MEAN ROW
        # =====================================================

        mean_row = table_start + len(df_out) + 2

        worksheet.cell(
            row=mean_row,
            column=1
        ).value = "OVERALL MEAN"

        worksheet.cell(
            row=mean_row,
            column=7
        ).value = overall_mean

        # =====================================================
        # ADD COMPOSITE SCORE ROW
        # =====================================================

        composite_row = mean_row + 1

        worksheet.cell(
            row=composite_row,
            column=1
        ).value = "COMPOSITE SCORE (%)"

        worksheet.cell(
            row=composite_row,
            column=7
        ).value = composite_score

        # =====================================================
        # QUALITATIVE SECTION
        # =====================================================

        qual_start = composite_row + 3

        worksheet.cell(

            row=qual_start,

            column=1

        ).value = "MOST LIKED:"

        worksheet.cell(

            row=qual_start + 1,

            column=1

        ).value = likes

        worksheet.cell(

            row=qual_start + 3,

            column=1

        ).value = "SUGGESTIONS:"

        worksheet.cell(

            row=qual_start + 4,

            column=1

        ).value = suggestions

    print(f"Overall Mean: {overall_mean}")
    print(f"Composite Score: {composite_score}%")
    print(f"Highest Aspect: {highest_indicator}")
    print(f"Lowest Aspect: {lowest_indicator}")

    print(f"✅ Analysis saved: {output_file}")

    return output_file