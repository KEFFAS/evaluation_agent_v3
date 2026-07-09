import os
import pandas as pd
from openpyxl.styles import Font


def analyze_eee(cleaned_file, output_folder="outputs"):
    """
    Analyze a cleaned End of Event Evaluation dataset.

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
    rating_cols = [

        col for col in df.columns

        if df[col].dtype in [

            "int64",
            "float64",
            "Int64"

        ]

    ]

    exclude = [

        "Timetable No"

    ]

    rating_cols = [

        col for col in rating_cols

        if col not in exclude

    ]

    # =====================================================
    # DETECT SECTION COLUMNS
    # =====================================================
    objective_col = next(

        (

            col for col in rating_cols

            if "objective" in col.lower()

        ),

        None

    )

    expectation_col = next(

        (

            col for col in rating_cols

            if "expectation" in col.lower()

        ),

        None

    )

    comparison_col = next(

        (

            col for col in rating_cols

            if "similar institution" in col.lower()

        ),

        None

    )

    # =====================================================
    # SECTION 1
    # =====================================================
    section1 = []

    if objective_col:

        counts = df[objective_col].value_counts().to_dict()

        total = sum(counts.values())

        labels = {

            5: "Excellent",

            4: "Very Good",

            3: "Satisfactory",

            2: "Poor",

            1: "Very Poor"

        }

        for score in [5,4,3,2,1]:

            pct = (

                round(

                    (counts.get(score,0)/total)*100,

                    1

                )

                if total > 0

                else 0

            )

            section1.append([

                labels[score],

                pct

            ])

    section1_df = pd.DataFrame(

        section1,

        columns=[

            "Rating",

            "Percentage of Respondents"

        ]

    )

    # =====================================================
    # SECTION 2
    # =====================================================
    section2 = []

    if expectation_col:

        counts = df[expectation_col].value_counts().to_dict()

        total = sum(counts.values())

        labels = {

            5:"5 - Great Extent",

            4:"4 - Some Extent",

            3:"3 - Satisfactory",

            2:"2 - Not Sure",

            1:"1 - Not at All"

        }

        for score in [5,4,3,2,1]:

            pct = (

                round(

                    (counts.get(score,0)/total)*100,

                    1

                )

                if total > 0

                else 0

            )

            section2.append([

                labels[score],

                pct

            ])

    section2_df = pd.DataFrame(

        section2,

        columns=[

            "Rating",

            "Percentage of Respondents"

        ]

    )

    # =====================================================
    # SECTION 3
    # =====================================================
    specific_aspects = []

    for col in rating_cols:

        if col not in [

            objective_col,

            expectation_col,

            comparison_col

        ]:

            counts = df[col].value_counts().to_dict()

            total = sum(counts.values())

            specific_aspects.append([

                col,

                round((counts.get(5,0)/total)*100,1) if total else 0,

                round((counts.get(4,0)/total)*100,1) if total else 0,

                round((counts.get(3,0)/total)*100,1) if total else 0,

                round((counts.get(2,0)/total)*100,1) if total else 0,

                round((counts.get(1,0)/total)*100,1) if total else 0

            ])

    section3_df = pd.DataFrame(

        specific_aspects,

        columns=[

            "ASPECT OF THE PROGRAM",

            "Excellent %",

            "Very Good %",

            "Satisfactory %",

            "Poor %",

            "Very Poor %"

        ]

    )

    # =====================================================
    # SECTION 8
    # Institution Comparison
    # =====================================================
    section8 = []

    if comparison_col:

        counts = df[comparison_col].value_counts().to_dict()

        total = sum(counts.values())

        labels = {

            5: "5 - Very High",

            4: "4 - High",

            3: "3 - Average",

            2: "2 - Low",

            1: "1 - Very Low"

        }

        for score in [5, 4, 3, 2, 1]:

            pct = (

                round(

                    (counts.get(score, 0) / total) * 100,

                    1

                )

                if total > 0

                else 0

            )

            section8.append([

                labels[score],

                pct

            ])

    section8_df = pd.DataFrame(

        section8,

        columns=[

            "Rating",

            "Percentage of Respondents"

        ]

    )

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

            qualitative_mapping["Suggestions"] = col

        elif "other areas you would like added" in col_lower:

            qualitative_mapping["Areas to Add"] = col

        elif "other ksg training programs" in col_lower:

            qualitative_mapping["Interest in Other KSG Programmes"] = col

        elif "other training programs not currently offered" in col_lower:

            qualitative_mapping["Additional Training Areas"] = col

        elif "other comments" in col_lower:

            qualitative_mapping["General Comments"] = col

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

            responses = responses[responses != ""]

            qualitative_outputs[section] = " ".join(responses)

        else:

            qualitative_outputs[section] = ""

    # =====================================================
    # SAVE ANALYSIS
    # =====================================================
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
        # SECTION SHEETS
        # =================================================
        section1_df.to_excel(

            writer,

            sheet_name="Section 1 Objectives",

            index=False

        )

        section2_df.to_excel(

            writer,

            sheet_name="Section 2 Expectations",

            index=False

        )

        section3_df.to_excel(

            writer,

            sheet_name="Section 3 Specific Aspects",

            index=False

        )

        section8_df.to_excel(

            writer,

            sheet_name="Section 8 Comparison",

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

    print(f"\nAnalysis completed: {output_file}")

    return output_file