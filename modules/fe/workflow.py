"""
Facilitator Evaluation Workflow
Runs the complete Facilitator Evaluation pipeline.

Pipeline:
Raw Excel
    ↓
Cleaning
    ↓
Analysis
    ↓
Final Report Generation
"""

import os

from .clean import clean_fe
from .analyze import analyze_fe
from .report import generate_fe_report
from .report_llm import generate_fe_report_llm


def run_fe(
    excel_file,
    programme_title,
    programme_code="",
    duration="",
    venue="",
    coordinator="",
    assistant="",
    total_participants=1,
    use_llm=True,
    generate_analysis=True,
    output_folder="outputs"
):
    """
    Runs the complete Facilitator Evaluation workflow.

    Parameters
    ----------
    excel_file : str
        Path to uploaded raw Excel file.

    programme_title : str
        Name of the programme/course.

    programme_code : str
        Programme code.

    duration : str
        Programme duration.

    venue : str
        Training venue.

    coordinator : str
        Programme coordinator.

    assistant : str
        Assistant coordinator.

    total_participants : int
        Total number of participants enrolled in the class.

    use_llm : bool
        Whether to use AI-enhanced qualitative analysis.

    generate_analysis : bool
        Whether to generate the analysis Excel workbook.

    output_folder : str
        Folder where generated files will be saved.

    Returns
    -------
    dict
        Dictionary containing:
        - cleaned_file
        - analysis_file
        - report_file
    """

    # =====================================================
    # CREATE OUTPUT FOLDER
    # =====================================================

    os.makedirs(
        output_folder,
        exist_ok=True
    )

    # =====================================================
    # STEP 1: CLEANING
    # =====================================================

    print("\n" + "=" * 60)
    print("STEP 1 : CLEANING FACILITATOR EVALUATION DATA")
    print("=" * 60)

    cleaned_file = clean_fe(
        file_name=excel_file,
        output_folder=output_folder
    )

    print(
        f"✓ Cleaning completed: {cleaned_file}"
    )

    # =====================================================
    # STEP 2: ANALYSIS
    # =====================================================

    analysis_file = None

    if generate_analysis:

        print("\n" + "=" * 60)
        print("STEP 2 : ANALYSING FACILITATOR EVALUATION DATA")
        print("=" * 60)

        analysis_file = analyze_fe(
            cleaned_file=cleaned_file,
            total_participants=total_participants,
            output_folder=output_folder
        )

        print(
            f"✓ Analysis completed: {analysis_file}"
        )

    else:

        print(
            "\nSTEP 2 : ANALYSIS SKIPPED"
        )

    # =====================================================
    # STEP 3: REPORT GENERATION
    # =====================================================

    print("\n" + "=" * 60)
    print("STEP 3 : GENERATING FACILITATOR EVALUATION REPORT")
    print("=" * 60)

    if use_llm:

        print(
            "Using AI-enhanced qualitative analysis..."
        )

        report_file = generate_fe_report_llm(

            cleaned_file=cleaned_file,

            programme_title=programme_title,

            programme_code=programme_code,

            duration=duration,

            venue=venue,

            coordinator=coordinator,

            assistant=assistant,

            total_participants=total_participants,

            output_folder=output_folder
        )

    else:

        print(
            "Using standard report generation..."
        )

        report_file = generate_fe_report(

            cleaned_file=cleaned_file,

            programme_title=programme_title,

            programme_code=programme_code,

            duration=duration,

            venue=venue,

            coordinator=coordinator,

            assistant=assistant,

            total_participants=total_participants,

            output_folder=output_folder
        )

    print(
        f"✓ Report generated: {report_file}"
    )

    # =====================================================
    # COMPLETION
    # =====================================================

    print("\n" + "=" * 60)
    print("FACILITATOR EVALUATION COMPLETED SUCCESSFULLY")
    print("=" * 60)

    print(
        f"Cleaned File  : {cleaned_file}"
    )

    if analysis_file:

        print(
            f"Analysis File : {analysis_file}"
        )

    print(
        f"Report File   : {report_file}"
    )

    # =====================================================
    # RETURN ALL FILES
    # =====================================================

    return {

        "cleaned_file":
            cleaned_file,

        "analysis_file":
            analysis_file,

        "report_file":
            report_file
    }