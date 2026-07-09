"""
EEE Workflow
Runs the complete End of Event Evaluation pipeline.
"""

from .clean import clean_eee
from .analyze import analyze_eee
from .report import generate_eee_report
from .report_llm import generate_eee_report_llm

from modules.common.details import extract_program_details


def run_eee(
    excel_file,
    duration,
    use_llm=True,
    generate_analysis=True,
    output_folder="outputs"
):
    """
    Complete End of Event Evaluation workflow.

    Parameters
    ----------
    excel_file : str
        Path to uploaded Excel file.

    use_llm : bool
        True = AI-enhanced report
        False = Standard report

    generate_analysis : bool
        Whether to generate the analysis workbook.

    output_folder : str

    Returns
    -------
    str
        Path to generated report.
    """

    print("=" * 60)
    print("STEP 1 : CLEANING")
    print("=" * 60)

    cleaned_file = clean_eee(
        excel_file,
        output_folder
    )

    print("=" * 60)
    print("STEP 2 : EXTRACTING PROGRAM DETAILS")
    print("=" * 60)

    details = extract_program_details(cleaned_file)

    print(f"Programme Title : {details['programme_title']}")
    print(f"Programme Code  : {details['programme_code']}")
    print(f"Venue           : {details['venue']}")
    print(f"Coordinator     : {details['coordinator']}")
    print(f"Assistant       : {details['assistant']}")
    print(f"Duration        : {details['duration']}")

    if generate_analysis:

        print("=" * 60)
        print("STEP 3 : ANALYSIS")
        print("=" * 60)

        analyze_eee(
            cleaned_file,
            output_folder
        )

    print("=" * 60)
    print("STEP 4 : REPORT GENERATION")
    print("=" * 60)

    if use_llm:

        report = generate_eee_report_llm(

            cleaned_file=cleaned_file,

            programme_title=details["programme_title"],

            programme_code=details["programme_code"],

            duration=details["duration"],

            venue=details["venue"],

            coordinator=details["coordinator"],

            assistant=details["assistant"],

            output_folder=output_folder

        )

    else:

        report = generate_eee_report(

            cleaned_file=cleaned_file,

            programme_title=details["programme_title"],

            programme_code=details["programme_code"],

            duration=details["duration"],

            venue=details["venue"],

            coordinator=details["coordinator"],

            assistant=details["assistant"],

            output_folder=output_folder

        )

    print("=" * 60)
    print("EEE WORKFLOW COMPLETED SUCCESSFULLY")
    print("=" * 60)
    print(f"Report saved to: {report}")

    return report