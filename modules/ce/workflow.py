"""
Coordinator Evaluation Workflow
Runs the complete Coordinator Evaluation pipeline.
"""

from .clean import clean_ce
from .analyze import analyze_ce
from .report import generate_ce_report


def run_ce(
    excel_file,
    programme_title,
    programme_code,
    duration,
    venue,
    coordinator,
    assistant,
    generate_analysis=True,
    output_folder="outputs"
):
    """
    Complete Coordinator Evaluation workflow.

    Parameters
    ----------
    excel_file : str
        Uploaded Excel file.

    programme_title : str

    programme_code : str

    duration : str

    venue : str

    coordinator : str

    assistant : str

    generate_analysis : bool

    output_folder : str

    Returns
    -------
    str
        Generated report path.
    """

    print("=" * 60)
    print("STEP 1 : CLEANING")
    print("=" * 60)

    cleaned_file = clean_ce(

        excel_file,

        output_folder

    )

    analysis_file = None

    if generate_analysis:

        print("=" * 60)
        print("STEP 2 : ANALYSIS")
        print("=" * 60)

        analysis_file = analyze_ce(

            cleaned_file,

            output_folder

        )

    print("=" * 60)
    print("STEP 3 : REPORT GENERATION")
    print("=" * 60)

    report = generate_ce_report(

        cleaned_file=cleaned_file,

        programme_title=programme_title,

        programme_code=programme_code,

        duration=duration,

        venue=venue,

        coordinator=coordinator,

        assistant=assistant,

        output_folder=output_folder

    )

    print("=" * 60)
    print("COORDINATOR EVALUATION COMPLETED")
    print("=" * 60)

    if analysis_file:

        print(f"Analysis File : {analysis_file}")

    print(f"Report File   : {report}")

    return report