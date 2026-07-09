"""
Facilitator Evaluation Workflow
Runs the complete Facilitator Evaluation pipeline.
"""

from .clean import clean_fe
from .analyze import analyze_fe
from .report import generate_fe_report
from .report_llm import generate_fe_report_llm


def run_fe(
    excel_file,
    programme_title,
    programme_code,
    duration,
    venue,
    coordinator,
    assistant,
    total_participants,
    use_llm=True,
    generate_analysis=True,
    output_folder="outputs"
):
    """
    Complete Facilitator Evaluation workflow.

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

    total_participants : int

    use_llm : bool

    generate_analysis : bool

    output_folder : str

    Returns
    -------
    str
        Path to generated report.
    """

    print("=" * 60)
    print("STEP 1 : CLEANING")
    print("=" * 60)

    cleaned_file = clean_fe(

        excel_file,

        output_folder

    )

    analysis_file = None

    if generate_analysis:

        print("=" * 60)
        print("STEP 2 : ANALYSIS")
        print("=" * 60)

        analysis_file = analyze_fe(

            cleaned_file,

            total_participants,

            output_folder

        )

    print("=" * 60)
    print("STEP 3 : REPORT GENERATION")
    print("=" * 60)

    if use_llm:

        report = generate_fe_report_llm(

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

        report = generate_fe_report(

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

    print("=" * 60)
    print("FACILITATOR EVALUATION COMPLETED")
    print("=" * 60)

    if analysis_file:

        print(f"Analysis File : {analysis_file}")

    print(f"Report File   : {report}")

    return report