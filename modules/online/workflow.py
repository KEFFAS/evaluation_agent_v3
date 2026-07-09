"""
Online End of Event Evaluation Workflow
Runs the complete Online EEE pipeline.
"""

from .clean import clean_online
from .report import generate_online_report


def run_online(
    csv_file,
    programme_title,
    programme_code,
    duration,
    venue,
    coordinator,
    assistant,
    output_folder="outputs"
):
    """
    Complete Online End of Event Evaluation workflow.

    Parameters
    ----------
    csv_file : str
        Uploaded CSV file.

    programme_title : str

    programme_code : str

    duration : str

    venue : str

    coordinator : str

    assistant : str

    output_folder : str

    Returns
    -------
    str
        Path to generated report.
    """

    print("=" * 60)
    print("STEP 1 : CLEANING")
    print("=" * 60)

    cleaned_file = clean_online(

        csv_file,

        output_folder

    )

    print("=" * 60)
    print("STEP 2 : REPORT GENERATION")
    print("=" * 60)

    report = generate_online_report(

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
    print("ONLINE EVALUATION COMPLETED")
    print("=" * 60)

    print(f"Report File : {report}")

    return report