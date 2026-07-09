import os

from .clean import clean_eee
from .analyze import analyze_eee
from .report import generate_eee_report
from .report_llm import generate_eee_report_llm


def run_eee(
    excel_file,
    programme_title,
    programme_code,
    duration,
    venue,
    coordinator,
    assistant,
    use_llm=True
):
    """
    Complete End of Event Evaluation workflow.

    Returns
    -------
    str
        Path to generated report.
    """

    # -----------------------------
    # STEP 1 : CLEAN
    # -----------------------------
    cleaned_file = clean_eee(excel_file)

    # -----------------------------
    # STEP 2 : ANALYSE
    # -----------------------------
    analysed_file = analyze_eee(cleaned_file)

    # -----------------------------
    # STEP 3 : REPORT
    # -----------------------------
    if use_llm:

        report = generate_eee_report_llm(

            analysis_file=analysed_file,

            programme_title=programme_title,

            programme_code=programme_code,

            duration=duration,

            venue=venue,

            coordinator=coordinator,

            assistant=assistant

        )

    else:

        report = generate_eee_report(

            analysis_file=analysed_file,

            programme_title=programme_title,

            programme_code=programme_code,

            duration=duration,

            venue=venue,

            coordinator=coordinator,

            assistant=assistant

        )

    return report