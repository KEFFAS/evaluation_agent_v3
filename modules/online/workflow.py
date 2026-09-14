from .clean import clean_online
from .analyze import analyze_online
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

    print("=" * 60)
    print("STEP 1 : CLEANING")
    print("=" * 60)

    cleaned_file = clean_online(
        csv_file,
        output_folder
    )

    print("=" * 60)
    print("STEP 2 : ANALYSIS")
    print("=" * 60)

    analysis_file = analyze_online(
        cleaned_file,
        output_folder
    )

    print("=" * 60)
    print("STEP 3 : REPORT GENERATION")
    print("=" * 60)

    report_file = generate_online_report(

        cleaned_file=cleaned_file,

        analysis_file=analysis_file,

        programme_title=programme_title,

        programme_code=programme_code,

        duration=duration,

        venue=venue,

        coordinator=coordinator,

        assistant=assistant,

        output_folder=output_folder

    )

    return {

        "cleaned_file": cleaned_file,

        "analysis_file": analysis_file,

        "report_file": report_file

    }