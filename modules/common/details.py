import pandas as pd


def extract_program_details(cleaned_file):
    """
    Extract programme details from cleaned EEE data.

    Returns a dictionary.
    """

    df = pd.read_excel(cleaned_file)

    details = {

        "programme_title": "",

        "programme_code": "",

        "venue": "",

        "coordinator": "",

        "assistant": "",

        "duration": ""

    }

    column_mapping = {

        "Program Title": "programme_title",

        "Program Code": "programme_code",

        "Venue / Campus": "venue",

        "Coordinator Name": "coordinator",

        "Program Assistant Name": "assistant",

        "Course Duration (Days)": "duration"

    }

    for column, key in column_mapping.items():

        if column in df.columns:

            value = df[column].dropna()

            if not value.empty:

                details[key] = str(value.iloc[0]).strip()

    return details