import streamlit as st
import os
import shutil
from datetime import datetime

from modules.fe.workflow import run_fe


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="KSG Evaluation System",
    page_icon="📊",
    layout="wide"
)


# =========================================================
# TITLE
# =========================================================

st.title("📊 KSG Evaluation System")

st.subheader(
    "Facilitator Evaluation Module"
)

st.write(
    "Upload the Facilitator Evaluation Excel file, "
    "enter programme details, and generate the analysis "
    "and final evaluation report."
)


# =========================================================
# CREATE FOLDERS
# =========================================================

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(
    UPLOAD_FOLDER,
    exist_ok=True
)

os.makedirs(
    OUTPUT_FOLDER,
    exist_ok=True
)


# =========================================================
# PROGRAMME DETAILS
# =========================================================

st.markdown("---")

st.subheader("1. Programme Details")


col1, col2 = st.columns(2)


with col1:

    programme_title = st.text_input(
        "Programme Title *"
    )

    programme_code = st.text_input(
        "Programme Code"
    )

    duration = st.text_input(
        "Duration"
    )


with col2:

    venue = st.text_input(
        "Venue"
    )

    coordinator = st.text_input(
        "Coordinator"
    )

    assistant = st.text_input(
        "Assistant Coordinator"
    )


# =========================================================
# PARTICIPANTS
# =========================================================

st.markdown("---")

st.subheader("2. Participant Information")


total_participants = st.number_input(
    "Total Participants in the Class *",
    min_value=1,
    step=1,
    help=(
        "Enter the total number of participants "
        "enrolled in the class. "
        "This is different from the number of respondents."
    )
)


st.info(
    "Response rates will be calculated automatically "
    "from the number of participants who submitted "
    "evaluation responses."
)


# =========================================================
# AI OPTION
# =========================================================

st.markdown("---")

st.subheader("3. Report Options")


use_llm = st.checkbox(
    "Use AI for qualitative analysis",
    value=True,
    help=(
        "AI will analyse participant comments and "
        "generate professional qualitative findings."
    )
)


generate_analysis = st.checkbox(
    "Generate Analysis Excel",
    value=True
)


# =========================================================
# FILE UPLOAD
# =========================================================

st.markdown("---")

st.subheader("4. Upload Facilitator Evaluation File")


uploaded_file = st.file_uploader(
    "Upload Raw Facilitator Evaluation Excel File *",
    type=["xlsx", "xls"]
)


# =========================================================
# GENERATE BUTTON
# =========================================================

st.markdown("---")


generate_button = st.button(
    "🚀 Generate Facilitator Evaluation",
    type="primary",
    use_container_width=True
)


# =========================================================
# PROCESS
# =========================================================

if generate_button:

    # =====================================================
    # VALIDATION
    # =====================================================

    if not programme_title.strip():

        st.error(
            "Please enter the Programme Title."
        )

    elif uploaded_file is None:

        st.error(
            "Please upload the Facilitator Evaluation Excel file."
        )

    else:

        try:

            # =================================================
            # CREATE UNIQUE FILE NAME
            # =================================================

            timestamp = datetime.now().strftime(
                "%Y%m%d_%H%M%S"
            )

            original_name = os.path.basename(
                uploaded_file.name
            )

            file_extension = os.path.splitext(
                original_name
            )[1]

            input_file_name = (
                f"FE_{timestamp}{file_extension}"
            )

            input_file_path = os.path.join(
                UPLOAD_FOLDER,
                input_file_name
            )

            # =================================================
            # SAVE UPLOADED FILE
            # =================================================

            with open(
                input_file_path,
                "wb"
            ) as f:

                f.write(
                    uploaded_file.getvalue()
                )

            # =================================================
            # CREATE UNIQUE OUTPUT FOLDER
            # =================================================

            run_output_folder = os.path.join(

                OUTPUT_FOLDER,

                f"FE_{timestamp}"
            )

            os.makedirs(
                run_output_folder,
                exist_ok=True
            )

            # =================================================
            # PROCESS
            # =================================================

            with st.spinner(
                "Processing Facilitator Evaluation data..."
            ):

                results = run_fe(

                    excel_file=input_file_path,

                    programme_title=programme_title,

                    programme_code=programme_code,

                    duration=duration,

                    venue=venue,

                    coordinator=coordinator,

                    assistant=assistant,

                    total_participants=int(
                        total_participants
                    ),

                    use_llm=use_llm,

                    generate_analysis=generate_analysis,

                    output_folder=run_output_folder
                )

            # =================================================
            # SUCCESS
            # =================================================

            st.success(
                "Facilitator Evaluation completed successfully!"
            )

            # =================================================
            # RESULTS
            # =================================================

            st.markdown("---")

            st.subheader(
                "5. Download Results"
            )

            download_col1, download_col2, download_col3 = st.columns(3)


            # =================================================
            # CLEANED FILE
            # =================================================

            cleaned_file = results.get(
                "cleaned_file"
            )

            if cleaned_file and os.path.exists(
                cleaned_file
            ):

                with open(
                    cleaned_file,
                    "rb"
                ) as f:

                    download_col1.download_button(

                        label="📄 Download Cleaned Excel",

                        data=f.read(),

                        file_name=os.path.basename(
                            cleaned_file
                        ),

                        mime=(
                            "application/"
                            "vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),

                        use_container_width=True
                    )


            # =================================================
            # ANALYSIS FILE
            # =================================================

            analysis_file = results.get(
                "analysis_file"
            )

            if analysis_file and os.path.exists(
                analysis_file
            ):

                with open(
                    analysis_file,
                    "rb"
                ) as f:

                    download_col2.download_button(

                        label="📊 Download Analysis Excel",

                        data=f.read(),

                        file_name=os.path.basename(
                            analysis_file
                        ),

                        mime=(
                            "application/"
                            "vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),

                        use_container_width=True
                    )

            else:

                download_col2.info(
                    "Analysis Excel was not generated."
                )


            # =================================================
            # FINAL REPORT
            # =================================================

            report_file = results.get(
                "report_file"
            )

            if report_file and os.path.exists(
                report_file
            ):

                with open(
                    report_file,
                    "rb"
                ) as f:

                    download_col3.download_button(

                        label="🤖 Download Final AI Report",

                        data=f.read(),

                        file_name=os.path.basename(
                            report_file
                        ),

                        mime=(
                            "application/"
                            "vnd.openxmlformats-officedocument."
                            "wordprocessingml.document"
                        ),

                        use_container_width=True
                    )

            else:

                download_col3.error(
                    "Final report was not generated."
                )


            # =================================================
            # SHOW FILE DETAILS
            # =================================================

            st.markdown("---")

            st.caption(
                f"Total Participants: {int(total_participants)}"
            )

            st.caption(
                f"Programme: {programme_title}"
            )

        except Exception as e:

            st.error(
                "An error occurred while generating the evaluation."
            )

            st.exception(e)