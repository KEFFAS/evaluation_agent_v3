import gradio as gr
import os

from modules.eee.workflow import run_eee
from modules.fe.workflow import run_fe
from modules.ce.workflow import run_ce
from modules.online.workflow import run_online


# =====================================================
# APPLICATION DETAILS
# =====================================================

APP_TITLE = "Kenya School of Government Evaluation Agent"

APP_DESCRIPTION = """
Generate institutional evaluation reports automatically.

Supported modules:

• End of Event Evaluation (EEE)

• Facilitator Evaluation (FE)

• Coordinator Evaluation (CE)

• Online End of Event Evaluation

Upload the evaluation file, enter the programme details,
and generate professionally formatted evaluation outputs.
"""


# =====================================================
# EXTRACT OUTPUT FILES
# =====================================================

def get_output_files(result):
    """
    Standardizes workflow output.

    Expected workflow output:

    {
        "cleaned_file": "...",
        "analysis_file": "...",
        "report_file": "..."
    }
    """

    # -----------------------------------------------
    # WORKFLOW RETURNS DICTIONARY
    # -----------------------------------------------

    if isinstance(result, dict):

        cleaned_file = result.get("cleaned_file")

        analysis_file = result.get("analysis_file")

        report_file = result.get("report_file")

        return (
            cleaned_file,
            analysis_file,
            report_file
        )

    # -----------------------------------------------
    # WORKFLOW RETURNS ONLY REPORT PATH
    # Backward compatibility
    # -----------------------------------------------

    elif isinstance(result, (str, os.PathLike)):

        return (
            None,
            None,
            str(result)
        )

    # -----------------------------------------------
    # INVALID OUTPUT
    # -----------------------------------------------

    else:

        raise gr.Error(
            f"Unexpected workflow output type: {type(result)}"
        )


# =====================================================
# PROCESS REPORT
# =====================================================

def process_report(

    evaluation_type,
    uploaded_file,
    programme_title,
    programme_code,
    duration,
    venue,
    coordinator,
    assistant,
    total_participants,
    use_llm

):

    # =================================================
    # VALIDATE FILE
    # =================================================

    if uploaded_file is None:

        raise gr.Error(
            "Please upload an evaluation file."
        )

    file_path = uploaded_file.name


    # =================================================
    # END OF EVENT EVALUATION
    # =================================================

    if evaluation_type == "End of Event Evaluation":

        result = run_eee(

            excel_file=file_path,

            duration=duration,

            use_llm=use_llm,

            generate_analysis=True,

            output_folder="outputs"

        )

        return get_output_files(result)


    # =================================================
    # FACILITATOR EVALUATION
    # =================================================

    elif evaluation_type == "Facilitator Evaluation":

        if not total_participants or total_participants <= 0:

            raise gr.Error(
                "Please enter the total number of participants."
            )

        result = run_fe(

            excel_file=file_path,

            programme_title=programme_title,

            programme_code=programme_code,

            duration=duration,

            venue=venue,

            coordinator=coordinator,

            assistant=assistant,

            total_participants=int(total_participants),

            use_llm=use_llm,

            generate_analysis=True,

            output_folder="outputs"

        )

        return get_output_files(result)


    # =================================================
    # COORDINATOR EVALUATION
    # =================================================

    elif evaluation_type == "Coordinator Evaluation":

        result = run_ce(

            excel_file=file_path,

            programme_title=programme_title,

            programme_code=programme_code,

            duration=duration,

            venue=venue,

            coordinator=coordinator,

            assistant=assistant,

            generate_analysis=True,

            output_folder="outputs"

        )

        return get_output_files(result)


    # =================================================
    # ONLINE END OF EVENT EVALUATION
    # =================================================

    elif evaluation_type == "Online End of Event Evaluation":

        result = run_online(

            csv_file=file_path,

            programme_title=programme_title,

            programme_code=programme_code,

            duration=duration,

            venue=venue,

            coordinator=coordinator,

            assistant=assistant,

            output_folder="outputs"

        )

        return get_output_files(result)


    # =================================================
    # INVALID TYPE
    # =================================================

    else:

        raise gr.Error(
            "Invalid evaluation type selected."
        )


# =====================================================
# USER INTERFACE
# =====================================================

with gr.Blocks(
    title=APP_TITLE
) as app:

    # =================================================
    # TITLE
    # =================================================

    gr.Markdown(
        f"# {APP_TITLE}"
    )

    gr.Markdown(
        APP_DESCRIPTION
    )


    # =================================================
    # EVALUATION TYPE
    # =================================================

    evaluation_type = gr.Dropdown(

        choices=[

            "End of Event Evaluation",

            "Facilitator Evaluation",

            "Coordinator Evaluation",

            "Online End of Event Evaluation"

        ],

        label="Evaluation Type",

        value="End of Event Evaluation"

    )


    # =================================================
    # FILE UPLOAD
    # =================================================

    uploaded_file = gr.File(

        label="Upload Evaluation File"

    )


    # =================================================
    # PROGRAMME DETAILS
    # =================================================

    gr.Markdown(
        "## Programme Details"
    )


    # =================================================
    # ROW 1
    # =================================================

    with gr.Row():

        programme_title = gr.Textbox(

            label="Programme Title"

        )

        programme_code = gr.Textbox(

            label="Programme Code"

        )


    # =================================================
    # ROW 2
    # =================================================

    with gr.Row():

        duration = gr.Textbox(

            label="Duration"

        )

        venue = gr.Textbox(

            label="Venue"

        )


    # =================================================
    # ROW 3
    # =================================================

    with gr.Row():

        coordinator = gr.Textbox(

            label="Coordinator"

        )

        assistant = gr.Textbox(

            label="Programme Assistant"

        )


    # =================================================
    # TOTAL PARTICIPANTS
    # =================================================

    total_participants = gr.Number(

        label="Total Participants",

        value=0,

        precision=0

    )


    # =================================================
    # AI OPTION
    # =================================================

    use_llm = gr.Checkbox(

        label="Use AI-enhanced Report",

        value=True

    )


    # =================================================
    # GENERATE BUTTON
    # =================================================

    generate_btn = gr.Button(

        "Generate Evaluation Outputs",

        variant="primary"

    )


    # =================================================
    # DOWNLOAD SECTION
    # =================================================

    gr.Markdown(
        "## Downloads"
    )


    with gr.Row():

        cleaned_output = gr.File(

            label="📥 Download Cleaned File"

        )

        analysis_output = gr.File(

            label="📊 Download Analysis File"

        )

        report_output = gr.File(

            label="📄 Download Final Report"

        )


    # =================================================
    # BUTTON EVENT
    # =================================================

    generate_btn.click(

        fn=process_report,

        inputs=[

            evaluation_type,

            uploaded_file,

            programme_title,

            programme_code,

            duration,

            venue,

            coordinator,

            assistant,

            total_participants,

            use_llm

        ],

        outputs=[

            cleaned_output,

            analysis_output,

            report_output

        ]

    )


# =====================================================
# LAUNCH APPLICATION
# =====================================================

if __name__ == "__main__":

    app.launch(

        server_name="0.0.0.0",

        server_port=int(
            os.environ.get(
                "PORT",
                7860
            )
        ),

        show_error=True

    )