import gradio as gr
import os

from modules.eee.workflow import run_eee
from modules.fe.workflow import run_fe
from modules.ce.workflow import run_ce
from modules.online.workflow import run_online


APP_TITLE = "Kenya School of Government Evaluation Agent"

APP_DESCRIPTION = """
Generate institutional evaluation reports automatically.

Supported modules:

• End of Event Evaluation (EEE)

• Facilitator Evaluation (FE)

• Coordinator Evaluation (CE)

• Online End of Event Evaluation

Upload the evaluation file, enter the programme details, and generate a professionally formatted report.
"""


# =====================================================
# EXTRACT REPORT FILE
# =====================================================

def get_report_file(result):

    # If workflow returns a dictionary
    if isinstance(result, dict):

        report_file = result.get("report_file")

        if report_file:

            return report_file

        raise gr.Error(
            "Report generation completed but no report file was returned."
        )

    # If workflow returns a string path
    if isinstance(result, (str, os.PathLike)):

        return str(result)

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

    if uploaded_file is None:

        raise gr.Error(
            "Please upload an evaluation file."
        )

    file_path = uploaded_file.name


    # =====================================================
    # END OF EVENT EVALUATION
    # =====================================================

    if evaluation_type == "End of Event Evaluation":

        result = run_eee(

            excel_file=file_path,

            duration=duration,

            use_llm=use_llm,

            generate_analysis=True,

            output_folder="outputs"

        )

        return get_report_file(result)


    # =====================================================
    # FACILITATOR EVALUATION
    # =====================================================

    elif evaluation_type == "Facilitator Evaluation":

        if total_participants <= 0:

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

        return get_report_file(result)


    # =====================================================
    # COORDINATOR EVALUATION
    # =====================================================

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

        return get_report_file(result)


    # =====================================================
    # ONLINE END OF EVENT EVALUATION
    # =====================================================

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

        return get_report_file(result)


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

    gr.Markdown(f"# {APP_TITLE}")

    gr.Markdown(APP_DESCRIPTION)


    # =====================================================
    # EVALUATION TYPE
    # =====================================================

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


    # =====================================================
    # FILE UPLOAD
    # =====================================================

    uploaded_file = gr.File(

        label="Upload Evaluation File"

    )


    # =====================================================
    # PROGRAMME DETAILS
    # =====================================================

    gr.Markdown("## Programme Details")


    with gr.Row():

        programme_title = gr.Textbox(

            label="Programme Title"

        )

        programme_code = gr.Textbox(

            label="Programme Code"

        )


    with gr.Row():

        duration = gr.Textbox(

            label="Duration"

        )

        venue = gr.Textbox(

            label="Venue"

        )


    with gr.Row():

        coordinator = gr.Textbox(

            label="Coordinator"

        )

        assistant = gr.Textbox(

            label="Programme Assistant"

        )


    total_participants = gr.Number(

        label="Total Participants",

        value=0,

        precision=0

    )


    # =====================================================
    # AI OPTION
    # =====================================================

    use_llm = gr.Checkbox(

        label="Use AI-enhanced Report",

        value=True

    )


    # =====================================================
    # GENERATE BUTTON
    # =====================================================

    generate_btn = gr.Button(

        "Generate Report",

        variant="primary"

    )


    # =====================================================
    # OUTPUT
    # =====================================================

    output_file = gr.File(

        label="Download Generated Report"

    )


    # =====================================================
    # BUTTON EVENT
    # =====================================================

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

        outputs=output_file

    )


# =====================================================
# LAUNCH APPLICATION
# =====================================================

if __name__ == "__main__":

    app.launch(

        server_name="0.0.0.0",

        server_port=int(
            os.environ.get("PORT", 7860)
        ),

        show_error=True

    )