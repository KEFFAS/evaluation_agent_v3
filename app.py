import os
import shutil
import gradio as gr

from modules.eee.workflow import run_eee


UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def generate_report(
    uploaded_file,
    duration,
    use_llm
):
    """
    Generates an EEE report.
    """

    if uploaded_file is None:
        return None, "Please upload an Excel file."

    # Copy uploaded file
    destination = os.path.join(
        UPLOAD_FOLDER,
        os.path.basename(uploaded_file)
    )

    shutil.copy(uploaded_file, destination)

    try:

        report = run_eee(

            excel_file=destination,

            duration=duration,

            use_llm=use_llm,

            generate_analysis=True,

            output_folder=OUTPUT_FOLDER

        )

        return report, "✅ Report generated successfully."

    except Exception as e:

        return None, f"❌ {str(e)}"


with gr.Blocks(
    title="KSG Evaluation Intelligence System"
) as app:

    gr.Markdown(
        """
# Kenya School of Government

## Evaluation Intelligence System

Generate End-of-Event Evaluation reports in one click.
"""
    )

    with gr.Row():

        uploaded_file = gr.File(
            label="Upload Evaluation Excel File",
            file_types=[".xlsx", ".xls"]
        )

    duration = gr.Textbox(

        label="Duration",

        placeholder="Example: 3rd May 2026 to 10th May 2026"

    )

    use_llm = gr.Checkbox(

        value=True,

        label="AI Enhanced Report"

    )

    generate = gr.Button(

        "🚀 Generate Report",

        variant="primary"

    )

    status = gr.Textbox(

        label="Status"

    )

    report_file = gr.File(

        label="Download Report"

    )

    generate.click(

        fn=generate_report,

        inputs=[

            uploaded_file,

            duration,

            use_llm

        ],

        outputs=[

            report_file,

            status

        ]

    )

app.launch(share=True)