import streamlit as st
import os

st.title("KSG Evaluation System")

# ===== UPLOAD =====
uploaded_file = st.file_uploader("Upload Cleaned Excel File", type=["xlsx"])

# ===== INPUTS =====
program_title = st.text_input("Program Title")
report_date = st.text_input("Report Date")
participants = st.number_input("Total Participants", min_value=1)

# ===== BUTTON =====
if uploaded_file and st.button("Generate Report"):

    # Save uploaded file
    with open("input.xlsx", "wb") as f:
        f.write(uploaded_file.read())

    st.info("Processing... please wait")

    # ===== RUN YOUR EXISTING SCRIPT =====
    os.system("python generate_report_llm.py")

    # ===== OUTPUT =====
    output_file = "input_KSG_report_LLM.docx"  # adjust if needed

    if os.path.exists(output_file):
        with open(output_file, "rb") as f:
            st.download_button(
                label="Download Report",
                data=f,
                file_name=output_file
            )

    st.success("Done!")