import os
import pandas as pd

from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from openai import OpenAI
from dotenv import load_dotenv


# =========================================================
# LOAD ENVIRONMENT
# =========================================================
load_dotenv()

client = OpenAI(
    api_key=os.getenv("OPENAI_API_KEY")
)


# =========================================================
# TABLE BORDER FUNCTION
# =========================================================
def set_table_borders(table):

    tbl = table._tbl
    tblPr = tbl.tblPr

    borders = OxmlElement("w:tblBorders")

    for border_name in [

        "top",
        "left",
        "bottom",
        "right",
        "insideH",
        "insideV"

    ]:

        border = OxmlElement(f"w:{border_name}")

        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "8")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "000000")

        borders.append(border)

    tblPr.append(borders)


# =========================================================
# LLM QUALITATIVE ANALYSIS
# =========================================================
def analyze_qualitative(text):

    if not text.strip():

        return "No comments provided."

    text = text[:4000]

    prompt = f"""
The following comments were provided by participants during evaluation of a facilitator at the Kenya School of Government.

Participant Feedback:

{text}

Write exactly TWO concise professional paragraphs.

Paragraph 1:
Summarize the most recurring positive feedback regarding facilitation, delivery style, subject mastery, participant engagement, communication, responsiveness and overall teaching effectiveness.

Paragraph 2:
Summarize the most recurring suggestions for improvement.

Requirements

- Formal institutional language
- Human tone
- Evidence based
- No bullet points
- No headings
- No repetition
- Suitable for an official KSG report
"""

    response = client.chat.completions.create(

        model="gpt-5-nano-2025-08-07",

        messages=[

            {

                "role": "system",

                "content":
                (
                    "You are an institutional monitoring and evaluation officer "
                    "writing formal Kenya School of Government evaluation reports. "
                    "Write in a professional, concise, evidence-based and human tone."
                )

            },

            {

                "role": "user",

                "content": prompt

            }

        ]

    )

    return response.choices[0].message.content.strip()


# =========================================================
# REPORT FUNCTION
# =========================================================
def generate_fe_report_llm(

    cleaned_file,

    programme_title,

    programme_code,

    duration,

    venue,

    coordinator,

    assistant,

    total_participants,

    output_folder="outputs"

):

    os.makedirs(output_folder, exist_ok=True)

    # =====================================================
    # LOAD DATA
    # =====================================================
    df = pd.read_excel(cleaned_file)

    df.columns = (

        df.columns

        .astype(str)

        .str.strip()

        .str.title()

    )

    rename_map = {

        "Topic": "Topic Description",

        "Session Topic": "Topic Description",

        "Facilitator": "Lecturer Name",

        "Lecturer": "Lecturer Name"

    }

    df = df.rename(columns=rename_map)

    df = df.sort_values(

        by=[

            "Lecturer Name",

            "Topic Description"

        ]

    )

    rating_cols = [

        "Punctuality",

        "Presentation Flow",

        "Handling Questions",

        "Active Participation Of Learners",

        "Use Of Visual Aids",

        "Relevance Of Subject To Workplace",

        "Use Of Relevant Examples",

        "Knowledge Of Subject",

        "Treats Participants With Dignity And Respect",

        "Variety And Appropriateness Of Training Methods"

    ]

    rating_cols = [

        c

        for c in rating_cols

        if c in df.columns

    ]

    grouped = df.groupby(

        [

            "Lecturer Name",

            "Topic Description"

        ]

    )

    # =====================================================
    # CREATE DOCUMENT
    # =====================================================
    doc = Document()

    style = doc.styles["Normal"]

    style.font.name = "Times New Roman"

    style.font.size = Pt(11)

    # =====================================================
    # LOOP THROUGH EACH SESSION
    # =====================================================
    for (facilitator, session), group in grouped:

        # -------------------------------------------------
        # HEADER
        # -------------------------------------------------
        doc.add_paragraph("KSG/17/FER/08")
        doc.add_paragraph("KENYA SCHOOL OF GOVERNMENT")
        doc.add_paragraph("MATUGA")
        doc.add_heading(
            "FACILITATOR EVALUATION REPORT",
            level=1
        )

        # -------------------------------------------------
        # DETAILS TABLE
        # -------------------------------------------------
        details_table = doc.add_table(
            rows=3,
            cols=2
        )

        details_table.cell(0, 0).text = "PROGRAM TITLE:"
        details_table.cell(0, 1).text = programme_title

        details_table.cell(1, 0).text = "SESSION'S TOPIC:"
        details_table.cell(1, 1).text = session

        details_table.cell(2, 0).text = f"FACILITATOR: {facilitator}"
        details_table.cell(2, 1).text = f"DURATION: {duration}"

        set_table_borders(details_table)

        doc.add_paragraph()

        # -------------------------------------------------
        # INTRODUCTION
        # -------------------------------------------------
        doc.add_paragraph(
            "This report provides information to KSG management for decision making and action. "
            "The Facilitator Evaluation forms are filled by participants during the course of "
            "training programmes. The Head of Department – Training is expected to discuss the "
            "evaluation results with individual facilitators where necessary."
        )

        doc.add_paragraph()

        doc.add_heading(
            "I. Participants' Ratings",
            level=2
        )

        # -------------------------------------------------
        # RATINGS TABLE
        # -------------------------------------------------
        table = doc.add_table(
            rows=1,
            cols=10
        )

        headers = [

            "SPECIFIC ASPECTS",

            "Total no of participants",

            "Non response",

            "5",

            "4",

            "3",

            "2",

            "1",

            "Total valid responses",

            "% of Scores 4 & 5"

        ]

        for i, header in enumerate(headers):

            table.rows[0].cells[i].text = header

        # -------------------------------------------------
        # POPULATE TABLE
        # -------------------------------------------------
        for col in rating_cols:

            counts = group[col].value_counts().to_dict()

            count_5 = counts.get(5, 0)
            count_4 = counts.get(4, 0)
            count_3 = counts.get(3, 0)
            count_2 = counts.get(2, 0)
            count_1 = counts.get(1, 0)

            total_valid = (

                count_5 +

                count_4 +

                count_3 +

                count_2 +

                count_1

            )

            non_response = (

                total_participants -

                total_valid

            )

            percent_45 = (

                ((count_5 + count_4) / total_valid * 100)

                if total_valid > 0

                else 0

            )

            row = table.add_row().cells

            values = [

                col,

                total_participants,

                non_response,

                count_5,

                count_4,

                count_3,

                count_2,

                count_1,

                total_valid,

                round(percent_45, 1)

            ]

            for i, value in enumerate(values):

                row[i].text = str(value)

        set_table_borders(table)

        doc.add_paragraph()
        # =====================================================
        # QUALITATIVE (LLM ENHANCED)
        # =====================================================
        likes_raw = "; ".join(

            group["Like"]

            .dropna()

            .astype(str)

        )

        suggestions_raw = "; ".join(

            group["Suggestions"]

            .dropna()

            .astype(str)

        )

        combined_text = f"""
Most Liked:
{likes_raw}

Suggestions:
{suggestions_raw}
"""

        qualitative = analyze_qualitative(combined_text)

        paragraphs = qualitative.split("\n\n")

        doc.add_paragraph("Most liked about the facilitator")

        if len(paragraphs) > 0:

            doc.add_paragraph(paragraphs[0])

        else:

            doc.add_paragraph("No comments provided.")

        doc.add_paragraph()

        doc.add_paragraph("Suggestions on areas of improvement")

        if len(paragraphs) > 1:

            doc.add_paragraph(paragraphs[1])

        else:

            doc.add_paragraph(
                "Participants expressed minimal suggestions for improvement."
            )

        # =====================================================
        # HEAD OF DEPARTMENT COMMENTS
        # =====================================================
        doc.add_paragraph()

        doc.add_paragraph(
            "II. Head of Department – Training's comments:"
        )

        doc.add_paragraph(
            "................................................................................................................................."
        )

        doc.add_paragraph()

        doc.add_paragraph(
            "III. Head of Department – Training's proposals or recommendations:"
        )

        doc.add_paragraph(
            "................................................................................................................................."
        )

        # =====================================================
        # PAGE BREAK
        # =====================================================
        doc.add_page_break()

    # =========================================================
    # SAVE REPORT
    # =========================================================
    base_name = os.path.splitext(

        os.path.basename(cleaned_file)

    )[0]

    output_file = os.path.join(

        output_folder,

        f"{base_name}_FE_Report_LLM.docx"

    )

    doc.save(output_file)

    print(f"✅ LLM Report saved: {output_file}")

    return output_file