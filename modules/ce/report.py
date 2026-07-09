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
# TABLE BORDERS
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
The following comments were provided by participants during evaluation of programme coordination and administration at the Kenya School of Government.

Participant Feedback

{text}

Write exactly TWO concise professional paragraphs.

Paragraph 1
Summarize the most recurring positive feedback regarding programme coordination, administration, communication, organization, participant support and overall management of the programme.

Paragraph 2
Summarize the most recurring suggestions for improvement.

Requirements

- Formal institutional language
- Human tone
- Evidence based
- No bullet points
- No headings
- Avoid repetition
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
def generate_ce_report(

    cleaned_file,

    programme_title,

    programme_code,

    duration,

    venue,

    coordinator,

    assistant,

    output_folder="outputs"

):

    os.makedirs(output_folder, exist_ok=True)

    # =====================================================
    # LOAD DATA
    # =====================================================
    df = pd.read_excel(cleaned_file)

    coordinator_name = (

        df["Coordinator Name"].iloc[0]

        if "Coordinator Name" in df.columns

        else "N/A"

    )

    # =====================================================
    # IDENTIFY RATING COLUMNS
    # =====================================================
    rating_cols = [

        col

        for col in df.columns

        if df[col].dtype in [

            "int64",

            "float64",

            "Int64"

        ]

    ]

    exclude = [

        "Timetable No"

    ]

    rating_cols = [

        col

        for col in rating_cols

        if col not in exclude

    ]

    # =====================================================
    # CREATE DOCUMENT
    # =====================================================
    doc = Document()

    style = doc.styles["Normal"]

    style.font.name = "Times New Roman"

    style.font.size = Pt(11)

    # =====================================================
    # HEADER
    # =====================================================
    doc.add_paragraph("KSG/17/FEF/09")
    doc.add_paragraph("KENYA SCHOOL OF GOVERNMENT")
    doc.add_paragraph("CAMPUS / INSTITUTE: MATUGA")
    doc.add_heading(
        "PROGRAM COORDINATOR EVALUATION REPORT",
        level=1
    )

    # =====================================================
    # DETAILS TABLE
    # =====================================================
    details_table = doc.add_table(
        rows=3,
        cols=2
    )

    details_table.cell(0, 0).text = "PROGRAMME TITLE:"
    details_table.cell(0, 1).text = programme_title

    details_table.cell(1, 0).text = "DURATION:"
    details_table.cell(1, 1).text = duration

    details_table.cell(2, 0).text = "COORDINATOR'S NAME:"
    details_table.cell(2, 1).text = coordinator_name

    # Bold first column
    for row in details_table.rows:

        for paragraph in row.cells[0].paragraphs:

            for run in paragraph.runs:

                run.bold = True

    set_table_borders(details_table)

    doc.add_paragraph()

    # =====================================================
    # INTRODUCTION
    # =====================================================
    doc.add_paragraph(

        "This report provides information to Kenya School of Government "
        "management for decision making and continuous improvement. "
        "The Coordinator Evaluation forms are completed by participants "
        "during the training programme. The findings provide useful "
        "feedback on programme coordination, administration and overall "
        "management of the programme."

    )

    doc.add_paragraph()

    doc.add_paragraph(

        "1. Ratings on the following aspects of programme coordination "
        "(5 = Excellent, 4 = Very Good, 3 = Good, 2 = Fair, 1 = Poor)."

    )

    # =====================================================
    # RATINGS TABLE
    # =====================================================
    table = doc.add_table(
        rows=1,
        cols=6
    )

    headers = [

        "Specific Aspects",

        "Excellent % : 5",

        "Very Good % : 4",

        "Good % : 3",

        "Fair % : 2",

        "Poor % : 1"

    ]

    for i, header in enumerate(headers):

        table.rows[0].cells[i].text = header

    # =====================================================
    # POPULATE TABLE
    # =====================================================
    for col in rating_cols:

        counts = df[col].value_counts().to_dict()

        count5 = counts.get(5, 0)
        count4 = counts.get(4, 0)
        count3 = counts.get(3, 0)
        count2 = counts.get(2, 0)
        count1 = counts.get(1, 0)

        total = (

            count5 +
            count4 +
            count3 +
            count2 +
            count1

        )

        if total > 0:

            p5 = round(count5 / total * 100, 1)
            p4 = round(count4 / total * 100, 1)
            p3 = round(count3 / total * 100, 1)
            p2 = round(count2 / total * 100, 1)
            p1 = round(count1 / total * 100, 1)

        else:

            p5 = p4 = p3 = p2 = p1 = 0

        row = table.add_row().cells

        row[0].text = col
        row[1].text = str(p5)
        row[2].text = str(p4)
        row[3].text = str(p3)
        row[4].text = str(p2)
        row[5].text = str(p1)

    set_table_borders(table)

    doc.add_paragraph()

    # =====================================================
    # QUALITATIVE FEEDBACK (LLM)
    # =====================================================
    likes = (

        "; ".join(

            df["Like"]

            .dropna()

            .astype(str)

        )

        if "Like" in df.columns

        else ""

    )

    suggestions = (

        "; ".join(

            df["Suggestions"]

            .dropna()

            .astype(str)

        )

        if "Suggestions" in df.columns

        else ""

    )

    combined = f"""
Most liked:
{likes}

Suggestions:
{suggestions}
"""

    analysis = analyze_qualitative(combined)

    paragraphs = analysis.split("\n\n")

    # =====================================================
    # MOST LIKED
    # =====================================================
    doc.add_paragraph()

    doc.add_paragraph(
        "2. Most liked about the programme coordination and overall administration"
    )

    if len(paragraphs) > 0:

        doc.add_paragraph(paragraphs[0])

    else:

        doc.add_paragraph("No comments provided.")

    # =====================================================
    # SUGGESTIONS
    # =====================================================
    doc.add_paragraph()

    doc.add_paragraph(
        "3. Suggestions on how the coordinator(s) can improve these aspects."
    )

    if len(paragraphs) > 1:

        doc.add_paragraph(paragraphs[1])

    else:

        doc.add_paragraph(
            "Participants expressed minimal suggestions for improvement."
        )

    # =====================================================
    # HOD RECOMMENDATIONS
    # =====================================================
    doc.add_paragraph()

    doc.add_paragraph(
        "4. Head of Department – Training's recommendations:"
    )

    doc.add_paragraph(
        "........................................................................................................................"
    )

    doc.add_paragraph()

    doc.add_paragraph(
        "Head of Department: Name………………. Signed: ………………… Date: …………………"
    )

    # =====================================================
    # SAVE REPORT
    # =====================================================
    base_name = os.path.splitext(

        os.path.basename(cleaned_file)

    )[0]

    output_file = os.path.join(

        output_folder,

        f"{base_name}_CE_Report.docx"

    )

    doc.save(output_file)

    print(f"✅ CE Report saved: {output_file}")

    return output_file



