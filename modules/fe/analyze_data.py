import pandas as pd
import os
from openpyxl.styles import Font

# ===== LOAD DATA =====
file_name = input("Enter cleaned file name: ").strip()
df = pd.read_excel(file_name)

# ===== STANDARDIZE =====
df.columns = df.columns.astype(str).str.strip().str.title()

rename_map = {
    "Topic": "Topic Description",
    "Session Topic": "Topic Description",
    "Facilitator": "Lecturer Name",
    "Lecturer": "Lecturer Name"
}
df = df.rename(columns=rename_map)

# ===== VALIDATION =====
if "Topic Description" not in df.columns or "Lecturer Name" not in df.columns:
    raise ValueError("❌ Required columns missing. Use cleaned file.")

# ===== INPUT =====
total_participants = int(input("Enter total number of participants: "))

# ===== RATING COLS =====
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

rating_cols = [col for col in rating_cols if col in df.columns]

# ===== GROUP =====
grouped = df.groupby("Topic Description")

# ===== OUTPUT =====
base_name = os.path.splitext(file_name)[0]
output_file = f"{base_name}_analyzed_sessions.xlsx"

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

    for session, group in grouped:

        facilitator = group["Lecturer Name"].iloc[0]

        results = []

        for col in rating_cols:
            counts = group[col].value_counts().to_dict()

            count_5 = counts.get(5, 0)
            count_4 = counts.get(4, 0)
            count_3 = counts.get(3, 0)
            count_2 = counts.get(2, 0)
            count_1 = counts.get(1, 0)

            total_valid = count_5 + count_4 + count_3 + count_2 + count_1
            non_response = total_participants - total_valid
            percent_45 = ((count_5 + count_4) / total_valid * 100) if total_valid > 0 else 0

            results.append([
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
            ])

        result_df = pd.DataFrame(results, columns=[
            "Indicator",
            "Total no of participants",
            "Non response",
            "5",
            "4",
            "3",
            "2",
            "1",
            "Total valid responses",
            "% of Scores 4 & 5"
        ])

        # ===== QUALITATIVE JOIN =====
        likes_text = "; ".join(group["Like"].dropna().astype(str))
        suggestions_text = "; ".join(group["Suggestions"].dropna().astype(str))

        # ===== WRITE SHEET =====
        sheet_name = str(session)[:31]
        result_df.to_excel(writer, sheet_name=sheet_name, startrow=3, index=False)

        worksheet = writer.sheets[sheet_name]

        # ===== HEADER (TOP) =====
        worksheet["A1"] = f"SESSION: {session}"
        worksheet["A2"] = f"FACILITATOR: {facilitator}"

        worksheet["A1"].font = Font(bold=True)
        worksheet["A2"].font = Font(bold=True)

        # ===== AFTER TABLE =====
        last_row = len(result_df) + 5

        worksheet[f"A{last_row}"] = "MOST LIKED:"
        worksheet[f"A{last_row+1}"] = likes_text

        worksheet[f"A{last_row+3}"] = "SUGGESTIONS:"
        worksheet[f"A{last_row+4}"] = suggestions_text

        worksheet[f"A{last_row}"].font = Font(bold=True)
        worksheet[f"A{last_row+3}"].font = Font(bold=True)

print(f"\n✅ Final analyzed file saved as {output_file}")