import pandas as pd
import os

# ===== INPUT =====
file_name = input("Enter cleaned CE file name: ").strip()
df = pd.read_excel(file_name)

print("Loaded file:", file_name)
print("Shape:", df.shape)

# ===== EXTRACT DETAILS =====
program_title = df["Program Title"].iloc[0] if "Program Title" in df.columns else "N/A"
coordinator = df["Coordinator Name"].iloc[0] if "Coordinator Name" in df.columns else "N/A"

# ===== IDENTIFY RATING COLUMNS =====
rating_cols = [
    col for col in df.columns
    if df[col].dtype in ["int64", "float64", "Int64"]
]

exclude_cols = ["Timetable No"]
rating_cols = [col for col in rating_cols if col not in exclude_cols]

# ===== ANALYSIS =====
results = []

for col in rating_cols:

    counts = df[col].value_counts().to_dict()

    count_5 = counts.get(5, 0)
    count_4 = counts.get(4, 0)
    count_3 = counts.get(3, 0)
    count_2 = counts.get(2, 0)
    count_1 = counts.get(1, 0)

    total_valid = count_5 + count_4 + count_3 + count_2 + count_1

    if total_valid > 0:
        p5 = round((count_5 / total_valid) * 100, 1)
        p4 = round((count_4 / total_valid) * 100, 1)
        p3 = round((count_3 / total_valid) * 100, 1)
        p2 = round((count_2 / total_valid) * 100, 1)
        p1 = round((count_1 / total_valid) * 100, 1)
    else:
        p5 = p4 = p3 = p2 = p1 = 0

    results.append([
        col,
        p5,
        p4,
        p3,
        p2,
        p1
    ])

# ===== TABLE =====
df_out = pd.DataFrame(results, columns=[
    "Specific Aspects",
    "Excellent % : 5",
    "Very Good % : 4",
    "Good % : 3",
    "Fair % : 2",
    "Poor % : 1"
])

# ===== QUALITATIVE =====
likes = "; ".join(df["Like"].dropna().astype(str)) if "Like" in df.columns else "No responses"
suggestions = "; ".join(df["Suggestions"].dropna().astype(str)) if "Suggestions" in df.columns else "No responses"

# ===== SAVE =====
base_name = os.path.splitext(file_name)[0]
output_file = f"{base_name}_ce_analysis.xlsx"

if os.path.exists(output_file):
    output_file = f"{base_name}_ce_analysis_new.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

    # Start writing manually using row positions
    start_row = 0

    # ===== HEADER INFO =====
    header_df = pd.DataFrame({
        "A": [
            f"Program Title: {program_title}",
            f"Coordinator Name: {coordinator}"
        ]
    })

    header_df.to_excel(writer, sheet_name="Analysis", index=False, header=False, startrow=start_row)

    # ===== ANALYSIS TABLE =====
    table_start = start_row + 4  # leave 2 blank rows

    df_out.to_excel(writer, sheet_name="Analysis", index=False, startrow=table_start)

    # ===== QUALITATIVE =====
    qual_start = table_start + len(df_out) + 3

    worksheet = writer.sheets["Analysis"]

    worksheet.cell(row=qual_start, column=1).value = "Most Liked:"
    worksheet.cell(row=qual_start + 1, column=1).value = likes

    worksheet.cell(row=qual_start + 3, column=1).value = "Suggestions:"
    worksheet.cell(row=qual_start + 4, column=1).value = suggestions

print(f"\n✅ CE analysis saved as: {output_file}")