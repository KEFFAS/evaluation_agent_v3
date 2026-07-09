import pandas as pd
import os
from openpyxl.styles import Font

# ===== USER INPUT =====
file_name = input("Enter Excel file name (e.g. RPP 25 FE.xlsx): ").strip()
program_title = input("Enter Program Title: ").strip()

# ===== LOAD FILE =====
df = pd.read_excel(file_name, header=None)

print("Original shape:", df.shape)

# ===== REMOVE FIRST ROW =====
df = df.iloc[1:].reset_index(drop=True)

# ===== SET HEADER =====
df.columns = df.iloc[0]
df = df[1:].reset_index(drop=True)

# ===== REMOVE EMPTY ROWS =====
df = df.dropna(how='all')

# ===== REMOVE 'UNNAMED' COLUMNS =====
df = df.loc[:, ~df.columns.astype(str).str.contains('unnamed', case=False)]

# ===== CLEAN COLUMN NAMES =====
df.columns = (
    df.columns
    .astype(str)
    .str.strip()
    .str.title()
)

# ===== CLEAN TEXT VALUES =====
for col in df.columns:
    if df[col].dtype == "object":
        df[col] = df[col].astype(str).str.strip()

# ===== CONVERT NUMERIC COLUMNS =====
for col in df.columns:
    converted = pd.to_numeric(df[col], errors='coerce')
    if converted.notna().sum() > len(df) * 0.5:
        df[col] = converted

# ===== FORCE INTEGER FORMAT WHERE POSSIBLE =====
for col in df.select_dtypes(include='number').columns:
    if (df[col].dropna() % 1 == 0).all():
        df[col] = df[col].astype('Int64')

# ===== REMOVE DUPLICATES =====
#df = df.drop_duplicates()

# ===== SET PROGRAM TITLE =====
df["Program Title"] = program_title

# ===== FIX COMMON COLUMN NAMES =====
rename_map = {
    "Programme Title": "Program Title",
    "Facilitator": "Lecturer Name",
    "Lecturer": "Lecturer Name",
    "Topic": "Topic Description"
}
df = df.rename(columns=rename_map)

# ===== REORDER COLUMNS =====
desired_order = [
    "Date",
    "Program Title",
    "Topic Description",
    "Lecturer Name",
    "Punctuality",
    "Presentation Flow",
    "Handling Questions",
    "Active Participation Of Learners",
    "Use Of Visual Aids",
    "Relevance Of Subject To Workplace",
    "Use Of Relevant Examples",
    "Knowledge Of Subject",
    "Treats Participants With Dignity And Respect",
    "Variety And Appropriateness Of Training Methods",
    "Like",
    "Suggestions",
    "Status",
    "Session Code",
    "Timetable No",
    "Campus"
]

existing_cols = [col for col in desired_order if col in df.columns]
remaining_cols = [col for col in df.columns if col not in existing_cols]

df = df[existing_cols + remaining_cols]

# ===== SORT BY TOPIC DESCRIPTION =====
if "Topic Description" in df.columns:
    df = df.sort_values(by="Topic Description").reset_index(drop=True)

print("Cleaned shape:", df.shape)

# ===== SAVE FILE =====
base_name = os.path.splitext(file_name)[0]
output_file = f"{base_name}_cleaned.xlsx"

# Avoid overwrite error
if os.path.exists(output_file):
    output_file = f"{base_name}_cleaned_new.xlsx"

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Cleaned Data')

    worksheet = writer.sheets['Cleaned Data']

    # Bold headers
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

print("✅ Cleaned file saved as:", output_file)