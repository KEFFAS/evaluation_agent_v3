import pandas as pd
import os
from openpyxl.styles import Font

# ===== USER INPUT =====
file_name = input("Enter Excel file name (e.g. RPP 25 CE.xlsx): ").strip()
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

# ===== FORCE INTEGER FORMAT =====
for col in df.select_dtypes(include='number').columns:
    if (df[col].dropna() % 1 == 0).all():
        df[col] = df[col].astype('Int64')

# ===== REMOVE DUPLICATES =====
#df = df.drop_duplicates()

# ===== SET PROGRAM TITLE =====
df["Program Title"] = program_title

# ===== DERIVE COORDINATOR NAME FROM EMAIL =====
for col in df.columns:
    if "email" in col.lower():
        df[col] = df[col].astype(str)

        df["Coordinator Name"] = df[col].apply(lambda x:
            x.split("@")[0]
            .replace(".", " ")
            .replace("_", " ")
            .replace("-", " ")
            .title()
            if "@" in x else ""
        )
        break

# ===== STANDARD CE RATING MAP =====
ce_rating_map = {
    "Organization of program opening and closing": "Organization Of Program Opening And Closing",
    "Briefing participants and orientation": "Briefing Participants And Orientation",
    "Leveling of participant expectations": "Leveling Of Participant Expectations",
    "Communication and provision of feedback to participants": "Communication And Feedback",
    "Management of program timetable and facilitators": "Management Of Timetable And Facilitators",
    "Monitoring participants’ attendance": "Monitoring Participants Attendance",
    "Program evaluation": "Program Evaluation",
    "Action planning": "Action Planning",
    "General administration of the program": "General Administration Of Program"
}

# ===== RENAME LONG QUESTION COLUMNS =====
new_columns = {}
for col in df.columns:
    col_str = str(col)

    for key in ce_rating_map:
        if key.lower() in col_str.lower():
            new_columns[col] = ce_rating_map[key]

df = df.rename(columns=new_columns)

# ===== DETECT & FIX QUALITATIVE =====
qualitative_cols = []

for col in df.columns:
    col_clean = str(col).strip().lower()

    if "like" in col_clean:
        df = df.rename(columns={col: "Like"})
        qualitative_cols.append("Like")

    elif "suggest" in col_clean:
        df = df.rename(columns={col: "Suggestions"})
        qualitative_cols.append("Suggestions")

# Remove duplicates
qualitative_cols = list(dict.fromkeys(qualitative_cols))

# ===== DEFINE RATING ORDER =====
rating_order = list(ce_rating_map.values())

# ===== IDENTIFY OTHER COLUMNS =====
other_cols = [
    col for col in df.columns
    if col not in rating_order + qualitative_cols
]

# ===== FINAL ORDER =====
final_cols = (
    other_cols +
    [col for col in rating_order if col in df.columns] +
    [col for col in qualitative_cols if col in df.columns]
)

df = df[final_cols]

print("Cleaned shape:", df.shape)

# ===== SAVE FILE =====
base_name = os.path.splitext(file_name)[0]
output_file = f"{base_name}_ce_cleaned.xlsx"

if os.path.exists(output_file):
    output_file = f"{base_name}_ce_cleaned_new.xlsx"

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Cleaned Data')

    worksheet = writer.sheets['Cleaned Data']

    # Bold headers
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

print("CE Cleaned file saved as:", output_file)