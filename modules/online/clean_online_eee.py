import pandas as pd
import os

# =========================================================
# INPUT FILE
# =========================================================
file_name = input("Enter Online EEE CSV file: ").strip()

# =========================================================
# LOAD CSV
# =========================================================
df = pd.read_csv(file_name)

print("\nLoaded file successfully.")

# =========================================================
# REMOVE EMPTY ROWS/COLUMNS
# =========================================================
df.dropna(how="all", inplace=True)
df.dropna(axis=1, how="all", inplace=True)

# =========================================================
# REMOVE DUPLICATE HEADER ROWS
# =========================================================
first_col = df.columns[0]

df = df[
    df[first_col].astype(str).str.strip()
    != first_col
]

# =========================================================
# STRIP COLUMN NAMES
# =========================================================
df.columns = [
    str(col).strip()
    for col in df.columns
]

# =========================================================
# REMOVE UNWANTED COLUMNS
# =========================================================
drop_cols = [
    "Timestamp",
    "Email Address",
    "Name"
    "Response number"
]

existing_drop_cols = [
    col for col in drop_cols
    if col in df.columns
]

df.drop(columns=existing_drop_cols, inplace=True)

# =========================================================
# STANDARDIZE TEXT VALUES
# =========================================================
for col in df.columns:

    if df[col].dtype == object:

        df[col] = (
            df[col]
            .astype(str)
            .str.strip()
        )

# =========================================================
# LIKERT SCALE MAPPING
# =========================================================
likert_map = {

    # Quality Scale
    "Excellent": 5,
    "Very Good": 4,
    "Good": 3,
    "Neutral": 3,
    "Satisfactory": 3,
    "Fair": 2,
    "Poor": 2,
    "Very Poor": 1,

    # Extent Scale
    "Great Extent": 5,
    "To Some Extent": 4,
    "Moderate Extent": 3,
    "Not Sure": 2,
    "Not At All": 1,
    "Not at All": 1
}

# =========================================================
# DETECT QUALITATIVE COLUMNS
# =========================================================
qualitative_keywords = [

    "suggest",
    "comment",
    "additional",
    "topic",
    "interest",
    "improve",
    "experience"
]

qualitative_cols = []

for col in df.columns:

    col_lower = str(col).lower()

    if any(
        keyword in col_lower
        for keyword in qualitative_keywords
    ):

        qualitative_cols.append(col)

# =========================================================
# CONVERT ONLY RATING COLUMNS
# =========================================================
rating_cols = [
    col for col in df.columns
    if col not in qualitative_cols
]

for col in rating_cols:

    try:

        df[col] = (
            df[col]
            .replace(likert_map)
        )

    except:
        pass

# =========================================================
# CONVERT NUMERIC COLUMNS
# =========================================================

for col in rating_cols:

    df[col] = pd.to_numeric(
        df[col],
        errors="coerce"
    )
# =========================================================
# SAVE CLEANED FILE
# =========================================================
base_name = os.path.splitext(
    os.path.basename(file_name)
)[0]

output_file = f"{base_name}_cleaned.xlsx"

df.to_excel(
    output_file,
    index=False
)

# =========================================================
# SUMMARY
# =========================================================
print("\n===================================")
print("ONLINE EEE CLEANING COMPLETE")
print("===================================")

print(f"\nRows: {df.shape[0]}")
print(f"Columns: {df.shape[1]}")

print("\nQualitative Columns Detected:")
for col in qualitative_cols:
    print(f"- {col}")

print(f"\nSaved cleaned file as:")
print(output_file)