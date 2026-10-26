import pandas as pd
import re
import string

# -----------------------------
# Function to clean text
# -----------------------------
def clean_text(text):
    if pd.isna(text):
        return ""

    text = str(text).lower()
    text = text.translate(str.maketrans('', '', string.punctuation))
    text = re.sub(r'\s+', ' ', text).strip()

    return text

# -----------------------------
# Read Excel Files
# -----------------------------
print("Reading Dataset...")

subjects = pd.read_excel("Dataset/Subject_Details.xlsx")
co = pd.read_excel("Dataset/CO_Dataset-1.xlsx")
po = pd.read_excel("Dataset/PO_Dataset.xlsx")
peo = pd.read_excel("Dataset/PEO_Dataset.xlsx")

# -----------------------------
# Display Dataset Information
# -----------------------------
print("\nSubject Details")
print(subjects.head())

print("\nCourse Outcomes")
print(co.head())

print("\nProgram Outcomes")
print(po.head())

print("\nProgram Educational Objectives")
print(peo.head())

# -----------------------------
# Remove Missing Values
# -----------------------------
subjects = subjects.fillna("")
co = co.fillna("")
po = po.fillna("")
peo = peo.fillna("")

# -----------------------------
# Remove Duplicate Rows
# -----------------------------
subjects = subjects.drop_duplicates()
co = co.drop_duplicates()
po = po.drop_duplicates()
peo = peo.drop_duplicates()

# -----------------------------
# Clean All Text Columns
# -----------------------------
for column in subjects.columns:
    if subjects[column].dtype == "object":
        subjects[column] = subjects[column].apply(clean_text)

for column in co.columns:
    if co[column].dtype == "object":
        co[column] = co[column].apply(clean_text)

for column in po.columns:
    if po[column].dtype == "object":
        po[column] = po[column].apply(clean_text)

for column in peo.columns:
    if peo[column].dtype == "object":
        peo[column] = peo[column].apply(clean_text)

# -----------------------------
# Save Preprocessed CSV Files
# -----------------------------
subjects.to_csv("Dataset/subjects_preprocessed.csv", index=False)

co.to_csv("Dataset/co_preprocessed.csv", index=False)

po.to_csv("Dataset/po_preprocessed.csv", index=False)

peo.to_csv("Dataset/peo_preprocessed.csv", index=False)

# -----------------------------
# Display Summary
# -----------------------------
print("\n========== Dataset Summary ==========")
print("Subjects :", len(subjects))
print("CO :", len(co))
print("PO :", len(po))
print("PEO :", len(peo))

print("\nPreprocessing Completed Successfully!")

print("\nCSV Files Generated:")
print("subjects_preprocessed.csv")
print("co_preprocessed.csv")
print("po_preprocessed.csv")
print("peo_preprocessed.csv")