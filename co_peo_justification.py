import pandas as pd

print("=" * 50)
print("AI-Based CO-PEO Justification")
print("=" * 50)

# Read CO-PEO Mapping
mapping = pd.read_csv("Dataset/co_peo_mapping.csv")

justification = []

for _, row in mapping.iterrows():

    co = str(row.iloc[0])

    for col in mapping.columns[1:]:

        value = int(row[col])

        if value == 3:
            text = "High semantic similarity between " + co + " and " + col
        elif value == 2:
            text = "Medium semantic similarity between " + co + " and " + col
        elif value == 1:
            text = "Low semantic similarity between " + co + " and " + col
        else:
            text = "No significant relationship"

        justification.append([co, col, value, text])

# Create DataFrame
df = pd.DataFrame(
    justification,
    columns=[
        "CO ID",
        "PEO ID",
        "Mapping Value",
        "Justification"
    ]
)

# Save CSV
df.to_csv(
    "Dataset/co_peo_justification.csv",
    index=False
)

print("CO-PEO Justification Created Successfully!")
print("Saved in Dataset/co_peo_justification.csv")