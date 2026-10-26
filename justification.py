import pandas as pd

# Read CO-PO Mapping
mapping = pd.read_csv("Dataset/co_po_mapping.csv")

justification = []

for _, row in mapping.iterrows():

    co = row.iloc[0]

    for col in mapping.columns[1:]:

        value = row[col]

        if value == 3:
            text = "High semantic similarity between CO and " + col
        elif value == 2:
            text = "Medium semantic similarity between CO and " + col
        elif value == 1:
            text = "Low semantic similarity between CO and " + col
        else:
            text = "No significant relationship"

        justification.append([co, col, value, text])

# Create DataFrame
justification_df = pd.DataFrame(
    justification,
    columns=[
        "CO ID",
        "PO ID",
        "Mapping Value",
        "Justification"
    ]
)

# Save CSV
justification_df.to_csv(
    "Dataset/co_po_justification.csv",
    index=False
)

print("CO-PO Justification Report Created Successfully!")