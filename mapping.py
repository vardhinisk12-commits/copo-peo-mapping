import pandas as pd

similarity = pd.read_csv("Dataset/co_po_similarity.csv", index_col=0)

mapping = pd.DataFrame(
    0,
    index=similarity.index,
    columns=similarity.columns
)

for co in similarity.index:

    top3 = similarity.loc[co].sort_values(ascending=False).head(3)

    mapping.loc[co, top3.index[0]] = 3
    mapping.loc[co, top3.index[1]] = 2
    mapping.loc[co, top3.index[2]] = 1

mapping.to_csv("Dataset/co_po_mapping.csv")

print(mapping)