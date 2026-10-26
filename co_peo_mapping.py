import pandas as pd
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

print("=" * 50)
print("AI-Based CO-PEO Mapping")
print("=" * 50)

# Load AI model
print("\nLoading Sentence-BERT Model...")
model = SentenceTransformer("all-MiniLM-L6-v2")

# Read datasets
print("\nReading datasets...")

co = pd.read_csv("Dataset/co_preprocessed.csv")
peo = pd.read_csv("Dataset/peo_preprocessed.csv")

# Create Unique CO ID
co["Unique_CO"] = co["Subject Code"].astype(str) + "_" + co["CO ID"].astype(str)

# Read description columns
co_text = co.iloc[:, -2].fillna("").astype(str)
peo_text = peo.iloc[:, -1].fillna("").astype(str)

print("\nGenerating Embeddings...")

co_embeddings = model.encode(
    co_text.tolist(),
    show_progress_bar=True
)

peo_embeddings = model.encode(
    peo_text.tolist(),
    show_progress_bar=True
)

print("\nCalculating Similarity...")

similarity_matrix = cosine_similarity(
    co_embeddings,
    peo_embeddings
)

# Create similarity DataFrame
similarity_df = pd.DataFrame(
    similarity_matrix,
    index=co["Unique_CO"],
    columns=peo.iloc[:, 0]
)

# Create mapping DataFrame
mapping = pd.DataFrame(
    0,
    index=similarity_df.index,
    columns=similarity_df.columns
)

# Top-3 Mapping
for co_id in similarity_df.index:

    top = similarity_df.loc[co_id].sort_values(ascending=False)

    mapping.loc[co_id, top.index[0]] = 3

    if len(top) > 1:
        mapping.loc[co_id, top.index[1]] = 2

    if len(top) > 2:
        mapping.loc[co_id, top.index[2]] = 1

# Save files
similarity_df.to_csv("Dataset/co_peo_similarity.csv")
mapping.to_csv("Dataset/co_peo_mapping.csv")

print("\nSimilarity Matrix")
print(similarity_df)

print("\nCO-PEO Mapping Matrix")
print(mapping)

print("\nFiles Created Successfully!")
print("1. Dataset/co_peo_similarity.csv")
print("2. Dataset/co_peo_mapping.csv")