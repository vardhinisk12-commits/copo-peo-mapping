import pandas as pd
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

print("=" * 50)
print("AI-Based CO-PO Semantic Similarity")
print("=" * 50)

# Load AI Model
print("\nLoading Sentence-BERT Model...")
model = SentenceTransformer("all-MiniLM-L6-v2")

# Read datasets
print("\nReading datasets...")

co = pd.read_csv("Dataset/co_preprocessed.csv")
po = pd.read_csv("Dataset/po_preprocessed.csv")

# Display dataset information
print("\nCO Dataset Shape :", co.shape)
print("PO Dataset Shape :", po.shape)

# Create Unique CO ID (SubjectCode_COID)
co["Unique_CO"] = co["Subject Code"].astype(str) + "_" + co["CO ID"].astype(str)

# Read description columns
co_text = co.iloc[:, -2].fillna("").astype(str)
po_text = po.iloc[:, -1].fillna("").astype(str)

print("\nGenerating Embeddings...")

co_embeddings = model.encode(
    co_text.tolist(),
    show_progress_bar=True
)

po_embeddings = model.encode(
    po_text.tolist(),
    show_progress_bar=True
)

print("\nCalculating Cosine Similarity...")

similarity_matrix = cosine_similarity(
    co_embeddings,
    po_embeddings
)

# Row names = Unique CO IDs
co_ids = co["Unique_CO"]

# Column names = PO IDs
po_ids = po["PO ID"]

similarity_df = pd.DataFrame(
    similarity_matrix,
    index=co_ids,
    columns=po_ids
)

print("\nSimilarity Matrix")
print(similarity_df)

# Save CSV
similarity_df.to_csv("Dataset/co_po_similarity.csv")

print("\nSimilarity matrix saved successfully!")
print("Output File : Dataset/co_po_similarity.csv")

print("\nBest Matching PO for each CO")

for index, row in similarity_df.iterrows():
    highest_po = row.idxmax()
    highest_score = row.max()

    print(f"{index} --> {highest_po} : {highest_score:.4f}")

print("\nProcess Completed Successfully!")