import numpy as np
from FlagEmbedding import BGEM3FlagModel
import faiss, os


def get_related_chunks(
    model: BGEM3FlagModel,
    vector_db_dir,
    query: str,
    k: int = 3,
    vector_db_name: str = "",
):

    if not vector_db_name:
        vector_db_name = os.getenv("vector_db_name") or "latest.index"

    vector_db_name = os.path.join(vector_db_dir, vector_db_name)
    index = faiss.read_index(vector_db_name)
    query_embedding = model.encode(
        query,
        batch_size=int(os.getenv("batch_size") or 12),
        max_length=int(os.getenv("max_length") or 8192),
    )["dense_vecs"]
    query_embedding = np.array(query_embedding).astype("float32").reshape(1, -1)

    D, I = index.search(np.array(query_embedding).astype("float32"), k=k)
    return D, I
