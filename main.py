from src.load_docs import create_vector_db, get_document_chunks
from src.retriever import get_related_chunks
from FlagEmbedding import BGEM3FlagModel
from pathlib import Path
from pprint import pprint


MODEL = BGEM3FlagModel("BAAI/bge-m3", use_fp16=True)


def get_knn_chunks(source_file_dir: str, vector_db_dir: str, query, threshold=0.7, k=3):
    chunks = get_document_chunks(source_file_dir)
    D, I = get_related_chunks(MODEL, str(vector_db_dir), query, k=k)
    result = [
        {
            "text": chunks[I[0][i]],
            "distance": D[0][i],
            "index": I[0][i],
        }
        for i in range(len(I[0]))
        if D[0][i] < threshold
    ]
    return result or []


def create_vector_db_from_text_file(source_file_path: str, vector_db_path: str):
    model = BGEM3FlagModel("BAAI/bge-m3", use_fp16=True)
    create_vector_db(model, source_file_path, vector_db_path)


if __name__ == "__main__":
    root_dir = Path(__file__).resolve(strict=True).parent

    source_file_dir = str(root_dir / "source_files")
    vector_db_dir = str(root_dir / "vector_dbs")

    # get_knn_chunks(source_file_dir, vector_db_dir, "How to hide columns", k=5)
    create_vector_db_from_text_file(source_file_dir , vector_db_dir)
