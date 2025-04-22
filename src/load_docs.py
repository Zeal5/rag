from langchain_community.document_loaders import TextLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from FlagEmbedding import BGEM3FlagModel
import numpy as np
import faiss
import json, os
from pathlib import Path
from datetime import datetime
import shutil


def _load_files_from_dir(source_file_dir: str):
    docs = []
    sorted_file_names = sorted(os.listdir(source_file_dir))
    for file_name in sorted_file_names:
        if file_name.endswith(".md"):
            file_path = os.path.join(source_file_dir, file_name)
            loader = TextLoader(file_path)
            documents = loader.load()
            docs.extend(documents)

    return docs


def _create_chunks(documents):
    def remove_empty_chunks(chunks):
        chunk_json = [json.loads(chunk.json()) for chunk in chunks]
        chunk_json = [
            chunk["page_content"]
            for chunk in chunk_json
            if chunk["page_content"].strip()
        ]

        return chunk_json

    splitter = RecursiveCharacterTextSplitter(chunk_size=500, chunk_overlap=50)
    chunks = splitter.split_documents(documents)

    return remove_empty_chunks(chunks)


def _create_embeddings(model: BGEM3FlagModel, chunks, vector_db_dir: str):
    os.makedirs(vector_db_dir, exist_ok=True)
    embeddings = model.encode(
        chunks,
        batch_size=int(os.getenv("batch_size") or 12),
        max_length=int(os.getenv("max_length") or 8192),
    )["dense_vecs"]
    root = Path(__file__).resolve(strict=True).parent.parent
    print(f"Creating vector db {vector_db_dir}...")

    index = faiss.IndexFlatL2(1024)
    index.add(np.array(embeddings).astype("float32"))
    created_at = datetime.now().strftime("%Y-%m-%d_%H:%M")
    faiss.write_index(index, f"{vector_db_dir}/{created_at}.index")
    shutil.copy(f"{vector_db_dir}/{created_at}.index", f"{vector_db_dir}/latest.index")


def get_document_chunks(source_file_dir: str):
    doc = _load_files_from_dir(source_file_dir)
    chunks = _create_chunks(doc)

    return chunks


def create_vector_db(model: BGEM3FlagModel, source_file_dir, vector_db_dir):

    chunks = get_document_chunks(source_file_dir)
    _create_embeddings(model, chunks, vector_db_dir)
