from langchain_community.document_loaders import TextLoader
from langchain.text_splitter import (
    RecursiveCharacterTextSplitter,
    MarkdownHeaderTextSplitter,
)
from FlagEmbedding import BGEM3FlagModel
import numpy as np
import faiss
import json, os
from datetime import datetime
import shutil
from typing import List
from langchain.schema import Document


def _load_files_from_dir(source_file_dir: str) -> List[Document]:
    docs = []
    sorted_file_names = sorted(os.listdir(source_file_dir))
    for file_name in sorted_file_names:
        if file_name.endswith(".md"):
            file_path = os.path.join(source_file_dir, file_name)
            loader = TextLoader(file_path)
            documents = loader.load()
            docs.extend(documents)

    return docs


def _create_chunks(documents: List[Document]) -> List[str]:
    def remove_empty_chunks(chunks):
        chunk_json = [json.loads(chunk.json()) for chunk in chunks]
        chunk_list = [
            chunk["page_content"]
            for chunk in chunk_json
            if chunk["page_content"].strip()
        ]
        return chunk_list

    headers_to_split_on = [
        ("#", "Header 1"),
        ("##", "Header 2"),
        ("###", "Header 3"),
        ("####", "Header 4"),
        ("#####", "Header 5"),
        ("######", "Header 6"),
    ]

    markdown_splitter = MarkdownHeaderTextSplitter(
        headers_to_split_on=headers_to_split_on, return_each_line=False
    )
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=500,
        chunk_overlap=100,
        is_separator_regex=False,
        separators=["\n\n", "\n", " ", ""],  # Explicitly define splitting hierarchy
    )

    chunks = []
    for doc in documents:
        chunks.extend(markdown_splitter.split_text(doc.page_content))

    final_chunks = []
    for chunk in chunks:
        if len(chunk.page_content) > 500:
            sub_chunks = text_splitter.split_text(chunk.page_content)
            for sub_chunk in sub_chunks:
                final_chunks.append(
                    Document(page_content=sub_chunk, metadata=chunk.metadata.copy())
                )

        else:
            final_chunks.append(chunk)

    return remove_empty_chunks(final_chunks)


def _create_embeddings(model: BGEM3FlagModel, chunks: List[str], vector_db_dir: str):
    os.makedirs(vector_db_dir, exist_ok=True)
    embeddings = model.encode(
        chunks,
        batch_size=int(os.getenv("batch_size") or 12),
        max_length=int(os.getenv("max_length") or 8192),
    )["dense_vecs"]
    print(f"Creating vector db {vector_db_dir}...")

    index = faiss.IndexFlatL2(1024)
    index.add(np.array(embeddings).astype("float32"))
    created_at = datetime.now().strftime("%Y-%m-%d_%H:%M")
    faiss.write_index(index, f"{vector_db_dir}/{created_at}.index")
    shutil.copy(f"{vector_db_dir}/{created_at}.index", f"{vector_db_dir}/latest.index")


def get_document_chunks(source_file_dir: str):
    documents_list = _load_files_from_dir(source_file_dir)
    chunks = _create_chunks(documents_list)

    return chunks


def create_vector_db(model: BGEM3FlagModel, source_file_dir, vector_db_dir):

    chunks = get_document_chunks(source_file_dir)
    _create_embeddings(model, chunks, vector_db_dir)
