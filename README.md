# Dependencies

This project uses [uv](https://github.com/astral-sh/uv) as package manager.

To create vector databases [BGE-M3](https://huggingface.co/BAAI/bge-m3).

To store vector databases [faiss](https://github.com/facebookresearch/faiss).

For documentation loading [LangChain Document Loader](https://python.langchain.com/v0.1/docs/modules/data_connection/document_loaders/).

Optional Dependency [Marker](https://github.com/VikParuchuri/marker) is used to convert pdf -> md.

## Install Marker (Optional)
    `uv pip install marker-pdf`
    
If conversion from any other format is required.
    
    `pip install marker-pdf[full]` 


### Convert docs.
To convert all the pdfs from sourcesPDF to md (marker will create a sub dir in "/" for each pdf
It uses GPU to render all images. (might be slow on low end pcs)
        
    `marker sourcesPDF/ --output_dir ./  --disable_image_extraction --debug`  

Move the .md file to `source_files`

Other  [Marker Flags](https://github.com/VikParuchuri/marker/blob/master/README.md)



## Quick Start
### Clone Repo
    `git clone https://github.com/Zeal5/rag.git`

### Install uv
Linux | macOS   

    `curl -LsSf https://astral.sh/uv/install.sh | sh`

Windows 

    `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`
    
From PyPi

    `pip install uv`
    
### Install Rag
1. **Create a virtual environment**
    `uv venv`.

3. **Install dependencies**
    `uv sync`
   
This will:
- Create a fresh virtual environment (`.venv` by default)
- Install all required dependencies from `pyproject.toml` and `uv.lock`

    
There are 2 main Functions in [main.py](/main.py).  
- create_vector_db_from_text_file:  

        Inputs:  
       - source_file_dir:str # dir str path to where .md | .txt files are stored  
       - vector_db_dir:str   # default is default.index  

        Outputs:  
       - None
  
- get_knn_chunks:  

        Inputs:  
       - source_file_dir:str # dir str path to where .md | .txt files are stored  
       - vector_db_dir:str   # default is default.index  
       - query : str         # query to match against db  
       - k :int              # k - nearest results to output  
       
        Output:  
       - None  
   
   
    

        
        

