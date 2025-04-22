# Dependencies

This project uses [uv](https://github.com/astral-sh/uv) as package manager.
To create vector databases [BGE-M3](https://huggingface.co/BAAI/bge-m3)
To store vector databases [faiss](https://github.com/facebookresearch/faiss)
For documentation loadind [LangChain Document Loader](https://python.langchain.com/v0.1/docs/modules/data_connection/document_loaders/)
Optional Dependency [Marker](https://github.com/VikParuchuri/marker) Is used to convert pdf -> md.

## Install Marker:
    `uv pip install marker-pdf`
    `pip install marker-pdf[full]` (if conversion from any other format is reuired)
    
    Convert docs.
        to convert all the pdfs from sourcesPDF to md (marker will create a sub dir in "/" for each pdf
        Move the .md file to `source_files`
        `marker sourcesPDF/ --output_dir ./  --disable_image_extraction --debug`

    It uses GPU to render all images other [Marker  Flags](https://github.com/VikParuchuri/marker/blob/master/README.md)



## Quick Start
### Clone Repo
    `git clone https://github.com/Zeal5/rag.git`

### Install uv
    Linux | macOS `curl -LsSf https://astral.sh/uv/install.sh | sh`
    Windows `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`
    From PyPi
    `pip install uv`
    
### Install Rag
1. **Create a virtual environment**:
`uv venv`

2. **Install dependencies**:
`uv sync`
This will:
- Create a fresh virtual environment (`.venv` by default)
- Install all required dependencies from `pyproject.toml` and `uv.lock`

    

        
        

