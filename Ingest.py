# ingest.py – loads PPTX files and stores embeddings in ./vector_db
import sys, pathlib, datetime
from langchain_community.document_loaders import UnstructuredPowerPointLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_openai import OpenAIEmbeddings
from langchain_community.vectorstores import Chroma

if len(sys.argv) < 2:
    print("Usage: python ingest.py path\\to\\file1.pptx [file2.pptx …]")
    sys.exit(1)

splitter = RecursiveCharacterTextSplitter(chunk_size=800, chunk_overlap=100)
embeddings = OpenAIEmbeddings(model="text-embedding-3-small")

all_docs = []
for pptx_path in sys.argv[1:]:
    pptx_path = pathlib.Path(pptx_path)
    if not pptx_path.exists():
        print(f"⚠  {pptx_path} not found – skipping")
        continue
    print(f"⏳  Loading {pptx_path.name}")
    docs = UnstructuredPowerPointLoader(str(pptx_path)).load()
    for d in docs:
        d.metadata.update(
            source=str(pptx_path.name),
            date=datetime.date.today().isoformat()
        )
    all_docs.extend(docs)

if not all_docs:
    print("No documents ingested.")
    sys.exit(0)

chunks = splitter.split_documents(all_docs)
db = Chroma.from_documents(chunks, embeddings, persist_directory="vector_db")
db.persist()
print(f"✅  Ingested {len(chunks)} chunks → ./vector_db")
