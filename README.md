# Word to LaTeX API

Convert Word documents (`.docx`) to compilable LaTeX + BibTeX automatically using Claude AI.

## Features

- Upload a `.docx` file → receive a ZIP with `document.tex` + `references.bib`
- Preserves headings, bold, italic, tables, lists, quotes
- Auto-detects reference style (APA, IEEE, MLA, Vancouver…) and generates proper BibTeX
- Powered by `claude-opus-4-6` with adaptive thinking

## Installation

```bash
pip install -r requirements.txt
```

## Configuration

```bash
export ANTHROPIC_API_KEY="sk-ant-..."
```

## Run

```bash
uvicorn main:app --reload
```

API docs available at `http://localhost:8000/docs`

## Usage

```bash
curl -X POST http://localhost:8000/convert \
  -F "file=@document.docx" \
  -o result.zip
```

```python
import requests

with open("document.docx", "rb") as f:
    r = requests.post("http://localhost:8000/convert", files={"file": f})

with open("result.zip", "wb") as out:
    out.write(r.content)

# Check if references were found
print(r.headers.get("X-Has-Bibliography"))  # "true" or "false"
```

## Endpoints

| Method | Path | Description |
|--------|------|-------------|
| POST | `/convert` | Upload `.docx` → ZIP (LaTeX + BibTeX) |
| GET | `/health` | Health check |
