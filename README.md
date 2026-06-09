# IR Assignment 1 – Group 52

End-to-end Information Retrieval system built with Streamlit.

## Features

| Section | What's implemented |
|---------|-------------------|
| A | Streamlit workflow: upload, view, query, select options |
| B | Tokenization, lowercasing, stop-word removal, hyphen handling, stemming, lemmatization + comparison |
| C | Biword index & positional index phrase queries with side-by-side comparison |
| D | BST and B-Tree dictionary search with performance benchmarking |
| E | Wildcard (k-gram), spelling correction, edit distance, k-gram demo, phonetic (Soundex) |
| G | Full inference & discussion section |

## Installation

```bash
pip install -r requirements.txt
```

> NLTK corpora (`punkt`, `stopwords`, `wordnet`) are downloaded automatically on first run.

## Running the app

```bash
streamlit run app.py
```

The app opens in your browser at `http://localhost:8501`.

## Usage

1. **Dataset** – Upload your own `.txt` files via the sidebar, or use the built-in 15-document demo dataset.
2. Navigate through the six tabs to explore each IR component.
3. All options (preprocessing flags, query input, tree parameters) are interactive.

## File structure

```
Assignment/
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
└── README.md           # This file
```
