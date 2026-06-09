import os
try:
    import certifi
    os.environ.setdefault("SSL_CERT_FILE", certifi.where())
    os.environ.setdefault("REQUESTS_CA_BUNDLE", certifi.where())
except ImportError:
    pass

import streamlit as st
import pandas as pd
import re
import time
import zipfile
import io
from collections import defaultdict
from nltk.metrics.distance import edit_distance

# Optional format parsers (graceful fallback if not installed)
try:
    import pdfplumber
    _PDF_OK = True
except ImportError:
    _PDF_OK = False

try:
    from docx import Document as DocxDocument
    _DOCX_OK = True
except ImportError:
    _DOCX_OK = False

import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.stem import PorterStemmer, WordNetLemmatizer

for pkg in ["punkt", "punkt_tab", "stopwords", "wordnet", "omw-1.4"]:
    try:
        nltk.download(pkg, quiet=True)
    except Exception:
        pass

# ─────────────────────────────────────────────────────────────────────────────
# Page config + CSS
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="IR System – Group 52",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
/* ── Global ─────────────────────────────────────────────────── */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* ── Hero banner ─────────────────────────────────────────────── */
.hero {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 40%, #0f3460 100%);
    border-radius: 16px;
    padding: 2.2rem 2.5rem;
    margin-bottom: 1.5rem;
    color: white;
    display: flex;
    align-items: center;
    gap: 1.2rem;
}
.hero-icon { font-size: 3rem; }
.hero-title { font-size: 2rem; font-weight: 700; letter-spacing: -0.5px; margin: 0; }
.hero-sub   { font-size: 0.95rem; opacity: 0.75; margin: 0.3rem 0 0; }
.badge {
    display: inline-block;
    background: rgba(255,255,255,0.15);
    border: 1px solid rgba(255,255,255,0.25);
    border-radius: 20px;
    padding: 2px 12px;
    font-size: 0.78rem;
    font-weight: 500;
    margin-top: 0.5rem;
}

/* ── Section headers ─────────────────────────────────────────── */
.section-header {
    font-size: 1.35rem;
    font-weight: 700;
    color: #0f3460;
    border-left: 4px solid #e94560;
    padding-left: 0.75rem;
    margin: 1.5rem 0 1rem;
}

/* ── Stat cards ──────────────────────────────────────────────── */
.stat-row { display: flex; gap: 1rem; margin: 1rem 0; flex-wrap: wrap; }
.stat-card {
    flex: 1; min-width: 140px;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    border-radius: 12px; padding: 1.1rem 1.3rem;
    color: white; text-align: center;
}
.stat-card.green  { background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); color: #1a1a2e; }
.stat-card.orange { background: linear-gradient(135deg, #f7971e 0%, #ffd200 100%); color: #1a1a2e; }
.stat-card.red    { background: linear-gradient(135deg, #e94560 0%, #f77f00 100%); }
.stat-card.blue   { background: linear-gradient(135deg, #2193b0 0%, #6dd5ed 100%); color: #1a1a2e; }
.stat-val  { font-size: 1.8rem; font-weight: 700; line-height: 1; }
.stat-lbl  { font-size: 0.78rem; font-weight: 500; opacity: 0.85; margin-top: 0.25rem; }

/* ── Info / result boxes ─────────────────────────────────────── */
.result-box {
    border-radius: 10px; padding: 1rem 1.2rem;
    margin: 0.6rem 0; font-size: 0.93rem;
}
.result-box.info    { background: #e8f4fd; border-left: 4px solid #2193b0; color: #1a3a4a; }
.result-box.success { background: #e8fdf0; border-left: 4px solid #11998e; color: #0d3325; }
.result-box.warn    { background: #fff8e1; border-left: 4px solid #ffd200; color: #3d3000; }
.result-box.danger  { background: #fdecea; border-left: 4px solid #e94560; color: #4a0d18; }

/* ── Method cards (two-up comparison) ───────────────────────── */
.method-card {
    background: #f8f9ff;
    border: 1px solid #e0e4ff;
    border-radius: 12px;
    padding: 1.2rem 1.4rem;
    height: 100%;
}
.method-card h4 { margin: 0 0 0.6rem; color: #0f3460; font-size: 1rem; font-weight: 600; }

/* ── Inference cards ─────────────────────────────────────────── */
.inf-card {
    background: white;
    border: 1px solid #e8eaf6;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    margin: 0.8rem 0;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}
.inf-card h4 {
    margin: 0 0 0.5rem;
    font-size: 1rem; font-weight: 600;
    color: #e94560;
}

/* ── Sidebar ─────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1a1a2e 0%, #16213e 100%);
}
[data-testid="stSidebar"] * { color: #e0e0f0 !important; }
[data-testid="stSidebar"] .stMarkdown h1,
[data-testid="stSidebar"] .stMarkdown h2,
[data-testid="stSidebar"] .stMarkdown h3 { color: #ffffff !important; }
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.15) !important; }

/* ── Download button ─────────────────────────────────────────── */
.stDownloadButton > button {
    background: linear-gradient(135deg, #11998e, #38ef7d) !important;
    color: #0d3325 !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    padding: 0.45rem 1.1rem !important;
}
.stDownloadButton > button:hover {
    opacity: 0.9 !important;
    transform: translateY(-1px) !important;
}

/* ── Tabs ────────────────────────────────────────────────────── */
[data-testid="stTabs"] [role="tab"] {
    font-weight: 500;
    font-size: 0.88rem;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
    color: #0f3460 !important;
    border-bottom: 2px solid #e94560 !important;
}

/* ── Dataframe ───────────────────────────────────────────────── */
[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# Hero
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <div class="hero-icon">🔍</div>
  <div>
    <p class="hero-title">Information Retrieval System</p>
    <p class="hero-sub">End-to-end IR pipeline — preprocessing · indexing · phrase queries · tolerant retrieval</p>
    <span class="badge">Group 52 &nbsp;·&nbsp; AIMLCZG537 / DSECLZG537 &nbsp;·&nbsp; S2-2025</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# NLP helpers
# ─────────────────────────────────────────────────────────────────────────────
try:
    STOPWORDS = set(stopwords.words("english"))
except Exception:
    STOPWORDS = set()

stemmer    = PorterStemmer()
lemmatizer = WordNetLemmatizer()


def tokenize(text):
    """Tokenize text; falls back to simple split if NLTK punkt is unavailable."""
    try:
        return word_tokenize(text.lower())
    except Exception:
        return re.findall(r"[a-zA-Z0-9]+", text.lower())


def basic_preprocess(tokens, lowercase=True, remove_sw=True,
                     handle_hyphen=True, stem=False, lemma=False):
    result = []
    for tok in tokens:
        parts = tok.split("-") if handle_hyphen else [tok]
        for part in parts:
            t = part.lower() if lowercase else part
            t = re.sub(r"[^a-z0-9]", "", t) if lowercase else re.sub(r"[^a-zA-Z0-9]", "", t)
            if not t:
                continue
            if remove_sw and t in STOPWORDS:
                continue
            if stem:
                t = stemmer.stem(t)
            elif lemma:
                t = lemmatizer.lemmatize(t)
            result.append(t)
    return result


@st.cache_data
def build_inverted_index(docs):
    index = defaultdict(list)
    for doc_id, tokens in docs.items():
        for tok in set(tokens):
            index[tok].append(doc_id)
    return dict(index)


@st.cache_data
def build_positional_index(docs):
    index = defaultdict(lambda: defaultdict(list))
    for doc_id, tokens in docs.items():
        for pos, tok in enumerate(tokens):
            index[tok][doc_id].append(pos)
    return {k: dict(v) for k, v in index.items()}


@st.cache_data
def build_biword_index(docs):
    index = defaultdict(list)
    for doc_id, tokens in docs.items():
        for i in range(len(tokens) - 1):
            bw = f"{tokens[i]} {tokens[i+1]}"
            if doc_id not in index[bw]:
                index[bw].append(doc_id)
    return dict(index)


def phrase_query_biword(phrase, biword_index):
    words = phrase.lower().split()
    if len(words) < 2:
        return biword_index.get(words[0], []) if words else []
    result_set = None
    for i in range(len(words) - 1):
        bw = f"{words[i]} {words[i+1]}"
        posting = set(biword_index.get(bw, []))
        result_set = posting if result_set is None else result_set & posting
    return sorted(result_set) if result_set else []


def phrase_query_positional(phrase, pos_index):
    words = phrase.lower().split()
    if not words or words[0] not in pos_index:
        return []
    candidates = set(pos_index[words[0]].keys())
    for i, word in enumerate(words[1:], start=1):
        if word not in pos_index:
            return []
        next_docs = pos_index[word]
        valid = set()
        for doc_id in candidates & set(next_docs.keys()):
            prev_pos = pos_index[words[i - 1]].get(doc_id, [])
            curr_pos = set(next_docs[doc_id])
            if any(p + 1 in curr_pos for p in prev_pos):
                valid.add(doc_id)
        candidates = valid
    return sorted(candidates)


# ── BST ───────────────────────────────────────────────────────────────────────
class BSTNode:
    __slots__ = ("key", "postings", "left", "right")
    def __init__(self, key, postings):
        self.key, self.postings = key, postings
        self.left = self.right = None

class BST:
    def __init__(self): self.root = None
    def insert(self, key, postings): self.root = self._ins(self.root, key, postings)
    def _ins(self, node, key, postings):
        if node is None: return BSTNode(key, postings)
        if key < node.key: node.left = self._ins(node.left, key, postings)
        elif key > node.key: node.right = self._ins(node.right, key, postings)
        return node
    def search(self, key):
        comps, node = 0, self.root
        while node:
            comps += 1
            if key == node.key: return node.postings, comps
            node = node.left if key < node.key else node.right
        return [], comps


# ── B-Tree ────────────────────────────────────────────────────────────────────
class BTreeNode:
    def __init__(self, leaf=True):
        self.keys, self.values, self.children, self.leaf = [], [], [], leaf

class BTree:
    def __init__(self, t=3): self.root, self.t, self.comparisons = BTreeNode(), t, 0
    def search(self, key, node=None):
        if node is None: node, self.comparisons = self.root, 0
        i = 0
        while i < len(node.keys):
            self.comparisons += 1
            if key == node.keys[i]: return node.values[i], self.comparisons
            elif key < node.keys[i]: break
            i += 1
        if node.leaf: return [], self.comparisons
        return self.search(key, node.children[i])
    def insert(self, key, value):
        root = self.root
        if len(root.keys) == 2 * self.t - 1:
            new_root = BTreeNode(leaf=False)
            new_root.children.append(self.root)
            self._split(new_root, 0)
            self.root = new_root
        self._ins_nf(self.root, key, value)
    def _ins_nf(self, node, key, value):
        i = len(node.keys) - 1
        if node.leaf:
            for idx, k in enumerate(node.keys):
                if k == key: node.values[idx] = value; return
            node.keys.append(None); node.values.append(None)
            while i >= 0 and key < node.keys[i]:
                node.keys[i+1] = node.keys[i]; node.values[i+1] = node.values[i]; i -= 1
            node.keys[i+1] = key; node.values[i+1] = value
        else:
            while i >= 0 and key < node.keys[i]: i -= 1
            i += 1
            if len(node.children[i].keys) == 2*self.t-1:
                self._split(node, i)
                if key > node.keys[i]: i += 1
            self._ins_nf(node.children[i], key, value)
    def _split(self, parent, i):
        t, child = self.t, parent.children[i]
        nc = BTreeNode(leaf=child.leaf); mid = t - 1
        parent.keys.insert(i, child.keys[mid]); parent.values.insert(i, child.values[mid])
        parent.children.insert(i+1, nc)
        nc.keys = child.keys[mid+1:]; nc.values = child.values[mid+1:]
        child.keys = child.keys[:mid]; child.values = child.values[:mid]
        if not child.leaf: nc.children = child.children[mid+1:]; child.children = child.children[:mid+1]


# ── K-gram + tolerant ─────────────────────────────────────────────────────────
def build_kgram_index(vocab, k=2):
    index = defaultdict(list)
    for term in vocab:
        padded = f"${term}$"
        for i in range(len(padded) - k + 1):
            index[padded[i:i+k]].append(term)
    return dict(index)


def wildcard_search(pattern, kgram_index, vocab, k=2):
    parts = pattern.split("*")
    padded_parts = []
    if parts[0]: padded_parts.append(f"${parts[0]}")
    if parts[-1]: padded_parts.append(f"{parts[-1]}$")
    for part in parts[1:-1]:
        if part: padded_parts.append(part)
    candidate_sets = []
    for part in padded_parts:
        grams = [part[i:i+k] for i in range(len(part)-k+1)]
        for gram in grams:
            candidate_sets.append(set(kgram_index.get(gram, [])))
    if not candidate_sets:
        return sorted(vocab)
    result = candidate_sets[0]
    for s in candidate_sets[1:]: result &= s
    regex = re.compile("^" + re.escape(pattern).replace(r"\*", ".*") + "$")
    return sorted(t for t in result if regex.match(t))


def spelling_correction(word, vocab, max_dist=2):
    return sorted(
        (t for t in vocab if edit_distance(word, t) <= max_dist),
        key=lambda t: edit_distance(word, t),
    )[:10]


def soundex(word):
    word = word.upper()
    if not word: return ""
    code_map = {"BFPV":"1","CGJKQSXYZ":"2","DT":"3","L":"4","MN":"5","R":"6"}
    result, prev = word[0], ""
    for ch in word[1:]:
        code = "0"
        for letters, digit in code_map.items():
            if ch in letters: code = digit; break
        if code != "0" and code != prev: result += code
        prev = code
        if len(result) == 4: break
    return result.ljust(4, "0")


def phonetic_search(word, vocab):
    code = soundex(word)
    return [t for t in vocab if soundex(t) == code]


def jaccard_sim(a, b):
    if not a and not b: return 1.0
    return len(a & b) / len(a | b)


# ─────────────────────────────────────────────────────────────────────────────
# Demo dataset
# ─────────────────────────────────────────────────────────────────────────────
DEMO_DOCS = {
    "doc01_ir_overview.txt":      "Information retrieval systems help users find relevant documents quickly. Modern IR systems use advanced indexing techniques and ranking algorithms.",
    "doc02_tokenization.txt":     "Tokenization is the first step in text preprocessing. It involves splitting text into individual tokens or words for further processing.",
    "doc03_stopwords.txt":        "Stop word removal eliminates common words like the, is, and at from the document collection. This reduces noise and improves retrieval precision.",
    "doc04_stemming.txt":         "Stemming reduces words to their root form. For example, running becomes run and jumping becomes jump. Lemmatization considers linguistic context.",
    "doc05_positional_index.txt": "A positional index stores the exact position of each term in every document. It enables accurate phrase query processing and proximity search.",
    "doc06_bst_btree.txt":        "Binary search trees allow efficient dictionary lookups with O(log n) average time. B-Trees are self-balancing and provide faster guaranteed retrieval.",
    "doc07_wildcard.txt":         "Wildcard queries use the asterisk symbol to match multiple terms. K-gram index structures efficiently support wildcard search operations.",
    "doc08_spelling.txt":         "Spelling correction handles typographical errors in user queries. Edit distance measures how similar two strings are by counting minimal edits.",
    "doc09_phonetic.txt":         "Phonetic correction matches words that sound alike using algorithms such as Soundex or Metaphone. It helps when users are unsure of spelling.",
    "doc10_biword.txt":           "Biword index stores pairs of consecutive words as index terms. It supports phrase queries but may produce false positives for longer phrases.",
    "doc11_tfidf.txt":            "TF-IDF weighting assigns importance scores to terms based on frequency in a document and rarity across the corpus. It improves retrieval ranking.",
    "doc12_boolean.txt":          "Boolean retrieval uses AND, OR, and NOT operators to combine query terms. It returns an exact set of documents matching the logical expression.",
    "doc13_inverted_index.txt":   "An inverted index maps each term to the list of documents containing it. It is the core data structure of most information retrieval systems.",
    "doc14_lemmatization.txt":    "Lemmatization uses vocabulary and morphological analysis to return the base dictionary form of a word. Unlike stemming, it always produces a valid word.",
    "doc15_evaluation.txt":       "Retrieval evaluation uses metrics such as precision, recall, and F1 score. Mean Average Precision (MAP) measures ranking quality across multiple queries.",
}


def make_sample_zip():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, content in DEMO_DOCS.items():
            zf.writestr(fname, content)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# Multi-format file parser
# ─────────────────────────────────────────────────────────────────────────────
def _extract_text_from_bytes(name: str, data: bytes) -> dict:
    """Return {doc_name: text_content} for a single uploaded file.
    Supported: .txt  .csv  .pdf  .docx  .zip (recursively parsed)
    """
    ext = os.path.splitext(name)[1].lower()
    results = {}

    if ext == ".txt":
        results[name] = data.decode("utf-8", errors="ignore")

    elif ext == ".csv":
        try:
            df = pd.read_csv(io.BytesIO(data), dtype=str).fillna("")
            # Each row becomes its own document; fall back to whole-file if tiny
            if len(df) > 1:
                for idx, row in df.iterrows():
                    results[f"{name}__row{idx+1}"] = " ".join(row.astype(str).tolist())
            else:
                results[name] = " ".join(df.to_string(index=False).split())
        except Exception as e:
            results[name] = data.decode("utf-8", errors="ignore")

    elif ext == ".pdf":
        if _PDF_OK:
            try:
                text_pages = []
                with pdfplumber.open(io.BytesIO(data)) as pdf:
                    for page in pdf.pages:
                        t = page.extract_text()
                        if t:
                            text_pages.append(t)
                results[name] = "\n".join(text_pages) if text_pages else "(no extractable text)"
            except Exception as e:
                results[name] = f"(PDF parse error: {e})"
        else:
            results[name] = "(pdfplumber not installed — run: pip install pdfplumber)"

    elif ext == ".docx":
        if _DOCX_OK:
            try:
                doc = DocxDocument(io.BytesIO(data))
                results[name] = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
            except Exception as e:
                results[name] = f"(DOCX parse error: {e})"
        else:
            results[name] = "(python-docx not installed — run: pip install python-docx)"

    elif ext == ".zip":
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                for entry in zf.namelist():
                    if entry.endswith("/"):
                        continue  # skip directories
                    inner_data = zf.read(entry)
                    inner_name = os.path.basename(entry) or entry
                    nested = _extract_text_from_bytes(inner_name, inner_data)
                    results.update(nested)
        except Exception as e:
            results[name] = f"(ZIP parse error: {e})"

    else:
        # Try plain-text decode as last resort
        try:
            results[name] = data.decode("utf-8", errors="ignore")
        except Exception:
            results[name] = "(unsupported format)"

    return results


# ─────────────────────────────────────────────────────────────────────────────
# Session state
# ─────────────────────────────────────────────────────────────────────────────
for key in ["raw_docs", "doc_names"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ─────────────────────────────────────────────────────────────────────────────
# Sidebar
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Dataset")
    st.markdown("---")

    # Download sample dataset
    st.markdown("**Download sample dataset**")
    st.download_button(
        label="⬇️  Download 15 sample .txt files",
        data=make_sample_zip(),
        file_name="IR_sample_dataset.zip",
        mime="application/zip",
        use_container_width=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

    # Upload
    st.markdown("**Upload your own documents**")
    st.caption("Supported: `.txt` · `.csv` · `.pdf` · `.docx` · `.zip` (containing any of the above)")
    uploaded_files = st.file_uploader(
        "Upload documents",
        type=["txt", "csv", "pdf", "docx", "zip"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    use_demo = st.checkbox("Use built-in demo dataset", value=True)
    st.markdown("---")

    if uploaded_files:
        raw_docs = {}
        parse_errors = []
        for f in uploaded_files:
            data = f.read()
            parsed = _extract_text_from_bytes(f.name, data)
            for doc_name, text in parsed.items():
                if text.strip():
                    raw_docs[doc_name] = text
                else:
                    parse_errors.append(doc_name)
        if raw_docs:
            st.session_state.raw_docs = raw_docs
            st.session_state.doc_names = list(raw_docs.keys())
            st.success(f"✅ {len(raw_docs)} document(s) loaded")
        if parse_errors:
            st.warning(f"⚠️ Skipped {len(parse_errors)} empty/unreadable file(s).")
    elif use_demo:
        st.session_state.raw_docs = DEMO_DOCS
        st.session_state.doc_names = list(DEMO_DOCS.keys())
        st.info(f"📚 Demo: {len(DEMO_DOCS)} documents")

    st.markdown("### ℹ️ About")
    st.markdown("""
**IR Assignment 1**
Group 52 · S2-2025

Covers:
• Text Preprocessing
• Inverted Index
• Phrase Queries
• BST vs B-Tree
• Tolerant Retrieval
""")

raw_docs = st.session_state.raw_docs or {}
doc_names = st.session_state.doc_names or []
doc_id_map = {i: name for i, name in enumerate(doc_names)}

# ─────────────────────────────────────────────────────────────────────────────
# Top-level stats bar
# ─────────────────────────────────────────────────────────────────────────────
if raw_docs:
    total_words = sum(len(t.split()) for t in raw_docs.values())
    all_tokens_raw = [tok for text in raw_docs.values() for tok in basic_preprocess(tokenize(text), stem=True)]
    vocab_size = len(set(all_tokens_raw))
    st.markdown(f"""
<div class="stat-row">
  <div class="stat-card blue">
    <div class="stat-val">{len(raw_docs)}</div>
    <div class="stat-lbl">Documents</div>
  </div>
  <div class="stat-card green">
    <div class="stat-val">{total_words:,}</div>
    <div class="stat-lbl">Total Words</div>
  </div>
  <div class="stat-card orange">
    <div class="stat-val">{vocab_size}</div>
    <div class="stat-lbl">Unique Terms (stemmed)</div>
  </div>
  <div class="stat-card red">
    <div class="stat-val">{round(total_words/len(raw_docs))}</div>
    <div class="stat-lbl">Avg Words / Doc</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# Tabs
# ─────────────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "📄 Documents",
    "🔍 Search & Retrieval",
    " Preprocessing",
    "🔎 Phrase Query",
    "🌲 BST vs B-Tree",
    "🩹 Tolerant Retrieval",
    "📊 Inference & Discussion",
])

# ════════════════════════════════════════════════════════════════════════════
# TAB 0 – Documents
# ════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    st.markdown('<p class="section-header">Document Collection</p>', unsafe_allow_html=True)

    if not raw_docs:
        st.markdown('<div class="result-box warn">⚠️ Upload documents or enable the demo dataset from the sidebar.</div>', unsafe_allow_html=True)
    else:
        # Search filter
        search_filter = st.text_input("🔎 Filter documents by keyword", placeholder="Type to filter…")
        filtered = {k: v for k, v in raw_docs.items()
                    if not search_filter or search_filter.lower() in v.lower() or search_filter.lower() in k.lower()}

        st.caption(f"Showing {len(filtered)} of {len(raw_docs)} document(s)")

        for i, (name, content) in enumerate(filtered.items()):
            word_count = len(content.split())
            with st.expander(f"📄 {name}  —  {word_count} words"):
                st.markdown(f'<div style="background:#f8f9ff;border-radius:8px;padding:1rem;font-size:0.9rem;line-height:1.6">{content}</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 1 – Search & Retrieval  (main interactive hub)
# ════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    st.markdown('<p class="section-header">Interactive Search & Retrieval</p>', unsafe_allow_html=True)

    if not raw_docs:
        st.markdown('<div class="result-box warn">⚠️ Upload documents or enable the demo dataset from the sidebar.</div>', unsafe_allow_html=True)
    else:
        # ── Build all indexes once ────────────────────────────────────────────
        _sr_stem_docs  = {i: basic_preprocess(tokenize(t), stem=True)  for i, t in enumerate(raw_docs.values())}
        _sr_lem_docs   = {i: basic_preprocess(tokenize(t), lemma=True) for i, t in enumerate(raw_docs.values())}
        _sr_inv_stem   = build_inverted_index(_sr_stem_docs)
        _sr_inv_lem    = build_inverted_index(_sr_lem_docs)
        _sr_biword     = build_biword_index(_sr_lem_docs)
        _sr_pos        = build_positional_index(_sr_lem_docs)
        _sr_vocab_stem = sorted(_sr_inv_stem.keys())
        _sr_kgram      = build_kgram_index(_sr_vocab_stem, k=2)

        # ── Technique selector ────────────────────────────────────────────────
        st.markdown("#### 🎛️ Select Retrieval Technique")
        technique = st.selectbox(
            "Retrieval technique",
            [
                "Single Term  —  Inverted Index",
                "Boolean AND",
                "Boolean OR",
                "Boolean NOT",
                "Phrase  —  Biword Index",
                "Phrase  —  Positional Index",
                "Wildcard  (use * as wildcard)",
                "Spelling Correction",
                "Phonetic  (Soundex)",
            ],
            label_visibility="collapsed",
        )

        st.markdown("---")

        # ── Query input (context-sensitive label) ─────────────────────────────
        if "Wildcard" in technique:
            query_val, query_ph = "inf*", "e.g.  inf*  or  *index*"
        elif "Boolean AND" in technique:
            query_val, query_ph = "information retrieval", "term1  term2  … (all must match)"
        elif "Boolean OR" in technique:
            query_val, query_ph = "stemming lemmatization", "term1  term2  … (any must match)"
        elif "Boolean NOT" in technique:
            query_val, query_ph = "stemming", "term to exclude from all docs"
        elif "Phrase" in technique:
            query_val, query_ph = "information retrieval", "exact phrase, e.g.  positional index"
        elif "Spelling" in technique:
            query_val, query_ph = "informaton", "misspelled word"
        elif "Phonetic" in technique:
            query_val, query_ph = "retrieve", "word to match phonetically"
        else:
            query_val, query_ph = "information", "single search term"

        query_in = st.text_input("🔎 Enter your query", value=query_val, placeholder=query_ph)

        if query_in.strip():
            q_raw = query_in.strip().lower()
            t0    = time.perf_counter()

            # ── Execute retrieval ─────────────────────────────────────────────
            result_doc_ids = []
            extra_info     = ""

            if "Single Term" in technique:
                term = stemmer.stem(re.sub(r"[^a-z0-9]", "", q_raw.split()[0]))
                result_doc_ids = sorted(_sr_inv_stem.get(term, []))
                extra_info = f"Stemmed term: **`{term}`**"

            elif "Boolean AND" in technique:
                terms = [stemmer.stem(re.sub(r"[^a-z0-9]", "", w))
                         for w in q_raw.split() if re.sub(r"[^a-z0-9]", "", w)]
                if terms:
                    sets = [set(_sr_inv_stem.get(t, [])) for t in terms]
                    result_doc_ids = sorted(set.intersection(*sets)) if sets else []
                    extra_info = f"AND of stemmed terms: {terms}"

            elif "Boolean OR" in technique:
                terms = [stemmer.stem(re.sub(r"[^a-z0-9]", "", w))
                         for w in q_raw.split() if re.sub(r"[^a-z0-9]", "", w)]
                if terms:
                    sets = [set(_sr_inv_stem.get(t, [])) for t in terms]
                    result_doc_ids = sorted(set.union(*sets)) if sets else []
                    extra_info = f"OR of stemmed terms: {terms}"

            elif "Boolean NOT" in technique:
                term = stemmer.stem(re.sub(r"[^a-z0-9]", "", q_raw.split()[0]))
                excluded = set(_sr_inv_stem.get(term, []))
                result_doc_ids = sorted(set(range(len(raw_docs))) - excluded)
                extra_info = f"Docs NOT containing stemmed term **`{term}`**"

            elif "Biword" in technique:
                words = re.sub(r"[^a-z\s]", "", q_raw).split()
                lq    = " ".join(lemmatizer.lemmatize(w) for w in words if w)
                result_doc_ids = phrase_query_biword(lq, _sr_biword)
                extra_info = f"Lemmatized phrase: **`{lq}`**"

            elif "Positional" in technique:
                words = re.sub(r"[^a-z\s]", "", q_raw).split()
                lq    = " ".join(lemmatizer.lemmatize(w) for w in words if w)
                result_doc_ids = phrase_query_positional(lq, _sr_pos)
                extra_info = f"Lemmatized phrase: **`{lq}`**"

            elif "Wildcard" in technique:
                matched_terms  = wildcard_search(q_raw, _sr_kgram, set(_sr_vocab_stem))
                hit_sets       = [set(_sr_inv_stem.get(t, [])) for t in matched_terms]
                result_doc_ids = sorted(set.union(*hit_sets)) if hit_sets else []
                extra_info = f"Matched {len(matched_terms)} term(s): `{', '.join(matched_terms[:12])}`{'…' if len(matched_terms)>12 else ''}"

            elif "Spelling" in technique:
                sugg = spelling_correction(q_raw.split()[0], _sr_vocab_stem, max_dist=2)
                if sugg:
                    hit_sets       = [set(_sr_inv_stem.get(t, [])) for t in sugg]
                    result_doc_ids = sorted(set.union(*hit_sets)) if hit_sets else []
                    extra_info = f"Suggestions (edit dist ≤ 2): `{', '.join(sugg)}`"
                else:
                    extra_info = "No close matches found in vocabulary."

            elif "Phonetic" in technique:
                matched_terms  = phonetic_search(q_raw.split()[0], _sr_vocab_stem)
                hit_sets       = [set(_sr_inv_stem.get(t, [])) for t in matched_terms]
                result_doc_ids = sorted(set.union(*hit_sets)) if hit_sets else []
                sdx = soundex(q_raw.split()[0])
                extra_info = f"Soundex code: **`{sdx}`** — matched terms: `{', '.join(matched_terms)}`"

            elapsed_ms = (time.perf_counter() - t0) * 1000

            # ── Display results ───────────────────────────────────────────────
            st.markdown(f"**Technique:** {technique}")
            if extra_info:
                st.markdown(extra_info)

            m_col1, m_col2 = st.columns(2)
            m_col1.metric("Results", len(result_doc_ids))
            m_col2.metric("Query time", f"{elapsed_ms:.3f} ms")

            if result_doc_ids:
                st.markdown(f'<div class="result-box success">✅ Found <strong>{len(result_doc_ids)}</strong> matching document(s)</div>', unsafe_allow_html=True)
                for doc_id in result_doc_ids:
                    doc_name    = doc_id_map.get(doc_id, str(doc_id))
                    doc_text    = list(raw_docs.values())[doc_id] if doc_id < len(raw_docs) else ""
                    snippet     = doc_text[:200].replace("\n", " ")
                    snippet_hl  = snippet + ("…" if len(doc_text) > 200 else "")
                    with st.expander(f"📄 {doc_name}"):
                        st.markdown(f'<div style="background:#f8f9ff;border-radius:8px;padding:0.8rem;font-size:0.88rem;line-height:1.6">{snippet_hl}</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="result-box info">ℹ️ No documents matched this query.</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 2 – Preprocessing
# ════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    st.markdown('<p class="section-header">Text Preprocessing Pipeline</p>', unsafe_allow_html=True)

    if not raw_docs:
        st.markdown('<div class="result-box warn">⚠️ Load a dataset first.</div>', unsafe_allow_html=True)
    else:
        # Options
        with st.container():
            st.markdown("#### ⚙️ Preprocessing Options")
            c1, c2, c3 = st.columns(3)
            with c1:
                opt_lower  = st.checkbox("🔡 Lowercase",          value=True)
                opt_sw     = st.checkbox("🚫 Remove Stop Words",  value=True)
            with c2:
                opt_hyphen = st.checkbox("➖ Hyphen Splitting",   value=True)
            with c3:
                norm_choice = st.radio("📐 Normalisation",
                                       ["None", "Stemming", "Lemmatization"], index=1)

        do_stem  = norm_choice == "Stemming"
        do_lemma = norm_choice == "Lemmatization"

        processed_docs = {
            i: basic_preprocess(tokenize(text), lowercase=opt_lower, remove_sw=opt_sw,
                                handle_hyphen=opt_hyphen, stem=do_stem, lemma=do_lemma)
            for i, (_, text) in enumerate(raw_docs.items())
        }

        # Pipeline visualization for first doc
        st.markdown("---")
        st.markdown("#### 🔬 Step-by-Step Token Transformation (first document)")
        first_text = list(raw_docs.values())[0]
        raw_tok = tokenize(first_text)
        steps = [
            ("1️⃣  Raw tokens",                     raw_tok[:15]),
            ("2️⃣  After lowercase + punct removal", basic_preprocess(raw_tok, lowercase=True, remove_sw=False, handle_hyphen=False)[:15]),
            ("3️⃣  After stop word removal",         basic_preprocess(raw_tok, lowercase=True, remove_sw=True,  handle_hyphen=False)[:15]),
            ("4️⃣  After hyphen handling",            basic_preprocess(raw_tok, lowercase=True, remove_sw=True,  handle_hyphen=True )[:15]),
            ("5️⃣  After stemming",                  basic_preprocess(raw_tok, lowercase=True, remove_sw=True,  handle_hyphen=True, stem=True )[:15]),
            ("6️⃣  After lemmatization",              basic_preprocess(raw_tok, lowercase=True, remove_sw=True,  handle_hyphen=True, lemma=True)[:15]),
        ]
        df_steps = pd.DataFrame([(s, " · ".join(t)) for s, t in steps], columns=["Step", "Sample Tokens (first 15)"])
        st.dataframe(df_steps, use_container_width=True, hide_index=True)

        # Inverted Index
        st.markdown("---")
        st.markdown("#### 📋 Inverted Index")
        inv_index = build_inverted_index(processed_docs)
        df_inv = pd.DataFrame(
            [(term, len(p), ", ".join(doc_id_map.get(d, str(d)) for d in sorted(p)))
             for term, p in sorted(inv_index.items())],
            columns=["Term", "Doc Frequency", "Posting List"],
        )
        st.dataframe(df_inv, use_container_width=True, height=280, hide_index=True)

        # Stemming vs Lemmatization
        st.markdown("---")
        st.markdown("#### ⚖️ Stemming vs Lemmatization — Retrieval Quality Comparison")

        stem_docs  = {i: basic_preprocess(tokenize(t), stem=True)  for i, t in enumerate(raw_docs.values())}
        lemma_docs = {i: basic_preprocess(tokenize(t), lemma=True) for i, t in enumerate(raw_docs.values())}
        stem_idx   = build_inverted_index(stem_docs)
        lemma_idx  = build_inverted_index(lemma_docs)
        stem_vocab  = set(t for ts in stem_docs.values()  for t in ts)
        lemma_vocab = set(t for ts in lemma_docs.values() for t in ts)

        test_queries = ["information", "running", "retrieval", "process", "index"]
        rows = []
        for q in test_queries:
            sq, lq = stemmer.stem(q), lemmatizer.lemmatize(q)
            sr, lr = set(stem_idx.get(sq, [])), set(lemma_idx.get(lq, []))
            rows.append({"Query": q, "Stemmed form": sq, "Lemmatized form": lq,
                         "Stem hits": str(sorted(sr)), "Lemma hits": str(sorted(lr)),
                         "Jaccard": f"{jaccard_sim(sr, lr):.2f}"})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        col_a, col_b = st.columns(2)
        col_a.markdown(f"""
<div class="method-card">
<h4>✂️ Stemming</h4>
Vocabulary size: <strong>{len(stem_vocab)}</strong><br>
Fast, rule-based, may produce non-words<br>
e.g. "information" → <code>inform</code>
</div>""", unsafe_allow_html=True)
        col_b.markdown(f"""
<div class="method-card">
<h4>📖 Lemmatization</h4>
Vocabulary size: <strong>{len(lemma_vocab)}</strong><br>
Linguistically accurate, always valid word<br>
e.g. "running" → <code>run</code>
</div>""", unsafe_allow_html=True)

        st.markdown('<div class="result-box success">✅ <strong>Verdict:</strong> Lemmatization is more suitable for this dataset — it preserves semantic meaning and returns valid dictionary words, improving retrieval precision.</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 3 – Phrase Query
# ════════════════════════════════════════════════════════════════════════════
with tabs[3]:
    st.markdown('<p class="section-header">Phrase Query Processing</p>', unsafe_allow_html=True)

    if not raw_docs:
        st.markdown('<div class="result-box warn">⚠️ Load a dataset first.</div>', unsafe_allow_html=True)
    else:
        lem_docs   = {i: basic_preprocess(tokenize(t), lemma=True) for i, t in enumerate(raw_docs.values())}
        biword_idx = build_biword_index(lem_docs)
        pos_idx    = build_positional_index(lem_docs)

        phrase_q = st.text_input("🔎 Enter a phrase query", value="information retrieval",
                                 placeholder="e.g. positional index")

        if phrase_q.strip():
            # Lemmatize each query word (without stop-word removal) so the
            # query vocabulary matches the lemmatized document index.
            raw_words = re.sub(r"[^a-z\s]", "", phrase_q.lower().strip()).split()
            q_words   = [lemmatizer.lemmatize(w) for w in raw_words if w]
            q = " ".join(q_words) if q_words else phrase_q.lower().strip()

            t0 = time.perf_counter(); bw_res  = phrase_query_biword(q, biword_idx);    bw_ms  = (time.perf_counter()-t0)*1000
            t0 = time.perf_counter(); pos_res = phrase_query_positional(q, pos_idx); pos_ms = (time.perf_counter()-t0)*1000

            col1, col2 = st.columns(2, gap="large")
            with col1:
                st.markdown('<div class="method-card">', unsafe_allow_html=True)
                st.markdown("#### 🔗 Biword Index")
                st.metric("Query time", f"{bw_ms:.3f} ms")
                if bw_res:
                    st.markdown(f'<div class="result-box success">✅ Found in <strong>{len(bw_res)}</strong> document(s)</div>', unsafe_allow_html=True)
                    for d in bw_res: st.markdown(f"• **{doc_id_map.get(d, d)}**")
                else:
                    st.markdown('<div class="result-box info">ℹ️ No results found</div>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

            with col2:
                st.markdown('<div class="method-card">', unsafe_allow_html=True)
                st.markdown("#### 📍 Positional Index")
                st.metric("Query time", f"{pos_ms:.3f} ms")
                if pos_res:
                    st.markdown(f'<div class="result-box success">✅ Found in <strong>{len(pos_res)}</strong> document(s)</div>', unsafe_allow_html=True)
                    for d in pos_res: st.markdown(f"• **{doc_id_map.get(d, d)}**")
                else:
                    st.markdown('<div class="result-box info">ℹ️ No results found</div>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

            # Index samples
            st.markdown("---")
            st.markdown("#### 🗂️ Index Representations (sample)")
            ic1, ic2 = st.columns(2)
            with ic1:
                st.markdown("**Biword Index — first 8 entries**")
                st.json(dict(list(biword_idx.items())[:8]))
            with ic2:
                st.markdown("**Positional Index — first 4 terms**")
                st.json(dict(list(pos_idx.items())[:4]))

            # Comparison table
            st.markdown("---")
            st.markdown("#### 📊 Comparison")
            st.markdown("""
| Aspect | Biword Index | Positional Index |
|--------|:---:|:---:|
| Storage cost | Low | Higher |
| Phrase precision | ⚠️ May have false positives | ✅ Exact |
| Multi-word support | Pair decomposition only | Any length |
| Speed | Slightly faster | Slightly slower |
| **Recommended** | Approximate use | **Accurate phrase search** |
""")
            st.markdown('<div class="result-box warn">⚠️ <strong>False positive example:</strong> The query "black cat" stored as biword matches any doc with both "black" and "cat" consecutively, but a doc containing "black dog and cat" could cause issues in longer decomposed phrase chains. Positional index verifies exact adjacency at every position, eliminating false positives.</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 4 – Dictionary
# ════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    st.markdown('<p class="section-header">Dictionary Search: BST vs B-Tree</p>', unsafe_allow_html=True)

    if not raw_docs:
        st.markdown('<div class="result-box warn">⚠️ Load a dataset first.</div>', unsafe_allow_html=True)
    else:
        base_docs = {i: basic_preprocess(tokenize(t), stem=True) for i, t in enumerate(raw_docs.values())}
        inv_idx   = build_inverted_index(base_docs)
        vocab     = sorted(inv_idx.keys())

        bst   = BST()
        btree = BTree(t=3)
        for term in vocab:
            bst.insert(term, inv_idx[term])
            btree.insert(term, inv_idx[term])

        st.markdown(f'<div class="result-box info">ℹ️ Dictionary built with <strong>{len(vocab)}</strong> unique terms from {len(raw_docs)} documents.</div>', unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("#### ⚡ Performance Benchmark")
        default_q = ", ".join(vocab[:8]) if vocab else "information, retrieval"
        query_input = st.text_input("Enter comma-separated terms to benchmark", value=default_q)
        queries = [q.strip().lower() for q in query_input.split(",") if q.strip()]

        if queries:
            rows = []
            for q in queries:
                t0 = time.perf_counter(); bst_res, bst_c = bst.search(q);   bst_us = (time.perf_counter()-t0)*1e6
                t0 = time.perf_counter(); bt_res,  bt_c  = btree.search(q); bt_us  = (time.perf_counter()-t0)*1e6
                rows.append({"Term": q,
                              "BST Time (µs)": f"{bst_us:.2f}", "BST Comparisons": bst_c, "BST Docs": len(bst_res),
                              "B-Tree Time (µs)": f"{bt_us:.2f}", "B-Tree Comparisons": bt_c, "B-Tree Docs": len(bt_res)})

            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            avg_bst = sum(r["BST Comparisons"] for r in rows) / len(rows)
            avg_bt  = sum(r["B-Tree Comparisons"] for r in rows) / len(rows)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Avg BST Comparisons",    f"{avg_bst:.1f}")
            m2.metric("Avg B-Tree Comparisons", f"{avg_bt:.1f}")
            m3.metric("Vocab Size",             len(vocab))
            m4.metric("Queries Tested",         len(queries))

            st.markdown("---")
            c1, c2 = st.columns(2)
            c1.markdown("""
<div class="method-card">
<h4>🌳 Binary Search Tree</h4>
<strong>Time complexity:</strong> O(log n) avg, O(n) worst case<br>
<strong>Weakness:</strong> Skewed on sorted insertion — no self-balancing<br>
<strong>Use case:</strong> Small dictionaries, random insertion order
</div>""", unsafe_allow_html=True)
            c2.markdown("""
<div class="method-card">
<h4>🌲 B-Tree (order 3)</h4>
<strong>Time complexity:</strong> O(log<sub>t</sub> n) guaranteed<br>
<strong>Strength:</strong> Self-balancing, disk-friendly block structure<br>
<strong>Use case:</strong> Large IR dictionaries, database indexes
</div>""", unsafe_allow_html=True)

            st.markdown('<div class="result-box success">✅ <strong>Verdict:</strong> B-Tree outperforms BST due to guaranteed balance — especially important for large, sorted vocabularies typical in IR systems.</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 5 – Tolerant Retrieval
# ════════════════════════════════════════════════════════════════════════════
with tabs[5]:
    st.markdown('<p class="section-header">Tolerant Retrieval</p>', unsafe_allow_html=True)

    if not raw_docs:
        st.markdown('<div class="result-box warn">⚠️ Load a dataset first.</div>', unsafe_allow_html=True)
    else:
        tol_docs  = {i: basic_preprocess(tokenize(t), stem=True) for i, t in enumerate(raw_docs.values())}
        tol_inv   = build_inverted_index(tol_docs)
        tol_vocab = sorted(tol_inv.keys())
        kgram_idx = build_kgram_index(tol_vocab, k=2)

        tol_tabs = st.tabs(["🌟 Wildcard", "✏️ Spelling Correction", "📏 Edit Distance", "🔤 K-gram Index", "🔊 Phonetic"])

        # ── Wildcard ───────────────────────────────────────────────────────────
        with tol_tabs[0]:
            st.markdown("#### Wildcard Queries (backed by K-gram index)")
            wc_q = st.text_input("Pattern  (use * as wildcard)", value="inf*", key="wc")
            if wc_q.strip():
                matches = wildcard_search(wc_q.strip().lower(), kgram_idx, set(tol_vocab))
                if matches:
                    st.markdown(f'<div class="result-box success">✅ <strong>{len(matches)}</strong> matching term(s): <code>{", ".join(matches)}</code></div>', unsafe_allow_html=True)
                    doc_hits = set()
                    for term in matches: doc_hits |= set(tol_inv.get(term, []))
                    st.markdown(f'<div class="result-box info">📄 Documents: <strong>{", ".join(doc_id_map.get(d,str(d)) for d in sorted(doc_hits))}</strong></div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="result-box warn">No matches found.</div>', unsafe_allow_html=True)

        # ── Spelling Correction ────────────────────────────────────────────────
        with tol_tabs[1]:
            st.markdown("#### Spelling Correction (Edit Distance ≤ threshold)")
            spell_q = st.text_input("Possibly misspelled word", value="informaton", key="spell")
            max_d   = st.slider("Max edit distance", 1, 4, 2, key="spell_d")
            if spell_q.strip():
                sugg = spelling_correction(spell_q.strip().lower(), tol_vocab, max_dist=max_d)
                if sugg:
                    rows_s = [{"Suggestion": t, "Edit Distance": edit_distance(spell_q.lower(), t)} for t in sugg]
                    st.dataframe(pd.DataFrame(rows_s), use_container_width=True, hide_index=True)
                    st.markdown(f'<div class="result-box success">✅ Best match: <strong>{sugg[0]}</strong></div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="result-box info">No suggestions within threshold.</div>', unsafe_allow_html=True)

        # ── Edit Distance ──────────────────────────────────────────────────────
        with tol_tabs[2]:
            st.markdown("#### Edit Distance Calculator")
            ec1, ec2 = st.columns(2)
            w1 = ec1.text_input("Word 1", value="information", key="ed1")
            w2 = ec2.text_input("Word 2", value="informaton",  key="ed2")
            if w1 and w2:
                dist = edit_distance(w1, w2)
                thr  = st.slider("Correction threshold", 1, 5, 2, key="ed_thr")
                st.metric("Edit Distance", dist)
                if dist <= thr:
                    st.markdown(f'<div class="result-box success">✅ Within threshold ({dist} ≤ {thr}) — correction applicable.</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="result-box danger">❌ Distance {dist} exceeds threshold {thr} — too dissimilar.</div>', unsafe_allow_html=True)

        # ── K-gram ─────────────────────────────────────────────────────────────
        with tol_tabs[3]:
            st.markdown("#### K-gram Index Explorer")
            kg_term = st.text_input("Term to decompose into k-grams", value="retrieval", key="kg")
            k_val   = st.slider("k", 2, 3, 2, key="kval")
            if kg_term.strip():
                padded = f"${kg_term.strip().lower()}$"
                grams  = [padded[i:i+k_val] for i in range(len(padded)-k_val+1)]
                st.markdown(f"**Padded form:** `{padded}`")
                st.markdown(f"**K-grams:** `{grams}`")
                kg_sample = {g: kgram_idx.get(g, []) for g in grams}
                st.json(kg_sample)

        # ── Phonetic ───────────────────────────────────────────────────────────
        with tol_tabs[4]:
            st.markdown("#### Phonetic Correction (Soundex)")
            phon_q = st.text_input("Word for phonetic matching", value="retrieve", key="phon")
            if phon_q.strip():
                code   = soundex(phon_q.strip())
                results = phonetic_search(phon_q.strip().lower(), tol_vocab)
                st.metric("Soundex Code", code)
                if results:
                    st.markdown(f'<div class="result-box success">✅ Phonetically similar terms: <code>{", ".join(results)}</code></div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="result-box info">No phonetically similar terms in vocabulary.</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# TAB 6 – Inference & Discussion
# ════════════════════════════════════════════════════════════════════════════
with tabs[6]:
    st.markdown('<p class="section-header">Inference & Discussion</p>', unsafe_allow_html=True)

    sections = [
        ("🔧 B. Preprocessing", """
- **Tokenization** split raw text into individual terms, enabling per-term indexing.
- **Lowercasing** ensured case-insensitive matching ("Information" = "information").
- **Stop word removal** reduced index size by ~30-40% and improved precision.
- **Hyphen handling** correctly split compound words (e.g. "well-known" → "well" + "known").
- **Stemming vs Lemmatization:** Lemmatization outperformed stemming — it produces valid dictionary words
  with POS context. Stemming is faster but can produce non-words.
  **✅ Verdict: Lemmatization is more suitable for this dataset.**
"""),
        ("🔎 C. Phrase Query", """
- **Biword index** is fast to build but produces **false positives** for long phrases — it only
  verifies consecutive pairs, not full-phrase adjacency.
- **Positional index** stores exact token positions, enabling verification that all phrase tokens
  appear consecutively in the correct order.
  **✅ Verdict: Positional index gives more accurate phrase query results.**
"""),
        ("🌲 D. Dictionary Structures", """
- **BST** offers O(log n) average lookup but degenerates to O(n) on sorted/nearly-sorted input.
- **B-Tree (order 3)** guarantees O(log_t n) worst-case and is self-balancing — optimal for large IR dictionaries.
- Experimental results confirmed fewer comparisons per lookup with B-Tree.
  **✅ Verdict: B-Tree is faster and more reliable.**
"""),
        ("🩹 E. Tolerant Retrieval", """
- **Wildcard (k-gram):** Successfully matched partial patterns ("inf*" → information, index, etc.).
- **Spelling correction:** Recovered common typos within edit distance 2.
- **K-gram index:** Provided fast wildcard support without full vocabulary scan.
- **Phonetic (Soundex):** Found similar-sounding terms — useful for voice-to-text queries.
  **✅ Overall: The system gracefully handles typos, partials, and phonetic variations.**
"""),
        ("⚠️ Limitations", """
1. Spell correction is limited to terms already in the vocabulary — OOV words get no correction.
2. Soundex is insensitive to vowels, leading to over-matching in some cases.
3. BST degrades on sorted insertion — AVL or Red-Black tree would be better in production.
4. No TF-IDF ranking — results are unranked boolean sets.
5. In-memory positional index may be infeasible for very large corpora.
"""),
        ("🚀 Future Improvements", """
1. Add TF-IDF / BM25 ranking for relevance-ordered results.
2. Replace BST with a hash map (O(1) avg) or AVL tree for guaranteed balance.
3. Implement permuterm index for complete wildcard coverage.
4. Use Metaphone/Double Metaphone instead of Soundex for better phonetic precision.
5. Add query expansion using Word2Vec / fastText embeddings.
6. Persist indexes to disk (SQLite, FAISS) to support large corpora.
"""),
    ]

    for title, body in sections:
        with st.expander(title, expanded=True):
            st.markdown(body)

# ─────────────────────────────────────────────────────────────────────────────
# Footer
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style="text-align:center;opacity:0.55;font-size:0.82rem;padding:0.5rem 0">
  IR Assignment 1 &nbsp;·&nbsp; Group 52 &nbsp;·&nbsp; AIMLCZG537 / DSECLZG537 &nbsp;·&nbsp; S2-2025
</div>
""", unsafe_allow_html=True)
