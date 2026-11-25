"""
IraPrep — The Sophisticated IRaMuTeQ Studio
Minimalist, elegant, step-by-step corpus compiler.
Zero cloud leakage. 100% Private. Fully self-contained.
"""

import io
import re
import json
import html
import unicodedata
from typing import List, Dict, Any, Optional, Tuple
import streamlit as st
import pandas as pd

try:
    import docx
except ImportError:
    docx = None

try:
    import pypdf
except ImportError:
    pypdf = None


# =============================================================================
# Core Engine (Embedded for zero-import-error reliability)
# =============================================================================

def remove_accents(text: str) -> str:
    """Removes accents from a string (e.g. 'santé' -> 'sante')."""
    if not text:
        return ""
    nfkd_form = unicodedata.normalize('NFKD', str(text))
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])


def sanitize_variable_name(name: str) -> str:
    """Sanitizes variable names for IRaMuTeQ: lowercase alphanumeric and underscores."""
    if not name:
        return "var"
    name = remove_accents(str(name)).lower().strip()
    name = name.lstrip("*")
    name = re.sub(r"[^a-z0-9_]", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")
    return name if name else "var"


def sanitize_variable_value(val: Any) -> str:
    """Sanitizes variable modality values for IRaMuTeQ."""
    if val is None or pd.isna(val):
        return "unknown"
    val = remove_accents(str(val)).lower().strip()
    val = re.sub(r"[^a-z0-9_]", "_", val)
    val = re.sub(r"_+", "_", val).strip("_")
    return val if val else "unknown"


def clean_text_for_iramuteq(
    text: str,
    compound_words: Optional[List[str]] = None,
    strip_urls: bool = True,
    strip_emojis: bool = True,
    normalize_quotes: bool = True,
    fix_punctuation_spacing: bool = True,
    remove_in_text_asterisks: bool = True,
    strip_parenthetical_counts: bool = True,
    normalize_symbols: bool = True
) -> str:
    """Cleans and normalizes text for IRaMuTeQ compliance."""
    if not text:
        return ""

    text = html.unescape(str(text))
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"```[\s\S]*?```", " ", text)
    text = re.sub(r"`[^`\n]+`", " ", text)

    if strip_urls:
        text = re.sub(r"!\[.*?\]\(.*?\)", " ", text)
        text = re.sub(r"\[([^\]]+)\]\(https?://[^\)]+\)", r"\1", text)
        text = re.sub(r"https?://\S+", " ", text)
        text = re.sub(r"www\.\S+", " ", text)

    text = re.sub(r"^\s*#{1,6}\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"^\s*[-*_]{3,}\s*$", " ", text, flags=re.MULTILINE)
    text = re.sub(r"^\s*>\s*", " ", text, flags=re.MULTILINE)
    text = re.sub(r"\[deleted\]|\[removed\]", " ", text, flags=re.IGNORECASE)

    if strip_parenthetical_counts:
        text = re.sub(r"\s*\(\d+\)", "", text)

    text = text.replace("[", "(").replace("]", ")").replace("{", "(").replace("}", ")")

    if normalize_symbols:
        text = re.sub(r"\b(\d+)\s*°\s*([cCfF])\b", r"\1_degrees_\2", text)
        text = text.replace("°", " degrees ")
        text = re.sub(r"&", " and ", text)
        text = re.sub(r"\$", " usd ", text)
        text = re.sub(r"€", " eur ", text)
        text = re.sub(r"£", " gbp ", text)
        text = re.sub(r"%", " percent ", text)
        text = re.sub(r"\+", " plus ", text)

    if normalize_quotes:
        text = re.sub(r"[“”«»]", '"', text)
        text = re.sub(r"[‘’`´]", "'", text)

    text = re.sub(r"[—–]", " - ", text)
    text = re.sub(r"…", "... ", text)

    if strip_emojis:
        text = re.sub(r"[\U00010000-\U0010ffff]", " ", text)
        text = re.sub(r"[\u2600-\u27bf]", " ", text)

    text = re.sub(r"[\u200b\u200c\u200d\ufeff\u00ad]", "", text)

    if compound_words:
        sorted_compounds = sorted([c.strip() for c in compound_words if c.strip()], key=len, reverse=True)
        for comp in sorted_compounds:
            parts = [re.escape(p) for p in re.split(r"[\s\-_]+", comp) if p]
            if len(parts) >= 2:
                pattern = r"\b" + r"[\s\-_]+".join(parts) + r"\b"
                replacement = "_".join(parts).lower()
                text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)

    if remove_in_text_asterisks:
        text = text.replace("*", "")

    text = re.sub(r"(?<![a-zA-Z0-9])_|_(?![a-zA-Z0-9])", " ", text)

    if fix_punctuation_spacing:
        text = re.sub(r"([.,?!;:~])(?=[^\s\d])", r"\1 ", text)

    text = re.sub(r",\s*,+", ", ", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n\n", text)

    return text.strip(" ,.\n\t")


class UCIUnit:
    """Represents a single Initial Context Unit (UCI) in IRaMuTeQ."""
    def __init__(self, variables: Dict[str, str], body_text: str, source_doc: str = ""):
        self.variables = variables
        self.body_text = body_text
        self.source_doc = source_doc

    def to_iramuteq_header(self) -> str:
        parts = ["****"]
        for var_name, var_val in self.variables.items():
            s_name = sanitize_variable_name(var_name)
            s_val = sanitize_variable_value(var_val)
            parts.append(f"*{s_name}_{s_val}")
        return " ".join(parts)

    def to_iramuteq_block(self) -> str:
        return f"{self.to_iramuteq_header()}\n{self.body_text}\n"

    def word_count(self) -> int:
        return len(self.body_text.split())


class CorpusLintReport:
    """Contains diagnostics and warnings about the compiled corpus."""
    def __init__(self):
        self.total_ucis: int = 0
        self.total_words: int = 0
        self.avg_words_per_uci: float = 0.0
        self.short_ucis: List[Tuple[int, int]] = []
        self.empty_ucis: List[int] = []
        self.variable_keys: List[str] = []


def compile_corpus(
    ucis: List[UCIUnit],
    min_words_warning: int = 25
) -> Tuple[str, CorpusLintReport]:
    """Compiles a list of UCIUnit objects into a single IRaMuTeQ text corpus."""
    report = CorpusLintReport()
    report.total_ucis = len(ucis)

    all_keys = set()
    total_words = 0
    corpus_blocks = []

    for idx, uci in enumerate(ucis, start=1):
        w_count = uci.word_count()
        total_words += w_count
        all_keys.update(uci.variables.keys())

        if w_count == 0:
            report.empty_ucis.append(idx)
        elif w_count < min_words_warning:
            report.short_ucis.append((idx, w_count))

        corpus_blocks.append(uci.to_iramuteq_block())

    report.total_words = total_words
    report.avg_words_per_uci = round(total_words / max(report.total_ucis, 1), 1)
    report.variable_keys = sorted(list(all_keys))

    full_corpus_text = "\n".join(corpus_blocks).strip() + "\n"
    return full_corpus_text, report


def split_corpus_by_language(ucis: List[UCIUnit]) -> Dict[str, List[UCIUnit]]:
    """Groups UCIs by language if language variable is detected in metadata."""
    lang_groups: Dict[str, List[UCIUnit]] = {}
    for uci in ucis:
        found_lang = "general"
        for k, v in uci.variables.items():
            if "lang" in k.lower():
                found_lang = sanitize_variable_value(v)
                break
        if found_lang not in lang_groups:
            lang_groups[found_lang] = []
        lang_groups[found_lang].append(uci)
    return lang_groups


def filter_or_merge_short_ucis(
    ucis: List[UCIUnit],
    min_word_threshold: int = 0,
    action: str = "keep"
) -> List[UCIUnit]:
    """Handles micro-entries to avoid Reinert segmentation failures."""
    if min_word_threshold <= 0 or action == "keep":
        return ucis

    if action == "filter":
        return [u for u in ucis if u.word_count() >= min_word_threshold]

    if action == "merge":
        merged = []
        buffer_uci: Optional[UCIUnit] = None
        for uci in ucis:
            if buffer_uci is None:
                buffer_uci = uci
            else:
                if buffer_uci.word_count() < min_word_threshold:
                    buffer_uci.body_text = f"{buffer_uci.body_text} . {uci.body_text}"
                else:
                    merged.append(buffer_uci)
                    buffer_uci = uci
        if buffer_uci:
            merged.append(buffer_uci)
        return merged

    return ucis


def extract_text_from_docx(file_bytes_or_path) -> List[Tuple[str, str]]:
    if docx is None:
        raise ImportError("python-docx is required to read .docx files.")
    doc = docx.Document(file_bytes_or_path)
    paragraphs = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt:
            style = p.style.name if p.style else "Normal"
            paragraphs.append((style, txt))
    return paragraphs


def extract_text_from_pdf(file_bytes_or_path) -> List[Tuple[int, str]]:
    if pypdf is None:
        raise ImportError("pypdf is required to read PDF files.")
    reader = pypdf.PdfReader(file_bytes_or_path)
    pages = []
    for idx, page in enumerate(reader.pages, start=1):
        txt = page.extract_text() or ""
        txt = txt.strip()
        if txt:
            pages.append((idx, txt))
    return pages


def split_text_by_pattern(full_text: str, pattern: str) -> List[Tuple[str, str]]:
    if not pattern:
        return [("1", full_text.strip())]
    regex = re.compile(pattern, re.IGNORECASE | re.MULTILINE)
    splits = regex.split(full_text)
    matches = regex.findall(full_text)
    if len(splits) == 1:
        return [("1", full_text.strip())]
    sections = []
    preamble = splits[0].strip()
    if preamble:
        sections.append(("intro", preamble))
    for idx, body in enumerate(splits[1:], start=1):
        tag = sanitize_variable_value(matches[idx - 1] if idx - 1 < len(matches) else f"sec_{idx}")
        b_clean = body.strip()
        if b_clean:
            sections.append((tag, b_clean))
    return sections


# =============================================================================
# Streamlit Interface
# =============================================================================

st.set_page_config(
    page_title="IraPrep — Elegant IRaMuTeQ Studio",
    page_icon="✨",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        color: #27272A;
        background-color: #FAFAF9;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container {
        padding-top: 2rem;
        padding-bottom: 4rem;
        max-width: 760px;
    }

    .brand-title {
        font-size: 2.2rem;
        font-weight: 700;
        letter-spacing: -0.03em;
        color: #18181B;
        text-align: center;
        margin-bottom: 0.2rem;
    }
    .brand-subtitle {
        font-size: 0.95rem;
        color: #71717A;
        text-align: center;
        font-weight: 400;
        margin-bottom: 2rem;
    }
    .chic-badge {
        display: inline-block;
        background: #F4EBE8;
        color: #8C4A51;
        font-size: 0.72rem;
        font-weight: 600;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        padding: 4px 12px;
        border-radius: 999px;
        margin-bottom: 0.8rem;
    }

    .step-card {
        background: #FFFFFF;
        border: 1px solid #F0EEEC;
        border-radius: 18px;
        padding: 24px 28px;
        margin-bottom: 24px;
        box-shadow: 0 4px 20px -4px rgba(0, 0, 0, 0.03);
    }
    .step-number {
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        color: #A1A1AA;
        margin-bottom: 4px;
    }
    .step-heading {
        font-size: 1.15rem;
        font-weight: 600;
        color: #18181B;
        margin-bottom: 14px;
    }

    .metric-container {
        display: flex;
        justify-content: space-between;
        background: #FDFBFB;
        border: 1px solid #F5EFEB;
        border-radius: 14px;
        padding: 16px 20px;
        margin: 16px 0;
    }
    .metric-item {
        text-align: center;
        flex: 1;
    }
    .metric-val {
        font-size: 1.5rem;
        font-weight: 700;
        color: #18181B;
    }
    .metric-lbl {
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        color: #A1A1AA;
        margin-top: 2px;
    }

    .info-pill {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 10px 14px;
        font-size: 0.85rem;
        color: #475569;
        margin-bottom: 16px;
    }

    div.stButton > button {
        background-color: #18181B;
        color: #FFFFFF;
        border: none;
        border-radius: 12px;
        font-weight: 500;
        font-size: 0.95rem;
        padding: 0.6rem 1.4rem;
        transition: all 0.2s ease;
    }
    div.stButton > button:hover {
        background-color: #3F3F46;
        color: #FFFFFF;
        border: none;
        transform: translateY(-1px);
    }
</style>
""", unsafe_allow_html=True)

if "metadata_list" not in st.session_state:
    st.session_state.metadata_list = [{"key": "date", "val": "2024_11"}]

if "compound_list" not in st.session_state:
    st.session_state.compound_list = "mental health, harm reduction, substance use, artificial intelligence, climate finance"


def add_metadata_field():
    st.session_state.metadata_list.append({"key": "", "val": ""})


def remove_metadata_field(idx: int):
    if len(st.session_state.metadata_list) > 1:
        st.session_state.metadata_list.pop(idx)


# Header
st.markdown("""
<div style="text-align: center;">
    <span class="chic-badge">Private & Local • IRaMuTeQ Studio</span>
    <h1 class="brand-title">IraPrep</h1>
    <p class="brand-subtitle">Effortless, publication-ready corpus formatting for any document.</p>
</div>
""", unsafe_allow_html=True)

# Step 1: Upload
st.markdown("""
<div class="step-card">
    <div class="step-number">Step 01</div>
    <div class="step-heading">Choose your document</div>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Upload Word (.docx), Excel (.xlsx), CSV, PDF, or Text (.txt)",
    type=["docx", "xlsx", "xls", "csv", "tsv", "pdf", "txt"],
    accept_multiple_files=True,
    help="Upload a single file or a batch of files.",
    label_visibility="collapsed"
)

# Step 2: Configure
if uploaded_files:
    first_file = uploaded_files[0]
    fname = first_file.name.lower()
    is_spreadsheet = fname.endswith((".xlsx", ".xls", ".csv", ".tsv"))
    is_multidoc = len(uploaded_files) > 1

    st.markdown("""
    <div class="step-card">
        <div class="step-number">Step 02</div>
        <div class="step-heading">Configure your format & metadata</div>
    </div>
    """, unsafe_allow_html=True)

    if is_spreadsheet and not is_multidoc:
        if fname.endswith(".csv"):
            df = pd.read_csv(first_file)
        elif fname.endswith(".tsv"):
            df = pd.read_csv(first_file, sep="\t")
        else:
            df = pd.read_excel(first_file)

        st.markdown(f"""
        <div class="info-pill">
            💡 <b>Spreadsheet Mode:</b> Each row in your sheet represents <b>1 section (UCI)</b> in IRaMuTeQ.
            Your file has <b>{len(df):,} rows</b>, so it will generate <b>{len(df):,} sections</b>.
        </div>
        """, unsafe_allow_html=True)

        str_lengths = {}
        for c in df.columns:
            if df[c].dtype == object or pd.api.types.is_string_dtype(df[c]):
                avg_len = df[c].astype(str).str.len().mean()
                str_lengths[c] = avg_len

        candidate_cols = [c for c in df.columns if c.lower() in ["full_text", "text", "body", "transcript", "speech", "content"]]
        if candidate_cols:
            auto_text_col = candidate_cols[0]
        else:
            auto_text_col = max(str_lengths, key=str_lengths.get) if str_lengths else df.columns[0]

        auto_meta_cols = [
            c for c in df.columns 
            if c != auto_text_col and df[c].nunique() < 100
        ]

        col_t1, col_t2 = st.columns([2, 1])
        with col_t1:
            chosen_text_col = st.selectbox(
                "Text Column (Speech / Content):",
                options=list(df.columns),
                index=list(df.columns).index(auto_text_col) if auto_text_col in df.columns else 0,
                help="Choose the column containing the full natural sentences."
            )
        with col_t2:
            st.caption("✨ Prioritize natural sentences for rich Reinert clustering.")

        st.markdown("#### 🏷️ Metadata Variables (`*var_val`)")
        chosen_meta_cols = st.multiselect(
            "Selected Columns:",
            options=[c for c in df.columns if c != chosen_text_col],
            default=auto_meta_cols[:4],
            label_visibility="collapsed"
        )

        st.markdown("#### ➕ Additional Global Metadata (Optional)")
        for idx, item in enumerate(st.session_state.metadata_list):
            m_c1, m_c2, m_c3 = st.columns([1, 1, 0.2])
            with m_c1:
                item["key"] = st.text_input("Name", value=item["key"], placeholder="e.g. date", key=f"m_key_{idx}", label_visibility="collapsed")
            with m_c2:
                item["val"] = st.text_input("Value", value=item["val"], placeholder="e.g. 2024_11", key=f"m_val_{idx}", label_visibility="collapsed")
            with m_c3:
                if len(st.session_state.metadata_list) > 1:
                    if st.button("✕", key=f"del_{idx}"):
                        remove_metadata_field(idx)
                        st.rerun()

        st.button("+ Add Another Metadata Variable", on_click=add_metadata_field)

    else:
        st.markdown(f"""
        <div class="info-pill">
            📄 <b>Document Mode:</b> Processing <b>{len(uploaded_files)}</b> document(s). You can define whether each document is one whole section, or divided by speaker/question.
        </div>
        """, unsafe_allow_html=True)

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            split_method = st.selectbox(
                "How should sections be separated?",
                [
                    "Entire Document = 1 Section (e.g. 1 interview per file)",
                    "By Speaker Turns (e.g. Participant:, Speaker A:)",
                    "By Document Headings (Heading 1, 2)",
                    "By Custom Separator (---, ===)"
                ]
            )

        with col_s2:
            if "Speaker" in split_method:
                speaker_pat = st.text_input("Speaker Pattern (Regex):", value=r"(?:Participant|Speaker|Interviewer|Person)\s*[A-Za-z0-9_-]*:")
            elif "Custom" in split_method:
                custom_delim = st.text_input("Separator Text:", value="---")
            else:
                st.caption("✨ Clean separation applied automatically")

        st.markdown("#### 🏷️ Metadata Variables (`*var_val`)")
        for idx, item in enumerate(st.session_state.metadata_list):
            m_c1, m_c2, m_c3 = st.columns([1, 1, 0.2])
            with m_c1:
                item["key"] = st.text_input("Name", value=item["key"], placeholder="e.g. date", key=f"doc_m_key_{idx}", label_visibility="collapsed")
            with m_c2:
                item["val"] = st.text_input("Value", value=item["val"], placeholder="e.g. 2024_11", key=f"doc_m_val_{idx}", label_visibility="collapsed")
            with m_c3:
                if len(st.session_state.metadata_list) > 1:
                    if st.button("✕", key=f"doc_del_{idx}"):
                        remove_metadata_field(idx)
                        st.rerun()

        st.button("+ Add Another Metadata Variable", on_click=add_metadata_field)

        if is_multidoc or fname.endswith(".docx"):
            extract_tags = st.checkbox(
                "Auto-extract tags from filenames (e.g. P01_Female_30.docx)",
                value=True
            )
        else:
            extract_tags = False

    # Step 3: Refine & Export
    st.markdown("""
    <div class="step-card">
        <div class="step-number">Step 03</div>
        <div class="step-heading">Refine & Export</div>
    </div>
    """, unsafe_allow_html=True)

    with st.expander("🛠️ Advanced Cleaning & Lexicon Refinements", expanded=True):
        c_ref1, c_ref2 = st.columns(2)
        with c_ref1:
            strip_counts = st.checkbox("Strip count tags like (1), (2)", value=True, help="Removes numeric noise from frequency extractions")
            norm_symbols = st.checkbox("Convert & / ° symbols to words", value=True, help="Changes 1.5°c -> 1_5_degrees_c, & -> and")
        with c_ref2:
            short_handling = st.selectbox(
                "Short Entry Strategy:",
                [
                    "Keep all entries as-is",
                    "Filter out entries under 10 words",
                    "Merge consecutive short entries"
                ],
                help="IRaMuTeQ's Reinert method prefers at least 20-30 words per section."
            )

        st.caption("Word Binding (locutions composées) with `_`:")
        compounds_raw = st.text_input(
            "Compound words (comma-separated):",
            value=st.session_state.compound_list,
            label_visibility="collapsed"
        )
        st.session_state.compound_list = compounds_raw

    compounds = [c.strip() for c in st.session_state.compound_list.split(",") if c.strip()]

    # Generate Button
    if st.button("Generate IRaMuTeQ Corpus ✨", type="primary", use_container_width=True):
        generated_ucis: List[UCIUnit] = []

        with st.spinner("Processing, sanitizing, and compiling your corpus..."):
            global_meta = {
                sanitize_variable_name(m["key"]): sanitize_variable_value(m["val"])
                for m in st.session_state.metadata_list
                if m["key"].strip()
            }

            if is_spreadsheet and not is_multidoc:
                for _, row in df.iterrows():
                    raw_text = str(row[chosen_text_col]) if pd.notna(row[chosen_text_col]) else ""
                    clean_body = clean_text_for_iramuteq(
                        raw_text,
                        compound_words=compounds,
                        strip_parenthetical_counts=strip_counts,
                        normalize_symbols=norm_symbols
                    )
                    if not clean_body:
                        continue

                    row_vars = dict(global_meta)
                    for col in chosen_meta_cols:
                        v_name = sanitize_variable_name(col)
                        v_val = sanitize_variable_value(row[col])
                        row_vars[v_name] = v_val

                    generated_ucis.append(UCIUnit(row_vars, clean_body, first_file.name))

            else:
                for doc_file in uploaded_files:
                    f_name = doc_file.name
                    f_base = f_name.rsplit(".", 1)[0]
                    file_bytes = io.BytesIO(doc_file.read())

                    doc_vars = dict(global_meta)
                    if extract_tags:
                        parts = [p for p in re.split(r"[\s\-_]+", f_base) if p]
                        for i, p in enumerate(parts):
                            doc_vars[f"tag_{i+1}"] = sanitize_variable_value(p)

                    if f_name.lower().endswith(".docx"):
                        paragraphs = extract_text_from_docx(file_bytes)
                        if "Entire Document" in split_method:
                            full_raw = "\n\n".join([t for _, t in paragraphs])
                            clean_b = clean_text_for_iramuteq(full_raw, compound_words=compounds, strip_parenthetical_counts=strip_counts, normalize_symbols=norm_symbols)
                            if clean_b:
                                generated_ucis.append(UCIUnit(doc_vars, clean_b, f_name))
                        elif "Speaker" in split_method:
                            full_raw = "\n\n".join([t for _, t in paragraphs])
                            sections = split_text_by_pattern(full_raw, speaker_pat)
                            for s_tag, s_body in sections:
                                clean_b = clean_text_for_iramuteq(s_body, compound_words=compounds, strip_parenthetical_counts=strip_counts, normalize_symbols=norm_symbols)
                                if clean_b:
                                    s_vars = dict(doc_vars)
                                    s_vars["speaker"] = sanitize_variable_value(s_tag)
                                    generated_ucis.append(UCIUnit(s_vars, clean_b, f_name))
                        elif "Headings" in split_method:
                            cur_body, cur_heading = [], "intro"
                            for style, txt in paragraphs:
                                if "Heading" in style:
                                    if cur_body:
                                        clean_b = clean_text_for_iramuteq("\n\n".join(cur_body), compound_words=compounds)
                                        if clean_b:
                                            h_vars = dict(doc_vars)
                                            h_vars["section"] = sanitize_variable_value(cur_heading)
                                            generated_ucis.append(UCIUnit(h_vars, clean_b, f_name))
                                        cur_body = []
                                    cur_heading = txt[:25]
                                else:
                                    cur_body.append(txt)
                            if cur_body:
                                clean_b = clean_text_for_iramuteq("\n\n".join(cur_body), compound_words=compounds)
                                if clean_b:
                                    h_vars = dict(doc_vars)
                                    h_vars["section"] = sanitize_variable_value(cur_heading)
                                    generated_ucis.append(UCIUnit(h_vars, clean_b, f_name))

                    elif f_name.lower().endswith(".pdf"):
                        pages = extract_text_from_pdf(file_bytes)
                        for page_num, p_text in pages:
                            clean_b = clean_text_for_iramuteq(p_text, compound_words=compounds)
                            if clean_b:
                                p_vars = dict(doc_vars)
                                p_vars["page"] = str(page_num)
                                generated_ucis.append(UCIUnit(p_vars, clean_b, f_name))

                    else:
                        txt_str = file_bytes.getvalue().decode("utf-8", errors="replace")
                        clean_b = clean_text_for_iramuteq(txt_str, compound_words=compounds)
                        if clean_b:
                            generated_ucis.append(UCIUnit(doc_vars, clean_b, f_name))

        action_map = {
            "Keep all entries as-is": "keep",
            "Filter out entries under 10 words": "filter",
            "Merge consecutive short entries": "merge"
        }
        generated_ucis = filter_or_merge_short_ucis(
            generated_ucis,
            min_word_threshold=10,
            action=action_map.get(short_handling, "keep")
        )

        if generated_ucis:
            corpus_output, report = compile_corpus(generated_ucis)

            st.markdown(f"""
            <div class="metric-container">
                <div class="metric-item">
                    <div class="metric-val">{report.total_ucis:,}</div>
                    <div class="metric-lbl">Total Units</div>
                </div>
                <div class="metric-item">
                    <div class="metric-val">{report.total_words:,}</div>
                    <div class="metric-lbl">Words</div>
                </div>
                <div class="metric-item">
                    <div class="metric-val">{report.avg_words_per_uci}</div>
                    <div class="metric-lbl">Avg Words / Unit</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            with st.expander("👀 View First Formatted Section", expanded=True):
                preview_block = generated_ucis[0].to_iramuteq_block()
                st.code(preview_block, language="text")

            lang_groups = split_corpus_by_language(generated_ucis)

            if len(lang_groups) > 1 and "general" not in lang_groups:
                st.markdown("### 🌐 Multi-Language Corpus Detected")
                st.info(
                    "IRaMuTeQ requires selecting a **single dictionary language** per analysis. "
                    "We've created separate, ready-to-run files for each language:"
                )
                l_cols = st.columns(len(lang_groups) + 1)
                with l_cols[0]:
                    st.download_button(
                        label=f"💾 All ({len(generated_ucis)})",
                        data=corpus_output.encode("utf-8"),
                        file_name="corpus_iramuteq_all.txt",
                        mime="text/plain",
                        use_container_width=True
                    )

                for idx, (lang_name, l_ucis) in enumerate(lang_groups.items(), start=1):
                    l_corpus, _ = compile_corpus(l_ucis)
                    with l_cols[idx]:
                        st.download_button(
                            label=f"📥 {lang_name.title()} ({len(l_ucis):,})",
                            data=l_corpus.encode("utf-8"),
                            file_name=f"corpus_{lang_name}.txt",
                            mime="text/plain",
                            type="primary",
                            use_container_width=True
                        )
            else:
                st.download_button(
                    label="📥 Download corpus_iramuteq.txt",
                    data=corpus_output.encode("utf-8"),
                    file_name="corpus_iramuteq.txt",
                    mime="text/plain",
                    type="primary",
                    use_container_width=True
                )
        else:
            st.warning("No text content could be extracted. Please check the uploaded file.")

else:
    st.markdown("""
    <div style="text-align: center; color: #A1A1AA; font-size: 0.9rem; margin-top: 1.5rem;">
        Drop any file above to get started. Everything is processed locally on your device.
    </div>
    """, unsafe_allow_html=True)
