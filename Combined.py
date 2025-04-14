import streamlit as st
import os
import tempfile
from docx import Document
import re
from io import BytesIO
from spellchecker import SpellChecker
import pandas as pd
import base64
from fuzzywuzzy import fuzz
import fitz  # PyMuPDF
from nltk.tokenize import sent_tokenize
import nltk
nltk_data_path = os.path.join(os.getcwd(), "nltk_data")
nltk.data.path.append(nltk_data_path)


# --- Constants ---
STANDARD_SECTIONS = [
    "Table of content", "Table of Contents", "contents", "Content",
    "Introduction", "Background", "Objective",
    "Methodology", "Approach", "technical approach",
    "Project Team", "About Sahel", "Budget", "Work Plan"
]

# --- Helper Functions ---
def extract_text(file):
    text = ""
    if file.name.endswith('.docx'):
        temp_path = os.path.join(tempfile.gettempdir(), file.name)
        with open(temp_path, 'wb') as f:
            f.write(file.read())
        if not is_valid_docx(temp_path):
            raise ValueError("The uploaded file is not a valid .docx file.")
        doc = Document(temp_path)
        for para in doc.paragraphs:
            text += para.text.strip().lower() + '\n'
    else:
        raise ValueError("Unsupported file type.")
    return text

def is_valid_docx(file_path):
    try:
        Document(file_path)
        return True
    except Exception:
        return False

def extract_text_with_formatting(file):
    """Extract text and formatting attributes from an RFP file (.docx or .pdf)."""
    text_with_formatting = []
    is_pdf = False

    if file.name.endswith(".docx"):
        # Extract text paragraph by paragraph for Word documents
        temp_path = os.path.join(tempfile.gettempdir(), file.name)
        with open(temp_path, "wb") as f:
            f.write(file.read())

        # Validate the file
        if not is_valid_docx(temp_path):
            raise ValueError("The uploaded file is not a valid .docx file. Please ensure it is properly formatted.")

        doc = Document(temp_path)

        # Extract text from paragraphs
        for para in doc.paragraphs:
            text_with_formatting.append({
                "text": para.text.strip(),
                "bold": any(run.bold for run in para.runs),
                "style": para.style.name if para.style else None
            })

        # Extract text from tables (if any)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        text_with_formatting.append({
                            "text": para.text.strip(),
                            "bold": any(run.bold for run in para.runs),
                            "style": para.style.name if para.style else None
                        })

    elif file.name.endswith(".pdf"):
        # Extract text line by line for PDFs
        is_pdf = True
        with fitz.open(stream=file.read(), filetype="pdf") as pdf:
            for page in pdf:
                for line in page.get_text("text").split("\n"):
                    text_with_formatting.append({
                        "text": line.strip(),
                        "bold": False,
                        "style": None
                    })

    else:
        raise ValueError("Unsupported file type. Please upload a .docx or .pdf file.")

    return text_with_formatting, is_pdf

import re

def extract_rfp_expectations(text_with_formatting, is_pdf=False):
    """Extract expectations from the RFP with hierarchical structure and formatting cues."""
    expectations = []
    current_section = None
    current_subsection = None
    current_content = []
    section_hierarchy = []

    # Regex for multi-level numbering (e.g., 1., 1.1, 1.1.1)
    numbering_pattern = re.compile(r"^\d+(\.\d+)*\s")

    def save_current_section():
        """Save the current section and its content."""
        if current_section and current_content:
            expectations.append({
                "section": " > ".join(section_hierarchy),
                "content": " ".join(current_content).strip()
            })

    for item in text_with_formatting:
        text = item["text"]
        bold = item.get("bold", False)
        style = item.get("style", "")
        is_numbered = bool(numbering_pattern.match(text))

        if not text:  # Skip empty lines
            continue

        # Detect section headers (e.g., bold text, numbered headers, or specific styles)
        if bold or is_numbered or style in ['Heading 1', 'Heading 2', 'Heading 3']:
            # Save the current section before moving to the next
            save_current_section()
            current_content = []

            # Update the section hierarchy
            if is_numbered:
                # Extract the numbering level (e.g., 1., 1.1)
                numbering = numbering_pattern.match(text).group().strip()
                level = numbering.count(".") + 1

                # Adjust the hierarchy based on the level
                while len(section_hierarchy) >= level:
                    section_hierarchy.pop()
                section_hierarchy.append(text.strip())
            else:
                # Treat bold or styled text as a new top-level section
                section_hierarchy = [text.strip()]

            current_section = text.strip()
            continue

        # Combine lines that are part of the same bullet or paragraph
        if current_content and not text.startswith("-") and not is_numbered:
            current_content[-1] += " " + text.strip()
        else:
            current_content.append(text.strip())

    # Save the last section
    save_current_section()

    return expectations

def check_expectations_coverage(expectations, proposal_text):
    addressed = []
    missing = []
    proposal_sentences = [s.strip().lower() for s in sent_tokenize(proposal_text) if s.strip()]
    for exp in expectations:
        exp_text = exp["content"].lower()
        best_score = max([fuzz.partial_ratio(exp_text, sentence) for sentence in proposal_sentences] or [0])
        if best_score >= 70:
            addressed.append({"expectation": exp})
        else:
            missing.append({"expectation": exp})
    score = (len(addressed) / len(expectations)) * 100 if expectations else 0
    return score, addressed, missing

def evaluate_proposal(text, required_sections, doc):
    lower_text = text.lower()
    section_results = {}
    for sec in required_sections:
        found = any(sec.lower() in para.text.lower() for para in doc.paragraphs)
        section_results[sec] = found
    section_score = sum(section_results.values())
    section_percentage = (section_score / len(required_sections)) * 100
    formatting_results = formatting_check(doc)
    total_score = 0
    max_score = 100
    methodology_components = [
        "project kick-off", "project inception", "desk review",
        "data collection", "data analysis", "data management",
        "report development", "deliverables", "output", "outputs"
    ]
    section_weight = 0.35
    total_score += section_percentage * section_weight
    spelling_weight = 0.20
    spell_score = 100 if not formatting_results['spelling_issues'] else max(0, 100 - len(formatting_results['spelling_issues']) * 10)
    total_score += spell_score * spelling_weight
    methodology_weight = 0.25
    methodology_text = "\n".join(
        para.text for para in doc.paragraphs if "methodology" in para.text.lower() or "approach" in para.text.lower()
    ).lower()
    missing_components = [comp for comp in methodology_components if comp not in methodology_text]
    methodology_score = 100 if not missing_components else 100 - (len(missing_components) * 10)
    total_score += methodology_score * methodology_weight
    formatting_weight = 0.20
    font_style_score = 100 if formatting_results['font_ok'] else 0
    font_size_score = 100 if formatting_results['font_size_ok'] else 0
    formatting_score = (font_style_score + font_size_score) / 2
    total_score += round(formatting_score * formatting_weight)
    missing_sections = [sec for sec, present in section_results.items() if not present]
    recommendations = []
    if missing_sections:
        recommendations.append(f"Kindly include the following missing sections: {', '.join(missing_sections)}")
    if formatting_results['spelling_issues']:
        recommendations.append("Spelling issues found in the document.")
    if not formatting_results['font_ok']:
        recommendations.append("Document should use font 'Tenorite' or 'Candara' throughout.")
    if not formatting_results['font_size_ok']:
        recommendations.append("Body text should use font size 11.")
    if missing_components:
        recommendations.append(f"The methodology section is missing the following components: {', '.join(set(missing_components)).title()}")
    return {
        'sections': section_results,
        'score': total_score,
        'recommendations': recommendations,
        'formatting': formatting_results
    }

def formatting_check(doc):
    spell = SpellChecker()
    text = "\n".join([para.text for para in doc.paragraphs])
    words = re.findall(r'\b\w+\b', text.lower())
    misspelled = spell.unknown(words)
    spelling_issues = list(misspelled)[:15]
    font_sizes = []
    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.size:
                font_sizes.append(run.font.size.pt)
    body_font_size = max(set(font_sizes), key=font_sizes.count) if font_sizes else 11
    font_ok = True
    font_size_ok = True
    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.name and run.font.name.lower() not in ["tenorite", "candara"]:
                font_ok = False
            if run.font.size and run.font.size.pt != body_font_size:
                if para.style.name not in ['Heading 1', 'Heading 2', 'Heading 3']:
                    font_size_ok = False
        if not font_ok or not font_size_ok:
            break
    return {
        "spelling_issues": spelling_issues,
        "font_ok": font_ok,
        "font_size_ok": font_size_ok
    }

def create_word_report(evaluation, rfp_score=None, rfp_missing=None):
    doc = Document()
    doc.add_heading("Proposal Evaluation Report", level=1)
    doc.add_heading("Section Check", level=2)
    for section, found in evaluation['sections'].items():
        doc.add_paragraph(f"{section}: {'Present' if found else 'Missing'}")
    doc.add_heading("Formatting & Presentation", level=2)
    if evaluation['formatting']['spelling_issues']:
        doc.add_paragraph("Spelling Issues Detected:")
        doc.add_paragraph(", ".join(evaluation['formatting']['spelling_issues']))
    else:
        doc.add_paragraph("No major spelling issues detected.")
    if evaluation['formatting']['font_ok'] and evaluation['formatting']['font_size_ok']:
        doc.add_paragraph("Font style and size meet organizational standards (Tenorite or Candara, size 11).")
    else:
        doc.add_paragraph("Font style does not match standard (Tenorite or Candara) or font size is not 11 in body text.")
    doc.add_heading("Overall Score", level=2)
    doc.add_paragraph(f"{evaluation['score']}%")
    doc.add_heading("Recommendations", level=2)
    if evaluation['recommendations']:
        for rec in evaluation['recommendations']:
            doc.add_paragraph(f"- {rec}")
    else:
        doc.add_paragraph("All criteria met. Great job!")
    if rfp_score is not None and rfp_missing:
        doc.add_heading("Missing RFP Expectations", level=2)
        doc.add_paragraph("The following expectations from the RFP were not addressed in the proposal:")
        for miss in rfp_missing:
            if isinstance(miss, dict) and 'expectation' in miss and 'section' in miss['expectation']:
                doc.add_paragraph(f"- {miss['expectation']['content']} (Section: {miss['expectation']['section']})")
            else:
                doc.add_paragraph(f"- {miss}")
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def truncate_text(text, max_words=25):
    words = text.split()
    return " ".join(words[:max_words]) + ("..." if len(words) > max_words else "")

# --- Streamlit Interface ---
st.set_page_config(page_title="Strategy Unit Toolkit", page_icon=":briefcase:", layout="wide")

current_dir = os.path.dirname(__file__)
with open(os.path.join(current_dir, "background.jpg"), "rb") as file:
    encoded_string = base64.b64encode(file.read()).decode()
    st.markdown(f"""
        <style>
        .stApp {{
            background-image: linear-gradient(rgba(255, 255, 255, 0.94), rgba(255, 255, 255, 0.94)),
            url("data:image/jpg;base64,{encoded_string}");
            background-attachment: fixed;
            background-size: cover;
            background-repeat: no-repeat;
            background-position: center;
        }}
        </style>
    """, unsafe_allow_html=True)

st.image("Sahel Consulting (Official).png", width=300)
st.title(":green[Strategy Unit Toolkit]")
st.write(":orange[Welcome! Upload the Proposal and RFP to evaluate alignment and quality.]")

uploaded_proposal = st.file_uploader("Upload Proposal (.docx only)", type=["docx"])
uploaded_rfp = st.file_uploader("Upload RFP (.docx or .pdf)", type=["docx", "pdf"])

evaluation = None
rfp_score = None
rfp_missing = []
rfp_addressed = []
org_score = None

if uploaded_proposal:
    try:
        prop_text = extract_text(uploaded_proposal)
    except ValueError as e:
        st.error(f"Error: {e}")
    except Exception as e:
        st.error("An unexpected error occurred while processing the file.")

    if st.button("Evaluate Proposal"):
        st.success("Proposal uploaded successfully.")
        doc = Document(uploaded_proposal)

        if uploaded_rfp:
            with st.spinner("Processing RFP..."):
                try:
                    rfp_text_with_formatting, is_pdf = extract_text_with_formatting(uploaded_rfp)
                    rfp_expectations = extract_rfp_expectations(rfp_text_with_formatting, is_pdf=is_pdf)
                except ValueError as e:
                    st.error(f"Error: {e}")
                except Exception as e:
                    st.error(f"An unexpected error occurred: {e}")

            with st.spinner("Checking alignment with RFP..."):
                try:
                    rfp_score, rfp_addressed, rfp_missing = check_expectations_coverage(rfp_expectations, prop_text)
                except ValueError as e:
                    st.error(f"Error: {e}")
                except Exception as e:
                    st.error(f"An unexpected error occurred: {e}")

        with st.spinner("Evaluating proposal against organizational standards..."):
            evaluation = evaluate_proposal(prop_text, STANDARD_SECTIONS, doc)
            org_score = evaluation['score']

if evaluation or rfp_score is not None:
    st.subheader("Evaluation Results")

    if rfp_score is not None:
        st.write("### RFP Alignment")
        st.info(f"RFP Coverage Score: **{round(rfp_score)}%**")

        if rfp_addressed:
            st.success("Addressed Expectations from RFP:")
            for addr in rfp_addressed:
                truncated = truncate_text(addr['expectation']['content'])
                st.write(f"- **{truncated}** (Section: {addr['expectation']['section']})")

        if rfp_missing:
            st.warning("Missing Expectations from RFP:")
            for miss in rfp_missing:
                truncated = truncate_text(miss['expectation']['content'])
                st.write(f"- **{truncated}** (Section: {miss['expectation']['section']})")

    if evaluation:
        st.write("### Proposal Evaluation Against Organizational Standards")
        st.info(f"Organizational Standards Score: **{round(org_score)}%**")

        st.write("### Section Check")
        for section, found in evaluation['sections'].items():
            st.write(f"- **{section}**: {'✅' if found else '❌'}")

        st.write("### Formatting & Presentation")
        if evaluation['formatting']['spelling_issues']:
            st.warning("Spelling Issues Detected:")
            st.write(", ".join(evaluation['formatting']['spelling_issues']))
        else:
            st.success("No major spelling issues detected.")

        if evaluation['formatting']['font_ok'] and evaluation['formatting']['font_size_ok']:
            st.success("Font style and size meet organizational standards (Tenorite or Candara, size 11).")
        else:
            st.warning("Font style or font size issue detected.")

        st.write("### Recommendations")
        if evaluation['recommendations']:
            for rec in evaluation['recommendations']:
                truncated_rec = truncate_text(rec)  # Truncate the recommendation
                st.warning(truncated_rec)
        else:
            st.success("All criteria met. Great job!")

        if rfp_missing:
            st.warning("The following expectations from the RFP were not addressed in the proposal:")
            for miss in rfp_missing:
                st.write(f"- **{miss['expectation']['content']}** (Section: {miss['expectation']['section']})")
        else:
            st.success("Your proposal aligns well with the RFP expectations!")

    word_buffer = create_word_report(
        evaluation,
        rfp_score,
        rfp_missing
    )
    st.download_button(
        label="Download Evaluation Report (.docx)",
        data=word_buffer,
        file_name="proposal_evaluation.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
