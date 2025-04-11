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
import fitz  # PyMuPDF for PDF processing

# --- Constants ---
STANDARD_SECTIONS = [
    "Table of content" or "Table of Contents" or "contents" or "Content",
    "Introduction",
    "Background",
    "Objective",
    "Methodology" or "Approach" or "technical approach",
    "Project Team",
    "About Sahel",
    "Budget",
    "Work Plan",
]

# --- Helper Functions ---
def extract_text(file):
    """Extract text from a Word document (.docx only)."""
    text = ""
    if file.name.endswith('.docx'):
        temp_path = os.path.join(tempfile.gettempdir(), file.name)
        with open(temp_path, 'wb') as f:
            f.write(file.read())

        # Validate the file
        if not is_valid_docx(temp_path):
            raise ValueError("The uploaded file is not a valid .docx file. Please ensure it is properly formatted.")

        doc = Document(temp_path)
        for para in doc.paragraphs:
            text += para.text.strip().lower() + '\n'  # Convert to lowercase
    else:
        raise ValueError("Unsupported file type. Please upload a .docx file.")
    return text

def is_valid_docx(file_path):
    """Check if the file is a valid .docx file."""
    try:
        Document(file_path)  # Try opening the file with python-docx
        return True
    except Exception:
        return False

def extract_text_with_formatting(file):
    """Extract text and formatting attributes from an RFP file (.docx or .pdf)."""
    text_with_formatting = []

    if file.name.endswith(".docx"):
        temp_path = os.path.join(tempfile.gettempdir(), file.name)
        with open(temp_path, "wb") as f:
            f.write(file.read())

        # Validate the file
        if not is_valid_docx(temp_path):
            raise ValueError("The uploaded file is not a valid .docx file. Please ensure it is properly formatted.")

        doc = Document(temp_path)

        # Extract text from paragraphs
        for para in doc.paragraphs:
            for run in para.runs:
                text_with_formatting.append({
                    "text": run.text.strip().lower(),  # Convert to lowercase
                    "bold": run.bold,
                    "font_size": run.font.size.pt if run.font.size else None
                })

        # Extract text from tables (if any)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            text_with_formatting.append({
                                "text": run.text.strip().lower(),  # Convert to lowercase
                                "bold": run.bold,
                                "font_size": run.font.size.pt if run.font.size else None
                            })

    elif file.name.endswith(".pdf"):
        # Extract plain text from PDF
        with fitz.open(stream=file.read(), filetype="pdf") as pdf:
            for page in pdf:
                for line in page.get_text("dict")["blocks"]:
                    if "lines" not in line or not line["lines"]:
                        continue
                    for span in line["lines"][0]["spans"]:
                        text_with_formatting.append({
                            "text": span["text"].strip().lower(),  # Convert to lowercase
                            "bold": False,  # PDFs don't provide bold information
                            "font_size": span["size"]  # Font size from PDF
                        })

    else:
        raise ValueError("Unsupported file type. Please upload a .docx or .pdf file.")

    return text_with_formatting

def extract_text_from_rfp(file):
    """Extract text from an RFP file (.docx only)."""
    text = ""
    if file.name.endswith(".docx"):
        temp_path = os.path.join(tempfile.gettempdir(), file.name)
        with open(temp_path, "wb") as f:
            f.write(file.read())
        doc = Document(temp_path)
        for para in doc.paragraphs:
            text += para.text + "\n"
    else:
        raise ValueError("Unsupported file type. Please upload a .docx file.")

    return text

def extract_rfp_expectations(text_with_formatting):
    """Extract expectations from the RFP with section headings for context."""
    expectations = []
    keywords = ["deliverable", "budget", "timeline", "expected", "scope of work", "methodology", "objective", "goal", "requirements", "outcomes"]

    current_section = "General"
    seen_expectations = set()  # To track duplicates

    # Determine the most common font size (assumed to be the body font size)
    font_sizes = [item["font_size"] for item in text_with_formatting if item["font_size"]]
    body_font_size = max(set(font_sizes), key=font_sizes.count) if font_sizes else 11  # Default to 11 if no font size info

    for item in text_with_formatting:
        text = item["text"]
        bold = item["bold"]
        font_size = item["font_size"]

        if not text:  # Skip empty lines
            continue

        # Detect section headings based on formatting (bold or larger font size than body text)
        if bold or (font_size and font_size > body_font_size):
            current_section = text.strip(":").title()
            continue

        # Detect expectations based on keywords
        if any(k in text.lower() for k in keywords):
            if text.lower() not in seen_expectations:  # Check for duplicates
                expectations.append({"section": current_section, "expectation": text})
                seen_expectations.add(text.lower())

    return expectations

def check_expectations_coverage(expectations, proposal_text):
    """Check if expectations from the RFP are addressed in the proposal using fuzzy matching."""
    missing = []
    addressed = []
    proposal_paragraphs = proposal_text.split("\n")  # Split proposal into paragraphs

    for exp in expectations:
        exp_text = exp["expectation"]  # Already in lowercase
        best_match_score = 0

        # Compare the expectation with each paragraph in the proposal
        for para in proposal_paragraphs:
            para_text = para  # Already in lowercase
            match_score = fuzz.partial_ratio(exp_text, para_text)
            if match_score > best_match_score:
                best_match_score = match_score

        # Determine if the expectation is addressed based on a threshold
        if best_match_score >= 70:  # Threshold for alignment
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
        "project kick-off" or "project inception", 
        "desk review", 
        "data collection",
        "data analysis",
        "data management",
         "report development",
        "deliverables" or "output" or "outputs"
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

    # Determine the most common font size (assumed to be the body font size)
    font_sizes = []
    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.size:
                font_sizes.append(run.font.size.pt)

    body_font_size = max(set(font_sizes), key=font_sizes.count) if font_sizes else 11  # Default to 11 if no font size info

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

    # Section Check
    doc.add_heading("Section Check", level=2)
    for section, found in evaluation['sections'].items():
        doc.add_paragraph(f"{section}: {'Present' if found else 'Missing'}")

    # Formatting & Presentation
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

    # Overall Score
    doc.add_heading("Overall Score", level=2)
    doc.add_paragraph(f"{evaluation['score']}%")

    # Recommendations
    doc.add_heading("Recommendations", level=2)
    if evaluation['recommendations']:
        for rec in evaluation['recommendations']:
            doc.add_paragraph(f"- {rec}")
    else:
        doc.add_paragraph("All criteria met. Great job!")

    # Include Missing Expectations in Recommendations
    if rfp_score is not None and rfp_missing:
        doc.add_heading("Missing RFP Expectations", level=2)
        doc.add_paragraph("The following expectations from the RFP were not addressed in the proposal:")
        for miss in rfp_missing:
            # Handle cases where `miss` is not a dictionary
            if isinstance(miss, dict) and 'expectation' in miss and 'section' in miss['expectation']:
                doc.add_paragraph(f"- {miss['expectation']['expectation']} (Section: {miss['expectation']['section']})")
            else:
                doc.add_paragraph(f"- {miss}")  # Fallback for unexpected data

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def truncate_text(text, max_words=25):
    """Truncate text to a maximum number of words for display purposes."""
    words = text.split()
    return " ".join(words[:max_words]) + ("..." if len(words) > max_words else "")

# --- Streamlit Interface ---
st.set_page_config(page_title="Strategy Unit Toolkit", page_icon=":briefcase:", layout="wide")

# Background Image
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

# File Uploaders
uploaded_proposal = st.file_uploader("Upload Proposal (.docx only)", type=["docx"])
uploaded_rfp = st.file_uploader("Upload RFP (.docx or .pdf)", type=["docx", "pdf"])

# Initialize variables
evaluation = None
rfp_score = None
rfp_missing = []
rfp_addressed = []
org_score = None

# --- Evaluate Proposal ---
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

        # Part 1: RFP Alignment
        if uploaded_rfp:
            with st.spinner("Processing RFP..."):
                try:
                    rfp_text_with_formatting = extract_text_with_formatting(uploaded_rfp)
                    rfp_expectations = extract_rfp_expectations(rfp_text_with_formatting)
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

        # Part 2: Proposal Evaluation Against Organizational Standards
        with st.spinner("Evaluating proposal against organizational standards..."):
            evaluation = evaluate_proposal(prop_text, STANDARD_SECTIONS, doc)
            org_score = evaluation['score']

# --- Display Results ---
if evaluation or rfp_score is not None:
    st.subheader("Evaluation Results")

    # Part 1: RFP Alignment
    if rfp_score is not None:
        st.write("### RFP Alignment")
        st.info(f"RFP Coverage Score: **{round(rfp_score)}%**")

        if rfp_addressed:
            st.success("Addressed Expectations from RFP:")
            for addr in rfp_addressed:
                truncated = truncate_text(addr['expectation']['expectation'])
                st.write(f"- **{truncated}** (Section: {addr['expectation']['section']})")

        if rfp_missing:
            st.warning("Missing Expectations from RFP:")
            for miss in rfp_missing:
                truncated = truncate_text(miss['expectation']['expectation'])
                st.write(f"- **{truncated}** (Section: {miss['expectation']['section']})")

    # Part 2: Proposal Evaluation Against Organizational Standards
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
                st.warning(rec)

        # Include Missing Expectations in Recommendations
        if rfp_missing:
            st.warning("The following expectations from the RFP were not addressed in the proposal:")
            for miss in rfp_missing:
                st.write(f"- **{miss['expectation']['expectation']}** (Section: {miss['expectation']['section']})")
        else:
            st.success("Your proposal aligns well with the RFP expectations!")

    # Download Evaluation Report
    word_buffer = create_word_report(
        evaluation,
        rfp_score,
        rfp_missing  # Pass the full `rfp_missing` list directly
    )
    st.download_button(
        label="Download Evaluation Report (.docx)",
        data=word_buffer,
        file_name="proposal_evaluation.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"    )
