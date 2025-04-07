import streamlit as st
import pandas as pd
import spacy
import os
from spellchecker import SpellChecker
from docx import Document
from pdfminer.high_level import extract_text

# Load spaCy model
nlp = spacy.load("en_core_web_sm")

# Define required sections
required_sections = ["Executive Summary", "Technical Approach", "Methodology", "Key Personnel", "Workplan", "Budget"]

# Helper functions
def extract_text_from_pdf(pdf_file):
    return extract_text(pdf_file)

def extract_text_from_docx(docx_file):
    doc = Document(docx_file)
    full_text = "\n".join([para.text for para in doc.paragraphs])
    return full_text

def analyze_text(text):
    doc = nlp(text)
    sentences = list(doc.sents)
    word_count = len([token.text for token in doc if token.is_alpha])
    return len(sentences), word_count

def formatting_check(doc):
    font_ok = all(run.font.name == 'Tenorite' for para in doc.paragraphs for run in para.runs if run.font.name)
    font_size_ok = all(
        run.font.size and run.font.size.pt == 11
        for para in doc.paragraphs for run in para.runs if run.font.size
    )
    text = "\n".join([para.text for para in doc.paragraphs])
    spell = SpellChecker()
    words = text.split()
    misspelled = spell.unknown(words)
    return {
        'font_ok': font_ok,
        'font_size_ok': font_size_ok,
        'spelling_issues': list(misspelled)
    }

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
    max_score = 4

    total_score += section_percentage * 0.50

    spell_score = 0
    if len(formatting_results['spelling_issues']) == 0:
        spell_score = 100
    else:
        spell_score = max(0, 100 - len(formatting_results['spelling_issues']) * 10)
    total_score += spell_score * 0.25

    font_style_score = 100 if formatting_results['font_ok'] else 0
    font_size_score = 100 if formatting_results['font_size_ok'] else 0
    total_score += (font_style_score + font_size_score) * 0.25

    total_score = round(total_score)

    missing_sections = [sec for sec, present in section_results.items() if not present]
    recommendations = []
    if missing_sections:
        recommendations.append(f"Kindly include the following missing sections: {', '.join(missing_sections)}")
    if formatting_results['spelling_issues']:
        recommendations.append("Spelling issues found in the document.")
    if not formatting_results['font_ok']:
        recommendations.append("Document should use font 'Tenorite' throughout.")
    if not formatting_results['font_size_ok']:
        recommendations.append("Body text should use font size 11.")

    # Methodology component check
    methodology_components = [
        "project kick-off", "project inception",
        "desk review",
        "data collection",
        "data analysis", "data management",
        "report development"
    ]
    methodology_text = "\n".join(
        para.text for para in doc.paragraphs if "methodology" in para.text.lower()
    ).lower()

    missing_components = []
    for comp in methodology_components:
        if comp not in methodology_text:
            missing_components.append(comp)

    if missing_components:
        recommendations.append(
            f"The methodology section is missing the following components: {', '.join(set(missing_components)).title()}"
        )

    return {
        'sections': section_results,
        'score': total_score,
        'recommendations': recommendations,
        'formatting': formatting_results
    }

# Streamlit App Navigation
st.title("Proposal Evaluation and RFP Extractor")

page = st.radio("Go to", ("Proposal Evaluator", "RFP Extractor"))

if page == "Proposal Evaluator":
    st.header("📄 Proposal Evaluator")
    uploaded_file = st.file_uploader("Upload your proposal (DOCX only)", type=["docx"])

    if uploaded_file is not None:
        with open("temp.docx", "wb") as f:
            f.write(uploaded_file.getbuffer())

        doc = Document("temp.docx")
        text = extract_text_from_docx("temp.docx")

        evaluation = evaluate_proposal(text, required_sections, doc)

        st.subheader("Evaluation Results")
        st.write("**Sections Present:**")
        st.write(evaluation['sections'])

        st.write(f"**Total Score:** {evaluation['score']}%")

        st.subheader("Recommendations")
        for rec in evaluation['recommendations']:
            st.warning(rec)

elif page == "RFP Extractor":
    st.header("📄 RFP Extractor")
    uploaded_rfp = st.file_uploader("Upload an RFP Document (PDF)", type=["pdf"])

    if uploaded_rfp is not None:
        with open("temp.pdf", "wb") as f:
            f.write(uploaded_rfp.getbuffer())

        extracted_text = extract_text_from_pdf("temp.pdf")
        sentence_count, word_count = analyze_text(extracted_text)

        st.subheader("Extracted Text")
        st.text_area("Content:", extracted_text[:3000], height=300)

        st.subheader("Text Summary")
        st.write(f"Number of Sentences: {sentence_count}")
        st.write(f"Number of Words: {word_count}")
