import streamlit as st
import pandas as pd
import json
import io
from openai import OpenAI
from docx import Document
from PyPDF2 import PdfReader

st.set_page_config(page_title="GenAI Case Study Evaluator", layout="wide")
st.title("📊 GenAI Case Study Evaluator")

api_key = st.text_input("OpenAI API Key", type="password")
custom_prompt = st.text_area("Additional Evaluator Instructions (Optional)")

case_file = st.file_uploader("Case Study Problem", type=["pdf","docx"])
rubric_file = st.file_uploader("Rubric (Excel)", type=["xlsx"])
submissions = st.file_uploader(
    "Participant Submissions",
    type=["pdf","docx"],
    accept_multiple_files=True
)

def read_pdf(uploaded_file):
    reader = PdfReader(uploaded_file)
    text = []
    for p in reader.pages:
        t = p.extract_text()
        if t:
            text.append(t)
    return "\n".join(text)

def read_docx(uploaded_file):
    doc = Document(uploaded_file)
    return "\n".join([p.text for p in doc.paragraphs])

def read_file(uploaded_file):
    if uploaded_file.name.lower().endswith(".pdf"):
        return read_pdf(uploaded_file)
    return read_docx(uploaded_file)

def rubric_to_text(df):
    lines = []
    for _, r in df.iterrows():
        lines.append(
            f"Criterion: {r['Criterion']}\n"
            f"Max Score: {r['Max Score']}\n"
            f"Description: {r['Description']}\n"
        )
    return "\n".join(lines)

SYSTEM_PROMPT = """
You are an expert evaluator for a GenAI-enabled business case study.

Evaluate ONLY against the rubric.
Do NOT compare against a model answer.
Different valid approaches may receive high scores.
Score only based on evidence found.

Return ONLY JSON:

{
  "evidence": {},
  "scores": {},
  "strengths": [],
  "improvements": [],
  "overall_rating": ""
}
"""

if st.button("Evaluate"):
    if not api_key or not case_file or not rubric_file or not submissions:
        st.error("Please provide API key, case study, rubric, and submissions.")
        st.stop()

    client = OpenAI(api_key=api_key)

    case_text = read_file(case_file)

    rubric_df = pd.read_excel(rubric_file)
    rubric_text = rubric_to_text(rubric_df)

    results = []

    for submission in submissions:
        submission_text = read_file(submission)

        prompt = f"""
{custom_prompt}

CASE STUDY
{case_text}

RUBRIC
{rubric_text}

PARTICIPANT SUBMISSION
{submission_text}
"""

        response = client.chat.completions.create(
            model="gpt-4.1",
            response_format={"type":"json_object"},
            messages=[
                {"role":"system","content":SYSTEM_PROMPT},
                {"role":"user","content":prompt}
            ],
            temperature=0
        )

        data = json.loads(response.choices[0].message.content)

        row = {"Participant": submission.name}

        total = 0
        for _, r in rubric_df.iterrows():
            criterion = r["Criterion"]
            max_score = int(r["Max Score"])

            score = float(data.get("scores", {}).get(criterion, 0))
            score = max(0, min(score, max_score))

            row[criterion] = score
            total += score

        row["Total"] = total
        row["Rating"] = data.get("overall_rating", "")
        row["Strengths"] = "; ".join(data.get("strengths", []))
        row["Improvements"] = "; ".join(data.get("improvements", []))

        results.append(row)

    df = pd.DataFrame(results)

    st.dataframe(df, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)

    st.download_button(
        "Download Evaluation Report",
        output.getvalue(),
        file_name="evaluation_report.xlsx"
    )
