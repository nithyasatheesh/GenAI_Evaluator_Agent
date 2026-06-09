# Updated app.py
# Changes:
# - Uses st.secrets['OPENAI_API_KEY']
# - No API key textbox
# - Rubric validation
# - Safer scoring
# - Total (100) column display

import streamlit as st
import pandas as pd
import json, io
from openai import OpenAI
from docx import Document
from PyPDF2 import PdfReader

st.set_page_config(page_title="GenAI Case Study Evaluator", layout="wide")
st.title("📊 GenAI Case Study Evaluator")

def read_pdf(file):
    reader = PdfReader(file)
    return "\n".join([p.extract_text() or "" for p in reader.pages])

def read_docx(file):
    doc = Document(file)
    return "\n".join([p.text for p in doc.paragraphs])

def read_file(file):
    return read_pdf(file) if file.name.lower().endswith(".pdf") else read_docx(file)

def rubric_to_text(df):
    rows = []
    for _, r in df.iterrows():
        rows.append(
            f"Criterion: {r['Criterion']}\nMax Score: {r['Max Score']}\nDescription: {r['Description']}"
        )
    return "\n\n".join(rows)

def rating(score):
    if score >= 90: return "Exceptional"
    if score >= 80: return "Proficient"
    if score >= 70: return "Competent"
    if score >= 60: return "Developing"
    return "Needs Improvement"

SYSTEM_PROMPT = '''
You are an expert evaluator for a GenAI business case study.
Evaluate ONLY against the rubric.
Do not compare against a model answer.
Return ONLY valid JSON:
{"scores":{}, "strengths":[], "improvements":[]}
'''

try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception:
    st.error("OPENAI_API_KEY not configured in Streamlit Secrets.")
    st.stop()

case_file = st.file_uploader("Case Study Problem", type=["pdf","docx"])
rubric_file = st.file_uploader("Rubric (Excel)", type=["xlsx"])
participant_files = st.file_uploader(
    "Participant Submissions", type=["pdf","docx"], accept_multiple_files=True
)

custom_prompt = st.text_area(
    "Custom Evaluator Instructions",
    placeholder="Example: Be stricter on governance and validation scoring."
)

if st.button("Evaluate"):
    rubric_df = pd.read_excel(rubric_file)

    required = ["Criterion","Max Score","Description"]
    missing = [c for c in required if c not in rubric_df.columns]
    if missing:
        st.error(f"Rubric missing columns: {missing}")
        st.stop()

    case_text = read_file(case_file)
    rubric_text = rubric_to_text(rubric_df)

    results = []
    max_total = int(pd.to_numeric(rubric_df["Max Score"], errors="coerce").fillna(0).sum())

    for submission in participant_files:
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

        try:
            resp = client.chat.completions.create(
                model="gpt-4.1",
                temperature=0,
                response_format={"type":"json_object"},
                messages=[
                    {"role":"system","content":SYSTEM_PROMPT},
                    {"role":"user","content":prompt}
                ]
            )
            result = json.loads(resp.choices[0].message.content)
        except Exception as e:
            result = {"scores":{}, "strengths":[], "improvements":[str(e)]}

        row = {"Participant": submission.name}
        total = 0

        for _, r in rubric_df.iterrows():
            criterion = str(r["Criterion"]).strip()
            max_score = int(float(r["Max Score"]))

            raw_score = result.get("scores", {}).get(criterion, 0)
            try:
                score = float(str(raw_score).split("/")[0])
            except:
                score = 0

            score = max(0, min(score, max_score))
            row[f"{criterion} ({max_score})"] = score
            total += score

        row[f"Total ({max_total})"] = total
        row["Rating"] = rating(total)
        row["Strengths"] = "; ".join(result.get("strengths", []))
        row["Improvements"] = "; ".join(result.get("improvements", []))

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
