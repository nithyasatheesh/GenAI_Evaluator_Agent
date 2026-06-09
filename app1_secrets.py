
import streamlit as st
import pandas as pd
import json
import io
from openai import OpenAI
from docx import Document
from PyPDF2 import PdfReader

st.set_page_config(page_title="GenAI Case Study Evaluator", layout="wide")
st.title("📊 GenAI Case Study Evaluator")

# ---------- Readers ----------
def read_pdf(file):
    try:
        reader = PdfReader(file)
        return "\n".join([(p.extract_text() or "") for p in reader.pages])
    except Exception as e:
        return f"PDF Read Error: {e}"

def read_docx(file):
    try:
        doc = Document(file)
        return "\n".join([p.text for p in doc.paragraphs])
    except Exception as e:
        return f"DOCX Read Error: {e}"

def read_file(file):
    if file.name.lower().endswith(".pdf"):
        return read_pdf(file)
    return read_docx(file)

def rubric_to_text(df):
    rows = []
    for _, r in df.iterrows():
        rows.append(
            f"Criterion: {r['Criterion']}\n"
            f"Max Score: {r['Max Score']}\n"
            f"Description: {r['Description']}"
        )
    return "\n\n".join(rows)

def get_rating(score):
    if score >= 90:
        return "Exceptional"
    elif score >= 80:
        return "Proficient"
    elif score >= 70:
        return "Competent"
    elif score >= 60:
        return "Developing"
    return "Needs Improvement"

SYSTEM_PROMPT = """
You are an expert evaluator for a GenAI business case study.

Evaluate ONLY against the rubric.
Do NOT compare against a model answer.
Different valid approaches may receive high scores.
Score only based on evidence present.

Return ONLY valid JSON:

{
  "scores": {},
  "strengths": [],
  "improvements": []
}
"""

# ---------- OpenAI ----------
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception:
    st.error("OPENAI_API_KEY not configured in Streamlit Secrets.")
    st.stop()

# ---------- UI ----------
case_file = st.file_uploader("Case Study Problem", type=["pdf", "docx"])
rubric_file = st.file_uploader("Rubric (Excel)", type=["xlsx"])

participant_files = st.file_uploader(
    "Participant Submissions",
    type=["pdf", "docx"],
    accept_multiple_files=True
)

custom_prompt = st.text_area(
    "Custom Evaluator Instructions",
    placeholder="Example: Be stricter on validation and governance scoring."
)

# ---------- Evaluate ----------
if st.button("Evaluate"):

    if not case_file:
        st.error("Please upload a case study.")
        st.stop()

    if not rubric_file:
        st.error("Please upload a rubric.")
        st.stop()

    if not participant_files:
        st.error("Please upload participant submissions.")
        st.stop()

    rubric_df = pd.read_excel(rubric_file)

    required_cols = ["Criterion", "Max Score", "Description"]

    missing = [c for c in required_cols if c not in rubric_df.columns]

    if missing:
        st.error(f"Rubric missing required columns: {missing}")
        st.stop()

    # Safe cleanup
    rubric_df["Max Score"] = pd.to_numeric(
        rubric_df["Max Score"],
        errors="coerce"
    )

    rubric_df = rubric_df.dropna(
        subset=["Criterion", "Max Score"]
    )

    rubric_df["Max Score"] = rubric_df["Max Score"].astype(int)

    rubric_df = rubric_df[
        rubric_df["Criterion"].astype(str).str.strip() != ""
    ]

    with st.expander("Rubric Preview"):
        st.dataframe(rubric_df)

    max_total = int(rubric_df["Max Score"].sum())

    case_text = read_file(case_file)
    rubric_text = rubric_to_text(rubric_df)

    results = []

    progress = st.progress(0)

    for idx, submission in enumerate(participant_files):

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
            response = client.chat.completions.create(
                model="gpt-4.1",
                temperature=0,
                response_format={"type": "json_object"},
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": prompt}
                ]
            )

            try:
                result = json.loads(
                    response.choices[0].message.content
                )
            except Exception as e:
                result = {
                    "scores": {},
                    "strengths": [],
                    "improvements": [f"JSON Parse Error: {e}"]
                }

        except Exception as e:
            result = {
                "scores": {},
                "strengths": [],
                "improvements": [str(e)]
            }

        row = {"Participant": submission.name}
        total = 0

        for _, r in rubric_df.iterrows():

            criterion = str(r["Criterion"]).strip()
            max_score = int(r["Max Score"])

            raw_score = result.get(
                "scores",
                {}
            ).get(
                criterion,
                0
            )

            try:
                cleaned = (
                    str(raw_score)
                    .replace("/100", "")
                    .replace("/20", "")
                    .replace("/15", "")
                    .replace("/10", "")
                    .replace("/5", "")
                    .strip()
                )
                score = float(cleaned)
            except:
                score = 0

            score = max(
                0,
                min(score, max_score)
            )

            row[f"{criterion} ({max_score})"] = score
            total += score

        row[f"Total ({max_total})"] = total
        row["Rating"] = get_rating(total)
        row["Strengths"] = "; ".join(result.get("strengths", []))
        row["Improvements"] = "; ".join(result.get("improvements", []))

        results.append(row)

        progress.progress((idx + 1) / len(participant_files))

    df = pd.DataFrame(results)

    st.success("Evaluation Complete")
    st.dataframe(df, use_container_width=True)

    excel = io.BytesIO()

    with pd.ExcelWriter(excel, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Results")

    st.download_button(
        "📥 Download Evaluation Report",
        excel.getvalue(),
        file_name="evaluation_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
