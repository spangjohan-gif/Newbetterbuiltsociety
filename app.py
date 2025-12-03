
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import fitz  # PyMuPDF for PDF export

# Load Excel data
excel_file = "Better Built Society_v.0.1.xlsx"
calc_df = pd.read_excel(excel_file, sheet_name="Beräkningar", engine="openpyxl")
indata_df = pd.read_excel(excel_file, sheet_name="Indata", engine="openpyxl")
indata_df["Råvärde"] = None  # Clear raw values

# Prepare parameter list
parameters = calc_df["Parameter"].dropna().unique().tolist()

# Initialize session state
if "step" not in st.session_state:
    st.session_state.step = 0
if "answers" not in st.session_state:
    st.session_state.answers = {}

# Function to calculate normalized score based on user answers
def calculate_scores(answers):
    scores = {}
    for param in parameters:
        sub_df = calc_df[calc_df["Parameter"] == param]
        total_score = 0
        for _, row in sub_df.iterrows():
            weight = row["Vikt(%)"]
            if param in answers:
                normalized = answers[param] / 5  # scale 1-5 to 0.2-1.0
            else:
                normalized = row["Normaliserat värde"]
            total_score += normalized * weight
        scores[param] = round(total_score, 3)
    return scores

# Function to assign color-coded label
def get_label(score):
    if score >= 0.75:
        return "Bra", "green"
    elif score >= 0.5:
        return "Godtagbar", "yellow"
    elif score >= 0.25:
        return "Bristfällig", "orange"
    else:
        return "Dålig", "red"

# Function to export PDF
def export_pdf(scores, answers, total_score):
    doc = fitz.open()
    page = doc.new_page()

    text = f"""Better Built Society - Resultat
# for line breaks
Totalpoäng: {round(total_score, 3)}

Detaljer per parameter:
"""
    for param, score in scores.items():
        label, _ = get_label(score)
        text += f"""{param}: {score} ({label})
Dina svar:

    for param, ans in answers.items():
        {param}: {ans}
"""

    page.insert_text((50, 50), text)
    pdf_path = "resultat.pdf"
    doc.save(pdf_path)
    doc.close()
    return pdf_path

# Streamlit UI
st.title("Better Built Society - Prototyp med beräkningar och export")

if st.session_state.step == 0:
    st.write("### Välkommen!")
    st.write("Svara på frågor om ditt område. När du är klar får du ett betyg och en visuell sammanställning.")
    if st.button("Starta enkäten"):
        st.session_state.step = 1

elif 1 <= st.session_state.step <= len(parameters):
    current_param = parameters[st.session_state.step - 1]
    st.write(f"### Fråga {st.session_state.step} av {len(parameters)}")
    st.write(f"Hur upplever du {current_param}?")
    answer = st.radio("Välj ett betyg", [1, 2, 3, 4, 5], index=2)
    if st.button("Nästa"):
        st.session_state.answers[current_param] = answer
        st.session_state.step += 1

else:
    st.write("### Resultat")
    scores = calculate_scores(st.session_state.answers)
    total_score = sum(scores.values()) / len(scores)

    fig_gauge = go.Figure(go.Indicator(
        mode="gauge+number",
        value=total_score,
        title={'text': "Totalpoäng"},
        gauge={'axis': {'range': [0, 1]}, 'bar': {'color': "green"}}
    ))
    st.plotly_chart(fig_gauge)

    fig_bar = go.Figure(go.Bar(
        x=list(scores.keys()),
        y=list(scores.values()),
        marker_color="blue"
    ))
    fig_bar.update_layout(title="Resultat per parameter", xaxis_title="Parameter", yaxis_title="Poäng", yaxis_range=[0,1])
    st.plotly_chart(fig_bar)

    st.write("### Dina svar")
    st.write(st.session_state.answers)

    st.write("### Beräknade poäng")
    for param, score in scores.items():
        label, color = get_label(score)
        st.markdown(f"**{param}:** {score} - <span style='color:{color}'>{label}</span>", unsafe_allow_html=True)

    if st.button("Exportera som PDF"):
        pdf_path = export_pdf(scores, st.session_state.answers, total_score)
        with open(pdf_path, "rb") as f:
            st.download_button("Ladda ner PDF", f, file_name="resultat.pdf")
