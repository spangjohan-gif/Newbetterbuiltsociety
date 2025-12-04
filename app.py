
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import fitz  # PyMuPDF for PDF export
from pathlib import Path

# -----------------------------------
# Konfiguration / filinläsning
# -----------------------------------
excel_file = "Better Built Society_v.0.1.xlsx"

@st.cache_data
def load_data(path):
    df_calc = pd.read_excel(path, sheet_name="Beräkningar", engine="openpyxl")
    df_indata = pd.read_excel(path, sheet_name="Indata", engine="openpyxl")
    return df_calc, df_indata

try:
    calc_df, indata_df = load_data(excel_file)
except FileNotFoundError:
    st.error(f"Hittar inte Excel-filen: {excel_file}. Kontrollera sökvägen/filnamnet.")
    st.stop()
except Exception as e:
    st.error(f"Kunde inte läsa Excel-data: {e}")
    st.stop()

if "Råvärde" in indata_df.columns:
    indata_df["Råvärde"] = None

if "Parameter" not in calc_df.columns:
    st.error("Kolumnen 'Parameter' saknas i fliken 'Beräkningar'.")
    st.stop()

parameters = (
    calc_df["Parameter"]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

# -----------------------------------
# NYTT: Läs frågetext från kolumn D (fjärde kolumnen) i 'Beräkningar'
# -----------------------------------
def get_free_text_prompts(df: pd.DataFrame) -> dict:
    """
    Hämtar frågetext/fritext-etikett per parameter från kolumn D.
    - Försöker först via kolumnnamn som kan vara relevanta (Fritext, Frågetext, Kommentar).
    - Faller annars tillbaka till fjärde kolumnen (index 3), dvs kolumn D.
    Returnerar dict: {parameter: text (str)}.
    """
    candidate_names = ["Fritext", "Fritextfråga", "Frågetext", "Kommentar"]
    col_name = None
    for name in candidate_names:
        if name in df.columns:
            col_name = name
            break

    if col_name is None:
        if len(df.columns) >= 4:
            col_name = df.columns[3]  # kolumn D
            st.info(f"Använder kolumn D ('{col_name}') som fritextfråga.")
        else:
            st.warning("Hittar ingen kolumn D (fjärde kolumn) i 'Beräkningar'—fritext inaktiverad.")
            return {}

    prompts = {}
    for param, sub in df.groupby("Parameter"):
        vals = sub[col_name].dropna()
        if not vals.empty:
            prompts[param] = str(vals.iloc[0])
        else:
            prompts[param] = ""  # tom etikett => använd generisk label senare
    return prompts

free_text_prompts = get_free_text_prompts(calc_df)

# -----------------------------------
# Session state
# -----------------------------------
if "step" not in st.session_state:
    st.session_state.step = 0
if "answers" not in st.session_state:
    st.session_state.answers = {}

# -----------------------------------
# Hjälpfunktioner
# -----------------------------------
def calculate_scores(answers: dict) -> dict:
    """
    Räknar ut poäng per parameter.
    Stödjer att answers[param] kan vara:
      - int (betyg 1–5), eller
      - dict {"rating": int, "text": str}
    Antaganden:
      - 'Vikt(%)' i procent (ex: 20 = 20%)
      - 'Normaliserat värde' i [0..1]
      - Användarens betyg (1–5) normaliseras till [0.2..1.0] och appliceras på alla delfaktorer för den parametern.
    """
    scores = {}
    if "Vikt(%)" not in calc_df.columns:
        st.error("Kolumnen 'Vikt(%)' saknas i fliken 'Beräkningar'.")
        return scores
    if "Normaliserat värde" not in calc_df.columns:
        st.error("Kolumnen 'Normaliserat värde' saknas i fliken 'Beräkningar'.")
        return scores

    for param in parameters:
        sub_df = calc_df[calc_df["Parameter"] == param]
        total_score = 0.0

        # Plocka ut ev. betyg från answers[param]
        rating_value = None
        if param in answers:
            if isinstance(answers[param], dict):
                rating_value = answers[param].get("rating", None)
            else:
                rating_value = answers[param]  # bakåtkompatibelt (int)

        for _, row in sub_df.iterrows():
            weight_pct = row.get("Vikt(%)", 0)
            try:
                weight = float(weight_pct) / 100.0
            except Exception:
                weight = 0.0

            if rating_value is not None:
                normalized = float(rating_value) / 5.0
            else:
                normalized = row.get("Normaliserat värde", 0.0)
                try:
                    normalized = float(normalized)
                except Exception:
                    normalized = 0.0

            total_score += normalized * weight

        scores[param] = round(total_score, 3)

    return scores

def get_label(score: float):
    if score >= 0.75:
        return "Bra", "green"
    elif score >= 0.5:
        return "Godtagbar", "yellow"
    elif score >= 0.25:
        return "Bristfällig", "orange"
    else:
        return "Dålig", "red"

def export_pdf(scores: dict, answers: dict, total_score: float) -> str:
    """
    Skapar PDF-rapport med totalpoäng, parameterpoäng och fritextsvar.
    """
    doc = fitz.open()
    page = doc.new_page()

    lines = []
    lines.append("Better Built Society - Resultat")
    lines.append("")
    lines.append(f"Totalpoäng: {round(total_score, 3)}")
    lines.append("")
    lines.append("Detaljer per parameter:")
    for param, score in scores.items():
        label, _ = get_label(score)
        lines.append(f"- {param}: {score} ({label})")

    lines.append("")
    lines.append("Dina svar (betyg + fritext):")
    if answers:
        for param in parameters:
            ans = answers.get(param)
            if ans is None:
                continue
            if isinstance(ans, dict):
                rating = ans.get("rating", "")
                text = ans.get("text", "")
            else:
                rating = ans
                text = ""
            # Ta med frågetexten från kolumn D som rubrik/kontext, om den finns
            prompt = free_text_prompts.get(param, "").strip()
            if prompt:
                lines.append(f"- {param} | Fråga: {prompt}")
            else:
                lines.append(f"- {param}")
            lines.append(f"    Betyg: {rating}")
            if text:
                lines.append(f"    Fritext: {text}")
    else:
        lines.append("- Inga svar registrerade.")

    text = "\n".join(lines)
    rect = fitz.Rect(50, 50, 550, 800)
    page.insert_textbox(rect, text, fontsize=11, fontname="helv", align=0)

    pdf_path = "resultat.pdf"
    doc.save(pdf_path)
    doc.close()
    return pdf_path

# -----------------------------------
# UI
# -----------------------------------
st.title("Better Built Society - Prototyp med beräkningar, fritext och export")

if len(parameters) > 0 and st.session_state.step > 0 and st.session_state.step <= len(parameters):
    st.progress(st.session_state.step / len(parameters))

if st.session_state.step == 0:
    st.write("### Välkommen!")
    st.write("Svara på frågor om ditt område. Varje fråga har både betyg (1–5) och ett valfritt fritextsvar.")
    if st.button("Starta enkäten"):
        st.session_state.step = 1

elif 1 <= st.session_state.step <= len(parameters):
    current_param = parameters[st.session_state.step - 1]
    st.write(f"### Fråga {st.session_state.step} av {len(parameters)}")
    st.write(f"Hur upplever du **{current_param}**?")

    # Betyg
    answer_rating = st.radio(
        "Välj ett betyg",
        [1, 2, 3, 4, 5],
        index=2,
        horizontal=True
    )

    # Fritext (etikett från kolumn D om tillgänglig)
    prompt = free_text_prompts.get(current_param, "").strip()
    if prompt:
        text_label = prompt
    else:
        text_label = "Skriv en kommentar (valfritt)"

    answer_text = st.text_area(
        text_label,
        key=f"text_{current_param}",
        placeholder="Skriv din kommentar här...",
        height=120
    )

    col_next, col_back = st.columns([1, 1])
    with col_next:
        if st.button("Nästa"):
            st.session_state.answers[current_param] = {
                "rating": int(answer_rating),
                "text": (answer_text or "").strip()
            }
            st.session_state.step += 1

    with col_back:
        if st.button("Tillbaka", type="secondary", help="Gå till föregående fråga"):
            if st.session_state.step > 1:
                st.session_state.step -= 1

else:
    st.write("### Resultat")
    scores = calculate_scores(st.session_state.answers)
    total_score = round(sum(scores.values()) / len(scores), 3) if len(scores) > 0 else 0.0
    total_label, total_color = get_label(total_score)

    # Gauge
    fig_gauge = go.Figure(go.Indicator(
        mode="gauge+number",
        value=total_score,
        title={'text': f"Totalpoäng ({total_label})"},
        number={'valueformat': '.3f'},
        gauge={
            'axis': {'range': [0, 1]},
            'bar': {'color': total_color},
            'steps': [
                {'range': [0.0, 0.25], 'color': '#ffcccc'},
                {'range': [0.25, 0.5], 'color': '#ffe0b3'},
                {'range': [0.5, 0.75], 'color': '#ffffb3'},
                {'range': [0.75, 1.0], 'color': '#ccffcc'},
            ]
        }
    ))
    st.plotly_chart(fig_gauge, use_container_width=True)

    # Staplar
    fig_bar = go.Figure()
    fig_bar.add_trace(go.Bar(
        x=list(scores.keys()),
        y=list(scores.values()),
        marker_color="steelblue"
    ))
    fig_bar.update_layout(
        title="Resultat per parameter",
        xaxis_title="Parameter",
        yaxis_title="Poäng (0–1)",
        yaxis_range=[0, 1],
        bargap=0.2,
    )
    fig_bar.update_xaxes(tickangle=30)
    st.plotly_chart(fig_bar, use_container_width=True)

    # Visa svaren (betyg + fritext)
    st.write("### Dina svar")
    for param in parameters:
        ans = st.session_state.answers.get(param)
        if ans is None:
            continue
        rating = ans["rating"] if isinstance(ans, dict) else ans
        text = ans["text"] if isinstance(ans, dict) else ""
        prompt = free_text_prompts.get(param, "").strip()

        st.markdown(f"**{param}**")
        if prompt:
            st.caption(f"Fråga (från kolumn D): {prompt}")
        st.write(f"- Betyg: {rating}")
        if text:
            st.write(f"- Fritext: {text}")

    # Beräknade poäng med etiketter
    st.write("### Beräknade poäng")
    for param, score in scores.items():
        label, color = get_label(score)
        st.markdown(
            f"**{param}:** {score:.3f} - <span style='color:{color}'>{label}</span>",
            unsafe_allow_html=True
        )

    # Exportera PDF
    if st.button("Exportera som PDF"):
        pdf_path = export_pdf(scores, st.session_state.answers, total_score)
        if Path(pdf_path).exists():
            with open(pdf_path, "rb") as f:
                st.download_button("Ladda ner PDF", f, file_name="resultat.pdf")
        else:
            st.error("Misslyckades att skapa PDF. Försök igen.")

    # Reset
    if st.button("Starta om"):
        st.session_state.step = 0
        st.session_state.answers = {}
        st.rerun()
