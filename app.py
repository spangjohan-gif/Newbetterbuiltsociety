
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import fitz  # PyMuPDF for PDF export
from pathlib import Path
import re

# -----------------------------------
# Konfiguration / filinläsning
# -----------------------------------
excel_file = "Better Built Society_v.0.1.xlsx"

@st.cache_data
def load_data(path):
    df_indata = pd.read_excel(path, sheet_name="Indata", engine="openpyxl")
    df_calc = pd.read_excel(path, sheet_name="Beräkningar", engine="openpyxl")
    return df_indata, df_calc

try:
    indata_df, calc_df = load_data(excel_file)
except FileNotFoundError:
    st.error(f"Hittar inte Excel-filen: {excel_file}. Kontrollera sökvägen/filnamnet.")
    st.stop()
except Exception as e:
    st.error(f"Kunde inte läsa Excel-data: {e}")
    st.stop()

required_indata_cols = {"Parameter", "Delparameter", "Råvärde", "Enhet", "Kommentar"}
if not required_indata_cols.issubset(set(indata_df.columns)):
    st.error(f"Fliken 'Indata' saknar någon av kolumnerna: {required_indata_cols}")
    st.stop()

required_calc_cols = {"Parameter", "Delparameter", "Råvärde", "Vikt(%)", "Max/referensvärde", "Optimalt värde", "Logik", "Normaliserat värde"}
if not required_calc_cols.issubset(set(calc_df.columns)):
    st.error(f"Fliken 'Beräkningar' saknar någon av kolumnerna: {required_calc_cols}")
    st.stop()

# Parametrar (i den ordning de förekommer i Indata)
parameters = indata_df["Parameter"].dropna().astype(str).unique().tolist()

# -----------------------------------
# Hjälpfunktioner
# -----------------------------------
def parse_scale(enhet: str):
    """Returnerar (min_val, max_val) om 'enhet' är av typen Skala X-Y eller Skala X-Y (med 0 som lägre gräns)."""
    if not isinstance(enhet, str):
        return None
    s = enhet.lower().strip()
    m = re.search(r"skala\s*([0-9]+)\s*[-–]\s*([0-9]+)", s)
    if m:
        lo = int(m.group(1))
        hi = int(m.group(2))
        return (lo, hi)
    # "Skala 0-10" varianten fångas av regex ovan. Om formatet skulle vara "Skala 1–5" med annorlunda tecken hanteras också.
    return None

def infer_input_type(enhet: str, parameter: str, delparam: str):
    """Avgör vilken kontroll som ska visas baserat på 'Enhet'."""
    if not isinstance(enhet, str):
        return {"type": "number", "min": 0, "step": 1}

    s = enhet.lower().strip()
    # Fritext
    if "fritext" in s or s == "text" or "kommentar" == s:
        return {"type": "text"}

    # Skala
    scale = parse_scale(enhet)
    if scale:
        lo, hi = scale
        # Välj slider för breda skala (t.ex. 0-10), radio för kort skala (1-3 / 1-5)
        if hi - lo <= 5:
            return {"type": "radio", "options": list(range(lo, hi + 1))}
        else:
            return {"type": "slider", "min": lo, "max": hi, "step": 1}

    # Kategori (special för Trafiktyp: 3=Spår, 2=BRT, 1=Buss)
    if "kategori" in s and parameter.lower() == "kollektivtrafik" and delparam.lower() == "trafiktyp":
        return {"type": "select", "options": [("Spår", 3), ("BRT", 2), ("Buss", 1)]}

    # Numeriska enheter
    if "procent" in s:
        return {"type": "number", "min": 0, "max": 100, "step": 1}
    if "meter" in s:
        return {"type": "number", "min": 0, "step": 1}
    if "minuter" in s:
        return {"type": "number", "min": 0, "step": 1}
    if "antal" in s:
        return {"type": "number", "min": 0, "step": 1}
    if "poäng" in s:
        return {"type": "number", "min": 0, "max": 100, "step": 1}
    if "dB".lower() in s.lower():
        return {"type": "number", "min": 0, "step": 1}
    if "km/h" in enhet or "kmh" in s:
        return {"type": "number", "min": 0, "step": 1}
    if "index" in s or "hushållstyper" in s or "boendeformer" in s:
        return {"type": "number", "min": 0, "step": 1}

    # Default: numeriskt
    return {"type": "number", "min": 0, "step": 1}

def normalize_value(raw, logic, max_ref, optimal):
    """Normaliserar råvärde till [0,1] enligt logik och referenser från 'Beräkningar'."""
    try:
        rv = float(raw)
    except Exception:
        return None  # kan ej normalisera t.ex. ren fritext

    if pd.isna(max_ref) and logic.lower() != "opt":
        return None

    if isinstance(max_ref, str):
        try:
            max_ref = float(max_ref)
        except Exception:
            return None

    logic = str(logic).lower().strip()
    if logic == "max":
        if max_ref == 0:
            return None
        val = rv / max_ref
    elif logic == "min":
        if max_ref == 0:
            return None
        val = 1 - (rv / max_ref)
    elif logic == "opt":
        if pd.isna(optimal):
            return None
        if isinstance(optimal, str):
            try:
                optimal = float(optimal)
            except Exception:
                return None
        denom = max_ref - optimal if max_ref is not None else None
        if denom is None or denom == 0:
            return None
        val = 1 - abs(rv - optimal) / denom
    else:
        return None

    # Klampa till [0,1]
    val = max(0.0, min(1.0, val))
    return val

def detect_weight_scale(weights_series: pd.Series) -> bool:
    """
    Returnerar True om vikter ser ut att vara fraktioner (0..1).
    Om någon vikt > 1 antas procenttal (0..100).
    """
    try:
        max_w = pd.to_numeric(weights_series, errors="coerce").max()
        return max_w <= 1.0
    except Exception:
        return True

def calculate_scores(answers: dict) -> dict:
    """
    Räknar poäng per parameter baserat på användarens råvärden.
    Hämtar 'Logik', 'Max/referensvärde', 'Optimalt värde', 'Vikt(%)' från 'Beräkningar'.
    Om svar saknas för en delparameter, används Råvärde från 'Beräkningar' eller dess redan
    normaliserade värde som fallback.
    """
    scores = {}
    weights_are_fractions = detect_weight_scale(calc_df["Vikt(%)"])

    for param, sub in calc_df.groupby("Parameter"):
        total_score = 0.0

        for _, row in sub.iterrows():
            delp = row["Delparameter"]
            logic = row["Logik"]
            max_ref = row["Max/referensvärde"]
            optimal = row["Optimalt värde"]
            weight = row["Vikt(%)"]
            if pd.isna(weight):
                weight = 0.0
            else:
                weight = float(weight)
                if not weights_are_fractions:
                    weight = weight / 100.0

            # Användarens svar (råvärde) om det finns
            key = (str(param), str(delp))
            if key in answers and answers[key].get("type") != "text":
                raw = answers[key].get("value")
                normalized = normalize_value(raw, logic, max_ref, optimal)
            else:
                # Fallback: använd normaliserat från kalkbladen
                normalized = row.get("Normaliserat värde", None)
                try:
                    normalized = float(normalized)
                except Exception:
                    # Försök normalisera med råvärdet i kalkbladen
                    raw = row.get("Råvärde", None)
                    normalized = normalize_value(raw, logic, max_ref, optimal)

            if normalized is None:
                normalized = 0.0

            total_score += weight * normalized

        scores[str(param)] = round(total_score, 3)

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
    lines.append("Dina svar:")
    # Gör lista per Parameter
    by_param = {}
    for (p, d), val in answers.items():
        by_param.setdefault(p, []).append((d, val))

    for p in parameters:
        if p not in by_param:
            continue
        lines.append(f"{p}:")
        for d, val in by_param[p]:
            vtype = val.get("type")
            enhet = val.get("enhet", "")
            if vtype == "text":
                txt = val.get("text", "").strip()
                lines.append(f"  - {d} ({enhet}): {txt if txt else '(tom)'}")
            else:
                v = val.get("value")
                lines.append(f"  - {d} ({enhet}): {v}")

    text = "\n".join(lines)
    page.insert_textbox(fitz.Rect(50, 50, 550, 800), text, fontsize=11, fontname="helv", align=0)

    pdf_path = "resultat.pdf"
    doc.save(pdf_path)
    doc.close()
    return pdf_path

# -----------------------------------
# Session state
# -----------------------------------
if "step" not in st.session_state:
    st.session_state.step = 0
if "answers" not in st.session_state:
    st.session_state.answers = {}

# -----------------------------------
# UI
# -----------------------------------
st.title("Better Built Society – Enkät enligt Indata (kolumn D)")

if st.session_state.step == 0:
    st.write("### Välkommen!")
    st.write("Du svarar per delparameter med rätt typ (fritext, skala eller numeriskt) baserat på 'Enhet' i Excel.")
    if st.button("Starta enkäten"):
        st.session_state.step = 1

elif 1 <= st.session_state.step <= len(parameters):
    current_param = parameters[st.session_state.step - 1]
    st.write(f"### {current_param} ({st.session_state.step} / {len(parameters)})")

    # Alla delparametrar för aktuell parameter från Indata
    rows = indata_df[indata_df["Parameter"] == current_param]

    for _, r in rows.iterrows():
        delparam = str(r["Delparameter"])
        enhet = str(r["Enhet"]) if not pd.isna(r["Enhet"]) else ""
        kommentar = str(r["Kommentar"]) if not pd.isna(r["Kommentar"]) else ""
        default_raw = r["Råvärde"]

        spec = infer_input_type(enhet, current_param, delparam)

        st.caption(f"**{delparam}** — {kommentar}" if kommentar else f"**{delparam}**")
        key = f"{current_param}__{delparam}"

        if spec["type"] == "text":
            val = st.text_area("Fritext", key=key+"_text", placeholder="Skriv din kommentar...", height=120)
            st.session_state.answers[(current_param, delparam)] = {
                "type": "text",
                "text": (val or "").strip(),
                "enhet": enhet
            }

        elif spec["type"] == "radio":
            options = spec["options"]
            # default: när options innehåller default_raw, annars mitten-värde
            default_idx = options.index(int(default_raw)) if default_raw in options else len(options)//2
            val = st.radio("Välj betyg", options, index=default_idx, horizontal=True, key=key+"_radio")
            st.session_state.answers[(current_param, delparam)] = {
                "type": "scale",
                "value": int(val),
                "enhet": enhet
            }

        elif spec["type"] == "slider":
            min_v, max_v, step_v = spec["min"], spec["max"], spec.get("step", 1)
            start = int(default_raw) if pd.notna(default_raw) else (min_v + max_v) // 2
            val = st.slider("Välj betyg", min_value=min_v, max_value=max_v, value=start, step=step_v, key=key+"_slider")
            st.session_state.answers[(current_param, delparam)] = {
                "type": "scale",
                "value": int(val),
                "enhet": enhet
            }

        elif spec["type"] == "select":
            options = spec["options"]  # list of (label, value)
            labels = [o[0] for o in options]
            values = [o[1] for o in options]
            try:
                default_index = values.index(int(default_raw)) if pd.notna(default_raw) else 0
            except Exception:
                default_index = 0
            chosen_label = st.selectbox("Välj kategori", labels, index=default_index, key=key+"_select")
            val = dict(options)[chosen_label]
            st.session_state.answers[(current_param, delparam)] = {
                "type": "number",
                "value": int(val),
                "enhet": enhet
            }

        else:  # number
            min_v = spec.get("min", 0)
            max_v = spec.get("max", None)
            step_v = spec.get("step", 1)
            kwargs = {"min_value": min_v, "step": step_v, "key": key+"_num"}
            if max_v is not None:
                kwargs["max_value"] = max_v
            # default_raw kan vara float eller int
            default_val = default_raw if pd.notna(default_raw) else min_v
            if isinstance(default_val, float) and step_v != 1:
                # behåll som float
                pass
            elif isinstance(default_val, float):
                default_val = int(default_val)
            val = st.number_input("Ange värde", value=default_val, **kwargs)
            st.session_state.answers[(current_param, delparam)] = {
                "type": "number",
                "value": float(val),
                "enhet": enhet
            }

        st.divider()

    # Navigering
    col_next, col_back = st.columns([1, 1])
    with col_next:
        if st.button("Nästa"):
            st.session_state.step += 1
    with col_back:
        if st.button("Tillbaka", type="secondary"):
            if st.session_state.step > 1:
                st.session_state.step -= 1

else:
    st.write("### Resultat")
    scores = calculate_scores(st.session_state.answers)
    total_score = round(sum(scores.values()) / len(scores), 3) if len(scores) > 0 else 0.0
    total_label, total_color = get_label(total_score)

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

    fig_bar = go.Figure(go.Bar(
        x=list(scores.keys()),
        y=list(scores.values()),
        marker_color="steelblue"
    ))
    fig_bar.update_layout(title="Resultat per parameter", xaxis_title="Parameter", yaxis_title="Poäng", yaxis_range=[0,1])
    fig_bar.update_xaxes(tickangle=30)
    st.plotly_chart(fig_bar, use_container_width=True)

    st.write("### Dina svar")
    for (p, d), val in st.session_state.answers.items():
        enhet = val.get("enhet", "")
        if val.get("type") == "text":
            st.markdown(f"**{p} / {d}** ({enhet}): {val.get('text','')}")
        else:
            st.markdown(f"**{p} / {d}** ({enhet}): {val.get('value')}")

    st.write("### Beräknade poäng")
    for param, score in scores.items():
        label, color = get_label(score)
        st.markdown(f"**{param}:** {score:.3f} - <span style='color:{color}'>{label}</span>", unsafe_allow_html=True)

    if st.button("Exportera som PDF"):
        pdf_path = export_pdf(scores, st.session_state.answers, total_score)
        if Path(pdf_path).exists():
            with open(pdf_path, "rb") as f:
                st.download_button("Ladda ner PDF", f, file_name="resultat.pdf")
        else:
            st.error("Misslyckades att skapa PDF. Försök igen.")

    if st.button("Starta om"):
        st.session_state.step = 0
        st.session_state.answers = {}
        st.rerun()
