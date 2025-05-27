# app.py
import streamlit as st
import pandas as pd
import time
import re
from io import BytesIO
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# → Clé OpenAI depuis les Secrets Streamlit
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

st.set_page_config(page_title="AI Excel Processor", layout="wide")
st.title("🔧 AI Excel Processor")

# 1) Upload & cache les bytes, le méta et le workbook
uploaded = st.file_uploader("📂 Chargez votre fichier Excel", type=["xlsx"])
if not uploaded:
    st.stop()

# on redétecte un nouveau fichier si le nom ou la taille change
if (
    "bytes" not in st.session_state
    or st.session_state.filename != uploaded.name
    or st.session_state.filesize != uploaded.size
):
    st.session_state.bytes     = uploaded.read()
    st.session_state.filename  = uploaded.name
    st.session_state.filesize  = uploaded.size

    # on charge le workbook en mode écriture directement
    wb = load_workbook(
        filename=BytesIO(st.session_state.bytes),
        read_only=False,
        data_only=False
    )
    st.session_state.wb = wb
    st.session_state.sheet_names = wb.sheetnames

sheet_names = st.session_state.sheet_names

# 2) Sélection de la feuille (lazy loading)
selected_sheet = st.selectbox("🗂 Sélectionnez l'onglet", sheet_names)

@st.cache_data(show_spinner=False)
def load_sheet(excel_bytes: bytes, sheet: str) -> pd.DataFrame:
    return pd.read_excel(BytesIO(excel_bytes), engine="openpyxl", sheet_name=sheet)

df = load_sheet(st.session_state.bytes, selected_sheet).copy()

st.success(f"Onglet « {selected_sheet} » : {df.shape[0]} lignes × {df.shape[1]} colonnes")
st.dataframe(df.head(50), height=250)

# --- Prépare le prompt et les placeholders ---
if "prompt_text" not in st.session_state:
    st.session_state.prompt_text = ""

st.markdown("### ✏️ Rédigez votre prompt")
st.text_area(
    "Prompt (utilisez #Colonne# pour insérer un placeholder)",
    height=200,
    key="prompt_text"
)

st.markdown("### ➕ Sélectionnez vos placeholders")
st.multiselect(
    "Colonnes à insérer",
    options=list(df.columns),
    key="cols_to_insert"
)

def insert_placeholders_bulk():
    """Callback : ajoute les placeholders sélectionnés au prompt."""
    for col in st.session_state.cols_to_insert:
        ph = f"#{col}#"
        if ph not in st.session_state.prompt_text:
            st.session_state.prompt_text += ph + " "

st.button(
    "Ajouter tous les placeholders",
    on_click=insert_placeholders_bulk,
    key="btn_add_placeholders"
)

# Validation basique des placeholders (nouvelle syntaxe #Colonne#)
prompt_template = st.session_state.prompt_text
placeholders = re.findall(r"#([^#]+)#", prompt_template)
invalid = [c for c in placeholders if c not in df.columns]
if invalid:
    st.error(f"Colonnes invalides détectées : {', '.join(invalid)}")
    st.stop()
if not placeholders:
    st.warning("Aucun placeholder détecté pour l’instant.")

# 3) Prépare la colonne résultat
output_col = st.text_input("Nom de la colonne résultat", "Réponse IA")
if output_col not in df.columns:
    df[output_col] = ""

# 4) Config API & rate-limit
model       = st.selectbox("Modèle", ["gpt-4o-mini", "gpt-3.5-turbo"])
temperature = st.slider("Température", 0.0, 1.0, 0.0)
rate_limit  = st.number_input("Pause entre requêtes (s)", min_value=0.0, step=0.1, value=1.0)

# 5) Boutons Run / Stop
col1, col2 = st.columns(2)
do_run     = col1.button("▶️ Lancer")
do_stop    = col2.button("⏹️ Arrêter")
if "stop_flag" not in st.session_state:
    st.session_state.stop_flag = False
if do_stop:
    st.session_state.stop_flag = True

live_table   = st.empty()
progress_bar = st.progress(0)

# Cache local des prompts déjà exécutés
if "prompt_cache" not in st.session_state:
    st.session_state.prompt_cache = {}

def call_chat(prompt: str) -> str:
    if prompt in st.session_state.prompt_cache:
        return st.session_state.prompt_cache[prompt]
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=temperature,
            messages=[
                {"role": "system", "content": "Vous êtes un assistant utile et précis."},
                {"role": "user",   "content": prompt}
            ]
        )
        text = resp.choices[0].message.content.strip()
    except Exception as e:
        text = f"Erreur API : {e}"
    st.session_state.prompt_cache[prompt] = text
    return text

# 6) Boucle de traitement avec remplacement manuel des #placeholders#
if do_run:
    st.session_state.stop_flag = False
    total = len(df)
    try:
        for i, row in df.iterrows():
            if st.session_state.stop_flag:
                st.warning("⚠️ Traitement interrompu.")
                break

            if not row.get(output_col):
                data = {c: ("" if pd.isna(v) else str(v)) for c, v in row.items()}
                filled = prompt_template
                for col in placeholders:
                    filled = filled.replace(f"#{col}#", data.get(col, ""))
                df.at[i, output_col] = call_chat(filled)

            live_table.dataframe(df.head(50), height=250)
            progress_bar.progress(int((i + 1) / total * 100))
            time.sleep(rate_limit)
    except Exception as e:
        st.exception(e)

    st.success("✅ Traitement terminé.")
    live_table.dataframe(df.head(50), height=250)

# 7) Export en réutilisant st.session_state.wb pour ne pas écraser les autres onglets
if st.button("💾 Télécharger les résultats (tous onglets)"):
    buf = BytesIO()
    wb = st.session_state.wb  # même objet, avec vos modifications

    # Supprime l’ancienne version de la feuille traitée
    if selected_sheet in wb.sheetnames:
        old_ws = wb[selected_sheet]
        wb.remove(old_ws)

    # Recrée la feuille en première position
    new_ws = wb.create_sheet(selected_sheet, 0)
    for r in dataframe_to_rows(df, index=False, header=True):
        new_ws.append(r)

    wb.save(buf)
    buf.seek(0)
    st.download_button(
        "Télécharger Excel",
        data=buf,
        file_name="output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
