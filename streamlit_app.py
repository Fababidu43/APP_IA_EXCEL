# app.py
import streamlit as st
import pandas as pd
import time
import re
from io import BytesIO
from openai import OpenAI

# → Récupère la clé depuis les Secrets Streamlit (jamais committée)
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

st.set_page_config(page_title="AI Excel Processor", layout="wide")
st.title("🔧 AI Excel Processor")

# 1) Upload & load all sheets
uploaded = st.file_uploader("📂 Chargez votre fichier Excel", type=["xlsx"])
if not uploaded:
    st.stop()
all_sheets = pd.read_excel(uploaded, engine="openpyxl", sheet_name=None)
sheet_names = list(all_sheets.keys())

# 2) Choix de l'onglet
selected_sheet = st.selectbox("🗂 Sélectionnez l'onglet", sheet_names, key="select_sheet")
df = all_sheets[selected_sheet].copy()
st.success(f"Onglet « {selected_sheet} » : {df.shape[0]} lignes × {df.shape[1]} colonnes")
st.dataframe(df.head(100), height=300)  # affiche seulement les 100 premières lignes pour la perf

# --- Prépare session_state pour le prompt ---
if "prompt_text" not in st.session_state:
    st.session_state.prompt_text = ""

# 3) Zone de saisie du prompt
st.markdown("### ✏️ Rédigez votre prompt")
st.text_area(
    "Prompt (utilisez {Colonne} pour insérer un placeholder)",
    height=200,
    key="prompt_text"
)

# 4) Insertion de placeholders
st.markdown("### ➕ Ajouter un placeholder")

# 4A) Sélection simple d'une colonne
col_to_insert = st.selectbox("Sélectionnez la colonne :", df.columns, key="select_placeholder")
def insert_placeholder():
    placeholder = f"{{{col_to_insert}}}"
    if placeholder not in st.session_state.prompt_text:
        st.session_state.prompt_text += placeholder + " "
st.button("Ajouter `{Colonne}`", on_click=insert_placeholder)

# 4B) Sélection multiple de colonnes
cols_to_insert = st.multiselect(
    "Sélectionnez plusieurs colonnes à ajouter d’un coup",
    options=df.columns
)
def insert_placeholders_bulk():
    for col in cols_to_insert:
        placeholder = f"{{{col}}}"
        if placeholder not in st.session_state.prompt_text:
            st.session_state.prompt_text += placeholder + " "
st.button("Ajouter tous les placeholders", on_click=insert_placeholders_bulk)

# 5) Récupère et valide le prompt
prompt_template = st.session_state.prompt_text
placeholders = re.findall(r"\{([^}]+)\}", prompt_template)
if not placeholders:
    st.warning("Aucun placeholder détecté pour le moment.")
invalid = [c for c in placeholders if c not in df.columns]
if invalid:
    st.error(f"Colonnes invalides détectées : {', '.join(invalid)}")
    st.stop()

# 6) Prépare la colonne résultat
output_col = st.text_input("Nom de la colonne résultat", "Réponse IA")
if output_col not in df.columns:
    df[output_col] = ""

# 7) Configuration de l’API
model       = st.selectbox("Modèle", ["gpt-4o-mini", "gpt-3.5-turbo"])
temperature = st.slider("Température", 0.0, 1.0, 0.0)
rate_limit  = st.number_input("Pause entre requêtes (s)", min_value=0.0, step=0.1, value=1.0)

# 8) Lancer / Arrêter
col1, col2 = st.columns(2)
do_run     = col1.button("▶️ Lancer")
do_stop    = col2.button("⏹️ Arrêter")
if "stop_flag" not in st.session_state:
    st.session_state.stop_flag = False
if do_stop:
    st.session_state.stop_flag = True

def call_chat(prompt: str) -> str:
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=temperature,
            messages=[
                {"role": "system", "content": "Vous êtes un assistant utile et précis."},
                {"role": "user",   "content": prompt}
            ]
        )
        return resp.choices[0].message.content.strip()
    except Exception as e:
        return f"Erreur API : {e}"

# 9) Boucle de traitement avec progress bar (pas de live table)
if do_run:
    st.session_state.stop_flag = False
    progress_bar = st.progress(0)
    total = len(df)
    for i, row in df.iterrows():
        if st.session_state.stop_flag:
            st.warning("⚠️ Traitement interrompu.")
            break

        if not row.get(output_col):
            data = {c: ("" if pd.isna(v) else str(v)) for c, v in row.items()}
            try:
                filled = prompt_template.format(**data)
            except KeyError as e:
                df.at[i, output_col] = f"Placeholder manquant : {e}"
                continue

            df.at[i, output_col] = call_chat(filled)

        progress_bar.progress(int((i + 1) / total * 100))
        time.sleep(rate_limit)

    st.success("✅ Traitement terminé.")
    st.dataframe(df.head(100), height=300)

# 10) Prépare l’export de tous les onglets
all_sheets[selected_sheet] = df
buf = BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    for name, sheet_df in all_sheets.items():
        sheet_df.to_excel(writer, sheet_name=name, index=False)
buf.seek(0)

st.download_button(
    "💾 Télécharger les résultats (tous onglets)",
    data=buf,
    file_name="output.xlsx",
    mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet"
)
