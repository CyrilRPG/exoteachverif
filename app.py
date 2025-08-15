
import io
import json
import re
from typing import List, Tuple, Dict, Any, Optional

import pandas as pd
import streamlit as st

# -------------------------------
# Streamlit App Config
# -------------------------------
st.set_page_config(page_title="Vérif Groupes Étudiants", page_icon="✅", layout="wide")

# Minimal pleasant styling
st.markdown("""
<style>
:root { --radius: 14px; }
.block-container { padding-top: 1.2rem; }
.stButton>button, .stDownloadButton>button { border-radius: var(--radius); padding: 0.55rem 0.9rem; }
.kpi { border:1px solid #e5e7eb; border-radius: var(--radius); padding:0.8rem; background:#fafafa; }
small.dim { color:#6b7280; }
</style>
""", unsafe_allow_html=True)

# -------------------------------
# Liste officielle (numéro -> (nom filière, type))
# -------------------------------
OFFICIEL: Dict[int, Tuple[str, str]] = {
    5944: ("USPN", "Classe"),
    5943: ("USPN", "Classe"),
    5942: ("USPN", "Classe"),
    5018: ("USPN", "Filière"),
    5017: ("USPN", "Filière"),
    5016: ("USPN", "Filière"),
    5935: ("PASS UPC", "Classe"),
    5934: ("PASS UPC", "Classe"),
    5933: ("PASS UPC", "Classe"),
    5932: ("PASS UPC", "Classe"),
    5013: ("PASS UPC", "Filière"),
    5012: ("PASS UPC", "Filière"),
    5940: ("PASS SU", "Classe"),
    5939: ("PASS SU", "Classe"),
    5938: ("PASS SU", "Classe"),
    5937: ("PASS SU", "Classe"),
    5936: ("PASS SU", "Classe"),
    5014: ("PASS SU", "Filière"),
    5941: ("PASS UVSQ", "Classe"),
    5015: ("PASS UVSQ", "Filière"),
    5945: ("PASS UPS", "Classe"),
    5019: ("PASS UPS", "Filière"),
    5953: ("LSPS2 UPEC", "Classe"),
    5952: ("LSPS2 UPEC", "Classe"),
    5951: ("LSPS2 UPEC", "Classe"),
    5950: ("LSPS1 UPEC", "Classe"),
    5949: ("LSPS1 UPEC", "Classe"),
    5948: ("LSPS1 UPEC", "Classe"),
    5947: ("LSPS1 UPEC", "Classe"),
    5946: ("LAS1 UPEC", "Classe"),
    5032: ("LSPS3 UPEC", "Filière"),
    5022: ("LSPS2 UPEC", "Filière"),
    5021: ("LSPS1 UPEC", "Filière"),
    5020: ("LAS1 UPEC", "Filière"),
    6127: ("PAES Distanciel", "Classe"),
    6125: ("PAES Présentiel", "Classe"),
    6124: ("PAES Présentiel", "Classe"),
    6123: ("PAES Présentiel", "Classe"),
    6122: ("PAES Présentiel", "Classe"),
    5024: ("PAES Distanciel", "Filière"),
    5023: ("PAES Présentiel", "Filière"),
    6120: ("Terminale Santé Distanciel", "Classe"),
    6119: ("Terminale Santé Présentiel", "Classe"),
    6118: ("Terminale Santé Présentiel", "Classe"),
    6117: ("Terminale Santé Présentiel", "Classe"),
    6116: ("Terminale Santé Présentiel", "Classe"),
    6115: ("Terminale Santé Présentiel", "Classe"),
    6114: ("Terminale Santé Présentiel", "Classe"),
    6113: ("Terminale Santé Présentiel", "Classe"),
    6112: ("Terminale Santé Présentiel", "Classe"),
    5026: ("Terminale Santé Distanciel", "Filière"),
    5025: ("Terminale Santé Présentiel", "Filière"),
    6128: ("Première Élite", "Classe"),
    5027: ("Première Élite", "Filière"),
}

# Regex pour extraire tous les nombres (plus robuste que split sur espaces)
NUM_RE = re.compile(r"\\d+")

# -------------------------------
# Helpers Analyse
# -------------------------------
def parse_numeros(groupes_str: Any) -> List[int]:
    if pd.isna(groupes_str):
        return []
    return [int(m.group(0)) for m in NUM_RE.finditer(str(groupes_str))]

def analyser_groupes(groupes_str: Any) -> str:
    nums = parse_numeros(groupes_str)
    filieres = [n for n in nums if n in OFFICIEL and OFFICIEL[n][1] == "Filière"]
    classes  = [n for n in nums if n in OFFICIEL and OFFICIEL[n][1] == "Classe"]

    if len(filieres) == 0 and len(classes) == 0:
        return "Pas de classe ni de filière"
    if len(filieres) == 0 and len(classes) > 0:
        return "Pas de filière"
    if len(classes) == 0 and len(filieres) > 0:
        return "Pas de classe"
    if len(filieres) > 1 and len(classes) > 1:
        return "Plusieurs filières et plusieurs classes"
    if len(filieres) > 1:
        return "Plusieurs filières"
    if len(classes) > 1:
        return "Plusieurs classes"

    # 1 filière + 1 classe -> vérifier cohérence du label
    filiere_nom = OFFICIEL[filieres[0]][0]
    classe_nom  = OFFICIEL[classes[0]][0]
    if filiere_nom != classe_nom:
        return "Classe et filière incohérents"
    return "OK"

def extra_info(groupes_str: Any) -> Dict[str, Any]:
    nums = parse_numeros(groupes_str)
    connus = [n for n in nums if n in OFFICIEL]
    inconnus = [n for n in nums if n not in OFFICIEL]
    filieres = [n for n in nums if n in OFFICIEL and OFFICIEL[n][1]=="Filière"]
    classes  = [n for n in nums if n in OFFICIEL and OFFICIEL[n][1]=="Classe"]
    filiere_label = None
    classe_label = None
    if len(filieres) == 1:
        filiere_label = f"{OFFICIEL[filieres[0]][0]} ({filieres[0]})"
    if len(classes) == 1:
        classe_label = f"{OFFICIEL[classes[0]][0]} ({classes[0]})"
    return {
        "NumerosTrouvés": nums,
        "NumerosConnus": connus,
        "NumerosInconnus": inconnus,
        "FiliereDéduite": filiere_label,
        "ClasseDéduite": classe_label,
    }

# -------------------------------
# Barre latérale (paramètres)
# -------------------------------
with st.sidebar:
    st.header("⚙️ Options d'import")
    header_row_display = st.number_input("Ligne d'en-tête (1 = première ligne)", min_value=1, value=1, step=1)
    header_row = header_row_display - 1  # pandas header index
    use_sheet = st.text_input("Nom de l'onglet (laisser vide pour auto)", value="")
    show_debug = st.checkbox("Afficher colonnes techniques", value=False)
    st.markdown("---")
    st.caption("Si votre fichier a des en-têtes sur plusieurs lignes, changez la ligne d'en-tête.")

st.title("Vérification des groupes étudiants")
st.caption("Upload Excel → choix de l'onglet et de la colonne → diagnostic automatique → export JSON/CSV.")

uploaded = st.file_uploader("Dépose un fichier Excel (.xlsx, .xls)", type=["xlsx", "xls"])

if not uploaded:
    st.info("Charge un fichier pour commencer.")
    st.stop()

# Lecture des onglets si xlsx multi-sheets
xl = pd.ExcelFile(uploaded)
sheet_name: Optional[str] = None
if use_sheet and use_sheet in xl.sheet_names:
    sheet_name = use_sheet
else:
    # heuristique: prendre le premier onglet
    sheet_name = xl.sheet_names[0]

try:
    df_raw = pd.read_excel(uploaded, sheet_name=sheet_name, header=header_row)
except Exception as e:
    st.error(f"Erreur de lecture: {e}")
    st.stop()

st.write(f"**Onglet lu:** `{sheet_name}` — **ligne d'en-tête:** {header_row_display}")

# Proposer la sélection de la colonne "Groupes"
# Heuristique d'auto-détection: choisir la colonne avec le plus de nombres >= 4 chiffres
def score_col(s: pd.Series) -> int:
    try:
        return s.astype(str).str.count(r"\\d{4,}").fillna(0).astype(int).sum()
    except Exception:
        return 0

candidate_cols = sorted(df_raw.columns, key=lambda c: score_col(df_raw[c]), reverse=True)
default_col = candidate_cols[0] if candidate_cols else None
groupes_col = st.selectbox("Choisis la colonne contenant les numéros de groupes", options=list(df_raw.columns), index=(list(df_raw.columns).index(default_col) if default_col in df_raw.columns else 0))

# Analyse
df = df_raw.copy()
df["Diagnostic"] = df[groupes_col].apply(analyser_groupes)
extras = df[groupes_col].apply(extra_info).apply(pd.Series)
df = pd.concat([df, extras], axis=1)

# KPIs
c1, c2, c3, c4, c5 = st.columns(5)
total = len(df)
n_ok = (df["Diagnostic"] == "OK").sum()
n_incoh = (df["Diagnostic"] == "Classe et filière incohérents").sum()
n_pas_fil = (df["Diagnostic"] == "Pas de filière").sum()
n_pas_cls = (df["Diagnostic"] == "Pas de classe").sum()

with c1: st.markdown(f'<div class="kpi"><b>Total</b><br><span style="font-size:1.4rem">{total}</span></div>', unsafe_allow_html=True)
with c2: st.markdown(f'<div class="kpi"><b>OK</b><br><span style="font-size:1.4rem">{n_ok}</span></div>', unsafe_allow_html=True)
with c3: st.markdown(f'<div class="kpi"><b>Incohérents</b><br><span style="font-size:1.4rem">{n_incoh}</span></div>', unsafe_allow_html=True)
with c4: st.markdown(f'<div class="kpi"><b>Pas de filière</b><br><span style="font-size:1.4rem">{n_pas_fil}</span></div>', unsafe_allow_html=True)
with c5: st.markdown(f'<div class="kpi"><b>Pas de classe</b><br><span style="font-size:1.4rem">{n_pas_cls}</span></div>', unsafe_allow_html=True)

st.markdown("### Aperçu des données")
base_cols = [c for c in df.columns if c not in [groupes_col, "Diagnostic", "FiliereDéduite", "ClasseDéduite", "NumerosTrouvés", "NumerosConnus", "NumerosInconnus"]]
display_cols = base_cols + [groupes_col, "Diagnostic"]
if show_debug:
    display_cols += ["FiliereDéduite", "ClasseDéduite", "NumerosTrouvés", "NumerosConnus", "NumerosInconnus"]
st.dataframe(df[display_cols], use_container_width=True)

# Exports
records = df.to_dict(orient="records")
json_bytes = json.dumps(records, ensure_ascii=False, indent=2).encode("utf-8")
st.download_button("⬇️ Télécharger JSON (complet)", data=json_bytes, file_name="export_verifie.json", mime="application/json")

erreurs = df[df["Diagnostic"] != "OK"]
if not erreurs.empty:
    csv_err = erreurs[display_cols].to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Télécharger uniquement les erreurs (CSV)", data=csv_err, file_name="erreurs_groupes.csv", mime="text/csv")

st.markdown("""
#### Rappels de règles
- **OK** : 1 filière + 1 classe officielles, cohérentes entre elles.
- **Pas de filière** : classe officielle détectée mais aucune filière.
- **Pas de classe** : filière officielle détectée mais aucune classe.
- **Pas de classe ni de filière** : aucun numéro officiel détecté (les autres numéros sont ignorés).
- **Plusieurs filières / classes** : plus d'une filière ou plus d'une classe détectée.
- **Classe et filière incohérents** : la classe détectée n'appartient pas à la filière détectée.
""")
