# app.py (v3) — robust I3/I4+ with auto data-start detection & clean errors CSV
import json
import re
from typing import List, Tuple, Dict, Any, Optional

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Vérif Groupes Étudiants — I3/I4+", page_icon="✅", layout="wide")
st.markdown("""
<style>
:root { --radius: 14px; }
.block-container { padding-top: 1rem; }
.stButton>button, .stDownloadButton>button { border-radius: var(--radius); padding:0.55rem 0.9rem; }
.kpi { border:1px solid #e5e7eb; border-radius: var(--radius); padding:0.8rem; background:#fafafa; }
small.dim { color:#6b7280; }
</style>
""", unsafe_allow_html=True)

st.title("Vérification des groupes étudiants — format I3/I4+")
st.caption("Repère I3 comme libellé, détecte automatiquement la première ligne de données sous I3, et garantit des exports propres.")

# ---------- Référentiel officiel ----------
OFFICIEL: Dict[int, Tuple[str, str]] = {
    5944: ("USPN", "Classe"), 5943: ("USPN", "Classe"), 5942: ("USPN", "Classe"),
    5018: ("USPN", "Filière"), 5017: ("USPN", "Filière"), 5016: ("USPN", "Filière"),
    5935: ("PASS UPC", "Classe"), 5934: ("PASS UPC", "Classe"), 5933: ("PASS UPC", "Classe"), 5932: ("PASS UPC", "Classe"),
    5013: ("PASS UPC", "Filière"), 5012: ("PASS UPC", "Filière"),
    5940: ("PASS SU", "Classe"), 5939: ("PASS SU", "Classe"), 5938: ("PASS SU", "Classe"), 5937: ("PASS SU", "Classe"), 5936: ("PASS SU", "Classe"),
    5014: ("PASS SU", "Filière"),
    5941: ("PASS UVSQ", "Classe"), 5015: ("PASS UVSQ", "Filière"),
    5945: ("PASS UPS", "Classe"), 5019: ("PASS UPS", "Filière"),
    5953: ("LSPS2 UPEC", "Classe"), 5952: ("LSPS2 UPEC", "Classe"), 5951: ("LSPS2 UPEC", "Classe"),
    5950: ("LSPS1 UPEC", "Classe"), 5949: ("LSPS1 UPEC", "Classe"), 5948: ("LSPS1 UPEC", "Classe"), 5947: ("LSPS1 UPEC", "Classe"),
    5946: ("LAS1 UPEC", "Classe"),
    5032: ("LSPS3 UPEC", "Filière"), 5022: ("LSPS2 UPEC", "Filière"), 5021: ("LSPS1 UPEC", "Filière"), 5020: ("LAS1 UPEC", "Filière"),
    6127: ("PAES Distanciel", "Classe"),
    6125: ("PAES Présentiel", "Classe"), 6124: ("PAES Présentiel", "Classe"), 6123: ("PAES Présentiel", "Classe"), 6122: ("PAES Présentiel", "Classe"),
    5024: ("PAES Distanciel", "Filière"), 5023: ("PAES Présentiel", "Filière"),
    6120: ("Terminale Santé Distanciel", "Classe"),
    6119: ("Terminale Santé Présentiel", "Classe"), 6118: ("Terminale Santé Présentiel", "Classe"), 6117: ("Terminale Santé Présentiel", "Classe"),
    6116: ("Terminale Santé Présentiel", "Classe"), 6115: ("Terminale Santé Présentiel", "Classe"), 6114: ("Terminale Santé Présentiel", "Classe"),
    6113: ("Terminale Santé Présentiel", "Classe"), 6112: ("Terminale Santé Présentiel", "Classe"),
    5026: ("Terminale Santé Distanciel", "Filière"), 5025: ("Terminale Santé Présentiel", "Filière"),
    6128: ("Première Élite", "Classe"), 5027: ("Première Élite", "Filière"),
}

NUM_RE = re.compile(r"\d+")

def parse_numeros(groupes_str: Any) -> List[int]:
    if pd.isna(groupes_str):
        return []
    s = str(groupes_str)
    return [int(m.group(0)) for m in NUM_RE.finditer(s)]

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
    filiere_label = f"{OFFICIEL[filieres[0]][0]} ({filieres[0]})" if len(filieres)==1 else None
    classe_label = f"{OFFICIEL[classes[0]][0]} ({classes[0]})" if len(classes)==1 else None
    return {"NumerosTrouvés": nums, "NumerosConnus": connus, "NumerosInconnus": inconnus,
            "FiliereDéduite": filiere_label, "ClasseDéduite": classe_label}

def excel_col_to_index(col_letter: str) -> int:
    col_letter = col_letter.strip().upper()
    total = 0
    for ch in col_letter:
        if not ('A' <= ch <= 'Z'):
            raise ValueError("Lettre de colonne invalide.")
        total = total * 26 + (ord(ch) - ord('A') + 1)
    return total - 1

def make_unique(cols: List[str]) -> List[str]:
    seen: Dict[str, int] = {}
    out: List[str] = []
    for c in cols:
        c = str(c)
        if c in seen:
            seen[c] += 1
            out.append(f"{c}.{seen[c]}")
        else:
            seen[c] = 0
            out.append(c)
    return out

def autodetect_name_columns(columns: List[str]) -> Tuple[Optional[str], Optional[str]]:
    lower_map = {c: str(c).strip().lower() for c in columns}
    nom_candidates = [c for c, l in lower_map.items() if any(k in l for k in ["nom", "last name"])]
    prenom_candidates = [c for c, l in lower_map.items() if any(k in l for k in ["prénom", "prenom", "first name"])]
    return (nom_candidates[0] if nom_candidates else None,
            prenom_candidates[0] if prenom_candidates else None)

def detect_data_start(raw: pd.DataFrame, groupes_col_idx: int, header_row_idx: int) -> int:
    """
    Cherche la première ligne non vide sous l'en-tête dans la colonne Groupes.
    Si rien n'est trouvé, retourne header_row_idx+1 (ancienne logique).
    """
    start_probe = header_row_idx + 1
    max_probe = min(len(raw), header_row_idx + 50)  # on scanne 50 lignes max
    for r in range(start_probe, max_probe):
        val = raw.iat[r, groupes_col_idx] if groupes_col_idx < raw.shape[1] else None
        if pd.notna(val) and str(val).strip() != "":
            return r
    return start_probe

# ---------- Sidebar ----------
with st.sidebar:
    st.header("⚙️ Import")
    use_sheet = st.text_input("Nom de l'onglet (laisser vide pour auto)", value="")
    col_letter_override = st.text_input("Colonne Groupes (défaut I)", value="I")
    start_row_manual = st.number_input("Forcer ligne de départ (0 = auto)", min_value=0, value=0, step=1)
    show_debug = st.checkbox("Afficher colonnes techniques", value=False)
    st.markdown("---")
    st.header("🧭 Colonnes Nom/Prénom")
    export_semicolon = st.checkbox("CSV erreurs avec point-virgule (;)", value=True)
    st.caption("Encodage UTF-8-SIG pour Excel FR.")

uploaded = st.file_uploader("Dépose un fichier Excel (.xlsx, .xls)", type=["xlsx", "xls"])
if not uploaded:
    st.info("Charge un fichier pour commencer.")
    st.stop()

xl = pd.ExcelFile(uploaded)
sheet_name = use_sheet if (use_sheet and use_sheet in xl.sheet_names) else xl.sheet_names[0]

try:
    raw = pd.read_excel(uploaded, sheet_name=sheet_name, header=None)
except Exception as e:
    st.error(f"Erreur de lecture: {e}")
    st.stop()

st.write(f"**Onglet lu:** `{sheet_name}`")

# I3 / headers
header_row_idx = 2
try:
    groupes_col_idx = excel_col_to_index(col_letter_override or "I")
except Exception:
    groupes_col_idx = 8

if header_row_idx >= len(raw):
    st.error("La ligne d'en-tête (3) n'existe pas dans ce fichier.")
    st.stop()
if groupes_col_idx >= raw.shape[1]:
    st.error("La colonne Groupes dépasse le nombre de colonnes du fichier.")
    st.stop()

# Détermine la première ligne de données (auto ou forcée)
auto_start_row_idx = detect_data_start(raw, groupes_col_idx, header_row_idx)
start_row_idx = int(start_row_manual) - 1 if start_row_manual > 0 else auto_start_row_idx

headers = make_unique(list(raw.iloc[header_row_idx].astype(str)))
data = raw.iloc[start_row_idx:, :].reset_index(drop=True)
if data.shape[1] > len(headers):
    headers = headers + [f"COL_{i}" for i in range(data.shape[1] - len(headers))]
else:
    headers = headers[: data.shape[1]]
data.columns = headers

GROUPES_COL_NAME = "Groupes (détecté I3/auto)"
data[GROUPES_COL_NAME] = raw.iloc[start_row_idx:, groupes_col_idx].reset_index(drop=True)

# Sanity check: la colonne Groupes paraît-elle vide ?
digits4 = data[GROUPES_COL_NAME].astype(str).str.count(r"\d{4,}").sum()
if digits4 == 0:
    st.warning("⚠️ La colonne **Groupes** semble vide ou mal alignée. "
               "Vérifie l’onglet et/ou force la ligne de départ dans la barre latérale.")
    st.write("Aperçu des 10 premières valeurs de la colonne Groupes :")
    st.write(data[GROUPES_COL_NAME].head(10))

# Nom / Prénom
nom_guess, prenom_guess = autodetect_name_columns(list(data.columns))
col1, col2 = st.columns(2)
with col1:
    nom_col = st.selectbox("Colonne Nom", options=["—"] + list(data.columns),
                           index=(["—"] + list(data.columns)).index(nom_guess) if nom_guess in (["—"] + list(data.columns)) else 0)
with col2:
    prenom_col = st.selectbox("Colonne Prénom", options=["—"] + list(data.columns),
                              index=(["—"] + list(data.columns)).index(prenom_guess) if prenom_guess in (["—"] + list(data.columns)) else 0)
nom_col = None if nom_col == "—" else nom_col
prenom_col = None if prenom_col == "—" else prenom_col
if not nom_col or not prenom_col:
    st.warning("⚠️ Choisis/valide les colonnes **Nom** et **Prénom** pour un export d'erreurs correct.")

# Analyse
df = data.copy()
df["Diagnostic"] = df[GROUPES_COL_NAME].apply(analyser_groupes)
extras = df[GROUPES_COL_NAME].apply(extra_info).apply(pd.Series)
df = pd.concat([df, extras], axis=1)

# Répartition complète
counts = df["Diagnostic"].value_counts().sort_index()
total = int(len(df))

c0, = st.columns(1)
with c0:
    st.markdown(f'<div class="kpi"><b>Total</b><br><span style="font-size:1.4rem">{total}</span></div>', unsafe_allow_html=True)

st.markdown("#### Répartition par diagnostic")
rep_df = counts.reset_index()
rep_df.columns = ["Diagnostic", "Effectif"]
rep_df.loc[len(rep_df)] = ["Total", total]
st.dataframe(rep_df, use_container_width=True)

# Tableau principal
base_cols = [c for c in df.columns if c not in [GROUPES_COL_NAME, "Diagnostic", "FiliereDéduite", "ClasseDéduite", "NumerosTrouvés", "NumerosConnus", "NumerosInconnus"]]
display_cols = base_cols + [GROUPES_COL_NAME, "Diagnostic"]
if st.checkbox("Afficher colonnes techniques", value=show_debug):
    display_cols += ["FiliereDéduite", "ClasseDéduite", "NumerosTrouvés", "NumerosConnus", "NumerosInconnus"]
st.markdown("### Aperçu des données")
st.dataframe(df[display_cols], use_container_width=True)

# Export JSON
records = df.to_dict(orient="records")
json_bytes = json.dumps(records, ensure_ascii=False, indent=2).encode("utf-8")
st.download_button("⬇️ Télécharger JSON (complet)", data=json_bytes, file_name="export_verifie.json", mime="application/json")

# Export erreurs (Nom, Prénom, Diagnostic)
erreurs = df[df["Diagnostic"] != "OK"].copy()

def safe_col(s: pd.Series) -> pd.Series:
    return s.astype(str).fillna("").replace({"nan": ""})

if erreurs.empty:
    st.info("Aucune erreur à exporter 🎉")
else:
    nom_for_export = nom_col if nom_col in erreurs.columns else None
    prenom_for_export = prenom_col if prenom_col in erreurs.columns else None
    if not nom_for_export:
        erreurs["__NOM__"] = ""
        nom_for_export = "__NOM__"
        st.warning("La colonne Nom sélectionnée n’existe pas — exportera une colonne vide.")
    if not prenom_for_export:
        erreurs["__PRENOM__"] = ""
        prenom_for_export = "__PRENOM__"
        st.warning("La colonne Prénom sélectionnée n’existe pas — exportera une colonne vide.")
    export_df = pd.DataFrame({
        "Nom": safe_col(erreurs[nom_for_export]),
        "Prénom": safe_col(erreurs[prenom_for_export]),
        "Diagnostic": safe_col(erreurs["Diagnostic"]),
    })
    sep = ";" if export_semicolon else ","
    csv_text = export_df.to_csv(index=False, sep=sep)
    csv_bytes = csv_text.encode("utf-8-sig")
    st.download_button("⬇️ Télécharger uniquement les erreurs (CSV) — 3 colonnes", data=csv_bytes,
                       file_name="erreurs_groupes.csv", mime="text/csv")

st.markdown("""
**Rappels :**  
- Auto-détection de la première ligne de données sous I3 (tu peux aussi forcer une ligne de départ dans la sidebar).  
- Si la colonne Groupes paraît vide, l’app te le signale avec un aperçu pour corriger facilement.  
- Export erreurs = **Nom, Prénom, Diagnostic** (UTF-8-SIG, `;` activable).
""")
