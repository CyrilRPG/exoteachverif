# app.py — Vérification I3/I4+ + PDF (1 page = 1 classe, jusqu'à 42 élèves)
import json
import re
from typing import List, Tuple, Dict, Any, Optional, Set, DefaultDict
from collections import defaultdict
import io

import pandas as pd
import streamlit as st

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether

# ------------------------------- UI / THEME -------------------------------
st.set_page_config(page_title="Vérif Groupes Étudiants — I3/I4+ & PDF", page_icon="✅", layout="wide")
st.markdown("""
<style>
:root { --radius: 14px; }
.block-container { padding-top: 1rem; }
.stButton>button, .stDownloadButton>button { border-radius: var(--radius); padding:0.55rem 0.9rem; }
.kpi { border:1px solid #e5e7eb; border-radius: var(--radius); padding:0.8rem; background:#fafafa; }
small.dim { color:#6b7280; }
</style>
""", unsafe_allow_html=True)

st.title("Vérification des groupes étudiants — format I3/I4+ & Générateur PDF")

# ==================== RÉFÉRENTIEL (FILIERES ↔ CLASSES) ====================
FILIERE_NAMES: Dict[int, str] = {
    # USPN
    5016: "LAS - USPN 25/26",
    5017: "PASS - USPN 25/26",
    5018: "LSPS - USPN 25/26",
    # UPC
    5012: "PASS - UPC 25/26",
    5013: "LAS - UPC 25/26",
    # SU
    5014: "PASS - SU (TC) 25/26",
    # UVSQ
    5015: "PASS - UVSQ 25/26",
    # UPS
    5019: "PASS - UPS 25/26",
    # UPEC
    5020: "LAS1 Majeure disciplinaire - UPEC 25/26",
    5021: "LSPS1 - UPEC 25-26",
    5022: "LSPS2 - UPEC 25-26",
    5032: "LSPS3 - UPEC - 25-26",
    # PAES
    5023: "PAES - Présentiel 25-26",
    5024: "PAES - Distanciel 25-26",
    # Terminale Santé
    5025: "Terminale Santé 25-26 - Présentiel",
    5026: "Terminale Santé 25-26 - Distanciel",
    # Première Élite
    5027: "Première Élite 25-26",
}

CLASS_NAMES: Dict[int, str] = {
    # USPN (classes) — LAS 1 corrigé = 5944
    5944: "USPN - Classe 1 (LAS) 25/26",
    5943: "USPN - Classe 2 (PASS/LSPS) 25/26",
    5942: "USPN - Classe 1 (PASS/LSPS) 25/26",
    # UPC (PASS)
    5935: "PASS UPC - Classe 4 25/26",
    5934: "PASS UPC - Classe 3 25/26",
    5933: "PASS UPC - Classe 2 25/26",
    5932: "PASS UPC - Classe 1 25/26",
    # UPC (LAS)
    5931: "LAS UPC - Classe 1 25/26",
    # SU
    5940: "PASS SU - Classe 5 (Mineure Sciences) 25/26",
    5939: "PASS SU - Classe 4 (Mineure Lettres) 25/26",
    5938: "PASS SU - Classe 3 (Mineure Sciences) 25/26",
    5937: "PASS SU - Classe 2 (Mineure Sciences) 25/26",
    5936: "PASS SU - Classe 1 (Mineure Sciences) 25/26",
    # UVSQ
    5941: "PASS UVSQ - Classe 1 25/26",
    # UPS
    5945: "PASS UPS - Classe 1 25/26",
    # UPEC
    5953: "LSPS2 UPEC - Classe 3 (25-26)",
    5952: "LSPS2 UPEC - Classe 2 (25-26)",
    5951: "LSPS2 UPEC - Classe 1 (25-26)",
    5950: "LSPS1 UPEC - Classe 4 25/26",
    5949: "LSPS1 UPEC - Classe 3 25/26",
    5948: "LSPS1 UPEC - Classe 2 25/26",
    5947: "LSPS1 UPEC - Classe 1 25-26",
    5946: "LAS1 Majeure disciplinaire - UPEC - Classe 1 25/26",
    # PAES
    6127: "PAES Distanciel - Classe 1 25/26",
    6125: "PAES Présentiel - Classe 4 25/26",
    6124: "PAES Présentiel - Classe 2 25/26",
    6123: "PAES Présentiel - Classe 3 25/26",
    6122: "PAES Présentiel - Classe 1 25/26",
    # Terminale Santé
    6120: "Terminale Santé Distanciel - Classe 1 25/26",
    6119: "Terminale Santé Présentiel - Classe 8 25/26",
    6118: "Terminale Santé Présentiel - Classe 7 25/26",
    6117: "Terminale Santé Présentiel - Classe 6 25/26",
    6116: "Terminale Santé Présentiel - Classe 5 25/26",
    6115: "Terminale Santé Présentiel - Classe 4 25/26",
    6114: "Terminale Santé Présentiel - Classe 3 25/26",
    6113: "Terminale Santé Présentiel - Classe 2 25/26",
    6112: "Terminale Santé Présentiel - Classe 1 25/26",
    # Première Élite
    6128: "Première Elite - Classe 1 25/26",
}

# FILIERE -> CLASSES autorisées
FILIERE_TO_CLASSES: Dict[int, Set[int]] = {
    # USPN
    5016: {5944},           # LAS - USPN -> Classe 1 (LAS)
    5017: {5942, 5943},     # PASS - USPN -> Classes PASS/LSPS
    5018: {5942, 5943},     # LSPS - USPN -> Classes PASS/LSPS
    # UPC
    5012: {5932, 5933, 5934, 5935},  # PASS - UPC
    5013: {5931},                    # LAS - UPC
    # SU
    5014: {5936, 5937, 5938, 5939, 5940},
    # UVSQ
    5015: {5941},
    # UPS
    5019: {5945},
    # UPEC
    5020: {5946},
    5021: {5947, 5948, 5949, 5950},
    5022: {5951, 5952, 5953},
    5032: set(),
    # PAES
    5023: {6122, 6123, 6124, 6125},
    5024: {6127},
    # Terminale Santé
    5025: {6112, 6113, 6114, 6115, 6116, 6117, 6118, 6119},
    5026: {6120},
    # Première Élite
    5027: {6128},
}

# CLASSE -> FILIERES (inverse)
CLASSES_TO_FILIERES: Dict[int, Set[int]] = defaultdict(set)
for fcode, cls_set in FILIERE_TO_CLASSES.items():
    for c in cls_set:
        CLASSES_TO_FILIERES[c].add(fcode)

# OFFICIEL (tous codes)
OFFICIEL: Dict[int, Tuple[str, str]] = {}
for f_code, f_name in FILIERE_NAMES.items():
    OFFICIEL[f_code] = (f_name, "Filière")
for c_code, c_name in CLASS_NAMES.items():
    OFFICIEL[c_code] = (c_name, "Classe")

# ---------- Exceptions : classe sans filière => OK ----------
EXCEPTION_OK_IF_CLASS_ONLY: Set[int] = {
    4538, 4537, 4388, 4386, 4385, 4384, 4383, 4382, 4381, 4380, 4379, 4378, 4377, 4376, 4375
}

NUM_RE = re.compile(r"\d+")

def parse_numeros(groupes_str: Any) -> List[int]:
    if pd.isna(groupes_str):
        return []
    return [int(m.group(0)) for m in NUM_RE.finditer(str(groupes_str))]

# ============================= ANALYSE =============================
def analyser_groupes(groupes_str: Any) -> str:
    nums = parse_numeros(groupes_str)
    has_exception = any(n in EXCEPTION_OK_IF_CLASS_ONLY for n in nums)

    filieres = [n for n in nums if n in OFFICIEL and OFFICIEL[n][1] == "Filière"]
    classes  = [n for n in nums if n in OFFICIEL and OFFICIEL[n][1] == "Classe"]

    if len(filieres) == 0 and len(classes) == 0:
        return "Pas de classe ni de filière"

    if len(filieres) == 0 and len(classes) > 0:
        if has_exception:
            return "OK"  # classe seule mais exception autorisée
        return "Pas de filière"

    if len(classes) == 0 and len(filieres) > 0:
        return "Pas de classe"

    if len(filieres) > 1 and len(classes) > 1:
        return "Plusieurs filières et plusieurs classes"
    if len(filieres) > 1:
        return "Plusieurs filières"
    if len(classes) > 1:
        return "Plusieurs classes"

    # 1 filière et 1 classe -> vérifie appartenance par mapping
    f = filieres[0]
    c = classes[0]
    if c in CLASSES_TO_FILIERES and f in CLASSES_TO_FILIERES[c]:
        return "OK"
    else:
        return "Classe et filière incohérents"

def extra_info(groupes_str: Any) -> Dict[str, Any]:
    nums = parse_numeros(groupes_str)
    connus = [n for n in nums if n in OFFICIEL]
    inconnus = [n for n in nums if n not in OFFICIEL]
    filieres = [n for n in nums if n in OFFICIEL and OFFICIEL[n][1]=="Filière"]
    classes  = [n for n in nums if n in OFFICIEL and OFFICIEL[n][1]=="Classe"]
    filiere_label = FILIERE_NAMES[filieres[0]] if len(filieres)==1 and filieres[0] in FILIERE_NAMES else None
    classe_label = CLASS_NAMES[classes[0]] if len(classes)==1 and classes[0] in CLASS_NAMES else None
    return {
        "NumerosTrouvés": nums,
        "NumerosConnus": connus,
        "NumerosInconnus": inconnus,
        "FiliereDéduite": filiere_label,
        "ClasseDéduite": classe_label,
    }

# ============================= IMPORT I3/I4+ =============================
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

def autodetect_phone_column(columns: List[str]) -> Optional[str]:
    lower_map = {c: str(c).strip().lower() for c in columns}
    keys = ["téléphone", "telephone", "tel", "phone", "portable", "mobile"]
    for c, l in lower_map.items():
        if any(k in l for k in keys):
            return c
    return None

def detect_data_start(raw: pd.DataFrame, groupes_col_idx: int, header_row_idx: int) -> int:
    start_probe = header_row_idx + 1
    max_probe = min(len(raw), header_row_idx + 50)
    for r in range(start_probe, max_probe):
        val = raw.iat[r, groupes_col_idx] if groupes_col_idx < raw.shape[1] else None
        if pd.notna(val) and str(val).strip() != "":
            return r
    return start_probe

# --------------------------- Sidebar (commune) ---------------------------
with st.sidebar:
    st.header("⚙️ Import")
    use_sheet = st.text_input("Nom de l'onglet (laisser vide pour auto)", value="")
    col_letter_override = st.text_input("Colonne Groupes (défaut I)", value="I")
    start_row_manual = st.number_input("Forcer ligne de départ (0 = auto)", min_value=0, value=0, step=1)
    show_debug = st.checkbox("Afficher colonnes techniques", value=False)
    st.markdown("---")
    st.header("🧭 Colonnes Nom/Prénom/Téléphone")
    st.caption("Auto-détection, mais tu peux forcer plus bas dans chaque onglet.")
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

# --- I3 / headers / data cut ---
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

digits4 = data[GROUPES_COL_NAME].astype(str).str.count(r"\d{4,}").sum()
if digits4 == 0:
    st.warning("⚠️ La colonne **Groupes** semble vide ou mal alignée. "
               "Vérifie l’onglet et/ou force la ligne de départ dans la barre latérale.")
    st.write("Aperçu des 10 premières valeurs de la colonne Groupes :")
    st.write(data[GROUPES_COL_NAME].head(10))

# --------------------------- Onglets ---------------------------
tab_verif, tab_pdf = st.tabs(["✅ Vérification", "🧾 Listes PDF (1 page = 1 classe)"])

# =========================
# Onglet 1 : Vérification
# =========================
with tab_verif:
    st.subheader("Paramètres colonnes (Vérification)")
    nom_guess, prenom_guess = autodetect_name_columns(list(data.columns))
    col1, col2 = st.columns(2)
    with col1:
        nom_col = st.selectbox("Colonne Nom", options=["—"] + list(data.columns),
                               index=(["—"] + list(data.columns)).index(nom_guess) if nom_guess in (["—"] + list(data.columns)) else 0,
                               key="nom_verif")
    with col2:
        prenom_col = st.selectbox("Colonne Prénom", options=["—"] + list(data.columns),
                                  index=(["—"] + list(data.columns)).index(prenom_guess) if prenom_guess in (["—"] + list(data.columns)) else 0,
                                  key="prenom_verif")
    nom_col = None if nom_col == "—" else nom_col
    prenom_col = None if prenom_col == "—" else prenom_col
    if not nom_col or not prenom_col:
        st.warning("⚠️ Choisis/valide les colonnes **Nom** et **Prénom** pour un export d'erreurs correct.")

    # Analyse
    df = data.copy()
    df["Diagnostic"] = df[GROUPES_COL_NAME].apply(analyser_groupes)
    extras = df[GROUPES_COL_NAME].apply(extra_info).apply(pd.Series)
    df = pd.concat([df, extras], axis=1)

    # Répartition
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

    # Tableau
    base_cols = [c for c in df.columns if c not in [GROUPES_COL_NAME, "Diagnostic", "FiliereDéduite", "ClasseDéduite", "NumerosTrouvés", "NumerosConnus", "NumerosInconnus"]]
    display_cols = base_cols + [GROUPES_COL_NAME, "Diagnostic"]
    if st.checkbox("Afficher colonnes techniques", value=show_debug, key="tech_verif"):
        display_cols += ["FiliereDéduite", "ClasseDéduite", "NumerosTrouvés", "NumerosConnus", "NumerosInconnus"]
    st.markdown("### Données vérifiées")
    st.dataframe(df[display_cols], use_container_width=True)

    # Export JSON
    records = df.to_dict(orient="records")
    json_bytes = json.dumps(records, ensure_ascii=False, indent=2).encode("utf-8")
    st.download_button("⬇️ Télécharger JSON (complet)", data=json_bytes, file_name="export_verifie.json", mime="application/json", key="json_verif")

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
                           file_name="erreurs_groupes.csv", mime="text/csv", key="csv_erreurs")

# =========================
# Onglet 2 : PDF (1 page = 1 classe, jusqu'à 42 élèves)
# =========================
with tab_pdf:
    st.subheader("Paramètres colonnes (PDF)")
    nom_guess2, prenom_guess2 = autodetect_name_columns(list(data.columns))
    tel_guess = autodetect_phone_column(list(data.columns))
    c1, c2, c3 = st.columns(3)
    with c1:
        nom_col_pdf = st.selectbox("Colonne Nom", options=["—"] + list(data.columns),
                                   index=(["—"] + list(data.columns)).index(nom_guess2) if nom_guess2 in (["—"] + list(data.columns)) else 0,
                                   key="nom_pdf")
    with c2:
        prenom_col_pdf = st.selectbox("Colonne Prénom", options=["—"] + list(data.columns),
                                      index=(["—"] + list(data.columns)).index(prenom_guess2) if prenom_guess2 in (["—"] + list(data.columns)) else 0,
                                      key="prenom_pdf")
    with c3:
        tel_col_pdf = st.selectbox("Colonne Téléphone", options=["—"] + list(data.columns),
                                   index=(["—"] + list(data.columns)).index(tel_guess) if tel_guess in (["—"] + list(data.columns)) else 0,
                                   key="tel_pdf")
    nom_col_pdf = None if nom_col_pdf == "—" else nom_col_pdf
    prenom_col_pdf = None if prenom_col_pdf == "—" else prenom_col_pdf
    tel_col_pdf = None if tel_col_pdf == "—" else tel_col_pdf

    st.markdown("#### Aperçu (10 lignes)")
    st.dataframe(data.head(10), use_container_width=True)

    # Préparer : classes -> liste d'étudiants (Nom, Prénom, Téléphone)
    classes_to_students: Dict[int, list] = defaultdict(list)

    def classes_for_row(nums: List[int]) -> Set[int]:
        return {n for n in nums if n in CLASS_NAMES}

    for _, row in data.iterrows():
        nums = parse_numeros(row.get(GROUPES_COL_NAME))
        cls = classes_for_row(nums)
        if not cls:
            continue
        nom_v = "" if not nom_col_pdf else str(row.get(nom_col_pdf, "") or "")
        prenom_v = "" if not prenom_col_pdf else str(row.get(prenom_col_pdf, "") or "")
        tel_v = "" if not tel_col_pdf else str(row.get(tel_col_pdf, "") or "")
        for c in cls:
            classes_to_students[c].append((nom_v, prenom_v, tel_v))

    # Génération du PDF : 1 page = 1 classe, police petite, max 42 élèves/page
    MAX_PER_PAGE = 42

    if st.button("🧾 Générer le PDF (1 page = 1 classe, 42 élèves max)"):
        if not nom_col_pdf or not prenom_col_pdf:
            st.error("Sélectionne d'abord **Nom** et **Prénom**.")
        else:
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(
                buffer, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36,
                title="Listes par classe", author="Auto"
            )
            styles = getSampleStyleSheet()
            class_style = ParagraphStyle("ClassTitle", parent=styles["Heading1"], fontSize=16, leading=19, spaceAfter=8)

            elements = []

            ordered_classes = sorted(classes_to_students.keys(), key=lambda c: CLASS_NAMES.get(c, str(c)))

            for ccode in ordered_classes:
                class_title = CLASS_NAMES.get(ccode, f"Classe {ccode}")
                rows_sorted = sorted(classes_to_students[ccode], key=lambda t: ((t[0] or "").lower(), (t[1] or "").lower()))
                page_rows = rows_sorted[:MAX_PER_PAGE]  # <= 42 élèves

                page_block = []
                page_block.append(Paragraph(class_title, class_style))
                page_block.append(Spacer(1, 6))

                data_rows = [["Nom", "Prénom", "Téléphone"]]
                data_rows += [[n, p, t] for (n, p, t) in page_rows]

                # Largeurs adaptées A4; petite police pour tenir 42 lignes
                tbl = Table(data_rows, colWidths=[220, 220, 83])
                tbl.setStyle(TableStyle([
                    ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
                    ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
                    ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                    ("ALIGN", (0,0), (-1,0), "CENTER"),
                    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                    ("FONTSIZE", (0,0), (-1,-1), 9),
                    ("BOTTOMPADDING", (0,0), (-1,0), 4),
                    ("TOPPADDING", (0,1), (-1,-1), 2),
                    ("BOTTOMPADDING", (0,1), (-1,-1), 2),
                ]))
                page_block.append(tbl)

                elements.append(KeepTogether(page_block))
                elements.append(PageBreak())

            if elements and isinstance(elements[-1], PageBreak):
                elements = elements[:-1]

            doc.build(elements)
            buffer.seek(0)
            st.download_button("⬇️ Télécharger le PDF", data=buffer, file_name="listes_par_classe.pdf",
                               mime="application/pdf", key="pdf_download_classes")
