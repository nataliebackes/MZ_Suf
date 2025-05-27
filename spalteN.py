import pandas as pd
import unicodedata
import re

# ======================================
# 📁 Pfade zu den Excel-Dateien
# ======================================
excel1_path = r"C:\Users\backesne\Desktop\MFB Spalte N\Mappe1.xlsx"
excel2_path = r"C:\Users\backesne\Desktop\MFB Spalte N\ziel_dsb_2023.xlsx"
excel3_path = r"C:\Users\backesne\Desktop\MFB Spalte N\MZ_2022_Masterfragebogen_FINALFINAL.xlsx"
ergebnis_path = r"C:\Users\backesne\Desktop\MFB Spalte N\Ergebnis_Zuordnung.xlsx"
report_path = r"C:\Users\backesne\Desktop\MFB Spalte N\Ergebnis_Report.xlsx"

# ======================================
# 📄 Dateien laden und Spaltennamen bereinigen
# ======================================
df1 = pd.read_excel(excel1_path)
df2 = pd.read_excel(excel2_path)
df3 = pd.read_excel(excel3_path)

df1.columns = df1.columns.str.strip().str.replace('\n', '', regex=False)
df2.columns = df2.columns.str.strip().str.replace('\n', '', regex=False)
df3.columns = df3.columns.str.strip().str.replace('\n', '', regex=False)

# ======================================
# 🔧 Relevante Spaltennamen definieren
# ======================================
col_n_2023 = "Variablenname 2023 (SUF)"
col_p_2023 = "Variablenname 2023 (Onsite)"
col_n_2022 = "Variablenname 2022 (SUF)"
col_p_2022 = "Variablenname 2022 (Onsite)"
col_ja_flag = "im_suf_2023"
col_varname = "Variablenname"
col_bez = "Bezeichnung"
col_j_text = "Fragetext / Erläuterungen / Antwortvorgaben"

# ======================================
# 🔎 Hilfsfunktion zur Vereinheitlichung
# ======================================
def normalize(val):
    if not isinstance(val, str):
        return ""
    val = val.strip()
    val = re.sub(r'\s+', ' ', val)
    val = unicodedata.normalize("NFKC", val)
    return val.casefold()

# ======================================
# 📌 Schritt 0: Vorauswahl – nur "ja"-markierte Variablen aus Excel 2
# ======================================
df2_ja = df2[df2[col_ja_flag].astype(str).str.lower().str.strip() == 'ja'].copy()
df2_ja['Zuordnung'] = 'nicht zugeordnet'

# ======================================
# ✅ Schritt 1: Direkter Abgleich (Excel 1 ↔ Excel 2)
# ======================================
eingetragen_1 = 0

for idx, row in df1.iterrows():
    eintraege = str(row.get(col_p_2023, '')).split(',')
    gefundene_variablen = []
    for eintrag in eintraege:
        eintrag_clean = normalize(eintrag)
        match = df2_ja[df2_ja[col_varname].apply(normalize) == eintrag_clean]
        if not match.empty:
            gefundene_variablen.append(eintrag.strip())
            df2_ja.loc[match.index, 'Zuordnung'] = 'Excel 1'
    if gefundene_variablen:
        df1.at[idx, col_n_2023] = ', '.join(gefundene_variablen)
        eingetragen_1 += 1

print(f"✅ Schritt 1: {eingetragen_1} Einträge durch direkten Vergleich eingetragen.")

# ======================================
# ✅ Schritt 2a: Abgleich über Fragetext
# ======================================
eingetragen_2 = 0
zugewiesene_vars = df1[col_n_2023].dropna().astype(str).str.split(',').explode().str.strip().str.lower().unique()

for idx, row in df2_ja[df2_ja['Zuordnung'] == 'nicht zugeordnet'].iterrows():
    varname = normalize(row[col_varname])
    if varname in zugewiesene_vars:
        continue

    text_b = normalize(row[col_bez])
    if not text_b:
        continue  # Bezeichnung leer → überspringen

    match = df1[df1[col_j_text].apply(normalize) == text_b]
    if not match.empty:
        for m_idx in match.index:
            exist = str(df1.at[m_idx, col_n_2023]) if pd.notna(df1.at[m_idx, col_n_2023]) else ''
            if varname not in exist.lower():
                neu = (exist + ', ' + row[col_varname]).strip(', ')
                df1.at[m_idx, col_n_2023] = neu
        df2_ja.at[idx, 'Zuordnung'] = 'Excel 2'
        eingetragen_2 += 1

print(f"✅ Schritt 2a: {eingetragen_2} Einträge über Fragetext erkannt.")

# ======================================
# ✅ Schritt 2b: Rückverlinkung über Excel 3
# ======================================
eingetragen_3 = 0

def contains_varname_in_suf_list(cell, target):
    if not isinstance(cell, str):
        return False
    parts = [normalize(part) for part in cell.split(';')]
    return normalize(target) in parts

for idx, row in df2_ja[df2_ja['Zuordnung'] == 'nicht zugeordnet'].iterrows():
    varname = normalize(row[col_varname])
    if not varname:
        continue

    match_excel3 = df3[df3[col_n_2022].apply(lambda cell: contains_varname_in_suf_list(cell, varname))]
    if match_excel3.empty:
        continue

    for _, r3 in match_excel3.iterrows():
        pot_onsite = r3[col_p_2022]
        if pd.isna(pot_onsite) or not str(pot_onsite).strip():
            continue

        pot_onsite_norm = normalize(str(pot_onsite))
        match_excel1 = df1[df1[col_p_2023].apply(normalize) == pot_onsite_norm]

        if not match_excel1.empty:
            for m_idx in match_excel1.index:
                exist = str(df1.at[m_idx, col_n_2023]) if pd.notna(df1.at[m_idx, col_n_2023]) else ''
                if varname not in exist.lower():
                    neu = (exist + ', ' + row[col_varname]).strip(', ')
                    df1.at[m_idx, col_n_2023] = neu
                    eingetragen_3 += 1
                    df2_ja.at[idx, 'Zuordnung'] = 'Excel 3'
            break  # nur erste passende Onsite-Zeile verwenden

print(f"✅ Schritt 2b: {eingetragen_3} Einträge über Rückverlinkung erkannt.")

# ======================================
# 💾 Ergebnis speichern
# ======================================
df1_out_path = excel1_path.replace(".xlsx", "_mit_Eintraegen.xlsx")
df1.to_excel(df1_out_path, index=False)
df2_ja.to_excel(ergebnis_path, index=False)
df2_ja.to_excel(report_path, index=False)

print(f"\n🎯 Ergebnisdateien erstellt:")
print(f"🔹 Excel 1 mit Einträgen: {df1_out_path}")
print(f"🔹 Zuordnungsdetails:     {ergebnis_path}")
print(f"🔹 Zuordnungsreport:      {report_path}")
