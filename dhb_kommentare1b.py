import pandas as pd

# === Dateien laden ===
df_kommentare = pd.read_excel("transformierte_tabelle_mit_kommentaren.xlsx")
df_dsb = pd.read_excel("ziel_dsb_NUR2023.xlsx")

# === DSB-Datei bereinigen ===
df_dsb = df_dsb[['Variablenname', 'Änderungsbeschreibung (zu Vorjahr)']].copy()
df_dsb['Variablenname'] = df_dsb['Variablenname'].astype(str).str.strip()
df_dsb['Änderungsbeschreibung (zu Vorjahr)'] = df_dsb['Änderungsbeschreibung (zu Vorjahr)'].astype(str).str.strip()

# Nur Einträge mit echtem Inhalt behalten
df_dsb = df_dsb[
    df_dsb['Änderungsbeschreibung (zu Vorjahr)'].notna() &
    (df_dsb['Änderungsbeschreibung (zu Vorjahr)'] != "") &
    (df_dsb['Änderungsbeschreibung (zu Vorjahr)'].str.lower() != "nan")
]

# === Kommentartabelle vorbereiten ===
df_kommentare['Variablenname'] = df_kommentare['Variablenname'].astype(str).str.strip()
df_kommentare['Ziel DSB Kommentare'] = ""  # Neue Spalte

# Ergebnisliste
ergebnis = []

# Vorhandene Variablennamen merken
vorhandene_variablen = set(df_kommentare['Variablenname'])

# === Teil 1: Kommentare einfügen unter vorhandene Variablen ===
bereits_eingefügt = set()  # Set für (Variablenname, Kommentar)

for idx in range(len(df_kommentare)):
    zeile = df_kommentare.iloc[idx]
    varname = zeile['Variablenname']
    ergebnis.append(zeile)

    # Nur hinzufügen, wenn Kommentar existiert und noch nicht eingefügt wurde
    matches = df_dsb[df_dsb['Variablenname'] == varname]
    for _, match_row in matches.iterrows():
        kommentar = match_row['Änderungsbeschreibung (zu Vorjahr)']
        key = (varname, kommentar)
        if kommentar and kommentar.lower() != "nan" and key not in bereits_eingefügt:
            neue_zeile = zeile.copy()
            neue_zeile[:] = ""
            neue_zeile['Variablenname'] = varname
            neue_zeile['Ziel DSB Kommentare'] = kommentar
            ergebnis.append(neue_zeile)
            bereits_eingefügt.add(key)

# === Teil 2: Variablen, die bisher nicht vorkamen ===
fehlende_vars = df_dsb[~df_dsb['Variablenname'].isin(vorhandene_variablen)]

for _, row in fehlende_vars.iterrows():
    kommentar = row['Änderungsbeschreibung (zu Vorjahr)']
    key = (row['Variablenname'], kommentar)
    if kommentar and kommentar.lower() != "nan" and key not in bereits_eingefügt:
        neue_zeile = pd.Series({col: "" for col in df_kommentare.columns})
        neue_zeile['Variablenname'] = row['Variablenname']
        neue_zeile['Ziel DSB Kommentare'] = kommentar
        ergebnis.append(neue_zeile)
        bereits_eingefügt.add(key)

# === Zusammenführen und leere Zellen sichern ===
df_final = pd.DataFrame([z.to_dict() for z in ergebnis])
df_final.fillna("", inplace=True)

# === Export ===
df_final.to_excel("transformierte_tabelle_mit_kommentaren_und_zieldsb.xlsx", index=False)
print("Datei 'transformierte_tabelle_mit_kommentaren_und_zieldsb.xlsx' wurde erfolgreich erstellt.")
