import pandas as pd

# === Schritt 1: Originaldatei laden ===
df = pd.read_excel(r"C:\Users\backesne\Desktop\DHB Kommentare 1\Kommentare1.xlsx")

# Entferne leere Variablennamen
df = df[df['Variablenname 2023 (SUF)'].notna()].copy()
df['Variablenname 2023 (SUF)'] = df['Variablenname 2023 (SUF)'].astype(str)

# Mehrere Variablen aufteilen
df_split = df.assign(
    Variablenname=df['Variablenname 2023 (SUF)'].str.split(r',\s*')
).explode('Variablenname').copy()

# Bereinigen
df_split['Variablenname'] = df_split['Variablenname'].str.strip()

# Spalten extrahieren + Fragenummer bereinigen
df_result = df_split[['Variablenname', 'Fragenummer (Master) fortlaufend']].copy()
df_result['Fragenummer (Master) fortlaufend'] = (
    df_result['Fragenummer (Master) fortlaufend']
    .astype(str)
    .str.replace(r"\.0$", "", regex=True)
    .str.strip()
)

# === Schritt 2: Kommentardatei einlesen ===
df_kommentare = pd.read_excel(
    r"C:\Users\backesne\Desktop\DHB Kommentare 1\fb_mz2022_2023_mitMFBNR.xlsx",
    header=5
)

# Spaltennamen bereinigen
df_kommentare.columns = (
    df_kommentare.columns
    .str.replace(u"\xa0", " ", regex=False)
    .str.replace(r"\s+", " ", regex=True)
    .str.strip()
)

# Kommentarspalten + neue Spalte
kommentar_spalten = [
    'Kommentartyp',
    'Kommentar',
    'Interne Hinweise / Hinweise für Readme',
    'Abgleich Kommentar'
]

# Spaltenprüfung
missing = [col for col in kommentar_spalten if col not in df_kommentare.columns]
if missing:
    raise ValueError(f"Kommentarspalten fehlen oder falsch benannt: {missing}")

# Leere Kommentarzeilen filtern (alle 4 Spalten gleichzeitig prüfen)
df_kommentare_filtered = df_kommentare.dropna(subset=kommentar_spalten, how='all').copy()

# MFB-Spalte bereinigen
df_kommentare_filtered['MFB NR fortlaufend'] = (
    df_kommentare_filtered['MFB NR fortlaufend']
    .astype(str)
    .str.replace(r"\.0$", "", regex=True)
    .str.strip()
)

# === Schritt 3: M:N-Merge durchführen ===
df_merged = df_result.merge(
    df_kommentare_filtered[['MFB NR fortlaufend'] + kommentar_spalten],
    left_on='Fragenummer (Master) fortlaufend',
    right_on='MFB NR fortlaufend',
    how='left',
    validate="many_to_many"
)

# Spalte für Join entfernen
df_final = df_merged.drop(columns=['MFB NR fortlaufend'])

# === Schritt 4: Export ===
df_final.to_excel("transformierte_tabelle_mit_kommentaren.xlsx", index=False)
print("Datei 'transformierte_tabelle_mit_kommentaren.xlsx' wurde erfolgreich erstellt.")
