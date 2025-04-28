import pandas as pd

def search_file1(file1_path, variable_list):
    # Lade Datei 1 mit Header in Zeile 2 (header=1)
    df1 = pd.read_excel(file1_path, header=1)
    df1.columns = df1.columns.str.strip()  # Entfernt unnötige Leerzeichen

    # Spalten P und Q haben in pandas iloc-Index 15 und 16
    col_idx_P = 15
    col_idx_Q = 16

    # Überprüfe, ob die Spalten existieren
    if df1.shape[1] <= col_idx_Q:
        raise IndexError(f"Die Tabelle hat nur {df1.shape[1]} Spalten; Index 16 (Q) ist nicht vorhanden.")

    # Werte in P und Q als Strings
    values_P = df1.iloc[:, col_idx_P].astype(str).values
    values_Q = df1.iloc[:, col_idx_Q].astype(str).values

    # Ergebnis dict: var → True/False
    return {var: {"found_in_P": var in values_P, "found_in_Q": var in values_Q} for var in variable_list}


def search_file2(file2_path, variable_list):
    # Lade Datei 2 mit Header in Zeile 1 (header=0)
    df2 = pd.read_excel(file2_path, header=0)
    df2.columns = df2.columns.str.strip()  # Entfernt unnötige Leerzeichen

    # Spaltennamen
    key_col    = "Variablenname"
    ret_col1   = "im_suf_2022"
    ret_col2   = "im_suf_2023"
    for c in (key_col, ret_col1, ret_col2):
        if c not in df2.columns:
            raise KeyError(f"Spalte '{c}' nicht gefunden. Verfügbare Spalten: {list(df2.columns)}")

    results = {}
    for var in variable_list:
        # Maske auf Schlüsselspalte
        mask = df2[key_col].astype(str) == var
        if mask.any():
            row = df2.loc[mask].iloc[0]
            results[var] = {
                "found":      True,
                ret_col1:     row[ret_col1],
                ret_col2:     row[ret_col2]
            }
        else:
            results[var] = {"found": False, ret_col1: None, ret_col2: None}
    return results


if __name__ == "__main__":
    # Pfade anpassen
    file1 = r"J:\Work\GML\Mikrozensus\MZ_2023\Ziel_DSB\mfb23_2.xlsx"
    file2 = r"J:\Work\GML\Mikrozensus\MZ_2023\Ziel_DSB\ziel_dsb_2023_1004.xlsx"

    variable_list = ["bc0100h","bc0600h","bd0101h","bd0201h","bd0401h","be0301h",
    "fb0201p","fc0201p","fc0401p","fc0601p","fc0701p","fc1001p",
    "fc1100pu01","fc1100pu02","fc1100pu03","fc1100pu04","fc1100pu05","fc1100pu06","fc1100pu07",
    "fc1201p","fd0101p","fd1001p","fd1201p","fd1401p",
    "hs120","hs150","hi010","hi040",
    "pw010","pw191",
    "rch010","rch020","rk050","rk060","pk020",
    "ph122","ph132","ph142","ph152","ph171","ph180",
    "pw241","pw030","pw160","pw120","pw230","pw090",
    "ps010","ps020","ps030","ps040","ps042","ps041", "index_finanz_situation", 
    "index_sozialkapital", 
    "index_lebenssituation", 
    "index_kinderbetreuung_std", 
    "index_gesundheitszustand", 
    "index_zahlungsrueckstaende", 
    "index_oeffentliche_zahlungen", 
    "index_arztkosten", 
    "index_gesundheitsausgaben", 
    "index_bmi", 
    "index_koerperl_beeintraechtig", 
    "index_soziale_kontakte", 
    "index_engagement"]  # Liste von Variablen, die du suchen möchtest

    # Suche in Datei 1 (Spalten P und Q)
    try:
        res1 = search_file1(file1, variable_list)
        print("Ergebnisse in Datei 1 (Spalten P und Q):")
        for v, info in res1.items():
            found_in = []
            if info["found_in_P"]: found_in.append("P")
            if info["found_in_Q"]: found_in.append("Q")
            if found_in:
                print(f"  {v}: {' & '.join(found_in)} gefunden")
            else:
                print(f"  {v}: nicht gefunden")
    except KeyError as e:
        print(f"Fehler in Datei 1: {e}")
    except IndexError as e:
        print(f"Fehler in Datei 1: {e}")

    # Suche in Datei 2
    try:
        res2 = search_file2(file2, variable_list)
        print("\nErgebnisse in Datei 2 (Variablenname → im_suf_2022, im_suf_2023):")
        for v, info in res2.items():
            if info["found"]:
                print(f"  {v}: im_suf_2022={info['im_suf_2022']}, im_suf_2023={info['im_suf_2023']}")
            else:
                print(f"  {v}: nicht gefunden")
    except KeyError as e:
        print(f"Fehler in Datei 2: {e}")
        
        
        
def print_summary(variable_list, res1, res2):
    # Breiten ermitteln
    name_width   = max(len(v) for v in variable_list) + 2
    col1_width   = len("Datei1 (P/Q)") + 2
    col2_width   = len("Datei2 (2022/2023)") + 2

    # Kopfzeile
    header = (
        f"{'Variable'.ljust(name_width)}| "
        f"{'Datei1 (P/Q)'.ljust(col1_width)}| "
        f"{'Datei2 (2022/2023)'.ljust(col2_width)}"
    )
    sep = "-" * len(header)
    print(header)
    print(sep)

    # Zeilen
    for var in variable_list:
        info1 = res1.get(var, {"found_in_P": False, "found_in_Q": False})
        locs = []
        if info1["found_in_P"]: locs.append("P")
        if info1["found_in_Q"]: locs.append("Q")
        status1 = " & ".join(locs) if locs else "–"

        info2 = res2.get(var, {"found": False})
        if info2["found"]:
            status2 = f"{info2['im_suf_2022']}/{info2['im_suf_2023']}"
        else:
            status2 = "–"

        print(
            f"{var.ljust(name_width)}| "
            f"{status1.ljust(col1_width)}| "
            f"{status2.ljust(col2_width)}"
        )
print_summary(variable_list, res1, res2)


data = []
for var in variable_list:
    info1 = res1[var]
    data.append({
        "Variable":       var,
        "in_MFB23":       "ja"  if res2[var] else "nein",
        "in_ZielDSB_P":   "ja"  if info1["found_in_P"] else "nein",
        "in_ZielDSB_Q":   "ja"  if info1["found_in_Q"] else "nein",
    })

df_summary = pd.DataFrame(data)
df_summary.to_excel("summary_report.xlsx", index=False)
print("summary_report.xlsx geschrieben")