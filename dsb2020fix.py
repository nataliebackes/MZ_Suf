import pandas as pd

# Liste der Variablennamen, die überprüft werden sollen
variables_to_check = [
    "bc0100h", "bc0600h", "bd0101h", "bd0201h", "bd0401h", "be0301h", "ed0701p", "ed0702p", "fb0201p",
    "fc0201p", "fc1201p", "fc0401p", "fc0701p", "fc0601p", "fd1401p", "fc1001p", "fc1100pu01", "fc1100pu02",
    "fc1100pu03", "fc1100pu04", "fc1100pu05", "fc1100pu06", "fc1100pu07", "fd0101p", "fd1001p", "fd1201p"
]

# Excel-Datei laden (hier wird angenommen, dass die Datei 'daten.xlsx' heißt)
df = pd.read_excel(r'J:\Work\GML\Mikrozensus\MZ_2023\Ziel_DSB\report.xlsx')

# Überprüfen, ob der Variablenname in Spalte A enthalten ist und in Spalte D eintragen
df['D'] = df['Variable'].apply(lambda x: 'Ja' if x in variables_to_check else 'Nein')

# Excel-Datei mit den Ergebnissen speichern
df.to_excel(r'J:\Work\GML\Mikrozensus\MZ_2023\Ziel_DSB\report.xlsx', index=False)

print("Die Excel-Datei wurde erfolgreich aktualisiert und gespeichert.")
