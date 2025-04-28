import requests
from bs4 import BeautifulSoup

# Die Liste der Variablen, die überprüft werden sollen
variables_to_check = [
    "hs090", "hs110", "hs040", "hs050", "hh050", "hd080", 
    "ev0100p", "ev0200p", "pd020", "pd080", "rl030", "rl060", 
    "ph010", "ph030", "hs011", "hs031", "hs021", "bd0101h", 
    "bd0201h", "bd0301h", "bd0401h", "er0908p", "er0905p", 
    "ph040", "ph050", "ph060", "ph070", "hs200", "hs210", 
    "hs220", "ph110a", "ph110b", "ph101", "ph111", "ph121", 
    "ph131", "ph141", "ph151", "ps050", "ps060", "ps070", 
    "ps080", "ps110", "ps111", "ps102", "hs180", "hs190",
    "hd100", "hd110", "hd120", "hd140", "hd150", "hd160", "hd170", "hd180", "hd190", "hd200", "hd240",
    "cb0900h", "cb0600h", "cb1000h", "cb0700h", "cb1100h", "cb0800h", "hh040", "hs160", "hs170"
]

# URL der Seite
url = "https://www.gesis.org/en/missy/metadata/EU-SILC/2023/Cross-sectional/original"

# HTML von der Seite herunterladen
response = requests.get(url)

# Überprüfen, ob die Anfrage erfolgreich war
if response.status_code == 200:
    # HTML-Inhalt mit BeautifulSoup parsen
    soup = BeautifulSoup(response.text, "html.parser")
    
    # Alle Textinhalte der Seite durchsuchen (z.B. Tabellen, Links, Absätze)
    page_text = soup.get_text().lower()  # Wir konvertieren alles zu Kleinbuchstaben, um die Suche case-insensitiv zu machen

    # Überprüfen, welche Variablen in der Liste auf der Seite vorkommen
    found_variables = []
    not_found_variables = []

    for var in variables_to_check:
        if var.lower() in page_text:
            found_variables.append(var)
        else:
            not_found_variables.append(var)
    
    # Zusammenfassung der gefundenen und nicht gefundenen Variablen
    print("Gefundene Variablen:")
    for var in found_variables:
        print(f"- {var}")

    print("\nNicht gefundene Variablen:")
    for var in not_found_variables:
        print(f"- {var}")
else:
    print(f"Fehler beim Abrufen der Seite: {response.status_code}")