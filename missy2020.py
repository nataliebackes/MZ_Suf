import requests
from bs4 import BeautifulSoup

# Liste der Variablen, die überprüft werden sollen
variables_to_check = [
    'aa0201p', 'aa0500h', 'aa1201p', 'aa1202p', 'aa1300p', 'aa1401h', 'aa1403h', 
    'aa1402h', 'aa1404h', 'aa1521h', 'aa1522h', 'aa1523h', 'aa1524h', 'aa1525h',
    'aa1701h', 'aa1702h', 'aa1703h', 'aa1704h', 'aa1705h', 'aa1801h', 'aa1802h',
    'aa1803h', 'aa1804h', 'aa1805h', 'aa1811h', 'aa1812h', 'aa1813h', 'aa1814h', 
    'aa1815h', 'aa3602h', 'aa3604h', 'aa3601h', 'aa3603h', 'aa3721h', 'aa3722h', 
    'aa3723h', 'aa3724h', 'aa3725h', 'aa2400h', 'aa2501h', 'aa2502h', 'aa2600h', 
    'aa2700h', 'aa2801h', 'aa2802h', 'ba0400h', 'ba0500h', 'ba0600h', 'ba0800h', 
    'ba1100h', 'ba1800h', 'ba1801h', 'ba1802h', 'ba1803h', 'ba1902h', 'ba2201h', 
    'ba2003h', 'ba2004h', 'ba2302h', 'ba2304h', 'ba2306h', 'ba2308h', 'ba2611h', 
    'ba2602h', 'ba2603h', 'ba3001h', 'ba2900h', 'ba3200h', 'ba3301h', 'ba3302h', 
    'ba3303h', 'ba3304h', 'ba3305h', 'ba3306h', 'ba3800h', 'ba4001h', 'ba4002h', 
    'ba4003h', 'ba4004h', 'bb0101h', 'bb0102h', 'bb0201h', 'bb0202h', 'bb0203h', 
    'bb0204h', 'bb0300h', 'bb0400h', 'bb0600h', 'bb0700h', 'bb0733h', 'ba4201h', 
    'ba4202h', 'ba4203h', 'ba4204h', 'ba4205h', 'ba4301h', 'ba4302h', 'ba4303h', 
    'bc0100h', 'bc0200h', 'bc0300h', 'bc0400h', 'bc0600h', 'bc0601p', 'bc0602p', 
    'bc0603p', 'bc0800h', 'bc0801p', 'bc0802p', 'bc1000h', 'bc1001p', 'bc1002p', 
    'bc1003p', 'bc1400h', 'bc1401p', 'bc1402p', 'bc1403p', 'bc1100h', 'bc1200h', 
    'bd0101h', 'bd0102h', 'bd0105h', 'bd0103h', 'bd0107h', 'bd0201h', 'bd0202h', 
    'bd0204h', 'bd0203h', 'bd0206h', 'bd0301h', 'bd0302h', 'bd0304h', 'bd0303h', 
    'bd0306h', 'bd0401h', 'bd0402h', 'bd0404h', 'bd0403h', 'bd0406h', 'be0101h', 
    'be0102h', 'be0104h', 'be0103h', 'be0106h', 'be0201h', 'be0202h', 'be0204h', 
    'be0203h', 'be0206h', 'be0301h', 'be0302h', 'be0304h', 'be0306h', 'be0307h', 
    'be0400h', 'be0500h', 'be0600h', 'be0600p', 'be0603p', 'be0602p', 'be0604p', 
    'be0606p', 'be0900h', 'be0607p', 'be0608p', 'be0609p', 'be0610p', 'be0612p', 
    'be0700h', 'be0800h', 'bf0100h', 'bf0101h', 'bf0102h', 'bf0103h', 'bf0104h', 
    'bf0301h', 'bf0302h', 'bf0303h', 'bf0401h', 'bf0500h', 'bf1000h', 'bf0600h', 
    'bf1200h', 'bf0700h', 'bf0800h', 'bf1300h', 'bf0900h', 'bf1500h', 'ca0400h', 
    'ca0700h', 'ca0502p', 'ca0504p', 'ca0506p', 'ca0508p', 'ca0512p', 'ca0514p', 
    'ca0515p', 'ca0520p', 'cc0000h', 'cc0100h', 'cc0200h', 'cc0300h', 'dc0201p', 
    'dc0202p', 'dc0801p', 'eb0800p', 'eb1000p', 'ed0701p', 'ed0702p', 'ed1801p', 
    'ed1802p', 'eh0000p', 'eh0100p', 'eh0200p', 'ej0500p', 'ea1200p', 'ea1300p', 
    'ea1400p', 'ej0501p', 'ej0502p', 'ej0508p', 'ej0513p', 'ea1401p', 'dg0302h', 
    'dg0303h', 'dg0304h', 'dg0305h', 'dg0306h', 'ep1105p', 'ep1106p', 'ep1107p', 
    'er1001p', 'de0602p', 'de0601pu01', 'de0603p', 'de0601pu02', 'de0604p', 'de0601pu03', 
    'de0605p', 'de0601pu04', 'de0606p', 'de0601pu05', 'de0607p', 'de0601pu06', 'de0608p', 
    'de0601pu07', 'de0609p', 'de0601pu08', 'de0610p', 'de0601pu09', 'de0611p', 'de0601pu10', 
    'er0900p', 'er0901p', 'er0902p', 'er0904p', 'er0905p', 'er0906p', 'er0907p', 'er0908p', 
    'er0909p', 'er0910p', 'eu0301p', 'eu0302p', 'eu0303p', 'eu0304p', 'eu0305p', 'eu0306p', 
    'eu2000p', 'eu2001p', 'eu2022p', 'ev0100p', 'ev0200p', 'eu2024p', 'eu2025p', 'eu2026p', 
    'eu2027p', 'eu2028p', 'eu2029p', 'eu2030p', 'eu2031p', 'eu2032p', 'eu2033p', 'eu2034p', 
    'eu2035p', 'eu2036p', 'eu2037p', 'eu2038p', 'eu2039p', 'eu2040p', 'eu2041p', 'eu2042p', 
    'fa0101p', 'fa0102p', 'fa0103p', 'fa0104p', 'fa0105p', 'fa0106p', 'fa0107p', 'fa0108p', 
    'fa0109p', 'fa0110p', 'fa0111p', 'fa0112p', 'fa0113p', 'fa0200p', 'fa0201p', 'fa0203p', 
    'fa0204p', 'fb0100p', 'fb0201p', 'fb0202p', 'fb0301p', 'fb0302p', 'fb0601p', 'fb0701p', 
    'fb0801p', 'fb0901p', 'fb1001p', 'fb1101p', 'fb1600p', 'fb0507p', 'fb0401p', 'fb0402p', 
    'fb1201p', 'fb1301p', 'fb1302p', 'fc0100p', 'fc0201p', 'fc0202p', 'fc0203p', 'fc0204p', 
    'fc0205p', 'fc1201p', 'fc1202p', 'fc1203p', 'fc1204p', 'fc1205p', 'fc0301p', 'fc0302p', 
    'fc0303p', 'fc0304p', 'fc0305p', 'fc0401p', 'fc0402p', 'fc0403p', 'fc0404p', 'fc0405p', 
    'fc0501p', 'fc0502p', 'fc0503p', 'fc0504p', 'fc0505p', 'fc0801p', 'fc0802p', 'fc0803p', 
    'fc0804p', 'fc0805p', 'fc0701p', 'fc0702p', 'fc1301p', 'fc1302p', 'fc0601p', 'fc0602p', 
    'fd1401p', 'fd1402p', 'fc1001p', 'fc1002p', 'fc1003p', 'fc1004p', 'fc1005p', 
    'fc1100pu01', 'fc1100pu02', 'fc1100pu03', 'fc1100pu04', 'fc1100pu05', 'fc1100pu06', 
    'fc1100pu07', 'fd0101p', 'fd0102p', 'fd0201p', 'fd0202p', 'fd0301p', 'fd0302p', 
    'fd0401p', 'fd0402p', 'fd0501p', 'fd0502p', 'fd0601p', 'fd0602p', 'fd0701p', 
    'fd0702p', 'fd0801p', 'fd0802p', 'fd0803p', 'fd0901p', 'fd0902p', 'fd1001p', 
    'fd1002p', 'fd1201p', 'fd1101p', 'fd1102p', 'fd1103p', 'fd1104p', 'fd1105p', 'fd1106p',
    'fd1107p', 'fd1108p', 'fd1109p', 'fd1110p', 'fd1111p', 'fd1112p', 'fd1113p',
    'fd1114p', 'fd1116p', 'fd1118p', 'fb0306p', 'fb0307p', 'fb1501p', 'fb1502p',
    'fb1505p', 'fe0101p', 'fe0102p', 'fe0201p', 'fe0202p'
]

# URL der Seite
url = "https://www.gesis.org/missy/metadata/MZ/2020/Cross-sectional/original"

# Abrufen der HTML-Seite
response = requests.get(url)

# Überprüfen, ob die Anfrage erfolgreich war
if response.status_code == 200:
    # Parse den HTML-Inhalt mit BeautifulSoup
    soup = BeautifulSoup(response.text, 'html.parser')

    # Finde alle Textabschnitte, die Variablennamen enthalten könnten
    found_variables = []
    for var in variables_to_check:
        if var.lower() in soup.get_text().lower():  # Ignoriere Groß- und Kleinschreibung
            found_variables.append(var)

    # Ausgabe der gefundenen Variablen
    if found_variables:
        print("Gefundene Variablen:")
        for var in found_variables:
            print(var)
    else:
        print("Keine der Variablen wurden auf der Seite gefunden.")
else:
    print(f"Fehler beim Abrufen der Seite: {response.status_code}")