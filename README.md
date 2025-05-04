# 2. Ziel DSB 
## 2.5 SILC:

### 1) 2022: welche Vars waren drinnen? -> DHB, Fachkonzept
### 1.1 Variablenliste 
  - DHB: Varnamen ausgeben lassen, bei denen es -4 nicht in Silc gibt (-> noch nicht final)
  - Fachkonzept: Vars händisch rauskopieren
  - aus beidem auch die zugrundeliegenden Vars der Indizes rausschreiben
  => daraus Variablenliste in Excel "silc_vars2022" erstellen

### 1.2 Abgleich: welche sind wieder drinnen, welche nicht?
  - MFB Abgleich: wird es noch gefragt?
  - Ziel DSB: ist es noch im SUF?
  - MFB muss angepasst werden, weil python sonst wegen zusammengefassten Spalten rumnervt -> einfach in Ziel DSB Ordner neu abspeichern und 1. Zeile entfernen 
  => dafür python skript benutzen "silcvars.py" (-> achtung der summary report buggt noch, aber lässt sich easy aus konsole kopieren)
  - infos in silc_vars2022 Excel integrieren
  => dann checken: passt das alles?
   
### 1.3 SILC Vars für Indizes
  - gibt es die zugrundeliegenden Vars in 2023?
  => alle Vars der Indizes in SILC Missy suchen -> python skript missytest.py und in excel silc_vars2022 ergänzen 

### 1.4 Was ist 2023 auf Ja?
  - händisch aus MFB alle Onsite SILC Vars rausholen
- Filtern: Fragenummer nur in SILC, dann alle Vars aus Spalte P und Q
  - sind die jetzt, Vorjahr und 3 Jahre vorher (3-jährliches Silc Modul) in DSB auf ja/nein? 
  - Prüfung: dsb2020.py => report Excel
  => hier checken: wenn alle gleich (alles ja/alles nein) OK
  => ansonsten checken ob zB ZP oÄ 

 	- Da ziel dsb 2020 schrott -> missy variablenübersicht für skript genommen 
  - missy2020.py
  - die dann in report excel eingefügt mit dsb2020fix.py

