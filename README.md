# stochdorphia-kettenbrief

### Folgendes muss installiert sein
- Python 3.10 + pip (pip muss meistens beim installieren extra ausgewählt werden)
- pipenv ("pip install pipenv" auf der shell eingeben)

### Ausführen
Lade den Inhalt dieses Repositories runter und entpacke ihn ggf.

Es werden die Dateien example.xlsx und template.docx erwartet.
Die Dateinamen können in der main.py konfiguriert werden wie auch die Zeile bei der begonnen werden soll.

Die Spalten der Excel werden der Reihe nach, aufsteigend mit den suchzeichen ( z.B. wird "%2%" durch den Wert in der 2. Spalte ersetzt) ersetzt.
Eine Ausnahme bilder hier das suchzeichen "%anrede%". Dieser Wert wird anhand der ersten Spalte (Herr/Frau) entschieden.
Bei der ersten Reihe ohne einen Wert in der ersten Spalte wird abgebrochen. 

Führe folgende Befehle in der shell aus (bei 1. muss der pfad entsprechend ersetzt werden):
1. cd /pfad/zum/ordner
2. pipenv install
3. python main.py