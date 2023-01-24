# Bewertungen
A small (and shitty) python script to compile a nice list of of grades and other stats of participants of Lina I of a group in a `.xlsx` file.
### Usage

1. Download the list of the participants as a `.xlsx` file. To do this, navigate to `Dieser Kurs -> Teilnehmer`, select `Alle Nutzer/innen auswählen` and then select `Tabellendaten herunterladen als Microsoft Excel (.xlsx)`.
2. Download the grades table. To do this, navigate to `Dieser Kurs -> Berwertungen`. From the dropdown menu select `Export`. 
3. Open the grades table in Excel, and save as `.xlsx`.
4. run ```bash
pip install openpyxl
python main.py
``` And follow the instructions.