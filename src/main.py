import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
factures = pd.read_excel("data/factures.xlsx")
total1 = factures["Montant"].sum()
impayees = factures[factures["Statut"] == "Impayee"]
total2 = impayees["Montant"].sum()
payees = factures[factures["Statut"] == "Payee"]
total3 = payees["Montant"].sum()
factures["Date_echeance"] = pd.to_datetime(
    factures["Date_echeance"],
    format="%d/%m/%Y"
)
aujourdhui = pd.Timestamp.today().normalize()
en_retard = factures[
    (factures["Statut"] == "Impayee") &
    (factures["Date_echeance"] < aujourdhui)
]
total_retard = en_retard["Montant"].sum()
en_retard = en_retard.copy()
en_retard["Jours_retard"] = (aujourdhui - en_retard["Date_echeance"]).dt.days
en_retard.to_excel("output/factures_en_retard.xlsx", index=False)
classeur = load_workbook("output/factures_en_retard.xlsx")
feuille = classeur.active

for cellule in feuille[1]:
    cellule.font = Font(bold=True, color="FFFFFF")
    cellule.fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )
feuille.column_dimensions["A"].width = 15
feuille.column_dimensions["B"].width = 20
feuille.column_dimensions["C"].width = 20
feuille.column_dimensions["D"].width = 20
feuille.column_dimensions["E"].width = 20
feuille.column_dimensions["F"].width = 20
feuille.column_dimensions["G"].width = 20
feuille["I1"] = "RESUME FINANCIER"

feuille["I2"] = "Montant total"
feuille["J2"] = total1

feuille["I3"] = "Montant impayé"
feuille["J3"] = total2

feuille["I4"] = "Montant payé"
feuille["J4"] = total3

feuille["I5"] = "Montant en retard"
feuille["J5"] = total_retard
feuille.column_dimensions["I"].width = 25
feuille.column_dimensions["J"].width = 18
for cellule in feuille["E"][1:]:
    cellule.number_format = "DD/MM/YYYY"
for cellule in feuille["J"][1:5]:
    cellule.number_format = '#,##0.00 "€"'
feuille["I1"].font = Font(bold=True, color="FFFFFF")
feuille["I1"].fill = PatternFill(
    start_color="1F4E78",
    end_color="1F4E78",
    fill_type="solid"
)
feuille["J3"].font = Font(bold=True, color="FF0000")
feuille["J4"].font = Font(bold=True, color="008000")
feuille["J5"].font = Font(bold=True, color="FF8C00")
classeur.save("output/factures_en_retard.xlsx")
print("===== FACTUREPILOT =====")
print("Montant total :", total1, "€")
print("Montant total des factures impayees :", total2, "€")
print("Montant total des factures payees :", total3, "€")
print("Montant total en retard :", total_retard, "€")
print("Rapport Excel genere dans output/factures_en_retard.xlsx")