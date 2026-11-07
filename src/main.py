import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
chemin = input("Entrez le chemin du fichier Excel : ")
if chemin.lower().endswith(".xlsx"):
    factures = pd.read_excel(chemin)
elif chemin.lower().endswith(".csv"):
    factures = pd.read_csv(chemin)
else:
    print("Erreur : format de fichier non accepté")
    exit()
colonnes_obligatoires = [
    "Client",
    "Numero_facture",
    "Montant",
    "Date_facture",
    "Date_echeance",
    "Statut"
]

for colonne in colonnes_obligatoires:
    if colonne not in factures.columns:
        print("Erreur : colonne manquante :", colonne)
        exit()
factures["Montant"] = pd.to_numeric(
    factures["Montant"],
    errors="coerce"
)

if factures["Montant"].isna().any():
    print("Erreur : certains montants sont invalides")
    exit()
for colonne in ["Date_facture", "Date_echeance"]:
    factures[colonne] = pd.to_datetime(
        factures[colonne],
        format="%d/%m/%Y",
        errors="coerce"
    )

    if factures[colonne].isna().any():
        print("Erreur : dates invalides dans", colonne)
        exit()
statuts_valides = ["Payee", "Impayee"]

for statut in factures["Statut"]:
    if statut not in statuts_valides:
        print("Erreur : statut invalide :", statut)
        exit()
total1 = factures["Montant"].sum()
impayees = factures[factures["Statut"] == "Impayee"]
total2 = impayees["Montant"].sum()
payees = factures[factures["Statut"] == "Payee"]

total3 = payees["Montant"].sum()
aujourdhui = pd.Timestamp.today().normalize()
# 1. Factures déjà payées
payees = factures[factures["Statut"] == "Payee"]

# 2. Factures encore impayées
impayees = factures[factures["Statut"] == "Impayee"]

# 3. Factures en retard
en_retard = impayees[
    impayees["Date_echeance"] < aujourdhui
].copy()

# 4. Factures à payer aujourd'hui ou prochainement
a_payer = impayees[
    impayees["Date_echeance"] >= aujourdhui
].copy()
# Trier les factures à payer par date d'échéance
a_payer = a_payer.sort_values(
    by="Date_echeance",
    ascending=True
)

# Calculer le nombre de jours restants
a_payer["Jours_restants"] = (
    a_payer["Date_echeance"] - aujourdhui
).dt.days

# Compter les factures à payer
nombre_a_payer = len(a_payer)

# Calculer leur montant total
total_a_payer = a_payer["Montant"].sum()
total_retard = en_retard["Montant"].sum()
en_retard = en_retard.sort_values(
    by="Date_echeance",
    ascending=True
)
en_retard = en_retard.reset_index(drop=True)

en_retard.insert(
    0,
    "Priorite",
    range(1, len(en_retard) + 1)
)
nombre_retard = len(en_retard)

en_retard["Jours_retard"] = (aujourdhui - en_retard["Date_echeance"]).dt.days
# Ajouter une priorité aux factures à payer
a_payer = a_payer.reset_index(drop=True)

a_payer.insert(
    0,
    "Priorite",
    range(1, len(a_payer) + 1)
)
os.makedirs("output", exist_ok=True)
for colonne in ["Date_facture", "Date_echeance"]:
    en_retard[colonne] = en_retard[colonne].dt.strftime("%d/%m/%Y")
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
feuille.column_dimensions["H"].width = 20
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
# Préparer les dates pour le rapport
a_payer_export = a_payer.copy()

for colonne in ["Date_facture", "Date_echeance"]:
    a_payer_export[colonne] = (
        a_payer_export[colonne].dt.strftime("%d/%m/%Y")
    )

# Créer un rapport avec trois feuilles
with pd.ExcelWriter(
    "output/rapport_facturepilot.xlsx",
    engine="openpyxl"
) as writer:

    # Première feuille : factures en retard
    en_retard.to_excel(
        writer,
        sheet_name="Factures en retard",
        index=False
    )

    # Deuxième feuille : factures à payer
    a_payer_export.to_excel(
        writer,
        sheet_name="A payer",
        index=False
    )

    # Troisième feuille : factures payées
    payees_export = payees.copy()

    for colonne in ["Date_facture", "Date_echeance"]:
        payees_export[colonne] = (
            payees_export[colonne].dt.strftime("%d/%m/%Y")
        )

    payees_export.to_excel(
        writer,
        sheet_name="Factures payees",
        index=False
    )
print("Nouveau rapport : output/rapport_facturepilot.xlsx")
print("===== FACTUREPILOT =====")
print("Montant total :", total1, "€")
print("Montant total des factures impayees :", total2, "€")
print("Montant total des factures payees :", total3, "€")
print("Montant total en retard :", total_retard, "€")
print("Rapport Excel genere dans output/factures_en_retard.xlsx")
print("Nombre de factures en retard :", nombre_retard)
print("Nombre de factures à payer :", nombre_a_payer)
print("Montant à payer :", total_a_payer, "€")
print(a_payer[["Client", "Date_echeance", "Jours_restants"]])