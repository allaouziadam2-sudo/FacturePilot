from openpyxl import Workbook

classeur = Workbook()
feuille = classeur.active
feuille.append(["Client", "Numero_facture", "Montant",
                "Date_facture", "Date_echeance", "Statut"])

feuille.append(["Client A", "F001", 250,
                "01/10/2026", "15/10/2026", "Payee"])

feuille.append(["Client B", "F002", 450,
                "05/10/2026", "20/10/2026", "Impayee"])

feuille.append(["Client C", "F003", 300,
                "10/10/2026", "25/10/2026", "Payee"])

feuille.append(["Client D", "F004", 700,
                "15/10/2026", "30/10/2026", "Impayee"])

feuille.append(["Client E", "F005", 150,
                "20/10/2026", "15/11/2026", "Impayee"])

classeur.save("data/factures.xlsx")