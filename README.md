# FacturePilot

FacturePilot est un outil Python qui automatise
l'analyse des factures et la génération de rapports Excel.
## Fonctionnalités

- Lecture automatique des factures depuis un fichier Excel.
- Calcul du montant total des factures.
- Calcul des montants payés et impayés.
- Détection automatique des factures en retard.
- Calcul du nombre de jours de retard.
- Génération d'un rapport Excel avec une mise en forme professionnelle.
- Création d'un résumé financier.
## Installation

### 1. Télécharger le projet

```bash
git clone https://github.com/allaouziadam2-sudo/FacturePilot.git
```

### 2. Entrer dans le dossier

```bash
cd FacturePilot
```

### 3. Installer les bibliothèques Python

```bash
python -m pip install pandas openpyxl
```

### 4. Générer les factures d'exemple

```bash
python src/creer_excel.py
```

### 5. Lancer FacturePilot

```bash
python src/main.py
```

Le rapport Excel sera généré dans le dossier `output/`.